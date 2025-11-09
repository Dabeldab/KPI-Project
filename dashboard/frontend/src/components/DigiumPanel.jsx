import React, { useState, useEffect } from 'react';
import { digiumApi } from '../api';
import { Phone, PhoneCall, PhoneIncoming, PhoneOutgoing, Clock, RefreshCw, Copy } from 'lucide-react';
import CallMonitoring from './CallMonitoring';
import './DigiumPanel.css';

const DigiumPanel = () => {
  const [queueStatus, setQueueStatus] = useState(null);
  const [currentCalls, setCurrentCalls] = useState([]);
  const [monitoringInfo, setMonitoringInfo] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [copiedValue, setCopiedValue] = useState(null);

  const copyToClipboard = (value, label) => {
    navigator.clipboard.writeText(value).then(() => {
      setCopiedValue(label);
      setTimeout(() => setCopiedValue(null), 2000);
    });
  };

  const fetchData = async () => {
    try {
      setLoading(true);
      const [queue, callsResponse, monitoring] = await Promise.all([
        digiumApi.getQueueStatus().catch(() => null),
        digiumApi.getCurrentCalls().catch(() => ({ calls: [] })),
        digiumApi.getMonitoringInfo().catch(() => null)
      ]);
      
      setQueueStatus(queue);
      
      // Handle the new response format with parsed calls
      const calls = callsResponse?.calls || [];
      setCurrentCalls(Array.isArray(calls) ? calls : []);
      
      setMonitoringInfo(monitoring);
      setError(null);
      
      console.log('[DigiumPanel] Fetched calls:', calls);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
    const interval = setInterval(fetchData, 15000); // Refresh every 15 seconds
    return () => clearInterval(interval);
  }, []);

  if (loading && !currentCalls.length) {
    return (
      <div className="digium-panel">
        <h2>📞 Digium Switchvox</h2>
        <div className="loading">Loading...</div>
      </div>
    );
  }

  // Parse call data from API response
  const activeCalls = currentCalls.length || 0;
  const incomingCalls = currentCalls.filter(c => c.direction === 'inbound' || c.direction === 'incoming').length || 0;
  const outgoingCalls = currentCalls.filter(c => c.direction === 'outbound' || c.direction === 'outgoing').length || 0;

  return (
    <div className="digium-panel">
      <div className="panel-header">
        <h2>📞 Digium Switchvox</h2>
        <div className="header-actions">
          <div className="last-update">
            Last updated: {new Date().toLocaleTimeString()}
          </div>
          <button 
            className="refresh-btn-panel" 
            onClick={fetchData}
            disabled={loading}
            title="Refresh now"
          >
            <RefreshCw size={16} className={loading ? 'spinning' : ''} />
          </button>
        </div>
      </div>

      {error && <div className="error-message">⚠️ {error}</div>}

      <div className="stats-grid">
        <div className="stat-card phone">
          <div className="stat-icon">
            <PhoneCall size={32} />
          </div>
          <div className="stat-content">
            <div className="stat-label">Active Calls</div>
            <div className="stat-value">{activeCalls}</div>
          </div>
        </div>

        <div className="stat-card incoming">
          <div className="stat-icon">
            <PhoneIncoming size={32} />
          </div>
          <div className="stat-content">
            <div className="stat-label">Incoming</div>
            <div className="stat-value">{incomingCalls}</div>
          </div>
        </div>

        <div className="stat-card outgoing">
          <div className="stat-icon">
            <PhoneOutgoing size={32} />
          </div>
          <div className="stat-content">
            <div className="stat-label">Outgoing</div>
            <div className="stat-value">{outgoingCalls}</div>
          </div>
        </div>
      </div>

      {/* Call Monitoring Controls - Separate component that doesn't refresh */}
      <CallMonitoring />

      {/* Active Calls List */}
      <div className="calls-section">
        <h3>
          <Phone size={20} /> Active Calls
        </h3>
        {currentCalls.length === 0 ? (
          <div className="empty-state">No active calls</div>
        ) : (
          <div className="calls-list">
            {currentCalls.map((call, index) => {
              const isIncoming = call.direction === 'inbound' || call.direction === 'incoming';
              const displayDuration = call.duration ? 
                (typeof call.duration === 'number' ? 
                  `${Math.floor(call.duration / 60)}:${String(call.duration % 60).padStart(2, '0')}` : 
                  call.duration) : 
                '00:00';
              
              return (
                <div key={call.id || index} className="call-card">
                  <div className="call-header">
                    <div className="call-direction">
                      {isIncoming ? (
                        <PhoneIncoming size={20} className="incoming-icon" />
                      ) : (
                        <PhoneOutgoing size={20} className="outgoing-icon" />
                      )}
                      <span className="call-type">
                        {call.direction || 'Active'}
                      </span>
                    </div>
                    <span className="call-status active">{call.status || 'In Progress'}</span>
                  </div>
                  <div className="call-details">
                    {call.extension && (
                      <div className="detail-row">
                        <Phone size={16} />
                        <span>Extension: <strong>{call.extension}</strong></span>
                      </div>
                    )}
                    {call.accountId && (
                      <div className="detail-row">
                        <Phone size={16} />
                        <span>Account ID: <strong>{call.accountId}</strong></span>
                      </div>
                    )}
                    <div className="detail-row">
                      <Clock size={16} />
                      <span>Duration: {displayDuration}</span>
                    </div>
                    {call.callerName && (
                      <div className="detail-row">
                        <span>Caller: {call.callerName}</span>
                      </div>
                    )}
                    {call.callerNumber && (
                      <div className="detail-row">
                        <span>Number: {call.callerNumber}</span>
                      </div>
                    )}
                    {call.calledNumber && (
                      <div className="detail-row">
                        <span>Called: {call.calledNumber}</span>
                      </div>
                    )}
                  </div>
                  {(call.extension || call.accountId) && (
                    <div className="call-extension-info">
                      <div className="call-id-row">
                        {call.extension && (
                          <div className="call-id-item">
                            <span>Extension: <strong>{call.extension}</strong></span>
                            <button 
                              className="copy-btn"
                              onClick={() => copyToClipboard(call.extension, `ext-${call.id}`)}
                              title="Copy extension"
                            >
                              <Copy size={14} />
                              {copiedValue === `ext-${call.id}` && <span className="copied">✓</span>}
                            </button>
                          </div>
                        )}
                        {call.accountId && (
                          <div className="call-id-item">
                            <span>Account ID: <strong>{call.accountId}</strong></span>
                            <button 
                              className="copy-btn"
                              onClick={() => copyToClipboard(call.accountId, `acc-${call.id}`)}
                              title="Copy account ID"
                            >
                              <Copy size={14} />
                              {copiedValue === `acc-${call.id}` && <span className="copied">✓</span>}
                            </button>
                          </div>
                        )}
                      </div>
                      <div className="monitor-hint">
                        💡 Click copy buttons, then paste into Call Monitoring controls above
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* Queue Status */}
      {queueStatus && (
        <div className="queue-section">
          <h3>📊 Queue Status</h3>
          <div className="queue-info">
            <pre>{JSON.stringify(queueStatus, null, 2)}</pre>
          </div>
        </div>
      )}
    </div>
  );
};

export default DigiumPanel;
