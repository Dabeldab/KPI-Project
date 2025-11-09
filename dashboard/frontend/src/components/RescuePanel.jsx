import React, { useState, useEffect } from 'react';
import { rescueApi } from '../api';
import { Users, UserCheck, Clock, Activity, RefreshCw } from 'lucide-react';
import './RescuePanel.css';

const RescuePanel = () => {
  const [techAvailable, setTechAvailable] = useState(null);
  const [sessions, setSessions] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  const fetchData = async () => {
    try {
      setLoading(true);
      const [techData, sessionsData] = await Promise.all([
        rescueApi.getTechAvailability().catch(() => null),
        rescueApi.getSessions().catch(() => [])
      ]);
      
      setTechAvailable(techData);
      setSessions(Array.isArray(sessionsData) ? sessionsData : []);
      setError(null);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
    const interval = setInterval(fetchData, 30000); // Refresh every 30 seconds
    return () => clearInterval(interval);
  }, []);

  if (loading && !sessions.length) {
    return (
      <div className="rescue-panel">
        <h2>🛟 LogMeIn Rescue</h2>
        <div className="loading">Loading...</div>
      </div>
    );
  }

  return (
    <div className="rescue-panel">
      <div className="panel-header">
        <h2>🛟 LogMeIn Rescue</h2>
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
        <div className="stat-card">
          <div className="stat-icon">
            <UserCheck size={32} />
          </div>
          <div className="stat-content">
            <div className="stat-label">Techs Available</div>
            <div className="stat-value">
              {techAvailable?.available ? '✅ Yes' : '❌ No'}
            </div>
          </div>
        </div>

        <div className="stat-card">
          <div className="stat-icon">
            <Activity size={32} />
          </div>
          <div className="stat-content">
            <div className="stat-label">Active Sessions</div>
            <div className="stat-value">{sessions.length}</div>
          </div>
        </div>
      </div>

      <div className="sessions-section">
        <h3>
          <Users size={20} /> Active Sessions
        </h3>
        {sessions.length === 0 ? (
          <div className="empty-state">No active sessions</div>
        ) : (
          <div className="sessions-list">
            {sessions.map((session, index) => (
              <div key={session.id || index} className="session-card">
                <div className="session-header">
                  <span className="session-id">
                    Session #{session.id || index + 1}
                  </span>
                  <span className={`session-status ${session.status}`}>
                    {session.status || 'Active'}
                  </span>
                </div>
                <div className="session-details">
                  <div className="detail-row">
                    <Clock size={16} />
                    <span>
                      Duration: {session.duration || 'N/A'}
                    </span>
                  </div>
                  {session.technician && (
                    <div className="detail-row">
                      <UserCheck size={16} />
                      <span>Tech: {session.technician}</span>
                    </div>
                  )}
                  {session.customer && (
                    <div className="detail-row">
                      <Users size={16} />
                      <span>Customer: {session.customer}</span>
                    </div>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};

export default RescuePanel;
