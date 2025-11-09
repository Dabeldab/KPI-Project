import React, { useState } from 'react';
import { digiumApi } from '../api';
import { Eye, EyeOff, Phone, AlertCircle } from 'lucide-react';
import './CallMonitoring.css';

// Extension to Account ID mapping
const EXTENSION_TO_ACCOUNT = {
  '252': '1381',
  '304': '1442',
  '305': '1430',
  '306': '1436',
  '308': '1421',
  '322': '1423',
  '355': '1439',
  '356': '1351'
};

const CallMonitoring = () => {
  const [monitoringExtension, setMonitoringExtension] = useState('');
  const [targetExtension, setTargetExtension] = useState('');
  const [isMonitoring, setIsMonitoring] = useState(false);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [debugLog, setDebugLog] = useState([]);

  const addLog = (message, type = 'info') => {
    const timestamp = new Date().toLocaleTimeString();
    const logEntry = { timestamp, message, type };
    setDebugLog(prev => [logEntry, ...prev].slice(0, 10)); // Keep last 10 logs
    console.log(`[${timestamp}] ${type.toUpperCase()}: ${message}`);
  };

  const handleStartMonitoring = async (callExtension = null) => {
    const target = callExtension || targetExtension;
    if (!monitoringExtension || !target) {
      const msg = '⚠️ Please enter both your extension and the target extension';
      setMessage(msg);
      addLog('Missing extension or target', 'warning');
      setTimeout(() => setMessage(''), 3000);
      return;
    }

    // Convert extensions to account IDs
    const yourAccountId = EXTENSION_TO_ACCOUNT[monitoringExtension];
    const targetAccountId = EXTENSION_TO_ACCOUNT[target];

    addLog(`Looking up extensions - Your: ${monitoringExtension}, Target: ${target}`, 'info');
    
    if (!yourAccountId) {
      const msg = `❌ Extension ${monitoringExtension} not found in mapping. Available extensions: ${Object.keys(EXTENSION_TO_ACCOUNT).join(', ')}`;
      setMessage(msg);
      addLog(msg, 'error');
      setTimeout(() => setMessage(''), 5000);
      return;
    }

    if (!targetAccountId) {
      const msg = `❌ Target extension ${target} not found in mapping. Available extensions: ${Object.keys(EXTENSION_TO_ACCOUNT).join(', ')}`;
      setMessage(msg);
      addLog(msg, 'error');
      setTimeout(() => setMessage(''), 5000);
      return;
    }

    addLog(`Mapped extensions - Your Account: ${yourAccountId}, Target Account: ${targetAccountId}`, 'success');

    setLoading(true);
    try {
      addLog(`Calling API to start monitoring...`, 'info');
      await digiumApi.startMonitoring(yourAccountId, targetAccountId);
      const successMsg = `✅ Monitoring started! Your ext ${monitoringExtension} (Account ${yourAccountId}) is monitoring ext ${target} (Account ${targetAccountId})`;
      setMessage(successMsg);
      addLog(successMsg, 'success');
      setIsMonitoring(true);
      setTimeout(() => setMessage(''), 5000);
    } catch (err) {
      const errorMsg = `❌ Failed to start monitoring: ${err.message}`;
      setMessage(errorMsg);
      addLog(`API Error: ${err.response?.data ? JSON.stringify(err.response.data) : err.message}`, 'error');
      setTimeout(() => setMessage(''), 5000);
    } finally {
      setLoading(false);
    }
  };

  const handleStopMonitoring = async () => {
    if (!monitoringExtension) {
      const msg = '⚠️ Please enter your extension';
      setMessage(msg);
      addLog('Missing extension for stop', 'warning');
      setTimeout(() => setMessage(''), 3000);
      return;
    }

    const yourAccountId = EXTENSION_TO_ACCOUNT[monitoringExtension];
    
    if (!yourAccountId) {
      const msg = `❌ Extension ${monitoringExtension} not found in mapping`;
      setMessage(msg);
      addLog(msg, 'error');
      setTimeout(() => setMessage(''), 5000);
      return;
    }

    addLog(`Stopping monitoring for extension ${monitoringExtension} (Account ${yourAccountId})`, 'info');

    setLoading(true);
    try {
      await digiumApi.stopMonitoring(yourAccountId);
      const successMsg = `✅ Monitoring stopped for extension ${monitoringExtension}`;
      setMessage(successMsg);
      addLog(successMsg, 'success');
      setIsMonitoring(false);
      setTimeout(() => setMessage(''), 3000);
    } catch (err) {
      const errorMsg = `❌ Failed to stop monitoring: ${err.message}`;
      setMessage(errorMsg);
      addLog(`API Error: ${err.response?.data ? JSON.stringify(err.response.data) : err.message}`, 'error');
      setTimeout(() => setMessage(''), 5000);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="call-monitoring">
      <h3>
        <Eye size={20} /> Call Monitoring Controls
      </h3>
      
      {message && (
        <div className={`message ${message.includes('✅') ? 'success' : message.includes('❌') ? 'error' : 'warning'}`}>
          {message}
        </div>
      )}

      <div className="monitoring-form">
        <div className="input-row">
          <div className="input-group">
            <label htmlFor="your-ext">Your Extension:</label>
            <input
              id="your-ext"
              type="text"
              value={monitoringExtension}
              onChange={(e) => setMonitoringExtension(e.target.value)}
              placeholder="e.g., 1001"
              disabled={loading}
            />
          </div>
          <div className="input-group">
            <label htmlFor="target-ext">Target Extension:</label>
            <input
              id="target-ext"
              type="text"
              value={targetExtension}
              onChange={(e) => setTargetExtension(e.target.value)}
              placeholder="e.g., 1002"
              disabled={loading}
            />
          </div>
        </div>

        <div className="button-group">
          <button 
            className={`btn btn-primary ${isMonitoring ? 'active' : ''}`}
            onClick={() => handleStartMonitoring()}
            disabled={loading || !monitoringExtension || !targetExtension}
          >
            <Eye size={16} /> 
            {loading ? 'Starting...' : 'Start Monitoring'}
          </button>
          <button 
            className="btn btn-secondary" 
            onClick={handleStopMonitoring}
            disabled={loading || !monitoringExtension}
          >
            <EyeOff size={16} /> 
            {loading ? 'Stopping...' : 'Stop Monitoring'}
          </button>
        </div>

        {isMonitoring && (
          <div className="monitoring-active">
            <Phone size={16} />
            <span>Currently monitoring extension {targetExtension} (Account {EXTENSION_TO_ACCOUNT[targetExtension]})</span>
          </div>
        )}

        <div className="extension-mapping">
          <details>
            <summary>
              <AlertCircle size={14} />
              Available Extensions (click to expand)
            </summary>
            <div className="mapping-list">
              {Object.entries(EXTENSION_TO_ACCOUNT).map(([ext, accountId]) => (
                <div key={ext} className="mapping-item">
                  <span className="ext">Ext {ext}</span>
                  <span className="arrow">→</span>
                  <span className="account">Account {accountId}</span>
                </div>
              ))}
            </div>
          </details>
        </div>
      </div>

      {/* Debug Log */}
      {debugLog.length > 0 && (
        <div className="debug-log">
          <details>
            <summary>
              <AlertCircle size={14} />
              Debug Log ({debugLog.length} entries)
            </summary>
            <div className="log-entries">
              {debugLog.map((log, index) => (
                <div key={index} className={`log-entry ${log.type}`}>
                  <span className="log-time">{log.timestamp}</span>
                  <span className="log-message">{log.message}</span>
                </div>
              ))}
            </div>
          </details>
        </div>
      )}
    </div>
  );
};

export default CallMonitoring;
