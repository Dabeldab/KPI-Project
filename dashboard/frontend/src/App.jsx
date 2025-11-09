import { useState, useEffect } from 'react'
import RescuePanel from './components/RescuePanel'
import DigiumPanel from './components/DigiumPanel'
import { Activity, RefreshCw } from 'lucide-react'
import { healthCheck } from './api'
import './App.css'

function App() {
  const [health, setHealth] = useState(null);
  const [refreshing, setRefreshing] = useState(false);

  useEffect(() => {
    checkHealth();
  }, []);

  const checkHealth = async () => {
    try {
      const status = await healthCheck();
      setHealth(status);
    } catch (err) {
      console.error('Health check failed:', err);
    }
  };

  const handleRefresh = () => {
    setRefreshing(true);
    window.location.reload();
  };

  return (
    <div className="app">
      <header className="app-header">
        <div className="header-content">
          <div className="header-title">
            <Activity size={32} />
            <h1>DevOps Live Dashboard</h1>
          </div>
          <div className="header-actions">
            {health && (
              <div className="health-status">
                <span className={`status-dot ${health.status === 'ok' ? 'healthy' : 'error'}`}></span>
                <span>System {health.status === 'ok' ? 'Online' : 'Offline'}</span>
              </div>
            )}
            <button 
              className="refresh-btn" 
              onClick={handleRefresh}
              disabled={refreshing}
            >
              <RefreshCw size={20} className={refreshing ? 'spinning' : ''} />
              Refresh
            </button>
          </div>
        </div>
      </header>

      <main className="app-main">
        <div className="dashboard-grid">
          <div className="panel-wrapper">
            <RescuePanel />
          </div>
          <div className="panel-wrapper">
            <DigiumPanel />
          </div>
        </div>
      </main>

      <footer className="app-footer">
        <p>Real-time DevOps Dashboard • Auto-refresh: LogMeIn (30s) | Digium (15s) • Click panel refresh buttons for instant updates</p>
        <p className="timestamp">Dashboard loaded: {new Date().toLocaleString()}</p>
      </footer>
    </div>
  )
}

export default App
