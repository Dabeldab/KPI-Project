import axios from 'axios';

const API_BASE_URL = '/api';

// LogMeIn Rescue API calls
export const rescueApi = {
  getTechAvailability: async () => {
    const response = await axios.get(`${API_BASE_URL}/rescue/tech-available`);
    return response.data;
  },
  
  getSessions: async () => {
    const response = await axios.get(`${API_BASE_URL}/rescue/sessions`);
    return response.data;
  }
};

// Digium/Switchvox API calls
export const digiumApi = {
  getQueueStatus: async () => {
    const response = await axios.get(`${API_BASE_URL}/digium/queue-status`);
    return response.data;
  },
  
  getCurrentCalls: async () => {
    const response = await axios.get(`${API_BASE_URL}/digium/current-calls`);
    return response.data;
  },
  
  getMonitoringInfo: async () => {
    const response = await axios.get(`${API_BASE_URL}/digium/monitoring-info`);
    return response.data;
  },
  
  startMonitoring: async (extension, targetExtension) => {
    const response = await axios.post(`${API_BASE_URL}/digium/start-monitoring`, {
      extension,
      targetExtension
    });
    return response.data;
  },
  
  stopMonitoring: async (extension) => {
    const response = await axios.post(`${API_BASE_URL}/digium/stop-monitoring`, {
      extension
    });
    return response.data;
  }
};

// Health check
export const healthCheck = async () => {
  const response = await axios.get(`${API_BASE_URL}/health`);
  return response.data;
};
