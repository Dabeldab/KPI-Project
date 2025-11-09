import express from 'express';
import cors from 'cors';
import axios from 'axios';
import dotenv from 'dotenv';
import { parseString } from 'xml2js';
import { promisify } from 'util';
import rateLimit from 'express-rate-limit';

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3001;
const parseXml = promisify(parseString);

// Rate limiter for authentication endpoints
const authLimiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 5, // Limit each IP to 5 login requests per windowMs
  message: 'Too many login attempts, please try again after 15 minutes',
  standardHeaders: true,
  legacyHeaders: false,
});

// General API rate limiter
const apiLimiter = rateLimit({
  windowMs: 1 * 60 * 1000, // 1 minute
  max: 100, // Limit each IP to 100 requests per minute
  message: 'Too many requests, please try again later',
  standardHeaders: true,
  legacyHeaders: false,
});

// Middleware
app.use(cors());
app.use(express.json());
app.use('/api/', apiLimiter); // Apply general rate limiting to all API routes

// Helper function to create basic auth header
const createAuthHeader = (username, password) => {
  const credentials = Buffer.from(`${username}:${password}`).toString('base64');
  return `Basic ${credentials}`;
};

// ============================================
// LogMeIn Rescue API Authentication
// ============================================

// Store for LogMeIn Rescue session tokens
let rescueSession = {
  token: null,
  cookieJar: null,
  expiresAt: null
};

// Login to LogMeIn Rescue API and get session token
const loginToRescue = async () => {
  try {
    console.log('[Rescue Login] Attempting to login to LogMeIn Rescue API...');
    
    // According to LogMeIn Rescue API documentation, we need to call the login endpoint
    // The login endpoint returns a session cookie/token that we use for subsequent requests
    const response = await axios.post(
      `${process.env.LOGMEIN_API_URL}/login`,
      null,
      {
        params: {
          userName: process.env.LOGMEIN_USERNAME,
          password: process.env.LOGMEIN_PASSWORD
        },
        // Enable cookie handling
        withCredentials: true,
        maxRedirects: 0,
        validateStatus: (status) => status >= 200 && status < 400
      }
    );

    console.log('[Rescue Login] Login response status:', response.status);
    console.log('[Rescue Login] Response headers:', response.headers);
    
    // Extract session token/cookie from response
    const cookies = response.headers['set-cookie'];
    if (cookies) {
      rescueSession.cookieJar = cookies.join('; ');
      rescueSession.token = response.data?.token || 'session-active';
      // Set expiration to 55 minutes (tokens typically last 1 hour)
      rescueSession.expiresAt = Date.now() + (55 * 60 * 1000);
      console.log('[Rescue Login] ✅ Login successful, session token obtained');
      return true;
    }
    
    // Some implementations return token in response body
    if (response.data?.token || response.data?.sessionId) {
      rescueSession.token = response.data.token || response.data.sessionId;
      rescueSession.expiresAt = Date.now() + (55 * 60 * 1000);
      console.log('[Rescue Login] ✅ Login successful, token from response body');
      return true;
    }
    
    console.log('[Rescue Login] ⚠️ Login response did not contain expected token/cookie');
    return false;
  } catch (error) {
    console.error('[Rescue Login] ❌ Login failed:', error.message);
    if (error.response) {
      console.error('[Rescue Login] Status:', error.response.status);
      console.error('[Rescue Login] Response:', error.response.data);
    }
    return false;
  }
};

// Check if session is valid and login if needed
const ensureRescueSession = async () => {
  // Check if we have a valid session
  if (rescueSession.token && rescueSession.expiresAt && Date.now() < rescueSession.expiresAt) {
    console.log('[Rescue Session] Using existing session');
    return true;
  }
  
  console.log('[Rescue Session] No valid session, logging in...');
  return await loginToRescue();
};

// Helper function to make authenticated Rescue API calls
const makeRescueApiCall = async (endpoint, method = 'GET', data = null) => {
  // Ensure we have a valid session
  const sessionValid = await ensureRescueSession();
  if (!sessionValid) {
    throw new Error('Failed to establish Rescue API session');
  }
  
  const config = {
    method,
    url: `${process.env.LOGMEIN_API_URL}${endpoint}`,
    headers: {}
  };
  
  // Add session token/cookie to request
  if (rescueSession.cookieJar) {
    config.headers['Cookie'] = rescueSession.cookieJar;
  }
  if (rescueSession.token && rescueSession.token !== 'session-active') {
    config.headers['Authorization'] = `Bearer ${rescueSession.token}`;
  }
  
  // Add basic auth as fallback
  if (!rescueSession.cookieJar && !config.headers['Authorization']) {
    config.headers['Authorization'] = createAuthHeader(
      process.env.LOGMEIN_USERNAME,
      process.env.LOGMEIN_PASSWORD
    );
  }
  
  if (data) {
    config.data = data;
  }
  
  try {
    const response = await axios(config);
    return response.data;
  } catch (error) {
    // If we get 401, session might have expired - try to re-login once
    if (error.response?.status === 401) {
      console.log('[Rescue API] Session expired, attempting re-login...');
      rescueSession.token = null;
      rescueSession.expiresAt = null;
      
      const sessionValid = await ensureRescueSession();
      if (sessionValid) {
        // Retry the request with new session
        if (rescueSession.cookieJar) {
          config.headers['Cookie'] = rescueSession.cookieJar;
        }
        if (rescueSession.token && rescueSession.token !== 'session-active') {
          config.headers['Authorization'] = `Bearer ${rescueSession.token}`;
        }
        const retryResponse = await axios(config);
        return retryResponse.data;
      }
    }
    throw error;
  }
};

// ============================================
// LogMeIn Rescue API Routes
// ============================================

// Test login endpoint with rate limiting
app.post('/api/rescue/login', authLimiter, async (req, res) => {
  try {
    const success = await loginToRescue();
    if (success) {
      res.json({
        success: true,
        message: 'Login successful',
        expiresAt: rescueSession.expiresAt
      });
    } else {
      res.status(401).json({
        success: false,
        error: 'Login failed - check credentials'
      });
    }
  } catch (error) {
    console.error('Login endpoint error:', error.message);
    res.status(500).json({
      success: false,
      error: 'Login failed',
      details: error.message
    });
  }
});

// Check if any tech is available on channel
app.get('/api/rescue/tech-available', async (req, res) => {
  try {
    const data = await makeRescueApiCall('/isAnyTechAvailableOnChannel');
    res.json(data);
  } catch (error) {
    console.error('LogMeIn Rescue tech-available error:', error.message);
    res.status(error.response?.status || 500).json({
      error: 'Failed to fetch tech availability',
      details: error.message
    });
  }
});

// Get active sessions
app.get('/api/rescue/sessions', async (req, res) => {
  try {
    const data = await makeRescueApiCall('/getSession_v2');
    res.json(data);
  } catch (error) {
    console.error('LogMeIn Rescue sessions error:', error.message);
    res.status(error.response?.status || 500).json({
      error: 'Failed to fetch sessions',
      details: error.message
    });
  }
});

// ============================================
// Digium/Switchvox API Routes
// ============================================

// Helper function to make Digium API calls
const makeDigiumApiCall = async (method, params = {}) => {
  // Check credentials
  if (!process.env.DIGIUM_USERNAME || !process.env.DIGIUM_PASSWORD) {
    throw new Error('Digium credentials not configured. Please set DIGIUM_USERNAME and DIGIUM_PASSWORD in .env file');
  }

  if (process.env.DIGIUM_USERNAME === 'your_username_here' || process.env.DIGIUM_PASSWORD === 'your_password_here') {
    throw new Error('Digium credentials are still set to placeholder values. Please update your .env file with real credentials');
  }

  const xmlPayload = `<?xml version="1.0" encoding="UTF-8"?>
<request>
  <authenticate>
    <username>${process.env.DIGIUM_USERNAME}</username>
    <password>${process.env.DIGIUM_PASSWORD}</password>
  </authenticate>
  <method>${method}</method>
  <parameters>${Object.keys(params).length > 0 ? 
    Object.entries(params).map(([key, value]) => 
      `<${key}>${value}</${key}>`
    ).join('') : ''
  }</parameters>
</request>`;

  console.log(`[Digium API] Calling method: ${method}`);
  console.log(`[Digium API] URL: ${process.env.DIGIUM_API_URL}`);
  console.log(`[Digium API] Username: ${process.env.DIGIUM_USERNAME}`);
  console.log(`[Digium API] Parameters:`, params);

  try {
    const response = await axios.post(
      process.env.DIGIUM_API_URL,
      xmlPayload,
      {
        headers: {
          'Content-Type': 'text/xml'
        }
      }
    );

  console.log(`[Digium API] Success: ${method}`);
    const parsed = await parseXml(response.data);
    
    // Check for API-level errors in the response
    if (parsed.response?.result?.[0] === 'failure') {
      const errorMsg = parsed.response?.error?.[0] || 'Unknown API error';
      throw new Error(`Digium API returned failure: ${errorMsg}`);
    }
    
    return parsed;
    if (error.response) {
      console.error(`[Digium API] HTTP ${error.response.status}:`, error.response.data);
      
      if (error.response.status === 401) {
        throw new Error('Authentication failed. Please check your Digium username and password in the .env file');
      }
      
      throw new Error(`Digium API error (${error.response.status}): ${error.message}`);
    }
  } catch (error) {
    throw new Error(`Digium API error: ${error.message}`);
  }
};

// Get call queue status
app.get('/api/digium/queue-status', async (req, res) => {
  try {
    const result = await makeDigiumApiCall('switchvox.callQueues.getCurrentStatus');
    res.json(result);
  } catch (error) {
    console.error('Digium queue-status error:', error.message);
    res.status(500).json({
      error: 'Failed to fetch queue status',
      details: error.message
    });
  }
});

// Get current calls list
app.get('/api/digium/current-calls', async (req, res) => {
  try {
    const result = await makeDigiumApiCall('switchvox.currentCalls.getList');
    
    console.log('[Current Calls] Raw API Response:', JSON.stringify(result, null, 2));
    
    // Parse the XML response to extract call details
    // The response structure varies, so we'll handle it gracefully
    let calls = [];
    
    if (result && result.response) {
      const response = result.response;
      
      // Check if there are calls in the response
      if (response.calls && response.calls[0] && response.calls[0].call) {
        const callData = response.calls[0].call;
        calls = Array.isArray(callData) ? callData : [callData];
        
        // Map calls to a more usable format
        calls = calls.map(call => ({
          id: call.id ? call.id[0] : null,
          accountId: call.account_id ? call.account_id[0] : null,
          extension: call.extension ? call.extension[0] : null,
          callerNumber: call.caller_number ? call.caller_number[0] : null,
          callerName: call.caller_name ? call.caller_name[0] : null,
          calledNumber: call.called_number ? call.called_number[0] : null,
          direction: call.direction ? call.direction[0] : null,
          status: call.status ? call.status[0] : null,
          duration: call.duration ? call.duration[0] : null,
          startTime: call.start_time ? call.start_time[0] : null
        }));
        
        console.log('[Current Calls] Parsed calls:', JSON.stringify(calls, null, 2));
      }
    }
    
    res.json({ calls, raw: result });
  } catch (error) {
    console.error('Digium current-calls error:', error.message);
    res.status(500).json({
      error: 'Failed to fetch current calls',
      details: error.message
    });
  }
});

// Get call monitoring info
app.get('/api/digium/monitoring-info', async (req, res) => {
  try {
    const result = await makeDigiumApiCall('switchvox.extensions.featureCodes.callMonitoring.getInfo');
    res.json(result);
  } catch (error) {
    console.error('Digium monitoring-info error:', error.message);
    res.status(500).json({
      error: 'Failed to fetch monitoring info',
      details: error.message
    });
  }
});

// Start call monitoring
app.post('/api/digium/start-monitoring', async (req, res) => {
  try {
    const { extension, targetExtension } = req.body;
    
    if (!extension || !targetExtension) {
      return res.status(400).json({
        error: 'Missing required parameters: extension (your account ID) and targetExtension (target account ID)'
      });
    }

    console.log(`[Call Monitoring] Starting: Your Account ID: ${extension}, Target Account ID: ${targetExtension}`);

    const result = await makeDigiumApiCall(
      'switchvox.extensions.featureCodes.callMonitoring.add',
      {
        account_id: extension,
        target_account_id: targetExtension
      }
    );
    
    console.log(`[Call Monitoring] Success:`, JSON.stringify(result, null, 2));
    res.json(result);
  } catch (error) {
    console.error('[Call Monitoring] Error:', error.message);
    console.error('[Call Monitoring] Stack:', error.stack);
    res.status(500).json({
      error: 'Failed to start monitoring',
      details: error.message
    });
  }
});

// Stop call monitoring
app.post('/api/digium/stop-monitoring', async (req, res) => {
  try {
    const { extension } = req.body;
    
    if (!extension) {
      return res.status(400).json({
        error: 'Missing required parameter: extension (account ID)'
      });
    }

    console.log(`[Call Monitoring] Stopping: Account ID: ${extension}`);

    const result = await makeDigiumApiCall(
      'switchvox.extensions.featureCodes.callMonitoring.remove',
      { account_id: extension }
    );
    
    console.log(`[Call Monitoring] Stopped successfully`);
    res.json(result);
  } catch (error) {
    console.error('[Call Monitoring] Stop Error:', error.message);
    res.status(500).json({
      error: 'Failed to stop monitoring',
      details: error.message
    });
  }
});

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    services: {
      logmein: !!process.env.LOGMEIN_USERNAME,
      digium: !!process.env.DIGIUM_USERNAME
    }
  });
});

app.listen(PORT, () => {
  console.log(`🚀 DevOps Dashboard Backend running on port ${PORT}`);
  console.log(`📊 Health check: http://localhost:${PORT}/api/health`);
  
  // Check if credentials are configured
  if (!process.env.LOGMEIN_USERNAME || !process.env.DIGIUM_USERNAME) {
    console.warn('⚠️  Warning: API credentials not configured. Copy .env.example to .env and add your credentials.');
  }
});
