#!/usr/bin/env node

/**
 * API Credentials Tester
 * Tests your LogMeIn Rescue and Digium/Switchvox credentials
 */

import dotenv from 'dotenv';
import axios from 'axios';
import { parseString } from 'xml2js';
import { promisify } from 'util';

dotenv.config();
const parseXml = promisify(parseString);

console.log('🔐 API Credentials Tester\n');
console.log('========================================\n');

// Check environment variables
console.log('📋 Checking environment variables...\n');

const checks = {
  logmein_username: !!process.env.LOGMEIN_USERNAME && process.env.LOGMEIN_USERNAME !== 'your_username_here',
  logmein_password: !!process.env.LOGMEIN_PASSWORD && process.env.LOGMEIN_PASSWORD !== 'your_password_here',
  digium_username: !!process.env.DIGIUM_USERNAME && process.env.DIGIUM_USERNAME !== 'your_username_here',
  digium_password: !!process.env.DIGIUM_PASSWORD && process.env.DIGIUM_PASSWORD !== 'your_password_here',
  logmein_url: !!process.env.LOGMEIN_API_URL,
  digium_url: !!process.env.DIGIUM_API_URL
};

console.log(`LOGMEIN_USERNAME: ${checks.logmein_username ? '✅ Set' : '❌ Not set or placeholder'} ${process.env.LOGMEIN_USERNAME ? `(${process.env.LOGMEIN_USERNAME})` : ''}`);
console.log(`LOGMEIN_PASSWORD: ${checks.logmein_password ? '✅ Set' : '❌ Not set or placeholder'} ${process.env.LOGMEIN_PASSWORD ? '(****)' : ''}`);
console.log(`LOGMEIN_API_URL: ${checks.logmein_url ? '✅ Set' : '❌ Not set'} ${process.env.LOGMEIN_API_URL ? `(${process.env.LOGMEIN_API_URL})` : ''}`);
console.log(`DIGIUM_USERNAME: ${checks.digium_username ? '✅ Set' : '❌ Not set or placeholder'} ${process.env.DIGIUM_USERNAME ? `(${process.env.DIGIUM_USERNAME})` : ''}`);
console.log(`DIGIUM_PASSWORD: ${checks.digium_password ? '✅ Set' : '❌ Not set or placeholder'} ${process.env.DIGIUM_PASSWORD ? '(****)' : ''}`);
console.log(`DIGIUM_API_URL: ${checks.digium_url ? '✅ Set' : '❌ Not set'} ${process.env.DIGIUM_API_URL ? `(${process.env.DIGIUM_API_URL})` : ''}`);

console.log('\n========================================\n');

// Test LogMeIn Rescue
if (checks.logmein_username && checks.logmein_password && checks.logmein_url) {
  console.log('🛟 Testing LogMeIn Rescue API...\n');
  
  // Test 1: Try the login endpoint
  console.log('   Test 1: Attempting login via /login endpoint...');
  try {
    const loginResponse = await axios.post(
      `${process.env.LOGMEIN_API_URL}/login`,
      null,
      {
        params: {
          userName: process.env.LOGMEIN_USERNAME,
          password: process.env.LOGMEIN_PASSWORD
        },
        maxRedirects: 0,
        validateStatus: (status) => status >= 200 && status < 400,
        timeout: 10000
      }
    );
    
    console.log('   ✅ Login endpoint successful!');
    console.log(`   Status: ${loginResponse.status}`);
    
    // Check for session cookie or token
    const cookies = loginResponse.headers['set-cookie'];
    const token = loginResponse.data?.token || loginResponse.data?.sessionId;
    
    if (cookies || token) {
      console.log('   ✅ Session token/cookie obtained');
      if (token) {
        console.log(`   Token: ${token.substring(0, 20)}...`);
      }
      
      // Test 2: Try using the session for an API call
      console.log('\n   Test 2: Making authenticated API call...');
      const headers = {};
      if (cookies) {
        headers['Cookie'] = cookies.join('; ');
      }
      if (token) {
        headers['Authorization'] = `Bearer ${token}`;
      }
      
      try {
        const apiResponse = await axios.get(
          `${process.env.LOGMEIN_API_URL}/isAnyTechAvailableOnChannel`,
          { headers, timeout: 10000 }
        );
        console.log('   ✅ API call with session successful!');
        console.log(`   Response: ${JSON.stringify(apiResponse.data)}\n`);
      } catch (apiError) {
        console.log('   ⚠️  API call with session failed, but login worked');
        console.log(`   Error: ${apiError.message}\n`);
      }
    } else {
      console.log('   ⚠️  Login successful but no token/cookie found in response\n');
    }
  } catch (loginError) {
    // If login endpoint fails, try basic auth as fallback
    console.log(`   ⚠️  Login endpoint failed: ${loginError.message}`);
    console.log('   Trying basic authentication as fallback...\n');
    
    try {
      const credentials = Buffer.from(`${process.env.LOGMEIN_USERNAME}:${process.env.LOGMEIN_PASSWORD}`).toString('base64');
      const response = await axios.get(
        `${process.env.LOGMEIN_API_URL}/isAnyTechAvailableOnChannel`,
        {
          headers: {
            'Authorization': `Basic ${credentials}`
          },
          timeout: 10000
        }
      );
      
      console.log('✅ LogMeIn Rescue: Basic authentication successful!');
      console.log(`   Response: ${JSON.stringify(response.data)}\n`);
    } catch (error) {
      console.log('❌ LogMeIn Rescue: All authentication methods failed');
      if (error.response) {
        console.log(`   Status: ${error.response.status}`);
        console.log(`   Error: ${error.response.statusText}`);
        if (error.response.status === 401) {
          console.log('   ⚠️  Invalid username or password');
        } else if (error.response.status === 404) {
          console.log('   ⚠️  API endpoint not found - check URL');
        }
      } else {
        console.log(`   Error: ${error.message}`);
      }
      console.log('');
    }
  }
} else {
  console.log('⏭️  Skipping LogMeIn Rescue test (credentials not configured)\n');
}

// Test Digium/Switchvox
if (checks.digium_username && checks.digium_password && checks.digium_url) {
  console.log('📞 Testing Digium/Switchvox API...\n');
  
  const xmlPayload = `<?xml version="1.0" encoding="UTF-8"?>
<request>
  <authenticate>
    <username>${process.env.DIGIUM_USERNAME}</username>
    <password>${process.env.DIGIUM_PASSWORD}</password>
  </authenticate>
  <method>switchvox.extensions.getInfo</method>
  <parameters>
    <account_id>${process.env.DIGIUM_USERNAME}</account_id>
  </parameters>
</request>`;

  try {
    const response = await axios.post(
      process.env.DIGIUM_API_URL,
      xmlPayload,
      {
        headers: {
          'Content-Type': 'text/xml'
        },
        timeout: 10000
      }
    );
    
    const parsed = await parseXml(response.data);
    
    if (parsed.response?.result?.[0] === 'success') {
      console.log('✅ Digium/Switchvox: Authentication successful!');
      console.log(`   Response: ${JSON.stringify(parsed.response.result[0])}\n`);
    } else if (parsed.response?.result?.[0] === 'failure') {
      console.log('❌ Digium/Switchvox: API returned failure');
      console.log(`   Error: ${parsed.response?.error?.[0] || 'Unknown error'}\n`);
    } else {
      console.log('⚠️  Digium/Switchvox: Unexpected response format');
      console.log(`   Response: ${JSON.stringify(parsed, null, 2)}\n`);
    }
  } catch (error) {
    console.log('❌ Digium/Switchvox: Authentication failed');
    if (error.response) {
      console.log(`   Status: ${error.response.status}`);
      console.log(`   Error: ${error.response.statusText}`);
      if (error.response.status === 401) {
        console.log('   ⚠️  Invalid username or password');
      }
    } else {
      console.log(`   Error: ${error.message}`);
    }
    console.log('');
  }
} else {
  console.log('⏭️  Skipping Digium/Switchvox test (credentials not configured)\n');
}

console.log('========================================\n');
console.log('📝 Summary:\n');

const allConfigured = Object.values(checks).every(v => v);
if (allConfigured) {
  console.log('✅ All credentials are configured');
  console.log('   If tests failed, check that credentials are correct\n');
} else {
  console.log('❌ Some credentials are missing or still have placeholder values');
  console.log('   Please edit backend/.env file with your actual credentials\n');
  console.log('Steps to fix:');
  console.log('1. cd /workspaces/KPI-Project/dashboard/backend');
  console.log('2. nano .env');
  console.log('3. Replace placeholder values with real credentials');
  console.log('4. Save and run this test again\n');
}
