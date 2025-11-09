# 🔐 LogMeIn Rescue API Login Implementation

## Overview

The LogMeIn Rescue API login method has been implemented according to the official API documentation:
https://secure.logmeinrescue.com/welcome/webhelp/EN/RescueAPI/API/API_Rescue_login.html

## What Changed

### Session-Based Authentication

Previously, the application used Basic Authentication for every API call. Now it implements the recommended login flow:

1. **Login Phase**: Call the `/login` endpoint with credentials to obtain a session token
2. **API Calls**: Use the session token for subsequent API calls
3. **Auto-Refresh**: Automatically renew session when it expires (every ~55 minutes)

### Benefits

- ✅ **More Secure**: Session tokens are shorter-lived than persistent credentials
- ✅ **Better Performance**: Login once, reuse session for multiple calls
- ✅ **API Compliant**: Follows LogMeIn Rescue official documentation
- ✅ **Automatic Retry**: Handles expired sessions gracefully

## How It Works

### Backend Implementation

The `server.js` file now includes:

#### 1. Login Function
```javascript
loginToRescue()
```
- Calls `POST /API/login` with username and password as query parameters
- Extracts session token/cookie from response
- Stores session with 55-minute expiration

#### 2. Session Management
```javascript
ensureRescueSession()
```
- Checks if current session is valid
- Automatically logs in if session expired or missing
- Returns true if valid session available

#### 3. API Call Helper
```javascript
makeRescueApiCall(endpoint, method, data)
```
- Ensures valid session before making calls
- Adds session cookie/token to request headers
- Falls back to basic auth if session not available
- Automatically retries with new session on 401 errors

#### 4. API Endpoints

**New Login Endpoint**:
```
POST /api/rescue/login
```
Returns:
```json
{
  "success": true,
  "message": "Login successful",
  "expiresAt": 1699500000000
}
```

**Updated Existing Endpoints**:
- `GET /api/rescue/tech-available` - Now uses session auth
- `GET /api/rescue/sessions` - Now uses session auth

### Frontend Implementation

Updated `api.js` with new login method:

```javascript
import { rescueApi } from './api';

// Manually trigger login (optional - backend handles this automatically)
const result = await rescueApi.login();
console.log(result);

// Use API as before - login happens automatically
const sessions = await rescueApi.getSessions();
```

## Configuration

### Setting Up Credentials

1. **Navigate to backend directory**:
   ```bash
   cd /home/runner/work/KPI-Project/KPI-Project/dashboard/backend
   ```

2. **Edit the .env file**:
   ```bash
   nano .env
   ```

3. **Add your LogMeIn Rescue credentials**:
   ```env
   # Replace with your actual credentials
   LOGMEIN_USERNAME=your_email@company.com
   LOGMEIN_PASSWORD=your_password_or_api_key
   LOGMEIN_API_URL=https://secure.logmeinrescue.com/API
   ```

4. **Save the file** (Ctrl+X, then Y, then Enter)

### Known Credentials (from previous testing)

Based on the troubleshooting documentation, these credentials were used:

```env
LOGMEIN_USERNAME=darius@novapointofsale.com
LOGMEIN_PASSWORD=[need to be provided]

DIGIUM_USERNAME=Darius_Parlor
DIGIUM_PASSWORD=[need to be provided]
```

**⚠️ Security Note**: Never commit the `.env` file with real credentials. It's already in `.gitignore`.

## Testing Authentication

### Option 1: Run the Test Script

```bash
cd /home/runner/work/KPI-Project/KPI-Project/dashboard/backend
npm run test-creds
```

This will:
1. Check if credentials are configured
2. Attempt login via `/login` endpoint
3. Test an API call with the session
4. Fall back to basic auth if needed
5. Test Digium credentials as well

### Option 2: Start the Server and Test Manually

**Terminal 1 - Start Backend**:
```bash
cd /home/runner/work/KPI-Project/KPI-Project/dashboard/backend
npm run dev
```

**Terminal 2 - Test Login**:
```bash
# Test login endpoint
curl -X POST http://localhost:3001/api/rescue/login

# Test tech availability
curl http://localhost:3001/api/rescue/tech-available

# Test sessions
curl http://localhost:3001/api/rescue/sessions
```

### Option 3: Use the Frontend

```bash
# Terminal 1 - Backend (must be running)
cd /home/runner/work/KPI-Project/KPI-Project/dashboard/backend
npm run dev

# Terminal 2 - Frontend
cd /home/runner/work/KPI-Project/KPI-Project/dashboard/frontend
npm run dev
```

Open browser to http://localhost:5173 and the dashboard will automatically:
1. Login to LogMeIn Rescue when making first API call
2. Display tech availability and active sessions
3. Refresh data periodically

## Troubleshooting

### Error: 404 Not Found on /login endpoint

**Possible Causes**:
- The `/login` endpoint might use different URL structure
- API version might be different

**Solutions**:
1. Check LogMeIn Rescue admin panel for exact API endpoint
2. Try alternative formats:
   - `POST /API/login?userName=X&password=Y`
   - `POST /api/v1/login`
   - `POST /API/authenticate`

**Current Implementation**: The code falls back to basic auth if login fails

### Error: 401 Unauthorized

**Causes**:
- Incorrect username or password
- Account doesn't have API access enabled

**Solutions**:
1. Verify credentials in LogMeIn Rescue admin panel
2. Check if API access is enabled for your account
3. Contact LogMeIn support to enable API access

### Error: No session token in response

**Cause**: Login succeeded but token/cookie not in expected format

**Solution**: Check the response structure:
```javascript
console.log('Login response:', response.headers);
console.log('Login data:', response.data);
```

The implementation checks for:
- Cookies in `set-cookie` header
- Token in response body as `response.data.token`
- Session ID in response body as `response.data.sessionId`

### Session Expires Too Quickly

**Current Setting**: 55 minutes

**To Adjust**: Edit `server.js`:
```javascript
// Change this line in loginToRescue()
rescueSession.expiresAt = Date.now() + (55 * 60 * 1000);
// To (for example, 30 minutes):
rescueSession.expiresAt = Date.now() + (30 * 60 * 1000);
```

## API Documentation Reference

Official LogMeIn Rescue API Login Documentation:
https://secure.logmeinrescue.com/welcome/webhelp/EN/RescueAPI/API/API_Rescue_login.html

## Implementation Details

### Login Request Format

According to the LogMeIn Rescue API documentation, the login endpoint expects:

```
POST https://secure.logmeinrescue.com/API/login?userName=USERNAME&password=PASSWORD
```

**Parameters**:
- `userName` (query param): Your LogMeIn Rescue username/email
- `password` (query param): Your password or API key

**Response**:
- Session cookie in `Set-Cookie` header
- Or session token in response body
- Status 200 for success
- Status 401 for invalid credentials

### Session Token Usage

Once logged in, subsequent requests should include:

**Option A**: Cookie-based
```
Cookie: [session cookie from login response]
```

**Option B**: Token-based
```
Authorization: Bearer [token from login response]
```

**Option C**: Basic Auth Fallback
```
Authorization: Basic [base64(username:password)]
```

Our implementation tries all three methods for maximum compatibility.

## Next Steps

1. ✅ **Add your credentials** to `.env` file
2. ✅ **Run the test script**: `npm run test-creds`
3. ✅ **Start the server**: `npm run dev`
4. ✅ **Open the dashboard** in your browser
5. ✅ **Verify** data is loading from LogMeIn Rescue

## Support

If you encounter issues:

1. **Check backend logs**: Look for `[Rescue Login]` and `[Rescue API]` messages
2. **Run test script**: `npm run test-creds` for detailed diagnostics
3. **Verify credentials**: Double-check username/password in `.env`
4. **Check API access**: Ensure your LogMeIn account has API access enabled
5. **Contact support**: Refer to LogMeIn Rescue API documentation

---

**Implementation Date**: November 9, 2025  
**Status**: ✅ Complete - Ready for Testing
