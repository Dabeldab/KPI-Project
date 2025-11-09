# 🚀 Quick Test Guide - LogMeIn Rescue API Login

## What Was Implemented

✅ **LogMeIn Rescue API Login Method** - Session-based authentication as per official API documentation  
✅ **Automatic Session Management** - Backend handles login, token storage, and auto-refresh  
✅ **Enhanced Test Script** - Tests both login endpoint and basic auth fallback  
✅ **Comprehensive Documentation** - See RESCUE_LOGIN_IMPLEMENTATION.md for details  

## Quick Start - Test Authentication Now

### Option 1: Interactive Setup Script (Easiest)

```bash
cd /home/runner/work/KPI-Project/KPI-Project/dashboard/backend
./setup-and-test.sh
```

This script will:
1. Check if credentials are configured
2. Let you enter credentials interactively
3. Save them to .env file
4. Run authentication tests automatically

### Option 2: Manual Configuration

```bash
cd /home/runner/work/KPI-Project/KPI-Project/dashboard/backend

# Edit the .env file
nano .env

# Add your credentials:
# LOGMEIN_USERNAME=your_email@company.com
# LOGMEIN_PASSWORD=your_password
# DIGIUM_USERNAME=your_username
# DIGIUM_PASSWORD=your_password

# Save and exit (Ctrl+X, Y, Enter)

# Run tests
npm run test-creds
```

### Option 3: Quick Test with Environment Variables

```bash
cd /home/runner/work/KPI-Project/KPI-Project/dashboard/backend

LOGMEIN_USERNAME="your_email@company.com" \
LOGMEIN_PASSWORD="your_password" \
DIGIUM_USERNAME="your_username" \
DIGIUM_PASSWORD="your_password" \
npm run test-creds
```

## What to Expect

### ✅ Successful Test Output

```
🛟 Testing LogMeIn Rescue API...

   Test 1: Attempting login via /login endpoint...
   ✅ Login endpoint successful!
   Status: 200
   ✅ Session token/cookie obtained
   Token: abc123def456...

   Test 2: Making authenticated API call...
   ✅ API call with session successful!
   Response: {...}

📞 Testing Digium/Switchvox API...

✅ Digium/Switchvox: Authentication successful!
   Response: success

📝 Summary:

✅ All credentials are configured
```

### ❌ Failed Test Output

```
❌ LogMeIn Rescue: All authentication methods failed
   Status: 401
   Error: Unauthorized
   ⚠️  Invalid username or password
```

## Understanding the Implementation

### How Session-Based Auth Works

1. **First Call**: Backend calls `POST /API/login?userName=X&password=Y`
2. **Session Token**: Stores the session cookie/token (expires in 55 minutes)
3. **Subsequent Calls**: Uses the session token
4. **Auto-Refresh**: If token expires (401), automatically re-authenticates
5. **Fallback**: If login fails, uses basic authentication

### Endpoints That Use Session Auth

- `GET /api/rescue/tech-available` - Check technician availability
- `GET /api/rescue/sessions` - Get active support sessions
- `POST /api/rescue/login` - Manual login test (optional)

### Backend Logs to Watch

When running the server (`npm run dev`), you'll see:

```
[Rescue Login] Attempting to login to LogMeIn Rescue API...
[Rescue Login] Login response status: 200
[Rescue Login] ✅ Login successful, session token obtained
[Rescue Session] Using existing session
```

## Troubleshooting

### Issue: 404 Not Found on /login

**Cause**: Login endpoint URL might be different

**Check**: The implementation tries `/API/login` but also falls back to basic auth if it fails

**Solution**: This is already handled - basic auth will be used as fallback

### Issue: 401 Unauthorized

**Causes**:
- Wrong username or password
- Account doesn't have API access

**Solutions**:
1. Verify credentials in LogMeIn Rescue admin panel
2. Check if API access is enabled for your account
3. Contact LogMeIn support

### Issue: Connection Timeout

**Cause**: Network issue or wrong API URL

**Check**: 
```bash
curl -I https://secure.logmeinrescue.com/API
```

**Current API URL**: `https://secure.logmeinrescue.com/API`

## Testing the Full Dashboard

Once authentication tests pass:

```bash
# Terminal 1 - Backend
cd /home/runner/work/KPI-Project/KPI-Project/dashboard/backend
npm run dev

# Terminal 2 - Frontend
cd /home/runner/work/KPI-Project/KPI-Project/dashboard/frontend
npm install  # if not already done
npm run dev

# Open browser to http://localhost:5173
```

The dashboard will:
- Automatically login to LogMeIn Rescue on first API call
- Display technician availability
- Show active support sessions
- Refresh data every 10 seconds
- Handle session expiration automatically

## Manual API Testing

With the backend running:

```bash
# Test login
curl -X POST http://localhost:3001/api/rescue/login

# Test tech availability
curl http://localhost:3001/api/rescue/tech-available

# Test sessions
curl http://localhost:3001/api/rescue/sessions

# Health check
curl http://localhost:3001/api/health
```

## Known Credentials (from previous testing)

Based on the troubleshooting documentation:

```env
LOGMEIN_USERNAME=darius@novapointofsale.com
LOGMEIN_PASSWORD=[needs to be provided by you]

DIGIUM_USERNAME=Darius_Parlor
DIGIUM_PASSWORD=[needs to be provided by you]
```

If these are your credentials, update the .env file with the passwords.

## Files Modified

1. **server.js** - Implements session-based authentication
2. **test-credentials.js** - Enhanced test script with login endpoint testing
3. **api.js** (frontend) - Added login method
4. **.gitignore** - Ensures .env files are never committed

## Documentation Added

1. **RESCUE_LOGIN_IMPLEMENTATION.md** - Comprehensive implementation guide
2. **API_CONFIG.md** - Updated with new authentication method
3. **QUICK_TEST_GUIDE.md** - This file
4. **setup-and-test.sh** - Interactive setup script

## Need Help?

1. **Run interactive setup**: `./setup-and-test.sh`
2. **Read full documentation**: `cat RESCUE_LOGIN_IMPLEMENTATION.md`
3. **Check API config**: `cat API_CONFIG.md`
4. **View backend logs**: `npm run dev` (shows detailed auth logs)
5. **Test credentials**: `npm run test-creds`

## Summary

✅ **Implementation Complete**: LogMeIn Rescue API login method is implemented  
✅ **Ready for Testing**: Just add your credentials and run tests  
✅ **Automatic Handling**: Backend manages sessions automatically  
✅ **Multiple Fallbacks**: Login endpoint → Basic auth → Clear error messages  
✅ **Well Documented**: Multiple guides and documentation files created  

**Next Step**: Add your credentials to `.env` file and run `npm run test-creds` or `./setup-and-test.sh`
