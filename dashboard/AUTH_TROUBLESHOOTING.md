# 🔧 API Authentication Troubleshooting Guide

## Current Status

Based on the credential test, here's what we found:

### LogMeIn Rescue API
- ❌ **Status**: 404 Not Found
- **Issue**: API endpoint not found
- **Current URL**: `https://secure.logmeinrescue.com/API/isAnyTechAvailableOnChannel`

### Digium/Switchvox API  
- ❌ **Status**: 401 Unauthorized
- **Issue**: Invalid credentials or authentication method
- **Current URL**: `https://nova.digiumcloud.net/xml`
- **Username**: `Darius_Parlor`

---

## 🛟 LogMeIn Rescue - Fixing 404 Error

### Problem
The API endpoint returns 404, which usually means:
1. The URL structure is incorrect
2. The API version changed
3. Different API base URL needed

### Solutions to Try

#### Option 1: Check API Documentation
LogMeIn Rescue API endpoints might use a different structure:

**Try these URL formats**:
```bash
# Format 1: With action parameter
https://secure.logmeinrescue.com/API?action=isAnyTechAvailableOnChannel

# Format 2: Different base path
https://secure.logmeinrescue.com/api/v1/isAnyTechAvailableOnChannel

# Format 3: REST style
https://secure.logmeinrescue.com/API/technicians/available
```

#### Option 2: Test with curl

```bash
# Test with your credentials
curl -X GET "https://secure.logmeinrescue.com/API/isAnyTechAvailableOnChannel" \
  -u "darius@novapointofsale.com:YOUR_PASSWORD" \
  -v
```

Look for:
- Redirect (301/302) to different URL
- Error message with correct endpoint
- API key requirement instead of basic auth

#### Option 3: Check for API Key
Some LogMeIn APIs use API keys instead of username/password:

```env
# Instead of:
LOGMEIN_USERNAME=darius@novapointofsale.com
LOGMEIN_PASSWORD=your_password

# Try:
LOGMEIN_API_KEY=your_api_key_here
```

#### Option 4: Contact LogMeIn Support
- Get exact API endpoint URLs
- Verify your account has API access
- Check if API version changed

---

## 📞 Digium/Switchvox - Fixing 401 Unauthorized

### Problem
401 Unauthorized means authentication is failing. Common causes:

1. **Username format incorrect**
2. **Password incorrect**
3. **Account doesn't have API access**
4. **Different authentication method required**

### Solutions to Try

#### Option 1: Verify Credentials in Switchvox Admin Panel

1. Log into Switchvox web interface
2. Go to **Settings** → **Users & Extensions**
3. Find user **`Darius_Parlor`**
4. Check:
   - ✅ User account is active
   - ✅ API access is enabled
   - ✅ Password is correct
   - ✅ User has admin/API permissions

#### Option 2: Try Different Username Formats

Switchvox might need username in different format:

```env
# Current (not working):
DIGIUM_USERNAME=Darius_Parlor

# Try these alternatives:
DIGIUM_USERNAME=darius_parlor          # lowercase
DIGIUM_USERNAME=Darius_Parlor@nova     # with domain
DIGIUM_USERNAME=darius@novapointofsale.com  # email format
DIGIUM_USERNAME=admin                   # admin account
```

#### Option 3: Check API Authentication Method

Some Switchvox versions use different auth:

**Test with curl:**
```bash
# Test current credentials
curl -X POST "https://nova.digiumcloud.net/xml" \
  -H "Content-Type: text/xml" \
  -d '<?xml version="1.0" encoding="UTF-8"?>
<request>
  <authenticate>
    <username>Darius_Parlor</username>
    <password>YOUR_PASSWORD</password>
  </authenticate>
  <method>switchvox.extensions.getInfo</method>
</request>' \
  -v
```

#### Option 4: Use Session-Based Authentication

Some Switchvox versions require session tokens:

1. First, get a session token
2. Use token for subsequent requests

```xml
<!-- Step 1: Login -->
<request>
  <method>switchvox.users.login</method>
  <parameters>
    <username>Darius_Parlor</username>
    <password>YOUR_PASSWORD</password>
  </parameters>
</request>
```

#### Option 5: Check Account Permissions

User might need specific permissions:
- API Access permission
- Extension management
- Call monitoring permission
- Admin role

Contact your Switchvox administrator to verify.

---

## 🧪 Testing Commands

### Test LogMeIn Rescue (run from backend folder)
```bash
cd /workspaces/KPI-Project/dashboard/backend
node test-credentials.js
```

### Test Digium with curl
```bash
# Replace YOUR_PASSWORD with actual password
curl -X POST "https://nova.digiumcloud.net/xml" \
  -H "Content-Type: text/xml" \
  -d '<?xml version="1.0" encoding="UTF-8"?>
<request>
  <authenticate>
    <username>Darius_Parlor</username>
    <password>YOUR_PASSWORD</password>
  </authenticate>
  <method>switchvox.ping</method>
</request>'
```

### Check if Digium API is accessible
```bash
curl -v https://nova.digiumcloud.net/xml
```

---

## 📝 Next Steps

### For LogMeIn Rescue:
1. ✅ Verify API documentation URL structure
2. ✅ Test with curl to see actual error
3. ✅ Check if API key authentication is required
4. ✅ Contact LogMeIn support if needed

### For Digium/Switchvox:
1. ✅ Log into Switchvox admin panel
2. ✅ Verify user account and permissions
3. ✅ Try different username formats
4. ✅ Test with curl to isolate issue
5. ✅ Contact Switchvox administrator

---

## 🔍 Understanding Error Codes

| Code | Meaning | Common Cause |
|------|---------|--------------|
| 401 | Unauthorized | Wrong username/password |
| 403 | Forbidden | No API access permission |
| 404 | Not Found | Wrong URL or endpoint doesn't exist |
| 500 | Server Error | API server issue |

---

## 💡 Quick Fixes

### If Digium works but LogMeIn doesn't:
- Comment out LogMeIn code temporarily
- Dashboard will still show Digium data

### If LogMeIn works but Digium doesn't:
- Comment out Digium code temporarily  
- Dashboard will still show LogMeIn data

### To disable either API:
Edit `backend/server.js` and comment out the routes you don't need.

---

## 📞 Getting Help

1. **LogMeIn Support**: https://secure.logmeinrescue.com/welcome/webhelp/
2. **Digium Support**: http://developers.digium.com/switchvox/wiki/
3. **Your IT Admin**: They may have specific credentials/permissions needed

---

## ✅ When Authentication Works

You'll see:
```
✅ LogMeIn Rescue: Authentication successful!
   Response: {...}

✅ Digium/Switchvox: Authentication successful!
   Response: success
```

Then restart your dashboard and it should work!

---

**Run the credential tester after any changes:**
```bash
cd /workspaces/KPI-Project/dashboard/backend
node test-credentials.js
```
