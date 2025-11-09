# 🔧 API Configuration Guide

## LogMeIn Rescue API Setup

### Getting Your Credentials

1. Log into your LogMeIn Rescue account
2. Navigate to **Administration** → **API Access**
3. Generate or retrieve your API credentials
4. You'll need:
   - Username (your LogMeIn Rescue username)
   - Password (your LogMeIn Rescue password or API key)

### API Base URL
```
https://secure.logmeinrescue.com/API
```

### Endpoints Used

#### 1. isAnyTechAvailableOnChannel
**Purpose**: Check if any technicians are available

**Request**:
```
GET /API/isAnyTechAvailableOnChannel
Authorization: Basic <base64(username:password)>
```

**Response** (expected):
```json
{
  "available": true,
  "channel": "default",
  "techCount": 5
}
```

#### 2. getSession_v2
**Purpose**: Get list of active support sessions

**Request**:
```
GET /API/getSession_v2
Authorization: Basic <base64(username:password)>
```

**Response** (expected):
```json
[
  {
    "id": "12345",
    "status": "Active",
    "technician": "John Doe",
    "customer": "Jane Smith",
    "duration": "00:15:30"
  }
]
```

### Testing Your LogMeIn Credentials

```bash
# Replace with your actual credentials
curl -u "your_username:your_password" \
  https://secure.logmeinrescue.com/API/isAnyTechAvailableOnChannel
```

---

## Digium/Switchvox API Setup

### Getting Your Credentials

1. Log into your Switchvox admin panel
2. Navigate to **Settings** → **API Access**
3. Create or retrieve API credentials
4. You'll need:
   - Username (admin or API user)
   - Password (corresponding password)

### API Base URL
```
https://nova.digiumcloud.net/xml
```

### XML Request Format

Digium uses XML-RPC format:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<request>
  <authenticate>
    <username>your_username</username>
    <password>your_password</password>
  </authenticate>
  <method>method_name_here</method>
  <parameters>
    <!-- method-specific parameters -->
  </parameters>
</request>
```

### Endpoints Used

#### 1. switchvox.callQueues.getCurrentStatus
**Purpose**: Get current status of all call queues

**Method**: `switchvox.callQueues.getCurrentStatus`

**Parameters**: None

**Response** (expected XML):
```xml
<response>
  <queue>
    <name>Support Queue</name>
    <calls_waiting>3</calls_waiting>
    <agents_available>5</agents_available>
  </queue>
</response>
```

#### 2. switchvox.currentCalls.getList
**Purpose**: Get list of all current active calls

**Method**: `switchvox.currentCalls.getList`

**Parameters**: None

**Response** (expected XML):
```xml
<response>
  <call>
    <id>12345</id>
    <extension>1001</extension>
    <direction>incoming</direction>
    <caller>555-1234</caller>
    <duration>120</duration>
  </call>
</response>
```

#### 3. switchvox.extensions.featureCodes.callMonitoring.getInfo
**Purpose**: Get information about call monitoring feature codes

**Method**: `switchvox.extensions.featureCodes.callMonitoring.getInfo`

**Parameters**: None

#### 4. switchvox.extensions.featureCodes.callMonitoring.add
**Purpose**: Start monitoring a call

**Method**: `switchvox.extensions.featureCodes.callMonitoring.add`

**Parameters**:
```xml
<parameters>
  <extension>1001</extension>
  <target_extension>1002</target_extension>
</parameters>
```

#### 5. switchvox.extensions.featureCodes.callMonitoring.remove
**Purpose**: Stop monitoring a call

**Method**: `switchvox.extensions.featureCodes.callMonitoring.remove`

**Parameters**:
```xml
<parameters>
  <extension>1001</extension>
</parameters>
```

### Testing Your Digium Credentials

```bash
curl -X POST https://nova.digiumcloud.net/xml \
  -H "Content-Type: text/xml" \
  -d '<?xml version="1.0" encoding="UTF-8"?>
<request>
  <authenticate>
    <username>your_username</username>
    <password>your_password</password>
  </authenticate>
  <method>switchvox.currentCalls.getList</method>
  <parameters></parameters>
</request>'
```

---

## Environment Variables Configuration

### Complete `.env` Template

```env
# LogMeIn Rescue API Credentials
# Get these from: Administration → API Access
LOGMEIN_USERNAME=your_rescue_username
LOGMEIN_PASSWORD=your_rescue_password_or_api_key
LOGMEIN_API_URL=https://secure.logmeinrescue.com/API

# Digium/Switchvox API Credentials  
# Get these from: Settings → API Access
DIGIUM_USERNAME=your_switchvox_admin_username
DIGIUM_PASSWORD=your_switchvox_admin_password
DIGIUM_API_URL=https://nova.digiumcloud.net/xml

# Server Configuration
PORT=3001
```

### Setting Up Your `.env` File

```bash
cd /workspaces/KPI-Project/dashboard/backend

# Create .env from template
cp .env.example .env

# Edit with your credentials
nano .env
```

---

## Troubleshooting API Issues

### LogMeIn Rescue

**401 Unauthorized**
- ❌ Incorrect username or password
- ✅ Verify credentials in LogMeIn admin panel
- ✅ Ensure API access is enabled for your account

**403 Forbidden**
- ❌ API access not enabled
- ✅ Contact LogMeIn support to enable API access

**404 Not Found**
- ❌ Incorrect API endpoint
- ✅ Verify base URL: `https://secure.logmeinrescue.com/API`

**Empty Response**
- ✅ No active sessions - this is normal!
- ✅ Start a test session to verify

### Digium/Switchvox

**Authentication Failed**
- ❌ Incorrect credentials in XML request
- ✅ Verify username/password
- ✅ Check if user has API permissions

**Method Not Found**
- ❌ Incorrect method name
- ✅ Check exact spelling: `switchvox.currentCalls.getList`
- ✅ Refer to Digium API documentation

**XML Parse Error**
- ❌ Malformed XML request
- ✅ Validate XML structure
- ✅ Check for special characters in credentials

**Connection Timeout**
- ❌ Incorrect API URL
- ✅ Verify: `https://nova.digiumcloud.net/xml`
- ✅ Check firewall/network settings

---

## Testing the Dashboard Backend

### Health Check
```bash
curl http://localhost:3001/api/health
```

**Expected Response**:
```json
{
  "status": "ok",
  "timestamp": "2025-11-09T...",
  "services": {
    "logmein": true,
    "digium": true
  }
}
```

### Test LogMeIn Endpoint
```bash
curl http://localhost:3001/api/rescue/tech-available
```

### Test Digium Endpoint
```bash
curl http://localhost:3001/api/digium/current-calls
```

---

## API Rate Limits

### LogMeIn Rescue
- Typically: 60 requests per minute
- Dashboard polls every 10 seconds = 6 requests/minute ✅

### Digium/Switchvox
- Typically: No hard limit, but be reasonable
- Dashboard polls every 5 seconds = 12 requests/minute ✅

**Both are well within limits!**

---

## Security Best Practices

1. ✅ **Never commit `.env` file**
   ```bash
   # Already in .gitignore
   ```

2. ✅ **Use environment variables**
   - No hardcoded credentials
   - Easy to change per environment

3. ✅ **Backend proxies all API calls**
   - Credentials never exposed to frontend
   - No CORS issues

4. ⚠️ **For Production**:
   - Use HTTPS
   - Add rate limiting
   - Implement user authentication
   - Use secret management (AWS Secrets, Azure Key Vault, etc.)

---

## Need More Help?

### Official Documentation
- [LogMeIn Rescue API Docs](https://secure.logmeinrescue.com/welcome/webhelp/EN/RescueAPI/API/)
- [Digium Switchvox API Wiki](http://developers.digium.com/switchvox/wiki/)

### Support
- Check backend logs: `cd backend && npm run dev`
- Check browser console: F12 in browser
- Test APIs directly with curl commands above

---

**Once your credentials are set, the dashboard will automatically pull data! 🎉**
