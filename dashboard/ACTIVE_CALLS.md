# Active Calls Display - Technical Documentation

## Overview

The active calls list now displays complete information from the Digium/Switchvox API, including both **extensions** and **account IDs** for easy monitoring.

## API Response Structure

### Digium API Call
**Method**: `switchvox.currentCalls.getList`

### Response Format (XML → JSON)
```javascript
{
  calls: [
    {
      id: "12345",
      accountId: "1381",           // Account ID for monitoring
      extension: "252",             // Extension number
      callerNumber: "+15551234567", // Caller's phone number
      callerName: "John Doe",       // Caller's name
      calledNumber: "+15559876543", // Called number
      direction: "inbound",         // "inbound" or "outbound"
      status: "connected",          // Call status
      duration: 125,                // Duration in seconds
      startTime: "2025-11-09T10:30:00Z"
    }
  ]
}
```

## Display Features

### 1. Call Card Information
Each active call shows:
- ✅ **Direction** (incoming/outgoing with icons)
- ✅ **Extension** (if available)
- ✅ **Account ID** (if available)
- ✅ **Duration** (formatted as MM:SS)
- ✅ **Caller Name** (if available)
- ✅ **Caller Number** (if available)
- ✅ **Called Number** (if available)
- ✅ **Status** (connected, ringing, etc.)

### 2. Copy-to-Clipboard Feature
- Click the copy button next to extension or account ID
- Value is copied to clipboard
- Shows checkmark ✓ for 2 seconds
- Paste directly into Call Monitoring controls

### 3. Visual Hints
- 💡 Helper text reminds users how to monitor
- Color-coded by call direction (green=incoming, blue=outgoing)
- Highlighted extension/account ID values

## Backend Processing

### Raw XML Parsing
```javascript
// Backend extracts these fields from XML:
{
  id: call.id[0],
  accountId: call.account_id[0],      // KEY: Used for monitoring
  extension: call.extension[0],        // KEY: User-friendly display
  callerNumber: call.caller_number[0],
  callerName: call.caller_name[0],
  calledNumber: call.called_number[0],
  direction: call.direction[0],
  status: call.status[0],
  duration: call.duration[0],
  startTime: call.start_time[0]
}
```

### Logging
Backend logs show:
```
[Current Calls] Raw API Response: {...}
[Current Calls] Parsed calls: [...]
```

## Frontend Display Logic

### Duration Formatting
```javascript
// Converts seconds to MM:SS
const displayDuration = call.duration ? 
  (typeof call.duration === 'number' ? 
    `${Math.floor(call.duration / 60)}:${String(call.duration % 60).padStart(2, '0')}` : 
    call.duration) : 
  '00:00';
```

### Direction Detection
```javascript
const isIncoming = call.direction === 'inbound' || call.direction === 'incoming';
```

### Statistics Calculation
```javascript
const activeCalls = currentCalls.length;
const incomingCalls = currentCalls.filter(c => 
  c.direction === 'inbound' || c.direction === 'incoming'
).length;
const outgoingCalls = currentCalls.filter(c => 
  c.direction === 'outbound' || c.direction === 'outgoing'
).length;
```

## Monitoring Workflow

### Step-by-Step Usage

1. **View Active Calls**
   - See list of all active calls
   - Each shows extension + account ID

2. **Copy Values**
   - Click copy button next to extension or account ID
   - Value is copied to clipboard

3. **Start Monitoring**
   - Scroll to Call Monitoring section (top of panel)
   - Paste values into:
     - "Your Extension" field
     - "Target Extension" field
   - Click "Start Monitoring"

4. **Automatic Mapping**
   - Dashboard maps extensions to account IDs
   - Sends correct account IDs to Digium API
   - Your phone starts monitoring the call

## API Field Mapping

| Digium API Field | Our Display Name | Used For |
|------------------|------------------|----------|
| `account_id` | Account ID | Monitoring API calls |
| `extension` | Extension | User display |
| `caller_number` | Number | Caller identification |
| `caller_name` | Caller | Caller identification |
| `called_number` | Called | Call destination |
| `direction` | Direction | Incoming/Outgoing icon |
| `status` | Status | Call state badge |
| `duration` | Duration | Time display |

## Error Handling

### No Calls Available
```javascript
if (currentCalls.length === 0) {
  // Show: "No active calls"
}
```

### Missing Fields
```javascript
// All fields are optional and checked before display
{call.extension && <span>Extension: {call.extension}</span>}
{call.accountId && <span>Account ID: {call.accountId}</span>}
```

### API Errors
```javascript
// Caught and logged in backend
console.error('Digium current-calls error:', error.message);
```

## Debugging

### Backend Logs
Check terminal for:
```
[Current Calls] Raw API Response: {...}
[Current Calls] Parsed calls: [...]
```

### Frontend Console
Check browser console (F12) for:
```javascript
[DigiumPanel] Fetched calls: [...]
```

### Common Issues

**No calls showing despite active calls**
- Check backend logs for API response structure
- XML structure may differ from expected
- Add console.log to see raw response

**Extension vs Account ID confusion**
- **Extension**: User-facing number (e.g., 252)
- **Account ID**: Internal ID for API (e.g., 1381)
- Monitoring requires Account ID
- Dashboard handles mapping automatically

**Copy button not working**
- Requires HTTPS or localhost
- Check browser clipboard permissions
- Fallback: manually type values

## Future Enhancements

Possible improvements:
- [ ] Click call card to auto-fill monitoring form
- [ ] Show which calls are being monitored
- [ ] Call history/log
- [ ] Filter calls by extension
- [ ] Sort calls by duration
- [ ] Export call list

---

**The active calls list now provides all the information needed to monitor any call with a single click!** 🎯
