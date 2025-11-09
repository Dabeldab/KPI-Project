# Extension to Account ID Mapping

## Overview

The Digium/Switchvox call monitoring feature requires **Account IDs** (not extension numbers) to work properly. This dashboard automatically maps extension numbers to their corresponding account IDs.

## Current Mappings

| Extension | Account ID |
|-----------|------------|
| 252 | 1381 |
| 304 | 1442 |
| 305 | 1430 |
| 306 | 1436 |
| 308 | 1421 |
| 322 | 1423 |
| 355 | 1439 |
| 356 | 1351 |

## How It Works

1. **You enter extensions** (e.g., 252)
2. **Dashboard looks up account IDs** (e.g., 1381)
3. **API call uses account IDs** for monitoring

## Adding New Extensions

To add more extensions, edit the mapping in:

**File**: `/dashboard/frontend/src/components/CallMonitoring.jsx`

**Location**: Lines 5-14

```javascript
const EXTENSION_TO_ACCOUNT = {
  '252': '1381',
  '304': '1442',
  // Add new mappings here:
  '999': '1234',  // New extension → account ID
};
```

## Finding Account IDs

To find an account ID for an extension:

1. Use the Digium API: `switchvox.extensions.getInfo`
2. Or check your Switchvox admin panel
3. Or contact your Switchvox administrator

## Error Messages

### "Extension XXX not found in mapping"
**Solution**: The extension you entered is not in the mapping table. Add it to the `EXTENSION_TO_ACCOUNT` object.

### "Available extensions: 252, 304, 305..."
**Info**: This shows all currently mapped extensions you can use.

## Debug Log

The Call Monitoring component includes a built-in debug log that shows:

- ✅ **Extension lookups** - What extensions are being converted
- ✅ **Account ID mappings** - What account IDs are being used
- ✅ **API calls** - When monitoring starts/stops
- ✅ **Errors** - Full error details from the API

### Viewing the Debug Log

1. Open the dashboard
2. Go to the Call Monitoring section
3. Click "Debug Log (X entries)" to expand
4. See real-time logs with timestamps

### Log Colors

- 🔵 **Blue (info)**: General information
- 🟢 **Green (success)**: Successful operations
- 🟡 **Yellow (warning)**: Warnings
- 🔴 **Red (error)**: Errors with details

## Backend Logging

The backend also logs to the terminal:

```bash
[Call Monitoring] Starting: Your Account ID: 1381, Target Account ID: 1442
[Call Monitoring] Success: {...}
[Call Monitoring] Error: {...}
```

Check your backend terminal for detailed server-side logs.

## Testing

### Test Monitoring

1. Enter your extension (e.g., **252**)
2. Enter target extension (e.g., **304**)
3. Click "Start Monitoring"
4. Check the debug log to see:
   ```
   Looking up extensions - Your: 252, Target: 304
   Mapped extensions - Your Account: 1381, Target Account: 1442
   Calling API to start monitoring...
   ✅ Monitoring started!
   ```

### Test Error Handling

1. Enter an invalid extension (e.g., **999**)
2. Try to start monitoring
3. You'll see:
   ```
   ❌ Extension 999 not found in mapping. 
   Available extensions: 252, 304, 305, 306, 308, 322, 355, 356
   ```

## API Parameters

### Old (Incorrect)
```javascript
{
  extension: "252",
  target_extension: "304"
}
```

### New (Correct)
```javascript
{
  account_id: "1381",      // Mapped from extension 252
  target_account_id: "1442" // Mapped from extension 304
}
```

## Troubleshooting

### Monitoring doesn't start
1. Check debug log for errors
2. Verify extensions are in the mapping
3. Check backend terminal logs
4. Verify account IDs are correct in Switchvox

### Wrong call is monitored
- The account ID mapping is incorrect
- Update the `EXTENSION_TO_ACCOUNT` object with correct IDs

### Need to add many extensions
1. Export extension list from Switchvox
2. Use Digium API to get account IDs
3. Update the mapping object in bulk

---

**Pro Tip**: Keep the extension mapping expandable open in the dashboard to see all available extensions while working!
