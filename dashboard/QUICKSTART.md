# 🎯 QUICK START GUIDE

## Initial Setup (First Time Only)

### Step 1: Configure API Credentials

```bash
cd /workspaces/KPI-Project/dashboard/backend
cp .env.example .env
```

Edit the `.env` file with your actual credentials:

```bash
nano .env
```

**Required credentials:**
- LogMeIn Rescue username & password
- Digium/Switchvox username & password

Save and exit (Ctrl+X, Y, Enter in nano)

### Step 2: Install Dependencies

```bash
cd /workspaces/KPI-Project/dashboard
./install.sh
```

**Note**: You may see warnings about vulnerabilities - these are from dependencies and don't affect functionality.

## Running the Dashboard

### Option 1: Automatic (Recommended)

```bash
cd /workspaces/KPI-Project/dashboard
./start.sh
```

This will start both backend and frontend servers automatically!

### Option 2: Manual (Two Terminals)

**Terminal 1 - Backend:**
```bash
cd /workspaces/KPI-Project/dashboard
npm run backend
```

**Terminal 2 - Frontend:**
```bash
cd /workspaces/KPI-Project/dashboard
npm run frontend
```

## Accessing the Dashboard

Once both servers are running:

1. Open your browser
2. Go to: **http://localhost:3000**
3. Enjoy your live dashboard! 🎉

## What You'll See

### LogMeIn Rescue Panel
- ✅ Tech availability status
- 📊 Count of active support sessions
- 📋 List of all active sessions with details

### Digium/Switchvox Panel
- 📞 Active call count
- ↗️ Incoming calls
- ↙️ Outgoing calls
- 🎧 One-click call monitoring controls
- 📋 Detailed list of all active calls

## Using Call Monitoring

1. Enter **your extension** (e.g., 1001)
2. Enter **target extension** to monitor (e.g., 1002)
   - OR click "Monitor This Call" on any active call
3. Click **"Start Monitoring"**
4. Your phone will automatically join the call as a monitor
5. Click **"Stop Monitoring"** when done

## Refresh Rates

- LogMeIn Rescue: Updates every **10 seconds**
- Digium Calls: Updates every **5 seconds**
- Click the **"Refresh"** button in the header to reload immediately

## Troubleshooting

### "Failed to fetch" errors
- Make sure the backend server is running (port 3001)
- Check that your `.env` file has valid credentials

### Blank dashboard
- Check the browser console (F12) for errors
- Verify both frontend and backend are running
- Check that APIs are returning data (visit http://localhost:3001/api/health)

### Call monitoring not working
- Verify your extension number is correct
- Make sure you have permissions in Digium
- Check that the target extension is active

## Next Steps

- Customize the refresh intervals in the component files
- Modify the CSS to match your brand colors
- Add more API endpoints as needed
- Set up authentication for production use

## Need Help?

Check the main README.md for detailed documentation:
```bash
cat /workspaces/KPI-Project/dashboard/README.md
```

---

**Happy Monitoring! 🚀**
