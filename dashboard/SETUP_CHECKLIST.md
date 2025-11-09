# ✅ First-Time Setup Checklist

Follow this checklist to get your dashboard up and running!

## 📋 Pre-Setup

- [ ] You have Node.js 18+ installed
  - Check: `node --version`
  - If not installed: [Download Node.js](https://nodejs.org/)

- [ ] You have LogMeIn Rescue API credentials ready
  - Username: __________________
  - Password/API Key: __________________

- [ ] You have Digium/Switchvox API credentials ready
  - Username: __________________
  - Password: __________________

## 🔧 Setup Steps

### Step 1: Configure Backend Credentials

- [ ] Navigate to backend directory
  ```bash
  cd /workspaces/KPI-Project/dashboard/backend
  ```

- [ ] Open `.env` file for editing
  ```bash
  nano .env
  ```

- [ ] Update `LOGMEIN_USERNAME` with your LogMeIn username
- [ ] Update `LOGMEIN_PASSWORD` with your LogMeIn password
- [ ] Update `DIGIUM_USERNAME` with your Digium username  
- [ ] Update `DIGIUM_PASSWORD` with your Digium password
- [ ] Save file (Ctrl+X, Y, Enter)

### Step 2: Verify Installation

- [ ] Check that backend dependencies are installed
  ```bash
  cd /workspaces/KPI-Project/dashboard/backend
  ls node_modules/
  ```
  Should see folders like `express`, `axios`, etc.

- [ ] Check that frontend dependencies are installed
  ```bash
  cd /workspaces/KPI-Project/dashboard/frontend
  ls node_modules/
  ```
  Should see folders like `react`, `vite`, etc.

- [ ] If missing, install dependencies:
  ```bash
  cd /workspaces/KPI-Project/dashboard
  npm run install-all
  ```

### Step 3: Test Backend

- [ ] Start the backend server
  ```bash
  cd /workspaces/KPI-Project/dashboard/backend
  npm start
  ```

- [ ] Should see:
  ```
  🚀 DevOps Dashboard Backend running on port 3001
  📊 Health check: http://localhost:3001/api/health
  ```

- [ ] Test health endpoint (in another terminal):
  ```bash
  curl http://localhost:3001/api/health
  ```
  Should return JSON with `"status": "ok"`

- [ ] Stop the backend (Ctrl+C)

### Step 4: Run Full Application

- [ ] Use the startup script
  ```bash
  cd /workspaces/KPI-Project/dashboard
  ./start.sh
  ```

- [ ] Wait for both servers to start:
  - Backend: `http://localhost:3001`
  - Frontend: `http://localhost:3000`

- [ ] Open browser to `http://localhost:3000`

## ✨ Verification Checklist

Once the dashboard loads in your browser:

### Visual Check
- [ ] Page loads without errors
- [ ] Header shows "DevOps Live Dashboard"
- [ ] System status shows "Online" with green dot
- [ ] Two panels visible: LogMeIn Rescue and Digium Switchvox

### LogMeIn Rescue Panel
- [ ] Panel shows tech availability status
- [ ] Panel shows count of active sessions
- [ ] If no sessions: Shows "No active sessions" (this is normal!)
- [ ] If sessions exist: Shows session details
- [ ] No error messages in panel

### Digium Panel  
- [ ] Panel shows active call count
- [ ] Panel shows incoming/outgoing calls
- [ ] Shows call monitoring controls
- [ ] If no calls: Shows "No active calls" (this is normal!)
- [ ] If calls exist: Shows call details
- [ ] No error messages in panel

### Functionality Check
- [ ] Click "Refresh" button - page refreshes
- [ ] Health indicator pulses
- [ ] No console errors (press F12 to check)
- [ ] Data updates automatically after 5-10 seconds

## 🐛 Troubleshooting

If something's not working, check these:

### Backend Issues
- [ ] Check terminal for backend errors
- [ ] Verify `.env` file has all credentials filled in
- [ ] Test credentials with curl commands in API_CONFIG.md
- [ ] Check port 3001 is not in use: `lsof -i :3001`

### Frontend Issues
- [ ] Check browser console (F12) for errors
- [ ] Verify frontend is running on port 3000
- [ ] Check backend is running on port 3001
- [ ] Try clearing browser cache

### API Issues
- [ ] Verify API URLs are correct in `.env`
- [ ] Test credentials directly with API providers
- [ ] Check network connectivity
- [ ] Review logs for authentication errors

### Common Errors

**"Failed to fetch"**
- → Backend is not running
- → Check backend terminal for errors

**"401 Unauthorized"**
- → Invalid credentials in `.env`
- → Double-check username/password

**"Cannot GET /api/..."**
- → Backend route missing
- → Check backend is running

## 📚 Next Steps

Once everything works:

- [ ] Read the full README.md
- [ ] Customize refresh intervals if desired
- [ ] Adjust styling/colors to match your branding
- [ ] Share the dashboard URL with your team
- [ ] Set up on a dedicated server for 24/7 monitoring

## 🎉 Success Criteria

You're done when:

✅ Dashboard loads in browser
✅ No error messages
✅ Data displays (or "No active" messages if no data)
✅ Auto-refresh works
✅ Call monitoring controls respond
✅ System health shows "Online"

---

## 🆘 Still Need Help?

1. Check console logs: `F12` in browser
2. Check backend logs in terminal
3. Review API_CONFIG.md for API troubleshooting
4. Test APIs directly with curl commands
5. Verify credentials with your API providers

---

**Happy Monitoring! 🚀**

---

## 📝 Notes

Use this space to write down your specific configuration:

**Your Extension Number**: ________________

**Common Target Extensions**:
- ________________
- ________________
- ________________

**API Notes**:
- ________________________________
- ________________________________

**Customizations Made**:
- ________________________________
- ________________________________
