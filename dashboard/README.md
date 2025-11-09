# 🚀 DevOps Live Dashboard

A beautiful, real-time dashboard for monitoring LogMeIn Rescue sessions and Digium/Switchvox phone systems.

![Dashboard Preview](https://via.placeholder.com/800x400/667eea/ffffff?text=DevOps+Live+Dashboard)

## ✨ Features

### LogMeIn Rescue Integration
- 👥 **Tech Availability Monitoring** - Real-time check if technicians are available
- 📊 **Active Sessions Display** - View all current support sessions
- 🔄 **Auto-refresh** - Updates every 10 seconds

### Digium/Switchvox Integration
- 📞 **Active Call Monitoring** - See all current calls in real-time
- 📈 **Call Statistics** - Track incoming/outgoing calls
- 🎧 **One-Click Call Monitoring** - Jump into calls with a button click
- 📋 **Queue Status** - Monitor call queue statistics
- 🔄 **Auto-refresh** - Updates every 5 seconds

## 🛠️ Tech Stack

- **Frontend**: React 18 + Vite
- **Backend**: Node.js + Express
- **Styling**: Custom CSS with animations
- **Icons**: Lucide React
- **HTTP Client**: Axios

## 📋 Prerequisites

- Node.js 18+ installed
- LogMeIn Rescue API credentials
- Digium/Switchvox API credentials
- Basic auth username and password for both services

## 🚀 Quick Start

### 1. Clone and Setup

```bash
cd /workspaces/KPI-Project/dashboard
```

### 2. Backend Setup

```bash
cd backend

# Install dependencies
npm install

# Configure environment variables
cp .env.example .env

# Edit .env with your credentials
nano .env
```

**Update `.env` with your credentials:**

```env
# LogMeIn Rescue API Credentials
LOGMEIN_USERNAME=your_username_here
LOGMEIN_PASSWORD=your_password_here
LOGMEIN_API_URL=https://secure.logmeinrescue.com/API

# Digium/Switchvox API Credentials
DIGIUM_USERNAME=your_username_here
DIGIUM_PASSWORD=your_password_here
DIGIUM_API_URL=https://nova.digiumcloud.net/xml

# Server Configuration
PORT=3001
```

### 3. Frontend Setup

```bash
cd ../frontend

# Install dependencies
npm install
```

### 4. Run the Application

**Terminal 1 - Backend:**
```bash
cd backend
npm run dev
```

**Terminal 2 - Frontend:**
```bash
cd frontend
npm run dev
```

### 5. Access the Dashboard

Open your browser to: **http://localhost:3000**

The backend API will be running on: **http://localhost:3001**

## 📡 API Endpoints

### Backend API Routes

#### LogMeIn Rescue
- `GET /api/rescue/tech-available` - Check tech availability
- `GET /api/rescue/sessions` - Get active sessions

#### Digium/Switchvox
- `GET /api/digium/queue-status` - Get call queue status
- `GET /api/digium/current-calls` - Get current active calls
- `GET /api/digium/monitoring-info` - Get call monitoring information
- `POST /api/digium/start-monitoring` - Start monitoring a call
- `POST /api/digium/stop-monitoring` - Stop monitoring

#### Health Check
- `GET /api/health` - Check server and credentials status

## 🎨 Dashboard Features

### Real-Time Updates
- LogMeIn Rescue updates every 10 seconds
- Digium/Switchvox updates every 5 seconds
- Visual indicators for system health
- Smooth animations and transitions

### Call Monitoring Workflow
1. Enter your extension number
2. Enter the target extension to monitor (or click "Monitor This Call" on any active call)
3. Click "Start Monitoring"
4. Your phone will automatically start monitoring the call
5. Click "Stop Monitoring" to end monitoring

## 🔧 Configuration

### Adjusting Refresh Intervals

Edit the interval values in the component files:

**RescuePanel.jsx** (line ~29):
```javascript
const interval = setInterval(fetchData, 10000); // 10 seconds
```

**DigiumPanel.jsx** (line ~33):
```javascript
const interval = setInterval(fetchData, 5000); // 5 seconds
```

### Customizing Styles

Each component has its own CSS file:
- `App.css` - Main layout and header
- `RescuePanel.css` - LogMeIn Rescue panel styles
- `DigiumPanel.css` - Digium panel styles

## 🐛 Troubleshooting

### Backend not starting
- Verify Node.js is installed: `node --version`
- Check if port 3001 is available
- Ensure `.env` file exists with valid credentials

### Frontend not connecting to backend
- Verify backend is running on port 3001
- Check browser console for errors
- Verify proxy configuration in `vite.config.js`

### API Authentication Errors
- Verify credentials in `.env` file
- Test credentials directly with API documentation
- Check API URLs are correct

### CORS Issues
- Backend has CORS enabled by default
- If issues persist, check backend logs

## 📦 Production Build

### Backend
```bash
cd backend
npm start
```

### Frontend
```bash
cd frontend
npm run build
npm run preview
```

## 🔐 Security Notes

- Never commit `.env` file to version control
- Use environment variables for all sensitive data
- Backend proxies API calls to protect credentials
- Consider adding authentication for production use

## 🎯 Future Enhancements

- [ ] User authentication and login
- [ ] Historical data and analytics
- [ ] Customizable dashboards
- [ ] Email/SMS notifications
- [ ] Dark mode
- [ ] Mobile app version
- [ ] WebSocket for real-time updates (instead of polling)
- [ ] Call recording integration
- [ ] Advanced filtering and search

## 📄 License

MIT License - feel free to use and modify!

## 🤝 Contributing

This is a custom internal tool, but suggestions and improvements are welcome!

## 📞 Support

For issues or questions, refer to the official API documentation:
- [LogMeIn Rescue API](https://secure.logmeinrescue.com/welcome/webhelp/EN/RescueAPI/API/)
- [Digium/Switchvox API](http://developers.digium.com/switchvox/wiki/)

---

**Built with ❤️ for DevOps teams**
