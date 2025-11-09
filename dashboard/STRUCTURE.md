# 📁 Project Structure

```
dashboard/
├── backend/                    # Express API server
│   ├── node_modules/          # Backend dependencies
│   ├── .env                   # API credentials (DO NOT COMMIT)
│   ├── .env.example           # Template for credentials
│   ├── package.json           # Backend dependencies
│   └── server.js              # Express server with API routes
│
├── frontend/                  # React + Vite application
│   ├── node_modules/          # Frontend dependencies
│   ├── public/                # Static assets
│   ├── src/
│   │   ├── components/
│   │   │   ├── RescuePanel.jsx      # LogMeIn Rescue component
│   │   │   ├── RescuePanel.css      # Rescue panel styles
│   │   │   ├── DigiumPanel.jsx      # Digium/Switchvox component
│   │   │   └── DigiumPanel.css      # Digium panel styles
│   │   ├── App.jsx            # Main application component
│   │   ├── App.css            # Main application styles
│   │   ├── api.js             # API client functions
│   │   ├── main.jsx           # React entry point
│   │   └── index.css          # Global styles
│   ├── index.html             # HTML template
│   ├── package.json           # Frontend dependencies
│   └── vite.config.js         # Vite configuration
│
├── .gitignore                 # Git ignore rules
├── package.json               # Root package.json with scripts
├── start.sh                   # Startup script (runs both servers)
├── README.md                  # Full documentation
└── QUICKSTART.md              # Quick start guide

```

## 🎯 Key Files Explained

### Backend Files

#### `server.js`
The Express server that handles:
- API routes for LogMeIn Rescue
- API routes for Digium/Switchvox
- Authentication with Basic Auth
- XML parsing for Digium responses
- CORS configuration

#### `.env`
Stores sensitive credentials:
```
LOGMEIN_USERNAME=your_username
LOGMEIN_PASSWORD=your_password
DIGIUM_USERNAME=your_username
DIGIUM_PASSWORD=your_password
```

### Frontend Files

#### `App.jsx`
Main dashboard container with:
- Header with health status
- Grid layout for panels
- Refresh functionality

#### `RescuePanel.jsx`
LogMeIn Rescue monitoring:
- Tech availability check
- Active sessions list
- Auto-refresh every 10s

#### `DigiumPanel.jsx`
Digium/Switchvox monitoring:
- Call statistics
- Active calls list
- Call monitoring controls
- Auto-refresh every 5s

#### `api.js`
Centralized API calls:
- `rescueApi.*` - LogMeIn functions
- `digiumApi.*` - Digium functions
- `healthCheck()` - Server status

### Configuration Files

#### `vite.config.js`
- Development server on port 3000
- Proxy to backend on port 3001
- React plugin configuration

#### `package.json` (root)
Convenient npm scripts:
- `npm run install-all` - Install all dependencies
- `npm run backend` - Start backend only
- `npm run frontend` - Start frontend only
- `npm start` - Run start.sh script

## 🔄 Data Flow

```
┌─────────────┐
│   Browser   │ (http://localhost:3000)
└──────┬──────┘
       │
       ▼
┌─────────────────────┐
│  React Frontend     │
│  - RescuePanel      │
│  - DigiumPanel      │
└──────┬──────────────┘
       │ /api/* requests
       ▼
┌─────────────────────┐
│  Express Backend    │ (http://localhost:3001)
│  - CORS enabled     │
│  - Basic Auth       │
└──────┬──────────────┘
       │
       ├─────► LogMeIn Rescue API
       │       (https://secure.logmeinrescue.com/API)
       │
       └─────► Digium/Switchvox API
               (https://nova.digiumcloud.net/xml)
```

## 🌐 API Endpoints

### Backend Endpoints

| Method | Endpoint | Purpose |
|--------|----------|---------|
| GET | `/api/health` | Health check |
| GET | `/api/rescue/tech-available` | Check tech availability |
| GET | `/api/rescue/sessions` | Get active sessions |
| GET | `/api/digium/queue-status` | Get call queue status |
| GET | `/api/digium/current-calls` | Get current calls |
| GET | `/api/digium/monitoring-info` | Get monitoring info |
| POST | `/api/digium/start-monitoring` | Start call monitoring |
| POST | `/api/digium/stop-monitoring` | Stop call monitoring |

## 🎨 Styling Architecture

- **CSS Modules**: Each component has its own CSS file
- **Gradient Theme**: Purple/blue gradient (#667eea → #764ba2)
- **Animations**: Fade-in, slide-up, pulse effects
- **Responsive**: Mobile-friendly breakpoints
- **Glass Morphism**: Semi-transparent elements with blur

## 📦 Dependencies

### Backend
- `express` - Web server
- `cors` - Cross-origin requests
- `axios` - HTTP client
- `dotenv` - Environment variables
- `xml2js` - XML parsing for Digium

### Frontend
- `react` - UI framework
- `react-dom` - React renderer
- `axios` - HTTP client
- `lucide-react` - Icon library
- `vite` - Build tool

## 🚀 Development Workflow

1. **Edit backend**: Modify `backend/server.js`
2. **Edit frontend**: Modify components in `frontend/src/`
3. **Add API routes**: Add to `server.js` and `api.js`
4. **Style changes**: Edit component `.css` files
5. **Test**: Servers auto-reload on changes

## 🔐 Security Checklist

- ✅ `.env` in `.gitignore`
- ✅ Backend proxies API calls
- ✅ Basic Auth handled server-side
- ✅ CORS configured properly
- ⚠️ Add user authentication for production
- ⚠️ Use HTTPS in production
- ⚠️ Rate limiting recommended

---

**Need to modify something? All files are well-organized and documented!**
