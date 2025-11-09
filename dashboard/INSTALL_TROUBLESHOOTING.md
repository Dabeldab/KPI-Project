# 🔧 Installation Troubleshooting Guide

## Common Installation Issues

### ✅ **Everything Installed Successfully!**

If you see this message, you're good to go:
```
🎉 All dependencies installed successfully!
```

The warnings about vulnerabilities are normal and don't affect functionality.

---

## Issue: "npm: command not found"

**Problem**: Node.js/npm is not installed

**Solution**:
```bash
# Check if Node.js is installed
node --version

# If not installed, install Node.js 18+
# For Ubuntu/Debian:
curl -fsSL https://deb.nodesource.com/setup_18.x | sudo -E bash -
sudo apt-get install -y nodejs

# Verify installation
node --version
npm --version
```

---

## Issue: "Permission denied" when running ./install.sh

**Problem**: Script is not executable

**Solution**:
```bash
chmod +x /workspaces/KPI-Project/dashboard/install.sh
./install.sh
```

---

## Issue: Vulnerability Warnings

**Message**: "2 moderate severity vulnerabilities"

**Is this a problem?**: No! This is normal.

**Explanation**: 
- These are vulnerabilities in development dependencies
- They don't affect the running application
- They're from third-party packages (React, Vite, etc.)
- The application is safe to use

**If you want to fix them anyway**:
```bash
cd /workspaces/KPI-Project/dashboard/frontend
npm audit fix

# If that doesn't work and you understand the risks:
npm audit fix --force
```

⚠️ **Warning**: `npm audit fix --force` may break things. Only use if you know what you're doing.

---

## Issue: "EACCES: permission denied"

**Problem**: Permission issues with npm global packages

**Solution**:
```bash
# Option 1: Use the local install (recommended)
cd /workspaces/KPI-Project/dashboard
./install.sh

# Option 2: Fix npm permissions
mkdir ~/.npm-global
npm config set prefix '~/.npm-global'
echo 'export PATH=~/.npm-global/bin:$PATH' >> ~/.bashrc
source ~/.bashrc
```

---

## Issue: "ENOSPC: no space left on device"

**Problem**: Disk is full

**Solution**:
```bash
# Check disk space
df -h

# Clean npm cache
npm cache clean --force

# Remove node_modules and reinstall
cd /workspaces/KPI-Project/dashboard
rm -rf backend/node_modules frontend/node_modules
./install.sh
```

---

## Issue: "Cannot find module" errors

**Problem**: Dependencies not installed correctly

**Solution**:
```bash
cd /workspaces/KPI-Project/dashboard

# Clean install
rm -rf backend/node_modules backend/package-lock.json
rm -rf frontend/node_modules frontend/package-lock.json

# Reinstall
./install.sh
```

---

## Issue: Network/Proxy Errors

**Message**: "network timeout", "ETIMEDOUT", "ENOTFOUND"

**Solution**:
```bash
# Check internet connection
ping registry.npmjs.org

# Try with different registry
npm config set registry https://registry.npmjs.org/

# If behind proxy, configure npm
npm config set proxy http://proxy.company.com:8080
npm config set https-proxy http://proxy.company.com:8080

# Retry installation
./install.sh
```

---

## Issue: Version Conflicts

**Message**: "peer dependency" warnings

**Solution**: These are usually just warnings, not errors. The app will still work.

If you want to resolve them:
```bash
cd /workspaces/KPI-Project/dashboard/frontend
npm install --legacy-peer-deps
```

---

## Issue: Package Lock Errors

**Message**: "corrupt package-lock.json"

**Solution**:
```bash
cd /workspaces/KPI-Project/dashboard

# Remove lock files
rm -f backend/package-lock.json frontend/package-lock.json

# Reinstall
./install.sh
```

---

## Issue: Node Version Incompatibility

**Problem**: "requires node version X" errors

**Solution**:
```bash
# Check current version
node --version

# Need Node.js 18 or higher
# Install/upgrade Node.js if needed

# Using nvm (Node Version Manager)
nvm install 18
nvm use 18

# Verify
node --version

# Reinstall dependencies
./install.sh
```

---

## Verification Steps

After installation, verify everything is working:

### 1. Check Backend Dependencies
```bash
cd /workspaces/KPI-Project/dashboard/backend
ls node_modules/ | wc -l
# Should show 80+ packages
```

### 2. Check Frontend Dependencies
```bash
cd /workspaces/KPI-Project/dashboard/frontend
ls node_modules/ | wc -l
# Should show 60+ packages
```

### 3. Test Backend Startup
```bash
cd /workspaces/KPI-Project/dashboard/backend
npm start
# Should show: "🚀 DevOps Dashboard Backend running on port 3001"
# Press Ctrl+C to stop
```

### 4. Test Frontend Startup
```bash
cd /workspaces/KPI-Project/dashboard/frontend
npm run dev
# Should show: "VITE v5.x.x ready in Xms"
# Press Ctrl+C to stop
```

---

## Still Having Issues?

### Check the logs

**Backend logs**:
```bash
cd /workspaces/KPI-Project/dashboard/backend
npm start 2>&1 | tee backend-log.txt
```

**Frontend logs**:
```bash
cd /workspaces/KPI-Project/dashboard/frontend
npm run dev 2>&1 | tee frontend-log.txt
```

### Common Quick Fixes

```bash
# Nuclear option - clean everything and start fresh
cd /workspaces/KPI-Project/dashboard

# Remove all installed packages
rm -rf backend/node_modules frontend/node_modules
rm -rf backend/package-lock.json frontend/package-lock.json

# Clear npm cache
npm cache clean --force

# Reinstall everything
./install.sh
```

---

## System Requirements

Make sure your system meets these requirements:

- ✅ **Node.js**: Version 18 or higher
- ✅ **npm**: Version 9 or higher (comes with Node.js)
- ✅ **Disk Space**: At least 500MB free
- ✅ **RAM**: At least 2GB available
- ✅ **OS**: Linux, macOS, or Windows with WSL
- ✅ **Internet**: Required for downloading packages

Check your system:
```bash
node --version    # Should be 18.x or higher
npm --version     # Should be 9.x or higher
df -h .           # Check disk space
free -h           # Check RAM (Linux)
```

---

## Success Indicators

You know installation worked when you see:

```
✅ Backend dependencies installed successfully!
✅ Frontend dependencies installed successfully!
🎉 All dependencies installed successfully!
```

And these directories exist:
- `/workspaces/KPI-Project/dashboard/backend/node_modules/`
- `/workspaces/KPI-Project/dashboard/frontend/node_modules/`

---

## Next Steps After Successful Installation

1. ✅ Configure API credentials: `nano backend/.env`
2. ✅ Start the dashboard: `./start.sh`
3. ✅ Open browser: `http://localhost:3000`

---

**If nothing here helped, please check:**
- Node.js and npm are properly installed
- You have internet connectivity
- You're in the correct directory
- You have read/write permissions

---

*Updated: November 9, 2025*
