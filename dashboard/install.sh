#!/bin/bash

# Installation script for DevOps Dashboard
# This script installs all dependencies for both backend and frontend

echo "📦 Installing DevOps Dashboard Dependencies"
echo ""

# Backend installation
echo "🔧 Installing backend dependencies..."
cd backend
if npm install; then
    echo "✅ Backend dependencies installed successfully!"
else
    echo "❌ Backend installation failed!"
    exit 1
fi
cd ..

echo ""

# Frontend installation
echo "🎨 Installing frontend dependencies..."
cd frontend
if npm install; then
    echo "✅ Frontend dependencies installed successfully!"
else
    echo "❌ Frontend installation failed!"
    exit 1
fi
cd ..

echo ""
echo "🎉 All dependencies installed successfully!"
echo ""
echo "⚠️  Note: You may see some vulnerability warnings. These are from dependencies"
echo "   and don't affect functionality. Run 'npm audit' in backend/frontend for details."
echo ""
echo "Next steps:"
echo "1. Configure your API credentials in backend/.env"
echo "2. Run './start.sh' to start the dashboard"
echo ""
