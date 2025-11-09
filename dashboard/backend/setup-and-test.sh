#!/bin/bash

# Setup and Test Script for LogMeIn Rescue & Digium APIs
# This script helps you configure credentials and test authentication

echo "🔧 API Credentials Setup and Test"
echo "=================================="
echo ""

# Check if .env file exists
if [ ! -f .env ]; then
    echo "📝 Creating .env file from template..."
    cp .env.example .env
    echo "✅ .env file created"
    echo ""
fi

# Function to update .env file
update_env() {
    local key=$1
    local value=$2
    local file=.env
    
    if grep -q "^${key}=" "$file"; then
        # Update existing value
        sed -i "s|^${key}=.*|${key}=${value}|" "$file"
    else
        # Add new value
        echo "${key}=${value}" >> "$file"
    fi
}

# Check current credentials
echo "📋 Current Credentials Status:"
echo ""

source .env 2>/dev/null

if [ "$LOGMEIN_USERNAME" = "your_username_here" ] || [ -z "$LOGMEIN_USERNAME" ]; then
    echo "❌ LogMeIn Rescue credentials not configured"
    NEED_LOGMEIN=true
else
    echo "✅ LogMeIn Rescue: $LOGMEIN_USERNAME"
    NEED_LOGMEIN=false
fi

if [ "$DIGIUM_USERNAME" = "your_username_here" ] || [ -z "$DIGIUM_USERNAME" ]; then
    echo "❌ Digium/Switchvox credentials not configured"
    NEED_DIGIUM=true
else
    echo "✅ Digium/Switchvox: $DIGIUM_USERNAME"
    NEED_DIGIUM=false
fi

echo ""
echo "=================================="
echo ""

# Offer to configure credentials
if [ "$NEED_LOGMEIN" = true ] || [ "$NEED_DIGIUM" = true ]; then
    echo "Would you like to configure credentials now? (y/n)"
    read -r CONFIGURE
    
    if [ "$CONFIGURE" = "y" ] || [ "$CONFIGURE" = "Y" ]; then
        echo ""
        
        if [ "$NEED_LOGMEIN" = true ]; then
            echo "🛟 LogMeIn Rescue Configuration:"
            echo ""
            echo "Enter LogMeIn Rescue username (email):"
            read -r LOGMEIN_USER
            echo "Enter LogMeIn Rescue password:"
            read -rs LOGMEIN_PASS
            echo ""
            
            update_env "LOGMEIN_USERNAME" "$LOGMEIN_USER"
            update_env "LOGMEIN_PASSWORD" "$LOGMEIN_PASS"
            
            echo "✅ LogMeIn credentials saved"
            echo ""
        fi
        
        if [ "$NEED_DIGIUM" = true ]; then
            echo "📞 Digium/Switchvox Configuration:"
            echo ""
            echo "Enter Digium username:"
            read -r DIGIUM_USER
            echo "Enter Digium password:"
            read -rs DIGIUM_PASS
            echo ""
            
            update_env "DIGIUM_USERNAME" "$DIGIUM_USER"
            update_env "DIGIUM_PASSWORD" "$DIGIUM_PASS"
            
            echo "✅ Digium credentials saved"
            echo ""
        fi
        
        echo "✅ Credentials saved to .env file"
        echo ""
    fi
fi

echo "=================================="
echo ""
echo "🧪 Running Authentication Tests..."
echo ""

# Run the test script
npm run test-creds

echo ""
echo "=================================="
echo ""
echo "📝 Next Steps:"
echo ""
echo "1. If tests passed, start the server:"
echo "   npm run dev"
echo ""
echo "2. If tests failed, check your credentials:"
echo "   nano .env"
echo ""
echo "3. Re-run this script to test again:"
echo "   ./setup-and-test.sh"
echo ""
echo "4. View detailed documentation:"
echo "   cat ../RESCUE_LOGIN_IMPLEMENTATION.md"
echo ""
