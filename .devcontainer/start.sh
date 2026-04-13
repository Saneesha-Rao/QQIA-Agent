#!/bin/bash
# Codespace startup script for QQIA Agent
# Detects the Codespace URL, provisions the Teams bot, and starts the server.

set -e

echo "🚀 QQIA Agent - Codespace Setup"
echo "================================"

# ---- Step 1: Detect Codespace URL ----
if [ -n "$CODESPACE_NAME" ]; then
  BOT_DOMAIN="${CODESPACE_NAME}-3978.${GITHUB_CODESPACES_PORT_FORWARDING_DOMAIN}"
  echo "✅ Codespace URL: https://${BOT_DOMAIN}"
else
  BOT_DOMAIN="localhost:3978"
  echo "⚠️  Not in Codespace, using localhost"
fi

# ---- Step 2: Update .env.dev with Codespace domain ----
ENV_FILE="env/.env.dev"
if [ -f "$ENV_FILE" ]; then
  # Update BOT_DOMAIN
  if grep -q "^BOT_DOMAIN=" "$ENV_FILE"; then
    sed -i "s|^BOT_DOMAIN=.*|BOT_DOMAIN=${BOT_DOMAIN}|" "$ENV_FILE"
  else
    echo "BOT_DOMAIN=${BOT_DOMAIN}" >> "$ENV_FILE"
  fi
  echo "✅ Updated BOT_DOMAIN in ${ENV_FILE}"
fi

# ---- Step 3: Set port to public (required for Bot Framework) ----
# Bot Framework Service needs to POST messages to this endpoint.
# Authentication is handled by AAD JWT validation in CloudAdapter — 
# only Microsoft-signed tokens are accepted, so public port is safe.
if [ -n "$CODESPACE_NAME" ]; then
  gh codespace ports visibility 3978:public -c "$CODESPACE_NAME" 2>/dev/null || true
  echo "✅ Port 3978 set to public (secured by Bot Framework AAD auth)"
fi

# ---- Step 4: Provision Teams app (if not already done) ----
if [ -z "$(grep '^BOT_ID=.' "$ENV_FILE" 2>/dev/null)" ]; then
  echo ""
  echo "📱 First-time setup: Provisioning Teams bot..."
  echo "   You'll be prompted to sign in with your Microsoft 365 account."
  echo ""
  npx teamsapp provision --env dev
  echo "✅ Teams bot provisioned"
else
  echo "✅ Teams bot already provisioned (BOT_ID found in ${ENV_FILE})"
  # Still update the messaging endpoint in case Codespace URL changed
  BOT_ID=$(grep '^BOT_ID=' "$ENV_FILE" | cut -d= -f2)
  if [ -n "$BOT_ID" ]; then
    echo "🔄 Updating Bot Framework endpoint to new Codespace URL..."
    npx teamsapp provision --env dev 2>/dev/null || true
  fi
fi

# ---- Step 5: Load env vars and start the bot ----
echo ""
echo "🔧 Loading environment..."

# Export provisioned credentials as env vars for the bot
if [ -f "$ENV_FILE" ]; then
  export MICROSOFT_APP_ID=$(grep '^BOT_ID=' "$ENV_FILE" | cut -d= -f2)
  export MICROSOFT_APP_PASSWORD=$(grep '^SECRET_BOT_PASSWORD=' "$ENV_FILE" | cut -d= -f2)
  export MICROSOFT_APP_TENANT_ID=$(grep '^MICROSOFT_APP_TENANT_ID=' "$ENV_FILE" | cut -d= -f2 || echo "")
fi

echo "✅ Bot ID: ${MICROSOFT_APP_ID:-not set}"
echo ""
echo "🚀 Starting QQIA Agent..."
echo "   Endpoint: https://${BOT_DOMAIN}/api/messages"
echo ""

npm start
