#!/bin/bash
# Codespace startup script for QQIA Agent (Outgoing Webhook mode)
# No Azure or Teams Toolkit required — just start the bot.

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

# ---- Step 2: Set port to public ----
# Required for Teams Outgoing Webhook to POST messages to this endpoint.
# HMAC-SHA256 validation ensures only requests signed with the webhook secret are accepted.
if [ -n "$CODESPACE_NAME" ]; then
  gh codespace ports visibility 3978:public -c "$CODESPACE_NAME" 2>/dev/null || true
  echo "✅ Port 3978 set to public (secured by HMAC-SHA256 validation)"
fi

# ---- Step 3: Start the bot ----
echo ""
echo "🚀 Starting QQIA Agent..."
echo ""
echo "   Webhook endpoint: https://${BOT_DOMAIN}/api/webhook"
echo "   Health check:     https://${BOT_DOMAIN}/api/health"
echo ""
echo "📱 To connect to Teams:"
echo "   1. Go to your Teams channel → ⋯ → Manage channel → Connectors"
echo "   2. Search 'Outgoing Webhook' → Configure"
echo "   3. Name: QQIA Agent"
echo "   4. Callback URL: https://${BOT_DOMAIN}/api/webhook"
echo "   5. Save the HMAC secret — set it as WEBHOOK_HMAC_SECRET env var for validation"
echo ""

npm start
