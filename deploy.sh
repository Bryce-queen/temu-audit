#!/bin/bash
set -e

echo "=== Temu Listing AI Audit — One-Click Deploy ==="

# Check Docker
if ! command -v docker &>/dev/null; then
    echo "[ERROR] Docker not found. Install: https://docs.docker.com/get-docker/"
    exit 1
fi

cd "$(dirname "$0")"

# Copy .env if missing
if [ ! -f .env ]; then
    echo "[INFO] .env not found, creating with defaults..."
    cat > .env <<'EOF'
STRIPE_SECRET_KEY=sk_test_YOUR_KEY_HERE
STRIPE_PUBLISHABLE_KEY=pk_test_YOUR_KEY_HERE
STRIPE_WEBHOOK_SECRET=
ADMIN_TOKEN=change_me_to_a_long_random_value
SECRET_KEY=change_me_to_another_long_random_value
ZHIPU_API_KEY=
SILICONFLOW_API_KEY=
COZE_BOT_ID=
COZE_TOKEN=
AUDIT_PRICE_CENTS=499
AUDIT_PRICE_CURRENCY=usd
PUBLIC_DOMAIN=http://localhost:8080
EOF
    echo "[INFO] Edit .env with your real Stripe keys, then re-run."
    exit 0
fi

# Build & start
echo "[INFO] Building Docker image..."
docker compose build

echo "[INFO] Starting services..."
docker compose up -d

echo ""
echo "=== Deploy Complete ==="
echo "  App:  http://localhost:8080"
echo "  Audit: http://localhost:8080/audit"
echo "  Orders: http://localhost:8080/orders"
echo ""
echo "Commands:"
echo "  docker compose logs -f     # tail logs"
echo "  docker compose down        # stop"
echo "  docker compose up -d       # restart"
