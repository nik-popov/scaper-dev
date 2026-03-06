#!/bin/bash
set -euo pipefail

# ============================================================
# Super Scraper - Systemd Deployment Installer
# 
# Usage:
#   sudo ./install.sh
#
# Prerequisites:
#   - Docker installed and running
#   - Logged into Docker Hub:
#     docker login -u nikiconluxury
# ============================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CONFIG_DIR="/etc/super-scraper"
SYSTEMD_DIR="/etc/systemd/system"
COMPOSE_FILE="${CONFIG_DIR}/docker-compose.yml"

# Check running as root
if [ "$EUID" -ne 0 ]; then
    echo "ERROR: Please run as root (sudo ./install.sh)"
    exit 1
fi

# Check Docker is available
if ! command -v docker &> /dev/null; then
    echo "ERROR: Docker is not installed. Install Docker first."
    exit 1
fi

echo "==> Preparing ${CONFIG_DIR}..."
mkdir -p "${CONFIG_DIR}"

echo "==> Installing compose and systemd files..."
install -m 0644 "${SCRIPT_DIR}/docker-compose.yml" "${COMPOSE_FILE}"
install -m 0644 "${SCRIPT_DIR}/super-scraper.service" "${SYSTEMD_DIR}/super-scraper.service"
install -m 0644 "${SCRIPT_DIR}/super-scraper-update.service" "${SYSTEMD_DIR}/super-scraper-update.service"
install -m 0644 "${SCRIPT_DIR}/super-scraper-update.timer" "${SYSTEMD_DIR}/super-scraper-update.timer"

# Remove stale standalone consumer unit from older deployments
rm -f "${SYSTEMD_DIR}/super-scraper-consumer.service"

# Create env file directory if it doesn't exist
if [ ! -f "${CONFIG_DIR}/env" ]; then
    echo "Creating environment file at ${CONFIG_DIR}/env"
    cat > "${CONFIG_DIR}/env" << 'EOF'
# Add your environment variables here, one per line:
# DB_PASSWORD=your_password
# GOOGLE_API_KEY=your_key
EOF
    chmod 600 "${CONFIG_DIR}/env"
    echo ""
    echo "WARNING: Edit ${CONFIG_DIR}/env with your environment variables before starting."
    echo ""
fi

# Reload systemd to pick up new units
systemctl daemon-reload

echo "==> Enabling services..."
systemctl enable super-scraper.service
systemctl enable super-scraper-update.timer

echo "==> Restarting Super Scraper compose stack..."
systemctl restart super-scraper.service

echo "==> Starting update timer (checks every 5 min)..."
systemctl restart super-scraper-update.timer

echo "==> Current compose status..."
docker compose -f "${COMPOSE_FILE}" ps

echo ""
echo "============================================================"
echo " Super Scraper deployed successfully!"
echo "============================================================"
echo ""
echo " Useful commands:"
echo "   systemctl status super-scraper          # Check service status"
echo "   journalctl -u super-scraper -f          # Follow service logs"
echo "   docker logs -f super-scraper            # Follow container logs"
echo "   docker compose -f ${COMPOSE_FILE} ps    # Check compose status"
echo "   systemctl list-timers                   # Check timer status"
echo "   journalctl -u super-scraper-update      # Update check logs"
echo "   systemctl restart super-scraper         # Manual restart"
echo ""
echo " To uninstall:"
echo "   sudo ./uninstall.sh"
echo " To reinstall after pulling new source:"
echo "   sudo ./reinstall.sh"
echo ""

