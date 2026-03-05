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
IMAGE="nikiconluxury/scaper-dev:latest"

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

# Create env file directory if it doesn't exist
if [ ! -f /etc/super-scraper/env ]; then
    echo "Creating environment file at /etc/super-scraper/env"
    mkdir -p /etc/super-scraper
    cat > /etc/super-scraper/env << 'EOF'
# Add your environment variables here, one per line:
# DB_PASSWORD=your_password
# GOOGLE_API_KEY=your_key
EOF
    chmod 600 /etc/super-scraper/env
    echo ""
    echo "WARNING: Edit /etc/super-scraper/env with your environment variables before starting."
    echo ""
fi

echo "==> Installing systemd units..."

# Copy unit files to systemd directory
cp "${SCRIPT_DIR}/super-scraper.service" /etc/systemd/system/
cp "${SCRIPT_DIR}/super-scraper-consumer.service" /etc/systemd/system/
cp "${SCRIPT_DIR}/super-scraper-update.service" /etc/systemd/system/
cp "${SCRIPT_DIR}/super-scraper-update.timer" /etc/systemd/system/

# Reload systemd to pick up new units
systemctl daemon-reload

echo "==> Pulling latest image..."
docker pull "${IMAGE}"

echo "==> Enabling and starting super-scraper service..."
systemctl enable super-scraper.service
systemctl start super-scraper.service

echo "==> Enabling and starting RabbitMQ consumer service..."
systemctl enable super-scraper-consumer.service
systemctl start super-scraper-consumer.service

echo "==> Enabling and starting update timer (checks every 5 min)..."
systemctl enable super-scraper-update.timer
systemctl start super-scraper-update.timer

echo ""
echo "============================================================"
echo " Super Scraper deployed successfully!"
echo "============================================================"
echo ""
echo " Useful commands:"
echo "   systemctl status super-scraper          # Check service status"
echo "   journalctl -u super-scraper -f          # Follow service logs"
echo "   docker logs -f super-scraper            # Follow container logs"
echo "   systemctl list-timers                   # Check timer status"
echo "   journalctl -u super-scraper-update      # Update check logs"
echo "   systemctl restart super-scraper         # Manual restart"
echo ""
echo " To uninstall:"
echo "   sudo ./uninstall.sh"
echo ""

