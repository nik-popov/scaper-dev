#!/bin/bash
set -euo pipefail

# ============================================================
# Super Scraper - Systemd Deployment Uninstaller
# Usage: sudo ./uninstall.sh
# ============================================================

CONFIG_DIR="/etc/super-scraper"
SYSTEMD_DIR="/etc/systemd/system"
COMPOSE_FILE="${CONFIG_DIR}/docker-compose.yml"

if [ "$EUID" -ne 0 ]; then
    echo "ERROR: Please run as root (sudo ./uninstall.sh)"
    exit 1
fi

echo "==> Stopping and disabling services..."
systemctl disable --now super-scraper-update.timer 2>/dev/null || true
systemctl stop super-scraper-update.service 2>/dev/null || true
systemctl disable --now super-scraper-consumer.service 2>/dev/null || true
systemctl disable --now super-scraper.service 2>/dev/null || true

echo "==> Stopping compose stack and cleaning containers..."
if command -v docker &> /dev/null && [ -f "${COMPOSE_FILE}" ]; then
    docker compose -f "${COMPOSE_FILE}" down --remove-orphans --volumes 2>/dev/null || true
fi

echo "==> Removing systemd unit files..."
rm -f "${SYSTEMD_DIR}/super-scraper.service"
rm -f "${SYSTEMD_DIR}/super-scraper-consumer.service"
rm -f "${SYSTEMD_DIR}/super-scraper-update.service"
rm -f "${SYSTEMD_DIR}/super-scraper-update.timer"

systemctl daemon-reload

echo "==> Removing app containers and images..."
if command -v docker &> /dev/null; then
    docker rm -f super-scraper 2>/dev/null || true
    docker rm -f super-scraper-consumer 2>/dev/null || true
    docker image rm -f \
        nikiconluxury/scaper-dev:latest \
        nikiconluxury/scaper-dev-api:latest \
        nikiconluxury/scaper-dev-consumer:latest 2>/dev/null || true
fi

echo "==> Removing deployed compose file..."
rm -f "${COMPOSE_FILE}"

echo ""
echo "Super Scraper uninstalled."
echo "Note: ${CONFIG_DIR}/env was preserved. Remove manually if desired."


