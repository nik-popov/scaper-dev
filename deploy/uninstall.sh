#!/bin/bash
set -euo pipefail

# ============================================================
# Super Scraper - Systemd Deployment Uninstaller
# Usage: sudo ./uninstall.sh
# ============================================================

if [ "$EUID" -ne 0 ]; then
    echo "ERROR: Please run as root (sudo ./uninstall.sh)"
    exit 1
fi

echo "==> Stopping and disabling services..."
systemctl stop super-scraper-update.timer 2>/dev/null || true
systemctl disable super-scraper-update.timer 2>/dev/null || true
systemctl stop super-scraper-consumer.service 2>/dev/null || true
systemctl disable super-scraper-consumer.service 2>/dev/null || true
systemctl stop super-scraper.service 2>/dev/null || true
systemctl disable super-scraper.service 2>/dev/null || true

echo "==> Removing systemd unit files..."
rm -f /etc/systemd/system/super-scraper.service
rm -f /etc/systemd/system/super-scraper-consumer.service
rm -f /etc/systemd/system/super-scraper-update.service
rm -f /etc/systemd/system/super-scraper-update.timer

systemctl daemon-reload

echo "==> Stopping and removing containers..."
docker stop super-scraper 2>/dev/null || true
docker rm super-scraper 2>/dev/null || true
docker stop super-scraper-consumer 2>/dev/null || true
docker rm super-scraper-consumer 2>/dev/null || true

echo ""
echo "Super Scraper uninstalled."
echo "Note: /etc/super-scraper/env was preserved. Remove manually if desired."

