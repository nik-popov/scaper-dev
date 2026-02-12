#!/bin/bash
set -e

echo "🔨 Rebuilding Docker image with debug code..."
docker build -t scaper-dev:gen-file-resizepic .

echo ""
echo "🔄 Restarting container..."
docker restart scaper-prod

echo ""
echo "✓ Deployment complete!"
echo ""
echo "📋 Now test with:"
echo "curl -X POST 'http://localhost:8080/generate-download-file/?file_id=1083&row_offset=0'"
echo ""
echo "📊 Monitor logs with:"
echo "docker exec scaper-prod sh -c 'tail -f jobs/1083/*/logs/*.log'"
