#!/bin/bash

# File Merger Docker Build Script

set -e  # Exit on any error

echo "🐳 Building File Merger Docker Image"
echo "====================================="

# Build the Docker image
echo "📦 Building Docker image..."
docker build -t file-merger:latest .

echo "✅ Docker image built successfully!"
echo ""

# Test the image
echo "🧪 Testing Docker image..."
docker run --rm file-merger:latest --help

echo ""
echo "🎉 Build and test completed successfully!"
echo ""
echo "📋 Next steps:"
echo "1. Test with your data: docker run -v \$(pwd)/Data:/app/data file-merger:latest"
echo "2. Push to Docker Hub: docker tag file-merger:latest your-username/file-merger:latest"
echo "3. Push: docker push your-username/file-merger:latest"
echo ""
echo "💡 Usage examples:"
echo "   docker run -v \$(pwd)/Data:/app/data file-merger:latest"
echo "   docker run -v \$(pwd)/Data:/app/data file-merger:latest --verbose"
echo "   docker run -v \$(pwd)/Data:/app/data file-merger:latest --lowercase-columns"
