#!/bin/bash

# File Merger Docker Build Script
# Hybrid approach: tries xlwings, falls back to pandas/openpyxl

set -e  # Exit on any error

echo "🐳 Building File Merger Docker Image (Hybrid Approach)"
echo "======================================================"

# Build the Docker image
echo "📦 Building Docker image with hybrid xlwings support..."
echo "   - Tries to install xlwings (may fail in Linux containers)"
echo "   - Falls back to pandas/openpyxl if xlwings fails"
echo "   - Works on Windows/Mac (with xlwings) and Linux (pandas only)"
echo ""

docker build -t file-merger:latest .

echo ""
echo "✅ Docker image built successfully!"
echo ""

# Test the image
echo "🧪 Testing Docker image..."
docker run --rm file-merger:latest --help

echo ""
echo "🎉 Build and test completed successfully!"
echo ""
echo "📋 Next steps:"
echo "1. Test with your data: docker run -v \$(pwd)/Data:/app/data file-merger:latest /app/data"
echo "2. Push to Docker Hub: docker tag file-merger:latest your-username/file-merger:latest"
echo "3. Push: docker push your-username/file-merger:latest"
echo ""
echo "💡 Usage examples:"
echo "   # Basic usage"
echo "   docker run -v \$(pwd)/Data:/app/data file-merger:latest /app/data"
echo ""
echo "   # With verbose output"
echo "   docker run -v \$(pwd)/Data:/app/data file-merger:latest /app/data --verbose"
echo ""
echo "   # With column cleaning options"
echo "   docker run -v \$(pwd)/Data:/app/data file-merger:latest /app/data --lowercase-columns"
echo "   docker run -v \$(pwd)/Data:/app/data file-merger:latest /app/data --remove-special-chars"
echo ""
echo "🔍 Note: xlwings may not work in Linux containers, but pandas/openpyxl will handle most Excel files"
