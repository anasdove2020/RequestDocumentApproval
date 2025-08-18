#!/bin/bash

# Script to create a clean zip archive of the SPFx project
# This respects .gitignore and excludes node_modules, build artifacts, etc.
# without actually deleting them from your working directory

PROJECT_NAME="RequestDocumentApproval"
TIMESTAMP=$(date +"%Y%m%d_%H%M%S")
ZIP_NAME="${PROJECT_NAME}_clean_${TIMESTAMP}.zip"

echo "🚀 Creating clean project archive..."
echo "📦 Archive name: ${ZIP_NAME}"

# Use git archive to create a clean zip that respects .gitignore
# This only includes tracked files and excludes everything in .gitignore
git archive --format=zip --output="${ZIP_NAME}" HEAD

if [ $? -eq 0 ]; then
    echo "✅ Clean archive created successfully!"
    echo "📁 Location: $(pwd)/${ZIP_NAME}"
    echo "📊 Archive size: $(du -h "${ZIP_NAME}" | cut -f1)"
    echo ""
    echo "🎯 This archive contains:"
    echo "   ✅ All source code"
    echo "   ✅ Configuration files"
    echo "   ✅ Package.json dependencies"
    echo "   ❌ node_modules (excluded)"
    echo "   ❌ Build artifacts (excluded)"
    echo "   ❌ Temp files (excluded)"
    echo ""
    echo "📧 Ready to zip!"
else
    echo "❌ Failed to create archive"
    exit 1
fi
