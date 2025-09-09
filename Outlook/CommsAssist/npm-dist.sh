#!/bin/bash
# A robust shell script to inline CSS and JS into a single HTML file.
# This script uses a here document to handle multiline content safely.

# Exit immediately if a command exits with a non-zero status.
set -e

# Define source and destination paths
SRC_DIR="src/outlook"
RESOURCES_DIR="$SRC_DIR/resources"
DIST_DIR="dist"
DIST_PUBLIC_DIR="$DIST_DIR/public"

echo "Starting build process..."

# Step 1: Clean and create the distribution directories
echo "Cleaning old dist directory..."
rm -rf "$DIST_DIR"

echo "Creating new dist directories: $DIST_DIR and $DIST_PUBLIC_DIR"
mkdir -p "$DIST_DIR"
mkdir -p "$DIST_PUBLIC_DIR"
mkdir -p "$DIST_PUBLIC_DIR/config"

# Step 2: Read the content of the CSS and JS files
echo "Reading source files..."
css_content=$(<"$RESOURCES_DIR/pane.css")
js_content=$(<"$SRC_DIR/pane.js")

# Step 3: Perform the inlining using a here document and redirect to the final file
echo "Inlining CSS and JS into pane.html..."

cat "$RESOURCES_DIR/pane.html" | while read -r line; do
    if [[ "$line" =~ "<link rel=\"stylesheet\" href=\"./pane.css\" />" ]]; then
        echo "<style>"
        echo "$css_content"
        echo "</style>"
    elif [[ "$line" =~ "<script src=\"./pane.js\"></script>" ]]; then
        echo "<script>"
        echo "$js_content"
        echo "</script>"
    else
        echo "$line"
    fi
done > "$DIST_PUBLIC_DIR/pane.html"

# Step 4: Copy remaining resource files to the correct locations
echo "Copying manifest.xml and other resources..."
cp "$RESOURCES_DIR/manifest.xml" "$DIST_DIR/manifest.xml"
cp $RESOURCES_DIR/*.png "$DIST_PUBLIC_DIR/"
cp $RESOURCES_DIR/config.json "$DIST_PUBLIC_DIR/config/"

echo "Build process completed successfully! Final files are in the 'dist' folder."
