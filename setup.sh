#!/bin/bash
echo "Setting up Excel MCP Server..."
pip install -r requirements.txt

DOC_DIR="$HOME/Documents/excel_files"
if [ ! -d "$DOC_DIR" ]; then
    mkdir -p "$DOC_DIR"
    echo "Created working directory at: $DOC_DIR"
fi

echo "Setup complete!"