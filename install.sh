#!/bin/bash
# One-line installer for Excel MCP Server (Linux/macOS)
# Usage: bash -c "$(curl -fsSL https://raw.githubusercontent.com/arv-fazriansyah/excel-mcp-server-code/main/install.sh)"

set -e

REPO_URL="https://github.com/arv-fazriansyah/excel-mcp-server-code.git"
INSTALL_DIR="$HOME/.local/share/ExcelMCPServer"

echo "🚀 Installing Excel MCP Server..."

# 1. Clone or update repo
if [ -d "$INSTALL_DIR" ]; then
    echo "📁 Updating existing installation..."
    git -C "$INSTALL_DIR" pull
else
    echo "📥 Cloning repository..."
    git clone "$REPO_URL" "$INSTALL_DIR"
fi

# 2. Install dependencies
echo "📦 Installing Python dependencies..."
pip install -r "$INSTALL_DIR/requirements.txt"

# 3. Create working directory
DOC_DIR="$HOME/Documents/excel_files"
if [ ! -d "$DOC_DIR" ]; then
    mkdir -p "$DOC_DIR"
    echo "📂 Created working directory: $DOC_DIR"
fi

# 4. Show MCP configuration
SERVER_PATH="$INSTALL_DIR/excel_mcp_server.py"
echo ""
echo "✅ Installation complete!"
echo ""
echo "📋 Add this to your MCP client config (Claude Desktop, Cursor, etc.):"
cat <<EOF
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["$SERVER_PATH"]
    }
  }
}
EOF

echo "📁 Your Excel files go in: $DOC_DIR"