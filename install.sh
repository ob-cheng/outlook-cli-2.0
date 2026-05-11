#!/bin/bash
set -e

REPO_URL="https://github.com/ob-cheng/outlook-cli-2.0.git"
SKILL_URL="https://raw.githubusercontent.com/ob-cheng/outlook-cli-2.0/main/SKILL.md"
INSTALL_DIR="$HOME/.local/share/outlook-cli"

echo "Installing Outlook CLI 2.0..."

# Step 1: Check Python
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is required but not installed."
    exit 1
fi

PYTHON_VERSION=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
echo "Found Python $PYTHON_VERSION"

# Step 2: Clone or update repo
if [ -d "$INSTALL_DIR" ]; then
    echo "Updating existing installation..."
    cd "$INSTALL_DIR"
    git pull --quiet
else
    echo "Cloning repository..."
    git clone --quiet "$REPO_URL" "$INSTALL_DIR"
    cd "$INSTALL_DIR"
fi

# Step 3: Install dependencies
echo "Installing dependencies..."
pip3 install -q -r requirements.txt

# Step 4: Create launcher script
LAUNCHER="$HOME/.local/bin/outlook-cli"
mkdir -p "$HOME/.local/bin"

cat > "$LAUNCHER" << 'EOF'
#!/bin/bash
INSTALL_DIR="$HOME/.local/share/outlook-cli"
python3 "$INSTALL_DIR/outlook.py" "$@"
EOF

chmod +x "$LAUNCHER"

# Add to PATH if needed
case ":$PATH:" in
    *":$HOME/.local/bin:"*) ;;
    *)
        SHELL_RC="$HOME/.bashrc"
        [ -n "$ZSH_VERSION" ] && SHELL_RC="$HOME/.zshrc"
        [ "$(uname -s)" = "Darwin" ] && SHELL_RC="$HOME/.zshrc"

        if ! grep -qF '$HOME/.local/bin' "$SHELL_RC" 2>/dev/null; then
            echo 'export PATH="$HOME/.local/bin:$PATH"' >> "$SHELL_RC"
            echo "Added ~/.local/bin to PATH in $SHELL_RC"
        fi
        ;;
esac

# Step 5: Install AI agent skills
SKILL_MARKER="$INSTALL_DIR/.skills-installed"
if [ ! -f "$SKILL_MARKER" ]; then
    echo "Installing AI agent skills..."

    AGENTS=(
        "$HOME/.claude:Claude Code"
        "$HOME/.cursor:Cursor"
        "$HOME/.windsurf:Windsurf"
        "$HOME/.copilot:GitHub Copilot"
        "$HOME/.hermes:Hermes Agent"
        "$HOME/.openclaw:OpenClaw"
    )

    for agent_entry in "${AGENTS[@]}"; do
        dir="${agent_entry%%:*}"
        name="${agent_entry##*:}"

        if [ -d "$dir" ]; then
            echo "  Found $name"
            mkdir -p "$dir/skills/outlook-cli"
            cp "$INSTALL_DIR/SKILL.md" "$dir/skills/outlook-cli/SKILL.md"
            echo "    Installed: $dir/skills/outlook-cli/SKILL.md"
        fi
    done

    touch "$SKILL_MARKER"
fi

echo ""
echo "Outlook CLI 2.0 installed successfully!"
echo ""
echo "Usage:"
echo "  outlook-cli search --unread"
echo "  outlook-cli send --to user@example.com --subject 'Hello' --body 'Hi there'"
echo "  outlook-cli cal list"
echo ""
echo "Note: This tool requires Windows with Outlook desktop installed."
echo "Run 'outlook-cli --help' for more commands."
