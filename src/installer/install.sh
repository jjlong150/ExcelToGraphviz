#!/usr/bin/env bash
#
# Relationship Visualizer installer for macOS.
#
# Run this from inside the unzipped "Relationship Visualizer" folder:
#
#     cd path/to/Relationship\ Visualizer
#     bash install.sh
#
# Running it this way (from an already-open Terminal) avoids macOS
# Gatekeeper's "unidentified developer" warnings, since Gatekeeper only
# checks apps launched from Finder, not scripts run from a shell.
#
# What this does:
#   1. Shows the MIT license and asks you to accept it.
#   2. Finds where Graphviz's `dot` command is installed.
#   3. Updates ExcelToGraphviz.applescript to use that path.
#   4. Copies ExcelToGraphviz.applescript into the folder Microsoft Excel's
#      sandbox requires: ~/Library/Application Scripts/com.microsoft.Excel
#   5. Asks whether to keep the sample workbooks (they take extra disk
#      space) and removes any SQL-based samples, since macOS Excel has no
#      ActiveX/ADO support and those samples cannot run here.
#
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APPLESCRIPT_SRC="$SCRIPT_DIR/ExcelToGraphviz.applescript"
LICENSE_FILE="$SCRIPT_DIR/license.txt"
SANDBOX_DIR="$HOME/Library/Application Scripts/com.microsoft.Excel"
SANDBOX_FILE="$SANDBOX_DIR/ExcelToGraphviz.applescript"
SAMPLES_DIR="$SCRIPT_DIR/samples"

info()  { printf '%s\n' "$*"; }
error() { printf 'Error: %s\n' "$*" >&2; }

if [[ "$(uname -s)" != "Darwin" ]]; then
    error "This installer is for macOS only. On Windows, follow the instructions at https://exceltographviz.com/install-win/"
    exit 1
fi

if [[ ! -f "$APPLESCRIPT_SRC" ]]; then
    error "Cannot find ExcelToGraphviz.applescript next to this installer."
    error "Run install.sh from inside the unzipped 'Relationship Visualizer' folder."
    exit 1
fi

info "Relationship Visualizer Installer"
info "=================================="
info

# ---- 1. License ----
if [[ -f "$LICENSE_FILE" ]]; then
    cat "$LICENSE_FILE"
    info
else
    error "license.txt not found next to this installer; continuing without displaying it."
fi

read -r -p "Do you accept the terms of the above MIT License? [y/N]: " ACCEPT
case "$ACCEPT" in
    [yY]|[yY][eE][sS]) ;;
    *) info "License not accepted. Installation cancelled."; exit 0 ;;
esac
info

# ---- 2. Locate Graphviz's dot command ----
DOT_PATH=""
if command -v dot >/dev/null 2>&1; then
    DOT_PATH="$(command -v dot)"
else
    for candidate in /opt/homebrew/bin/dot /usr/local/bin/dot /opt/local/bin/dot; do
        if [[ -x "$candidate" ]]; then
            DOT_PATH="$candidate"
            break
        fi
    done
fi

if [[ -z "$DOT_PATH" ]]; then
    error "Could not find Graphviz's 'dot' command."
    info "Install Graphviz first, for example with Homebrew:"
    info "    brew install graphviz"
    info "    sudo dot -c"
    info "Then run this installer again."
    info "Full instructions: https://exceltographviz.com/install-mac/"
    exit 1
fi

DOT_VERSION="$("$DOT_PATH" -V 2>&1 || true)"
info "Found Graphviz: $DOT_PATH"
info "  $DOT_VERSION"
info

# ---- 3. Patch and install the AppleScript ----
TMP_SCRIPT="$(mktemp)"
trap 'rm -f "$TMP_SCRIPT"' EXIT

ESCAPED_DOT_PATH="$(printf '%s\n' "$DOT_PATH" | sed 's/[&\]/\\&/g')"
sed -E "s#do shell script \"[^\"]*dot \"#do shell script \"${ESCAPED_DOT_PATH} \"#" "$APPLESCRIPT_SRC" > "$TMP_SCRIPT"

if ! grep -qF "do shell script \"${DOT_PATH} \"" "$TMP_SCRIPT"; then
    error "Failed to update the dot path in ExcelToGraphviz.applescript."
    error "Please edit the file manually and change the path on the first line to: $DOT_PATH"
    exit 1
fi

cp "$TMP_SCRIPT" "$APPLESCRIPT_SRC"

mkdir -p "$SANDBOX_DIR"
cp "$APPLESCRIPT_SRC" "$SANDBOX_FILE"
info "Installed ExcelToGraphviz.applescript to:"
info "  $SANDBOX_FILE"
info

# ---- 4. Samples ----
if [[ -d "$SAMPLES_DIR" ]]; then
    read -r -p "Install the sample workbooks? They're useful for learning but take up extra disk space. [y/N]: " KEEP_SAMPLES
    case "$KEEP_SAMPLES" in
        [yY]|[yY][eE][sS])
            REMOVED_SQL=0
            for d in "$SAMPLES_DIR"/*/; do
                base="$(basename "$d")"
                case "$base" in
                    *"Using SQL"*)
                        rm -rf "$d"
                        REMOVED_SQL=1
                        ;;
                esac
            done
            if [[ "$REMOVED_SQL" -eq 1 ]]; then
                info "Removed the SQL-based samples: macOS Excel has no ActiveX/ADO support, so those samples cannot run here."
            fi
            info "Sample workbooks kept in: $SAMPLES_DIR"
            ;;
        *)
            rm -rf "$SAMPLES_DIR"
            info "Removed the sample workbooks to save disk space."
            ;;
    esac
    info
fi

# ---- 5. Done ----
info "Installation complete."
info
info "Next steps:"
info "  1. Open 'Relationship Visualizer.xlsm' in Excel."
info "  2. When prompted, enable macros and grant permissions."
info "  3. From the File menu, choose 'Save as Template...' to keep a clean copy for future use."
info
info "Full documentation: https://exceltographviz.com/install-mac/"
