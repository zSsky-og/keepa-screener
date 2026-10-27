#!/bin/bash
# Keepa screener - Mac launcher. Double-click this file to start.
# (The first time, macOS may ask: right-click the file, choose Open, then Open again.)

cd "$(dirname "$0")" || exit 1
clear
echo ""
echo "  Keepa screener"
echo "  =============="
echo ""

finish() {
    echo ""
    read -r -p "  Press Enter to close this window. " _
    exit "${1:-1}"
}

# ---- 1. Find Python 3.9 or newer --------------------------------------------
PY=""
for candidate in python3 /opt/homebrew/bin/python3 /usr/local/bin/python3 \
                 /Library/Frameworks/Python.framework/Versions/Current/bin/python3 python; do
    if command -v "$candidate" >/dev/null 2>&1 &&
       "$candidate" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)' >/dev/null 2>&1; then
        PY="$candidate"
        break
    fi
done
if [ -z "$PY" ]; then
    echo "  Python 3.9 or newer is not installed on this Mac."
    echo ""
    echo "  1. The Python download page will open now. Download the macOS installer and run it."
    echo "  2. When it has finished, double-click start.command again."
    open "https://www.python.org/downloads/macos/" 2>/dev/null
    finish 1
fi

# ---- 2. Private environment so nothing else on the Mac is touched ------------
if [ ! -x ".venv/bin/python" ]; then
    echo "  First start: setting things up. This takes a minute or two, only once."
    echo ""
    if ! "$PY" -m venv .venv; then
        echo "  Could not create the Python environment in this folder."
        echo "  Move the folder somewhere you can write to, such as Documents, and try again."
        finish 1
    fi
fi
VPY=".venv/bin/python"

# ---- 3. Install or update the packages when requirements.txt changes ---------
if ! cmp -s requirements.txt .venv/installed-requirements.txt; then
    echo "  Installing the packages the screener needs..."
    "$VPY" -m pip install --disable-pip-version-check -q --upgrade pip >/dev/null 2>&1
    if ! "$VPY" -m pip install --disable-pip-version-check -q -r requirements.txt; then
        echo ""
        echo "  Installing the packages failed. Check the internet connection and try again."
        finish 1
    fi
    cp requirements.txt .venv/installed-requirements.txt
    echo "  Done."
    echo ""
fi

# ---- 4. Start the app ----------------------------------------------------------
"$VPY" app.py
status=$?
if [ $status -ne 0 ] && [ $status -ne 130 ]; then
    echo ""
    echo "  The screener stopped with an error, see the message above."
    finish $status
fi
