#!/usr/bin/env bash

# Optimized HoanCauGroup 2026 Formatting Script
# Delegating to Python version for superior performance (One read, one batch write)

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
FILE=$1

if [ -z "$FILE" ]; then
    echo "Usage: $0 <file.docx>"
    exit 1
fi

python3 "$SCRIPT_DIR/format_hoancau_v2.py" "$FILE"
