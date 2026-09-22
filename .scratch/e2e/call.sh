#!/bin/bash
# e2e helper: call bridge tool, args JSON on argv[2] or stdin
TOOL="$1"; shift
ARGS="${1:-{}}"
cd "/Users/Python/Office Agent Bridge" || exit 1
python3 skills/office-agent-bridge/scripts/bridge_client.py "$TOOL" "$ARGS"
