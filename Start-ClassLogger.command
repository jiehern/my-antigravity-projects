#!/bin/bash
# Double-click this file to launch ClassLogger anytime
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

PORT=5173

# Check if server is already running on port 5173
if lsof -Pi :$PORT -sTCP:LISTEN -t >/dev/null 2>&1 ; then
    open "http://localhost:$PORT/"
else
    # Prefer npx vite if available, otherwise seamlessly use python3 http.server
    if command -v npx >/dev/null 2>&1 ; then
        nohup npx vite --port $PORT >/dev/null 2>&1 &
    elif command -v python3 >/dev/null 2>&1 ; then
        nohup python3 -m http.server $PORT >/dev/null 2>&1 &
    elif command -v python >/dev/null 2>&1 ; then
        nohup python -m http.server $PORT >/dev/null 2>&1 &
    fi
    sleep 1.2
    open "http://localhost:$PORT/"
fi

