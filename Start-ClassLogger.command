#!/bin/bash
# Double-click this file to launch ClassLogger anytime
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

# Check if Vite is already running on port 5173
if lsof -Pi :5173 -sTCP:LISTEN -t >/dev/null ; then
    open "http://localhost:5173/"
else
    # Start dev server and open in default browser
    npx vite --port 5173 &
    sleep 1.5
    open "http://localhost:5173/"
fi
