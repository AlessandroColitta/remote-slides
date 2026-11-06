@echo off
title Remote Slides Local Server
echo Starting Remote Slides server at http://localhost:8000...
start "" "http://localhost:8000"
python -m http.server 8000
