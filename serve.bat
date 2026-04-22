@echo off
echo Starting local server at http://localhost:8080
echo Press Ctrl+C to stop.
start http://localhost:8080/index.html
"C:\Users\Daddy\Anaconda3\python.exe" -m http.server 8080
