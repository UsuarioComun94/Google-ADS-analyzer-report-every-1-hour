@echo off
set "NODE_DIR=C:\Program Files\nodejs"
if exist "%NODE_DIR%" set "PATH=%NODE_DIR%;%PATH%"

cd /d "%~dp0hma-desktop"

if not exist "node_modules" (
  npm install
)

npm run dev

