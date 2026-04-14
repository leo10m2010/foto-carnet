@echo off
cd /d "%~dp0"

if not exist node_modules (
  echo Instalando dependencias...
  call npm install
)

echo Iniciando app de escritorio...
call npm run electron:start
