@echo off
cd /d "%~dp0"

if not exist node_modules (
  echo Instalando dependencias...
  call npm install
)

echo Generando EXE portable...
call npm run electron:portable

echo.
echo Listo. Revisa:
echo dist-electron\FotoCarnet-Portable-1.0.0.exe
pause
