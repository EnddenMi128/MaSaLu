@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo =========================
echo    GAS CLASP PUSH
echo =========================
echo.

call clasp push
if errorlevel 1 (
  echo.
  echo ❌ clasp push 失敗
  pause
  exit /b 1
)

call clasp deploy -i AKfycby6RKpYPoQGyV6dp6Nw1lwDaT634uddhro65nmorws19MmDx0kX_NQbee4ThQPcEAag
if errorlevel 1 (
  echo.
  echo ❌ clasp deploy 失敗
  pause
  exit /b 1
)

echo.
echo ✅ 完成！已 push 並 deploy
pause