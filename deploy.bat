@echo off
setlocal
cd /d "%~dp0"
title Deploy Hamdan Tour Check

set "REPO=https://github.com/kamilzhamdan-stafvisa/hamdantour.git"
set "SITE=https://kamilzhamdan-stafvisa.github.io/hamdantour/"
set "ACTIONS=https://github.com/kamilzhamdan-stafvisa/hamdantour/actions"

echo.
echo === Deploy Hamdan Tour Check ===
echo Folder: %CD%
echo.

where git >nul 2>&1
if errorlevel 1 goto nogit

if not exist ".git" git init -b main

set "GUN="
for /f "delims=" %%i in ('git config user.name') do set "GUN=%%i"
if not defined GUN git config user.name "Hamdan Tour"
set "GUE="
for /f "delims=" %%i in ('git config user.email') do set "GUE=%%i"
if not defined GUE git config user.email "hamdantour@users.noreply.github.com"

git remote get-url origin >nul 2>&1
if errorlevel 1 (git remote add origin %REPO%) else (git remote set-url origin %REPO%)

git branch -M main
git add -A
git add -f .github
git commit -m "Update Hamdan Tour Check %date% %time%"

echo.
echo Mengirim ke GitHub...
git push -u origin main
if not errorlevel 1 goto pushed

echo.
echo Push ditolak. Mencoba menggabungkan perubahan dari GitHub...
git pull origin main --rebase --allow-unrelated-histories
git push -u origin main
if errorlevel 1 goto failed

:pushed
where gh >nul 2>&1
if not errorlevel 1 gh api -X POST repos/kamilzhamdan-stafvisa/hamdantour/pages -f build_type=workflow >nul 2>&1

echo.
echo Selesai. Situs diperbarui otomatis dalam sekitar 1 menit:
echo %SITE%
echo.
start "" "%ACTIONS%"
timeout /t 3 >nul
start "" "%SITE%"
goto end

:nogit
echo Git belum terpasang. Unduh dan pasang dari https://git-scm.com/download/win
echo lalu jalankan file ini lagi.
goto end

:failed
echo.
echo Gagal mengirim. Salin pesan error di atas dan kirim ke Claude.

:end
echo.
pause
endlocal
