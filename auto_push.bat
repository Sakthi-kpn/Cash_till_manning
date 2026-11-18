@echo off
REM Daily auto-push for the Cash Till dashboard.
REM Commits any new file in data\ (and config\) and pushes to GitHub -> the Action rebuilds data.json -> same Pages link updates.
cd /d "%~dp0"
set "GIT=git"
where git >nul 2>nul
if errorlevel 1 (
  for /d %%D in ("%LOCALAPPDATA%\GitHubDesktop\app-*") do set "GIT=%%D\resources\app\git\cmd\git.exe"
)
echo ==== %date% %time% ====>> auto_push.log
"%GIT%" pull --rebase --autostash origin main >> auto_push.log 2>&1
"%GIT%" add data config >> auto_push.log 2>&1
"%GIT%" diff --cached --quiet
if not errorlevel 1 (
  echo No new data to push.>> auto_push.log
  exit /b 0
)
"%GIT%" commit -m "Daily data %date%" >> auto_push.log 2>&1
"%GIT%" push origin main >> auto_push.log 2>&1
if errorlevel 1 (echo PUSH FAILED - open GitHub Desktop and push manually.>> auto_push.log & exit /b 1)
echo Pushed.>> auto_push.log
