@echo off
REM Set the path to your Git repository folder
set REPO_PATH="C:\Users\amarr\OneDrive\Desktop\expertise3000\expertises2"

REM Go to the repository folder
cd /d "%REPO_PATH%"

git pull
pause