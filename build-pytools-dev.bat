@echo off
setlocal
call "%~dp0build-release.bat" %*
exit /b %ERRORLEVEL%
