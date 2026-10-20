@echo off
rem Double-click to start, or drag today's files (raw data, trackers, schedules, OT reports / .zip, the workbook) onto this file.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0DailyOps.ps1" %*
if errorlevel 1 pause