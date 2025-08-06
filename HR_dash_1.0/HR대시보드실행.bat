@echo off
chcp 65001 > nul
title HR 대시보드

echo ╔══════════════════════════════════════════╗
echo ║        HR 대시보드 실행 프로그램         ║
echo ╚══════════════════════════════════════════╝
echo.

:: Python 실행
python run_dashboard.py

:: Python이 없을 경우
if errorlevel 1 (
    echo.
    echo ❌ Python이 설치되어 있지 않습니다!
    echo.
    echo Python을 설치하려면:
    echo 1. https://www.python.org 방문
    echo 2. Download Python 클릭
    echo 3. 설치 시 "Add Python to PATH" 체크
    echo.
    pause
)