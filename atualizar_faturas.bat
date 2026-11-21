@echo off
chcp 65001 > nul
title Faturas Shadow_77
cd /d "%~dp0"
echo A iniciar Faturas Shadow_77...
echo ----------------------------------------------------
if exist "%~dp0dist\Faturas Shadow_77.exe" (
  "%~dp0dist\Faturas Shadow_77.exe"
) else (
  py "%~dp0scripts\faturas_shadow_77.py"
)
echo ----------------------------------------------------
echo Aplicacao terminada.
pause