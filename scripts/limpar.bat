@echo off
chcp 65001 >nul
echo ========================================================
echo A executar limpeza de manutencao do Faturas Shadow_77...
echo ========================================================
echo [1/2] A limpar caches de compilacao do Python...
for /d /r "%~dp0" %%d in (__pycache__) do (
    if exist "%%d" (
        rd /s /q "%%d"
        echo    -> Removido: %%d
    )
)
set "DATA_DIR=%LOCALAPPDATA%\Faturas Shadow_77"
echo [2/2] A limpar ficheiros temporarios de entrada...
del /q /f "%DATA_DIR%\entrada\*.tmp" 2>nul
del /q /f "%DATA_DIR%\entrada\~*" 2>nul
del /q /f "%DATA_DIR%\whatsapp\*.tmp" 2>nul
del /q /f "%DATA_DIR%\whatsapp\~*" 2>nul
echo ========================================================
echo Limpeza concluida com sucesso, shadow!
echo ========================================================
pause