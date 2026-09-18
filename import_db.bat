@echo off
chcp 65001 >nul
setlocal
set "PYTHONUTF8=1"
set "PROJECT=C:\Users\czy\Desktop\github-repositories\Tainted-Cain"
cd /d "%PROJECT%" || goto :no_project

rem Password source, use either one:
rem   1) environment variable ISAAC_DB_PASSWORD  -> set once: setx ISAAC_DB_PASSWORD "your_password"
rem   2) a file named db_password.txt in this folder with the password on one line
if not "%ISAAC_DB_PASSWORD%"=="" goto :has_pwd
if not exist "%PROJECT%\db_password.txt" goto :no_pwd
set /p ISAAC_DB_PASSWORD=<"%PROJECT%\db_password.txt"
goto :has_pwd

:no_project
echo [ERROR] Project folder not found: %PROJECT%
echo.
pause
exit /b 1

:no_pwd
echo [ERROR] Database password not found.
echo.
echo   Option 1 - run once in PowerShell or cmd:
echo       setx ISAAC_DB_PASSWORD "your_db_password"
echo   Option 2 - create db_password.txt in this folder,
echo              containing the password on a single line
echo.
pause
exit /b 1

:has_pwd
if exist "%PROJECT%\.venv\Scripts\python.exe" goto :venv_py
uv run import_to_mysql.py
goto :done

:venv_py
"%PROJECT%\.venv\Scripts\python.exe" "%PROJECT%\import_to_mysql.py"

:done
echo.
echo ---- finished, press any key to close ----
pause >nul
