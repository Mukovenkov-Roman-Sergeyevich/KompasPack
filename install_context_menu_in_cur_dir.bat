@echo off
SETLOCAL ENABLEDELAYEDEXPANSION

SET CONTEXT_MENU_NAME=^&Pack Kompas Project
SET SCRIPT_FILENAME=kompas_pack-n-go.py
SET FILE_EXTENSIONS=.a3d .cdw

SET "SCRIPT_DIR=%~dp0"
SET "PYTHON_SCRIPT_FULL_PATH=%SCRIPT_DIR%%SCRIPT_FILENAME%"

IF NOT EXIST "%PYTHON_SCRIPT_FULL_PATH%" (
    echo ОШИБКА: Не найден "%PYTHON_SCRIPT_FULL_PATH%".
    pause
    EXIT /B 1
)

python.exe --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo ПРЕДУПРЕЖДЕНИЕ: python возможно не установлен
    pause
)

SET "PYTHON_EXECUTABLE="
FOR /F "delims=" %%i IN ('where python.exe 2^>nul') DO (
    SET "PYTHON_EXECUTABLE=%%i"
    echo Нашли полный путь к python: !PYTHON_EXECUTABLE!
    GOTO PythonFound
)

IF NOT DEFINED PYTHON_EXECUTABLE (
    echo ОШИБКА: Не удалось автоматически найти python.exe в вашем PATH.
    EXIT /B 1
)

:PythonFound
echo.
echo Установливаем контекстное меню...
echo Расположение скрипта: %PYTHON_SCRIPT_FULL_PATH%
echo Python: %PYTHON_EXECUTABLE%
echo.

SET "TEMP_REG_FILE=%TEMP%\kompas_packer_temp.reg"

(
  echo Windows Registry Editor Version 5.00
  echo.
) > "%TEMP_REG_FILE%"

FOR %%E IN (%FILE_EXTENSIONS%) DO (
  (
    echo [HKEY_CLASSES_ROOT\SystemFileAssociations\%%E\shell\KompasPackAndGo]
    echo @="%CONTEXT_MENU_NAME%"
    echo "Icon"="\"%PYTHON_EXECUTABLE%\",0"
    echo.
    echo [HKEY_CLASSES_ROOT\SystemFileAssociations\%%E\shell\KompasPackAndGo\command]
    SET "COMMAND_PATH_PY_ESCAPED=%PYTHON_SCRIPT_FULL_PATH:\=\\%"
    SET "COMMAND_PYTHON_EXE_ESCAPED=%PYTHON_EXECUTABLE:\=\\%"
    echo @="\"!COMMAND_PYTHON_EXE_ESCAPED!\" \"!COMMAND_PATH_PY_ESCAPED!\" \"%%1\""
    echo.
  ) >> "%TEMP_REG_FILE%"
)

echo.
type "%TEMP_REG_FILE%"
echo.
echo Пытаемся применить изменения в регистре...

reg import "%TEMP_REG_FILE%"
IF %ERRORLEVEL% EQU 0 (
    echo.
    echo Контекстное меню успешно установлено
) ELSE (
    echo.
    echo ОШИБКА: Не удалось применить изменения в регистре.
    echo Попробуйте запустить в режиме администратора.
)

IF EXIST "%TEMP_REG_FILE%" del "%TEMP_REG_FILE%"

echo.
pause
EXIT /B 0