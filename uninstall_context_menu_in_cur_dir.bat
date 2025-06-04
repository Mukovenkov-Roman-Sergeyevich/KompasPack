@echo off
SETLOCAL

SET FILE_EXTENSIONS=.a3d .cdw

SET "TEMP_REG_FILE=%TEMP%\kompas_packer_uninstall_temp.reg"

(
  echo Windows Registry Editor Version 5.00
  echo.
) > "%TEMP_REG_FILE%"

FOR %%E IN (%FILE_EXTENSIONS%) DO (
  (
    echo [-HKEY_CLASSES_ROOT\SystemFileAssociations\%%E\shell\KompasPackAndGo]
    echo.
  ) >> "%TEMP_REG_FILE%"
)

reg import "%TEMP_REG_FILE%"
IF %ERRORLEVEL% EQU 0 (
    echo Успешно удалено контекстное меню
) ELSE (
    echo ОШИБКА: Не удалось удалить или контекстное меню не установлено.
    echo Запустите с помощью администратора.
)

IF EXIST "%TEMP_REG_FILE%" del "%TEMP_REG_FILE%"

echo.
pause
EXIT /B 0