@echo off
chcp 65001 >nul
color 0a

echo %USERNAME% | findstr /r "^[a-zA-Z0-9-_]*$" 1>nul 2>&1
if %errorlevel% equ 0 (
	color 0E
    echo Имя пользователя Windows - %USERNAME% - содержит недопустимые символы. Замените его в пути при установке.
    echo.
)
echo %COMPUTERNAME% | findstr /r "^[a-zA-Z0-9-_]*$" 1>nul 2>&1
if %errorlevel% equ 0 (
	color 0E
    echo Название ПК - %COMPUTERNAME% - содержит недопустимые символы. Необходимо переименование до установки.
    echo.
)
 
if exist E:\app (
	color 0E
	echo E:\app\ существует.  
	echo.
)
if exist D:\app (
	color 0E
	echo D:\app\ существует.
	echo.
)
if exist C:\app (
	color 0E
	echo C:\app\ существует.
	echo.
)


reg query "HKLM\SOFTWARE\WOW6432Node\ORACLE" 2>nul
if %errorlevel% equ 0 (
	color 0E
	echo.
	echo В реестре обнаружен раздел Oracle
	echo.
)

if defined ORACLE_HOME (
	color 0E
	echo Переменная среды ORACLE_HOME определена.
	echo.
)

if defined ORACLE_BASE (
	color 0E
	echo Переменная среды ORACLE_BASE определена.
	echo.
)

if defined NLS_LANG (
	color 0E
	echo Переменная среды NLS_LANG определена.
	echo.
)

echo "%PATH%" | findstr :\\app\\ 2>nul
if %errorlevel% equ 0 (
	color 0E
	echo.
	echo В PATH обнаружен каталог app
	echo.
)

echo Конец предварительной проверки
pause >nul
