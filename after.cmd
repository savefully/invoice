@echo off
chcp 65001 >nul
cd %~dp0

net session 1>nul 2>&1
if %errorlevel% neq 0 (
    echo Ошибка: cкрипт запущен без прав администратора
    pause >nul
    exit
)

set diskpart=noapp
if exist E:\app (set diskpart=E)
if exist D:\app (set diskpart=D)
if exist C:\app (set diskpart=C)
if %diskpart%==noapp (
    color 0c
	echo Ошибка. Папка app не обнаружена в возможных расположениях. Вероятно setup еще не запускался.
	pause >nul
	exit
)

echo App: %diskpart%:\app

call :safely copy /y ^"Invoice_GUI_48_2_CC\INV_Clients_%diskpart%.lnk^" ^"C:\Users\Public\Desktop\INV_Clients_2.lnk^" 1^>nul 2^>^&1
echo Ярлык размещен.
call :safely move /y ^"Invoice_GUI_48_2_CC^" ^"%diskpart%:\^" 1^>nul 2^>^&1
echo Новый клиент перенесен.
call :safely move /y ^"GUI_Invoice^" ^"%diskpart%:\^" 1^>nul 2^>^&1
echo Старый клиент перенесен.
echo.

echo После установки setup нажмите любую клавишу дважды для продолжения.
pause >nul
pause >nul

echo Ниже представлены разделы KEY_OraClient12Home*_32bit в реестре.
echo Введите число актуального раздела (только число, которое должно быть вместо звездочки)
echo К примеру 1, чтобы выбрать KEY_OraClient12Home1_32bit
echo Как правило актуальный раздел с наибольшим числом из доступных.
echo.

call :safely reg query ^"HKLM\SOFTWARE\WOW6432Node\ORACLE^" 2^>nul

echo.
set oracle_home_key_number=1
set /p oracle_home_key_number=Введите число:
set absolute_oracle_home_key=HKLM\SOFTWARE\WOW6432Node\ORACLE\KEY_OraClient12Home%oracle_home_key_number%_32bit
echo Выбран: %absolute_oracle_home_key%

call :safely reg add ^"%absolute_oracle_home_key%^" /v NLS_LANG /t REG_SZ /d ^"American_America.CL8MSWIN1251^" /f ^>nul
echo NLS_LANG изменен.

set oracle_home_folder=
for /f "skip=2 tokens=2,*" %%a in ('reg query "%absolute_oracle_home_key%" /v "ORACLE_HOME" 2^>nul') do (
    set oracle_home_folder=%%b
)

if defined oracle_home_folder (
    call :safely copy /y ^"tnsnames.ora^" ^"%oracle_home_folder%\network\admin\tnsnames.ora^" 1^>nul 2^>^&1
    echo tnsnames.ora размещен.
) else (
    color 0c
    echo ORACLE_HOME не обнаружен в %absolute_oracle_home_key%
    pause >nul
    exit
)

set oracle_base_folder=
for /f "skip=2 tokens=2,*" %%a in ('reg query "%absolute_oracle_home_key%" /v "ORACLE_BASE" 2^>nul') do (
    set oracle_base_folder=%%b
)

if defined oracle_base_folder (
    icacls "%oracle_base_folder%" /grant Everyone:F /t /c 1>nul 2>&1
    if %errorlevel% neq 0 (
        call :safely icacls ^"%oracle_base_folder%^" /grant Все:F /t /c 1^>nul 2^>^&1          
    )
    echo Группе "Все" предоставлен полный доступ к %oracle_base_folder% 
) else (
    color 0c
    echo ORACLE_BASE не обнаружен в %absolute_oracle_home_key%
    pause >nul
    exit
)

color 0a
echo.
echo.
echo Успешно
pause >nul
goto :eof

:safely
%*
if %errorlevel% neq 0 (
	color 0c
    echo [Ошибка] Команда: %*
    echo Код: %errorlevel%
    pause
    exit
)
