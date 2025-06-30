# "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -NoExit -Command "[Console]::InputEncoding = [System.Text.Encoding]::UTF8"
[console]::InputEncoding = [System.Text.Encoding]::UTF8
[console]::OutputEncoding = [System.Text.Encoding]::UTF8
$tnsnamesValue = @'

####  KTK  DB for Invoice and other client applications
KTKDB_CLIENTS =
  (DESCRIPTION =
    (FAILOVER=on)
    (LOAD_BALANCE=off)
    (CONNECT_TIMEOUT=5)(TRANSPORT_CONNECT_TIMEOUT=3)(RETRY_COUNT=3)
    (ADDRESS_LIST =
      (ADDRESS = (PROTOCOL = TCP)(HOST = 10.82.0.16)(PORT = 1521))
      (ADDRESS = (PROTOCOL = TCP)(HOST = 10.82.0.17)(PORT = 1521))
    )
    (CONNECT_DATA =
      (SERVER = DEDICATED)
      (SERVICE_NAME = ktkdb_clients)
    )
  )


#### DB WIN MOBILE
wmdb_CLIENTS =
  (DESCRIPTION =
    (FAILOVER=on)
    (LOAD_BALANCE=off)
    (CONNECT_TIMEOUT=5)(TRANSPORT_CONNECT_TIMEOUT=3)(RETRY_COUNT=3)
    (ADDRESS_LIST =
      (ADDRESS = (PROTOCOL = TCP)(HOST = 10.54.193.59)(PORT = 1521))
      (ADDRESS = (PROTOCOL = TCP)(HOST = 10.54.193.60)(PORT = 1521))
    )
    (CONNECT_DATA =
      (SERVER = DEDICATED)
      (SERVICE_NAME = wmdb_clients)
    )
  )

'@

$menuString = @"


0 - Выйти
1 > Полная пост-установка
- 2 > NLS_LANG
- 3 > Ярлыки GUI
- 4 > tnsnames.ora

Введите номер нужного пункта
"@

function EnsureAdmin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if (-not $isAdmin) {
        Write-Warning "Требуется перезапуск от имени администратора."
        pause
        exit
    }
}

# Функция проверки пути
function ThrowIfPathDoNotExist {
    param ([string]$path, [string]$message)
    if (-not (Test-Path $path)) {
        throw $message
    }
}

function AbortOrContinue {
    while ($true) {
        $answer = Read-Host "1 - Прервать`n2 - Продолжить`nВведите значение"
        if ($answer -eq "1") {
            throw "Прервано"
        } elseif ($answer -eq "2") {
            break;
        }
    }
}

function EditNLSLANGinRegistry {    
    $registryPath = "HKLM:\SOFTWARE\WOW6432Node\ORACLE\KEY_OraClient12Home1_32bit"
    $valueName = "NLS_LANG"
    $newValue = "AMERICAN_AMERICA.CL8MSWIN1251"
    $defaultValue = "AMERICAN_AMERICA.WE8MSWIN1252"
    Write-Host "Замена значения $valueName в реестре"

    # Проверяем, что ключ есть
    ThrowIfPathDoNotExist $registryPath "Не найден ключ реестра: $registryPath"
    Write-Host "Путь в реестре существует: $registryPath"

    # Первичное значение 
    $initialValue = Get-ItemProperty -Path $registryPath -Name $valueName | Select-Object -ExpandProperty $valueName
    if ($initialValue -cne $defaultValue) {
        Write-Warning "Неожиданное первичное значение $valueName - $initialValue"
        AbortOrContinue
    }
    Write-Host "Значение $valueName - $initialValue"

    # Устанавливаем значение
    Set-ItemProperty -Path $registryPath -Name $valueName -Value $newValue
    Write-Host "Изменение значения на $newValue"

    # Читаем значение
    $readValue = Get-ItemProperty -Path $registryPath -Name $valueName | Select-Object -ExpandProperty $valueName
    if ($readValue -cne $newValue) {
        Write-Warning "Прочитанное значение $valueName - $readValue - не совпадает с $newValue"
        AbortOrContinue
    }
    Write-Host "Прочитанное значение: $readValue" 
}

function CreateGUIshortcuts { 
    Write-Host "Ярлыки GUI"
    $guiDriveLetter = Read-Host "Введите букву диска, на котором расположен GUI"
    $exePath1 = $guiDriveLetter + ":\GUI_Invoice\GUI_Volna\INV_Clients.exe"
    $exePath2 = $guiDriveLetter + ":\GUI_Invoice\GUI_WIN\INV_Clients_WIN.exe"

    $WshShell = New-Object -ComObject WScript.Shell

    function CreateShortcut {
        param ([string]$name, [string]$exePath)
        $desktopPath = "C:\Users\Public\Desktop\" + $name + ".lnk"
        $Shortcut = $WshShell.CreateShortcut($desktopPath)
        $Shortcut.TargetPath = $exePath
        $Shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($exePath)
        $Shortcut.WindowStyle = 1  # Открывать в обычном окне
        $Shortcut.Description = "Ярлык для " + $name
        $Shortcut.Save()
    }

    ThrowIfPathDoNotExist $exePath1 "Не найден файл: $exePath1"
    Write-Host "Файл обнаружен: $exePath1"
    ThrowIfPathDoNotExist $exePath2 "Не найден файл: $exePath2"
    Write-Host "Файл обнаружен: $exePath2"

    CreateShortcut "INV_Clients" $exePath1
    Write-Host "Создын ярлык для $exePath1" 
    CreateShortcut "INV_Clients_WIN" $exePath2
    Write-Host "Создын ярлык для $exePath2" 
}

function WriteTNSnamesFile {
    Write-Host "Запись tnsnames.ora"
    $appDriveLetter = Read-Host "Введите букву диска, на котором расположен app"
    $appClientPath = $appDriveLetter + ":\app\client\"
    $targetPathPart = "\product\12.2.0\client_1\network\admin\tnsnames.ora"

    # Проверяем существование app/client
    ThrowIfPathDoNotExist $appClientPath "Не найден путь: $appClientPath"
    Write-Host "Путь обнаружен: $appClientPath"

    # Проверяем папки пользователей в app/client
    $clientFolders = Get-ChildItem -Path $appClientPath -Directory
    $clientFolderNames = $clientFolders.Name
    Write-Host "Обнаруженные папки: $clientFolderNames"
    $user = ""
    if ($clientFolderNames.Length -eq 0) {
        throw "Нет папок пользователей: $appClientPath" 
    } elseif ($clientFolderNames -is [string]) {
        $user = $clientFolderNames
    } else { # Если не 0 и не строка, значит массив с несколькими значениями
        $index = Read-Host "Введите индекс нужной папки (начиная с 0)"
        $index = [int]$index
        $user = $clientFolderNames[$index]
    }
    Write-Host "Выбранный пользовтель - $user"

    # Пишем файл
    $targetPath = $appClientPath + $user +  $targetPathPart
    Set-Content -Path $targetPath -Value $tnsnamesValue
    Write-Host "Файл записан: $targetPath"
}

function Attempt {
    param([ScriptBlock]$Callback)
    try { $Callback.Invoke() }
    catch {
        if ($_ -eq "Прервано") {
            Write-Host $_
            pause
            return
        }
        Write-Host $_ -ForegroundColor Red -BackgroundColor Black
        Write-Host $_.ScriptStackTrace -ForegroundColor Red -BackgroundColor Black
        pause
    }
}

function Case {
    $case = Read-Host $menuString
    if ($case -eq "0") {
        $global:work = $false;
    } elseif ($case -eq "1") {
        EditNLSLANGinRegistry
        CreateGUIshortcuts
        WriteTNSnamesFile
        Write-Host "Успех" -ForegroundColor Green
    } elseif ($case -eq "2") {
        EditNLSLANGinRegistry
    } elseif ($case -eq "3") {
        CreateGUIshortcuts
    } elseif ($case -eq "4") {
        WriteTNSnamesFile
    } else {
        Write-Host "Некорректный вариант"
    }
}

function Run {
    EnsureAdmin
    while ($global:work) { Attempt -Callback ${function:Case} }
}

$work = $true;
Attempt -Callback ${function:Run}
