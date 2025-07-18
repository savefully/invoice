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

0 - Exit
1 - Full installation

Custom installation:
2 - NLS_LANG
3 - GUI Links
4 - tnsnames.ora

Input:
"@
$oracleRegistryPath = "HKLM:\SOFTWARE\WOW6432Node\ORACLE"

function EnsureAdmin {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host "Run as Administrator"
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
        Write-Host "1 - Abort`n2 - Continue`nInput:"
        $answer = Read-Host ":"
        if ($answer -eq "1") {
            throw "Aborted"
        } elseif ($answer -eq "2") {
            break;
        }
    }
}

function GetOracleHomeRegistryKeys {
    $basePath = "HKLM:\SOFTWARE\WOW6432Node\ORACLE"
    if (-not (Test-Path $oracleRegistryPath)) {
        throw "Path is not found: $oracleRegistryPath"
        return @()
    }

    $oracleKeys = Get-ChildItem -Path $oracleRegistryPath -ErrorAction SilentlyContinue
    $foundKeys = @()
    $pattern = "^KEY_OraClient12Home\d+_32bit$"

    foreach ($key in $oracleKeys) {
        $keyName = $key.PSChildName
        
        if ($keyName -match $pattern) {
            $foundKeys += $keyName
        }
    }
    if ($foundKeys.length -eq 0) {
        throw "No registry keys are found: $oracleRegistryPath\$pattern"
    }
    return foundKeys
}

function EditNLSLANGinRegistry {    
    $registryPath = $null;
    $valueName = "NLS_LANG"
    $newValue = "AMERICAN_AMERICA.CL8MSWIN1251"
    $defaultValue = "AMERICAN_AMERICA.WE8MSWIN1252"
    Write-Host "Handle $valueName in registry"

    $oracleHomeRegistryKeys = GetOracleHomeRegistryKeys
    Write-Host "Found oracle home keys: $oracleHomeRegistryKeys"
    if ($oracleHomeRegistryKeys.length -eq 1) {
        $registryPath = "$oracleRegistryPath\$($oracleHomeRegistryKeys[0])"
    } else {
        Write-Host "Input oracle home key index starting from 0:"
        $index = Read-Host ":"
        $index = [int]$index
        $registryPath = "$oracleRegistryPath\$($oracleHomeRegistryKeys[$index])"
    }
     
    $initialValue = Get-ItemProperty -Path $registryPath -Name $valueName | Select-Object -ExpandProperty $valueName
    if ($initialValue -cne $defaultValue) {
        Write-Warning "Unexpected initial value $valueName - $initialValue"
        AbortOrContinue
    }
    Write-Host "Initial value $valueName - $initialValue"

    # Устанавливаем значение
    Set-ItemProperty -Path $registryPath -Name $valueName -Value $newValue
    Write-Host "Change value: $newValue"

    # Читаем значение
    $readValue = Get-ItemProperty -Path $registryPath -Name $valueName | Select-Object -ExpandProperty $valueName
    if ($readValue -cne $newValue) {
        Write-Warning "Value that is read $valueName - $readValue - is not $newValue"
        AbortOrContinue
    }
    Write-Host "Read value: $readValue" 
}

function CreateGUIshortcuts { 
    Write-Host "GUI shortcuts"
    Write-Host 'Input the letter of the disk partition where "GUI_Invoice" is located:'
    $guiDriveLetter = Read-Host ":"
    $exePath1 = $guiDriveLetter + ":\GUI_Invoice\GUI_Volna\INV_Clients.exe"
    $exePath2 = $guiDriveLetter + ":\GUI_Invoice\GUI_WIN\INV_Clients_WIN.exe"

    $WshShell = New-Object -ComObject WScript.Shell

    function CreateShortcut {
        param ([string]$name, [string]$exePath)
        $desktopPath = "C:\Users\Public\Desktop\" + $name + ".lnk"
        $Shortcut = $WshShell.CreateShortcut($desktopPath)
        $Shortcut.TargetPath = $exePath
        $Shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($exePath)
        $Shortcut.WindowStyle = 1
        $Shortcut.Description = $name
        $Shortcut.Save()
    }

    ThrowIfPathDoNotExist $exePath1 "File is not found: $exePath1"
    Write-Host "File is found: $exePath1"
    ThrowIfPathDoNotExist $exePath2 "File is not found: $exePath2"
    Write-Host "File is found: $exePath2"

    CreateShortcut "INV_Clients" $exePath1
    Write-Host "Shortcut has been created for $exePath1" 
    CreateShortcut "INV_Clients_WIN" $exePath2
    Write-Host "Shortcut has been created for $exePath2" 
}

function WriteTNSnamesFile {
    Write-Host "Writing tnsnames.ora"
    Write-Host 'Input the letter of the disk partition where "app" is located:'
    $appDriveLetter = Read-Host ":"
    $appClientPath = $appDriveLetter + ":\app\client\"
    $targetPathPart = "\product\12.2.0\client_1\network\admin\tnsnames.ora"

    # Проверяем существование app/client
    ThrowIfPathDoNotExist $appClientPath "Path is not found: $appClientPath"
    Write-Host "Path is found: $appClientPath"

    # Проверяем папки пользователей в app/client
    $clientFolders = Get-ChildItem -Path $appClientPath -Directory
    $clientFolderNames = $clientFolders.Name
    Write-Host "Located folders: $clientFolderNames"
    $user = ""
    if ($clientFolderNames.Length -eq 0) {
        throw "No user-folders: $appClientPath" 
    } elseif ($clientFolderNames -is [string]) {
        $user = $clientFolderNames
    } else { # Если не 0 и не строка, значит массив с несколькими значениями
        Write-Host "Input folder index starting from 0:"
        $index = Read-Host ":"
        $index = [int]$index
        $user = $clientFolderNames[$index]
    }
    Write-Host "Selected user - $user"

    # Пишем файл
    $targetPath = $appClientPath + $user +  $targetPathPart
    Set-Content -Path $targetPath -Value $tnsnamesValue
    Write-Host "File is written: $targetPath"
}

function Attempt {
    param([ScriptBlock]$Callback)
    try { $Callback.Invoke() }
    catch {
        if ($_ -eq "Aborted") {
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
    Write-Host $menuString
    $case = Read-Host ":"
    if ($case -eq "0") {
        $global:work = $false;
    } elseif ($case -eq "1") {
        EditNLSLANGinRegistry
        CreateGUIshortcuts
        WriteTNSnamesFile
        Write-Host "Success." -ForegroundColor Green
    } elseif ($case -eq "2") {
        EditNLSLANGinRegistry
    } elseif ($case -eq "3") {
        CreateGUIshortcuts
    } elseif ($case -eq "4") {
        WriteTNSnamesFile
    } else {
        Write-Host "No such option: $case"
    }
}

function Run {
    EnsureAdmin
    while ($global:work) { Attempt -Callback ${function:Case} }
}

$work = $true;
Attempt -Callback ${function:Run}
