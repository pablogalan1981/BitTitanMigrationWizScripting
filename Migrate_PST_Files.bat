
@ECHO OFF

SET AzureStorageAccountName=<REPLACE WITH YOUR OWN VALUE>
SET AzurePrimaryAccessKey=<REPLACE WITH YOUR OWN VALUE>
SET AzureBlobContainerName=<REPLACE WITH YOUR OWN VALUE>

ECHO BiTitan's Automated PST File Migration

TIMEOUT /t 10
ECHO.
Powershell "Write-host '1. Disconnecting all PST files from your Outlook.';try{$Outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop}catch{};if($Outlook){$Namespace = $Outlook.getNamespace('MAPI');$all_psts = $Namespace.Stores | Where-Object {($_.ExchangeStoreType -eq '3') -and ($_.FilePath -like '*.pst') -and ($_.IsDataFileStore -eq $true)}; ForEach ($pst in $all_psts){write-host 'PST file disconnected:' $pst.FilePath;try{$Outlook.Session.RemoveStore($pst.GetRootFolder())}catch{Write-Host -ForeGroundColor Red "ERROR: Failed to disconnect PST. Please close Outlook client."};}}" 
SET dir=%userprofile%\AppData\Local

IF EXIST %dir%\BitTitan\UploaderWiz\UploaderWiz.exe ( 

ECHO.
ECHO 2. UploaderWiz already downloaded to c:\BitTitan\UploaderWiz\UploaderWiz.exe
ECHO    Launching UploaderWiz desktop agent again to complete your PST migration

) ELSE  ( 

ECHO.
ECHO 2. Downloading and unzipping the latest version of the UploaderWiz desktop agent

if not exist "%dir%\BitTitan\UploaderWiz\" mkdir %dir%\BitTitan\UploaderWiz\ 

Powershell "try{Invoke-WebRequest -Uri https://api.bittitan.com/secure/downloads/UploaderWiz.zip -OutFile %dir%\BitTitan\UploaderWiz\UploaderWiz.zip -ErrorAction Stop}catch{Write-Host -ForeGroundColor Red "ERROR: Failed to execute Invoke-WebRequest, connection closed. Are you connecting through proxy?"}"
Powershell.exe -nologo -noprofile -command "& { Add-Type -A 'System.IO.Compression.FileSystem'; [IO.Compression.ZipFile]::ExtractToDirectory('%dir%\BitTitan\UploaderWiz\UploaderWiz.zip','%dir%\BitTitan\UploaderWiz\'); }"
copy migrate_pst_files.bat %dir%\BitTitan\UploaderWiz\

ECHO.
ECHO 3. Adding Windows Registry Key to relaunch UploaderWiz upon reboot
REG QUERY "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "PSTMigration" /t REG_SZ 
IF ERRORLEVEL 1 (REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "PSTMigration" /t REG_SZ /d "%dir%\BitTitan\UploaderWiz\migrate_pst_files.bat")

)

:LOOP1
tasklist | find /i "UploaderWiz" >nul 2>&1
IF ERRORLEVEL 1 (
  GOTO CONTINUE1
) ELSE (
  ECHO.
  ECHO WARNING: UploaderWiz already running. It will exit
  Timeout /T 15 /Nobreak  >nul 2>&1
  EXIT
)
:CONTINUE1

TIMEOUT /t 10
ECHO.
ECHO 4. Executing UploaderWiz desktop agent to discover all PST files to generate Voleer PST assessment report
START %dir%\BitTitan\UploaderWiz\UploaderWiz.exe -type azureblobs -accesskey %AzureStorageAccountName% -secretkey %AzurePrimaryAccessKey% -container %AzureBlobContainerName% -autodiscover true -interactive false -filefilter "*.pst" -force True -command GenerateMetadata

:LOOP2
tasklist | find /i "UploaderWiz" >nul 2>&1
IF ERRORLEVEL 1 (
  GOTO CONTINUE2
) ELSE (
  Timeout /T 5 /Nobreak  >nul 2>&1
  GOTO LOOP2
)
:CONTINUE2

TIMEOUT /t 10
ECHO.
ECHO 5. Executing UploaderWiz desktop agent to upload all PST files to Azure blob container %AzureBlobContainerName% 
START %dir%\BitTitan\UploaderWiz\UploaderWiz.exe -type azureblobs -accesskey %AzureStorageAccountName% -secretkey %AzurePrimaryAccessKey% -container %AzureBlobContainerName%  -autodiscover true -interactive false -filefilter "*.pst"

:LOOP3
tasklist | find /i "UploaderWiz" >nul 2>&1
IF ERRORLEVEL 1 (
  GOTO CONTINUE3
) ELSE (
  Timeout /T 5 /Nobreak  >nul 2>&1
  GOTO LOOP3
)
:CONTINUE3

ECHO.
ECHO 6. Removing Windows Registry Key to relaunch UploaderWiz upon reboot and batch file
REG DELETE "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "PSTMigration" /f

ECHO.
ECHO 7. The upload of all your PST files has been completed. You can close this window
TIMEOUT /t 99999 /Nobreak >nul 2>&1
