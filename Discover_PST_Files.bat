
@ECHO OFF

SET AzureStorageAccountName=<REPLACE WITH YOUR OWN VALUE>
SET AzurePrimaryAccessKey=<REPLACE WITH YOUR OWN VALUE>
SET AzureBlobContainerName=<REPLACE WITH YOUR OWN VALUE>

ECHO BiTitan's Automated PST File Migration

TIMEOUT /t 10
ECHO.

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
