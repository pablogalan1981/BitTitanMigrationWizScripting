@ECHO OFF

SET GoogleClientId=<REPLACE WITH YOUR OWN VALUE>
SET GoogleClientSecret=<REPLACE WITH YOUR OWN VALUE>

SET AzureStorageAccountName=<REPLACE WITH YOUR OWN VALUE>
SET AzurePrimaryAccessKey=<REPLACE WITH YOUR OWN VALUE>
SET AzureBlobContainerName=<REPLACE WITH YOUR OWN VALUE>

ECHO BiTitan's Automated Google Vault Migration

TIMEOUT /t 10
ECHO.

:LOOP1
tasklist | find /i "GoogleVaultExport.exe" >nul 2>&1
IF ERRORLEVEL 1 (
GOTO CONTINUE1
) ELSE (
ECHO.
ECHO WARNING: GoogleVaultExport.exe already running. It will exit
Timeout /T 10 /Nobreak  >nul 2>&1
EXIT
)
:CONTINUE1

for /F "usebackq tokens=*" %%i in (`PowerShell -NoLogo -NonInteractive "Write-Host -nonewline -Separator '' $([Environment]::GetFolderPath('Desktop') )" `) do set desktopPath=%%i
for /F "usebackq tokens=*" %%i in (`PowerShell -NoLogo -NonInteractive "Write-Host -nonewline -Separator '' $($([Environment]::GetFolderPath('Desktop')) + """\GoogleVaultExport.lnk""")" `) do set shortcutPath=%%i

IF EXIST C:\GoogleVaultExport\EXE\GoogleVaultExportSetUp.exe ( 

ECHO.
ECHO 1. GoogleVaultExtractor already downloaded to c:\BitTitan\GoogleVaultExport\GoogleVaultExportSetUp.exe

IF EXIST C:\GoogleVaultExport\EXE\UploaderWiz.exe ( 
ECHO.
ECHO 2. UploaderWiz already downloaded to c:\BitTitan\GoogleVaultExport\EXE\UploaderWiz.exe
ECHO.
ECHO    Launching GoogleVaultExtractor desktop agent again to complete your Google Vault migration
ECHO    Launching UploaderWiz desktop agent again to complete your Google Vault migration

) ELSE  ( 
ECHO.
ECHO 2. Downloading and unzipping the latest version of the UploaderWiz desktop agent

if not exist "C:\GoogleVaultExport\EXE" mkdir C:\GoogleVaultExport\EXE 

Powershell "try{Invoke-WebRequest -Uri https://api.bittitan.com/secure/downloads/UploaderWiz.zip -OutFile C:\GoogleVaultExport\EXE\UploaderWiz.zip -ErrorAction Stop}catch{Write-Host -ForeGroundColor Red "ERROR: Failed to execute Invoke-WebRequest, connection closed. Are you connecting through proxy?"}"
Powershell.exe -nologo -noprofile -command "& { Add-Type -A 'System.IO.Compression.FileSystem'; [IO.Compression.ZipFile]::ExtractToDirectory('C:\GoogleVaultExport\EXE\UploaderWiz.zip','C:\GoogleVaultExport\EXE'); }"

)

ECHO.
ECHO 3. Adding Windows Registry Key to relaunch Migrate_GoogleVault.bat upon reboot
REG QUERY "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "GoogleVaultMigration" /t REG_SZ 
IF ERRORLEVEL 1 (REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "GoogleVaultMigration" /t REG_SZ /d "C:\GoogleVaultExport\EXE\Migrate_GoogleVault.bat")

) ELSE  ( 

ECHO.
ECHO 1. Downloading the latest version of the GoogleVaultExport desktop agent

if not exist "C:\GoogleVaultExport\EXE" mkdir C:\GoogleVaultExport\EXE 

Powershell "try{Invoke-WebRequest -Uri https://api.bittitan.com/public/downloads/GoogleVaultExport/GoogleVaultExportSetUp.exe -OutFile C:\GoogleVaultExport\EXE\GoogleVaultExportSetUp.exe -ErrorAction Stop}catch{Write-Host -ForeGroundColor Red "ERROR: Failed to execute Invoke-WebRequest, connection closed. Are you connecting through proxy?"}"

IF EXIST C:\GoogleVaultExport\EXE\UploaderWiz.exe ( 

ECHO.
ECHO 2. UploaderWiz already downloaded to c:\BitTitan\GoogleVaultExport\UploaderWiz.exe
ECHO.
ECHO    Launching GoogleVaultExtractor desktop agent again to complete your Google Vault migration
ECHO    Launching UploaderWiz desktop agent again to complete your Google Vault migration

) ELSE  ( 
ECHO.
ECHO 2. Downloading and unzipping the latest version of the UploaderWiz desktop agent

if not exist "C:\GoogleVaultExport\EXE" mkdir C:\GoogleVaultExport\EXE 

Powershell "try{Invoke-WebRequest -Uri https://api.bittitan.com/secure/downloads/UploaderWiz.zip -OutFile C:\GoogleVaultExport\EXE\UploaderWiz.zip -ErrorAction Stop}catch{Write-Host -ForeGroundColor Red "ERROR: Failed to execute Invoke-WebRequest, connection closed. Are you connecting through proxy?"}"
Powershell.exe -nologo -noprofile -command "& { Add-Type -A 'System.IO.Compression.FileSystem'; [IO.Compression.ZipFile]::ExtractToDirectory('C:\GoogleVaultExport\EXE\UploaderWiz.zip','C:\GoogleVaultExport\EXE'); }"

)

ECHO.
ECHO 3. Adding Windows Registry Key to relaunch Migrate_GoogleVault.bat upon reboot
REG QUERY "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "GoogleVaultMigration" /t REG_SZ 
IF ERRORLEVEL 1 (REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "GoogleVaultMigration" /t REG_SZ /d "C:\GoogleVaultExport\EXE\Migrate_GoogleVault.bat")

)

ECHO.
for /F "usebackq tokens=1" %%i in (`PowerShell -NoLogo -NonInteractive "Write-Host -nonewline -Separator '' "GoogleVaultExport\app-"$((get-item C:\GoogleVaultExport\EXE\GoogleVaultExportSetUp.exe).VersionInfo.FileVersion)"`) do set exeVersion=%%i
SET installDir=%userprofile%\AppData\Local\%exeVersion%\

ECHO.
ECHO 4. Executing GoogleVaultExportSetUp.exe to create GoogleVaultExport.lnk in your desktop. 
ECHO    Don't worry if GoogleVaultExportSetUp.exe displays an error.

START C:\GoogleVaultExport\EXE\GoogleVaultExportSetUp.exe

:LOOP2
tasklist | find /i "GoogleVaultExport.exe" >nul 2>&1
IF ERRORLEVEL 1 (
Timeout /T 15 /Nobreak  >nul 2>&1
GOTO LOOP2
) ELSE (
ECHO.
taskkill /f /im "GoogleVaultExport.exe"
)
:CONTINUE2

TIMEOUT /t 10
ECHO.
ECHO 5. Import all users to export from Google Vault

if not exist "C:\GoogleVaultExport\ExportedGoogleVault" mkdir "C:\GoogleVaultExport\ExportedGoogleVault"
if not exist "C:\GoogleVaultExport\UsersToProcess.txt" (
@echo off
break>"C:\GoogleVaultExport\UsersToProcess.txt
)
echo Enter the user email addresses you want to extract from Google Vault in a single column separated by , in UsersToProcess.txt.

notepad "C:\GoogleVaultExport\UsersToProcess.txt"

echo Press any key to Continue.
pause > nul

ECHO.
ECHO 6. Copying files to their corresponding folders
ECHO    GoogleVaultExportSetUp.exe original installation folder path %installDir%
ECHO    Copying file Migrate-GoogleVault.bat to C:\GoogleVaultExport\EXE\ folder
copy Migrate-GoogleVault.bat C:\GoogleVaultExport\EXE
ECHO    Copying GoogleVaultExport.lnk to C:\GoogleVaultExport\EXE\ folder
copy "%shortcutPath%" C:\GoogleVaultExport\EXE\

for /F "usebackq tokens=1" %%i in (`PowerShell -NoLogo -NonInteractive "$sh = New-Object -COM WScript.Shell;$WorkingDirectory = $sh.CreateShortcut('C:\GoogleVaultExport\EXE\GoogleVaultExport.lnk').WorkingDirectory;Write-Host $WorkingDirectory "`) do set installDir=%%i

ECHO    Copying UploaderWiz.exe to %installDir%\ folder
copy "C:\GoogleVaultExport\EXE\UploaderWiz.exe" %installDir%
copy "C:\GoogleVaultExport\EXE\UploaderWiz.exe.config" %installDir%

ECHO.
ECHO 7. Exporting Google Vault data to c:\GoogleVaultExport\ExportedGoogleVault\ 
ECHO    And uploading it to Azure blob container %AzureBlobContainerName%

START C:\GoogleVaultExport\EXE\GoogleVaultExport.lnk -process-start-args "-command exportandupload -clientid %GoogleClientId% -clientSecret %GoogleClientSecret% -NewEmailsTimeout 86400000 -CompressionTimeout 86400000 -inputFile C:\GoogleVaultExport\UsersToProcess.txt -outputFolder c:\GoogleVaultExport\ExportedGoogleVault -searchTerms ""label:^deleted"" -uploadAccessKey %AzureStorageAccountName% -uploadSecretKey %AzurePrimaryAccessKey% -uploadBucketName %AzureBlobContainerName% "

TIMEOUT /t 5 /nobreak >nul 2>&1

:LOOP4
tasklist | find /i "GoogleVaultExport.exe" >nul 2>&1
IF ERRORLEVEL 1 (
GOTO CONTINUE4
) ELSE (
Timeout /T 10 /Nobreak  >nul 2>&1
GOTO LOOP4
)
:CONTINUE4

ECHO.
ECHO 8. Removing Windows Registry Key to relaunch Migrate_GoogleVault.bat upon reboot and batch file
REG DELETE "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "GoogleVaultMigration" /f

ECHO.
ECHO 9. The extraction and upload of all Google Vault users has been completed. You can close this window.
TIMEOUT /t 99999 /Nobreak >nul 2>&1
