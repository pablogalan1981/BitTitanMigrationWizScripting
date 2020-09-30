<#


.SYNOPSIS
     
Copyright 2020 BitTitan, Inc.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. 

You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.

.DESCRIPTION
    This script will create a MigrationWiz project to migrate FileServer Home Directories to Google Drive accounts.
    It will generate a CSV file with the MigrationWiz project and all the migrations that will be used by the script 
    Start-MW_FileServerToGoogleDrive.ps1 to submit all the migrations.
    
.NOTES
    Author          Pablo Galan Sabugo <pablog@bittitan.com> from the BitTitan Technical Sales Specialist Team
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
#>

Param
(
    [Parameter(Mandatory = $false)] [String]$WorkingDirectory,
    [Parameter(Mandatory = $false)] [Boolean]$downloadLatestVersion,
    [Parameter(Mandatory = $false)] [ValidateSet('NorthAmerica','WesternEurope','AsiaPacific','Australia','Japan','SouthAmerica','Canada','NorthernEurope','China','France','SouthAfrica')] [String]$BitTitanAzureDatacenter,
    [Parameter(Mandatory = $false)] [String]$BitTitanWorkgroupId,
    [Parameter(Mandatory = $false)] [String]$BitTitanCustomerId,
    [Parameter(Mandatory = $false)] [String]$BitTitanSourceEndpointId,
    [Parameter(Mandatory = $false)] [String]$AzureStorageAccessKey,
    [Parameter(Mandatory = $false)] [String]$BitTitanDestinationEndpointId,
    [Parameter(Mandatory = $false)] [String]$FileServerRootFolderPath,
    [Parameter(Mandatory = $false)] [String]$HomeDirectorySeachPattern,
    [Parameter(Mandatory = $false)] [Boolean]$CheckFileServer,
    [Parameter(Mandatory = $false)] [Boolean]$CheckOneDriveAccounts,
    [Parameter(Mandatory = $false)] [String]$MigrationWizFolderMapping,
    [Parameter(Mandatory = $false)] [Boolean]$ApplyUserMigrationBundle
)


#######################################################################################################################
#                  HELPER FUNCTIONS                          
#######################################################################################################################
Function Import-PowerShellModules{
    if (!(((Get-Module -Name "MSOnline") -ne $null) -or ((Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue) -ne $null))) {
        Write-Host
        $msg = "INFO: MSOnline PowerShell module not installed."
        Write-Host $msg     
        $msg = "INFO: Installing MSOnline PowerShell module."
        Write-Host $msg

        Sleep 5
    
        try{
            Install-Module -Name MSOnline -force -ErrorAction Stop
        }
        catch{
            $msg = "ERROR: Failed to install MSOnline module. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Write-Host
            $msg = "ACTION: Run this script 'As administrator' to intall the MSOnline module."
            Write-Host -ForegroundColor Yellow $msg
            Exit
        }
        Import-Module MSOnline
    }

    if (!(((Get-Module -Name "AzureAD") -ne $null) -or ((Get-InstalledModule -Name "AzureAD" -ErrorAction SilentlyContinue) -ne $null))) {
        Write-Host
        $msg = "INFO: AzureAD PowerShell module not installed."
        Write-Host $msg     
        $msg = "INFO: Installing AzureAD PowerShell module."
        Write-Host $msg

        Sleep 5
    
        try{
            Install-Module -Name AzureAD -force -ErrorAction Stop
        }
        catch{
            $msg = "ERROR: Failed to install AzureAD module. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Write-Host
            $msg = "ACTION: Run this script 'As administrator' to intall the AzureAD module."
            Write-Host -ForegroundColor Yellow $msg
            Exit
        }
        Import-Module AzureAD
    }
}

function Import-MigrationWizPowerShellModule {
        if (( $null -ne (Get-Module -Name "BitTitanPowerShell")) -or ( $null -ne (Get-InstalledModule -Name "BitTitanManagement" -ErrorAction SilentlyContinue))) {
        return
    }

    $currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
    $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll",  "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
    foreach ($moduleLocation in $moduleLocations) {
        if (Test-Path $moduleLocation) {
            Import-Module -Name $moduleLocation
            return
        }
    }
    
    $msg = "INFO: BitTitan PowerShell SDK not installed."
    Write-Host -ForegroundColor Red $msg 

    Write-Host
    $msg = "ACTION: Install BitTitan PowerShell SDK 'bittitanpowershellsetup.msi' downloaded from 'https://www.bittitan.com'."
    Write-Host -ForegroundColor Yellow $msg

    Start-Sleep 5

    $url = "https://www.bittitan.com/downloads/bittitanpowershellsetup.msi " 
    $result= Start-Process $url
    Exit

}

Function Get-CsvFile {
    Write-Host
    Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import the user email addresses."
    Get-FileName $script:workingDir -ExtensionFilter "CSV (*.csv)| *.csv"

    # Import CSV and validate if headers are according the requirements
    try {
        $lines = @(Import-Csv $script:inputFile)
    }
    catch {
        $msg = "ERROR: Failed to import '$inputFile' CSV file. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message   
        Exit   
    }

    # Validate if CSV file is empty
    if ( $lines.count -eq 0 ) {
        $msg = "ERROR: '$inputFile' CSV file exist but it is empty. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Exit
    }

    # Validate CSV Headers
    $CSVHeaders = "UserEmailAddress,FirstName"
    foreach ($header in $CSVHeaders) {
        if ($lines.$header -eq "" ) {
            $msg = "ERROR: '$inputFile' CSV file does not have all the required columns. Required columns are: '$($CSVHeaders -join "', '")'. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg   
            Exit
        }
    }

    Return $lines
 
 }

 # Function to get a CSV file name or to create a new CSV file
 Function Get-FileName($initialDirectory,$DefaultColumnName,$extensionFilter) {

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = $extensionFilter
    $OpenFileDialog.ShowDialog() | Out-Null
    $script:inputFile = $OpenFileDialog.filename

    if($OpenFileDialog.filename -eq "") {

            if($defaultColumnName -eq "PrimarySmtpAddress") {
                # create new import file
	            $inputFileName = "FilteredPrimarySmtpAddress-$((Get-Date).ToString("yyyyMMddHHmmss")).csv"
                $script:inputFile = "$initialDirectory\$inputFileName"

                $csv = "PrimarySmtpAddress`r`n"
                $file = New-Item -Path $initialDirectory -name $inputFileName -ItemType file -value $csv -force 

                $msg = "SUCCESS: Empty CSV file '$script:inputFile' created."
                Write-Host -ForegroundColor Green  $msg
                
                $msg = "WARNING: Populate the CSV file with the source 'PrimarySmtpAddress' you want to process in a single column and save it as`r`n         File Type: 'CSV (Comma Delimited) (.csv)'`r`n         File Name: '$script:inputFile'."
                Write-Host -ForegroundColor Yellow $msg
            }
            elseif($defaultColumnName -eq "MailNickName") {
                # create new import file
	            $inputFileName = "FilteredTeamMailNickName-$((Get-Date).ToString("yyyyMMddHHmmss")).csv"
                $script:inputFile = "$initialDirectory\$inputFileName"

                $csv = "MailNickName`r`n"
                $file = New-Item -Path $initialDirectory -name $inputFileName -ItemType file -value $csv -force 

                $msg = "SUCCESS: Empty CSV file '$script:inputFile' created."
                Write-Host -ForegroundColor Green  $msg
                
                $msg = "WARNING: Populate the CSV file with the source 'MailNickName' you want to process in a single column and save it as`r`n         File Type: 'CSV (Comma Delimited) (.csv)'`r`n         File Name: '$script:inputFile'."
                Write-Host -ForegroundColor Yellow $msg  
            }
            elseif($defaultColumnName -eq "SourceEmailAddress,DestinationEmailAddress") {
                # create new import file
	            $inputFileName = "UserMapping-$((Get-Date).ToString("yyyyMMddHHmmss")).csv"
                $script:inputFile = "$initialDirectory\$inputFileName"

                $csv = "SourceEmailAddress,DestinationEmailAddress`r`n"
                $file = New-Item -Path $initialDirectory -name $inputFileName -ItemType file -value $csv -force 

                $msg = "SUCCESS: Empty CSV file '$script:inputFile' created."
                Write-Host -ForegroundColor Green  $msg
                
                $msg = "WARNING: Populate the CSV file with the 'SourceEmailAddress', 'DestinationEmailAddress' columns and save it as`r`n         File Type: 'CSV (Comma Delimited) (.csv)'`r`n         File Name: '$script:inputFile'."
                Write-Host -ForegroundColor Yellow $msg  
            }
            elseif($defaultColumnName -eq "SourceFolder,DestinationEmailAddress") {
                # create new import file
	            $inputFileName = "UserMapping-$((Get-Date).ToString("yyyyMMddHHmmss")).csv"
                $script:inputFile = "$initialDirectory\$inputFileName"

                $csv = "SourceFolder,DestinationEmailAddress`r`n"
                $file = New-Item -Path $initialDirectory -name $inputFileName -ItemType file -value $csv -force 

                $msg = "SUCCESS: Empty CSV file '$script:inputFile' created."
                Write-Host -ForegroundColor Green  $msg
                
                $msg = "WARNING: Populate the CSV file with the 'SourceFolder', 'DestinationEmailAddress' columns and save it as`r`n         File Type: 'CSV (Comma Delimited) (.csv)'`r`n         File Name: '$script:inputFile'."
                Write-Host -ForegroundColor Yellow $msg  
            }
            else{
                Return $false
            }            

		    # open file for editing
		    Start-Process $file 

		    do {
			    $confirm = (Read-Host -prompt "Are you done editing the import CSV file?  [Y]es or [N]o")
		        if($confirm -eq "Y") {
			        $importConfirm = $true
		        }

		        if($confirm -eq "N") {
			        $importConfirm = $false
                    try {
                        #Open the CSV file for editing
                        Start-Process -FilePath $script:inputFile
                    }
                    catch {
                        $msg = "ERROR: Failed to open '$script:inputFile' CSV file. Script aborted."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg
                        Write-Host -ForegroundColor Red $_.Exception.Message
                        Log-Write -Message $_.Exception.Message
                    }
		        }
		    }
		    while(-not $importConfirm)
            
            $msg = "SUCCESS: CSV file '$script:inputFile' saved."
            Write-Host -ForegroundColor Green  $msg

            Return $true
    }
    else{
        $msg = "SUCCESS: CSV file '$($OpenFileDialog.filename)' selected."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
        Return $true
    }
}

# Function to write information to the Log File
Function Log-Write {
    param
    (
        [Parameter(Mandatory=$true)]    [string]$Message
    )

    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
	Add-Content -Path $logFile -Value $lineItem
}
 
# Function to create the working and log directories
Function Create-Working-Directory {    
    param 
    (
        [CmdletBinding()]
        [parameter(Mandatory=$true)] [string]$workingDir,
        [parameter(Mandatory=$true)] [string]$logDir,
        [parameter(Mandatory=$false)] [string]$metadataDir
    )
    if ( !(Test-Path -Path $workingDir)) {
		try {
			$suppressOutput = New-Item -ItemType Directory -Path $workingDir -Force -ErrorAction Stop
            $msg = "SUCCESS: Folder '$($workingDir)' for CSV files has been created."
            Write-Host -ForegroundColor Green $msg
		}
		catch {
            $msg = "ERROR: Failed to create '$workingDir'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
		}
    }
    if ( !(Test-Path -Path $logDir)) {
        try {
            $suppressOutput = New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop      

            $msg = "SUCCESS: Folder '$($logDir)' for log files has been created."
            Write-Host -ForegroundColor Green $msg 
        }
        catch {
            $msg = "ERROR: Failed to create log directory '$($logDir)'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
        } 
    }
    if ( $metadataDir -and !(Test-Path -Path $metadataDir)) {
        try {
            $suppressOutput = New-Item -ItemType Directory -Path $metadataDir -Force -ErrorAction Stop      

            $msg = "SUCCESS: Folder '$($metadataDir)' for PST metadata files has been created."
            Write-Host -ForegroundColor Green $msg 
        }
        catch {
            $msg = "ERROR: Failed to create log directory '$($metadataDir)'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
        } 
    }
}

# Function to wait for the user to press any key to continue
Function WaitForKeyPress{
    param 
    (      
        [parameter(Mandatory=$true)] [string]$message
    )
    
    Write-Host $message
    Log-Write -Message $message 
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
}

# Function to download a file from a URL
Function Download-File {
    param 
    (      
        [parameter(Mandatory=$true)] [String]$url,
        [parameter(Mandatory=$true)] [String]$outFile
    )

    $fileName = $($url.split("/")[-1])
    $folderName = $fileName.split(".")[0]

    $msg = "INFO: Downloading the latest version of '$fileName' agent (~12MB) from BitTitan..."
    Write-Host $msg
    Log-Write -Message $msg 

    #Download the latest version of UploaderWiz from BitTitan server
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    try {
        $result = Invoke-WebRequest -Uri $url -OutFile $outFile
        $msg = "SUCCESS: '$fileName' file downloaded into '$PSScriptRoot'."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg 
    }
    catch {
        $msg = "ERROR: Failed to download '$fileName'."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg    
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Unzip-File $outFile 

    #Open the zip file 
    try {
    
            Start-Process -FilePath "$PSScriptRoot\$folderName"


        }
        catch {
            $msg = "ERROR: Failed to open '$PSScriptRoot' folder."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message 
            Exit
        }

    
   # else {
   #     $msg = 
   #     "ERROR: Failed to download  UploaderWiz agent from BitTitan."
   #     Write-Host -ForegroundColor Red  $msg
   #     Log-Write -Message $msg 
   # }
 
 }

 # Function to unzip a file
Function Unzip-File {
    param 
    (      
        [parameter(Mandatory=$true)] [String]$zipfile
    )

    $folderName = (Get-Item $zipfile).Basename
    $fileName = $($zipfile.split("\")[-1])

    $result = New-Item -ItemType directory -Path $folderName -Force 

    try {
        $result = Expand-Archive $zipfile -DestinationPath $folderName -Force

        $msg = "SUCCESS: '$fileName' file unzipped into '$PSScriptRoot\$folderName'."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg 
    }
    catch {
        $msg = "ERROR: Failed to unzip '$fileName' file."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message
        Exit
    }
}

 # Function to remove invalid chars from folders and files in the File Server
Function Check-FileServerInvalidCharacters ($Path) {

    if([string]::IsNullOrEmpty($CheckFileServer) -or !$CheckFileServer){
        Write-Host "INFO: Analyzing invalid characters in all files and folders under File Server '$Path'. "

        do {        
            $confirm = (Read-Host -prompt "Do you want to remove invalid characters from folders and files?  [Y]es or [N]o")

            if($confirm.ToLower() -eq "y") {
                $removeInvalidChars = $true
            }
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    }

    #Get all files and folders under the path specified
    $items = Get-ChildItem -Path $Path -Recurse

    foreach ($item in $items) {
        #Check if the item is a file or a folder
        if ($item.PSIsContainer) { $type = "Folder" }
        else { $type = "File" }
   
        #Check if item name is 248 characters or more in length
        #UploaderWiz can only support file path names that are shorter than 248 characters.
        if ($item.Name.Length -gt 248) {
            $msg = "INFO: $type $($item.Name) is 248 characters or over item name and will need to be truncated to be uploaded by UploaderWiz." 
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile 
        }
        elseif($item.VersionInfo.FileName.length -gt 248){
            $msg = "INFO: $type $($item.VersionInfo.FileName) is a 248 characters or over file path and will need to be truncated to be uploaded by UploaderWiz." 
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile 
        }
        else {
            #Characters that aren't allowed in file and folder names  in OneDrive, Google Drive on Office 365, and SharePoint Online
            #   " * : < > ? / \ |# %
            #If the filename contains no bytes above 0x7F, then it's ASCII. 
            # gci -recurse . | where {$_.Name -match '[^\u0000-\u007F]'}
            
            $windowsIllegalChars = '[!&{}~#%]'
            $spoIllegalChars = '["*:<>?/\\|]'
            $nonAsciiChars = '[^\u0000-\u007F]'

            #Replace illegal characters with legal characters where found
            $newFileName = $item.Name

            filter Matches($nonAsciiChars) {
                $newFileName | Select-String -AllMatches $nonAsciiChars |
                Select-Object -ExpandProperty Matches
                Select-Object -ExpandProperty Values
            }
                        
            Matches $nonAsciiChars | ForEach-Object {

                Write-Host "ERROR: $type $item.FullName has an non ASCII character $_" -ForegroundColor Red
                #These characters may be used on the file system but not SharePoint
                if ($_ -match '[^\u0000-\u007F]') { $newFileName = ($newFileName -replace '[^\u0000-\u007F]', "") }
            }            

            filter Matches($spoIllegalChars) {
                $newFileName | Select-String -AllMatches $spoIllegalChars |
                Select-Object -ExpandProperty Matches
                Select-Object -ExpandProperty Values
            }
                        
            Matches $spoIllegalChars | ForEach-Object {

                Write-Host "ERROR: $type $item.FullName has the illegal character $($_.Value)" -ForegroundColor Red
                #These characters may be used on the file system but not SharePoint
                if ($_.Value -match '"') { $newFileName = ($newFileName -replace '"', "") }
                elseif ($_.Value -match '"') { $newFileName = ($newFileName -replace '"', '') }
                elseif ($_.Value -match "*") { $newFileName = ($newFileName -replace "*", "") }
                elseif ($_.Value -match ":") { $newFileName = ($newFileName -replace ":", "") }
                elseif ($_.Value -match "<") { $newFileName = ($newFileName -replace "<", "") }
                elseif ($_.Value -match ">") { $newFileName = ($newFileName -replace ">", "") }
                elseif ($_.Value -match "?") { $newFileName = ($newFileName -replace "?", "") }
                elseif ($_.Value -match "/") { $newFileName = ($newFileName -replace "/", "") }
                elseif ($_.Value -match "\") { $newFileName = ($newFileName -replace "\", "") }
                elseif ($_.Value -match "|") { $newFileName = ($newFileName -replace "|", "") }
            }


            #Check for start, end and double periods
            if ($newFileName.StartsWith(".")) { Write-Host "ERROR: $type $item.FullName starts with a period" -ForegroundColor Red }
            while ($newFileName.StartsWith(".")) { $newFileName = $newFileName.TrimStart(".") }
            
            if ($newFileName.EndsWith(".")) { Write-Host "ERROR: $type $item.FullName ends with a period" -ForegroundColor Red }
            while ($newFileName.EndsWith("."))   { $newFileName = $newFileName.TrimEnd(".") }
            
            if ($newFileName.Contains("..")) { Write-Host "ERROR: $type $item.FullName contains double periods" -ForegroundColor Red }
            while ($newFileName.Contains(".."))  { $newFileName = $newFileName.Replace("..", ".") }
            
            #Fix file and folder names if found and the Fix switch is specified
            if (($newFileName -ne $item.Name) -and ($removeInvalidChars)) {
                try {
                Rename-Item $item.FullName -NewName ($newFileName)
                Write-Host "SUCCESS: $type $item.Name has been changed to $newFileName" -ForegroundColor Green
                }
                catch {
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg    
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $_.Exception.Message 
                }
                
            }
        }
    }
}

# Function to get the licenses of each of the Office 365 users
Function Get-GoogleDriveAccounts {

    $result = Get-FileName $script:workingDir -DefaultColumnName "SourceFolder,DestinationEmailAddress" -ExtensionFilter "CSV (*.csv)| *.csv"
    
    if($result) {
        ##Import the CSV file
        Try{ 
            $GoogleDriveArray = @(Import-Csv $script:inputFile | Where-Object { $_.PSObject.Properties.Value -ne ""} )
        }
        Catch [Exception ]{
            Write-Host -ForegroundColor Red "ERROR: Failed to import the CSV file '$script:inputFile'."
            Write-Host -ForegroundColor Red $_.Exception.Message
            Exit
        }

        #Check if CSV is formated properly
        if ($GoogleDriveArray.SourceFolder -eq $null -or $GoogleDriveArray.DestinationEmailAddress -eq $null) {
            Write-Host -ForegroundColor Red "ERROR: The CSV file format is invalid. It must have 2 columns: 'SourceFolder' and 'DestinationEmailAddress' columns."
            Exit 
        }

        $GoogleDriveArray = $GoogleDriveArray | sort-object DestinationEmailAddress -Unique
        Return $GoogleDriveArray 
    }
    else {
        Return $false 
    }
}

######################################################################################################################################
#                                    CONNECTION TO BITTITAN
######################################################################################################################################

# Function to authenticate to BitTitan SDK
Function Connect-BitTitan {
    #[CmdletBinding()]

    #Install Packages/Modules for Windows Credential Manager if required
    If(!(Get-PackageProvider -Name 'NuGet')){
        Install-PackageProvider -Name NuGet -Force
    }
    If(!(Get-Module -ListAvailable -Name 'CredentialManager')){
        Install-Module CredentialManager -Force
    } 
    else { 
        Import-Module CredentialManager
    }

    # Authenticate
    $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'
    
    if(!$script:creds){
        $credentials = (Get-Credential -Message "Enter BitTitan credentials")
        if(!$credentials) {
            $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Exit
        }
        New-StoredCredential -Target 'https://migrationwiz.bittitan.com' -Persist 'LocalMachine' -Credentials $credentials | Out-Null
        
        $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' stored in Windows Credential Manager."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg

        $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'

        $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' retrieved from Windows Credential Manager."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }
    else{
        $msg = "SUCCESS: BitTitan credentials for target 'https://migrationwiz.bittitan.com' retrieved from Windows Credential Manager."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }

    try { 
        # Get a ticket and set it as default
        $script:ticket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction Stop
        # Get a MW ticket
        $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction Stop 
    }
    catch {

        $currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
        $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll",  "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
        foreach ($moduleLocation in $moduleLocations) {
            if (Test-Path $moduleLocation) {
                Import-Module -Name $moduleLocation

                # Get a ticket and set it as default
                $script:ticket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction SilentlyContinue
                # Get a MW ticket
                $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction SilentlyContinue 

                if(!$script:ticket -or !$script:mwTicket) {
                    $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                    Exit
                }
                else {
                    $msg = "SUCCESS: Connected to BitTitan."
                    Write-Host -ForegroundColor Green  $msg
                    Log-Write -Message $msg
                }

                return
            }
        }

        $msg = "ACTION: Install BitTitan PowerShell SDK 'bittitanpowershellsetup.msi' downloaded from 'https://www.bittitan.com' and execute the script from there."
        Write-Host -ForegroundColor Yellow $msg
        Write-Host

        Start-Sleep 5

        $url = "https://www.bittitan.com/downloads/bittitanpowershellsetup.msi " 
        $result= Start-Process $url

        Exit
    }  

    if(!$script:ticket -or !$script:mwTicket) {
        $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        Exit
    }
    else {
        $msg = "SUCCESS: Connected to BitTitan."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }
}

# Function to create an endpoint under a customer
# Configuration Table in https://www.bittitan.com/doc/powershell.html#PagePowerShellmspcmd%20
Function Create-MSPC_Endpoint {
    param 
    (      
        [parameter(Mandatory=$true)] [guid]$CustomerOrganizationId,
        [parameter(Mandatory=$false)] [String]$endpointType,
        [parameter(Mandatory=$false)] [String]$endpointName,
        [parameter(Mandatory=$false)] [object]$endpointConfiguration,
        [parameter(Mandatory=$false)] [String]$exportOrImport
    )

    $customerTicket  = Get-BT_Ticket -OrganizationId $customerOrganizationId
    
    if($endpointType -eq "AzureFileSystem"){
        
        ########################################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")

            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg 

            if([string]::IsNullOrEmpty($script:secretKey)) {
                do {
                    $script:secretKey = (Read-Host -prompt "Please enter the Azure storage account access key ").trim()
                }while ($script:secretKey -eq "")
            }

            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $azureFileSystemConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername" = $azureAccountName; #Azure Storage Account Name        
                "AccessKey" = $script:secretKey; #Azure Storage Account SecretKey         
                "ContainerName" = $ContainerName #Container Name
            }
        }
        else {
            $azureFileSystemConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername; #Azure Storage Account Name        
                "AccessKey" = $endpointConfiguration.AccessKey; #Azure Storage Account SecretKey         
                "ContainerName" = $endpointConfiguration.ContainerName #Container Name
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $customerTicket -Name $endpointName -Type $endpointType -Configuration $azureFileSystemConfiguration 

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
            

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message 
        }    
    }
    elseif($endpointType -eq "AzureSubscription"){
           
        ########################################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $script:AzureSubscriptionPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($script:AzureSubscriptionPassword -eq "")

            do {
                $global:btAzureSubscriptionID = (Read-Host -prompt "Please enter the Azure subscription ID").trim()
            }while ($global:btAzureSubscriptionID -eq "")

            $msg = "INFO: Azure subscription ID is '$global:btAzureSubscriptionID'."
            Write-Host $msg
            Log-Write -Message $msg 

            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureSubscriptionConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername" = $adminUsername;     
                "AdministrativePassword" = $script:AzureSubscriptionPassword;         
                "SubscriptionID" = $global:btAzureSubscriptionID
            }
        }
        else {
            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureSubscriptionConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;  
                "AdministrativePassword" = $endpointConfiguration.AdministrativePassword;    
                "SubscriptionID" = $endpointConfiguration.SubscriptionID 
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $customerTicket -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration 

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
            

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message 
        }   
    }
    elseif($endpointType -eq "BoxStorage"){
		########################################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $boxStorageConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.BoxStorageConfiguration' -Property @{  
            }
        }
        else {
            $boxStorageConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.BoxStorageConfiguration' -Property @{  
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $customerTicket -Name $endpointName -Type $endpointType -Configuration $boxStorageConfiguration 

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message 
        }  
    }
    elseif($endpointType -eq "DropBox"){

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $dropBoxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.DropBoxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativePassword" = ""
            }
        }
        else {
            $dropBoxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.DropBoxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativePassword" = ""
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $dropBoxConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }      
    elseif($endpointType -eq "Gmail"){

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $domains = (Read-Host -prompt "Please enter the domain or domains (separated by comma)").trim()
            }while ($domains -eq "")
        
            $msg = "INFO: Domain(s) is (are) '$domains'."
            Write-Host $msg
            Log-Write -Message $msg 
                         
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $GoogleMailboxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleMailboxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $adminUsername;
                "Domains" = $Domains;
            }
        }
        else {
            $GoogleMailboxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleMailboxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;
                "Domains" = $endpointConfiguration.Domains;
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $GoogleMailboxConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }
    elseif($endpointType -eq "GoogleDrive"){

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $domains = (Read-Host -prompt "Please enter the domain or domains (separated by comma)").trim()
            }while ($domains -eq "")
        
            $msg = "INFO: Domain(s) is (are) '$domains'."
            Write-Host $msg
            Log-Write -Message $msg 
                         
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $GoogleDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleDriveConfiguration' -Property @{              
                "AdminEmailAddress" = $adminUsername;
                "Domains" = $Domains;
            }
        }
        else {
            $GoogleDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleDriveConfiguration' -Property @{              
                "AdminEmailAddress" = $endpointConfiguration.AdminEmailAddress;
                "Domains" = $endpointConfiguration.Domains;
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $GoogleDriveConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }
    elseif($endpointType -eq "GoogleDriveCustomerTenant"){

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                Write-host -NoNewline "Please enter the file path to the Google service account credentials using JSON file:"
                $workingDir = "C:\scripts"
                Get-FileName $workingDir

                #Read CSV file
                try {
                    $jsonFileContent = get-content $script:inputFile 
                }
                catch {
                    $msg = "ERROR: Failed to import the CSV file '$script:inputFile'."
                    Write-Host -ForegroundColor Red  $msg
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $msg 
                    Log-Write -Message $_.Exception.Message
                    Return -1    
                } 
            }while ($jsonFileContent -eq "")
        
            $msg = "INFO: The file path to the JSON file is '$script:inputFile'."
            Write-Host $msg
            Log-Write -Message $msg 
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################
          
            $GoogleDriveCustomerTenantConfiguration = New-BT_GSuiteConfiguration -AdministrativeUsername $adminUsername `
                                                              -CredentialsFileName $script:inputFile `
                                                              -Credentials $jsonFileContent.ToString()   

        }
        else {
            $adminUsername = $endpointConfiguration.AdministrativeUsername
            do {
                Write-host -NoNewline "Please enter the file path to the Google service account credentials using JSON file:"
                $workingDir = "C:\scripts"
                Get-FileName $workingDir

                #Read CSV file
                try {
                    $jsonFileContent = get-content $script:inputFile 
                }
                catch {
                    $msg = "ERROR: Failed to import the CSV file '$script:inputFile'."
                    Write-Host -ForegroundColor Red  $msg
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $msg 
                    Log-Write -Message $_.Exception.Message
                    Return -1 
                } 
            }while ($jsonFileContent -eq "")
        
            $msg = "INFO: The file path to the JSON file is '$script:inputFile'."
            Write-Host $msg
            Log-Write -Message $msg 
            $GoogleDriveCustomerTenantConfiguration = New-BT_GoogleDriveCustomerTenantConfiguration -AdministrativeUsername $adminUsername `
                                                              -CredentialsFileName $script:inputFile `
                                                              -Credentials $jsonFileContent.ToString()   
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $GoogleDriveCustomerTenantConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return -1              
        }
    }
    elseif($endpointType -eq "ExchangeOnline2"){
        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $script:o365AdminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $exchangeOnline2Configuration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangeConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $adminUsername;
                "AdministrativePassword" = $script:o365AdminPassword
            }
        }
        else {
            $exchangeOnline2Configuration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangeConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword" = $endpointConfiguration.AdministrativePassword
            }
        }

        try {
            
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $exchangeOnline2Configuration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }

    }
    elseif($endpointType -eq "Office365Groups") {

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $url = (Read-Host -prompt "Please enter the Office 365 group URL").trim()
            }while ($url -eq "")
        
            $msg = "INFO: Office 365 group URL is '$url'."
            Write-Host $msg
            Log-Write -Message $msg 
        
        
            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg 

            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $office365GroupsConfiguration = New-Object -TypeName "ManagementProxy.ManagementService.SharePointConfiguration" -Property @{
                "Url" = $url;
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $adminUsername;
                "AdministrativePassword" = $adminPassword
            }
        }
        else {
            $office365GroupsConfiguration = New-Object -TypeName "ManagementProxy.ManagementService.SharePointConfiguration" -Property @{
                "Url" = $endpointConfiguration.Url;
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword" = $endpointConfiguration.AdministrativePassword
            }
        }

        try {
            
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $office365GroupsConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }
    elseif($endpointType -eq "OneDrivePro"){

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $script:adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($script:adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$script:adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg 
                         
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $oneDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $adminUsername;
                "AdministrativePassword" = $script:adminPassword
            }
        }
        else {
            $oneDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword" = $endpointConfiguration.AdministrativePassword
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $oneDriveConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }
    elseif($endpointType -eq "OneDriveProAPI"){

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 
            
            do {
                $administrativePasswordSecureString = (Read-Host -prompt "Please enter the Office 365 admin password" -AsSecureString)
            }while ($administrativePasswordSecureString -eq "") 
    
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($administrativePasswordSecureString)
            $script:adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            do {
                $confirm = (Read-Host -prompt "Do you want to use your own Azure Storage account?  [Y]es or [N]o")
                if($confirm.ToLower() -eq "y") {
                    $script:microsoftStorage = $false
                }
                else {
                    $script:microsoftStorage = $true
                }
            } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
            
            if(!$script:microsoftStorage) {
                do {
                    $script:dstAzureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
                }while ($script:dstAzureAccountName -eq "")
            
                $msg = "INFO: Azure storage account name is '$script:dstAzureAccountName'."
                Write-Host $msg
                Log-Write -Message $msg 

                do {
                    $secretKeySecureString = (Read-Host -prompt "Please enter the Azure storage account access key" -AsSecureString)
                }while ($secretKeySecureString -eq "")

                $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secretKeySecureString)
                $script:secretKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            }
    
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            if(!$script:microsoftStorage) {
                $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                    "UseAdministrativeCredentials" = $true;
                    "AdministrativeUsername" = $adminUsername;
                    "AdministrativePassword" = $script:adminPassword;
                    "AzureStorageAccountName" = $script:dstAzureAccountName;
                    "AzureAccountKey" = $script:secretKey
                    "UseSharePointOnlineProvidedStorage" = $true
                }
            }
            else {
                $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                    "UseAdministrativeCredentials" = $true;
                    "AdministrativeUsername" = $adminUsername;
                    "AdministrativePassword" = $script:adminPassword;
                    "UseSharePointOnlineProvidedStorage" = $true
                }
            }            
        }
        else {
            $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword" = $endpointConfiguration.AdministrativePassword;
                #"AzureStorageAccountName" = $endpointConfiguration.AzureStorageAccountName;
                #"AzureAccountKey" = $endpointConfiguration.AzureAccountKey
                "UseSharePointOnlineProvidedStorage" = $true
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $oneDriveProAPIConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }
    elseif($endpointType -eq "SharePoint"){
        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg 
                         
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{   
                "Url" = $Url;           
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $adminUsername;
                "AdministrativePassword" = $adminPassword
            }
        }
        else {
            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{  
                "Url" = $endpointConfiguration.Url;             
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword" = $endpointConfiguration.AdministrativePassword
            }
        }

        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $spoConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
    }
    elseif($endpointType -eq "SharePointOnlineAPI"){

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")
        
            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $script:secretKey = (Read-Host -prompt "Please enter the Azure storage account access key").trim()
            }while ($script:secretKey -eq "")
    
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{   
                "Url" = $Url;               
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $adminUsername;
                "AdministrativePassword" = $adminPassword;
                #"AzureStorageAccountName" = $azureAccountName;
                #"AzureAccountKey" = $script:secretKey
                "UseSharePointOnlineProvidedStorage" = $true 
            }
        }
        else {
            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{   
                 "Url" = $endpointConfiguration.Url;              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword" = $endpointConfiguration.AdministrativePassword;
                #"AzureStorageAccountName" = $endpointConfiguration.AzureStorageAccountName;
                #"AzureAccountKey" = $endpointConfiguration.AzureAccountKey
                "UseSharePointOnlineProvidedStorage" = $true 
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $spoConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
 
    }
    elseif($endpointType -eq "MicrosoftTeamsSource"){

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")
        
            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key").trim()
            }while ($secretKey -eq "")
    
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsSourceConfiguration' -Property @{          
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $adminUsername;
                "AdministrativePassword" = $adminPassword;
            }
        }
        else {
            $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsSourceConfiguration' -Property @{             
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword" = $endpointConfiguration.AdministrativePassword;
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $teamsConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
 
    }
    elseif($endpointType -eq "MicrosoftTeamsDestination"){

        ########################################################################################################################################
        # Prompt for endpoint data. 
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")
        
            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key").trim()
            }while ($secretKey -eq "")
        
            $msg = "INFO: Azure storage account access key is '$secretKey'."
            Write-Host $msg
            Log-Write -Message $msg 
    
            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsDestinationConfiguration' -Property @{          
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $adminUsername;
                "AdministrativePassword" = $adminPassword;
                #"AzureStorageAccountName" = $azureAccountName;
                #"AzureAccountKey" = $secretKey
                "UseSharePointOnlineProvidedStorage" = $true 
            }
        }
        else {
            $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsDestinationConfiguration' -Property @{             
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword" = $endpointConfiguration.AdministrativePassword;
                #"AzureStorageAccountName" = $endpointConfiguration.AzureStorageAccountName;
                #"AzureAccountKey" = $endpointConfiguration.AzureAccountKey
                "UseSharePointOnlineProvidedStorage" = $true 
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $teamsConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                
        }
 
    }
    elseif($endpointType -eq "Pst"){

        ########################################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        ########################################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

             do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")

            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg 

            if([string]::IsNullOrEmpty($script:secretKey)) {
                do {
                    $script:secretKey = (Read-Host -prompt "Please enter the Azure storage account access key ").trim()
                }while ($script:secretKey -eq "")
            }


            do {
                $containerName = (Read-Host -prompt "Please enter the container name").trim()
            }while ($containerName -eq "")

            $msg = "INFO: Azure container name is '$containerName'."
            Write-Host $msg
            Log-Write -Message $msg 

            ########################################################################################################################################
            # Create endpoint. 
            ########################################################################################################################################

            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername" = $azureAccountName;     
                "AccessKey" = $script:secretKey;  
                "ContainerName" = $containerName;       
            }
        }
        else {
            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;  
                "AccessKey" = $endpointConfiguration.AccessKey;    
                "ContainerName" = $endpointConfiguration.ContainerName 
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $customerTicket -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration 

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                Return $checkEndpoint.Id
            }
            

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message 
        }  
    }

    <#
        elseif(endpointType -eq "WorkMail"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name WorkMailRegion -Value $endpoint.WorkMailRegion

             
        }
        elseif(endpointType -eq "Zimbra"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            return $endpointCredentials  
        }
        elseif(endpointType -eq "ExchangeOnlinePublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            
        }
        elseif(endpointType -eq "ExchangeOnlineUsGovernment"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            
        }
        elseif(endpointType -eq "ExchangeOnlineUsGovernmentPublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            
        }
        elseif(endpointType -eq "ExchangeServer"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            
        }
        elseif(endpointType -eq "ExchangeServerPublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            
        }  
        elseif(endpointType -eq "GoogleDrive"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdminEmailAddress -Value $endpoint.AdminEmailAddress
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            return $endpointCredentials   
        }
        elseif(endpointType -eq "GoogleVault"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdminEmailAddress -Value $endpoint.AdminEmailAddress
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            return $endpointCredentials   
        }
        elseif(endpointType -eq "GroupWise"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name TrustedAppKey -Value $endpoint.AdministrativePassword

            return $endpointCredentials  
        }
        elseif(endpointType -eq "IMAP"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Host -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Port -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseSsl -Value $endpoint.AdministrativePassword

            return $endpointCredentials  
        }
        elseif(endpointType -eq "Lotus"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ExtractorName -Value $endpoint.UseAdministrativeCredentials

            return $endpointCredentials  
        }  
        elseif(endpointType -eq "PstInternalStorage") {

            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessKey -Value $endpoint.AccessKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ContainerName -Value $endpoint.ContainerName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
   
        }
        elseif(endpointType -eq "OX"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url

        }#>
}

# Function to create a connector under a customer
Function Create-MW_Connector {
    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerOrganizationId,
        [parameter(Mandatory=$true)] [String]$ProjectName,
        [parameter(Mandatory=$true)] [String]$ProjectType,
        [parameter(Mandatory=$true)] [String]$importType,
        [parameter(Mandatory=$true)] [String]$exportType,   
        [parameter(Mandatory=$true)] [guid]$exportEndpointId,
        [parameter(Mandatory=$true)] [guid]$importEndpointId,  
        [parameter(Mandatory=$true)] [object]$exportConfiguration,
        [parameter(Mandatory=$true)] [object]$importConfiguration,
        [parameter(Mandatory=$false)] [String]$advancedOptions,   
        [parameter(Mandatory=$false)] [String]$folderFilter="",
        [parameter(Mandatory=$false)] [String]$maximumSimultaneousMigrations=100,
        [parameter(Mandatory=$false)] [String]$MaxLicensesToConsume=10,
        [parameter(Mandatory=$false)] [int64]$MaximumDataTransferRate,
        [parameter(Mandatory=$false)] [String]$Flags,
        [parameter(Mandatory=$false)] [String]$ZoneRequirement,
        [parameter(Mandatory=$false)] [Boolean]$updateConnector   
        
    )
    try{
        $connector = @(Get-MW_MailboxConnector -ticket $script:MwTicket `
        -UserId $script:MwTicket.UserId `
        -OrganizationId $global:btCustomerOrganizationId `
        -Name "$ProjectName" `
        -ErrorAction SilentlyContinue
        #-SelectedExportEndpointId $exportEndpointId `
        #-SelectedImportEndpointId $importEndpointId `        
        #-ProjectType $ProjectType `
        #-ExportType $exportType `
        #-ImportType $importType `

        ) 

        if($connector.Count -eq 1) {
            $msg = "WARNING: Connector '$($connector.Name)' already exists with the same configuration." 
            write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg 

            if($updateConnector) {
                $connector = Set-MW_MailboxConnector -ticket $script:MwTicket `
                    -MailboxConnector $connector `
                    -Name $ProjectName `
                    -ExportType $exportType `
                    -ImportType $importType `
                    -SelectedExportEndpointId $exportEndpointId `
                    -SelectedImportEndpointId $importEndpointId `
                    -ExportConfiguration $exportConfiguration `
                    -ImportConfiguration $importConfiguration `
                    -AdvancedOptions $advancedOptions `
                    -FolderFilter $folderFilter `
                    -MaximumDataTransferRate ([int]::MaxValue) `
                    -MaximumDataTransferRateDuration 600000 `
                    -MaximumSimultaneousMigrations $maximumSimultaneousMigrations `
                    -PurgePeriod 180 `
                    -MaximumItemFailures 1000 `
                    -ZoneRequirement $ZoneRequirement `
                    -MaxLicensesToConsume $MaxLicensesToConsume  
                    #-Flags $Flags `

                $msg = "SUCCESS: Connector '$($connector.Name)' updated." 
                write-Host -ForegroundColor Blue $msg
                Log-Write -Message $msg 

                return $connector.Id
            }
            else { 
                return $connector.Id 
            }
        }
        elseif($connector.Count -gt 1) {
            $msg = "WARNING: $($connector.Count) connectors '$ProjectName' already exist with the same configuration." 
            write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg 

            return $null

        } else {
            try { 
                $connector = Add-MW_MailboxConnector -ticket $script:MwTicket `
                -UserId $script:MwTicket.UserId `
                -OrganizationId $global:btCustomerOrganizationId `
                -Name $ProjectName `
                -ProjectType $ProjectType `
                -ExportType $exportType `
                -ImportType $importType `
                -SelectedExportEndpointId $exportEndpointId `
                -SelectedImportEndpointId $importEndpointId `
                -ExportConfiguration $exportConfiguration `
                -ImportConfiguration $importConfiguration `
                -AdvancedOptions $advancedOptions `
                -FolderFilter $folderFilter `
                -MaximumDataTransferRate ([int]::MaxValue) `
                -MaximumDataTransferRateDuration 600000 `
                -MaximumSimultaneousMigrations $maximumSimultaneousMigrations `
                -PurgePeriod 180 `
                -MaximumItemFailures 1000 `
                -ZoneRequirement $ZoneRequirement `
                -MaxLicensesToConsume $MaxLicensesToConsume  
                #-Flags $Flags `

                $msg = "SUCCESS: Connector '$($connector.Name)' created." 
                write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                return $connector.Id
            }
            catch{
                $msg = "ERROR: Failed to create mailbox connector '$($connector.Name)'."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg 
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message  
            }
        }
    }
    catch {
        $msg = "ERROR: Failed to get mailbox connector '$($connector.Name)'."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message  
    }

}

# Function to get the ULR to the MSPComplete Customerdashboard
Function Get-CustomerUrlId {
    param 
    (      
        [parameter(Mandatory=$true)] [String]$customerOrganizationId
    )

    $customerUrlId = (Get-BT_Customer -OrganizationId $customerOrganizationId).Id

    Return $customerUrlId

}

# Function to display the workgroups created by the user
Function Select-MSPC_Workgroup {

    #######################################
    # Display all mailbox workgroups
    #######################################

    $workgroupPageSize = 100
  	$workgroupOffSet = 0
	$workgroups = @()

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC workgroups..."

   do {
       try {
            #default workgroup in the 1st position
            $workgroupsPage = @(Get-BT_Workgroup -PageOffset $workgroupOffset -PageSize 1 -IsDeleted false -CreatedBySystemUserId $script:ticket.SystemUserId )
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC workgroups."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }

        if($workgroupsPage) {
            $workgroups += @($workgroupsPage)
        }

        $workgroupOffset += 1
    } while($workgroupsPage)

    $workgroupOffSet = 0

    do { 
        try{
            #add all the workgroups including the default workgroup, so there will be 2 default workgroups
            $workgroupsPage = @(Get-BT_Workgroup -PageOffset $workgroupOffSet -PageSize $workgroupPageSize -IsDeleted false | Where-Object  { $_.CreatedBySystemUserId -ne $script:ticket.SystemUserId })   
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC workgroups."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }
        if($workgroupsPage) {
            $workgroups += @($workgroupsPage)
            foreach($Workgroup in $workgroupsPage) {
                Write-Progress -Activity ("Retrieving workgroups (" + $($workgroups.Length -1) + ")") -Status $Workgroup.Id
            }

            $workgroupOffset += $workgroupPageSize
        }
    } while($workgroupsPage)

    if($workgroups -ne $null -and $workgroups.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $($workgroups.Length -1).ToString() + " Workgroup(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No workgroups found." 
        Exit
    }

    #######################################
    # Prompt for the mailbox Workgroup
    #######################################
    if($workgroups -ne $null)
    {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a Workgroup:" 
        Write-Host -Object "INFO: A default workgroup has no name, only Id. Your default workgroup is the number 0 in yellow." 

        for ($i=0; $i -lt $workgroups.Length; $i++) {
            
            $Workgroup = $workgroups[$i]

            if([string]::IsNullOrEmpty($Workgroup.Name)) {
                if($i -eq 0) {
                    $defaultWorkgroupId = $Workgroup.Id.Guid
                    Write-Host -ForegroundColor Yellow -Object $i,"-",$defaultWorkgroupId
                }
                else {
                    if($Workgroup.Id -ne $defaultWorkgroupId) {
                        Write-Host -Object $i,"-",$Workgroup.Id
                    }
                }
            }
            else {
                Write-Host -Object $i,"-",$Workgroup.Name
            }
        }
        Write-Host -Object "x - Exit"
        Write-Host

        do {
            if($workgroups.count -eq 1) {
                $msg = "INFO: There is only one workgroup. Selected by default."
                Write-Host -ForegroundColor yellow  $msg
                Log-Write -Message $msg
                $Workgroup=$workgroups[0]
                Return $Workgroup.Id
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($workgroups.Length-1) + ", or x")
            }
            
            if($result -eq "x") {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $workgroups.Length)) {
                $Workgroup=$workgroups[$result]
                $global:btWorkgroupOrganizationId = $Workgroup.WorkgroupOrganizationId
                Return $Workgroup.Id
            }
        }
        while($true)

    }

}

# Function to display all customers
Function Select-MSPC_Customer {

    param 
    (      
        [parameter(Mandatory=$true)] [String]$WorkgroupId
    )

    #######################################
    # Display all mailbox customers
    #######################################

    $customerPageSize = 100
  	$customerOffSet = 0
	$customers = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC customers..."

    do
    {   
        try { 
            $customersPage = @(Get-BT_Customer -WorkgroupId $global:btWorkgroupId -IsDeleted False -IsArchived False -PageOffset $customerOffSet -PageSize $customerPageSize)
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC customers."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }
    
        if($customersPage) {
            $customers += @($customersPage)
            foreach($customer in $customersPage) {
                Write-Progress -Activity ("Retrieving customers (" + $customers.Length + ")") -Status $customer.CompanyName
            }
            
            $customerOffset += $customerPageSize
        }

    } while($customersPage)

    

    if($customers -ne $null -and $customers.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $customers.Length.ToString() + " customer(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No customers found." 
        Return "-1"
    }

    #######################################
    # {Prompt for the mailbox customer
    #######################################
    if($customers -ne $null)
    {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a customer:" 

        for ($i=0; $i -lt $customers.Length; $i++)
        {
            $customer = $customers[$i]
            Write-Host -Object $i,"-",$customer.CompanyName
        }
        Write-Host -Object "b - Go back to workgroup selection menu"
        Write-Host -Object "x - Exit"
        Write-Host

        do
        {
            if($customers.count -eq 1) {
                $msg = "INFO: There is only one customer. Selected by default."
                Write-Host -ForegroundColor yellow  $msg
                Log-Write -Message $msg
                $customer=$customers[0]

                try{
                    if($script:confirmImpersonation) {
                        $global:btCustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ImpersonateId $global:btMspcSystemUserId -ErrorAction Stop
                    }
                    else{
                        $global:btCustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ErrorAction Stop
                    }
                }
                Catch{
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket under Select-MSPC_Customer()." 
                }

                $global:btCustomerName = $Customer.CompanyName

                Return $customer
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($customers.Length-1) + ", b or x")
            }

            if($result -eq "b") {
                Return "-1"
            }
            if($result -eq "x") {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $customers.Length)) {
                $customer=$customers[$result]
    
                try{
                    if($script:confirmImpersonation) {
                        $global:btCustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ImpersonateId $global:btMspcSystemUserId -ErrorAction Stop
                    }
                    else{ 
                        $global:btCustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ErrorAction Stop
                    }
                }
                Catch{
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket under Select-MSPC_Customer()." 
                }

                $global:btCustomerName = $Customer.CompanyName

                Return $Customer
            }
        }
        while($true)

    }

}

# Function to display all endpoints under a customer
Function Select-MSPC_Endpoint {
    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerOrganizationId,
        [parameter(Mandatory=$false)] [String]$endpointType,
        [parameter(Mandatory=$false)] [String]$endpointName,
        [parameter(Mandatory=$false)] [object]$endpointConfiguration,
        [parameter(Mandatory=$false)] [String]$exportOrImport,
        [parameter(Mandatory=$false)] [String]$projectType,
        [parameter(Mandatory=$false)] [boolean]$deleteEndpointType

    )

    ########################################################################################################################################
    # Display all MSPC endpoints. If $endpointType is provided, only endpoints of that type
    ########################################################################################################################################

    $endpointPageSize = 100
  	$endpointOffSet = 0
	$endpoints = $null

    $sourceMailboxEndpointList = @("ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment","Gmail","IMAP","GroupWise","zimbra","OX","WorkMail","Lotus","Office365Groups")
    $destinationeMailboxEndpointList = @("ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment","Gmail","IMAP","OX","WorkMail","Office365Groups","Pst")
    $sourceStorageEndpointList = @("OneDrivePro","OneDriveProAPI","SharePoint","SharePointOnlineAPI","GoogleDrive","GoogleDriveCustomerTenant","AzureFileSystem","BoxStorage"."DropBox","Office365Groups")
    $destinationStorageEndpointList = @("OneDrivePro","OneDriveProAPI","SharePoint","SharePointOnlineAPI","GoogleDrive","GoogleDriveCustomerTenant","BoxStorage"."DropBox","Office365Groups")
    $sourceArchiveEndpointList = @("ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment","GoogleVault","PstInternalStorage","Pst")
    $destinationArchiveEndpointList =  @("ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment","Gmail","IMAP","OX","WorkMail","Office365Groups","Pst")
    $sourcePublicFolderEndpointList = @("ExchangeServerPublicFolder","ExchangeOnlinePublicFolder","ExchangeOnlineUsGovernmentPublicFolder")
    $destinationPublicFolderEndpointList = @("ExchangeServerPublicFolder","ExchangeOnlinePublicFolder","ExchangeOnlineUsGovernmentPublicFolder","ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment")

    Write-Host
    if($endpointType -ne "") {
        Write-Host -Object  "INFO: Retrieving MSPC $exportOrImport $endpointType endpoints..."
    }else {
        Write-Host -Object  "INFO: Retrieving MSPC $exportOrImport endpoints..."

        if($projectType -ne "") {
            switch($projectType) {
                "Mailbox" {
                    if($exportOrImport -eq "Source") {
                        $availableEndpoints = $sourceMailboxEndpointList 
                    }
                    else {
                        $availableEndpoints = $destinationeMailboxEndpointList
                    }
                }

                "Storage" {
                    if($exportOrImport -eq "Source") { 
                        $availableEndpoints = $sourceStorageEndpointList
                    }
                    else {
                        $availableEndpoints = $destinationStorageEndpointList
                    }
                }

                "Archive" {
                    if($exportOrImport -eq "Source") {
                        $availableEndpoints = $sourceArchiveEndpointList 
                    }
                    else {
                        $availableEndpoints = $destinationArchiveEndpointList
                    }
                }

                "PublicFolder" {
                    if($exportOrImport -eq "Source") { 
                        $availableEndpoints = $publicfolderEndpointList
                    }
                    else {
                        $availableEndpoints = $publicfolderEndpointList
                    }
                } 
            }          
        }
    }

    $customerTicket = Get-BT_Ticket -OrganizationId $customerOrganizationId

    do {
        try{
            if($endpointType -ne "") {
                $endpointsPage = @(Get-BT_Endpoint -Ticket $customerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize -type $endpointType )
            }else{
                $endpointsPage = @(Get-BT_Endpoint -Ticket $customerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize | Sort-Object -Property Type)
            }
        }

        catch {
            $msg = "ERROR: Failed to retrieve MSPC endpoints."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message 
            Exit
        }

        if($endpointsPage) {
            
            $endpoints += @($endpointsPage)

            foreach($endpoint in $endpointsPage) {
                Write-Progress -Activity ("Retrieving endpoint (" + $endpoints.Length + ")") -Status $endpoint.Name
            }
            
            $endpointOffset += $endpointPageSize
        }
    } while($endpointsPage)

    Write-Progress -Activity " " -Completed

    if($endpoints -ne $null -and $endpoints.Length -ge 1) {
        Write-Host -ForegroundColor Green "SUCCESS: $($endpoints.Length) endpoint(s) found."
    }
    else {
        Write-Host -ForegroundColor Red "INFO: No endpoints found." 
    }

    ########################################################################################################################################
    # Prompt for the endpoint. If no endpoints found and endpointType provided, ask for endpoint creation
    ########################################################################################################################################
    if($endpoints -ne $null) {


        if($endpointType -ne "") {
            
            Write-Host -ForegroundColor Yellow -Object "ACTION: Select the $exportOrImport $endpointType endpoint:" 

            for ($i=0; $i -lt $endpoints.Length; $i++) {
                $endpoint = $endpoints[$i]
                Write-Host -Object $i,"-",$endpoint.Name
            }
        }
        elseif($endpointType -eq "" -and $projectType -ne "") {

            Write-Host -ForegroundColor Yellow -Object "ACTION: Select the $exportOrImport $projectType endpoint:" 

           for ($i=0; $i -lt $endpoints.Length; $i++) {
                $endpoint = $endpoints[$i]
                if($endpoint.Type -in $availableEndpoints) {
                    
                    Write-Host $i,"- Type: " -NoNewline 
                    Write-Host -ForegroundColor White $endpoint.Type -NoNewline                      
                    Write-Host "- Name: " -NoNewline                    
                    Write-Host -ForegroundColor White $endpoint.Name   
                }
            }
        }


        Write-Host -Object "c - Create a new $endpointType endpoint"
        Write-Host -Object "x - Exit"
        Write-Host

        do
        {
            if($endpoints.count -eq 1) {
                $result = Read-Host -Prompt ("Select 0, c or x")
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($endpoints.Length-1) + ", c or x")
            }
            
            if($result -eq "c") {
                if ($endpointName -eq "") {
                
                    if($endpointConfiguration  -eq $null) {
                        $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType                     
                    }
                    else {
                        $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration          
                    }        
                }
                else {
                    $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
                }
                Return $endpointId
            }
            elseif($result -eq "x") {
                Exit
            }
            elseif(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $endpoints.Length)) {
                $endpoint=$endpoints[$result]
                Return $endpoint.Id
            }
        }
        while($true)

    } 
    elseif($endpoints -eq $null -and $endpointType -ne "") {

        do {
            $confirm = (Read-Host -prompt "Do you want to create a $endpointType endpoint ?  [Y]es or [N]o")
            if($confirm.ToLower() -eq "y") {
                if ($endpointName -eq "") {
                    $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration 
                }
                else {
                    $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
                }
                Return $endpointId
            }
            elseif($confirm.ToLower() -eq "n") {
                Return -1
            }
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    }
}

# Function to get endpoint data
Function Get-MSPC_EndpointData {
    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerOrganizationId,
        [parameter(Mandatory=$true)] [guid]$endpointId
    )

    $global:btCustomerTicket  = Get-BT_Ticket -OrganizationId $global:btCustomerOrganizationId

    try {
        $endpoint = Get-BT_Endpoint -Ticket $global:btCustomerTicket -Id $endpointId -IsDeleted False -IsArchived False | Select-Object -Property Name, Type -ExpandProperty Configuration   
        
        $msg = "SUCCESS: Endpoint '$($endpoint.Name)' '$endpointId' retrieved." 
        write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg  

        if($endpoint.Type -eq "AzureFileSystem") {

            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessKey -Value $endpoint.AccessKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ContainerName -Value $endpoint.ContainerName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername

            return $endpointCredentials        
        }
        elseif($endpoint.Type -eq "AzureSubscription"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name SubscriptionID -Value $endpoint.SubscriptionID

            return $endpointCredentials
        
        } 
        elseif($endpoint.Type -eq "BoxStorage"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessToken -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name RefreshToken -Value $administrativePassword
            return $endpointCredentials
        }
        elseif($endpoint.Type -eq "DropBox"){
            $endpointCredentials = New-Object PSObject
            return $endpointCredentials
        }
        elseif($endpoint.Type -eq "ExchangeOnline2"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif($endpoint.Type -eq "ExchangeOnlinePublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif($endpoint.Type -eq "ExchangeOnlineUsGovernment"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif($endpoint.Type -eq "ExchangeOnlineUsGovernmentPublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif($endpoint.Type -eq "ExchangeServer"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif($endpoint.Type -eq "ExchangeServerPublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif($endpoint.Type -eq "Gmail"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            if($script:userMailboxesWithResourceMailboxes -or $script:resourceMailboxes) {
                Export-GoogleResources $endpoint.UseAdministrativeCredentials
            }
            
            return $endpointCredentials   
        }
        elseif($endpoint.Type -eq "GSuite"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name CredentialsFileName -Value $endpoint.CredentialsFileName

            if($script:userMailboxesWithResourceMailboxes -or $script:resourceMailboxes) {
                Export-GoogleResources $endpoint.UseAdministrativeCredentials
            }
            
            return $endpointCredentials   
        }
        elseif($endpoint.Type -eq "GoogleDrive"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdminEmailAddress -Value $endpoint.AdminEmailAddress
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            return $endpointCredentials   
        }
        elseif($endpoint.Type -eq "GoogleDriveCustomerTenant"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name CredentialsFileName -Value $endpoint.CredentialsFileName

            return $endpointCredentials   
        }
        elseif($endpoint.Type -eq "GoogleVault"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdminEmailAddress -Value $endpoint.AdminEmailAddress
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            return $endpointCredentials   
        }
        elseif($endpoint.Type -eq "GroupWise"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name TrustedAppKey -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif($endpoint.Type -eq "IMAP"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Host -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Port -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseSsl -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif($endpoint.Type -eq "Lotus"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ExtractorName -Value $endpoint.UseAdministrativeCredentials

            $msg = "INFO: Extractor Name '$($endpoint.ExtractorName)'." 
            write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg  

            return $endpointCredentials  
        }
        elseif($endpoint.Type -eq "Office365Groups"){
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials     
        }
        elseif($endpoint.Type -eq "OneDrivePro"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials   
        }
        elseif($endpoint.Type -eq "OneDriveProAPI"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureStorageAccountName -Value $endpoint.AzureStorageAccountName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureAccountKey -Value $azureAccountKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseSharePointOnlineProvidedStorage -Value $endpoint.UseSharePointOnlineProvidedStorage

            return $endpointCredentials   
        }
        elseif($endpoint.Type -eq "OX"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url

            return $endpointCredentials  
        }
        elseif($endpoint.Type -eq "Pst") {

            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessKey -Value $endpoint.AccessKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ContainerName -Value $endpoint.ContainerName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials

            return $endpointCredentials        
        }
        elseif($endpoint.Type -eq "PstInternalStorage") {

            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessKey -Value $endpoint.AccessKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ContainerName -Value $endpoint.ContainerName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials

            return $endpointCredentials        
        }
        elseif($endpoint.Type -eq "SharePoint"){
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials     
        }
        elseif($endpoint.Type -eq "SharePointOnlineAPI"){
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureStorageAccountName -Value $endpoint.AzureStorageAccountName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureAccountKey -Value $azureAccountKey

            return $endpointCredentials     
        }
        elseif($endpoint.Type -eq "MicrosoftTeamsSource"){
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials     
        }
        elseif($endpoint.Type -eq "MicrosoftTeamsDestination"){
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureStorageAccountName -Value $endpoint.AzureStorageAccountName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureAccountKey -Value $azureAccountKey

            return $endpointCredentials     
        }
        elseif($endpoint.Type -eq "WorkMail"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name WorkMailRegion -Value $endpoint.WorkMailRegion

            return $endpointCredentials  
        }
        elseif($endpoint.Type -eq "Zimbra"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        else{
            Return -1
        }

    }
    catch {
        $msg = "ERROR: Failed to retrieve endpoint '$($endpoint.Name)' credentials."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
    }
}


######################################################################################################################################
#                                    CONNECTION TO AZURE
######################################################################################################################################

 # Function to connect to Azure
Function Connect-Azure{
    param(
        [Parameter(Mandatory=$true)] [PSObject]$azureCredentials,
        [Parameter(Mandatory=$false)] [String]$subscriptionID
    )

    $msg = "INFO: Connecting to Azure to create a blob container."
    Write-Host $msg
    Log-Write -Message $msg   

    Load-Module ("AzureRm")

    Try {
        if($subscriptionID -eq $null) {
            $result = Login-AzureRMAccount  -Environment "AzureCloud" -ErrorAction Stop -Credential $azureCredentials
        }
        else {
            $result = Login-AzureRMAccount -Environment "AzureCloud" -SubscriptionId $subscriptionID -ErrorAction Stop -Credential $azureCredentials
        }
    }
    catch {
        $msg = "ERROR: Failed to connect to Azure. You must use multi-factor authentication to access Azure subscription '$subscriptionID'."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        
        Try {
            if($subscriptionID -eq $null) {
                $result = Login-AzureRMAccount  -Environment "AzureCloud" -ErrorAction Stop 
            }
            else {
                $result = Login-AzureRMAccount -Environment "AzureCloud" -SubscriptionId $subscriptionID -ErrorAction Stop 
            }
        }
        catch {
            $msg = "ERROR: Failed to connect to Azure. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg   
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message   
            Exit
        }
    }

    try {

        #If there are multiple Azure Subscriptions in the tenant ensure the current context is set correctly.
        $result = Get-AzureRmSubscription -SubscriptionID $subscriptionID | Set-AzureRmContext

        $azureAccount = (Get-AzureRmContext).Account.Id
        $subscriptionName = (Get-AzureRmSubscription -SubscriptionID $subscriptionID).Name
        $msg = "SUCCESS: Connection to Azure: Account: $azureAccount Subscription: '$subscriptionName'."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg   

    }
    catch {
        $msg = "ERROR: Failed to get the Azure subscription. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message   
        Exit
    }
}

 # Function to check if AzureRM is installed
Function Check-AzureRM {
     Try {
        $result = get-module -ListAvailable -name AzureRM -ErrorAction Stop
        if ($result) {
            $msg = "INFO: Ready to execute Azure PowerShell module $($result.moduletype), $($result.version), $($result.name)"
            Write-Host $msg
            Log-Write -Message $msg 
        }
        Else {
            $msg = "INFO: AzureRM module is not installed."
            Write-Host -ForegroundColor Red $msg
            Log-Write -Message $msg 

            Install-Module AzureRM
            Import-Module AzureRM

            Try {
                
                $result = get-module -ListAvailable -name AzureRM -ErrorAction Stop
                
                If ($result) {
                    write-information "INFO: Ready to execute PowerShell module $($result.moduletype), $($result.version), $($result.name)"
                }
                Else {
                    $msg = "ERROR: Failed to install and import the AzureRM module. Script aborted."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg    
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $_.Exception.Message 
                    Exit
                }
            }
            Catch {
                $msg = "ERROR: Failed to check if the AzureRM module is installed. Script aborted."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg    
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message 
                Exit
            }
        }

    }
    Catch {
        $msg = "ERROR: Failed to check if the AzureRM module is installed. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg    
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message 
        Exit
    } 
}

# Function to check if a blob container exists
Function Check-BlobContainer{
    param 
    (      
        [parameter(Mandatory=$true)] [String]$blobContainerName,
        [parameter(Mandatory=$true)] [PSObject]$storageAccount
    )   

    try {
        $result = Get-AzureStorageContainer -Name $blobContainerName -Context $storageAccount.Context -ErrorAction SilentlyContinue

        if($result){
            $msg = "SUCCESS: Blob container '$($blobContainerName)' found under the Storage account '$($storageAccount.StorageAccountName)'."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg   
            Return $true
        }
        else {
            Return $false
        }
    }
    catch {
        $msg = "ERROR: Failed to get the blob container '$($blobContainerName)' under the Storage account '$($storageAccount.StorageAccountName)'. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg    
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message 
        Exit

    }
}

# Function to check if a StorageAccount exists
Function Check-StorageAccount{
    param 
    (      
        [parameter(Mandatory=$true)] [String]$storageAccountName,
        [parameter(Mandatory=$true)] [String]$userPrincipalName
    )   

    try {
        $storageAccount = Get-AzureRmStorageAccount -ErrorAction Stop |? {$_.StorageAccountName -eq $storageAccountName}
        $resourceGroupName = $storageAccount.ResourceGroupName

        $msg = "SUCCESS: Azure storage account '$storageAccountName' found."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg   

        $msg = "SUCCESS: Resource Group Name '$resourceGroupName' found."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg   

    }
    catch {
        $msg = "ERROR: Failed to find the Azure storage account '$storageAccountName'. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message 
        Exit

    }

    try {
        $userId = Get-AzureRmADUser -UserPrincipalName $userPrincipalName | Select ID

        if($userId) {
            $msg = "SUCCESS: User Id '$($userId.Id)' retrieved for '$userPrincipalName' found."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg  
        }
        else{
            $msg = "ERROR: Failed to get ObjectId with Get-AzureRmADUser for user '$userPrincipalName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 

            Return $storageAccount
        }
    }
    catch {
        $msg = "ERROR: Failed to get ObjectId with Get-AzureRmADUser for user '$userPrincipalName'."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message   
        
        Return $storageAccount
    }

    try {

        $result = New-AzureRmRoleAssignment -ResourceGroupName $resourceGroupName -RoleDefinitionName "Contributor" -ObjectID $userId.Id -ErrorAction SilentlyContinue

        $msg = "INFO: Registering all Azure resource providers for $userPrincipalName to allow non-subscription administrator to create new resources."
        Write-Host $msg
        Log-Write -Message $msg   
    }
    catch {
        $msg = "ERROR: Failed to register all Azure resource providers for '$userPrincipalName' with ObjectId $($userId.Id). Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message   
        Exit

    }

    try {

        $resutl = Get-AzureRmResourceProvider -ListAvailable | Where-Object { $_.RegistrationState -eq 'NotRegistered'} | Select-Object ProviderNamespace | Foreach-Object { Register-AzureRmResourceProvider -ProviderName $_.ProviderNamespace }

        $msg = "SUCCESS: All Azure resource providers registered for '$userPrincipalName'."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg  
    }
    catch {
        $msg = "ERROR: Failed to get the Azure Resource Provider. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message   
        Exit

    }

    if($storageAccount){
        Return $storageAccount
    }
    else {
        Return $false
    }
}

# Function to create a Blob Container
Function Create-BlobContainer{
    param 
    (      
        [parameter(Mandatory=$true)] [String]$blobContainerName,
        [parameter(Mandatory=$true)] [PSObject]$storageAccount
    )   

    try {
        $result = New-AzureStorageContainer -Name $blobContainerName -Context $storageAccount.Context -ErrorAction Stop

        $msg = "SUCCESS: Blob container '$($blobContainerName)' created under the Storage account '$($storageAccount.StorageAccountName)'."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg    
    }
    catch {
        $msg = "ERROR: Failed to create blob container '$($blobContainerName)' under the Storage account '$($storageAccount.StorageAccountName)'."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg    
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message 
        
        $msg = "ACTION: Create blob container '$($blobContainerName)' under the Storage account '$($storageAccount.StorageAccountName)'. Script aborted."
        Write-Host -ForegroundColor Yellow  $msg

        Exit

    }

}

# Function to create a SAS Token
Function Create-SASToken {
    param 
    (      
        [parameter(Mandatory=$true)] [String]$blobContainerName,
        [parameter(Mandatory=$true)] [String]$BlobName,
        [parameter(Mandatory=$true)] [PSObject]$storageAccount
    )   

    $StartTime = Get-Date
    $EndTime = $startTime.AddHours(8760.0) #1 year

    # Read access - https://docs.microsoft.com/en-us/powershell/module/azure.storage/new-azurestoragecontainersastoken
    $SasToken = New-AzureStorageContainerSASToken -Name $blobContainerName `
    -Context $storageAccount.Context -Permission rl -StartTime $StartTime -ExpiryTime $EndTime
    $SasToken | clip

    # Construnct the URL & Test
    $url = "$($storageAccount.Context.BlobEndPoint)$($blobContainerName)/$($BlobName)$($SasToken)"
    $url | clip

    Return $url
}



Function Load-Module ($m) {

    # If module is imported say that and do nothing
    if (Get-Module | Where-Object {$_.Name -eq $m}) {
        write-host "INFO: Module $m is already imported."
    }
    else {

        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $m}) {
            Import-Module $m -Verbose
        }
        else {

            # If module is not imported, not available on disk, but is in online gallery then install and import
            if (Find-Module -Name $m | Where-Object {$_.Name -eq $m}) {
                Install-Module -Name $m -Force -Verbose -Scope CurrentUser
                Import-Module $m -Verbose
            }
            else {

                # If module is not imported, not available and not in online gallery then abort
                write-host -ForegroundColor Red "ERROR: Module $m not imported, not available and not in online gallery, exiting."
                EXIT 1
            }
        }
    }
}

function Invoke-Command() {
    param ( [string]$program = $(throw "Please specify a program" ),
            [string]$argumentString = "",
            [switch]$waitForExit )

    $psi = new-object "Diagnostics.ProcessStartInfo"
    $psi.FileName = $program 
    $psi.Arguments = $argumentString
    $proc = [Diagnostics.Process]::Start($psi)
    if ( $waitForExit ) {
        $proc.WaitForExit();
    }
}

#######################################################################################################################
#                                               MAIN PROGRAM
#######################################################################################################################
Import-PowerShellModules
Import-MigrationWizPowerShellModule

#######################################################################################################################
#                   CUSTOMIZABLE VARIABLES  
#######################################################################################################################
$script:srcGermanyCloud = $false
$script:srcUsGovernment = $False

$script:dstGermanyCloud = $False
$script:dstUsGovernment = $false
                        
$ZoneRequirement1  = "NorthAmerica"   #North America (Virginia). For Azure: Both AZNAE and AZNAW.
$ZoneRequirement2  = "WesternEurope"  #Western Europe (Amsterdam for Azure, Ireland for AWS). For Azure: AZEUW.
$ZoneRequirement3  = "AsiaPacific"    #Asia Pacific (Singapore). For Azure: AZSEA
$ZoneRequirement4  = "Australia"      #Australia (Asia Pacific Sydney). For Azure: AZAUE - NSW.
$ZoneRequirement5  = "Japan"          #Japan (Asia Pacific Tokyo). For Azure: AZJPE - Saltiama.
$ZoneRequirement6  = "SouthAmerica"   #South America (Sao Paolo). For Azure: AZSAB.
$ZoneRequirement7  = "Canada"         #Canada. For Azure: AZCAD.
$ZoneRequirement8  = "NorthernEurope" #Northern Europe (Dublin). For Azure: AZEUN.
$ZoneRequirement9  = "China"          #China.
$ZoneRequirement10 = "France"         #France.
$ZoneRequirement11 = "SouthAfrica"    #South Africa.

$ZoneRequirement = $ZoneRequirement1

#######################################################################################################################
#                       SELECT WORKING DIRECTORY  
#######################################################################################################################

Write-Host
Write-Host
Write-Host -ForegroundColor Yellow "             BitTitan File Server to Google Drive migration project creation tool."
Write-Host

write-host 
$msg = "#######################################################################################################################`
                       SELECT WORKING DIRECTORY                  `
#######################################################################################################################"
Write-Host $msg
write-host 

#Working Directorys
$script:workingDir = "C:\scripts"

if([string]::IsNullOrEmpty($WorkingDirectory)){
    if(!$global:btCheckWorkingDirectory) {
        do {
            $confirm = (Read-Host -prompt "The default working directory is '$script:workingDir'. Do you want to change it?  [Y]es or [N]o")
            if($confirm.ToLower() -eq "y") {
                #Working Directory
                $script:workingDir = [environment]::getfolderpath("desktop")
                Get-Directory $script:workingDir            
            }

            $global:btCheckWorkingDirectory = $true

        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    }
}
else {
    $script:workingDir = $WorkingDirectory
}

#Logs directory
$logDirName = "LOGS"
$logDir = "$script:workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format "yyyyMMddTHHmmss")_Create-FileServerToGoogleDriveMWMigrationProjects.log"
$script:logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $script:workingDir -logDir $logDir

Write-Host 
Write-Host -ForegroundColor Yellow "WARNING: Minimal output will appear on the screen." 
Write-Host -ForegroundColor Yellow "         Please look at the log file '$($script:logFile)'."
Write-Host -ForegroundColor Yellow "         Generated CSV file will be in folder '$($script:workingDir)'."
Write-Host 
Start-Sleep -Seconds 1

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg 

#######################################################################################################################
#         CONNECTION TO YOUR BITTITAN ACCOUNT 
#######################################################################################################################

write-host 
$msg = "#######################################################################################################################`
                       CONNECTION TO YOUR BITTITAN ACCOUNT                  `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "CONNECTION TO YOUR BITTITAN ACCOUNT" 
write-host 

Connect-BitTitan

write-host 
$msg = "#######################################################################################################################`
                       AZURE CLOUD SELECTION                 `
#######################################################################################################################"
Write-Host $msg
Write-Host

if($script:srcGermanyCloud) {
    Write-Host -ForegroundColor Magenta "WARNING: Connecting to (source) Azure Germany Cloud." 

    Write-Host
    do {
        $confirm = (Read-Host -prompt "Do you want to switch to (destination) Azure Cloud (global service)?  [Y]es or [N]o")  
        if($confirm.ToLower() -eq "y") {
            $script:srcGermanyCloud = $false
        }  
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    Write-Host 
}
elseif($script:srcUsGovernment ){
    Write-Host -ForegroundColor Magenta "WARNING: Connecting to (source) Azure Goverment Cloud." 

    Write-Host
    do {
        $confirm = (Read-Host -prompt "Do you want to switch to (destination) Azure Cloud (global service)?  [Y]es or [N]o")  
        if($confirm.ToLower() -eq "y") {
            $script:srcUsGovernment = $false
        }  
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    Write-Host 
}

if($script:dstGermanyCloud) {
    Write-Host -ForegroundColor Magenta "WARNING: Connecting to (destination) Azure Germany Cloud." 

    Write-Host
    do {
        $confirm = (Read-Host -prompt "Do you want to switch to (destination) Azure Cloud (global service)?  [Y]es or [N]o")  
        if($confirm.ToLower() -eq "y") {
            $script:dstGermanyCloud = $false
        }  
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    Write-Host 
}
elseif($script:dstUsGovernment){
    Write-Host -ForegroundColor Magenta "WARNING: Connecting to (destination) Azure Goverment Cloud." 

    Write-Host
    do {
        $confirm = (Read-Host -prompt "Do you want to switch to (destination) Azure Cloud (global service)?  [Y]es or [N]o")  
        if($confirm.ToLower() -eq "y") {
            $script:dstUsGovernment = $false
        }  
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    Write-Host 
}

Write-Host -ForegroundColor Green "INFO: Using Azure $ZoneRequirement Datacenter." 

if([string]::IsNullOrEmpty($BitTitanAzureDatacenter)){
    if(!$global:btCheckAzureDatacenter) {
        Write-Host
        do {
            $confirm = (Read-Host -prompt "Do you want to switch the Azure Datacenter to another region?  [Y]es or [N]o")  
            if($confirm.ToLower() -eq "y") {
                do{
                    $ZoneRequirementNumber = (Read-Host -prompt "`
        1. NorthAmerica   #North America (Virginia). For Azure: Both AZNAE and AZNAW.
        2. WesternEurope  #Western Europe (Amsterdam for Azure, Ireland for AWS). For Azure: AZEUW.
        3. AsiaPacific    #Asia Pacific (Singapore). For Azure: AZSEA
        4. Australia      #Australia (Asia Pacific Sydney). For Azure: AZAUE - NSW.
        5. Japan          #Japan (Asia Pacific Tokyo). For Azure: AZJPE - Saltiama.
        6. SouthAmerica   #South America (Sao Paolo). For Azure: AZSAB.
        7. Canada         #Canada. For Azure: AZCAD.
        8. NorthernEurope #Northern Europe (Dublin). For Azure: AZEUN.
        9. China          #China.
        10. France         #France.
        11. SouthAfrica    #South Africa.

        Select 0-11")
                    switch ($ZoneRequirementNumber) {
                                1 {  $ZoneRequirement = 'NorthAmerica'   }
                                2 {  $ZoneRequirement = 'WesternEurope'  }
                                3 {  $ZoneRequirement = 'AsiaPacific'    }
                                4 {  $ZoneRequirement = 'Australia'      }
                                5 {  $ZoneRequirement = 'Japan'          }
                                6 {  $ZoneRequirement = 'SouthAmerica'   }
                                7 {  $ZoneRequirement = 'Canada'         }
                                8 {  $ZoneRequirement = 'NorthernEurope' }
                                9 {  $ZoneRequirement = 'China'          }
                                10 {  $ZoneRequirement = 'France'         }
                                11 {  $ZoneRequirement = 'SouthAfrica'    }
                            }
                } while(!(isNumeric($ZoneRequirementNumber)) -or !($ZoneRequirementNumber -in 1..11))

                Write-Host 
                Write-Host -ForegroundColor Yellow "WARNING: Now using Azure $ZoneRequirement Datacenter." 

                $global:checkAzureDatacenter = $false
            }  
            if($confirm.ToLower() -eq "n") {
                $global:btCheckAzureDatacenter = $true
            }
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    }
    else{
        $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different Azure datacenter."
        Write-Host -ForegroundColor Yellow $msg
    }
}
else{
    $ZoneRequirement = $BitTitanAzureDatacenter
}

write-host 
$msg = "#######################################################################################################################`
                       WORKGROUP AND CUSTOMER SELECTION              `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "WORKGROUP AND CUSTOMER SELECTION"   

if(-not [string]::IsNullOrEmpty($BitTitanWorkgroupId) -and -not [string]::IsNullOrEmpty($BitTitanCustomerId)){
    $global:btWorkgroupId = $BitTitanWorkgroupId
    # To apply UMB licenses
    $global:btWorkgroupOrganizationId = (Get-BT_Workgroup -ticket $script:ticket -id $global:btWorkgroupId).WorkgroupOrganizationId.Guid
    $global:btCustomerOrganizationId = $BitTitanCustomerId
    
    Write-Host
    $msg = "INFO: Selected workgroup '$global:btWorkgroupId' and customer '$global:btCustomerOrganizationId'."
    Write-Host -ForegroundColor Green $msg
}
else{
    if(!$global:btCheckCustomerSelection -or !$global:btWorkgroupId -or !$global:btWorkgroupOrganizationId -or !$global:btCustomerOrganizationId) {
        do {
            #Select workgroup
            $global:btWorkgroupId = Select-MSPC_WorkGroup

            Write-Host
            $msg = "INFO: Selected workgroup '$global:btWorkgroupId'."
            Write-Host -ForegroundColor Green $msg

            Write-Progress -Activity " " -Completed

            #Select customer
            $customer = Select-MSPC_Customer -WorkgroupId $global:btWorkgroupId

            $global:btCustomerOrganizationId = $customer.OrganizationId.Guid

            Write-Host
            $msg = "INFO: Selected customer '$($customer.CompanyName)'."
            Write-Host -ForegroundColor Green $msg

            Write-Progress -Activity " " -Completed
        }
        while ($customer -eq "-1")
        
        $global:btCheckCustomerSelection = $true  
    }
    else{
        Write-Host
        $msg = "INFO: Already selected workgroup '$global:btWorkgroupId' and customer '$global:btCustomerOrganizationId'."
        Write-Host -ForegroundColor Green $msg

        Write-Host
        $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different workgroups/customers."
        Write-Host -ForegroundColor Yellow $msg

    }
}

#Create a ticket for project sharing
try{
    $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -WorkgroupId $global:btWorkgroupId -IncludeSharedProjects
}
catch{
    $msg = "ERROR: Failed to create MigrationWiz ticket for project sharing. Script aborted."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg 
    Exit
}
#Create a ticket for workgroup ticket to apply UMB
try{
    $global:btWorkgroupTicket  = Get-BT_Ticket -Ticket $script:ticket -OrganizationId $global:btWorkgroupOrganizationId
}
catch{
    $msg = "ERROR: Failed to create #Create a ticket for workgroup ticket to apply UMB. Script aborted."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg 
    Exit
}

write-host 
$msg = "#######################################################################################################################`
                       UPLOADERWIZ DOWNLOAD AND UNZIPPING              `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "AZURE AND PST ENDPOINT SELECTION"   
Write-Host

$url = "https://api.bittitan.com/secure/downloads/UploaderWiz.zip"   
$outFile = "$PSScriptRoot\UploaderWiz.zip" 
$path = "$PSScriptRoot\UploaderWiz"

$downloadUploaderWiz = $false

if([string]::IsNullOrEmpty($downloadLatestVersion)) {
    if(!$global:btCheckPath) {
        $checkPath = Test-Path $outFile 
        if($checkPath) {
            $lastWriteTime = (get-Item -Path $path).LastWriteTime

            do {
                $confirm = (Read-Host -prompt "UploaderWiz was downloaded on $lastWriteTime. Do you want to download it again?  [Y]es or [N]o")

                if($confirm.ToLower() -eq "y") {
                    $downloadUploaderWiz = $true
                }
                elseif($confirm.ToLower() -eq "n"){
                    $global:btCheckPath = $true
                }

            } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
        }else {
            $downloadUploaderWiz = $true
        }
    }
    else{
        $checkPath = Test-Path $outFile 
        if($checkPath) {
            $lastWriteTime = (get-Item -Path $path).LastWriteTime
            $msg = "INFO: UploaderWiz was downloaded on $lastWriteTime."
            Write-Host -ForegroundColor Green $msg

            Write-Host
            $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to download it again."
            Write-Host -ForegroundColor Yellow $msg
        }
        else{
            $downloadUploaderWiz = $true   
        }
    }
}
else{    
    $checkPath = Test-Path $outFile 
    if($checkPath) {
        $lastWriteTime = (get-Item -Path $path).LastWriteTime
        $msg = "INFO: UploaderWiz was downloaded on $lastWriteTime."
        Write-Host -ForegroundColor Green $msg

        $downloadUploaderWiz = $downloadLatestVersion
    }
    else{
        $downloadUploaderWiz = $true   
    }    
}

if($downloadUploaderWiz) {    
    Download-File -Url $url -OutFile $outFile
    $global:btCheckPath = $true
}

write-host 
$msg = "#######################################################################################################################`
                       AZURE AND FILE SYSTEM ENDPOINT SELECTION              `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "AZURE AND FILE SYSTEM ENDPOINT SELECTION"   
Write-Host

if([string]::IsNullOrEmpty($AzureStorageAccessKey)) {
    $msg = "INFO: Getting the connection information to the Azure Storage Account."
    Write-Host $msg
    Log-Write -Message $msg   

    $skipAzureCheck = $false
    if(!$global:btAzureCredentials) {
        #Select source endpoint
        $azureSubscriptionEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointType "AzureSubscription"

        if($azureSubscriptionEndpointId -eq "-1") {    
            do {
            $confirm = (Read-Host -prompt "Do you want to skip the Azure Check ?  [Y]es or [N]o")
                if($confirm.ToLower() -eq "n") {
                    $skipAzureCheck = $false    
        
                    Write-Host
                    $msg = "ACTION: Provide the following credentials that cannot be retrieved from endpoints:"
                    Write-Host -ForegroundColor Yellow $msg
                    Log-Write -Message $msg 
        
                    Write-Host
                    do {
                        $administrativeUsername = (Read-Host -prompt "Please enter the Azure account email address")
                    }while ($administrativeUsername -eq "")
                }
                if($confirm.ToLower() -eq "y") {
                    $skipAzureCheck = $true
                }    
            } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))    
        }
        else {
            $skipAzureCheck = $false
        }
    }
    else{
        Write-Host
        $msg = "INFO: Already selected 'AzureSubscription' endpoint '$azureSubscriptionEndpointId'."
        Write-Host -ForegroundColor Green $msg

        Write-Host
        $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to 'AzureSubscription'."
        Write-Host -ForegroundColor Yellow $msg
    }

    if(!$skipAzureCheck) {
        if(!$global:btAzureCredentials) {
            #Get source endpoint credentials
            [PSObject]$azureSubscriptionEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointId $azureSubscriptionEndpointId 

            #Create a PSCredential object to connect to Azure Active Directory tenant
            $administrativeUsername = $azureSubscriptionEndpointData.AdministrativeUsername
            
            if(!$script:AzureSubscriptionPassword) {
                Write-Host
                $msg = "ACTION: Provide the following credentials that cannot be retrieved from endpoints:"
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                do {
                    $AzureAccountPassword = (Read-Host -prompt "Please enter the Azure Account Password" -AsSecureString)
                }while ($AzureAccountPassword -eq "")
            }
            else{
                $AzureAccountPassword = $script:AzureSubscriptionPassword
            }

            $global:btAzureCredentials = New-Object System.Management.Automation.PSCredential ($administrativeUsername, $AzureAccountPassword)
        }

        if(!$global:btAzureSubscriptionID) {
            do {
                $global:btAzureSubscriptionID = (Read-Host -prompt "Please enter the Azure Subscription ID").trim()
            }while ($global:btAzureSubscriptionID -eq "")
        }

        if(!$script:secretKey) {
            Write-Host
            do {
                $script:secretKeySecureString = (Read-Host -prompt "Please enter the Azure Storage Account Primary Access Key" -AsSecureString)
        
                $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($script:secretKeySecureString)
                $script:secretKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        
            }while ($script:secretKey -eq "")
        }
    }
}
else{
    $script:secretKey = $AzureStorageAccessKey
}

if(!$global:btExportEndpointId){
    if([string]::IsNullOrEmpty($BitTitanSourceEndpointId)){
        #Select source endpoint
        $global:btExportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType "AzureFileSystem"
    } 
    else{
        $global:btExportEndpointId = $BitTitanSourceEndpointId
    }   
}
else{
    Write-Host
    $msg = "INFO: Already selected 'AzureFileSystem' endpoint '$global:btExportEndpointId'."
    Write-Host -ForegroundColor Green $msg

    Write-Host
    $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different 'AzureFileSystem'."
    Write-Host -ForegroundColor Yellow $msg
    Write-Host 
}

#Get source endpoint credentials
[PSObject]$exportEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointId $global:btExportEndpointId 

if([string]::IsNullOrEmpty($AzureStorageAccessKey)) {
if(!$skipAzureCheck) {
    write-host 
    $msg = "#######################################################################################################################`
                           CONNECTION TO YOUR AZURE ACCOUNT                  `
#######################################################################################################################"
    Write-Host $msg
    Log-Write -Message "CONNECTION TO YOUR AZURE ACCOUNT" 
    write-host 

    if(!$global:btAzureStorageAccountChecked) {
        $msg = "INFO: Checking the Azure Storage Account."
        Write-Host $msg
        Log-Write -Message $msg 
        Write-Host

        # AzureRM module installation
        Check-AzureRM
        # Azure log in
        if($azureSubscriptionEndpointData.SubscriptionID){
            Connect-Azure -AzureCredentials $global:btAzureCredentials -SubscriptionID $azureSubscriptionEndpointData.SubscriptionID
        }
        elseif($global:btAzureSubscriptionID){
            Connect-Azure -AzureCredentials $global:btAzureCredentials -SubscriptionID $global:btAzureSubscriptionID
        }
        else{
            $msg = "ERROR: Wrong Azure Subscription ID provided."
            Write-Host $msg -ForegroundColor Red
            Log-Write -Message $msg     
            Exit
        }
        #Azure storage account
        $storageAccount = Check-StorageAccount -StorageAccountName $exportEndpointData.AdministrativeUsername -UserPrincipalName $global:btAzureCredentials.UserName

        if($storageAccount) {
            $global:btAzureStorageAccountChecked = $true  
        }
    }
    else{
        $msg = "INFO: Already validated Azure Storage account with subscription ID: '$global:btAzureSubscriptionID'."
        Write-Host -ForegroundColor Green $msg
    
        Write-Host
        $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different Azure Storage account."
        Write-Host -ForegroundColor Yellow $msg      
    }
}
}

write-host 
$msg = "#######################################################################################################################`
                       UPLOADERWIZ CONFIGURATION AND FILE SERVER REMEDIATION               `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "UPLOADERWIZ CONFIGURATION AND FILE SERVER REMEDIATION" 
write-host 

if([string]::IsNullOrEmpty($FileServerRootFolderPath)){
    do {    
        do {
            Write-host -ForegroundColor Yellow  "ACTION: Enter the folder path to the FileServer root folder: "  -NoNewline
            $fileServerPath = Read-Host
            $rootPath = $fileServerPath
            #$rootPath = "`'$fileServerPath`'"

        } while($rootPath -eq "")
        
        Write-host -ForegroundColor Yellow  "ACTION: If $rootPath is correct press [C] to continue. If not, press any key to re-enter: " -NoNewline
        $confirm = Read-Host 

    } while($confirm -ne "C")
}
else{
    $fileServerPath = $FileServerRootFolderPath
    $rootPath = $fileServerPath
    #$rootPath = "`'$fileServerPath`'"
}

$alreadyProcessedUsers = @(Import-CSV "$script:workingDir\AllAlreadyProccessedHomeDirectories.csv" | where-Object { $_.PSObject.Properties.Value -ne ""} | select SourceFolder -unique | sort  { $_.SourceFolder} )

if([string]::IsNullOrEmpty($HomeDirectorySeachPattern)){
    $applyHomeDirFilter = $false
    do {
        Write-host
        $confirm = (Read-Host -prompt "Do you want to apply home directory filtering?  [Y]es or [N]o")
        if($confirm.ToLower() -eq "y") {
            $applyHomeDirFilter = $true  
            
            :SearchPatternLoop while($true) {
                Write-host -ForegroundColor Yellow  "ACTION: Enter the search pattern. It can be a combination of literal and wildcard characters (* or ?): "  -NoNewline
                $searchPattern = Read-Host
            
                $filteredHomeDirectories = (Get-ChildItem -path "$fileServerPath" -filter "$searchPattern")

                Write-host -ForegroundColor Green  "SUCCESS: $($filteredHomeDirectories.Count) home directories found with search pattern '$searchPattern'."
                
                $alreadyProcessedUsersCount = $alreadyProcessedUsers.Count
                
                if($alreadyProcessedUsersCount -ne 0) {
                    $filteredHomeDirectoryAlreadyProccessed = 0
                    foreach ($alreadyProcessedUser in $alreadyProcessedUsers.SourceFolder) {
                        if($filteredHomeDirectories.name -contains $alreadyProcessedUser){
                            $filteredHomeDirectoryAlreadyProccessed += 1
                        }
                    }
                    if($filteredHomeDirectoryAlreadyProccessed -ne 0) {
                        Write-host -ForegroundColor Yellow  "WARNING: $filteredHomeDirectoryAlreadyProccessed home directories found that were previously processed."
                    }    
                }             

                do {
                    $confirmSearchPattern = (Read-Host -prompt "Do you want to change the current search pattern?  [Y]es or [N]o")
                    if($confirmSearchPattern.ToLower() -eq "y") {
                        Continue SearchPatternLoop 
                    }
                } while(($confirmSearchPattern.ToLower() -ne "y") -and ($confirmSearchPattern.ToLower() -ne "n"))    

                Break
            }  
        }
        if($confirm.ToLower() -eq "n") {
            $filteredHomeDirectories = (Get-ChildItem -path "$fileServerPath")

            Write-host -ForegroundColor Green  "SUCCESS: $($filteredHomeDirectories.Count) home directories found."

            if($alreadyProcessedUsersCount -ne 0) {
                $filteredHomeDirectoryAlreadyProccessed = 0
                foreach ($alreadyProcessedUser in $alreadyProcessedUsers.SourceFolder) {
                    if($filteredHomeDirectories.name -contains $alreadyProcessedUser){
                        $filteredHomeDirectoryAlreadyProccessed += 1
                    }
                }
                if($filteredHomeDirectoryAlreadyProccessed -ne 0) {
                    Write-host -ForegroundColor Yellow  "WARNING: $filteredHomeDirectoryAlreadyProccessed home directories found that were previously processed."
                }    
            }    

            do {
                $confirmNewSearchPattern = (Read-Host -prompt "Do you want to apply a search pattern to narrow down home directories to process?  [Y]es or [N]o")
                if($confirmNewSearchPattern.ToLower() -eq "y") {
                    $applyHomeDirFilter = $true 

                    :SearchPatternLoop while($true) {
                        Write-host -ForegroundColor Yellow  "ACTION: Enter the search pattern. It can be a combination of literal and wildcard characters (* or ?): "  -NoNewline
                        $searchPattern = Read-Host
                    
                        $filteredHomeDirectories = (Get-ChildItem -path "$fileServerPath" -filter "$searchPattern")
            
                        Write-host -ForegroundColor Green  "SUCCESS: $($filteredHomeDirectories.Count) home directories found with search pattern '$searchPattern'."
                        
                        $alreadyProcessedUsersCount = $alreadyProcessedUsers.Count
                        
                        if($alreadyProcessedUsersCount -ne 0) {
                            $filteredHomeDirectoryAlreadyProccessed = 0
                            foreach ($alreadyProcessedUser in $alreadyProcessedUsers.SourceFolder) {
                                if($filteredHomeDirectories.name -contains $alreadyProcessedUser){
                                    $filteredHomeDirectoryAlreadyProccessed += 1
                                }
                            }
                            if($filteredHomeDirectoryAlreadyProccessed -ne 0) {
                                Write-host -ForegroundColor Yellow  "WARNING: $filteredHomeDirectoryAlreadyProccessed home directories found that were previously processed."
                            }    
                        }             
            
                        do {
                            $confirmSearchPattern = (Read-Host -prompt "Do you want to change the current search pattern?  [Y]es or [N]o")
                            if($confirmSearchPattern.ToLower() -eq "y") {
                                Continue SearchPatternLoop 
                            }
                        } while(($confirmSearchPattern.ToLower() -ne "y") -and ($confirmSearchPattern.ToLower() -ne "n"))    
            
                        Break
                    }  
                }
            } while(($confirmNewSearchPattern.ToLower() -ne "y") -and ($confirmNewSearchPattern.ToLower() -ne "n"))    

            Break
        }
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
}
else{
    $searchPattern = $HomeDirectorySeachPattern
    $applyHomeDirFilter = $true

    $filteredHomeDirectories = (Get-ChildItem -path "$fileServerPath" -filter $searchPattern)
            
    Write-host -ForegroundColor Green  "SUCCESS: $($filteredHomeDirectories.Count) home directories found with search pattern '$searchPattern'."
    
    $alreadyProcessedUsersCount = $alreadyProcessedUsers.Count
    
    if($alreadyProcessedUsersCount -ne 0) {
        $filteredHomeDirectoryAlreadyProccessed = 0
        foreach ($alreadyProcessedUser in $alreadyProcessedUsers.SourceFolder) {
            if($filteredHomeDirectories.name -contains $alreadyProcessedUser){
                $filteredHomeDirectoryAlreadyProccessed += 1
            }
        }
        if($filteredHomeDirectoryAlreadyProccessed -ne 0) {
            Write-host -ForegroundColor Yellow  "WARNING: $filteredHomeDirectoryAlreadyProccessed home directories found that were previously processed."
        }    
    }   
}


Write-Host
Check-FileServerInvalidCharacters -Path $fileServerPath

$uploaderwizCommandFilePath = ".\UploaderWiz\UploaderWiz.exe"
if($applyHomeDirFilter){    
    $uploaderwizCommandArgumentList = "-type azureblobs -accesskey " + $exportEndpointData.AdministrativeUsername + " -secretkey " + $script:secretkey + " -rootPath `"$rootpath`" -homedrive true -force True -Pathfilter $searchPattern" 
Write-Host $uploaderwizCommandArgumentList
}
else{
    $uploaderwizCommandArgumentList = "-type azureblobs -accesskey " + $exportEndpointData.AdministrativeUsername + " -secretkey " + $script:secretkey + " -rootPath `"$rootpath`" -homedrive true -force True" 
}

write-host 
$msg = "#######################################################################################################################`
                       UPLOADERWIZ EXECUTION               `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "UPLOADERWIZ EXECUTION" 
write-host 

#Run the UploaderWiz command line with parameters

$msg = "INFO: Launching UploaderWiz with these parameters:`r`n$uploaderwizCommand"
Write-Host $msg
Log-Write -Message $msg   
Write-Host

#"-type azureblobs -accesskey pablog -secretkey 97sP378PgpZd9WgCURXqURbZnQQ+xh4cHOMY/zm4+HZpRkUqSUfcSxbfGMjSi2kRW8ylU271QNk8wFP7lCf1YQ== -rootPath `"C:\Users\PabloG\OneDrive - BitTitan\Desktop\Desktop\PowerShell\FS to ODFB\File Server`" -homedrive true -force True -Pathfilter debra*" -Wait
Start-Process -FilePath $uploaderwizCommandFilePath -ArgumentList $uploaderwizCommandArgumentList

$msg = "ACTION: Once the uploaded has been completed (new window is closed when you press <Enter>), press any key to continue."
Write-Host -ForegroundColor Yellow $msg
Log-Write -Message $msg   
Write-Host


do {
    try {
        Sleep -Seconds 5
        $result = Get-Process UploaderWiz -ErrorAction Stop
    } catch{ 
        $msg = "SUCCESS: File Server Home Directories have been uploaded"
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg   
        Write-Host
        Break
    }
}while($true)

if([string]::IsNullOrEmpty($FileServerRootFolderPath)){
    $msg = "INFO: UploaderWiz log file in folder '$Env:temp\UploaderWiz'."
    Write-Host $msg
    Log-Write -Message $msg   
    #Open the CSV file
    try {    
        Start-Process -FilePath "$Env:temp\UploaderWiz"

        $msg = "SUCCESS: Folder '$Env:temp\UploaderWiz' opened."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg   
    }
    catch {
        $msg = "ERROR: Failed to open folder '$Env:temp\UploaderWiz."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message   
        Exit
    }
}

write-host 
$msg = "#######################################################################################################################`
                       SELECTING GOOGLE DRIVE ACCOUNTS                 `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "SELECTING GOOGLE DRIVE ACCOUNTS " 
write-host 

$msg = "INFO: Creating or selecting existing Google Drive endpoint."
Write-Host $msg
Log-Write -Message $msg 

if(!$global:btImportEndpointId) {
    if([string]::IsNullOrEmpty($BitTitanDestinationEndpointId)){
        #Select destination endpoint
        $global:btImportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType "GoogleDriveCustomerTenant"
    } 
    else{
        $global:btImportEndpointId = $BitTitanDestinationEndpointId
    } 
}
else{
    [PSObject]$importEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointId $global:btImportEndpointId 
    if($importEndpointData -eq -1) {
        $global:btImportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType "GoogleDriveCustomerTenant"
    }
}

#Get source endpoint credentials
if(!$importEndpointData){[PSObject]$importEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointId $global:btImportEndpointId}

if(!$global:btJsonFileContent -and !$global:btJsonFile) {
    do {
        if(!$script:inputFile) {
            Write-host -NoNewline "Please enter the file path to the Google service account credentials using JSON file:"
            Get-FileName $script:workingDir -ExtensionFilter "JSON (*.json)| *.json"
            $global:btJsonFilePath = $script:inputFile
        }
    
        #Read CSV file
        try {
            $global:btJsonFileContent = get-content $script:inputFile 
        }
        catch {
            $msg = "ERROR: Failed to import the CSV file '$script:inputFile'."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg 
            Log-Write -Message $_.Exception.Message
            Return -1    
        } 
    }while ($global:btJsonFileContent -eq "")
}

if(!$global:btGoogleDriveArray) {
    Write-Host
    $msg = "ACTION: Select a CSV file with the Home Directory -> Google Drive account mapping."
    Write-Host -ForegroundColor Yellow $msg
    Log-Write -Message $msg 

    $global:btGoogleDriveArray = @(Get-GoogleDriveAccounts)   
}

$date = (Get-Date -Format yyyyMMddHHmm)

if($global:btGoogleDriveArray) {
    do {
        try{
            if($filterHomeDirs) {
                #Export users with GoogleDrive to CSV file filtering by the folder names in the root path
                $global:btGoogleDriveArray | Select-Object SourceFolder,DestinationEmailAddress | sort { $_.DestinationEmailAddress } | where {$filteredHomeDirectories.Name -match $_.SourceFolder} | Export-Csv -Path $script:workingDir\GoogleDriveAccounts-$date.csv -NoTypeInformation -Force
            }
            else {
                $global:btGoogleDriveArray | Select-Object SourceFolder,DestinationEmailAddress | sort { $_.DestinationEmailAddress } | Export-Csv -Path $script:workingDir\GoogleDriveAccounts-$date.csv -NoTypeInformation -Force
            }
            Break
        }
        catch {
            $msg = "WARNING: Close CSV file '$script:workingDir\GoogleDriveAccounts-$date.csv' open."
            Write-Host -ForegroundColor Yellow $msg

            Start-Sleep 5
        }
    } while ($true)
}
else{

    Write-Host
    $msg = "INFO: Retrieving already processed home directories from File Server."
    Write-Host $msg
    Log-Write -Message $msg 

    foreach($folderName in $filteredHomeDirectories.Name) {
        [array]$GoogleDriveArray += New-Object PSObject -Property @{DestinationEmailAddress='';SourceFolder=$folderName;}
    }

    $global:btGoogleDriveArray | Export-Csv -Path $script:workingDir\GoogleDriveAccounts-$date.csv -NoTypeInformation -Force

    $msg = "ACTION: Provide the DestinationEmailAddress in the CSV file for each home directory processed."
    Write-Host -ForegroundColor Yellow $msg
    Log-Write -Message $msg   
}

if([string]::IsNullOrEmpty($CheckOneDriveAccounts)){
    #Open the CSV file
    try {
        
        Start-Process -FilePath $script:workingDir\GoogleDriveAccounts-$date.csv

        $msg = "SUCCESS: CSV file '$script:workingDir\GoogleDriveAccounts-$date.csv' processed, exported and open."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg   
    }
    catch {
        $msg = "ERROR: Failed to open '$script:workingDir\GoogleDriveAccounts-$date.csv' CSV file."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message   
        Exit
    }

    WaitForKeyPress -Message "ACTION: If you have edited and saved the CSV file then press any key to continue." 
}

#Re-import the edited CSV file
Try{
    $users = @(Import-CSV "$script:workingDir\GoogleDriveAccounts-$date.csv" | where-Object { $_.PSObject.Properties.Value -ne ""} | sort { $_.DestinationEmailAddress.length } )
    $totalLines = $users.Count

    if($totalLines -eq 0) {
        $msg = "INFO: No Office 365 users found with Google Drive matching the home directory names. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        
        $output = "$script:workingDir\FileServerToGoogleDriveProjects-$date.csv"
        
        write-host $output

        Return $output 
    } 
}
Catch [Exception] {
    $msg = "ERROR: Failed to import the CSV file '$script:workingDir\GoogleDriveAccounts-$date.csv'."
    Write-Host -ForegroundColor Red  $msg
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $msg   
    Log-Write -Message $_.Exception.Message   
    Exit
}

write-host 
$msg = "#######################################################################################################################`
                       MIGRATIONWIZ PROJECT CREATION                 `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "MIGRATIONWIZ PROJECT CREATION" 
write-host 

#Create AzureFileSystem-OneDriveProAPI Document project
Write-Host
$msg = "INFO: Creating MigrationWiz FileServer to Google Drive projects."
Write-Host $msg
Log-Write -Message $msg   
Write-Host

if([string]::IsNullOrEmpty($MigrationWizFolderMapping) -or !$MigrationWizFolderMapping){
    $applyCustomFolderMapping = $false
    do {
        $confirm = (Read-Host -prompt "Do you want to add a custom folder mapping to move the home directory under a folder?  [Y]es or [N]o")

        if($confirm.ToLower() -eq "y") {
            $applyCustomFolderMapping = $true
            
            do {
                Write-host -ForegroundColor Yellow  "ACTION: Enter the destination folder name: "  -NoNewline
                $destinationFolder = Read-Host

            } while($destinationFolder -eq "")
            
        }

    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
}
else{
    if($MigrationWizFolderMapping ) {
        $applyCustomFolderMapping = $true  
        $destinationFolder =  $MigrationWizFolderMapping 
    }
    else{
        $applyCustomFolderMapping = $false
    }
}

$msg = "INFO: Checking User Migration Bundle licenses available in the BitTItan account:"
Write-Host $msg
Log-Write -Message $msg   

if(-not [string]::IsNullOrEmpty($ApplyUserMigrationBundle) -and $ApplyUserMigrationBundle) {
    #Get the product ID
    #$productId = Get-BT_ProductSkuId -Ticket $script:ticket -ProductName MspcEndUserYearlySubscription
    $productId = '39854d8c-b41d-11e6-a82f-e4b31882dc3b'
    <#If ($productid) {
        $msg = "SUCCESS: Product ID obtained..."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }
    Else {
        $msg = "ERRO: Invalid Product ID"
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        Break
    }#>

    # Validate if the account have the required subscriptions
    # On this query, all the expired Subscriptions licenses will be excluded
    $curDate = Get-Date
    $licensesPack = @(Get-MW_LicensePack -Ticket $script:MwTicket -WorkgroupOrganizationId $global:btWorkgroupOrganizationId -ProductSkuId $productId | Where-Object {$_.ExpireDate -gt $curDate} | where {(($_.Purchased -eq 1 -or $_.Granted -eq 1) -and $_.Revoked -eq 0) -and ($_.Used -eq 1)})
    $licensesAvailable = 0

    if (!($licensesPack)) {
        $msg = "ERROR: No valid license pack found on this BitTitan Workgroup / Account"
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
    }
    else {
        $licensesAvailable = $licensesPack.Count 
        $msg = "INFO: $licensesAvailable User Migration Bundle licenses found on this MSPC Workgroup / Account"
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }        
}

Write-Host

$processedLines = 0
$existingMigrationList = @()
$FileServerToGoogleDriveProjects = @()

foreach ($user in $users) {      

    $ProjectName = "FS-GoogleDrive-$($user.SourceFolder)" #-$(Get-Date -Format yyyyMMddHHmm)
    $ProjectType = "Storage"   
    $exportType = "AzureFileSystem" 
    $importType = "GoogleDriveCustomerTenant" 
    $containerName = $user.SourceFolder

    $exportTypeName = "MigrationProxy.WebApi.AzureConfiguration"
    $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
        "AdministrativeUsername" = $exportEndpointData.AdministrativeUsername;
        "AccessKey" = $script:SecretKey;
        "ContainerName" = $containerName;
        "UseAdministrativeCredentials" = $true
    }


    
    $msg = "INFO: The file path to the JSON file is '$global:btJsonFilePath'."
    Write-Host $msg
    Log-Write -Message $msg 
    $importTypeName = "MigrationProxy.WebApi.GoogleDriveCustomerTenantConfiguration"
    $importConfiguration = New-Object -TypeName $importTypeName -Property @{
        "UseAdministrativeCredentials" = $true;
        "AdministrativeUsername" = $importEndpointData.AdminEmailAddress;
        "CredentialsFileName" = $global:btJsonFilePath;
        "Credentials" = $global:btJsonFileContent.ToString()
    }

    #Double Quotation Marks
    [string]$CH34=[CHAR]34
    if ($applyCustomFolderMapping) {

        $folderMapping= "FolderMapping=" + $CH34 + "^" + "Documents->Documents/" + $destinationFolder + $CH34
    }
    
    $advancedOptions = "InitializationTimeout=8 DocumentBrowsingMode=FullCopy ShouldConvertToDestinationFormat=1"
    
    $connectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
    -ProjectName $ProjectName `
    -ProjectType $ProjectType `
    -importType $importType `
    -exportType $exportType `
    -exportEndpointId $global:btExportEndpointId `
    -importEndpointId $global:btImportEndpointId `
    -exportConfiguration $exportConfiguration `
    -importConfiguration $importConfiguration `
    -advancedOptions $advancedOptions `
    -maximumSimultaneousMigrations $totalLines `
    -ZoneRequirement $ZoneRequirement

    $msg = "INFO: Adding advanced options '$advancedOptions' to the project."
    Write-Host $msg
    Log-Write -Message $msg   
    
    $msg = "INFO: Adding migration to the project:"
    Write-Host $msg
    Log-Write -Message $msg    

    $SourceFolder= $user.SourceFolder
    $importEmailAddress =  $user.DestinationEmailAddress 

    if($SourceFolder -ne "" -and $importEmailAddress -ne "") {
    
        $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -ImportEmailAddress $importEmailAddress -ErrorAction SilentlyContinue
        if(!$result) {
            try {
                $suppressOutput = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId  -ImportEmailAddress $importEmailAddress

                $tab = [char]9
                $msg = "SUCCESS: Migration '$SourceFolder->$importEmailAddress' added to the project $ProjectName."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg   

                $ProcessedLines += 1

                [array]$FileServerToGoogleDriveProjects += New-Object PSObject -Property @{ProjectName=$ProjectName;ProjectType='Storage';ConnectorId=$connectorId;MailboxId=$result.id;SourceFolder=$SourceFolder;EmailAddress=$importEmailAddress;CreateDate=$(Get-Date -Format yyyyMMddHHmm)} 
            }
            catch {
                $msg = "ERROR: Failed to add source folder and destination primary SMTP address." 
                write-Host -ForegroundColor Red $msg
                Log-Write -Message $msg     
                Exit
            }
        }
        else{
            $msg = "WARNING: Home Directory to Google Drive migration '$SourceFolder->$importEmailAddress' already exists in connector."  
            write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg  

            $existingMigrationList += "'$SourceFolder->$importEmailAddress'`n"
            $existingMigrationCount += 1

            [array]$FileServerToGoogleDriveProjects += New-Object PSObject -Property @{ProjectName=$ProjectName;ProjectType='Storage';ConnectorId=$connectorId;MailboxId=$result.id;SourceFolder=$SourceFolder;EmailAddress=$importEmailAddress;CreateDate=$(Get-Date -Format yyyyMMddHHmm)} 
        }

        ########################################################################    
        #                            APPLY UMB
        ######################################################################## 
        if(-not [string]::IsNullOrEmpty($ApplyUserMigrationBundle) -and $ApplyUserMigrationBundle) {

            $msg = "INFO: Checking and applying User Migration Bundle."
            Write-Host $msg
            Log-Write -Message $msg    

            ########################################################################    
            # If mailbox is added and was previously licensed    
            ######################################################################## 
            $mspcUser = (Get-BT_CustomerEndUser -Ticket $script:ticket -OrganizationId $global:btCustomerOrganizationId -id $result.CustomerEndUserId -Environment "BT" -IsDeleted $false) 

            if(!$mspcUser) {
                Write-host -ForegroundColor Red "ERROR: User '$importEmailAddress' not found in MSPComplete."
            }

            if($mspcUser) {
                $subscriptionEndDate = (Get-BT_Subscription -Ticket $script:ticket -Id $mspcuser.SubscriptionId.guid).SubscriptionEndDate

                if ( $mspcuser.ActiveSubscriptionId -eq "00000000-0000-0000-0000-000000000000" ) {
                    Write-host -ForegroundColor Yellow "WARNING: User '$($mspcuser.PrimaryEmailAddress)' does not have a subscription applied."

                    $isUmbApplied = $false
                }
                else {
                    Write-host -ForegroundColor Green "SUCCESS: User '$($mspcuser.PrimaryEmailAddress)' has a subscription applied that will expire in '$subscriptionEndDate'. "

                    $isUmbApplied = $true
                } 

                if(!$isUmbApplied){

                    Try {
                        $subscription = Add-BT_Subscription -ticket $global:btWorkgroupTicket -ReferenceEntityType CustomerEndUser -ReferenceEntityId $mspcuser.Id -ProductSkuId $productId -WorkgroupOrganizationId $global:btWorkgroupOrganizationId -ErrorAction Stop
                        
                        $msg = "SUCCESS: User Migration Bundle subscription assigned to MSPC User '$($mspcUser.PrimaryEmailAddress)' and migration '$SourceFolder->$importEmailAddress'."
                        Write-Host -ForegroundColor Blue  $msg
                        Log-Write -Message $msg 

                        $changeCount += 1 
                    }
                    Catch {
                        $msg =  "ERROR: Failed to assign User Migration License subscription to MSPC User '$($mspcUser.PrimaryEmailAddress)'."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg
                        Write-Host -ForegroundColor Red $($_.Exception.Message)
                        Log-Write -Message $($_.Exception.Message) 
                    }
                }    
            }
        }

    }
    else{
        if($SourceFolder -eq "") {
            $msg = "ERROR: Missing source folder in the CSV file. Skipping '$importEmailAddress' user processing."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg   
            Continue    
        } 
        if($importEmailAddress -eq "") {
            $msg = "ERROR: Missing destination Google Drive email address in the CSV file. Skipping '$SourceFolder' source folder processing."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg   
            Continue    
        }         
    }              
}

if($ProcessedLines -gt 0) {
    write-Host
    $msg = "SUCCESS: $ProcessedLines out of $totalLines migrations have been processed." 
    write-Host -ForegroundColor Green $msg
    Log-Write -Message $msg 
}
if(-not ([string]::IsNullOrEmpty($existingMigrationList)) -and $existingMigrationCount -ne 0 ) {
    write-Host
    $msg = "WARNING: $existingMigrationCount out of $totalLines migrations have not been added because they already exist: `n$existingMigrationList" 
    write-Host -ForegroundColor Yellow $msg
    Log-Write -Message $msg 
}

if([string]::IsNullOrEmpty($ApplyUserMigrationBundle)) {
    $customerUrlId = Get-CustomerUrlId -CustomerOrganizationId $global:btCustomerOrganizationId

    $url = "https://manage.bittitan.com/customers/$customerUrlId/users?qp_currentWorkgroupId=$workgroupId"

    Write-Host
    $msg = "ACTION: Apply User Migration Bundle licenses to the Google Drive email addresses in MSPComplete."
    Write-Host -ForegroundColor Yellow $msg
    Log-Write -Message $msg   
    $msg = "INFO: Opening '$url' in your default web browser."
    Write-Host $msg
    Log-Write -Message $msg   

    $result= Start-Process $url
    Start-Sleep 5
    WaitForKeyPress -Message "ACTION: If you have applied the User Migration Bundle to the users, press any key to continue"
    Write-Host

    $url = "https://migrationwiz.bittitan.com/app/projects" + "?qp_currentWorkgroupId=$workgroupId"  

    Write-Host
    $msg = "INFO: MigrationWiz projects created."
    Write-Host $msg
    Log-Write -Message $msg   
    $msg = "INFO: Opening '$url' in your default web browser."
    Write-Host $msg
    Log-Write -Message $msg  

    $result= Start-Process $url
    Start-Sleep 5

    Write-Host
    $msg = "INFO: Opening the CSV file that will be used by 'Start-MWMigrationsFromCSVFile.ps1' script."
    Write-Host $msg
    Log-Write -Message $msg  
} 
do{    
    try {
        #export the project info to CSV file
        $FileServerToGoogleDriveProjects| Select-Object ProjectName,ProjectType,ConnectorId,MailboxId,SourceFolder,EmailAddress | sort { $_.DestinationEmailAddress } | Export-Csv -Path $script:workingDir\FileServerToGoogleDriveProjects-$date.csv -NoTypeInformation -force
        $FileServerToGoogleDriveProjects| Select-Object ProjectName,ProjectType,ConnectorId,MailboxId,SourceFolder,EmailAddress | sort { $_.DestinationEmailAddress } | Export-Csv -Path $script:workingDir\AllAlreadyProccessedHomeDirectories.csv -NoTypeInformation -Append

        if([string]::IsNullOrEmpty($ApplyUserMigrationBundle)) {
            #Open the CSV file
            Start-Process -FilePath $script:workingDir\FileServerToGoogleDriveProjects-$date.csv
        }

        $msg = "SUCCESS: CSV file CSV file with the script output '$script:workingDir\FileServerToGoogleDriveProjects-$date.csv' opened."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg   
        $msg = "INFO: This CSV file will be used by Start-MWMigrationsFromCSVFile.ps1 script to automatically submit all home directories for migration."
        Write-Host $msg
        Log-Write -Message $msg   
        Write-Host

        Break
    }
    catch {
        $msg = "WARNING: Close CSV file '$script:workingDir\FileServerToGoogleDriveProjects-$date.csv' open."
        Write-Host -ForegroundColor Yellow $msg

        Start-Sleep 5
    }
} while ($true)

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg   

$output = "$script:workingDir\FileServerToGoogleDriveProjects-$date.csv"

Return $output 


##END SCRIPT
