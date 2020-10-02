<#
.SYNOPSIS
    This script will remove all invalid and non-ASCII characters from a folder path to be able to migrate the folders 
    and documents to SharePoint Online or OneDrive For Business

.DESCRIPTION
    This script will remove all invalid and non-ASCII characters from a folder path to be able to migrate the folders 
    and documents to SharePoint Online or OneDrive For Business
	
.NOTES
    Author          Pablo Galan Sabugo <pablogalanscripts@gmail.com>
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
#>

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

#Function to remove invalid characters for SharePoint Online or OneDrive For Business
Function Check-FileServerInvalidCharacters ($Path) {

    Write-Host
    $msg = "INFO: Analyzing invalid characters in all files and folders under File Server '$Path'. "
    Write-Host $msg
    Log-Write -Message $msg -LogFile $logFile 

    do {        
        $confirm = (Read-Host -prompt "If an invalid character is found, do you want to remove it?  [Y]es or [N]o")

        if($confirm.ToLower() -eq "y") {
            $removeInvalidChars = $true
        }
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    #Get all files and folders under the path specified
    #$items = Get-ChildItem -Path $Path -Recurse
    #foreach ($item in $items) {
    #Check if the item is a file or a folder
    #if ($item.PSIsContainer) { $type = "Folder" }
    #else { $type = "File" }
    #}
    
    $folders = @(Get-ChildItem -Path $Path -Recurse | where {$_.PSIsContainer -eq $true})
    $files = @(Get-ChildItem -Path $Path -Recurse | where {$_.PSIsContainer -eq $false})

    Write-Host
        

    foreach ($file in $files) {

        #Check if item name is 248 characters or more in length
        #UploaderWiz can only support file path names that are shorter than 248 characters.
        if ($file.Name.Length -gt 248) {
            $msg = "INFO: File $($file.Name) is 248 characters or over and will need to be truncated." 
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile 
        }
        else {
            #Characters that aren't allowed in file and folder names  in OneDrive, OneDrive for Business on Office 365, and SharePoint Online
            #   " * : < > ? / \ |# %
            #If the filename contains no bytes above 0x7F, then it's ASCII. 
            # gci -recurse . | where {$_.Name -match '[^\u0000-\u007F]'}

            $windowsIllegalChars = '[!&{}~#%]'
            $spoIllegalChars = '["*:<>?/\\|]'
            $nonAsciiChars = '[^\u0000-\u007F]'

            #Replace illegal characters with legal characters where found
            $newFileName = $file.Name

            filter Matches($nonAsciiChars) {
                $newFileName | Select-String -AllMatches $nonAsciiChars |
                Select-Object -ExpandProperty Matches
                Select-Object -ExpandProperty Values
            }
                
            Matches $nonAsciiChars | ForEach-Object {
                $msg = "INFO: File '$($file.FullName)' has an non ASCII character '$_'"
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile 

                #These characters may be used on the file system but not SharePoint
                if ($_ -match '[^\u0000-\u007F]') { $newFileName = ($newFileName -replace '[^\u0000-\u007F]', "") }
            }            

            filter Matches($spoIllegalChars) {
                $newFileName | Select-String -AllMatches $spoIllegalChars |
                Select-Object -ExpandProperty Matches
                Select-Object -ExpandProperty Values
            }
                
            Matches $spoIllegalChars | ForEach-Object {
                $msg = "INFO: File '$($file.FullName)' has the illegal character '$($_.Value)'" 
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile 

                #These characters may be used on the file system but not SharePoint
                if ($_.Value -match '"') { $newFileName = ($newFileName -replace '"', "") }
                elseif ($_.Value -match '"') { $newFileName = ($newFileName -replace '"', '') }
                elseif ($_.Value -match "*") { $newFileName = ($newFileName -replace "\*", "") }
                elseif ($_.Value -match ":") { $newFileName = ($newFileName -replace ":", "") }
                elseif ($_.Value -match "<") { $newFileName = ($newFileName -replace "<", "") }
                elseif ($_.Value -match ">") { $newFileName = ($newFileName -replace ">", "") }
                elseif ($_.Value -match "?") { $newFileName = ($newFileName -replace "\?", "") }
                elseif ($_.Value -match "/") { $newFileName = ($newFileName -replace "/", "") }
                elseif ($_.Value -match "\") { $newFileName = ($newFileName -replace "\\", "") }
                elseif ($_.Value -match "|") { $newFileName = ($newFileName -replace "\|", "") }
            }

            #Check for start, end and double periods

            if ($newFileName.StartsWith("~$")) { 
                $msg = "INFO: File '$($file.FullName)' starts with ~$"
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile 
            }
            while ($newFileName.StartsWith("~$")) { $newFileName = $newFileName.TrimStart("~$") }

            if ($newFileName.StartsWith(".")) { 
                $msg = "INFO: File '$($file.FullName)' starts with a period"
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile 
            }
            while ($newFileName.StartsWith(".")) { $newFileName = $newFileName.TrimStart(".") }
    
            if ($newFileName.EndsWith(".")) { 
                $msg = "INFO: File '$($file.FullName)' ends with a period"
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile 
            }
            while ($newFileName.EndsWith("."))   { $newFileName = $newFileName.TrimEnd(".") }
    
            if ($newFileName.Contains("..")) { 
                $msg = "INFO: File '$($file.FullName)' contains double periods"
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile 
            }
            while ($newFileName.Contains(".."))  { $newFileName = $newFileName.Replace("..", ".") }
    
            $errorActionPreference = 'Stop'

            #Fix file and folder names if found and the Fix switch is specified
            if (($newFileName -ne $file.Name) -and ($removeInvalidChars)) {
        
                try {
                    Rename-Item $file.FullName -NewName ($newFileName)  -ErrorAction Stop
            
                    $msg = "      SUCCESS: File '$($file.Name)' has been renamed to '$newFileName'"
                    Write-Host -ForegroundColor Green  $msg
                    Log-Write -Message $msg -LogFile $logFile
                    Write-Host 
                }
                Catch {
                    $msg = "      ERROR: Failed to rename File." 
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg -LogFile $logFile
                    Write-Host -ForegroundColor Red "      $($_.Exception.Message)"
                    Log-Write -Message "      $($_.Exception.Message)" -LogFile $logFile

                    Rename-Item $file.FullName -NewName ($newFileName+"_")  

                    $msg = "      SUCCESS: File '$($file.Name)' has been renamed to '$($newFileName)_'."
                    Write-Host -ForegroundColor Green  $msg
                    Log-Write -Message $msg -LogFile $logFile
                    Write-Host 

                }
            }
        }

    }
        
    for ($i=$folders.Count-1; $i -ge 0; $i--) {
        $folder=$folders[$i]

        #Check if item name is 248 characters or more in length
        #UploaderWiz can only support file path names that are shorter than 248 characters.
        if ($folder.Name.Length -gt 248) {
            $msg = "INFO: Folder $($folder.Name) is 248 characters or over and will need to be truncated." 
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile 
        }
        else {
            #Characters that aren't allowed in file and folder names  in OneDrive, OneDrive for Business on Office 365, and SharePoint Online
            #   " * : < > ? / \ |# %
            #If the filename contains no bytes above 0x7F, then it's ASCII. 
            # gci -recurse . | where {$_.Name -match '[^\u0000-\u007F]'}
    
            $windowsIllegalChars = '[!&{}~#%]'
            $spoIllegalChars = '["*:<>?/\\|]'
            $nonAsciiChars = '[^\u0000-\u007F]'

            #Replace illegal characters with legal characters where found
            $newFileName = $folder.Name

            filter Matches($nonAsciiChars) {
                $newFileName | Select-String -AllMatches $nonAsciiChars |
                Select-Object -ExpandProperty Matches
                Select-Object -ExpandProperty Values
            }
                
            Matches $nonAsciiChars | ForEach-Object {
                $msg = "INFO: Folder '$($folder.FullName)' has an non ASCII character '$_'"
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile 

                #These characters may be used on the file system but not SharePoint
                if ($_ -match '[^\u0000-\u007F]') { $newFileName = ($newFileName -replace '[^\u0000-\u007F]', "") }
            }            

            filter Matches($spoIllegalChars) {
                $newFileName | Select-String -AllMatches $spoIllegalChars |
                Select-Object -ExpandProperty Matches
                Select-Object -ExpandProperty Values
            }
                
            Matches $spoIllegalChars | ForEach-Object {
                $msg = "INFO: Folder '$($folder.FullName)' has the illegal character '$($_.Value)'" 
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile 

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
            if ($newFileName.StartsWith(".")) { 
                $msg = "INFO: Folder '$($folder.FullName)' starts with a period"
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile 
            }
            while ($newFileName.StartsWith(".")) { $newFileName = $newFileName.TrimStart(".") }
    
            if ($newFileName.EndsWith(".")) { 
                $msg = "INFO: Folder '$($folder.FullName)' ends with a period"
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile 
            }
            while ($newFileName.EndsWith("."))   { $newFileName = $newFileName.TrimEnd(".") }
    
            if ($newFileName.Contains("..")) { 
                $msg = "INFO: Folder '$($folder.FullName)' contains double periods"
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile 
            }
            while ($newFileName.Contains(".."))  { $newFileName = $newFileName.Replace("..", ".") }
    
            $errorActionPreference = 'Stop'

            #Fix file and folder names if found and the Fix switch is specified
            if (($newFileName -ne $folder.Name) -and ($removeInvalidChars)) {
        
                try {
                    Rename-Item $folder.FullName -NewName ($newFileName)  -ErrorAction Stop
            
                    $msg = "      SUCCESS: Folder '$($folder.Name)' has been renamed to '$newFileName'"
                    Write-Host -ForegroundColor Green  $msg
                    Log-Write -Message $msg -LogFile $logFile
                    Write-Host 
                }
                Catch {
                    $msg = "      ERROR: Failed to rename Folder." 
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg -LogFile $logFile
                    Write-Host -ForegroundColor Red "      $($_.Exception.Message)"
                    Log-Write -Message "      $($_.Exception.Message)" -LogFile $logFile

                    Rename-Item $folder.FullName -NewName ($newFileName+"_")  

                    $msg = "      SUCCESS: Folder '$($folder.Name)' has been renamed to '$($newFileName)_'."
                    Write-Host -ForegroundColor Green  $msg
                    Log-Write -Message $msg -LogFile $logFile
                    Write-Host 

                }
            }
        }
    }     
}

# Function to write information to the Log File
Function Log-Write {
    param
    (
        [Parameter(Mandatory=$true)]    [string]$Message,
        [Parameter(Mandatory=$true)]    [string]$LogFile
    )
    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
	Add-Content -Path $logFile -Value $lineItem
}

#######################################################################################################################
#                                               MAIN PROGRAM
#######################################################################################################################

#Working Directory
$global:workingDir = "C:\scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Check-FileServerInvalidCharacters.log"
$global:logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg -LogFile $logFile

Write-Host
do {
    
    do {
        Write-host -ForegroundColor Yellow  "ACTION: Enter the folder path to the FileServer root folder: "  -NoNewline
        $fileServerPath = Read-Host
        $rootPath = "`'$fileServerPath`'"

    } while($rootPath -eq "")
      
    Write-host -ForegroundColor Yellow  "ACTION: If $rootPath is correct press [C] to continue. If not, press any key to re-enter: " -NoNewline
    $confirm = Read-Host 

} while($confirm -ne "C")

Check-FileServerInvalidCharacters -Path $fileServerPath


$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg -LogFile $logFile

##END SCRIPT
