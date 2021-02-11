<#
.SYNOPSIS
     .SYNOPSIS
    This script will create a MigrationWiz project to migrate FileServer Home Directories to OneDrive For Business accounts.
    It will generate a CSV file with the MigrationWiz project and all the migrations that will be used by the script 
    Start-MWMigrationsFromCSVFile.ps1 to submit all the migrations.
    Another output 

.DESCRIPTION
    This script will download the UploaderWiz exe file and execute it to create Azure blob containers per each home directory 
    found in the File Server and upload each home directory to the corresponding blob container. After that the script will 
    create the MigrationWiz projects to migrate from Azure blob containers to the OneDrive For Business accounts.
    The output of this script will be 4 CSV files:
    1. FileServerToOD4BProject-<date>.csv with the project name of the home directory processed that will be passed to Start-MWMigrationsFromCSVFile.ps1
                                          to start automatically all created MigrationWiz projects. 
    2. ProccessedHomeDirectories-<date>.csv with the project names of the home directories processed the same day that will be passed to Get-MWMigrationProjectStatistics.ps1
                                          to generate the migration statistics and error reports of MigrationWiz projects processed that day. 
    3. AllFailedHomeDirectories.csv with ALL the project names of the home directories processed to date.
    4. AllProccessedHomeDirectories.csv with ALL the project names of the home directories failed to process to date.

.PARAMETER WorkingDirectory
    This parameter defines the folder path where the CSV files generated during the script execution will be placed.
    This parameter is optional. If you don't specify an output folder path, the script will output the CSV files to 'C:\scripts'.  
.PARAMETER BitTitanAzureDatacenter
    This parameter defines the Azure data center the MigrationWiz project will be using. These are the accepted Azure data centers:
    'NorthAmerica','WesternEurope','AsiaPacific','Australia','Japan','SouthAmerica','Canada','NorthernEurope','China','France','SouthAfrica'
    This parameter is optional. If you don't specify an Azure data center, the script will use 'NorthAmerica' by default.  
.PARAMETER BitTitanWorkgroupId
    This parameter defines the BitTitan Workgroup Id.
    This parameter is optional. If you don't specify a BitTitan Workgroup Id, the script will display a menu for you to manually select the workgroup.  
.PARAMETER DownloadLatestVersion
    This parameter defines if the script must download the latest version of UploaderWiz before starting the UploaderWiz execution for file server upload. 
    This parameter is optional. If you don't specify the parameter with true value, the script will check if UploaderWIz was downloaded and if it was, 
    it will skip the new download. If UploaderWiz wasn´t previously downloaded it will download it for the first time. 
.PARAMETER BitTitanWorkgroupId
    This parameter defines the BitTitan Workgroup Id.
    This parameter is optional. If you don't specify a BitTitan Workgroup Id, the script will display a menu for you to manually select the workgroup.  
.PARAMETER BitTitanCustomerId
    This parameter defines the BitTitan Customer Id.
    This parameter is optional. If you don't specify a BitTitan Customer Id, the script will display a menu for you to manually select the customer.  
.PARAMETER BitTitanSourceEndpointId
    This parameter defines the BitTitan source endpoint Id.
    This parameter is optional. If you don't specify a source endpoint Id, the script will display a menu for you to manually select the source endpoint.  
    The selected source endpoint Id will be save in a global variable for next executions.
 .PARAMETER AzureStorageAccessKey
    This parameter defines the Azure storage primary access key.
    This parameter is mandatory in unattended execution. If you don't specify the Azure storage primary access key, the script will prompt for it in every execution
    making the unattended execution an interactive execution.  
.PARAMETER BitTitanDestinationEndpointId
    This parameter defines the BitTitan destination endpoint Id.
    This parameter is optional. If you don't specify a destination endpoint Id, the script will display a menu for you to manually select the destination endpoint.  
    The selected destination endpoint Id will be save in a global variable for next executions.
 .PARAMETER FileServerRootFolderPath
    This parameter defines the folder path to the file server root folder holding all end user home directories. 
    This parameter is mandatory in unattended execution. If you don't specify the folder path to the file server root folder, the script will prompt for it in every execution
    making the unattended execution an interactive execution.  
.PARAMETER HomeDirectorySearchPattern
    This parameter defines which projects you want to process, based on the project name search term. There is no limit on the number of characters you define on the search term.
    This parameter is optional. If you don't specify a project search term, all projects in the customer will be processed.
    Example: to process all projects starting with "Batch" you enter '-ProjectSearchTerm Batch'  
.PARAMETER CheckFileServer
    This parameter defines if the script must analyze the file server and remove all invalid characters both in Azure blob container and in OneDrive. 
    This parameter is optional. If you don't specify the parameter with true value, the file server folder and file names won´t be analyzed and invalid characters won´t be removed.
.PARAMETER CheckOneDriveAccounts
    This parameter defines if the home directory name exist as a OneDrive for Business account (Home Directory name = User Principal Name prefix).
    This parameter is mandatory. If you don't specify the paramter with a true value, you have to specify a CSV file name with the home directory and OneDrive for Business mapping.
.PARAMETER MigrationWizFolderMapping
    This parameter defines if the home directory must be migrated under a destination subfolder.
    This parameter is optional. If you don't specify the parameter with the destination subfolder name, the home directory will be directly migrated under the OneDrive.
.PARAMETER OwnAzureStorageAccount
    This parameter defines if the destination endpoint must use the custom Azure storage account specified in the endpoint.
    This parameter is optional. If you don't specify the parameter with true value, the destination endpoint will use the Microsoft provided Azure storage.
.PARAMETER ApplyUserMigrationBundle
    This parameter defines if the migration added to the MigrationWiz project must be licensed with an existing User Migration Bundle.
    This parameter is optional. If you don't specify the parameter with true value, the migration won't be automatically licensed. 
.PARAMETER BitTitanMigrationScope
    This parameter defines the BitTitan migration status.
    This paramenter only accepts 'All', 'NotStarted', 'Failed','ErrorItems' and 'NotSuccessfull' as valid arguments.
    This parameter is optional. If you don't specify a BitTitan migration scope type, the script will display a menu for you to manually select the migration scope.  

.PARAMETER BitTitanMigrationType
    This parameter defines the BitTitan migration submission type.
    This paramenter only accepts 'Verify', 'PreStage', 'Full', 'RetryErrors', 'Pause' and 'Reset' as valid arguments.
    This parameter is optional. If you don't specify a BitTitan migration submission type, the script will display a menu for you to manually select the migration scope.  


.NOTES
    Author          Pablo Galan Sabugo <pablogalanscripts@gmail.com>
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
#>

Param
(
    [Parameter(Mandatory = $true)] [String]$WorkingDirectory,
    [Parameter(Mandatory = $true)] [Boolean]$DownloadLatestVersion,
    [Parameter(Mandatory = $true)] [ValidateSet('NorthAmerica','WesternEurope','AsiaPacific','Australia','Japan','SouthAmerica','Canada','NorthernEurope','China','France','SouthAfrica')] [String]$BitTitanAzureDatacenter,
    [Parameter(Mandatory = $true)] [String]$BitTitanWorkgroupId,
    [Parameter(Mandatory = $true)] [String]$BitTitanCustomerId,
    [Parameter(Mandatory = $true)] [String]$BitTitanSourceEndpointId,
    [Parameter(Mandatory = $true)] [String]$AzureStorageAccessKey,
    [Parameter(Mandatory = $true)] [String]$BitTitanDestinationEndpointId,
    [Parameter(Mandatory = $true)] [String]$FileServerRootFolderPath,
    [Parameter(Mandatory = $true)] [Boolean]$CheckFileServer,
    [Parameter(Mandatory = $true)] [Boolean]$CheckOneDriveAccounts,
    [Parameter(Mandatory = $false)] [String]$MigrationWizFolderMapping,
    [Parameter(Mandatory = $true)] [Boolean]$OwnAzureStorageAccount,
    [Parameter(Mandatory = $true)] [Boolean]$ApplyUserMigrationBundle,
    [Parameter(Mandatory = $true)] [ValidateSet('All', 'NotStarted', 'Failed', 'ErrorItems', 'NotSuccessfull')] [String]$BitTitanMigrationScope,
    [Parameter(Mandatory = $true)] [ValidateSet('Verify', 'PreStage', 'Full', 'RetryErrors', 'Pause', 'Reset')] [String]$BitTitanMigrationType
)

#######################################################################################################################
#                  HELPER FUNCTIONS                          
#######################################################################################################################
Function Get-CsvFile {
    Write-Host
    Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import the HomeDirectoryFolderNames and their corresponding UserPrincipalNames."
    Get-FileName $script:workingDir

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
    $CSVHeaders = "HomeDirectoryFolderName,UserPrincipalName"
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
Function Get-FileName {
    param 
    (      
        [parameter(Mandatory=$true)] [String]$initialDirectory

    )

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $script:inputFile = $OpenFileDialog.filename

    if($OpenFileDialog.filename -ne "") {		    
        $msg = "SUCCESS: CSV file '$($OpenFileDialog.filename)' selected."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg 
    }
    else {
        $msg = "ERROR: CSV file has not been selected. Script aborted"
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
        Exit
    }
}

# Function to write information to the Log File
Function Log-Write {
    param
    (
        [Parameter(Mandatory = $true)]    [string]$Message
    )

    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
    Add-Content -Path $logFile -Value $lineItem
}
 
# Function to create the working and log directories
Function Create-Working-Directory {    
    param 
    (
        [CmdletBinding()]
        [parameter(Mandatory = $true)] [string]$workingDir,
        [parameter(Mandatory = $true)] [string]$logDir,
        [parameter(Mandatory = $false)] [string]$metadataDir
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

#######################################################################################################################
#                                               MAIN PROGRAM
#######################################################################################################################
 
#######################################################################################################################
#                       SELECT WORKING DIRECTORY  
#######################################################################################################################

Write-Host
Write-Host
Write-Host -ForegroundColor Yellow "             BitTitan File Server to OneDrive For Business migration project creation + start migrations tool."
Write-Host

write-host 
$msg = "#######################################################################################################################`
                       SELECT WORKING DIRECTORY                  `
#######################################################################################################################"
Write-Host $msg
write-host 

#Working Directorys
$script:workingDir = "C:\scripts"

if ([string]::IsNullOrEmpty($WorkingDirectory)) {
    if (!$global:btCheckWorkingDirectory) {
        do {
            $confirm = (Read-Host -prompt "The default working directory is '$script:workingDir'. Do you want to change it?  [Y]es or [N]o")
            if ($confirm.ToLower() -eq "y") {
                #Working Directory
                $script:workingDir = [environment]::getfolderpath("desktop")
                Get-Directory $script:workingDir            
            }

            $global:btCheckWorkingDirectory = $true

        } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    }
}
else {
    $script:workingDir = $WorkingDirectory
}

#Logs directory
$logDirName = "LOGS"
$logDir = "$script:workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format "yyyyMMddTHHmmss")_Concatenate-Create-Start-FileServerToOD4B.log"
$script:logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $script:workingDir -logDir $logDir

$users = Get-CsvFile

foreach($user in $users){
    .\Start-MWMigrationsFromCSVFile.ps1 -BitTitanWorkGroupId $BitTitanWorkGroupId `
                                        -BitTitanCustomerId $BitTitanCustomerId `
                                        -BitTitanMigrationScope $BitTitanMigrationScope `
                                        -BitTitanMigrationType $BitTitanMigrationType `
                                        -ProjectsCSVFilePath $(    .\Create-FileServerToOD4BMWMigrationProjects.ps1 -WorkingDirectory 'C:\Scripts' `
                                                                                                                    -downloadLatestVersion $downloadLatestVersion `
                                                                                                                    -BitTitanAzureDatacenter $BitTitanAzureDatacenter `
                                                                                                                    -BitTitanWorkgroupId $BitTitanWorkgroupId `
                                                                                                                    -BitTitanCustomerId $BitTitanCustomerId `
                                                                                                                    -BitTitanSourceEndpointId  $BitTitanSourceEndpointId `
                                                                                                                    -AzureStorageAccessKey $AzureStorageAccessKey `
                                                                                                                    -FileServerRootFolderPath $FileServerRootFolderPath `
                                                                                                                    -HomeDirectorySearchPattern $user.SourceFolder `
                                                                                                                    -CheckFileServer $CheckFileServer `
                                                                                                                    -BitTitanDestinationEndpointId $BitTitanDestinationEndpointId `
                                                                                                                    -CheckOneDriveAccounts $CheckOneDriveAccounts `
                                                                                                                    -HomeDirToUserPrincipalNameMapping $user `
                                                                                                                    -MigrationWizFolderMapping $MigrationWizFolderMapping `
                                                                                                                    -OwnAzureStorageAccount $OwnAzureStorageAccount `
                                                                                                                    -ApplyUserMigrationBundle $ApplyUserMigrationBundle   )
}
