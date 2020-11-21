<#
.SYNOPSIS
    Script to analyze a source Office 365 tenant and create automatically all MigrationWiz projects to migrate all supported workloads to another Office 365 tenant.
    
.DESCRIPTION
    This script will will analyze an Office 365 tenant to discover all the objects that can be migrated with MigrationWiz and will create automatically all MigrationWiz 
    projects from exported CSV files containing all the Office 365 tenant objects like mailboxes, OneDrive For Business accounts, Classic SPO sites, Office 365 unified groups, Microsoft Teams 
    and Public Folders. Once the projects are created it will output a CSV file with the MigrationWiz project names to be used by the script Start-MW_Migrations_From_CSVFile.ps1 which will start 
    automatically all MigrationWiz projects created by previous script from the CSV with all MigrationWiz project names.
       
.NOTES
    Author          Pablo Galan Sabugo <pablogalanscripts@gmail.com>
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
    
    <#
#>

#######################################################################################################################
#                  HELPER FUNCTIONS                          
#######################################################################################################################
Function Import-PowerShellModules {
    if (!(((Get-Module -Name "MSOnline") -ne $null) -or ((Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue) -ne $null))) {
        Write-Host
        $msg = "INFO: MSOnline PowerShell module not installed."
        Write-Host $msg     
        $msg = "INFO: Installing MSOnline PowerShell module."
        Write-Host $msg

        Sleep 5
    
        try {
            Install-Module -Name MSOnline -force -ErrorAction Stop
        }
        catch {
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
    
        try {
            Install-Module -Name AzureAD -force -ErrorAction Stop
        }
        catch {
            $msg = "ERROR: Failed to install AzureAD module. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Write-Host
            $msg = "ACTION: Run this script 'As administrator' to intall the AzureAD module."
            Write-Host -ForegroundColor Yellow $msg
            Exit
        }
        Import-Module AzureAD
    }

    if (!(((Get-Module -Name "MicrosoftTeams") -ne $null) -or ((Get-InstalledModule -Name "MicrosoftTeams" -ErrorAction SilentlyContinue) -ne $null))) {
        Write-Host
        $msg = "INFO: MicrosoftTeams PowerShell module not installed."
        Write-Host $msg     
        $msg = "INFO: Installing MicrosoftTeams PowerShell module."
        Write-Host $msg

        Sleep 5
    
        try {
            Install-Module -Name MicrosoftTeams -force -ErrorAction Stop
        }
        catch {
            $msg = "ERROR: Failed to install MicrosoftTeams module. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Write-Host
            $msg = "ACTION: Run this script 'As administrator' to intall the MicrosoftTeams module."
            Write-Host -ForegroundColor Yellow $msg
            Exit
        }
        Import-Module MicrosoftTeams
    }

    if (!(((Get-Module -Name "Microsoft.Online.SharePoint.PowerShell") -ne $null) -or ((Get-InstalledModule -Name "Microsoft.Online.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -ne $null))) {
        Write-Host
        $msg = "INFO: Microsoft SharePoint Online Services PowerShell module not installed."
        Write-Host $msg     
        $msg = "INFO: Installing Microsoft SharePoint Online Services PowerShell module."
        Write-Host $msg

        Sleep 5
    
        try {
            Install-Module -Name Microsoft.Online.SharePoint.PowerShell -force -ErrorAction Stop
        }
        catch {
            $msg = "ERROR: Failed to install Microsoft SharePoint Online Services PowerShell module. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Write-Host
            $msg = "ACTION: Run this script 'As administrator' to intall the Microsoft SharePoint Online Services PowerShell module."
            Write-Host -ForegroundColor Yellow $msg
            Exit
        }
        Import-Module Microsoft.Online.SharePoint.PowerShell
    }

    if (!(((Get-Module -Name "SharePointPnPPowerShellOnline") -ne $null) -or ((Get-InstalledModule -Name "SharePointPnPPowerShellOnline" -ErrorAction SilentlyContinue) -ne $null))) {
        Write-Host
        $msg = "INFO: SharePointPnPPowerShellOnline PowerShell module not installed."
        Write-Host $msg     
        $msg = "INFO: Installing SharePointPnPPowerShellOnline PowerShell module."
        Write-Host $msg

        Sleep 5
    
        try {
            Install-Module -Name SharePointPnPPowerShellOnline -force -ErrorAction Stop
        }
        catch {
            $msg = "ERROR: Failed to install SharePointPnPPowerShellOnline module. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Write-Host
            $msg = "ACTION: Run this script 'As administrator' to intall the SharePointPnPPowerShellOnline module."
            Write-Host -ForegroundColor Yellow $msg
            Exit
        }
        Import-Module SharePointPnPPowerShellOnline
    }
    else {    
        $PnPversion = (Get-Module SharePointPnPPowerShellOnline -ListAvailable | Select-Object Name, Version | Sort-Object Version -Descending | Select-Object -first 1).Version
        if ($PnPversion -lt '3.26.2010.0') {
            Write-Host
            $msg = "INFO: SharePointPnPPowerShellOnline PowerShell module not updated. Current version $PnPversion"
            Write-Host $msg     
            do {        
                $confirm = (Read-Host -prompt "Do you want to update SharePointPnPPowerShellOnline PowerShell module?  [Y]es or [N]o")
                if ($confirm.ToLower() -eq "y") {
                    $msg = "ACTION: Execute this script as an Admin to update SharePointPnPPowerShellOnline PowerShell module."
                    Write-Host -foregorouncolor Yellow $msg
                    
                    $msg = "INFO: Updating SharePointPnPPowerShellOnline PowerShell module."
                    Write-Host $msg
                    Update-Module SharePointPnPPowerShell* -force
                }
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
        }
    }
}
function Import-MigrationWizPowerShellModule {
    if (( $null -ne (Get-Module -Name "BitTitanPowerShell")) -or ( $null -ne (Get-InstalledModule -Name "BitTitanManagement" -ErrorAction SilentlyContinue))) {
        return
    }

    $currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
    $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll", "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
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
    $result = Start-Process $url
    Exit

}

# Function to create the working and log directories
Function Create-Working-Directory {    
    param 
    (
        [CmdletBinding()]
        [parameter(Mandatory = $true)] [string]$workingDir,
        [parameter(Mandatory = $true)] [string]$logDir
    )
    if ( !(Test-Path -Path $script:workingDir)) {
        try {
            $suppressOutput = New-Item -ItemType Directory -Path $script:workingDir -Force -ErrorAction Stop
            $msg = "SUCCESS: Folder '$($script:workingDir)' for CSV files has been created."
            Write-Host -ForegroundColor Green $msg
        }
        catch {
            $msg = "ERROR: Failed to create '$script:workingDir'. Script will abort."
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
}

# Function to write information to the Log File
Function Log-Write {
    param
    (
        [Parameter(Mandatory = $true)]    [string]$Message
    )
    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
    Add-Content -Path $script:logFile -Value $lineItem
}

Function Get-FileName($initialDirectory, $DefaultColumnName) {

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $script:inputFile = $OpenFileDialog.filename

    if ($OpenFileDialog.filename -eq "") {

        if ($defaultColumnName -eq "PrimarySmtpAddress") {
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
        elseif ($defaultColumnName -eq "MailNickName") {
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
        elseif ($defaultColumnName -eq "SourceEmailAddress,DestinationEmailAddress") {
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
        else {
            Return $false
        }            

        # open file for editing
        Start-Process $file 

        do {
            $confirm = (Read-Host -prompt "Are you done editing the import CSV file?  [Y]es or [N]o")
            if ($confirm -eq "Y") {
                $importConfirm = $true
            }

            if ($confirm -eq "N") {
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
        while (-not $importConfirm)
            
        $msg = "SUCCESS: CSV file '$script:inputFile' saved."
        Write-Host -ForegroundColor Green  $msg

        Return $true
    }
    else {
        $msg = "SUCCESS: CSV file '$($OpenFileDialog.filename)' selected."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
        Return $true
    }
}

Function Get-Directory($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null    
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.ShowDialog() | Out-Null

    if ($FolderBrowser.SelectedPath -ne "") {
        $workingDir = $FolderBrowser.SelectedPath               
    }
    Write-Host -ForegroundColor Gray  "INFO: Directory '$workingDir' selected."
}

# Function to wait for the user to press any key to continue
Function WaitForKeyPress {
    $msg = "ACTION: If you have edited and saved the CSV file then press any key to continue. Press 'Ctrl + C' to exit." 
    Write-Host $msg
    Log-Write -Message $msg
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
}

Function Import-CSV_RecipientMapping {

    $result = Get-FileName $script:workingDir -DefaultColumnName "SourceEmailAddress,DestinationEmailAddress"

    if ($result) {
        ##Import the CSV file
        Try {
            $script:emailAddressMappingCSVFile = @(Import-Csv $script:inputFile | Where-Object { $_.PSObject.Properties.Value -ne "" } )
        }
        Catch [Exception] {
            Write-Host -ForegroundColor Red "ERROR: Failed to import the CSV file '$script:inputFile'."
            Write-Host -ForegroundColor Red $_.Exception.Message
            Exit
        }

        #Check if CSV is formated properly
        If ($script:emailAddressMappingCSVFile.SourceEmailAddress -eq $null -or $script:emailAddressMappingCSVFile.DestinationEmailAddress -eq $null) {
            Write-Host -ForegroundColor Red "ERROR: The CSV file format is invalid. It must have 2 columns: 'SourceEmailAddress' and 'DestinationEmailAddress' columns."
            Exit 
        }

        #Load existing advanced options
        $ADVOPTString += $Connector.AdvancedOptions
        $ADVOPTString += "`n"

        $count = 0

        #Processing CSV into string
        Write-Host "         INFO: Applying RecipientMappings from CSV File:"

        $script:emailAddressMappingCSVFile | ForEach-Object {

            $sourceAddress = $_.SourceEmailAddress
            $destinationAddress = $_.DestinationEmailAddress

            $recipientMapping = "RecipientMapping=`"$sourceAddress->$destinationAddress`""

            $count += 1

            Write-Host -ForegroundColor Green "         SUCCESS: Email address mapping $sourceAddress->$destinationAddress found." 
                   
            $allRecipientMappings += $recipientMapping
            $allRecipientMappings += "`n"

        }

        Write-Host -ForegroundColor Green "         SUCCESS: CSV file '$script:inputFile' succesfully processed. $count recipient mappings applied."

        Return $allRecipientMappings
    }
}

Function Import-CSV_MailNickNames {

    $result = Get-FileName $script:workingDir -DefaultColumnName "MailNickName"

    if ($result) {
        ##Import the CSV file
        Try {
            $script:importedMailNickNames = @(Import-Csv -Path $script:inputFile) 
            $script:importedMailNickNames            
        }
        Catch [Exception] {
            Write-Host -ForegroundColor Red "ERROR: Failed to import the CSV file '$script:inputFile '."
            Write-Host -ForegroundColor Red $_.Exception.Message

            Exit
        }

        #Check if CSV is formated properly
        If ($script:importedMailNickNames.MailNickName -eq $null) {
            Write-Host -ForegroundColor Red "ERROR: The CSV file format is invalid. It must have 1 'MailNickName' column."
            
            Exit
        }         
    }
}

Function isNumeric($x) {
    $x2 = 0
    $isNum = [System.Int32]::TryParse($x, [ref]$x2)
    return $isNum
}

# Function to rewrite domain in CSV file
Function ReWriteFile ($File, $Source, $Target) {
    do {
        try {

            (Get-Content $File) -replace $Source, $Target | Set-Content $File
            if ($Source -ne $Target ) {
                $msg = "      WARNING: '$Source' has been renamed to '$Target' in '$File'."
                Write-Host -ForegroundColor Yellow $msg
            }

            Break
        }
        catch {
            $msg = "WARNING: Close CSV file '$File' open."
            Write-Host -ForegroundColor Yellow $msg

            Start-Sleep 5
        }
    } while ($true)
}

# Function to query destination email addresses
Function Apply-EmailAddressMapping {
    do {
        $confirm = (Read-Host -prompt "Are you migrating to the same email addresses?  [Y]es or [N]o")
    } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    if ($confirm.ToLower() -eq "y") {
        $script:sameEmailAddresses = $true
        $script:sameUserName = $true
        $script:differentDomain = $false
        
        $script:selectedDomains = @(Select-Domain -Credentials $global:btSourceO365Creds -EmailAddressMapping $true -sourceOrDestination "source")
      
        if ($script:selectedDomains -and $script:selectedDomains.Count -eq 1) {        
            $msg = "INFO: There is only 1 verified domain that that can be migrated to destination '$script:selectedDomains'."
            Write-Host $msg
            Log-Write -Message $msg 

            Return $true
        }
        if ($script:selectedDomains -and $script:selectedDomains.Count -gt 1) {        
            $msg = "INFO: There are several verified domains that can be migrated to destination '$script:selectedDomains'."
            Write-Host $msg
            Log-Write -Message $msg 

            Return $true
        }
        else {
            Return $false
        }
    }
    elseif ($confirm.ToLower() -eq "n") {
        
        $script:sameEmailAddresses = $false

        do {
            $confirm = (Read-Host -prompt "Are you migrating to a different domain?  [Y]es or [N]o")
        } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        # If destination Domain is different
        if ($confirm.ToLower() -eq "y") {
            
            $script:differentDomain = $true

            do {
                $domains = (Read-Host -prompt "Please enter the destination domain (or domains separated by comma)")

                $script:selectedDomains = @($domains.split(","))

            }while ($script:selectedDomains -eq "")
            
            $msg = "INFO: The destination domain(s) is/are '$script:selectedDomains'."
            Write-Host $msg
            Log-Write -Message $msg             
            
        }
        else {
            $script:differentDomain = $false         
        } 
        
        do {
            $confirm = (Read-Host -prompt "Are the destination email addresses keeping the same user prefix?  [Y]es or [N]o")
        } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        if ($confirm.ToLower() -eq "y") {
            $script:sameUserName = $true
            Return $false 
        }
        else {
            $script:sameUserName = $false
            Return $true 
        }    
    }
}

#######################################################################################################################
#                    BITTITAN
#######################################################################################################################

# Function to authenticate to BitTitan SDK
Function Connect-BitTitan {
    #Install Packages/Modules for Windows Credential Manager if required
    If (!(Get-PackageProvider -Name 'NuGet')) {
        Install-PackageProvider -Name NuGet -Force
    }
    If (!(Get-Module -ListAvailable -Name 'CredentialManager')) {
        Install-Module CredentialManager -Force
    } 
    else { 
        Import-Module CredentialManager
    }

    # Authenticate
    $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'
    
    if (!$script:creds) {
        $credentials = (Get-Credential -Message "Enter BitTitan credentials")
        if (!$credentials) {
            $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Exit
        }
        New-StoredCredential -Target 'https://migrationwiz.bittitan.com' -Persist 'LocalMachine' -Credentials $credentials | Out-Null
        
        $msg = "SUCCESS: BitTitan credentials stored in Windows Credential Manager."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg

        $script:creds = Get-StoredCredential -Target 'https://migrationwiz.bittitan.com'

        $msg = "SUCCESS: BitTitan credentials retrieved from Windows Credential Manager."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }
    else {
        $msg = "SUCCESS: BitTitan credentials retrieved from Windows Credential Manager."
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
        $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll", "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
        foreach ($moduleLocation in $moduleLocations) {
            if (Test-Path $moduleLocation) {
                Import-Module -Name $moduleLocation

                # Get a ticket and set it as default
                $script:ticket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction SilentlyContinue
                # Get a MW ticket
                $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction SilentlyContinue 

                if (!$script:ticket -or !$script:mwTicket) {
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
        $result = Start-Process $url

        Exit
    }  

    if (!$script:ticket -or !$script:mwTicket) {
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

# Function to create a connector under a customer
Function Create-MW_Connector {
    param 
    (      
        [parameter(Mandatory = $true)] [guid]$customerOrganizationId,
        [parameter(Mandatory = $true)] [String]$ProjectName,
        [parameter(Mandatory = $true)] [String]$ProjectType,
        [parameter(Mandatory = $true)] [String]$importType,
        [parameter(Mandatory = $true)] [String]$exportType,   
        [parameter(Mandatory = $true)] [guid]$exportEndpointId,
        [parameter(Mandatory = $true)] [guid]$importEndpointId,  
        [parameter(Mandatory = $true)] [object]$exportConfiguration,
        [parameter(Mandatory = $true)] [object]$importConfiguration,
        [parameter(Mandatory = $false)] [String]$advancedOptions,   
        [parameter(Mandatory = $false)] [String]$folderFilter = "",
        [parameter(Mandatory = $false)] [String]$maximumSimultaneousMigrations = 100,
        [parameter(Mandatory = $false)] [String]$MaxLicensesToConsume = 10,
        [parameter(Mandatory = $false)] [int64]$MaximumDataTransferRate,
        [parameter(Mandatory = $false)] [String]$Flags,
        [parameter(Mandatory = $false)] [String]$ZoneRequirement,
        [parameter(Mandatory = $false)] [Boolean]$updateConnector   
        
    )
    try {
        $connector = @(Get-MW_MailboxConnector -ticket $script:MwTicket `
                -UserId $script:MwTicket.UserId `
                -OrganizationId $customerOrganizationId `
                -Name "$ProjectName" `
                -ErrorAction SilentlyContinue
            #-SelectedExportEndpointId $exportEndpointId `
            #-SelectedImportEndpointId $importEndpointId `        
            #-ProjectType $ProjectType `
            #-ExportType $exportType `
            #-ImportType $importType `

        ) 

        if ($connector.Count -eq 1) {
            $msg = "WARNING: Connector '$($connector.Name)' already exists with the same configuration." 
            write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg 

            if ($updateConnector) {
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
        elseif ($connector.Count -gt 1) {
            $msg = "WARNING: $($connector.Count) connectors '$ProjectName' already exist with the same configuration." 
            write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg 

            return $null

        }
        else {
            try { 
                $connector = Add-MW_MailboxConnector -ticket $script:MwTicket `
                    -UserId $script:MwTicket.UserId `
                    -OrganizationId $customerOrganizationId `
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
            catch {
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

# Function to create an endpoint under a customer. Configuration Table in https://www.bittitan.com/doc/powershell.html#PagePowerShellmspcmd%20
Function Create-MSPC_Endpoint {
    param 
    (      
        [parameter(Mandatory = $true)] [guid]$customerOrganizationId,
        [parameter(Mandatory = $false)] [String]$endpointType,
        [parameter(Mandatory = $false)] [String]$endpointName,
        [parameter(Mandatory = $false)] [object]$endpointConfiguration,
        [parameter(Mandatory = $false)] [String]$exportOrImport,
        [parameter(Mandatory = $false)] [Boolean]$updateEndpoint
    )

    $script:CustomerTicket = Get-BT_Ticket -OrganizationId $customerOrganizationId
    
    if ($endpointType -eq "AzureFileSystem") {
        
        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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

            do {
                $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key ").trim()
            }while ($secretKey -eq "")

            $msg = "INFO: Azure storage account access key is '$secretKey'."
            Write-Host $msg
            Log-Write -Message $msg 

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $azureFileSystemConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $azureAccountName; #Azure Storage Account Name        
                "AccessKey"                    = $secretKey; #Azure Storage Account SecretKey         
                "ContainerName"                = $ContainerName #Container Name
            }
        }
        else {
            $azureFileSystemConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername; #Azure Storage Account Name        
                "AccessKey"                    = $endpointConfiguration.AccessKey; #Azure Storage Account SecretKey         
                "ContainerName"                = $endpointConfiguration.ContainerName #Container Name
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $azureFileSystemConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $azureFileSystemConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "AzureSubscription") {
           
        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($secretKey -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            do {
                $azureSubscriptionID = (Read-Host -prompt "Please enter the Azure subscription ID").trim()
            }while ($azureSubscriptionID -eq "")

            $msg = "INFO: Azure subscription ID is '$azureSubscriptionID'."
            Write-Host $msg
            Log-Write -Message $msg 

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureSubscriptionConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $adminUsername;     
                "AdministrativePassword"       = $adminPassword;         
                "SubscriptionID"               = $azureSubscriptionID
            }
        }
        else {
            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureSubscriptionConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;  
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword;    
                "SubscriptionID"               = $endpointConfiguration.SubscriptionID 
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "BoxStorage") {
        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $boxStorageConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.BoxStorageConfiguration' -Property @{  
            }
        }
        else {
            $boxStorageConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.BoxStorageConfiguration' -Property @{  
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $boxStorageConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $boxStorageConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "DropBox") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $dropBoxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.DropBoxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativePassword"       = ""
            }
        }
        else {
            $dropBoxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.DropBoxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativePassword"       = ""
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $dropBoxConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $dropBoxConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "Gmail") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $GoogleMailboxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleMailboxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "Domains"                      = $Domains;
                "ContactHandling"              = 'MigrateSuggestedContacts';
            }
        }
        else {
            $adminUsername = $endpointConfiguration.AdministrativeUsername
            $GoogleMailboxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleMailboxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "Domains"                      = $endpointConfiguration.Domains;
                "ContactHandling"              = 'MigrateSuggestedContacts';
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $GoogleMailboxConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $GoogleMailboxConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "GSuite") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                
                $result = Get-FileName $script:workingDir

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
          
            $GoogleMailboxConfiguration = New-BT_GSuiteConfiguration -AdministrativeUsername $adminUsername `
                -CredentialsFileName $script:inputFile `
                -Credentials $jsonFileContent.ToString()   

        }
        else {
            $adminUsername = $endpointConfiguration.AdministrativeUsername
            do {
                Write-host -NoNewline "Please enter the file path to the Google service account credentials using JSON file:"

                $result = Get-FileName $script:workingDir

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
            $GoogleMailboxConfiguration = New-BT_GSuiteConfiguration -AdministrativeUsername $adminUsername `
                -CredentialsFileName $script:inputFile `
                -Credentials $jsonFileContent.ToString()   
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $GoogleMailboxConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $GoogleMailboxConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "GoogleDrive") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $GoogleDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleDriveConfiguration' -Property @{              
                "AdminEmailAddress" = $adminUsername;
                "Domains"           = $Domains;
            }
        }
        else {
            $GoogleDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.GoogleDriveConfiguration' -Property @{              
                "AdminEmailAddress" = $endpointConfiguration.AdminEmailAddress;
                "Domains"           = $endpointConfiguration.Domains;
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $GoogleDriveConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $GoogleDriveConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "ExchangeServer") {
        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $url = (Read-Host -prompt "Please enter the Exchange Server 2003+ URL").trim()
            }while ($url -eq "")
        
            $msg = "INFO: Exchange Server 2003+ URL is '$url'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $exchangeOnline2Configuration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangeConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "Url"                          = $url
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword
            }
        }
        else {
            $exchangeOnline2Configuration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangeConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "Url"                          = $url
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword
            }
        }

        try {
            
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $exchangeOnline2Configuration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $exchangeOnline2Configuration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "ExchangeOnline2") {
        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $exchangeOnline2Configuration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangeConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword
            }
        }
        else {
            $exchangeOnline2Configuration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangeConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword
            }
        }

        try {
            
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $exchangeOnline2Configuration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $exchangeOnline2Configuration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "ExchangeOnlinePublicFolder") {
        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $exchangePublicFolderConfiguration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangePublicFolderConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword
            }
        }
        else {
            $exchangePublicFolderConfiguration = New-Object -TypeName "ManagementProxy.ManagementService.ExchangePublicFolderConfiguration" -Property @{
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword
            }
        }

        try {
            
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $exchangePublicFolderConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $exchangePublicFolderConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "Office365Groups") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $office365GroupsConfiguration = New-Object -TypeName "ManagementProxy.ManagementService.SharePointConfiguration" -Property @{
                "Url"                          = $url;
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword
            }
        }
        else {
            $office365GroupsConfiguration = New-Object -TypeName "ManagementProxy.ManagementService.SharePointConfiguration" -Property @{
                "Url"                          = $endpointConfiguration.Url;
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword
            }
        }

        try {
            
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $office365GroupsConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 
                
                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $office365GroupsConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "OneDrivePro") {
        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $oneDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword
            }
        }
        else {
            $oneDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $oneDriveConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $oneDriveConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "OneDriveProAPI") {
        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            do {
                $confirm = (Read-Host -prompt "Do you want to use your own Azure Storage account?  [Y]es or [N]o")
                if ($confirm.ToLower() -eq "y") {
                    $microsoftStorage = $false
                }
                else {
                    $microsoftStorage = $true
                }
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 

            if (!$microsoftStorage) {
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
            }
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            if ($microsoftStorage) {
                $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $adminUsername;
                    "AdministrativePassword"             = $adminPassword;
                    "UseSharePointOnlineProvidedStorage" = $true
                }
            }
            else {
                $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                    "UseAdministrativeCredentials" = $true;
                    "AdministrativeUsername"       = $adminUsername;
                    "AdministrativePassword"       = $adminPassword;
                    "AzureStorageAccountName"      = $azureAccountName;
                    "AzureAccountKey"              = $secretKey
                }
            }
        }
        else {
            if ($endpointConfiguration.UseSharePointOnlineProvidedStorage) {
                $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "UseSharePointOnlineProvidedStorage" = $true
                }
            }
            else {
                $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                    "UseAdministrativeCredentials" = $true;
                    "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"       = $endpointConfiguration.administrativePassword;
                    "AzureStorageAccountName"      = $endpointConfiguration.AzureStorageAccountName;
                    "AzureAccountKey"              = $azureAccountKey
                }
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $oneDriveProAPIConfiguration -ErrorAction Stop
                 
                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $oneDriveProAPIConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "SharePoint") {
        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{   
                "Url"                          = $Url;           
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword
            }
        }
        else {
            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{  
                "Url"                          = $endpointConfiguration.Url;             
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword
            }
        }

        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $spoConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 
                
                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $spoConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "SharePointOnlineAPI") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            <#
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
            #>
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{   
                "Url"                                = $Url;               
                "UseAdministrativeCredentials"       = $true;
                "AdministrativeUsername"             = $adminUsername;
                "AdministrativePassword"             = $adminPassword;
                #"AzureStorageAccountName" = $azureAccountName;
                #"AzureAccountKey" = $secretKey
                "UseSharePointOnlineProvidedStorage" = $true 
            }
        }
        else {
            if ($endpointConfiguration.UseSharePointOnlineProvidedStorage) {
                $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{   
                    "Url"                                = $endpointConfiguration.Url;              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "UseSharePointOnlineProvidedStorage" = $true 
                }
            }
            else {
                $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{   
                    "Url"                                = $endpointConfiguration.Url;              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "AzureStorageAccountName"            = $endpointConfiguration.AzureStorageAccountName;
                    "AzureAccountKey"                    = $endpointConfiguration.azureAccountKey
                    "UseSharePointOnlineProvidedStorage" = $false 
                }            
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $spoConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $spoConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "SharePointBeta") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            <#
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
            #>
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointBetaConfiguration' -Property @{   
                "Url"                                = $Url;               
                "UseAdministrativeCredentials"       = $true;
                "AdministrativeUsername"             = $adminUsername;
                "AdministrativePassword"             = $adminPassword;
                #"AzureStorageAccountName" = $azureAccountName;
                #"AzureAccountKey" = $secretKey
                "UseSharePointOnlineProvidedStorage" = $true 
            }
        }
        else {
            if ($endpointConfiguration.UseSharePointOnlineProvidedStorage) {
                $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointBetaConfiguration' -Property @{   
                    "Url"                                = $endpointConfiguration.Url;              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "UseSharePointOnlineProvidedStorage" = $true 
                }
            }
            else {
                $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointBetaConfiguration' -Property @{   
                    "Url"                                = $endpointConfiguration.Url;              
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "AzureStorageAccountName"            = $endpointConfiguration.AzureStorageAccountName;
                    "AzureAccountKey"                    = $endpointConfiguration.azureAccountKey
                    "UseSharePointOnlineProvidedStorage" = $false 
                }            
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $spoConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $spoConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "MicrosoftTeamsSourceParallel") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            <#
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
            #>
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsSourceConfiguration' -Property @{          
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $adminUsername;
                "AdministrativePassword"       = $adminPassword;
            }
        }
        else {
            $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsSourceConfiguration' -Property @{             
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword"       = $endpointConfiguration.administrativePassword;
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $teamsConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $teamsConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "MicrosoftTeamsDestinationParallel") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adminPassword)
            $adminPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            <#
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
            #>
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsDestinationConfiguration' -Property @{          
                "UseAdministrativeCredentials"       = $true;
                "AdministrativeUsername"             = $adminUsername;
                "AdministrativePassword"             = $adminPassword;
                #"AzureStorageAccountName" = $azureAccountName;
                #"AzureAccountKey" = $secretKey
                "UseSharePointOnlineProvidedStorage" = $true 
            }
        }
        else {
            if ($endpointConfiguration.UseSharePointOnlineProvidedStorage) {
                $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsDestinationConfiguration' -Property @{             
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "UseSharePointOnlineProvidedStorage" = $true 
                }
            }
            else {
                $teamsConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.MsTeamsDestinationConfiguration' -Property @{             
                    "UseAdministrativeCredentials"       = $true;
                    "AdministrativeUsername"             = $endpointConfiguration.AdministrativeUsername;
                    "AdministrativePassword"             = $endpointConfiguration.administrativePassword;
                    "AzureStorageAccountName"            = $endpointConfiguration.AzureStorageAccountName;
                    "AzureAccountKey"                    = $azureAccountKey
                    "UseSharePointOnlineProvidedStorage" = $false 
                }       
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $teamsConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $teamsConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "Pst") {

        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
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

            do {
                $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key ").trim()
            }while ($secretKey -eq "")


            do {
                $containerName = (Read-Host -prompt "Please enter the container name").trim()
            }while ($containerName -eq "")

            $msg = "INFO: Azure subscription ID is '$containerName'."
            Write-Host $msg
            Log-Write -Message $msg 

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $azureAccountName;     
                "AccessKey"                    = $secretKey;  
                "ContainerName"                = $containerName;       
            }
        }
        else {
            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername"       = $endpointConfiguration.AdministrativeUsername;  
                "AccessKey"                    = $endpointConfiguration.AccessKey;    
                "ContainerName"                = $endpointConfiguration.ContainerName 
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "IMAP") {

        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $hostName = (Read-Host -prompt "Please enter the server name").trim()
            }while ($hostName -eq "")

            $msg = "INFO: Server name is '$hostName'."
            Write-Host $msg
            Log-Write -Message $msg 

            do {
                $portNumber = (Read-Host -prompt "Please enter server port").trim()
            }while ($portNumber -eq "" -and (isNumeric($portNumber)))

            do {
                $confirm = (Read-Host -prompt "Is SSL enabled?  [Y]es or [N]o")
                if ($confirm.ToLower() -eq "y") {
                    $UseSsl = $true
                }
                else {
                    $UseSsl = $false
                }
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $imapConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.HostConfiguration' -Property @{  
                "Host"                         = $hostName;     
                "Port"                         = $portNumber; 
                "UseSsl"                       = $UseSsl; 
                "UseAdministrativeCredentials" = $false;       
            }
        }
        else {
            $imapConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.HostConfiguration' -Property @{         
                "Host"                         = $endpointConfiguration.hostName;     
                "Port"                         = $endpointConfiguration.portNumber;  
                "UseSsl"                       = .$endpointConfiguration.UseSsl; 
                "UseAdministrativeCredentials" = $false;   
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $imapConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $imapConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    elseif ($endpointType -eq "Lotus") {

        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if ($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $extractorName = (Read-Host -prompt "Please enter the Lotus Extractor name (bt- identified)").trim()
            }while ($extractorName -eq "")

            $msg = "INFO: Lotus Extractor name is '$extractorName'."
            Write-Host $msg
            Log-Write -Message $msg 

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $imapConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.ExtractorConfiguration' -Property @{  
                "ExtractorName"                = $extractorName;     
                "UseAdministrativeCredentials" = $true;       
            }
        }
        else {
            $imapConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.ExtractorConfiguration' -Property @{         
                "ExtractorName"                = $endpointConfiguration.extractorName;     
                "UseAdministrativeCredentials" = $true;   
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if ( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $script:CustomerTicket -Name $endpointName -Type $endpointType -Configuration $imapConfiguration -ErrorAction Stop

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 

                if ($updateEndpoint) {
                    $updatedEnpoint = Set-BT_Endpoint -Ticket $script:CustomerTicket -endpoint $checkEndpoint -Name $endpointName -Type $endpointType -Configuration $imapConfiguration 

                    $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' updated."
                    Write-Host -ForegroundColor Blue $msg
                    Log-Write -Message $msg 
                
                }

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
    <#
        elseif(endpointType -eq "WorkMail"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name WorkMailRegion -Value $endpoint.WorkMailRegion

             
        }
        elseif(endpointType -eq "Zimbra"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif(endpointType -eq "ExchangeOnlinePublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            
        }
        elseif(endpointType -eq "ExchangeOnlineUsGovernment"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            
        }
        elseif(endpointType -eq "ExchangeOnlineUsGovernmentPublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            
        }
        elseif(endpointType -eq "ExchangeServer"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            
        }
        elseif(endpointType -eq "ExchangeServerPublicFolder"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            
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
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name TrustedAppKey -Value "ChangeMe"

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

# Function to get endpoint data
Function Get-MSPC_EndpointData {
    param 
    (      
        [parameter(Mandatory = $true)] [guid]$customerOrganizationId,
        [parameter(Mandatory = $true)] [guid]$endpointId
    )

    $script:CustomerTicket = Get-BT_Ticket -OrganizationId $customerOrganizationId

    try {
        $endpoint = Get-BT_Endpoint -Ticket $script:CustomerTicket -Id $endpointId -IsDeleted False -IsArchived False | Select-Object -Property Name, Type -ExpandProperty Configuration   
        
        $msg = "SUCCESS: Endpoint '$($endpoint.Name)' Administrative Username retrieved." 
        write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg  

        if ($endpoint.Type -eq "AzureFileSystem") {

            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessKey -Value $endpoint.AccessKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ContainerName -Value $endpoint.ContainerName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername

            return $endpointCredentials        
        }
        elseif ($endpoint.Type -eq "AzureSubscription") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name SubscriptionID -Value $endpoint.SubscriptionID

            return $endpointCredentials
        
        } 
        elseif ($endpoint.Type -eq "BoxStorage") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessToken -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name RefreshToken -Value $administrativePassword
            return $endpointCredentials
        }
        elseif ($endpoint.Type -eq "DropBox") {
            $endpointCredentials = New-Object PSObject
            return $endpointCredentials
        }
        elseif ($endpoint.Type -eq "ExchangeOnline2") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "ExchangeOnlinePublicFolder") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "ExchangeOnlineUsGovernment") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "ExchangeOnlineUsGovernmentPublicFolder") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "ExchangeServer") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "ExchangeServerPublicFolder") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "Gmail") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            if ($script:userMailboxesWithResourceMailboxes -or $script:resourceMailboxes) {
                Export-GoogleResources $endpoint.UseAdministrativeCredentials
            }
            
            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "GSuite") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name CredentialsFileName -Value $endpoint.CredentialsFileName

            if ($script:userMailboxesWithResourceMailboxes -or $script:resourceMailboxes) {
                Export-GoogleResources $endpoint.UseAdministrativeCredentials
            }
            
            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "GoogleDrive") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdminEmailAddress -Value $endpoint.AdminEmailAddress
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "GoogleVault") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdminEmailAddress -Value $endpoint.AdminEmailAddress
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Domains -Value $endpoint.Domains

            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "GroupWise") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name TrustedAppKey -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "IMAP") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Host -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Port -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseSsl -Value $administrativePassword

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "Lotus") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ExtractorName -Value $endpoint.UseAdministrativeCredentials

            $msg = "INFO: Extractor Name '$($endpoint.ExtractorName)'." 
            write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg  

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "Office365Groups") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials     
        }
        elseif ($endpoint.Type -eq "OneDrivePro") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "OneDriveProAPI") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureStorageAccountName -Value $endpoint.AzureStorageAccountName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureAccountKey -Value $azureAccountKey

            return $endpointCredentials   
        }
        elseif ($endpoint.Type -eq "OX") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "Pst") {

            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessKey -Value $endpoint.AccessKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ContainerName -Value $endpoint.ContainerName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials

            return $endpointCredentials        
        }
        elseif ($endpoint.Type -eq "PstInternalStorage") {

            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessKey -Value $endpoint.AccessKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ContainerName -Value $endpoint.ContainerName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials

            return $endpointCredentials        
        }
        elseif ($endpoint.Type -eq "SharePoint") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials     
        }
        elseif ($endpoint.Type -eq "SharePointOnlineAPI") {
            
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
        elseif ($endpoint.Type -eq "MicrosoftTeamsSource") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials     
        }
        elseif ($endpoint.Type -eq "MicrosoftTeamsDestination") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureStorageAccountName -Value $endpoint.AzureStorageAccountName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureAccountKey -Value $azureAccountKey

            return $endpointCredentials     
        }
        elseif ($endpoint.Type -eq "MicrosoftTeamsSourceParallel") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials     
        }
        elseif ($endpoint.Type -eq "MicrosoftTeamsDestinationParallel") {
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureStorageAccountName -Value $endpoint.AzureStorageAccountName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureAccountKey -Value $azureAccountKey

            return $endpointCredentials     
        }
        elseif ($endpoint.Type -eq "WorkMail") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name WorkMailRegion -Value $endpoint.WorkMailRegion

            return $endpointCredentials  
        }
        elseif ($endpoint.Type -eq "Zimbra") {
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $administrativePassword

            return $endpointCredentials  
        }

    }
    catch {
        $msg = "ERROR: Failed to retrieve endpoint '$($endpoint.Name)' credentials."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
    }
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

        if ($workgroupsPage) {
            $workgroups += @($workgroupsPage)
        }

        $workgroupOffset += 1
    } while ($workgroupsPage)

    $workgroupOffSet = 0

    do { 
        try {
            #add all the workgroups including the default workgroup, so there will be 2 default workgroups
            $workgroupsPage = @(Get-BT_Workgroup -PageOffset $workgroupOffSet -PageSize $workgroupPageSize -IsDeleted false | Where-Object { $_.CreatedBySystemUserId -ne $script:ticket.SystemUserId })   
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC workgroups."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }
        if ($workgroupsPage) {
            $workgroups += @($workgroupsPage)
            foreach ($Workgroup in $workgroupsPage) {
                Write-Progress -Activity ("Retrieving workgroups (" + $($workgroups.Length - 1) + ")") -Status $Workgroup.Id
            }

            $workgroupOffset += $workgroupPageSize
        }
    } while ($workgroupsPage)

    if ($workgroups -ne $null -and $workgroups.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: " + $($workgroups.Length - 1).ToString() + " Workgroup(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No workgroups found." 
        Exit
    }

    #######################################
    # Prompt for the mailbox Workgroup
    #######################################
    if ($workgroups -ne $null) {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a Workgroup:" 
        Write-Host -Object "INFO: A default workgroup has no name, only Id. Your default workgroup is the number 0 in yellow." 

        for ($i = 0; $i -lt $workgroups.Length; $i++) {
            
            $Workgroup = $workgroups[$i]

            if ([string]::IsNullOrEmpty($Workgroup.Name)) {
                if ($i -eq 0) {
                    $defaultWorkgroupId = $Workgroup.Id.Guid
                    Write-Host -ForegroundColor Yellow -Object $i, "-", $defaultWorkgroupId
                }
                else {
                    if ($Workgroup.Id -ne $defaultWorkgroupId) {
                        Write-Host -Object $i, "-", $Workgroup.Id
                    }
                }
            }
            else {
                Write-Host -Object $i, "-", $Workgroup.Name
            }
        }
        Write-Host -Object "x - Exit"
        Write-Host

        do {
            if ($workgroups.count -eq 1) {
                $msg = "INFO: There is only one workgroup. Selected by default."
                Write-Host -ForegroundColor yellow  $msg
                Log-Write -Message $msg
                $Workgroup = $workgroups[0]
                $global:btWorkgroupOrganizationId = $Workgroup.WorkgroupOrganizationId
                Return $Workgroup.Id
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($workgroups.Length - 1) + ", or x")
            }
            
            if ($result -eq "x") {
                Exit
            }
            if (($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $workgroups.Length)) {
                $Workgroup = $workgroups[$result]
                $global:btWorkgroupOrganizationId = $Workgroup.WorkgroupOrganizationId
                Return $Workgroup.Id
            }
        }
        while ($true)

    }

}

# Function to display all customers
Function Select-MSPC_Customer {

    param 
    (      
        [parameter(Mandatory = $true)] [String]$WorkgroupId
    )

    #######################################
    # Display all mailbox customers
    #######################################

    $customerPageSize = 100
    $customerOffSet = 0
    $customers = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC customers..."

    do {   
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
    
        if ($customersPage) {
            $customers += @($customersPage)
            foreach ($customer in $customersPage) {
                Write-Progress -Activity ("Retrieving customers (" + $customers.Length + ")") -Status $customer.CompanyName
            }
            
            $customerOffset += $customerPageSize
        }

    } while ($customersPage)

    

    if ($customers -ne $null -and $customers.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: " + $customers.Length.ToString() + " customer(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No customers found." 
        Return "-1"
    }

    #######################################
    # {Prompt for the mailbox customer
    #######################################
    if ($customers -ne $null) {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a customer:" 

        for ($i = 0; $i -lt $customers.Length; $i++) {
            $customer = $customers[$i]
            Write-Host -Object $i, "-", $customer.CompanyName
        }
        Write-Host -Object "b - Go back to workgroup selection menu"
        Write-Host -Object "x - Exit"
        Write-Host

        do {
            if ($customers.count -eq 1) {
                $msg = "INFO: There is only one customer. Selected by default."
                Write-Host -ForegroundColor yellow  $msg
                Log-Write -Message $msg
                $customer = $customers[0]

                try {
                    if ($script:confirmImpersonation) {
                        $script:CustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ImpersonateId $global:btMspcSystemUserId -ErrorAction Stop
                    }
                    else {
                        $script:CustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ErrorAction Stop
                    }
                }
                Catch {
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket under Select-MSPC_Customer()." 
                }

                $global:btcustomerName = $Customer.CompanyName

                Return $customer
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($customers.Length - 1) + ", b or x")
            }

            if ($result -eq "b") {
                Return "-1"
            }
            if ($result -eq "x") {
                Exit
            }
            if (($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $customers.Length)) {
                $customer = $customers[$result]
    
                try {
                    if ($script:confirmImpersonation) {
                        $script:CustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ImpersonateId $global:btMspcSystemUserId -ErrorAction Stop
                    }
                    else { 
                        $script:CustomerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId.Guid -ErrorAction Stop
                    }
                }
                Catch {
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket under Select-MSPC_Customer()." 
                }

                $global:btcustomerName = $Customer.CompanyName

                Return $Customer
            }
        }
        while ($true)

    }

}

# Function to display all endpoints under a customer
Function Select-MSPC_Endpoint {
    param 
    (      
        [parameter(Mandatory = $true)] [guid]$customerOrganizationId,
        [parameter(Mandatory = $false)] [String]$endpointType,
        [parameter(Mandatory = $false)] [String]$endpointName,
        [parameter(Mandatory = $false)] [object]$endpointConfiguration,
        [parameter(Mandatory = $false)] [String]$exportOrImport,
        [parameter(Mandatory = $false)] [String]$projectType,
        [parameter(Mandatory = $false)] [boolean]$deleteEndpointType

    )

    #####################################################################################################################
    # Display all MSPC endpoints. If $endpointType is provided, only endpoints of that type
    #####################################################################################################################

    $endpointPageSize = 100
    $endpointOffSet = 0
    $endpoints = $null

    $sourceMailboxEndpointList = @("ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment", "Gmail", "IMAP", "GroupWise", "zimbra", "OX", "WorkMail", "Lotus", "Office365Groups")
    $destinationeMailboxEndpointList = @("ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment", "Gmail", "IMAP", "OX", "WorkMail", "Office365Groups", "Pst")
    $sourceStorageEndpointList = @("OneDrivePro", "OneDriveProAPI", "SharePoint", "SharePointOnlineAPI", "GoogleDrive", "AzureFileSystem", "BoxStorage"."DropBox", "Office365Groups")
    $destinationStorageEndpointList = @("OneDrivePro", "OneDriveProAPI", "SharePoint", "SharePointOnlineAPI", "GoogleDrive", "BoxStorage"."DropBox", "Office365Groups")
    $sourceArchiveEndpointList = @("ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment", "GoogleVault", "PstInternalStorage", "Pst")
    $destinationArchiveEndpointList = @("ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment", "Gmail", "IMAP", "OX", "WorkMail", "Office365Groups", "Pst")
    $sourcePublicFolderEndpointList = @("ExchangeServerPublicFolder", "ExchangeOnlinePublicFolder", "ExchangeOnlineUsGovernmentPublicFolder")
    $destinationPublicFolderEndpointList = @("ExchangeServerPublicFolder", "ExchangeOnlinePublicFolder", "ExchangeOnlineUsGovernmentPublicFolder", "ExchangeServer", "ExchangeOnline2", "ExchangeOnlineUsGovernment")
    $sourceTeamworkEndpointList = @("MicrosoftTeamsSourceParallel")
    $destinationTeamworkEndpointList = @("MicrosoftTeamsDestinationParallel")

    Write-Host
    if ($endpointType -ne "") {
        Write-Host -Object  "INFO: Retrieving MSPC $exportOrImport $endpointType endpoints..."
    }
    else {
        Write-Host -Object  "INFO: Retrieving MSPC $exportOrImport endpoints..."

        if ($projectType -ne "") {
            switch ($projectType) {
                "Mailbox" {
                    if ($exportOrImport -eq "Source") {
                        $availableEndpoints = $sourceMailboxEndpointList 
                    }
                    else {
                        $availableEndpoints = $destinationeMailboxEndpointList
                    }
                }

                "Storage" {
                    if ($exportOrImport -eq "Source") { 
                        $availableEndpoints = $sourceStorageEndpointList
                    }
                    else {
                        $availableEndpoints = $destinationStorageEndpointList
                    }
                }

                "Archive" {
                    if ($exportOrImport -eq "Source") {
                        $availableEndpoints = $sourceArchiveEndpointList 
                    }
                    else {
                        $availableEndpoints = $destinationArchiveEndpointList
                    }
                }

                "PublicFolder" {
                    if ($exportOrImport -eq "Source") { 
                        $availableEndpoints = $publicfolderEndpointList
                    }
                    else {
                        $availableEndpoints = $publicfolderEndpointList
                    }
                } 

                "Teamwork" {
                    if ($exportOrImport -eq "Source") { 
                        $availableEndpoints = $sourceTeamworkEndpointList
                    }
                    else {
                        $availableEndpoints = $destinationTeamworkEndpointList
                    }
                } 
            }          
        }
    }

    $script:CustomerTicket = Get-BT_Ticket -OrganizationId $global:btCustomerOrganizationId

    do {
        try {
            if ($endpointType -ne "") {
                $endpointsPage = @(Get-BT_Endpoint -Ticket $script:CustomerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize -type $endpointType )
            }
            else {
                $endpointsPage = @(Get-BT_Endpoint -Ticket $script:CustomerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize | Sort-Object -Property Type)
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

        if ($endpointsPage) {
            
            $endpoints += @($endpointsPage)

            foreach ($endpoint in $endpointsPage) {
                Write-Progress -Activity ("Retrieving endpoint (" + $endpoints.Length + ")") -Status $endpoint.Name
            }
            
            $endpointOffset += $endpointPageSize
        }
    } while ($endpointsPage)

    Write-Progress -Activity " " -Completed

    if ($endpoints -ne $null -and $endpoints.Length -ge 1) {
        Write-Host -ForegroundColor Green "SUCCESS: $($endpoints.Length) endpoint(s) found."
    }
    else {
        Write-Host -ForegroundColor Red "INFO: No endpoints found." 
    }

    #####################################################################################################################
    # Prompt for the endpoint. If no endpoints found and endpointType provided, ask for endpoint creation
    #####################################################################################################################
    if ($endpoints -ne $null) {


        if ($endpointType -ne "") {
            
            Write-Host -ForegroundColor Yellow -Object "ACTION: Select the $exportOrImport $endpointType endpoint:" 

            for ($i = 0; $i -lt $endpoints.Length; $i++) {
                $endpoint = $endpoints[$i]
                Write-Host -Object $i, "-", $endpoint.Name
            }
        }
        elseif ($endpointType -eq "" -and $projectType -ne "") {

            Write-Host -ForegroundColor Yellow -Object "ACTION: Select the $exportOrImport $projectType endpoint:" 

            for ($i = 0; $i -lt $endpoints.Length; $i++) {
                $endpoint = $endpoints[$i]
                if ($endpoint.Type -in $availableEndpoints) {
                    
                    Write-Host $i, "- Type: " -NoNewline 
                    Write-Host -ForegroundColor White $endpoint.Type -NoNewline                      
                    Write-Host "- Name: " -NoNewline                    
                    Write-Host -ForegroundColor White $endpoint.Name   
                }
            }
        }


        Write-Host -Object "c - Create a new $endpointType endpoint"
        Write-Host -Object "x - Exit"
        Write-Host

        do {
            if ($endpoints.count -eq 1) {
                $result = Read-Host -Prompt ("Select 0, c or x")
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($endpoints.Length - 1) + ", c or x")
            }
            
            if ($result -eq "c") {
                if ($endpointName -eq "") {
                
                    if ($endpointConfiguration -eq $null) {
                        $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType                     
                    }
                    else {
                        $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration          
                    }        
                }
                else {
                    $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
                }
                Return $endpointId
            }
            elseif ($result -eq "x") {
                Exit
            }
            elseif (($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $endpoints.Length)) {
                $endpoint = $endpoints[$result]
                Return $endpoint.Id
            }
        }
        while ($true)

    } 
    elseif ($endpoints -eq $null -and $endpointType -ne "") {

        do {
            $confirm = (Read-Host -prompt "Do you want to create a $endpointType endpoint ?  [Y]es or [N]o")
        } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        if ($confirm.ToLower() -eq "y") {
            if ($endpointName -eq "") {
                $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration 
            }
            else {
                $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
            }
            Return $endpointId
        }
    }
}

# Function to
Function Get-CustomerUrlId {
    param 
    (      
        [parameter(Mandatory = $true)] [String]$customerOrganizationId
    )

    $customerUrlId = (Get-BT_Customer -OrganizationId $customerOrganizationId).Id

    Return $customerUrlId

}

# Function to delete all endpoints under a customer
Function Remove-MSPC_Endpoints {
    param 
    (      
        [parameter(Mandatory = $true)] [guid]$customerOrganizationId,
        [parameter(Mandatory = $false)] [String]$endpointType,
        [parameter(Mandatory = $false)] [String]$endpointName
    )

    $endpointPageSize = 100
    $endpointOffSet = 0
    $endpoints = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC $endpointType endpoints matching '$endpointName' endpoint name..."

    do {
        
        $endpointsPage = @(Get-BT_Endpoint -Ticket $script:CustomerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize -type $endpointType)

        if ($endpointsPage) {
            
            $endpoints += @($endpointsPage)

            foreach ($endpoint in $endpointsPage) {
                Write-Progress -Activity ("Retrieving endpoint (" + $endpoints.Length + ")") -Status $endpoint.Name
            }
            
            $endpointOffset += $endpointPageSize
        }
    } while ($endpointsPage)

    

    if ($endpoints -ne $null -and $endpoints.Length -ge 1) {
        Write-Host -ForegroundColor Green "SUCCESS: $($endpoints.Length) endpoint(s) found."
    }
    else {
        Write-Host -ForegroundColor Red "INFO: No endpoints found." 
    }

    $deletedEndpointsCount = 0

    if ($endpoints -ne $null) {
        Write-Host -ForegroundColor Yellow -Object "INFO: Deleting $endpointType endpoints:" 

        for ($i = 0; $i -lt $endpoints.Length; $i++) {
            $endpoint = $endpoints[$i]

            Try {
                if (($endpoint.Name -match "SRC-OD4B-" -and $endpointName -match "SRC-OD4B-") -or `
                    ($endpoint.Name -match "DST-OD4B-" -and $endpointName -match "DST-OD4B-") -or `
                    ($endpoint.Name -match "SRC-SPO-" -and $endpointName -match "SRC-SPO-") -or `
                    ($endpoint.Name -match "DST-SPO-" -and $endpointName -match "DST-SPO-") -or `
                    ($endpoint.Name -match "SRC-PF-" -and $endpointName -match "SRC-PF-") -or `
                    ($endpoint.Name -match "DST-PF-" -and $endpointName -match "DST-PF-") -or `
                    ($endpoint.Name -match "SRC-Teams-" -and $endpointName -match "SRC-Teams-") -or `
                    ($endpoint.Name -match "DST-Teams-" -and $endpointName -match "DST-Teams-") -or `
                    ($endpoint.Name -match "SRC-O365G-" -and $endpointName -match "SRC-O365G-") -or `
                    ($endpoint.Name -match "DST-O365G-" -and $endpointName -match "DST-O365G-")) {

                    remove-BT_Endpoint -Ticket $script:CustomerTicket -Id $endpoint.Id -force
             
                    Write-Host -ForegroundColor Green "SUCCESS: $($endpoint.Name) endpoint deleted." 
                    $deletedEndpointsCount += 1
                }

            }
            catch {
                $msg = "ERROR: Failed to delete endpoint $($endpoint.Name)."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg 
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message   
            }            
        }

        if ($deletedEndpointsCount -ge 1 ) {
            Write-Host
            Write-Host -ForegroundColor Green "SUCCESS: $deletedEndpointsCount $endpointType endpoint(s) deleted." 
        }
        elseif ($deletedEndpointsCount -eq 0) {
            Write-Host -ForegroundColor Blue "INFO: No $endpointType endpoint was deleted. They were not created by Create-O365T2TMigrationWizProjects.ps1" 
        }
    }
}

# Function to delete all mailbox connectors created by scripts
Function Remove-MW_Connectors {

    param 
    (      
        [parameter(Mandatory = $true)] [guid]$CustomerOrganizationId,
        [parameter(Mandatory = $false)] [String]$ProjectType,
        [parameter(Mandatory = $false)] [String]$ProjectName
    )
   
    $connectorPageSize = 100
    $connectorOffSet = 0
    $connectors = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving $projectType connectors matching '$ProjectName' project name..."
    
    do {   

        if ($projectType -eq "Mailbox") {
            $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:MwTicket -OrganizationId $customerOrganizationId -ProjectType "Mailbox" -PageOffset $connectorOffSet -PageSize $connectorPageSize | Sort-Object Name)
        }
        elseif ($projectType -eq "Storage") {
            $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:MwTicket -OrganizationId $customerOrganizationId -ProjectType "Storage" -PageOffset $connectorOffSet -PageSize $connectorPageSize | Sort-Object Name)
        }
        elseif ($projectType -eq "Archive") {
            $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:MwTicket -OrganizationId $customerOrganizationId -ProjectType "Archive" -PageOffset $connectorOffSet -PageSize $connectorPageSize | Sort-Object Name)
        }
        elseif ($projectType -eq "Teamwork") {
            $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:MwTicket -OrganizationId $customerOrganizationId -ProjectType "Teamwork" -PageOffset $connectorOffSet -PageSize $connectorPageSize | Sort-Object Name)
        }
        elseif ($projectType -eq "PublicFolder") {
            $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:MwTicket -OrganizationId $customerOrganizationId -ProjectType "PublicFolder" -PageOffset $connectorOffSet -PageSize $connectorPageSize | Sort-Object Name)
        }

        if ($connectorsPage) {
            $connectors += @($connectorsPage)
            foreach ($connector in $connectorsPage) {
                Write-Progress -Activity ("Retrieving connectors (" + $connectors.Length + ")") -Status $connector.Name
            }

            $connectorOffset += $connectorPageSize
        }

    } while ($connectorsPage)

    if ($connectors -ne $null -and $connectors.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: " + $connectors.Length.ToString() + " $projectType connector(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No $projectType connectors found." 
        Return
    }


    $deletedMailboxConnectorsCount = 0
    $deletedDocumentConnectorsCount = 0
    if ($connectors -ne $null) {
        
        Write-Host -ForegroundColor Yellow -Object "INFO: Deleting $projectType connectors:" 

        for ($i = 0; $i -lt $connectors.Length; $i++) {
            $connector = $connectors[$i]

            Try {
                if ($projectType -eq "Storage") {
                    if ($ProjectName -match "FS-DropBox-" -and $connector.Name -match "FS-DropBox-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                    elseif ($ProjectName -match "FS-OD4B-" -and $connector.Name -match "FS-OD4B-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                    elseif ($ProjectName -match "FS-GoogleDrive-" -and $connector.Name -match "FS-GoogleDrive-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                    elseif ($ProjectName -match "OneDrive-Document-" -and $connector.Name -match "OneDrive-Document-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                    elseif ($ProjectName -match "ClassicSPOSite-Document-" -and $connector.Name -match "ClassicSPOSite-Document-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                    elseif ($ProjectName -match "O365Group-Document-" -and $connector.Name -match "O365Group-Document-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                }                    
                
                if ($projectType -eq "Mailbox") {
                    if ($ProjectName -match "Mailbox-All conversations" -and $connector.Name -match "Mailbox-All conversations") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id  -force -ErrorAction Stop

                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedMailboxConnectorsCount += 1
                    }
                    elseif ($ProjectName -match "O365-Mailbox-User Mailboxes-" -and $connector.Name -match "O365-Mailbox-User Mailboxes-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id  -force -ErrorAction Stop

                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedMailboxConnectorsCount += 1
                    }
                    elseif ($ProjectName -match "O365-RecoverableItems-User Mailboxes-" -and $connector.Name -match "O365-RecoverableItems-User Mailboxes-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id  -force -ErrorAction Stop

                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedMailboxConnectorsCount += 1
                    }
                    elseif ($ProjectName -match "O365-Mailbox-Shared, Room and Equipment Mailboxes-" -and $connector.Name -match "O365-Mailbox-Shared, Room and Equipment Mailboxes-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id  -force -ErrorAction Stop

                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedMailboxConnectorsCount += 1
                    }
                }       
                
                if ($projectType -eq "Archive") {
                    if ($ProjectName -match "O365-Archive-User Mailboxes-" -and $connector.Name -match "O365-Archive-User Mailboxes-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id  -force -ErrorAction Stop

                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedMailboxConnectorsCount += 1
                    }
                }     

                if ($projectType -eq "Teamwork") {
                    if ($ProjectName -match "Teams-Collaboration-" -and $connector.Name -match "Teams-Collaboration-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                }

                if ($projectType -eq "PublicFolder") {
                    if ($ProjectName -match "O365-PublicFolder" -and $connector.Name -match "O365-PublicFolder") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $script:MwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                }

            }
            catch {
                $msg = "ERROR: Failed to delete $projectType connector $($connector.Name)."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg 
            } 
        }

        
        if (($deletedDocumentConnectorsCount -ge 1 -and $projectType -eq "Storage") -or ($deletedMailboxConnectorsCount -ge 1 -and $projectType -eq "Mailbox")) {
            Write-Host
            Write-Host -ForegroundColor Green "SUCCESS: $deletedDocumentConnectorsCount $projectType connector(s) deleted." 
        }
        elseif (($deletedDocumentConnectorsCount -eq 0 -and $projectType -eq "Storage") -or ($deletedMailboxConnectorsCount -eq 0 -and $projectType -eq "Mailbox")) {
            if ($projectName -match "FS-OD4B-") {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Migrate-MW_AzureBlobContainerToOD4B.ps1."    
            }
            elseif ($projectName -match "FS-DropBox-") {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Create-MW_AzureBlobContainerToDropBox.ps1."    
            }    
            elseif ($projectName -match "O365Group-Document-" -or $projectName -match "ClassicSPOSite-Document-") {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Create-O365T2TMigrationWizProjects.ps1."   
            } 
            elseif ($projectName -match "OneDrive-Document-") {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Create-O365T2TMigrationWizProjects.ps1."   
            }     
            elseif ($projectName -match "O365-Archive-User Mailboxes-") {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Create-O365T2TMigrationWizProjects.ps1."   
            }
            elseif ($projectName -match "O365-Mailbox-User Mailboxes-" -or $projectName -match "O365-RecoverableItems-User Mailboxes-" -or $projectName -match "O365Group-Mailbox-All conversations" ) {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Create-O365T2TMigrationWizProjects.ps1."   
            }
            elseif ($projectName -match "Teams-Collaboration-") {
                Write-Host -ForegroundColor Blue "INFO: No $projectType connector was deleted. They were not created by Create-O365T2TMigrationWizProjects.ps1."   
            }  
        }
    }
}

#######################################################################################################################
#        CONNECTION TO SOURCE AND/OR DESTINATION O365 / SPO
#######################################################################################################################

# Function to create source EXO PowerShell session
Function Connect-SourceExchangeOnline {

    write-host 
    $msg = "#######################################################################################################################`
                       CONNECTION TO SOURCE OFFICE 365 TENANT             `
#######################################################################################################################"
    Write-Host $msg
    Log-Write -Message $msg

    #Prompt for source Office 365 script admin Credentials

    try {
        $loginAttempts = 0
        do {
            $loginAttempts++

            # Connect to Source Exchange Online via MSPC endpoint
            if ($useMspcEndpoints) {
                #Select source endpoint
                if ($script:srcGermanyCloud) {
                    $global:btExportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType "ExchangeOnlineGermany"
                    if ($global:btExportEndpointId -eq -1) { Return -1 }
                }
                elseif ($script:srcUsGovernment) {
                    $global:btExportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType "ExchangeOnlineUsGovernment"
                    if ($global:btExportEndpointId -eq -1) { Return -1 }
                }
                else {
                    $global:btExportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType "ExchangeOnline2"
                    if ($global:btExportEndpointId -eq -1) { Return -1 }
                }
                #Get source endpoint credentials
                [PSObject]$exportEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointId $global:btExportEndpointId 

                #Create a PSCredential object to connect to source Office 365 tenant
                $script:SrcAdministrativeUsername = $exportEndpointData.AdministrativeUsername
               
                $global:btSourceO365Creds = Get-Credential -Message "Enter Your Source Office 365 Admin Credentials." -UserName $script:SrcAdministrativeUsername
                if (!($global:btSourceO365Creds)) {
                    $msg = "ERROR: Cancel button or ESC was pressed while asking for Credentials. Script will abort."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                    Exit
                }
            }
            # Connect to Source Exchange Online via manual credentials entry
            else {
                $global:btSourceO365Creds = Get-Credential -Message "Enter Your Source Office 365 Admin Credentials." -UserName $exportEndpointData.AdministrativeUsername
                if (!($global:btSourceO365Creds)) {
                    $msg = "ERROR: Cancel button or ESC was pressed while asking for Credentials. Script will abort."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                    Exit
                }
            }

            $SecureString = $global:btSourceO365Creds.Password
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
            $script:sourcePlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            $script:sourceTenantDomain = (Get-O365TenantDomain -Credentials $global:btSourceO365Creds -SourceOrDestination "source").ToLower()
                        
            if ($script:srcGermanyCloud) {
                $script:sourceTenantName = $script:sourceTenantDomain.replace(".onmicrosoft.de", "")
            }
            elseif ($script:srcUsGovernment) {
                $script:sourceTenantName = $script:sourceTenantDomain.replace(".onmicrosoft.us", "")
            }        
            else {
                $script:sourceTenantName = $script:sourceTenantDomain.replace(".onmicrosoft.com", "")
            }
            
            if (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue) {
                # The new module is installed
                Import-Module ExchangeOnlineManagement;

                if ($script:srcGermanyCloud) {                
                    Connect-ExchangeOnline -Credential $global:btSourceO365Creds -ExchangeEnvironmentName O365GermanyCloud -ShowBanner:$false  
                    $sourceEXOSession = $true
                }
                elseif ($script:srcUsGovernment) {               
                    Connect-ExchangeOnline -Credential $global:btSourceO365Creds -ExchangeEnvironmentName O365USGovGCCHigh -ShowBanner:$false  
                    $sourceEXOSession = $true
                }
                else {                
                    Connect-ExchangeOnline -Credential $global:btSourceO365Creds -ShowBanner:$false  
                    $sourceEXOSession = $true
                }

                $msg = "SUCCESS: Connection to source Office 365 '$script:sourceTenantDomain' Remote PowerShell V2."
                Write-Host -ForegroundColor Green  $msg
                Log-Write -Message $msg

                Connect-SourceMicrosoftTeams

            } 
            else {
                # the new module is not installed
            
                if ($script:srcGermanyCloud) { $script:sourceO365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office.de/powershell-liveid/ -Credential $global:btSourceO365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue }
                elseif ($script:srcUsGovernment) { $script:sourceO365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.us/powershell-liveid/ -Credential $global:btSourceO365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue }
                else { $script:sourceO365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $global:btSourceO365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue }
            
                $result = Import-PSSession -Session $script:sourceO365Session -AllowClobber -ErrorAction Stop -WarningAction silentlyContinue -DisableNameChecking -Prefix DST 

                $msg = "SUCCESS: Connection to source Office 365 '$script:sourceTenantDomain' Remote PowerShell."
                Write-Host -ForegroundColor Green  $msg
                Log-Write -Message $msg

                $msg = "WARNING: You are not using the modern management shell. It is strongly recommended that you install the new shell using Install-Module -Name ExchangeOnlineManagement"
                Write-Host -ForegroundColor Yellow  $msg
                Log-Write -Message $msg

                Connect-SourceMicrosoftTeams
            }
        }
        until (($loginAttempts -ge 3) -or ($($script:sourceO365Session.State) -eq "Opened" -or $sourceEXOSession))

        # Only 3 attempts allowed
        if ($loginAttempts -ge 3) {
            $msg = "ERROR: Failed to connect to the Source Office 365. Review your source Office 365 admin credentials and try again."
            Write-Host $msg -ForegroundColor Red
            Log-Write -Message $msg
            Start-Sleep -Seconds 5
            Exit
        }        
    }
    catch {
        $msg = "ERROR: Failed to connect to source Office 365."
        Write-Host -ForegroundColor Red $msg
        Log-Write -Message $msg
        Write-Host -ForegroundColor Red $($_.Exception.Message)
        Log-Write -Message $($_.Exception.Message)
        Get-PSSession | Remove-PSSession
        Exit
    }

    return $script:sourceO365Session

}

# Function to create destination EXO PowerShell session
Function Connect-DestinationExchangeOnline {
    write-host 
    $msg = "#######################################################################################################################`
                       CONNECTION TO DESTINATION OFFICE 365 TENANT             `
#######################################################################################################################"
    Write-Host $msg
    Log-Write -Message $msg

    #Prompt for destination Office 365 script admin Credentials

    try {
        $loginAttempts = 0
        do {
            $loginAttempts++

            # Connect to Source Exchange Online via MSPC endpoint
            if ($useMspcEndpoints) {
                #Select destination endpoint
                if ($script:dstGermanyCloud) {
                    $global:btImportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType "ExchangeOnlineGermany"
                    if ($global:btImportEndpointId -eq -1) { Return -1 }
                }
                elseif ($script:dstUsGovernment) {
                    $global:btImportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType "ExchangeOnlineUsGovernment"
                    if ($global:btImportEndpointId -eq -1) { Return -1 }
                }
                else {
                    $global:btImportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType "ExchangeOnline2"
                    if ($global:btImportEndpointId -eq -1) { Return -1 }
                }
                #Get destination endpoint credentials
                [PSObject]$importEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointId $global:btImportEndpointId 

                #Create a PSCredential object to connect to destination Office 365 tenant
                $script:DstAdministrativeUsername = $importEndpointData.AdministrativeUsername

                $global:btDestinationO365Creds = Get-Credential -Message "Enter Your Destination Office 365 Admin Credentials." -UserName $script:DstAdministrativeUsername
                if (!($global:btDestinationO365Creds)) {
                    $msg = "ERROR: Cancel button or ESC was pressed while asking for Credentials. Script will abort."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                    Exit
                }
            }
            # Connect to Source Exchange Online via manual credentials entry
            else {                
                $global:btDestinationO365Creds = Get-Credential -Message "Enter Your Destination Office 365 Admin Credentials." -UserName $importEndpointData.AdministrativeUsername
                if (!($global:btDestinationO365Creds)) {
                    $msg = "ERROR: Cancel button or ESC was pressed while asking for Credentials. Script will abort."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                    Exit
                }            
            }
            
            $SecureString = $global:btDestinationO365Creds.Password
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
            $script:destinationPlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

            $script:destinationTenantDomain = (Get-O365TenantDomain -Credentials $global:btDestinationO365Creds -SourceOrDestination "destination").ToLower()
            
            if ($script:srcGermanyCloud) {
                $script:destinationTenantName = $script:destinationTenantDomain.replace(".onmicrosoft.de", "") 
            }
            elseif ($script:srcUsGovernment) {
                $script:destinationTenantName = $script:destinationTenantDomain.replace(".onmicrosoft.us", "")
            }
            else {
                $script:destinationTenantName = $script:destinationTenantDomain.replace(".onmicrosoft.com", "")
            }

            if (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue) {
                # The new module is installed
                Import-Module ExchangeOnlineManagement;

                if ($script:dstGermanyCloud) {                
                    Connect-ExchangeOnline -Credential $global:btDestinationO365Creds -ExchangeEnvironmentName O365GermanyCloud -ShowBanner:$false -Prefix DST 
                    $destinationEXOSession = $true 
                }
                elseif ($script:dstUsGovernment) {                
                    Connect-ExchangeOnline -Credential $global:btDestinationO365Creds -ExchangeEnvironmentName O365USGovGCCHigh -ShowBanner:$false -Prefix DST
                    $destinationEXOSession = $true 
                }
                else {                
                    Connect-ExchangeOnline -Credential $global:btDestinationO365Creds -ShowBanner:$false -Prefix DST  
                    $destinationEXOSession = $true
                }

                $msg = "SUCCESS: Connection to destination Office 365 '$script:destinationTenantDomain' Remote PowerShell V2."
                Write-Host -ForegroundColor Green  $msg
                Log-Write -Message $msg

            }
            else {
                # the new module is not installed
            
                if ($script:dstGermanyCloud) { $script:destinationO365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office.de/powershell-liveid/ -Credential $global:btDestinationO365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue }
                elseif ($script:dstUsGovernment) { $script:destinationO365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.us/powershell-liveid/ -Credential $global:btDestinationO365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue }
                else { $script:destinationO365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $global:btDestinationO365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue }
            
                $result = Import-PSSession -Session $script:destinationO365Session -AllowClobber -ErrorAction Stop -WarningAction silentlyContinue -DisableNameChecking -Prefix DST 

                $msg = "SUCCESS: Connection to destination Office 365 '$script:destinationTenantDomain' Remote PowerShell."
                Write-Host -ForegroundColor Green  $msg
                Log-Write -Message $msg

                $msg = "WARNING: You are not using the modern management shell. It is strongly recommended that you install the new shell using Install-Module -Name ExchangeOnlineManagement"
                Write-Host -ForegroundColor Yellow  $msg
                Log-Write -Message $msg
            }
        }
        until (($loginAttempts -ge 3) -or ($($script:destinationO365Session.State) -eq "Opened" -or $destinationEXOSession))

        # Only 3 attempts allowed
        if ($loginAttempts -ge 3) {
            $msg = "ERROR: Failed to connect to the destination Office 365. Review your destination Office 365 admin credentials and try again."
            Write-Host $msg -ForegroundColor Red
            Log-Write -Message $msg
            Start-Sleep -Seconds 5
            Exit
        }
    }
    catch {
        $msg = "ERROR: Failed to connect to destination Office 365."
        Write-Host -ForegroundColor Red $msg
        Log-Write -Message $msg
        Write-Host -ForegroundColor Red $($_.Exception.Message)
        Log-Write -Message $($_.Exception.Message)
        Get-PSSession | Remove-PSSession
        Exit
    }

    return $script:destinationO365Session
}

# Function to connect to SharePoint Online
function Connect-SPO {

    Try {

        $msg = "INFO: Connecting to SPOService."
        Write-Host $msg
        Log-Write -Message $msg

        $spoAdminCenterUrl = "https://$script:sourceTenantName-admin.sharepoint.com/"

        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
        Connect-SPOService -url $spoAdminCenterUrl -Credential $global:btSourceO365Creds

        $msg = "SUCCESS: Connection to SPOService '$spoAdminCenterUrl'."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg

        $msg = "INFO: Connecting to PnPOnline."
        Write-Host $msg
        Log-Write -Message $msg

        if (Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline) {
            Write-Host "INFO: SharePointPnPPowerShellOnline module already installed"
        } 
        else {
            Install-Module SharePointPnPPowerShellOnline -Force
        }

        if ($useModernAuthentication) {
            Connect-PnPOnline -Url $spoAdminCenterUrl -ClientSecret $script:ClientSecret  -ClientId $script:ClientId -ErrorAction Stop
        }
        else {
            $tenantContext = Connect-PnPOnline -url $spoAdminCenterUrl -Credentials $global:btSourceO365Creds -Scopes "Group.Read.All", "User.ReadBasic.All"
        }

        #Gets the OAuth 2.0 Access Token to consume the Microsoft Graph API
        $accesstoken = Get-PnPAccessToken
        
        $msg = "SUCCESS: Connection to PnPOnline '$spoAdminCenterUrl'."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }
    catch {
        $msg = "ERROR: Failed to connect to PnPOnline. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        Write-Host -ForegroundColor Red $($_.Exception.Message)
        Log-Write -Message $($_.Exception.Message)
        Exit
    }    

    return $accesstoken
}

# Function to connect to SharePoint Online
function Connect-SourceMicrosoftTeams {
    Write-Host
    $msg = "INFO: Connecting to MicrosoftTeams."
    Write-Host $msg
    Log-Write -Message $msg 

    Try {
        $connectTeams = Connect-MicrosoftTeams -Credential $global:btSourceO365Creds

        $msg = "SUCCESS: Connection to MicrosoftTeams."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg 

        Return $connectTeams
    }
    catch {
        $msg = "ERROR: Failed to connect to MicrosoftTeams. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message 
        Exit
    }
}

#######################################################################################################################
#        FUNCTIONS WITH SOURCE AND/OR DESTINATION O365 / SPO
#######################################################################################################################

Function Create-MigrationWizServiceAccount {
    param 
    (      
        [parameter(Mandatory = $true)] [Boolean]$isSource,  
        [parameter(Mandatory = $true)] [Boolean]$needsO365License     
    )

    if ($isSource) {
        Write-Host
        
        $adminEmail = $AdministrativeUsername

        $adminMailbox = (Get-SRCMailbox -Identity $adminEmail -ErrorAction SilentlyContinue)
        if ($adminMailbox -eq $null) {
            Write-Host "INFO: Creating migration admin account $adminEmail"

            $adminMailbox = New-SRCMailbox -Name $adminEmail.split("@")[0] -MicrosoftOnlineServicesID $adminEmail -Password $script:SourceO365Creds.Password -RemotePowerShellEnabled $true -WarningAction SilentlyContinue

            Write-Host -ForegroundColor Green "SUCCESS: Migration admin account $adminEmail created in source Office 365."

        }
        else {
            Write-Host -ForegroundColor Yellow "WARNING: Migration admin account $adminEmail already exists in source Office 365."
        }

        if ($global:srcGermanyCloud) {
            Connect-MsolService -Credential $script:sourceO365Creds -AzureEnvironment AzureGermanyCloud -ErrorAction Stop
        }
        elseif ($global:srcUsGovernment) {
            Connect-MsolService -Credential $script:sourceO365Creds -AzureEnvironment USGovernment -ErrorAction Stop
        }
        else {
            Connect-MsolService -Credential $script:sourceO365Creds -ErrorAction Stop
        }
        
        $sourceMailboxLocation = (Get-MsolUser -UserPrincipalName  $script:SourceO365Creds.UserName | Select-Object UsageLocation).UsageLocation

        while ($true) {
            $adminMsolUser = Get-MsolUser -UserPrincipalName $adminEmail -ErrorAction SilentlyContinue
            if ($adminMsolUser -ne $null) {
                Break
            }

            Write-Host  "INFO: Waiting for Office 365 replication. Retry in 5 seconds."
            Start-Sleep -Seconds 5
        }

        if ($needsO365License) {            
            if ($adminMsolUser.Licenses -eq $null -or $adminMsolUser.Licenses.Count -le 0) { 
                
                if ($sourceMailboxLocation -eq $null -or $sourceMailboxLocation -eq "") {
                    do {
                        $usageLocation = (Read-Host "ACTION: Enter the two letter country location of the mailbox (i.e US)")
                    } while ($value.Length -ge 1)                    
                }
                else {
                    $usageLocation = $sourceMailboxLocation
                }

                Set-MsolUser -UserPrincipalName $adminEmail -UsageLocation $usageLocation

                do {
                    $sku = Select-O365Sku
                    if ($sku -ne $null) {
                        Write-Host
                        Write-Host "INFO: Assigning license to mailbox"

                        $licenseOption = New-MsolLicenseOptions -AccountSkuId $sku.AccountSkuId
                   
                        try { 
                            Set-MsolUserLicense -UserPrincipalName $adminEmail -AddLicenses $sku.AccountSkuId -LicenseOptions $licenseOption
                            Break
                        }
                        catch {
                            $msg = "ERROR: Failed to assign this license to $adminEmail."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg 
                            Write-Host -ForegroundColor Red $_.Exception.Message
                            Log-Write -Message $_.Exception.Message  
                        }
                    }
                    else {
                        Write-Host -ForegroundColor Red "ERROR: No Office 365 licenses found"
                    }
                }    
                while ($true)
            }
            else {
                Write-Host -ForegroundColor Yellow "WARNING: Office 365 license is already assigned"
            }

            # SPO Admin at source
            $roleName = "SharePoint Service Administrator"
            Add-MsolRoleMember -RoleMemberEmailAddress $adminEmail -RoleName $roleName -ErrorAction SilentlyContinue
        }
        else {
            Grant-O365MigrationPermissions -userPrincipalName $adminEmail -IsSource $isSource  
        }    
    }
    else {

        Write-Host
        
        $adminEmail = $AdministrativeUsername

        $adminMailbox = (Get-Mailbox -Identity $adminEmail -ErrorAction SilentlyContinue)
        if ($adminMailbox -eq $null) {
            Write-Host "INFO: Creating migration admin account $adminEmail"

            $adminMailbox = New-Mailbox -Name $adminEmail.split("@")[0] -MicrosoftOnlineServicesID $adminEmail -Password $script:DestinationO365Creds.Password -RemotePowerShellEnabled $true -WarningAction SilentlyContinue

            Write-Host -ForegroundColor Green "SUCCESS: Migration admin account $adminEmail created in destination Office 365."

        }
        else { 
            Write-Host -ForegroundColor Yellow "WARNING: Migration admin account $adminEmail already exists in destination Office 365." 
        }

        Connect-MsolService -Credential $script:destinationO365Creds -ErrorAction Stop
       
        if ($global:dstGermanyCloud) {
            Connect-MsolService -Credential $script:destinationO365Creds  -AzureEnvironment AzureGermanyCloud -ErrorAction Stop
        }
        elseif ($global:dstUsGovernment) {
            Connect-MsolService -Credential $script:destinationO365Creds  -AzureEnvironment USGovernment -ErrorAction Stop
        }
        else {
            Connect-MsolService -Credential $script:destinationO365Creds  -ErrorAction Stop
        }

        $destinationMailboxLocation = (Get-MsolUser -UserPrincipalName  $script:DestinationO365Creds.UserName | Select-Object UsageLocation).UsageLocation

        while ($true) {
            $adminMsolUser = Get-MsolUser -UserPrincipalName $adminEmail -ErrorAction SilentlyContinue
            if ($adminMsolUser -ne $null) {
                Break
            }

            Write-Host  "INFO: Waiting for Office 365 replication. Retry in 5 seconds."
            Start-Sleep -Seconds 5
        }
        
        if ($needsO365License) {
            if ($adminMsolUser.Licenses -eq $null -or $adminMsolUser.Licenses.Count -le 0) {
            
                if ($destinationMailboxLocation -eq $null -or $destinationMailboxLocation -eq "") {
                    do {
                        $usageLocation = (Read-Host "ACTION: Enter the two letter country location of the mailbox (i.e US)")
                    } while ($value.Length -ge 1)                    
                }
                else {
                    $usageLocation = $destinationMailboxLocation
                }

                Set-MsolUser -UserPrincipalName $adminEmail -UsageLocation $usageLocation

                do {
                    $sku = Select-O365Sku
                    if ($sku -ne $null) {
                        Write-Host
                        Write-Host "INFO: Assigning license to mailbox"

                        $licenseOption = New-MsolLicenseOptions -AccountSkuId $sku.AccountSkuId
                    
                        Set-MsolUserLicense -UserPrincipalName $adminEmail -AddLicenses $sku.AccountSkuId -LicenseOptions $licenseOption
                        Break
                    }
                    else {
                        throw "ERROR: No Office 365 licenses found"
                    }
                }    
                while ($true)
            }
            else {
                Write-Host -ForegroundColor Yellow "WARNING: Office 365 license is already assigned"
            }

            # O365 Global Admin at destination
            $roleName = "Company Administrator"
            Add-MsolRoleMember -RoleMemberEmailAddress $adminEmail -RoleName $roleName -ErrorAction SilentlyContinue
        }
        else {
            Grant-O365MigrationPermissions -userPrincipalName $adminEmail -IsSource $isSource 
        } 
    }

    Enable-OrganizationCustomization -ErrorAction SilentlyContinue

    New-ManagementRoleAssignment -Role "ApplicationImpersonation" -User $adminEmail
}

# Function to get the tenant domain
Function Get-O365TenantDomain {
    param 
    (      
        [parameter(Mandatory = $true)] [Object]$Credentials,
        [parameter(Mandatory = $false)] [String]$SourceOrDestination

    )

    try {
        if ($SourceOrDestination -eq "source") {
            if ($script:srcGermanyCloud) {
                Connect-MsolService -Credential $Credentials -AzureEnvironment AzureGermanyCloud -ErrorAction Stop

                $tenantDomain = @((Get-MsolDomain | Where-Object { $_.Name -match '.onmicrosoft.de' -and $_.Name -notmatch '.mail.' }).Name)    
            }
            elseif ($script:srcUsGovernment) {
                Connect-MsolService -Credential $Credentials -AzureEnvironment USGovernment -ErrorAction Stop
            
                $tenantDomain = @((Get-MsolDomain | Where-Object { $_.Name -match '.onmicrosoft.us' -and $_.Name -notmatch '.mail.' }).Name)    
            }
            else {
                Connect-MsolService -Credential $Credentials -ErrorAction Stop
            
                $tenantDomain = @((Get-MsolDomain | Where-Object { $_.Name -match '.onmicrosoft.com' -and $_.Name -notmatch '.mail.' }).Name)
            }
        }

        if ($SourceOrDestination -eq "destination") {
            if ($script:dstGermanyCloud) {
                Connect-MsolService -Credential $Credentials -AzureEnvironment AzureGermanyCloud -ErrorAction Stop

                $tenantDomain = @((Get-MsolDomain | Where-Object { $_.Name -match '.onmicrosoft.de' -and $_.Name -notmatch '.mail.' }).Name)    
            }
            elseif ($script:dstUsGovernment) {
                Connect-MsolService -Credential $Credentials -AzureEnvironment USGovernment -ErrorAction Stop
            
                $tenantDomain = @((Get-MsolDomain | Where-Object { $_.Name -match '.onmicrosoft.us' -and $_.Name -notmatch '.mail.' }).Name)    
            }
            else {
                Connect-MsolService -Credential $Credentials -ErrorAction Stop
            
                $tenantDomain = @((Get-MsolDomain | Where-Object { $_.Name -match '.onmicrosoft.com' -and $_.Name -notmatch '.mail.' }).Name)
            }
        }

        $geoDomains = @()

        if ($tenantDomain.Count -gt 1) {
            $geoLocations = @("APC"; "AUS"; "CAN"; "EUR"; "FRA"; "IND"; "JPN"; "KOR"; "NAM"; "ZAF"; "ARE"; "GBR")

            $domainArray = @()

            foreach ($domain in $tenantDomain) {
                foreach ($geoLocation in $geoLocations) {
                    if ($domain -match $geoLocation) {
                        $geoDomains += $domain
                        switch ($geoLocation) {
                            "APC" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in Asia-Pacific"
                                $apcDomain = $domain
                                $domainArray += $apcDomain
                            }
                            "AUS" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in Australia"
                                $ausDomain = $domain
                                $domainArray += $ausDomain
                            }
                            "CAN" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in Canada"
                                $canDomain = $domain
                                $domainArray += $canDomain
                            }
                            "EUR" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in Europe"
                                $eurDomain = $domain
                                $domainArray += $eurDomain
                            }
                            "FRA" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in France"
                                $fraDomain = $domain
                                $domainArray += $fraDomain
                            }
                            "IND" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in India"
                                $indDomain = $domain
                                $domainArray += $indDomain
                            }
                            "JPN" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in Japan"
                                $jpnDomain = $domain
                                $domainArray += $jpnDomain
                            }
                            "KOR" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in Korea"
                                $korDomain = $domain
                                $domainArray += $korDomain
                            }
                            "NAM" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in North America"
                                $namDomain = $domain
                                $domainArray += $namDomain
                            }
                            "ZAF" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in South Africa"
                                $zafDomain = $domain
                                $domainArray += $zafDomain
                            }
                            "ARE" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in United Arab Emirates"
                                $areDomain = $domain
                                $domainArray += $areDomain
                            }
                            "GBR" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in United Kingdom"
                                $gbrDomain = $domain
                                $domainArray += $gbrDomain
                            }      
                            default {
  
                            }                     
                        }
                    }
                }      
            }

            $mainDomain = ($tenantDomain | Where-Object { $geoDomains -NotContains $_ }) + ($geoDomains | Where-Object { $tenantDomain -NotContains $_ })
            $domainArray += $mainDomain

            write-host -ForegroundColor Yellow  "WARNING: Office 365 tenant Multi-Geo. Main domain '$mainDomain' in United Estates" 

            Write-Host 
            write-host -ForegroundColor Yellow  "ACTION: Select the domain you want to connect to:" 
            for ($i = 0; $i -lt $domainArray.Length; $i++) {
                $geoDomain = $domainArray[$i]
                Write-Host -Object $i, "-", $geoDomain
            }
            Write-Host
            $result = Read-Host -Prompt ("Select 0-" + ($domainArray.Length - 1) + ", a or x")
            $tenantDomain = $domainArray[$result]

            Return $tenantDomain
        }            
            
    }
    catch {
        $msg = "ERROR: Failed to connect to Azure Active Directory to get the tenant domain."
        Write-Host $msg -ForegroundColor Red
        Log-Write -Message $msg 
        
        do {
            $tenantDomain = Read-Host -Prompt ("Enter tenant domain or [C] to cancel")
        } while ($tenantDomain -ne "C" -and $tenantDomain -eq "")

        if ($tenantDomain -eq "C") {
            Exit
        }
    }

    Return $tenantDomain
}

# Function to get all validated domains in the tenant
Function Get-VanityDomains {
    param 
    (      
        [parameter(Mandatory = $true)] [Object]$Credentials,
        [parameter(Mandatory = $false)] [String]$SourceOrDestination
    )

    try {
        if ($SourceOrDestination -eq "source") {
            if ($script:srcGermanyCloud) {
                Connect-MsolService -Credential $Credentials -AzureEnvironment AzureGermanyCloud -ErrorAction Stop
            
                $vanityDomains = @(Get-MsolDomain | Where-Object { $_.Name -notmatch 'onmicrosoft.de' }).Name    
            }
            elseif ($script:srcUsGovernment) {
                Connect-MsolService -Credential $Credentials -AzureEnvironment USGovernment -ErrorAction Stop

                $vanityDomains = @(Get-MsolDomain | Where-Object { $_.Name -notmatch 'onmicrosoft.us' }).Name    
            }
            else {
                Connect-MsolService -Credential $Credentials -ErrorAction Stop
            
                $vanityDomains = @(Get-MsolDomain | Where-Object { $_.Name -notmatch 'onmicrosoft.com' }).Name
            }
        }

        if ($SourceOrDestination -eq "destination") {
            if ($script:dstGermanyCloud) {
                Connect-MsolService -Credential $Credentials -AzureEnvironment AzureGermanyCloud -ErrorAction Stop
            
                $vanityDomains = @(Get-MsolDomain | Where-Object { $_.Name -notmatch 'onmicrosoft.de' }).Name    
            }
            elseif ($script:dstUsGovernment) {
                Connect-MsolService -Credential $Credentials -AzureEnvironment USGovernment -ErrorAction Stop

                $vanityDomains = @(Get-MsolDomain | Where-Object { $_.Name -notmatch 'onmicrosoft.us' }).Name    
            }
            else {
                Connect-MsolService -Credential $Credentials -ErrorAction Stop
            
                $vanityDomains = @(Get-MsolDomain | Where-Object { $_.Name -notmatch 'onmicrosoft.com' }).Name
            }
        }
    }
    catch {
        $msg = "ERROR: Failed to connect to Azure Active Directory to get the vanity domains."
        Write-Host $msg -ForegroundColor Red
        Log-Write -Message $msg
        Start-Sleep -Seconds 5
        Exit
    }

    Return $vanityDomains
}

# Function to select the domain
Function Select-Domain {
    param 
    (      
        [parameter(Mandatory = $true)] [Object]$Credentials,
        [parameter(Mandatory = $false)] [Boolean]$DisplayAll,
        [parameter(Mandatory = $false)] [Boolean]$EmailAddressMapping,
        [parameter(Mandatory = $false)] [String]$sourceOrDestination

    )
    $tenantDomain = Get-O365TenantDomain -Credentials $Credentials -sourceOrDestination $sourceOrDestination
    $vanityDomains = @(Get-VanityDomains -Credentials $Credentials -sourceOrDestination $sourceOrDestination)
    $domainLength = $vanityDomains.Length

    #######################################
    # {Prompt for the domain to delete
    #######################################
    
    if ($vanityDomains -ne $null) {
        if ($vanityDomains.Count -gt 1) {            
            if ($DisplayAll) {
                Write-Host
                Write-Host -Object "INFO: Current domains added to the Office 365 tenant:" 

                for ($i = 0; $i -lt $domainLength; $i++) {
                    $vanityDomain = $vanityDomains[$i]
                    Write-Host -Object $i, "-", $vanityDomain
                }
            }
            else {
                Write-Host
                Write-Host -ForegroundColor Yellow -Object "ACTION: Select the source verified domain:" 
            
                for ($i = 0; $i -lt $domainLength; $i++) {
                    $vanityDomain = $vanityDomains[$i]
                    Write-Host -Object $i, "-", $vanityDomain
                }
                Write-Host                
                if ($EmailAddressMapping) {
                    do {
                        if ($domainLength -eq 1) {
                            Write-Host -Object "INFO: There is only 1 domain '$vanityDomains', selected by default."
                            Return $vanityDomains
                        }
                        else {
                            $result = Read-Host -Prompt ("Select the domain to migrate: 0-" + ($domainLength - 1) + ", a for [a]ll of them or c for [c]ancel")
                            if ($result -eq "a") { 
                                Return $vanityDomains
                            }
                            if ($result -eq "c") { 
                                Return $false
                            }
                            if (($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $domainLength)) {
                                $vanityDomain = $vanityDomains[$result]
                                Return $vanityDomain
                            }
                        }
                    }
                    while ($true)                
                }
                else {  
                    Write-Host  
                    do {
                        if ($domainLength -eq 1) {
                            $result = Read-Host -Prompt ("Select 0 or x")
                        }
                        else {
                            $result = Read-Host -Prompt ("Select 0-" + ($domainLength - 1) + " or x")
                        }

                        if ($result -eq "x") {
                            Exit
                        }
                        elseif (($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $domainLength)) {
                            $vanityDomain = $vanityDomains[$result]
                            Return $vanityDomain
                        }
                    }
                    while ($true)
                }

            }
        }
        elseif ($vanityDomains.Count -eq 1) {
            Return $vanityDomains        
        }

    }
    else {
        Write-Host
        Write-Host -ForegroundColor Red "INFO: There is no domain attached to the Office 365 tenant. The default domain is $tenantDomain." 
        Return $false        
    }
}

# Function to perform an Office 365 tenant assessment
Function Get-SourceTenantAssessment {
    param 
    (
        [CmdletBinding()]
        [parameter(Mandatory = $false)] [Boolean]$migrateUserMailboxes,
        [parameter(Mandatory = $false)] [Boolean]$migrateArchiveMailboxes,
        [parameter(Mandatory = $false)] [Boolean]$migrateRecoverableItems,
        [parameter(Mandatory = $false)] [Boolean]$migrateNonUserMailboxes,
        [parameter(Mandatory = $false)] [Boolean]$migrateOd4b,
        [parameter(Mandatory = $false)] [Boolean]$migratePublicFolders,
        [parameter(Mandatory = $false)] [Boolean]$migratePublicFoldersSharedMailboxes,
        [parameter(Mandatory = $false)] [Boolean]$migrateSpoTeamSites,
        [parameter(Mandatory = $false)] [Boolean]$migrateO365Groups,
        [parameter(Mandatory = $false)] [Boolean]$migrateMicrosoftTeams
    )

    $openCSVFile = !$script:migrateEntireTenant 

    if ($migrateSpoTeamSites) {
        write-Host
        $msg = "INFO: Exporting Classic SharePoint Online sites."
        write-Host $msg
        Log-Write -Message $msg 

        Export-SPOTeamSites

        $script:importedClassicTeamSites = @()
        $script:importedDocumentLibraries = @()

        ###################################################################################################
        #                        CLASSIC TEAM SITES
        ###################################################################################################
        Try {
            $script:importedClassicTeamSites = @(Import-CSV "$script:workingDir\SPOclassicTeamSites-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } )
            [Int]$script:classicTeamSitesCount = $script:importedClassicTeamSites.Count
        }
        Catch [Exception] {
            $msg = "ERROR: No Classic Team Sites have been found in CSV file generated by Assess-O365TenantAndBitTitanLicenses.ps1."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg  
            [Int]$script:classicTeamSitesCount = 0
        } 

        Try {
            $script:allDocumentLibraries = @(Import-CSV "$script:workingDir\SPODocumentLibraries-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" })
            $script:allDocumentLibrariesCount = $script:allDocumentLibraries.Count
            $script:importedDocumentLibraries = @(Import-CSV "$script:workingDir\SPODocumentLibraries-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } | Where-Object { $_.isUnderRoot -eq $false })
            $importedSubsites = @($script:importedDocumentLibraries | Where-Object { $_.isSubsite -eq $true } )            
            $script:importedRootSubSiteDocumentLibraries = @(Import-CSV "$script:workingDir\SPODocumentLibraries-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } | Where-Object { $_.isUnderRoot -eq $true -and $_.isSubsite -eq $true } )
            $script:importedRootDocumentLibraries = @(Import-CSV "$script:workingDir\SPODocumentLibraries-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } | Where-Object { $_.isUnderRoot -eq $true -and $_.isSubsite -eq $false } )

            [Int]$script:documentLibrariesCount = @($script:importedDocumentLibraries | Where-Object { $_.isSubsite -eq $false } ).Count
            [Int]$script:subsitesCount = ($importedSubsites | Select-Object TeamSiteUrl -unique).Count
            [Int]$script:subsiteDocumentLibrariesCount = $importedSubsites.Count
            [Int]$script:rootDocumentLibrariesCount = @($script:importedRootDocumentLibraries).Count
            [Int]$script:rootSubSiteDocumentLibrariesCount = @($script:importedRootSubSiteDocumentLibraries).Count
            [Int]$script:rootSubsitesCount = ($script:importedRootSubSiteDocumentLibraries | Select-Object TeamSiteUrl -unique).Count
        }
        Catch [Exception] {
            [Int]$script:documentLibrariesCount = 0
            [Int]$script:subsitesCount = 0
            [Int]$script:subsiteDocumentLibrariesCount = 0

            [Int]$script:rootDocumentLibrariesCount = 0
            [Int]$script:rootSubSiteDocumentLibrariesCount = 0
            [Int]$script:rootSubsitesCount = 0
        }   

        if ($script:rootDocumentLibrariesCount -ne 0) {
            $msg = "SUCCESS: $script:rootDocumentLibrariesCount Root Document Libraries have been found in source Office 365."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg 
        }
                
        if ($script:rootSubsitesCount -ne 0) {
            $msg = "SUCCESS: $script:rootSubsitesCount Root Subsites with $script:rootSubSiteDocumentLibrariesCount Document Libraries have been found in source Office 365."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg 
        }

        if ($script:classicTeamSitesCount -ne 0) {
            $msg = "SUCCESS: $script:classicTeamSitesCount Classic SPO Team Sites with $script:documentLibrariesCount Document Libraries have been found in source Office 365`
         - $script:subsitesCount Subsites with $script:subsiteDocumentLibrariesCount Subsites Document Libraries."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg 
        }   

        if ($openCSVFile) {        
            Write-Host
            do {
                $confirm = (Read-Host "ACTION:  If you have reviewed, edited and saved the CSV file then press [C] to continue" ) 
            } while ($confirm -ne "C")

            $script:importedClassicTeamSites = @()
            $script:importedDocumentLibraries = @()

            ###################################################################################################
            #                        CLASSIC TEAM SITES
            ###################################################################################################
            Try {
                $script:importedClassicTeamSites = @(Import-CSV "$script:workingDir\SPOclassicTeamSites-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } )
                [Int]$script:classicTeamSitesCount = $script:importedClassicTeamSites.Count
            }
            Catch [Exception] {
                $msg = "ERROR: No Classic Team Sites have been found in the saved CSV file."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg  
                [Int]$script:classicTeamSitesCount = 0
            } 

            Try {
                $script:importedDocumentLibraries = @(Import-CSV "$script:workingDir\SPODocumentLibraries-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } | Where-Object { $_.isUnderRoot -eq $false })
                $importedSubsites = @($script:importedDocumentLibraries | Where-Object { $_.isSubsite -eq $true } )            
                $script:importedRootSubSiteDocumentLibraries = @(Import-CSV "$script:workingDir\SPODocumentLibraries-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } | Where-Object { $_.isUnderRoot -eq $true -and $_.isSubsite -eq $true } )
                $script:importedRootDocumentLibraries = @(Import-CSV "$script:workingDir\SPODocumentLibraries-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } | Where-Object { $_.isUnderRoot -eq $true -and $_.isSubsite -eq $false } )

                [Int]$script:documentLibrariesCount = @($script:importedDocumentLibraries | Where-Object { $_.isSubsite -eq $false } ).Count
                [Int]$script:subsitesCount = ($importedSubsites | Select-Object TeamSiteUrl -unique).Count
                [Int]$script:subsiteDocumentLibrariesCount = $importedSubsites.Count
                [Int]$script:rootDocumentLibrariesCount = @($script:importedRootDocumentLibraries).Count
                [Int]$script:rootSubSiteDocumentLibrariesCount = @($script:importedRootSubSiteDocumentLibraries).Count
                [Int]$script:rootSubsitesCount = ($script:importedRootSubSiteDocumentLibraries | Select-Object TeamSiteUrl -unique).Count
            }
            Catch [Exception] {
                [Int]$script:documentLibrariesCount = 0
                [Int]$script:subsitesCount = 0
                [Int]$script:subsiteDocumentLibrariesCount = 0

                [Int]$script:rootDocumentLibrariesCount = 0
                [Int]$script:rootSubSiteDocumentLibrariesCount = 0
                [Int]$script:rootSubsitesCount = 0
            }   

            if ($script:rootDocumentLibrariesCount -ne 0) {
                $msg = "SUCCESS: $script:rootDocumentLibrariesCount Root Document Libraries have been found in the saved CSV file."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 
            }
                
            if ($script:rootSubsitesCount -ne 0) {
                $msg = "SUCCESS: $script:rootSubsitesCount Root Subsites with $script:rootSubSiteDocumentLibrariesCount Document Libraries have been found in the saved CSV file."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 
            }

            if ($script:classicTeamSitesCount -ne 0) {
                $msg = "SUCCESS: $script:classicTeamSitesCount Classic SPO Team Sites with $script:documentLibrariesCount Document Libraries have been found in the saved CSV file. 
         - $script:subsitesCount Subsites with $script:subsiteDocumentLibrariesCount Subsites Document Libraries."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 
            }   
        
        } # End if($openCSVFile){

    }

    if ($migrateO365Groups) {
    
        write-Host
        $msg = "INFO: Exporting Office 365 (unified) Groups."
        write-Host $msg
        Log-Write -Message $msg 

        Export-O365UnifiedGroups
        
        if ($openCSVFile) {
            $readO365GroupsCSVFile = $false
            do {
                $confirm = (Read-Host -prompt "Do you want to import a CSV file with the Office 365 (unified) group 'PrimarySmtpAddress' you want to process (tenant divestiture)?  [Y]es or [N]o")

                if ($confirm.ToLower() -eq "y") {
                    $readO365GroupsCSVFile = $true

                    Write-Host -ForegroundColor yellow "ACTION: Select the CSV file with 'PrimarySmtpAddress' column to import (Press cancel to create one)"

                    $result = Get-FileName $script:workingDir -DefaultColumnName "PrimarySmtpaddress"

                    try {
                        $unifiedGroupsInCSV = @(Import-Csv $script:inputFile)
                        $unifiedGroupsInCSV = @($unifiedGroupsInCSV.PrimarySmtpaddress.Split("@") | Select-Object -unique | Where-Object { $_ -ne $script:sourceTenantDomain -and $_ -ne $script:destinationTenantDomain })
                    }
                    catch {
                        $msg = "ERROR: Failed to import the CSV file."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg 
                    }  
                }

            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
        } #End if($openCSVFile){

        $script:importedUnifiedGroups = @()
        $script:importedUnifiedGroupDocumentLibraries = @()

        $importedTeamEnabledUnifiedGroups = @()

        ###################################################################################################
        #                        OFFICE 365 UNIFIED GROUPS
        ###################################################################################################

        
        Try {
            if ($readO365GroupsCSVFile) { 

                $importedUnifiedGroupsTemp = @()
                $importedUnifiedGroupsTemp = @(Import-CSV "$script:workingDir\O365UnifiedGroups-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } )
                $script:importedUnifiedGroups = @()   
                foreach ($importedUnifiedGroup in $importedUnifiedGroupsTemp) {            
                    if ($unifiedGroupsInCSV -notcontains $importedUnifiedGroup.PrimarySmtpAddress.Split("@")[0]) { Continue }
                    else {
                        $script:importedUnifiedGroups += $importedUnifiedGroup   
                    }
                }
                [Int]$script:unifiedGroupsCount = $script:importedUnifiedGroups.Count

                if ($script:unifiedGroupsCount -eq 0) {
                    $msg = "ERROR: No O365 (unified) group specified in the CSV file or CSV file has not been provided."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg 
                }
            }
            else {
                $script:importedUnifiedGroups = @(Import-CSV "$script:workingDir\O365UnifiedGroups-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } )
                [Int]$script:unifiedGroupsCount = $script:importedUnifiedGroups.Count
            }
        }
        Catch [Exception] {
            $msg = "ERROR: No Office 365 (unified) Groups have been found in CSV file generated by Assess-O365TenantAndBitTitanLicenses.ps1."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg  
            [Int]$script:unifiedGroupsCount = 0
        } 

        Try {
            if ($readO365GroupsCSVFile) {               
                $unifiedGroupDocumentLibraries = @(Import-CSV "$script:workingDir\O365UnifiedGroupDocumentLibraries-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } )
                foreach ($importedUnifiedGroupSharePointSiteUrl in $script:importedUnifiedGroups.SharePointSiteUrl) {
                    if (!$importedUnifiedGroupSharePointSiteUrl) { Continue }
                    $script:importedUnifiedGroupDocumentLibraries += $unifiedGroupDocumentLibraries  | Where-Object { $_.TeamSiteUrl -match ($importedUnifiedGroupSharePointSiteUrl -split ".sharepoint.com")[1] } 
                }
            }
            else {
                $script:importedUnifiedGroupDocumentLibraries = @(Import-CSV "$script:workingDir\O365UnifiedGroupDocumentLibraries-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } )                
            }
            $script:allUnifiedGroupDocumentLibrariesCount = $script:importedUnifiedGroupDocumentLibraries.Count
            $importedUnifiedGroupSites = @($script:importedUnifiedGroupDocumentLibraries | Where-Object { $_.isSubsite -eq $false } | Where-Object { $_.DocumentLibraryName -ne "Teams Wiki Data" -and $_.DocumentLibraryName -ne "SiteAssets" })
            [Int]$script:unifiedGroupDocumentLibrariesCount = $importedUnifiedGroupSites.Count
            [Int]$script:unifiedGroupSitesCount = ($importedUnifiedGroupSites | Select-Object TeamSiteUrl -unique).Count

            $importedUnifiedGroupSubsites = @($script:importedUnifiedGroupDocumentLibraries | Where-Object { $_.isSubsite -eq $true } | Where-Object { $_.DocumentLibraryName -ne "Teams Wiki Data" -and $_.DocumentLibraryName -ne "SiteAssets" })            
            [Int]$script:unifiedGroupSubSiteDocumentLibrariesCount = $importedUnifiedGroupSubsites.Count
            [Int]$script:unifiedGroupSubsitesCount = ($importedUnifiedGroupSubsites | Select-Object TeamSiteUrl -unique).Count
        }
        Catch [Exception] {
            [Int]$script:unifiedGroupDocumentLibrariesCount = 0
            [Int]$script:unifiedGroupSitesCount = 0
            [Int]$script:unifiedGroupSubSiteDocumentLibrariesCount = 0
            [Int]$script:unifiedGroupSubsitesCount = 0
        } 

        if ($script:unifiedGroupsCount -ne 0) {
            $msg = "SUCCESS: $script:unifiedGroupsCount Office 365 (unified) Groups with $script:unifiedGroupDocumentLibrariesCount Document Libraries have been found in source Office 365` 
         - $script:unifiedGroupSubsitesCount Subsites with $script:unifiedGroupSubSiteDocumentLibrariesCount Subsites Document Libraries."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg 
        }

        if ($openCSVFile) {

            $msg = "SUCCESS: CSV file '$workingDir\O365UnifiedGroups-$script:sourceTenantName.csv' processed, exported and open."
            if ($details) { Write-Host -ForegroundColor Green $msg }
            Log-Write -Message $msg
    
            try {
                if ($openCSVFile) { Start-Process -FilePath $workingDir\O365UnifiedGroups-$script:sourceTenantName.csv }
            }
            catch {
                $msg = "ERROR: Failed to find the CSV file '$workingDir\O365UnifiedGroups-$script:sourceTenantName.csv'."    
                if ($details) { Write-Host -ForegroundColor Red $msg }
                return
            }  

            $msg = "SUCCESS: CSV file '$workingDir\O365UnifiedGroupSPOSites-$script:sourceTenantName.csv' processed, exported and open."
            if ($details) { Write-Host -ForegroundColor Green $msg }
            Log-Write -Message $msg

            try {
                if ($openCSVFile) { Start-Process -FilePath $workingDir\O365UnifiedGroupSPOSites-$script:sourceTenantName.csv }
            }
            catch {
                $msg = "ERROR: Failed to find the CSV file '$workingDir\O365UnifiedGroupSPOSites-$script:sourceTenantName.csv'."    
                if ($details) { Write-Host -ForegroundColor Red $msg }
                return
            }  

            $msg = "SUCCESS: CSV file '$workingDir\O365UnifiedGroupDocumentLibraries-$script:sourceTenantName.csv' processed, exported and open."
            if ($details) { Write-Host -ForegroundColor Green $msg }
            Log-Write -Message $msg
    
            try {
                if ($openCSVFile) { Start-Process -FilePath $workingDir\O365UnifiedGroupDocumentLibraries-$script:sourceTenantName.csv }
            }
            catch {
                $msg = "ERROR: Failed to find the CSV file '$workingDir\O365UnifiedGroupDocumentLibraries-$script:sourceTenantName.csv'."    
                if ($details) { Write-Host -ForegroundColor Red $msg }
                return
            }              

            Write-Host
            do {
                $confirm = (Read-Host "ACTION:  If you have reviewed, edited and saved the CSV file then press [C] to continue" ) 
            } while ($confirm -ne "C")

            try {
                $script:importedUnifiedGroups = @(Import-CSV "$script:workingDir\O365UnifiedGroups-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } )
                [Int]$script:unifiedGroupsCount = $script:importedUnifiedGroups.Count
            }
            Catch [Exception] {
                $msg = "ERROR: No Office 365 (unified) Groups have been found in the saved CSV file."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg  
                [Int]$script:unifiedGroupsCount = 0
            }   
        
            Try {
                $script:importedUnifiedGroupDocumentLibraries = @(Import-CSV "$script:workingDir\O365UnifiedGroupDocumentLibraries-$script:sourceTenantName.csv" -Encoding UTF8 | Where-Object { $_.PSObject.Properties.Value -ne "" } )
                $script:allUnifiedGroupDocumentLibrariesCount = $script:importedUnifiedGroupDocumentLibraries.Count
                $importedUnifiedGroupSites = @($script:importedUnifiedGroupDocumentLibraries | Where-Object { $_.isSubsite -eq $false } | Where-Object { $_.DocumentLibraryName -ne "Teams Wiki Data" -and $_.DocumentLibraryName -ne "SiteAssets" })
                [Int]$script:unifiedGroupDocumentLibrariesCount = $importedUnifiedGroupSites.Count
                [Int]$script:unifiedGroupSitesCount = ($importedUnifiedGroupSites | Select-Object TeamSiteUrl -unique).Count

                $importedUnifiedGroupSubsites = @($script:importedUnifiedGroupDocumentLibraries | Where-Object { $_.isSubsite -eq $true } | Where-Object { $_.DocumentLibraryName -ne "Teams Wiki Data" -and $_.DocumentLibraryName -ne "SiteAssets" })            
                [Int]$script:unifiedGroupSubSiteDocumentLibrariesCount = $importedUnifiedGroupSubsites.Count
                [Int]$script:unifiedGroupSubsitesCount = ($importedUnifiedGroupSubsites | Select-Object TeamSiteUrl -unique).Count
            }
            Catch [Exception] {
                [Int]$script:unifiedGroupDocumentLibrariesCount = 0
                [Int]$script:unifiedGroupSitesCount = 0
                [Int]$script:unifiedGroupSubSiteDocumentLibrariesCount = 0
                [Int]$script:unifiedGroupSubsitesCount = 0
            } 

            if ($script:unifiedGroupsCount -ne 0) {
                $msg = "SUCCESS: $script:unifiedGroupsCount Office 365 (unified) Groups with $script:unifiedGroupDocumentLibrariesCount Document Libraries have been found in the saved CSV files 
         - $script:unifiedGroupSubsitesCount Subsites with $script:unifiedGroupSubSiteDocumentLibrariesCount Subsites Document Libraries."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg 
            }

        } # End if($openCSVFile){
        
    }
    
    Return
}

# Determine if the identity can be used to retrieve public folders

# Function to export classic SPO sites
Function Export-SPOTeamSites { 
 
    $details = $false
    $openCSVFile = !$script:migrateEntireTenant 

    $unifiedGroupArray = @()   
    $classicTeamSitesArray = @() 
    $documentLibrariesArray = @()
        
    $systemLists = @("Maintenance Log Library", "appdata", "TaxonomyHiddenList", "User Information List", "Composed Looks", "MicroFeed", "appfiles", "Converted Forms", "Customized Reports", "Form Templates", "Images", "List Template Gallery", "Master Page Gallery", "Pages", "Reporting Templates", "Site Assets", "Site Collection Documents", "Site Collection Images", "Site Pages", "Solution Gallery", "Style Library", "Theme Gallery", "Web Part Gallery", "wfpub")
        
    if ($script:srcGermanyCloud) {
        $sSPOAdminCenterUrl = "https://$script:sourceTenantName-admin.sharepoint.de/"
        $sSPOUrl = "https://$script:sourceTenantName.sharepoint.de/"
    }
    elseif ($script:srcUsGovernment) {
        $sSPOAdminCenterUrl = "https://$script:sourceTenantName-admin.sharepoint.us/"
        $sSPOUrl = "https://$script:sourceTenantName.sharepoint.us/"
    }
    else {
        $sSPOAdminCenterUrl = "https://$script:sourceTenantName-admin.sharepoint.com/"
        $sSPOUrl = "https://$script:sourceTenantName.sharepoint.com/"
    }

    try {
        if ($useModernAuthentication) {
            Connect-PnPOnline -Url $sSPOUrl -ClientSecret $script:ClientSecret  -ClientId $script:ClientId -ErrorAction Stop
        }
        else {
            Connect-PnPOnline -Url $sSPOUrl -Credentials $global:btSourceO365Creds -ErrorAction Stop
        }
    }
    catch {
        $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access '$sSPOUrl' root with Connect-PnPOnline."
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
        Log-Write -Message $msg
        Log-Write -Message "         $($_.Exception.Message)"
        Write-Host

        Continue
    }

    if ($details) { Write-host }
    $msg = "INFO: Exporting Root Document Libraries from source SharePoint Online to CSV file."
    if ($details) { Write-host $msg } 
    Log-Write -Message $msg 

    try {        
        $rootDocumentLibraries = @(Get-PnPList -Includes RootFolder.ServerRelativeUrl | Where-Object { $_.Title -notin $systemLists } | Where-Object { ($_.BaseType -eq 'DocumentLibrary' -and $_.BaseType -ne 'GenericList') -and $_.EntityTypeName -NotMatch 'List' -and $_.EntityTypeName -ne 'Style_x0020_Library' -and $_.EntityTypeName -ne 'Teams_x0020_Wiki_x0020_Data' }) 
        $script:rootDocumentLibrariesCount = $rootDocumentLibraries.Count
    }
    catch {
        $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($sSPOUrl.Url) root Document Library with Get-PnPList."
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
        Log-Write -Message $msg
        Log-Write -Message "         $($_.Exception.Message)"
    }
        
    if ($script:rootDocumentLibrariesCount -eq 0) {
        $msg = "INFO: No Root Document Libraries have been retrieved from Office 365."
        if ($details) { Write-host -ForegroundColor Red $msg } 
        Log-Write -Message $msg
    }
    else {
        $msg = "SUCCESS: $script:rootDocumentLibrariesCount Root Document Libraries have been retrieved from Office 365."
        if ($details) { Write-host -ForegroundColor Green $msg } 
        #Log-Write -Message $msg
    } 

    $currentRootDocumentLibrary = 0
    if ($rootDocumentLibraries -ne $null) {
        Foreach ($rootDocumentLibrary in $rootDocumentLibraries) { 

            $currentRootDocumentLibrary += 1 

            $rootDocumentLibraryName = $($rootDocumentLibrary.RootFolder.ServerRelativeUrl.replace("/", ""))
            $DocumentLibraryUrl = "$sSPOUrl$rootDocumentLibraryName"
            
            $msg = "         INFO: Exporting Root Document Library $currentRootSubWebDocLib $currentRootDocumentLibrary/$script:rootDocumentLibrariesCount of Root :  '$($rootDocumentLibrary.Title)' '$DocumentLibraryUrl'."
            if ($details) { Write-host $msg } 
            #Log-Write -Message $msg 

            $msg = "            SUCCESS: Document Library '$($rootDocumentLibrary.Title)' found under '$sSPOUrl'."
            if ($details) { Write-Host -ForegroundColor Green  $msg }
            #Log-Write -Message $msg

            $documentLibrariesArray += [PSCustomObject] @{ 
                Title                     = $rootDocumentLibrary.Title  
                DocumentLibraryName       = $rootDocumentLibraryName                      
                EntityTypeName            = $rootDocumentLibrary.EntityTypeName 
                BaseType                  = $rootDocumentLibrary.BaseType
                TeamSiteUrl               = $sSPOUrl
                NewDestinationTeamSiteUrl = $sSPOUrl
                isUnderRoot               = "TRUE"
                isSubsite                 = "FALSE"
                SubsiteDepth              = "0"
                SubsiteNumber             = "0"
                DocumentLibraryUrl        = $DocumentLibraryUrl          
            }
        }    
    }

    if ($details) { Write-host }
    $msg = "INFO: Exporting Root Subsites and their Document Libraries from source SharePoint Online to CSV file."
    if ($details) { Write-host $msg } 
    Log-Write -Message $msg 
        
    try {
        $rootSubWebs = @(Get-PnPSubWebs -Recurse -ErrorAction Stop)
        $rootSubWebsCount = $rootSubWebs.Count
    }
    catch {
        $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($sSPOUrl) SubWebs with Get-PnPSubWebs."
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
        Log-Write -Message $msg
        Log-Write -Message "         $($_.Exception.Message)"
    }

    if ($rootSubWebsCount -eq 0) {
        $msg = "INFO: No Root Subsites have been retrieved from Office 365."
        if ($details) { Write-host -ForegroundColor Red $msg } 
        Log-Write -Message $msg
    }
    else {
        $msg = "SUCCESS: $rootSubWebsCount Root Subsites have been retrieved from Office 365."
        #Write-host -ForegroundColor Green $msg 
        Log-Write -Message $msg
    } 

    if ($details) { Write-host }
    $msg = "INFO: Exporting Classic Team Sites and their Document Libraries from source SharePoint Online to CSV file."
    if ($details) { Write-host $msg } 
    Log-Write -Message $msg 
    
    $classicTeamSubsiteCount = 0
    $classicTeamRootSubsiteCount = 0
    
    $currentRootSubWeb = 1
    $currentRootSubWebDocLib = 1
    if ($rootSubWebs -ne $null) {
        if ($details) { Write-host }
        Foreach ($rootSubWeb in $rootSubWebs) { 
            
            $msg = "         INFO: Exporting Root Subsite Document Library $currentRootSubWebDocLib of $currentRootSubWeb/$rootSubWebsCount Root Subsites :  '$($rootSubWeb.Title)' '$($rootSubWeb.Url)'."
            if ($details) { Write-host $msg } 
            Log-Write -Message $msg 
            try {
                if ($useModernAuthentication) {
                    Connect-PnPOnline -Url $rootSubWeb.Url -ClientSecret $script:ClientSecret  -ClientId $script:ClientId -ErrorAction Stop
                }
                else {
                    Connect-PnPOnline -Url $rootSubWeb.Url -Credentials $global:btSourceO365Creds -ErrorAction Stop 
                }              
            }
            catch {
                $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($rootSubWeb.Url) root subsite with Connect-PnPOnline."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                Log-Write -Message $msg
                Log-Write -Message "         $($_.Exception.Message)"
                Write-Host
                Continue
            }

            try {
                $rootSubWebDocumentLibraries = @(Get-PnPList | Where-Object { $_.Title -notin $systemLists })
            }
            catch {
                $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($rootSubWeb.Url) root subsite with Get-PnPList."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                Log-Write -Message $msg
                Log-Write -Message "         $($_.Exception.Message)"
            }

            $rootSubSiteUrl = $rootSubWeb.Url
            
            $additionalRootSubsites = @()
            $urlparts = $rootSubSiteUrl.split("/")
            $rootSubsiteName = $urlparts[3]
            if ($urlparts.count -gt 4) {
                for ($i = 4; $i -lt $urlparts.length; $i++) {
                    $additionalRootSubsites += $urlparts[$i]
                    
                }
            }
            $rootUrl = ("$rootSubSiteUrl/" -split "/$rootSubsiteName/")[0]  
            #$additionalRootSubsiteNames = ($rootSubSiteUrl -split $rootSubsiteName)[2] 
            $rebuiltRootSubSiteUrl = ''
            $relativeRootSubSiteUrl = ''
            if ($urlparts.count -eq 4) {  
                $rebuiltRootSubSiteUrl = "$rootUrl/$rootSubsiteName"
            }
            elseif ($urlparts.count -gt 4) {
                foreach ($additionalRootSubsite in $additionalRootSubsites) {
                    $relativeRootSubSiteUrl += $additionalRootSubsite + "/"
                }
                $rebuiltRootSubSiteUrl = "$rootUrl/$rootSubsiteName/$relativeRootSubSiteUrl"
                $rebuiltRootSubSiteUrl = $rebuiltRootSubSiteUrl.TrimEnd("/")
            }

            if ($rootSubSiteUrl -ne $rebuiltRootSubSiteUrl) {
                $msg = "         ERROR: Unable to rebuild $($rootSubWeb.Url) subsite URL."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                Log-Write -Message $msg
                Log-Write -Message "         $($_.Exception.Message)"  
                Continue  
            }

            if ($additionalRootSubsites) {
                $subsitesArray = @($additionalRootSubsites.split('/'))
                $subsitesDepth = $subsitesArray.Count
            }
            else {
                $subsitesDepth = 0
            }            

            foreach ($rootSubWebDocumentLibrary in $rootSubWebDocumentLibraries) {

                if ($rootSubWebDocumentLibrary.DocumentTemplateUrl -eq $null) { Continue }
                if ($rootSubWebDocumentLibrary.EntityTypeName -match "SiteAssets") { Continue }
                if ($rootSubWebDocumentLibrary.EntityTypeName -match "Translation_x0020_Packages") { Continue }
                if ($rootSubWebDocumentLibrary.EntityTypeName -match "List") { Continue }
                if ($rootSubWebDocumentLibrary.DocumentTemplateUrl -notmatch "/Forms/template.dot") { Continue }
                
                $documentLibraryName = $rootSubWebDocumentLibrary.DocumentTemplateUrl.Replace("/$rootSubsiteName/$relativeRootSubSiteUrl", "").Replace("/Forms/template.dotx", "") #.Replace("/$rootSubsiteName/$relativeRootSubSiteUrl/","").split("/")[0] 
                $documentLibraryUrl = "$rebuiltRootSubSiteUrl/$documentLibraryName"  

                $msg = "            SUCCESS: Document Library '$($rootSubWebDocumentLibrary.Title)' found under '$documentLibraryUrl'."
                if ($details) { Write-Host -ForegroundColor Green  $msg }
                Log-Write -Message $msg

                $documentLibraryNames += "Subsite:$documentLibraryName"
                
                $documentLibrariesArray += [PSCustomObject] @{ 
                    Title                     = $rootSubWebDocumentLibrary.Title  
                    DocumentLibraryName       = $documentLibraryName                     
                    EntityTypeName            = $rootSubWebDocumentLibrary.EntityTypeName 
                    BaseType                  = $rootSubWebDocumentLibrary.BaseType
                    TeamSiteUrl               = $rootSubSiteUrl.replace(" ", "%20")
                    NewDestinationTeamSiteUrl = $rootSubSiteUrl.replace(" ", "%20")
                    isUnderRoot               = "TRUE"
                    isSubsite                 = "TRUE"
                    SubsiteDepth              = $subsitesDepth
                    SubsiteNumber             = "$currentRootSubWebDocLib of $rootSubWebsCount"
                    DocumentLibraryUrl        = $documentLibraryUrl          
                }

                $currentRootSubWebDocLib += 1
                $classicTeamRootSubsiteCount += 1
            }
            
            $currentRootSubWeb += 1        
        }
    }

    try {
        Connect-SPOService -Url $sSPOAdminCenterUrl -Credential $global:btSourceO365Creds -ErrorAction Stop
    }
    catch {
        $msg = "ERROR: Failed to connect to SPOService."    
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message 
        return
    }
    try {
        if ($useModernAuthentication) {
            Connect-PnPOnline -Url $sSPOAdminCenterUrl -ClientSecret $script:ClientSecret  -ClientId $script:ClientId -ErrorAction Stop
        }
        else {
            Connect-PnPOnline -Url $sSPOAdminCenterUrl -Credentials $global:btSourceO365Creds -ErrorAction Stop
        }
    }
    catch {
        $msg = "ERROR: Failed to connect to SPOService."    
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message 
        return
    }

    $classicTeamSites = Get-PnPTenantSite -Detailed | Where-Object { $_.Template -match "STS#" -and $_.Url -match $sSPOUrl }
    $classicTeamSiteCount = $classicTeamSites.Count

    if ($classicTeamSiteCount -ne 0) {
        $sessionStartTime = Get-Date
        $sessionEndTime = (Get-Date).AddHours(+5)

        if ($details) { Write-host }
        Foreach ($classicTeamSite in $classicTeamSites) {
            if ($sessionStartTime -gt $SessionEndTime) {
                Write-Host 
                $msg = "INFO: Start another Exchange Online session."
                Write-Host -ForegroundColor Magenta  $msg
                Log-Write -Message $msg
                Write-Host

                $script:sourceO365Session = Connect-SourceExchangeOnline

                if ($script:sourceO365Session.toString() -ne "-1") {
                    $script:destinationO365Session = Connect-DestinationExchangeOnline
                    if ($script:destinationO365Session.toString() -ne "-1") { Break }            
                }

                $sessionStartTime = Get-Date
                $sessionEndTime = (Get-Date).AddHours(+5)
            }

            $currentClassicTeamSite += 1 

            $msg = "      INFO: Exporting Classic Team Sites and their Document Libraries $currentClassicTeamSite/$classicTeamSiteCount :  '$($classicTeamSite.Title)' '$($classicTeamSite.Url)'."
            if ($details) { Write-host $msg } 
            Log-Write -Message $msg 

            try {        
                $result = Set-SPOUser -Site $classicTeamSite.Url -LoginName $global:btSourceO365Creds.UserName -IsSiteCollectionAdmin $true

                $msg = "      SUCCESS: '$($global:btSourceO365Creds.UserName)' added to get SPO Site Group '$($classicTeamSite.Url)' as SiteCollectionAdmin."
                if ($details) { Write-Host -ForegroundColor Green  $msg }
                Log-Write -Message $msg 

                $removeSiteCollectionAdmin = $true
            }
            catch {
                $msg = "      ERROR: Failed to set '$($global:btSourceO365Creds.UserName)' as SiteCollectionAdmin of SPO Site Group '$($classicTeamSite.Url)'. Access denied."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg 

                $removeSiteCollectionAdmin = $false
            }
        
            try {
                if ($useModernAuthentication) {
                    Connect-PnPOnline -Url $classicTeamSite.Url -ClientSecret $script:ClientSecret  -ClientId $script:ClientId -ErrorAction Stop
                }
                else {
                    Connect-PnPOnline -Url $classicTeamSite.Url -Credentials $global:btSourceO365Creds
                }
            }
            catch {
                $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($classicTeamSite.Url) site with Connect-PnPOnline."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                Log-Write -Message $msg
                Log-Write -Message "         $($_.Exception.Message)"
                Write-Host 
                Continue
            }

            try {
                $documentLibraries = @(Get-PnPList -ErrorAction Stop | Where-Object { $_.Title -notin $systemLists })  
            }
            catch {
                $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($classicTeamSite.Url) site with Get-PnPList."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                Log-Write -Message $msg
                Log-Write -Message "         $($_.Exception.Message)"
            }

            $siteUrl = $classicTeamSite.Url

            <#
            $urlparts =$siteUrl.split("/")        
            $sitesOrTeams = $urlparts[3]
            $rootUrl = ($siteUrl -split $sitesOrTeams)[0]  
            $siteName = ($siteUrl -split $sitesOrTeams)[1]    
            $rebuiltSiteUrl = "$rootUrl$sitesOrTeams$siteName"
            #>

            $additionalSubsites = @()
            $urlparts = $siteUrl.split("/")
            $sitesOrTeams = $urlparts[3]        
            $rootUrl = ("$siteUrl/" -split "/$sitesOrTeams/")[0]  
            #$siteName = ($siteUrl -split $sitesOrTeams)[1] 
            $siteName = $urlparts[4]
            if ($urlparts.count -gt 5) {
                for ($i = 5; $i -lt $urlparts.length; $i++) {
                    $additionalSubsites += $urlparts[$i]
                }
            }
            $rebuiltSiteUrl = ''
            $relativeSubSiteUrl = ''
            if ($urlparts.count -eq 5) {  
                $rebuiltSiteUrl = "$rootUrl/$sitesOrTeams/$siteName"
            }
            elseif ($urlparts.count -gt 5) {
                foreach ($additionalSubsite in $additionalSubsites) {
                    $relativeSubSiteUrl += "/" + $additionalSubsite
                }
                $rebuiltSiteUrl = "$rootUrl/$sitesOrTeams/$siteName$relativeSubSiteUrl"
            }
             
            foreach ($documentLibrary in $documentLibraries) {

                if ($documentLibrary.DocumentTemplateUrl -eq $null) { Continue }
                if ($documentLibrary.EntityTypeName -match "SiteAssets") { Continue }
                if ($documentLibrary.EntityTypeName -match "Translation_x0020_Packages") { Continue }
                if ($documentLibrary.EntityTypeName -match "List") { Continue }
                if ($documentLibrary.DocumentTemplateUrl -notmatch "/Forms/template.dot") { Continue }

                $documentLibraryName = $documentLibrary.DocumentTemplateUrl.Replace("/$sitesOrTeams/$siteName/", "").split("/")[0] 
                $documentLibraryUrl = "$rebuiltSiteUrl/$documentLibraryName"  

                $msg = "         SUCCESS: Document Library '$($documentLibrary.Title)' found under '$documentLibraryUrl'."
                if ($details) { Write-Host -ForegroundColor Green  $msg }
                Log-Write -Message $msg
                    
                $documentLibrariesArray += [PSCustomObject] @{ 
                    Title                     = $documentLibrary.Title  
                    DocumentLibraryName       = $documentLibraryName                     
                    EntityTypeName            = $documentLibrary.EntityTypeName 
                    BaseType                  = $documentLibrary.BaseType
                    TeamSiteUrl               = $classicTeamSite.Url.replace(" ", "%20")
                    NewDestinationTeamSiteUrl = $classicTeamSite.Url.replace(" ", "%20")
                    isUnderRoot               = "FALSE"
                    isSubsite                 = "FALSE"
                    SubsiteDepth              = "0"
                    SubsiteNumber             = "0"
                    DocumentLibraryUrl        = $documentLibraryUrl          
                }
            }

            try {
                $SubWebs = @()
                $SubWebs = (Get-PnPSubWebs -Recurse -ErrorAction Stop)
                $SubWebsCount = $SubWebs.Count
                $currentSubWeb = 0
            }
            catch {
                $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($classicTeamSite.Url) SubWebs with Get-PnPSubWebs."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                Log-Write -Message $msg
                Log-Write -Message "         $($_.Exception.Message)"
            }

            if ($SubWebs -ne $null) {
                Foreach ($SubWeb in $SubWebs) { 
                    $currentSubWeb += 1
    
                    $msg = "         INFO: Exporting '$($classicTeamSite.Title)' Subsite $currentSubWeb/$SubWebsCount and their Document Libraries :  '$($SubWeb.Title)' '$($SubWeb.Url)'."
                    if ($details) { Write-host $msg } 
                    Log-Write -Message $msg 
                    try {
                        if ($useModernAuthentication) {
                            Connect-PnPOnline -Url $SubWeb.Url -ClientSecret $script:ClientSecret  -ClientId $script:ClientId -ErrorAction Stop
                        }
                        else {
                            Connect-PnPOnline -Url $SubWeb.Url -Credentials $global:btSourceO365Creds
                        }
                    }
                    catch {
                        $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($SubWeb.Url) site with Connect-PnPOnline."
                        Write-Host -ForegroundColor Red  $msg
                        Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                        Log-Write -Message $msg
                        Log-Write -Message "         $($_.Exception.Message)"
                        Write-Host
                        Continue
                    }

                    try {
                        $subWebDocumentLibraries = @(Get-PnPList | Where-Object { $_.Title -notin $systemLists })  
                    }
                    catch {
                        $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($SubWeb.Url) site with Get-PnPList."
                        Write-Host -ForegroundColor Red  $msg
                        Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                        Log-Write -Message $msg
                        Log-Write -Message "         $($_.Exception.Message)"
                    }

                    $subSiteUrl = $SubWeb.Url

                    $urlparts = $subSiteUrl.split("/")      
                    $sitesOrTeams = $urlparts[3]
                    $rootUrl = ($subSiteUrl -split $sitesOrTeams)[0]  
                    $siteNames = ($subSiteUrl -split $sitesOrTeams)[1] 
                    $rebuiltSubSiteUrl = "$rootUrl$sitesOrTeams$siteNames"

                    if ($SubWeb.Url -ne $rebuiltSubSiteUrl) {
                        $msg = "         ERROR: Unable to rebuild $($SubWeb.Url) subsite URL."
                        Write-Host -ForegroundColor Red  $msg
                        Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                        Log-Write -Message $msg
                        Log-Write -Message "         $($_.Exception.Message)"  
                        Continue  
                    }

                    $subsitesArray = @($siteNames.split('/'))
                    $subsitesDepth = $subsitesArray.Count - 2

                    foreach ($subWebDocumentLibrary in $subWebDocumentLibraries) {
    
                        if ($subWebDocumentLibrary.DocumentTemplateUrl -eq $null) { Continue }
                        if ($subWebDocumentLibrary.EntityTypeName -match "SiteAssets") { Continue }
                        if ($subWebDocumentLibrary.EntityTypeName -match "Translation_x0020_Packages") { Continue }
                        if ($subWebDocumentLibrary.EntityTypeName -match "List") { Continue }
                        if ($subWebDocumentLibrary.DocumentTemplateUrl -notmatch "/Forms/template.dot") { Continue }


                        $documentLibraryName = $subWebDocumentLibrary.DocumentTemplateUrl.Replace("/$sitesOrTeams$siteNames/", "").split("/")[0] 
                        $documentLibraryUrl = "$rebuiltSubSiteUrl/$documentLibraryName"  

                        $msg = "            SUCCESS: Document Library '$($subWebDocumentLibrary.Title)' found under '$documentLibraryUrl'."
                        if ($details) { Write-Host -ForegroundColor Green  $msg }
                        Log-Write -Message $msg

                        $documentLibraryNames += "Subsite:$documentLibraryName"
                    
                        $documentLibrariesArray += [PSCustomObject] @{ 
                            Title                     = $subWebDocumentLibrary.Title  
                            DocumentLibraryName       = $documentLibraryName                     
                            EntityTypeName            = $subWebDocumentLibrary.EntityTypeName 
                            BaseType                  = $subWebDocumentLibrary.BaseType
                            TeamSiteUrl               = $subSiteUrl.replace(" ", "%20")
                            NewDestinationTeamSiteUrl = $subSiteUrl.replace(" ", "%20")
                            isUnderRoot               = "FALSE"
                            isSubsite                 = "TRUE"
                            SubsiteDepth              = $subsitesDepth
                            SubsiteNumber             = "$currentSubWeb of $SubWebsCount"
                            DocumentLibraryUrl        = $documentLibraryUrl          
                        }

                        $classicTeamSubsiteCount += 1
                    }        
                }
            }

            $classicTeamSitesArray += [PSCustomObject] @{ 
                CompatibilityLevel                       = $classicTeamSite.CompatibilityLevel                       
                LocaleId                                 = $classicTeamSite.LocaleId 
                Owner                                    = $classicTeamSite.Owner    
                ResourceQuota                            = $classicTeamSite.ResourceQuota 
                StorageQuota                             = $classicTeamSite.StorageQuota
                Template                                 = $classicTeamSite.Template  
                TimeZoneId                               = $classicTeamSite.TimeZoneId                                                                          
                Title                                    = $classicTeamSite.Title 
                Url                                      = $classicTeamSite.Url
                NewDestinationUrl                        = $classicTeamSite.Url

                AllowEditing                             = $classicTeamSite.AllowEditing       
                AllowSelfServiceUpgrade                  = $classicTeamSite.AllowSelfServiceUpgrade  
                BlockDownloadOfNonViewableFiles          = $classicTeamSite.AllowDownloadingNonWebViewableFiles  
                CommentsOnSitePagesDisabled              = $classicTeamSite.CommentsOnSitePagesDisabled  
                ConditionalAccessPolicy                  = $classicTeamSite.ConditionalAccessPolicy   
                DefaultLinkPermission                    = $classicTeamSite.DefaultLinkPermission             
                DefaultSharingLinkType                   = $classicTeamSite.DefaultSharingLinkType    
                DenyAddAndCustomizePages                 = $classicTeamSite.DenyAddAndCustomizePages   
                DisableAppViews                          = $classicTeamSite.DisableAppViews 
                DisableCompanyWideSharingLinks           = $classicTeamSite.DisableCompanyWideSharingLinks   
                DisableFlows                             = $classicTeamSite.DisableFlows   
                LimitedAccessFileType                    = $classicTeamSite.LimitedAccessFileType 
                LockState                                = $classicTeamSite.LockState                                
                RestrictedToRegion                       = $classicTeamSite.RestrictedToRegion 
                SandboxedCodeActivationCapability        = $classicTeamSite.SandboxedCodeActivationCapability        
                SharingAllowedDomainList                 = $classicTeamSite.SharingAllowedDomainList              
                SharingBlockedDomainList                 = $classicTeamSite.SharingBlockedDomainList    
                SharingCapability                        = $classicTeamSite.SharingCapability                        
                SharingDomainRestrictionMode             = $classicTeamSite.SharingDomainRestrictionMode             
                ShowPeoplePickerSuggestionsForGuestUsers = $classicTeamSite.ShowPeoplePickerSuggestionsForGuestUsers                     
                SocialBarOnSitePagesDisabled             = $classicTeamSite.SocialBarOnSitePagesDisabled                                                   
                StorageQuotaWarningLevel                 = $classicTeamSite.StorageQuotaWarningLevel              
                ResourceQuotaWarningLevel                = $classicTeamSite.ResourceQuotaWarningLevel                
            }   

            #if($removeSiteCollectionAdmin) {Remove-SiteCollecionAdmin -srcAdministrativeUsername $global:btSourceO365Creds.Username -srcSpoSiteUrl $classicTeamSite.Url -details $details}
        
        } 

    }

    if ($details) { Write-host }
    if ($script:rootDocumentLibrariesCount -ne 0 -or $classicTeamRootSubsiteCount -ne 0 -or $classicTeamSiteCount -ne 0) {
        #Export Classic Team Sites Array to CSV file
        do {
            try {
                $classicTeamSitesArray | Export-Csv -Path $workingDir\SPOclassicTeamSites-$script:sourceTenantName.csv -NoTypeInformation -force -Encoding UTF8 -ErrorAction Stop

                Break
            }
            catch {
                $msg = "WARNING: Close opened CSV file '$workingDir\SPOclassicTeamSites-$script:sourceTenantName.csv'."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg
                Write-Host

                Start-Sleep 5
            }
        } while ($true)

        do {
            try {
                $documentLibrariesArray | Export-Csv -Path $workingDir\SPODocumentLibraries-$script:sourceTenantName.csv -NoTypeInformation -force -Encoding UTF8 -ErrorAction Stop 

                Break
            }
            catch {
                $msg = "WARNING: Close opened CSV file '$workingDir\SPODocumentLibraries-$script:sourceTenantName.csv'."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg
                Write-Host

                Start-Sleep 5
            }
        } while ($true)
    
        $msg = "SUCCESS: CSV file '$workingDir\SPOclassicTeamSites-$script:sourceTenantName.csv' processed, exported and open."
        if ($details) { Write-Host -ForegroundColor Green $msg }
        Log-Write -Message $msg
    
        try {
            if ($openCSVFile) { Start-Process -FilePath $workingDir\SPOclassicTeamSites-$script:sourceTenantName.csv }
        }
        catch {
            $msg = "ERROR: Failed to find the CSV file '$workingDir\SPOclassicTeamSites-$script:sourceTenantName.csv'."    
            if ($details) { Write-Host -ForegroundColor Red $msg }
            return
        }  

        $msg = "SUCCESS: CSV file '$workingDir\SPODocumentLibraries-$script:sourceTenantName.csv' processed, exported and open."
        if ($details) { Write-Host -ForegroundColor Green $msg }
        Log-Write -Message $msg
    
        try {
            if ($openCSVFile) { Start-Process -FilePath $workingDir\SPODocumentLibraries-$script:sourceTenantName.csv }
        }
        catch {
            $msg = "ERROR: Failed to find the CSV file '$workingDir\SPODocumentLibraries-$script:sourceTenantName.csv'."    
            if (!$detail) { Write-Host -ForegroundColor Red $msg }
            return
        }  
    }
}

# Function to export O365 Groups
Function Export-O365UnifiedGroups {  

    $details = $false
    $openCSVFile = !$script:migrateEntireTenant 

    $groups = @(Get-UnifiedGroup -ResultSize Unlimited )
    $groups | ForEach-Object {
        $currentGroup = $_

        $serviceEndPoints = @()
        if ($currentGroup.ServiceEndpointUris -ne $null) {
            $serviceEndPoints = $currentGroup.ServiceEndpointUris
        }

        $hasTeams = ($serviceEndPoints | Where-Object { $_.IndexOf("MicrosoftTeams.TeamHomeURL") -gt -1 }) -ne $null
        $hasMail = [String]::IsNullOrEmpty($currentGroup.InboxUrl) -eq $false
        $hasSharePoint = [String]::IsNullOrEmpty($currentGroup.SharePointSiteUrl) -eq $false

        $currentGroup | Add-Member -NotePropertyName HasTeams -NotePropertyValue $hasTeams
        $currentGroup | Add-Member -NotePropertyName HasMail -NotePropertyValue $hasMail
        $currentGroup | Add-Member -NotePropertyName HasSharePoint -NotePropertyValue $hasSharePoint
    }
    
    $unifiedGroups = $groups | Where-Object { $_.HasTeams -eq $false }

    $unifiedGroupArray = @()   
    
    if ($script:srcGermanyCloud) {
        $sSPOAdminCenterUrl = "https://$script:sourceTenantName-admin.sharepoint.de/"
        $sSPOUrl = "https://$script:sourceTenantName.sharepoint.de/"
    }
    elseif ($script:srcUsGovernment) {
        $sSPOAdminCenterUrl = "https://$script:sourceTenantName-admin.sharepoint.us/"
        $sSPOUrl = "https://$script:sourceTenantName.sharepoint.us/"
    }
    else {
        $sSPOAdminCenterUrl = "https://$script:sourceTenantName-admin.sharepoint.com/"
        $sSPOUrl = "https://$script:sourceTenantName.sharepoint.com/"
    }

    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
    try {
        Connect-SPOService -Url $sSPOAdminCenterUrl -Credential $global:btSourceO365Creds -ErrorAction Stop
    }
    catch {
        $msg = "ERROR: Failed to connect to SPOService because modern authentication is required."    
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 

        $msg = "ACTION: Re-enter the credentials for the SPOService."    
        Write-Host -ForegroundColor Yellow  $msg
        Log-Write -Message $msg 

        try {
            Connect-SPOService -Url $sSPOAdminCenterUrl -ErrorAction Stop
        }
        catch {
            $msg = "ERROR: Failed to connect to SPOService."    
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message 
            return
        }
    }
    try {
        Connect-PnPOnline -Url $sSPOAdminCenterUrl -Credentials $global:btSourceO365Creds -ErrorAction Stop
    }
    catch {
        $msg = "ERROR: Failed to connect to PnPOnline because modern authentication is required."    
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 

        $msg = "ACTION: Enter ClientSecret and ClientId for the PnPOnline with modern authentication."    
        Write-Host -ForegroundColor Yellow  $msg
        Log-Write -Message $msg 

        do {
            $script:ClientSecret = (Read-Host -prompt "Please enter the ClientSecret").trim()
        }while ($script:ClientSecret -eq "")

        $msg = "INFO: ClientSecret is '$script:ClientSecret'."
        Write-Host $msg
        Log-Write -Message $msg 

        do {
            $script:ClientId = (Read-Host -prompt "Please enter the ClientId").trim()
        }while ($script:ClientId -eq "")

        $msg = "INFO: ClientId is '$script:ClientId'."
        Write-Host $msg
        Log-Write -Message $msg 

        $useModernAuthentication = $true

        try {
            Connect-PnPOnline -Url $sSPOAdminCenterUrl -ClientSecret $script:ClientSecret -ClientId $script:ClientId -ErrorAction Stop
        }
        catch {
            $msg = "ERROR: Failed to connect to PnPOnline."    
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message 
            return
        }
    }
    
    $unifiedGroupSPOSiteArray = @()
    $documentLibrariesArray = @()

    #Exclude Document libraries
    $systemLists = @("Maintenance Log Library", "appdata", "TaxonomyHiddenList", "User Information List", "Composed Looks", "MicroFeed", "appfiles", "Converted Forms", "Customized Reports", "Form Templates", "Images", "List Template Gallery", "Master Page Gallery", "Pages", "Reporting Templates", "Site Assets", "Site Collection Documents", "Site Collection Images", "Site Pages", "Solution Gallery", "Style Library", "Theme Gallery", "Web Part Gallery", "wfpub", "_catalogs/hubsite")
 
    $o365UnifiedGroupSubsiteCount = 0

    $unlinkServiceAccount = $false

    $sessionStartTime = Get-Date
    $sessionEndTime = (Get-Date).AddHours(+5)

    Foreach ($UnifiedGroup in $UnifiedGroups) {

        if ($sessionStartTime -gt $SessionEndTime) {
            Write-Host 
            $msg = "INFO: Start another Exchange Online session."
            Write-Host -ForegroundColor Magenta  $msg
            Log-Write -Message $msg
            Write-Host

            $script:sourceO365Session = Connect-SourceExchangeOnline

            if ($script:sourceO365Session.toString() -ne "-1") {
                $script:destinationO365Session = Connect-DestinationExchangeOnline
                if ($script:destinationO365Session.toString() -ne "-1") { Break }            
            }

            $sessionStartTime = Get-Date
            $sessionEndTime = (Get-Date).AddHours(+5)
        }

        $currentUnifiedGroup += 1 
        $o365UnifiedGroupSubsiteWithDomainCount = 0

        if ($UnifiedGroup.TeamEnabled -eq $true) {
            $msg = "      INFO: Exporting Team-enabled Office 365 (unified) group $currentUnifiedGroup/$teamEnabledUnifiedGroupsCount : '$($UnifiedGroup.DisplayName)' URL: $($UnifiedGroup.SharePointSiteUrl) EmailAddress: $($UnifiedGroup.PrimarySmtpAddress)."
        }
        else {
            $msg = "      INFO: Exporting Office 365 (unified) group $currentUnifiedGroup/$nonTeamEnabledUnifiedGroupsCount : '$($UnifiedGroup.DisplayName)' URL: $($UnifiedGroup.SharePointSiteUrl) EmailAddress: $($UnifiedGroup.PrimarySmtpAddress)."
        }
       
        if ($details) { Write-host $msg } 
        Log-Write -Message $msg

        Link-O365Groups -groupName $UnifiedGroup.DisplayName -srcPrimarySMTPAddress $unifiedGroup.PrimarySmtpAddress -srcAdministrativeUsername $global:btSourceO365Creds.Username

        $unlinkServiceAccount = $true

        If ($UnifiedGroup.SharePointSiteUrl -ne $null) { 
           
            try {
                $UnifiedGroupSPOSite = Get-SPOSite -Identity $UnifiedGroup.SharePointSiteUrl -ErrorAction Stop
            }
            catch {
                $msg = "      ERROR: Failed to get SPOSite '$($UnifiedGroup.SharePointSiteUrl)'. Access denied."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg 
            }

            try {
                $result = Set-SPOUser -Site $UnifiedGroup.SharePointSiteUrl -LoginName $global:btSourceO365Creds.UserName -IsSiteCollectionAdmin $true
       
                $msg = "      SUCCESS: '$($global:btSourceO365Creds.UserName)' added to get SPO Site Group '$($UnifiedGroup.SharePointSiteUrl)' as SiteCollectionAdmin."
                if ($details) { Write-Host -ForegroundColor Green  $msg }
                Log-Write -Message $msg 

                $removeSiteCollectionAdmin = $true
            }
            catch {
                $msg = "      ERROR: Failed to set '$($global:btSourceO365Creds.UserName)' as SiteCollectionAdmin of SPO Site Group '$($UnifiedGroup.SharePointSiteUrl)'. Access denied."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg 

                $removeSiteCollectionAdmin = $false
            }
           
            $documentLibraryNames = @()
            try {
                if ($useModernAuthentication) {
                    Connect-PnPOnline -Url $UnifiedGroup.SharePointSiteUrl -ClientSecret $script:ClientSecret  -ClientId $script:ClientId -ErrorAction Stop
                }
                else {
                    Connect-PnPOnline -Url $UnifiedGroup.SharePointSiteUrl -Credentials $global:btSourceO365Creds -ErrorAction Stop 
                }                         
            }
            catch {
                $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($UnifiedGroup.SharePointSiteUrl) site with Connect-PnPOnline."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                Log-Write -Message $msg
                Log-Write -Message "         $($_.Exception.Message)"

                $documentLibraryNames += "Shared Documents"

                $documentLibrariesArray += [PSCustomObject] @{ 
                    Title                     = "Documents"
                    DocumentLibraryName       = "Shared Documents"                  
                    EntityTypeName            = "Shared_x0020_Documents" 
                    BaseType                  = "DocumentLibrary"
                    TeamSiteUrl               = $UnifiedGroup.SharePointSiteUrl
                    NewDestinationTeamSiteUrl = $UnifiedGroup.SharePointSiteUrl
                    isSubsite                 = "FALSE"
                    SubsiteDepth              = "0"
                    SubsiteNumber             = 0
                    DocumentLibraryUrl        = $($UnifiedGroup.SharePointSiteUrl) + "/Shared Documents"       
                }
                Write-Host
                Continue
            } 
            
            try {
                $documentLibraries = @(Get-PnPList -ErrorAction Stop | Where-Object { $_.Title -notin $systemLists } ) 
            }
            catch {
                $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($UnifiedGroup.SharePointSiteUrl) site with Get-PnPList."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                Log-Write -Message $msg
                Log-Write -Message "         $($_.Exception.Message)"
            }

            $siteUrl = $UnifiedGroup.SharePointSiteUrl 

            $urlparts = $siteUrl.split("/")        
            $sitesOrTeams = $urlparts[3]
            $rootUrl = ($siteUrl -split $sitesOrTeams)[0]  
            $siteName = ($siteUrl -split $sitesOrTeams)[1]    
            $rebuiltSiteUrl = "$rootUrl$sitesOrTeams$siteName"
          
            foreach ($documentLibrary in $documentLibraries) {

                if ($documentLibrary.DocumentTemplateUrl -eq $null) { Continue }
                if ($documentLibrary.EntityTypeName -match "SiteAssets") { Continue }
                if ($documentLibrary.EntityTypeName -match "Translation_x0020_Packages") { Continue }
                if ($documentLibrary.EntityTypeName -match "List") { Continue }
                if ($documentLibrary.DocumentTemplateUrl -notmatch "/Forms/template.dot") { Continue }

                $documentLibraryName = $documentLibrary.DocumentTemplateUrl.Replace("/$sitesOrTeams$siteName/", "").split("/")[0] 
                $documentLibraryUrl = "$rebuiltSiteUrl/$documentLibraryName"  

                $msg = "            SUCCESS: Document Library '$documentLibraryName' found under '$documentLibraryUrl'."
                if ($details) { write-Host -ForegroundColor Green  $msg }
                Log-Write -Message $msg

                $documentLibraryNames += $documentLibraryName
         
                $documentLibrariesArray += [PSCustomObject] @{ 
                    Title                     = $documentLibrary.Title  
                    DocumentLibraryName       = $documentLibraryName                     
                    EntityTypeName            = $documentLibrary.EntityTypeName 
                    BaseType                  = $documentLibrary.BaseType
                    TeamSiteUrl               = $UnifiedGroup.SharePointSiteUrl
                    NewDestinationTeamSiteUrl = $UnifiedGroup.SharePointSiteUrl
                    isSubsite                 = "FALSE"
                    SubsiteDepth              = "0"
                    SubsiteNumber             = "0"
                    DocumentLibraryUrl        = $documentLibraryUrl          
                }
            }

            try {
                $SubWebs = @()
                $SubWebs = (Get-PnPSubWebs -Recurse -ErrorAction Stop)
                $SubWebsCount = $SubWebs.Count
                $currentSubWeb = 0
            }
            catch {
                $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($UnifiedGroup.SharePointSiteUrl) SubWebs with Get-PnPSubWebs."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                Log-Write -Message $msg
                Log-Write -Message "         $($_.Exception.Message)"
            }

            if ($SubWebs -ne $null) {
                Foreach ($SubWeb in $SubWebs) { 
                    $currentSubWeb += 1

                    $msg = "         INFO: Exporting '$($UnifiedGroup.DisplayName)' Subsites and their Document Libraries $currentSubWeb/$SubWebsCount :  '$($SubWeb.Title)' '$($SubWeb.Url)'."
                    if ($details) { Write-host $msg } 
                    Log-Write -Message $msg 

                    try {
                        if ($useModernAuthentication) {
                            Connect-PnPOnline -Url $SubWeb.Url -ClientSecret $script:ClientSecret  -ClientId $script:ClientId -ErrorAction Stop
                        }
                        else {
                            Connect-PnPOnline -Url $SubWeb.Url -Credentials $global:btSourceO365Creds -ErrorAction Stop
                        }
                    }
                    catch {
                        $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($SubWeb.Url) subsite with Connect-PnPOnline."
                        Write-Host -ForegroundColor Red  $msg
                        Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                        Log-Write -Message $msg
                        Log-Write -Message "         $($_.Exception.Message)"

                        $documentLibraryNames += "Subsite:Shared Documents"

                        $documentLibrariesArray += [PSCustomObject] @{ 
                            Title                     = "Documents"
                            DocumentLibraryName       = "Shared Documents"                  
                            EntityTypeName            = "Shared_x0020_Documents" 
                            BaseType                  = "DocumentLibrary"
                            TeamSiteUrl               = $SubWeb.Url
                            NewDestinationTeamSiteUrl = $SubWeb.Url
                            isSubsite                 = "TRUE"
                            SubsiteDepth              = $subsitesDepth
                            SubsiteNumber             = "$currentSubWeb of $SubWebsCount"
                            DocumentLibraryUrl        = $SubWeb.Url + "/Shared Documents"       
                        }

                        $o365UnifiedGroupSubsiteCount += 1
                        $o365UnifiedGroupSubsiteWithDomainCount += 1
                    }  

                    try {
                        $subWebDocumentLibraries = @(Get-PnPList -ErrorAction Stop | Where-Object { $_.Title -notin $systemLists } ) 
                    }
                    catch {
                        $msg = "      ERROR: Service account '$($global:btSourceO365Creds.UserName)' unauthorized to access $($SubWeb.Url) site with Get-PnPList."
                        Write-Host -ForegroundColor Red  $msg
                        Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                        Log-Write -Message $msg
                        Log-Write -Message "         $($_.Exception.Message)"
                    }
                    try {
                        $subSiteUrl = $SubWeb.Url

                        $urlparts = $subSiteUrl.split("/")      
                        $sitesOrTeams = $urlparts[3]
                        $rootUrl = ($subSiteUrl -split $sitesOrTeams)[0]  
                        $siteNames = ($subSiteUrl -split $sitesOrTeams)[1] 
                        $rebuiltSubSiteUrl = "$rootUrl$sitesOrTeams$siteNames"

                        if ($subSiteUrl -ne $rebuiltSubSiteUrl) {
                            $msg = "      ERROR: Unable to rebuild $($SubWeb.Url) subsite URL."
                            Write-Host -ForegroundColor Red  $msg
                            Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                            Log-Write -Message $msg
                            Log-Write -Message "         $($_.Exception.Message)"  
                            Continue  
                        }

                        $subsitesArray = @($siteNames.split('/'))
                        $subsitesDepth = $subsitesArray.Count - 2

                        foreach ($subWebDocumentLibrary in $subWebDocumentLibraries) {

                            if ($subWebDocumentLibrary.DocumentTemplateUrl -eq $null) { Continue }
                            if ($subWebDocumentLibrary.EntityTypeName -match "SiteAssets") { Continue }
                            if ($subWebDocumentLibrary.EntityTypeName -match "Translation_x0020_Packages") { Continue }
                            if ($subWebDocumentLibrary.EntityTypeName -match "List") { Continue }
                            if ($subWebDocumentLibrary.DocumentTemplateUrl -notmatch "/Forms/template.dot") { Continue }

                            $documentLibraryName = $subWebDocumentLibrary.DocumentTemplateUrl.Replace("/$sitesOrTeams$siteNames/", "").split("/")[0] 
                            $documentLibraryUrl = "$rebuiltSubSiteUrl/$documentLibraryName"  


                            $msg = "            SUCCESS: Document Library '$($subWebDocumentLibrary.Title)' found under '$documentLibraryUrl'."
                            if ($details) { Write-Host -ForegroundColor Green  $msg }
                            Log-Write -Message $msg

                            $documentLibraryNames += "Subsite:$documentLibraryName"
        
                            $documentLibrariesArray += [PSCustomObject] @{ 
                                Title                     = $subWebDocumentLibrary.Title  
                                DocumentLibraryName       = $documentLibraryName                     
                                EntityTypeName            = $subWebDocumentLibrary.EntityTypeName 
                                BaseType                  = $subWebDocumentLibrary.BaseType
                                TeamSiteUrl               = $subSiteUrl
                                NewDestinationTeamSiteUrl = $subSiteUrl
                                isSubsite                 = "TRUE"
                                SubsiteDepth              = $subsitesDepth
                                SubsiteNumber             = "$currentSubWeb of $SubWebsCount"
                                DocumentLibraryUrl        = $documentLibraryUrl          
                            }

                            $o365UnifiedGroupSubsiteCount += 1
                            $o365UnifiedGroupSubsiteWithDomainCount += 1
                        }                      
                    }
                    catch {
                        $msg = "         ERROR: Unable to get PnPList of $($SubWeb.Url) subsite ."
                        Write-Host -ForegroundColor Red  $msg
                        Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
                        Log-Write -Message $msg
                        Log-Write -Message "         $($_.Exception.Message)"

                        $documentLibraryNames += "Subsite:Shared Documents"

                        $documentLibrariesArray += [PSCustomObject] @{ 
                            Title                     = "Documents"
                            DocumentLibraryName       = "Shared Documents"                  
                            EntityTypeName            = "Shared_x0020_Documents" 
                            BaseType                  = "DocumentLibrary"
                            TeamSiteUrl               = $SubWeb.Url
                            NewDestinationTeamSiteUrl = $SubWeb.Url
                            isSubsite                 = "TRUE"
                            SubsiteDepth              = $subsitesDepth
                            SubsiteNumber             = "$currentSubWeb of $SubWebsCount"
                            DocumentLibraryUrl        = $SubWeb.Url + "/Shared Documents"       
                        }
                    }      
                }
            }
           

            $documentLibraryNames = $documentLibraryNames -join ";" 
            $msg = "         INFO: All Document Library names found: '$documentLibraryNames'."
            if ($details) { Write-Host $msg }
            Log-Write -Message $msg                      

            $unifiedGroupSPOSiteArray += [PSCustomObject] @{ 
                CompatibilityLevel                       = $UnifiedGroupSPOSite.CompatibilityLevel                       
                LocaleId                                 = $UnifiedGroupSPOSite.LocaleId 
                Owner                                    = $UnifiedGroupSPOSite.Owner    
                ResourceQuota                            = $UnifiedGroupSPOSite.ResourceQuota 
                StorageQuota                             = $UnifiedGroupSPOSite.StorageQuota
                Template                                 = $UnifiedGroupSPOSite.Template  
                TimeZoneId                               = $UnifiedGroupSPOSite.TimeZoneId                                                                          
                Title                                    = $UnifiedGroupSPOSite.Title 
                Url                                      = $UnifiedGroupSPOSite.Url
                NewDestinationUrl                        = $UnifiedGroupSPOSite.Url
                DocumentLibraryName                      = $documentLibraryNames

                AllowEditing                             = $UnifiedGroupSPOSite.AllowEditing       
                AllowSelfServiceUpgrade                  = $UnifiedGroupSPOSite.AllowSelfServiceUpgrade  
                BlockDownloadOfNonViewableFiles          = $UnifiedGroupSPOSite.AllowDownloadingNonWebViewableFiles  
                CommentsOnSitePagesDisabled              = $UnifiedGroupSPOSite.CommentsOnSitePagesDisabled  
                ConditionalAccessPolicy                  = $UnifiedGroupSPOSite.ConditionalAccessPolicy   
                DefaultLinkPermission                    = $UnifiedGroupSPOSite.DefaultLinkPermission             
                DefaultSharingLinkType                   = $UnifiedGroupSPOSite.DefaultSharingLinkType    
                DenyAddAndCustomizePages                 = $UnifiedGroupSPOSite.DenyAddAndCustomizePages   
                DisableAppViews                          = $UnifiedGroupSPOSite.DisableAppViews 
                DisableCompanyWideSharingLinks           = $UnifiedGroupSPOSite.DisableCompanyWideSharingLinks   
                DisableFlows                             = $UnifiedGroupSPOSite.DisableFlows   
                LimitedAccessFileType                    = $UnifiedGroupSPOSite.LimitedAccessFileType 
                LockState                                = $UnifiedGroupSPOSite.LockState                                
                RestrictedToRegion                       = $UnifiedGroupSPOSite.RestrictedToRegion 
                SandboxedCodeActivationCapability        = $UnifiedGroupSPOSite.SandboxedCodeActivationCapability        
                SharingAllowedDomainList                 = $UnifiedGroupSPOSite.SharingAllowedDomainList              
                SharingBlockedDomainList                 = $UnifiedGroupSPOSite.SharingBlockedDomainList    
                SharingCapability                        = $UnifiedGroupSPOSite.SharingCapability                        
                SharingDomainRestrictionMode             = $UnifiedGroupSPOSite.SharingDomainRestrictionMode            
                ShowPeoplePickerSuggestionsForGuestUsers = $UnifiedGroupSPOSite.ShowPeoplePickerSuggestionsForGuestUsers                     
                SocialBarOnSitePagesDisabled             = $UnifiedGroupSPOSite.SocialBarOnSitePagesDisabled                                                   
                StorageQuotaWarningLevel                 = $UnifiedGroupSPOSite.StorageQuotaWarningLevel              
                ResourceQuotaWarningLevel                = $UnifiedGroupSPOSite.ResourceQuotaWarningLevel                
            }
        }

        $Mem = Get-UnifiedGroupLinks -Identity $UnifiedGroup.Identity -LinkType Members -ResultSize Unlimited
        $Own = Get-UnifiedGroupLinks -Identity $UnifiedGroup.Identity -LinkType Owners -ResultSize Unlimited
        $Sub = Get-UnifiedGroupLinks -Identity $UnifiedGroup.Identity -LinkType Subscribers -ResultSize Unlimited
        $Agg = Get-UnifiedGroupLinks -Identity $UnifiedGroup.Identity -LinkType Aggregators -ResultSize Unlimited

        if ($script:exportSingleDomain) {
            if (($Mem.PrimarySmtpAddress -join ";") -notmatch $script:exportDomain) { Continue }
        }

        if ($script:exportSingleDomain) {
            $script:numberUnifiedGroupsSubsites += $o365UnifiedGroupSubsiteWithDomainCount
        }
        	
        # Resolve Members
        $Members = @()
        foreach ($address in $Mem) { if ($address.PrimarySmtpAddress -ne "") { $Members += $address.PrimarySmtpAddress } }

        # Resolve Owners
        $Owners = @()
        foreach ($address in $Own) { if ($address.PrimarySmtpAddress -ne "") { $Owners += $address.PrimarySmtpAddress } }
        
        # Resolve Subscribers
        $Subscribers = @()
        foreach ($address in $Sub) { if ($address.PrimarySmtpAddress -ne "") { $Subscribers += $address.PrimarySmtpAddress } }

        # Resolve Aggregators
        $Aggregators = @()
        foreach ($address in $Agg) { if ($address.PrimarySmtpAddress -ne "" -and $address.PrimarySmtpAddress -notMatch "AggregateGroupMailbox.A.") { $Aggregators += $address.PrimarySmtpAddress } }
	  
                	
        # Resolve GrantSendOnBehalfTo
        $GrantSendOnBehalfTo = @()
        foreach ($address in $UnifiedGroup.GrantSendOnBehalfTo) { 
            $recipient = (Get-Recipient $address -ErrorAction SilentlyContinue).PrimarySmtpAddress
            if ($recipient) { $GrantSendOnBehalfTo += $recipient } 
            else {
                $msg = "      ERROR: '$address' not found in source Office 365 tenant."
                if ($details) { Write-Host -ForegroundColor Red  $msg }
                Log-Write -Message $msg
            } 
        }
	
        #Resolve ManagedBy
        $ManagedBy = @()
        foreach ($address in $UnifiedGroup.ManagedBy) { 
            $recipient = (Get-Recipient $address -ErrorAction SilentlyContinue).PrimarySmtpAddress
            if ($recipient) { $ManagedBy += $recipient } 
            else {
                $msg = "      ERROR: '$address' not found in source Office 365 tenant."
                if ($details) { Write-Host -ForegroundColor Red  $msg }
                Log-Write -Message $msg
            } 
        }

        # Resolve ManagedByDetails
        $ManagedByDetails = @()
        foreach ($address in $UnifiedGroup.ManagedByDetails) { 
            $recipient = (Get-Recipient $address -ErrorAction SilentlyContinue).PrimarySmtpAddress
            if ($recipient) { $ManagedByDetails += $recipient } 
            else {
                $msg = "      ERROR: '$address' not found in source Office 365 tenant."
                if ($details) { Write-Host -ForegroundColor Red  $msg }
                Log-Write -Message $msg
            } 
        }

        # Resolve ModeratedBy
        $ModeratedBy = @()
        foreach ($address in $UnifiedGroup.ModeratedBy) { 
            $recipient = (Get-Recipient $address -ErrorAction SilentlyContinue).PrimarySmtpAddress
            if ($recipient) { $ModeratedBy += $recipient } 
            else {
                $msg = "      ERROR: '$address' not found in source Office 365 tenant."
                if ($details) { Write-Host -ForegroundColor Red  $msg }
                Log-Write -Message $msg
            } 
        }
	
        # Resolve RejectMessagesFrom
        $RejectMessagesFrom = @()
        foreach ($address in $UnifiedGroup.RejectMessagesFrom) { 
            $recipient = (Get-Recipient $address -ErrorAction SilentlyContinue).PrimarySmtpAddress
            if ($recipient) { $RejectMessagesFrom += $recipient } 
            else {
                $msg = "      ERROR: '$address' not found in source Office 365 tenant."
                if ($details) { Write-Host -ForegroundColor Red  $msg }
                Log-Write -Message $msg
            } 
        }
	
        # Resolve RejectMessagesFromSendersOrMembers
        $RejectMessagesFromSendersOrMembers = @()
        foreach ($address in $UnifiedGroup.RejectMessagesFromSendersOrMembers) { 
            $recipient = (Get-Recipient $address -ErrorAction SilentlyContinue).PrimarySmtpAddress
            if ($recipient) { $RejectMessagesFromSendersOrMembers += $recipient } 
            else {
                $msg = "      ERROR: '$address' not found in source Office 365 tenant."
                if ($details) { Write-Host -ForegroundColor Red  $msg }
                Log-Write -Message $msg
            } 
        }
	
        # Resolve RejectMessagesFromDLMembers
        $RejectMessagesFromDLMembers = @()
        foreach ($address in $UnifiedGroup.RejectMessagesFromDLMembers) { 
            $recipient = (Get-Recipient $address -ErrorAction SilentlyContinue).PrimarySmtpAddress
            if ($recipient) { $RejectMessagesFromDLMembers += $recipient } 
            else {
                $msg = "      ERROR: '$address' not found in source Office 365 tenant."
                if ($details) { Write-Host -ForegroundColor Red  $msg }
                Log-Write -Message $msg
            } 
        }
	
        # Resolve AcceptMessagesOnlyFrom
        $AcceptMessagesOnlyFrom = @()
        foreach ($address in $UnifiedGroup.AcceptMessagesOnlyFrom) { 
            $recipient = (Get-Recipient $address -ErrorAction SilentlyContinue).PrimarySmtpAddress
            if ($recipient) { $AcceptMessagesOnlyFrom += $recipient } 
            else {
                $msg = "      ERROR: '$address' not found in source Office 365 tenant."
                if ($details) { Write-Host -ForegroundColor Red  $msg }
                Log-Write -Message $msg
            } 
        }
	
        # Resolve AcceptMessagesOnlyFromDLMembers
        $AcceptMessagesOnlyFromDLMembers = @()
        foreach ($address in $UnifiedGroup.AcceptMessagesOnlyFromDLMembers) { 
            $recipient = (Get-Recipient $address -ErrorAction SilentlyContinue).PrimarySmtpAddress
            if ($recipient) { $AcceptMessagesOnlyFromDLMembers += $recipient } 
            else {
                $msg = "      ERROR: '$address' not found in source Office 365 tenant."
                if ($details) { Write-Host -ForegroundColor Red  $msg }
                Log-Write -Message $msg
            } 
        }
	
        # Resolve AcceptMessagesOnlyFromSendersOrMembers
        $AcceptMessagesOnlyFromSendersOrMembers = @()
        foreach ($address in $UnifiedGroup.AcceptMessagesOnlyFromSendersOrMembers) { 
            $recipient = (Get-Recipient $address -ErrorAction SilentlyContinue).PrimarySmtpAddress
            if ($recipient) { $AcceptMessagesOnlyFromSendersOrMembers += $recipient } 
            else {
                $msg = "      ERROR: '$address' not found in source Office 365 tenant."
                if ($details) { Write-Host -ForegroundColor Red  $msg }
                Log-Write -Message $msg
            } 
        }
	    
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName Name -NotePropertyValue $UnifiedGroup.Name -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName DisplayName -NotePropertyValue $UnifiedGroup.DisplayName -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName AccessType -NotePropertyValue $UnifiedGroup.AccessType -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName PrimarySmtpAddress -NotePropertyValue $UnifiedGroup.PrimarySmtpAddress -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName EmailAddresses -NotePropertyValue ($UnifiedGroup.EmailAddresses -join "|") -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName SharePointSiteUrl -NotePropertyValue $UnifiedGroup.SharePointSiteUrl -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName NewDestinationSharePointSiteUrl -NotePropertyValue $UnifiedGroup.SharePointSiteUrl -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName SharePointDocumentsUrl -NotePropertyValue $UnifiedGroup.SharePointDocumentsUrl -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName SharePointNotebookUrl -NotePropertyValue $UnifiedGroup.SharePointNotebookUrl -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName Members -NotePropertyValue ($Members -join "|") -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName Owners -NotePropertyValue ($Owners -join "|") -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName Subscribers -NotePropertyValue ($Subscribers -join "|") -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName Aggregators -NotePropertyValue ($Aggregators -join "|") -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName ManagedBy -NotePropertyValue ($ManagedBy -join "|") -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName ManagedByDetails -NotePropertyValue ($ManagedByDetails -join "|") -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName ModeratedBy -NotePropertyValue ($ModeratedBy -join "|") -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName AcceptMessagesOnlyFrom -NotePropertyValue ($AcceptMessagesOnlyFrom -join ",") -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName AcceptMessagesOnlyFromDLMembers -NotePropertyValue ($AcceptMessagesOnlyFromDLMembers -join "|") -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName AcceptMessagesOnlyFromSendersOrMembers -NotePropertyValue ($AcceptMessagesOnlyFromSendersOrMembers -join "|") -Force    
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName RejectMessagesFrom -NotePropertyValue ($RejectMessagesFrom -join "|") -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName RejectMessagesFromDLMembers -NotePropertyValue ($RejectMessagesFromDLMembers -join "|") -Force
        $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName RejectMessagesFromSendersOrMembers -NotePropertyValue ($RejectMessagesFromSendersOrMembers -join "|") -Force


        if ($GrantSendOnBehalfTo -ne "") { $UnifiedGroup | Add-Member -TypeName NoteProperty -NotePropertyName GrantSendOnBehalfTo -NotePropertyValue ($GrantSendOnBehalfTo -join "|") -Force }
        
        $unifiedGroupArray += $UnifiedGroup  

        if ($UnifiedGroup.TeamEnabled -eq $true) {
            $teamEnabled = $true
        }

        $UnifiedGroup = $null

        if ($unlinkServiceAccount) { UnLink-O365Groups -srcPrimarySMTPAddress $unifiedGroup.PrimarySmtpAddress -srcAdministrativeUsername $global:btSourceO365Creds.Username }

    } 
       
    #Export Distribution Groups to CSV file
    do {
        try {
            if ($teamEnabled) {
                $unifiedGroupArray | Export-Csv -Path $workingDir\O365TeamEnabledUnifiedGroups-$script:sourceTenantName.csv -NoTypeInformation -force -Encoding UTF8 -ErrorAction Stop 
            }
            else {
                $unifiedGroupArray | Export-Csv -Path $workingDir\O365UnifiedGroups-$script:sourceTenantName.csv -NoTypeInformation -force -Encoding UTF8  -ErrorAction Stop
                 
                if ([string]::IsNullOrEmpty($unifiedGroupArray.Count)) { $script:numberUnifiedGroups = 0 } else { $script:numberUnifiedGroups = $unifiedGroupArray.Count }
            }

            Break
        }
        catch {
            if ($teamEnabled) {
                $msg = "WARNING: Close opened CSV file '$workingDir\O365TeamEnabledUnifiedGroups-$script:sourceTenantName.csv'."
            }
            else {
                $msg = "WARNING: Close opened CSV file '$workingDir\O365UnifiedGroups-$script:sourceTenantName.csv'."
            }

            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            Write-Host

            Start-Sleep 5
        }
    } while ($true)

    do {
        try {
            if ($teamEnabled) {
                $unifiedGroupSPOSiteArray | Export-Csv -Path $workingDir\O365TeamEnabledUnifiedGroupSPOSites-$script:sourceTenantName.csv -NoTypeInformation -force -Encoding UTF8 -ErrorAction Stop 
            }
            else {
                $unifiedGroupSPOSiteArray | Export-Csv -Path $workingDir\O365UnifiedGroupSPOSites-$script:sourceTenantName.csv -NoTypeInformation -force -Encoding UTF8  -ErrorAction Stop
            }

            Break
        }
        catch {
            if ($teamEnabled) {
                $msg = "WARNING: Close opened CSV file '$workingDir\O365TeamEnabledUnifiedGroupSPOSites-$script:sourceTenantName.csv'."
            }
            else {
                $msg = "WARNING: Close opened CSV file '$workingDir\O365UnifiedGroupSPOSites-$script:sourceTenantName.csv'."
            }

            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            Write-Host

            Start-Sleep 5
        }
    } while ($true)

    #Export Distribution Groups to CSV file
    do {
        try {
            if ($teamEnabled) {
                $documentLibrariesArray | Export-Csv -Path $workingDir\O365TeamEnabledUnifiedGroupDocumentLibraries-$script:sourceTenantName.csv -NoTypeInformation -force -Encoding UTF8 -ErrorAction Stop
            }
            else {
                $documentLibrariesArray | Export-Csv -Path $workingDir\O365UnifiedGroupDocumentLibraries-$script:sourceTenantName.csv -NoTypeInformation -force -Encoding UTF8 -ErrorAction Stop
            }

            Break
        }
        catch {
            if ($teamEnabled) {
                $msg = "WARNING: Close opened CSV file '$workingDir\O365TeamEnabledUnifiedGroupDocumentLibraries-$script:sourceTenantName.csv'."
            }
            else {
                $msg = "WARNING: Close opened CSV file '$workingDir\O365UnifiedGroupDocumentLibraries-$script:sourceTenantName.csv'."
            }
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            Write-Host

            Start-Sleep 5
        }
    } while ($true)
 
    Disconnect-SPOService 
}

Function Link-O365Groups {
    param 
    (             
        [parameter(Mandatory = $true)] [String]$groupName,
        [parameter(Mandatory = $false)] [String]$srcPrimarySMTPAddress,
        [parameter(Mandatory = $false)] [String]$srcAdministrativeUsername,
        [parameter(Mandatory = $false)] [String]$dstPrimarySMTPAddress,
        [parameter(Mandatory = $false)] [String]$dstAdministrativeUsername,
        [parameter(Mandatory = $false)] [Boolean]$details

    )
    
    if (-not ([string]::IsNullOrEmpty($srcPrimarySMTPAddress))) {  
        $msg = "         INFO: Adding admin '$srcAdministrativeUsername' as a member and owner of the source Office 365 Group '$srcPrimarySMTPAddress'."
        if ($details) { Write-Host $msg }
        Log-Write -Message $msg 
    
        try {
            Add-UnifiedGroupLinks -Identity $srcPrimarySMTPAddress -LinkType Members -Links $srcAdministrativeUsername -ErrorAction Stop
            Add-UnifiedGroupLinks -Identity $srcPrimarySMTPAddress -LinkType Owners -Links  $srcAdministrativeUsername -ErrorAction Stop          
        }
        catch {
            $msg = "      ERROR: Failed to add admin '$srcAdministrativeUsername' as a member and owner to source Office 365 group '$srcPrimarySMTPAddress' ."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
            Log-Write -Message $_.Exception.Message 
        }
    }

    if (-not ([string]::IsNullOrEmpty($dstPrimarySMTPAddress))) {
        $msg = "         INFO: Adding admin '$dstAdministrativeUsername' as a member and owner of the destination Office 365 Group '$dstPrimarySMTPAddress'."
        if ($details) { Write-Host $msg }
        Log-Write -Message $msg 

        try {
            Add-DSTUnifiedGroupLinks -Identity $dstPrimarySMTPAddress -LinkType Members -Links $dstAdministrativeUsername -ErrorAction Stop
            Add-DSTUnifiedGroupLinks -Identity $dstPrimarySMTPAddress -LinkType Owners -Links  $dstAdministrativeUsername -ErrorAction Stop
        }
        catch {
            $msg = "      ERROR: Failed to add admin '$dstAdministrativeUsername' as a member and owner to destination Office 365 group '$dstPrimarySMTPAddress' ."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
            Log-Write -Message $_.Exception.Message 
        }
    }
        
}

Function Unlink-O365Groups {
    param 
    (             
        [parameter(Mandatory = $false)] [String]$srcPrimarySMTPAddress,
        [parameter(Mandatory = $false)] [String]$srcAdministrativeUsername,
        [parameter(Mandatory = $false)] [String]$dstPrimarySMTPAddress,
        [parameter(Mandatory = $false)] [String]$dstAdministrativeUsername,
        [parameter(Mandatory = $false)] [Boolean]$details
    )
   
    if (-not ([string]::IsNullOrEmpty($srcPrimarySMTPAddress))) {  
        $msg = "         INFO: Removing admin '$srcAdministrativeUsername' as a member and owner of the source Office 365 Group '$srcPrimarySMTPAddress'."
        if ($details) { Write-Host $msg }
        Log-Write -Message $msg 
    
        try {
            Remove-UnifiedGroupLinks -Identity $srcPrimarySMTPAddress -LinkType Owners -Links  $srcAdministrativeUsername -ErrorAction Stop -Confirm:$false
            Remove-UnifiedGroupLinks -Identity $srcPrimarySMTPAddress -LinkType Members -Links $srcAdministrativeUsername -ErrorAction Stop -Confirm:$false 
            
            $script:srcServiceAccountRemovedCount += 1          
        }
        catch {
            $msg = "         ERROR: Failed to remove admin '$srcAdministrativeUsername' as a member and owner to source Office 365 group '$srcPrimarySMTPAddress' ."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
            Log-Write -Message $_.Exception.Message 
        }
    }

    if (-not ([string]::IsNullOrEmpty($dstPrimarySMTPAddress))) {
        $msg = "         INFO: Removing admin '$dstAdministrativeUsername' as a member and owner of the destination Office 365 Group '$dstPrimarySMTPAddress'."
        if ($details) { Write-Host $msg }
        Log-Write -Message $msg 

        try {
            Remove-DSTUnifiedGroupLinks -Identity $dstPrimarySMTPAddress -LinkType Owners -Links  $dstAdministrativeUsername -ErrorAction Stop
            Remove-DSTUnifiedGroupLinks -Identity $dstPrimarySMTPAddress -LinkType Members -Links $dstAdministrativeUsername -ErrorAction Stop 
            
            $script:dstServiceAccountRemovedCount += 1           
        }
        catch {
            $msg = "         ERROR: Failed to remove admin '$dstAdministrativeUsername' as a member and owner to destination Office 365 group '$dstPrimarySMTPAddress' ."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red "         $($_.Exception.Message)"
            Log-Write -Message $_.Exception.Message 
        }
    }
        
}

# Function to check if mailbox exists in destination Office 365
Function check-DestinationO365Mailbox {
    param 
    (
        [parameter(Mandatory = $true)] [string]$mailbox,
        [parameter(Mandatory = $false)] [string]$mailboxAlias,
        [parameter(Mandatory = $false)] [Boolean]$maxSendReceiveSize150MB
    )

    #$recipientList = @("UserMailbox","SharedMailbox","RoomMailbox","EquipmentMailbox","TeamMailbox","GroupMailbox","DiscoveryMailbox",
    #                   "MailContact","MailUser","GuestMailUser",
    #                   "MailUniversalDistributionGroup","MailUniversalSecurityGroup","DynamicDistributionGroup","RoomList",
    #                   "PublicFolder")

    $recipient = Get-DSTRecipient -identity $mailbox -ErrorAction SilentlyContinue
    if (!$recipient) {
        $recipient = Get-DSTRecipient -identity $mailboxAlias -ErrorAction SilentlyContinue
    }
    
    $mailboxList = @("UserMailbox", "SharedMailbox", "RoomMailbox", "EquipmentMailbox", "TeamMailbox", "GroupMailbox")

    If ($recipient.RecipientType -in $mailboxList -and $recipient.RecipientypeDetails -ne "DiscoveryMailbox") {  
        if ($maxSendReceiveSize150MB) {
            try {
                $recipient | Set-DSTMailbox -MaxReceiveSize 150MB -MaxSendSize 150MB -ErrorAction Stop -WarningAction SilentlyContinue

                $msg = "         SUCCESS: MaxSendSize and  MaxReceiveSize set to 150MBs for '$($recipient.PrimarySmtpAddress)'."
                Write-Host -ForegroundColor Green  $msg
                Log-Write -Message $msg 
            }
            catch {
                $msg = "         ERROR: Failed to set MaxSendSize and  MaxReceiveSize set to 150MBs for '$($recipient.PrimarySmtpAddress)'."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg  
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message 
            }
        }

        try {
            $dstMailbox = @(Get-DSTMailbox $mailbox -ErrorAction Stop)
            $existingUserPrincipalNames = @(($dstMailbox.split("@")[0]).UserPrincipalName) 
            if ($existingUserPrincipalNames -notmatch "@$script:destinationTenantName.onmicrosoft.com") {
                if ($existingUserPrincipalNames.Count -eq 1) {
                    $script:usersAlreadySyncedWithLocalAD += $mailbox
                }
                else {
                    $script:usersWithExistingUpn += $mailbox
                }
            }
        }
        catch {        
        }
          
        return $true
    }
    else {
        return $false
    }
}

# Function to check if distribution group exists in destination Office 365
Function check-DestinationO365Group {
    param 
    (
        [parameter(Mandatory = $true)]  [string]$group,
        [parameter(Mandatory = $false)] [string]$groupAlias
    )

    $recipient = Get-DSTRecipient -identity $group -ErrorAction SilentlyContinue
    if (!$recipient) {
        $recipient = Get-DSTRecipient -identity $groupAlias -ErrorAction SilentlyContinue
    }
    
    # DynamicDistributionGroup are not supported
    $groupList = @("MailUniversalDistributionGroup", "MailUniversalSecurityGroup", "RoomList")

    If ($recipient.RecipientType -in $groupList) {
        return $true
    }
    else {
        return $false
    }
}

# Function to check if SPO site exists in destination SharePoint Online
Function check-DestinationSPOSite {
    param 
    (
        [parameter(Mandatory = $true)] [string]$url,
        [parameter(Mandatory = $true)] [string]$AdminCenterUrl

    )

    try {
        Connect-SPOService -Url $AdminCenterUrl -Credential $global:btDestinationO365Creds

        $spoSite = Get-SPOSite -identity $url -ErrorAction SilentlyContinue
    }
    catch {

    }

    If ($spoSite) {
        return $true
    }
    else {
        return $false
    }
}

# Function to check if SPO subsite exists in destination SharePoint Online
Function check-DestinationSPOSubsite {
    param 
    (
        [parameter(Mandatory = $true)] [string]$siteUrl,
        [parameter(Mandatory = $true)] [string]$subSiteName

    )

    try {
        if ($useModernAuthentication) {
            Connect-PnPOnline -Url $siteUrl -ClientSecret $script:ClientSecret  -ClientId $script:ClientId -ErrorAction Stop
        }
        else {
            Connect-PnPOnline -Url $siteUrl -Credential $global:btDestinationO365Creds  -ErrorAction Stop 
        }     
    }
    catch {
        $msg = "ERROR: Failed to connect to subsite '$siteUrl'." 
        write-Host -ForegroundColor Red $msg
        Log-Write -Message $msg    
    }
    try {
        $spoSubsite = Get-PnPWeb -ErrorAction Stop
    }
    catch {
        $msg = "ERROR: Failed to get PnPWeb of subsite '$subSiteName' with URL '$siteUrl'." 
        write-Host -ForegroundColor Red $msg
        Log-Write -Message $msg    
    }

    if ($spoSubsite) {
        return $true
    }
    else {
        return $false
    }
}


#######################################################################################################################
#                                   MENU
#######################################################################################################################
# Function to display the main menu
Function Menu {
    #Main menu
    do {
        write-host 
        $msg = "#######################################################################################################################`
                                                           ACTION SELECTION                 `
#######################################################################################################################"
        Write-Host $msg
        
        $confirm = (Read-Host -prompt "
1. Office 365 tenant and BitTitan license analysis and create MigrationWiz projects for Office 365 tenant migration
2. Apply BitTitan Licenses to MigrationWiz migrations `
3. Delete created MigrationWiz endpoints and projects
-----------------------------------------------------------------------------------------------------------------------
4. Exit

Select 1-5")

        if ($confirm -eq 1) {
            $script:createMigrationWizProjects = $true
            $script:ApplyBitTitanLicenses = $false
        }
        elseif ($confirm -eq 2) {
            $script:createMigrationWizProjects = $false
            $script:ApplyBitTitanLicenses = $true
        }
        elseif ($confirm -eq 3) {
            $script:createMigrationWizProjects = $false
            $script:ApplyBitTitanLicenses = $false
        }
        elseif ($confirm -eq 4) {
            $script:createMigrationWizProjects = $false
            $script:ApplyBitTitanLicenses = $false
        }
        elseif ($confirm -eq 5) {
            write-Host
            Exit
        }

    } while (!(isNumeric($confirm)) -or $confirm -notmatch '[1-5]')
        
    Return 1
}

#######################################################################################################################
#                   MAIN PROGRAM
#######################################################################################################################

Import-PowerShellModules
Import-MigrationWizPowerShellModule

#######################################################################################################################
#                   CUSTOMIZABLE VARIABLES  
#######################################################################################################################

$script:migrateEntireTenant = $false
$script:useTenantAssessment = $false

$getAssessment = $true
$script:getMailboxStatistics = $true
$checkDestinationSpoStructure = $false
$checkDestinationMailbox = $false

$updateEndpoint = $true
$updateConnector = $true

$modernAuth = $false
$size150MB = $true

$numberBatches = 1
$userThreshold = 10000
$teamsThreshold = 500

$linkO365Groups = $true
$customDocumentLibrary = $false

$script:usersAlreadySyncedWithLocalAD = @()
$script:usersWithExistingUpn = @()

###################################################################################################################
$script:srcGermanyCloud = $false
$script:srcUsGovernment = $False

$script:dstGermanyCloud = $False
$script:dstUsGovernment = $false
                        
$ZoneRequirement1 = "NorthAmerica"   #North America (Virginia). For Azure: Both AZNAE and AZNAW.
$ZoneRequirement2 = "WesternEurope"  #Western Europe (Amsterdam for Azure, Ireland for AWS). For Azure: AZEUW.
$ZoneRequirement3 = "AsiaPacific"    #Asia Pacific (Singapore). For Azure: AZSEA
$ZoneRequirement4 = "Australia"      #Australia (Asia Pacific Sydney). For Azure: AZAUE - NSW.
$ZoneRequirement5 = "Japan"          #Japan (Asia Pacific Tokyo). For Azure: AZJPE - Saltiama.
$ZoneRequirement6 = "SouthAmerica"   #South America (Sao Paolo). For Azure: AZSAB.
$ZoneRequirement7 = "Canada"         #Canada. For Azure: AZCAD.
$ZoneRequirement8 = "NorthernEurope" #Northern Europe (Dublin). For Azure: AZEUN.
$ZoneRequirement9 = "China"          #China.
$ZoneRequirement10 = "France"         #France.
$ZoneRequirement11 = "SouthAfrica"    #South Africa.

if ([string]::IsNullOrEmpty($global:btZoneRequirement)) {
    $global:btZoneRequirement = $ZoneRequirement1
}
#######################################################################################################################
#                       SELECT WORKING DIRECTORY  
#######################################################################################################################

Write-Host
Write-Host
Write-Host -ForegroundColor Yellow "             BitTitan Office 365 Tenant to Tenant migration project creation tool."
Write-Host

write-host 
$msg = "#######################################################################################################################`
                       SELECT WORKING DIRECTORY                  `
#######################################################################################################################"
Write-Host $msg
write-host 

#Working Directorys
$script:workingDir = "C:\scripts"

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

#Logs directory
$logDirName = "LOGS"
$logDir = "$script:workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format "yyyyMMddTHHmmss")_Create-O365T2TMigrationWizProjects.log"
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

#######################################################################################################################
#         INFINITE LOOP FOR MENU
#######################################################################################################################

# keep looping until specified to exit
do {
    #Select Action
    $action = Menu
    if ($action -ne $null) {

        if ($script:createMigrationWizProjects) {
            write-host 
            $msg = "#######################################################################################################################`
                       AZURE CLOUD SELECTION                 `
#######################################################################################################################"
            Write-Host $msg
            Write-Host

            if ($script:srcGermanyCloud) {
                Write-Host -ForegroundColor Magenta "WARNING: Connecting to (source) Azure Germany Cloud." 

                Write-Host
                do {
                    $confirm = (Read-Host -prompt "Do you want to switch to (destination) Azure Cloud (global service)?  [Y]es or [N]o")  
                    if ($confirm.ToLower() -eq "y") {
                        $script:srcGermanyCloud = $false
                    }  
                } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
                Write-Host 
            }
            elseif ($script:srcUsGovernment ) {
                Write-Host -ForegroundColor Magenta "WARNING: Connecting to (source) Azure Goverment Cloud." 

                Write-Host
                do {
                    $confirm = (Read-Host -prompt "Do you want to switch to (destination) Azure Cloud (global service)?  [Y]es or [N]o")  
                    if ($confirm.ToLower() -eq "y") {
                        $script:srcUsGovernment = $false
                    }  
                } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
                Write-Host 
            }

            if ($script:dstGermanyCloud) {
                Write-Host -ForegroundColor Magenta "WARNING: Connecting to (destination) Azure Germany Cloud." 

                Write-Host
                do {
                    $confirm = (Read-Host -prompt "Do you want to switch to (destination) Azure Cloud (global service)?  [Y]es or [N]o")  
                    if ($confirm.ToLower() -eq "y") {
                        $script:dstGermanyCloud = $false
                    }  
                } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
                Write-Host 
            }
            elseif ($script:dstUsGovernment) {
                Write-Host -ForegroundColor Magenta "WARNING: Connecting to (destination) Azure Goverment Cloud." 

                Write-Host
                do {
                    $confirm = (Read-Host -prompt "Do you want to switch to (destination) Azure Cloud (global service)?  [Y]es or [N]o")  
                    if ($confirm.ToLower() -eq "y") {
                        $script:dstUsGovernment = $false
                    }  
                } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
                Write-Host 
            }

            Write-Host -ForegroundColor Green "INFO: Using Azure $global:btZoneRequirement Datacenter." 

            if ([string]::IsNullOrEmpty($BitTitanAzureDatacenter)) {
                if (!$global:btCheckAzureDatacenter) {
                    Write-Host
                    do {
                        $confirm = (Read-Host -prompt "Do you want to switch the Azure Datacenter to another region?  [Y]es or [N]o")  
                        if ($confirm.ToLower() -eq "y") {
                            do {
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
        10. France        #France.
        11. SouthAfrica   #South Africa.

        Select 0-11")
                                switch ($ZoneRequirementNumber) {
                                    1 { $ZoneRequirement = 'NorthAmerica' }
                                    2 { $ZoneRequirement = 'WesternEurope' }
                                    3 { $ZoneRequirement = 'AsiaPacific' }
                                    4 { $ZoneRequirement = 'Australia' }
                                    5 { $ZoneRequirement = 'Japan' }
                                    6 { $ZoneRequirement = 'SouthAmerica' }
                                    7 { $ZoneRequirement = 'Canada' }
                                    8 { $ZoneRequirement = 'NorthernEurope' }
                                    9 { $ZoneRequirement = 'China' }
                                    10 { $ZoneRequirement = 'France' }
                                    11 { $ZoneRequirement = 'SouthAfrica' }
                                }
                            } while (!(isNumeric($ZoneRequirementNumber)) -or !($ZoneRequirementNumber -in 1..11))

                            $global:btZoneRequirement = $ZoneRequirement
                
                            Write-Host 
                            Write-Host -ForegroundColor Yellow "WARNING: Now using Azure $global:btZoneRequirement Datacenter." 

                            $global:btCheckAzureDatacenter = $true
                        }  
                        if ($confirm.ToLower() -eq "n") {
                            $global:btCheckAzureDatacenter = $true
                        }
                    } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
                }
                else {
                    Write-Host
                    $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different Azure datacenter."
                    Write-Host -ForegroundColor Yellow $msg
                }
            }
            else {
                $global:btZoneRequirement = $BitTitanAzureDatacenter
            }

        }

        do {
            write-host 
            $msg = "#######################################################################################################################`
                       PROJECT TYPE CREATION                `
#######################################################################################################################"
            Write-Host $msg
            Log-Write -Message "PROJECT TYPE CREATION" 

            #################################
            #SWITCHES
            #################################

            write-host 
            do {
                if (!$script:createMigrationWizProjects -and !$script:ApplyBitTitanLicenses ) {
                    $confirm = (Read-Host -prompt "Do you want to delete SharePoint Online Classic Team Sites MigrationWiz projects?  [Y]es or [N]o")
                }    
                else {     
                    $confirm = (Read-Host -prompt "Do you want to migrate SharePoint Online Classic Team Sites?  [Y]es or [N]o")
                }
                if ($confirm.ToLower() -eq "y") {
                    #$migrateSpoTeamSites = $true
                    $migrateNewSpoTeamSites = $true
                }
                else {
                    $migrateSpoTeamSites = $false 
                    $migrateNewSpoTeamSites = $false 
                }
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 

            do {
                if (!$script:createMigrationWizProjects -and !$script:ApplyBitTitanLicenses ) {
                    $confirm = (Read-Host -prompt "Do you want to delete O365 (unified) Groups  MigrationWiz projects?  [Y]es or [N]o")
                }    
                else {     
                    $confirm = (Read-Host -prompt "Do you want to migrate O365 (unified) Groups?  [Y]es or [N]o")
                }
                if ($confirm.ToLower() -eq "y") {
                    #$migrateO365Groups = $true
                    $migrateNewO365Groups = $true
                }
                else {
                    $migrateO365Groups = $false
                    $migrateNewO365Groups = $false
                }
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 

            write-host 

        }
        while (!$migrateSpoTeamSites -and !$migrateO365Groups -and !$migrateNewSpoTeamSites -and !$migrateNewO365Groups)

        if ($script:createMigrationWizProjects -or $script:ApplyBitTitanLicenses) {
            write-host 
            $msg = "#######################################################################################################################`
                       PROJECT SCOPE               `
#######################################################################################################################"
            Write-Host $msg
            Log-Write -Message "PROJECT SCOPE" 
            write-host 

            do {
                $confirm = (Read-Host -prompt "Do you want to migrate the entire Office 365 tenant?  [Y]es or [N]o")
                if ($confirm.ToLower() -eq "y") {
                    $script:migrateEntireTenant = $true
                }
                else {
                    $script:migrateEntireTenant = $false 
                }
            } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n")) 
        }


        write-host 
        $msg = "#######################################################################################################################`
                       WORKGROUP, CUSTOMER AND ENDPOINTS SELECTION              `
#######################################################################################################################"
        Write-Host $msg
        Log-Write -Message "WORKGROUP, CUSTOMER AND ENDPOINTS SELECTION" 

        if (!$global:btCheckCustomerSelection) {
            do {
                #Select workgroup
                $global:btWorkgroupId = Select-MSPC_WorkGroup

                #Select customer
                $customer = Select-MSPC_Customer -Workgroup $global:btWorkgroupId
            }
            while ($customer -eq "-1")

            $global:btCustomerOrganizationId = $customer.OrganizationId.Guid
    
            $global:btCheckCustomerSelection = $true  
        }
        else {
            Write-Host
            $msg = "INFO: Already selected workgroup '$global:btWorkgroupId' and customer '$global:btcustomerName'."
            Write-Host -ForegroundColor Green $msg

            Write-Host
            $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different workgroups/customers."
            Write-Host -ForegroundColor Yellow $msg
        }

        $script:customerTicket = Get-BT_Ticket -Ticket $script:ticket -OrganizationId $global:btCustomerOrganizationId #-ElevatePrivilege
        $script:workgroupTicket = Get-BT_Ticket -Ticket $script:ticket -OrganizationId $global:btWorkgroupOrganizationId #-ElevatePrivilege

        if (!$global:btCheckO365Connection) {
            $useMspcEndpoints = $true

            $script:sourceO365Session = Connect-SourceExchangeOnline
            $script:destinationO365Session = Connect-DestinationExchangeOnline

            $global:btCheckO365Connection = $true
        }
        else {
            write-host 
            $msg = "#######################################################################################################################`
                       CONNECTION TO SOURCE/DESTINATION OFFICE 365 TENANTS             `
#######################################################################################################################"
            Write-Host $msg
            Log-Write -Message $msg

            Write-Host
            $msg = "INFO: Already connected to source and destination Office 365 tenants."
            Write-Host -ForegroundColor Green $msg

            Write-Host
            $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different tenants."
            Write-Host -ForegroundColor Yellow $msg

        }

        $script:SrcAdministrativeUsername = $global:btSourceO365Creds.UserName
        $SecureString = $global:btSourceO365Creds.Password
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        $script:sourcePlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

        $script:DstAdministrativeUsername = $global:btDestinationO365Creds.UserName
        $SecureString = $global:btDestinationO365Creds.Password
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        $script:destinationPlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
 
        $exportEndpointId = $global:btExportEndpointId

        $script:sourceTenantDomain = (Get-O365TenantDomain -Credentials $global:btSourceO365Creds -SourceOrDestination "source").ToLower()

        if ($script:srcGermanyCloud) {
            $script:sourceTenantName = $script:sourceTenantDomain.replace(".onmicrosoft.de", "")
            $sSPOAdminCenterUrl = "https://$script:sourceTenantName-admin.sharepoint.de/"
            $sSPOUrl = "https://$script:sourceTenantName.sharepoint.de/"
        }
        elseif ($script:srcUsGovernment) {
            $script:sourceTenantName = $script:sourceTenantDomain.replace(".onmicrosoft.us", "")
            $sSPOAdminCenterUrl = "https://$script:sourceTenantName-admin.sharepoint.us/"
            $sSPOUrl = "https://$script:sourceTenantName.sharepoint.us/"
        }
        else {            
            $script:sourceTenantName = $script:sourceTenantDomain.replace(".onmicrosoft.com", "")
            $sSPOAdminCenterUrl = "https://$script:sourceTenantName-admin.sharepoint.com/"
            $sSPOUrl = "https://$script:sourceTenantName.sharepoint.com/"
        }

        $sourceVanityDomains = @(Get-VanityDomains -Credentials $global:btSourceO365Creds -SourceOrDestination "source")

        $importEndpointId = $global:btImportEndpointId

        $script:destinationTenantDomain = (Get-O365TenantDomain -Credentials $global:btDestinationO365Creds -SourceOrDestination "destination").ToLower()

        if ($script:dstGermanyCloud) {
            $script:destinationTenantName = $script:destinationTenantDomain.replace(".onmicrosoft.de", "")
            $dSPOAdminCenterUrl = "https://$script:destinationTenantName-admin.sharepoint.de/"
            $dSPOUrl = "https://$script:destinationTenantName.sharepoint.de/"
        }
        elseif ($script:dstUsGovernment) {
            $script:destinationTenantName = $script:destinationTenantDomain.replace(".onmicrosoft.us", "")
            $dSPOAdminCenterUrl = "https://$script:destinationTenantName-admin.sharepoint.us/"
            $dSPOUrl = "https://$script:destinationTenantName.sharepoint.us/"
        }
        else {            
            $script:destinationTenantName = $script:destinationTenantDomain.replace(".onmicrosoft.com", "")
            $dSPOAdminCenterUrl = "https://$script:destinationTenantName-admin.sharepoint.com/"
            $dSPOUrl = "https://$script:destinationTenantName.sharepoint.com/"
        }

        if (!$script:createMigrationWizProjects -and !$script:ApplyBitTitanLicenses ) {

            #######################################################################################################################
            #         MIGRATIONWIZ ACCOUNT CLEAN-UP
            #######################################################################################################################
            write-host 
            $msg = "#######################################################################################################################`
           MIGRATIONWIZ ACCOUNT CLEAN-UP                  `
#######################################################################################################################"
            Write-Host $msg
            Log-Write -Message "MIGRATIONWIZ ACCOUNT CLEAN-UP" 
            Write-Host
    
            Write-Host
            $msg = "INFO: Deleting MigrationWiz projects."
            Write-Host $msg
            Log-Write -Message $msg 
            Write-Host 
            #delete projects

            if ($migrateSpoTeamSites -and $migrateNewSpoTeamSites) {
                Remove-MW_Connectors -CustomerOrganizationId $global:btCustomerOrganizationId -ProjectType "Storage" -ProjectName "ClassicSPOSite-Document-"
            }
            if ($migrateO365Groups -and $migrateNewO365Groups) {
                Remove-MW_Connectors -CustomerOrganizationId $global:btCustomerOrganizationId -ProjectType "Mailbox" -ProjectName "O365Group-Mailbox-All conversations"
                Remove-MW_Connectors -CustomerOrganizationId $global:btCustomerOrganizationId -ProjectType "Storage" -ProjectName "O365Group-Document-"
            }

            Write-Host 
            Write-Host
            $msg = "INFO: Deleting MigrationWiz endpoints."
            Write-Host $msg
            Log-Write -Message $msg 
            Write-Host 

            #delete endpoints

            if ($migrateSpoTeamSites) {
                Remove-MSPC_Endpoints -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointType "SharePoint" -EndPointName "SRC-SPO-"
                Remove-MSPC_Endpoints -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointType "SharePointOnlineAPI" -EndPointName "DST-SPO-"
            }
            if ($migrateO365Groups) {
                Remove-MSPC_Endpoints -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointType "Office365Groups" -EndpointName "SRC-O365G-"
                Remove-MSPC_Endpoints -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointType "SharePointOnlineAPI" -EndpointName "DST-O365G-"
            }

            Continue
        }

        #######################################################################################################################
        #         PRE-MIGRATION ASSESSMENT
        #######################################################################################################################
        if ($getAssessment) {
            write-host 
            $msg = "#######################################################################################################################`
                       PRE-MIGRATION OFFICE 365 TENANT ASSESSMENT                  `
#######################################################################################################################"
            Write-Host $msg
            Log-Write -Message "PRE-MIGRATION OFFICE 365 TENANT ASSESSMENT" 
   
            if ($migrateSpoTeamSites -or $migrateNewSpoTeamSites) {
                Get-SourceTenantAssessment -migrateSpoTeamSites $($migrateSpoTeamSites -or $migrateNewSpoTeamSites)
            } 
    
            if ($migrateO365Groups -or $migrateNewO365Groups) {
                Get-SourceTenantAssessment -migrateO365Groups $($migrateO365Groups -or $migrateNewO365Groups)
            } 

        }

        #######################################################################################################################
        #         BITTITAN LICENSES
        #######################################################################################################################
        if ($script:unifiedGroupsCount -ne 0) {
            write-host 
            $msg = "#######################################################################################################################`
      BITTITAN LICENSES                  `
#######################################################################################################################"
            Write-Host $msg
            Log-Write -Message "BITTITAN LICENSES" 
            Write-Host

            if ($migrateNewSpoTeamSites) {
                $msg = "ACTION: Shared Document licenses required to migrate Classic SPO sites: $script:allDocumentLibrariesCount (50 GBs transfer limit)."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 
            }
            
            if ($migrateNewO365Groups) {
                $msg = "ACTION: Shared Document licenses required to migrate Office 365 Groups document libraries: $script:allUnifiedGroupDocumentLibrariesCount (50 GBs transfer limit)."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 
            }

            if ($migrateO365Groups) {
                $msg = "ACTION: User Migration Bundle licenses required to migrate Office 365 Groups: $script:unifiedGroupsCount (without data transfer limit)."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 
            }
        }

        #######################################################################################################################
        #      EMAIL ADDRESS MAPPING (SRC->DST)
        #######################################################################################################################
        if ($migrateO365Groups -or $migrateNewO365Groups) {
            write-host 
            $msg = "#######################################################################################################################`
 EMAIL ADDRESS MAPPING (SRC->DST)                 `
#######################################################################################################################"
            Write-Host $msg
            Log-Write -Message "EMAIL ADDRESS MAPPING (SRC->DST)" 

            $sourceAlreadyProccesedDomains = @()
            $destinationAlreadyProccesedDomains = @()

            write-host 
            $needRecipientMapping = Apply-EmailAddressMapping

            if ($needRecipientMapping) {
        
                $recipientMapping = @()

                if ($script:sameEmailAddresses) {
            
                    if ($script:selectedDomains.Count -eq 1) {
                        $recipientMapping = "RecipientMapping=`"@$script:sourceTenantDomain->@$script:selectedDomains`""

                        $msg = "INFO: Since you are migrating to the same email addresses, this '$recipientMapping' will be applied."
                        Write-Host $msg
                        Log-Write -Message $msg 

                    }
                    else {
                        $msg = "INFO: Since you are migrating to the same email addresses, but there are several source domains, you may need a RecipientMapping CSV file with SourceEmailAddress and DestinationEmailAddress."
                        Write-Host $msg
                        Log-Write -Message $msg 
                        $msg = "      If you don't need a RecipientMapping just press Cancel."
                        Write-Host $msg
                        Log-Write -Message $msg 

                        $recipientMapping = Import-CSV_RecipientMapping ($script:mailboxes) 
                
                        if (!$recipientMapping) {
                            $msg = "WARNING: You have not provided the RecipientMapping CSV file with SourceEmailAddress and DestinationEmailAddress. You may need to add it to the project later."
                            Write-Host -ForegroundColor Yellow $msg
                            Log-Write -Message $msg 
                        } 
                    }
                }
                elseif (!$script:sameEmailAddresses -and $script:sameUserName -and $script:differentDomain ) {
                    if ($sourceVanityDomains.Count -eq 1 -and $script:selectedDomains.count -eq 1 ) {

                        $recipientMapping = "RecipientMapping=`"@$script:sourceTenantDomain->@$script:selectedDomains`""

                        $msg = "INFO: Since you are migrating to a different domain but with same email prefixes, this recipient mapping will be applied: "
                        Write-Host $msg
                        Log-Write -Message $msg 

                        $msg = "$recipientMapping"
                        Write-Host -ForegroundColor Yellow $msg
                        Log-Write -Message $msg 
                    }
                    elseif ($sourceVanityDomains.Count -gt 1 -or $script:selectedDomains.count -gt 1) {

                        $msg = "INFO: Since you are migrating to the same email email prefixes but there are several selected source domains, please select the RecipientMapping CSV file with SourceEmailAddress and DestinationEmailAddress."
                        Write-Host $msg
                        Log-Write -Message $msg 

                        $recipientMapping = Import-CSV_RecipientMapping ($script:mailboxes) 

                        $msg = "$recipientMapping"
                        Write-Host -ForegroundColor Yellow $msg
                        Log-Write -Message $msg 
  
                    }
                }
                elseif (!$script:sameEmailAddresses -and !$script:sameUserName) {

                    $msg = "ACTION: Since you are migrating to different email prefixes, please select the RecipientMapping CSV file with SourceEmailAddress and DestinationEmailAddress."
                    Write-Host -ForegroundColor Yellow $msg
                    Log-Write -Message $msg 

                    $recipientMapping = Import-CSV_RecipientMapping ($script:mailboxes) 

                    Write-Host
                    $msg = "These RecipientMappings will be applied to the mailbox projects:"
                    Write-Host -ForegroundColor Yellow $msg
                    Log-Write -Message $msg 
                    $msg = "$recipientMapping"
                    Write-Host -ForegroundColor Yellow $msg
                    Log-Write -Message $msg 
                }
            }
        }

        if ($script:createMigrationWizProjects) {

            #######################################################################################################################
            #         BITTITAN LICENSE PACK
            #######################################################################################################################
            if ($script:ApplyBitTitanLicenses) {
                write-host 
                $msg = "#######################################################################################################################`
                  GETTING BITTITAN LICENSE PACK                 `
#######################################################################################################################"
                Write-Host $msg
                Log-Write -Message "GETTING BITTITAN LICENSE PACK" 
                write-host     
    
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
                $licensesPack = @(Get-MW_LicensePack -Ticket $script:MwTicket -WorkgroupOrganizationId $global:btWorkgroupOrganizationId -ProductSkuId $productId | Where-Object { $_.ExpireDate -gt $curDate } | where { (($_.Purchased -eq 1 -or $_.Granted -eq 1) -and $_.Revoked -eq 0) -and ($_.Used -eq 1) })
                $licensesAvailable = 0

                if ( ! ($licensesPack) ) {
                    $msg = "ERROR: No valid license pack found on this MSPC Workgroup / Account"
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                }
                else {
                    $licensesAvailable = 0
                    foreach ( $license in $licensesPack ) {
                        # Ignoring the Refunded and revoked. Don't know if important for the calculations or not.
                        $licensesAvailable += $license.purchased + $license.granted - $license.used - $license.revoked
                    }
                    $msg = "INFO: $licensesAvailable User Migration Bundle licenses found on this MSPC Workgroup / Account"
                    Write-Host -ForegroundColor Green  $msg
                    Log-Write -Message $msg
                }
            }

            $MigrationWizProjectArray = @()

            ##########################################################################################################################################
            #            SPO Team Sites Document Libraries
            ##########################################################################################################################################
            $connectorId = $null

            if ($migrateSpoTeamSites) {
                write-host 
                $msg = "#######################################################################################################################`
                  CREATING CLASSIC TEAM SITE PROJECTS                  `
#######################################################################################################################"
                Write-Host $msg
                Log-Write -Message "CREATING CLASSIC TEAM SITE PROJECTS" 
                Write-Host
    
                $msg = "INFO: Processing Classic Team site migration."
                Write-Host $msg
                Log-Write -Message $msg 

                if ($script:importedClassicTeamSites -eq $null -or $script:importedDocumentLibraries -eq $null) {
                    if ($script:importedClassicTeamSites -eq $null) {
                        $msg = "INFO: No SPO Team Sites found. Skipping SPO Team Site project creation."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg 
                    }
                    if ($script:importedDocumentLibraries -eq $null) {
                        $msg = "INFO: No Root Document Libraries or Root Subsites and their Document Libraries found. Skipping SPO Team Site project creation."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg 
                    }
                }
                else {
                    if (!$UseOwnAzureStorage -and !$AzureStorageSelected) {
                        Write-Host   
                        do {
                            $confirm = (Read-Host -prompt "Do you want to use Microsoft provided Azure Storage?  [Y]es or [N]o")
                            if ($confirm.ToLower() -eq "n") {
                                $UseOwnAzureStorage = $true

                                $AzureStorageAccountName = ''
                                $AzureAccountKey = ''
                
                                do {
                                    $AzureStorageAccountName = (Read-Host -prompt "         Enter Azure Storage Account Name")
                                } while ($AzureStorageAccountName -eq "")

                                do {
                                    $AzureAccountKey = (Read-Host -prompt "         Enter Azure Storage Primary Access Key")
                                } while ($AzureAccountKey -eq "")

                                $AzureStorageSelected = $true
                            }
                            else {
                                $UseOwnAzureStorage = $false
                                $AzureStorageSelected = $true
                            }
                        } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
                    }

                    $alreadyCreated = $false

                    $subsiteDocumentLibraries = @($script:importedDocumentLibraries | Where-Object { $_.isSubsite -eq $true })
               
                    $processedRootDocumentLibraries = 0
                    $processedRootSubsites = 0
                    $processedRootSubsitesDocumentLibraries = 0
                    $processedClassicTeamSites = 0
                    $processedDocumentLibraries = 0
                    $processedSubsite = 0
                    $allProcessedSubsites = 0
                    $processedSubsiteDocumentLibraries = 0

                    $currentSubsite = 0

                    ##################################################################################################
                    #    ROOT DOCUMENT LIBRARIES
                    ##################################################################################################
                    $currentRootDocumentLibrary = 0
        
                    if ($script:importedRootDocumentLibraries -ne $null) {
                        foreach ($rootDocumentLibrary in $script:importedRootDocumentLibraries) {

                            $currentRootDocumentLibrary += 1

                            $urlparts = $($rootDocumentLibrary.DocumentLibraryUrl).split("/")
                            $rootDocumentLibraryName = $urlparts[-1]

                            if ($rootDocumentLibrary.TeamSiteURL -eq $rootDocumentLibrary.NewDestinationTeamSiteURL) {    
                                $sSPOUrl = $rootDocumentLibrary.TeamSiteUrl -replace $script:destinationTenantName, $script:sourceTenantName
                                $dSPOUrl = $sSPOUrl -replace $script:sourceTenantName, $script:destinationTenantName
                            }
                            else {
                                $sSPOUrl = $rootDocumentLibrary.TeamSiteUrl -replace $script:destinationTenantName, $script:sourceTenantName
                                $dSPOUrl = $rootDocumentLibrary.NewDestinationTeamSiteURL -replace $script:sourceTenantName, $script:destinationTenantName  
                            }

                
                            write-host 
                            $msg = "INFO: Processing $currentRootDocumentLibrary/$script:rootDocumentLibrariesCount Root Document Library '$rootDocumentLibraryName' of Root Site '$sSPOUrl'."
                            Write-Host $msg
                            Log-Write -Message $msg 
                
                            #Create SPO endpoints
                            $msg = "INFO: Creating MSPC endpoints for Root Document Library '$sSPOUrl'."
                            Write-Host $msg
                            Log-Write -Message $msg 
                
                            $exportEndpointName = "SRC-SPO-Root-RootDocumentLibrary-$sSPOUrl-$script:sourceTenantName"
                            $exportType = "SharePoint"
                            $exportTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                            $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
                                "Url"                          = $sSPOUrl;
                                "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                                "AdministrativePassword"       = $script:sourcePlainPassword;
                                "UseAdministrativeCredentials" = $true
                            }
                
                            $importEndpointName = "DST-SPO-Root-RootDocumentLibrary-$dSPOUrl-$script:destinationTenantName"
                            $importType = "SharePointOnlineAPI"
                            $importTypeName = "MigrationProxy.WebApi.SharePointOnlineConfiguration"
                            if (!$UseOwnAzureStorage) {
                                $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                    "Url"                                = $dSPOUrl;
                                    "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                    "AdministrativePassword"             = $script:destinationPlainPassword;
                                    "UseAdministrativeCredentials"       = $true;
                                    "UseSharePointOnlineProvidedStorage" = $true 
                                }
                            }
                            else {
                                $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                    "Url"                                = $dSPOUrl;
                                    "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                    "AdministrativePassword"             = $script:destinationPlainPassword;
                                    "UseAdministrativeCredentials"       = $true;
                                    "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                    "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;    
                                    "UseSharePointOnlineProvidedStorage" = $false
                                }
                            }
                
                            #Create SPO Team Sites source endpoint
                            $exportEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType $exportType  -EndpointName $exportEndpointName -EndpointConfiguration $exportConfiguration -update $updateEndpoint
                            #Create SPO Team Sites destination endpoint
                            $importEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType $importType -EndpointName $importEndpointName -EndpointConfiguration $importConfiguration -update $updateEndpoint
                
                            #Create SPO Team Sites Document project
                
                            $ProjectName = "ClassicSPOSite-Document-Root-RootDocumentLibrary-$sSPOUrl-$script:sourceTenantName"
                            $ProjectType = "Storage"
                
                            if ($enableModernAuth) {
                                $advancedOptions = "$modernAuth InitializationTimeout=8 IncreasePathLengthLimit=1"
                            }
                            else {
                                $advancedOptions = "InitializationTimeout=8 IncreasePathLengthLimit=1"
                            }
                
                            if ($applicationPermissions) {
                                $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                            }
                            if ($UseDelegatePermission) {
                                $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                            }
                
                            $connectorId = $null
                            $connectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                                -ProjectName $ProjectName `
                                -ProjectType $ProjectType `
                                -exportType $exportType `
                                -importType $importType `
                                -exportEndpointId $exportEndpointId `
                                -importEndpointId $importEndpointId `
                                -exportConfiguration $exportConfiguration `
                                -importConfiguration $importConfiguration `
                                -advancedOptions $advancedOptions `
                                -maximumSimultaneousMigrations 100 `
                                -ZoneRequirement $global:btZoneRequirement `
                                -MaxLicensesToConsume 1 `
                                -updateConnector $true                        

                            if ($connectorId -ne $null) {

                                $documentLibraryName = $rootDocumentLibrary.DocumentLibraryName

                                if ([string]::IsNullOrEmpty($documentLibraryName)) { Continue }

                                $msg = "      INFO: Processing Document Library '$documentLibraryName'."
                                Write-Host $msg
                                Log-Write -Message $msg 

                                if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                    # Renew MW ticket
                                    Connect-BitTitan 

                                    Write-Host
                                    $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                    Write-Host -ForegroundColor Magenta $msg
                                    Log-Write -Message $msg 
                                    Write-Host                    
                                }   

                                try {
                                    $ImportLibrary = $documentLibraryName
                                    $ExportLibrary = $documentLibraryName

                                    $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
                                    if (!$result) {
                                        $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId  -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary

                                        $msg = "      SUCCESS: Document Library migration '$ExportLibrary->$ImportLibrary' added to connector." 
                                        write-Host -ForegroundColor Green $msg
                                        Log-Write -Message $msg  

                                        $processedRootDocumentLibraries += 1
                                    }
                                    else {
                                        $msg = "      WARNING: Document Library migration '$ExportLibrary->$ImportLibrary' already exists in connector." 
                                        write-Host -ForegroundColor Yellow $msg
                                        Log-Write -Message $msg  

                                        $processedRootDocumentLibraries += 1
                                    }
                                }
                                catch {
                                    $msg = "      ERROR: Failed to add Document Library migration '$ExportLibrary->$ImportLibrary' to connector." 
                                    write-Host -ForegroundColor Red $msg
                                    Log-Write -Message $msg     
                                }

                                [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $ProjectName; ProjectType = $ProjectType; ConnectorId = $connectorId; }
                            }



                        }
                    }

                    ##################################################################################################
                    #    ROOT SUBSITES
                    ##################################################################################################
                    $currentRootSubsiteDocumentLibrary = 0
                    $subsiteRootDocumentLibraryCount = $subsiteDocumentLibraries.Count
                    $totalSubsitesCount += $subsiteRootDocumentLibraryCount

                    $previousSubSiteName = ""
        
                    if ($script:importedRootSubSiteDocumentLibraries -ne $null) {
                        foreach ($rootSubsiteDocumentLibrary in $script:importedRootSubSiteDocumentLibraries) {

                            if ($ProcessedRootSubsites -ne $rootSubsiteDocumentLibrary.SubsiteNumber.SubString(0, 1)) {
                                $ProcessedRootSubsites += 1 
                            }  
                
                            $currentRootSubsiteDocumentLibrary += 1

                
                            $urlparts = $($rootSubsiteDocumentLibrary.TeamSiteUrl).split("/")
                            $rootSubsiteDepth = $urlparts.Count - 4
                            $rootSubsiteName = $urlparts[-1]
                            $parentRootSubsiteUrl = ($rootSubsiteDocumentLibrary.TeamSiteUrl -split "/$rootSubsiteName")[0] 
                            $rootSubsiteUrl = $parentRootSubsiteUrl + "/" + $rootSubsiteName

                            if ($rootDocumentLibrary.TeamSiteURL -ne $rootDocumentLibrary.NewDestinationTeamSiteURL) { 
                                $dstUrlparts = $($rootSubsiteDocumentLibrary.NewDestinationTeamSiteURL).split("/")
                                $dstRootSubsiteName = $dstUrlparts[-1]
                                $dstParentRootSubsiteUrl = ($rootSubsiteDocumentLibrary.NewDestinationTeamSiteURL -split "/$dstRootSubsiteName")[0] 
                                $dstRootSubsiteUrl = $dstParentRootSubsiteUrl + "/" + $dstRootSubsiteName
                            }               
                
                            if ($rootSubsiteName -ne $previousRootSubsiteName) {
                                $currentRootSubsite += 1

                                if ($rootDocumentLibrary.TeamSiteURL -eq $rootDocumentLibrary.NewDestinationTeamSiteURL) { 
                                    $srcRootSubsiteUrl = $rootSubsiteUrl -replace $script:destinationTenantName, $script:sourceTenantName
                                    $dstRootSubsiteUrl = $rootSubsiteUrl -replace $script:sourceTenantName, $script:destinationTenantName
                                }
                                else {
                                    $srcRootSubsiteUrl = $rootSubsiteUrl -replace $script:destinationTenantName, $script:sourceTenantName
                                    $dstRootSubsiteUrl = $dstRootSubsiteUrl -replace $script:sourceTenantName, $script:destinationTenantName    
                                }

                                if ($rootSubsiteDepth -eq 0) {  
                                    write-host 
                                    $msg = "INFO: Processing $currentRootSubsite/$script:rootSubsitesCount Root Subsite '$rootSubsiteName' level-$rootSubsiteDepth of Root Site '$sSPOUrl' with URL '$srcRootSubsiteUrl'."
                                    Write-Host $msg
                                    Log-Write -Message $msg 
                                }
                                elseif ($rootSubsiteDepth -gt 0) {
                                    write-host 
                                    $msg = "INFO: Processing $currentRootSubsite/$script:rootSubsitesCount Root Subsite '$rootSubsiteName' level-$rootSubsiteDepth  of Root Subsite '$parentRootSubsiteUrl' with URL '$srcRootSubsiteUrl'."
                                    Write-Host $msg
                                    Log-Write -Message $msg 
                                }

                                if ($checkDestinationSPOStructure) { $result = check-DestinationSPOSubsite -siteUrl $dstRootSubsiteUrl -SubSiteName $rootSubsiteName }
                                if (!$result -and $checkDestinationSpoStructure) {
                                    $msg = "INFO: Skipping classic Team Root Subsite project creation. Root Subsite '$rootSubsiteName' with URL '$dstRootSubsiteUrl' not found in destination SharePoint Online."
                                    Write-Host -ForegroundColor Red $msg
                                    Log-Write -Message $msg 

                                    Continue
                                }   

                                #Create SPO endpoints
                                $msg = "INFO: Creating MSPC endpoints for Root Site '$sSPOUrl'."
                                Write-Host $msg
                                Log-Write -Message $msg 

                                $exportEndpointName = "SRC-SPO-Root-RootSubsite$rootSubsiteDepth-$rootSubsiteUrl-$script:sourceTenantName"
                                $exportType = "SharePoint"
                                $exportTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                                $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
                                    "Url"                          = $srcRootSubsiteUrl;
                                    "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                                    "AdministrativePassword"       = $script:sourcePlainPassword;
                                    "UseAdministrativeCredentials" = $true
                                }

                                $importEndpointName = "DST-SPO-Root-RootSubsite$rootSubsiteDepth-$rootSubsiteUrl-$script:destinationTenantName"
                                $importType = "SharePointOnlineAPI"
                                $importTypeName = "MigrationProxy.WebApi.SharePointOnlineConfiguration"
                                if (!$UseOwnAzureStorage) {
                                    $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                        "Url"                                = $dstRootSubsiteUrl;
                                        "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                        "AdministrativePassword"             = $script:destinationPlainPassword;
                                        "UseAdministrativeCredentials"       = $true;
                                        "UseSharePointOnlineProvidedStorage" = $true 
                                    }
                                }
                                else {
                                    $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                        "Url"                                = $dstRootSubsiteUrl;
                                        "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                        "AdministrativePassword"             = $script:destinationPlainPassword;
                                        "UseAdministrativeCredentials"       = $true;
                                        "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                        "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;
                                        "UseSharePointOnlineProvidedStorage" = $false
                                    }
                                }

                                #Create SPO Team Sites source endpoint
                                $exportEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType $exportType  -EndpointName $exportEndpointName -EndpointConfiguration $exportConfiguration -update $updateEndpoint
                                #Create SPO Team Sites destination endpoint
                                $importEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType $importType -EndpointName $importEndpointName -EndpointConfiguration $importConfiguration -update $updateEndpoint

                                #Create SPO Team Sites Document project

                                $ProjectName = "ClassicSPOSite-Document-Root-RootSubsite$rootSubsiteDepth-$rootSubsiteUrl-$script:sourceTenantName"
                                $ProjectType = "Storage"

                                if ($enableModernAuth) {
                                    $advancedOptions = "$modernAuth InitializationTimeout=8 IncreasePathLengthLimit=1"
                                }
                                else {
                                    $advancedOptions = "InitializationTimeout=8 IncreasePathLengthLimit=1 "
                                }

                                if ($applicationPermissions) {
                                    $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                                }
                                if ($UseDelegatePermission) {
                                    $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                                }

                                $connectorId = $null
                                $connectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                                    -ProjectName $ProjectName `
                                    -ProjectType $ProjectType `
                                    -exportType $exportType `
                                    -importType $importType `
                                    -exportEndpointId $exportEndpointId `
                                    -importEndpointId $importEndpointId `
                                    -exportConfiguration $exportConfiguration `
                                    -importConfiguration $importConfiguration `
                                    -advancedOptions $advancedOptions `
                                    -maximumSimultaneousMigrations 100 `
                                    -ZoneRequirement $global:btZoneRequirement `
                                    -MaxLicensesToConsume 1 `
                                    -updateConnector $true   

                            }    

                            if ($connectorId -ne $null) {

                                $documentLibraryName = $rootSubsiteDocumentLibrary.DocumentLibraryName

                                if ([string]::IsNullOrEmpty($documentLibraryName)) { Continue }

                                $msg = "      INFO: Processing Document Library '$documentLibraryName'."
                                Write-Host $msg
                                Log-Write -Message $msg 

                                if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                    # Renew MW ticket
                                    Connect-BitTitan 

                                    Write-Host
                                    $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                    Write-Host -ForegroundColor Magenta $msg
                                    Log-Write -Message $msg 
                                    Write-Host                    
                                }   

                                try {
                                    $ImportLibrary = $documentLibraryName
                                    $ExportLibrary = $documentLibraryName

                                    $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
                                    if (!$result) {
                                        $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId  -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary

                                        $msg = "      SUCCESS: Document Library migration '$ExportLibrary->$ImportLibrary' added to connector." 
                                        write-Host -ForegroundColor Green $msg
                                        Log-Write -Message $msg  

                                        $processedRootSubsitesDocumentLibraries += 1
                                    }
                                    else {
                                        $msg = "      WARNING: Document Library migration '$ExportLibrary->$ImportLibrary' already exists in connector." 
                                        write-Host -ForegroundColor Yellow $msg
                                        Log-Write -Message $msg  

                                        $processedRootSubsitesDocumentLibraries += 1
                                    }
                                }
                                catch {
                                    $msg = "      ERROR: Failed to add Document Library migration '$ExportLibrary->$ImportLibrary' to connector." 
                                    write-Host -ForegroundColor Red $msg
                                    Log-Write -Message $msg     
                                }

                                [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $ProjectName; ProjectType = $ProjectType; ConnectorId = $connectorId; }
                            }



                        }
                    }

                    ##################################################################################################
                    #    SITES
                    ##################################################################################################

                    $classicTeamSitesIndex = 1
      
                    foreach ($classicTeamSite in $script:importedClassicTeamSites) {

                        $url = $classicTeamSite.url
                        $srcClassicTeamSiteUrl = $url -replace $script:destinationTenantName, $script:sourceTenantName
                        if ($classicTeamSite.url -eq $classicTeamSite.NewDestinationUrl) {
                            $dstClassicTeamSiteUrl = $url -replace $script:sourceTenantName, $script:destinationTenantName
                        }
                        else {
                            $dstClassicTeamSiteUrl = $classicTeamSite.NewDestinationUrl
                        }
        
                        $classicTeamSiteName = $classicTeamSite.Title

                        $documentLibraries = @()
                        $documentLibraries = @($script:importedDocumentLibraries | Where-Object { $_.TeamSiteUrl -eq $classicTeamSite.url } | Where-Object { $_.isSubsite -eq $false })
            
                        $subsiteDocumentLibraries = @($script:importedDocumentLibraries | Where-Object { $_.isSubsite -eq $true } | Where-Object { $_.TeamSiteUrl -match $classicTeamSite.url })

                        write-host 
                        $msg = "INFO: Processing classic Team Site $classicTeamSitesIndex/$script:classicTeamSitesCount '$classicTeamSiteName' with URL '$dstClassicTeamSiteUrl'."
                        Write-Host $msg
                        Log-Write -Message $msg 
                        $classicTeamSitesIndex += 1
    
                        $doNotSkipMailbox = $false

                        if ($classicTeamSiteName -eq "" -or $classicTeamSiteSrcUrl -eq "" -or $classicTeamSiteDstUrl -eq "") {
                            $msg = "INFO: Skipping SharePoint  classicTeamSite '$classicTeamSiteName'. Missing data in the CSV file."
                            Write-Host -ForegroundColor Red $msg
                            Log-Write -Message $msg 

                            Continue
                        }    

                        if ($checkDestinationSPOStructure) { $result = check-DestinationSPOSite -url $dstClassicTeamSiteUrl -AdminCenterUrl "https://$script:destinationTenantName-admin.sharepoint.com/" }
                        if (!$result -and $checkDestinationSpoStructure) {
                            $msg = "INFO: Skipping classic Team Site project creation. classic Team Site '$classicTeamSiteName' with URL '$url' not found in destination SharePoint Online."
                            Write-Host -ForegroundColor Red $msg
                            Log-Write -Message $msg 

                            Continue
                        } 
                    
                        #Create SPO endpoints
                        $msg = "INFO: Creating MSPC endpoints for classic Team Site '$classicTeamSiteName'."
                        Write-Host $msg
                        Log-Write -Message $msg 

                        $exportEndpointName = "SRC-SPO-$classicTeamSiteName-$script:sourceTenantName"
                        $exportType = "SharePoint"
                        $exportTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                        $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
                            "Url"                          = $srcClassicTeamSiteUrl;
                            "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                            "AdministrativePassword"       = $script:sourcePlainPassword;
                            "UseAdministrativeCredentials" = $true
                        }
        
                        $importEndpointName = "DST-SPO-$classicTeamSiteName-$script:destinationTenantName"
                        $importType = "SharePointOnlineAPI"
                        $importTypeName = "MigrationProxy.WebApi.SharePointOnlineConfiguration"
                        if (!$UseOwnAzureStorage) {
                            $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                "Url"                                = $dstClassicTeamSiteUrl;
                                "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                "AdministrativePassword"             = $script:destinationPlainPassword;
                                "UseAdministrativeCredentials"       = $true;
                                "UseSharePointOnlineProvidedStorage" = $true 
                            }
                        }
                        else {
                            $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                "Url"                                = $dstClassicTeamSiteUrl;
                                "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                "AdministrativePassword"             = $script:destinationPlainPassword;
                                "UseAdministrativeCredentials"       = $true;
                                "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;    
                                "UseSharePointOnlineProvidedStorage" = $false
                            }
                        }

                        #Create SPO Team Sites source endpoint
                        $exportEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType $exportType  -EndpointName $exportEndpointName -EndpointConfiguration $exportConfiguration -update $updateEndpoint
                        #Create SPO Team Sites destination endpoint
                        $importEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType $importType -EndpointName $importEndpointName -EndpointConfiguration $importConfiguration -update $updateEndpoint

                        #Create SPO Team Sites Document project

                        $ProjectName = "ClassicSPOSite-Document-$classicTeamSiteName-$script:sourceTenantName"
                        $ProjectType = "Storage"

                        if ($enableModernAuth) {
                            $advancedOptions = "$modernAuth InitializationTimeout=8 IncreasePathLengthLimit=1"
                        }
                        else {
                            $advancedOptions = "InitializationTimeout=8 IncreasePathLengthLimit=1 "
                        }

                        if ($applicationPermissions) {
                            $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                        }
                        if ($UseDelegatePermission) {
                            $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                        }

                        $connectorId = $null
                        $connectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                            -ProjectName $ProjectName `
                            -ProjectType $ProjectType `
                            -exportType $exportType `
                            -importType $importType `
                            -exportEndpointId $exportEndpointId `
                            -importEndpointId $importEndpointId `
                            -exportConfiguration $exportConfiguration `
                            -importConfiguration $importConfiguration `
                            -advancedOptions $advancedOptions `
                            -maximumSimultaneousMigrations 100 `
                            -ZoneRequirement $global:btZoneRequirement `
                            -MaxLicensesToConsume 1 `
                            -updateConnector $true   

                        if ($connectorId -ne $null) {
                        
                            foreach ($documentLibrary in $documentLibraries) {

                                $documentLibraryName = $documentLibrary.DocumentLibraryName

                                if ([string]::IsNullOrEmpty($documentLibraryName)) { Continue }

                                $msg = "      INFO: Processing Document Library '$documentLibraryName'."
                                Write-Host $msg
                                Log-Write -Message $msg 

                                if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                    # Renew MW ticket
                                    Connect-BitTitan 

                                    Write-Host
                                    $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                    Write-Host -ForegroundColor Magenta $msg
                                    Log-Write -Message $msg 
                                    Write-Host                    
                                }   

                                try {
                                    $ImportLibrary = $documentLibraryName
                                    $ExportLibrary = $documentLibraryName

                                    $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
                                    if (!$result) {
                                        $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId  -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary

                                        $msg = "      SUCCESS: Document Library migration '$ExportLibrary->$ImportLibrary' added to connector." 
                                        write-Host -ForegroundColor Green $msg
                                        Log-Write -Message $msg  

                                        $processedDocumentLibraries += 1
                                    }
                                    else {
                                        $msg = "      WARNING: Document Library migration '$ExportLibrary->$ImportLibrary' already exists in connector." 
                                        write-Host -ForegroundColor Yellow $msg
                                        Log-Write -Message $msg  

                                        $processedDocumentLibraries += 1
                                    }
                                }
                                catch {
                                    $msg = "      ERROR: Failed to add Document Library migration '$ExportLibrary->$ImportLibrary' to connector." 
                                    write-Host -ForegroundColor Red $msg
                                    Log-Write -Message $msg     
                                }
                            }

                            $processedClassicTeamSites += 1   

                            [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $ProjectName; ProjectType = $ProjectType; ConnectorId = $connectorId; }
                        }

                        ##################################################################################################
                        #    SUBSITES
                        ##################################################################################################
                        $currentSubsiteDocumentLibrary = 0
                        $subsiteDocumentLibraryCount = $subsiteDocumentLibraries.Count
                        $totalSubsitesCount += $subsiteDocumentLibraryCount

                        $previousSubSiteName = ""

                        if ($subsiteDocumentLibraries -ne $null) {
                            foreach ($subsiteDocumentLibrary in $subsiteDocumentLibraries  ) {

                                if ($subsiteDocumentLibrary.SubsiteNumber.SubString(0, 1) -eq 1) {
                                    if (!$alreadyProcessed) {
                                        $ProcessedSubsite = 0
                                        $alreadyProcessed = $true
                                    } 
                                }
                                else {
                                    $alreadyProcessed = $false
                                }

                                if ($ProcessedSubsite -ne $subsiteDocumentLibrary.SubsiteNumber.SubString(0, 1)) {
                                    $ProcessedSubsite += 1 
                                    $allProcessedSubsites += 1 
                                }  
                    
                                $currentSubsiteDocumentLibrary += 1

                                $urlparts = $($subSiteDocumentLibrary.TeamSiteUrl).split("/")
                                $subsiteDepth = $urlparts.Count - 5
                                $subSiteName = $urlparts[-1]
                                $parentSubSiteUrl = ($subSiteDocumentLibrary.TeamSiteUrl -split "/$subSiteName")[0] 
                                $subsiteUrl = $parentSubSiteUrl + "/" + $subSiteName

                                if ($subSiteName -ne $previousSubSiteName) {
                                    $currentSubsite += 1

                                    $srcSubsiteUrl = $subsiteUrl -replace $script:destinationTenantName, $script:sourceTenantName
                                    $dstSubsiteUrl = $subsiteUrl -replace $script:sourceTenantName, $script:destinationTenantName

                                    if ($subsiteDepth -eq 1) {  
                                        Write-Host 
                                        write-host 
                                        $msg = "INFO: Processing $currentSubsite/$script:subsitesCount Subsite '$subSiteName' level-$subsiteDepth of Team Site '$classicTeamSiteName' with URL '$srcSubsiteUrl'."
                                        Write-Host $msg
                                        Log-Write -Message $msg 
                                    }
                                    elseif ($subsiteDepth -gt 1) {
                                        Write-Host
                                        write-host 
                                        $msg = "INFO: Processing $currentSubsite/$script:subsitesCount Subsite '$subSiteName' level-$subsiteDepth  of SubSite '$parentSubSiteUrl' with URL '$srcSubsiteUrl'."
                                        Write-Host $msg
                                        Log-Write -Message $msg 
                                    }

                                    if ($checkDestinationSPOStructure) { $result = check-DestinationSPOSubsite -siteUrl $dstSubsiteUrl -SubSiteName $subSiteName }
                                    if (!$result -and $checkDestinationSpoStructure) {
                                        $msg = "INFO: Skipping classic Team Subsite project creation. Subsite '$subSiteName' with URL '$dstSubsiteUrl' not found in destination SharePoint Online."
                                        Write-Host -ForegroundColor Red $msg
                                        Log-Write -Message $msg 

                                        Continue
                                    }   

                                    #Create SPO endpoints
                                    $msg = "INFO: Creating MSPC endpoints for classic Team Site '$classicTeamSiteName'."
                                    Write-Host $msg
                                    Log-Write -Message $msg 

                                    $subsiteRelativeURL = $subsiteUrl.Replace($srcClassicTeamSiteUrl + "/", "").Replace($dstClassicTeamSiteUrl + "/", "")

                                    $exportEndpointName = "SRC-SPO-$classicTeamSiteName-Subsite$subsiteDepth-$subsiteRelativeURL-$script:sourceTenantName"
                                    $exportType = "SharePoint"
                                    $exportTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                                    $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
                                        "Url"                          = $srcSubsiteUrl;
                                        "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                                        "AdministrativePassword"       = $script:sourcePlainPassword;
                                        "UseAdministrativeCredentials" = $true
                                    }

                                    $importEndpointName = "DST-SPO-$classicTeamSiteName-Subsite$subsiteDepth-$subsiteRelativeURL-$script:destinationTenantName"
                                    $importType = "SharePointOnlineAPI"
                                    $importTypeName = "MigrationProxy.WebApi.SharePointOnlineConfiguration"
                                    if (!$UseOwnAzureStorage) {
                                        $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                            "Url"                                = $dstSubsiteUrl;
                                            "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                            "AdministrativePassword"             = $script:destinationPlainPassword;
                                            "UseAdministrativeCredentials"       = $true;
                                            "UseSharePointOnlineProvidedStorage" = $true 
                                        }
                                    }
                                    else {
                                        $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                            "Url"                                = $dstSubsiteUrl;
                                            "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                            "AdministrativePassword"             = $script:destinationPlainPassword;
                                            "UseAdministrativeCredentials"       = $true;
                                            "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                            "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;    
                                            "UseSharePointOnlineProvidedStorage" = $false
                                        }
                                    }

                                    #Create SPO Team Sites source endpoint
                                    $exportEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType $exportType  -EndpointName $exportEndpointName -EndpointConfiguration $exportConfiguration -update $updateEndpoint
                                    #Create SPO Team Sites destination endpoint
                                    $importEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType $importType -EndpointName $importEndpointName -EndpointConfiguration $importConfiguration -update $updateEndpoint

                                    #Create SPO Team Sites Document project

                                    $ProjectName = "ClassicSPOSite-Document-$classicTeamSiteName-Subsite$subsiteDepth-$subsiteRelativeURL-$script:sourceTenantName"
                                    $ProjectType = "Storage"

                                    if ($enableModernAuth) {
                                        $advancedOptions = "$modernAuth InitializationTimeout=8 IncreasePathLengthLimit=1"
                                    }
                                    else {
                                        $advancedOptions = "InitializationTimeout=8 IncreasePathLengthLimit=1 "
                                    }

                                    if ($applicationPermissions) {
                                        $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                                    }
                                    if ($UseDelegatePermission) {
                                        $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                                    }

                                    $connectorId = $null
                                    $connectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                                        -ProjectName $ProjectName `
                                        -ProjectType $ProjectType `
                                        -exportType $exportType `
                                        -importType $importType `
                                        -exportEndpointId $exportEndpointId `
                                        -importEndpointId $importEndpointId `
                                        -exportConfiguration $exportConfiguration `
                                        -importConfiguration $importConfiguration `
                                        -advancedOptions $advancedOptions `
                                        -maximumSimultaneousMigrations 100 `
                                        -ZoneRequirement $global:btZoneRequirement `
                                        -MaxLicensesToConsume 1 `
                                        -updateConnector $true   

                                }    

                                if ($connectorId -ne $null) {
            
                                    $documentLibraryName = $subsiteDocumentLibrary.DocumentLibraryName

                                    if ([string]::IsNullOrEmpty($documentLibraryName)) { Continue }

                                    $msg = "      INFO: Processing Document Library '$documentLibraryName'."
                                    Write-Host $msg
                                    Log-Write -Message $msg 

                                    if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                        # Renew MW ticket
                                        Connect-BitTitan 

                                        Write-Host
                                        $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                        Write-Host -ForegroundColor Magenta $msg
                                        Log-Write -Message $msg 
                                        Write-Host                    
                                    }   

                                    try {
                                        $ImportLibrary = $documentLibraryName
                                        $ExportLibrary = $documentLibraryName

                                        $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
                                        if (!$result) {
                                            $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId  -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary

                                            $msg = "      SUCCESS: Document Library migration '$ExportLibrary->$ImportLibrary' added to connector." 
                                            write-Host -ForegroundColor Green $msg
                                            Log-Write -Message $msg  

                                            $processedSubsiteDocumentLibraries += 1
                                        }
                                        else {
                                            $msg = "      WARNING: Document Library migration '$ExportLibrary->$ImportLibrary' already exists in connector." 
                                            write-Host -ForegroundColor Yellow $msg
                                            Log-Write -Message $msg  

                                            $processedSubsiteDocumentLibraries += 1
                                        }
                                    }
                                    catch {
                                        $msg = "      ERROR: Failed to add Document Library migration '$ExportLibrary->$ImportLibrary' to connector." 
                                        write-Host -ForegroundColor Red $msg
                                        Log-Write -Message $msg     
                                    }

                                    [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $ProjectName; ProjectType = $ProjectType; ConnectorId = $connectorId; }
                                }
                            }
                        }
                    }

                    if ($processedClassicTeamSites -ne 0) {
                        write-Host
                        $msg = "SUCCESS: $processedRootDocumentLibraries out of $script:rootDocumentLibrariesCount Root Document Libraries have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                        $msg = "SUCCESS: $currentRootSubsite out of $script:rootSubsitesCount Root Subsites have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                        $msg = "SUCCESS: $processedRootSubsitesDocumentLibraries out of $script:rootSubSiteDocumentLibrariesCount Root Subsites Document Libraries have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 

                        write-Host
            
                        $msg = "SUCCESS: $processedClassicTeamSites out of $script:classicTeamSitesCount classic Team Sites have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                        $msg = "SUCCESS: $processedDocumentLibraries out of $script:documentLibrariesCount Document Libraries have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                        $msg = "SUCCESS: $allProcessedSubsites out of $script:subsitesCount classic Team Subsites have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg             
                        if ([string]::IsNullOrEmpty(($processedSubsites))) { $processedSubsiteDocumentLibraries = 0 }
                        $msg = "SUCCESS: $processedSubsiteDocumentLibraries out of $script:subsiteDocumentLibrariesCount classic Team Subsites Document Libraries have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                    }

                }  
            }
            elseif ($migrateNewSpoTeamSites) {
                write-host 
                $msg = "#######################################################################################################################`
                  CREATING CLASSIC TEAM SITE PROJECTS                  `
#######################################################################################################################"
                Write-Host $msg
                Log-Write -Message "CREATING CLASSIC TEAM SITE PROJECTS" 
                Write-Host
    
                $msg = "INFO: Processing Classic Team site migration."
                Write-Host $msg
                Log-Write -Message $msg 

                if ($script:importedClassicTeamSites -eq $null -or $script:importedDocumentLibraries -eq $null) {
                    if ($script:importedClassicTeamSites -eq $null) {
                        $msg = "INFO: No SPO Team Sites found. Skipping SPO Team Site project creation."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg 
                    }
                    if ($script:importedDocumentLibraries -eq $null) {
                        $msg = "INFO: No Root Document Libraries or Root Subsites and their Document Libraries found. Skipping SPO Team Site project creation."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg 
                    }
                }
                else {
                    if (!$UseOwnAzureStorage -and !$AzureStorageSelected) {
                        Write-Host   
                        do {
                            $confirm = (Read-Host -prompt "Do you want to use Microsoft provided Azure Storage?  [Y]es or [N]o")
                            if ($confirm.ToLower() -eq "n") {
                                $UseOwnAzureStorage = $true

                                $AzureStorageAccountName = ''
                                $AzureAccountKey = ''
                
                                do {
                                    $AzureStorageAccountName = (Read-Host -prompt "         Enter Azure Storage Account Name")
                                } while ($AzureStorageAccountName -eq "")

                                do {
                                    $AzureAccountKey = (Read-Host -prompt "         Enter Azure Storage Primary Access Key")
                                } while ($AzureAccountKey -eq "")

                                $AzureStorageSelected = $true
                            }
                            else {
                                $UseOwnAzureStorage = $false
                                $AzureStorageSelected = $true
                            }
                        } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
                    }

                    $alreadyCreated = $false

                    $subsiteDocumentLibraries = @($script:importedDocumentLibraries | Where-Object { $_.isSubsite -eq $true })
               
                    $processedRootDocumentLibraries = 0
                    $processedRootSubsites = 0
                    $processedRootSubsitesDocumentLibraries = 0
                    $processedClassicTeamSites = 0
                    $processedDocumentLibraries = 0
                    $processedSubsites = 0
                    $processedSubsiteDocumentLibraries = 0

                    $currentSubsite = 0

                    ##################################################################################################
                    #    ROOT DOCUMENT LIBRARIES
                    ##################################################################################################
                    $currentRootDocumentLibrary = 0
        
                    if ($script:importedRootDocumentLibraries -ne $null) {
                        foreach ($rootDocumentLibrary in $script:importedRootDocumentLibraries) {

                            $currentRootDocumentLibrary += 1

                            $urlparts = $($rootDocumentLibrary.DocumentLibraryUrl).split("/")
                            $rootDocumentLibraryName = $urlparts[-1]

                            if ($rootDocumentLibrary.TeamSiteURL -eq $rootDocumentLibrary.NewDestinationTeamSiteURL) {    
                                $sSPOUrl = $rootDocumentLibrary.TeamSiteUrl -replace $script:destinationTenantName, $script:sourceTenantName
                                $dSPOUrl = $sSPOUrl -replace $script:sourceTenantName, $script:destinationTenantName
                            }
                            else {
                                $sSPOUrl = $rootDocumentLibrary.TeamSiteUrl -replace $script:destinationTenantName, $script:sourceTenantName
                                $dSPOUrl = $rootDocumentLibrary.NewDestinationTeamSiteURL -replace $script:sourceTenantName, $script:destinationTenantName  
                            }

                
                            write-host 
                            $msg = "INFO: Processing $currentRootDocumentLibrary/$script:rootDocumentLibrariesCount Root Document Library '$rootDocumentLibraryName' of Root Site '$sSPOUrl'."
                            Write-Host $msg
                            Log-Write -Message $msg 
                
                            if ($null -eq $connectorId) {
                                #Create SPO endpoints
                                $msg = "INFO: Creating MSPC endpoints for Root Document Library '$sSPOUrl'."
                                Write-Host $msg
                                Log-Write -Message $msg 
                    
                                $exportEndpointName = "SRC-SPO-$script:sourceTenantName"
                                $exportType = "SharePoint"
                                $exportTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                                $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
                                    "Url"                          = $sSPOUrl;
                                    "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                                    "AdministrativePassword"       = $script:sourcePlainPassword;
                                    "UseAdministrativeCredentials" = $true
                                }
                    
                                $importEndpointName = "DST-SPO-$script:destinationTenantName"
                                $importType = "SharePointBeta"
                                $importTypeName = "MigrationProxy.WebApi.SharePointBetaConfiguration"
                                if (!$UseOwnAzureStorage) {
                                    $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                        "Url"                                = $dSPOUrl;
                                        "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                        "AdministrativePassword"             = $script:destinationPlainPassword;
                                        "UseAdministrativeCredentials"       = $true;
                                        "UseSharePointOnlineProvidedStorage" = $true 
                                    }
                                }
                                else {
                                    $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                        "Url"                                = $dSPOUrl;
                                        "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                        "AdministrativePassword"             = $script:destinationPlainPassword;
                                        "UseAdministrativeCredentials"       = $true;
                                        "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                        "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;    
                                        "UseSharePointOnlineProvidedStorage" = $false
                                    }
                                }
                    
                                #Create SPO Team Sites source endpoint
                                $exportEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType $exportType  -EndpointName $exportEndpointName -EndpointConfiguration $exportConfiguration -update $updateEndpoint
                                #Create SPO Team Sites destination endpoint
                                $importEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType $importType -EndpointName $importEndpointName -EndpointConfiguration $importConfiguration -update $updateEndpoint
                    
                                #Create SPO Team Sites Document project
                    
                                $ProjectName = "ClassicSPOSite-$script:sourceTenantName"
                                $ProjectType = "Storage"
                    
                                if ($enableModernAuth) {
                                    $advancedOptions = "$modernAuth InitializationTimeout=8 IncreasePathLengthLimit=1"
                                }
                                else {
                                    $advancedOptions = "InitializationTimeout=8 IncreasePathLengthLimit=1"
                                }
                    
                                if ($applicationPermissions) {
                                    $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                                }
                                if ($UseDelegatePermission) {
                                    $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                                }
                    
                                $connectorId = $null
                                $connectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                                    -ProjectName $ProjectName `
                                    -ProjectType $ProjectType `
                                    -exportType $exportType `
                                    -importType $importType `
                                    -exportEndpointId $exportEndpointId `
                                    -importEndpointId $importEndpointId `
                                    -exportConfiguration $exportConfiguration `
                                    -importConfiguration $importConfiguration `
                                    -advancedOptions $advancedOptions `
                                    -maximumSimultaneousMigrations 100 `
                                    -ZoneRequirement $global:btZoneRequirement `
                                    -MaxLicensesToConsume 1 `
                                    -updateConnector $true                        
                            }

                            if ($connectorId -ne $null) {

                                $documentLibraryName = $rootDocumentLibrary.DocumentLibraryName

                                if ([string]::IsNullOrEmpty($documentLibraryName)) { Continue }

                                $msg = "      INFO: Processing Document Library '$documentLibraryName'."
                                Write-Host $msg
                                Log-Write -Message $msg 

                                if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                    # Renew MW ticket
                                    Connect-BitTitan 

                                    Write-Host
                                    $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                    Write-Host -ForegroundColor Magenta $msg
                                    Log-Write -Message $msg 
                                    Write-Host                    
                                }   

                                try {
                                    $ImportLibrary = $documentLibraryName
                                    $ExportLibrary = $documentLibraryName

                                    $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
                                    if (!$result) {
                                        $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId  -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary

                                        $msg = "      SUCCESS: Document Library migration '$ExportLibrary->$ImportLibrary' added to connector." 
                                        write-Host -ForegroundColor Green $msg
                                        Log-Write -Message $msg  

                                        $processedRootDocumentLibraries += 1
                                    }
                                    else {
                                        $msg = "      WARNING: Document Library migration '$ExportLibrary->$ImportLibrary' already exists in connector." 
                                        write-Host -ForegroundColor Yellow $msg
                                        Log-Write -Message $msg  

                                        $processedRootDocumentLibraries += 1
                                    }
                                }
                                catch {
                                    $msg = "      ERROR: Failed to add Document Library migration '$ExportLibrary->$ImportLibrary' to connector." 
                                    write-Host -ForegroundColor Red $msg
                                    Log-Write -Message $msg     
                                }

                                [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $ProjectName; ProjectType = $ProjectType; ConnectorId = $connectorId; }
                            }
                        }
                    }

                    ##################################################################################################
                    #    ROOT SUBSITES
                    ##################################################################################################
                    $currentRootSubsiteDocumentLibrary = 0
                    $subsiteRootDocumentLibraryCount = $subsiteDocumentLibraries.Count
                    $totalSubsitesCount += $subsiteRootDocumentLibraryCount

                    $previousSubSiteName = ""
        
                    if ($script:importedRootSubSiteDocumentLibraries -ne $null) {
                        foreach ($rootSubsiteDocumentLibrary in $script:importedRootSubSiteDocumentLibraries) {

                            if ($ProcessedRootSubsites -ne $rootSubsiteDocumentLibrary.SubsiteNumber.SubString(0, 1)) {
                                $ProcessedRootSubsites += 1 
                            }  
                
                            $currentRootSubsiteDocumentLibrary += 1

                            $urlparts = $($rootSubsiteDocumentLibrary.TeamSiteUrl).split("/")
                            $rootSubsiteDepth = $urlparts.Count - 4
                            $rootSubsiteName = $urlparts[-1]
                            $parentRootSubsiteUrl = ($rootSubsiteDocumentLibrary.TeamSiteUrl -split "/$rootSubsiteName")[0] 
                            $rootSubsiteUrl = $parentRootSubsiteUrl + "/" + $rootSubsiteName

                            if ($rootDocumentLibrary.TeamSiteURL -ne $rootDocumentLibrary.NewDestinationTeamSiteURL) { 
                                $dstUrlparts = $($rootSubsiteDocumentLibrary.NewDestinationTeamSiteURL).split("/")
                                $dstRootSubsiteName = $dstUrlparts[-1]
                                $dstParentRootSubsiteUrl = ($rootSubsiteDocumentLibrary.NewDestinationTeamSiteURL -split "/$dstRootSubsiteName")[0] 
                                $dstRootSubsiteUrl = $dstParentRootSubsiteUrl + "/" + $dstRootSubsiteName
                            }               
                
                            if ($rootSubsiteName -ne $previousRootSubsiteName) {
                                $currentRootSubsite += 1

                                if ($rootDocumentLibrary.TeamSiteURL -eq $rootDocumentLibrary.NewDestinationTeamSiteURL) { 
                                    $srcRootSubsiteUrl = $rootSubsiteUrl -replace $script:destinationTenantName, $script:sourceTenantName
                                    $dstRootSubsiteUrl = $rootSubsiteUrl -replace $script:sourceTenantName, $script:destinationTenantName
                                }
                                else {
                                    $srcRootSubsiteUrl = $rootSubsiteUrl -replace $script:destinationTenantName, $script:sourceTenantName
                                    $dstRootSubsiteUrl = $dstRootSubsiteUrl -replace $script:sourceTenantName, $script:destinationTenantName    
                                }

                                if ($rootSubsiteDepth -eq 0) {  
                                    write-host 
                                    $msg = "INFO: Processing $currentRootSubsite/$script:rootSubsitesCount Root Subsite '$rootSubsiteName' level-$rootSubsiteDepth of Root Site '$sSPOUrl' with URL '$srcRootSubsiteUrl'."
                                    Write-Host $msg
                                    Log-Write -Message $msg 
                                }
                                elseif ($rootSubsiteDepth -gt 0) {
                                    write-host 
                                    $msg = "INFO: Processing $currentRootSubsite/$script:rootSubsitesCount Root Subsite '$rootSubsiteName' level-$rootSubsiteDepth  of Root Subsite '$parentRootSubsiteUrl' with URL '$srcRootSubsiteUrl'."
                                    Write-Host $msg
                                    Log-Write -Message $msg 
                                }

                                if ($checkDestinationSPOStructure) { $result = check-DestinationSPOSubsite -siteUrl $dstRootSubsiteUrl -SubSiteName $rootSubsiteName }
                                if (!$result -and $checkDestinationSpoStructure) {
                                    $msg = "INFO: Skipping classic Team Root Subsite project creation. Root Subsite '$rootSubsiteName' with URL '$dstRootSubsiteUrl' not found in destination SharePoint Online."
                                    Write-Host -ForegroundColor Red $msg
                                    Log-Write -Message $msg 

                                    Continue
                                }   

                                if ($null -eq $connectorId) {
                                    #Create SPO endpoints
                                    $msg = "INFO: Creating MSPC endpoints for Root Site '$sSPOUrl'."
                                    Write-Host $msg
                                    Log-Write -Message $msg 

                                    $exportEndpointName = "SRC-SPO-$script:sourceTenantName"
                                    $exportType = "SharePoint"
                                    $exportTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                                    $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
                                        "Url"                          = $sSPOUrl;
                                        "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                                        "AdministrativePassword"       = $script:sourcePlainPassword;
                                        "UseAdministrativeCredentials" = $true
                                    }

                                    $importEndpointName = "DST-SPO-$script:destinationTenantName"
                                    $importType = "SharePointBeta"
                                    $importTypeName = "MigrationProxy.WebApi.SharePointBetaConfiguration"
                                    if (!$UseOwnAzureStorage) {
                                        $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                            "Url"                                = $dSPOUrl;
                                            "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                            "AdministrativePassword"             = $script:destinationPlainPassword;
                                            "UseAdministrativeCredentials"       = $true;
                                            "UseSharePointOnlineProvidedStorage" = $true 
                                        }
                                    }
                                    else {
                                        $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                            "Url"                                = $dSPOUrl;
                                            "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                            "AdministrativePassword"             = $script:destinationPlainPassword;
                                            "UseAdministrativeCredentials"       = $true;
                                            "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                            "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;
                                            "UseSharePointOnlineProvidedStorage" = $false
                                        }
                                    }

                                    #Create SPO Team Sites source endpoint
                                    $exportEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType $exportType  -EndpointName $exportEndpointName -EndpointConfiguration $exportConfiguration -update $updateEndpoint
                                    #Create SPO Team Sites destination endpoint
                                    $importEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType $importType -EndpointName $importEndpointName -EndpointConfiguration $importConfiguration -update $updateEndpoint

                                    #Create SPO Team Sites Document project

                                    $ProjectName = "ClassicSPOSite-$script:sourceTenantName"
                                    $ProjectType = "Storage"

                                    if ($enableModernAuth) {
                                        $advancedOptions = "$modernAuth InitializationTimeout=8 IncreasePathLengthLimit=1"
                                    }
                                    else {
                                        $advancedOptions = "InitializationTimeout=8 IncreasePathLengthLimit=1 "
                                    }

                                    if ($applicationPermissions) {
                                        $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                                    }
                                    if ($UseDelegatePermission) {
                                        $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                                    }

                                    $connectorId = $null
                                    $connectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                                        -ProjectName $ProjectName `
                                        -ProjectType $ProjectType `
                                        -exportType $exportType `
                                        -importType $importType `
                                        -exportEndpointId $exportEndpointId `
                                        -importEndpointId $importEndpointId `
                                        -exportConfiguration $exportConfiguration `
                                        -importConfiguration $importConfiguration `
                                        -advancedOptions $advancedOptions `
                                        -maximumSimultaneousMigrations 100 `
                                        -ZoneRequirement $global:btZoneRequirement `
                                        -MaxLicensesToConsume 1 `
                                        -updateConnector $true   

                                }  
                            }  

                            if ($connectorId -ne $null) {

                                $documentLibraryName = $rootSubsiteDocumentLibrary.DocumentLibraryName

                                if ([string]::IsNullOrEmpty($documentLibraryName)) { Continue }

                                $msg = "      INFO: Processing Document Library '$documentLibraryName'."
                                Write-Host $msg
                                Log-Write -Message $msg 

                                if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                    # Renew MW ticket
                                    Connect-BitTitan 

                                    Write-Host
                                    $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                    Write-Host -ForegroundColor Magenta $msg
                                    Log-Write -Message $msg 
                                    Write-Host                    
                                }   

                                try {
                                    $relativeUrl = $rootSubsiteUrl -replace $sSPOUrl
                                    $ImportLibrary = $relativeUrl + "/" + $documentLibraryName
                                    $ExportLibrary = $relativeUrl + "/" + $documentLibraryName

                                    $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
                                    if (!$result) {
                                        $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId  -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary

                                        $msg = "      SUCCESS: Document Library migration '$ExportLibrary->$ImportLibrary' added to connector." 
                                        write-Host -ForegroundColor Green $msg
                                        Log-Write -Message $msg  

                                        $processedRootSubsitesDocumentLibraries += 1
                                    }
                                    else {
                                        $msg = "      WARNING: Document Library migration '$ExportLibrary->$ImportLibrary' already exists in connector." 
                                        write-Host -ForegroundColor Yellow $msg
                                        Log-Write -Message $msg  

                                        $processedRootSubsitesDocumentLibraries += 1
                                    }
                                }
                                catch {
                                    $msg = "      ERROR: Failed to add Document Library migration '$ExportLibrary->$ImportLibrary' to connector." 
                                    write-Host -ForegroundColor Red $msg
                                    Log-Write -Message $msg     
                                }

                                [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $ProjectName; ProjectType = $ProjectType; ConnectorId = $connectorId; }
                            }
                        }
                    }

                    ##################################################################################################
                    #    SITES
                    ##################################################################################################

                    $classicTeamSitesIndex = 1
      
                    foreach ($classicTeamSite in $script:importedClassicTeamSites) {

                        $url = $classicTeamSite.url
                        $srcClassicTeamSiteUrl = $url -replace $script:destinationTenantName, $script:sourceTenantName
                        if ($classicTeamSite.url -eq $classicTeamSite.NewDestinationUrl) {
                            $dstClassicTeamSiteUrl = $url -replace $script:sourceTenantName, $script:destinationTenantName
                        }
                        else {
                            $dstClassicTeamSiteUrl = $classicTeamSite.NewDestinationUrl
                        }
        
                        $classicTeamSiteName = $classicTeamSite.Title

                        $documentLibraries = @()
                        $documentLibraries = @($script:importedDocumentLibraries | Where-Object { $_.TeamSiteUrl -eq $classicTeamSite.url } | Where-Object { $_.isSubsite -eq $false })
            
                        $subsiteDocumentLibraries = @($script:importedDocumentLibraries | Where-Object { $_.isSubsite -eq $true } | Where-Object { $_.TeamSiteUrl -match $classicTeamSite.url })

                        write-host 
                        $msg = "INFO: Processing classic Team Site $classicTeamSitesIndex/$script:classicTeamSitesCount '$classicTeamSiteName' with URL '$dstClassicTeamSiteUrl'."
                        Write-Host $msg
                        Log-Write -Message $msg 
                        $classicTeamSitesIndex += 1
    
                        $doNotSkipMailbox = $false

                        if ($classicTeamSiteName -eq "" -or $classicTeamSiteSrcUrl -eq "" -or $classicTeamSiteDstUrl -eq "") {
                            $msg = "INFO: Skipping SharePoint  classicTeamSite '$classicTeamSiteName'. Missing data in the CSV file."
                            Write-Host -ForegroundColor Red $msg
                            Log-Write -Message $msg 

                            Continue
                        }    

                        if ($checkDestinationSPOStructure) { $result = check-DestinationSPOSite -url $dstClassicTeamSiteUrl -AdminCenterUrl "https://$script:destinationTenantName-admin.sharepoint.com/" }
                        if (!$result -and $checkDestinationSpoStructure) {
                            $msg = "INFO: Skipping classic Team Site project creation. classic Team Site '$classicTeamSiteName' with URL '$url' not found in destination SharePoint Online."
                            Write-Host -ForegroundColor Red $msg
                            Log-Write -Message $msg 

                            Continue
                        } 
                    
                        if ($null -eq $connectorId) {
                            #Create SPO endpoints
                            $msg = "INFO: Creating MSPC endpoints for classic Team Site '$classicTeamSiteName'."
                            Write-Host $msg
                            Log-Write -Message $msg 

                            $exportEndpointName = "SRC-SPO-$script:sourceTenantName"
                            $exportType = "SharePoint"
                            $exportTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                            $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
                                "Url"                          = $sSPOUrl;
                                "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                                "AdministrativePassword"       = $script:sourcePlainPassword;
                                "UseAdministrativeCredentials" = $true
                            }
            
                            $importEndpointName = "DST-SPO-$script:destinationTenantName"
                            $importType = "SharePointBeta"
                            $importTypeName = "MigrationProxy.WebApi.SharePointBetaConfiguration"
                            if (!$UseOwnAzureStorage) {
                                $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                    "Url"                                = $dSPOUrl;
                                    "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                    "AdministrativePassword"             = $script:destinationPlainPassword;
                                    "UseAdministrativeCredentials"       = $true;
                                    "UseSharePointOnlineProvidedStorage" = $true 
                                }
                            }
                            else {
                                $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                    "Url"                                = $dSPOUrl;
                                    "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                    "AdministrativePassword"             = $script:destinationPlainPassword;
                                    "UseAdministrativeCredentials"       = $true;
                                    "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                    "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;    
                                    "UseSharePointOnlineProvidedStorage" = $false
                                }
                            }

                            #Create SPO Team Sites source endpoint
                            $exportEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType $exportType  -EndpointName $exportEndpointName -EndpointConfiguration $exportConfiguration -update $updateEndpoint
                            #Create SPO Team Sites destination endpoint
                            $importEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType $importType -EndpointName $importEndpointName -EndpointConfiguration $importConfiguration -update $updateEndpoint

                            #Create SPO Team Sites Document project

                            $ProjectName = "ClassicSPOSite-$script:sourceTenantName"
                            $ProjectType = "Storage"

                            if ($enableModernAuth) {
                                $advancedOptions = "$modernAuth InitializationTimeout=8 IncreasePathLengthLimit=1"
                            }
                            else {
                                $advancedOptions = "InitializationTimeout=8 IncreasePathLengthLimit=1 "
                            }

                            if ($applicationPermissions) {
                                $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                            }
                            if ($UseDelegatePermission) {
                                $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                            }

                            $connectorId = $null
                            $connectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                                -ProjectName $ProjectName `
                                -ProjectType $ProjectType `
                                -exportType $exportType `
                                -importType $importType `
                                -exportEndpointId $exportEndpointId `
                                -importEndpointId $importEndpointId `
                                -exportConfiguration $exportConfiguration `
                                -importConfiguration $importConfiguration `
                                -advancedOptions $advancedOptions `
                                -maximumSimultaneousMigrations 100 `
                                -ZoneRequirement $global:btZoneRequirement `
                                -MaxLicensesToConsume 1 `
                                -updateConnector $true   
                        }

                        if ($connectorId -ne $null) {
                        
                            foreach ($documentLibrary in $documentLibraries) {

                                $documentLibraryName = $documentLibrary.DocumentLibraryName

                                if ([string]::IsNullOrEmpty($documentLibraryName)) { Continue }

                                $msg = "      INFO: Processing Document Library '$documentLibraryName'."
                                Write-Host $msg
                                Log-Write -Message $msg 

                                if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                    # Renew MW ticket
                                    Connect-BitTitan 

                                    Write-Host
                                    $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                    Write-Host -ForegroundColor Magenta $msg
                                    Log-Write -Message $msg 
                                    Write-Host                    
                                }   

                                try {
                                    $relativeUrl = $srcClassicTeamSiteUrl -replace $sSPOUrl
                                    $ImportLibrary = $relativeUrl + "/" + $documentLibraryName
                                    $ExportLibrary = $relativeUrl + "/" + $documentLibraryName

                                    $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
                                    if (!$result) {
                                        $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId  -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary

                                        $msg = "      SUCCESS: Document Library migration '$ExportLibrary->$ImportLibrary' added to connector." 
                                        write-Host -ForegroundColor Green $msg
                                        Log-Write -Message $msg  

                                        $processedDocumentLibraries += 1
                                    }
                                    else {
                                        $msg = "      WARNING: Document Library migration '$ExportLibrary->$ImportLibrary' already exists in connector." 
                                        write-Host -ForegroundColor Yellow $msg
                                        Log-Write -Message $msg  

                                        $processedDocumentLibraries += 1
                                    }
                                }
                                catch {
                                    $msg = "      ERROR: Failed to add Document Library migration '$ExportLibrary->$ImportLibrary' to connector." 
                                    write-Host -ForegroundColor Red $msg
                                    Log-Write -Message $msg     
                                }
                            }

                            $processedClassicTeamSites += 1   

                            [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $ProjectName; ProjectType = $ProjectType; ConnectorId = $connectorId; }
                        }

                        ##################################################################################################
                        #    SUBSITES
                        ##################################################################################################
                        $currentSubsiteDocumentLibrary = 0
                        $subsiteDocumentLibraryCount = $subsiteDocumentLibraries.Count
                        $totalSubsitesCount += $subsiteDocumentLibraryCount

                        $previousSubSiteName = ""


                        if ($subsiteDocumentLibraries -ne $null) {
                            foreach ($subsiteDocumentLibrary in $subsiteDocumentLibraries  ) {

                                if ($subsiteDocumentLibrary.SubsiteNumber.SubString(0, 1) -eq 1) {
                                    if (!$alreadyProcessed) {
                                        $ProcessedSubsite = 0
                                        $alreadyProcessed = $true
                                    } 
                                }
                                else {
                                    $alreadyProcessed = $false
                                }

                                if ($ProcessedSubsite -ne $subsiteDocumentLibrary.SubsiteNumber.SubString(0, 1)) {
                                    $ProcessedSubsite += 1 
                                    $allProcessedSubsites += 1 
                                }  
                    
                                $currentSubsiteDocumentLibrary += 1

                                $urlparts = $($subSiteDocumentLibrary.TeamSiteUrl).split("/")
                                $subsiteDepth = $urlparts.Count - 5
                                $subSiteName = $urlparts[-1]
                                $parentSubSiteUrl = ($subSiteDocumentLibrary.TeamSiteUrl -split "/$subSiteName")[0] 
                                $subsiteUrl = $parentSubSiteUrl + "/" + $subSiteName

                                if ($subSiteName -ne $previousSubSiteName) {
                                    $currentSubsite += 1

                                    $srcSubsiteUrl = $subsiteUrl -replace $script:destinationTenantName, $script:sourceTenantName
                                    $dstSubsiteUrl = $subsiteUrl -replace $script:sourceTenantName, $script:destinationTenantName

                                    if ($subsiteDepth -eq 1) {  
                                        Write-Host 
                                        write-host 
                                        $msg = "INFO: Processing $currentSubsite/$script:subsitesCount Subsite '$subSiteName' level-$subsiteDepth of Team Site '$classicTeamSiteName' with URL '$srcSubsiteUrl'."
                                        Write-Host $msg
                                        Log-Write -Message $msg 
                                    }
                                    elseif ($subsiteDepth -gt 1) {
                                        Write-Host
                                        write-host 
                                        $msg = "INFO: Processing $currentSubsite/$script:subsitesCount Subsite '$subSiteName' level-$subsiteDepth  of SubSite '$parentSubSiteUrl' with URL '$srcSubsiteUrl'."
                                        Write-Host $msg
                                        Log-Write -Message $msg 
                                    }

                                    if ($checkDestinationSPOStructure) { $result = check-DestinationSPOSubsite -siteUrl $dstSubsiteUrl -SubSiteName $subSiteName }
                                    if (!$result -and $checkDestinationSpoStructure) {
                                        $msg = "INFO: Skipping classic Team Subsite project creation. Subsite '$subSiteName' with URL '$dstSubsiteUrl' not found in destination SharePoint Online."
                                        Write-Host -ForegroundColor Red $msg
                                        Log-Write -Message $msg 

                                        Continue
                                    }   

                                    if ($null -eq $connectorId) {
                                        #Create SPO endpoints
                                        $msg = "INFO: Creating MSPC endpoints for classic Team Site '$classicTeamSiteName'."
                                        Write-Host $msg
                                        Log-Write -Message $msg 

                                        $subsiteRelativeURL = $subsiteUrl.Replace($srcClassicTeamSiteUrl + "/", "").Replace($dstClassicTeamSiteUrl + "/", "")

                                        $exportEndpointName = "SRC-SPO-$script:sourceTenantName"
                                        $exportType = "SharePoint"
                                        $exportTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                                        $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
                                            "Url"                          = $sSPOUrl;
                                            "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                                            "AdministrativePassword"       = $script:sourcePlainPassword;
                                            "UseAdministrativeCredentials" = $true
                                        }

                                        $importEndpointName = "DST-SPO-$script:destinationTenantName"
                                        $importType = "SharePointBeta"
                                        $importTypeName = "MigrationProxy.WebApi.SharePointBetaConfiguration"
                                        if (!$UseOwnAzureStorage) {
                                            $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                                "Url"                                = $dSPOUrl;
                                                "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                                "AdministrativePassword"             = $script:destinationPlainPassword;
                                                "UseAdministrativeCredentials"       = $true;
                                                "UseSharePointOnlineProvidedStorage" = $true 
                                            }
                                        }
                                        else {
                                            $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                                "Url"                                = $dSPOUrl;
                                                "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                                "AdministrativePassword"             = $script:destinationPlainPassword;
                                                "UseAdministrativeCredentials"       = $true;
                                                "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                                "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;    
                                                "UseSharePointOnlineProvidedStorage" = $false
                                            }
                                        }

                                        #Create SPO Team Sites source endpoint
                                        $exportEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType $exportType  -EndpointName $exportEndpointName -EndpointConfiguration $exportConfiguration -update $updateEndpoint
                                        #Create SPO Team Sites destination endpoint
                                        $importEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType $importType -EndpointName $importEndpointName -EndpointConfiguration $importConfiguration -update $updateEndpoint

                                        #Create SPO Team Sites Document project

                                        $ProjectName = "ClassicSPOSite-$script:sourceTenantName"
                                        $ProjectType = "Storage"

                                        if ($enableModernAuth) {
                                            $advancedOptions = "$modernAuth InitializationTimeout=8 IncreasePathLengthLimit=1"
                                        }
                                        else {
                                            $advancedOptions = "InitializationTimeout=8 IncreasePathLengthLimit=1 "
                                        }

                                        if ($applicationPermissions) {
                                            $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                                        }
                                        if ($UseDelegatePermission) {
                                            $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                                        }

                                        $connectorId = $null
                                        $connectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                                            -ProjectName $ProjectName `
                                            -ProjectType $ProjectType `
                                            -exportType $exportType `
                                            -importType $importType `
                                            -exportEndpointId $exportEndpointId `
                                            -importEndpointId $importEndpointId `
                                            -exportConfiguration $exportConfiguration `
                                            -importConfiguration $importConfiguration `
                                            -advancedOptions $advancedOptions `
                                            -maximumSimultaneousMigrations 100 `
                                            -ZoneRequirement $global:btZoneRequirement `
                                            -MaxLicensesToConsume 1 `
                                            -updateConnector $true   
                                    }    

                                }    

                                if ($connectorId -ne $null) {
            
                                    $documentLibraryName = $subsiteDocumentLibrary.DocumentLibraryName

                                    if ([string]::IsNullOrEmpty($documentLibraryName)) { Continue }

                                    $msg = "      INFO: Processing Document Library '$documentLibraryName'."
                                    Write-Host $msg
                                    Log-Write -Message $msg 

                                    if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                        # Renew MW ticket
                                        Connect-BitTitan 

                                        Write-Host
                                        $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                        Write-Host -ForegroundColor Magenta $msg
                                        Log-Write -Message $msg 
                                        Write-Host                    
                                    }   

                                    try {
                                        $relativeUrl = $srcSubsiteUrl -replace $sSPOUrl
                                        $ImportLibrary = $relativeUrl + "/" + $documentLibraryName
                                        $ExportLibrary = $relativeUrl + "/" + $documentLibraryName

                                        $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
                                        if (!$result) {
                                            $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId  -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary

                                            $msg = "      SUCCESS: Document Library migration '$ExportLibrary->$ImportLibrary' added to connector." 
                                            write-Host -ForegroundColor Green $msg
                                            Log-Write -Message $msg  

                                            $processedSubsiteDocumentLibraries += 1
                                        }
                                        else {
                                            $msg = "      WARNING: Document Library migration '$ExportLibrary->$ImportLibrary' already exists in connector." 
                                            write-Host -ForegroundColor Yellow $msg
                                            Log-Write -Message $msg  

                                            $processedSubsiteDocumentLibraries += 1
                                        }
                                    }
                                    catch {
                                        $msg = "      ERROR: Failed to add Document Library migration '$ExportLibrary->$ImportLibrary' to connector." 
                                        write-Host -ForegroundColor Red $msg
                                        Log-Write -Message $msg     
                                    }

                                    [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $ProjectName; ProjectType = $ProjectType; ConnectorId = $connectorId; }
                                }
                            }
                        }
                    }

                    if ($processedClassicTeamSites -ne 0) {
                        write-Host
                        $msg = "SUCCESS: $processedRootDocumentLibraries out of $script:rootDocumentLibrariesCount Root Document Libraries have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                        $msg = "SUCCESS: $currentRootSubsite out of $script:rootSubsitesCount Root Subsites have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                        $msg = "SUCCESS: $processedRootSubsitesDocumentLibraries out of $script:rootSubSiteDocumentLibrariesCount Root Subsites Document Libraries have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 

                        write-Host
            
                        $msg = "SUCCESS: $processedClassicTeamSites out of $script:classicTeamSitesCount classic Team Sites have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                        $msg = "SUCCESS: $processedDocumentLibraries out of $script:documentLibrariesCount Document Libraries have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                        $msg = "SUCCESS: $allProcessedSubsites out of $script:subsitesCount classic Team Subsites have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg             
                        if ([string]::IsNullOrEmpty(($processedSubsites))) { $processedSubsiteDocumentLibraries = 0 }
                        $msg = "SUCCESS: $processedSubsiteDocumentLibraries out of $script:subsiteDocumentLibrariesCount classic Team Subsites Document Libraries have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                    }

                }  
            }

            ##########################################################################################################################################
            #            Office 365 Groups
            ##########################################################################################################################################
            $connectorId = $null

            if ($migrateO365Groups) {
                write-host 
                $msg = "#######################################################################################################################`
                   CREATING OFFICE 365 (UNIFIED) GROUP PROJECTS                  `
#######################################################################################################################"
                Write-Host $msg
                Log-Write -Message "CREATING OFFICE 365 (UNIFIED) GROUP PROJECTS" 
                write-host 

                $msg = "INFO: Processing Office 365 (unified) group migration."
                Write-Host $msg
                Log-Write -Message $msg 

                if ($importedUnifiedGroups -eq $null) {
                    $msg = "INFO: No Office 365 (unified) Groups found. Skipping Office 365 (unified) Group project creation"
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg 
                }  
                elseif ($script:importedUnifiedGroupDocumentLibraries -eq $null) {
                    $msg = "INFO: No Office 365 (unified) Group Document Libraries found. Skipping Office 365 (unified) Group project creation"
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg 
                } 
                else {
            
                    if (!$UseOwnAzureStorage -and !$AzureStorageSelected) {
                        Write-Host 
                        do {
                            $confirm = (Read-Host -prompt "Do you want to use Microsoft provided Azure Storage?  [Y]es or [N]o")
                            if ($confirm.ToLower() -eq "n") {
                                $UseOwnAzureStorage = $true

                                $AzureStorageAccountName = ''
                                $AzureAccountKey = ''
                
                                do {
                                    $AzureStorageAccountName = (Read-Host -prompt "         Enter Azure Storage Account Name")
                                } while ($AzureStorageAccountName -eq "")

                                do {
                                    $AzureAccountKey = (Read-Host -prompt "         Enter Azure Storage Primary Access Key")
                                } while ($AzureAccountKey -eq "")

                                $AzureStorageSelected = $true
                            }
                            else {
                                $UseOwnAzureStorage = $false
                                $AzureStorageSelected = $true
                            }
                        } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
                    }

                    ##################################################################################################
                    #    SITES AND GROUP MAILBOXES
                    ##################################################################################################

                    $alreadyCreated = $false

                    $currentSubsite = 0
        
                    $ProcessedO365GroupSites = 0
                    $processedO365GroupMailboxes = 0

                    $ProcessedSubsite = 0
                    $allProcessedSubsites = 0
                    $processedDocumentLibraries = 0

                    $UnifiedGroupsIndex = 1

                    $mailboxIDs = @() 

                    foreach ($unifiedGroup in $script:importedUnifiedGroups) {

                        $url = $unifiedGroup.SharePointSiteUrl

                        $srcUnifiedGroupUrl = $url -replace $script:destinationTenantName, $script:sourceTenantName
                        $dstUnifiedGroupUrl = $url -replace $script:sourceTenantName, $script:destinationTenantName
            
                        ##########################################################################################

                        #source and destination email addresses

                        # Source email addresses: UPN, PrimarySmtpAddress and TenantAddress
                        $filteredEmailAddresses = @()                    
                        $emailAddresses = @()
                        $emailAddresses = @($unifiedGroup.EmailAddresses).split("|")

                        foreach ($emailAddress in $emailAddresses) { 
                            if (($emailAddress -cmatch "smtp:" -or $emailAddress -cmatch "SMTP:") -and ($emailAddress -match ".onmicrosoft.com" -and $emailAddress -notmatch ".mail.onmicrosoft.com")) {
                                $filteredEmailAddresses += $emailAddress.replace("SMTP:", "").replace("smtp:", "")
                            }
                        } 

                        $tenantEmailAddress = @($filteredEmailAddresses | Select-Object -Unique)[0]
                        if ($tenantEmailAddress) {
                            $srcTenantEmailAddress = $tenantEmailAddress
                        }
                        else {
                            $srcTenantEmailAddress = $unifiedGroup.PrimarySmtpAddress
                        }
                        $srcPrimarySmtpAddress = $unifiedGroup.PrimarySmtpAddress

                        if ($srcPrimarySmtpAddress.split("@")[0] -eq $srcTenantEmailAddress.split("@")[0]) {
                            $changPrimarySmtpAddress2TenantEmailAddress = $true
                        }
                        else {
                            $changPrimarySmtpAddress2TenantEmailAddress = $false
                        }    

                        $srcMigrationWizAddress = $srcPrimarySmtpAddress                             

                        # Destination email address 
                        if ($script:sameUserName) {      
                            $dstMigrationWizAddress = $unifiedGroup.PrimarySmtpAddress.split("@")[0] + "@" + $script:destinationTenantDomain
                        }
                        else {                
                            $newDstMigrationWizAddress = ($script:emailAddressMappingCSVFile | Where-Object { $_.SourceEmailAddress -eq $unifiedGroup.PrimarySmtpAddress }).DestinationEmailAddress 

                            if (-not ([string]::IsNullOrEmpty($newDstMigrationWizAddress))) {                           
                                $dstMigrationWizAddress = $newDstMigrationWizAddress.split("@")[0] + "@" + $script:destinationTenantDomain 
                            }
                            else {
                                $wrongSourceEmailAddressInCSV = ($script:emailAddressMappingCSVFile | Where-Object { ($_.SourceEmailAddress).split("@")[0] -eq $unifiedGroup.PrimarySmtpAddress.split("@")[0] }).SourceEmailAddress 

                                if (-not ([string]::IsNullOrEmpty($wrongSourceEmailAddressInCSV))) {
                                    write-host -ForegroundColor Red "      ERROR: Invalid SourceEmailAddress '$wrongSourceEmailAddressInCSV' in mapping file for retrieved '$($unifiedGroup.PrimarySmtpAddress)' mailbox."
                                }
                                else {
                                    $wrongSourceEmailAddressInCSV = "<UserNotFound>"
                                    write-host -ForegroundColor Red "      ERROR: '$wrongSourceEmailAddressInCSV' in mapping file for retrieved '$($unifiedGroup.PrimarySmtpAddress)' mailbox."
                                }

                                if ($wrongSourceEmailAddressInCSV -ne "<UserNotFound>") {
                                    $wrongSourceEmailAddressInCSVList += "$wrongSourceEmailAddressInCSV must be replaced in mapping file by retrieved $($unifiedGroup.PrimarySmtpAddress)`n"
                                }
                                else {
                                    $wrongSourceEmailAddressInCSVList += "$wrongSourceEmailAddressInCSV in mapping file for retrieved $($unifiedGroup.PrimarySmtpAddress)`n"
                                }
                                $wrongSourceEmailAddressInCSVCount += 1

                                Continue
                            } 
                        }

                        ##########################################################################################

                        # Source Vanity Domain for Output CSV file

                        if ($unifiedGroup.PrimarySmtpAddress.split("@")[1] -notlike $sourceAlreadyProccesedDomains -and $unifiedGroup.PrimarySmtpAddress.split("@")[1] -notmatch ".onmicrosoft.com") {
                            $sourceAlreadyProccesedDomains += $unifiedGroup.PrimarySmtpAddress.split("@")[1]
                        }            



                        $groupName = $unifiedGroup.Alias

                        if ($checkDestinationSpoStructure) { $result = check-DestinationO365Group -group $dstMigrationWizAddress -groupAlias $groupName }
                        if (!$result -and $checkDestinationSpoStructure) {

                            write-host 
                            $msg = "INFO: Processing Office 365 (unified) Group $UnifiedGroupsIndex/$script:unifiedGroupsCount : '$groupName' '$($unifiedGroup.PrimarySmtpAddress)'."
                            Write-Host $msg
                            Log-Write -Message $msg 

                            if ($recipientMapping.count -ne 0) {
                                $msg = "      INFO: Skipping Office 365 (unified) Group '$groupName' project creation. Office 365 (unified) group not found in destination Office 365."
                                Write-Host -ForegroundColor Red $msg
                                Log-Write -Message $msg 
                            }
                            else {
                                $msg = "      INFO: Skipping Office 365 (unified) Group '$groupName' project creation. Office 365 (unified) group not found in destination Office 365."
                                Write-Host -ForegroundColor Red $msg
                                Log-Write -Message $msg 
                            }
                            $UnifiedGroupsIndex += 1
                            Continue
                        } 

                        write-host 
                        $msg = "INFO: Processing Office 365 (unified) Group $UnifiedGroupsIndex/$script:unifiedGroupsCount : '$groupName' '$($unifiedGroup.PrimarySmtpAddress)' $srcUnifiedGroupUrl."
                        Write-Host $msg
                        Log-Write -Message $msg 
                        $UnifiedGroupsIndex += 1

                        if ($linkO365Groups) {
                            if ($checkDestinationSpoStructure) {
                                Link-O365Groups $groupName $unifiedGroup.PrimarySmtpAddress $script:SrcAdministrativeUsername $dstMigrationWizAddress $script:DstAdministrativeUsername
                            }
                            else {
                                Link-O365Groups $groupName $unifiedGroup.PrimarySmtpAddress $script:SrcAdministrativeUsername 
                            }
                        }            
    
                        $doNotSkipMailbox = $false

                        if ($groupName -eq "" -or $srcUnifiedGroupUrl -eq "" -or $dstUnifiedGroupUrl -eq "" -or $srcMigrationWizAddress -eq "" -or $dstMigrationWizAddress -eq "") {
                            $msg = "INFO: Office 365 (unified) Group '$groupName' has no SharePointSiteUrl. Skipping endpoint and Document connector creation."
                            Write-Host -ForegroundColor Red $msg
                            Log-Write -Message $msg 

                            if ($srcMigrationWizAddress -ne "" -and $dstMigrationWizAddress -ne "" ) {
                                #-and $ProcessedO365GroupSites -ge 1
                                $doNotSkipMailbox = $true
                            }
                            else {
                                $doNotSkipMailbox = $false
                                #Continue
                            }  
                        }     
            
                        #Create O365 unified Group Mailbox project    

                        if ($alreadyCreated -eq $false) {

                            #$ProjectName = "Mailbox-$groupName"
                            $mailboxProjectName = "O365Group-Mailbox-All conversations-$script:sourceTenantName"
                            $projectTypeName = "MigrationProxy.WebApi.ExchangeConfiguration"
                            $ProjectType = "Mailbox"

                            if ($script:dstGermanyCloud) { $importType = "ExchangeOnlineGermany" }
                            elseif ($script:dstUsGovernment) { $importType = "ExchangeOnlineUsGovernment" }
                            else { $importType = "ExchangeOnline2" }

                            if ($script:srcGermanyCloud) { $exportType = "ExchangeOnlineGermany" }
                            elseif ($script:srcUsGovernment) { $exportType = "ExchangeOnlineUsGovernment" }
                            else { $exportType = "ExchangeOnline2" }

                            $exportEndpointId = $exportEndpointId
                            $importEndpointId = $importEndpointId
                            $exportConfiguration = New-Object -TypeName $projectTypeName -Property @{
                                "Url"                          = $srcUnifiedGroupUrl;
                                "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                                "AdministrativePassword"       = $script:sourcePlainPassword;
                                "UseAdministrativeCredentials" = $true
                            }
                            $importConfiguration = New-Object -TypeName $projectTypeName -Property @{
                                "Url"                          = $dstUnifiedGroupUrl;
                                "AdministrativeUsername"       = $script:DstAdministrativeUsername;
                                "AdministrativePassword"       = $script:destinationPlainPassword;
                                "UseAdministrativeCredentials" = $true
                            }
                            $folderFilter = "^(?!Inbox|Calendar)"

                            if ($enableModernAuth) {
                                $advancedOptions = "$modernAuth "
                                $advancedOptions += $recipientMapping
                            }
                            else {
                                $advancedOptions = $recipientMapping
                            }

                            if ($applicationPermissions) {
                                $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                            }
                            if ($UseDelegatePermission) {
                                $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                            }

                            if ($totalLines -ge 400) {
                                $maximumSimultaneousMigrations = 400
                            }
                            elseif ($totalLines -le 10) {
                                $maximumSimultaneousMigrations = 10
                            }
                            else {
                                $maximumSimultaneousMigrations = $totalLines
                            }

                            $mailboxConnectorId = $null
                            $mailboxConnectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                                -ProjectName $mailboxProjectName `
                                -ProjectType $ProjectType `
                                -importType $importType `
                                -exportType $exportType `
                                -exportEndpointId $exportEndpointId `
                                -importEndpointId $importEndpointId `
                                -exportConfiguration $exportConfiguration `
                                -importConfiguration $importConfiguration `
                                -advancedOptions $advancedOptions `
                                -folderFilter $folderFilter `
                                -ZoneRequirement $global:btZoneRequirement `
                                -maximumSimultaneousMigrations $maximumSimultaneousMigrations `
                                -updateConnector $true   

                            $alreadyCreated = $true 
                        }             
            

                        if (!$doNotSkipMailbox) {       
                            #Create O365 unified Group source endpoint

                            #Create SPO endpoints
                            $msg = "INFO: Creating MSPC endpoints for Office 365 (unified) Group '$groupName'."
                            Write-Host $msg
                            Log-Write -Message $msg 
    
                            $exportEndpointName = "SRC-O365G-$srcMigrationWizAddress"
                            $endpointTypeName = "ManagementProxy.ManagementService.SharePointConfiguration"
                            $endpointType = "Office365Groups"
                            $exportConfiguration = New-Object -TypeName $endpointTypeName -Property @{
                                "Url"                          = $srcUnifiedGroupUrl;
                                "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                                "AdministrativePassword"       = $script:sourcePlainPassword;
                                "UseAdministrativeCredentials" = $true
                            }

                            [guid]$exportEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType $endpointType  -EndpointName $exportEndpointName -EndpointConfiguration $exportConfiguration -update $updateEndpoint

                            #Create O365 unified Group destination endpoint

                            if (!$UseOwnAzureStorage) {
                                $importEndpointName = "DST-O365G-$dstMigrationWizAddress"
                                #$endpointTypeName = "ManagementProxy.ManagementService.SharePointConfiguration"
                                $endpointTypeName = "ManagementProxy.ManagementService.SharePointOnlineConfiguration"
                                #$endpointType = = "Office365Groups"
                                $endpointType = "SharePointOnlineAPI"
                                $importConfiguration = New-Object -TypeName $endpointTypeName -Property @{
                                    "Url"                                = $dstUnifiedGroupUrl;
                                    "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                    "AdministrativePassword"             = $script:destinationPlainPassword;
                                    "UseAdministrativeCredentials"       = $true;
                                    "UseSharePointOnlineProvidedStorage" = $true
                                }
                            }
                            else {
                                $importEndpointName = "DST-O365G-$dstMigrationWizAddress"
                                #$endpointTypeName = "ManagementProxy.ManagementService.SharePointConfiguration"
                                $endpointTypeName = "ManagementProxy.ManagementService.SharePointOnlineConfiguration"
                                #$endpointType = = "Office365Groups"
                                $endpointType = "SharePointOnlineAPI"
                                $importConfiguration = New-Object -TypeName $endpointTypeName -Property @{
                                    "Url"                                = $dstUnifiedGroupUrl;
                                    "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                    "AdministrativePassword"             = $script:destinationPlainPassword;
                                    "UseAdministrativeCredentials"       = $true;
                                    "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                    "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;    
                                    "UseSharePointOnlineProvidedStorage" = $false
                                }
                            }
    
                            [guid]$importEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType $endpointType -EndpointName $importEndpointName -EndpointConfiguration $importConfiguration -update $updateEndpoint
    
                            #Create O365 unified Group Document project
                    
                            $docProjectName = "O365Group-Document-$groupName-$script:sourceTenantName"
                            $ProjectType = "Storage"    
                            $exportType = "Office365Groups"
                            #$importType = "Office365Groups"
                            $importType = "SharePointOnlineAPI"
    
                            $exportEndpointId = $exportEndpointId
                            $importEndpointId = $importEndpointId

                            $exportTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                            $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
                                "Url"                          = $srcUnifiedGroupUrl;
                                "AdministrativeUsername"       = $exportEndpointData.AdministrativeUsername;
                                "AdministrativePassword"       = $script:sourcePlainPassword;
                                "UseAdministrativeCredentials" = $true
                            }
                            #$importTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                            $importTypeName = "MigrationProxy.WebApi.SharePointOnlineConfiguration"
                            $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                "Url"                                = $dstUnifiedGroupUrl;
                                "AdministrativeUsername"             = $importEndpointData.AdministrativeUsername;
                                "AdministrativePassword"             = $script:destinationPlainPassword;
                                "UseAdministrativeCredentials"       = $true;
                                "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;    
                                "UseSharePointOnlineProvidedStorage" = $false
                            }

                            if ($enableModernAuth) {
                                $AdvancedOptions = "$modernAuth InitializationTimeout=8"
                            }
                            else {
                                $AdvancedOptions = "InitializationTimeout=8"
                            }

                            if ($applicationPermissions) {
                                $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                            }
                            if ($UseDelegatePermission) {
                                $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                            }

                            $docConnectorId = $null
                            $docConnectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                                -ProjectName $docProjectName `
                                -ProjectType $ProjectType `
                                -importType $importType `
                                -exportType $exportType `
                                -exportEndpointId $exportEndpointId `
                                -importEndpointId $importEndpointId `
                                -exportConfiguration $exportConfiguration `
                                -importConfiguration $importConfiguration `
                                -advancedOptions $advancedOptions `
                                -maximumSimultaneousMigrations 100 `
                                -ZoneRequirement $global:btZoneRequirement `
                                -MaxLicensesToConsume 10 `
                                -updateConnector $true   

                            if ($docConnectorId -ne $null) {   

                                $ProcessedO365GroupSites += 1  

                                [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $docProjectName; ProjectType = $ProjectType; ConnectorId = $docConnectorId; }

                                $documentLibraries = @()
                                $documentLibraryNames = ($script:importedUnifiedGroupDocumentLibraries | Where-Object { $_.TeamSiteUrl -eq $srcUnifiedGroupUrl } | Select-Object DocumentLibraryName).DocumentLibraryName
                  
                                if ($documentLibraryNames -eq $null) { $documentLibraryNames = ($script:importedUnifiedGroupDocumentLibraries | Where-Object { $_.TeamSiteUrl -eq $dstUnifiedGroupUrl } | Select-Object DocumentLibraryName).DocumentLibraryName }

                                foreach ($documentLibraryName in $documentLibraryNames) {

                                    if ([string]::IsNullOrEmpty($documentLibraryName)) { Continue }
             
                                    $msg = "      INFO: Processing Document Library '$documentLibraryName'."
                                    Write-Host $msg
                                    Log-Write -Message $msg 

                                    if ($documentLibraryName -eq "SiteAssets") { Continue }
                                    if ($documentLibraryName -eq "Teams Wiki Data") { Continue }
                        
                                    if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                        # Renew MW ticket
                                        Connect-BitTitan 

                                        Write-Host
                                        $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                        Write-Host -ForegroundColor Magenta $msg
                                        Log-Write -Message $msg 
                                        Write-Host                    
                                    }   

                                    try {
                                        if ($customDocumentLibrary) {
                                            $ExportLibrary = "Shared Documents"
                                            $ImportLibrary = "Shared Documents"
                                        }
                                        else {    
                                            $ExportLibrary = $documentLibraryName
                                            $ImportLibrary = $documentLibraryName
                                        }
        
                                        $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $docConnectorId -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
                                        if (!$result) {
                                            $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $docConnectorId  -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary

                                            $msg = "      SUCCESS: Document Library migration '$ExportLibrary->$ImportLibrary' added to connector." 
                                            write-Host -ForegroundColor Green $msg
                                            Log-Write -Message $msg   
                                
                                            $processedDocumentLibraries += 1                    
                                        }
                                        else {
                                            $msg = "      WARNING: Document Library migration '$ExportLibrary->$ImportLibrary' already exists in connector." 
                                            write-Host -ForegroundColor Yellow $msg
                                            Log-Write -Message $msg  

                                            $processedDocumentLibraries += 1 
                                        }                      
                                    }
                                    catch {
                                        $msg = "      ERROR: Failed to add Document Library migration '$ExportLibrary->$ImportLibrary' to connector." 
                                        write-Host -ForegroundColor Red $msg
                                        Log-Write -Message $msg     
                                    }    
                                }                     

                            }  

                            ##################################################################################################
                            #    SUBSITES
                            ##################################################################################################

                            $subsiteDocumentLibraries = ($script:importedUnifiedGroupDocumentLibraries | Where-Object { $_.isSubsite -eq $true } | Where-Object { $_.TeamSiteUrl -match $srcUnifiedGroupUrl }) 
                            if ($subsiteDocumentLibraries -eq $null) { $subsiteDocumentLibraries = ($script:importedUnifiedGroupDocumentLibraries | Where-Object { $_.isSubsite -eq $true } | Where-Object { $_.TeamSiteUrl -match $dstUnifiedGroupUrl }) }

                            $currentSubsiteDocumentLibrary = 0
                            $subsiteDocumentLibraryCount = $subsiteDocumentLibraries.Count
                            $totalSubsitesCount += $subsiteDocumentLibraryCount

                            $previousSubSiteName = ""

                            if ($subsiteDocumentLibraries -ne $null) {

                                foreach ($subsiteDocumentLibrary in $subsiteDocumentLibraries) {

                                    if ($subsiteDocumentLibrary.SubsiteNumber.SubString(0, 1) -eq 1) {
                                        if (!$alreadyProcessed) {
                                            $ProcessedSubsite = 0
                                            $alreadyProcessed = $true
                                        } 
                                    }
                                    else {
                                        $alreadyProcessed = $false
                                    }
    
                                    if ($ProcessedSubsite -ne $subsiteDocumentLibrary.SubsiteNumber.SubString(0, 1)) {
                                        $ProcessedSubsite += 1 
                                        $allProcessedSubsites += 1 
                                    }  

                                    $currentSubsiteDocumentLibrary += 1

                                    $urlparts = $($subSiteDocumentLibrary.TeamSiteUrl).split("/")
                                    $subsiteDepth = $urlparts.Count - 5
                                    $subSiteName = $urlparts[-1]
                                    $parentSubSiteUrl = ($subSiteDocumentLibrary.TeamSiteUrl -split "/$subSiteName")[0] 
                                    $subsiteUrl = $parentSubSiteUrl + "/" + $subSiteName

                                    if ($subSiteName -ne $previousSubSiteName) {
                                        $currentSubsite += 1
                                    }

                                    $srcSubsiteUrl = $subsiteUrl -replace $script:destinationTenantName, $script:sourceTenantName
                                    $dstSubsiteUrl = $subsiteUrl -replace $script:sourceTenantName, $script:destinationTenantName

                                    if ($subsiteDepth -eq 1) {  
                                        Write-Host 
                                        write-host 
                                        $msg = "INFO: Processing $currentSubsite/$script:unifiedGroupSubsitesCount Subsite '$subSiteName' level-$subsiteDepth of Office 365 Group '$groupName' with URL '$srcSubsiteUrl'."
                                        Write-Host $msg
                                        Log-Write -Message $msg 
                                    }
                                    elseif ($subsiteDepth -gt 1) {
                                        Write-Host
                                        write-host 
                                        $msg = "INFO: Processing $currentSubsite/$script:unifiedGroupSubsitesCount Subsite '$subSiteName' level-$subsiteDepth  of SubSite '$parentSubSiteUrl' with URL '$srcSubsiteUrl'."
                                        Write-Host $msg
                                        Log-Write -Message $msg 
                                    }
                        
                                    $doNotSkipMailbox = $false

                                    if ($checkDestinationSpoStructure) { $result = check-DestinationSPOSubsite -siteUrl $dstSubsiteUrl -SubSiteName $subSiteName }
                                    if (!$result -and $checkDestinationSpoStructure) {
                                        $msg = "INFO: Skipping Office 365 Group Subsite project creation. Subsite '$subSiteName' with URL '$dstSubsiteUrl' not found in destination SharePoint Online."
                                        Write-Host -ForegroundColor Red $msg
                                        Log-Write -Message $msg 

                                        Continue
                                    }                     
                
                                    #Create SPO endpoints
                                    $msg = "INFO: Creating MSPC endpoints for Office 365 Group Subsite '$subSiteName'."
                                    Write-Host $msg
                                    Log-Write -Message $msg 

                                    $exportEndpointName = "SRC-O365G-$groupName-Subsite$subsiteDepth-$subSiteName-$script:sourceTenantName" #-Subsite$subsiteDepth-
                                    $exportType = "Office365Groups"
                                    $exportTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                                    $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
                                        "Url"                          = $srcSubsiteUrl;
                                        "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                                        "AdministrativePassword"       = $script:sourcePlainPassword;
                                        "UseAdministrativeCredentials" = $true
                                    }
    
                                    $importEndpointName = "DST-O365G-$groupName-Subsite$subsiteDepth-$subSiteName-$script:destinationTenantName" #-Subsite$subsiteDepth-
                                    $importType = "SharePointOnlineAPI"
                                    $importTypeName = "MigrationProxy.WebApi.SharePointOnlineConfiguration"
                                    if (!$UseOwnAzureStorage) {
                                        $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                            "Url"                                = $dstSubsiteUrl;
                                            "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                            "AdministrativePassword"             = $script:destinationPlainPassword;
                                            "UseAdministrativeCredentials"       = $true;
                                            "UseSharePointOnlineProvidedStorage" = $true 
                                        }
                                    }
                                    else {
                                        $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                            "Url"                                = $dstSubsiteUrl;
                                            "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                            "AdministrativePassword"             = $script:destinationPlainPassword;
                                            "UseAdministrativeCredentials"       = $true;
                                            "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                            "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;    
                                            "UseSharePointOnlineProvidedStorage" = $false
                                        }
                                    }

                                    #Create SPO Team Sites source endpoint
                                    $exportEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType $exportType  -EndpointName $exportEndpointName -EndpointConfiguration $exportConfiguration -update $updateEndpoint
                                    #Create SPO Team Sites destination endpoint
                                    $importEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType $importType -EndpointName $importEndpointName -EndpointConfiguration $importConfiguration -update $updateEndpoint

                                    #Create SPO Team Sites Document project

                                    $ProjectName = "O365Group-Document-$groupName-Subsite$subsiteDepth-$subSiteName-$script:sourceTenantName" #-Subsite$subsiteDepth-
                                    $ProjectType = "Storage"

                                    if ($enableModernAuth) {
                                        $advancedOptions = "$modernAuth InitializationTimeout=8 IncreasePathLengthLimit=1"
                                    }
                                    else {
                                        $advancedOptions = "InitializationTimeout=8 IncreasePathLengthLimit=1"
                                    }

                                    if ($applicationPermissions) {
                                        $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                                    }
                                    if ($UseDelegatePermission) {
                                        $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                                    }

                                    $connectorId = $null
                                    $connectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                                        -ProjectName $ProjectName `
                                        -ProjectType $ProjectType `
                                        -exportType $exportType `
                                        -importType $importType `
                                        -exportEndpointId $exportEndpointId `
                                        -importEndpointId $importEndpointId `
                                        -exportConfiguration $exportConfiguration `
                                        -importConfiguration $importConfiguration `
                                        -advancedOptions $advancedOptions `
                                        -maximumSimultaneousMigrations 100 `
                                        -ZoneRequirement $global:btZoneRequirement `
                                        -MaxLicensesToConsume 1 `
                                        -updateConnector $true

                                    if ($connectorId -ne $null) {
        
                                        $documentLibraryName = $subsiteDocumentLibrary.DocumentLibraryName
    
                                        if ([string]::IsNullOrEmpty($documentLibraryName)) { Continue }
    
                                        $msg = "      INFO: Processing $currentSubsiteDocumentLibrary/$subsiteDocumentLibraryCount Subsite Document Library '$documentLibraryName'."
                                        Write-Host $msg
                                        Log-Write -Message $msg 
    
                                        if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                            # Renew MW ticket
                                            Connect-BitTitan 
    
                                            Write-Host
                                            $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                            Write-Host -ForegroundColor Magenta $msg
                                            Log-Write -Message $msg 
                                            Write-Host                    
                                        }    
                                                            
                                        try {
                                            $ImportLibrary = $documentLibraryName
                                            $ExportLibrary = $documentLibraryName
    
                                            $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
                                            if (!$result) {
                                                $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId  -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
    
                                                $msg = "      SUCCESS: Document Library migration '$ExportLibrary->$ImportLibrary' added to connector." 
                                                write-Host -ForegroundColor Green $msg
                                                Log-Write -Message $msg  
    
                                                $processedSubsiteDocumentLibraries += 1
                                            }
                                            else {
                                                $msg = "      WARNING: Document Library migration '$ExportLibrary->$ImportLibrary' already exists in connector." 
                                                write-Host -ForegroundColor Yellow $msg
                                                Log-Write -Message $msg  
    
                                                $processedSubsiteDocumentLibraries += 1
                                            }
    
                                        }
                                        catch {
                                            $msg = "      ERROR: Failed to add Document Library migration '$ExportLibrary->$ImportLibrary' to connector." 
                                            write-Host -ForegroundColor Red $msg
                                            Log-Write -Message $msg     
                                        }                         
    
                                        [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $ProjectName; ProjectType = $ProjectType; ConnectorId = $connectorId; }
                                    }            
                        
                                    $previousSubSiteName = $subSiteName
                                    
                                }
                            }             
                        } 

                        if ($mailboxConnectorId -ne $null) {
            
                            $msg = "      INFO: Processing Office 365 (unified) group mailbox '$srcMigrationWizAddress'."
                            Write-Host $msg
                            Log-Write -Message $msg 

                            ##########################################################################################
                    
                            #MW ticket renewal
                    
                            if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                # Renew MW ticket
                                Connect-BitTitan 

                                Write-Host
                                $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                Write-Host -ForegroundColor Magenta $msg
                                Log-Write -Message $msg 
                                Write-Host                    
                            }
                            ##########################################################################################

                            $mailbox = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $mailboxConnectorId -ImportEmailAddress $dstMigrationWizAddress
                            if (!$mailbox) {
                                $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $mailboxConnectorId -ImportEmailAddress $dstMigrationWizAddress -ExportEmailAddress $srcMigrationWizAddress 
                
                                $mailboxIDs += $result.Id.Guid
                                $mailboxIDsInBatch += $result.Id.Guid
                
                                $msg = "      SUCCESS: Office 365 (unified) group migration '$srcMigrationWizAddress->$dstMigrationWizAddress' added to connector."  
                                write-Host -ForegroundColor Green $msg
                                Log-Write -Message $msg  
                
                                $processedO365GroupMailboxes += 1

                                ########################################################################
                                # If mailbox is added and was previously licensed    
                                $mspcUser = (Get-BT_CustomerEndUser -Ticket $script:ticket -OrganizationId $global:btCustomerOrganizationId -id $result.CustomerEndUserId -Environment "BT" -IsDeleted $false) 
                        
                                if (!$mspcUser) {
                                    Write-host -ForegroundColor Red "      ERROR: User '$(($result.ExportEmailAddress))' not found in MSPComplete."
                                }

                                if ($mspcUser) {
                                    $subscriptionEndDate = (Get-BT_Subscription -Ticket $script:ticket -Id $mspcuser.SubscriptionId.guid).SubscriptionEndDate

                                    if ( $mspcuser.ActiveSubscriptionId -eq "00000000-0000-0000-0000-000000000000" ) {
                                        Write-host -ForegroundColor Yellow "      WARNING: User '$($mspcuser.PrimaryEmailAddress)' does not have a subscription applied."

                                        $isUmbApplied = $false
                                    }
                                    else {
                                        Write-host -ForegroundColor Green "      SUCCESS: User '$($mspcuser.PrimaryEmailAddress)' has a subscription applied that will expire in '$subscriptionEndDate'. "

                                        $isUmbApplied = $true
                                    } 

                                    if ($isUmbApplied) {
                                        if ($result) {    			
                                            if ($changeUPN2TenantEmailAddress) {		
                                                $Result = Set-MW_Mailbox -Ticket $script:MwTicket -ConnectorId $mailboxConnectorId -mailbox $result -ExportEmailAddress $srcTenantEmailAddress
                                            }
                                            else {
                                                $Result = Set-MW_Mailbox -Ticket $script:MwTicket -ConnectorId $mailboxConnectorId -mailbox $result -ExportEmailAddress $srcTenantEmailAddress -Categories "tag-3"
                                            }
                                            Write-host -ForegroundColor Blue "      SUCCESS: Source MigrationWiz address for user '$($mspcuser.PrimaryEmailAddress)' changed to '$srcTenantEmailAddress'."
                                        }
                                        else {
                                            $msg = "      ERROR: Failed to change mailbox ExportEmailAddress to '$srcTenantEmailAddress'."
                                            Write-Host -ForegroundColor Red  $msg
                                            Log-Write -Message $msg 
                                        }
                                    }
                                }
                            }
                            else {
                                $msg = "      WARNING: Office 365 (unified) group migration '$($mailbox.ExportEmailAddress)->$dstMigrationWizAddress' already exists in connector."  
                                write-Host -ForegroundColor Yellow $msg
                                Log-Write -Message $msg  
                
                                $existingMigrationsList += "$srcMigrationWizAddress->$dstMigrationWizAddress`n"
                    
                                $existingUserMailboxesCount += 1

                                $processedO365GroupMailboxes += 1
                
                                $mailboxIDs += $mailbox.Id.Guid
                                $mailboxIDsInBatch += $mailbox.Id.Guid
                
                                $mspcUser = (Get-BT_CustomerEndUser -Ticket $script:ticket -OrganizationId $global:btCustomerOrganizationId -id $mailbox.CustomerEndUserId -Environment "BT" -IsDeleted $false) 
                    
                                if (!$mspcUser) {
                                    Write-host -ForegroundColor Red "      ERROR: User '$(($mailbox.ExportEmailAddress))' not found in MSPComplete."
                                }
                
                                if ($mspcUser) {
                                    $subscriptionEndDate = (Get-BT_Subscription -Ticket $script:ticket -Id $mspcuser.SubscriptionId.guid).SubscriptionEndDate
                
                                    if ( $mspcuser.ActiveSubscriptionId -eq "00000000-0000-0000-0000-000000000000" ) {
                                        Write-host -ForegroundColor Yellow "      WARNING: User '$($mspcuser.PrimaryEmailAddress)' does not have a subscription applied."
                
                                        $isUmbApplied = $false
                                    }
                                    else {
                                        Write-host -ForegroundColor Green "      SUCCESS: User '$($mspcuser.PrimaryEmailAddress)' has a subscription applied that will expire in '$subscriptionEndDate'. "
                
                                        $isUmbApplied = $true
                                    } 
                                }
                
                                if ($mspcUser -and !$isUmbApplied -and $script:ApplyBitTitanLicenses ) {
                
                                    Try {
                                        $subscription = Add-BT_Subscription -ticket $script:WorkgroupTicket -ReferenceEntityType CustomerEndUser -ReferenceEntityId $mspcuser.Id -ProductSkuId $productId -WorkgroupOrganizationId $global:btWorkgroupOrganizationId -ErrorAction Stop
                            
                                        $msg = "      SUCCESS: User Migration Bundle subscription assigned to MSPC User '$($mspcUser.PrimaryEmailAddress)' and migration '$srcMigrationWizAddress->$dstMigrationWizAddress'."
                                        Write-Host -ForegroundColor Green  $msg
                                        Log-Write -Message $msg 
                
                                        $changeCount += 1 
                                    }
                                    Catch {
                                        $msg = "      ERROR: Failed to assign User Migration License subscription to MSPC User '$($mspcUser.PrimaryEmailAddress)'."
                                        Write-Host -ForegroundColor Red  $msg
                                        Log-Write -Message $msg
                                        Write-Host -ForegroundColor Red $($_.Exception.Message)
                                        Log-Write -Message $($_.Exception.Message) 
                                    }
                
                                    if ($subscription -and $mailbox) {    			
                                        if ($changePrimarySmtpAddress2TenantEmailAddress) {		
                                            $Result = Set-MW_Mailbox -Ticket $script:MwTicket -ConnectorId $mailboxConnectorId -mailbox $mailbox -ExportEmailAddress $srcTenantEmailAddress
                                        }
                                        else {
                                            $Result = Set-MW_Mailbox -Ticket $script:MwTicket -ConnectorId $mailboxConnectorId -mailbox $mailbox -ExportEmailAddress $srcTenantEmailAddress -Categories "tag-3"
                                        }
                                        Write-host -ForegroundColor Blue "      SUCCESS: Source MigrationWiz address for user '$($mspcuser.PrimaryEmailAddress)' changed to '$srcTenantEmailAddress'."
                                    }
                                    else {
                                        $msg = "      ERROR: Failed to change mailbox ExportEmailAddress to '$srcTenantEmailAddress'."
                                        Write-Host -ForegroundColor Red  $msg
                                        Log-Write -Message $msg 
                                    }
                                }
                                elseif ($mspcUser -and $isUmbApplied) {
                                    if ($mailbox) {    			
                                        if ($changePrimarySmtpAddress2TenantEmailAddress) {		
                                            $Result = Set-MW_Mailbox -Ticket $script:MwTicket -ConnectorId $mailboxConnectorId -mailbox $mailbox -ExportEmailAddress $srcTenantEmailAddress
                                        }
                                        else {
                                            $Result = Set-MW_Mailbox -Ticket $script:MwTicket -ConnectorId $mailboxConnectorId -mailbox $mailbox -ExportEmailAddress $srcTenantEmailAddress -Categories "tag-3"
                                        }
                                        Write-host -ForegroundColor Blue "      SUCCESS: Source MigrationWiz address for user '$($mspcuser.PrimaryEmailAddress)' changed to '$srcTenantEmailAddress'."
                                    }
                                    else {
                                        $msg = "      ERROR: Failed to change mailbox ExportEmailAddress to '$srcTenantEmailAddress'."
                                        Write-Host -ForegroundColor Red  $msg
                                        Log-Write -Message $msg 
                                    }
                                }
                            }
                        }              
                    }

                    if ($ProcessedO365GroupSites -ne 0) {
                        write-Host
                        $msg = "SUCCESS: $ProcessedO365GroupSites out of $script:unifiedGroupsCount Office 365 Groups have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 

                        $msg = "SUCCESS: $processedO365GroupMailboxes out of $script:unifiedGroupsCount Office 365 Group mailboxes have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 

                        if ([string]::IsNullOrEmpty(($processedDocumentLibraries))) { $processedDocumentLibraries = 0 }
                        $msg = "SUCCESS: $processedDocumentLibraries out of $script:unifiedGroupDocumentLibrariesCount Office 365 Group Document Libraries have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                        write-Host
                        $msg = "SUCCESS: $allProcessedSubsites out of $script:unifiedGroupSubsitesCount Office 365 Groups Subsites have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 

                        if ([string]::IsNullOrEmpty(($allProcessedSubsites))) { $processedSubsiteDocumentLibraries = 0 }
                        $msg = "SUCCESS: $processedSubsiteDocumentLibraries out of $script:unifiedGroupSubSiteDocumentLibrariesCount Office 365 Group Subsites Document Libraries have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                    }

                    if ($mailboxConnectorId -ne $null) {
                        [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $mailboxProjectName; ProjectType = $ProjectType; ConnectorId = $mailboxConnectorId; }
                    }   
                }  
            }
            elseif ($migrateNewO365Groups) {
                write-host 
                $msg = "#######################################################################################################################`
                   CREATING OFFICE 365 (UNIFIED) GROUP PROJECTS                  `
#######################################################################################################################"
                Write-Host $msg
                Log-Write -Message "CREATING OFFICE 365 (UNIFIED) GROUP PROJECTS" 
                write-host 

                $msg = "INFO: Processing Office 365 (unified) group migration."
                Write-Host $msg
                Log-Write -Message $msg 

                if ($importedUnifiedGroups -eq $null) {
                    $msg = "INFO: No Office 365 (unified) Groups found. Skipping Office 365 (unified) Group project creation"
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg 
                }  
                elseif ($script:importedUnifiedGroupDocumentLibraries -eq $null) {
                    $msg = "INFO: No Office 365 (unified) Group Document Libraries found. Skipping Office 365 (unified) Group project creation"
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg 
                } 
                else {
            
                    if (!$UseOwnAzureStorage -and !$AzureStorageSelected) {
                        Write-Host 
                        do {
                            $confirm = (Read-Host -prompt "Do you want to use Microsoft provided Azure Storage?  [Y]es or [N]o")
                            if ($confirm.ToLower() -eq "n") {
                                $UseOwnAzureStorage = $true

                                $AzureStorageAccountName = ''
                                $AzureAccountKey = ''
                
                                do {
                                    $AzureStorageAccountName = (Read-Host -prompt "         Enter Azure Storage Account Name")
                                } while ($AzureStorageAccountName -eq "")

                                do {
                                    $AzureAccountKey = (Read-Host -prompt "         Enter Azure Storage Primary Access Key")
                                } while ($AzureAccountKey -eq "")

                                $AzureStorageSelected = $true
                            }
                            else {
                                $UseOwnAzureStorage = $false
                                $AzureStorageSelected = $true
                            }
                        } while (($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
                    }

                    ##################################################################################################
                    #    SITES AND GROUP MAILBOXES
                    ##################################################################################################

                    $alreadyCreated = $false

                    $currentSubsite = 0
        
                    $ProcessedO365GroupSites = 0
                    $processedO365GroupMailboxes = 0

                    $ProcessedSubsite = 0
                    $allProcessedSubsites = 0
                    $processedDocumentLibraries = 0

                    $UnifiedGroupsIndex = 1

                    $mailboxIDs = @() 

                    foreach ($unifiedGroup in $script:importedUnifiedGroups) {

                        $url = $unifiedGroup.SharePointSiteUrl

                        $srcUnifiedGroupUrl = $url -replace $script:destinationTenantName, $script:sourceTenantName
                        $dstUnifiedGroupUrl = $url -replace $script:sourceTenantName, $script:destinationTenantName
            
                        ##########################################################################################

                        #source and destination email addresses

                        # Source email addresses: UPN, PrimarySmtpAddress and TenantAddress
                        $filteredEmailAddresses = @()                    
                        $emailAddresses = @()
                        $emailAddresses = @($unifiedGroup.EmailAddresses).split("|")

                        foreach ($emailAddress in $emailAddresses) { 
                            if (($emailAddress -cmatch "smtp:" -or $emailAddress -cmatch "SMTP:") -and ($emailAddress -match ".onmicrosoft.com" -and $emailAddress -notmatch ".mail.onmicrosoft.com")) {
                                $filteredEmailAddresses += $emailAddress.replace("SMTP:", "").replace("smtp:", "")
                            }
                        } 

                        $tenantEmailAddress = @($filteredEmailAddresses | Select-Object -Unique)[0]
                        if ($tenantEmailAddress) {
                            $srcTenantEmailAddress = $tenantEmailAddress
                        }
                        else {
                            $srcTenantEmailAddress = $unifiedGroup.PrimarySmtpAddress
                        }
                        $srcPrimarySmtpAddress = $unifiedGroup.PrimarySmtpAddress

                        if ($srcPrimarySmtpAddress.split("@")[0] -eq $srcTenantEmailAddress.split("@")[0]) {
                            $changPrimarySmtpAddress2TenantEmailAddress = $true
                        }
                        else {
                            $changPrimarySmtpAddress2TenantEmailAddress = $false
                        }    

                        $srcMigrationWizAddress = $srcPrimarySmtpAddress                             

                        # Destination email address 
                        if ($script:sameUserName) {      
                            $dstMigrationWizAddress = $unifiedGroup.PrimarySmtpAddress.split("@")[0] + "@" + $script:destinationTenantDomain
                        }
                        else {                
                            $newDstMigrationWizAddress = ($script:emailAddressMappingCSVFile | Where-Object { $_.SourceEmailAddress -eq $unifiedGroup.PrimarySmtpAddress }).DestinationEmailAddress 

                            if (-not ([string]::IsNullOrEmpty($newDstMigrationWizAddress))) {                           
                                $dstMigrationWizAddress = $newDstMigrationWizAddress.split("@")[0] + "@" + $script:destinationTenantDomain 
                            }
                            else {
                                $wrongSourceEmailAddressInCSV = ($script:emailAddressMappingCSVFile | Where-Object { ($_.SourceEmailAddress).split("@")[0] -eq $unifiedGroup.PrimarySmtpAddress.split("@")[0] }).SourceEmailAddress 

                                if (-not ([string]::IsNullOrEmpty($wrongSourceEmailAddressInCSV))) {
                                    write-host -ForegroundColor Red "      ERROR: Invalid SourceEmailAddress '$wrongSourceEmailAddressInCSV' in mapping file for retrieved '$($unifiedGroup.PrimarySmtpAddress)' mailbox."
                                }
                                else {
                                    $wrongSourceEmailAddressInCSV = "<UserNotFound>"
                                    write-host -ForegroundColor Red "      ERROR: '$wrongSourceEmailAddressInCSV' in mapping file for retrieved '$($unifiedGroup.PrimarySmtpAddress)' mailbox."
                                }

                                if ($wrongSourceEmailAddressInCSV -ne "<UserNotFound>") {
                                    $wrongSourceEmailAddressInCSVList += "$wrongSourceEmailAddressInCSV must be replaced in mapping file by retrieved $($unifiedGroup.PrimarySmtpAddress)`n"
                                }
                                else {
                                    $wrongSourceEmailAddressInCSVList += "$wrongSourceEmailAddressInCSV in mapping file for retrieved $($unifiedGroup.PrimarySmtpAddress)`n"
                                }
                                $wrongSourceEmailAddressInCSVCount += 1

                                Continue
                            } 
                        }

                        ##########################################################################################

                        # Source Vanity Domain for Output CSV file

                        if ($unifiedGroup.PrimarySmtpAddress.split("@")[1] -notlike $sourceAlreadyProccesedDomains -and $unifiedGroup.PrimarySmtpAddress.split("@")[1] -notmatch ".onmicrosoft.com") {
                            $sourceAlreadyProccesedDomains += $unifiedGroup.PrimarySmtpAddress.split("@")[1]
                        }            

                        $groupName = $unifiedGroup.Alias

                        if ($checkDestinationSpoStructure) { $result = check-DestinationO365Group -group $dstMigrationWizAddress -groupAlias $groupName }
                        if (!$result -and $checkDestinationSpoStructure) {

                            write-host 
                            $msg = "INFO: Processing Office 365 (unified) Group $UnifiedGroupsIndex/$script:unifiedGroupsCount : '$groupName' '$($unifiedGroup.PrimarySmtpAddress)'."
                            Write-Host $msg
                            Log-Write -Message $msg 

                            if ($recipientMapping.count -ne 0) {
                                $msg = "      INFO: Skipping Office 365 (unified) Group '$groupName' project creation. Office 365 (unified) group not found in destination Office 365."
                                Write-Host -ForegroundColor Red $msg
                                Log-Write -Message $msg 
                            }
                            else {
                                $msg = "      INFO: Skipping Office 365 (unified) Group '$groupName' project creation. Office 365 (unified) group not found in destination Office 365."
                                Write-Host -ForegroundColor Red $msg
                                Log-Write -Message $msg 
                            }
                            $UnifiedGroupsIndex += 1
                            Continue
                        } 

                        write-host 
                        $msg = "INFO: Processing Office 365 (unified) Group $UnifiedGroupsIndex/$script:unifiedGroupsCount : '$groupName' '$($unifiedGroup.PrimarySmtpAddress)' $srcUnifiedGroupUrl."
                        Write-Host $msg
                        Log-Write -Message $msg 
                        $UnifiedGroupsIndex += 1

                        if ($linkO365Groups) {
                            if ($checkDestinationSpoStructure) {
                                Link-O365Groups $groupName $unifiedGroup.PrimarySmtpAddress $script:SrcAdministrativeUsername $dstMigrationWizAddress $script:DstAdministrativeUsername
                            }
                            else {
                                Link-O365Groups $groupName $unifiedGroup.PrimarySmtpAddress $script:SrcAdministrativeUsername 
                            }
                        }            
    
                        $doNotSkipMailbox = $false

                        if ($groupName -eq "" -or $srcUnifiedGroupUrl -eq "" -or $dstUnifiedGroupUrl -eq "" -or $srcMigrationWizAddress -eq "" -or $dstMigrationWizAddress -eq "") {
                            $msg = "INFO: Office 365 (unified) Group '$groupName' has no SharePointSiteUrl. Skipping endpoint and Document connector creation."
                            Write-Host -ForegroundColor Red $msg
                            Log-Write -Message $msg 

                            if ($srcMigrationWizAddress -ne "" -and $dstMigrationWizAddress -ne "" ) {
                                #-and $ProcessedO365GroupSites -ge 1
                                $doNotSkipMailbox = $true
                            }
                            else {
                                $doNotSkipMailbox = $false
                                #Continue
                            }  
                        }     
            
                        #Create O365 unified Group Mailbox project    

                        if ($alreadyCreated -eq $false) {

                            #$ProjectName = "Mailbox-$groupName"
                            $mailboxProjectName = "O365Group-Mailbox-All conversations-$script:sourceTenantName"
                            $projectTypeName = "MigrationProxy.WebApi.ExchangeConfiguration"
                            $ProjectType = "Mailbox"

                            if ($script:dstGermanyCloud) { $importType = "ExchangeOnlineGermany" }
                            elseif ($script:dstUsGovernment) { $importType = "ExchangeOnlineUsGovernment" }
                            else { $importType = "ExchangeOnline2" }

                            if ($script:srcGermanyCloud) { $exportType = "ExchangeOnlineGermany" }
                            elseif ($script:srcUsGovernment) { $exportType = "ExchangeOnlineUsGovernment" }
                            else { $exportType = "ExchangeOnline2" }

                            $exportEndpointId = $global:btExportEndpointId
                            $importEndpointId = $global:btImportEndpointId
                            $exportConfiguration = New-Object -TypeName $projectTypeName -Property @{
                                "Url"                          = $srcUnifiedGroupUrl;
                                "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                                "AdministrativePassword"       = $script:sourcePlainPassword;
                                "UseAdministrativeCredentials" = $true
                            }
                            $importConfiguration = New-Object -TypeName $projectTypeName -Property @{
                                "Url"                          = $dstUnifiedGroupUrl;
                                "AdministrativeUsername"       = $script:DstAdministrativeUsername;
                                "AdministrativePassword"       = $script:destinationPlainPassword;
                                "UseAdministrativeCredentials" = $true
                            }
                            $folderFilter = "^(?!Inbox|Calendar)"

                            if ($enableModernAuth) {
                                $advancedOptions = "$modernAuth "
                                $advancedOptions += $recipientMapping
                            }
                            else {
                                $advancedOptions = $recipientMapping
                            }

                            if ($applicationPermissions) {
                                $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                            }
                            if ($UseDelegatePermission) {
                                $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                            }

                            if ($totalLines -ge 400) {
                                $maximumSimultaneousMigrations = 400
                            }
                            elseif ($totalLines -le 10) {
                                $maximumSimultaneousMigrations = 10
                            }
                            else {
                                $maximumSimultaneousMigrations = $totalLines
                            }

                            $mailboxConnectorId = $null
                            $mailboxConnectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                                -ProjectName $mailboxProjectName `
                                -ProjectType $ProjectType `
                                -importType $importType `
                                -exportType $exportType `
                                -exportEndpointId $exportEndpointId `
                                -importEndpointId $importEndpointId `
                                -exportConfiguration $exportConfiguration `
                                -importConfiguration $importConfiguration `
                                -advancedOptions $advancedOptions `
                                -folderFilter $folderFilter `
                                -ZoneRequirement $global:btZoneRequirement `
                                -maximumSimultaneousMigrations $maximumSimultaneousMigrations `
                                -updateConnector $true   

                            $alreadyCreated = $true 
                        }             
            
                        if (!$doNotSkipMailbox) {       
                            #Create O365 unified Group source endpoint

                            if ($null -eq $connectorId) {

                                #Create SPO endpoints
                                $msg = "INFO: Creating MSPC endpoints for Office 365 (unified) Group '$groupName'."
                                Write-Host $msg
                                Log-Write -Message $msg 
        
                                $exportEndpointName = "SRC-O365G-$script:sourceTenantName"
                                $endpointTypeName = "ManagementProxy.ManagementService.SharePointConfiguration"
                                $endpointType = "Office365Groups"
                                $exportConfiguration = New-Object -TypeName $endpointTypeName -Property @{
                                    "Url"                          = $srcUnifiedGroupUrl;
                                    "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                                    "AdministrativePassword"       = $script:sourcePlainPassword;
                                    "UseAdministrativeCredentials" = $true
                                }

                                [guid]$exportEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType $endpointType  -EndpointName $exportEndpointName -EndpointConfiguration $exportConfiguration -update $updateEndpoint

                                #Create O365 unified Group destination endpoint

                                if (!$UseOwnAzureStorage) {
                                    $importEndpointName = "DST-O365G-$script:sourceTenantName"
                                    #$endpointTypeName = "ManagementProxy.ManagementService.SharePointConfiguration"
                                    $endpointTypeName = "ManagementProxy.ManagementService.SharePointOnlineConfiguration"
                                    #$endpointType = = "Office365Groups"
                                    $endpointType = "SharePointOnlineAPI"
                                    $importConfiguration = New-Object -TypeName $endpointTypeName -Property @{
                                        "Url"                                = $dstUnifiedGroupUrl;
                                        "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                        "AdministrativePassword"             = $script:destinationPlainPassword;
                                        "UseAdministrativeCredentials"       = $true;
                                        "UseSharePointOnlineProvidedStorage" = $true
                                    }
                                }
                                else {
                                    $importEndpointName = "DST-O365G-$dstMigrationWizAddress"
                                    #$endpointTypeName = "ManagementProxy.ManagementService.SharePointConfiguration"
                                    $endpointTypeName = "ManagementProxy.ManagementService.SharePointOnlineConfiguration"
                                    #$endpointType = = "Office365Groups"
                                    $endpointType = "SharePointOnlineAPI"
                                    $importConfiguration = New-Object -TypeName $endpointTypeName -Property @{
                                        "Url"                                = $dstUnifiedGroupUrl;
                                        "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                        "AdministrativePassword"             = $script:destinationPlainPassword;
                                        "UseAdministrativeCredentials"       = $true;
                                        "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                        "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;    
                                        "UseSharePointOnlineProvidedStorage" = $false
                                    }
                                }
        
                                [guid]$importEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType $endpointType -EndpointName $importEndpointName -EndpointConfiguration $importConfiguration -update $updateEndpoint
        
                                #Create O365 unified Group Document project
                        
                                $docProjectName = "O365Group-Document-$script:sourceTenantName"
                                $ProjectType = "Storage"    
                                $exportType = "Office365Groups"
                                #$importType = "Office365Groups"
                                $importType = "SharePointBeta"
        
                                $exportEndpointId = $exportEndpointId
                                $importEndpointId = $importEndpointId

                                $exportTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                                $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
                                    "Url"                          = $srcUnifiedGroupUrl;
                                    "AdministrativeUsername"       = $exportEndpointData.AdministrativeUsername;
                                    "AdministrativePassword"       = $script:sourcePlainPassword;
                                    "UseAdministrativeCredentials" = $true
                                }
                                #$importTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                                $importTypeName = "MigrationProxy.WebApi.SharePointBetaConfiguration"
                                $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                    "Url"                                = $dstUnifiedGroupUrl;
                                    "AdministrativeUsername"             = $importEndpointData.AdministrativeUsername;
                                    "AdministrativePassword"             = $script:destinationPlainPassword;
                                    "UseAdministrativeCredentials"       = $true;
                                    "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                    "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;    
                                    "UseSharePointOnlineProvidedStorage" = $false
                                }

                                if ($enableModernAuth) {
                                    $AdvancedOptions = "$modernAuth InitializationTimeout=8"
                                }
                                else {
                                    $AdvancedOptions = "InitializationTimeout=8"
                                }

                                if ($applicationPermissions) {
                                    $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                                }
                                if ($UseDelegatePermission) {
                                    $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                                }

                                $connectorId = $null
                                $connectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                                    -ProjectName $docProjectName `
                                    -ProjectType $ProjectType `
                                    -importType $importType `
                                    -exportType $exportType `
                                    -exportEndpointId $exportEndpointId `
                                    -importEndpointId $importEndpointId `
                                    -exportConfiguration $exportConfiguration `
                                    -importConfiguration $importConfiguration `
                                    -advancedOptions $advancedOptions `
                                    -maximumSimultaneousMigrations 100 `
                                    -ZoneRequirement $global:btZoneRequirement `
                                    -MaxLicensesToConsume 10 `
                                    -updateConnector $true   
                            }

                            if ($connectorId -ne $null) {   

                                $ProcessedO365GroupSites += 1  

                                [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $docProjectName; ProjectType = $ProjectType; ConnectorId = $connectorId; }

                                $documentLibraries = @()
                                $documentLibraryNames = ($script:importedUnifiedGroupDocumentLibraries | Where-Object { $_.TeamSiteUrl -eq $srcUnifiedGroupUrl } | Select-Object DocumentLibraryName).DocumentLibraryName
                  
                                if ($documentLibraryNames -eq $null) { $documentLibraryNames = ($script:importedUnifiedGroupDocumentLibraries | Where-Object { $_.TeamSiteUrl -eq $dstUnifiedGroupUrl } | Select-Object DocumentLibraryName).DocumentLibraryName }

                                foreach ($documentLibraryName in $documentLibraryNames) {

                                    if ([string]::IsNullOrEmpty($documentLibraryName)) { Continue }
             
                                    $msg = "      INFO: Processing Document Library '$documentLibraryName'."
                                    Write-Host $msg
                                    Log-Write -Message $msg 

                                    if ($documentLibraryName -eq "SiteAssets") { Continue }
                                    if ($documentLibraryName -eq "Teams Wiki Data") { Continue }
                        
                                    if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                        # Renew MW ticket
                                        Connect-BitTitan 

                                        Write-Host
                                        $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                        Write-Host -ForegroundColor Magenta $msg
                                        Log-Write -Message $msg 
                                        Write-Host                    
                                    }   

                                    try {
                                        $relativeUrl = $srcUnifiedGroupUrl -replace $sSPOUrl
                                        $ImportLibrary = $relativeUrl + "/" + $documentLibraryName
                                        $ExportLibrary = $relativeUrl + "/" + $documentLibraryName
        
                                        $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
                                        if (!$result) {
                                            $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId  -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary

                                            $msg = "      SUCCESS: Document Library migration '$ExportLibrary->$ImportLibrary' added to connector." 
                                            write-Host -ForegroundColor Green $msg
                                            Log-Write -Message $msg   
                                
                                            $processedDocumentLibraries += 1                    
                                        }
                                        else {
                                            $msg = "      WARNING: Document Library migration '$ExportLibrary->$ImportLibrary' already exists in connector." 
                                            write-Host -ForegroundColor Yellow $msg
                                            Log-Write -Message $msg  

                                            $processedDocumentLibraries += 1 
                                        }                      
                                    }
                                    catch {
                                        $msg = "      ERROR: Failed to add Document Library migration '$ExportLibrary->$ImportLibrary' to connector." 
                                        write-Host -ForegroundColor Red $msg
                                        Log-Write -Message $msg     
                                    }    
                                }                     

                            }  

                            ##################################################################################################
                            #    SUBSITES
                            ##################################################################################################

                            $subsiteDocumentLibraries = ($script:importedUnifiedGroupDocumentLibraries | Where-Object { $_.isSubsite -eq $true } | Where-Object { $_.TeamSiteUrl -match $srcUnifiedGroupUrl }) 
                            if ($subsiteDocumentLibraries -eq $null) { $subsiteDocumentLibraries = ($script:importedUnifiedGroupDocumentLibraries | Where-Object { $_.isSubsite -eq $true } | Where-Object { $_.TeamSiteUrl -match $dstUnifiedGroupUrl }) }

                            $currentSubsiteDocumentLibrary = 0
                            $subsiteDocumentLibraryCount = $subsiteDocumentLibraries.Count
                            $totalSubsitesCount += $subsiteDocumentLibraryCount

                            $previousSubSiteName = ""

                            if ($subsiteDocumentLibraries -ne $null) {

                                foreach ($subsiteDocumentLibrary in $subsiteDocumentLibraries) {

                                    if ($subsiteDocumentLibrary.SubsiteNumber.SubString(0, 1) -eq 1) {
                                        if (!$alreadyProcessed) {
                                            $ProcessedSubsite = 0
                                            $alreadyProcessed = $true
                                        } 
                                    }
                                    else {
                                        $alreadyProcessed = $false
                                    }
    
                                    if ($ProcessedSubsite -ne $subsiteDocumentLibrary.SubsiteNumber.SubString(0, 1)) {
                                        $ProcessedSubsite += 1 
                                        $allProcessedSubsites += 1 
                                    }    

                                    $currentSubsiteDocumentLibrary += 1

                                    $urlparts = $($subSiteDocumentLibrary.TeamSiteUrl).split("/")
                                    $subsiteDepth = $urlparts.Count - 5
                                    $subSiteName = $urlparts[-1]
                                    $parentSubSiteUrl = ($subSiteDocumentLibrary.TeamSiteUrl -split "/$subSiteName")[0] 
                                    $subsiteUrl = $parentSubSiteUrl + "/" + $subSiteName

                                    if ($subSiteName -ne $previousSubSiteName) {
                                        $currentSubsite += 1
                                    }

                                    $srcSubsiteUrl = $subsiteUrl -replace $script:destinationTenantName, $script:sourceTenantName
                                    $dstSubsiteUrl = $subsiteUrl -replace $script:sourceTenantName, $script:destinationTenantName

                                    if ($subsiteDepth -eq 1) {  
                                        Write-Host 
                                        write-host 
                                        $msg = "INFO: Processing $currentSubsite/$script:unifiedGroupSubsitesCount Subsite '$subSiteName' level-$subsiteDepth of Office 365 Group '$groupName' with URL '$srcSubsiteUrl'."
                                        Write-Host $msg
                                        Log-Write -Message $msg 
                                    }
                                    elseif ($subsiteDepth -gt 1) {
                                        Write-Host
                                        write-host 
                                        $msg = "INFO: Processing $currentSubsite/$script:unifiedGroupSubsitesCount Subsite '$subSiteName' level-$subsiteDepth  of SubSite '$parentSubSiteUrl' with URL '$srcSubsiteUrl'."
                                        Write-Host $msg
                                        Log-Write -Message $msg 
                                    }
                        
                                    $doNotSkipMailbox = $false

                                    if ($checkDestinationSpoStructure) { $result = check-DestinationSPOSubsite -siteUrl $dstSubsiteUrl -SubSiteName $subSiteName }
                                    if (!$result -and $checkDestinationSpoStructure) {
                                        $msg = "INFO: Skipping Office 365 Group Subsite project creation. Subsite '$subSiteName' with URL '$dstSubsiteUrl' not found in destination SharePoint Online."
                                        Write-Host -ForegroundColor Red $msg
                                        Log-Write -Message $msg 

                                        Continue
                                    }                     
                
                                    if ($null -eq $connectorId) {
                                        #Create SPO endpoints
                                        $msg = "INFO: Creating MSPC endpoints for Office 365 Group Subsite '$subSiteName'."
                                        Write-Host $msg
                                        Log-Write -Message $msg 

                                        $exportEndpointName = "SRC-O365G-$script:sourceTenantName" #-Subsite$subsiteDepth-
                                        $exportType = "Office365Groups"
                                        $exportTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
                                        $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
                                            "Url"                          = $srcSubsiteUrl;
                                            "AdministrativeUsername"       = $script:SrcAdministrativeUsername;
                                            "AdministrativePassword"       = $script:sourcePlainPassword;
                                            "UseAdministrativeCredentials" = $true
                                        }
        
                                        $importEndpointName = "DST-O365G-$script:destinationTenantName" #-Subsite$subsiteDepth-
                                        $importType = "SharePointBeta"
                                        $importTypeName = "MigrationProxy.WebApi.SharePointBetaConfiguration"
                                        if (!$UseOwnAzureStorage) {
                                            $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                                "Url"                                = $dstSubsiteUrl;
                                                "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                                "AdministrativePassword"             = $script:destinationPlainPassword;
                                                "UseAdministrativeCredentials"       = $true;
                                                "UseSharePointOnlineProvidedStorage" = $true 
                                            }
                                        }
                                        else {
                                            $importConfiguration = New-Object -TypeName $importTypeName -Property @{
                                                "Url"                                = $dstSubsiteUrl;
                                                "AdministrativeUsername"             = $script:DstAdministrativeUsername;
                                                "AdministrativePassword"             = $script:destinationPlainPassword;
                                                "UseAdministrativeCredentials"       = $true;
                                                "AzureStorageAccountName"            = $AzureStorageAccountName #$importEndpointData.AzureStorageAccountName;
                                                "AzureAccountKey"                    = $AzureAccountKey #$importEndpointData.AzureAccountKey;    
                                                "UseSharePointOnlineProvidedStorage" = $false
                                            }
                                        }

                                        #Create SPO Team Sites source endpoint
                                        $exportEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType $exportType  -EndpointName $exportEndpointName -EndpointConfiguration $exportConfiguration -update $updateEndpoint
                                        #Create SPO Team Sites destination endpoint
                                        $importEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType $importType -EndpointName $importEndpointName -EndpointConfiguration $importConfiguration -update $updateEndpoint

                                        #Create SPO Team Sites Document project

                                        $ProjectName = "O365Group-Document-$script:sourceTenantName" #-Subsite$subsiteDepth-
                                        $ProjectType = "Storage"

                                        if ($enableModernAuth) {
                                            $advancedOptions = "$modernAuth InitializationTimeout=8 IncreasePathLengthLimit=1"
                                        }
                                        else {
                                            $advancedOptions = "InitializationTimeout=8 IncreasePathLengthLimit=1"
                                        }

                                        if ($applicationPermissions) {
                                            $advancedOptions = "$advancedOptions UseApplicationPermission=1"
                                        }
                                        if ($UseDelegatePermission) {
                                            $advancedOptions = "$advancedOptions UseDelegatePermission=1"    
                                        }

                                        $connectorId = $null
                                        $connectorId = Create-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId `
                                            -ProjectName $ProjectName `
                                            -ProjectType $ProjectType `
                                            -exportType $exportType `
                                            -importType $importType `
                                            -exportEndpointId $exportEndpointId `
                                            -importEndpointId $importEndpointId `
                                            -exportConfiguration $exportConfiguration `
                                            -importConfiguration $importConfiguration `
                                            -advancedOptions $advancedOptions `
                                            -maximumSimultaneousMigrations 100 `
                                            -ZoneRequirement $global:btZoneRequirement `
                                            -MaxLicensesToConsume 1 `
                                            -updateConnector $true
                                    }    

                                    if ($connectorId -ne $null) {
        
                                        $documentLibraryName = $subsiteDocumentLibrary.DocumentLibraryName
    
                                        if ([string]::IsNullOrEmpty($documentLibraryName)) { Continue }
    
                                        $msg = "      INFO: Processing $currentSubsiteDocumentLibrary/$subsiteDocumentLibraryCount Subsite Document Library '$documentLibraryName'."
                                        Write-Host $msg
                                        Log-Write -Message $msg 
    
                                        if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                            # Renew MW ticket
                                            Connect-BitTitan 
    
                                            Write-Host
                                            $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                            Write-Host -ForegroundColor Magenta $msg
                                            Log-Write -Message $msg 
                                            Write-Host                    
                                        }    
                                                            
                                        try {
                                            $relativeUrl = $srcSubsiteUrl -replace $sSPOUrl
                                            $ImportLibrary = $relativeUrl + "/" + $documentLibraryName
                                            $ExportLibrary = $relativeUrl + "/" + $documentLibraryName
    
                                            $result = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
                                            if (!$result) {
                                                $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $connectorId  -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary
    
                                                $msg = "      SUCCESS: Document Library migration '$ExportLibrary->$ImportLibrary' added to connector." 
                                                write-Host -ForegroundColor Green $msg
                                                Log-Write -Message $msg  
    
                                                $processedSubsiteDocumentLibraries += 1
                                            }
                                            else {
                                                $msg = "      WARNING: Document Library migration '$ExportLibrary->$ImportLibrary' already exists in connector." 
                                                write-Host -ForegroundColor Yellow $msg
                                                Log-Write -Message $msg  
    
                                                $processedSubsiteDocumentLibraries += 1
                                            }
    
                                        }
                                        catch {
                                            $msg = "      ERROR: Failed to add Document Library migration '$ExportLibrary->$ImportLibrary' to connector." 
                                            write-Host -ForegroundColor Red $msg
                                            Log-Write -Message $msg     
                                        }                         
    
                                        [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $ProjectName; ProjectType = $ProjectType; ConnectorId = $connectorId; }
                                    }            
                        
                                    $previousSubSiteName = $subSiteName
                                    
                                }
                            }             
                        } 

                        if ($mailboxConnectorId -ne $null) {
            
                            $msg = "      INFO: Processing Office 365 (unified) group mailbox '$srcMigrationWizAddress'."
                            Write-Host $msg
                            Log-Write -Message $msg 

                            ##########################################################################################
                    
                            #MW ticket renewal
                    
                            if ([System.DateTime]::UtcNow -ge $script:MwTicket.ExpirationDate.AddSeconds(-60)) {
                                # Renew MW ticket
                                Connect-BitTitan 

                                Write-Host
                                $msg = "      INFO: MigrationWiz ticket renewed until $($script:MwTicket.ExpirationDate)."
                                Write-Host -ForegroundColor Magenta $msg
                                Log-Write -Message $msg 
                                Write-Host                    
                            }
                            ##########################################################################################

                            $mailbox = Get-MW_Mailbox -ticket $script:MwTicket -ConnectorId $mailboxConnectorId -ImportEmailAddress $dstMigrationWizAddress
                            if (!$mailbox) {
                                $result = Add-MW_Mailbox -ticket $script:MwTicket -ConnectorId $mailboxConnectorId -ImportEmailAddress $dstMigrationWizAddress -ExportEmailAddress $srcMigrationWizAddress 
                
                                $mailboxIDs += $result.Id.Guid
                                $mailboxIDsInBatch += $result.Id.Guid
                
                                $msg = "      SUCCESS: Office 365 (unified) group migration '$srcMigrationWizAddress->$dstMigrationWizAddress' added to connector."  
                                write-Host -ForegroundColor Green $msg
                                Log-Write -Message $msg  
                
                                $processedO365GroupMailboxes += 1

                                ########################################################################
                                # If mailbox is added and was previously licensed    
                                $mspcUser = (Get-BT_CustomerEndUser -Ticket $script:ticket -OrganizationId $global:btCustomerOrganizationId -id $result.CustomerEndUserId -Environment "BT" -IsDeleted $false) 
                        
                                if (!$mspcUser) {
                                    Write-host -ForegroundColor Red "      ERROR: User '$(($result.ExportEmailAddress))' not found in MSPComplete."
                                }

                                if ($mspcUser) {
                                    $subscriptionEndDate = (Get-BT_Subscription -Ticket $script:ticket -Id $mspcuser.SubscriptionId.guid).SubscriptionEndDate

                                    if ( $mspcuser.ActiveSubscriptionId -eq "00000000-0000-0000-0000-000000000000" ) {
                                        Write-host -ForegroundColor Yellow "      WARNING: User '$($mspcuser.PrimaryEmailAddress)' does not have a subscription applied."

                                        $isUmbApplied = $false
                                    }
                                    else {
                                        Write-host -ForegroundColor Green "      SUCCESS: User '$($mspcuser.PrimaryEmailAddress)' has a subscription applied that will expire in '$subscriptionEndDate'. "

                                        $isUmbApplied = $true
                                    } 

                                    if ($isUmbApplied) {
                                        if ($result) {    			
                                            if ($changeUPN2TenantEmailAddress) {		
                                                $Result = Set-MW_Mailbox -Ticket $script:MwTicket -ConnectorId $mailboxConnectorId -mailbox $result -ExportEmailAddress $srcTenantEmailAddress
                                            }
                                            else {
                                                $Result = Set-MW_Mailbox -Ticket $script:MwTicket -ConnectorId $mailboxConnectorId -mailbox $result -ExportEmailAddress $srcTenantEmailAddress -Categories "tag-3"
                                            }
                                            Write-host -ForegroundColor Blue "      SUCCESS: Source MigrationWiz address for user '$($mspcuser.PrimaryEmailAddress)' changed to '$srcTenantEmailAddress'."
                                        }
                                        else {
                                            $msg = "      ERROR: Failed to change mailbox ExportEmailAddress to '$srcTenantEmailAddress'."
                                            Write-Host -ForegroundColor Red  $msg
                                            Log-Write -Message $msg 
                                        }
                                    }
                                }
                            }
                            else {
                                $msg = "      WARNING: Office 365 (unified) group migration '$($mailbox.ExportEmailAddress)->$dstMigrationWizAddress' already exists in connector."  
                                write-Host -ForegroundColor Yellow $msg
                                Log-Write -Message $msg  
                
                                $existingMigrationsList += "$srcMigrationWizAddress->$dstMigrationWizAddress`n"
                    
                                $existingUserMailboxesCount += 1

                                $processedO365GroupMailboxes += 1
                
                                $mailboxIDs += $mailbox.Id.Guid
                                $mailboxIDsInBatch += $mailbox.Id.Guid
                
                                $mspcUser = (Get-BT_CustomerEndUser -Ticket $script:ticket -OrganizationId $global:btCustomerOrganizationId -id $mailbox.CustomerEndUserId -Environment "BT" -IsDeleted $false) 
                    
                                if (!$mspcUser) {
                                    Write-host -ForegroundColor Red "      ERROR: User '$(($mailbox.ExportEmailAddress))' not found in MSPComplete."
                                }
                
                                if ($mspcUser) {
                                    $subscriptionEndDate = (Get-BT_Subscription -Ticket $script:ticket -Id $mspcuser.SubscriptionId.guid).SubscriptionEndDate
                
                                    if ( $mspcuser.ActiveSubscriptionId -eq "00000000-0000-0000-0000-000000000000" ) {
                                        Write-host -ForegroundColor Yellow "      WARNING: User '$($mspcuser.PrimaryEmailAddress)' does not have a subscription applied."
                
                                        $isUmbApplied = $false
                                    }
                                    else {
                                        Write-host -ForegroundColor Green "      SUCCESS: User '$($mspcuser.PrimaryEmailAddress)' has a subscription applied that will expire in '$subscriptionEndDate'. "
                
                                        $isUmbApplied = $true
                                    } 
                                }
                
                                if ($mspcUser -and !$isUmbApplied -and $script:ApplyBitTitanLicenses ) {
                
                                    Try {
                                        $subscription = Add-BT_Subscription -ticket $script:WorkgroupTicket -ReferenceEntityType CustomerEndUser -ReferenceEntityId $mspcuser.Id -ProductSkuId $productId -WorkgroupOrganizationId $global:btWorkgroupOrganizationId -ErrorAction Stop
                            
                                        $msg = "      SUCCESS: User Migration Bundle subscription assigned to MSPC User '$($mspcUser.PrimaryEmailAddress)' and migration '$srcMigrationWizAddress->$dstMigrationWizAddress'."
                                        Write-Host -ForegroundColor Green  $msg
                                        Log-Write -Message $msg 
                
                                        $changeCount += 1 
                                    }
                                    Catch {
                                        $msg = "      ERROR: Failed to assign User Migration License subscription to MSPC User '$($mspcUser.PrimaryEmailAddress)'."
                                        Write-Host -ForegroundColor Red  $msg
                                        Log-Write -Message $msg
                                        Write-Host -ForegroundColor Red $($_.Exception.Message)
                                        Log-Write -Message $($_.Exception.Message) 
                                    }
                
                                    if ($subscription -and $mailbox) {    			
                                        if ($changePrimarySmtpAddress2TenantEmailAddress) {		
                                            $Result = Set-MW_Mailbox -Ticket $script:MwTicket -ConnectorId $mailboxConnectorId -mailbox $mailbox -ExportEmailAddress $srcTenantEmailAddress
                                        }
                                        else {
                                            $Result = Set-MW_Mailbox -Ticket $script:MwTicket -ConnectorId $mailboxConnectorId -mailbox $mailbox -ExportEmailAddress $srcTenantEmailAddress -Categories "tag-3"
                                        }
                                        Write-host -ForegroundColor Blue "      SUCCESS: Source MigrationWiz address for user '$($mspcuser.PrimaryEmailAddress)' changed to '$srcTenantEmailAddress'."
                                    }
                                    else {
                                        $msg = "      ERROR: Failed to change mailbox ExportEmailAddress to '$srcTenantEmailAddress'."
                                        Write-Host -ForegroundColor Red  $msg
                                        Log-Write -Message $msg 
                                    }
                                }
                                elseif ($mspcUser -and $isUmbApplied) {
                                    if ($mailbox) {    			
                                        if ($changePrimarySmtpAddress2TenantEmailAddress) {		
                                            $Result = Set-MW_Mailbox -Ticket $script:MwTicket -ConnectorId $mailboxConnectorId -mailbox $mailbox -ExportEmailAddress $srcTenantEmailAddress
                                        }
                                        else {
                                            $Result = Set-MW_Mailbox -Ticket $script:MwTicket -ConnectorId $mailboxConnectorId -mailbox $mailbox -ExportEmailAddress $srcTenantEmailAddress -Categories "tag-3"
                                        }
                                        Write-host -ForegroundColor Blue "      SUCCESS: Source MigrationWiz address for user '$($mspcuser.PrimaryEmailAddress)' changed to '$srcTenantEmailAddress'."
                                    }
                                    else {
                                        $msg = "      ERROR: Failed to change mailbox ExportEmailAddress to '$srcTenantEmailAddress'."
                                        Write-Host -ForegroundColor Red  $msg
                                        Log-Write -Message $msg 
                                    }
                                }
                            }
                        }              
                    }

                    if ($ProcessedO365GroupSites -ne 0) {
                        write-Host
                        $msg = "SUCCESS: $ProcessedO365GroupSites out of $script:unifiedGroupsCount Office 365 Groups have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 

                        $msg = "SUCCESS: $processedO365GroupMailboxes out of $script:unifiedGroupsCount Office 365 Group mailboxes have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 

                        if ([string]::IsNullOrEmpty(($processedDocumentLibraries))) { $processedDocumentLibraries = 0 }
                        $msg = "SUCCESS: $processedDocumentLibraries out of $script:unifiedGroupDocumentLibrariesCount Office 365 Group Document Libraries have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                        write-Host
                        $msg = "SUCCESS: $allProcessedSubsites out of $script:unifiedGroupSubsitesCount Office 365 Groups Subsites have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 

                        if ([string]::IsNullOrEmpty(($allProcessedSubsites))) { $processedSubsiteDocumentLibraries = 0 }
                        $msg = "SUCCESS: $processedSubsiteDocumentLibraries out of $script:unifiedGroupSubSiteDocumentLibrariesCount Office 365 Group Subsites Document Libraries have been processed." 
                        write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                    }

                    if ($mailboxConnectorId -ne $null) {
                        [array]$MigrationWizProjectArray += New-Object PSObject -Property @{ProjectName = $mailboxProjectName; ProjectType = $ProjectType; ConnectorId = $mailboxConnectorId; }
                    }   
                }  
            }

            ##########################################################################################################################################
            #            EXPORTING MW PROJECTS
            ##########################################################################################################################################

            if ($MigrationWizProjectArray -ne $null) { 
                write-host 
                $msg = "#######################################################################################################################`
                EXPORTING PROJECTS TO CSV FILE                  `
#######################################################################################################################"
                Write-Host $msg
                Log-Write -Message "EXPORTING PROJECTS TO CSV FILE" 
                write-host 

                do {
                    try {
                        $MigrationWizProjectArray  | sort ProjectName, ConnectorId -Unique | sort ProjectType | Export-Csv -Path $script:workingDir\AllMWProjects-$script:sourceTenantName.csv -NoTypeInformation -force

                        $msg = "SUCCESS: CSV file '$script:workingDir\AllMWProjects-$script:sourceTenantName.csv' processed, exported and open."
                        Write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 
                        $msg = "INFO: This CSV file will be used by Start-MW_Office365GroupMigrations.ps1 script to automatically submit all migrations for migration."
                        Write-Host $msg
                        Log-Write -Message $msg 
                        Write-Host

                        Break
                    }
                    catch {
                        $msg = "WARNING: Close opened CSV file '$script:workingDir\AllMWProjects-$script:sourceTenantName.csv'."
                        Write-Host -ForegroundColor Yellow $msg
                        Log-Write -Message $msg
                        Write-Host

                        Start-Sleep 5
                    }
                } while ($true) 

                try {
                    #Open the CSV file
                    Start-Process -FilePath $script:workingDir\AllMWProjects-$script:sourceTenantName.csv
                }
                catch {
                    $msg = "ERROR: Failed to open '$script:workingDir\AllMWProjects-$script:sourceTenantName.csv' CSV file."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg 
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $_.Exception.Message 
                    Exit
                }
            }

    
        }

    }
    #End if($action -ne $null)
    else {
        ##END SCRIPT 
        Write-Host

        $msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
        Log-Write -Message $msg

        if ($script:sourceO365Session) {
            try {
                Write-Host "INFO: Opening directory $script:workingDir where you will find all the generated CSV files."
                Invoke-Item $script:workingDir
                Write-Host
            }
            catch {
                $msg = "ERROR: Failed to open directory '$script:workingDir'. Script will abort."
                Write-Host -ForegroundColor Red $msg
                Exit
            }

            Remove-PSSession $script:sourceO365Session
            if ($script:destinationO365Session) {
                Remove-PSSession $script:destinationO365Session
            }
        }

        Exit
    }
}
while ($true)

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg 

##END SCRIPT

