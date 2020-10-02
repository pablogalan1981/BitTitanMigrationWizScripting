
<#
.SYNOPSIS
    Script to automate the PST migration to Gsuite mailbox with BitTitan MigrationWiz.
    
.DESCRIPTION
    This script will create an Azure blob container in case it does not exist and a .bat file. The .bat file will have to be clicked by each end user 
    or silently executed by a GPO to automatically download the UploaderWiz agent from BitTitan server, unzip it, disconnect all PST files from 
    the Outlook profile and discover and upload all PST files to the Azure blob container. 
    
    After creating the PST batch file, the script will create the MigrationWiz personal archive project connected to the Azure Blob container 
    to migrate the PST files to GSuite.
    
.NOTES
    Author          Pablo Galan Sabugo <pablogalanscripts@gmail.com> 
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied.
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
#>

######################################################################################################################################
#                                              HELPER FUNCTIONS                                                                                  
######################################################################################################################################

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

Function Get-Directory($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null    
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.ShowDialog()| Out-Null

    if($FolderBrowser.SelectedPath -ne "") {
        $global:btOutputDir = $FolderBrowser.SelectedPath               
    }
    else{
        $global:btOutputDir = $initialDirectory
    }
    Write-Host -ForegroundColor Gray  "INFO: Directory '$global:btOutputDir' selected."
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

Function isNumeric($x) {
    $x2 = 0
    $isNum = [System.Int32]::TryParse($x, [ref]$x2)
    return $isNum
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
        
        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
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
                "AdministrativeUsername" = $azureAccountName; #Azure Storage Account Name        
                "AccessKey" = $secretKey; #Azure Storage Account SecretKey         
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
           
        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
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
                $script:AzureSubscriptionPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($script:AzureSubscriptionPassword -eq "")

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
                "AdministrativeUsername" = $adminUsername;     
                "AdministrativePassword" = $script:AzureSubscriptionPassword;         
                "SubscriptionID" = $azureSubscriptionID
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
		#####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if($endpointConfiguration -eq $null) {
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

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if($endpointConfiguration -eq $null) {
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
    elseif($endpointType -eq "ExchangeOnline2"){
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
                $script:o365AdminPassword = (Read-Host -prompt "Please enter the admin password" -AsSecureString)
            }while ($adminPassword -eq "")
        
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

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

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

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
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $oneDriveConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $adminUsername;
                "AdministrativePassword" = $adminPassword
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
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $adminUsername;
                "AdministrativePassword" = $adminPassword;
                #"AzureStorageAccountName" = $azureAccountName;
                #"AzureAccountKey" = $secretKey
                "UseSharePointOnlineProvidedStorage" = $true
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
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($adminPassword -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg
                         
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

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
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $spoConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{   
                "Url" = $Url;               
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $adminUsername;
                "AdministrativePassword" = $adminPassword;
                #"AzureStorageAccountName" = $azureAccountName;
                #"AzureAccountKey" = $secretKey
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
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

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
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

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

        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
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
                "AdministrativeUsername" = $azureAccountName;     
                "AccessKey" = $secretKey;  
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

                $global:btcustomerName = $Customer.CompanyName

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

                $global:btcustomerName = $Customer.CompanyName

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

    #####################################################################################################################
    # Display all MSPC endpoints. If $endpointType is provided, only endpoints of that type
    #####################################################################################################################

    $endpointPageSize = 100
  	$endpointOffSet = 0
	$endpoints = $null

    $sourceMailboxEndpointList = @("ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment","Gmail","IMAP","GroupWise","zimbra","OX","WorkMail","Lotus","Office365Groups")
    $destinationeMailboxEndpointList = @("ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment","Gmail","IMAP","OX","WorkMail","Office365Groups","Pst")
    $sourceStorageEndpointList = @("OneDrivePro","OneDriveProAPI","SharePoint","SharePointOnlineAPI","GoogleDrive","AzureFileSystem","BoxStorage"."DropBox","Office365Groups")
    $destinationStorageEndpointList = @("OneDrivePro","OneDriveProAPI","SharePoint","SharePointOnlineAPI","GoogleDrive","BoxStorage"."DropBox","Office365Groups")
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

    #####################################################################################################################
    # Prompt for the endpoint. If no endpoints found and endpointType provided, ask for endpoint creation
    #####################################################################################################################
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
        
        if(!$endpoint){
            $msg = "ERROR: Saved endpoint '$endpointId' has been deleted." 
            write-Host -ForegroundColor Red $msg
            Log-Write -Message $msg 
            Return -1
        }

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


 #######################################################################################################################
#                                               MAIN PROGRAM
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
###################################################################################################################

#Working Directory
$global:workingDir = "C:\scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Create-PSTBatchFile_PSTtoGSuiteMWMigrationProject.log"
$script:logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

Write-Host 
Write-Host -ForegroundColor Yellow "WARNING: Minimal output will appear on the screen." 
Write-Host -ForegroundColor Yellow "         Please look at the log file '$($logFile)'."
Write-Host -ForegroundColor Yellow "         Generated CSV file will be in folder '$($workingDir)'."
Write-Host 
Start-Sleep -Seconds 1

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg 

write-host 
$msg = "####################################################################################################`
                       CONNECTION TO YOUR BITTITAN ACCOUNT                  `
####################################################################################################"
Write-Host $msg
Log-Write -Message "CONNECTION TO YOUR BITTITAN ACCOUNT"
write-host 

Connect-BitTitan

write-host 
$msg = "#######################################################################################################################`
                       WORKGROUP AND CUSTOMER SELECTION              `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "WORKGROUP AND CUSTOMER SELECTION"   

if(-not [string]::IsNullOrEmpty($BitTitanWorkgroupId) -and -not [string]::IsNullOrEmpty($BitTitanCustomerId)){
    $global:btWorkgroupId = $BitTitanWorkgroupId
    $global:btCustomerOrganizationId = $BitTitanCustomerId
    
    Write-Host
    $msg = "INFO: Selected workgroup '$global:btWorkgroupId' and customer '$global:btCustomerOrganizationId'."
    Write-Host -ForegroundColor Green $msg
}
else{
    if(!$global:btCheckCustomerSelection) {
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
}

write-host 
$msg = "#######################################################################################################################`
                       AZURE AND PST ENDPOINT SELECTION              `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "AZURE AND PST ENDPOINT SELECTION"   
Write-Host

if([string]::IsNullOrEmpty($AzureStorageAccessKey)) {
    $msg = "INFO: Getting the connection information to the Azure Storage Account."
    Write-Host $msg
    Log-Write -Message $msg   

    $skipAzureCheck = $false
    if(!$global:btAzureCredentials) {
        #Select source endpoint
        $azureSubscriptionEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointType "AzureSubscription"
        if($azureSubscriptionEndpointId.count -gt 1){$azureSubscriptionEndpointId = $azureSubscriptionEndpointId[1]}

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
        $global:btExportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "source" -EndpointType "PST"
        if($global:btExportEndpointId.count -gt 1){$global:btExportEndpointId = $global:btExportEndpointId[1]}
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

        $msg = "INFO: Checking the Azure Blob Container 'migrationwizpst'."
        Write-Host $msg
        Log-Write -Message $msg 
        Write-Host
    
        if(!$global:btAzureStorageAccountChecked -and !$global:btAzureStorageAccount) {
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
$msg = "####################################################################################################`
                       UPLOADERWIZ EXECUTION MODE SELECTION                `
####################################################################################################"
    Write-Host $msg
    Log-Write -Message "UPLOADERWIZ EXECUTION MODE SELECTION" 
    write-host 

$useProxy = $false
do {
    $confirm = (Read-Host -prompt "Do the end user computers connect to internet through proxy ?  [Y]es or [N]o")
    if($confirm.ToLower() -eq "y") {
        $useProxy = $true    
    }

} while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

$ignorenetworkdrives = $false
do {
    $confirm = (Read-Host -prompt "Do you want to include network mapped drive scanning during the PST discovery ?  [Y]es or [N]o")
    if($confirm.ToLower() -eq "y") {
        $ignorenetworkdrives = ""  
    }
    if($confirm.ToLower() -eq "n") {
        $ignorenetworkdrives = " -ignorenetworkdrives true"
    }

} while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))


do {
    $confirm = (Read-Host -prompt "Do you want to only discover PST files and generate a PST assessment report but not upload them ?  [Y]es or [N]o")
    if($confirm.ToLower() -eq "y") {
        $onlyMetadata = $true    
    }
    if($confirm.ToLower() -eq "n") {
        $onlyMetadata = $false    
    }

} while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

do {
    $confirm = (Read-Host -prompt "Do you want to use all available bandwidth or a fraction of the bandwidth?  [a]ll or [f]raction")
    if($confirm.ToLower() -eq "a") {
        $controlBandwidth = ""   
    }
    if($confirm.ToLower() -eq "f") {  
        $controlBandwidth = " -threads 1 -noblock true"
    }

} while(($confirm.ToLower() -ne "a") -and ($confirm.ToLower() -ne "f"))

$useProxy = $false
do {
    $confirm = (Read-Host -prompt "Do you want to generate a batch file to execute UploaderWiz with additional functionality or just the UploaderWiz command line ?  [B]atch or [C]ommand")
    if($confirm.ToLower() -eq "b") {
        $generateBatchFile = $true    
        $generateCommandLine = $false 
    }
    elseif($confirm.ToLower() -eq "c") {
        $generateBatchFile = $false    
        $generateCommandLine = $true 
    }

} while(($confirm.ToLower() -ne "b") -and ($confirm.ToLower() -ne "c"))

if($generateCommandLine) {
    write-host 
$msg = "####################################################################################################`
                       UPLOADERWIZ COMMAND LINE GENERATION                `
####################################################################################################"
    Write-Host $msg
    Log-Write -Message "UPLOADERWIZ COMMAND LINE GENERATION" 
    write-host 

    if($exportEndpointData.AdministrativeUsername){$azureStorageAccountName = $exportEndpointData.AdministrativeUsername.ToLower()} else{$azureStorageAccountName = "<Fill This Field>"}
    if($exportEndpointData.AccessKey){$primaryAccessKey = $exportEndpointData.AccessKey} else{$primaryAccessKey = "<Fill This Field>"}
    if(!$container){$container = "<Fill This Field>"}

    $CH34 = '"'
    $uwMetadataCommand = "UploaderWiz.exe -type azureblobs -accesskey " + $CH34 + $azureStorageAccountName + $CH34 + " -secretkey " + $CH34 + $script:secretKey + $CH34 +" -container " + $CH34 + $container + $CH34 + " -autodiscover true -interactive false -filefilter " + $CH34 + "*.pst" + $CH34 + " -force True -command GenerateMetadata" + $ignorenetworkdrives + $controlBandwidth
    $uploaderwizCommand = "UploaderWiz.exe -type azureblobs -accesskey " + $CH34 + $azureStorageAccountName + $CH34 + " -secretkey " + $CH34 + $script:secretKey + $CH34 +" -container " + $CH34 + $container + $CH34 + " -autodiscover true -interactive false -filefilter " + $CH34 + "*.pst" + $CH34 + " -force True " + $ignorenetworkdrives + $controlBandwidth

    $msg = "INFO: Download UploaderWiz from https://api.bittitan.com/secure/downloads/UploaderWiz.zip and unzip it."
    Write-Host $msg
    Log-Write -Message $msg 
    Write-Host
   
    $msg = "INFO: The command line you need to execute UploaderWiz from a MS-DOS command prompt window."
    Write-Host $msg
    Log-Write -Message $msg 
    Write-Host

    if ($onlyMetadata)  {
        Write-Host -ForegroundColor Yellow $uwMetadataCommand
    }else{
        Write-Host -ForegroundColor Yellow $uploaderwizCommand
    }

}
elseif($generateBatchFile ){

write-host 
$msg = "####################################################################################################`
                       BATCH FILE GENERATION                `
####################################################################################################"
Write-Host $msg
Log-Write -Message "BATCH FILE GENERATION" 
write-host 

$msg = "INFO: Generating the batch file to discover and upload PST files to the Azure Blob container."
Write-Host $msg
Log-Write -Message $msg 
Write-Host

#######################################
# Get the directory
#######################################

if(-not [string]::IsNullOrEmpty($OutputPath)) {
    $global:btOutputDir = $OutputPath
    $global:btOpenCSVFile = $false

    Write-Host
    $msg = "INFO: The PST batch file reports will be placed in directory '$OutputPath'."
    Write-Host -ForegroundColor Green $msg
}
else{ 
    #output Directory
    if(!$global:btOutputDir) {
        $desktopDir = [environment]::getfolderpath("desktop")

        Write-Host
        Write-Host -ForegroundColor yellow "ACTION: Select the directory where the PST batch file will be placed in (Press cancel to use $desktopDir)"
        Get-Directory $desktopDir
    }    
    else{
        Write-Host
        $msg = "INFO: Already selected the directory '$global:btOutputDir' where the PST batch file will be placed in."
        Write-Host -ForegroundColor Green $msg

        Write-Host
        $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to place the PST batch file in another folder."
        Write-Host -ForegroundColor Yellow $msg
    }
}

$container = "migrationwizpst"

$tab = [char]9
$nl1 = "`r`n"
$nl2 = "`r`n`r`n"
$CH34 = '"'

$echoOff = "@ECHO OFF"
$echoDot = "ECHO."
$echoLine0 = "ECHO BiTitan's Automated PST File Migration"
$echoLine2 = "ECHO 2. UploaderWiz already downloaded to c:\BitTitan\UploaderWiz\UploaderWiz.exe"
$echoLine2b = "ECHO    Launching UploaderWiz desktop agent again to complete your PST migration"
$echoLine02 = "ECHO 2. Downloading and unzipping the latest version of the UploaderWiz desktop agent"
$echoLine3 = "ECHO 3. Adding Windows Registry Key to relaunch UploaderWiz upon reboot"
$echoLine4 = "ECHO 4. Executing UploaderWiz desktop agent to discover all PST files"
$echoLine5 = "ECHO 5. Executing UploaderWiz desktop agent to upload all PST files"
$echoLine6 = "ECHO 6. Removing Windows Registry Key to relaunch UploaderWiz upon reboot and batch file"
$echoLine7 = "ECHO 7. The upload of all your PST files has been completed. You can close this window"

$timeoutLine = "TIMEOUT /t 10"
$timeoutLine2 = "TIMEOUT /t 99999 /Nobreak >nul 2>&1"

$disconnectPSTs = "Powershell `"Write-host '1. Disconnecting all PST files from your Outlook.';try{`$Outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop}catch{};if(`$Outlook){`$Namespace = `$Outlook.getNamespace('MAPI');`$all_psts = `$Namespace.Stores | Where-Object {(`$_.ExchangeStoreType -eq '3') -and (`$_.FilePath -like '*.pst') -and (`$_.IsDataFileStore -eq `$true)}; ForEach (`$pst in `$all_psts){write-host 'PST file disconnected:' `$pst.FilePath;try{`$Outlook.Session.RemoveStore(`$pst.GetRootFolder())}catch{Write-Host -ForeGroundColor Red `"ERROR: Failed to disconnect PST. Please close Outlook client.`"};}}`" "

$setDir = 'SET dir=%userprofile%\AppData\Local'

$startIfClause = 'IF EXIST %dir%\BitTitan\UploaderWiz\UploaderWiz.exe' + ' ( '
$startElseClause = ' ELSE ' + ' ( '
$endClause = ")"

$createDir = 'if not exist "%dir%\BitTitan\UploaderWiz\" mkdir %dir%\BitTitan\UploaderWiz\'

$downloadUploaderWiz = 'Powershell "try{Invoke-WebRequest -Uri https://api.bittitan.com/secure/downloads/UploaderWiz.zip -OutFile %dir%\BitTitan\UploaderWiz\UploaderWiz.zip -ErrorAction Stop}catch{Write-Host -ForeGroundColor Red "ERROR: Failed to execute Invoke-WebRequest, connection closed. Are you connecting through proxy?"}"'
$downloadUploaderWizProxy = "Powershell `"`$ProxyAddress = [System.Net.WebProxy]::GetDefaultProxy().Address;[system.net.webrequest]::defaultwebproxy = New-Object system.net.webproxy(`$ProxyAddress);[system.net.webrequest]::defaultwebproxy.credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials;[system.net.webrequest]::defaultwebproxy.BypassProxyOnLocal = `$true;try{Invoke-WebRequest -ErrorAction Stop -UseBasicParsing -Proxy `$ProxyAddress -ProxyUseDefaultCredentials -Uri https://api.bittitan.com/secure/downloads/UploaderWiz.zip -OutFile %dir%\BitTitan\UploaderWiz\UploaderWiz.zip}catch{Write-Host -ForeGroundColor Red `"ERROR: Failed to execute Invoke-WebRequest, connection closed. Are you connecting through proxy?`"}`" "
if($useProxy) {
    $downloadUploaderWiz = $downloadUploaderWizProxy
}

#$unzipUploaderWiz = 'Powershell "Expand-Archive %dir%\BitTitan\UploaderWiz\UploaderWiz.zip -DestinationPath %dir%\BitTitan\UploaderWiz\ -force"' 
$unzipUploaderWiz = "Powershell.exe -nologo -noprofile -command `"& { Add-Type -A 'System.IO.Compression.FileSystem'; [IO.Compression.ZipFile]::ExtractToDirectory('%dir%\BitTitan\UploaderWiz\UploaderWiz.zip','%dir%\BitTitan\UploaderWiz\'); }`""
$copyBatchFile = "copy migrate_pst_files.bat %dir%\BitTitan\UploaderWiz\"

$queryRegKey = 'REG QUERY "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "PSTMigration" /t REG_SZ '
$addRegKey = 'IF ERRORLEVEL 1 (REG ADD "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"' + ' /v "PSTMigration" /t REG_SZ /d ' + '"' + "%dir%\BitTitan\UploaderWiz\migrate_pst_files.bat" + '")'
$deleteRegKey = 'REG DELETE "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"' + ' /v "PSTMigration" /f' 
$deleteBatchFile = 'DEL %userprofile%\Desktop\migrate_pst_files.bat'

if($exportEndpointData.AdministrativeUsername){$azureStorageAccountName = $exportEndpointData.AdministrativeUsername.ToLower()} else{$azureStorageAccountName = "<Fill This Field>"}
if($exportEndpointData.AccessKey){$primaryAccessKey = $exportEndpointData.AccessKey} else{$primaryAccessKey = "<Fill This Field>"}
if(!$container){$container = "<Fill This Field>"}

$uwMetadataCommand = "UploaderWiz.exe -type azureblobs -accesskey " + $CH34 + $azureStorageAccountName + $CH34 + " -secretkey " + $CH34 + $script:secretKey + $CH34 +" -container " + $CH34 + $container + $CH34 + " -autodiscover true -interactive false -filefilter " + $CH34 + "*.pst" + $CH34 + " -force True -command GenerateMetadata" + $ignorenetworkdrives
$uploaderwizCommand = "UploaderWiz.exe -type azureblobs -accesskey " + $CH34 + $azureStorageAccountName + $CH34 + " -secretkey " + $CH34 + $script:secretKey + $CH34 +" -container " + $CH34 + $container + $CH34 + " -autodiscover true -interactive false -filefilter " + $CH34 + "*.pst" + $CH34 + $ignorenetworkdrives

$startUploaderWiz = 'START ' + "%dir%\BitTitan\UploaderWiz\" + $uwMetadataCommand
$resumeUploaderWiz = 'START ' + "%dir%\BitTitan\UploaderWiz\" + $uploaderwizCommand

$loop1 = ':LOOP1
tasklist | find /i ' + '"UploaderWiz" >nul 2>&1
IF ERRORLEVEL 1 (
  GOTO CONTINUE1
) ELSE (
  ECHO.
  ECHO WARNING: UploaderWiz already running. It will exit
  Timeout /T 15 /Nobreak  >nul 2>&1
  EXIT
)
:CONTINUE1'

$loop2 = ':LOOP2
tasklist | find /i ' + '"UploaderWiz" >nul 2>&1
IF ERRORLEVEL 1 (
  GOTO CONTINUE2
) ELSE (
  Timeout /T 5 /Nobreak  >nul 2>&1
  GOTO LOOP2
)
:CONTINUE2'

$loop3 = ':LOOP3
tasklist | find /i ' + '"UploaderWiz" >nul 2>&1
IF ERRORLEVEL 1 (
  GOTO CONTINUE3
) ELSE (
  Timeout /T 5 /Nobreak  >nul 2>&1
  GOTO LOOP3
)
:CONTINUE3'


$batchFileCode = "$echoOff

$echoLine0

$timeoutLine
$echoDot
$disconnectPSTs

$setDir

$startIfClause

$echoDot
$echoLine2
$echoLine2b

$endClause$startElseClause

$echoDot
$echoLine02

$createDir 
$downloadUploaderWiz

$unzipUploaderWiz

$copyBatchFile

$echoDot
$echoLine3
$queryRegKey
$addRegKey

$endClause

$loop1

$timeoutLine
$echoDot
$echoLine4
$startUploaderWiz

$loop2"

if(!$onlyMetadata) {
$batchFileCode += "

$timeoutLine
$echoDot
$echoLine5
$resumeUploaderWiz

$loop3

$echoDot
$echoLine6
$deleteRegKey

$echoDot
$echoLine7
$timeoutLine2"
}
 
if(!$onlyMetadata) {
    $batFile = "$global:btOutputDir\Migrate_PST_Files.bat"
}
else{
    $batFile = "$global:btOutputDir\Discover_PST_Files.bat"    
}

try {
    Set-Content -Path $batFile -Value $batchFileCode -Encoding ASCII -ErrorAction Stop

    write-host
    $msg = "++++++++++++++++++++++++++++++++++++++++ BATCH FILE: Migrate_PST_files.bat  ++++++++++++++++++++++++++++++++++++++++`n"
    write-host $msg
    Log-Write -Message $msg 

    write-host $batchFileCode
    Log-Write -Message $batchFileCode 

    $msg = "++++++++++++++++++++++++++++++++++++++++++++++++ END BATCH FILE ++++++++++++++++++++++++++++++++++++++++++++++++++++`n"
    write-host $msg
    Log-Write -Message $msg 

    Write-Host
    $msg = "SUCCESS: Batch file '$batFile' created."
    Write-Host  -ForegroundColor Green $msg
    Log-Write -Message $msg 

    if($skipAzureCheck) {
        $msg = 'ACTION: Manually edit the batch file and provide values for -accesskey "" -secretkey "".'
        Write-Host  -ForegroundColor Yellow $msg
        Log-Write -Message $msg 
    }

    try {
        #Open the folder containing the .bat file
        Start-Process -FilePath "$global:btOutputDir\"
    }
    catch {
        $msg = "ERROR: Failed to open '$global:btOutputDir\Migrate_PST_files.bat' batch file."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message 
        Exit
    }
}
catch{
    $msg = "ERROR: Failed to create Batch file '$batFile'. You don't have permissions to create the batch file under '$global:btOutputDir\'."
    Write-Host -ForegroundColor Red  $msg
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $msg 
    Log-Write -Message $_.Exception.Message 

    Write-Host
    $msg = "ACTION: Copy the Batch file directly from the script output and create '$batFile' ."
    Write-Host -ForegroundColor Yellow  $msg
    Log-Write -Message $msg 
}

if(!$skipAzureCheck) {
    
    <#
    try {
        Write-Host
        if(!$storageAccount) {
            #Azure storage account
            $storageAccount = Check-StorageAccount -StorageAccountName $exportEndpointData.AdministrativeUsername -UserPrincipalName $global:btAzureCredentials.UserName
        }

        Write-Host
        # Azure FileShare
        $result = Check-BlobContainer -BlobContainerName "batchfile" -StorageAccount $storageAccount

        if(!$result) {
            Create-BlobContainer -BlobContainerName "batchfile" -StorageAccount $storageAccount 
        }

        # upload the batch file
        $result = Set-AzureStorageBlobContent -File $batFile -Container "batchfile" -Blob "migratePstFiles.bat" -Context $storageAccount.context -force
    }
    catch{
        $msg = "ERROR: Failed to upload Batch file '$batFile' to Azure Blob container 'batchfile'."
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $msg 
        Log-Write -Message $_.Exception.Message 

        Write-Host
        $msg = "ACTION: Copy the Batch file directly from the script output and create '$batFile' ."
        Write-Host -ForegroundColor Yellow  $msg
        Log-Write -Message $msg 
    }

    $url = Create-SASToken -BlobContainerName "batchfile" -BlobName "migratePstFiles.bat" -StorageAccount $storageAccount
        
    $msg = "SUCCESS: Batch file '$batFile' created to be sent to all end users for them to manually run it for PST file automated migration."
    Write-Host  -ForegroundColor Green $msg
    Log-Write -Message $msg 
    write-host

    $applyCustomFolderMapping = $false
    do {
        $confirm = (Read-Host -prompt "Do you want to send the .bat file to all your users via email from Office 365?  [Y]es or [N]o").trim()

        if($confirm.ToLower() -eq "y") {
            $sendBatchFile = $true        
        }

    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    if($sendBatchFile) {
    
        $url = Create-SASToken -BlobContainerName "batchfile" -BlobName "migratePstFiles.bat" -StorageAccount $storageAccount

        $users = Get-CsvFile
        Write-Host
        $msg = "ACTION: Provide your Office 365 admin credentials to send the emails."
        Write-Host -ForegroundColor Yellow $msg
        Log-Write -Message $msg 

        $smtpCreds = Get-Credential -Message "Enter Office 365 credentials to send the emails"

        Write-Host
        foreach ($user in $users) {
        
            $msg = "INFO: Sending email with .bat file to '$($user.userEmailAddress)'."
            Write-Host $msg
            Log-Write -Message $msg 

            $body = $null
            #################################################################################
            $smtpServer =  "smtp.office365.com" #"smtp.gmail.com"        
            $emailTo = $user.userEmailAddress
            $emailFrom = $smtpCreds.UserName
            $Subject = "Action required: Install the BitTitan UploaderWiz Agent on your computer."

            $body += "<tbody>"
            $body += "<center>"
            $body += "<table>"
            $body += "<tr>"
            $body += "<td align='left' valign='top'>"
            $body += "<p class='x_Logo'><a href='http://www.bittitan.com' target='_blank' rel='noopener noreferrer' data-auth='NotApplicable' title='BitTitan'><img data-imagetype='External' src='https://static.bittitan.com/Images/MSPC/MSPC_banner.png' width='600' height='50' class='x_LogoImg' alt='BitTitan' border='0'> </a></p>"
            $body += "<span style='font-family: Arial, Helvetica, sans-serif, serif, EmojiFont; font-size: 12px; color: rgb(10, 10, 10);'>"
            $body += "<p>Hello $($user.FirstName),</p>"
            $body += "<h3>Important Announcement</h3>"
            $body += "<p>We are currently planning a series of updates and improvements to our IT Services.</p>"
            $body += "<p>We are committed to creating and maintaining the best user experience with these changes. In order to do so, </br>" 
            $body += "so we will need to install an application (the BitTitan UploaderWiz Agent) that will disconnect all your PST files</br>"
            $body += "from you Outlook client and migrate them to your new Office 365 mailbox.</p>"
            $body += "<h3>Actions Required</h3>"
            $body += "<p>Complete the application installation by following these steps:</p>"
            $body += "<ol>"
            $body += "<li>Click on this link: <a href='$url' target='_blank' rel='noopener noreferrer' data-auth='NotApplicable' title='Install BitTitan UploaderWiz Agent'>Install BitTitan UploaderWiz Agent Application</a> `</li><li>Select Run. `</li></ol>"
            $body += "<p>The application will silently install. It will not impact any other work that you are doing.</p>"
            $body += "<p>We will contact you in the coming weeks if we determine that your computer requires any necessary updates.</p>"
            $body += "<hr>"
            $body += "<p>Thank you, $CustomerName IT Department</p>"
            $body += "</span></td>"
            $body += "</tr>"
            $body += "</table>"
            $body += "</td>"
            $body += "</tr>"
            $body += "</center>"
            $body += "</tbody>"

            #################################################################################

            try {

                $result = Send-MailMessage -To $emailTo -From $emailFrom -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpServer -Port 587 -Credential $smtpCreds -UseSsl -ErrorAction Stop #-Attachments $attachment 

                if ($error[0].ToString() -match "Spam abuse detected from IP range.") { 
                    #5.7.501 Access denied, spam abuse detected. The sending account has been banned due to detected spam activity. 
                    #For details, see Fix email delivery issues for error code 451 5.7.500-699 (ASxxx) in Office 365.
                    #https://support.office.com/en-us/article/fix-email-delivery-issues-for-error-code-451-4-7-500-699-asxxx-in-office-365-51356082-9fef-4639-a18a-fc7c5beae0c8 
                    $msg = "      ERROR: Failed to send email to user '$emailTo'. Access denied, spam abuse detected. The sending account has been banned. "
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg 
                }
                else {
                    $msg = "SUCCESS: Email with .bat file sent to end user '$emailTo'"
                    Write-Host -ForegroundColor Green $msg
                    Log-Write -Message $msg  
               }

            }
            catch {
                $msg = "ERROR: Failed to send email to user '$emailTo'."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $msg 
                Log-Write -Message $_.Exception.Message 
            }
        }
    }

    #>
}

}

write-host 
$msg = "####################################################################################################`
                       MIGRATIONWIZ PROJECT CREATION                 `
####################################################################################################"
Write-Host $msg
Log-Write -Message "MIGRATIONWIZ PROJECT CREATION" 
write-host 

#Create AzureFileSystem-OneDriveProAPI Document project
$msg = "INFO: Creating MigrationWiz PST to Office 365 project."
Write-Host $msg
Log-Write -Message $msg 
Write-Host

$ProjectName = "PST-Gsuite-$(Get-Date -Format yyyyMMddHHmm)"

# Export data
if(!$exportEndpointData.ContainerName){
    $containerName = "migrationwizpst"
}
else {
    $containerName = $exportEndpointData.ContainerName
}

$applyCustomFolderMapping = $false
do { 
    $confirm = (Read-Host -prompt "Do you want to migrate the PST file directly to the mailbox or to a subfolder?  [M]ailbox or [F]older")    
    if($confirm.ToLower() -eq "f") {
        $applyCustomFolderMapping = $true

        do {
            Write-host -ForegroundColor Yellow  "ACTION: Enter the destination folder name: "  -NoNewline
            $destinationFolder = Read-Host

        } while($destinationFolder -eq "")
    }
    elseif($confirm.ToLower() -eq "m") {
        $applyCustomFolderMapping = $fale
    }

} while(($confirm.ToLower() -ne "m") -and ($confirm.ToLower() -ne "f"))

$exportType = "Pst"
$exportTypeName = "MigrationProxy.WebApi.AzureConfiguration"
$exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
    "AdministrativeUsername" = $exportEndpointData.AdministrativeUsername;
    "AccessKey" = $script:secretKey;
    "ContainerName" = $containerName;
    "UseAdministrativeCredentials" = $true
}

<#
Write-Host
$msg = "INFO: Creating or selecting existing GSuite endpoint."
Write-Host $msg
Log-Write -Message $msg 

if(!$global:btImportEndpointId) {
    if([string]::IsNullOrEmpty($BitTitanDestinationEndpointId)){
        #Select destination endpoint
        $global:btImportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType "GSuite"
        if($global:btImportEndpointId.count -gt 1){$global:btImportEndpointId = $global:btImportEndpointId[1]}
    } 
    else{
        $global:btImportEndpointId = $BitTitanDestinationEndpointId
    } 
}
else{
    [PSObject]$importEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointId $global:btImportEndpointId 
    if($importEndpointData -eq -1) {
        if(!$global:btImportEndpointId){$global:btImportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType "GSuite"}
        if($global:btImportEndpointId.count -gt 1){$global:btImportEndpointId = $global:btImportEndpointId[1]}
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
$msg = "INFO: The file path to the JSON file is '$global:btJsonFilePath'."
Write-Host $msg
Log-Write -Message $msg 

$importType = "GSuite"
$importTypeName = "MigrationProxy.WebApi.GSuiteConfiguration"
$importConfiguration = New-Object -TypeName $importTypeName -Property @{
    "UseAdministrativeCredentials" = $true;
    "AdministrativeUsername" = $importEndpointData.AdministrativeUsername;
    "CredentialsFileName" = $global:btJsonFilePath;
    "Credentials" = $global:btJsonFileContent.ToString();
}
#>

if(!$global:btImportEndpointId) {
    if([string]::IsNullOrEmpty($BitTitanDestinationEndpointId)){
        #Select destination endpoint
        $global:btImportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType "Gmail"
        if($global:btImportEndpointId.count -gt 1){$global:btImportEndpointId = $global:btImportEndpointId[1]}
    } 
    else{
        $global:btImportEndpointId = $BitTitanDestinationEndpointId
    } 
}
else{
    [PSObject]$importEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointId $global:btImportEndpointId 
    if($importEndpointData -eq -1) {
        if(!$global:btImportEndpointId){$global:btImportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $global:btCustomerOrganizationId -ExportOrImport "destination" -EndpointType "Gmail"}
        if($global:btImportEndpointId.count -gt 1){$global:btImportEndpointId = $global:btImportEndpointId[1]}
    }
}

#Get source endpoint credentials
if(!$importEndpointData){[PSObject]$importEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $global:btCustomerOrganizationId -EndpointId $global:btImportEndpointId}

$importType = "Gmail"
$importTypeName = "MigrationProxy.WebApi.GoogleMailboxConfiguration"
$importConfiguration = New-Object -TypeName $importTypeName -Property @{
    "UseAdministrativeCredentials" = $true;
    "Domains" = $importEndpointData.Domains;
    "ContactHandling" = 'MigrateSuggestedContacts'
}

$ProjectType = "Archive"
$maximumSimultaneousMigrations = 400

#Double Quotation Marks
[string]$CH34=[CHAR]34
if ($applyCustomFolderMapping) {
    $folderMapping= "FolderMapping=" + $CH34 + "^->$destinationFolder/" + $CH34
}
    
$advancedOptions = "$folderMapping"

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

Write-Host -ForegroundColor Yellow "WARNING: Using Azure $ZoneRequirement Datacenter." 
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
10. France        #France.
11. SouthAfrica   #South Africa.

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
        Write-Host
	}  
} while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

Write-Host
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
-maximumSimultaneousMigrations $maximumSimultaneousMigrations `
-ZoneRequirement $ZoneRequirement `
-updateConnector $true

Write-Host
$msg = "ACTION: Click on 'Autodiscover Items' directly in MigrationWiz to import the PST files into the MigrationWiz project."
Write-Host -ForegroundColor Yellow $msg
Log-Write -Message $msg 

$url = "https://migrationwiz.bittitan.com/app/projects/$connectorId`?qp_currentWorkgroupId=$workgroupId"

$msg = "INFO: Opening '$url' in your default web browser."
Write-Host $msg
Log-Write -Message $msg 

Write-Host
$result= Start-Process $url
Start-Sleep 5
WaitForKeyPress -Message "ACTION: If you have imported the PST files into the MigrationWiz project '$ProjectName', press any key to continue"
Write-Host

$msg = "ACTION: Apply User Migration Bundle licenses to the Office 365 email addresses in MigrationWiz."
Write-Host -ForegroundColor Yellow $msg
Write-Host

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg 

##END SCRIPT
