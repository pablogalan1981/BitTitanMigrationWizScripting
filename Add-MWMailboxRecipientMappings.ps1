<#

.SYNOPSIS
    Script to add automatically all RecipientMapping's to a MigrationWiz Mailbox project read from a CSV file with column names: 'SourceEmailAddress','DestinationEmailAddress'.

.DESCRIPTION
    This script will export the migration line items under the selected project or for all projects to a CSV file for you to review.  
    After that you will be able to change the migration line items just by replacing the corresponding values under the columns with 'New' prefix.
    
.NOTES
    Author          Pablo Galan Sabugo <pablogalanscripts@gmail.com>
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
#>

###################################################################################################################
#                                              HELPER FUNCTIONS                                                                                  
###################################################################################################################

# Function to check is BitTitan PowerShell SDK is installed
function Import-MigrationWizModule {
    if (((Get-Module -Name "BitTitanPowerShell") -ne $null) -or ((Get-InstalledModule -Name "BitTitanManagement" -ErrorAction SilentlyContinue) -ne $null)) {
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

    Sleep 5

    $url = "https://www.bittitan.com/downloads/bittitanpowershellsetup.msi " 
    $result= Start-Process $url
    Exit

}

# Function to create the working and log directories
Function Create-Working-Directory {    
    param 
    (
        [CmdletBinding()]
        [parameter(Mandatory=$true)] [string]$workingDir,
        [parameter(Mandatory=$true)] [string]$logDir
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

Function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $script:inputFile = $OpenFileDialog.filename

    if($OpenFileDialog.filename -eq "") {
        Return $false
    }
    else{
        $msg = "SUCCESS: CSV file '$($OpenFileDialog.filename)' selected."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
        Return $true
    }
}

###################################################################################################################
#                                                BITTITAN
###################################################################################################################

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
            $workgroupsPage = @(Get-BT_Workgroup -PageOffset $workgroupOffSet -PageSize $workgroupPageSize -IsDeleted false | where { $_.CreatedBySystemUserId -ne $script:ticket.SystemUserId })   
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
            
            if($result -eq "x")
            {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $workgroups.Length))
            {
                $Workgroup=$workgroups[$result]
                Return $Workgroup.Id
            }
        }
        while($true)

    }

}

### Function to display all customers
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
            $customersPage = @(Get-BT_Customer -WorkgroupId $WorkgroupId -IsDeleted False -IsArchived False -PageOffset $customerOffSet -PageSize $customerPageSize)
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
        Exit
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
                    if($script:confirmImpersonation -ne $null -and $script:confirmImpersonation.ToLower() -eq "y") {
                        $script:customerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId -ImpersonateId $script:mspcSystemUserId -ErrorAction Stop
                    }
                    else{
                        $script:customerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId -ErrorAction Stop
                    }
                }
                Catch{
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket." 
                }

                $script:customerName = $Customer.CompanyName

                Return $customer
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($customers.Length-1) + ", or x")
            }

            if($result -eq "x")
            {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $customers.Length))
            {
                $customer=$customers[$result]
                                
                try{
                    if($script:confirmImpersonation -ne $null -and $script:confirmImpersonation.ToLower() -eq "y") {
                        $script:customerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId -ImpersonateId $script:mspcSystemUserId -ErrorAction Stop
                    }
                    else{
                        $script:customerTicket = Get-BT_Ticket -Credentials $script:creds -OrganizationId $Customer.OrganizationId -ErrorAction Stop
                    }
                }
                Catch{
                    Write-Host -ForegroundColor Red "ERROR: Cannot create the customer ticket." 
                }

                $script:customerName = $Customer.CompanyName

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
    $sourceStorageEndpointList = @("OneDrivePro","OneDriveProAPI","SharePoint","SharePointOnlineAPI","GoogleDrive","GoogleDriveCustomerTenant","AzureFileSystem","BoxStorage"."DropBox","Office365Groups")
    $destinationStorageEndpointList = @("OneDrivePro","OneDriveProAPI","SharePoint","SharePointOnlineAPI","GoogleDrive","GoogleDriveCustomerTenant","BoxStorage"."DropBox","Office365Groups")
    $sourceArchiveEndpointList = @("ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment","GoogleVault","PstInternalStorage","Pst")
    $destinationArchiveEndpointList =  @("ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment","Gmail","IMAP","OX","WorkMail","Office365Groups","Pst")
    $sourcePublicFolderEndpointList = @("ExchangeServerPublicFolder","ExchangeOnlinePublicFolder","ExchangeOnlineUsGovernmentPublicFolder")
    $destinationPublicFolderEndpointList = @("ExchangeServerPublicFolder","ExchangeOnlinePublicFolder","ExchangeOnlineUsGovernmentPublicFolder","ExchangeServer","ExchangeOnline2","ExchangeOnlineUsGovernment")
    $sourceTeamWorkEndpointList = @("MicrosoftTeamsSource","MicrosoftTeamsSourceParallel")
    $destinationTeamWorkEndpointList = @("MicrosoftTeamsDestination","MicrosoftTeamsDestinationParallel")

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
                "TeamWork" {
                    if($exportOrImport -eq "Source") { 
                        $availableEndpoints = $sourceTeamWorkEndpointList
                    }
                    else {
                        $availableEndpoints = $destinationTeamWorkEndpointList
                    }
                } 
            }          
        }
    }

    $customerTicket = Get-BT_Ticket -OrganizationId $CustomerOrganizationId

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
            Log-Write -Message $_.Exception.Message -LogFile $global:logFile
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
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        if($confirm.ToLower() -eq "y") {
            if ($endpointName -eq "") {
                $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration 
            }
            else {
                $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
            }
            Return $endpointId
        }
    }
}

### Function to display all mailbox connectors
Function Select-MW_Connector {

    param 
    (      
        [parameter(Mandatory=$true)] [guid]$CustomerOrganizationId
    )

    write-host 
$msg = "####################################################################################################`
                       SELECT CONNECTOR(S)              `
####################################################################################################"
Write-Host $msg

    #######################################
    # Display all mailbox connectors
    #######################################
    
    $connectorPageSize = 100
  	$connectorOffSet = 0
	$script:connectors = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving Mailbox connectors ..."
    
    do {
        $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $CustomerOrganizationId -PageOffset $connectorOffSet -PageSize $connectorPageSize | where {$_.ProjectType -eq 'Mailbox'} | sort ProjectType,Name )
    
        if($connectorsPage) {
            $script:connectors += @($connectorsPage)
            foreach($connector in $connectorsPage) {
                Write-Progress -Activity ("Retrieving connectors (" + $script:connectors.Length + ")") -Status $connector.Name
            }

            $connectorOffset += $connectorPageSize
        }

    } while($connectorsPage)

    if($script:connectors -ne $null -and $script:connectors.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $script:connectors.Length.ToString() + " Mailbox connector(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No mailbox connectors found." 
        Exit
    }

    #######################################
    # {Prompt for the mailbox connector
    #######################################
    $script:allConnectors = $false

    if($script:connectors -ne $null) {       

        for ($i=0; $i -lt $script:connectors.Length; $i++)
        {
            $connector = $script:connectors[$i]
            Write-Host -Object $i,"-",$connector.Name,"-",$connector.ProjectType
        }
        Write-Host -Object "x - Exit"
        Write-Host

        Write-Host -ForegroundColor Yellow -Object "ACTION: Select the Mailbox connector:" 

        do
        {
            $result = Read-Host -Prompt ("Select 0-" + ($script:connectors.Length-1) + " o x")
            if($result -eq "x") {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $script:connectors.Length)) {
                $script:ProjectsFromCSV = $false
                $script:allConnectors = $false

                $script:connector=$script:connectors[$result]

                Return 
            }
        }
        while($true)        
    }

}

Function Import-MW_MailboxRecipientMapping {
    param 
    (      
        [parameter(Mandatory=$true)] [Object]$Connector
    )

    Write-Host                                                                   
    Write-Host "INFO: Adding RecipientMapping(s) to the Mailbox MigrationWiz project '$($connector.Name)'"
    Write-Host 

        # add items to a MigrationWiz only user mailbox project
        do {
        $confirm = (Read-Host -prompt "Would you like to import Mailbox RecipientMapping from a CSV file with 'SourceEmailAddress','DestinationEmailAddress' columns?  [Y]es or [N]o")
            if($confirm -eq "Y") {            
                $RecipientMapping = $true

                $result = Get-FileName $script:workingDir
                
	            if($script:inputFile -eq "") {
		            # create new import file
	                $script:inputFile = $env:temp + "\MigrationWiz-"
	                $script:inputFile += "import-RecipientMapping"
	                $script:inputFile += (Get-Date).ToString("yyyyMMddHHmmss")
	                $script:inputFile += ".csv"
        
                    $csv = "SourceEmailAddress,DestinationEmailAddress`r`n"

		            $file = New-Item -Path $script:inputFile -ItemType file -force -value $csv

		            # open file for editing
		            Start-Process -FilePath $script:inputFile

		            do {
			            $confirm = (Read-Host -prompt "Are you done editing the import CSV file?  [Y]es or [N]o")
		                if($confirm -eq "Y") {
			                $importConfirm = $true
                                               
		                }

		                if($confirm -eq "N") {
			                $importConfirm = $false
		                }
		            }
		            while(-not $importConfirm)
	            }
            } 
            if($confirm -eq "n") { 
                Return
            }
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        if($script:inputFile -eq "") { 
            $msg = "ERROR: No CSV file path was entered."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg 
            Return 
        }
             try {
	            # read csv file
	            $users = @(Import-Csv -Path $script:inputFile -Encoding UTF8 | where-Object { $_.PSObject.Properties.Value -ne ""} )
            }
            catch {
                $msg = "ERROR: Failed to import '$script:inputFile' CSV file. Script aborted."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg  
                Write-Host -ForegroundColor Red $($_.Exception.Message) 
                Log-Write -Message $($_.Exception.Message)  
                Exit   
            }

            if($users) {
                Write-Host -ForegroundColor Green  "SUCCESS: $script:inputFile CSV file imported with $($users.Count) RecipientMapping." 

                # Validate CSV Headers
                $CSVHeaders = @("SourceEmailAddress","DestinationEmailAddress")

                foreach ($header in $CSVHeaders) {
                    if (($users | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name') -notcontains $header  ) {
                        $msg = "ERROR: '$script:inputFile' CSV file does not have all the required column name. Required columns are: 'SourceEmailAddress','DestinationEmailAddress'."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg 
                        Return
                    }
                }                
            }
            else {
                $msg = "ERROR: '$script:inputFile' CSV file is empty."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg 
                Return
            }   

    if($RecipientMapping) {
        #Load existing advanced options
        $AdvancedOptionsString += $connector.AdvancedOptions
        $AdvancedOptionsString += "`n"

        $count=0
        $alreadyExistCount=0

        Write-Host
        Write-Host "INFO: Applying $($users.Count) RecipientMapping added to MigrationWiz Mailbox project."

        #Processing CSV into string
        foreach($user in $users) {

            $sourceEmail = $user.SourceEmailAddress
            $destinationEmail = $user.DestinationEmailAddress

            $AddressString = "$sourceEmail->$destinationEmail"
            $Sep = '"'
            $FullString = "RecipientMapping=$Sep$AddressString$Sep"

            if($AdvancedOptionsString -notmatch $FullString ) {
                $count+=1
                Write-Host -ForegroundColor Green "SUCCESS: $FullString" 

                $AdvancedOptionsString += $FullString
                $AdvancedOptionsString += "`n"
            }
            else{
                $alreadyExistCount+=1
                Write-Host -ForegroundColor Yellow "WARNING: $FullString already exists. RecipientMapping skipped." 
            }
        }

        #Adding RecipientMappings into Advancedoptions   
        Try {
            $result = Set-MW_MailboxConnector -Ticket  $mwTicket -mailboxconnector $connector -AdvancedOptions $AdvancedOptionsString -ErrorAction Stop
            Write-Host
            Write-Host -ForegroundColor Green "SUCCESS: $count RecipientMapping added to MigrationWiz mailbox project."

            if($alreadyExistCount -gt 0){
                Write-Host
                Write-Host -ForegroundColor Yellow "WARNING: $alreadyExistCount RecipientMapping were not added to MigrationWiz mailbox project because they already existed." 
            }
        }
        Catch {
            Write-Host -ForegroundColor Red "ERROR: Failed to add RecipientMapping added to MigrationWiz mailbox project advanced options."
            Write-Host -ForegroundColor Red $_.Exception.Message
            Exit
        }            
    }
}

### Function to wait for the user to press any key to continue
Function WaitForKeyPress{
    $msg = "ACTION: If you have edited and saved the CSV file then press any key to continue." 
    Write-Host $msg
    Log-Write -Message $msg
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
}

###################################################################################################################
#                                               MAIN PROGRAM
###################################################################################################################

Import-MigrationWizModule

#Working Directory
$script:workingDir = "C:\Scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$script:workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Add-MWMailboxRecipientMappings.log"
$logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $script:workingDir -logDir $logDir

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg


Write-Host 
Write-Host -ForegroundColor Yellow "WARNING: Minimal output will appear on the screen." 
Write-Host -ForegroundColor Yellow "         Please look at the log file '$($logFile)'."
Write-Host -ForegroundColor Yellow "         Generated CSV file will be in folder '$($workingDir)'."
Write-Host 
Start-Sleep -Seconds 1

write-host 
$msg = "####################################################################################################`
                       CONNECTION TO YOUR BITTITAN ACCOUNT                  `
####################################################################################################"
Write-Host $msg
Log-Write -Message "CONNECTION TO YOUR BITTITAN ACCOUNT"
write-host 

Connect-BitTitan

write-host 
$msg = "####################################################################################################`
                       WORKGROUP AND CUSTOMER SELECTION              `
####################################################################################################"
Write-Host $msg
Log-Write -Message "WORKGROUP AND CUSTOMER SELECTION"   

if(-not [string]::IsNullOrEmpty($BitTitanWorkgroupId) -and -not [string]::IsNullOrEmpty($BitTitanCustomerId)){
    $global:btWorkgroupId = $BitTitanWorkgroupId
    $global:btCustomerOrganizationId = (Get-BT_Customer | where {$_.id -eq $BitTitanCustomerId}).OrganizationId
        
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

:allProjects
do {

#Select connector
Select-MW_Connector -CustomerOrganizationId $global:btCustomerOrganizationId 

write-host 
$msg = "####################################################################################################`
                         ADD MAILBOX RECIPIENT MAPPINGS                   `
####################################################################################################"
Write-Host $msg
Log-Write -Message "MAILBOX RECIPIENT MAPPING IMPORT" 

Write-Host
Import-MW_MailboxRecipientMapping $script:connector


} while ($true)

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg

##END SCRIPT
