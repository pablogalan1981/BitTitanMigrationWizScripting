<#
.SYNOPSIS
    Script to move mailboxes between MigrationWiz projects.

.DESCRIPTION
    This script will move mailboxes specified in a CSV file from a MigrationWiz project to a target project. The target project can be a project cloned by the script or an existing project
    selected from a list. The CSV file can be created by the script or an existing one by speciifying a file path.
    
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

### Function to create the working and log directories
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

### Function to write information to the Log File
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
		    # create new import file
	        $inputFileName = "users-$((Get-Date).ToString("yyyyMMddHHmmss")).csv"
            $script:inputFile = "$initialDirectory\$inputFileName"

		    #$csv = "primarySmtpAddress`r`n"
		    $file = New-Item -Path $initialDirectory -name $inputFileName -ItemType file -force #-value $csv

            $msg = "SUCCESS: Empty CSV file '$script:inputFile' created."
            Write-Host -ForegroundColor Green  $msg
            
            $msg = "WARNING: Populate the CSV file with the source email addresses you want to process in a single column and save it as`r`n         File Type: 'CSV (Comma Delimited) (.csv)'`r`n         File Name: '$script:inputFile'."
            Write-Host -ForegroundColor Yellow $msg
            

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
            
    }
    else{
        $msg = "SUCCESS: CSV file '$script:inputFile' selected."
        Write-Host -ForegroundColor Green  $msg        
    }
}

######################################################################################################################################
#                                                  BITTITAN
######################################################################################################################################

# Function to authenticate to BitTitan SDK
Function Connect-BitTitan {
    [CmdletBinding()]
    # Authenticate
    $script:creds = Get-Credential -Message "Enter BitTitan credentials"

    if(!$script:creds) {
        $msg = "ERROR: Failed to authenticate with BitTitan. Please enter valid BitTitan Credentials. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        Exit
    }
    try { 
        # Get a ticket and set it as default
        $script:ticket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction SilentlyContinue
        # Get a MW ticket
        $script:mwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction SilentlyContinue 
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

        Sleep 5

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
            $customersPage = @(Get-BT_Customer -WorkgroupId $WorkgroupId -IsDeleted False -IsArchived False -PageOffset $customerOffSet -PageSize $customerPageSize)
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC customers."
            Write-Host -ForegroundColor Red  $msg
            
            Write-Host -ForegroundColor Red $_.Exception.Message
            
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
                
                $customer=$customers[0]
                Return $customer.OrganizationId
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
                Return $Customer.OrganizationId
            }
        }
        while($true)

    }

}

### Function to source and destination connector to move mailboxes between them
Function Select-MW_SourceDestinationConnector {

    param 
    (      
        [parameter(Mandatory=$true)] [String]$customerId
    )

    #######################################
    # Display all mailbox connectors
    #######################################
    
    $connectorPageSize = 100
  	$connectorOffSet = 0
	$connectors = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving mailbox connectors ..."
    
    do
    {
        $connectorsPage = @(Get-MW_MailboxConnector -ticket $script:mwTicket -OrganizationId $customerId -PageOffset $connectorOffSet -PageSize $connectorPageSize | sort ProjectType,Name )
    
        if($connectorsPage) {
            $connectors += @($connectorsPage)
            foreach($connector in $connectorsPage) {
                Write-Progress -Activity ("Retrieving connectors (" + $connectors.Length + ")") -Status $connector.Name
            }

            $connectorOffset += $connectorPageSize
        }

    } while($connectorsPage)

    if($connectors -ne $null -and $connectors.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $connectors.Length.ToString() + " mailbox connector(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No mailbox connectors found." 
        Exit
    }

    #######################################
    # {Prompt for the mailbox connector
    #######################################
    if($connectors -ne $null) {
        

        for ($i=0; $i -lt $connectors.Length; $i++) {
            $connector = $connectors[$i]
            Write-Host -Object $i,"-",$connector.Name,"-",$connector.ProjectType
        }
        Write-Host -Object "x - Exit"
        Write-Host

        Write-Host -ForegroundColor Yellow -Object "ACTION: Select the source mailbox connector:" 

        do
        {
            $result = Read-Host -Prompt ("Select 0-" + ($connectors.Length-1) + " o x")
            if($result -eq "x") {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $connectors.Length)) {
                $script:sourceConnector=$connectors[$result]
                Break
            }
        }
        while($true)


        do {
            $confirm = (Read-Host -prompt "Do you want to clone the source project?  [Y]es or [N]o")
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        if($confirm.ToLower() -eq "y") {
            $script:targetConnector = clone-MW_Project -projectToClone $script:sourceConnector
        }
        else {

            Write-Host -ForegroundColor Yellow -Object "ACTION: Select the destination mailbox connector:" 

            do
            {
                $result = Read-Host -Prompt ("Select 0-" + ($connectors.Length-1) + " o x")
                if($result -eq "x")
                {
                    Exit
                }
                if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $connectors.Length))
                {
                    $script:targetConnector=$connectors[$result]

                    $result = compare_MW_connectors -sourceConnector $script:sourceConnector -targetConnector $script:targetConnector

                    if($result) {
                        Break
                    }
                    else {
                        Return $false
                    }
                    
                }
            }
            while($true)
        }

        Write-Host
        Write-Host -ForegroundColor yellow "ACTION: Select the CSV file with the users you want to move from one connector to another. Press <CANCEL> to create a new one."
        Get-FileName $workingDir

        return $true

    }
}

# Function to clone an existing connector under a customer
function Clone-MW_Project {

    param 
    (      
        [parameter(Mandatory=$true)] [MigrationProxy.WebApi.Entity]$projectToClone
    )

    $newId = [guid]::NewGuid()
 
    $projectToClone.Name = $projectToClone.Name + "_CLONED_$(Get-Date -Format "yyyyMMddTHHmmss")"
    $selectedProject_ImportConf = $projectToClone | select -ExpandProperty ImportConfiguration
    $selectedProject_ExportConf = $projectToClone | select -ExpandProperty ExportConfiguration

    $newProject = Add-MW_MailboxConnector -ticket $script:mwTicket -Name $projectToClone.Name -ProjectType $projectToClone.projecttype `
    -ImportType $projectToClone.ImportType -ExportConfiguration $selectedProject_ExportConf `
    -ExportType $projectToClone.ExportType -ImportConfiguration $selectedProject_ImportConf `
    -SelectedExportEndpointId $projectToClone.SelectedExportEndpointId `
    -SelectedImportEndpointId $projectToClone.SelectedImportEndpointId `
    -OrganizationId $projectToClone.OrganizationId -UserId $projectToClone.UserId `
    -ZoneRequirement $projectToClone.ZoneRequirement -MaxLicensesToConsume $projectToClone.MaxLicensesToConsume `
    -AdvancedOptions $projectToClone.AdvancedOptions -MaximumItemFailures $projectToClone.MaximumItemFailures `
    -ErrorAction Stop

    $msg = "SUCCESS: Mailbox connector '$($projectToClone.Name)' created." 
    write-Host -ForegroundColor Green $msg
    Log-Write -Message $msg 
 
    return $newProject
}

# Function to move mailboxes read from CSV file from one project to another
Function Move-MW_MailboxesInCsv() {
    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerId
    )

    $result = Select-MW_SourceDestinationConnector -customerId $customerId 

    if(!$result) { Return}

    #Read CSV file
    try {
        $mailboxesInCsv = @((import-CSV $script:inputFile | Select ExportEmailAddress -unique).ExportEmailAddress) 
        if(!$mailboxesInCsv) {$mailboxesInCsv = @((import-CSV $script:inputFile | Select ImportEmailAddress -unique).ImportEmailAddress) }
        if(!$mailboxesInCsv) {$mailboxesInCsv = @((import-CSV $script:inputFile | Select SourceEmailAddress -unique).SourceEmailAddress) }
        if(!$mailboxesInCsv) {$mailboxesInCsv = @((import-CSV $script:inputFile | Select DestinationEmailAddress -unique).DestinationEmailAddress) }
        if(!$mailboxesInCsv) {$mailboxesInCsv = @((import-CSV $script:inputFile | Select PrimarySmtpAddress -unique).PrimarySmtpAddress) }                           
        if(!$mailboxesInCsv) {$mailboxesInCsv = @(get-content $script:inputFile | where {$_ -ne "ExportEmailAddress" -and $_ -ne "ImportEmailAddress" -and $_ -ne "SourceEmailAddress" -and $_ -ne "DestinationEmailAddress" -and $_ -ne "PrimarySmtpAddress"})}
        if($mailboxesInCsv.Length -ge 1) {
            Write-Host -ForegroundColor Green "SUCCESS: $($mailboxesInCsv.Length) migrations imported." 
        }
        else {
            Write-Host -ForegroundColor Red "ERROR: $($mailboxesInCsv.Length) migrations imported." 
        }
    }
    catch {
        $msg = "ERROR: Failed to import the CSV file '$script:inputFile'."
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $msg 
        Log-Write -Message $_.Exception.Message
    }     


    Write-Host
    $msg = "INFO: Processing mailboxes..." 
    Write-Host -ForegroundColor Gray $msg
    Log-Write -Message $msg 

    foreach($mailboxInCsv in $mailboxesInCsv) {

        $mailbox = @(Get-MW_Mailbox -Ticket $script:mwTicket -ConnectorId $sourceConnector.Id -ExportEmailAddress $mailboxinCsv)

        if ($mailbox) {

            if($mailbox.count -gt 1) {
                
                $msg = "WARNING: Mailbox '$mailboxinCsv' appears $($mailbox.count) times in the CSV file." 
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg 
            }

            foreach($item in $mailbox) {

	            $lastMigrationAttempt = Get-MW_MailboxMigration -ticket $script:mwTicket -MailboxId $item.Id | Select -First 1

                if($lastMigrationAttempt -eq $null -or $lastMigrationAttempt.Status -eq "Failed" -or $lastMigrationAttempt.Status -eq "Completed" -or $lastMigrationAttempt.Status -eq "Stopped") {
                    Try{
                        $result = Set-MW_Mailbox -Ticket $script:mwTicket -mailbox $item -ConnectorId $targetConnector.Id -ErrorAction Stop
            
                        if($lastMigrationAttempt -eq $null) {
                        $msg = "SUCCESS: Mailbox $($item.ExportEmailAddress) in status 'NotSubmitted' moved to the target connector '$($targetConnector.Name)'." 
                        }
                        else {
                        $msg = "SUCCESS: Mailbox $($item.ExportEmailAddress) in status '$($lastMigrationAttempt.Status)' moved to the target connector '$($targetConnector.Name)'." 
                        }
                        Write-Host -ForegroundColor Green $msg
                        Log-Write -Message $msg 

                    }
                    Catch{
                        $msg = "ERROR: Failed to move mailbox $($item.ExportEmailAddress) to the target connector '$($targetConnector.Name)'." 

                        Write-Host -ForegroundColor Red $msg 
                        Write-Host -ForegroundColor Red $_.Exception.Message
                        Log-Write -Message $msg   
                        Log-Write -Message $_.Exception.Message          
                    }
                }
            }
        }
        else {

            $msg = "ERROR: Mailbox $mailboxinCsv not found." 
            Write-Host -ForegroundColor Red $msg
            Log-Write -Message $msg 
        }


    }

}

# Function to compare 2 existing connectors under the same customer
Function Compare_MW_connectors {
    param 
    (      
        [parameter(Mandatory=$true)] [MigrationProxy.WebApi.Entity]$sourceConnector,
        [parameter(Mandatory=$true)] [MigrationProxy.WebApi.Entity]$targetConnector

    )

    $sourceProjectType = $sourceConnector.projecttype
    $sourceImportType = $sourceConnector.ImportType
    $sourceExportType = $sourceConnector.ExportType
    $targetProjectType = $targetConnector.projecttype
    $targetImportType = $targetConnector.ImportType
    $targetExportType = $targetConnector.ExportType


    if(($sourceProjectType -eq $targetProjectType) -and ($sourceImportType -eq $targetImportType)  -and ($sourceExportType -eq $targetExportType)) {
        Return $true
    }
    else{
        $msg = "ERROR: Target connector type ($targetProjectType : $targetExportType->$targetImportType) does not match with source connector type ($sourceProjectType : $sourceExportType->$sourceImportType)."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg 
        Return $false
    }    
 }

######################################################################################################################################
#                                               MAIN PROGRAM
######################################################################################################################################

Import-MigrationWizModule

#Working Directory
$workingDir = "C:\scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format "yyyyMMddTHHmmss")_Move-MW_Migrations_Between_Projects_From_CSVFile.log"
$logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

write-host 
$msg = "####################################################################################################`
                       CONNECTION TO YOUR BITTITAN ACCOUNT                  `
####################################################################################################"
Write-Host $msg
write-host 

Connect-BitTitan

write-host 
$msg = "####################################################################################################`
                       WORKGROUP AND CUSTOMER SELECTION             `
####################################################################################################"
Write-Host $msg

#Select workgroup
$WorkgroupId = Select-MSPC_WorkGroup

#Create a ticket for project sharing
$script:mwTicket = Get-MW_Ticket -Credentials $script:creds -WorkgroupId $workgroupId -IncludeSharedProjects 

#Select customer
$customerOrganizationId = Select-MSPC_Customer -Workgroup $WorkgroupId

do {
write-host 
$msg = "####################################################################################################`
                       MOVE MAILBOXES BETWEEN PROJECTS             `
####################################################################################################"
Write-Host $msg
    Move-MW_MailboxesInCsv -customerId $customerOrganizationId
}while ($true)


$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"


##END SCRIPT
