<#
.SYNOPSIS
    This script will remove all MigrationWiz projects created by Create-FileServerToOD4BMWMigrationProjects.ps1 under a MigrationWiz Customer.
    
.DESCRIPTION
    This script will remove all MigrationWiz projects created by Create-FileServerToOD4BMWMigrationProjects.ps1 under a MigrationWiz Customer.
    
.NOTES
    Author          Pablo Galan Sabugo <pablogalanscripts@gmail.com> 
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
#>

# Function to write information to the Log File
Function Log-Write {
    param
    (
        [Parameter(Mandatory=$true)]    [string]$Message
    )
    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
	Add-Content -Path $script:logFile -Value $lineItem
}


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
        $global:btTicket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction SilentlyContinue
        # Get a MW ticket
        $global:btMwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction SilentlyContinue 
    }
    catch {

        $currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
        $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll",  "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
        foreach ($moduleLocation in $moduleLocations) {
            if (Test-Path $moduleLocation) {
                Import-Module -Name $moduleLocation

                # Get a ticket and set it as default
                $global:btTicket = Get-BT_Ticket -Credentials $script:creds -SetDefault -ServiceType BitTitan -ErrorAction SilentlyContinue
                # Get a MW ticket
                $global:btMwTicket = Get-MW_Ticket -Credentials $script:creds -ErrorAction SilentlyContinue 

                if(!$global:btTicket -or !$global:btMwTicket) {
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

    if(!$global:btTicket -or !$global:btMwTicket) {
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
            $workgroupsPage = @(Get-BT_Workgroup -PageOffset $workgroupOffset -PageSize 1 -IsDeleted false -CreatedBySystemUserId $global:btTicket.SystemUserId )
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
            $workgroupsPage = @(Get-BT_Workgroup -PageOffset $workgroupOffSet -PageSize $workgroupPageSize -IsDeleted false | Where-Object  { $_.CreatedBySystemUserId -ne $global:btTicket.SystemUserId })   
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

# Function to delete all mailbox connectors created by scripts
Function Remove-MW_Connectors {

    param 
    (      
        [parameter(Mandatory=$true)] [guid]$CustomerOrganizationId,
        [parameter(Mandatory=$false)] [String]$ProjectType,
        [parameter(Mandatory=$false)] [String]$ProjectName
    )
   
    $connectorPageSize = 100
  	$connectorOffSet = 0
	$connectors = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving $projectType connectors ..."
    
    do
    {   

        if($projectType -eq "Mailbox") {
            $connectorsPage = @(Get-MW_MailboxConnector -ticket $global:btMwTicket -OrganizationId $customerOrganizationId -ProjectType "Mailbox" -PageOffset $connectorOffSet -PageSize $connectorPageSize | Sort Name)
        }
        elseif($projectType -eq "Storage"){
            $connectorsPage = @(Get-MW_MailboxConnector -ticket $global:btMwTicket -OrganizationId $customerOrganizationId -ProjectType "Storage" -PageOffset $connectorOffSet -PageSize $connectorPageSize | Sort Name)
        }
                elseif($projectType -eq "Archive"){
            $connectorsPage = @(Get-MW_MailboxConnector -ticket $global:btMwTicket -OrganizationId $customerOrganizationId -ProjectType "Archive" -PageOffset $connectorOffSet -PageSize $connectorPageSize | Sort Name)
        }

        if($connectorsPage) {
            $connectors += @($connectorsPage)
            foreach($connector in $connectorsPage) {
                Write-Progress -Activity ("Retrieving connectors (" + $connectors.Length + ")") -Status $connector.Name
            }

            $connectorOffset += $connectorPageSize
        }

    } while($connectorsPage)

    if($connectors -ne $null -and $connectors.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $connectors.Length.ToString() + " $projectType connector(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No $projectType connectors found." 
        Return
    }


    $deletedMailboxConnectorsCount = 0
    $deletedDocumentConnectorsCount = 0
    if($connectors -ne $null) {
        
        Write-Host -ForegroundColor Yellow -Object "INFO: Deleting $projectType connectors:" 

        for ($i=0; $i -lt $connectors.Length; $i++) {
            $connector = $connectors[$i]

            Try {
                if($projectType -eq "Storage"){
                    if($ProjectName -match "FS-DropBox-" -and $connector.Name -match "FS-DropBox-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:btMwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                    elseif($ProjectName -match "FS-OD4B-" -and $connector.Name -match "FS-OD4B-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:btMwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                    elseif($ProjectName -match "FS-GoogleDrive-" -and $connector.Name -match "FS-GoogleDrive-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:btMwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                    elseif($ProjectName -match "OneDrive-Document-" -and $connector.Name -match "OneDrive-Document-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:btMwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                    elseif($ProjectName -match "TeamSite-Document-" -and $connector.Name -match "TeamSite-Document-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:btMwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                    elseif($ProjectName -match "O365Group-Document-" -and $connector.Name -match "O365Group-Document-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:btMwTicket -Id $connector.Id -force -ErrorAction Stop
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                }    
                
                
                if($projectType -eq "Mailbox") {
                    if($ProjectName -match "Mailbox-All conversations" -and $connector.Name -match "Mailbox-All conversations") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:btMwTicket -Id $connector.Id  -force -ErrorAction Stop

                         Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                         $deletedMailboxConnectorsCount += 1
                    }
                    elseif($ProjectName -match "O365-Mailbox-" -and $connector.Name -match "O365-Mailbox-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:btMwTicket -Id $connector.Id  -force -ErrorAction Stop

                         Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                         $deletedMailboxConnectorsCount += 1
                    }
                    elseif($ProjectName -match "O365-RecoverableItems-" -and $connector.Name -match "O365-RecoverableItems-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:btMwTicket -Id $connector.Id  -force  -ErrorAction Stop

                         Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                         $deletedMailboxConnectorsCount += 1
                    }
                }       
                
                if($projectType -eq "Archive") {
                    if($ProjectName -match "O365-Archive-" -and $connector.Name -match "O365-Archive-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:btMwTicket -Id $connector.Id  -force -ErrorAction Stop

                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedMailboxConnectorsCount += 1
                    }
                }                      

            }
            catch {
                $msg = "ERROR: Failed to delete $projectType connector $($connector.Name)."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message   
            } 
        }

        
       if($deletedDocumentConnectorsCount -ge 1) {
            Write-Host
            Write-Host -ForegroundColor Green "SUCCESS: $deletedDocumentConnectorsCount $projectType connector(s) deleted." 
        }
        if($deletedDocumentConnectorsCount -eq 0) {
            if ($projectName -match "FS-OD4B-") {
                Write-Host -ForegroundColor Red "INFO: No $projectType connector was deleted. They were not created by Migrate-MW_AzureBlobContainerToOD4B.ps1."    
            }
            elseif($projectName -match "FS-DropBox-") {
                Write-Host -ForegroundColor Red "INFO: No $projectType connector was deleted. They were not created by Create-MW_AzureBlobContainerToDropBox.ps1."    
            }    
            elseif($projectName -match "Document-") {
                Write-Host -ForegroundColor Red "INFO: No $projectType connector was deleted. They were not created by Create-MW_Office365Groups.ps1."    
            }      
        }

    }

}

#######################################################################################################################
#                                               MAIN PROGRAM
#######################################################################################################################

#Working Directory
$script:workingDir = "C:\scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$script:workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Remove-MW_FSToOD4BConnectors.log"
$script:logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

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
                       WORKGROUP, CUSTOMER AND ENDPOINTS SELECTION              `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "WORKGROUP, CUSTOMER AND ENDPOINTS SELECTION" 

if(!$global:btCheckCustomerSelection) {
    do {
        #Select workgroup
        $global:btWorkgroupId = Select-MSPC_WorkGroup

        #Select customer
        $customer = Select-MSPC_Customer -Workgroup $global:btWorkgroupId
    }
    while ($customer -eq "-1")

    $global:btCustomerOrganizationId = $customer.OrganizationId.Guid

    $global:btCustomerTicket  = Get-BT_Ticket -Ticket $global:btTicket -OrganizationId $global:btCustomerOrganizationId #-ElevatePrivilege

    $global:btWorkgroupTicket  = Get-BT_Ticket -Ticket $global:btTicket -OrganizationId $global:btWorkgroupOrganizationId #-ElevatePrivilege
    
    $global:btCheckCustomerSelection = $true  
}
else{
    Write-Host
    $msg = "INFO: Already selected workgroup '$global:btWorkgroupId' and customer '$global:btcustomerName'."
    Write-Host -ForegroundColor Green $msg

    Write-Host
    $msg = "INFO: Exit the execution and run 'Get-Variable bt* -Scope Global | Clear-Variable' if you want to connect to different workgroups/customers."
    Write-Host -ForegroundColor Yellow $msg

}

write-host 
$msg = "#######################################################################################################################`
                       DELETING MIGRATIONWIZ PROJECTS              `
#######################################################################################################################"
Write-Host $msg
Log-Write -Message "DELETING MIGRATIONWIZ PROJECTS" 

#delete connectors
Remove-MW_Connectors -CustomerOrganizationId $global:btCustomerOrganizationId -ProjectType "Storage" -ProjectName "FS-OD4B-"

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg

##END SCRIPT
