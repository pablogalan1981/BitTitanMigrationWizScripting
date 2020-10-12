<#
.SYNOPSIS
    This script will create a MigrationWiz project to migrate FileServer Home Directories to OneDrive For Business accounts.
    It will generate a CSV file with the MigrationWiz project and all the migrations that will be used by the script 
    Start-MWMigrationsFromCSVFile.ps1 to submit all the migrations.
    
.DESCRIPTION
    This script will download the UploaderWiz exe file and execute it to create Azure blob containers per each home directory 
    found in the File Server and upload each home directory to the corresponding blob container. After that the script will 
    create the MigrationWiz projects to migrate from Azure blob containers to the OneDrive For Business accounts.
    The output of this script will be a CSV file with the projects names that will be passed to Start-MWMigrationsFromCSVFile.ps1
    to start automatically all created MigrationWiz projects. 
    
.PARAMETER WorkingDirectory
    This parameter defines the folder path where the CSV files generated during the script execution will be placed.
    This parameter is optional. If you don't specify an output folder path, the script will output the CSV files to 'C:\scripts'.  
.PARAMETER downloadLatestVersion
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
.PARAMETER HomeDirectorySeachPattern
    This parameter defines which projects you want to process, based on the project name search term. There is no limit on the number of characters you define on the search term.
    This parameter is optional. If you don't specify a project search term, all projects in the customer will be processed.
    Example: to process all projects starting with "Batch" you enter '-ProjectSearchTerm Batch'  
.PARAMETER CheckFileServer
    This parameter defines if the script must analyze the file server and remove all invalid characters both in Azure blob container and in OneDrive.which projects you want to process, based on the project name search term. There is no limit on the number of characters you define on the search term.
    This parameter is optional. If you don't specify the parameter and the true value, the file server folder and file names won´t be analyzed and invalid characters won´t be removed.
.PARAMETER CheckOneDriveAccounts
    This parameter defines if the home directory name exist as a OneDrive for Business User Principal Name prefix.
    This parameter is mandatory. If you don't specify the paramter with a true value, you have to specify a CSV file name with the home directory and OneDrive for Business mapping.
.PARAMETER MigrationWizFolderMapping
    This parameter defines if the home directory must be migrated under a destination subfolder.
    This parameter is optional. If you don't specify the parameter with the destination subfolder name, the home directory will be directrly migrated under the OneDrive.
.PARAMETER OwnAzureStorageAccount
    This parameter defines if the destination endpoint must use the custom Azure storage account specified in the endpoint.
    This parameter is optional. If you don't specify the parameter with true value, the destination endpoint will use the Microsoft provided Azure storage.
.PARAMETER ApplyUserMigrationBundle
    This parameter defines if the migration added to the MigrationWiz project must be licensed with an existing User Migration Bundle.
    This parameter is optional. If you don't specify the parameter with true value, the migration won't be automatically licensed. 
        
.NOTES
    Author          Pablo Galan Sabugo <pablogalanscripts@gmail.com> 
    Date            June/2020
    Disclaimer:     This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    Version: 1.1
    Change log:
    1.0 - Intitial Draft
#>
