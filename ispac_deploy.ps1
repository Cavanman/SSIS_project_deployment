Set-StrictMode -Version Latest
$InformationPreference = "Continue";

# Variables
$SSISNamespace = "Microsoft.SqlServer.Management.IntegrationServices"
$sql_server = ""
$output_folder_path = ""
$target_folder_name = ""
$project_file_path = ""
$libraries_folder_path = ""
$project_name = "your_project_name"
$environment_name = "your_environment_name"

# Create output run log file folder
$year_folder = "{0:yyyy}" -f (Get-Date)
$month_folder = "{0:MM}-{0:MMM}" -f (Get-Date)
$day_folder = "{0:dd}-{0:ddd}" -f (Get-Date)
$run_folder_path = [IO.Path]::Combine($output_folder_path, $year_folder, $month_folder, $day_folder)
if(-not (Test-Path $run_folder_path)){    
   mkdir $run_folder_path | Out-Null
}

#Create output log file
$file_name = "{0:HH-mm-ss}__{1}_run_output_log.txt" -f (Get-Date), (Get-Item $PSCommandPath ).Basename
$output_log_file = Join-Path -Path $run_folder_path -ChildPath $file_name
if (!(Test-Path $output_log_file))
{
    New-Item -itemType File -Path $output_log_file
}

# Check that .ispac file exists
Write-Information ">>> Check that .ispac file exists <<<" 6>>$output_log_file
if(-not (Test-Path $project_file_path))
{
    Write-Information ">>> $project_file_path doesn't exist in deployment folder" 6>$output_log_file 
    Break Script
}

Write-Information ">>> Loading libraries <<<" 6>$output_log_file
# Supply path to a Microsoft.SqlServer.Dmf.Common.dll file 
$dmf_common_library = Join-Path -Path $libraries_folder_path -ChildPath "Microsoft.SqlServer.Dmf.Common.dll"
 
# Load the IntegrationServices assembly
[Reflection.Assembly]::LoadFrom($dmf_common_library) #| Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Management.IntegrationServices") | Out-Null;
    

Write-Information ">>> Setting up connection to hosting SQL Server <<<" 6>>$output_log_file
# Create a connection to the server
$sqlConnectionString = "Data Source=" + $sql_server + ";Initial Catalog=master;Integrated Security=SSPI;"
#$sqlConnectionString = "Data Source=" + $sql_server + ";Initial Catalog=master;User ID=data_eng_team_user;Password=password;"
$sqlConnection = New-Object System.Data.SqlClient.SqlConnection $sqlConnectionString
 
# Create the Integration Services object
$integrationServices = New-Object $SSISNamespace".IntegrationServices" $sqlConnection
 
# Get the Integration Services catalog
$catalog = $integrationServices.Catalogs["SSISDB"]

# Create catalog if it doesn't exists
Write-Information ">>> Create catalog if it doesn't exists <<<" 6>>$output_log_file

$folder = $catalog.Folders[$target_folder_name]

if (-not $folder) {
    # Create the target folder
    $folder = New-Object $SSISNamespace".CatalogFolder" ($catalog, $target_folder_name, "Folder description")
    $folder.Create()
}

Write-Host ">>> Deploying " $project_name " project <<<"
 
# Read the project file and deploy it
[byte[]] $projectFile = [System.IO.File]::ReadAllBytes($project_file_path)
$folder.DeployProject($project_name, $projectFile)
 
Write-Information ">>> Deployment of SSIS packakgs Done <<<" 6>>$output_log_file

# >>>>>>>>>>        Environment Variables       <<<<<<<<<<<

$script:environment = $folder.Environments[$environment_name]
 
if (-not $script:environment)
{
    Write-Information "Creating environment ..." 6>>$output_log_file
    $script:environment = New-Object $SSISNamespace".EnvironmentInfo" ($folder, $environment_name, $environment_name)
    $script:environment.Create()
}

Write-Host ">>> Setting reference to Environment File <<<"
Write-Host "Environment name: " $environment_name
Write-Host "Folder name: " $folder.Name 

$script:project = $folder.Projects[$project_name]
$ref = $script:project.References[$environment_name, $folder.Name]
Write-Host "Reference Environment file: $ref"
 
if (-not $ref)
{
    # making project refer to this environment
    Write-Host "Adding environment reference to project ..."
    $script:project.References.Add($environment_name, $folder.Name) #[Environment], [Environment folder]
    $script:project.Alter()
}

Write-Host ">>> Loading functions <<<"
function Add-Modify-Project-Parameter([string]$variable, [string] $value, [bool] $sensitive)
{ 
    Write-Host ">> Running Update-Package-Parameter-Value"    #| Out-File -FilePath C:\scripts\testlog.txt -Append
    Write-Host "Variable Name: " $Variable #6>> C:\scripts\testlog.txt -Append
    Write-Host "Varible Value: " $value #6>>
    Write-Host "Sensitive: $sensitive"

    $type_code = $value.GetType()
    $var = $script:environment.Variables[$variable];
    if (-not $var)
    {
        Write-Host "Adding environment variables ..."
        #                               ([variable_name], [data_type], [value], [sensitive], [description])
        $script:environment.Variables.Add("$variable", $type_code.Name, "$value", $sensitive, "Description: $variable | Value: $value")        
    }
    else {
        $var.Type = $type_code.Name
        $var.Value = "$value"
        $var.Description = "Description: $variable | Value: $value"
        $var.Sensitive = $sensitive
        # $script:environment.Variables[$variable].Remove()
        # $script:environment.Alter()
        # #                               ([variable_name], [data_type], [value], [sensitive], [description])
        # $script:environment.Variables.Add("$variable", $type_code, "$value", $sensitive, "Description: $variable | Value: $value")
        # $script:environment.Alter()
    }
    $script:environment.Alter()

    # Set the project parameters to take the values from the environment file
    #$script:project.Parameters[$variable].Set("Literal", "$value") 
    $script:project.Parameters[$variable].Set("Referenced", $variable) 
    $script:project.Alter()
}

function Update-Package-Parameter-Value([string]$package_name, [string]$variable, [string]$value)
{
    Write-Host ">> Running Update-Package-Parameter-Value"
    Write-Host "Package Name: " $package_name
    Write-Host "Variable Name: " $Variable
    Write-Host "Varible Value: " $value
    
    $package = $script:project.Packages[$package_name]
    $package.Parameters[$variable].Set([Microsoft.SqlServer.Management.IntegrationServices.ParameterInfo+ParameterValueType]::Literal, "$value")
    $package.Alter()
}

# ******** Don't seperate your parameters with commas ','. Just leave a space ********

#>>>   DEV Environment Paramater Settings - START     <<<

# Adding Project Level parameters to the environment file
# Constructor args:             variable name           default value                         sensitivity
Add-Modify-Project-Parameter "prm_proj_one_name" "value_one" $false 
Add-Modify-Project-Parameter "prm_proj_two_name" "value_two" $false
# Add-Modify-Project-Parameter "" ""
# Add-Modify-Project-Parameter "" ""
# Add-Modify-Project-Parameter "" ""

# e.g. Update-Package-Parameter-Value "package_name.dtsx" "parameter_name" "parameter_value"
Update-Package-Parameter-Value "package_one.dtsx" "prm_pkg_one" "value_one"
Update-Package-Parameter-Value "package_one.dtsx" "prm_pkg_two" "value_two"
Update-Package-Parameter-Value "package_three.dtsx" "prm_pkg_one" "value_one"
# Update-Package-Parameter-Value "", "", ""
# Update-Package-Parameter-Value "", "", ""
# Update-Package-Parameter-Value "", "", ""
# Update-Package-Parameter-Value "", "", ""
# Update-Package-Parameter-Value "", "", ""

#>>>   DEV Environment Paramater Settings - END      <<<

#>>>   PROD Environment Paramater Settings - START     <<<

# Adding Project Level parameters to the environment file
# Constructor args:             variable name           default value                         sensitivity
# Add-Modify-Project-Parameter "prm_proj_one_name" "value_one" $false 
# Add-Modify-Project-Parameter "prm_proj_two_name" "value_two" $false
# Add-Modify-Project-Parameter "" ""
# Add-Modify-Project-Parameter "" ""
# Add-Modify-Project-Parameter "" ""

# e.g. Update-Package-Parameter-Value "package_name.dtsx" "parameter_name" "parameter_value"
# Update-Package-Parameter-Value "package_one.dtsx" "prm_pkg_one" "value_one"
# Update-Package-Parameter-Value "package_one.dtsx" "prm_pkg_two" "value_two"
# Update-Package-Parameter-Value "package_three.dtsx" "prm_pkg_one" "value_one"
# Update-Package-Parameter-Value "", "", ""
# Update-Package-Parameter-Value "", "", ""
# Update-Package-Parameter-Value "", "", ""
# Update-Package-Parameter-Value "", "", ""
# Update-Package-Parameter-Value "", "", ""
 
#>>>   PROD Environment Paramater Settings - END     <<<

# Kill connection to SSIS
$integrationServices = $null 
Write-Host "Finished SSIS deployment"
