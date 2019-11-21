<#	
	.NOTES
	===========================================================================
	 Created on:   	10/29/2018
	 Author:		Andrew Jimenez (asjimene) - https://github.com/asjimene/
	 Filename:     	DellBIOSUpdateAppCreator.ps1
	===========================================================================
	.DESCRIPTION
    Creates an App Model Application in SCCM for All Dell Models found in an SCCM environment
#>

$Global:SCCMSite = "##SITE:##"

# BIOS Upgrade App Desired Name
$Global:BiosAppName = "Upgrade Dell BIOS $(get-date -Format "yyyyM")"

# Parameters for the Dell BIOS Installer
$Global:InstallCmdParams = "/s /f"

# Logging and Temp Path Settings
$Global:TempDir = "C:\Temp\BIOS"
$Global:LogPath = "C:\Temp\DellBIOSUpgradeLog.log"
$Global:MaxLogSize = 1000kb

# Package Location Vars
$Global:ContentLocationRoot = "\\Path\TO\Application\Share"
$Global:IconRepo = "\\Path\To\Application Icons Share"
$Global:ApplicationIcon = "$Global:IconRepo\DellBIOSUpdater.ico"
$Global:BiosAppContentRoot = "$Global:ContentLocationRoot\$Global:BiosAppName"

# SCCM Vars
$Global:PreferredDistributionLoc = "Distribution Point Group Name"

# Define Dell Download Sources
$Global:DellDownloadList = "http://downloads.dell.com/published/Pages/index.html"
$Global:DellDownloadBase = "http://downloads.dell.com"
$Global:DellBaseURL = "http://en.community.dell.com"
$Global:Dell64BIOSUtil = "http://en.community.dell.com/techcenter/enterprise-client/w/wiki/12237.64-bit-bios-installation-utility"

# Define Dell Download Sources
$Global:DellXMLCabinetSource = "http://downloads.dell.com/catalog/DriverPackCatalog.cab"
$Global:DellCatalogSource = "http://downloads.dell.com/catalog/CatalogPC.cab"

# Define Dell Cabinet/XL Names and Paths
$Global:DellCabFile = [string]($DellXMLCabinetSource | Split-Path -Leaf)
$Global:DellCatalogFile = [string]($DellCatalogSource | Split-Path -Leaf)
$Global:DellXMLFile = $DellCabFile.Trim(".cab")
$Global:DellXMLFile = $DellXMLFile + ".xml"
$Global:DellCatalogXMLFile = $DellCatalogFile.Trim(".cab") + ".xml"

# Define Dell Global Variables
$global:DellCatalogXML = $null
$global:DellModelXML = $null
$global:DellModelCabFiles = $null


function Add-LogContent {
    param
    (
        [parameter(Mandatory = $false)]
        [switch]$Load,
        [parameter(Mandatory = $true)]
        $Content
    )
    if ($Load) {
        if ((Get-Item $LogPath).length -gt $MaxLogSize) {
            Write-Output "$(Get-Date -Format G) - $Content" > $LogPath
        }
        else {
            Write-Output "$(Get-Date -Format G) - $Content" >> $LogPath
        }
    }
    else {
        Write-Output "$(Get-Date -Format G) - $Content" >> $LogPath
    }
}


function DellBiosFinder {
    param (
        [string]$Model
    )
    if ($global:DellCatalogXML -eq $null) {
        # Read XML File
        Add-LogContent "Info: Reading Driver Pack XML File - $Global:TempDir\$DellCatalogXMLFile"
        [xml]$global:DellCatalogXML = Get-Content -Path $Global:TempDir\$DellCatalogXMLFile
		
        # Set XML Object
        $global:DellCatalogXML.GetType().FullName
    }
	
	
    # Cater for multple bios version matches and select the most recent
    $DellBIOSFile = $global:DellCatalogXML.Manifest.SoftwareComponent | Where-Object {
        ($_.name.display."#cdata-section" -match "BIOS") -and ($_.name.display."#cdata-section" -match "$model")
    } | Sort-Object ReleaseDate | Select-Object -First 1
    if ($DellBIOSFile -eq $null) {
        # Attempt to find BIOS link via Dell model number
        $DellBIOSFile = $global:DellCatalogXML.Manifest.SoftwareComponent | Where-Object {
            ($_.name.display."#cdata-section" -match "BIOS") -and ($_.name.display."#cdata-section" -match "$($model.Split(" ") | Select-Object -Last 1)")
        } | Sort-Object ReleaseDate | Select-Object -First 1
    }
    elseif ($DellBIOSFile -eq $null) {
        # Attempt to find BIOS link via Dell model number (V-Pro / Non-V-Pro Condition)
        $DellBIOSFile = $global:DellCatalogXML.Manifest.SoftwareComponent | Where-Object {
            ($_.name.display."#cdata-section" -match "BIOS") -and ($_.name.display."#cdata-section" -match "$($model.Split("-")[0])")
        } | Sort-Object ReleaseDate | Select-Object -First 1
    }
    Add-LogContent "Found BIOS for $Model at $($DellBIOSFile.Path)"
    # Return BIOS file values
    Return $DellBIOSFile
}


function New-BIOSUpgradeApplication {
    param (
        [System.String]$ApplicationName

    )

    # Create Application
    $ApplicationPublisher = "Dell Inc."
    $ApplicationDescription = "Upgrades Dell BIOS to the latest available version"
    $ApplicationDocURL = "https://downloads.dell.com"
    $ApplicationAutoInstall = $True

    ## Create the Application
    Push-Location
    Set-Location $Global:SCCMSite
    Try {
        If ($Global:ApplicationIcon -like "$Global:IconRepo\*") {
            Add-LogContent "Command: New-CMApplication -Name $ApplicationName $ApplicationSWVersion -Description $ApplicationDescription -Publisher $ApplicationPublisher -SoftwareVersion $ApplicationSWVersion -OptionalReference $ApplicationDocURL -AutoInstall $ApplicationAutoInstall -ReleaseDate (Get-Date) -LocalizedName $ApplicationName $ApplicationSWVersion -LocalizedDescription $ApplicationDescription -UserDocumentation $ApplicationDocURL -IconLocationFile"
            New-CMApplication -Name "$ApplicationName $ApplicationSWVersion" -Description "$ApplicationDescription" -Publisher "$ApplicationPublisher" -SoftwareVersion $ApplicationSWVersion -OptionalReference $ApplicationDocURL -AutoInstall $ApplicationAutoInstall -ReleaseDate (Get-Date) -LocalizedName "$ApplicationName $ApplicationSWVersion" -LocalizedDescription "$ApplicationDescription" -UserDocumentation $ApplicationDocURL -IconLocationFile "$Global:ApplicationIcon"
        }
        Else {
            Add-LogContent "Command: New-CMApplication -Name $ApplicationName $ApplicationSWVersion -Description $ApplicationDescription -Publisher $ApplicationPublisher -SoftwareVersion $ApplicationSWVersion -OptionalReference $ApplicationDocURL -AutoInstall $ApplicationAutoInstall -ReleaseDate (Get-Date) -LocalizedName $ApplicationName $ApplicationSWVersion -LocalizedDescription $ApplicationDescription -UserDocumentation"
            New-CMApplication -Name "$ApplicationName $ApplicationSWVersion" -Description "$ApplicationDescription" -Publisher "$ApplicationPublisher" -SoftwareVersion $ApplicationSWVersion -OptionalReference $ApplicationDocURL -AutoInstall $ApplicationAutoInstall -ReleaseDate (Get-Date) -LocalizedName "$ApplicationName $ApplicationSWVersion" -LocalizedDescription "$ApplicationDescription" -UserDocumentation $ApplicationDocURL
        }
    }
    Catch {
        $ErrorMessage = $_.Exception.Message
        $FullyQualified = $_.FullyQualifiedErrorID
        Add-LogContent "ERROR: Application Creation Failed!"
        Add-LogContent "ERROR: $ErrorMessage"
        Add-LogContent "ERROR: $FullyQualified"
        Add-LogContent "ERROR: $($_.CategoryInfo.Category): $($_.CategoryInfo.Reason)"
    }
    Pop-Location
}

############################################################################################
##                                            MAIN                                        ##
############################################################################################

Add-LogContent -Load "Starting Dell BIOS Application Creator"

if (-not (Get-Module ConfigurationManager)) {
    Import-Module ConfigurationManager -ErrorAction SilentlyContinue
}

if (-not (Get-Module ConfigurationManager)) {
    Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1) -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
}

## Query for Dell Models
Push-location
Set-Location $Global:SCCMSite
$WMI = @"
select distinct SMS_G_System_COMPUTER_SYSTEM.Model from  SMS_R_System inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_COMPUTER_SYSTEM.Manufacturer like "%Dell%" and (SMS_G_System_COMPUTER_SYSTEM.Model like "%Optiplex%" or SMS_G_System_COMPUTER_SYSTEM.Model like "%Precision%" or SMS_G_System_COMPUTER_SYSTEM.Model like "%XPS%" or SMS_G_System_COMPUTER_SYSTEM.Model like "%Latitude%") order by SMS_G_System_COMPUTER_SYSTEM.Model
"@

Add-LogContent "Running Query for Dell Models"
$QueryResults = Invoke-CMWmiQuery -Query $WMI -Option Lazy

## Create Global Conditions
Add-LogContent "Creating Global Conditions if they do not exist"
if (-not (Get-CMGlobalCondition -Name "AutoPackage - Computer Manufacturer")) {
    New-CMGlobalConditionWqlQuery -DataType String -Class Win32_ComputerSystem -Namespace root\cimv2 -Property Manufacturer -Name "AutoPackage - Computer Manufacturer" -Description "Returns the Manufacturer from ComputerSystem\Manufacturer"
}

if (-not (Get-CMGlobalCondition -Name "AutoPackage - Computer Model")) {
    New-CMGlobalConditionWqlQuery -DataType String -Class Win32_ComputerSystem -Namespace root\cimv2 -Property Model -Name "AutoPackage - Computer Model" -Description "Returns the Model from ComputerSystem\Model"
}
Pop-Location


$ModelList = @()
foreach ($Model in $QueryResults) {
    $Row = New-Object -TypeName System.Management.Automation.PSObject
    $Row | Add-Member -MemberType NoteProperty -Name "Model" -Value $Model.Model
    $ModelList += $Row
}

$ModelList | Export-Csv -Path "$Global:TempDir\AllDellModels.csv" -NoTypeInformation
$AllDellModels = $ModelList | Select-Object Model -ExpandProperty Model

if (-not (Test-Path $Global:TempDir -ErrorAction SilentlyContinue)){
    New-Item -ItemType Directory -Path $Global:TempDir -Force -ErrorAction SilentlyContinue
}

$ProgressPreference = "SilentlyContinue"
Invoke-WebRequest -Uri $DellCatalogSource -OutFile "$Global:TempDir\$DellCatalogFile"
if ($?) {
    Add-LogContent "Download Succeeded"
}
else {
    Add-LogContent "Download Failed"
}
Start-Sleep 1

# Expand Cabinet File
Add-LogContent "Info: Expanding Dell Driver Pack Cabinet File: $DellCatalogFile"
Expand "$Global:TempDir\$DellCatalogFile" -F:* "$Global:TempDir\$DellCatalogXMLFile" | Out-Null
if ($?) {
    Add-LogContent "Expand Succeeded"
}
else {
    Add-LogContent "Expand Failed"
}

# Create Folder Structure
Add-LogContent "Creating directory structure for $Global:BiosAppName"
if (-not(Test-path "$Global:BiosAppContentRoot" -ErrorAction SilentlyContinue)){
    New-Item -ItemType Directory -Path "$Global:BiosAppContentRoot" -Force -ErrorAction SilentlyContinue
}

if (-not(Test-path "$Global:BiosAppContentRoot\Packages" -ErrorAction SilentlyContinue)) {
    New-Item -ItemType Directory -Path "$Global:BiosAppContentRoot\Packages" -Force -ErrorAction SilentlyContinue 
}

Push-Location
Set-location $Global:SCCMSite
if (-not (Get-CMApplication "$Global:BiosAppName")) {
    # Create the Application
    New-BIOSUpgradeApplication -ApplicationName $Global:BiosAppName
}
Pop-Location

Foreach ($DellModel in $AllDellModels) {
    Add-LogContent "Processing Model $DellModel"
    $Manufacturer = "Dell Inc."
    
    # Download Driver and Copy to Share
    $DellBIOSDownload = DellBiosFinder -Model $DellModel
    $DellBIOSVersion = $DellBIOSDownload.DellVersion
    $BIOSDownload = $DellDownloadBase + "/" + $($DellBIOSDownload.Path)
    $BIOSName = "$($DellModel.Replace(' ', '_'))_$DellBIOSVersion.exe"
    if (-not (Test-Path "$DestDir\$BIOSName")) {
        Add-LogContent "$BIOSDownload"
        Invoke-WebRequest -Uri "$BIOSDownload" -OutFile "$Global:TempDir\$BIOSName"
    }

    
    $ModelSpecificContentLoc = "$Global:BiosAppContentRoot\Packages\$DellModel"
    Add-LogContent "Content for $DellModel will be located at: $ModelSpecificContentLoc"
    if (-not (Test-Path $ModelSpecificContentLoc -ErrorAction SilentlyContinue)){
        New-Item -ItemType Directory -Path $ModelSpecificContentLoc -Force -ErrorAction SilentlyContinue
    }

    # Check if File Already Exists, continue with packaging if it does not
    if ((-not(Test-Path "$ModelSpecificContentLoc\$BIOSName")) -or ($BIOSName -notlike "*_.exe")) {
        # Copy the BIOS File to the Destination
        Add-logcontent "Copying $BiosName to $ModelSpecificContentLoc\$BIOSName"
        Copy-Item -Path "$Global:TempDir\$BIOSName" -Destination "$ModelSpecificContentLoc\$BIOSName" -Force -ErrorAction SilentlyContinue
        
        # Create Detection Method
        Push-Location
        Set-Location $Global:SCCMSite
        $DetectionClause = New-CMDetectionClauseRegistryKeyValue -Hive "LocalMachine" -ExpressionOperator IsEquals -PropertyType "String" -Value -KeyName "HARDWARE\DESCRIPTION\System\BIOS" -ValueName "BIOSVersion" -ExpectedValue "$DellBIOSVersion"

        #Create Deployment Type
        $ModelSpecificInstallCmd = "$BIOSName $Global:InstallCmdParams" 
        Add-Logcontent "Adding Deployment Type for $DellModel to $Global:BiosAppName"
        Add-CMScriptDeploymentType -ApplicationName $Global:BiosAppName -DeploymentTypeName $DellModel -ContentLocation "$ModelSpecificContentLoc" -InstallCommand "$ModelSpecificInstallCmd" -AddDetectionClause $DetectionClause -EstimatedRuntimeMins "15" -MaximumRuntimeMins "45" -ContentFallback -EnableBranchCache -LogonRequirementType WhetherOrNotUserLoggedOn -UserInteractionMode Hidden -SlowNetworkDeploymentMode Download -InstallationBehaviorType InstallForSystem -RebootBehavior BasedOnExitCode

        # Add Manufacturer Query to Deployment Type
        Add-LogContent "Adding manufacturer rule to $BIOSName"
        Add-LogContent "`"$Manufacturer`" is being added"
        $rule = Get-CMGlobalCondition -Name "AutoPackage - Computer Manufacturer" | New-CMRequirementRuleCommonValue -Value1 "$Manufacturer" -RuleOperator IsEquals 
        $rule.Name = "AutoPackage - Computer Manufacturer Equals $Manufacturer"
        Set-CMScriptDeploymentType -ApplicationName $Global:BiosAppName -DeploymentTypeName $DellModel -AddRequirement $rule

        # Add Model Query to Deployment Type
        Add-LogContent "Adding model rule to $BIOSName"
        Add-LogContent "`"$DellModel`" is being added"
        $rule = Get-CMGlobalCondition -Name "AutoPackage - Computer Model" | New-CMRequirementRuleCommonValue -Value1 "$DellModel" -RuleOperator IsEquals 
        $rule.Name = "AutoPackage - Computer Model Equals $DellModel"
        Set-CMScriptDeploymentType -ApplicationName $Global:BiosAppName -DeploymentTypeName $DellModel -AddRequirement $rule
        Pop-Location
    }
    Else {
        Add-LogContent "ERROR: Downloading BIOS File for $DellModel has Failed!"

    }
}

Add-LogContent "Distributing Content for $Global:BiosAppName"
Try {
    Push-Location
    Set-Location $Global:SCCMSite
    Start-CMContentDistribution -ApplicationName "$Global:BiosAppName" -DistributionPointGroupName $Global:PreferredDistributionLoc -ErrorAction Stop
    Pop-Location
}
Catch {
    $ErrorMessage = $_.Exception.Message
    Add-LogContent "ERROR: Content Distribution Failed!"
    Add-LogContent "ERROR: $ErrorMessage"
}

function Add-LogContent {
    param
    (
        [parameter(Mandatory = $false)]
        [switch]$Load,
        [parameter(Mandatory = $true)]
        $Content
    )
    if ($Load) {
        if ((Get-Item $LogPath).length -gt $MaxLogSize) {
            Write-Output "$(Get-Date -Format G) - $Content" > $LogPath
        }
        else {
            Write-Output "$(Get-Date -Format G) - $Content" >> $LogPath
        }
    }
    else {
        Write-Output "$(Get-Date -Format G) - $Content" >> $LogPath
    }
}


function DellBiosFinder {
    param (
        [string]$Model
    )
    if ($global:DellCatalogXML -eq $null) {
        # Read XML File
        Add-LogContent "Info: Reading Driver Pack XML File - $Global:TempDir\$DellCatalogXMLFile"
        [xml]$global:DellCatalogXML = Get-Content -Path $Global:TempDir\$DellCatalogXMLFile
		
        # Set XML Object
        $global:DellCatalogXML.GetType().FullName
    }
	
	
    # Cater for multple bios version matches and select the most recent
    $DellBIOSFile = $global:DellCatalogXML.Manifest.SoftwareComponent | Where-Object {
        ($_.name.display."#cdata-section" -match "BIOS") -and ($_.name.display."#cdata-section" -match "$model")
    } | Sort-Object ReleaseDate | Select-Object -First 1
    if ($DellBIOSFile -eq $null) {
        # Attempt to find BIOS link via Dell model number
        $DellBIOSFile = $global:DellCatalogXML.Manifest.SoftwareComponent | Where-Object {
            ($_.name.display."#cdata-section" -match "BIOS") -and ($_.name.display."#cdata-section" -match "$($model.Split(" ") | Select-Object -Last 1)")
        } | Sort-Object ReleaseDate | Select-Object -First 1
    }
    elseif ($DellBIOSFile -eq $null) {
        # Attempt to find BIOS link via Dell model number (V-Pro / Non-V-Pro Condition)
        $DellBIOSFile = $global:DellCatalogXML.Manifest.SoftwareComponent | Where-Object {
            ($_.name.display."#cdata-section" -match "BIOS") -and ($_.name.display."#cdata-section" -match "$($model.Split("-")[0])")
        } | Sort-Object ReleaseDate | Select-Object -First 1
    }
    Add-LogContent "Found BIOS for $Model at $($DellBIOSFile.Path)"
    # Return BIOS file values
    Return $DellBIOSFile
}


function New-BIOSUpgradeApplication {
    param (
        [System.String]$ApplicationName

    )

    # Create Application
    $ApplicationPublisher = "Dell Inc."
    $ApplicationDescription = "Upgrades Dell BIOS to the latest available version"
    $ApplicationDocURL = "https://downloads.dell.com"
    $ApplicationAutoInstall = $True

    ## Create the Application
    Push-Location
    Set-Location $Global:SCCMSite
    Try {
        If ($Global:ApplicationIcon -like "$Global:IconRepo\*") {
            Add-LogContent "Command: New-CMApplication -Name $ApplicationName $ApplicationSWVersion -Description $ApplicationDescription -Publisher $ApplicationPublisher -SoftwareVersion $ApplicationSWVersion -OptionalReference $ApplicationDocURL -AutoInstall $ApplicationAutoInstall -ReleaseDate (Get-Date) -LocalizedName $ApplicationName $ApplicationSWVersion -LocalizedDescription $ApplicationDescription -UserDocumentation $ApplicationDocURL -IconLocationFile"
            New-CMApplication -Name "$ApplicationName $ApplicationSWVersion" -Description "$ApplicationDescription" -Publisher "$ApplicationPublisher" -SoftwareVersion $ApplicationSWVersion -OptionalReference $ApplicationDocURL -AutoInstall $ApplicationAutoInstall -ReleaseDate (Get-Date) -LocalizedName "$ApplicationName $ApplicationSWVersion" -LocalizedDescription "$ApplicationDescription" -UserDocumentation $ApplicationDocURL -IconLocationFile "$Global:ApplicationIcon"
        }
        Else {
            Add-LogContent "Command: New-CMApplication -Name $ApplicationName $ApplicationSWVersion -Description $ApplicationDescription -Publisher $ApplicationPublisher -SoftwareVersion $ApplicationSWVersion -OptionalReference $ApplicationDocURL -AutoInstall $ApplicationAutoInstall -ReleaseDate (Get-Date) -LocalizedName $ApplicationName $ApplicationSWVersion -LocalizedDescription $ApplicationDescription -UserDocumentation"
            New-CMApplication -Name "$ApplicationName $ApplicationSWVersion" -Description "$ApplicationDescription" -Publisher "$ApplicationPublisher" -SoftwareVersion $ApplicationSWVersion -OptionalReference $ApplicationDocURL -AutoInstall $ApplicationAutoInstall -ReleaseDate (Get-Date) -LocalizedName "$ApplicationName $ApplicationSWVersion" -LocalizedDescription "$ApplicationDescription" -UserDocumentation $ApplicationDocURL
        }
    }
    Catch {
        $ErrorMessage = $_.Exception.Message
        $FullyQualified = $_.FullyQualifiedErrorID
        Add-LogContent "ERROR: Application Creation Failed!"
        Add-LogContent "ERROR: $ErrorMessage"
        Add-LogContent "ERROR: $FullyQualified"
        Add-LogContent "ERROR: $($_.CategoryInfo.Category): $($_.CategoryInfo.Reason)"
    }
}

############################################################################################
##                                            MAIN                                        ##
############################################################################################

Add-LogContent -Load "Starting Dell BIOS Application Creator"

if (-not (Get-Module ConfigurationManager)) {
    Import-Module ConfigurationManager -ErrorAction SilentlyContinue
}

if (-not (Get-Module ConfigurationManager)) {
    Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1) -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
}

## Query for Dell Models
Push-location
Set-Location $Global:SCCMSite
$WMI = @"
select distinct SMS_G_System_COMPUTER_SYSTEM.Model from  SMS_R_System inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_G_System_COMPUTER_SYSTEM.Manufacturer like "%Dell%" and (SMS_G_System_COMPUTER_SYSTEM.Model like "%Optiplex%" or SMS_G_System_COMPUTER_SYSTEM.Model like "%Precision%" or SMS_G_System_COMPUTER_SYSTEM.Model like "%XPS%" or SMS_G_System_COMPUTER_SYSTEM.Model like "%Latitude%") order by SMS_G_System_COMPUTER_SYSTEM.Model
"@

Add-LogContent "Running Query for Dell Models"
$QueryResults = Invoke-CMWmiQuery -Query $WMI -Option Lazy

## Create Global Conditions
if (-not (Get-CMGlobalCondition -Name "AutoPackage - Computer Manufacturer")) {
    New-CMGlobalConditionWqlQuery -DataType String -Class Win32_ComputerSystem -Namespace root\cimv2 -Property Manufacturer -Name "AutoPackage - Computer Manufacturer" -Description "Returns the Manufacturer from ComputerSystem\Manufacturer"
}

if (-not (Get-CMGlobalCondition -Name "AutoPackage - Computer Model")) {
    New-CMGlobalConditionWqlQuery -DataType String -Class Win32_ComputerSystem -Namespace root\cimv2 -Property Model -Name "AutoPackage - Computer Model" -Description "Returns the Model from ComputerSystem\Model"
}

$ModelList = @()
foreach ($Model in $QueryResults) {
    $Row = New-Object -TypeName System.Management.Automation.PSObject
    $Row | Add-Member -MemberType NoteProperty -Name "Model" -Value $Model.Model
    $ModelList += $Row
}
Pop-Location

$ModelList | Export-Csv -Path "$Global:TempDir\AllDellModels.csv" -NoTypeInformation
$AllDellModels = $ModelList | Select-Object Model -ExpandProperty Model | Select-Object -last 3

New-Item -ItemType Directory -Path $Global:TempDir -Force -ErrorAction SilentlyContinue
$ProgressPreference = "SilentlyContinue"
Invoke-WebRequest -Uri $DellCatalogSource -OutFile "$Global:TempDir\$DellCatalogFile"
if ($?) {
    Add-LogContent "Download Succeeded"
}
else {
    Add-LogContent "Download Failed"
}
Start-Sleep 1

# Expand Cabinet File
Add-LogContent "Info: Expanding Dell Driver Pack Cabinet File: $DellCatalogFile"
Expand "$Global:TempDir\$DellCatalogFile" -F:* "$Global:TempDir\$DellCatalogXMLFile" | Out-Null
if ($?) {
    Add-LogContent "Expand Succeeded"
}
else {
    Add-LogContent "Expand Failed"
}

Push-Location
Set-location $Global:SCCMSite
if (-not (Get-CMApplication "$Global:BiosAppName")) {
    Pop-Location
    # Create Folder Structure
    New-Item -ItemType Directory -Path "$Global:BiosAppContentRoot" -Force -ErrorAction SilentlyContinue

    New-Item -ItemType Directory -Path "$Global:BiosAppContentRoot\Packages" -Force -ErrorAction SilentlyContinue 
    
    # Create the Application
    New-BIOSUpgradeApplication -ApplicationName $Global:BiosAppName
}
else {
    Pop-location   
}

Foreach ($DellModel in $AllDellModels) {
    Add-LogContent "Processing Model $DellModel"
    $Manufacturer = "Dell Inc."
    
    # Download Driver and Copy to Share
    $DellBIOSDownload = DellBiosFinder -Model $DellModel
    $DellBIOSVersion = $DellBIOSDownload.DellVersion
    $BIOSDownload = $DellDownloadBase + "/" + $($DellBIOSDownload.Path)
    $BIOSName = "$($DellModel.Replace(' ', '_'))_$DellBIOSVersion.exe"
    if (-not (Test-Path "$DestDir\$BIOSName")) {
        Add-LogContent "$BIOSDownload"
        Invoke-WebRequest -Uri "$BIOSDownload" -OutFile "$Global:TempDir\$BIOSName"
    }

    Add-LogContent "$Global:BiosappContentRoot"
    $ModelSpecificContentLoc = "$Global:BiosAppContentRoot\Packages\$DellModel"
    if (-not (Test-Path $ModelSpecificContentLoc -ErrorAction SilentlyContinue)) {
        New-Item -ItemType Directory -Path "$Global:BiosAppContentRoot\Packages\$DellModel" -Force -ErrorAction SilentlyContinue
    }
    
    # Check if File Already Exists, continue with packaging if it does not
    if ((-not(Test-Path "$ModelSpecificContentLoc\$BIOSName")) -or ($BIOSName -notlike "*_.exe")) {
        # Copy the BIOS File to the Destination
        Add-logcontent "Copying $BiosName to $ModelSpecificContentLoc\$BIOSName"
        Copy-Item -Path "$Global:TempDir\$BIOSName" -Destination "$ModelSpecificContentLoc\$BIOSName" -Force -ErrorAction SilentlyContinue
        
        # Create Detection Method
        Push-Location
        Set-Location $Global:SCCMSite
        $DetectionClause = New-CMDetectionClauseRegistryKeyValue -Hive "LocalMachine" -ExpressionOperator IsEquals -PropertyType "String" -Value -KeyName "HARDWARE\DESCRIPTION\System\BIOS" -ValueName "BIOSVersion" -ExpectedValue "$DellBIOSVersion"

        #Create Deployment Type
        $ModelSpecificInstallCmd = "$BIOSName $Global:InstallCmdParams" 
        Add-CMScriptDeploymentType -ApplicationName $Global:BiosAppName -DeploymentTypeName $DellModel -ContentLocation "$ModelSpecificContentLoc" -InstallCommand "$ModelSpecificInstallCmd" -AddDetectionClause $DetectionClause -EstimatedRuntimeMins "15" -MaximumRuntimeMins "45" -ContentFallback -EnableBranchCache -LogonRequirementType WhetherOrNotUserLoggedOn -UserInteractionMode Hidden -SlowNetworkDeploymentMode Download -InstallationBehaviorType InstallForSystem -RebootBehavior BasedOnExitCode
        
        # Add Manufacturer Queries to Template
        Add-LogContent "Processing - Add Manufacturers to Template"
        Add-LogContent "`"$Manufacturer`" is being added"
        $rule = Get-CMGlobalCondition -Name "AutoPackage - Computer Manufacturer" | New-CMRequirementRuleCommonValue -Value1 "$Manufacturer" -RuleOperator IsEquals 
        $rule.Name = "AutoPackage - Computer Manufacturer Equals $Manufacturer"
        Set-CMScriptDeploymentType -ApplicationName $Global:BiosAppName -DeploymentTypeName $DellModel -AddRequirement $rule

        # Add Model Queries to Template
        Add-LogContent "Processing - Add Models to Template"
        Add-LogContent "`"$DellModel`" is being added"
        $rule = Get-CMGlobalCondition -Name "AutoPackage - Computer Model" | New-CMRequirementRuleCommonValue -Value1 "$DellModel" -RuleOperator IsEquals 
        $rule.Name = "AutoPackage - Computer Model Equals $DellModel"
        Set-CMScriptDeploymentType -ApplicationName $Global:BiosAppName -DeploymentTypeName $DellModel -AddRequirement $rule
        Pop-Location
    }
    Else {
        Add-LogContent "ERROR: Downloading BIOS File for $DellModel has Failed!"

    }
}

Add-LogContent "Distributing Content for $Global:BiosAppName"
Try {
    Start-CMContentDistribution -ApplicationName "$Global:BiosAppName" -DistributionPointGroupName $Global:PreferredDistributionLoc -ErrorAction Stop
}
Catch {
    $ErrorMessage = $_.Exception.Message
    Add-LogContent "ERROR: Content Distribution Failed!"
    Add-LogContent "ERROR: $ErrorMessage"
}