<#	
	.NOTES
	===========================================================================
	 Created on:   	10/29/2018
	 Author:		Andrew Jimenez (asjimene) - https://github.com/asjimene/
	 Filename:     	DellBIOSUpdateAppCreator.ps1
	===========================================================================
	.DESCRIPTION
    Creates an App Model Application in SCCM for All Dell Models found in an SCCM environment
    
    Uses a method Described by Adam Juelich here: https://deployables.net/2018/10/25/updating-dell-bios-via-configmgr-application-method/

    Uses Scripts and Functions Sourced from the Following:
        Copy-CMDeploymentTypeRule - https://janikvonrotz.ch/2017/10/20/configuration-manager-configure-requirement-rules-for-deployment-types-with-powershell/
        Modifying a CM Requriement - https://fredbainbridge.com/2016/06/21/configmgr-add-os-requirements-to-an-application-deployment-type/

    NOTES:
    Before Using this script, make sure to do the following:
    1. Create a Global Condition with the following properties:
        - Name: Computer Model
        - Device Type: Windows 
        - Setting Type: WQL Query
        - Data Type: String
        - Namespace: root\cimv2
        - Class: Win32_ComputerSystem
        - Property: Model

    2. Create an Application with the following:
        - Name: Same as provided for the Variable $Global:RequirementsTemplateAppName
        - Deployment Type: Script Deployment Type
        - Deployment Type Name: Templates
        - Install Command: hostname.exe
        - Add the requirement created in Step 1 to this Deployment Type
            - Requirement should be set as follows:
                - Condition: Computer Model
                - Rule Type: Value
                - Operator: Equals
                - Value: MODEL NAME  (This is in all Caps with a Space between)
#>

$Global:SCCMSite = "##SITE:##"

# BIOS Upgrade App Desired Name
$Global:BiosAppName = "Upgrade Dell BIOS"

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
$Global:RequirementsTemplateAppName = "Application Requirements Template"
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
	} elseif ($DellBIOSFile -eq $null) {
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


# Add Requirements
Function Copy-CMDeploymentTypeRule {
    <#
	Function taken from https://janikvonrotz.ch/2017/10/20/configuration-manager-configure-requirement-rules-for-deployment-types-with-powershell/ and modified
 	
     #>
	Param (
		[System.String]$SourceApplicationName,
		[System.String]$DestApplicationName,
		[System.String]$DestDeploymentTypeName,
		[System.String]$RuleName
	)
    Push-Location
    Set-Location $Global:SCCMSite
	$DestDeploymentTypeIndex = 0
 
    # get the applications
    $SourceApplication = Get-CMApplication -Name $SourceApplicationName | ConvertTo-CMApplication
    $DestApplication = Get-CMApplication -Name $DestApplicationName | ConvertTo-CMApplication
	
	# Get DestDeploymentTypeIndex by finding the Title
	$DestApplication.DeploymentTypes | ForEach-Object {
		$i = 0
	} {
		If ($_.Title -eq "$DestDeploymentTypeName") {
			$DestDeploymentTypeIndex = $i
		}
		$i++
	}
	
	# get requirement rules from source application
    $Requirements = $SourceApplication.DeploymentTypes[0].Requirements | Where-Object {$_.Name -match $RuleName}
 
    # apply requirement rules
    $Requirements | ForEach-Object {
     
        $RuleExists = $DestApplication.DeploymentTypes[$DestDeploymentTypeIndex].Requirements | Where-Object {$_.Name -match $RuleName}
        if($RuleExists) {
 
            Add-LogContent "WARN: The rule `"$($_.Name)`" already exists in target application deployment type"
 
        } else{
         
            Add-LogContent "Apply rule `"$($_.Name)`" on target application deployment type"
 
            # create new rule ID
            $_.RuleID = "Rule_$( [guid]::NewGuid())"
 
            $DestApplication.DeploymentTypes[$DestDeploymentTypeIndex].Requirements.Add($_)
        }
    }
 
    # push changes
    $CMApplication = ConvertFrom-CMApplication -Application $DestApplication
    $CMApplication.Put()
    Pop-Location
}

function Update-ModelRequirement {
	Param (
        [System.String]$ModelName,
        [System.String]$BIOSApplicationName
	)
    Push-Location
    Set-Location $Global:SCCMSite
    #Get the Application
    $Application = Get-CMApplication $BIOSApplicationName

    # Convert the App to XML
    $Appxml = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($Application.SDMPackageXML,$True)

    #Get the Deployment Types
    #Excerpts taken from here: https://fredbainbridge.com/2016/06/21/configmgr-add-os-requirements-to-an-application-deployment-type/
    $DeploymentTypes = $AppXML.DeploymentTypes

    $TargetedDT = $DeploymentTypes | Where-Object Title -eq "$ModelName"
    $TargetedRequirements = $TargetedDT.Requirements | Where-Object Name -EQ "Computer Model Equals MODEL NAME"
    $TargetedRequirements.Name = $TargetedRequirements.Name.Replace("MODEL NAME","$ModelName")
    $TargetedRequirements.SecondOperand.Value = $TargetedRequirements.SecondOperand.Value.Replace("MODEL NAME","$ModelName")
    $TargetedRequirements.RuleID = "Rule_$([guid]::NewGuid())"

    $UpdatedXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::SerializeToString($AppXML, $True) 
    $Application.SDMPackageXML = $UpdatedXML 
    $Application.put()
    $t = Set-CMApplication -InputObject $Application -PassThru
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

$ModelList = @()
foreach ($Model in $QueryResults) {
	$Row = New-Object -TypeName System.Management.Automation.PSObject
	$Row | Add-Member -MemberType NoteProperty -Name "Model" -Value $Model.Model
	$ModelList += $Row
}
Pop-Location

$ModelList | Export-Csv -Path "$Global:TempDir\AllDellModels.csv" -NoTypeInformation
$AllDellModels = $ModelList | Select-Object Model -ExpandProperty Model

New-Item -ItemType Directory -Path $Global:TempDir -Force -ErrorAction SilentlyContinue
$ProgressPreference = "SilentlyContinue"
#Invoke-WebRequest -Uri $DellCatalogSource -OutFile "$Global:TempDir\$DellCatalogFile"
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

Foreach ($DellModel in $AllDellModels){
    Add-LogContent "Processing Model $DellModel"

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
    New-Item -ItemType Directory -Path "$Global:BiosAppContentRoot\Packages\$DellModel" -Force -ErrorAction SilentlyContinue

    # Check if File Already Exists, continue with packaging if it does not
    if ((-not(Test-Path "$ModelSpecificContentLoc\$BIOSName")) -or ($BIOSName -notlike "*_.exe")){
        # Copy the BIOS File to the Destination
        Add-logcontent "Copying $BiosName to $ModelSpecificContentLoc\$BIOSName"
        Copy-Item -Path "$Global:TempDir\$BIOSName" -Destination "$ModelSpecificContentLoc\$BIOSName" -Force -ErrorAction SilentlyContinue
        
        # Create Detection Method
        Push-Location
        Set-Location $Global:SCCMSite
        $DetectionClause = New-CMDetectionClauseRegistryKeyValue -Hive "LocalMachine" -ExpressionOperator IsEquals -PropertyType "String" -Value -KeyName "HARDWARE\DESCRIPTION\System\BIOS" -ValueName "BIOSVersion" -ExpectedValue "$DellBIOSVersion"
        Pop-Location

        #Create Deployment Type
        $ModelSpecificInstallCmd = "$BIOSName $Global:InstallCmdParams" 
        Push-Location
        Set-Location $Global:SCCMSite
        Add-CMScriptDeploymentType -ApplicationName $Global:BiosAppName -DeploymentTypeName $DellModel -ContentLocation "$ModelSpecificContentLoc" -InstallCommand "$ModelSpecificInstallCmd" -AddDetectionClause $DetectionClause -EstimatedRuntimeMins "15" -MaximumRuntimeMins "45" -ContentFallback -EnableBranchCache -LogonRequirementType WhetherOrNotUserLoggedOn -UserInteractionMode Hidden -SlowNetworkDeploymentMode Download -InstallationBehaviorType InstallForSystem -RebootBehavior BasedOnExitCode
        Pop-Location

        # Copy Requirement
        Copy-CMDeploymentTypeRule -SourceApplicationName $Global:RequirementsTemplateAppName -DestApplicationName $Global:BiosAppName -DestDeploymentTypeName $DellModel -RuleName "Computer Model Equals MODEL NAME"

        # Update Requirement for Model
        Update-ModelRequirement -ModelName $DellModel -BIOSApplicationName $Global:BiosAppName
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