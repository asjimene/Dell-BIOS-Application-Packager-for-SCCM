# Dell-BIOS-Application-Packager-for-SCCM
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
