<#PSScriptInfo
.VERSION 1.0.0
.AUTHOR Damien VAN ROBAEYS
.DESCRIPTION Create a reference catalog for BIOS settings on your different models
.PROJECTURI 
#>

<#
.SYNOPSIS
    Create a reference catalog for BIOS settings on your different models
.DESCRIPTION
	This script helps you to create a reference catalog of BIOS settings on your different models.
	The idea is to run the script on each model and use the CSV each time.
	Don't forget to each the last version of CSV each time, meaning the CSV containing latest models.
	The script works only for Lenovo for now
.PARAMETER Create
	Creates the catalog (not used for updating catalog)
	This parameter is a switch
.PARAMETER CSV_Catalog
    Path of the CSV path (with .CSV)
	This parameter is a string	
.PARAMETER Show_Diff
    Display new values in PowerShell console
	This parameter is a switch		
.EXAMPLE
	Create the first release of the catalog from a T480s model
	.\BIOS_Settings_Catalog.ps1 -Create -CSV_Catalog "C:\MyCatalog.csv"
	
	Updates the existing catalog (containing T480s model) by running the script on another model: T14s
	.\BIOS_Settings_Catalog.ps1 -CSV_Catalog "C:\MyCatalog.csv"
	
	Updates the existing catalog (containing T480s model) by running the script on another model: T14s
	Displays new settings (if there are) on the PowerShell console
	.\BIOS_Settings_Catalog.ps1 -CSV_Catalog "C:\MyCatalog.csv"	-Show_Diff
	
	Updates the existing catalog (containing T480s, T14s models) by running the script on another model: T490s
	.\BIOS_Settings_Catalog.ps1 -CSV_Catalog "C:\MyCatalog.csv"	
#>

param(
[switch]$Create,
[string]$CSV_Catalog,
[switch]$Show_Diff
)
  
Function Write_Log
	{
		param(
		$Message_Type, 
		$Message
		)

		$MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
		write-host  "$MyDate - $Message_Type : $Message"
	} 
	
Write_Log -Message_Type "INFO" -Message "Running BIOS settings catalog builder"	

$Manufacturer = (gwmi win32_computersystem).Manufacturer
Write_Log -Message_Type "INFO" -Message "Your manufacturer is: $Manufacturer"	
If($Manufacturer -like "*lenovo*")
	{
		Write_Log -Message_Type "INFO" -Message "Your manufacturer is supported"		
	}
Else	
	{
		Write_Log -Message_Type "ERROR" -Message "Your manufacturer is not supported"	
		write-warning "Your manufacturer is not supported"	
		break		
	}

[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null
$Run_As_Admin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
If($Run_As_Admin -eq $False)
	{
		Write_Log -Message_Type "ERROR" -Message "The script has not been lauched with admin rights"
		write-warning "Please run the script with admin rigths"		
		break		
	}
Else
	{
		Write_Log -Message_Type "SUCCESS" -Message "Running with admin rights"
	}

# $CSV_Catalog = "D:\bioscsv.csv"

# Check current model friendly name
$Get_Current_Model_FamilyName = (gwmi win32_computersystem).SystemFamily.split(" ")[1]
Write_Log -Message_Type "INFO" -Message "Current model is: $Get_Current_Model_FamilyName"

# Function to get BIOS settings on the current model
Function Get_Lenovo_BIOS_Settings
	{
		$Script:Get_Available_Values = Get-WmiObject -Class Lenovo_GetBiosSelections -Namespace root\wmi
		$Script:Get_BIOS_Settings = gwmi -class Lenovo_BiosSetting -namespace root\wmi  | select-object currentsetting | Where-Object {$_.CurrentSetting -ne ""} |
		select-object @{label = "Setting"; expression = {$_.currentsetting.split(",")[0]}} ,
		@{label = "New"; expression = {"No"}},	
		@{label = "$Get_Current_Model_FamilyName"; expression = {$_.currentsetting.split(",*;[")[1]}},
		@{label = "Available values"; expression = {($Get_Available_Values.GetBiosSelections($_.currentsetting.split(",")[0])).Selections}}            
		$Get_BIOS_Settings | select *
	}
 
# Check if CSV catalog catalog path has been setted 
If($CSV_Catalog -eq "")
	{
		$CSV_Catalog = read-host "Please type the path of the CSV catalog to use"
	}
	
If($CSV_Catalog -eq "")
	{
		Write_Log -Message_Type "INFO" -Message "Please add the CSV catalog path using parameter -CSV_Catalog"
		write-warning "Please add the CSV catalog path using parameter -CSV_Catalog"
		Break
	}	
	
Write_Log -Message_Type "INFO" -Message "Catalog path: $CSV_Catalog"
	
# If Generate catalog is set, it means it's the first time you run the script and you want to create the catalog for a first model 
If($Create)
	{
		Write_Log -Message_Type "INFO" -Message "Catalog is in create mode"
		Get_Lenovo_BIOS_Settings | Export-CSV $CSV_Catalog -Delimiter ";" -NoTypeInformation
		Write_Log -Message_Type "SUCCESS" -Message "Catalog has been created from the current model: $Get_Current_Model_FamilyName"		
	}
# If Generate catalog is not setted, means you already have a CSV catalog and want to add BIOS settings of another model in the catalog
Else
	{ 			
		Write_Log -Message_Type "INFO" -Message "Catalog is in update mode"
	
		# Load the CSV catalog
		Write_Log -Message_Type "INFO" -Message "Importing existing catalog"				
		$Get_CSV_FirstLine = Get-Content $CSV_Catalog | Select -First 1
		$Get_Delimiter = If($Get_CSV_FirstLine.Split(";").Length -gt 1){";"}Else{","};
		$Get_CSV_Content = import-csv $CSV_Catalog -Delimiter $Get_Delimiter
		Write_Log -Message_Type "SUCCESS" -Message "Importing existing catalog"						
		
		# Create a new array to gather catalog and current model
		$Existing_Settings_Array = @()

		# Convert Lenovo CSV catalog to an array without knowing name of headers
		Write_Log -Message_Type "INFO" -Message "Converting existing catalog"						
		$CSV_Headers = ($Get_CSV_Content | Get-Member -MemberType NoteProperty).Name
		foreach ($Row in $Get_CSV_Content) {
			$Obj = New-Object PSObject
			foreach ($Header in $CSV_Headers) {
				Add-Member -InputObject $Obj -MemberType NoteProperty -Name $Header -Value $Row.$Header
			}
			Add-Member -InputObject $Obj -MemberType NoteProperty -Name $Get_Current_Model_FamilyName -Value ""
			$Existing_Settings_Array += $Obj
		}
		Write_Log -Message_Type "SUCCESS" -Message "Converting existing CSV catalog"								

		# Compare settings from catalog and current model
		# Parse settings from current model
		Write_Log -Message_Type "INFO" -Message "Comparing settings from catalog and current model: $Get_Current_Model_FamilyName"								
		ForEach($Setting1 in Get_Lenovo_BIOS_Settings)
			{
				$Setting1_Name = $Setting1.Setting
				$Setting1_Value = $Setting1.$Get_Current_Model_FamilyName
				$Setting1_All_Value = $Setting1."Available values"
				
				# Parse settings from CSV
				ForEach($Setting2 in $Existing_Settings_Array)
					{						
						$found = $false
						$Setting2_Name = $Setting2.Setting
						# If settings are equals, meaning if a setting from the current model has been found in the CSV catalog
						If($Setting2_Name -eq $Setting1_Name)
							{
								# Set New tab to No and add setting value in the tab called with the model friendly name 
								# If model is T480s, settings for the current model will be added in a tab T480s
								$found = $true
								$Setting2.New = "No"
								$Setting2.$Get_Current_Model_FamilyName = $Setting1_Value
								Break
							}
					}
					
					# If settings are not equals, meaning if a setting from the current model has not been found in the CSV catalog
					IF (-not $found) 
						{				
							# Add properties n the array: setting name, value, available values
							$Obj2 = New-Object PSObject
							Add-Member -InputObject $Obj2 -MemberType NoteProperty -Name "New" -Value "At least on $Get_Current_Model_FamilyName"
							Add-Member -InputObject $Obj2 -MemberType NoteProperty -Name "Setting" -Value $Setting1_Name
							Add-Member -InputObject $Obj2 -MemberType NoteProperty -Name "Available values" -Value $Setting1_All_Value
							Add-Member -InputObject $Obj2 -MemberType NoteProperty -Name $Get_Current_Model_FamilyName -Value $Setting1_Value
							$Existing_Settings_Array += $Obj2
							
							If($Show_Diff)
								{
									Write_Log -Message_Type "INFO" -Message "New setting: $Setting1_Name - Value: $Setting1_Value"								
								}
						}				
			}
			
			Write_Log -Message_Type "SUCCESS" -Message "Comparing settings from catalog and current model: $Get_Current_Model_FamilyName"								

			# Format the CSV in a correct order
			# Create a header filter with following properties: "Setting", "New", "Available values"
			$filter_headers = "Setting", "New", "Available values"
			# Add headers corresponding to models in the filter header: like T480s, T14s, T480s
			# At this step the filter header will be: Setting, New, Available values, T480s, T14s...
			$filter_headers += (($Existing_Settings_Array | Get-Member -MemberType NoteProperty).Name | where {$_ -ne "setting" -and $_ -notlike "*avai*" -and $_ -ne "new"}) 
			# Export array to CSV in the order defined in the filter_headers
			# $Existing_Settings_Array | select $filter_headers | Export-CSV "D:\Lenovo_BIOS_Catalog.csv" -Delimiter $Get_Delimiter -NoTypeInformation
			$Existing_Settings_Array | select $filter_headers | Export-CSV $CSV_Catalog -Delimiter $Get_Delimiter -NoTypeInformation

			# In this part we will check empty cell in the Csv
			# Empty cells means settings that are not present for a model
			# We will replace empty part with text "Not present"
			$Check_CSV = import-csv $CSV_Catalog -Delimiter $Get_Delimiter
			$Check_CSV | Foreach-Object {
				$_.PSObject.Properties | Foreach-Object{ 
				If($_.Value -eq ""){($_.Value) = "not present"}}  
			}
			# Export the final CSV catalog
			Write_Log -Message_Type "INFO" -Message "exporting new catalog to: $CSV_Catalog"											
			$Check_CSV | Export-Csv $CSV_Catalog -delimiter $Get_Delimiter -NoTypeInformation
			Write_Log -Message_Type "SUCCESS" -Message "exporting new catalog to: $CSV_Catalog"														
	}

