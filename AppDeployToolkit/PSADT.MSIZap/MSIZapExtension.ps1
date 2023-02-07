<#
.SYNOPSIS
	MSI Zap Extension script file, must be dot-sourced by the AppDeployToolkitExtension.ps1 script.
.DESCRIPTION
	Runs MSIZap at the end of the Windows Installer application removal process.
.NOTES
	Author:  Leonardo Franco Maragna
	Version: 1.0
	Date:    2023/02/07
#>
[CmdletBinding()]
Param (
)

##*===============================================
##* VARIABLE DECLARATION
##*===============================================
#region VariableDeclaration

## Variables: Extension Info
$MSIZapExtName = "MSIZapExtension"
$MSIZapExtScriptFriendlyName = "MSI Zap Extension"
$MSIZapExtScriptVersion = "1.0"
$MSIZapExtScriptDate = "2023/02/07"
$MSIZapExtSubfolder = "PSADT.MSIZap"
$MSIZapExtConfigFileName = "MSIZapConfig.xml"

## Variables: MSI Zap Script Dependency Files
[IO.FileInfo]$dirMSIZapExtFiles = Join-Path -Path $scriptRoot -ChildPath $MSIZapExtSubfolder
[IO.FileInfo]$dirMSIZapExtSupportFiles = Join-Path -Path $dirSupportFiles -ChildPath $MSIZapExtSubfolder
[IO.FileInfo]$MSIZapConfigFile = Join-Path -Path $dirMSIZapExtFiles -ChildPath $MSIZapExtConfigFileName
if (-not $MSIZapConfigFile.Exists) { throw "$($MSIZapExtScriptFriendlyName) XML configuration file [$MSIZapConfigFile] not found." }

## Variables: Required Support Files
$msizapApplicationPath = (Get-ChildItem -Path $dirMSIZapExtSupportFiles -Recurse -Include "*msizap*.exe").FullName | Select-Object -First 1

## Import variables from XML configuration file
[Xml.XmlDocument]$xmlMSIZapConfigFile = Get-Content -LiteralPath $MSIZapConfigFile -Encoding UTF8
[Xml.XmlElement]$xmlMSIZapConfig = $xmlMSIZapConfigFile.MSIZap_Config

#  Get Config File Details
[Xml.XmlElement]$configMSIZapConfigDetails = $xmlMSIZapConfig.Config_File

#  Check compatibility version
$configMSIZapConfigVersion = [string]$configMSIZapConfigDetails.Config_Version
#$configMSIZapConfigDate = [string]$configMSIZapConfigDetails.Config_Date

try {
	if ([version]$MSIZapExtScriptVersion -ne [version]$configMSIZapConfigVersion) {
		Write-Log -Message "The $($MSIZapExtScriptFriendlyName) version [$([version]$MSIZapExtScriptVersion)] is not the same as the $($MSIZapExtConfigFileName) version [$([version]$configMSIZapConfigVersion)]. Problems may occurs." -Severity 2 -Source ${CmdletName}
	}
}
catch {}

#  Get MSI Zap General Options
[Xml.XmlElement]$xmlMSIZapOptions = $xmlMSIZapConfig.MSIZap_Options
$configMSIZapGeneralOptions = [PSCustomObject]@{
	InvokeMSIZapAfterMsiUninstall         = Invoke-Expression -Command 'try { [boolean]::Parse([string]($xmlMSIZapOptions.InvokeMSIZapAfterMsiUninstall)) } catch { $true }'
	InvokeMSIZapIfUninstallFails          = Invoke-Expression -Command 'try { [boolean]::Parse([string]($xmlMSIZapOptions.InvokeMSIZapIfUninstallFails)) } catch { $false }'

	RemoveForAllUsersInUserContextIfAdmin = Invoke-Expression -Command 'try { [boolean]::Parse([string]($xmlMSIZapOptions.RemoveForAllUsersInUserContextIfAdmin)) } catch { $true }'
	RemoveForAllUsersInSystemContext      = Invoke-Expression -Command 'try { [boolean]::Parse([string]($xmlMSIZapOptions.RemoveForAllUsersInSystemContext)) } catch { $true }'
}

#  Defines the original functions to be renamed
$FunctionsToRename = @()
$FunctionsToRename += [PSCustomObject]@{
	Scope = "Script"
	Name  = "Execute-MSIOriginal"
	Value = $(${Function:Execute-MSI}.ToString().Replace("http://psappdeploytoolkit.com", ""))
}

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function New-DynamicFunction
Function New-DynamicFunction {
	<#
	.SYNOPSIS
		Defines a new function with the given name, scope and content given.
	.DESCRIPTION
		Defines a new function with the given name, scope and content given.
	.PARAMETER Name
		Function name.
	.PARAMETER Scope
		Scope where the function will be created.
	.PARAMETER Value
		Logic of the function.
	.PARAMETER ContinueOnError
		Continue if an error occured while trying to create new function. Default: $false.
	.EXAMPLE
		New-DynamicFunction -Name 'Exit-ScriptOriginal' -Scope 'Script' -Value ${Function:Exit-Script}
	.NOTES
		This is an internal script function and should typically not be called directly.
		Author: Leonardo Franco Maragna
		Part of Toast Notification Extension
	.LINK
		https://github.com/LFM8787/PSADT.MSIZap
		http://psappdeploytoolkit.com
	#>
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullorEmpty()]
		[ValidateSet("Global", "Local", "Script")]
		[string]$Scope,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullorEmpty()]
		[string]$Name,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullorEmpty()]
		[string]$Value,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[boolean]$ContinueOnError = $false
	)

	Begin {
		## Get the name of this function and write header
		[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
	}
	Process {
		try {
			$null = New-Item -Path function: -Name "$($Scope):$($Name)" -Value $Value -Force

			if ($?) {
				Write-Log -Message "Successfully created function [$Name] in scope [$Scope]." -Source ${CmdletName} -DebugMessage
			}
		}
		catch {
			Write-Log -Message "Failed when trying to create new function [$Name] in scope [$Scope].`r`n$(Resolve-Error)" -Severity 3 -Source ${CmdletName}
			if (-not $ContinueOnError) {
				throw "Failed when trying to create new function [$Name] in scope [$Scope]: $($_.Exception.Message)"
			}
		}
	}
	End {
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
	}
}
#endregion


#region Rename Original Functions
#  Called now, before functions subsitution
$FunctionsToRename | ForEach-Object { New-DynamicFunction -Name $_.Name -Scope $_.Scope -Value $_.Value }
#endregion


#region Function Execute-MSI
Function Execute-MSI {
	<#
	.SYNOPSIS
		Wraps the original function but removes the ExitOnProcessFailure and ContinueOnError before.
	.DESCRIPTION
		Wraps the original function but removes the ExitOnProcessFailure and ContinueOnError before.
	.PARAMETER Action
		The action to perform. Options: Install, Uninstall, Patch, Repair, ActiveSetup.
	.PARAMETER Path
		The path to the MSI/MSP file or the product code of the installed MSI.
	.PARAMETER Transform
		The name of the transform file(s) to be applied to the MSI. The transform file is expected to be in the same directory as the MSI file. Multiple transforms have to be separated by a semi-colon.
	.PARAMETER Patch
		The name of the patch (msp) file(s) to be applied to the MSI for use with the "Install" action. The patch file is expected to be in the same directory as the MSI file. Multiple patches have to be separated by a semi-colon.
	.PARAMETER Parameters
		Overrides the default parameters specified in the XML configuration file. Install default is: "REBOOT=ReallySuppress /QB!". Uninstall default is: "REBOOT=ReallySuppress /QN".
	.PARAMETER AddParameters
		Adds to the default parameters specified in the XML configuration file. Install default is: "REBOOT=ReallySuppress /QB!". Uninstall default is: "REBOOT=ReallySuppress /QN".
	.PARAMETER SecureParameters
		Hides all parameters passed to the MSI or MSP file from the toolkit Log file.
	.PARAMETER LoggingOptions
		Overrides the default logging options specified in the XML configuration file. Default options are: "/L*v".
	.PARAMETER LogName
		Overrides the default log file name. The default log file name is generated from the MSI file name. If LogName does not end in .log, it will be automatically appended.
		For uninstallations, by default the product code is resolved to the DisplayName and version of the application.
	.PARAMETER WorkingDirectory
		Overrides the working directory. The working directory is set to the location of the MSI file.
	.PARAMETER SkipMSIAlreadyInstalledCheck
		Skips the check to determine if the MSI is already installed on the system. Default is: $false.
	.PARAMETER IncludeUpdatesAndHotfixes
		Include matches against updates and hotfixes in results.
	.PARAMETER NoWait
		Immediately continue after executing the process.
	.PARAMETER PassThru
		Returns ExitCode, STDOut, and STDErr output from the process.
	.PARAMETER IgnoreExitCodes
		List the exit codes to ignore or * to ignore all exit codes.
	.PARAMETER PriorityClass	
		Specifies priority class for the process. Options: Idle, Normal, High, AboveNormal, BelowNormal, RealTime. Default: Normal
	.PARAMETER ExitOnProcessFailure
		Specifies whether the function should call Exit-Script when the process returns an exit code that is considered an error/failure. Default: $true
	.PARAMETER RepairFromSource
		Specifies whether we should repair from source. Also rewrites local cache. Default: $false
	.PARAMETER ContinueOnError
		Continue if an error occurred while trying to start the process. Default: $false.
	.EXAMPLE
		Execute-MSI -Action 'Install' -Path 'Adobe_FlashPlayer_11.2.202.233_x64_EN.msi'
		Installs an MSI
	.NOTES
		Author: Leonardo Franco Maragna
		Part of MSI Zap Extension
	.LINK
		https://github.com/LFM8787/PSADT.MSIZap
		http://psappdeploytoolkit.com
	#>
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory = $false)]
		[ValidateSet('Install', 'Uninstall', 'Patch', 'Repair', 'ActiveSetup')]
		[string]$Action = 'Install',
		[Parameter(Mandatory = $true, HelpMessage = 'Please enter either the path to the MSI/MSP file or the ProductCode')]
		[ValidateScript({ ($_ -match $MSIProductCodeRegExPattern) -or ('.msi', '.msp' -contains [IO.Path]::GetExtension($_)) })]
		[Alias('FilePath')]
		[string]$Path,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[string]$Transform,
		[Parameter(Mandatory = $false)]
		[Alias('Arguments')]
		[ValidateNotNullorEmpty()]
		[string]$Parameters,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[string]$AddParameters,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[switch]$SecureParameters = $false,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[string]$Patch,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[string]$LoggingOptions,
		[Parameter(Mandatory = $false)]
		[Alias('LogName')]
		[string]$private:LogName,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[string]$WorkingDirectory,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[switch]$SkipMSIAlreadyInstalledCheck = $false,
		[Parameter(Mandatory = $false)]
		[switch]$IncludeUpdatesAndHotfixes = $false,
		[Parameter(Mandatory = $false)]
		[switch]$NoWait = $false,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[switch]$PassThru = $false,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[string]$IgnoreExitCodes,
		[Parameter(Mandatory = $false)]
		[ValidateSet('Idle', 'Normal', 'High', 'AboveNormal', 'BelowNormal', 'RealTime')]
		[Diagnostics.ProcessPriorityClass]$PriorityClass = 'Normal',
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[boolean]$ExitOnProcessFailure = $true,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[boolean]$RepairFromSource = $false,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[boolean]$ContinueOnError = $false
	)
	
	Begin {
		## Get the name of this function and write header
		[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
	}
	Process {
		## Get the parameters passed to the function for invoking the original function
		[hashtable]$executeMSIParameters = $PSBoundParameters

		if ($configMSIZapGeneralOptions.InvokeMSIZapAfterMsiUninstall) {
			Write-Log -Message "The MSI Zap Extension is loaded." -Source ${CmdletName} -DebugMessage

			if ($Action -eq "Uninstall" -and [IO.Path]::GetExtension($Path) -ne ".msp") {
				if (-not $ExitOnProcessFailure -and $ContinueOnError) {
					Write-Log -Message "The original function parameters are correctly set to use the MSI Zap Extension loaded after the uninstallation." -Source ${CmdletName} -DebugMessage
				}
				elseif (-not $configMSIZapGeneralOptions.InvokeMSIZapIfUninstallFails) {
					Write-Log -Message "The MSI Zap Extension is loaded but the override parameter [InvokeMSIZapIfUninstallFails] is not correctly set, if the uninstallation fails no zap will be executed, check config file [$(Split-Path -Path $MSIZapConfigFile -Leaf)]." -Severity 2 -Source ${CmdletName}
				}
				else {
					#  Override ExitOnProcessFailure parameter.
					if ($executeMSIParameters.ContainsKey("ExitOnProcessFailure")) {
						$executeMSIParameters.Remove("ExitOnProcessFailure")
					}
					$executeMSIParameters.Add("ExitOnProcessFailure", $false)

					#  Override ContinueOnError parameter.
					if ($executeMSIParameters.ContainsKey("ContinueOnError")) {
						$executeMSIParameters.Remove("ContinueOnError")
					}
					$executeMSIParameters.Add("ContinueOnError", $true)
				}
			}
		}
		else {
			Write-Log -Message "The MSI Zap Extension is loaded but deactivated, check config file [$(Split-Path -Path $MSIZapConfigFile -Leaf)]." -Severity 2 -Source ${CmdletName}
		}

		## Execute original function
		Execute-MSIOriginal @executeMSIParameters

		if ($configMSIZapGeneralOptions.InvokeMSIZapAfterMsiUninstall) {
			if ($Action -eq "Uninstall" -and [IO.Path]::GetExtension($Path) -ne ".msp") {
				#Estoy desinstalando desde un msi o producto, tengo en $Path el valor
				Invoke-MSIZap -Path $Path -ExitOnProcessFailure $ExitOnProcessFailure -ContinueOnError $ContinueOnError

				#  Refresh environment variables for Windows Explorer process as Windows does not consistently update environment variables created by MSIs
				Update-Desktop
			}
		}
	}
	End {
		If ($PassThru) { Write-Output -InputObject $ExecuteResults }
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
	}
}
#endregion


#region Function Invoke-MSIZap
Function Invoke-MSIZap {
	<#
	.SYNOPSIS
		Executes msizap.exe to perform a deep clean after uninstallation.
	.DESCRIPTION
		Executes msizap.exe to perform a deep clean after uninstallation.
	.PARAMETER Path
		The path to the MSI file or the product code of the installed MSI.
	.PARAMETER ExitOnProcessFailure
		Specifies whether the function should call Exit-Script when the process returns an exit code that is considered an error/failure. Default: $true
	.PARAMETER ContinueOnError
		Continue if an error occurred while trying to start the process. Default: $false.
	.EXAMPLE
		Invoke-MSIZap -Path '{A1E2B44D-3F7A-93AF-9423-AB73411328DC}'
	.NOTES
		Author: Leonardo Franco Maragna
		Part of MSI Zap Extension
	.LINK
		https://github.com/LFM8787/PSADT.MSIZap
		http://psappdeploytoolkit.com
	#>
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory = $true, HelpMessage = 'Please enter either the path to the MSI file or the ProductCode')]
		[ValidateScript({ ($_ -match $MSIProductCodeRegExPattern) -or ([IO.Path]::GetExtension($_) -eq ".msi") })]
		[Alias('FilePath')]
		[string]$Path,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[boolean]$ExitOnProcessFailure = $true,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[boolean]$ContinueOnError = $false
	)

	Begin {
		## Get the name of this function and write header
		[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
	}
	Process {
		## Check if the required application exists
		if (-not (Test-Path -Path $msizapApplicationPath -ErrorAction SilentlyContinue)) {
			Write-Log -Message "Unable to locate msizap.exe application, the required file must be located under [..\SupportFiles\PSADT.MSIZap\] directory." -Severity 3 -Source ${CmdletName}
			break
		}

		## Defines the product part of the argument
		if ([IO.Path]::GetExtension($Path) -eq ".msi") {
			$Product = "`"$Path`""
		}
		else {
			$Path = $Path -replace "{|}", ""
			$Product = "{$($Path)}"
		}

		## Defines the scope of the argument
		$AllUsers = ""

		if ($CurrentLoggedOnUserSession) {
			if ($IsAdmin -and $configMSIZapGeneralOptions.RemoveForAllUsersInUserContextIfAdmin) {
				#  If the user has administrative rights
				$AllUsers = "W"
			}
		}
		else {
			#  Running in system context
			if ($configMSIZapGeneralOptions.RemoveForAllUsersInSystemContext) {
				$AllUsers = "W"
			}
		}

		if ($AllUsers -eq "") {
			Write-Log -Message "Executing MSI Zap to delete and clean the product [$Product] without removing it from all users." -Source ${CmdletName}

		}
		else {
			Write-Log -Message "Executing MSI Zap to delete and clean the product [$Product] from all users." -Source ${CmdletName}
		}

		## Function parameters used to call Execute-Process function
		[hashtable]$ExecuteProcessSplat = @{
			Path                 = $msizapApplicationPath
			Parameters           = "T$($AllUsers)! $($Product)"
			WindowStyle          = "Hidden"
			CreateNoWindow       = $true
			ExitOnProcessFailure = $ExitOnProcessFailure
			UseShellExecute      = $false
			PassThru             = $true
			ContinueOnError      = $ContinueOnError
		}

		## Execute function and get result
		$MSIZapReturnedObject = Execute-Process @ExecuteProcessSplat

		if (-not [string]::IsNullOrWhiteSpace($MSIZapReturnedObject.StdOut)) {
			$MSIZapReturnedObject.StdOut | ForEach-Object { Write-Log -Message $_ -Source ${CmdletName} -DebugMessage }
		}

		if (-not [string]::IsNullOrWhiteSpace($MSIZapReturnedObject.StdErr)) {
			$MSIZapReturnedObject.StdErr | ForEach-Object { Write-Log -Message $_ -Severity 3 -Source ${CmdletName} -DebugMessage }
		}
	}
	End {
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
	}
}
#endregion

#endregion
##*===============================================
##* END FUNCTION LISTINGS
##*===============================================

##*===============================================
##* SCRIPT BODY
##*===============================================
#region ScriptBody

if ($scriptParentPath) {
	Write-Log -Message "Script [$($MyInvocation.MyCommand.Definition)] dot-source invoked by [$(((Get-Variable -Name MyInvocation).Value).ScriptName)]" -Source $MSIZapExtName
}
else {
	Write-Log -Message "Script [$($MyInvocation.MyCommand.Definition)] invoked directly" -Source $MSIZapExtName
}

#endregion
##*===============================================
##* END SCRIPT BODY
##*===============================================