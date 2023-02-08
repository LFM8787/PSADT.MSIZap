# PSADT.MSIZap 1.0.1
Extension for PowerShell App Deployment Toolkit to execute MSIZap after Windows Installer uninstallation.

## Features
- Invokes Microsoft MSIZap utility after a Windows Installer uninstallation process.
- Wraps the original function, so no script modification needed.
- Used to zap even a failed uninstallation (default).
- Wraps *ContinueOnError* and *ExitOnProcessFailure* original function parameters.

## Disclaimer
```diff
- Test the functions before production.
- Make a backup before applying.
- Check the config file options description.
- Run AppDeployToolkitHelp.ps1 for more help and parameter descriptions.
```

## Functions
* **Execute-MSI** - Wraps the original function but removes the ExitOnProcessFailure and ContinueOnError before.
* **Invoke-MSIZap** - Executes msizap.exe to perform a deep clean after uninstallation.

## Usage
```PowerShell
# Automatically invokes by Execute-MSI but can be called directly
Invoke-MSIZap -Path '{A1E2B44D-3F7A-93AF-9423-AB73411328DC}'
```

## Internal functions
`This set of functions are internals and are not designed to be called directly`
* **New-DynamicFunction** - Defines a new function with the given name, scope and content given.

## Extension Exit Codes
|Exit Code|Function|Exit Code Detail|
|:----------:|:--------------------|:-|
|70501|Invoke-MSIZap|The following error has been returned by MSIZap.|

## How to Install
#### 1. Download and copy into Toolkit folder.
#### 2. Edit *AppDeployToolkitExtensions.ps1* file and add the following lines.
#### 3. Create an empty array (only once if multiple extensions):
```PowerShell
## Variables: Extensions to load
$ExtensionToLoad = @()
```
#### 4. Add Extension Path and Script filename (repeat for multiple extensions):
```PowerShell
$ExtensionToLoad += [PSCustomObject]@{
	Path   = "PSADT.MSIZap"
	Script = "MSIZapExtension.ps1"
}
```
#### 5. Complete with the remaining code to load the extension (only once if multiple extensions):
```PowerShell
## Loading extensions
foreach ($Extension in $ExtensionToLoad) {
	$ExtensionPath = $null
	if ($Extension.Path) {
		[IO.FileInfo]$ExtensionPath = Join-Path -Path $scriptRoot -ChildPath $Extension.Path | Join-Path -ChildPath $Extension.Script
	}
	else {
		[IO.FileInfo]$ExtensionPath = Join-Path -Path $scriptRoot -ChildPath $Extension.Script
	}
	if ($ExtensionPath.Exists) {
		try {
			. $ExtensionPath
		}
		catch {
			Write-Log -Message "An error occurred while trying to load the extension file [$($ExtensionPath)].`r`n$(Resolve-Error)" -Severity 3 -Source $appDeployToolkitExtName
		}
	}
	else {
		Write-Log -Message "Unable to locate the extension file [$($ExtensionPath)]." -Severity 2 -Source $appDeployToolkitExtName
	}
}
```

## Requirements
* Powershell 5.1+
* PSAppDeployToolkit 3.8.4+

## External Links
* [PowerShell App Deployment Toolkit](https://psappdeploytoolkit.com/)
* [Msizap.exe - Win32 apps | Microsoft Learn](https://learn.microsoft.com/en-us/windows/win32/msi/msizap-exe)