<#
    .SYNOPSIS
	Create CSV report from MTA-STS emails from sending MTAs

    .DESCRIPTION
	If an _smtp._tls record is configured with an rua email address, senders will send failure reports and statistics to this address as attachments containing GZipped JSON files.
	This script will Prompt for an Outlook mailbox folder the reports are delivered to.
	All attachments from emails in this folder will be saved to the specified folder.
	The JSON files will be extracted from the GZ files.
	A CSV report will be generated from the data in the JSON files.
	The .GZ files will be deleted after extraction unless the -nocleanup switch is specified.
	The .JSON files will be deleted after the report is generated unless the -nocleanup switch is specified.
	If the -reportonly switch is specified, the script will only try to generate a report from existing JSON files.
	Use this option if you obtain the JSON files another way, or the the report has previously failed to be generated.

    .PARAMETER jsonfolder
	Mandatory string. Folder used to save attachments and extract the JSON files

    .PARAMETER reportfile
	String. Output report file. Defaults to ".\mtastsreport.csv"

    .PARAMETER reportonly
	Switch. Only generate a report from existing JSON files. Do not save attachments or extract files from GZ files.

    .PARAMETER nocleanup
	Switch. Don't delete .GZ and .JSON files when they are no longer needed.

    .INPUTS
	None

    .OUTPUTS
	CSV report generated from JSON data.

    .NOTES
	V 1.20220216.1

    .EXAMPLE
	.\new-mtastsreport.ps1 "C:\temp\mta-tls reports"

	Prompt for folder in Outlook using GUI picker. Save all attachments from the folder to "C:\temp\mta-tls reports".
	Attempt to extract contents of all GZ files to "C:\temp\mta-tls reports" using 7Z.EXE or 7Z Powershell module.
	Generate a CSV report to the default location ".\mtastsreport.csv".
	Cleanup .GZ and .JSON files when they are no longer needed.

    .EXAMPLE
	.\new-mtastsreport.ps1 "C:\temp\mta-tls reports" -nocleanup -reportfile "z:\myreport.csv"

	Prompt for folder in Outlook using GUI picker. Save all attachments from the folder to "C:\temp\mta-tls reports".
	Attempt to extract contents of all GZ files to "C:\temp\mta-tls reports" using 7Z.EXE or 7Z Powershell module.
	Generate a CSV report to the location "z:\myreport.csv".
	Do not cleanup .GZ and .JSON files.

    .EXAMPLE
	.\new-mtastsreport.ps1 "C:\temp\mta-tls reports" -reportonly

	Generate a CSV report using JSON files found in "C:\temp\mta-tls reports" to the default location ".\mtastsreport.csv".
	Cleanup .JSON files when they are no longer needed.

#>

# ============================================================================
#region Parameters
# ============================================================================
Param(
	[Parameter(Mandatory=$true,Position=0)]
	[String] $jsonfolder,

	[Parameter()]
	[String] $reportfile=".\mtastsreport.csv",

	[Parameter()]
	[Switch] $reportonly,

	[Parameter()]
	[Switch] $nocleanup

)
#endregion Parameters

# ============================================================================
#region Functions
# ============================================================================

# Test if specified path can be written to. Return True/False
function test-writepath {
	param ( [Parameter(Mandatory=$true)]
		[string] $testpath
	)

	try {
		new-item -path $testpath -name "testfile" -force -erroraction stop|out-null
		remove-item -path $testpath"\testfile" -force -erroraction stop|out-null
		return $true
	} catch {
		Return $false
	}
}

# Test if 7zip is installed. Return path to 7Z.EXE if it is
function test-7zipinstalled {
	try {
		$7zpath=((Get-ItemProperty -ErrorAction Stop -Path HKLM:\SOFTWARE\7-Zip -Name "Path").path)+"7z.exe"
		if (-not (test-path $7zpath)) {
			$7zpath=((Get-ItemProperty -ErrorAction Stop -Path HKLM:\SOFTWARE\7-Zip -Name "Path64").path64)+"7z.exe"
			if (-not (test-path $7zpath)) {
				$7zpath=$null
			}
		}
	} catch {
		$7zpath=$null
	}
	return $7zpath
}

# Test if specified Powershell module is installed. Attempt to instal from the gallery if it isn't
function test-module {
	param (
		[string] $modulename
	)

	# If module is not installed, attempt to install it
	if (-not(Get-InstalledModule $modulename -ErrorAction silentlycontinue)) {
		write-warning "Module $module is not installed. Attempting to install."
		Install-Module -Name $modulename -allowclobber -force -erroraction silentlycontinue
	}
	# Test if installed again and return true or false
	if (-not(Get-InstalledModule $modulename -ErrorAction silentlycontinue)) {
		return $false
	} else {
		return $true
	}
}

# Show the Outlook folder picker. Save all attachments from the selected folder.
function get-attachment {
	$o = New-Object -comobject outlook.application
	$n = $o.GetNamespace("MAPI")

	$f = $n.PickFolder()
	$numemails=$f.items.count
	write-output "Exporting attachments from $numemails emails. Please wait..."
	$f.Items| foreach-object {
		$_.attachments|foreach-object {
			$_.saveasfile((Join-Path $jsonfolder $_.filename))
		}
	}
}

# Extract JSON files from GZ files using 7Z.EXE or the 7ZIP Powershell module
function export-json {
	# Extract the JSON files
	if ($7zexe=test-7zipinstalled) {
		# 7ZIP is installed, so use that to extract the JSON files
		write-output "7Z.EXE is installed. Using it to extract the .GZ files."
		& $7zexe e $jsonfolder -o"$jsonfolder" 2>&1 |out-null
	} elseif (test-module 7Zip4PowerShell) {
		write-output "7Zip powershell module is installed. Using it to extract the .GZ files."
		$gzfiles=get-childitem -Path $jsonfolder -Recurse -Force -filter *.gz
		foreach ($gzfile in $gzfiles) {
			Expand-7Zip $gzfile.fullname -TargetPath $jsonfolder
			$tarname=$gzfile.fullname.replace(".json.gz",".json.tar")
			$jsonname=$tarname.replace(".json.tar",".json")
			rename-item -path $tarname -newname $jsonname
		}
	} else {
		write-output "Neither 7Z.EXE or the 7Zip powershell module could be found. Unable to extract the .GZ files."
		write-output "Install 7Zip and try again."
	}
}

#endregion Functions

# ============================================================================
#region Execute
# ============================================================================

# Check we can write to the specified report path
if (-not (test-writepath $jsonfolder)) {
	write-error "Unable to write to the path $jsonfolder. Aborting."
	return
}

# If reportonly is specified, dont save attachments or extract GZ files. Assumes JSON files are already in the report folder.
if (-not $reportonly) {
	# Prompt for Outlook folder and save all attachments
	write-output "Please select the Outlook folder you want to save the attachments from..."
	get-attachment

	# Check if there are any .GZ files and extract them.
	if (test-path "$jsonfolder\*.gz") {
		export-json
		if (-not $nocleanup) {
			write-output "Cleaning up GZ files."
			remove-item -path "$jsonfolder\*.gz" -force
		}
	} else {
		write-output "No GZ Files found."
	}
}

# Check if there are any JSON files and create report.
if (test-path "$jsonfolder\*.json") {
	$jsonfiles=get-childitem -Path $jsonfolder -Recurse -Force -filter *.json
	$report=@()
	Foreach ($json in $jsonfiles) {
		$jobj=get-content $json.fullname|convertfrom-json
		$orgname=$jobj."organization-name"
		$start=$jobj."date-range"."start-datetime"
		$end=$jobj."date-range"."end-datetime"
		$contact=$jobj."contact-info"
		$reportid=$jobj."report-id"

		$policies=$jobj.policies

		foreach ($policy in $policies) {
			$poltype=$policy."policy"."policy-type"
			$polstring=$policy."policy"."policy-string"
			$poldomain=$policy."policy"."policy-domain"
			$sumsuccess=$policy."summary"."total-successful-session-count"
			$sumfailure=$policy."summary"."total-failure-session-count"

			$properties=[pscustomobject]@{
				"organization-name"=$orgname
				"start-datetime"=$start
				"end-datetime"=$end
				"contact-info"=$contact
				"report-id"=$reportid
				"policy-type"=$poltype
				"policy-string"=(($polstring) -join ' , ')
				"policy-domain"=$poldomain
				"total-successful-session-count"=$sumsuccess
				"total-failure-session-count"=$sumfailure
			}
			$report += $Properties
		}
	}
	#Export the report to a CSV file
	try {
		$report|Export-Csv -path $reportfile -NoTypeInformation
		write-output "Exported report file $reportfile."
		if (-not $nocleanup) {
			write-output "Cleaning up JSON files."
			remove-item -path "$jsonfolder\*.json" -force
		}
	} catch {
		write-output  "Unable to export file $reportfile. Check path and permissions."
		write-output "JSON files not cleaned up. Run again with -reportonly switch to generate the report."
	}
} else {
	write-output "No JSON Files found. Aborting."
}