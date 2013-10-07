<#
	Written by 	-	Scott Morton
	Date 		-	10/7/2013
	Notes		-	This script queries the registry for the value of the 
                    specified entry and writes out to a CSV. The source 
                    CSV file must have a header labeled Server
	
#>

# Query the user for elevated privelages
$Cred = Get-Credential

# Request the filename to read server list
$readF = New-Object system.windows.forms.openfiledialog
$readF.showdialog()
$readF.filename

# Request the filename to write data to
$writeF = New-Object system.windows.forms.savefiledialog
$writeF.showdialog()
$writeF.filename


# Create the CSV objects
$servers = Import-Csv $readF.FileName
$results = @()

# Loop through each server
foreach ($Server in $servers ) 
{   
	Write-Host "Processing server -" $Server.Server

	$rlist = New-Object System.Object
	$rlist | Add-Member -MemberType NoteProperty -Name "Server" -Value $Server.Server

    $entry = $(Get-ItemProperty "hklm:\SYSTEM\CurrentControlSet\Control\Terminal Server\").fDenyTSConnections
				

	$rlist | Add-Member -MemberType NoteProperty -Name "fDenyTSConnections" -Value $entry

  # Uncomment if you start having connection issues
	#Start-Sleep -Seconds 1
	
	$results += $rlist
}


 $results | Export-Csv -NoTypeInformation -Path $writeF.filename

