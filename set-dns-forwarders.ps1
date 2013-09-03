<#
	Written by 	-	Scott Morton
	Date 		-	8/26/2013
	Notes		-	This script sets the list of DNS forwarding addresses
                    for the list of stored in a CSV file.
	
    Prereqs     -   Requires Windows server 2012 with remote DNS management tools
                    A CSV file with one column of server names or IP addresses w/no header
                    The forwarding addresses assigned to the four variables below

#>

# !!! Set the four IP addresses used for fowarding here !!!

$FDR1 = ""
$FDR2 = ""
$FDR3 = ""
$FDR4 = ""

# Query the user for elevated privelages
#$Cred = Get-Credential

# Get the file containing the server list
$fd = New-Object system.windows.forms.openfiledialog
$fd.showdialog()
$fd.filename

# Setup the data
$DNSServers = @()
$DNSServers = Import-Csv -Path $fd.FileName -Header "Server"

# Setup forwarder addresses into an array of IPAddresses
$faddrs = [IPAddress]::parse($FDR1),[IPAddress]::parse($FDR2),[IPAddress]::parse($FDR3),[IPAddress]::parse($FDR4)

# Loop through each server
foreach ($child in $DNSServers ) 
{

    Write-Host 'Processing server -' $child.Server

    try 
    {
        set-DnsServerForwarder -IPAddress $faddrs -ComputerName $child.Server
    }

    catch
    {
        $ErrorMessage = $_.Exception.Message

        Write-Host "The failed to write the DNS forwarding entries"
        Write-Host "The message was: $ErrorMessage"

    }
}
 
