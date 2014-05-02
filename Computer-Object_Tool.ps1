<#
  Written by Scott Morton
  Date - 5/2/2014
  Can be used and redistributed as long as this header information is retained
  License is GPL 3
#>

$daysOld = 180        # Default # of days since password last set

$date = Get-Date

$array = @()            # used for selected and displayed data
$listMatching = @{}
$listOS = @{}


$creds = Get-Credential


Function Scan()
{
    switch ($numOfDays_DrpText.SelectedIndex) 
    { 
        0 {$daysOld = 180} 
        1 {$daysOld = 90} 
        2 {$daysOld = 60} 
        3 {$daysOld = 45} 
        4 {$daysOld = 30} 
        default {$daysOld = 180}
    }

    Get-ADComputer -Filter * -Credential $creds -Server $Server.Text -Properties Name,CanonicalName,Description,Enabled,Modified,OperatingSystem,PasswordLastSet |
         ForEach-Object {

            Write-Host "Checking - " $_.Name

            if ($_.PasswordLastSet -ne $null -AND ($date - $_.PasswordLastSet).Days -ge $daysOld -AND $_.Enabled -eq $true)
            {
                $_ | Add-Member -MemberType NoteProperty -Name "Pwd Age" -Value ($date - $_.PasswordLastSet).Days -Force

                $listMatching.add($_.CanonicalName,$_) # Use CanonicalName to capture duplicate entries
            
                if ($_.OperatingSystem -ne $null -AND $listOS.ContainsKey($_.OperatingSystem) -ne $true)
                {
                    $listOS.add($_.OperatingSystem,$_.OperatingSystem)
                }
            }


            #if ($listMatching.Count -ge 50) { break } # uncomment to debug
        }

    $Matches.Text = $listMatching.Count.ToString()
    LoadOSs


}

Function Perform_Operation()
{
    foreach ($child in $script:array ) 
    {
        if ($disableObject.Checked)
        {
            write-host "Disabling - "$child.Name
            Set-ADComputer -Identity $child.Name -Credential $creds -Server $Server.Text -enabled $False
        }
        elseif ($deleteObject.Checked)
        {
            Write-Host "Not Implemented"
        }
    }
}

Function Import_CSV()
{
    # Get the file containing the server list
    $fd = New-Object system.windows.forms.openfiledialog
    $fd.showdialog()
    $fd.filename


    # Setup the data
    $script:array = Import-Csv -Path $fd.FileName

}

Function Export_CSV()
{
    # Request the filename to write data to
    $fd = New-Object system.windows.forms.savefiledialog
    $fd.showdialog()
    $fd.filename

    $array | Export-Csv -Path $fd.filename –NoTypeInformation
}

Function LoadOSs()
{
    foreach ($child in $listOS.Values)
    {
        $OSlist.Items.Add($child)
    }
}

Function Display_Selections()
{
    $script:array.Clear()
    foreach ( $child in $listMatching.Values )
    {
        foreach ($item in $OSlist.SelectedItems)
        {
           if ($child.OperatingSystem -eq $item)
           {
                $script:array += $child
                break
           }
        }
    }

    $script:array | Out-GridView
}


# BEGIN view

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$Form1 = New-Object System.Windows.Forms.Form 
$Form1.Text = "Computer Object Tool"
$Form1.MinimumSize = New-Object System.Drawing.Size(800,300)
$Form1.MaximumSize = New-Object System.Drawing.Size(1200,300)
$Form1.StartPosition = "CenterScreen"

$Form1.KeyPreview = $True
$Form1.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$Form1.Close()}})


$ScanButton = New-Object System.Windows.Forms.Button
$ScanButton.Location = New-Object System.Drawing.Size(10,55)
$ScanButton.Size = New-Object System.Drawing.Size(75,23)
$ScanButton.Text = "Scan"
$ScanButton.Add_Click({Scan})
$Form1.Controls.Add($ScanButton)

$ModifyButton = New-Object System.Windows.Forms.Button
$ModifyButton.Location = New-Object System.Drawing.Size(10,210)
$ModifyButton.Size = New-Object System.Drawing.Size(140,23)
$ModifyButton.Text = "Perform Operation"
$ModifyButton.Enabled = $false
$ModifyButton.Add_Click({Perform_Operation})
$Form1.Controls.Add($ModifyButton)

$ImportCSVButton = New-Object System.Windows.Forms.Button
$ImportCSVButton.Location = New-Object System.Drawing.Size(155,210)
$ImportCSVButton.Size = New-Object System.Drawing.Size(140,23)
$ImportCSVButton.Text = "Import CSV"
$ImportCSVButton.Add_Click({Import_CSV;$ScanButton.Enabled = $false;$DisplayButton.Enabled = $false;$ExportCSVButton.Enabled = $false;$ModifyButton.Enabled = $true})
$Form1.Controls.Add($ImportCSVButton)

$DisplayButton = New-Object System.Windows.Forms.Button
$DisplayButton.Location = New-Object System.Drawing.Size(350,210)
$DisplayButton.Size = New-Object System.Drawing.Size(140,23)
$DisplayButton.Text = "Select and Display"
$DisplayButton.Add_Click({Display_Selections;$ModifyButton.Enabled = $true; $ExportCSVButton.Enabled = $true})
$Form1.Controls.Add($DisplayButton)

$ExportCSVButton = New-Object System.Windows.Forms.Button
$ExportCSVButton.Location = New-Object System.Drawing.Size(640,210)
$ExportCSVButton.Size = New-Object System.Drawing.Size(140,23)
$ExportCSVButton.Text = "Export CSV"
$ExportCSVButton.Enabled = $false
$ExportCSVButton.Add_Click({Export_CSV})
$Form1.Controls.Add($ExportCSVButton)

$numOfDays_DrpText= New-Object System.Windows.Forms.ComboBox
$numOfDays_DrpText.Location = New-Object System.Drawing.Size(10,25)
$numOfDays_DrpText.Size = New-Object System.Drawing.Size(50,20)
$numOfDays_DrpText.DropDownHeight = 100
[Void] $numOfDays_DrpText.Items.Add("180")
[Void] $numOfDays_DrpText.Items.Add("90")
[Void] $numOfDays_DrpText.Items.Add("60")
[Void] $numOfDays_DrpText.Items.Add("45")
[Void] $numOfDays_DrpText.Items.Add("30")
$numOfDays_DrpText.SelectedIndex = 0
$Form1.Controls.Add($numOfDays_DrpText)

$numOfDays_Label = New-Object System.Windows.Forms.Label
$numOfDays_Label.Location = New-Object System.Drawing.Size(10,6) 
$numOfDays_Label.Size = New-Object System.Drawing.Size(60,20) 
$numOfDays_Label.Text = "Days Old"
$Form1.Controls.Add($numOfDays_Label) 

$Server = New-Object System.Windows.Forms.TextBox
$Server.Location = New-Object System.Drawing.Size(70,25)
$Server.Size = New-Object System.Drawing.Size(270,20)
$Server.Text = ""
$Form1.Controls.Add($Server) 

$Server_Label = New-Object System.Windows.Forms.Label
$Server_Label.Location = New-Object System.Drawing.Size(70,6) 
$Server_Label.Size = New-Object System.Drawing.Size(270,20) 
$Server_Label.Text = "Server Name or IP address to query"
$Form1.Controls.Add($Server_Label) 

$OSlist = New-Object System.Windows.Forms.ListBox
$OSlist.Location = New-Object System.Drawing.Size(350,25)
$OSlist.Size = New-Object System.Drawing.Size(20,20)
$OSlist.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
$OSlist.Height = 180
$OSlist.Width = 430
$Form1.Controls.Add($OSlist) 

$OSlist_Label = New-Object System.Windows.Forms.Label
$OSlist_Label.Location = New-Object System.Drawing.Size(350,6) 
$OSlist_Label.Size = New-Object System.Drawing.Size(430,20) 
$OSlist_Label.Text = "Select the Operating Systems to modify (ctrl-click for multiple)"
$Form1.Controls.Add($OSlist_Label) 


$Matches = New-Object System.Windows.Forms.Label
$Matches.Location = New-Object System.Drawing.Size(10,90) 
$Matches.Size = New-Object System.Drawing.Size(60,20) 
$Matches.Text = "0"
$Form1.Controls.Add($Matches) 

$Matches_Label = New-Object System.Windows.Forms.Label
$Matches_Label.Location = New-Object System.Drawing.Size(100,90) 
$Matches_Label.Size = New-Object System.Drawing.Size(250,20) 
$Matches_Label.Text = "- Matching based on PasswordLastSet date"
$Form1.Controls.Add($Matches_Label) 

$Operation_Label = New-Object System.Windows.Forms.Label
$Operation_Label.Location = New-Object System.Drawing.Size(10,140) 
$Operation_Label.Size = New-Object System.Drawing.Size(250,20) 
$Operation_Label.Text = "Select the desired operation"
$Form1.Controls.Add($Operation_Label) 

$disableObject = New-Object System.Windows.Forms.RadioButton
$disableObject.Location = New-Object System.Drawing.Size(10,160)
$disableObject.Text = "Disable Objects"
$disableObject.Checked = $true
$Form1.Controls.Add($disableObject)

$deleteObject = New-Object System.Windows.Forms.RadioButton
$deleteObject.Location = New-Object System.Drawing.Size(160,160) 
$deleteObject.Text = "Delete Objects"
$Form1.Controls.Add($deleteObject)

$Form1.Topmost = $True

$Form1.Add_Shown({$Form1.Activate()})

[void]$Form1.ShowDialog()



