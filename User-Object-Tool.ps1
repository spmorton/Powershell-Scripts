<#
  Version 2
  Written by Scott Morton
  Date - 5/2/2014
#>

$UOTVersion = "2.1"

$filters = $false

$date = Get-Date

$array = @()            # used for selected and displayed data
$listMatching = @{}
$failures = @{}
$stop = $false

$creds = Get-Credential

Function Init_Sys()
{
    $filters = $false

    $date = Get-Date

    $array = @()            # used for selected and displayed data
    $listMatching = @{}
    $failures = @{}
    $ScanButton.Enabled = $true
    $ModifyButton.Enabled = $false
    $ExportCSVButton.Enabled = $false
    $DisplayButton.Enabled = $false

    $Selected.Text = $script:array.Count.ToString()
    $Matches.Text = $listMatching.Count.ToString()

}

Function Scan()
{
    $ScanButton.Enabled = $false

    switch ($numOfDays_DrpText.SelectedIndex) 
    { 
        0 {$daysOld = 180} 
        1 {$daysOld = 90} 
        2 {$daysOld = 60} 
        3 {$daysOld = 45} 
        4 {$daysOld = 30} 
        default {$daysOld = 180}
    }

    Get-ADUser -Filter * -Credential $creds -Server $Server.Text -Properties Name,CanonicalName,Description,Enabled,LastLogonDate,lastLogonTimeStamp,Modified,modifyTimeStamp,PasswordLastSet,pwdLastSet |
         ForEach-Object {

            Write-Host "Checking - " $_.Name

            if ($Disabled_Check.Checked)
            {
                    switch ($ModifiedDate_DrpText.SelectedIndex) 
                { 
                    0 {$LastModifiedDate = 180} 
                    1 {$LastModifiedDate = 90} 
                    2 {$LastModifiedDate = 60} 
                    3 {$LastModifiedDate = 45} 
                    4 {$LastModifiedDate = 30} 
                    default {$LastModifiedDate = 180}
                }
                if (($date - $_.modifyTimeStamp).Days -ge $LastModifiedDate -AND $_.Enabled -eq $false)
                {
                    $_ | Add-Member -MemberType NoteProperty -Name "Days Since Last Mod" -Value ($date - $_.modifyTimeStamp).Days -Force
                    $listMatching.add($_.CanonicalName,$_) # Use CanonicalName to capture duplicate entries
                }
            }

            else
            {
                if (($date - ([datetime]::FromFileTime($_.lastLogonTimeStamp))).Days -ge $daysOld -AND $_.Enabled -eq $true)
                {
                    if ($_.pwdLastSet -ne $null)
                    {
                        $_ | Add-Member -MemberType NoteProperty -Name "Pwd Age" -Value ($date - ([datetime]::FromFileTime($_.pwdLastSet))).Days -Force
                    }
                    else
                    {
                        $_ | Add-Member -MemberType NoteProperty -Name "Pwd Age" -Value $null -Force
                    }

                    $listMatching.add($_.CanonicalName,$_) # Use CanonicalName to capture duplicate entries
                }
            }
            
            # uncomment to debug
            #if ($listMatching.Count -ge 50) { break } 


        }
}

Function Perform_Operation()
{

    foreach ($child in $listMatching ) 
    {
        if ($disableObject.Checked)
        {
            try
            {
                write-host "Disabling - "$child.SamAccountName
                Set-ADUser -Identity $child.SamAccountName -Credential $creds -Server $Server.Text -enabled $False
            }

            catch
            {
                $failures.Add($child.SamAccountName,$child)
            }

        }
        elseif ($deleteObject.Checked)
        {
            try
            {
                Remove-ADUser -Identity $child.SamAccountName -Credential $creds -Server $Server.Text -Confirm:$False
            }
            catch
            {
                $failures.Add($child.SamAccountName,$child)
            }
        }
    }

    if ($failures.Count)
    {
        $OUTPUT = [System.Windows.Forms.MessageBox]::Show("Modification failures detected, click Yes to select destination file for report and no to disregard", "Status", 4)
        if ($OUTPUT -eq "YES")
        {
            # Request the filename to write data to
            $fd = New-Object system.windows.forms.savefiledialog
            $fd.showdialog()
            $fd.filename

            $failures.Values | Export-Csv -Path $fd.filename –NoTypeInformat
        }
    }

    [System.Windows.Forms.MessageBox]::Show("Modification process completed", "Status")

}

Function Import_CSV()
{
    # Get the file containing the server list
    $fd = New-Object system.windows.forms.openfiledialog
    $fd.showdialog()
    $fd.filename


    # Setup the data
    $array = Import-Csv -Path $fd.FileName

    [System.Windows.Forms.MessageBox]::Show("CSV import completed", "Status")

}

Function Export_CSV()
{
    # Request the filename to write data to
    $fd = New-Object system.windows.forms.savefiledialog
    $fd.showdialog()
    $fd.filename

    $listMatching.Values | Export-Csv -Path $fd.filename –NoTypeInformation

    [System.Windows.Forms.MessageBox]::Show("Export CSV completed", "Status")

}


Function Display_Selections()
{
    $listMatching.Values | Out-GridView
    
}


Function Filters()
{
    Write-Host "Applying Filters"
    if ($ModifiedDate_DrpText.Enabled)
    {
        switch ($ModifiedDate_DrpText.SelectedIndex) 
        { 
            0 {$LastModifiedDate = 180} 
            1 {$LastModifiedDate = 90} 
            2 {$LastModifiedDate = 60} 
            3 {$LastModifiedDate = 45} 
            4 {$LastModifiedDate = 30} 
            default {$LastModifiedDate = 180}
        }
        
        foreach ($child in $listMatching.Values)
        {
            if (($date - $_.modifyTimeStamp).Days -le $LastModifiedDate)
            {
                $listMatching.Remove($child.CanonicalName)
            }
        }
    }
}


# BEGIN view

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$Form1 = New-Object System.Windows.Forms.Form 
$Form1.Text = "User Object Tool - " + $UOTVersion
$Form1.MinimumSize = New-Object System.Drawing.Size(360,520)
$Form1.MaximumSize = New-Object System.Drawing.Size(360,520)
$Form1.StartPosition = "CenterScreen"
$Form1.KeyPreview = $True


$ScanButton = New-Object System.Windows.Forms.Button
$ScanButton.Location = New-Object System.Drawing.Size(10,200)
$ScanButton.Size = New-Object System.Drawing.Size(75,25)
$ScanButton.Text = "Scan"
$ScanButton.Add_Click(
                        {
                            $StopButton.Enabled = $true
                            Scan
 
                            if ($filters)
                            {
                                Filters
                            } 


                            $Matches.Text = $listMatching.Count.ToString()

                            $DisplayButton.Enabled = $true
                            $ExportCSVButton.Enabled = $true


                            [System.Windows.Forms.MessageBox]::Show("Scan completed", "Status")

                        })

$Form1.Controls.Add($ScanButton)

$StopButton = New-Object System.Windows.Forms.Button
$StopButton.Location = New-Object System.Drawing.Size(90,200)
$StopButton.Size = New-Object System.Drawing.Size(140,25)
$StopButton.Text = "Stop Scan"
$StopButton.Enabled = $false
$StopButton.Add_Click({$stop = $true})
$Form1.Controls.Add($StopButton)

$ModifyButton = New-Object System.Windows.Forms.Button
$ModifyButton.Location = New-Object System.Drawing.Size(10,445)
$ModifyButton.Size = New-Object System.Drawing.Size(140,25)
$ModifyButton.Text = "Perform Operation"
$ModifyButton.Enabled = $false
$ModifyButton.Add_Click({Perform_Operation})
$Form1.Controls.Add($ModifyButton)

$ImportCSVButton = New-Object System.Windows.Forms.Button
$ImportCSVButton.Location = New-Object System.Drawing.Size(155,445)
$ImportCSVButton.Size = New-Object System.Drawing.Size(140,25)
$ImportCSVButton.Text = "Import CSV"
$ImportCSVButton.Add_Click({Import_CSV;$ScanButton.Enabled = $false;$DisplayButton.Enabled = $false;$ExportCSVButton.Enabled = $false;$ModifyButton.Enabled = $true})
$Form1.Controls.Add($ImportCSVButton)

$DisplayButton = New-Object System.Windows.Forms.Button
$DisplayButton.Location = New-Object System.Drawing.Size(10,280)
$DisplayButton.Size = New-Object System.Drawing.Size(140,25)
$DisplayButton.Text = "Display Accounts"
$DisplayButton.Add_Click({Display_Selections;$ModifyButton.Enabled = $true; $ExportCSVButton.Enabled = $true})
$DisplayButton.Enabled = $false
$Form1.Controls.Add($DisplayButton)

$ExportCSVButton = New-Object System.Windows.Forms.Button
$ExportCSVButton.Location = New-Object System.Drawing.Size(160,280)
$ExportCSVButton.Size = New-Object System.Drawing.Size(140,25)
$ExportCSVButton.Text = "Export CSV"
#$ExportCSVButton.Enabled = $false
$ExportCSVButton.Add_Click({Export_CSV})
$Form1.Controls.Add($ExportCSVButton)

$ResetButton = New-Object System.Windows.Forms.Button
$ResetButton.Location = New-Object System.Drawing.Size(10,315)
$ResetButton.Size = New-Object System.Drawing.Size(100,25)
$ResetButton.Text = "Reset"
$ResetButton.Enabled = $true
$ResetButton.Add_Click({Init_Sys})
$Form1.Controls.Add($ResetButton)

$Server = New-Object System.Windows.Forms.TextBox
$Server.Location = New-Object System.Drawing.Size(10,25)
$Server.Size = New-Object System.Drawing.Size(270,20)
$Server.Text = ""
$Form1.Controls.Add($Server)

$Server_Label = New-Object System.Windows.Forms.Label
$Server_Label.Location = New-Object System.Drawing.Size(10,6) 
$Server_Label.Size = New-Object System.Drawing.Size(270,20) 
$Server_Label.Text = "Server Name or IP address to query"
$Form1.Controls.Add($Server_Label) 

$Filters_Label = New-Object System.Windows.Forms.Label
$Filters_Label.Location = New-Object System.Drawing.Size(10,50) 
$Filters_Label.Size = New-Object System.Drawing.Size(270,20) 
$Filters_Label.Text = "---------------------- Filters --------------------------------------"
$Form1.Controls.Add($Filters_Label) 

$numOfDays_DrpText= New-Object System.Windows.Forms.ComboBox
$numOfDays_DrpText.Location = New-Object System.Drawing.Size(10,85)
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
$numOfDays_Label.Location = New-Object System.Drawing.Size(70,80) 
$numOfDays_Label.Size = New-Object System.Drawing.Size(280,30) 
$numOfDays_Label.Text = "Days since LastLogonTimestamp for unused accounts (Primary search parameter per MicroSoft)"
$Form1.Controls.Add($numOfDays_Label) 

$LastModifiedDate_Check = New-Object System.Windows.Forms.CheckBox
$LastModifiedDate_Check.Location = New-Object System.Drawing.Size(10,120)
$LastModifiedDate_Check.Size = New-Object System.Drawing.Size(120,30)
$LastModifiedDate_Check.Text = "Enable LastModifiedDate"
$LastModifiedDate_Check.Add_CheckStateChanged({
    if ($LastModifiedDate_Check.Checked)
    {
        $ModifiedDate_DrpText.Enabled = $true
        $ModifiedDateDrpBx_Label.Enabled = $true
    }
    else
    {
        $ModifiedDate_DrpText.Enabled = $false
        $ModifiedDateDrpBx_Label.Enabled = $false
    }
    })
$Form1.Controls.Add($LastModifiedDate_Check)

$ModifiedDate_DrpText= New-Object System.Windows.Forms.ComboBox
$ModifiedDate_DrpText.Location = New-Object System.Drawing.Size(130,125)
$ModifiedDate_DrpText.Size = New-Object System.Drawing.Size(50,20)
$ModifiedDate_DrpText.DropDownHeight = 100
[Void] $ModifiedDate_DrpText.Items.Add("180")
[Void] $ModifiedDate_DrpText.Items.Add("90")
[Void] $ModifiedDate_DrpText.Items.Add("60")
[Void] $ModifiedDate_DrpText.Items.Add("45")
[Void] $ModifiedDate_DrpText.Items.Add("30")
$ModifiedDate_DrpText.SelectedIndex = 0
$ModifiedDate_DrpText.Enabled = $false
$Form1.Controls.Add($ModifiedDate_DrpText)

$ModifiedDateDrpBx_Label = New-Object System.Windows.Forms.Label
$ModifiedDateDrpBx_Label.Location = New-Object System.Drawing.Size(180,120) 
$ModifiedDateDrpBx_Label.Size = New-Object System.Drawing.Size(160,30) 
$ModifiedDateDrpBx_Label.Text = "Select the number of days since last modified"
$ModifiedDateDrpBx_Label.Enabled = $false
$Form1.Controls.Add($ModifiedDateDrpBx_Label) 

$Disabled_Check = New-Object System.Windows.Forms.CheckBox
$Disabled_Check.Location = New-Object System.Drawing.Size(10,160)
$Disabled_Check.Size = New-Object System.Drawing.Size(300,30)
$Disabled_Check.Text = "Search for disabled accounts with LastModifiedDate date greater than selected number of days"
$Disabled_Check.Add_CheckStateChanged({
    if ($Disabled_Check.Checked)
    {
        $LastModifiedDate_Check.Checked = $true
        $numOfDays_DrpText.Enabled = $false
        $deleteObject.Checked = $true
        $disableObject.Enabled = $false
    }
    else
    {
        $LastModifiedDate_Check.Checked = $false
        $numOfDays_DrpText.Enabled = $true
        $disableObject.Enabled = $true
        $disableObject.Checked = $true
    }
    })
$Form1.Controls.Add($Disabled_Check)

$Matches = New-Object System.Windows.Forms.Label
$Matches.Location = New-Object System.Drawing.Size(10,240) 
$Matches.Size = New-Object System.Drawing.Size(60,20) 
$Matches.Text = "0"
$Form1.Controls.Add($Matches) 

$Matches_Label = New-Object System.Windows.Forms.Label
$Matches_Label.Location = New-Object System.Drawing.Size(100,240) 
$Matches_Label.Size = New-Object System.Drawing.Size(250,20) 
$Matches_Label.Text = "- Matching users"
$Form1.Controls.Add($Matches_Label) 

$Operation_Label = New-Object System.Windows.Forms.Label
$Operation_Label.Location = New-Object System.Drawing.Size(10,380) 
$Operation_Label.Size = New-Object System.Drawing.Size(250,20) 
$Operation_Label.Text = "Select the desired operation"
$Form1.Controls.Add($Operation_Label) 

$disableObject = New-Object System.Windows.Forms.RadioButton
$disableObject.Location = New-Object System.Drawing.Size(10,400)
$disableObject.Text = "Disable Objects"
$disableObject.Checked = $true
$Form1.Controls.Add($disableObject)

$deleteObject = New-Object System.Windows.Forms.RadioButton
$deleteObject.Location = New-Object System.Drawing.Size(160,400) 
$deleteObject.Text = "Delete Objects"
$Form1.Controls.Add($deleteObject)

$Form1.Topmost = $False

Init_Sys 

$Form1.Add_Shown({$Form1.Activate()})

[void]$Form1.ShowDialog()



