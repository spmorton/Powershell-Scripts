# BEGIN access Section

# Connect to the API and authenticate.
$CookieContainer = New-Object System.Net.CookieContainer
$wsdlProxy = New-WebServiceProxy -uri $($wsdlPath)
$wsdlProxy.url = $wsdlPath
$wsdlProxy.CookieContainer = $CookieContainer

#!!!!!!! Begin Login 
#Production
#$wsdlPath = "http://your-server/Services/API?wsdl"
#$cred = Get-Credential
#$wsdlProxy.login($cred.UserName, $cred.GetNetworkCredential().password)


# user must be api enabled
# Test
$wsdlPath = "http://your-server/Services/API?wsdl"
$apiUsername = "api-admin"
$apiPassword = "admin"
$wsdlProxy.login($apiUsername, $apiPassword)

# END access Section

Function Reset()
{
	$script:selected = $wsdlProxy.getEntities(0, "Configuration", 0, 20)
}

Function Load_Listbox($list)
{
    foreach($item in $list)
    {
        $entry = ("id = ({0})    name = ({1}) properties = ({2})    type = ({3})" -f $item.id.ToString(), $item.name, $item.properties, $item.type )
       
        [Void] $objListBox.Items.Add($entry)
    }
    
}

Function Select_Obj()
{
    if ($script:selected[$objListBox.SelectedIndex].type -eq "View") 
    {
     $script:selectedDNS = $script:selected[$objListBox.SelectedIndex]
	 $DNStxt.text = $objListBox.SelectedItem
    }
    elseif ($script:selected[$objListBox.SelectedIndex].type -eq "IP4Network") 
    {
     $script:selectedIPNet = $script:selected[$objListBox.SelectedIndex]
	 $NETtxt.text = $objListBox.SelectedItem
    }
    else
    {
     $script:selectedOther = $script:selected[$objListBox.SelectedIndex]
	 $Othertxt.text = $objListBox.SelectedItem
    }
}

Function runit()
{
    $script:selectedItem = $script:selected[$objListBox.SelectedIndex]
    if (!$DNSButton.Enabled) 
    {
        $bs = $wsdlProxy.getEntities($script:selectedItem.id, "View", 0, 20)
    }
    elseif (!$NetButton.Enabled) 
    {
        $bs = $wsdlProxy.getEntities($script:selectedItem.id, $DrpText.SelectedItem, 0, 2000)
    }
    elseif (!$SearchButton.Enabled) 
    {
        $bs = $wsdlProxy.searchByObjectTypes($SearchBox.Text, $DrpText.SelectedItem, 0, 2000)
    }
    elseif (!$AssignIPButton.Enabled) 
    {
        Assign_IP4
    }
	
	if ($AssignIPButton.Enabled)
	{
		$objListBox.Items.Clear()
	    Load_Listbox $bs
	    
		$script:selected = $bs
	}
}

Function Form_Controls($view)
{
    switch ($view) 
        { 
            "DNS"           {
								Reset
                                $objListBox.Items.Clear()
                                $DNSButton.Enabled = $false
                                $NetButton.Enabled = $true
                                $SearchButton.Enabled = $true
                                $AssignIPButton.Enabled = $true
                                $begin = $wsdlProxy.getEntities(0, "Configuration", 0, 10)
                                Load_Listbox $begin
								$DrpText.Visible = $false
								$SearchBox.Visible = $false
                            } 
            "Net"        {
								Reset
                                $objListBox.Items.Clear()
                                $DNSButton.Enabled = $true
                                $NetButton.Enabled = $false
                                $SearchButton.Enabled = $true
                                $AssignIPButton.Enabled = $true
								$DrpText.items.clear()
								[Void] $DrpText.Items.Add("IP4Block")
								[Void] $DrpText.Items.Add("IP4Network")
								$DrpText.SelectedIndex = 0
								$DrpText.Visible = $true
                                $begin = $wsdlProxy.getEntities(0, "Configuration", 0, 10)
                                Load_Listbox $begin
								$SearchBox.Visible = $false
                            } 
            "Search"        {
                                $objListBox.Items.Clear()
                                $DNSButton.Enabled = $true
                                $NetButton.Enabled = $true
                                $SearchButton.Enabled = $false
                                $AssignIPButton.Enabled = $true
								$DrpText.Visible = $true
								$SearchBox.Visible = $true
								$DrpText.items.clear()
								[Void] $DrpText.Items.Add("IP4Block")
								[Void] $DrpText.Items.Add("IP4Network")
								[Void] $DrpText.Items.Add("IP4Address")
								[Void] $DrpText.Items.Add("Configuration")
								[Void] $DrpText.Items.Add("View")
								$DrpText.SelectedIndex = 0
                            } 
            "AssignIP"        {
                                $objListBox.Items.Clear()
                                $DNSButton.Enabled = $true
                                $NetButton.Enabled = $true
                                $SearchButton.Enabled = $true
                                $AssignIPButton.Enabled = $false
								$DrpText.Visible = $false
								$SearchBox.Visible = $false
                            } 
            default         {"error"} # Error
        }
}

Function GetCSV_File()
{
	# Request the filename to write data to
	$fd = New-Object system.windows.forms.openfiledialog
	$fd.showdialog()
	$fd.filename
	$script:csvList = Import-Csv -Path $fd.filename
	$objListBox.Items.Clear()
	foreach($item in $csvList)
    {
        $entry = ("deviceName = ({0})    deviceMac = ({1})    ipAddress = ({2})" -f $item.deviceName, $item.deviceMac, $item.ipAddress)
       
        [Void] $objListBox.Items.Add($entry)
    }
}

Function Assign_IP4()
{	
	foreach ($item in $script:csvList)
	{
		
		$wsdlProxy.assignIP4Address($script:selectedOther.id,$item.ipAddress,$item.deviceMac, $item.deviceName,"MAKE_DHCP_RESERVED","name=$($item.deviceName)")
		$wsdlProxy.addHostRecord($script:selectedDNS.id, $item.deviceName, $item.ipAddress, 900, "")
	}
}

Function Check_Reqs()
{
	if ($script:selectedDNS -eq $null -or $script:selectedOther -eq $null)
	{
		[Void] [System.Windows.Forms.MessageBox]::Show("Must select a DNS View and a Configuration to complete","Requirements Not Met","OK","Warning")
		Form_Controls "DNS"
		return $false
	}
	return $true
}

# BEGIN view

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Proteus Selector"
$objForm.MinimumSize = New-Object System.Drawing.Size(600,600)
$objForm.MaximumSize = New-Object System.Drawing.Size(600,600)
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {runit}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(20,545)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(100,545)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)

$DNSButton = New-Object System.Windows.Forms.Button
$DNSButton.Location = New-Object System.Drawing.Size(10,40)
$DNSButton.Size = New-Object System.Drawing.Size(75,23)
$DNSButton.Text = "DNS"
$DNSButton.Add_Click({Form_Controls("DNS")})
$objForm.Controls.Add($DNSButton)

$NetButton = New-Object System.Windows.Forms.Button
$NetButton.Location = New-Object System.Drawing.Size(90,40)
$NetButton.Size = New-Object System.Drawing.Size(75,23)
$NetButton.Text = "IP Net"
$NetButton.Add_Click({Form_Controls("Net")})
$objForm.Controls.Add($NetButton)

$SearchButton = New-Object System.Windows.Forms.Button
$SearchButton.Location = New-Object System.Drawing.Size(170,40)
$SearchButton.Size = New-Object System.Drawing.Size(75,23)
$SearchButton.Text = "Search"
$SearchButton.Add_Click({Form_Controls("Search")})
$objForm.Controls.Add($SearchButton)

$AssignIPButton = New-Object System.Windows.Forms.Button
$AssignIPButton.Location = New-Object System.Drawing.Size(250,40)
$AssignIPButton.Size = New-Object System.Drawing.Size(75,23)
$AssignIPButton.Text = "AssignIP"
$AssignIPButton.Add_Click(
{
	Form_Controls("AssignIP")
	$test = Check_Reqs
	if ($test)
	{
		GetCSV_File
	}
})
$objForm.Controls.Add($AssignIPButton)

$DrpText= New-Object System.Windows.Forms.ComboBox
$DrpText.Location = New-Object System.Drawing.Size(10,70)
$DrpText.Size = New-Object System.Drawing.Size(100,20)
$DrpText.DropDownHeight = 200
$objForm.Controls.Add($DrpText)

$GoButton = New-Object System.Windows.Forms.Button
$GoButton.Location = New-Object System.Drawing.Size(10,345)
$GoButton.Size = New-Object System.Drawing.Size(75,23)
$GoButton.Text = "Go"
$GoButton.Add_Click({runit})
$objForm.Controls.Add($GoButton)

$SelectButton = New-Object System.Windows.Forms.Button
$SelectButton.Location = New-Object System.Drawing.Size(90,345)
$SelectButton.Size = New-Object System.Drawing.Size(75,23)
$SelectButton.Text = "Select"
$SelectButton.Add_Click({Select_Obj})
$objForm.Controls.Add($SelectButton)

$ResetButton = New-Object System.Windows.Forms.Button
$ResetButton.Location = New-Object System.Drawing.Size(170,345)
$ResetButton.Size = New-Object System.Drawing.Size(75,23)
$ResetButton.Text = "Reset"
$ResetButton.Add_Click(
	{
	    if (!$DNSButton.Enabled) 
	    {
			Form_Controls "DNS"
	    }
	    if (!$NetButton.Enabled) 
	    {
			Form_Controls "Net"
	    }
	    if (!$SearchButton.Enabled) 
	    {
			Form_Controls "Search"
	    }
	    if (!$AssignIPButton.Enabled) 
	    {
			Form_Controls "AssignIP"
	    }
	})
$objForm.Controls.Add($ResetButton)

$TxtBoxLabel = New-Object System.Windows.Forms.Label
$TxtBoxLabel.Location = New-Object System.Drawing.Size(10,20) 
$TxtBoxLabel.Size = New-Object System.Drawing.Size(280,20) 
$TxtBoxLabel.Text = "Select Operation"
$objForm.Controls.Add($TxtBoxLabel) 

$objListBox = New-Object System.Windows.Forms.ListBox 
$objListBox.Location = New-Object System.Drawing.Size(10,100) 
$objListBox.Size = New-Object System.Drawing.Size(570,3000) 
$objListBox.Height = 250
$objListBox.HorizontalScrollbar = $true

$SearchBox = New-Object System.Windows.Forms.TextBox 
$SearchBox.Location = New-Object System.Drawing.Size(120,70) 
$SearchBox.Size = New-Object System.Drawing.Size(320,20) 
$objForm.Controls.Add($SearchBox) 

$DNSLabel = New-Object System.Windows.Forms.Label
$DNSLabel.Location = New-Object System.Drawing.Size(10,420) 
$DNSLabel.Size = New-Object System.Drawing.Size(110,20) 
$DNSLabel.Text = "Selected DNS View: "
$objForm.Controls.Add($DNSLabel) 

$DNStxt = New-Object System.Windows.Forms.Label
$DNStxt.Location = New-Object System.Drawing.Size(120,420) 
$DNStxt.Size = New-Object System.Drawing.Size(470,20) 
$DNStxt.Text = ""
$objForm.Controls.Add($DNStxt) 

$NETLabel = New-Object System.Windows.Forms.Label
$NETLabel.Location = New-Object System.Drawing.Size(10,450) 
$NETLabel.Size = New-Object System.Drawing.Size(110,20) 
$NETLabel.Text = "Selected Network: "
$objForm.Controls.Add($NETLabel) 

$NETtxt = New-Object System.Windows.Forms.Label
$NETtxt.Location = New-Object System.Drawing.Size(120,450) 
$NETtxt.Size = New-Object System.Drawing.Size(470,80) 
$NETtxt.Text = ""
$objForm.Controls.Add($NETtxt) 

$OtherLabel = New-Object System.Windows.Forms.Label
$OtherLabel.Location = New-Object System.Drawing.Size(10,390) 
$OtherLabel.Size = New-Object System.Drawing.Size(110,20) 
$OtherLabel.Text = "Selected Other: "
$objForm.Controls.Add($OtherLabel) 

$Othertxt = New-Object System.Windows.Forms.Label
$Othertxt.Location = New-Object System.Drawing.Size(120,390) 
$Othertxt.Size = New-Object System.Drawing.Size(470,80) 
$Othertxt.Text = ""
$objForm.Controls.Add($Othertxt) 

$objForm.Controls.Add($objListBox) 

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
Form_Controls "DNS"
[void]$objForm.ShowDialog()

$script:selectedDNS

$script:selectedIPNET

$script:selectedOther

#!!!! Create a host record

# params are (viewID, FQDN, IPADDR, TTL, properties)
# view ID is in the details tab of a Proteus DNS view
# example navigate to CHS - Internal > Internal
# click the details tab and retrieve the Object ID - 196923
# the ID below is for the test env.

#$wsdlProxy.addHostRecord(3660, "testme222.chs.net", "172.27.1.5", 900, "")

