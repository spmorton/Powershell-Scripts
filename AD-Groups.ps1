# Author - Scott P. Morton Ph.D.
# Date - on or about 4/20/2023
# Searches for circular memberships of groups
# See methods below, adjust the following parameters to suit
# your needs
#
# Copyright (C) 2023  Scott P. Morton PhD (spm3c at mtmail.mtsu.edu)
# 
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
# 
# The output file 'cyclical-groups.csv' has 'DistinguishedName' (DN)
# and 'Zgroupo x' The DN is the searched group, each 'Zgroup x' is
# an endpoint group that creates a loop back to the DN,
# where x is 0 through maxZgroups parameter below.
#
# The output file 'max-depth-groups.txt' can be used to narrow the
# succesive searches to a limited number of groups that exceed
# maxDepth through method 3 to search for the 'bottom'
#
# The output file 'error-groups.txt' needs no explanation

# Specific details of any DistinguishedName you want 
# to exclude from the search, like a foreign domain
$FD = "DC=adomain,DC=com"

# Maximum recursive depth, start shallow and work your way up
# or you'll be sorry
$script:depthLimit = 1

# Max number of endpoints to add to any group entry (DN) that has loops
$maxZgroups = 10

# dupGroups contains all cyclical groups
$dupGroups = @()
# errorGroups... duh
$script:errorGrpsp = @{}
# Any group that reaches max recursive depth
$script:depthGroupsm = @{}

# !!!! METHODS !!!!
# Identifies method to run: 
# Full domain (1)
# Subset of domain (2)
# Single group (3)
# Single user (4)
$method = 3

if ($method -eq 1) {
    $groups = Get-ADGroup -filter *
}
elseif ($method -eq 2) {
    $groups = Get-ADGroup -filter "Name -like '*Citrix*'"
}
elseif ($method -eq 3) {
    $id = "CN=anobject,OU=org1,OU=org2,DC=subdom,DC=dom,DC=net"
    $groups = Get-ADGroup $id
}
elseif ($method -eq 4) {
    $groups = @()
    $x = Get-ADUser auser -Properties "MemberOf"
    foreach ($id in $x["MemberOf"]) {
        $groups += Get-ADGroup $id
    }
}
else {
    Write-Host "Invalid Method chosen, check 'method' value"
    $groups = @()
}

function getGroups {
    param (
        $thisGroup,$count
    )
    $theseGroups = @()
    if($count -eq $script:depthLimit) {
        #Write-Host "Max depth reached at $thisGroup"
        if (-Not $script:depthGroupsm.Contains($thisGroup)){
            $script:depthGroupsm.Add($thisGroup,1)
        }
        return $thisGroup
    }
    if($thisGroup.Contains($FD)) {
        Write-Host "foreign domain group $thisGroup"
        return $theseGroups
    }
    try {
        $members = Get-ADGroupMember $thisGroup
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        #write-host "bad group $thisGroup"
        if (-Not $script:errorGrpsp.Contains($thisGroup)){
            $script:errorGrpsp.Add($thisGroup,1)
        }
        return $theseGroups
    }
    catch {
        # We have likely exceeded the 5k limit and must do this manually
        $members = @()
        $x = Get-ADGroup -Identity $thisGroup -Properties "Members"
        foreach ($i in $x["Members"]) {
            $members += Get-ADObject -Identity ($i)
        }
    }
    $count += 1
    foreach ($member in $members) {
        if ($member.ObjectClass -eq "group") {
            if($member.distinguishedName.Contains($FD)) {
                continue
            }
            $theseGroups += $member.DistinguishedName
            $theseGroups += getGroups $member.DistinguishedName $count
        }
    }
    return $theseGroups
}

foreach ($group in $groups) {
    $groupDN = @{}
    $count = 0
    $groupDN.Add($group.DistinguishedName,1)
    Write-Host "Processing - $group"
    $allSubGrps = getGroups $group.DistinguishedName 0
    if($allSubGrps.Count) {
        for ($i = 0; $i -lt $maxZgroups; $i++) {
            try {
                $group | Add-Member -MemberType NoteProperty -Name "Zgroup $i" -Value "" -Force
            }
            catch {
                continue
            }
        }
        foreach ($subGrp in $allSubGrps) {
            try {
                $groupDN.Add($subGrp,1)
            }
            catch {
                continue
            }
            try {
                $thisGroup = Get-ADGroup -Identity $subGrp -Properties "Members"
            }
            catch {
                if (-Not $script:errorGrpsp.Contains($subGrp)){
                    $script:errorGrpsp.Add($subGrp,1)
                }
                continue
            }
            if ($thisGroup["Members"].Contains($group.DistinguishedName)) {
                $group  | Add-Member -MemberType NoteProperty -Name "Zgroup $count" -Value $subGrp -Force
                $count += 1
            }
        }
    }
    # Keeps a running tab of findings in case of failure or manual stop
    if($count) {
        $dupGroups += $group
        $dupGroups | Export-Csv -Path "cyclical-groups.csv" -NoTypeInformat -Force
    }
    if ($script:errorGrpsp.Count) {
        $script:errorGrpsp.Keys | Out-File -FilePath "error-groups.txt" -Force
    }
    if ($script:depthGroupsm.Count) {
        $script:depthGroupsm.Keys | Out-File -FilePath "max-depth-groups.txt" -Force
    }
}

