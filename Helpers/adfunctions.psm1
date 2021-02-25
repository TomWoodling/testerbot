function Get-ADDirectReports
{
	<#
	.SYNOPSIS
		This function retrieve the directreports property from the IdentitySpecified.
		Optionally you can specify the Recurse parameter to find all the indirect
		users reporting to the specify account (Identity).
	
	.DESCRIPTION
		This function retrieve the directreports property from the IdentitySpecified.
		Optionally you can specify the Recurse parameter to find all the indirect
		users reporting to the specify account (Identity).
	
	.NOTES
		Francois-Xavier Cat
		www.lazywinadmin.com
		@lazywinadm
	
		VERSION HISTORY
		1.0 2014/10/05 Initial Version
	
	.PARAMETER Identity
		Specify the account to inspect
	
	.PARAMETER Recurse
		Specify that you want to retrieve all the indirect users under the account
	
	.EXAMPLE
		Get-ADDirectReports -Identity Test_director
	
Name                SamAccountName      Mail                Manager
----                --------------      ----                -------
test_managerB       test_managerB       test_managerB@la... test_director
test_managerA       test_managerA       test_managerA@la... test_director
		
	.EXAMPLE
		Get-ADDirectReports -Identity Test_director -Recurse
	
Name                SamAccountName      Mail                Manager
----                --------------      ----                -------
test_managerB       test_managerB       test_managerB@la... test_director
test_userB1         test_userB1         test_userB1@lazy... test_managerB
test_userB2         test_userB2         test_userB2@lazy... test_managerB
test_managerA       test_managerA       test_managerA@la... test_director
test_userA2         test_userA2         test_userA2@lazy... test_managerA
test_userA1         test_userA1         test_userA1@lazy... test_managerA
	
	#>
	[CmdletBinding()]
	PARAM (
		[Parameter(Mandatory)]
		[String[]]$Identity,
		[Switch]$Recurse
	)
	BEGIN
	{
		TRY
		{
			IF (-not (Get-Module -Name ActiveDirectory)) { Import-Module -Name ActiveDirectory -ErrorAction 'Stop' -Verbose:$false }
		}
		CATCH
		{
			Write-Verbose -Message "[BEGIN] Something wrong happened"
			Write-Verbose -Message $Error[0].Exception.Message
		}
	}
	PROCESS
	{
		foreach ($Account in $Identity)
		{
			TRY
			{
				IF ($PSBoundParameters['Recurse'])
				{
					# Get the DirectReports
					Write-Verbose -Message "[PROCESS] Account: $Account (Recursive)"
					Get-Aduser -identity $Account -Properties directreports |
					ForEach-Object -Process {
						$_.directreports | ForEach-Object -Process {
							# Output the current object with the properties Name, SamAccountName, Mail and Manager
							Get-ADUser -Identity $PSItem  -Properties mail, manager | Select-Object -Property Name, SamAccountName, Mail, @{ Name = "Manager"; Expression = { (Get-Aduser -identity $psitem.manager ).samaccountname } }
							# Gather DirectReports under the current object and so on...
							Get-ADDirectReports -Identity $PSItem -Recurse
						}
					}
				}#IF($PSBoundParameters['Recurse'])
				IF (-not ($PSBoundParameters['Recurse']))
				{
					Write-Verbose -Message "[PROCESS] Account: $Account"
					# Get the DirectReports
					Get-Aduser -identity $Account  -Properties directreports | Select-Object -ExpandProperty directReports |
					Get-ADUser -Properties mail, manager  | Select-Object -Property Name, SamAccountName, Mail, @{ Name = "Manager"; Expression = { (Get-Aduser -identity $psitem.manager ).samaccountname } }
				}#IF (-not($PSBoundParameters['Recurse']))
			}#TRY
			CATCH
			{
				Write-Verbose -Message "[PROCESS] Something wrong happened"
				Write-Verbose -Message $Error[0].Exception.Message
			}
		}
	}
	END
	{
		Remove-Module -Name ActiveDirectory -ErrorAction 'SilentlyContinue' -Verbose:$false | Out-Null
    }
    
}

function Get-UserGroupMembershipRecursive {
[CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [String[]]$UserName,
        [String]$Domain
    )
    begin {
        # introduce two lookup hashtables. First will contain cached AD groups,
        # second will contain user groups. We will reuse it for each user.
        # format: Key = group distinguished name, Value = ADGroup object
        $ADGroupCache = @{}
        $UserGroups = @{}
        # define recursive function to recursively process groups.
        function __findPath ([string]$currentGroup) {
            Write-Verbose "Processing group: $currentGroup"
            # we must do processing only if the group is not already processed.
            # otherwise we will get an infinity loop
            if (!$UserGroups.ContainsKey($currentGroup)) {
                # retrieve group object, either, from cache (if is already cached)
                # or from Active Directory
                $groupObject = if ($ADGroupCache.ContainsKey($currentGroup)) {
                    Write-Verbose "Found group in cache: $currentGroup"
                    $ADGroupCache[$currentGroup]
                } else {
                    Write-Verbose "Group: $currentGroup is not presented in cache. Retrieve and cache."
                    $g = Get-ADGroup -Identity $currentGroup -Property "MemberOf"
                    # immediately add group to local cache:
                    $ADGroupCache.Add($g.DistinguishedName, $g)
                    $g
                }
                # add current group to user groups
                $UserGroups.Add($currentGroup, $groupObject)
                Write-Verbose "Member of: $currentGroup"
                foreach ($p in $groupObject.MemberOf) {
                    __findPath $p
                }
            } else {Write-Verbose "Closed walk or duplicate on '$currentGroup'. Skipping."}
        }
    }
    process {
        foreach ($user in $UserName) {
            Write-Verbose "========== $user =========="
            # clear group membership prior to each user processing
            $UserObject = Get-ADUser -Identity $user -Property "MemberOf"
            $UserObject.MemberOf | ForEach-Object {__findPath $_}
            New-Object psobject -Property @{
                UserName = $UserObject.Name;
                MemberOf = $UserGroups.Values | % {$_}; # groups are added in no particular order
            }
            $UserGroups.Clear()
        }
    }
}

function Get-ADNestedGroups {
[cmdletbinding()]
param (
[String] $GroupName
)            


$grat = New-Object System.Collections.ArrayList

$Members = Get-ADGroupMember -Identity $GroupName
$members | % {
    if($_.ObjectClass -eq "group") {
        $grot = $_ | select @{n='name';e={$_.name}}, @{n='SAM';e={$_.samaccountname}}
        $grat.Add($grot) > $null
    } else {
    }
}            
return $grat
}