function Get-ADGroupsForUserBot {
    <#
    .Synopsis
        Gets AD Groups for username.
    .DESCRIPTION
        Gets AD Groups for username.
    .EXAMPLE
        Get-ADGroupsForUserBot.ps1 -username t.woodling
    #>
    
    [CmdletBinding()]
    [PoshBot.BotCommand(
        CommandName = 'Get-ADGroupsForUserBot',
        Aliases = ('get-usergroups', 'groups4user', 'membership'),
        Permissions = 'read'
    )]
    Param
    (
        $bot,
        [Parameter(Mandatory=$true, Position=0)]
        [string]$User
    )
    
    #Get details for snippet
    $path="$env:botroot\csv\"
    $title="Results_for_$($User.Replace('.','_')).csv"
    
    # Create a hashtable for the results
    $result = @{}

    $go = Get-ADUser -Identity $user

    try {
        # Use ErrorAction Stop to make sure we can catch any errors
        $groups = Get-UserGroupMembershipRecursive -UserName "$User"
    
        if ($groups.memberof) {
        # Set a successful result
        $result.success = $true
        $groups.memberof | select name | Export-Csv -Path "$path\$title" -Force -NoTypeInformation
        $result.output = "I have sent the results as a DM :bowtie:"        
        New-PoshBotFileUpload -Path "$path\$title" -Title $title -DM
        #Remove-Item -Path "$path\$title" -Force
        }
        else {
            $result.success = $false
            $result.output = "No results for $user :crying_cat_face:"        }
        }
    catch {
    
        $clib = ':cold_sweat:'
        $result.output = "I cannot get details for $User $clib"
        
        # Set a failed result
        $result.success = $false
        }
    # Return the result and convert it to json, then attach a snippet with the results

    return $result.output
    }