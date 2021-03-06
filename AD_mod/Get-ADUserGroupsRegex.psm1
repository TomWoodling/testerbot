function Get-ADUserGroupsRegex {
    <#
    .Synopsis
        Gets AD Groups for username.
    .DESCRIPTION
        Gets AD Groups for username.
    .EXAMPLE
        Get-ADGroupsForUserBot.ps1 -username t.woodling
    #>
    
    [PoshBot.BotCommand(
        Command = $false,
        CommandName = 'Get-ADUserGroupsRegex',
        TriggerType = 'regex',
        Regex = "%"
    )]
    [cmdletbinding()]
    param(
        $Bot,
        [parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Arguments
    )
    
    $User = $Arguments[5]
    #Get details for snippet
    $path="$env:botroot\csv\"
    $title="Results_for_$($User.Replace('.','_')).csv"
    
    # Create a hashtable for the results
    $result = @{}
    

    try {
        $go = Get-ADUser -Filter "samaccountname -like '$User'"
        if ($go) {
            $groups = Get-UserGroupMembershipRecursive -UserName "$User"
        
            if ($groups.memberof) {
            # Set a successful result
            $result.success = $true
            $groups.memberof | select name | Export-Csv -Path "$path\$title" -Force -NoTypeInformation
            $result.output = "I have sent the results for $user as a DM :bowtie:"        
            New-PoshBotFileUpload -Path "$path\$title" -Title $title -DM
            #Remove-Item -Path "$path\$title" -Force
            }
            else {
                $result.success = $false
                $result.output = "No results for $user :crying_cat_face: - they may not have been added to any AD groups yet"
                }
            }
        else {
            $result.success = $false
            $result.output = "$user not found - you could try the search AD command for a partial string match :crying_cat_face:"            
            }
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