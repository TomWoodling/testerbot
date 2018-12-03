function Get-ADDirectRepsBot {
    <#
    .Synopsis
        Gets direct reports in AD.
    .DESCRIPTION
        Gets direct reports in AD
    .EXAMPLE
        Get-ADDirectRepsBot -User a.testname
    #>
    
    [CmdletBinding()]
    [PoshBot.BotCommand(
        CommandName = 'Get-ADDirectRepsBot',
        Aliases = ('get-directreps', 'direct-reps', 'reports-to'),
        Permissions = 'read'
    )]
    Param
    (
        # Name of the Service
        $bot,
        [Parameter(Mandatory=$true, Position=0)]
        [string]$User
    )
    
    #Get details for snippet
    $path="$env:botroot\csv\"
    $title="Results_for_$($User.Replace('.','_')).csv"
    
    $go = Get-ADUser -Identity $user
    # Create a hashtable for the results
    $result = @{}
    
    try {
        # Use ErrorAction Stop to make sure we can catch any errors
        $reps = Get-ADDirectReports -Identity $user -Recurse
        if ($reps) {$reps | Export-Csv -Path "$path\$title" -Force -NoTypeInformation
            New-PoshBotFileUpload -Path "$path\$title" -Title $title -DM
            $result.output = "I have sent the results as a DM :bowtie:"
            }
        else {$result.output = "No results found :bowtie:"}
        # Set a successful result
        $result.success = $true
    
        
        }
    catch {
        # If this script fails we can try to match the name instead to see if we get any suggestions
        $result.output = "$User does not exist :cold_sweat:"
        
        # Set a failed result
        $result.success = $false
        }
    # Return the result and convert it to json, then attach a snippet with the results

    return $result.output
    }