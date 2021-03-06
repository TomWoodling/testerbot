function Get-ADDirectRepsRegex {
    <#
    .Synopsis
        Gets direct reports in AD.
    .DESCRIPTION
        Gets direct reports in AD
    .EXAMPLE
        Get-ADDirectRepsBot -User a.testname
    #>
    
    [PoshBot.BotCommand(
        Command = $false,
        CommandName = 'Get-ADDirectRepsRegex',
        TriggerType = 'regex',
        Regex = "%"
    )]
    [cmdletbinding()]
    param(
        $Bot,
        [parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Arguments
    )
    
    $user = $Arguments[6]
    #Get details for snippet
    $path="$env:botroot\csv\"
    $title="Results_for_$($User.Replace('.','_')).csv"
    
    $go = Get-ADUser -Filter "samaccountname -like '$User'"
    # Create a hashtable for the results
    $result = @{}
    if($go) {
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
        }
    else {$result.output = "$user not found - you could try the search AD command for a partial string match :crying_cat_face:"}
    # Return the result and convert it to json, then attach a snippet with the results

    return $result.output
    }