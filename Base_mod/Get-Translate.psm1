function Get-Translate {
<#
.Synopsis
    Translates a phrase to english
.DESCRIPTION
    Makes connection to Microsoft Cognitive Services Translator API
.EXAMPLE
    .\Get-Translation.ps1 -phrase 'Hola soy un gato'
#>

    [PoshBot.BotCommand(
        Command = $false,
        CommandName = 'Get-Translate',
        TriggerType = 'regex',
        Regex = "%"
    )]
    [cmdletbinding()]
    param(
        $Bot,
        [parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Arguments
    )
    # Create a hashtable for the results
    $result = @{}
    $phrase = $Arguments[2]
    # Use try/catch block            
    try
    {
        # Use ErrorAction Stop to make sure we can catch any errors
        $translate = Get-Translation -phrase "$phrase" -ErrorAction Stop
        
        # Create a string for sending back to slack. * and ` are used to make the output look nice in Slack. Details: http://bit.ly/MHSlackFormat
        $result.output = "I think it is ``$translate``"
        
        # Set a successful result
        $result.success = $true
    }
    catch
    {
        # If this script fails we can assume the service did not exist
        $result.output = "No match for ``$phrase``"
        
        # Set a failed result
        $result.success = $false
    }
    
    # Return the result and conver it to json
    return $result | ConvertTo-Json
    
    }