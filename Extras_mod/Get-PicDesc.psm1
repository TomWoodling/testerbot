function Get-PicDesc {
    [CmdletBinding()]
    [PoshBot.BotCommand(
        CommandName = 'Get-PicDesc',
        Aliases = ('describe', 'picdesc'),
        Permissions = 'read'
    )]
    Param
    (
        $Bot,
        [Parameter(Position=0)]
	    [string]$pic
    )

    # Create a hashtable for the results
    $result = @{}
    $result.output = "Generic Error"    
    # Use try/catch block            
    try
    {
        # Use ErrorAction Stop to make sure we can catch any errors
        $viewp = Get-DescriptionOfPic -image $pic
        
        # Create a string for sending back to slack. * and ` are used to make the output look nice in Slack. Details: http://bit.ly/MHSlackFormat
        $result.output = "I think it is ``$viewp``"
        
        # Set a successful result
        $result.success = $true
    }
    catch
    {
        # If this script fails we can assume the service did not exist
        $result.output = "It didn't work :crying_cat_face: $($_.Exception.message)"
        
        # Set a failed result
        $result.success = $false
    }
    
    # Return the result and conver it to json
    return $result.output
}