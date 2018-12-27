function Get-CatPic {
    [CmdletBinding()]
    [PoshBot.BotCommand(
        CommandName = 'Get-CatPic',
        Aliases = ('cat-me','cats'),
        Permissions = 'read'
    )]
    Param
    (
        $Bot = 'a',
        [Parameter(Position=0)]
	[Int]$limit = 1
    )

    # Some variables to use
    $apikey = $env:THE_CAT_API_KEY
    $baseurl = 'https://api.thecatapi.com/v1'
    if ($limit -gt 5) {$limit = 5}
    $rest = "/images/search?size=small&mime_types=jpg,png,gif&format=json&has_breeds=false&order=RANDOM&page=0&limit=$limit&api_key=$apikey"
    # Create a hashtable for the results
    $result = @{}
    #Headers for api call
    $headers = @{
        'Content-Type'="application/json"
        'x-api-key'="$apikey"
        }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Use try/catch block            
    try
    {
        # Use ErrorAction Stop to make sure we can catch any errors
        $res = irm -uri "$baseurl$rest" -Headers $headers -Method Get -ErrorAction Stop
        
        # Create a string for sending back to slack. * and ` are used to make the output look nice in Slack. Details: http://bit.ly/MHSlackFormat
        $result.output = "$($res.url)"
        
        # Set a successful result
        $result.success = $true
    }
    catch
    {
        # If this script fails we can assume the service did not exist
        $result.output = "It didn't work :crying_cat_face:"
        
        # Set a failed result
        $result.success = $false
    }
    
    # Return the result and conver it to json
    return $result.output
}