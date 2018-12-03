function Get-CatBomb {
    [CmdletBinding()]
    [PoshBot.BotCommand(
        CommandName = 'Get-CatBomb',
        Aliases = ('cat-bomb'),
        Permissions = 'read'
    )]
    Param
    (
        $Bot,
        [Parameter(Position=0)]
        [ValidateSet('boxes','caturday','clothes','dream','funny','hats','kittens','sinks','space','sunglasses','ties','all')]
	    [string]$category = 'all'
    )

    # Some variables to use
    $apikey = $env:THE_CAT_API_KEY
    $baseurl = 'https://api.thecatapi.com/v1'
    $limit = 5
    if ($category -eq 'all') {
        $skot=('5','6','15','9','3','1','10','14','2','4','7')
        $skit = $skot | Get-Random
        }
    else {
        switch($category) {
        'boxes' { $skit = '5' }
        'caturday' { $skit = '6' }
        'clothes' { $skit = '15' }
        'dream' { $skit = '9' }
        'funny' { $skit = '3' }
        'hats' { $skit = '1' }
        'kittens' { $skit = '10' }
        'sinks' { $skit = '14' }
        'space' { $skit = '2' }
        'sunglasses' { $skit = '4' }
        'ties' { $skit = '7' }
        }
    }
    $rest = "/images/search?size=small&mime_types=jpg,png,gif&format=json&has_breeds=false&order=RANDOM&page=0&limit=$limit&category_ids=$skit&api_key=$apikey"
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