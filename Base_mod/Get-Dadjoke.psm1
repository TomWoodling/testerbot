function Get-Dadjoke {
    <#
    .SYNOPSIS
        Displays a Grafana graph for a given type for a server.
    .EXAMPLE
        grafana cpu server01
    .EXAMPLE
        grafana disk myotherserver02
    #>
    [PoshBot.BotCommand(
        Command = $false,
        CommandName = 'Get-Dadjoke',
        TriggerType = 'regex',
        Regex = "%"
    )]
    [cmdletbinding()]
    param(
        $Bot,
        [parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Arguments
    )

    if ($Arguments) {
        $headers = @{
        'Accept'="application/json"
        }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $term = $arguments[4]
    $lulz = ((irm -uri "https://icanhazdadjoke.com/search?term=$term&limit=10" -Method Get -Headers $headers | select -ExpandProperty results) | Get-Random) | select -ExpandProperty joke
        }
    else {$lulz = (irm -uri https://icanhazdadjoke.com/slack -Method Get).attachments.text}

    Return $lulz
}