function Test-Regex {
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
        CommandName = 'Test-Regex',
        TriggerType = 'regex',
        Regex = "%"
    )]
    [cmdletbinding()]
    param(
        $Bot,
        [parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Arguments
    )

    $urbs = new-object system.collections.arraylist
    $i = 0
    foreach ($arg in $arguments) {
        $urg = "$i is $arg"
        $urbs.add($urg) > $null
        $i = $i+1
    }

    Return $urbs
}