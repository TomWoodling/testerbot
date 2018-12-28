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

    $first = $Arguments[1]
    $second = $Arguments[2]

    Return "I get $first for the first part and $second for the second - from $arguments"
}