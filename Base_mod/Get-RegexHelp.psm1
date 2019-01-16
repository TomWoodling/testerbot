function Get-RegexHelp {
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
        CommandName = 'Get-RegexHelp',
        TriggerType = 'regex',
        Regex = "%"
    )]
    [cmdletbinding()]
    param(
        $Bot,
        [parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Arguments
    )
    $base = ":adbot: ADbot attempts to identify commands using regex matches - try following these examples:
1. adbot get groups for t.woodling - gets all groups t.woodling is member of
2. adbot get users in gl-it-services - gets all members of gl-it-services
3. adbot get groups in gl-it-services - gets only nested groups in gl-it-services
4. adbot reports to t.woodling - gets any users reporting to t.woodling (if there are any)
5. adbot search AD t-serv - returns any group that contains the string 't-serv'
*Important note* quotes are no longer required for group names containing whitespace or special characters!
*Further Note* Any old commands you used with adbot will still work, but you should replace the adbot with !

"
    $usser = "User query examples:
1. adbot get groups for t.woodling - gets all groups t.woodling is member of.
Alternate command syntax:
 i) adbot membership t.woodling 
 ii) ! groups4user t.woodling
2. adbot reports to t.woodling - gets any users reporting to t.woodling (if there are any)
Altername command syntax:
 i) adbot get direct reports for user t.woodling
 ii) ! reports-to t.woodling

"

    $grup = "Group query examples:
1. adbot get users in gl-it-services - gets all members of gl-it-services
Alternate command syntax:
 i) adbot get members of group gl-it-services'
 ii) ! adgroup gl-it-services
2. adbot get groups in gl-it-services - gets only nested groups in gl-it-services
Altername command syntax:
 i) adbot groups in gl-it-services
 ii) ! groups-in gl-it-services"

    if ($Arguments) {
        if ($Arguments[5] -match 'user' -and $Arguments[5] -notmatch 'group') {$hulp = $base+$usser}
        elseif ($Arguments[5] -notmatch 'user' -and $Arguments[5] -match 'group') {$hulp = $base+$grup}
        elseif ($Arguments[5] -match 'user' -and $Arguments[5] -match 'group') {$hulp = $base+$usser+$grup}
        else {$hulp = $base}
        }
    else {$hulp = $base}

    Return $hulp
}