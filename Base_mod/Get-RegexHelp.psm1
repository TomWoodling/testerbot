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
    $base = ":adbot: $env:BOTNAME attempts to identify commands using regex matches - try following these examples:
1. $env:BOTNAME get groups for USERNAME - gets all groups *USERNAME* is member of
2. $env:BOTNAME get users in GROUPNAME - gets all users in *GROUPNAME*
3. $env:BOTNAME get groups in GROUPNAME - gets only nested groups in *GROUPNAME*
4. $env:BOTNAME reports to USERNAME - gets any users reporting to *USERNAME* (if there are any)
5. $env:BOTNAME search AD t-serv - returns any group that contains the string 't-serv'
*Important note* quotes are no longer required for group names containing whitespace or special characters.
*Further Note* Any old commands you used with $env:BOTNAME will still work, but you should replace the $env:BOTNAME with $env:ALT

$env:HELPURL
"
    $usser = "User query examples:
1. $env:BOTNAME get groups for USERNAME - gets all groups *USERNAME* is member of.
Alternate command syntax:
 i) $env:BOTNAME membership USERNAME 
 ii) $env:ALT groups4user USERNAME
2. $env:BOTNAME reports to USERNAME - gets any users reporting to *USERNAME* (if there are any)
Alternate command syntax:
 i) $env:BOTNAME get direct reports for user USERNAME
 ii) $env:ALT reports-to USERNAME

"

    $grup = "Group query examples:
1. $env:BOTNAME get users in GROUPNAME - gets all members of *GROUPNAME*
Alternate command syntax:
 i) $env:BOTNAME get members of group GROUPNAME
 ii) $env:ALT adgroup GROUPNAME
2. $env:BOTNAME get groups in GROUPNAME - gets only nested groups in *GROUPNAME*
Alternate command syntax:
 i) $env:BOTNAME groups in GROUPNAME
 ii) $env:ALT groups-in GROUPNAME
 "

    if ($Arguments) {
        if ($Arguments[4] -match 'user' -and $Arguments[4] -notmatch 'group') {$hulp = $base+$usser}
        elseif ($Arguments[4] -notmatch 'user' -and $Arguments[4] -match 'group') {$hulp = $base+$grup}
        elseif ($Arguments[4] -match 'user' -and $Arguments[4] -match 'group') {$hulp = $base+$usser+$grup}
        else {$hulp = $base}
        }
    else {$hulp = $base}

    Return $hulp
}