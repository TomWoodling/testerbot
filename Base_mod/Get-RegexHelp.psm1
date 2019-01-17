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
    $base = ":$env:BOTNAME: $env:BOTNAME attempts to identify commands using regex matches - try following these examples:
1. $env:BOTNAME get groups for a.person - gets all groups a.person is member of
2. $env:BOTNAME get users in group of users & things - gets all users in group of users & things
3. $env:BOTNAME get groups in group of users & things - gets only nested groups in group of users & things
4. $env:BOTNAME reports to a.person - gets any users reporting to a.person (if there are any)
5. $env:BOTNAME search AD t-serv - returns any group that contains the string 't-serv'
*Important note* quotes are no longer required for group names containing whitespace or special characters.
*Further Note* Any old commands you used with $env:BOTNAME will still work, but you should replace the $env:BOTNAME with $env:ALT

"
    $usser = "User query examples:
1. $env:BOTNAME get groups for a.person - gets all groups a.person is member of.
Alternate command syntax:
 i) $env:BOTNAME membership a.person 
 ii) $env:ALT groups4user a.person
2. $env:BOTNAME reports to a.person - gets any users reporting to a.person (if there are any)
Alternate command syntax:
 i) $env:BOTNAME get direct reports for user a.person
 ii) $env:ALT reports-to a.person

"

    $grup = "Group query examples:
1. $env:BOTNAME get users in group of users & things - gets all members of group of users & things
Alternate command syntax:
 i) $env:BOTNAME get members of group group of users & things
 ii) $env:ALT adgroup group of users & things
2. $env:BOTNAME get groups in group of users & things - gets only nested groups in group of users & things
Alternate command syntax:
 i) $env:BOTNAME groups in group of users & things
 ii) $env:ALT groups-in group of users & things
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