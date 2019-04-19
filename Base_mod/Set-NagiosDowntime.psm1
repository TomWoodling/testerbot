function Set-NagiosDowntime {
    [PoshBot.BotCommand(
        Command = $false,
        CommandName = 'Set-NagiosDowntime',
        TriggerType = 'regex',
        Regex = "%"
    )]
    [cmdletbinding()]
    param(
        $Bot,
        [parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Arguments
    )

        $target=$Arguments[2]
        $StartDate = (Get-Date)
        $duration=$Arguments[3]
        $nagkey = $env:NAGIOSKEY
     
    # Convert the dates
    $starter = ([DateTimeOffset](Get-Date)).ToUnixTimeSeconds()
    $endo = ([DateTimeOffset]((Get-Date).AddMinutes($duration))).ToUnixTimeSeconds()
    #Validate the target hostname
    $womps = (irm -uri "https://nagios.cool.blue/nagiosxi/api/v1/objects/host?apikey=$nagkey&pretty=1").host | select -ExpandProperty host_name
    if ($womps -contains $target) {
        $comment = "auto_downtime"
        $url = "https://nagios.cool.blue/nagiosxi/api/v1/system/scheduleddowntime?apikey=$nagkey&pretty=1&hosts[]=$target&start=$starter&end=$endo&comment=$comment&all_services=1"

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $res = irm -Uri $url  -Method Post
        return "$(($res | select -ExpandProperty scheduled).hosts) booked downtime"
        }
    else {
        $wimps = New-Object System.Collections.ArrayList
        foreach ($womp in $womps) {
            if ($target -match $womp) {
                $wimps.Add($womp) > $null
                }
            }
        if ($wimps) {
            return "I could not find the host :crying_bear: Maybe you could try one of these: $($wimps -join ', ')"
            }
        else {
            return "I could not find the host :crying_bear:"
            }
        }
    }