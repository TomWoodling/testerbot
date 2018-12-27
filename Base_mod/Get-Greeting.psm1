function Get-Greeting {
    [PoshBot.BotCommand(
        Command = $false,
        CommandName = 'Get-Greeting',
        TriggerType = 'regex',
        #Regex replaces placeholder below during docker build
        Regex = "%"
    )]
    [cmdletbinding()]
    Param
    (
        $bot,
        [parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Arguments
    )

$greetz = @('hi there :bowtie:','good day :sun_with_face:','nice to see you :robot_face:','yo :thumbsup_all:')

$outp = $greetz | Get-Random

Write-Output $outp
}