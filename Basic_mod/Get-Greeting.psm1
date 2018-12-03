function Get-Greeting {
    [cmdletbinding()]
[PoshBot.BotCommand(
    CommandName = 'Get-Greeting',
    Aliases = ('hi', 'hello'),
    Permissions = 'read'
)]
param(
    $Bot = 'a'
)

$greetz = @('hi there :bowtie:','good day :sun_with_face:','nice to see you :robot_face:','yo :thumbsup_all:')

$outp = $greetz | Get-Random

Write-Output $outp
}