function Get-FrontTeammates {
    [PoshBot.BotCommand(
        CommandName = 'Get-FrontTeammates',
        Aliases = ('frontteam','teammates'),
        Permissions = 'read'
    )]
    [cmdletbinding()]
    param(
        $Bot
    )

$frontkey = $env:FRONTAPI

$FrontTest = Invoke-WebRequest https://api2.frontapp.com/teammates -UseBasicParsing -Headers @{
    "Authorization"="Bearer $frontkey"
    "accept"="application/json"
    }

$path="$env:BOTROOT\csv\"

$FrontTest.Content | ConvertFrom-Json | Select-Object -ExpandProperty _results | Select-Object id,email,username,first_name,last_name | Export-Csv -NoTypeInformation $path\FrontUsers.csv -Force

New-PoshBotFileUpload -Path "$path\FrontUsers.csv" -Title Front_Teammates -DM

return "Message sent as DM :bowtie:"

}