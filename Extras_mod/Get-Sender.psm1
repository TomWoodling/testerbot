function Get-Sender {

    [PoshBot.BotCommand(
        CommandName = 'Get-Sender',
        Aliases = ('myname'),
        Permissions = 'read'
    )]
    [cmdletbinding()]
    param(
        $Bot
    )

    return $PoshBotContext.FromName
}