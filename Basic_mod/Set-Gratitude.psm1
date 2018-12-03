function Set-Gratitude {
    <#
    .Synopsis
        Bot graciously accepts your thanks
    .DESCRIPTION
        Bot graciously accepts your thanks - must be used with command to replace placeholder with correct regex (including bot name) in dockerfile
    .EXAMPLE
        Set-Gratitude
    #>
    
    [PoshBot.BotCommand(
        Command = $false,
        CommandName = 'Set-Gratitude',
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


    $phrases = (
        "You're welcome",
        "No problem",
        "No prob",
        "np",
        "Sure thing",
        "Anytime, human",
        "Anytime",
        "Anything for you",
        "De nada, amigo",
        "Don't worry about it",
        "My pleasure"
        )

    $punc = (
        " ",
        "!",
        ".",
        "!!",
        " - ",
        ", "
        )

    $emoji = (" ", " ", ":muscle:", ":smile:", ":+1:", ":ok_hand:", ":punch:",
    ":bowtie:", ":smiley:", ":joy_cat:", ":heart:", ":robot_face:",
    ":heartbeat:", ":sparkles:", ":star:", ":star2:", ":smirk:",
    ":grinning:", ":smiley_cat:", ":sunflower:", ":tulip:",
    ":hibiscus:", ":cherry_blossom:", ":ghost:", ":eyes:")

    $resp = "$($phrases | Get-Random)$($punc | Get-Random) $($emoji | Get-Random)"

    return $resp
}

