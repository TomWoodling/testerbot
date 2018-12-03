function Get-ADGrpMemBot {
    <#
    .Synopsis
        Ad group members query for ad_bot
    .DESCRIPTION
        Ad group members query
    .EXAMPLE
        Get-ADGroupMemberBot.ps1 -group 'SG-ITS-Ops'
    #>
    
    [CmdletBinding()]
    [PoshBot.BotCommand(
        CommandName = 'Get-ADGrpMemBot',
        Aliases = ('get-members', 'adgrpmems','adgroup'),
        Permissions = 'read'
    )]
    Param
    (
        $bot,
        # Name of the Service
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Group
    )
Add-Type @"
    public class DynParamQuotedString {

        public DynParamQuotedString(string quotedString) : this(quotedString, "'") {}
        public DynParamQuotedString(string quotedString, string quoteCharacter) {
            OriginalString = quotedString;
            _quoteCharacter = quoteCharacter;
        }

        public string OriginalString { get; set; }
        string _quoteCharacter;

        public override string ToString() {
            if (OriginalString.Contains(" ")) {
                return string.Format("{1}{0}{1}", OriginalString, _quoteCharacter);
            }
            else {
                return OriginalString;
            }
        }
    }
"@


    #Get details for snippet
    $path="$env:BOTROOT\csv\"
    $mitle = $Group.Replace(' ','_')
    $title = "$($mitle.replace('&amp;','-')).ps1"

    # Create a hashtable for the results
    $result = @{}
    
    $birp = noquotez -bloop $group

    $gwipe = $($birp.replace('&amp;','&'))

    $go = Get-ADGroup -Identity $gwipe
    

    try {
        $gwurp = "Get-ADGroupMember -Identity `"$gwipe`" -Recursive | select name,samaccountname"
        $gwoops = Invoke-Expression -Command $gwurp
        $outle = "$($mitle.replace('&amp;','-')).csv"
        $gwoops | Export-Csv -Path "$path\$outle" -Force -NoTypeInformation
        New-PoshBotFileUpload -Path "$path\$outle" -Title $outle -DM
        $result.output = "Request for $gwipe processed - results sent as a DM :bowtie:"
        # Set a successful result
        $result.success = $true
        }
    catch {
        $result.output = "Group $gwipe does not exist :cold_sweat:"        
        # Set a failed result
        $result.success = $false
        }
    return $result.output
    Remove-Item -Force -Path "$path$title"
}