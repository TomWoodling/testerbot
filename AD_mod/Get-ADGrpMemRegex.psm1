function Get-ADGrpMemRegex {
    <#
    .Synopsis
        Ad group members query for ad_bot
    .DESCRIPTION
        Ad group members query
    .EXAMPLE
        Get-ADGroupMemberBot.ps1 -group 'SG-ITS-Ops'
    #>
    
    [PoshBot.BotCommand(
        Command = $false,
        CommandName = 'Get-ADGrpMemRegex',
        TriggerType = 'regex',
        Regex = "%"
    )]
    [cmdletbinding()]
    param(
        $Bot,
        [parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Arguments
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

    $group = $Arguments[7]
    #Get details for snippet
    $path="$env:BOTROOT\csv\"
    $mitle = $Group.Replace(' ','_')
    $title = "$($mitle.replace('&amp;','-')).ps1"

    # Create a hashtable for the results
    $result = @{}
    
    $gwipe = $group.Replace("'","").Replace('"','')

    $go = Get-ADGroup -filter "samaccountname -like '$gwipe'"
    
    if ($go) {
        try {
            #$gwurp = "Get-ADGroupMember -Identity $gwipe -Recursive | select name,samaccountname"
            $gwoops = Get-ADGroupMember -Identity $gwipe -Recursive | select name,samaccountname | Sort-Object name
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
        }
    else {$result.output = "Group $gwipe does not exist :cold_sweat: - you can try search ad command for a partial string match"}
    return $result.output
    Remove-Item -Force -Path "$path$title"
}