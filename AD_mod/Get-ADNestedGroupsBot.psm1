function Get-ADNestedGroupsBot {
    <#
.Synopsis
    Gets service status for Hubot Script.
.DESCRIPTION
    Gets service status for Hubot Script.
.EXAMPLE
    Get-ServiceHubot -Name dhcp
#>

[CmdletBinding()]
[PoshBot.BotCommand(
    CommandName = 'Get-ADNestedGroupsBot',
    Aliases = ('get-nestedgroups', 'groups-in'),
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
    # Use ErrorAction Stop to make sure we can catch any errors
    $gurps = "Get-ADNestedGroups -GroupName `'$gwipe`' -ErrorAction stop | select name"
    $gwoops = Invoke-Expression -Command $gurps
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
# Return the result and convert it to json, then attach a snippet with the results
    return $result.output
    Remove-Item -Force -Path "$path$title"
}