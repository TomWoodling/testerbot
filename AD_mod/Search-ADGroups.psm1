function Search-ADGroups {
    <#
    .Synopsis
        Search for AD to match partial string
    .DESCRIPTION
        Ad group members query
    .EXAMPLE
        Get-ADGroupMemberBot.ps1 -group 'SG-ITS-Ops'
    #>
    
    [CmdletBinding()]
    [PoshBot.BotCommand(
        CommandName = 'Search-ADGroups',
        Aliases = ('search-ad','find'),
        Permissions = 'read'
    )]
    Param
    (
        $bot,
        # Name of the Service
        [Parameter(Mandatory=$true, Position=0)]
        [string]$object
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
    $mitle = $object.Replace(' ','_')
    $title = "$($mitle.replace('&amp;','-')).ps1"

    # Create a hashtable for the results
    $result = @{}
    
    $birp = noquotez -bloop $object

    $gwipe = $($birp.replace('&amp;','&'))

    

    try {
        $gwurp = "Get-ADGroup -Filter {name -like `'*$gwipe*`'} -Properties description | select name,description"
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
    }