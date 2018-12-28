function Ignore-SelfSignedCerts {
   add-type -TypeDefinition  @"
       using System.Net;
       using System.Security.Cryptography.X509Certificates;
       public class TrustAllCertsPolicy : ICertificatePolicy {
           public bool CheckValidationResult(
               ServicePoint srvPoint, X509Certificate certificate,
               WebRequest request, int certificateProblem) {
               return true;
           }
       }
"@
   [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
}


function Get-PicOCR {
[cmdletbinding()]
param (
    [Parameter(Mandatory=$true)]    
    [string]$image    
    )

$headers = @{
    'Content-Type'='application/json'
    'Ocp-Apim-Subscription-Key'=$env:COMVIS_KEY
    }

$json_data = @{
    'url'="$image"
    } | ConvertTo-Json # Test connection

$postParams =  @{json_data=$json_data} # Test connection

$url = "https://westus.api.cognitive.microsoft.com/vision/v1.0/ocr?language=unk&detectOrientation =true"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ris = Invoke-RestMethod -Uri $url -Headers $headers -Body $json_data -Method Post 

$script:OCRres = $ris.regions.lines.words | Select-Object -ExpandProperty text

$script:OCRres

}


function Get-DescriptionOfPic {
[cmdletbinding()]
param (
    [Parameter(Mandatory=$true)]    
    [string]$image    
    )

$headers = @{
    'Content-Type'='application/json'
    'Ocp-Apim-Subscription-Key'=$env:COMVIS_KEY
    }

$json_data = @{
    'url'="$image"
    } | ConvertTo-Json # Test connection

$postParams =  @{json_data=$json_data} # Test connection

$url = "https://westus.api.cognitive.microsoft.com/vision/v1.0/analyze?visualFeatures=Description&language=en"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ris = Invoke-RestMethod -Uri $url -Headers $headers -Body $json_data -Method Post 

$script:opine = $ris.description.captions.text

return $script:opine

}


function Send-SlackBotFile {
    [CmdletBinding()]
    Param
    (
        # Name of the Service
        [Parameter()]
        [string]$Channels="#bot-conversation",
        $path
    )

    Send-SlackFile -Token $env:BOT_SLACK_TOKEN -Channel $Channels -Path $path
    }
    
function Get-TranslateToken {

$headers = @{
    'Ocp-Apim-Subscription-Key'=$env:TRANSLATOR_KEY
    }

$url = "https://api.cognitive.microsoft.com/sts/v1.0/issueToken"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ris = Invoke-RestMethod -Uri $url -Headers $headers -Method Post 

$script:opine = $ris

return $script:opine

}


function Get-LanguageOfPhrase {
[cmdletbinding()]
param (
    [Parameter(Mandatory=$true)]    
    [string]$phrase    
    )

$miik = Get-TranslateToken

$muuk = "Bearer" + " " + $miik

$headers = @{
    'authorization'= $muuk
    }

$url = "https://api.microsofttranslator.com/v2/http.svc/Detect?text=$phrase"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ris = Invoke-RestMethod -Uri $url -Headers $headers -Method Get 

$script:opine = $ris.string.'#text'

return $script:opine

}


function Get-Translation {
[cmdletbinding()]
param (
    [Parameter(Mandatory=$true)]    
    [string]$phrase,
    $tolang='en'     
    )

$orglang = Get-LanguageOfPhrase -phrase $phrase

$miik = Get-TranslateToken

$muuk = "Bearer" + " " + $miik

$headers = @{
    'authorization'=$muuk
    'from'=$orglang
    }

$url = "https://api.microsofttranslator.com/v2/http.svc/Translate?text=$phrase&to=$tolang"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ris = Invoke-RestMethod -Uri $url -Headers $headers -Method Get 

$script:opine = $ris.string.'#text'

return $script:opine

}

function new-botplugin {
[cmdletbinding()]
param (
    $plugname,
    $mod
    )

$mod = $mod+"_mod"
cd 'C:\Program Files\WindowsPowerShell\Modules\'
mkdir $plugname -Force
cd $plugname
$params = @{
    Path = ".\$plugname.psd1"
    RootModule = ".\$plugname.psm1"
    ModuleVersion = "0.1.0"
    Guid = New-Guid
    RequiredModules = "PoshBot"
    Author = "tom"
    Description = "$plugname plugin"
    }

Write-Host -ForegroundColor Green url will be "https://raw.githubusercontent.com/TomWoodling/testerbot/master/$mod/$plugname.psm1"

New-ModuleManifest @params
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
irm -uri "https://raw.githubusercontent.com/TomWoodling/testerbot/master/$mod/$plugname.psm1" -OutFile "$plugname.psm1"
$plugname >> "C:\Users\ContainerAdministrator\Documents\plugs.txt"
}

function Set-BotPlugs {
    [cmdletbinding()]
    param(
        $klips
        )
    
    $loc = '~\.poshbot\plugins.psd1'
    
    $woop = "@{"
    $woop > $loc
    foreach ($klip in $klips) {
    $smarp = "  '$klip' = @{
      '0.1.0' = @{
        Version = '0.1.0'
        Name = '$klip'
        AdhocPermissions = @()
        ManifestPath = 'C:\Program Files\WindowsPowerShell\Modules\$klip.psm1'
        CommandPermissions = @{
          '$klip' = @()
        }
        Enabled = `$True
      }
    }"
    $smarp >> $loc
      }
    $goop = "}"
    $goop >> $loc
    }

function New-BuiltInPlug {
    [cmdletbinding()]
    param (
        $plugname,
        $mod
        )
    $mod = $mod+'_mod'
    $woop = (gci "C:\Program Files\WindowsPowerShell\Modules\PoshBot").name    
    cd "C:\Program Files\WindowsPowerShell\Modules\PoshBot\$woop\Plugins\Builtin\Public\"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    irm -uri "https://raw.githubusercontent.com/TomWoodling/testerbot/master/$mod/$plugname.psm1" -OutFile "$plugname.ps1"
    $plugname >> "C:\Users\ContainerAdministrator\Documents\plugs.txt"
}

function Remove-Quotes {
    [CmdletBinding()]
    param(
    )

    dynamicparam {
        $ParamDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
 
        $Attributes = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $Attributes.Add( (New-Object System.Management.Automation.ParameterAttribute) )

        # Convert each object into a DynParamQuotedString:
        $Attributes.Add( (New-Object System.Management.Automation.ValidateSetAttribute($__DynamicParamValidateSet | % { [DynParamQuotedString] $_.ToString() })) )

        $ParamDictionary.$__DynamicParamName = New-Object System.Management.Automation.RuntimeDefinedParameter (
            $__DynamicParamName,
            [DynParamQuotedString],  # Notice the type here
            $Attributes
        )

        return $ParamDictionary
    } 

    process {
        
        $ParamValue = $null
        if ($PSBoundParameters.ContainsKey($__DynamicParamName)) {
            # Get the original string back:
            $ParamValue = $PSBoundParameters.$__DynamicParamName.OriginalString
        }

        "$ParamValue"
    }
}

function noquotez {
    [CmdletBinding()]
    param(
        [string]$bloop
    )

    $__DynamicParamName = "DynamicParam"
    $__DynamicParamValidateSet = @(
    "A string with spaces"
    "Another string with spaces"
    "StringWithoutSpaces"
    $bloop
)
    Remove-Quotes -DynamicParam $bloop
}

function New-RegexPlug {
    [cmdletbinding()]
    param (
        $plugname,
        $mod
        )
    $mod = $mod+'_mod'    
    $woop = (gci "C:\Program Files\WindowsPowerShell\Modules\PoshBot").name
    $boop = (irm -Uri "https://raw.githubusercontent.com/TomWoodling/testerbot/master/Regex/$plugname.regex").replace('$env:BOTNAME',"$env:BOTNAME")
    $gip = (irm -Uri "https://raw.githubusercontent.com/TomWoodling/testerbot/master/$mod/$plugname.psm1").replace('%',$boop) 
    $gip | Out-File "C:\Program Files\WindowsPowerShell\Modules\PoshBot\$woop\Plugins\Builtin\Public\$plugname.ps1" -Force
    $plugname >> "C:\Users\ContainerAdministrator\Documents\plugs.txt"
}

function New-BuiltinMod {
    [cmdletbinding()]
    param (
        $pluggers
        )

    $woop = (gci "C:\Program Files\WindowsPowerShell\Modules\PoshBot").name        
    $kip = (Invoke-RestMethod -uri "https://raw.githubusercontent.com/TomWoodling/testerbot/master/Builtin/Builtin_template").replace("%1%",",`'$($pluggers -join "','")`'") 
    $kip | Out-File "C:\Program Files\WindowsPowerShell\Modules\PoshBot\$woop\Plugins\Builtin\Builtin.psd1" -Force

}