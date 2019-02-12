More testing for bots - based on the poshbot framework at https://github.com/poshbotio/PoshBot

The aim here is to produce a bot that can be called using docker, use regex to invoke commands, automatically build and add required plugins

To begin, the modules here can be used with a dockerfile to add extra built in commands to poshbot so they are immediately available.

**NB:** Sample dockerfile (for Windows Server 2016 - change the first line for other server versions).

**NB(again):** To run the AD commands you'll need to enable the container with AD support - see https://blogs.msdn.microsoft.com/containerstuff/2017/01/30/create-a-container-with-active-directory-support/

```
FROM microsoft/dotnet-framework:4.7.2-runtime-windowsservercore-ltsc2016

# $ProgressPreference: https://github.com/PowerShell/PowerShell/issues/2138#issuecomment-251261324
SHELL ["powershell", "-Command", "$ErrorActionPreference = 'Stop'; $ProgressPreference = 'SilentlyContinue';"]
ENV BOT_SLACK_TOKEN [SLACK_TOKEN]
ENV BOTADMINS [name]
ENV COMVIS_KEY [key]
ENV LUIS_SUB [key]
ENV LUIS_APP [key]
ENV TRANSLATOR_KEY [key]
ENV THE_CAT_API_KEY [key]
ENV BOTNAME testing
ENV BOTROOT C:\\testing
ENV ALT !
RUN New-Item -ItemType Directory -Path \"C:\Program Files\WindowsPowerShell\Modules\testbothelpers\"; \
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; \
    Invoke-RestMethod -uri \"https://raw.githubusercontent.com/TomWoodling/testerbot/master/Helpers/testbothelpers.psm1\" -OutFile \"C:\Program Files\WindowsPowerShell\Modules\testbothelpers\testbothelpers.psm1\"; \
    New-Item -ItemType Directory -Path \"C:\Program Files\WindowsPowerShell\Modules\adfunctions\"; \
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; \
    Invoke-RestMethod -uri \"https://raw.githubusercontent.com/TomWoodling/testerbot/master/Helpers/adfunctions.psm1\" -OutFile \"C:\Program Files\WindowsPowerShell\Modules\adfunctions\adfunctions.psm1\";
RUN Install-PackageProvider Nuget -Force; \
    Install-Module -Name PowerShellGet -Force; \
    Update-Module -Name PowerShellGet; \
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted; \
    New-Item -ItemType Directory -Path \"$ENV:BOTROOT\"; \
    New-Item -ItemType Directory -Path \"$ENV:BOTROOT\csv\";
RUN cd \"$ENV:BOTROOT\"; \
    Install-Module -Name PoshBot -Repository PSGallery; \
    Import-Module -Name PoshBot -ErrorAction SilentlyContinue -Force;
RUN Install-WindowsFeature RSAT-AD-PowerShell; \
    Import-Module -Name ActiveDirectory -Force; \
    Install-Module -Name PSSlack; \
    Import-Module -Name PSSlack -Force; \
    $TextInfo = (Get-Culture).TextInfo; \
    $admins = @($env:BOTADMINS); \
    $botParams = @{}; \
    $botParams.Add(\"Name\",\"$env:BOTNAME\"); \
    $botParams.Add(\"CommandPrefix\",\"$env:ALT\"); \
    $botParams.Add(\"BotAdmins\",$admins); \
    $botParams.Add(\"LogLevel\",\"Info\"); \
    $back = @{}; \
    $back.Add(\"Name\",\"SlackBackend\"); \
    $back.Add(\"Token\",\"$env:BOT_SLACK_TOKEN\"); \
    $botparams.Add(\"back\",$back); \
    $myBotConfig = New-PoshBotConfiguration @botParams; \
    $confp = \"C:\$env:BOTNAME\poshbotconfig.psd1\"; \
    Save-PoshBotConfiguration -InputObject $myBotConfig -Path $confp -Force;
RUN New-Item \"C:\Users\ContainerAdministrator\Documents\WindowsPowerShell\Modules\" -ItemType Directory; \
    Import-Module -Name testbothelpers -Force; \
    Import-Module -Name adfunctions -Force; \
    $woop = (gci \"C:\Program Files\WindowsPowerShell\Modules\PoshBot\").name; \
    New-BuiltInPlug -plugname \"Get-CatPic\" -mod \"Extras\"; \
    New-BuiltInPlug -plugname \"Get-PicDesc\" -mod \"Extras\"; \
    New-BuiltInPlug -plugname \"Get-CatBomb\" -mod \"Extras\"; \
    New-BuiltInPlug -plugname \"Get-ADDirectRepsBot\" -mod \"AD\"; \
    New-BuiltInPlug -plugname \"Get-ADGroupsForUserBot\" -mod \"AD\"; \
    New-BuiltInPlug -plugname \"Get-ADGrpMemBot\" -mod \"AD\"; \
    New-BuiltInPlug -plugname \"Get-ADNestedGroupsBot\" -mod \"AD\"; \
    New-BuiltInPlug -plugname \"Search-ADGroups\" -mod \"AD\"; \
    New-RegexPlug -plugname \"Set-Gratitude\" -mod \"Base\"; \
    New-RegexPlug -plugname \"Get-Greeting\" -mod \"Base\"; \
    New-RegexPlug -plugname \"Test-Regex\" -mod \"Base\"; \
    New-RegexPlug -plugname \"Get-CatRegex\" -mod \"Base\"; \
    New-RegexPlug -plugname \"Get-RegexHelp\" -mod \"Base\"; \
    New-RegexPlug -plugname \"Get-ADDirectRepsRegex\" -mod \"AD\"; \
    New-RegexPlug -plugname \"Get-ADGrpMemRegex\" -mod \"AD\"; \
    New-RegexPlug -plugname \"Get-ADNestedGroupsRegex\" -mod \"AD\"; \
    New-RegexPlug -plugname \"Get-ADUserGroupsRegex\" -mod \"AD\"; \
    New-RegexPlug -plugname \"Search-ADGroupsRegex\" -mod \"AD\"; \
    $grubs = Get-Content \"C:\Users\ContainerAdministrator\Documents\plugs.txt\"; \
    New-BuiltinMod -pluggers $grubs;
CMD Start-Job -Name PoshBot_RY -ScriptBlock {Import-Module -Name PoshBot; $pbc = Get-PoshBotConfiguration -Path "$ENV:BOTROOT\poshbotconfig.psd1"; Start-PoshBot -Configuration $pbc}; $o = 10 ; while ($o -eq 10) {Start-Sleep 300};
```