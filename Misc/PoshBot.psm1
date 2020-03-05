
# Track bot instnace(s) running as PS job
$script:botTracker = @{}

$script:pathSeperator = [IO.Path]::PathSeparator

$script:moduleBase = $PSScriptRoot

if (($null -eq $IsWindows) -or $IsWindows) {
    $homeDir = $env:USERPROFILE
} else {
    $homeDir = $env:HOME
}
$script:defaultPoshBotDir = (Join-Path -Path $homeDir -ChildPath '.poshbot')

$PSDefaultParameterValues = @{
    'ConvertTo-Json:Verbose' = $false
}

# Enforce TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Some enums
enum AccessRight {
    Allow
    Deny
}

enum ConnectionStatus {
    Connected
    Disconnected
}

enum TriggerType {
    Command
    Event
    Regex
}

enum Severity {
    Success
    Warning
    Error
    None
}

enum LogLevel {
    Info = 1
    Verbose = 2
    Debug = 4
}

enum LogSeverity {
    Normal
    Warning
    Error
}

enum ReactionType {
    Success
    Failure
    Processing
    Custom
    Warning
    ApprovalNeeded
    Cancelled
    Denied
}

# Unit of time for scheduled commands
enum TimeInterval {
    Days
    Hours
    Minutes
    Seconds
}

enum ApprovalState {
    AutoApproved
    Pending
    Approved
    Denied
}

enum MessageType {
    CardClicked
    ChannelRenamed
    Message
    PinAdded
    PinRemoved
    PresenceChange
    ReactionAdded
    ReactionRemoved
    StarAdded
    StarRemoved
}

enum MessageSubtype {
    None
    ChannelJoined
    ChannelLeft
    ChannelRenamed
    ChannelPurposeChanged
    ChannelTopicChanged
}

enum MiddlewareType {
    PreReceive
    PostReceive
    PreExecute
    PostExecute
    PreResponse
    PostResponse
}

class LogMessage {
    [datetime]$DateTime = (Get-Date).ToUniversalTime()
    [string]$Class
    [string]$Method
    [LogSeverity]$Severity = [LogSeverity]::Normal
    [LogLevel]$LogLevel = [LogLevel]::Info
    [string]$Message
    [object]$Data

    LogMessage() {
    }

    LogMessage([string]$Message) {
        $this.Message = $Message
    }

    LogMessage([string]$Message, [object]$Data) {
        $this.Message = $Message
        $this.Data = $Data
    }

    LogMessage([LogSeverity]$Severity, [string]$Message) {
        $this.Severity = $Severity
        $this.Message = $Message
    }

    LogMessage([LogSeverity]$Severity, [string]$Message, [object]$Data) {
        $this.Severity = $Severity
        $this.Message = $Message
        $this.Data = $Data
    }

    # Borrowed from https://github.com/PowerShell/PowerShell/issues/2736
    hidden [string]Compact([string]$Json) {
        $indent = 0
        $compacted = ($Json -Split '\n' | ForEach-Object {
            if ($_ -match '[\}\]]') {
                # This line contains  ] or }, decrement the indentation level
                $indent--
            }
            $line = (' ' * $indent * 2) + $_.TrimStart().Replace(':  ', ': ')
            if ($_ -match '[\{\[]') {
                # This line contains [ or {, increment the indentation level
                $indent++
            }
            $line
        }) -Join "`n"
        return $compacted
    }

    [string]ToJson() {
        $json = [ordered]@{
            DataTime = $this.DateTime.ToString('u')
            Class = $this.Class
            Method = $this.Method
            Severity = $this.Severity.ToString()
            LogLevel = $this.LogLevel.ToString()
            Message = $this.Message
            Data = foreach ($item in $this.Data) {
                # Summarize exceptions so they can be serialized to json correctly

                # Don't try to serialize jobs
                if ($item.GetType().BaseType.ToString() -eq 'System.Management.Automation.Job') {
                    continue
                }

                # Summarize Error records so the json is easier to read and doesn't
                # contain a ton of unnecessary infomation
                if ($item -is [System.Management.Automation.ErrorRecord]) {
                    [ExceptionFormatter]::Summarize($item)
                } else {
                    $item
                }
            }
        } | ConvertTo-Json -Depth 10 -Compress
        return $json
    }

    [string]ToString() {
        return $this.ToJson()
    }
}

class Logger {

    # The log directory
    [string]$LogDir

    hidden [string]$LogFile

    # Our logging level
    # Any log messages less than or equal to this will be logged
    [LogLevel]$LogLevel

    # The max size for the log files before rolling
    [int]$MaxSizeMB

    # Number of each log file type to keep
    [int]$FilesToKeep

    # Create logs files under provided directory
    Logger([string]$LogDir, [LogLevel]$LogLevel, [int]$MaxLogSizeMB, [int]$MaxLogsToKeep) {
        $this.LogDir = $LogDir
        $this.LogLevel = $LogLevel
        $this.MaxSizeMB = $MaxLogSizeMB
        $this.FilesToKeep = $MaxLogsToKeep
        $this.LogFile = Join-Path -Path $this.LogDir -ChildPath 'PoshBot.log'
        $this.CreateLogFile()
        $this.Log([LogMessage]::new("Log level set to [$($this.LogLevel)]"))
    }

    hidden Logger() { }

    # Create new log file or roll old log
    hidden [void]CreateLogFile() {
        if (Test-Path -Path $this.LogFile) {
            $this.RollLog($this.LogFile, $true)
        }
        Write-Debug -Message "[Logger:Logger] Creating log file [$($this.LogFile)]"
        New-Item -Path $this.LogFile -ItemType File -Force
    }

    # Log the message and optionally write to console
    [void]Log([LogMessage]$Message) {
        switch ($Message.Severity.ToString()) {
            'Normal' {
                if ($global:VerbosePreference -eq 'Continue') {
                    Write-Verbose -Message $Message.ToJson()
                } elseIf ($global:DebugPreference -eq 'Continue') {
                    Write-Debug -Message $Message.ToJson()
                }
                break
            }
            'Warning' {
                if ($global:WarningPreference -eq 'Continue') {
                    Write-Warning -Message $Message.ToJson()
                }
                break
            }
            'Error' {
                if ($global:ErrorActionPreference -eq 'Continue') {
                    Write-Error -Message $Message.ToJson()
                }
                break
            }
        }

        if ($Message.LogLevel.value__ -le $this.LogLevel.value__) {
            $this.RollLog($this.LogFile, $false)
            $json = $Message.ToJson()
            $this.WriteLine($json)
        }
    }

    [void]Log([LogMessage]$Message, [string]$LogFile, [int]$MaxLogSizeMB, [int]$MaxLogsToKeep) {
        $this.RollLog($LogFile, $false, $MaxLogSizeMB, $MaxLogSizeMB)
        $json = $Message.ToJson()
        $sw = [System.IO.StreamWriter]::new($LogFile, [System.Text.Encoding]::UTF8)
        $sw.WriteLine($json)
        $sw.Close()
    }

    # Write line to file
    hidden [void]WriteLine([string]$Message) {
        $sw = [System.IO.StreamWriter]::new($this.LogFile, [System.Text.Encoding]::UTF8)
        $sw.WriteLine($Message)
        $sw.Close()
    }

    hidden [void]RollLog([string]$LogFile, [bool]$Always) {
        $this.RollLog($LogFile, $Always, $this.MaxSizeMB, $this.FilesToKeep)
    }

    # Checks to see if file in question is larger than the max size specified for the logger.
    # If it is, it will roll the log and delete older logs to keep our number of logs per log type to
    # our max specifiex in the logger.
    # Specified $Always = $true will roll the log regardless
    hidden [void]RollLog([string]$LogFile, [bool]$Always, $MaxLogSize, $MaxFilesToKeep) {

        $keep = $MaxFilesToKeep - 1

        if (Test-Path -Path $LogFile) {
            if ((($file = Get-Item -Path $logFile) -and ($file.Length/1mb) -gt $MaxLogSize) -or $Always) {
                # Remove the last item if it would go over the limit
                if (Test-Path -Path "$logFile.$keep") {
                    Remove-Item -Path "$logFile.$keep"
                }
                foreach ($i in $keep..1) {
                    if (Test-path -Path "$logFile.$($i-1)") {
                        Move-Item -Path "$logFile.$($i-1)" -Destination "$logFile.$i"
                    }
                }
                Move-Item -Path $logFile -Destination "$logFile.$i"
                New-Item -Path $LogFile -Type File -Force > $null
            }
        }
    }
}

class BaseLogger {

    [Logger]$Logger

    BaseLogger() {}

    BaseLogger([string]$LogDirectory, [LogLevel]$LogLevel, [int]$MaxLogSizeMB, [int]$MaxLogsToKeep) {
        $this.Logger = [Logger]::new($LogDirectory, $LogLevel, $MaxLogSizeMB, $MaxLogsToKeep)
    }

    [void]LogInfo([string]$Message) {
        $logMessage = [LogMessage]::new($Message)
        $logMessage.LogLevel = [LogLevel]::Info
        $this.Log($logMessage)
    }

    [void]LogInfo([string]$Message, [object]$Data) {
        $logMessage = [LogMessage]::new($Message, $Data)
        $logMessage.LogLevel = [LogLevel]::Info
        $this.Log($logMessage)
    }

    [void]LogInfo([LogSeverity]$Severity, [string]$Message) {
        $logMessage = [LogMessage]::new($Severity, $Message)
        $logMessage.LogLevel = [LogLevel]::Info
        $this.Log($logMessage)
    }

    [void]LogInfo([LogSeverity]$Severity, [string]$Message, [object]$Data) {
        $logMessage = [LogMessage]::new($Severity, $Message, $Data)
        $logMessage.LogLevel = [LogLevel]::Info
        $this.Log($logMessage)
    }

    [void]LogVerbose([string]$Message) {
        $logMessage = [LogMessage]::new($Message)
        $logMessage.LogLevel = [LogLevel]::Verbose
        $this.Log($logMessage)
    }

    [void]LogVerbose([string]$Message, [object]$Data) {
        $logMessage = [LogMessage]::new($Message, $Data)
        $logMessage.LogLevel = [LogLevel]::Verbose
        $this.Log($logMessage)
    }

    [void]LogVerbose([LogSeverity]$Severity, [string]$Message) {
        $logMessage = [LogMessage]::new($Severity, $Message)
        $logMessage.LogLevel = [LogLevel]::Verbose
        $this.Log($logMessage)
    }

    [void]LogVerbose([LogSeverity]$Severity, [string]$Message, [object]$Data) {
        $logMessage = [LogMessage]::new($Severity, $Message, $Data)
        $logMessage.LogLevel = [LogLevel]::Verbose
        $this.Log($logMessage)
    }

    [void]LogDebug([string]$Message) {
        $logMessage = [LogMessage]::new($Message)
        $logMessage.LogLevel = [LogLevel]::Debug
        $this.Log($logMessage)
    }

    [void]LogDebug([string]$Message, [object]$Data) {
        $logMessage = [LogMessage]::new($Message, $Data)
        $logMessage.LogLevel = [LogLevel]::Debug
        $this.Log($logMessage)
    }

    [void]LogDebug([LogSeverity]$Severity, [string]$Message) {
        $logMessage = [LogMessage]::new($Severity, $Message)
        $logMessage.LogLevel = [LogLevel]::Debug
        $this.Log($logMessage)
    }

    [void]LogDebug([LogSeverity]$Severity, [string]$Message, [object]$Data) {
        $logMessage = [LogMessage]::new($Severity, $Message, $Data)
        $logMessage.LogLevel = [LogLevel]::Debug
        $this.Log($logMessage)
    }

    # Determine source class/method that called the log method
    # and add to log message before sending to logger
    hidden [void]Log([LogMessage]$Message) {
        $Message.Class = $this.GetType().Name
        $Message.Method = @(Get-PSCallStack)[2].FunctionName
        $this.Logger.Log($Message)
    }

    hidden [void]Log([LogMessage]$Message, [string]$LogFile, [int]$MaxLogSizeMB, [int]$MaxLogsToKeep) {
        $Message.Class = $this.GetType().Name
        $Message.Method = @(Get-PSCallStack)[2].FunctionName
        $this.Logger.Log($Message, $LogFile, $MaxLogSizeMB, $MaxLogsToKeep)
    }
}

class ExceptionFormatter {

    static [pscustomobject]Summarize([System.Management.Automation.ErrorRecord]$Exception) {
        return [pscustomobject]@{
            CommandName = $Exception.InvocationInfo.MyCommand.Name
            Message = $Exception.Exception.Message
            TargetObject = $Exception.TargetObject
            Position = $Exception.InvocationInfo.PositionMessage
            CategoryInfo = $Exception.CategoryInfo.ToString()
            FullyQualifiedErrorId = $Exception.FullyQualifiedErrorId
        }
    }

    static [string]ToJson([System.Management.Automation.ErrorRecord]$Exception) {
        return ([ExceptionFormatter]::Summarize($Exception) | ConvertTo-Json)
    }
}

# An Event is something that happended on a chat network. A person joined a room, a message was received
# Really any notification from the chat network back to the bot is considered an Event of some sort
class Event {
    [string]$Type
    [string]$ChannelId
    [pscustomobject]$Data
}


# Represents a person on a chat network
class Person {

    [string]$Id

    # The identifier for the device or client the person is using
    [string]$ClientId

    [string]$Nickname
    [string]$FirstName
    [string]$LastName
    [string]$FullName

    [string]ToString() {
        return "$($this.id):$($this.NickName):$($this.FullName)"
    }

    [hashtable]ToHash() {
        $hash = @{}
        (Get-Member -InputObject $this -MemberType Property).foreach({
            $hash.Add($_.Name, $this.($_.name))
        })
        return $hash
    }
}

class Room {

    [string]$Id

    # The name of the room
    [string]$Name

    # The room topic
    [string]$Topic

    # Indicates if this room already exists or not
    [bool]$Exists

    # Indicates if this room has already been joined
    [bool]$Joined

    [hashtable]$Members = @{}

    Room() {}

    [string]Join() {
        throw 'Must Override Method'
    }

    [string]Leave() {
        throw 'Must Override Method'
    }

    [string]Create() {
        throw 'Must Override Method'
    }

    [string]Destroy() {
        throw 'Must Override Method'
    }

    [string]Invite([string[]]$Invitees) {
        throw 'Must Override Method'
    }
}

# A response message that is sent back to the chat network.
class Response {
    [Severity]$Severity = [Severity]::Success
    [string[]]$Text
    [string]$MessageFrom
    [string]$To
    [Message]$OriginalMessage = [Message]::new()
    [pscustomobject[]]$Data = @()

    Response() {}

    Response([Message]$Message) {
        $this.MessageFrom = $Message.From
        $this.To = $Message.To
        $this.OriginalMessage = $Message
    }

    [pscustomobject] Summarize() {
        return [pscustomobject]@{
            Severity        = $this.Severity.ToString()
            Text            = $this.Text
            MessageFrom     = $this.MessageFrom
            To              = $this.To
            OriginalMessage = $this.OriginalMessage
            Data            = $this.Data
        }
    }
}

# A chat message that is received from the chat network
class Message {
    [string]$Id                 # MAY have this
    [MessageType]$Type = [MessageType]::Message
    [MessageSubtype]$Subtype = [MessageSubtype]::None    # Some messages have subtypes
    [string]$Text               # Text of message. This may be empty depending on the message type
    [string]$To                 # Id of user/channel the message is to
    [string]$ToName             # Name of user/channel the message is to
    [string]$From               # ID of user who sent the message
    [string]$FromName           # Name of user who sent the message
    [datetime]$Time             # The date/time (UTC) the message was received
    [bool]$IsDM                 # Denotes if message is a direct message
    [hashtable]$Options         # Any other bits of information about a message. This will be backend specific
    [pscustomobject]$RawMessage # The raw message as received by the backend. This can be usefull for the backend

    [Message]Clone () {
        $newMsg = [Message]::New()
        foreach ($prop in ($this | Get-Member -MemberType Property)) {
            if ('Clone' -in ($this.$($prop.Name) | Get-Member -MemberType Method -ErrorAction Ignore).Name) {
                $newMsg.$($prop.Name) = $this.$($prop.Name).Clone()
            } else {
                $newMsg.$($prop.Name) = $this.$($prop.Name)
            }
        }
        return $newMsg
    }

    [hashtable] ToHash() {
        return @{
            Id         = $this.Id
            Type       = $this.Type.ToString()
            Subtype    = $this.Subtype.ToString()
            Text       = $this.Text
            To         = $this.To
            ToName     = $this.ToName
            From       = $this.From
            FromName   = $this.FromName
            Time       = $this.Time.ToUniversalTime().ToString('u')
            IsDM       = $this.IsDM
            Options    = $this.Options
            RawMessage = $this.RawMessage
        }
    }

    [string] ToJson() {
        return $this.ToHash() | ConvertTo-Json -Depth 10 -Compress
    }
}

class Stream {
    [object[]]$Debug = @()
    [object[]]$Error = @()
    [object[]]$Information = @()
    [object[]]$Verbose = @()
    [object[]]$Warning = @()
}

# Represents the result of running a command
class CommandResult {
    [bool]$Success
    [object[]]$Errors = @()
    [object[]]$Output = @()
    [Stream]$Streams = [Stream]::new()
    [bool]$Authorized = $true
    [timespan]$Duration

    [pscustomobject]Summarize() {
        return [pscustomobject]@{
            Success = $this.Success
            Output = $this.Output
            Errors = foreach ($item in $this.Errors) {
                # Summarize exceptions so they can be serialized to json correctly
                if ($item -is [System.Management.Automation.ErrorRecord]) {
                    [ExceptionFormatter]::Summarize($item)
                } else {
                    $item
                }
            }
            Authorized = $this.Authorized
            Duration = $this.Duration.TotalSeconds
        }
    }

    [string]ToJson() {
        return $this.Summarize() | ConvertTo-Json -Depth 10 -Compress
    }
}

class ParsedCommand {
    [string]$CommandString
    [string]$Plugin = $null
    [string]$Command = $null
    [string]$Version = $null
    [hashtable]$NamedParameters = @{}
    [System.Collections.ArrayList]$PositionalParameters = (New-Object System.Collections.ArrayList)
    [datetime]$Time = (Get-Date).ToUniversalTime()
    [string]$From = $null
    [string]$FromName = $null
    [hashtable]$CallingUserInfo = @{}
    [string]$To = $null
    [string]$ToName = $null
    [Message]$OriginalMessage

    [pscustomobject]Summarize() {
        $o = $this | Select-Object -Property * -ExcludeProperty NamedParameters
        if ($this.Plugin -eq 'Builtin') {
            $np = $this.NamedParameters.GetEnumerator() | Where-Object {$_.Name -ne 'Bot'}
            $o | Add-Member -MemberType NoteProperty -Name NamedParameters -Value $np
        } else {
            $o | Add-Member -MemberType NoteProperty -Name NamedParameters -Value $this.NamedParameters
        }
        return [pscustomobject]$o
    }
}

class CommandParser {
    [ParsedCommand] static Parse([Message]$Message) {

        $commandString = [string]::Empty
        if (-not [string]::IsNullOrEmpty($Message.Text)) {
            $commandString = $Message.Text.Trim()
        }

        # The command is the first word of the message
        $cmdArray = $commandString.Split(' ')
        $command = $cmdArray[0]
        if ($cmdArray.Count -gt 1) {
            $commandArgs = $cmdArray[1..($cmdArray.length-1)] -join ' '
        } else {
            $commandArgs = [string]::Empty
        }

        # The first word of the message COULD be a URI, don't try and parse than into a command
        if ($command -notlike '*://*') {
            $arrCmdStr = $command.Split(':')
        } else {
            $arrCmdStr = @($command)
        }

        # Check if a specific version of the command was specified
        $version = $null
        if ($arrCmdStr[1] -as [Version]) {
            $version = $arrCmdStr[1]
        } elseif ($arrCmdStr[2] -as [Version]) {
            $version = $arrCmdStr[2]
        }

        # The command COULD be in the form of <command> or <plugin:command>
        # Figure out which one
        $plugin = [string]::Empty
        if ($Message.Type -eq [MessageType]::Message -and $Message.SubType -eq [MessageSubtype]::None ) {
            $plugin = $arrCmdStr[0]
        }
        if ($arrCmdStr[1] -as [Version]) {
            $command = $arrCmdStr[0]
            $plugin = $null
        } else {
            $command = $arrCmdStr[1]
            if (-not $command) {
                $command = $plugin
                $plugin = $null
            }
        }

        # Create the ParsedCommand instance
        $parsedCommand = [ParsedCommand]::new()
        $parsedCommand.CommandString = $commandString
        $parsedCommand.Plugin = $plugin
        $parsedCommand.Command = $command
        $parsedCommand.OriginalMessage = $Message
        $parsedCommand.Time = $Message.Time
        if ($version)          { $parsedCommand.Version  = $version }
        if ($Message.To)       { $parsedCommand.To       = $Message.To }
        if ($Message.ToName)   { $parsedCommand.ToName   = $Message.ToName }
        if ($Message.From)     { $parsedCommand.From     = $Message.From }
        if ($Message.FromName) { $parsedCommand.FromName = $Message.FromName }

        # Parse the message text using AST into named and positional parameters
        try {
            $positionalParams = @()
            $namedParams = @{}

            if (-not [string]::IsNullOrEmpty($commandArgs)) {

                # Create Abstract Syntax Tree of command string so we can parse out parameter names
                # and their values
                $astCmdStr = "fake-command $commandArgs" -Replace '(\s--([a-zA-Z0-9])*?)', ' -$2'
                $ast = [System.Management.Automation.Language.Parser]::ParseInput($astCmdStr, [ref]$null, [ref]$null)
                $commandAST = $ast.FindAll({$args[0] -as [System.Management.Automation.Language.CommandAst]},$false)

                for ($x = 1; $x -lt $commandAST.CommandElements.Count; $x++) {
                    $element = $commandAST.CommandElements[$x]

                    # The element is a command parameter (meaning -<ParamName>)
                    # Determine the values for it
                    if ($element -is [System.Management.Automation.Language.CommandParameterAst]) {

                        $paramName = $element.ParameterName
                        $paramValues = @()
                        $y = 1

                        # If the element after this one is another CommandParameterAst or this
                        # is the last element then assume this parameter is a [switch]
                        if ((-not $commandAST.CommandElements[$x+1]) -or ($commandAST.CommandElements[$x+1] -is [System.Management.Automation.Language.CommandParameterAst])) {
                            $paramValues = $true
                        } else {
                            # Inspect the elements immediately after this CommandAst as they are values
                            # for a named parameter and pull out the values (array, string, bool, etc)
                            do {
                                $elementValue = $commandAST.CommandElements[$x+$y]

                                if ($elementValue -is [System.Management.Automation.Language.VariableExpressionAst]) {
                                    # The element 'looks' like a variable reference
                                    # Get the raw text of the value
                                    $paramValues += $elementValue.Extent.Text
                                } else {
                                    if ($elementValue.Value) {
                                       $paramValues += $elementValue.Value
                                    } else {
                                        $paramValues += $elementValue.SafeGetValue()
                                    }
                                }
                                $y++
                            } until ((-not $commandAST.CommandElements[$x+$y]) -or $commandAST.CommandElements[$x+$y] -is [System.Management.Automation.Language.CommandParameterAst])
                        }

                        if ($paramValues.Count -eq 1) {
                            $paramValues = $paramValues[0]
                        }
                        $namedParams.Add($paramName, $paramValues)
                        $x += $y-1
                    } else {
                        # This element is a positional parameter value so just get the value
                        if ($element -is [System.Management.Automation.Language.VariableExpressionAst]) {
                            $positionalParams += $element.Extent.Text
                        } else {
                            if ($element.Value) {
                                $positionalParams += $element.Value
                            } else {
                                $positionalParams += $element.SafeGetValue()
                            }
                        }
                    }
                }
            }

            $parsedCommand.NamedParameters = $namedParams
            $parsedCommand.PositionalParameters = $positionalParams
        } catch {
            Write-Error -Message "Error parsing command [$CommandString]: $_"
        }

        return $parsedCommand
    }
}

class Permission {
    [string]$Name
    [string]$Plugin
    [string]$Description
    [bool]$Adhoc = $false

    Permission([string]$Name) {
        $this.Name = $Name
    }

    Permission([string]$Name, [string]$Plugin) {
        $this.Name = $Name
        $this.Plugin = $Plugin
    }

    Permission([string]$Name, [string]$Plugin, [string]$Description) {
        $this.Name = $Name
        $this.Plugin = $Plugin
        $this.Description = $Description
    }

    [hashtable]ToHash() {
        return @{
            Name = $this.Name
            Plugin = $this.Plugin
            Description = $this.Description
            Adhoc = $this.Adhoc
        }
    }

    [string]ToString() {
        return "$($this.Plugin):$($this.Name)"
    }
}

class CommandAuthorizationResult {
    [bool]$Authorized
    [string]$Message

    CommandAuthorizationResult() {
        $this.Authorized = $true
    }

    CommandAuthorizationResult([bool]$Authorized) {
        $this.Authorized = $Authorized
    }

    CommandAuthorizationResult([bool]$Authorized, [string]$Message) {
        $this.Authorized = $Authorized
        $this.Message = $Message
    }
}

# An access filter controls under what conditions a command can be run and who can run it.
class AccessFilter {

    [hashtable]$Permissions = @{}

    [CommandAuthorizationResult]Authorize([string]$PermissionName) {
        if ($this.Permissions.Count -eq 0) {
            return $true
        } else {
            if (-not $this.Permissions.ContainsKey($PermissionName)) {
                return [CommandAuthorizationResult]::new($false, "Permission [$PermissionName] is not authorized to execute this command")
            } else {
                return $true
            }
        }
    }

    [void]AddPermission([Permission]$Permission) {
        if (-not $this.Permissions.ContainsKey($Permission.ToString())) {
            $this.Permissions.Add($Permission.ToString(), $Permission)
        }
    }

    [void]RemovePermission([Permission]$Permission) {
        if ($this.Permissions.ContainsKey($Permission.ToString())) {
            $this.Permissions.Remove($Permission.ToString())
        }
    }
}

class Role : BaseLogger {
    [string]$Name
    [string]$Description
    [hashtable]$Permissions = @{}

    Role([string]$Name, [Logger]$Logger) {
        $this.Name = $Name
        $this.Logger = $Logger
    }

    Role([string]$Name, [string]$Description, [Logger]$Logger) {
        $this.Name = $Name
        $this.Description = $Description
        $this.Logger = $Logger
    }

    [void]AddPermission([Permission]$Permission) {
        if (-not $this.Permissions.ContainsKey($Permission.ToString())) {
            $this.LogVerbose("Adding permission [$($Permission.Name)] to role [$($this.Name)]")
            $this.Permissions.Add($Permission.ToString(), $Permission)
        }
    }

    [void]RemovePermission([Permission]$Permission) {
        if ($this.Permissions.ContainsKey($Permission.ToString())) {
            $this.LogVerbose("Removing permission [$($Permission.Name)] from role [$($this.Name)]")
            $this.Permissions.Remove($Permission.ToString())
        }
    }

    [hashtable]ToHash() {
        return @{
            Name = $this.Name
            Description = $this.Description
            Permissions = @($this.Permissions.Keys)
        }
    }
}

# A group contains a collection of users and a collection of roles
# those users will be a member of
class Group : BaseLogger {
    [string]$Name
    [string]$Description
    [hashtable]$Users = @{}
    [hashtable]$Roles = @{}

    Group([string]$Name, [Logger]$Logger) {
        $this.Name = $Name
        $this.Logger = $Logger
    }

    Group([string]$Name, [string]$Description, [Logger]$Logger) {
        $this.Name = $Name
        $this.Description = $Description
        $this.Logger = $Logger
    }

    [void]AddRole([Role]$Role) {
        if (-not $this.Roles.ContainsKey($Role.Name)) {
            $this.LogVerbose("Adding role [$($Role.Name)] to group [$($this.Name)]")
            $this.Roles.Add($Role.Name, $Role)
        } else {
            $this.LogVerbose([LogSeverity]::Warning, "Role [$($Role.Name)] is already in group [$($this.Name)]")
        }
    }

    [void]RemoveRole([Role]$Role) {
        if ($this.Roles.ContainsKey($Role.Name)) {
            $this.LogVerbose("Removing role [$($Role.Name)] from group [$($this.Name)]")
            $this.Roles.Remove($Role.Name)
        }
    }

    [void]AddUser([string]$Username) {
        if (-not $this.Users.ContainsKey($Username)) {
            $this.LogVerbose("Adding user [$Username)] to group [$($this.Name)]")
            $this.Users.Add($Username, $null)
        } else {
            $this.LogVerbose([LogSeverity]::Warning, "User [$Username)] is already in group [$($this.Name)]")
        }
    }

    [void]RemoveUser([string]$Username) {
        if ($this.Users.ContainsKey($Username)) {
            $this.LogVerbose("Removing user [$Username)] from group [$($this.Name)]")
            $this.Users.Remove($Username)
        }
    }

    [hashtable]ToHash() {
        return @{
            Name = $this.Name
            Description = $this.Description
            Users = $this.Users.Keys
            Roles = $this.Roles.Keys
        }
    }
}

class Trigger {
    [TriggerType]$Type
    [string]$Trigger
    [MessageType]$MessageType = [MessageType]::Message
    [MessageSubType]$MessageSubtype = [Messagesubtype]::None

    Trigger([TriggerType]$Type, [string]$Trigger) {
        $this.Type = $Type
        $this.Trigger = $Trigger
    }
}
#requires -Modules Configuration

class StorageProvider : BaseLogger {

    [string]$ConfigPath

    StorageProvider([Logger]$Logger) {
        $this.Logger = $Logger
        $this.ConfigPath = $script:defaultPoshBotDir
    }

    StorageProvider([string]$Dir, [Logger]$Logger) {
        $this.Logger = $Logger
        $this.ConfigPath = $Dir
    }

    [hashtable]GetConfig([string]$ConfigName) {
        $path = Join-Path -Path $this.ConfigPath -ChildPath "$($ConfigName).psd1"
        if (Test-Path -Path $path) {
            $this.LogDebug("Loading config [$ConfigName] from [$path]")
            $config = Get-Content -Path $path -Raw | ConvertFrom-Metadata
            return $config
        } else {
            $this.LogInfo([LogSeverity]::Warning, "Configuration file [$path] not found")
            return $null
        }
    }

    [void]SaveConfig([string]$ConfigName, [hashtable]$Config) {
        $path = Join-Path -Path $this.ConfigPath -ChildPath "$ConfigName.psd1"
        $meta = $config | ConvertTo-Metadata
        if (-not (Test-Path -Path $path)) {
            New-Item -Path $Path -ItemType File
        }
        $this.LogDebug("Saving config [$ConfigName] to [$path]")
        $meta | Out-file -FilePath $path -Force -Encoding utf8
    }
}

class RoleManager : BaseLogger {
    [hashtable]$Groups = @{}
    [hashtable]$Permissions = @{}
    [hashtable]$Roles = @{}
    [hashtable]$RoleUserMapping = @{}
    hidden [object]$_Backend
    hidden [StorageProvider]$_Storage
    hidden [string[]]$_AdminPermissions = @('manage-roles', 'show-help' ,'view', 'view-role', 'view-group',
                                           'manage-plugins', 'manage-groups', 'manage-permissions', 'manage-schedules')

    RoleManager([object]$Backend, [StorageProvider]$Storage, [Logger]$Logger) {
        $this._Backend = $Backend
        $this._Storage = $Storage
        $this.Logger = $Logger
        $this.Initialize()
    }

    [void]Initialize() {
        # Load in state from persistent storage
        $this.LogInfo('Initializing')

        $this.LoadState()

        # Create the initial state of the [Admin] role ONLY if it didn't get loaded from storage
        # This could be because this is the first time the bot has run and [roles.psd1] doesn't exist yet.
        # The bot admin could have modified the permissions for the role and we want to respect those changes
        if (-not $this.Roles['Admin']) {
            # Create the builtin Admin role and add all the permissions defined in the [Builtin] module
            $this.LogDebug('Creating builtin [Admin] role')
            $adminrole = [Role]::New('Admin', 'Bot administrator role', $this.Logger)

            # TODO
            # Get the builtin permissions from the module manifest rather than hard coding them in the class
            $this._AdminPermissions | foreach-object {
                $p = [Permission]::new($_, 'Builtin')
                $adminRole.AddPermission($p)
            }
            $this.LogDebug('Added builtin permissions to [Admin] role', $this._AdminPermissions)
            $this.Roles.Add($adminRole.Name, $adminRole)

            # Creat the builtin [Admin] group and add the [Admin] role to it
            $this.LogDebug('Creating builtin [Admin] group with [Admin] role')
            $adminGroup = [Group]::new('Admin', 'Bot administrators', $this.Logger)
            $adminGroup.AddRole($adminRole)
            $this.Groups.Add($adminGroup.Name, $adminGroup)
            $this.SaveState()
        } else {
            # Make sure all the admin permissions are added to the 'Admin' role
            # This is so if we need to add any permissions in future versions, they will automatically
            # be added to the role
            $adminRole = $this.Roles['Admin']
            foreach ($perm in $this._AdminPermissions) {
                if (-not $adminRole.Permissions.ContainsKey($perm)) {
                    $this.LogInfo("[Admin] role missing builtin permission [$perm]. Adding permission back.")
                    $p = [Permission]::new($perm, 'Builtin')
                    $adminRole.AddPermission($p)
                }
            }
        }
    }

    # Save state to storage
    [void]SaveState() {
        $this.LogDebug('Saving role manager state to storage')

        $permissionsToSave = @{}
        foreach ($permission in $this.Permissions.GetEnumerator()) {
            $permissionsToSave.Add($permission.Name, $permission.Value.ToHash())
        }
        $this._Storage.SaveConfig('permissions', $permissionsToSave)

        $rolesToSave = @{}
        foreach ($role in $this.Roles.GetEnumerator()) {
            $rolesToSave.Add($role.Name, $role.Value.ToHash())
        }
        $this._Storage.SaveConfig('roles', $rolesToSave)

        $groupsToSave = @{}
        foreach ($group in $this.Groups.GetEnumerator()) {
            $groupsToSave.Add($group.Name, $group.Value.ToHash())
        }
        $this._Storage.SaveConfig('groups', $groupsToSave)
    }

    # Load state from storage
    [void]LoadState() {
        $this.LogDebug('Loading role manager state from storage')

        $permissionConfig = $this._Storage.GetConfig('permissions')
        if ($permissionConfig) {
            foreach($permKey in $permissionConfig.Keys) {
                $perm = $permissionConfig[$permKey]
                $p = [Permission]::new($perm.Name, $perm.Plugin)
                if ($perm.Adhoc) {
                    $p.Adhoc = $perm.Adhoc
                }
                if ($perm.Description) {
                    $p.Description = $perm.Description
                }
                if (-not $this.Permissions.ContainsKey($p.ToString())) {
                    $this.Permissions.Add($p.ToString(), $p)
                }
            }
        }

        $roleConfig = $this._Storage.GetConfig('roles')
        if ($roleConfig) {
            foreach ($roleKey in $roleConfig.Keys) {
                $role = $roleConfig[$roleKey]
                $r = [Role]::new($roleKey, $this.Logger)
                if ($role.Description) {
                    $r.Description = $role.Description
                }
                if ($role.Permissions) {
                    foreach ($perm in $role.Permissions) {
                        if ($p = $this.Permissions[$perm]) {
                            $r.AddPermission($p)
                        }
                    }
                }
                if (-not $this.Roles.ContainsKey($r.Name)) {
                    $this.Roles.Add($r.Name, $r)
                }
            }
        }

        $groupConfig = $this._Storage.GetConfig('groups')
        if ($groupConfig) {
            foreach ($groupKey in $groupConfig.Keys) {
                $group = $groupConfig[$groupKey]
                $g = [Group]::new($groupKey, $this.Logger)
                if ($group.Description) {
                    $g.Description = $group.Description
                }
                if ($group.Users) {
                    foreach ($u in $group.Users) {
                        $g.AddUser($u)
                    }
                }
                if ($group.Roles) {
                    foreach ($r in $group.Roles) {
                        if ($ro = $this.GetRole($r)) {
                            $g.AddRole($ro)
                        }
                    }
                }
                if (-not $this.Groups.ContainsKey($g.Name)) {
                    $this.Groups.Add($g.Name, $g)
                }
            }
        }
    }

    [Group]GetGroup([string]$Groupname) {
        if ($g = $this.Groups[$Groupname]) {
            return $g
        } else {
            $this.LogInfo([LogSeverity]::Warning, "Group [$Groupname] not found")
            return $null
        }
    }

    [void]UpdateGroupDescription([string]$Groupname, [string]$Description) {
        if ($g = $this.Groups[$Groupname]) {
            $g.Description = $Description
            $this.SaveState()
        } else {
            $this.LogInfo([LogSeverity]::Warning, "Group [$Groupname] not found")
        }
    }

    [void]UpdateRoleDescription([string]$Rolename, [string]$Description) {
        if ($r = $this.Roles[$Rolename]) {
            $r.Description = $Description
            $this.SaveState()
        } else {
            $this.LogInfo([LogSeverity]::Warning, "Role [$Rolename] not found")
        }
    }

    [Permission]GetPermission([string]$PermissionName) {
        $p = $this.Permissions[$PermissionName]
        if ($p) {
            return $p
        } else {
            $this.LogInfo([LogSeverity]::Warning, "Permission [$PermissionName] not found")
            return $null
        }
    }

    [Role]GetRole([string]$RoleName) {
        $r = $this.Roles[$RoleName]
        if ($r) {
            return $r
        } else {
            $this.LogInfo([LogSeverity]::Warning, "Role [$RoleName] not found")
            return $null
        }
    }

    [void]AddGroup([Group]$Group) {
        if (-not $this.Groups.ContainsKey($Group.Name)) {
            $this.LogVerbose("Adding group [$($Group.Name)]")
            $this.Groups.Add($Group.Name, $Group)
            $this.SaveState()
        } else {
            $this.LogInfo([LogSeverity]::Warning, "Group [$($Group.Name)] is already loaded")
        }
    }

    [void]AddPermission([Permission]$Permission) {
        if (-not $this.Permissions.ContainsKey($Permission.ToString())) {
            $this.LogVerbose("Adding permission [$($Permission.Name)]")
            $this.Permissions.Add($Permission.ToString(), $Permission)
            $this.SaveState()
        } else {
            $this.LogInfo([LogSeverity]::Warning, "Permission [$($Permission.Name)] is already loaded")
        }
    }

    [void]AddRole([Role]$Role) {
        if (-not $this.Roles.ContainsKey($Role.Name)) {
            $this.LogVerbose("Adding role [$($Role.Name)]")
            $this.Roles.Add($Role.Name, $Role)
            $this.SaveState()
        } else {
            $this.LogInfo([LogSeverity]::Warning, "Role [$($Role.Name)] is already loaded")
        }
    }

    [void]RemoveGroup([Group]$Group) {
        if ($this.Groups.ContainsKey($Group.Name)) {
            $this.LogVerbose("Removing group [$($Group.Name)]")
            $this.Groups.Remove($Group.Name)
            $this.SaveState()
        } else {
            $this.LogInfo([LogSeverity]::Warning, "Group [$($Group.Name)] was not found")
        }
    }

    [void]RemovePermission([Permission]$Permission) {
        if (-not $this.Permissions.ContainsKey($Permission.ToString())) {
            # Remove the permission from roles
            foreach ($role in $this.Roles.GetEnumerator()) {
                if ($role.Value.Permissions.ContainsKey($Permission.ToString())) {
                    $this.LogVerbose("Removing permission [$($Permission.ToString())] from role [$($role.Value.Name)]")
                    $role.Value.RemovePermission($Permission)
                }
            }

            $this.LogVerbose("Removing permission [$($Permission.ToString())]")
            $this.Permissions.Remove($Permission.ToString())
            $this.SaveState()
        } else {
            $this.LogInfo([LogSeverity]::Warning, "Permission [$($Permission.ToString())] was not found")
        }
    }

    [void]RemoveRole([Role]$Role) {
        if ($this.Roles.ContainsKey($Role.Name)) {
            # Remove the role from groups
            foreach ($group in $this.Groups.GetEnumerator()) {
                if ($group.Value.Roles.ContainsKey($Role.Name)) {
                    $this.LogVerbose("Removing role [$($Role.Name)] from group [$($group.Value.Name)]")
                    $group.Value.RemoveRole($Role)
                }
            }

            $this.LogVerbose("Removing role [$($Role.Name)]")
            $this.Roles.Remove($Role.Name)
            $this.SaveState()
        } else {
            $this.LogInfo([LogSeverity]::Warning, "Role [$($Role.Name)] was not found")
        }
    }

    [void]AddRoleToGroup([string]$RoleName, [string]$GroupName) {
        try {
            if ($role = $this.GetRole($RoleName)) {
                if ($group = $this.Groups[$GroupName]) {
                    $this.LogVerbose("Adding role [$RoleName] to group [$($group.Name)]")
                    $group.AddRole($role)
                    $this.SaveState()
                } else {
                    $msg = "Unknown group [$GroupName]"
                    $this.LogInfo([LogSeverity]::Warning, $msg)
                    throw $msg
                }
            } else {
                $msg = "Unable to find role [$RoleName]"
                $this.LogInfo([LogSeverity]::Warning, $msg)
                throw $msg
            }
        } catch {
            throw $_
        }
    }

    [void]AddUserToGroup([string]$UserId, [string]$GroupName) {
        try {
            if ($this._Backend.GetUser($UserId)) {
                if ($group = $this.Groups[$GroupName]) {
                    $this.LogVerbose("Adding user [$UserId] to [$($group.Name)]")
                    $group.AddUser($UserId)
                    $this.SaveState()
                } else {
                    $msg = "Unknown group [$GroupName]"
                    $this.LogInfo([LogSeverity]::Warning, $msg)
                    throw $msg
                }
            } else {
                $msg = "Unable to find user [$UserId]"
                $this.LogInfo([LogSeverity]::Warning, $msg)
                throw $msg
            }
        } catch {
            $this.LogInfo([LogSeverity]::Error, "Exception adding [$UserId] to [$GroupName]", $_)
            throw $_
        }
    }

    [void]RemoveRoleFromGroup([string]$RoleName, [string]$GroupName) {
        try {
            if ($role = $this.GetRole($RoleName)) {
                if ($group = $this.Groups[$GroupName]) {
                    $this.LogVerbose("Removing role [$RoleName] from group [$($group.Name)]")
                    $group.RemoveRole($role)
                    $this.SaveState()
                } else {
                    $msg = "Unknown group [$GroupName]"
                    $this.LogInfo([LogSeverity]::Warning, $msg)
                    throw $msg
                }
            } else {
                $msg = "Unable to find role [$RoleName]"
                $this.LogInfo([LogSeverity]::Warning, $msg)
                throw $msg
            }
        } catch {
            $this.LogInfo([LogSeverity]::Error, "Exception removing [$RoleName] from [$GroupName]", $_)
            throw $_
        }
    }

    [void]RemoveUserFromGroup([string]$UserId, [string]$GroupName) {
        try {
            if ($group = $this.Groups[$GroupName]) {
                if ($group.Users.ContainsKey($UserId)) {
                    $this.LogVerbose("Removing user [$UserId] from group [$($group.Name)]")
                    $group.RemoveUser($UserId)
                    $this.SaveState()
                }
            } else {
                $msg = "Unknown group [$GroupName]"
                $this.LogInfo([LogSeverity]::Warning, $msg)
                throw $msg
            }
        } catch {
            $this.LogInfo([LogSeverity]::Error, "Exception removing [$UserId] from [$GroupName]", $_)
            throw $_
        }
    }

    [void]AddPermissionToRole([string]$PermissionName, [string]$RoleName) {
        try {
            if ($role = $this.GetRole($RoleName)) {
                if ($perm = $this.Permissions[$PermissionName]) {
                    $this.LogVerbose("Adding permission [$PermissionName] to role [$($role.Name)]")
                    $role.AddPermission($perm)
                    $this.SaveState()
                } else {
                    $msg = "Unknown permission [$perm]"
                    $this.LogInfo([LogSeverity]::Warning, $msg)
                    throw $msg
                }
            } else {
                $msg = "Unable to find role [$RoleName]"
                $this.LogInfo([LogSeverity]::Warning, $msg)
                throw $msg
            }
        } catch {
            $this.LogInfo([LogSeverity]::Error, "Exception adding [$PermissionName] to [$RoleName]", $_)
            throw $_
        }
    }

    [void]RemovePermissionFromRole([string]$PermissionName, [string]$RoleName) {
        try {
            if ($role = $this.GetRole($RoleName)) {
                if ($perm = $this.Permissions[$PermissionName]) {
                    $this.LogVerbose("Removing permission [$PermissionName] from role [$($role.Name)]")
                    $role.RemovePermission($perm)
                    $this.SaveState()
                } else {
                    $msg = "Unknown permission [$PermissionName]"
                    $this.LogInfo([LogSeverity]::Warning, $msg)
                    throw $msg
                }
            } else {
                $msg = "Unable to find role [$RoleName]"
                $this.LogInfo([LogSeverity]::Warning, $msg)
                throw $msg
            }
        } catch {
            $this.LogInfo([LogSeverity]::Error, "Exception removing [$PermissionName] from [$RoleName]", $_)
            throw $_
        }
    }

    [Group[]]GetUserGroups([string]$UserId) {
        $userGroups = New-Object System.Collections.ArrayList

        foreach ($group in $this.Groups.GetEnumerator()) {
            if ($group.Value.Users.ContainsKey($UserId)) {
                $userGroups.Add($group.Value)
            }
        }
        return $userGroups
    }

    [Role[]]GetUserRoles([string]$UserId) {
        $userRoles = New-Object System.Collections.ArrayList

        foreach ($group in $this.GetUserGroups($UserId)) {
            foreach ($role in $group.Roles.GetEnumerator()) {
                $userRoles.Add($role.Value)
            }
        }

        return $userRoles
    }

    [Permission[]]GetUserPermissions([string]$UserId) {
        $userPermissions = New-Object System.Collections.ArrayList

        if ($userRoles = $this.GetUserRoles($UserId)) {
            foreach ($role in $userRoles) {
                $userPermissions.AddRange($role.Permissions.Keys)
            }
        }

        return $userPermissions
    }

    # Resolve a username to their Id
    [string]ResolveUserIdToUserName([string]$Id) {
        return $this._Backend.UserIdToUsername($Id)
    }

    [string]ResolveUsernameToId([string]$Username) {
        return $this._Backend.UsernameToUserId($Username)
    }
}

# Some custom exceptions dealing with executing commands
class CommandException : Exception {
    CommandException() {}
    CommandException([string]$Message) : base($Message) {}
}
class CommandNotFoundException : CommandException {
    CommandNotFoundException() {}
    CommandNotFoundException([string]$Message) : base($Message) {}
}
class CommandFailed : CommandException {
    CommandFailed() {}
    CommandFailed([string]$Message) : base($Message) {}
}
class CommandDisabled : CommandException {
    CommandDisabled() {}
    CommandDisabled([string]$Message) : base($Message) {}
}
class CommandNotAuthorized : CommandException {
    CommandNotAuthorized() {}
    CommandNotAuthorized([string]$Message) : base($Message) {}
}
class CommandRequirementsNotMet : CommandException {
    CommandRequirementsNotMet() {}
    CommandRequirementsNotMet([string]$Message) : base($Message) {}
}

# Represent a command that can be executed
[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Scope='Function', Target='*')]
[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Scope='Function', Target='*')]
class Command : BaseLogger {

    # Unique (to the plugin) name of the command
    [string]$Name

    [string[]]$Aliases = @()

    [string]$Description

    [TriggerType]$TriggerType = [TriggerType]::Command

    [Trigger[]]$Triggers = @()

    [string[]]$Usage

    [bool]$KeepHistory = $true

    [bool]$HideFromHelp = $false

    [bool]$AsJob = $true

    # Fully qualified name of a cmdlet or function in a module to execute
    [string]$ModuleQualifiedCommand

    [string]$ManifestPath

    [System.Management.Automation.FunctionInfo]$FunctionInfo

    [System.Management.Automation.CmdletInfo]$CmdletInfo

    [AccessFilter]$AccessFilter = [AccessFilter]::new()

    [bool]$Enabled = $true

    # Cannot have a constructor called "Command". Lame
    # We need to set the Logger property separately
    # Command([Logger]$Logger) {
    #     $this.Logger = $Logger
    # }

    # Execute the command in a PowerShell job and return the running job
    [object]Invoke([ParsedCommand]$ParsedCommand, [bool]$InvokeAsJob = $this.AsJob, [string]$Backend) {

        $outer = {
            [cmdletbinding()]
            param(
                [hashtable]$Options
            )

            Import-Module -Name $Options.PoshBotManifestPath -Force -Verbose:$false -WarningAction SilentlyContinue -ErrorAction Stop

            Import-Module -Name $Options.ManifestPath -Scope Local -Force -Verbose:$false -WarningAction SilentlyContinue

            $namedParameters = $Options.NamedParameters
            $positionalParameters = $Options.PositionalParameters

            # Context for who/how the command was called
            $parsedCommandExcludes = @('From', 'FromName', 'To', 'ToName', 'CallingUserInfo', 'OriginalMessage')
            $global:PoshBotContext = [pscustomobject]@{
                Plugin = $options.ParsedCommand.Plugin
                Command = $options.ParsedCommand.Command
                From = $options.ParsedCommand.From
                FromName = $options.ParsedCommand.FromName
                To = $options.ParsedCommand.To
                ToName = $options.ParsedCommand.ToName
                CallingUserInfo = $options.CallingUserInfo
                ConfigurationDirectory = $options.ConfigurationDirectory
                ParsedCommand = $options.ParsedCommand | Select-Object -ExcludeProperty $parsedCommandExcludes
                OriginalMessage = $options.OriginalMessage
                BackendType = $options.BackendType
            }

            & $Options.ModuleQualifiedCommand @namedParameters @positionalParameters
        }

        [string]$sb = [string]::Empty
        $options = @{
            ManifestPath = $this.ManifestPath
            ParsedCommand = $ParsedCommand
            CallingUserInfo = $ParsedCommand.CallingUserInfo
            OriginalMessage = $ParsedCommand.OriginalMessage.ToHash()
            ConfigurationDirectory = $script:ConfigurationDirectory
            BackendType = $Backend
            PoshBotManifestPath = (Join-Path -Path $script:moduleBase -ChildPath "PoshBot.psd1")
            ModuleQualifiedCommand = $this.ModuleQualifiedCommand
        }
        if ($this.FunctionInfo) {
            $options.Function = $this.FunctionInfo
        } elseIf ($this.CmdletInfo) {
            $options.Function = $this.CmdletInfo
        }

        # Add named/positional parameters
        $options.NamedParameters = $ParsedCommand.NamedParameters
        $options.PositionalParameters = $ParsedCommand.PositionalParameters

        if ($InvokeAsJob) {
            $this.LogDebug("Executing command [$($this.ModuleQualifiedCommand)] as job")
            $fdt = Get-Date -Format FileDateTimeUniversal
            $jobName = "$($this.Name)_$fdt"
            $jobParams = @{
                Name = $jobName
                ScriptBlock = $outer
                ArgumentList = $options
            }
            return (Start-Job @jobParams)
        } else {
            $this.LogDebug("Executing command [$($this.ModuleQualifiedCommand)] in current PS session")
            $errors = $null
            $information = $null
            $warning = $null
            New-Variable -Name opts -Value $options
            $cmdParams = @{
                ScriptBlock = $outer
                ArgumentList = $Options
                ErrorVariable = 'errors'
                InformationVariable = 'information'
                WarningVariable = 'warning'
                Verbose = $true
                NoNewScope = $true
            }
            $output = Invoke-Command @cmdParams
            return @{
                Error = @($errors)
                Information = @($Information)
                Output = $output
                Warning = @($warning)
            }
        }
    }

    [bool]IsAuthorized([string]$UserId, [RoleManager]$RoleManager) {
        $isAuth = $false
        if ($this.AccessFilter.Permissions.Count -gt 0) {
            $perms = $RoleManager.GetUserPermissions($UserId)
            foreach ($perm in $perms) {
                $result = $this.AccessFilter.Authorize($perm.Name)
                if ($result.Authorized) {
                    $this.LogDebug("User [$UserId] authorized to execute command [$($this.Name)] via permission [$($perm.Name)]")
                    $isAuth = $true
                    break
                }
            }
        } else {
            $isAuth = $true
        }

        if ($isAuth) {
            return $true
        } else {
            $this.LogDebug("User [$UserId] not authorized to execute command [$($this.name)]")
            return $false
        }
    }

    [void]Activate() {
        $this.Enabled = $true
        $this.LogDebug("Command [$($this.Name)] activated")
    }

    [void]Deactivate() {
        $this.Enabled = $false
        $this.LogDebug("Command [$($this.Name)] deactivated")
    }

    [void]AddPermission([Permission]$Permission) {
        $this.LogDebug("Adding permission [$($Permission.Name)] to [$($this.Name)]")
        $this.AccessFilter.AddPermission($Permission)
    }

    [void]RemovePermission([Permission]$Permission) {
        $this.LogDebug("Removing permission [$($Permission.Name)] from [$($this.Name)]")
        $this.AccessFilter.RemovePermission($Permission)
    }

    # Search all the triggers for this command and return TRUE if we have a match
    # with the parsed command
    [bool]TriggerMatch([ParsedCommand]$ParsedCommand, [bool]$CommandSearch = $true) {
        $match = $false
        foreach ($trigger in $this.Triggers) {
            switch ($trigger.Type) {
                'Command' {
                    if ($CommandSearch) {
                        # Command tiggers only work with normal messages received from chat network
                        if ($ParsedCommand.OriginalMessage.Type -eq [MessageType]::Message) {
                            if ($trigger.Trigger -eq $ParsedCommand.Command) {
                                $this.LogDebug("Parsed command [$($ParsedCommand.Command)] matched to command trigger [$($trigger.Trigger)] on command [$($this.Name)]")
                                $match = $true
                                break
                            }
                        }
                    }
                }
                'Event' {
                    if ($trigger.MessageType -eq $ParsedCommand.OriginalMessage.Type) {
                        if ($trigger.MessageSubtype -eq $ParsedCommand.OriginalMessage.Subtype) {
                            $this.LogDebug("Parsed command event type [$($ParsedCommand.OriginalMessage.Type.Tostring())`:$($ParsedCommand.OriginalMessage.Subtype.ToString())] matched to command trigger [$($trigger.MessageType.ToString())`:$($trigger.MessageSubtype.ToString())] on command [$($this.Name)]")
                            $match = $true
                            break
                        }
                    }
                }
                'Regex' {
                    if ($ParsedCommand.CommandString -match $trigger.Trigger) {
                        $this.LogDebug("Parsed command string [$($ParsedCommand.CommandString)] matched to regex trigger [$($trigger.Trigger)] on command [$($this.Name)]")
                        $match = $true
                        break
                    }
                }
            }
        }

        return $match
    }
}

class CommandHistory {
    [string]$Id

    # ID of command
    [string]$CommandId

    # ID of caller
    [string]$CallerId

    # Command results
    [CommandResult]$Result

    [ParsedCommand]$ParsedCommand

    # Date/time command was executed
    [datetime]$Time

    CommandHistory([string]$CommandId, [string]$CallerId, [CommandResult]$Result, [ParsedCommand]$ParsedCommand) {
        $this.Id = (New-Guid).ToString() -Replace '-', ''
        $this.CommandId = $CommandId
        $this.CallerId = $CallerId
        $this.Result = $Result
        $this.ParsedCommand = $ParsedCommand
        $this.Time = Get-Date
    }
}
# The Plugin class holds a collection of related commands that came
# from a PowerShell module

# Some custom exceptions dealing with plugins
class PluginException : Exception {
    PluginException() {}
    PluginException([string]$Message) : base($Message) {}
}

class PluginNotFoundException : PluginException {
    PluginNotFoundException() {}
    PluginNotFoundException([string]$Message) : base($Message) {}
}

class PluginDisabled : PluginException {
    PluginDisabled() {}
    PluginDisabled([string]$Message) : base($Message) {}
}

class Plugin : BaseLogger {

    # Unique name for the plugin
    [string]$Name

    # Commands bundled with plugin
    [hashtable]$Commands = @{}

    [version]$Version

    [bool]$Enabled

    [hashtable]$Permissions = @{}

    hidden [string]$_ManifestPath

    Plugin([Logger]$Logger) {
        $this.Name = $this.GetType().Name
        $this.Logger = $Logger
        $this.Enabled = $true
    }

    Plugin([string]$Name, [Logger]$Logger) {
        $this.Name = $Name
        $this.Logger = $Logger
        $this.Enabled = $true
    }

    # Find the command
    [Command]FindCommand([Command]$Command) {
        return $this.Commands.($Command.Name)
    }

    # Add a new command
    [void]AddCommand([Command]$Command) {
        if (-not $this.FindCommand($Command)) {
            $this.LogDebug("Adding command [$($Command.Name)]")
            $this.Commands.Add($Command.Name, $Command)
        }
    }

    # Remove an existing command
    [void]RemoveCommand([Command]$Command) {
        $existingCommand = $this.FindCommand($Command)
        if ($existingCommand) {
            $this.LogDebug("Removing command [$($Command.Name)]")
            $this.Commands.Remove($Command.Name)
        }
    }

    # Activate a command
    [void]ActivateCommand([Command]$Command) {
        $existingCommand = $this.FindCommand($Command)
        if ($existingCommand) {
            $this.LogDebug("Activating command [$($Command.Name)]")
            $existingCommand.Activate()
        }
    }

    # Deactivate a command
    [void]DeactivateCommand([Command]$Command) {
        $existingCommand = $this.FindCommand($Command)
        if ($existingCommand) {
            $this.LogDebug("Deactivating command [$($Command.Name)]")
            $existingCommand.Deactivate()
        }
    }

    [void]AddPermission([Permission]$Permission) {
        if (-not $this.Permissions.ContainsKey($Permission.ToString())) {
            $this.LogDebug("Adding permission [$Permission.ToString()] to plugin [$($this.Name)`:$($this.Version.ToString())]")
            $this.Permissions.Add($Permission.ToString(), $Permission)
        }
    }

    [Permission]GetPermission([string]$Name) {
        return $this.Permissions[$Name]
    }

    [void]RemovePermission([Permission]$Permission) {
        if ($this.Permissions.ContainsKey($Permission.ToString())) {
            $this.LogDebug("Removing permission [$Permission.ToString()] from plugin [$($this.Name)`:$($this.Version.ToString())]")
            $this.Permissions.Remove($Permission.ToString())
        }
    }

    # Activate plugin and all commands
    [void]Activate() {
        $this.LogDebug("Activating plugin [$($this.Name)`:$($this.Version.ToString())]")
        $this.Enabled = $true
        $this.Commands.GetEnumerator() | ForEach-Object {
            $_.Value.Activate()
        }
    }

    # Deactivate plugin and all commands
    [void]Deactivate() {
        $this.LogDebug("Deactivating plugin [$($this.Name)`:$($this.Version.ToString())]")
        $this.Enabled = $false
        $this.Commands.GetEnumerator() | ForEach-Object {
            $_.Value.Deactivate()
        }
    }

    [hashtable]ToHash() {
        $cmdPerms = @{}
        $this.Commands.GetEnumerator() | Foreach-Object {
            $cmdPerms.Add($_.Name, $_.Value.AccessFilter.Permissions.Keys)
        }

        $adhocPerms = New-Object System.Collections.ArrayList
        $this.Permissions.GetEnumerator() | Where-Object {$_.Value.Adhoc -eq $true} | Foreach-Object {
            $adhocPerms.Add($_.Name) > $null
        }
        return @{
            Name = $this.Name
            Version = $this.Version.ToString()
            Enabled = $this.Enabled
            ManifestPath = $this._ManifestPath
            CommandPermissions = $cmdPerms
            AdhocPermissions = $adhocPerms
        }
    }
}

class PluginCommand {
    [Plugin]$Plugin
    [Command]$Command

    PluginCommand([Plugin]$Plugin, [Command]$Command) {
        $this.Plugin = $Plugin
        $this.Command = $Command
    }

    [string]ToString() {
        return "$($this.Plugin.Name):$($this.Command.Name):$($this.Plugin.Version.ToString())"
    }
}

class Approver {
    [string]$Id
    [string]$Name
}

# Represents the state of a currently executing command
class CommandExecutionContext {
    [string]$Id = (New-Guid).ToString().Split('-')[0]
    [bool]$Complete = $false
    [CommandResult]$Result
    [string]$FullyQualifiedCommandName
    [Command]$Command
    [ParsedCommand]$ParsedCommand
    [Message]$Message
    [bool]$IsJob
    [datetime]$Started
    [datetime]$Ended
    [object]$Job
    [ApprovalState]$ApprovalState = [ApprovalState]::AutoApproved
    [Approver]$Approver = [Approver]::new()
    [Response]$Response = [Response]::new()

    [pscustomobject]Summarize() {
        return [pscustomobject]@{
            Id                        = $this.Id
            Complete                  = $this.Complete
            Result                    = $this.Result.Summarize()
            FullyQualifiedCommandName = $this.FullyQualifiedCommandName
            ParsedCommand             = $this.ParsedCommand.Summarize()
            Message                   = $this.Message
            IsJob                     = $this.IsJob
            Started                   = $this.Started.ToUniversalTime().ToString('u')
            Ended                     = $this.Ended.ToUniversalTime().ToString('u')
            ApprovalState             = $this.ApprovalState.ToString()
            Approver                  = $this.Approver
            Response                  = $this.Response.Summarize()
        }
    }

    [string]ToJson() {
        return $this.Summarize() | ConvertTo-Json -Depth 10 -Compress
    }
 }

class MiddlewareHook {
    [string]$Name
    [string]$Path

    MiddlewareHook([string]$Name, [string]$Path) {
        $this.Name = $Name
        $this.Path = $Path
    }

    [CommandExecutionContext] Execute([CommandExecutionContext]$Context, [Bot]$Bot) {
        try {
            $fileContent = Get-Content -Path $this.Path -Raw
            $scriptBlock = [scriptblock]::Create($fileContent)
            $params = @{
                scriptblock  = $scriptBlock
                ArgumentList = @($Context, $Bot)
                ErrorAction  = 'Stop'
            }
            $Context = Invoke-Command @params
        } catch {
            throw $_
        }
        return $Context
    }

    [hashtable]ToHash() {
        return @{
            Name = $this.Name
            Path = $this.Path
        }
    }
}

class MiddlewareConfiguration {

    [object] $PreReceiveHooks   = [ordered]@{}
    [object] $PostReceiveHooks  = [ordered]@{}
    [object] $PreExecuteHooks   = [ordered]@{}
    [object] $PostExecuteHooks  = [ordered]@{}
    [object] $PreResponseHooks  = [ordered]@{}
    [object] $PostResponseHooks = [ordered]@{}

    [void] Add([MiddlewareHook]$Hook, [MiddlewareType]$Type) {
        if (-not $this."$($Type.ToString())Hooks".Contains($Hook.Name)) {
            $this."$($Type.ToString())Hooks".Add($Hook.Name, $Hook) > $null
        }
    }

    [void] Remove([MiddlewareHook]$Hook, [MiddlewareType]$Type) {
        if ($this."$($Type.ToString())Hooks".Contains($Hook.Name)) {
            $this."$($Type.ToString())Hooks".Remove($Hook.Name, $Hook) > $null
        }
    }

    [hashtable]ToHash() {
        $hash = @{}
        foreach ($type in [enum]::GetNames([MiddlewareType])) {
            $hash.Add(
                $type,
                $this."$($type)Hooks".GetEnumerator().foreach({$_.Value.ToHash()})
            )
        }
        return $hash
    }

    static [MiddlewareConfiguration] Serialize([hashtable]$DeserializedObject) {
        $mc = [MiddlewareConfiguration]::new()
        foreach ($type in [enum]::GetNames([MiddlewareType])) {
            $DeserializedObject.$type.GetEnumerator().foreach({
                $hook = [MiddlewareHook]::new($_.Name, $_.Path)
                $mc."$($type)Hooks".Add($hook.Name, $hook) > $null
            })
        }
        return $mc
    }
}

# In charge of executing and tracking progress of commands
class CommandExecutor : BaseLogger {

    [RoleManager]$RoleManager

    hidden [Bot]$_bot

    [int]$HistoryToKeep = 100

    [int]$ExecutedCount = 0

    # Recent history of commands executed
    [System.Collections.ArrayList]$History = (New-Object System.Collections.ArrayList)

    # Plugin commands get executed as PowerShell jobs
    # This is to keep track of those
    hidden [hashtable]$_jobTracker = @{}

    CommandExecutor([RoleManager]$RoleManager, [Logger]$Logger, [Bot]$Bot) {
        $this.RoleManager = $RoleManager
        $this.Logger = $Logger
        $this._bot = $Bot
    }

    # Execute a command
    [void]ExecuteCommand([CommandExecutionContext]$cmdExecContext) {

        # Verify command is not disabled
        if (-not $cmdExecContext.Command.Enabled) {
            $err = [CommandDisabled]::New("Command [$($cmdExecContext.Command.Name)] is disabled")
            $cmdExecContext.Complete = $true
            $cmdExecContext.Ended = (Get-Date).ToUniversalTime()
            $cmdExecContext.Result.Success = $false
            $cmdExecContext.Result.Errors += $err
            $this.LogInfo([LogSeverity]::Error, $err.Message, $err)
            $this.TrackJob($cmdExecContext)
            return
        }

        # Verify that all mandatory parameters have been provided for "command" type bot commands
        # This doesn't apply to commands triggered from regex matches, timers, or events
        if ($cmdExecContext.Command.TriggerType -eq [TriggerType]::Command) {
            if (-not $this.ValidateMandatoryParameters($cmdExecContext.ParsedCommand, $cmdExecContext.Command)) {
                $msg = "Mandatory parameters for [$($cmdExecContext.Command.Name)] not provided.`nUsage:`n"
                foreach ($usage in $cmdExecContext.Command.Usage) {
                    $msg += "    $usage`n"
                }
                $err = [CommandRequirementsNotMet]::New($msg)
                $cmdExecContext.Complete = $true
                $cmdExecContext.Ended = (Get-Date).ToUniversalTime()
                $cmdExecContext.Result.Success = $false
                $cmdExecContext.Result.Errors += $err
                $this.LogInfo([LogSeverity]::Error, $err.Message, $err)
                $this.TrackJob($cmdExecContext)
                return
            }
        }

        # If command is [command] or [regex] trigger types, verify that the caller is authorized to execute it
        if ($cmdExecContext.Command.TriggerType -in @('Command', 'Regex')) {
            $authorized = $cmdExecContext.Command.IsAuthorized($cmdExecContext.Message.From, $this.RoleManager)
        } else {
            $authorized = $true
        }

        if ($authorized) {

            # Check if approval(s) are needed to execute this command
            if ($this.ApprovalNeeded($cmdExecContext)) {
                $cmdExecContext.ApprovalState = [ApprovalState]::Pending
                $this._bot.Backend.AddReaction($cmdExecContext.Message, [ReactionType]::ApprovalNeeded)

                # Put this message in the deferred bucket until it is released by the [!approve] command from an authorized approver
                if (-not $this._bot.DeferredCommandExecutionContexts.ContainsKey($cmdExecContext.id)) {
                    $this._bot.DeferredCommandExecutionContexts.Add($cmdExecContext.id, $cmdExecContext)
                } else {
                    $this.LogInfo([LogSeverity]::Error, "This shouldn't happen, but command execution context [$($cmdExecContext.id)] is already in the deferred bucket")
                }

                $approverGroups = $this.GetApprovalGroups($cmdExecContext) -join ', '
                $prefix = $this._bot.Configuration.CommandPrefix
                $msg = "Approval is needed to run [$($cmdExecContext.ParsedCommand.CommandString)] from someone in the approval group(s) [$approverGroups]."
                $msg += "`nTo approve, say '$($prefix)approve $($cmdExecContext.Id)'."
                $msg += "`nTo deny, say '$($prefix)deny $($cmdExecContext.Id)'."
                $msg += "`nTo list pending approvals, say '$($prefix)pending'."
                $response = [Response]::new($cmdExecContext.Message)
                $response.Data = New-PoshBotCardResponse -Type Warning -Title "Approval Needed for [$($cmdExecContext.ParsedCommand.CommandString)]" -Text $msg
                $this._bot.SendMessage($response)
                return
            } else {

                # If command is [command] or [regex] trigger type, add reaction telling the user that the command is being executed
                # Reactions don't make sense for event triggered commands
                if ($cmdExecContext.Command.TriggerType -in @('Command', 'Regex')) {
                    if ($this._bot.Configuration.AddCommandReactions) {
                        $this._bot.Backend.AddReaction($cmdExecContext.Message, [ReactionType]::Processing)
                    }
                }

                if ($cmdExecContext.Command.AsJob) {
                    $this.LogDebug("Command [$($cmdExecContext.FullyQualifiedCommandName)] will be executed as a job")

                    # Kick off job and add to job tracker
                    $cmdExecContext.IsJob = $true
                    $cmdExecContext.Job = $cmdExecContext.Command.Invoke($cmdExecContext.ParsedCommand, $true,$this._bot.Backend.GetType().Name)
                    $this.LogDebug("Command [$($cmdExecContext.FullyQualifiedCommandName)] executing in job [$($cmdExecContext.Job.Id)]")
                    $cmdExecContext.Complete = $false
                } else {
                    # Run command in current session and get results
                    # This should only be 'builtin' commands
                    try {
                        $cmdExecContext.IsJob = $false
                        $hash = $cmdExecContext.Command.Invoke($cmdExecContext.ParsedCommand, $false,$this._bot.Backend.GetType().Name)
                        $cmdExecContext.Complete = $true
                        $cmdExecContext.Ended = (Get-Date).ToUniversalTime()
                        $cmdExecContext.Result.Errors = $hash.Error
                        $cmdExecContext.Result.Streams.Error = $hash.Error
                        $cmdExecContext.Result.Streams.Information = $hash.Information
                        $cmdExecContext.Result.Streams.Warning = $hash.Warning
                        $cmdExecContext.Result.Output = $hash.Output
                        if ($cmdExecContext.Result.Errors.Count -gt 0) {
                            $cmdExecContext.Result.Success = $false
                        } else {
                            $cmdExecContext.Result.Success = $true
                        }
                        $this.LogVerbose("Command [$($cmdExecContext.FullyQualifiedCommandName)] completed with successful result [$($cmdExecContext.Result.Success)]")
                    } catch {
                        $cmdExecContext.Complete = $true
                        $cmdExecContext.Result.Success = $false
                        $cmdExecContext.Result.Errors = $_.Exception.Message
                        $cmdExecContext.Result.Streams.Error = $_.Exception.Message
                        $this.LogInfo([LogSeverity]::Error, $_.Exception.Message, $_)
                    }
                }
            }
        } else {
            $msg = "Command [$($cmdExecContext.Command.Name)] was not authorized for user [$($cmdExecContext.Message.From)]"
            $err = [CommandNotAuthorized]::New($msg)
            $cmdExecContext.Complete = $true
            $cmdExecContext.Result.Errors += $err
            $cmdExecContext.Result.Success = $false
            $cmdExecContext.Result.Authorized = $false
            $this.LogInfo([LogSeverity]::Error, $err.Message, $err)
            $this.TrackJob($cmdExecContext)
            return
        }

        $this.TrackJob($cmdExecContext)
    }

    # Add the command execution context to the job tracker
    # So the status and results of it can be checked later
    [void]TrackJob([CommandExecutionContext]$CommandContext) {
        if (-not $this._jobTracker.ContainsKey($CommandContext.Id)) {
            $this.LogVerbose("Adding job [$($CommandContext.Id)] to tracker")
            $this._jobTracker.Add($CommandContext.Id, $CommandContext)
        }
    }

    # Receive any completed jobs from the job tracker
    [CommandExecutionContext[]]ReceiveJob() {
        $results = New-Object System.Collections.ArrayList

        if ($this._jobTracker.Count -ge 1) {
            $completedJobs = $this._jobTracker.GetEnumerator() |
                Where-Object {($_.Value.Complete -eq $true) -or
                              ($_.Value.IsJob -and (($_.Value.Job.State -eq 'Completed') -or ($_.Value.Job.State -eq 'Failed')))} |
                Select-Object -ExpandProperty Value

            foreach ($cmdExecContext in $completedJobs) {
                # If the command was executed in a PS job, get the output
                # Builtin commands are NOT executed as jobs so their output
                # was already recorded in the [Result] property in the ExecuteCommand() method
                if ($cmdExecContext.IsJob) {
                    if ($cmdExecContext.Job.State -eq 'Completed') {
                        $this.LogVerbose("Job [$($cmdExecContext.Id)] is complete")
                        $cmdExecContext.Complete = $true
                        $cmdExecContext.Ended = (Get-Date).ToUniversalTime()

                        # Capture all the streams
                        $cmdExecContext.Result.Errors = $cmdExecContext.Job.ChildJobs[0].Error.ReadAll()
                        $cmdExecContext.Result.Streams.Error = $cmdExecContext.Result.Errors
                        $cmdExecContext.Result.Streams.Information = $cmdExecContext.Job.ChildJobs[0].Information.ReadAll()
                        $cmdExecContext.Result.Streams.Verbose = $cmdExecContext.Job.ChildJobs[0].Verbose.ReadAll()
                        $cmdExecContext.Result.Streams.Warning = $cmdExecContext.Job.ChildJobs[0].Warning.ReadAll()
                        $cmdExecContext.Result.Output = $cmdExecContext.Job.ChildJobs[0].Output.ReadAll()

                        # Determine if job had any terminating errors
                        if ($cmdExecContext.Result.Streams.Error.Count -gt 0) {
                            $cmdExecContext.Result.Success = $false
                        } else {
                            $cmdExecContext.Result.Success = $true
                        }

                        $this.LogVerbose("Command [$($cmdExecContext.FullyQualifiedCommandName)] completed with successful result [$($cmdExecContext.Result.Success)]")

                        # Clean up the job
                        Remove-Job -Job $cmdExecContext.Job
                    } elseIf ($cmdExecContext.Job.State -eq 'Failed') {
                        $cmdExecContext.Complete = $true
                        $cmdExecContext.Result.Success = $false
                        $this.LogVerbose("Command [$($cmdExecContext.FullyQualifiedCommandName)] failed")
                    }
                }

                # Send a success, warning, or fail reaction
                if ($cmdExecContext.Command.TriggerType -in @('Command', 'Regex')) {
                    if ($this._bot.Configuration.AddCommandReactions) {
                        if (-not $cmdExecContext.Result.Success) {
                            $reaction = [ReactionType]::Failure
                        } elseIf ($cmdExecContext.Result.Streams.Warning.Count -gt 0) {
                            $reaction = [ReactionType]::Warning
                        } else {
                            $reaction = [ReactionType]::Success
                        }
                        $this._bot.Backend.AddReaction($cmdExecContext.Message, $reaction)
                    }
                }

                # Add to history
                if ($cmdExecContext.Command.KeepHistory) {
                    $this.AddToHistory($cmdExecContext)
                }

                $this.LogVerbose("Removing job [$($cmdExecContext.Id)] from tracker")
                $this._jobTracker.Remove($cmdExecContext.Id)

                # Remove the reaction specifying the command is in process
                if ($cmdExecContext.Command.TriggerType -in @('Command', 'Regex')) {
                    if ($this._bot.Configuration.AddCommandReactions) {
                        $this._bot.Backend.RemoveReaction($cmdExecContext.Message, [ReactionType]::Processing)
                    }
                }

                # Track number of commands executed
                if ($cmdExecContext.Result.Success) {
                    $this.ExecutedCount++
                }

                $cmdExecContext.Result.Duration = ($cmdExecContext.Ended - $cmdExecContext.Started)

                $results.Add($cmdExecContext) > $null
            }
        }

        return $results
    }

    # Add command result to history
    [void]AddToHistory([CommandExecutionContext]$CmdExecContext) {
        if ($this.History.Count -ge $this.HistoryToKeep) {
            $this.History.RemoveAt(0) > $null
        }
        $this.LogDebug("Adding command execution [$($CmdExecContext.Id)] to history")
        $this.History.Add($CmdExecContext)
    }

    # Validate that all mandatory parameters have been provided
    [bool]ValidateMandatoryParameters([ParsedCommand]$ParsedCommand, [Command]$Command) {
        $validated = $false

        if ($Command.FunctionInfo) {
            $parameterSets = $Command.FunctionInfo.ParameterSets
        } else {
            $parameterSets = $Command.CmdletInfo.ParameterSets
        }

        foreach ($parameterSet in $parameterSets) {
            $this.LogDebug("Validating parameters for parameter set [$($parameterSet.Name)]")
            $mandatoryParameters = @($parameterSet.Parameters | Where-Object {$_.IsMandatory -eq $true}).Name
            if ($mandatoryParameters.Count -gt 0) {
                # Remove each provided mandatory parameter from the list
                # so we can find any that will have to be coverd by positional parameters
                $this.LogDebug('Provided named parameters', $ParsedCommand.NamedParameters.Keys)
                foreach ($providedNamedParameter in $ParsedCommand.NamedParameters.Keys ) {
                    $this.LogDebug("Named parameter [$providedNamedParameter] provided")
                    $mandatoryParameters = @($mandatoryParameters | Where-Object {$_ -ne $providedNamedParameter})
                }
                if ($mandatoryParameters.Count -gt 0) {
                    if ($ParsedCommand.PositionalParameters.Count -lt $mandatoryParameters.Count) {
                        $validated = $false
                    } else {
                        $validated = $true
                    }
                } else {
                    $validated = $true
                }
            } else {
                $validated = $true
            }

            $this.LogDebug("Valid parameters for parameterset [$($parameterSet.Name)] - [$($validated.ToString())]")
            if ($validated) {
                break
            }
        }

        return $validated
    }

    # Check if command needs approval by checking against command expressions in the approval configuration
    # if peer approval is needed, always return $true regardless if calling user is in approval group
    [bool]ApprovalNeeded([CommandExecutionContext]$Context) {
        if ($Context.ApprovalState -ne [ApprovalState]::Approved) {
            foreach ($approvalConfig in $this._bot.Configuration.ApprovalConfiguration.Commands) {
                if ($Context.FullyQualifiedCommandName -like $approvalConfig.Expression) {

                    $approvalGroups = $this._bot.RoleManager.GetUserGroups($Context.ParsedCommand.From).Name
                    if (-not $approvalGroups) {
                        $approvalGroups = @()
                    }
                    $compareParams = @{
                        ReferenceObject = $this.GetApprovalGroups($Context)
                        DifferenceObject = $approvalGroups
                        PassThru = $true
                        IncludeEqual = $true
                        ExcludeDifferent = $true
                    }
                    $inApprovalGroup = (Compare-Object @compareParams).Count -gt 0

                    $Context.ApprovalState = [ApprovalState]::Pending
                    $this.LogDebug("Execution context ID [$($Context.Id)] needs approval from group(s) [$(($compareParams.ReferenceObject) -join ', ')]")

                    if ($inApprovalGroup) {
                        if ($approvalConfig.PeerApproval) {
                            $this.LogDebug("Execution context ID [$($Context.Id)] needs peer approval")
                        } else {
                            $this.LogInfo("Peer Approval not needed to execute context ID [$($Context.Id)]")
                        }
                        return $approvalConfig.PeerApproval
                    } else {
                        $this.LogInfo("Approval needed to execute context ID [$($Context.Id)]")
                        return $true
                    }
                }
            }
        }

        return $false
    }

    # Get list of approval groups for a command that needs approval
    [string[]]GetApprovalGroups([CommandExecutionContext]$Context) {
        foreach ($approvalConfig in $this._bot.Configuration.ApprovalConfiguration.Commands) {
            if ($Context.FullyQualifiedCommandName -like $approvalConfig.Expression) {
                return $approvalConfig.ApprovalGroups
            }
        }
        return @()
    }
}

# A scheduled message that the scheduler class will return when the time interval
# has elapsed. The bot will treat this message as though it was returned from the
# chat network like a normal message
class ScheduledMessage {

    [string]$Id = (New-Guid).ToString() -Replace '-', ''

    [TimeInterval]$TimeInterval

    [int]$TimeValue

    [Message]$Message

    [bool]$Enabled = $true

    [bool]$Once = $false

    [double]$IntervalMS

    [int]$TimesExecuted = 0

    [DateTime]$StartAfter = (Get-Date).ToUniversalTime()

    ScheduledMessage([TimeInterval]$Interval, [int]$TimeValue, [Message]$Message, [bool]$Enabled, [DateTime]$StartAfter) {
        $this.Init($Interval, $TimeValue, $Message, $Enabled, $StartAfter)
    }

    ScheduledMessage([TimeInterval]$Interval, [int]$TimeValue, [Message]$Message, [bool]$Enabled) {
        $this.Init($Interval, $TimeValue, $Message, $Enabled, (Get-Date).ToUniversalTime())
    }

    ScheduledMessage([TimeInterval]$Interval, [int]$TimeValue, [Message]$Message, [DateTime]$StartAfter) {
        $this.Init($Interval, $TimeValue, $Message, $true, $StartAfter)
    }

    ScheduledMessage([TimeInterval]$Interval, [int]$TimeValue, [Message]$Message) {
        $this.Init($Interval, $TimeValue, $Message, $true, (Get-Date).ToUniversalTime())
    }

    ScheduledMessage([Message]$Message, [Datetime]$StartAt) {
        $this.Message = $Message
        $this.Enabled = $true
        $this.Once = $true
        $this.StartAfter = $StartAt.ToUniversalTime()
    }

    [void]Init([TimeInterval]$Interval, [int]$TimeValue, [Message]$Message, [bool]$Enabled, [DateTime]$StartAfter) {
        $this.TimeInterval = $Interval
        $this.TimeValue = $TimeValue
        $this.Message = $Message
        $this.Enabled = $Enabled
        $this.StartAfter = $StartAfter.ToUniversalTime()

        switch ($this.TimeInterval) {
            'Days' {
                $this.IntervalMS = ($TimeValue * 86400000)
                break
            }
            'Hours' {
                $this.IntervalMS = ($TimeValue * 3600000)
                break
            }
            'Minutes' {
                $this.IntervalMS = ($TimeValue * 60000)
                break
            }
            'Seconds' {
                $this.IntervalMS = ($TimeValue * 1000)
                break
            }
        }
    }

    [bool]HasElapsed() {
        $now = (Get-Date).ToUniversalTime()
        if ($now -gt $this.StartAfter) {
            $this.TimesExecuted += 1
            return $true
        } else {
            return $false
        }
    }

    [void]Enable() {
        $this.Enabled = $true
    }

    [void]Disable() {
        $this.Enabled = $false
    }

    [void]RecalculateStartAfter() {
        $currentDate = (Get-Date).ToUniversalTime()
        $difference = (New-TimeSpan $this.StartAfter $currentDate)
        $elapsedIntervals = [int][Math]::Ceiling($difference.TotalMilliseconds / $this.IntervalMS)
        #Always move forward at least one interval
        if ($elapsedIntervals -lt 1) {
            $elapsedIntervals = 1
        }
        $this.StartAfter = $this.StartAfter.AddMilliseconds($this.IntervalMS * $elapsedIntervals)
    }

    [hashtable]ToHash() {
        return @{
            Id = $this.Id
            TimeInterval = $this.TimeInterval.ToString()
            TimeValue = $this.TimeValue
            StartAfter = $This.StartAfter.ToUniversalTime()
            Once = $this.Once
            Message = @{
                Id = $this.Message.Id
                Type = $this.Message.Type.ToString()
                Subtype = $this.Message.Subtype.ToString()
                Text = $this.Message.Text
                To = $this.Message.To
                From = $this.Message.From
            }
            Enabled = $this.Enabled
            IntervalMS = $this.IntervalMS
        }
    }
}

class Scheduler : BaseLogger {

    [hashtable]$Schedules = @{}

    hidden [StorageProvider]$_Storage

    Scheduler([StorageProvider]$Storage, [Logger]$Logger) {
        $this._Storage = $Storage
        $this.Logger = $Logger
        $this.Initialize()
    }

    [void]Initialize() {
        $this.LogInfo('Initializing')
        $this.LoadState()
    }

    [void]LoadState() {
        $this.LogVerbose('Loading scheduler state from storage')

        if ($scheduleConfig = $this._Storage.GetConfig('schedules')) {
            foreach($key in $scheduleConfig.Keys) {
                $sched = $scheduleConfig[$key]
                $msg = [Message]::new()
                $msg.Id = $sched.Message.Id
                $msg.Text = $sched.Message.Text
                $msg.To = $sched.Message.To
                $msg.From = $sched.Message.From
                $msg.Type = $sched.Message.Type
                $msg.Subtype = $sched.Message.Subtype
                if ($sched.Once) {
                    $newSchedule = [ScheduledMessage]::new($msg, $sched.StartAfter.ToUniversalTime())
                } else {
                    if (-not [string]::IsNullOrEmpty($sched.StartAfter)) {
                        $newSchedule = [ScheduledMessage]::new($sched.TimeInterval, $sched.TimeValue, $msg, $sched.Enabled, $sched.StartAfter.ToUniversalTime())

                        if ($newSchedule.StartAfter -lt (Get-Date).ToUniversalTime()) {
                            #Prevent reruns of commands initially scheduled at least one interval ago
                            $newSchedule.RecalculateStartAfter()
                        }
                    } else {
                        $newSchedule = [ScheduledMessage]::new($sched.TimeInterval, $sched.TimeValue, $msg, $sched.Enabled, (Get-Date).ToUniversalTime())
                    }
                }

                $newSchedule.Id = $sched.Id
                $this.ScheduleMessage($newSchedule, $false)
            }
            $this.SaveState()
        }
    }

    [void]SaveState() {
        $this.LogVerbose('Saving scheduler state to storage')

        $schedulesToSave = @{}
        foreach ($schedule in $this.Schedules.GetEnumerator()) {
            $schedulesToSave.Add("sched_$($schedule.Name)", $schedule.Value.ToHash())
        }
        $this._Storage.SaveConfig('schedules', $schedulesToSave)
    }

    [void]ScheduleMessage([ScheduledMessage]$ScheduledMessage) {
        $this.ScheduleMessage($ScheduledMessage, $true)
    }

    [void]ScheduleMessage([ScheduledMessage]$ScheduledMessage, [bool]$Save) {
        if (-not $this.Schedules.ContainsKey($ScheduledMessage.Id)) {
            $this.LogInfo("Scheduled message [$($ScheduledMessage.Id)]", $ScheduledMessage)
            $this.Schedules.Add($ScheduledMessage.Id, $ScheduledMessage)
        } else {
            $msg = "Id [$($ScheduledMessage.Id)] is already scheduled"
            $this.LogInfo([LogSeverity]::Error, $msg)
        }
        if ($Save) {
            $this.SaveState()
        }
    }

    [void]RemoveScheduledMessage([string]$Id) {
        if ($this.GetSchedule($Id)) {
            $this.Schedules.Remove($id)
            $this.LogInfo("Scheduled message [$($_.Id)] removed")
            $this.SaveState()
        }
    }

    [ScheduledMessage[]]ListSchedules() {
        $result = $this.Schedules.GetEnumerator() |
            Select-Object -ExpandProperty Value |
            Sort-Object -Property TimeValue -Descending

        return $result
    }

    [Message[]]GetTriggeredMessages() {
        $remove = @()
        $messages = $this.Schedules.GetEnumerator() | Foreach-Object {
            if ($_.Value.HasElapsed()) {
                $this.LogInfo("Timer reached on scheduled command [$($_.Value.Id)]")

                # Check if one time command
                if ($_.Value.Once) {
                    $remove += $_.Value.Id
                } else {
                    $_.Value.RecalculateStartAfter()
                }

                $newMsg = $_.Value.Message.Clone()
                $newMsg.Time = Get-Date
                $newMsg
            }
        }

        # Remove any one time commands that have triggered
        foreach ($id in $remove) {
            $this.RemoveScheduledMessage($id)
        }

        return $messages
    }

    [ScheduledMessage]GetSchedule([string]$Id) {
        if ($msg = $this.Schedules[$id]) {
            return $msg
        } else {
            $msg = "Unknown schedule Id [$Id]"
            $this.LogInfo([LogSeverity]::Warning, $msg)
            return $null
        }
    }

    [ScheduledMessage]SetSchedule([ScheduledMessage]$ScheduledMessage) {
        $existingMessage = $this.GetSchedule($ScheduledMessage.Id)
        $existingMessage.Init($ScheduledMessage.TimeInterval, $ScheduledMessage.TimeValue, $ScheduledMessage.Message, $ScheduledMessage.Enabled, $ScheduledMessage.StartAfter)
        $this.LogInfo("Scheduled message [$($ScheduledMessage.Id)] modified", $existingMessage)

        $this.SaveState()
        return $existingMessage
    }

    [ScheduledMessage]EnableSchedule([string]$Id) {
        if ($msg = $this.GetSchedule($Id)) {
            $this.LogInfo("Enabled scheduled command [$($_.Id)] enabled")
            $msg.Enable()
            $this.SaveState()
            return $msg
        } else {
            return $null
        }
    }

    [ScheduledMessage]DisableSchedule([string]$Id) {
        if ($msg = $this.GetSchedule($Id)) {
            $this.LogInfo("Disabled scheduled command [$($_.Id)] enabled")
            $msg.Disable()
            $this.SaveState()
            return $msg
        } else {
            return $null
        }
    }
}

class ConfigProvidedParameter {
    [PoshBot.FromConfig]$Metadata
    [System.Management.Automation.ParameterMetadata]$Parameter

    ConfigProvidedParameter([PoshBot.FromConfig]$Meta, [System.Management.Automation.ParameterMetadata]$Param) {
        $this.Metadata = $Meta
        $this.Parameter = $param
    }
}

class PluginManager : BaseLogger {

    [hashtable]$Plugins = @{}
    [hashtable]$Commands = @{}
    hidden [string]$_PoshBotModuleDir
    [RoleManager]$RoleManager
    [StorageProvider]$_Storage

    PluginManager([RoleManager]$RoleManager, [StorageProvider]$Storage, [Logger]$Logger, [string]$PoshBotModuleDir) {
        $this.RoleManager = $RoleManager
        $this._Storage = $Storage
        $this.Logger = $Logger
        $this._PoshBotModuleDir = $PoshBotModuleDir
        $this.Initialize()
    }

    # Initialize the plugin manager
    [void]Initialize() {
        $this.LogInfo('Initializing')
        $this.LoadState()
        $this.LoadBuiltinPlugins()
    }

    # Get the list of plugins to load and... wait for it... load them
    [void]LoadState() {
        $this.LogVerbose('Loading plugin state from storage')

        $pluginsToLoad = $this._Storage.GetConfig('plugins')
        if ($pluginsToLoad) {
            foreach ($pluginKey in $pluginsToLoad.Keys) {
                $pluginToLoad = $pluginsToLoad[$pluginKey]

                $pluginVersions = $pluginToLoad.Keys
                foreach ($pluginVersionKey in $pluginVersions) {
                    $pv = $pluginToLoad[$pluginVersionKey]
                    $manifestPath = $pv.ManifestPath
                    $adhocPermissions = $pv.AdhocPermissions
                    $this.CreatePluginFromModuleManifest($pluginKey, $manifestPath, $true, $false)

                    if ($newPlugin = $this.Plugins[$pluginKey]) {
                        # Add adhoc permissions back to plugin (all versions)
                        foreach ($version in $newPlugin.Keys) {
                            $npv = $newPlugin[$version]
                            foreach($permission in $adhocPermissions) {
                                if ($p = $this.RoleManager.GetPermission($permission)) {
                                    $npv.AddPermission($p)
                                }
                            }

                            # Add adhoc permissions back to the plugin commands (all versions)
                            $commandPermissions = $pv.CommandPermissions
                            foreach ($commandName in $commandPermissions.Keys ) {
                                $permissions = $commandPermissions[$commandName]
                                foreach ($permission in $permissions) {
                                    if ($p = $this.RoleManager.GetPermission($permission)) {
                                        $npv.AddPermission($p)
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    # Save the state of currently loaded plugins to storage
    [void]SaveState() {
        $this.LogVerbose('Saving loaded plugin state to storage')

        # Skip saving builtin plugin as it will always be loaded at initialization
        $pluginsToSave = @{}
        foreach($pluginKey in $this.Plugins.Keys | Where-Object {$_ -ne 'Builtin'}) {
            $versions = @{}
            foreach ($versionKey in $this.Plugins[$pluginKey].Keys) {
                $pv = $this.Plugins[$pluginKey][$versionKey]
                $versions.Add($versionKey, $pv.ToHash())
            }
            $pluginsToSave.Add($pluginKey, $versions)
        }
        $this._Storage.SaveConfig('plugins', $pluginsToSave)
    }

    # TODO
    # Given a PowerShell module definition, inspect it for commands etc,
    # create a plugin instance and load the plugin
    [void]InstallPlugin([string]$ManifestPath, [bool]$SaveAfterInstall = $false) {
        if (Test-Path -Path $ManifestPath) {
            $moduleName = (Get-Item -Path $ManifestPath).BaseName
            $this.CreatePluginFromModuleManifest($moduleName, $ManifestPath, $true, $SaveAfterInstall)
        } else {
            $msg = "Module manifest path [$manifestPath] not found"
            $this.LogInfo([LogSeverity]::Warning, $msg)
        }
    }

    # Add a plugin to the bot
    [void]AddPlugin([Plugin]$Plugin, [bool]$SaveAfterInstall = $false) {
        if (-not $this.Plugins.ContainsKey($Plugin.Name)) {
            $this.LogInfo("Attaching plugin [$($Plugin.Name)]")

            $pluginVersion = @{
                ($Plugin.Version).ToString() = $Plugin
            }
            $this.Plugins.Add($Plugin.Name, $pluginVersion)

            # Register the plugins permission set with the role manager
            foreach ($permission in $Plugin.Permissions.GetEnumerator()) {
                $this.LogVerbose("Adding permission [$($permission.Value.ToString())] to Role Manager")
                $this.RoleManager.AddPermission($permission.Value)
            }
        } else {
            if (-not $this.Plugins[$Plugin.Name].ContainsKey($Plugin.Version)) {
                # Install a new plugin version
                $this.LogInfo("Attaching version [$($Plugin.Version)] of plugin [$($Plugin.Name)]")
                $this.Plugins[$Plugin.Name].Add($Plugin.Version.ToString(), $Plugin)

                # Register the plugins permission set with the role manager
                foreach ($permission in $Plugin.Permissions.GetEnumerator()) {
                    $this.LogVerbose("Adding permission [$($permission.Value.ToString())] to Role Manager")
                    $this.RoleManager.AddPermission($permission.Value)
                }
            } else {
                $msg = "Plugin [$($Plugin.Name)] version [$($Plugin.Version)] is already loaded"
                $this.LogInfo([LogSeverity]::Warning, $msg)
                throw [PluginException]::New($msg)
            }
        }

        # # Reload commands and role from all currently loading (and active) plugins
        $this.LoadCommands()

        if ($SaveAfterInstall) {
            $this.SaveState()
        }
    }

    # Remove a plugin from the bot
    [void]RemovePlugin([Plugin]$Plugin) {
        if ($this.Plugins.ContainsKey($Plugin.Name)) {
            $pluginVersions = $this.Plugins[$Plugin.Name]
            if ($pluginVersions.Keys.Count -eq 1) {
                # Remove the permissions for this plugin from the role manaager
                # but only if this is the only version of the plugin loaded
                foreach ($permission in $Plugin.Permissions.GetEnumerator()) {
                    $this.LogVerbose("Removing permission [$($Permission.Value.ToString())]. No longer in use")
                    $this.RoleManager.RemovePermission($Permission.Value)
                }
                $this.LogInfo("Removing plugin [$($Plugin.Name)]")
                $this.Plugins.Remove($Plugin.Name)

                # Unload the PS module
                $moduleSpec = @{
                    ModuleName = $Plugin.Name
                    ModuleVersion = $pluginVersions
                }
                Remove-Module -FullyQualifiedName $moduleSpec -Verbose:$false -Force
            } else {
                if ($pluginVersions.ContainsKey($Plugin.Version)) {
                    $this.LogInfo("Removing plugin [$($Plugin.Name)] version [$($Plugin.Version)]")
                    $pluginVersions.Remove($Plugin.Version)

                    # Unload the PS module
                    $moduleSpec = @{
                        ModuleName = $Plugin.Name
                        ModuleVersion = $Plugin.Version
                    }
                    Remove-Module -FullyQualifiedName $moduleSpec -Verbose:$false -Force
                } else {
                    $msg = "Plugin [$($Plugin.Name)] version [$($Plugin.Version)] is not loaded in bot"
                    $this.LogInfo([LogSeverity]::Warning, $msg)
                    throw [PluginNotFoundException]::New($msg)
                }
            }
        }

        # Reload commands from all currently loading (and active) plugins
        $this.LoadCommands()

        $this.SaveState()
    }

    # Remove a plugin and optionally a specific version from the bot
    # If there is only one version, then remove any permissions defined in the plugin as well
    [void]RemovePlugin([string]$PluginName, [string]$Version) {
        if ($p = $this.Plugins[$PluginName]) {
            if ($pv = $p[$Version]) {
                if ($p.Keys.Count -eq 1) {
                    # Remove the permissions for this plugin from the role manaager
                    # but only if this is the only version of the plugin loaded
                    foreach ($permission in $pv.Permissions.GetEnumerator()) {
                        $this.LogVerbose("Removing permission [$($Permission.Value.ToString())]. No longer in use")
                        $this.RoleManager.RemovePermission($Permission.Value)
                    }
                    $this.LogInfo("Removing plugin [$($pv.Name)]")
                    $this.Plugins.Remove($pv.Name)
                } else {
                    $this.LogInfo("Removing plugin [$($pv.Name)] version [$Version]")
                    $p.Remove($pv.Version.ToString())
                }

                # Unload the PS module
                $unloadModuleParams = @{
                    FullyQualifiedName = @{
                        ModuleName    = $PluginName
                        ModuleVersion = $Version
                    }
                    Verbose = $false
                    Force   = $true
                }
                $this.LogDebug("Unloading module [$PluginName] version [$Version]")
                Remove-Module @unloadModuleParams
            } else {
                $msg = "Plugin [$PluginName] version [$Version] is not loaded in bot"
                $this.LogInfo([LogSeverity]::Warning, $msg)
                throw [PluginNotFoundException]::New($msg)
            }
        } else {
            $msg = "Plugin [$PluginName] is not loaded in bot"
            $this.LogInfo([LogSeverity]::Warning, $msg)
            throw [PluginNotFoundException]::New()
        }

        # Reload commands from all currently loading (and active) plugins
        $this.LoadCommands()

        $this.SaveState()
    }

    # Activate a plugin
    [void]ActivatePlugin([string]$PluginName, [string]$Version) {
        if ($p = $this.Plugins[$PluginName]) {
            if ($pv = $p[$Version]) {
                $this.LogInfo("Activating plugin [$PluginName] version [$Version]")
                $pv.Activate()

                # Reload commands from all currently loading (and active) plugins
                $this.LoadCommands()
                $this.SaveState()
            } else {
                $msg = "Plugin [$PluginName] version [$Version] is not loaded in bot"
                $this.LogInfo([LogSeverity]::Warning, $msg)
                throw [PluginNotFoundException]::New($msg)
            }
        } else {
            $msg = "Plugin [$PluginName] is not loaded in bot"
            $this.LogInfo([LogSeverity]::Warning, $msg)
            throw [PluginNotFoundException]::New()
        }
    }

    # Activate a plugin
    [void]ActivatePlugin([Plugin]$Plugin) {
        $p = $this.Plugins[$Plugin.Name]
        if ($p) {
            if ($pv = $p[$Plugin.Version.ToString()]) {
                $this.LogInfo("Activating plugin [$($Plugin.Name)] version [$($Plugin.Version)]")
                $pv.Activate()
            }
        } else {
            $msg = "Plugin [$($Plugin.Name)] version [$($Plugin.Version)] is not loaded in bot"
            $this.LogInfo([LogSeverity]::Warning, $msg)
            throw [PluginNotFoundException]::New($msg)
        }

        # Reload commands from all currently loading (and active) plugins
        $this.LoadCommands()

        $this.SaveState()
    }

    # Deactivate a plugin
    [void]DeactivatePlugin([Plugin]$Plugin) {
        $p = $this.Plugins[$Plugin.Name]
        if ($p) {
            if ($pv = $p[$Plugin.Version.ToString()]) {
                $this.LogInfo("Deactivating plugin [$($Plugin.Name)] version [$($Plugin.Version)]")
                $pv.Deactivate()
            }
        } else {
            $msg = "Plugin [$($Plugin.Name)] version [$($Plugin.Version)] is not loaded in bot"
            $this.LogInfo([LogSeverity]::Warning, $msg)
            throw [PluginNotFoundException]::New($msg)
        }

        # # Reload commands from all currently loading (and active) plugins
        $this.LoadCommands()

        $this.SaveState()
    }

    # Deactivate a plugin
    [void]DeactivatePlugin([string]$PluginName, [string]$Version) {
        if ($p = $this.Plugins[$PluginName]) {
            if ($pv = $p[$Version]) {
                $this.LogInfo("Deactivating plugin [$PluginName)] version [$Version]")
                $pv.Deactivate()

                # Reload commands from all currently loading (and active) plugins
                $this.LoadCommands()
                $this.SaveState()
            } else {
                $msg = "Plugin [$PluginName] version [$Version] is not loaded in bot"
                $this.LogInfo([LogSeverity]::Warning, $msg)
                throw [PluginNotFoundException]::New($msg)
            }
        } else {
            $msg = "Plugin [$PluginName] is not loaded in bot"
            $this.LogInfo([LogSeverity]::Warning, $msg)
            throw [PluginNotFoundException]::New($msg)
        }
    }

    # Match a parsed command to a command in one of the currently loaded plugins
    [PluginCommand]MatchCommand([ParsedCommand]$ParsedCommand, [bool]$CommandSearch = $true) {

        # Check builtin commands first
        $builtinKey = $this.Plugins['Builtin'].Keys | Select-Object -First 1
        $builtinPlugin = $this.Plugins['Builtin'][$builtinKey]
        foreach ($commandKey in $builtinPlugin.Commands.Keys) {
            $command = $builtinPlugin.Commands[$commandKey]
            if ($command.TriggerMatch($ParsedCommand, $CommandSearch)) {
                $this.LogInfo("Matched parsed command [$($ParsedCommand.Plugin)`:$($ParsedCommand.Command)] to builtin command [Builtin:$commandKey]")
                return [PluginCommand]::new($builtinPlugin, $command)
            }
        }

        # If parsed command is fully qualified with <plugin:command> syntax. Just look in that plugin
        if (($ParsedCommand.Plugin -ne [string]::Empty) -and ($ParsedCommand.Command -ne [string]::Empty)) {
            $plugin = $this.Plugins[$ParsedCommand.Plugin]
            if ($plugin) {
                if ($ParsedCommand.Version) {
                    # User specified a specific version of the plugin so get that one
                    $pluginVersion = $plugin[$ParsedCommand.Version]
                } else {
                    # Just look in the latest version of the plugin.
                    $latestVersionKey = $plugin.Keys | Sort-Object -Descending | Select-Object -First 1
                    $pluginVersion = $plugin[$latestVersionKey]
                }

                if ($pluginVersion) {
                    foreach ($commandKey in $pluginVersion.Commands.Keys) {
                        $command = $pluginVersion.Commands[$commandKey]
                        if ($command.TriggerMatch($ParsedCommand, $CommandSearch)) {
                            $this.LogInfo("Matched parsed command [$($ParsedCommand.Plugin)`:$($ParsedCommand.Command)] to plugin command [$($plugin.Name)`:$commandKey]")
                            return [PluginCommand]::new($pluginVersion, $command)
                        }
                    }
                }

                $this.LogInfo([LogSeverity]::Warning, "Unable to match parsed command [$($ParsedCommand.Plugin)`:$($ParsedCommand.Command)] to a command in plugin [$($plugin.Name)]")
            } else {
                $this.LogInfo([LogSeverity]::Warning, "Unable to match parsed command [$($ParsedCommand.Plugin)`:$($ParsedCommand.Command)] to a plugin command")
                return $null
            }
        } else {
            # Check all regular plugins/commands now
            foreach ($pluginKey in $this.Plugins.Keys) {
                $plugin = $this.Plugins[$pluginKey]
                $pluginVersion = $null
                if ($ParsedCommand.Version) {
                    # User specified a specific version of the plugin so get that one
                    $pluginVersion = $plugin[$ParsedCommand.Version]
                    foreach ($commandKey in $pluginVersion.Commands.Keys) {
                        $command = $pluginVersion.Commands[$commandKey]
                        if ($command.TriggerMatch($ParsedCommand, $CommandSearch)) {
                            $this.LogInfo("Matched parsed command [$($ParsedCommand.Plugin)`:$($ParsedCommand.Command)] to plugin command [$pluginKey`:$commandKey]")
                            return [PluginCommand]::new($pluginVersion, $command)
                        }
                    }
                } else {
                    # Just look in the latest version of the plugin.
                    foreach ($pluginVersionKey in $plugin.Keys | Sort-Object -Descending | Select-Object -First 1) {
                        $pluginVersion = $plugin[$pluginVersionKey]
                        foreach ($commandKey in $pluginVersion.Commands.Keys) {
                            $command = $pluginVersion.Commands[$commandKey]
                            if ($command.TriggerMatch($ParsedCommand, $CommandSearch)) {
                                $this.LogInfo("Matched parsed command [$($ParsedCommand.Plugin)`:$($ParsedCommand.Command)] to plugin command [$pluginKey`:$commandKey]")
                                return [PluginCommand]::new($pluginVersion, $command)
                            }
                        }
                    }
                }
            }
        }

        $this.LogInfo([LogSeverity]::Warning, "Unable to match parsed command [$($ParsedCommand.Plugin)`:$($ParsedCommand.Command)] to a plugin command")
        return $null
    }

    # Load in the available commands from all the loaded plugins
    [void]LoadCommands() {
        $allCommands = New-Object System.Collections.ArrayList
        foreach ($pluginKey in $this.Plugins.Keys) {
            $plugin = $this.Plugins[$pluginKey]

            foreach ($pluginVersionKey in $plugin.Keys | Sort-Object -Descending | Select-Object -First 1) {
                $pluginVersion = $plugin[$pluginVersionKey]
                if ($pluginVersion.Enabled) {
                    foreach ($commandKey in $pluginVersion.Commands.Keys) {
                        $command =  $pluginVersion.Commands[$commandKey]
                        $fullyQualifiedCommandName = "$pluginKey`:$CommandKey`:$pluginVersionKey"
                        $allCommands.Add($fullyQualifiedCommandName)
                        if (-not $this.Commands.ContainsKey($fullyQualifiedCommandName)) {
                            $this.LogVerbose("Loading command [$fullyQualifiedCommandName]")
                            $this.Commands.Add($fullyQualifiedCommandName, $command)
                        }
                    }
                }
            }
        }

        # Remove any commands that are not in any of the loaded (and active) plugins
        $remove = New-Object System.Collections.ArrayList
        foreach($c in $this.Commands.Keys) {
            if (-not $allCommands.Contains($c)) {
                $remove.Add($c)
            }
        }
        $remove | ForEach-Object {
            $this.LogVerbose("Removing command [$_]. Plugin has either been removed or is deactivated.")
            $this.Commands.Remove($_)
        }
    }

    # Create a new plugin from a given module manifest
    [void]CreatePluginFromModuleManifest([string]$ModuleName, [string]$ManifestPath, [bool]$AsJob = $true, [bool]$SaveAfterCreation = $false) {
        $manifest = Import-PowerShellDataFile -Path $ManifestPath -ErrorAction SilentlyContinue
        if ($manifest) {
            $this.LogInfo("Creating new plugin [$ModuleName]")
            $plugin = [Plugin]::new($this.Logger)
            $plugin.Name = $ModuleName
            $plugin._ManifestPath = $ManifestPath
            if ($manifest.ModuleVersion) {
                $plugin.Version = $manifest.ModuleVersion
            } else {
                $plugin.Version = '0.0.0'
            }

            # Load our plugin config
            $pluginConfig = $this.GetPluginConfig($plugin.Name, $plugin.Version)

            # Create new permissions from metadata in the module manifest
            $this.GetPermissionsFromModuleManifest($manifest) | ForEach-Object {
                $_.Plugin = $plugin.Name
                $plugin.AddPermission($_)
            }

            # Add any adhoc permissions that were previously defined back to the plugin
            if ($pluginConfig -and $pluginConfig.AdhocPermissions.Count -gt 0) {
                foreach ($permissionName in $pluginConfig.AdhocPermissions) {
                    if ($p = $this.RoleManager.GetPermission($permissionName)) {
                        $this.LogDebug("Adding adhoc permission [$permissionName] to plugin [$($plugin.Name)]")
                        $plugin.AddPermission($p)
                    } else {
                        $this.LogInfo([LogSeverity]::Warning, "Adhoc permission [$permissionName] not found in Role Manager. Can't attach permission to plugin [$($plugin.Name)]")
                    }
                }
            }

            # Add the plugin so the roles can be registered with the role manager
            $this.AddPlugin($plugin, $SaveAfterCreation)

            # Get exported cmdlets/functions from the module and add them to the plugin
            # Adjust bot command behaviour based on metadata as appropriate
            Import-Module -Name $ManifestPath -Scope Local -Verbose:$false -WarningAction SilentlyContinue -Force
            $moduleCommands = Microsoft.PowerShell.Core\Get-Command -Module $ModuleName -CommandType @('Cmdlet', 'Function') -Verbose:$false
            foreach ($command in $moduleCommands) {

                # Get any command metadata that may be attached to the command
                # via the PoshBot.BotCommand extended attribute
                # NOTE: This only works on functions, not cmdlets
                if ($command.CommandType -eq 'Function') {
                    $metadata = $this.GetCommandMetadata($command)
                } else {
                    $metadata = $null
                }

                $this.LogVerbose("Creating command [$($command.Name)] for new plugin [$($plugin.Name)]")
                $cmd                        = [Command]::new()
                $cmd.Name                   = $command.Name
                $cmd.ModuleQualifiedCommand = "$ModuleName\$($command.Name)"
                $cmd.ManifestPath           = $ManifestPath
                $cmd.Logger                 = $this.Logger
                $cmd.AsJob                  = $AsJob

                if ($command.CommandType -eq 'Function') {
                    $cmd.FunctionInfo = $command
                } elseIf ($command.CommandType -eq 'Cmdlet') {
                    $cmd.CmdletInfo = $command
                }

                # Triggers that will be added to the command
                $triggers = @()

                # Set command properties based on metadata from module
                if ($metadata) {

                    # Set the command name to what is defined in the metadata
                    if ($metadata.CommandName) {
                        $cmd.Name = $metadata.CommandName
                    }

                    # Add any alternate command names as aliases to the command
                    if ($metadata.Aliases) {
                        $metadata.Aliases | Foreach-Object {
                            $cmd.Aliases += $_
                            $triggers += [Trigger]::new([TriggerType]::Command, $_)
                        }
                    }

                    # Add any permissions defined within the plugin to the command
                    if ($metadata.Permissions) {
                        foreach ($item in $metadata.Permissions) {
                            $fqPermission = "$($plugin.Name):$($item)"
                            if ($p = $plugin.GetPermission($fqPermission)) {
                                $cmd.AddPermission($p)
                            } else {
                                $this.LogInfo([LogSeverity]::Warning, "Permission [$fqPermission] is not defined in the plugin module manifest. Command will not be added to plugin.")
                                continue
                            }
                        }
                    }

                    # Add any adhoc permissions that we may have been added in the past
                    # that is stored in our plugin configuration
                    if ($pluginConfig) {
                        foreach ($permissionName in $pluginConfig.AdhocPermissions) {
                            if ($p = $this.RoleManager.GetPermission($permissionName)) {
                                $this.LogDebug("Adding adhoc permission [$permissionName] to command [$($plugin.Name):$($cmd.name)]")
                                $cmd.AddPermission($p)
                            } else {
                                $this.LogInfo([LogSeverity]::Warning, "Adhoc permission [$permissionName] not found in Role Manager. Can't attach permission to command [$($plugin.Name):$($cmd.name)]")
                            }
                        }
                    }

                    $cmd.KeepHistory = $metadata.KeepHistory    # Default is $true
                    $cmd.HideFromHelp = $metadata.HideFromHelp  # Default is $false

                    # Set the trigger type to something other than 'Command'
                    if ($metadata.TriggerType) {
                        switch ($metadata.TriggerType) {
                            'Command' {
                                $cmd.TriggerType = [TriggerType]::Command
                                $cmd.Triggers += [Trigger]::new([TriggerType]::Command, $cmd.Name)

                                # Add any alternate command names as aliases to the command
                                if ($metadata.Aliases) {
                                    $metadata.Aliases | Foreach-Object {
                                        $cmd.Aliases += $_
                                        $triggers += [Trigger]::new([TriggerType]::Command, $_)
                                    }
                                }
                            }
                            'Event' {
                                $cmd.TriggerType = [TriggerType]::Event
                                $t = [Trigger]::new([TriggerType]::Event, $command.Name)

                                # The message type/subtype the command is intended to respond to
                                if ($metadata.MessageType) {
                                    $t.MessageType = $metadata.MessageType
                                }
                                if ($metadata.MessageSubtype) {
                                    $t.MessageSubtype = $metadata.MessageSubtype
                                }
                                $triggers += $t
                            }
                            'Regex' {
                                $cmd.TriggerType = [TriggerType]::Regex
                                $t = [Trigger]::new([TriggerType]::Regex, $command.Name)
                                $t.Trigger = $metadata.Regex
                                $triggers += $t
                            }
                        }
                    } else {
                        $triggers += [Trigger]::new([TriggerType]::Command, $cmd.Name)
                    }
                } else {
                    # No metadata defined so set the command name to the module function name
                    $cmd.Name = $command.Name
                    $triggers += [Trigger]::new([TriggerType]::Command, $cmd.Name)
                }

                # Get the command help so we can pull information from it
                # to construct the bot command
                $cmdHelp = Get-Help -Name $cmd.ModuleQualifiedCommand -ErrorAction SilentlyContinue
                if ($cmdHelp) {
                    $cmd.Description = $cmdHelp.Synopsis.Trim()
                }

                # Set the command usage differently for [Command] and [Regex] trigger types
                if ($cmd.TriggerType -eq [TriggerType]::Command) {
                    # Remove unneeded parameters from command syntax
                    if ($cmdHelp) {
                        $helpSyntax = ($cmdHelp.syntax | Out-String).Trim() -split "`n" | Where-Object {$_ -ne "`r"}
                        $helpSyntax = $helpSyntax -replace '\[\<CommonParameters\>\]', ''
                        $helpSyntax = $helpSyntax -replace '-Bot \<Object\> ', ''
                        $helpSyntax = $helpSyntax -replace '\[-Bot\] \<Object\> ', '['

                        # Replace the function name in the help syntax with
                        # what PoshBot will call the command
                        $helpSyntax = foreach ($item in $helpSyntax) {
                            $item -replace $command.Name, $cmd.Name
                        }
                        $cmd.Usage = $helpSyntax.ToLower().Trim()
                    } else {
                        $this.LogInfo([LogSeverity]::Warning, "Unable to parse help for command [$($command.Name)]")
                        $cmd.Usage = 'ERROR: Unable to parse command help'
                    }
                } elseIf ($cmd.TriggerType -eq [TriggerType]::Regex) {
                    $cmd.Usage = @($triggers | Select-Object -Expand Trigger) -join "`n"
                }

                # Add triggers based on command type and metadata
                $cmd.Triggers += $triggers

                $plugin.AddCommand($cmd)
            }

            # If the plugin was previously disabled in our plugin configuration, make sure it still is
            if ($pluginConfig -and (-not $pluginConfig.Enabled)) {
                $plugin.Deactivate()
            }

            $this.LoadCommands()

            if ($SaveAfterCreation) {
                $this.SaveState()
            }
        } else {
            $msg = "Unable to load module manifest [$ManifestPath]"
            $this.LogInfo([LogSeverity]::Error, $msg)
            Write-Error -Message $msg
        }
    }

    # Get the [Poshbot.BotComamnd()] attribute from the function if it exists
    [PoshBot.BotCommand]GetCommandMetadata([System.Management.Automation.FunctionInfo]$Command) {
        $attrs = $Command.ScriptBlock.Attributes
        $botCmdAttr = $attrs | ForEach-Object {
            if ($_.GetType().ToString() -eq 'PoshBot.BotCommand') {
                $_
            }
        }

        if ($botCmdAttr) {
            $this.LogDebug("Command [$($Command.Name)] has metadata defined")
        } else {
            $this.LogDebug("No metadata defined for command [$($Command.Name)]")
        }

        return $botCmdAttr
    }

    # Inspect the module manifest and return any permissions defined
    [Permission[]]GetPermissionsFromModuleManifest($Manifest) {
        $permissions = New-Object System.Collections.ArrayList
        foreach ($permission in $Manifest.PrivateData.Permissions) {
            if ($permission -is [string]) {
                $p = [Permission]::new($Permission)
                $permissions.Add($p)
            } elseIf ($permission -is [hashtable]) {
                $p = [Permission]::new($permission.Name)
                if ($permission.Description) {
                    $p.Description = $permission.Description
                }
                $permissions.Add($p)
            }
        }

        if ($permissions.Count -gt 0) {
            $this.LogDebug("Permissions defined in module manifest", $permissions)
        } else {
            $this.LogDebug('No permissions defined in module manifest')
        }

        return $permissions
    }

    # Load in the built in plugins
    # These will be marked so that they DON't execute in a PowerShell job
    # as they need access to the bot internals
    [void]LoadBuiltinPlugins() {
        $this.LogInfo('Loading builtin plugins')
        $builtinPlugin = Get-Item -Path "$($this._PoshBotModuleDir)/Plugins/Builtin"
        $moduleName = $builtinPlugin.BaseName
        $manifestPath = Join-Path -Path $builtinPlugin.FullName -ChildPath "$moduleName.psd1"
        $this.CreatePluginFromModuleManifest($moduleName, $manifestPath, $false, $false)
    }

    [hashtable]GetPluginConfig([string]$PluginName, [string]$Version) {
        $config = @{}
        if ($pluginConfig = $this._Storage.GetConfig('plugins')) {
            if ($thisPluginConfig = $pluginConfig[$PluginName]) {
                if (-not [string]::IsNullOrEmpty($Version)) {
                    if ($thisPluginConfig.ContainsKey($Version)) {
                        $pluginVersion = $Version
                    } else {
                        $this.LogDebug([LogSeverity]::Warning, "Plugin [$PluginName`:$Version] not defined in plugins.psd1")
                        return $null
                    }
                } else {
                    $pluginVersion = @($thisPluginConfig.Keys | Sort-Object -Descending)[0]
                }

                $pv = $thisPluginConfig[$pluginVersion]
                return $pv
            } else {
                $this.LogDebug([LogSeverity]::Warning, "Plugin [$PluginName] not defined in plugins.psd1")
                return $null
            }
        } else {
            $this.LogDebug([LogSeverity]::Warning, "No plugin configuration defined in storage")
            return $null
        }
    }
}

# This class holds the bare minimum information necesary to establish a connection to a chat network.
# Specific implementations MAY extend this class to provide more properties
class ConnectionConfig {

    [string]$Endpoint

    [pscredential]$Credential

    ConnectionConfig() {}

    ConnectionConfig([string]$Endpoint, [pscredential]$Credential) {
        $this.Endpoint = $Endpoint
        $this.Credential = $Credential
    }
}

# This class represents the connection to a backend Chat network

class Connection : BaseLogger {
    [ConnectionConfig]$Config
    [ConnectionStatus]$Status = [ConnectionStatus]::Disconnected

    [void]Connect() {}

    [void]Disconnect() {}
}

# This generic Backend class provides the base scaffolding to represent a chat network
class Backend : BaseLogger {

    [string]$Name

    [string]$BotId

    # Connection information for the chat network
    [Connection]$Connection

    [hashtable]$Users = @{}

    [hashtable]$Rooms = @{}

    [System.Collections.ArrayList]$IgnoredMessageTypes = (New-Object System.Collections.ArrayList)

    [bool]$LazyLoadUsers = $false

    Backend() {}

    # Send a message
    [void]SendMessage([Response]$Response) {
        # Must be extended by the specific Backend implementation
        throw 'Implement me!'
    }

    # Add a reaction to an existing chat message
    [void]AddReaction([Message]$Message, [ReactionType]$Type, [string]$Reaction) {
        # Must be extended by the specific Backend implementation
        throw 'Implement me!'
    }

    [void]AddReaction([Message]$Message, [ReactionType]$Type) {
        $this.AddReaction($Message, $Type, [string]::Empty)
    }

    # Add a reaction to an existing chat message
    [void]RemoveReaction([Message]$Message, [ReactionType]$Type, [string]$Reaction) {
        # Must be extended by the specific Backend implementation
        throw 'Implement me!'
    }

    [void]RemoveReaction([Message]$Message, [ReactionType]$Type) {
        $this.RemoveReaction($Message, $Type, [string]::Empty)
    }

    # Receive a message
    [Message[]]ReceiveMessage() {
        # Must be extended by the specific Backend implementation
        throw 'Implement me!'
    }

    # Send a ping on the chat network
    [void]Ping() {
        # Only implement this method to send a message back
        # to the chat network to keep the connection open
    }

    # Get a user by their Id
    [Person]GetUser([string]$UserId) {
        # Must be extended by the specific Backend implementation
        throw 'Implement me!'
    }

    # Connect to the chat network
    [void]Connect() {
        $this.Connection.Connect()
    }

    # Disconnect from the chat network
    [void]Disconnect() {
        $this.Connection.Disconnect()
    }

    # Populate the list of users on the chat network
    [void]LoadUsers() {
        # Must be extended by the specific Backend implementation
        throw 'Implement me!'
    }

    # Populate the list of channel or rooms on the chat network
    [void]LoadRooms() {
        # Must be extended by the specific Backend implementation
        throw 'Implement me!'
    }

    # Get the bot identity Id
    [string]GetBotIdentity() {
        # Must be extended by the specific Backend implementation
        throw 'Implement me!'
    }

    # Resolve a user name to user id
    [string]UsernameToUserId([string]$Username) {
        # Must be extended by the specific Backend implementation
        throw 'Implement me!'
    }

    # Resolve a user ID to a username/nickname
    [string]UserIdToUsername([string]$UserId) {
        # Must be extended by the specific Backend implementation
        throw 'Implement me!'
    }

    [hashtable]GetUserInfo([string]$UserId) {
        # Must be extended by the specific Backend implementation
        throw 'Implement me!'
    }

    [string]ChannelIdToName([string]$ChannelId) {
        # Must be extended by the specific Backend implementation
        throw 'Implement me!'
    }
}

class ApprovalCommandConfiguration {
    [string]$Expression
    [System.Collections.ArrayList]$ApprovalGroups
    [bool]$PeerApproval

    ApprovalCommandConfiguration() {
        $this.Expression = [string]::Empty
        $this.ApprovalGroups = New-Object -TypeName System.Collections.ArrayList
        $this.PeerApproval = $true
    }

    [hashtable]ToHash() {
        return @{
            Expression   = $this.Expression
            Groups       = $this.ApprovalGroups
            PeerApproval = $this.PeerApproval
        }
    }

    static [ApprovalCommandConfiguration] Serialize([PSObject]$DeserializedObject) {
        $acc = [ApprovalCommandConfiguration]::new()
        $acc.Expression     = $DeserializedObject.Expression
        $acc.ApprovalGroups = $DeserializedObject.ApprovalGroups
        $acc.PeerApproval   = $DeserializedObject.PeerApproval

        return $acc
    }
}

class ApprovalConfiguration {
    [int]$ExpireMinutes
    [System.Collections.ArrayList]$Commands

    ApprovalConfiguration() {
        $this.ExpireMinutes = 30
        $this.Commands = New-Object -TypeName System.Collections.ArrayList
    }

    [hashtable]ToHash() {
        $hash = @{
            ExpireMinutes = $this.ExpireMinutes
        }
        $cmds = New-Object -TypeName System.Collections.ArrayList
        $this.Commands | Foreach-Object {
            $cmds.Add($_.ToHash()) > $null
        }
        $hash.Commands = $cmds

        return $hash
    }

    static [ApprovalConfiguration] Serialize([hashtable]$DeserializedObject) {
        $ac = [ApprovalConfiguration]::new()
        $ac.ExpireMinutes = $DeserializedObject.ExpireMinutes
        $DeserializedObject.Commands.foreach({
            $ac.Commands.Add(
                [ApprovalCommandConfiguration]::Serialize($_)
            ) > $null
        })

        return $ac
    }
}

class ChannelRule {
    [string]$Channel
    [string[]]$IncludeCommands
    [string[]]$ExcludeCommands

    ChannelRule() {
        $this.Channel = '*'
        $this.IncludeCommands = @('*')
        $this.ExcludeCommands = @()
    }

    ChannelRule([string]$Channel, [string[]]$IncludeCommands, [string]$ExcludeCommands) {
        $this.Channel = $Channel
        $this.IncludeCommands = $IncludeCommands
        $this.ExcludeCommands = $ExcludeCommands
    }

    [hashtable]ToHash() {
        return @{
            Channel         = $this.Channel
            IncludeCommands = $this.IncludeCommands
            ExcludeCommands = $this.ExcludeCommands
        }
    }

    static [ChannelRule] Serialize([hashtable]$DeserializedObject) {
        $cr = [ChannelRule]::new()
        $cr.Channel = $DeserializedObject.Channel
        $cr.IncludeCommands = $DeserializedObject.IncludeCommands
        $cr.ExcludeCommands = $DeserializedObject.ExcludeCommands

        return $cr
    }
}

class BotConfiguration {

    [string]$Name = 'PoshBot'

    [string]$ConfigurationDirectory = $script:defaultPoshBotDir

    [string]$LogDirectory = $script:defaultPoshBotDir

    [string]$PluginDirectory = $script:defaultPoshBotDir

    [string[]]$PluginRepository = @('PSGallery')

    [string[]]$ModuleManifestsToLoad = @()

    [LogLevel]$LogLevel = [LogLevel]::Verbose

    [int]$MaxLogSizeMB = 10

    [int]$MaxLogsToKeep = 5

    [bool]$LogCommandHistory = $true

    [int]$CommandHistoryMaxLogSizeMB = 10

    [int]$CommandHistoryMaxLogsToKeep = 5

    [hashtable]$BackendConfiguration = @{}

    [hashtable]$PluginConfiguration = @{}

    [string[]]$BotAdmins = @()

    [char]$CommandPrefix = '!'

    [string[]]$AlternateCommandPrefixes = @('poshbot')

    [char[]]$AlternateCommandPrefixSeperators = @(':', ',', ';')

    [string[]]$SendCommandResponseToPrivate = @()

    [bool]$MuteUnknownCommand = $false

    [bool]$AddCommandReactions = $true

    [bool]$DisallowDMs = $false

    [int]$FormatEnumerationLimitOverride = -1

    [ChannelRule[]]$ChannelRules = @([ChannelRule]::new())

    [ApprovalConfiguration]$ApprovalConfiguration = [ApprovalConfiguration]::new()

    [MiddlewareConfiguration]$MiddlewareConfiguration = [MiddlewareConfiguration]::new()

    [BotConfiguration] SerializeInstance([PSObject]$DeserializedObject) {
        return [BotConfiguration]::Serialize($DeserializedObject)
    }

    [hashtable] ToHash() {
        $propertyNames = $this | Get-Member -MemberType Property | Select-Object -ExpandProperty Name
        $hash = @{}

        foreach ($property in $propertyNames) {
            if ($this.$property | Get-Member -MemberType Method -Name ToHash) {
                $hash.$property = $this.$property.ToHash()
            } else {
                $hash.$property = $this.$property
            }
        }

        return $hash
    }

    [BotConfiguration] Serialize([hashtable]$Hash) {

        $propertyNames = $this | Get-Member -MemberType Property | Select-Object -ExpandProperty Name

        $bc = [BotConfiguration]::new()

        foreach ($key in $Hash.keys) {
            if ($key -in $propertyNames) {

                $bc.Name                             = $hash.Name
                $bc.ConfigurationDirectory           = $hash.ConfigurationDirectory
                $bc.LogDirectory                     = $hash.LogDirectory
                $bc.PluginDirectory                  = $hash.PluginDirectory
                $bc.PluginRepository                 = $hash.PluginRepository
                $bc.ModuleManifestsToLoad            = $hash.ModuleManifestsToLoad
                $bc.LogLevel                         = $hash.LogLevel
                $bc.MaxLogSizeMB                     = $hash.MaxLogSizeMB
                $bc.MaxLogsToKeep                    = $hash.MaxLogsToKeep
                $bc.LogCommandHistory                = $hash.LogCommandHistory
                $bc.CommandHistoryMaxLogSizeMB       = $hash.CommandHistoryMaxLogSizeMB
                $bc.CommandHistoryMaxLogsToKeep      = $hash.CommandHistoryMaxLogsToKeep
                $bc.BackendConfiguration             = $hash.BackendConfiguration
                $bc.PluginConfiguration              = $hash.PluginConfiguration
                $bc.BotAdmins                        = $hash.BotAdmins
                $bc.CommandPrefix                    = $hash.CommandPrefix
                $bc.AlternateCommandPrefixes         = $hash.AlternateCommandPrefixes
                $bc.AlternateCommandPrefixSeperators = $hash.AlternateCommandPrefixSeperators
                $bc.SendCommandResponseToPrivate     = $hash.SendCommandResponseToPrivate
                $bc.MuteUnknownCommand               = $hash.MuteUnknownCommand
                $bc.AddCommandReactions              = $hash.AddCommandReactions
                $bc.DisallowDMs                      = $hash.DisallowDMs
                $bc.FormatEnumerationLimitOverride   = $hash.FormatEnumerationLimitOverride
                $bc.ChannelRules                     = $hash.ChannelRules.ForEach({[ChannelRule]::Serialize($_)})
                $bc.ApprovalConfiguration            = [ApprovalConfiguration]::Serialize($hash.ApprovalConfiguration)
                $bc.MiddlewareConfiguration          = [MiddlewareConfiguration]::Serialize($hash.MiddlewareConfiguration)
            } else {
                throw "Hash key [$key] is not a property in BotConfiguration"
            }
        }

        return $bc
    }

    static [BotConfiguration] Serialize([PSObject]$DeserializedObject) {
        $bc = [BotConfiguration]::new()
        $bc.Name                             = $DeserializedObject.Name
        $bc.ConfigurationDirectory           = $DeserializedObject.ConfigurationDirectory
        $bc.LogDirectory                     = $DeserializedObject.LogDirectory
        $bc.PluginDirectory                  = $DeserializedObject.PluginDirectory
        $bc.PluginRepository                 = $DeserializedObject.PluginRepository
        $bc.ModuleManifestsToLoad            = $DeserializedObject.ModuleManifestsToLoad
        $bc.LogLevel                         = $DeserializedObject.LogLevel
        $bc.MaxLogSizeMB                     = $DeserializedObject.MaxLogSizeMB
        $bc.MaxLogsToKeep                    = $DeserializedObject.MaxLogsToKeep
        $bc.LogCommandHistory                = $DeserializedObject.LogCommandHistory
        $bc.CommandHistoryMaxLogSizeMB       = $DeserializedObject.CommandHistoryMaxLogSizeMB
        $bc.CommandHistoryMaxLogsToKeep      = $DeserializedObject.CommandHistoryMaxLogsToKeep
        $bc.BackendConfiguration             = $DeserializedObject.BackendConfiguration
        $bc.PluginConfiguration              = $DeserializedObject.PluginConfiguration
        $bc.BotAdmins                        = $DeserializedObject.BotAdmins
        $bc.CommandPrefix                    = $DeserializedObject.CommandPrefix
        $bc.AlternateCommandPrefixes         = $DeserializedObject.AlternateCommandPrefixes
        $bc.AlternateCommandPrefixSeperators = $DeserializedObject.AlternateCommandPrefixSeperators
        $bc.SendCommandResponseToPrivate     = $DeserializedObject.SendCommandResponseToPrivate
        $bc.MuteUnknownCommand               = $DeserializedObject.MuteUnknownCommand
        $bc.AddCommandReactions              = $DeserializedObject.AddCommandReactions
        $bc.DisallowDMs                      = $DeserializedObject.DisallowDMs
        $bc.FormatEnumerationLimitOverride   = $DeserializedObject.FormatEnumerationLimitOverride
        $bc.ChannelRules                     = $DeserializedObject.ChannelRules.Foreach({[ChannelRule]::Serialize($_)})
        $bc.ApprovalConfiguration            = [ApprovalConfiguration]::Serialize($DeserializedObject.ApprovalConfiguration)
        $bc.MiddlewareConfiguration          = [MiddlewareConfiguration]::Serialize($DeserializedObject.MiddlewareConfiguration)

        return $bc
    }
}

class Bot : BaseLogger {

    # Friendly name for the bot
    [string]$Name

    # The backend system for this bot (Slack, HipChat, etc)
    [Backend]$Backend

    hidden [string]$_PoshBotDir

    [StorageProvider]$Storage

    [PluginManager]$PluginManager

    [RoleManager]$RoleManager

    [CommandExecutor]$Executor

    [Scheduler]$Scheduler

    # Queue of messages from the chat network to process
    [System.Collections.Queue]$MessageQueue = (New-Object System.Collections.Queue)

    [hashtable]$DeferredCommandExecutionContexts = @{}

    [System.Collections.Queue]$ProcessedDeferredContextQueue = (New-Object System.Collections.Queue)

    [BotConfiguration]$Configuration

    hidden [System.Diagnostics.Stopwatch]$_Stopwatch

    hidden [System.Collections.Arraylist] $_PossibleCommandPrefixes = (New-Object System.Collections.ArrayList)

    hidden [MiddlewareConfiguration] $_Middleware

    hidden [bool]$LazyLoadComplete = $false

    Bot([Backend]$Backend, [string]$PoshBotDir, [BotConfiguration]$Config)
        : base($Config.LogDirectory, $Config.LogLevel, $Config.MaxLogSizeMB, $Config.MaxLogsToKeep) {

        $this.Name = $config.Name
        $this.Backend = $Backend
        $this._PoshBotDir = $PoshBotDir
        $this.Storage = [StorageProvider]::new($Config.ConfigurationDirectory, $this.Logger)
        $this.Initialize($Config)
    }

    Bot([string]$Name, [Backend]$Backend, [string]$PoshBotDir, [string]$ConfigPath)
        : base($Config.LogDirectory, $Config.LogLevel, $Config.MaxLogSizeMB, $Config.MaxLogsToKeep) {

        $this.Name = $Name
        $this.Backend = $Backend
        $this._PoshBotDir = $PoshBotDir
        $this.Storage = [StorageProvider]::new((Split-Path -Path $ConfigPath -Parent), $this.Logger)
        $config = Get-PoshBotConfiguration -Path $ConfigPath
        $this.Initialize($config)
    }

    [void]Initialize([BotConfiguration]$Config) {
        $this.LogInfo('Initializing bot')

        # Attach the logger to the backend
        $this.Backend.Logger = $this.Logger
        $this.Backend.Connection.Logger = $this.Logger

        if ($null -eq $Config) {
            $this.LoadConfiguration()
        } else {
            $this.Configuration = $Config
        }
        $this.RoleManager = [RoleManager]::new($this.Backend, $this.Storage, $this.Logger)
        $this.PluginManager = [PluginManager]::new($this.RoleManager, $this.Storage, $this.Logger, $this._PoshBotDir)
        $this.Executor = [CommandExecutor]::new($this.RoleManager, $this.Logger, $this)
        $this.Scheduler = [Scheduler]::new($this.Storage, $this.Logger)
        $this.GenerateCommandPrefixList()

        # Register middleware hooks
        $this._Middleware = $Config.MiddlewareConfiguration

        # Ugly hack alert!
        # Store the ConfigurationDirectory property in a script level variable
        # so the command class as access to it.
        $script:ConfigurationDirectory = $this.Configuration.ConfigurationDirectory

        # Add internal plugin directory and user-defined plugin directory to PSModulePath
        if (-not [string]::IsNullOrEmpty($this.Configuration.PluginDirectory)) {
            $internalPluginDir = Join-Path -Path $this._PoshBotDir -ChildPath 'Plugins'
            $modulePaths = $env:PSModulePath.Split($script:pathSeperator)
            if ($modulePaths -notcontains $internalPluginDir) {
                $env:PSModulePath = $internalPluginDir + $script:pathSeperator + $env:PSModulePath
            }
            if ($modulePaths -notcontains $this.Configuration.PluginDirectory) {
                $env:PSModulePath = $this.Configuration.PluginDirectory + $script:pathSeperator + $env:PSModulePath
            }
        }

        # Set PS repository to trusted
        foreach ($repo in $this.Configuration.PluginRepository) {
            if ($r = Get-PSRepository -Name $repo -Verbose:$false -ErrorAction SilentlyContinue) {
                if ($r.InstallationPolicy -ne 'Trusted') {
                    $this.LogVerbose("Setting PowerShell repository [$repo] to [Trusted]")
                    Set-PSRepository -Name $repo -Verbose:$false -InstallationPolicy Trusted
                }
            } else {
                $this.LogVerbose([LogSeverity]::Warning, "PowerShell repository [$repo)] is not defined on the system")
            }
        }

        # Load in plugins listed in configuration
        if ($this.Configuration.ModuleManifestsToLoad.Count -gt 0) {
            $this.LogInfo('Loading in plugins from configuration')
            foreach ($manifestPath in $this.Configuration.ModuleManifestsToLoad) {
                if (Test-Path -Path $manifestPath) {
                    $this.PluginManager.InstallPlugin($manifestPath, $false)
                } else {
                    $this.LogInfo([LogSeverity]::Warning, "Could not find manifest at [$manifestPath]")
                }
            }
        }
    }

    [void]LoadConfiguration() {
        $botConfig = $this.Storage.GetConfig($this.Name)
        if ($botConfig) {
            $this.Configuration = $botConfig
        } else {
            $this.Configuration = [BotConfiguration]::new()
            $hash = @{}
            $this.Configuration | Get-Member -MemberType Property | ForEach-Object {
                $hash.Add($_.Name, $this.Configuration.($_.Name))
            }
            $this.Storage.SaveConfig('Bot', $hash)
        }
    }

    # Start the bot
    [void]Start() {
        $this._Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $this.LogInfo('Start your engines')
        $OldFormatEnumerationLimit = $global:FormatEnumerationLimit
        if($this.Configuration.FormatEnumerationLimitOverride -is [int]) {
            $global:FormatEnumerationLimit = $this.Configuration.FormatEnumerationLimitOverride
            $this.LogInfo("Setting global FormatEnumerationLimit to [$($this.Configuration.FormatEnumerationLimitOverride)]")
        }
        try {
            $this.Connect()

            # Start the loop to receive and process messages from the backend
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $this.LogInfo('Beginning message processing loop')
            while ($this.Backend.Connection.Connected) {

                # Receive message and add to queue
                $this.ReceiveMessage()

                # Get 0 or more scheduled jobs that need to be executed
                # and add to message queue
                $this.ProcessScheduledMessages()

                # Determine if any contexts that are deferred are expired
                $this.ProcessDeferredContexts()

                # Determine if message is for bot and handle as necessary
                $this.ProcessMessageQueue()

                # Receive any completed jobs and process them
                $this.ProcessCompletedJobs()

                Start-Sleep -Milliseconds 100

                # Send a ping every 5 seconds
                if ($sw.Elapsed.TotalSeconds -gt 5) {
                    $this.Backend.Ping()
                    $sw.Reset()
                }
            }
        } catch {
            $this.LogInfo([LogSeverity]::Error, "$($_.Exception.Message)", [ExceptionFormatter]::Summarize($_))
        } finally {
            $global:FormatEnumerationLimit = $OldFormatEnumerationLimit
            $this.Disconnect()
        }
    }

    # Connect the bot to the chat network
    [void]Connect() {
        $this.LogVerbose('Connecting to backend chat network')
        $this.Backend.Connect()

        # If the backend is not configured to lazy load
        # then add admins now
        if (-not $this.Backend.LazyLoadUsers) {
            $this._LoadAdmins()
        }
    }

    # Disconnect the bot from the chat network
    [void]Disconnect() {
        $this.LogVerbose('Disconnecting from backend chat network')
        $this.Backend.Disconnect()
    }

    # Receive messages from the backend chat network
    [void]ReceiveMessage() {
        foreach ($msg in $this.Backend.ReceiveMessage()) {

            # If the backend lazy loads and has done so
            if (($this.Backend.LazyLoadUsers) -and (-not $this.LazyLoadComplete)) {
                $this._LoadAdmins()
                $this.LazyLoadComplete = $true
            }

            # Ignore DMs if told to
            if ($msg.IsDM -and $this.Configuration.DisallowDMs) {
                $this.LogInfo('Ignoring message. DMs are disabled.', $msg)
                $this.AddReaction($msg, [ReactionType]::Denied)
                $response = [Response]::new($msg)
                $response.Severity = [Severity]::Warning
                $response.Data = New-PoshBotCardResponse -Type Warning -Text 'Sorry :( PoshBot has been configured to ignore DMs (direct messages). Please contact your bot administrator.'
                $this.SendMessage($response)
                return
            }

            # HTML decode message text
            # This will ensure characters like '&' that MAY have
            # been encoded as &amp; on their way in get translated
            # back to the original
            if (-not [string]::IsNullOrEmpty($msg.Text)) {
                $msg.Text = [System.Net.WebUtility]::HtmlDecode($msg.Text)
            }

            # Execute PreReceive middleware hooks
            $cmdExecContext = [CommandExecutionContext]::new()
            $cmdExecContext.Started = (Get-Date).ToUniversalTime()
            $cmdExecContext.Message = $msg
            $cmdExecContext = $this._ExecuteMiddleware($cmdExecContext, [MiddlewareType]::PreReceive)

            if ($cmdExecContext) {
                $this.LogDebug('Received bot message from chat network. Adding to message queue.', $cmdExecContext.Message)
                $this.MessageQueue.Enqueue($cmdExecContext.Message)
            }
        }
    }

    # Receive any messages from the scheduler that had their timer elapse and should be executed
    [void]ProcessScheduledMessages() {
        foreach ($msg in $this.Scheduler.GetTriggeredMessages()) {
            $this.LogDebug('Received scheduled message from scheduler. Adding to message queue.', $msg)
            $this.MessageQueue.Enqueue($msg)
        }
    }

    [void]ProcessDeferredContexts() {
        $now = (Get-Date).ToUniversalTime()
        $expireMinutes = $this.Configuration.ApprovalConfiguration.ExpireMinutes

        $toRemove = New-Object System.Collections.ArrayList
        foreach ($context in $this.DeferredCommandExecutionContexts.Values) {
            $expireTime = $context.Started.AddMinutes($expireMinutes)
            if ($now -gt $expireTime) {
                $msg = "[$($context.Id)] - [$($context.ParsedCommand.CommandString)] has been pending approval for more than [$expireMinutes] minutes. The command will be cancelled."

                # Add cancelled reation
                $this.RemoveReaction($context.Message, [ReactionType]::ApprovalNeeded)
                $this.AddReaction($context.Message, [ReactionType]::Cancelled)

                # Send message back to backend saying command context was cancelled due to timeout
                $this.LogInfo($msg)
                $response = [Response]::new($context.Message)
                $response.Data = New-PoshBotCardResponse -Type Warning -Text $msg
                $this.SendMessage($response)

                $toRemove.Add($context.Id)
            }
        }
        foreach ($id in $toRemove) {
            $this.DeferredCommandExecutionContexts.Remove($id)
        }

        while ($this.ProcessedDeferredContextQueue.Count -ne 0) {
            $cmdExecContext = $this.ProcessedDeferredContextQueue.Dequeue()
            $this.DeferredCommandExecutionContexts.Remove($cmdExecContext.Id)

            if ($cmdExecContext.ApprovalState -eq [ApprovalState]::Approved) {
                $this.LogDebug("Starting exeuction of context [$($cmdExecContext.Id)]")
                $this.RemoveReaction($cmdExecContext.Message, [ReactionType]::ApprovalNeeded)
                $this.Executor.ExecuteCommand($cmdExecContext)
            } elseif ($cmdExecContext.ApprovalState -eq [ApprovalState]::Denied) {
                $this.LogDebug("Context [$($cmdExecContext.Id)] was denied")
                $this.RemoveReaction($cmdExecContext.Message, [ReactionType]::ApprovalNeeded)
                $this.AddReaction($cmdExecContext.Message, [ReactionType]::Denied)
            }
        }
    }

    # Determine if message text is addressing the bot and should be
    # treated as a bot command
     [bool]IsBotCommand([Message]$Message) {
        $firstWord = ($Message.Text -split ' ')[0].Trim()
        foreach ($prefix in $this._PossibleCommandPrefixes ) {
            # If we've elected for a $null prefix, don't escape it
            # as [regex]::Escape() converts null chars into a space (' ')
            if ([char]$null -eq $prefix) {
                $prefix = ''
            } else {
                $prefix = [regex]::Escape($prefix)
            }

            if ($firstWord -match "^$prefix") {
                $this.LogDebug('Message is a bot command')
                return $true
            }
        }
        return $false
    }

    # Pull message(s) off queue and pass to handler
    [void]ProcessMessageQueue() {
        while ($this.MessageQueue.Count -ne 0) {
            $msg = $this.MessageQueue.Dequeue()
            $this.LogDebug('Dequeued message', $msg)
            $this.HandleMessage($msg)
        }
    }

    # Determine if the message received from the backend
    # is something the bot should act on
    [void]HandleMessage([Message]$Message) {
        # If message is intended to be a bot command
        # if this is false, and a trigger match is not found
        # then the message is just normal conversation that didn't
        # match a regex trigger. In that case, don't respond with an
        # error that we couldn't find the command
        $isBotCommand = $this.IsBotCommand($Message)

        $cmdSearch = $true
        if (-not $isBotCommand) {
            $cmdSearch = $false
            $this.LogDebug('Message is not a bot command. Command triggers WILL NOT be searched.')
        } else {
            # The message is intended to be a bot command
            $Message = $this.TrimPrefix($Message)
        }

        $parsedCommand = [CommandParser]::Parse($Message)
        $this.LogDebug('Parsed bot command', $parsedCommand)

        # Attempt to populate the parsed command with full user info from the backend
        $parsedCommand.CallingUserInfo = $this.Backend.GetUserInfo($parsedCommand.From)

        # Match parsed command to a command in the plugin manager
        $pluginCmd = $this.PluginManager.MatchCommand($parsedCommand, $cmdSearch)
        if ($pluginCmd) {

            # Create the command execution context
            $cmdExecContext = [CommandExecutionContext]::new()
            $cmdExecContext.Started = (Get-Date).ToUniversalTime()
            $cmdExecContext.Result = [CommandResult]::New()
            $cmdExecContext.Command = $pluginCmd.Command
            $cmdExecContext.FullyQualifiedCommandName = $pluginCmd.ToString()
            $cmdExecContext.ParsedCommand = $parsedCommand
            $cmdExecContext.Message = $Message

            # Execute PostReceive middleware hooks
            $cmdExecContext = $this._ExecuteMiddleware($cmdExecContext, [MiddlewareType]::PostReceive)

            if ($cmdExecContext) {
                # Check command is allowed in channel
                if (-not $this.CommandInAllowedChannel($parsedCommand, $pluginCmd)) {
                    $this.LogDebug('Igoring message. Command not approved in channel', $pluginCmd.ToString())
                    $this.AddReaction($Message, [ReactionType]::Denied)
                    $response = [Response]::new($Message)
                    $response.Severity = [Severity]::Warning
                    $response.Data = New-PoshBotCardResponse -Type Warning -Text 'Sorry :( PoshBot has been configured to not allow that command in this channel. Please contact your bot administrator.'
                    $this.SendMessage($response)
                    return
                }

                # Add the name of the plugin to the parsed command
                # if it wasn't fully qualified to begin with
                if ([string]::IsNullOrEmpty($parsedCommand.Plugin)) {
                    $parsedCommand.Plugin = $pluginCmd.Plugin.Name
                }

                # If the command trigger is a [regex], then we shoudn't parse named/positional
                # parameters from the message so clear them out. Only the regex matches and
                # config provided parameters are allowed.
                if ([TriggerType]::Regex -in $pluginCmd.Command.Triggers.Type) {
                    $parsedCommand.NamedParameters = @{}
                    $parsedCommand.PositionalParameters = @()
                    $regex = [regex]$pluginCmd.Command.Triggers[0].Trigger
                    $parsedCommand.NamedParameters['Arguments'] = $regex.Match($parsedCommand.CommandString).Groups | Select-Object -ExpandProperty Value
                }

                # Pass in the bot to the module command.
                # We need this for builtin commands
                if ($pluginCmd.Plugin.Name -eq 'Builtin') {
                    $parsedCommand.NamedParameters.Add('Bot', $this)
                }

                # Inspect the command and find any parameters that should
                # be provided from the bot configuration
                # Insert these as named parameters
                $configProvidedParams = $this.GetConfigProvidedParameters($pluginCmd)
                foreach ($cp in $configProvidedParams.GetEnumerator()) {
                    if (-not $parsedCommand.NamedParameters.ContainsKey($cp.Name)) {
                        $this.LogDebug("Inserting configuration provided named parameter", $cp)
                        $parsedCommand.NamedParameters.Add($cp.Name, $cp.Value)
                    }
                }

                # Execute PreExecute middleware hooks
                $cmdExecContext = $this._ExecuteMiddleware($cmdExecContext, [MiddlewareType]::PreExecute)

                if ($cmdExecContext) {
                    $this.Executor.ExecuteCommand($cmdExecContext)
                }
            }
        } else {
            if ($isBotCommand) {
                $msg = "No command found matching [$($Message.Text)]"
                $this.LogInfo([LogSeverity]::Warning, $msg, $parsedCommand)
                # Only respond with command not found message if configuration allows it.
                if (-not $this.Configuration.MuteUnknownCommand) {
                    $response = [Response]::new($Message)
                    $response.Severity = [Severity]::Warning
                    $response.Data = New-PoshBotCardResponse -Type Warning -Text $msg
                    $this.SendMessage($response)
                }
            }
        }
    }

    # Get completed jobs, determine success/error, then return response to backend
    [void]ProcessCompletedJobs() {
        $completedJobs = $this.Executor.ReceiveJob()

        $count = $completedJobs.Count
        if ($count -ge 1) {
            $this.LogInfo("Processing [$count] completed jobs")
        }

        foreach ($cmdExecContext in $completedJobs) {
            $this.LogInfo("Processing job execution [$($cmdExecContext.Id)]")

            # Execute PostExecute middleware hooks
            $cmdExecContext = $this._ExecuteMiddleware($cmdExecContext, [MiddlewareType]::PostExecute)

            if ($cmdExecContext) {
                $cmdExecContext.Response = [Response]::new($cmdExecContext.Message)

                if (-not $cmdExecContext.Result.Success) {
                    # Was the command authorized?
                    if (-not $cmdExecContext.Result.Authorized) {
                        $cmdExecContext.Response.Severity = [Severity]::Warning
                        $cmdExecContext.Response.Data = New-PoshBotCardResponse -Type Warning -Text "You do not have authorization to run command [$($cmdExecContext.Command.Name)] :(" -Title 'Command Unauthorized'
                        $this.LogInfo([LogSeverity]::Warning, 'Command unauthorized')
                    } else {
                        $cmdExecContext.Response.Severity = [Severity]::Error
                        if ($cmdExecContext.Result.Errors.Count -gt 0) {
                            $cmdExecContext.Response.Data = $cmdExecContext.Result.Errors | ForEach-Object {
                                if ($_.Exception) {
                                    New-PoshBotCardResponse -Type Error -Text $_.Exception.Message -Title 'Command Exception'
                                } else {
                                    New-PoshBotCardResponse -Type Error -Text $_ -Title 'Command Exception'
                                }
                            }
                        } else {
                            $cmdExecContext.Response.Data += New-PoshBotCardResponse -Type Error -Text 'Something bad happened :(' -Title 'Command Error'
                            $cmdExecContext.Response.Data += $cmdExecContext.Result.Errors
                        }
                        $this.LogInfo([LogSeverity]::Error, "Errors encountered running command [$($cmdExecContext.FullyQualifiedCommandName)]", $cmdExecContext.Result.Errors)
                    }
                } else {
                    $this.LogVerbose('Command execution result', $cmdExecContext.Result)
                    foreach ($resultOutput in $cmdExecContext.Result.Output) {
                        if ($null -ne $resultOutput) {
                            if ($this._IsCustomResponse($resultOutput)) {
                                $cmdExecContext.Response.Data += $resultOutput
                            } else {
                                # If the response is a simple type, just display it as a string
                                # otherwise we need remove auto-generated properties that show up
                                # from deserialized objects
                                if ($this._IsPrimitiveType($resultOutput)) {
                                    $cmdExecContext.Response.Text += $resultOutput.ToString().Trim()
                                } else {
                                    $deserializedProps = 'PSComputerName', 'PSShowComputerName', 'PSSourceJobInstanceId', 'RunspaceId'
                                    $resultText = $resultOutput | Select-Object -Property * -ExcludeProperty $deserializedProps
                                    $cmdExecContext.Response.Text += ($resultText | Format-List -Property * | Out-String).Trim()
                                }
                            }
                        }
                    }
                }

                # Write out this command execution to permanent storage
                if ($this.Configuration.LogCommandHistory) {
                    $logMsg = [LogMessage]::new("[$($cmdExecContext.FullyQualifiedCommandName)] was executed by [$($cmdExecContext.Message.From)]", $cmdExecContext.Summarize())
                    $cmdHistoryLogPath = Join-Path $this.Configuration.LogDirectory -ChildPath 'CommandHistory.log'
                    $this.Log($logMsg, $cmdHistoryLogPath, $this.Configuration.CommandHistoryMaxLogSizeMB, $this.Configuration.CommandHistoryMaxLogsToKeep)
                }

                # Send response back to user in private (DM) channel if this command
                # is marked to devert responses
                foreach ($rule in $this.Configuration.SendCommandResponseToPrivate) {
                    if ($cmdExecContext.FullyQualifiedCommandName -like $rule) {
                        $this.LogInfo("Deverting response from command [$($cmdExecContext.FullyQualifiedCommandName)] to private channel")
                        $cmdExecContext.Response.To = "@$($this.RoleManager.ResolveUserIdToUserName($cmdExecContext.Message.From))"
                        break
                    }
                }

                # Execute PreResponse middleware hooks
                $cmdExecContext = $this._ExecuteMiddleware($cmdExecContext, [MiddlewareType]::PreResponse)

                # Send response back to chat network
                if ($cmdExecContext) {
                    $this.SendMessage($cmdExecContext.Response)
                }

                # Execute PostResponse middleware hooks
                $cmdExecContext = $this._ExecuteMiddleware($cmdExecContext, [MiddlewareType]::PostResponse)
            }

            $this.LogInfo("Done processing command [$($cmdExecContext.FullyQualifiedCommandName)]")
        }
    }

    # Trim the command prefix or any alternate prefix or seperators off the message
    # as we won't need them anymore.
    [Message]TrimPrefix([Message]$Message) {
        if (-not [string]::IsNullOrEmpty($Message.Text)) {
            $firstWord = ($Message.Text -split ' ')[0].Trim()
            foreach ($prefix in $this._PossibleCommandPrefixes) {
                $prefixEscaped = [regex]::Escape($prefix)
                if ($firstWord -match "^$prefixEscaped") {
                    $Message.Text = $Message.Text.TrimStart($prefix).Trim()
                }
            }
        }
        return $Message
    }

    # Create complete list of command prefixes so we can quickly
    # evaluate messages from the chat network and determine if
    # they are bot commands
    [void]GenerateCommandPrefixList() {
        $this._PossibleCommandPrefixes.Add($this.Configuration.CommandPrefix)
        foreach ($alternatePrefix in $this.Configuration.AlternateCommandPrefixes) {
            $this._PossibleCommandPrefixes.Add($alternatePrefix) > $null
            foreach ($seperator in ($this.Configuration.AlternateCommandPrefixSeperators)) {
                $prefixPlusSeperator = "$alternatePrefix$seperator"
                $this._PossibleCommandPrefixes.Add($prefixPlusSeperator) > $null
            }
        }
        $this.LogDebug('Configured command prefixes', $this._PossibleCommandPrefixes)
    }

    # Send the response to the backend to execute
    [void]SendMessage([Response]$Response) {
        $this.LogInfo('Sending response to backend')
        $this.Backend.SendMessage($Response)
    }

    # Add a reaction to a message
    [void]AddReaction([Message]$Message, [ReactionType]$ReactionType) {
        if ($this.Configuration.AddCommandReactions) {
            $this.Backend.AddReaction($Message, $ReactionType)
        }
    }

    # Remove a reaction from a message
    [void]RemoveReaction([Message]$Message, [ReactionType]$ReactionType) {
        if ($this.Configuration.AddCommandReactions) {
            $this.Backend.RemoveReaction($Message, $ReactionType)
        }
    }

    # Get any parameters with the
    [hashtable]GetConfigProvidedParameters([PluginCommand]$PluginCmd) {
        if ($PluginCmd.Command.FunctionInfo) {
            $command = $PluginCmd.Command.FunctionInfo
        } else {
            $command = $PluginCmd.Command.CmdletInfo
        }
        $this.LogDebug("Inspecting command [$($PluginCmd.ToString())] for configuration-provided parameters")
        $configParams = foreach($param in $Command.Parameters.GetEnumerator() | Select-Object -ExpandProperty Value) {
            foreach ($attr in $param.Attributes) {
                if ($attr.GetType().ToString() -eq 'PoshBot.FromConfig') {
                    [ConfigProvidedParameter]::new($attr, $param)
                }
            }
        }

        $configProvidedParams = @{}
        if ($configParams) {
            $configParamNames = $configParams.Parameter | Select-Object -ExpandProperty Name
            $this.LogInfo("Command [$($PluginCmd.ToString())] has configuration provided parameters", $configParamNames)
            $pluginConfig = $this.Configuration.PluginConfiguration[$PluginCmd.Plugin.Name]
            if ($pluginConfig) {
                $this.LogDebug("Inspecting bot configuration for parameter values matching command [$($PluginCmd.ToString())]")
                foreach ($cp in $configParams) {
                    if (-not [string]::IsNullOrEmpty($cp.Metadata.Name)) {
                        $configParamName = $cp.Metadata.Name
                    } else {
                        $configParamName = $cp.Parameter.Name
                    }

                    if ($pluginConfig.ContainsKey($configParamName)) {
                        $configProvidedParams.Add($cp.Parameter.Name, $pluginConfig[$configParamName])
                    }
                }
                if ($configProvidedParams.Count -ge 0) {
                    $this.LogDebug('Configuration supplied parameter values', $configProvidedParams)
                }
            } else {
                # No plugin configuration defined.
                # Unable to provide values for these parameters
                $this.LogDebug([LogSeverity]::Warning, "Command [$($PluginCmd.ToString())] has requested configuration supplied parameters but none where found")
            }
        } else {
            $this.LogDebug("Command [$($PluginCmd.ToString())] has 0 configuration provided parameters")
        }

        return $configProvidedParams
    }

    # Check command against approved commands in channels
    [bool]CommandInAllowedChannel([ParsedCommand]$ParsedCommand, [PluginCommand]$PluginCommand) {

        # DMs won't be governed by the 'ApprovedCommandsInChannel' configuration property
        if ($ParsedCommand.OriginalMessage.IsDM) {
            return $true
        }

        $channel = $ParsedCommand.ToName
        $fullyQualifiedCommand = $PluginCommand.ToString()

        # Match command against included/excluded commands for the channel
        # If there is a channel match, assume command is NOT approved unless
        # it matches the included commands list and DOESN'T match the excluded list
        foreach ($ChannelRule in $this.Configuration.ChannelRules) {
            if ($channel -like $ChannelRule.Channel) {
                foreach ($includedCommand in $ChannelRule.IncludeCommands) {
                    if ($fullyQualifiedCommand -like $includedCommand) {
                        $this.LogDebug("Matched [$fullyQualifiedCommand] to included command [$includedCommand]")
                        foreach ($excludedCommand in $ChannelRule.ExcludeCommands) {
                            if ($fullyQualifiedCommand -like $excludedCommand) {
                                $this.LogDebug("Matched [$fullyQualifiedCommand] to excluded command [$excludedCommand]")
                                return $false
                            }
                        }

                        return $true
                    }
                }
                return $false
            }
        }

        return $false
    }

    # Determine if response from command is custom and the output should be formatted
    hidden [bool]_IsCustomResponse([object]$Response) {
        $isCustom = (($Response.PSObject.TypeNames[0] -eq 'PoshBot.Text.Response') -or
                     ($Response.PSObject.TypeNames[0] -eq 'PoshBot.Card.Response') -or
                     ($Response.PSObject.TypeNames[0] -eq 'PoshBot.File.Upload') -or
                     ($Response.PSObject.TypeNames[0] -eq 'Deserialized.PoshBot.Text.Response') -or
                     ($Response.PSObject.TypeNames[0] -eq 'Deserialized.PoshBot.Card.Response') -or
                     ($Response.PSObject.TypeNames[0] -eq 'Deserialized.PoshBot.File.Upload'))

        if ($isCustom) {
            $this.LogDebug("Detected custom response [$($Response.PSObject.TypeNames[0])] from command")
        }

        return $isCustom
    }

    # Test if an object is a primitive data type
    hidden [bool] _IsPrimitiveType([object]$Item) {
        $primitives = @('Byte', 'SByte', 'Int16', 'Int32', 'Int64', 'UInt16', 'UInt32', 'UInt64',
                        'Decimal', 'Single', 'Double', 'TimeSpan', 'DateTime', 'ProgressRecord',
                        'Char', 'String', 'XmlDocument', 'SecureString', 'Boolean', 'Guid', 'Uri', 'Version'
        )
        return ($Item.GetType().Name -in $primitives)
    }

    hidden [CommandExecutionContext] _ExecuteMiddleware([CommandExecutionContext]$Context, [MiddlewareType]$Type) {

        $hooks = $this._Middleware."$($Type.ToString())Hooks"

        # Execute PostResponse middleware hooks
        foreach ($hook in $hooks.Values) {
            try {
                $this.LogDebug("Executing [$($Type.ToString())] hook [$($hook.Name)]")
                if ($null -ne $Context) {
                    $Context = $hook.Execute($Context, $this)
                    if ($null -eq $Context) {
                        $this.LogInfo([LogSeverity]::Warning, "[$($Type.ToString())] middleware [$($hook.Name)] dropped message.")
                        break
                    }
                }
            } catch {
                $this.LogInfo([LogSeverity]::Error, "[$($Type.ToString())] middleware [$($hook.Name)] raised an exception. Command context dropped.", [ExceptionFormatter]::Summarize($_))
                return $null
            }
        }

        return $Context
    }

    # Resolve any bot administrators defined in configuration to their IDs
    # and add to the [admin] role
    hidden [void] _LoadAdmins() {
        foreach ($admin in $this.Configuration.BotAdmins) {
            if ($adminId = $this.RoleManager.ResolveUsernameToId($admin)) {
                try {
                    $this.RoleManager.AddUserToGroup($adminId, 'Admin')
                } catch {
                    $this.LogInfo([LogSeverity]::Warning, "Unable to add [$admin] to [Admin] group", [ExceptionFormatter]::Summarize($_))
                }
            } else {
                $this.LogInfo([LogSeverity]::Warning, "Unable to resolve ID for admin [$admin]")
            }
        }
    }
}

function Get-PoshBot {
    <#
    .SYNOPSIS
        Gets any currently running instances of PoshBot that are running as background jobs.
    .DESCRIPTION
        PoshBot can be run in the background with PowerShell jobs. This function returns
        any currently running PoshBot instances.
    .PARAMETER Id
        One or more job IDs to retrieve.
    .EXAMPLE
        PS C:\> Get-PoshBot

        Id         : 5
        Name       : PoshBot_3ddfc676406d40fca149019d935f065d
        State      : Running
        InstanceId : 3ddfc676406d40fca149019d935f065d
        Config     : BotConfiguration

    .EXAMPLE
        PS C:\> Get-PoshBot -Id 100

        Id         : 100
        Name       : PoshBot_eab96f2ad147489b9f90e110e02ad805
        State      : Running
        InstanceId : eab96f2ad147489b9f90e110e02ad805
        Config     : BotConfiguration

        Gets the PoshBot job instance with ID 100.
    .INPUTS
        System.Int32
    .OUTPUTS
        PSCustomObject
    .LINK
        Start-PoshBot
    .LINK
        Stop-PoshBot
    #>
    [OutputType([PSCustomObject])]
    [cmdletbinding()]
    param(
        [parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [int[]]$Id = @()
    )

    process {
        if ($Id.Count -gt 0) {
            foreach ($item in $Id) {
                if ($b = $script:botTracker.$item) {
                    [pscustomobject][ordered]@{
                        Id = $item
                        Name = $b.Name
                        State = (Get-Job -Id $b.jobId).State
                        InstanceId = $b.InstanceId
                        Config = $b.Config
                    }
                }
            }
        } else {
            $script:botTracker.GetEnumerator() | ForEach-Object {
                [pscustomobject][ordered]@{
                    Id = $_.Value.JobId
                    Name = $_.Value.Name
                    State = (Get-Job -Id $_.Value.JobId).State
                    InstanceId = $_.Value.InstanceId
                    Config = $_.Value.Config
                }
            }
        }
    }
}

Export-ModuleMember -Function 'Get-PoshBot'


function Get-PoshBotConfiguration {
    <#
    .SYNOPSIS
        Gets a PoshBot configuration from a file.
    .DESCRIPTION
        PoshBot configurations can be stored on the filesytem in PowerShell data (.psd1) files.
        This functions will load that file and return a [BotConfiguration] object.
    .PARAMETER Path
        One or more paths to a PoshBot configuration file.
    .PARAMETER LiteralPath
        Specifies the path(s) to the current location of the file(s). Unlike the Path parameter, the value of LiteralPath is used exactly as it is typed.
        No characters are interpreted as wildcards. If the path includes escape characters, enclose it in single quotation marks. Single quotation
        marks tell PowerShell not to interpret any characters as escape sequences.
    .EXAMPLE
        PS C:\> Get-PoshBotConfiguration -Path C:\Users\joeuser\.poshbot\Cherry2000.psd1

        Name                             : Cherry2000
        ConfigurationDirectory           : C:\Users\joeuser\.poshbot
        LogDirectory                     : C:\Users\joeuser\.poshbot\Logs
        PluginDirectory                  : C:\Users\joeuser\.poshbot
        PluginRepository                 : {PSGallery}
        ModuleManifestsToLoad            : {}
        LogLevel                         : Debug
        BackendConfiguration             : {Token, Name}
        PluginConfiguration              : {}
        BotAdmins                        : {joeuser}
        CommandPrefix                    : !
        AlternateCommandPrefixes         : {bender, hal}
        AlternateCommandPrefixSeperators : {:, ,, ;}
        SendCommandResponseToPrivate     : {}
        MuteUnknownCommand               : False
        AddCommandReactions              : True

        Gets the bot configuration located at [C:\Users\joeuser\.poshbot\Cherry2000.psd1].
    .EXAMPLE
        PS C:\> $botConfig = 'C:\Users\joeuser\.poshbot\Cherry2000.psd1' | Get-PoshBotConfiguration

        Gets the bot configuration located at [C:\Users\brand\.poshbot\Cherry2000.psd1].
    .INPUTS
        String
    .OUTPUTS
        BotConfiguration
    .LINK
        New-PoshBotConfiguration
    .LINK
        Start-PoshBot
    #>
    [cmdletbinding(DefaultParameterSetName = 'Path')]
    param(
        [parameter(
            Mandatory,
            ParameterSetName  = 'Path',
            Position = 0,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName
        )]
        [ValidateNotNullOrEmpty()]
        [SupportsWildcards()]
        [string[]]$Path,

        [parameter(
            Mandatory,
            ParameterSetName = 'LiteralPath',
            Position = 0,
            ValueFromPipelineByPropertyName
        )]
        [ValidateNotNullOrEmpty()]
        [Alias('PSPath')]
        [string[]]$LiteralPath
    )

    process {
        # Resolve path(s)
        if ($PSCmdlet.ParameterSetName -eq 'Path') {
            $paths = Resolve-Path -Path $Path | Select-Object -ExpandProperty Path
        } elseif ($PSCmdlet.ParameterSetName -eq 'LiteralPath') {
            $paths = Resolve-Path -LiteralPath $LiteralPath | Select-Object -ExpandProperty Path
        }

        foreach ($item in $paths) {
            if (Test-Path $item) {
                if ( (Get-Item -Path $item).Extension -eq '.psd1') {
                    Write-Verbose -Message "Loading bot configuration from [$item]"
                    $hash = Get-Content -Path $item -Raw | ConvertFrom-Metadata
                    $config = [BotConfiguration]::new()
                    foreach ($key in $hash.Keys) {
                        if ($config | Get-Member -MemberType Property -Name $key) {
                            switch ($key) {
                                'ChannelRules' {
                                    $config.ChannelRules = @()
                                    foreach ($item in $hash[$key]) {
                                        $config.ChannelRules += [ChannelRule]::new($item.Channel, $item.IncludeCommands, $item.ExcludeCommands)
                                    }
                                    break
                                }
                                'ApprovalConfiguration' {
                                    # Validate ExpireMinutes
                                    if ($hash[$key].ExpireMinutes -is [int]) {
                                        $config.ApprovalConfiguration.ExpireMinutes = $hash[$key].ExpireMinutes
                                    }
                                    # Validate ApprovalCommandConfiguration
                                    if ($hash[$key].Commands.Count -ge 1) {
                                        foreach ($approvalConfig in $hash[$key].Commands) {
                                            $acc = [ApprovalCommandConfiguration]::new()
                                            $acc.Expression = $approvalConfig.Expression
                                            $acc.ApprovalGroups = $approvalConfig.Groups
                                            $acc.PeerApproval = $approvalConfig.PeerApproval
                                            $config.ApprovalConfiguration.Commands.Add($acc) > $null
                                        }
                                    }
                                    break
                                }
                                'MiddlewareConfiguration' {
                                    foreach ($type in [enum]::GetNames([MiddlewareType])) {
                                        foreach ($item in $hash[$key].$type) {
                                            $config.MiddlewareConfiguration.Add([MiddlewareHook]::new($item.Name, $item.Path), $type)
                                        }
                                    }
                                    break
                                }
                                Default {
                                    $config.$Key = $hash[$key]
                                    break
                                }
                            }
                        }
                    }
                    $config
                } else {
                    Throw 'Path must be to a valid .psd1 file'
                }
            } else {
                Write-Error -Message "Path [$item] is not valid."
            }
        }
    }
}

Export-ModuleMember -Function 'Get-PoshBotConfiguration'


function Get-PoshBotStatefulData {
    <#
    .SYNOPSIS
        Get stateful data previously exported from a PoshBot command
    .DESCRIPTION
        Get stateful data previously exported from a PoshBot command

        Reads data from the PoshBot ConfigurationDirectory.
    .PARAMETER Name
        If specified, retrieve only this property from the stateful data
    .PARAMETER ValueOnly
        If specified, return only the value of the specified property Name
    .PARAMETER Scope
        Get stateful data from this scope:
            Module: Data scoped to this plugin
            Global: Data available to any Poshbot plugin
    .EXAMPLE
        $ModuleData = Get-PoshBotStatefulData

        Get all stateful data for the PoshBot plugin this runs from
    .EXAMPLE
        $Something = Get-PoshBotStatefulData -Name 'Something' -ValueOnly -Scope Global

        Set $Something to the value of the 'Something' property from Poshbot's global stateful data
    .LINK
        Set-PoshBotStatefulData
    .LINK
        Remove-PoshBotStatefulData
    .LINK
        Start-PoshBot
    #>
    [cmdletbinding()]
    param(
        [string]$Name = '*',

        [switch]$ValueOnly,

        [validateset('Global','Module')]
        [string]$Scope = 'Module'
    )
    process {
        if($Scope -eq 'Module') {
            $FileName = "$($global:PoshBotContext.Plugin).state"
        } else {
            $FileName = "PoshbotGlobal.state"
        }
        $Path = Join-Path $global:PoshBotContext.ConfigurationDirectory $FileName

        if(-not (Test-Path $Path)) {
            Write-Verbose "Requested stateful data file not found: [$Path]"
            return
        }
        Write-Verbose "Getting stateful data from [$Path]"
        $Output = Import-Clixml -Path $Path | Select-Object -Property $Name
        if($ValueOnly)
        {
            $Output = $Output.${Name}
        }
        $Output
    }
}

Export-ModuleMember -Function 'Get-PoshBotStatefulData'


function New-PoshBotCardResponse {
    <#
    .SYNOPSIS
        Tells PoshBot to send a specially formatted response.
    .DESCRIPTION
        Responses from PoshBot commands can either be plain text or formatted. Returning a response with New-PoshBotRepsonse will tell PoshBot
        to craft a specially formatted message when sending back to the chat network.
    .PARAMETER Type
        Specifies a preset color for the card response. If the [Color] parameter is specified as well, it will override this parameter.

        | Type    | Color  | Hex code |
        |---------|--------|----------|
        | Normal  | Greed  | #008000  |
        | Warning | Yellow | #FFA500  |
        | Error   | Red    | #FF0000  |
    .PARAMETER Text
        The text response from the command.
    .PARAMETER DM
        Tell PoshBot to redirect the response to a DM channel.
    .PARAMETER Title
        The title of the response. This will be the card title in chat networks like Slack.
    .PARAMETER ThumbnailUrl
        A URL to a thumbnail image to display in the card response.
    .PARAMETER ImageUrl
        A URL to an image to display in the card response.
    .PARAMETER LinkUrl
        Will turn the title into a hyperlink
    .PARAMETER Fields
        A hashtable to display as a table in the card response.
    .PARAMETER COLOR
        The hex color code to use for the card response. In Slack, this will be the color of the left border in the message attachment.
    .PARAMETER CustomData
        Any additional custom data you'd like to pass on. Useful for custom backends, in case you want to pass a specifically formatted response
        in the Data stream of the responses received by the backend. Any data sent here will be skipped by the built-in backends provided with PoshBot itself.
    .EXAMPLE
        function Do-Something {
            [cmdletbinding()]
            param(
                [parameter(mandatory)]
                [string]$MyParam
            )

            New-PoshBotCardResponse -Type Normal -Text 'OK, I did something.' -ThumbnailUrl 'https://www.streamsports.com/images/icon_green_check_256.png'
        }

        Tells PoshBot to send a formatted response back to the chat network. In Slack for example, this response will be a message attachment
        with a green border on the left, some text and a green checkmark thumbnail image.
    .EXAMPLE
        function Do-Something {
            [cmdletbinding()]
            param(
                [parameter(mandatory)]
                [string]$ComputerName
            )

            $info = Get-ComputerInfo -ComputerName $ComputerName -ErrorAction SilentlyContinue
            if ($info) {
                $fields = [ordered]@{
                    Name = $ComputerName
                    OS = $info.OSName
                    Uptime = $info.Uptime
                    IPAddress = $info.IPAddress
                }
                New-PoshBotCardResponse -Type Normal -Fields $fields
            } else {
                New-PoshBotCardResponse -Type Error -Text 'Something bad happended :(' -ThumbnailUrl 'http://p1cdn05.thewrap.com/images/2015/06/don-draper-shrug.jpg'
            }
        }

        Attempt to retrieve some information from a given computer and return a card response back to PoshBot. If the command fails for some reason,
        return a card response specified the error and a sad image.
    .OUTPUTS
        PSCustomObject
    .LINK
        New-PoshBotTextResponse
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='*')]
    [cmdletbinding()]
    param(
        [ValidateSet('Normal', 'Warning', 'Error')]
        [string]$Type = 'Normal',

        [switch]$DM,

        [string]$Text = [string]::empty,

        [string]$Title,

        [ValidateScript({
            $uri = $null
            if ([system.uri]::TryCreate($_, [System.UriKind]::Absolute, [ref]$uri)) {
                return $true
            } else {
                $msg = 'ThumbnailUrl must be a valid URL'
                throw [System.Management.Automation.ValidationMetadataException]$msg
            }
        })]
        [string]$ThumbnailUrl,

        [ValidateScript({
            $uri = $null
            if ([system.uri]::TryCreate($_, [System.UriKind]::Absolute, [ref]$uri)) {
                return $true
            } else {
                $msg = 'ImageUrl must be a valid URL'
                throw [System.Management.Automation.ValidationMetadataException]$msg
            }
        })]
        [string]$ImageUrl,

        [ValidateScript({
            $uri = $null
            if ([system.uri]::TryCreate($_, [System.UriKind]::Absolute, [ref]$uri)) {
                return $true
            } else {
                $msg = 'LinkUrl must be a valid URL'
                throw [System.Management.Automation.ValidationMetadataException]$msg
            }
        })]
        [string]$LinkUrl,

        [System.Collections.IDictionary]$Fields,

        [ValidateScript({
            if ($_ -match '^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$') {
                return $true
            } else {
                $msg = 'Color but be a valid hexidecimal color code e.g. ##008000'
                throw [System.Management.Automation.ValidationMetadataException]$msg
            }
        })]
        [string]$Color = '#D3D3D3',

        [object]$CustomData
    )

    $response = [ordered]@{
        PSTypeName = 'PoshBot.Card.Response'
        Type = $Type
        Text = $Text.Trim()
        Private = $PSBoundParameters.ContainsKey('Private')
        DM = $PSBoundParameters['DM']
    }
    if ($PSBoundParameters.ContainsKey('Title')) {
        $response.Title = $Title
    }
    if ($PSBoundParameters.ContainsKey('ThumbnailUrl')) {
        $response.ThumbnailUrl = $ThumbnailUrl
    }
    if ($PSBoundParameters.ContainsKey('ImageUrl')) {
        $response.ImageUrl = $ImageUrl
    }
    if ($PSBoundParameters.ContainsKey('LinkUrl')) {
        $response.LinkUrl = $LinkUrl
    }
    if ($PSBoundParameters.ContainsKey('Fields')) {
        $response.Fields = $Fields
    }
    if ($PSBoundParameters.ContainsKey('CustomData')) {
        $response.CustomData = $CustomData
    }
    if ($PSBoundParameters.ContainsKey('Color')) {
        $response.Color = $Color
    } else {
        switch ($Type) {
            'Normal' {
                $response.Color = '#008000'
            }
            'Warning' {
                $response.Color = '#FFA500'
            }
            'Error' {
                $response.Color = '#FF0000'
            }
        }
    }

    [pscustomobject]$response
}

Export-ModuleMember -Function 'New-PoshBotCardResponse'


function New-PoshBotConfiguration {
    <#
    .SYNOPSIS
        Creates a new PoshBot configuration object.
    .DESCRIPTION
        Creates a new PoshBot configuration object.
    .PARAMETER Name
        The name the bot instance will be known as.
    .PARAMETER ConfigurationDirectory
        The directory when PoshBot configuration data will be written to.
    .PARAMETER LogDirectory
        The log directory logs will be written to.
    .PARAMETER PluginDirectory
        The directory PoshBot will look for PowerShell modules.
        This path will be prepended to your $env:PSModulePath.
    .PARAMETER PluginRepository
        One or more PowerShell repositories to look in when installing new plugins (modules).
        These will be the repository name(s) as found in Get-PSRepository.
    .PARAMETER ModuleManifestsToLoad
        One or more paths to module manifest (.psd1) files. These modules will be automatically
        loaded when PoshBot starts.
    .PARAMETER LogLevel
        The level of logging that PoshBot will do.
    .PARAMETER MaxLogSizeMB
        The maximum log file size in megabytes.
    .PARAMETER MaxLogsToKeep
        The maximum number of logs to keep. Once this value is reached, logs will start rotating.
    .PARAMETER LogCommandHistory
        Enable command history to be logged to a separate file for convenience. The default it $true
    .PARAMETER CommandHistoryMaxLogSizeMB
        The maximum log file size for the command history
    .PARAMETER CommandHistoryMaxLogsToKeep
        The maximum number of logs to keep for command history. Once this value is reached, the logs will start rotating.
    .PARAMETER BackendConfiguration
        A hashtable of configuration options required by the backend chat network implementation.
    .PARAMETER PluginConfiguration
        A hashtable of configuration options used by the various plugins (modules) that are installed in PoshBot.
        Each key in the hashtable must be the name of a plugin. The value of that hashtable item will be another hashtable
        with each key matching a parameter name in one or more commands of that module. A plugin command can specifiy that a
        parameter gets its value from this configuration by applying the custom attribute [PoshBot.FromConfig()] on
        the parameter.

        The function below is stating that the parameter $MyParam will get its value from the plugin configuration. The user
        running this command in PoshBot does not need to specify this parameter. PoshBot will dynamically resolve and apply
        the matching value from the plugin configuration when the command is executed.

        function Get-Foo {
            [cmdletbinding()]
            param(
                [PoshBot.FromConfig()]
                [parameter(mandatory)]
                [string]$MyParam
            )

            Write-Output $MyParam
        }

        If the function below was part of the Demo plugin, PoshBot will look in the plugin configuration for a key matching Demo
        and a child key matching $MyParam.

        Example plugin configuration:
        @{
            Demo = @{
                MyParam = 'bar'
            }
        }
    .PARAMETER BotAdmins
        An array of chat handles that will be granted admin rights in PoshBot. Any user in this array will have full rights in PoshBot. At startup,
        PoshBot will resolve these handles into IDs given by the chat network.
    .PARAMETER CommandPrefix
        The prefix (single character) that must be specified in front of a command in order for PoshBot to recognize the chat message as a bot command.

        !get-foo --value bar
    .PARAMETER AlternateCommandPrefixes
        Some users may want to specify alternate prefixes when calling bot comamnds. Use this parameter to specify an array of words that PoshBot
        will also check when parsing a chat message.

        bender get-foo --value bar

        hal open-doors --type pod
    .PARAMETER AlternateCommandPrefixSeperators
        An array of characters that can also ben used when referencing bot commands.

        bender, get-foo --value bar

        hal; open-doors --type pod
    .PARAMETER SendCommandResponseToPrivate
        A list of fully qualified (<PluginName>:<CommandName>) plugin commands that will have their responses redirected back to a direct message
        channel with the calling user rather than a shared channel.

        @(
            demo:get-foo
            network:ping
        )
    .PARAMETER MuteUnknownCommand
        Instead of PoshBot returning a warning message when it is unable to find a command, use this to parameter to tell PoshBot to return nothing.
    .PARAMETER AddCommandReactions
        Add reactions to a chat message indicating the command is being executed, has succeeded, or failed.
    .PARAMETER ApprovalExpireMinutes
        The amount of time (minutes) that a command the requires approval will be pending until it expires.
    .PARAMETER DisallowDMs
        Disallow DMs (direct messages) with the bot. If a user tries to DM the bot it will be ignored.
    .PARAMETER FormatEnumerationLimitOverride
        Set $FormatEnumerationLimit to this.  Defaults to unlimited (-1)

        Determines how many enumerated items are included in a display.
        This variable does not affect the underlying objects; just the display.
        When the value of $FormatEnumerationLimit is less than the number of enumerated items, PowerShell adds an ellipsis (...) to indicate items not shown.
    .PARAMETER ApprovalCommandConfigurations
        Array of hashtables containing command approval configurations.

        @(
            @{
                Expression = 'MyModule:Execute-Deploy:*'
                Groups = 'platform-admins'
                PeerApproval = $true
            }
            @{
                Expression = 'MyModule:Deploy-HRApp:*'
                Groups = @('platform-managers', 'hr-managers')
                PeerApproval = $true
            }
        )
    .PARAMETER ChannelRules
        Array of channels rules that control what plugin commands are allowed in a channel. Wildcards are supported.
        Channel names that match against this list will be allowed to have Poshbot commands executed in them.

        Internally this uses the `-like` comparison operator, not `-match`. Regexes are not allowed.

        For best results, list channels and commands from most specific to least specific. PoshBot will
        evaluate the first match found.

        Note that the bot will still receive messages from all channels it is a member of. These message MAY
        be logged depending on your configured logging level.

        Example value:
        @(
            # Only allow builtin commands in the 'botadmin' channel
            @{
                Channel = 'botadmin'
                IncludeCommands = @('builtin:*')
                ExcludeCommands = @()
            }
            # Exclude builtin commands from any "projectX" channel
            @{
                Channel = '*projectx*'
                IncludeCommands = @('*')
                ExcludeCommands = @('builtin:*')
            }
            # It's the wild west in random, except giphy :)
            @{
                Channel = 'random'
                IncludeCommands = @('*')
                ExcludeCommands = @('*giphy*')
            }
            # All commands are otherwise allowed
            @{
                Channel = '*'
                IncludeCommands = @('*')
                ExcludeCommands = @()
            }
        )
    .PARAMETER PreReceiveMiddlewareHooks
        Array of middleware scriptblocks that will run before PoshBot "receives" the message from the backend implementation.
        This middleware will receive the original message sent from the chat network and have a chance to modify, analyze, and optionally drop the message before PoshBot continues processing it.
    .PARAMETER PostReceiveMiddlewareHooks
        Array of middleware scriptblocks that will run after a message is "received" from the backend implementation.
        This middleware runs after messages have been parsed and matched with a registered command in PoshBot.
    .PARAMETER PreExecuteMiddlewareHooks
        Array of middleware scriptblocks that will run before a command is executed.
        This middleware is a good spot to run extra authentication or validation processes before commands are executed.
    .PARAMETER PostExecuteMiddlewareHooks
        Array of middleware scriptblocks that will run after PoshBot commands have been executed.
        This middleware is a good spot for custom logging solutions to write command history to a custom location.
    .PARAMETER PreResponseMiddlewareHooks
        Array of middleware scriptblocks that will run before command responses are sent to the backend implementation.
        This middleware is a good spot for modifying or sanitizing responses before they are sent to the chat network.
    .PARAMETER PostResponseMiddlewareHooks
        Array of middleware scriptblocks that will run after command responses have been sent to the backend implementation.
        This middleware runs after all processing is complete for a command and is a good spot for additional custom logging.
    .EXAMPLE
        PS C:\> New-PoshBotConfiguration -Name Cherry2000 -AlternateCommandPrefixes @('Cherry', 'Sam')

        Name                             : Cherry2000
        ConfigurationDirectory           : C:\Users\brand\.poshbot
        LogDirectory                     : C:\Users\brand\.poshbot
        PluginDirectory                  : C:\Users\brand\.poshbot
        PluginRepository                 : {PSGallery}
        ModuleManifestsToLoad            : {}
        LogLevel                         : Verbose
        BackendConfiguration             : {}
        PluginConfiguration              : {}
        BotAdmins                        : {}
        CommandPrefix                    : !
        AlternateCommandPrefixes         : {Cherry, Sam}
        AlternateCommandPrefixSeperators : {:, ,, ;}
        SendCommandResponseToPrivate     : {}
        MuteUnknownCommand               : False
        AddCommandReactions              : True
        ApprovalConfiguration            : ApprovalConfiguration

        Create a new PoshBot configuration with default values except for the bot name and alternate command prefixes that it will listen for.
    .EXAMPLE
        PS C:\> $backend = @{Name = 'SlackBackend'; Token = 'xoxb-569733935137-njOPkyBThqOTTUnCZb7tZpKK'}
        PS C:\> $botParams = @{
                    Name = 'HAL9000'
                    LogLevel = 'Info'
                    BotAdmins = @('JoeUser')
                    BackendConfiguration = $backend
                }
        PS C:\> $myBotConfig = New-PoshBotConfiguration @botParams
        PS C:\> $myBotConfig

        Name                             : HAL9000
        ConfigurationDirectory           : C:\Users\brand\.poshbot
        LogDirectory                     : C:\Users\brand\.poshbot
        PluginDirectory                  : C:\Users\brand\.poshbot
        PluginRepository                 : {MyLocalRepo}
        ModuleManifestsToLoad            : {}
        LogLevel                         : Info
        BackendConfiguration             : {}
        PluginConfiguration              : {}
        BotAdmins                        : {JoeUser}
        CommandPrefix                    : !
        AlternateCommandPrefixes         : {poshbot}
        AlternateCommandPrefixSeperators : {:, ,, ;}
        SendCommandResponseToPrivate     : {}
        MuteUnknownCommand               : False
        AddCommandReactions              : True
        ApprovalConfiguration            : ApprovalConfiguration

        PS C:\> $myBotConfig | Start-PoshBot -AsJob

        Create a new PoshBot configuration with a Slack backend. Slack's backend only requires a bot token to be specified. Ensure the person
        with Slack handle 'JoeUser' is a bot admin.
    .OUTPUTS
        BotConfiguration
    .LINK
        Get-PoshBotConfiguration
    .LINK
        Save-PoshBotConfiguration
    .LINK
        New-PoshBotInstance
    .LINK
        Start-PoshBot
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='*')]
    [cmdletbinding()]
    param(
        [string]$Name = 'PoshBot',
        [string]$ConfigurationDirectory = $script:defaultPoshBotDir,
        [string]$LogDirectory = $script:defaultPoshBotDir,
        [string]$PluginDirectory = $script:defaultPoshBotDir,
        [string[]]$PluginRepository = @('PSGallery'),
        [string[]]$ModuleManifestsToLoad = @(),
        [LogLevel]$LogLevel = [LogLevel]::Verbose,
        [int]$MaxLogSizeMB = 10,
        [int]$MaxLogsToKeep = 5,
        [bool]$LogCommandHistory = $true,
        [int]$CommandHistoryMaxLogSizeMB = 10,
        [int]$CommandHistoryMaxLogsToKeep = 5,
        [hashtable]$BackendConfiguration = @{},
        [hashtable]$PluginConfiguration = @{},
        [string[]]$BotAdmins = @(),
        [char]$CommandPrefix = '!',
        [string[]]$AlternateCommandPrefixes = @('poshbot'),
        [char[]]$AlternateCommandPrefixSeperators = @(':', ',', ';'),
        [string[]]$SendCommandResponseToPrivate = @(),
        [bool]$MuteUnknownCommand = $false,
        [bool]$AddCommandReactions = $true,
        [int]$ApprovalExpireMinutes = 30,
        [switch]$DisallowDMs,
        [int]$FormatEnumerationLimitOverride = -1,
        [hashtable[]]$ApprovalCommandConfigurations = @(),
        [hashtable[]]$ChannelRules = @(),
        [MiddlewareHook[]]$PreReceiveMiddlewareHooks   = @(),
        [MiddlewareHook[]]$PostReceiveMiddlewareHooks  = @(),
        [MiddlewareHook[]]$PreExecuteMiddlewareHooks   = @(),
        [MiddlewareHook[]]$PostExecuteMiddlewareHooks  = @(),
        [MiddlewareHook[]]$PreResponseMiddlewareHooks  = @(),
        [MiddlewareHook[]]$PostResponseMiddlewareHooks = @()
    )

    Write-Verbose -Message 'Creating new PoshBot configuration'
    $config = [BotConfiguration]::new()
    $config.Name = $Name
    $config.ConfigurationDirectory = $ConfigurationDirectory
    $config.AlternateCommandPrefixes = $AlternateCommandPrefixes
    $config.AlternateCommandPrefixSeperators = $AlternateCommandPrefixSeperators
    $config.BotAdmins = $BotAdmins
    $config.CommandPrefix = $CommandPrefix
    $config.LogDirectory = $LogDirectory
    $config.LogLevel = $LogLevel
    $config.MaxLogSizeMB = $MaxLogSizeMB
    $config.MaxLogsToKeep = $MaxLogsToKeep
    $config.LogCommandHistory = $LogCommandHistory
    $config.CommandHistoryMaxLogSizeMB = $CommandHistoryMaxLogSizeMB
    $config.CommandHistoryMaxLogsToKeep = $CommandHistoryMaxLogsToKeep
    $config.BackendConfiguration = $BackendConfiguration
    $config.PluginConfiguration = $PluginConfiguration
    $config.ModuleManifestsToLoad = $ModuleManifestsToLoad
    $config.MuteUnknownCommand = $MuteUnknownCommand
    $config.PluginDirectory = $PluginDirectory
    $config.PluginRepository = $PluginRepository
    $config.SendCommandResponseToPrivate = $SendCommandResponseToPrivate
    $config.AddCommandReactions = $AddCommandReactions
    $config.ApprovalConfiguration.ExpireMinutes = $ApprovalExpireMinutes
    $config.DisallowDMs = ($DisallowDMs -eq $true)
    $config.FormatEnumerationLimitOverride = $FormatEnumerationLimitOverride
    if ($ChannelRules.Count -ge 1) {
        $config.ChannelRules = $null
        foreach ($item in $ChannelRules) {
            $config.ChannelRules += [ChannelRule]::new($item.Channel, $item.IncludeCommands, $item.ExcludeCommands)
        }
    }
    if ($ApprovalCommandConfigurations.Count -ge 1) {
        foreach ($item in $ApprovalCommandConfigurations) {
            $acc = [ApprovalCommandConfiguration]::new()
            $acc.Expression = $item.Expression
            $acc.ApprovalGroups = $item.Groups
            $acc.PeerApproval = $item.PeerApproval
            $config.ApprovalConfiguration.Commands.Add($acc) > $null
        }
    }

    # Add any middleware hooks
    foreach ($type in [enum]::GetNames([MiddlewareType])) {
        foreach ($item in $PSBoundParameters["$($type)MiddlewareHooks"]) {
            $config.MiddlewareConfiguration.Add($item, $type)
        }
    }

    $config
}

Export-ModuleMember -Function 'New-PoshBotConfiguration'


function New-PoshBotFileUpload {
    <#
    .SYNOPSIS
        Tells PoshBot to upload a file to the chat network.
    .DESCRIPTION
        Returns a custom object back to PoshBot telling it to upload the given file to the chat network. The custom object
        can also tell PoshBot to redirect the file upload to a DM channel with the calling user. This could be useful if
        the contents the bot command returns are sensitive and should not be visible to all users in the channel.
    .PARAMETER Path
        The path(s) to one or more files to upload. Wildcards are permitted.
    .PARAMETER LiteralPath
        Specifies the path(s) to the current location of the file(s). Unlike the Path parameter, the value of LiteralPath is used exactly as it is typed.
        No characters are interpreted as wildcards. If the path includes escape characters, enclose it in single quotation marks. Single quotation
        marks tell PowerShell not to interpret any characters as escape sequences.
    .PARAMETER Content
        The content of the file to send.
    .PARAMETER FileType
        If specified, override the file type determined by the filename.
    .PARAMETER FileName
        The name to call the uploaded file
    .PARAMETER Title
        The title for the uploaded file.
    .PARAMETER DM
        Tell PoshBot to redirect the file upload to a DM channel.
    .PARAMETER KeepFile
        If specified, keep the source file after calling Send-SlackFile. The source file is deleted without this
    .EXAMPLE
        function Do-Stuff {
            [cmdletbinding()]
            param()

            $myObj = [pscustomobject]@{
                value1 = 'foo'
                value2 = 'bar'
            }

            $csv = Join-Path -Path $env:TEMP -ChildPath "$((New-Guid).ToString()).csv"
            $myObj | Export-Csv -Path $csv -NoTypeInformation

            New-PoshBotFileUpload -Path $csv
        }

        Export a CSV file and tell PoshBot to upload the file back to the channel that initiated this command.

    .EXAMPLE
        function Get-SecretPlan {
            [cmdletbinding()]
            param()

            $myObj = [pscustomobject]@{
                Title = 'Secret moon base'
                Description = 'Plans for secret base on the dark side of the moon'
            }

            $csv = Join-Path -Path $env:TEMP -ChildPath "$((New-Guid).ToString()).csv"
            $myObj | Export-Csv -Path $csv -NoTypeInformation

            New-PoshBotFileUpload -Path $csv -Title 'YourEyesOnly.csv' -DM
        }

        Export a CSV file and tell PoshBot to upload the file back to a DM channel with the calling user.

    .EXAMPLE
        function Do-Stuff {
            [cmdletbinding()]
            param()

            $myObj = [pscustomobject]@{
                value1 = 'foo'
                value2 = 'bar'
            }

            $csv = Join-Path -Path $env:TEMP -ChildPath "$((New-Guid).ToString()).csv"
            $myObj | Export-Csv -Path $csv -NoTypeInformation

            New-PoshBotFileUpload -Path $csv -KeepFile
        }

        Export a CSV file and tell PoshBot to upload the file back to the channel that initiated this command.
        Keep the file after uploading it.
    .INPUTS
        String
    .OUTPUTS
        PSCustomObject
    .LINK
        New-PoshBotCardResponse
    .LINK
        New-PoshBotTextResponse
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='*')]
    [cmdletbinding(DefaultParameterSetName = 'Path')]
    param(
        [parameter(
            Mandatory,
            ParameterSetName  = 'Path',
            Position = 0,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName
        )]
        [ValidateNotNullOrEmpty()]
        [SupportsWildcards()]
        [string[]]$Path,

        [parameter(
            Mandatory,
            ParameterSetName = 'LiteralPath',
            Position = 0,
            ValueFromPipelineByPropertyName
        )]
        [ValidateNotNullOrEmpty()]
        [Alias('PSPath')]
        [string[]]$LiteralPath,

        [parameter(
            Mandatory,
            ParameterSetName = 'Content')]
        [string]$Content,

        [parameter(
            ParameterSetName = 'Content'
        )]
        [string]$FileType,

        [parameter(
            ParameterSetName = 'Content'
        )]
        [string]$FileName,

        [string]$Title = [string]::Empty,

        [switch]$DM,

        [switch]$KeepFile
    )

    process {
        if ($PSCmdlet.ParameterSetName -eq 'Content') {
            [pscustomobject][ordered]@{
                PSTypeName = 'PoshBot.File.Upload'
                Content    = $Content
                FileName   = $FileName
                FileType   = $FileType
                Title      = $Title
                DM         = $DM.IsPresent
                KeepFile   = $KeepFile.IsPresent
            }
        } else {
            # Resolve path(s)
            if ($PSCmdlet.ParameterSetName -eq 'Path') {
                $paths = Resolve-Path -Path $Path | Select-Object -ExpandProperty Path
            } elseIf ($PSCmdlet.ParameterSetName -eq 'LiteralPath') {
                $paths = Resolve-Path -LiteralPath $LiteralPath | Select-Object -ExpandProperty Path
            }

            foreach ($item in $paths) {
                [pscustomobject][ordered]@{
                    PSTypeName = 'PoshBot.File.Upload'
                    Path       = $item
                    Title      = $Title
                    DM         = $DM.IsPresent
                    KeepFile   = $KeepFile.IsPresent
                }
            }
        }
    }
}

Export-ModuleMember -Function 'New-PoshBotFileUpload'


function New-PoshBotInstance {
    <#
    .SYNOPSIS
        Creates a new instance of PoshBot
    .DESCRIPTION
        Creates a new instance of PoshBot from an existing configuration (.psd1) file or a configuration object.
    .PARAMETER Configuration
        The bot configuration object to create a new instance from.
    .PARAMETER Path
        The path to a PowerShell data (.psd1) file to create a new instance from.
    .PARAMETER LiteralPath
        Specifies the path(s) to the current location of the file(s). Unlike the Path parameter, the value of LiteralPath is used exactly as it is typed.
        No characters are interpreted as wildcards. If the path includes escape characters, enclose it in single quotation marks. Single quotation
        marks tell PowerShell not to interpret any characters as escape sequences.
    .PARAMETER Backend
        The backend object that hosts logic for receiving and sending messages to a chat network.
    .EXAMPLE
        PS C:\> New-PoshBotInstance -Path 'C:\Users\joeuser\.poshbot\Cherry2000.psd1' -Backend $backend

        Name          : Cherry2000
        Backend       : SlackBackend
        Storage       : StorageProvider
        PluginManager : PluginManager
        RoleManager   : RoleManager
        Executor      : CommandExecutor
        MessageQueue  : {}
        Configuration : BotConfiguration

        Create a new PoshBot instance from configuration file [C:\Users\joeuser\.poshbot\Cherry2000.psd1] and Slack backend object [$backend].
    .EXAMPLE
        PS C:\> $botConfig = Get-PoshBotConfiguration -Path (Join-Path -Path $env:USERPROFILE -ChildPath '.poshbot\Cherry2000.psd1')
        PS C:\> $backend = New-PoshBotSlackBackend -Configuration $botConfig.BackendConfiguration
        PS C:\> $myBot = $botConfig | New-PoshBotInstance -Backend $backend
        PS C:\> $myBot | Format-List

        Name          : Cherry2000
        Backend       : SlackBackend
        Storage       : StorageProvider
        PluginManager : PluginManager
        RoleManager   : RoleManager
        Executor      : CommandExecutor
        MessageQueue  : {}
        Configuration : BotConfiguration

        Gets a bot configuration from the filesytem, creates a chat backend object, and then creates a new bot instance.
    .EXAMPLE
        PS C:\> $botConfig = Get-PoshBotConfiguration -Path (Join-Path -Path $env:USERPROFILE -ChildPath '.poshbot\Cherry2000.psd1')
        PS C:\> $backend = $botConfig | New-PoshBotSlackBackend
        PS C:\> $myBotJob = $botConfig | New-PoshBotInstance -Backend $backend | Start-PoshBot -AsJob -PassThru

        Gets a bot configuration, creates a Slack backend from it, then creates a new PoshBot instance and starts it as a background job.
    .INPUTS
        String
    .INPUTS
        BotConfiguration
    .OUTPUTS
        Bot
    .LINK
        Get-PoshBotConfiguration
    .LINK
        New-PoshBotSlackBackend
    .LINK
        Start-PoshBot
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='*')]
    [cmdletbinding(DefaultParameterSetName = 'path')]
    param(
        [parameter(
            Mandatory,
            ParameterSetName  = 'Path',
            Position = 0,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName
        )]
        [ValidateNotNullOrEmpty()]
        [SupportsWildcards()]
        [string[]]$Path,

        [parameter(
            Mandatory,
            ParameterSetName = 'LiteralPath',
            Position = 0,
            ValueFromPipelineByPropertyName
        )]
        [ValidateNotNullOrEmpty()]
        [Alias('PSPath')]
        [string[]]$LiteralPath,

        [parameter(
            Mandatory,
            ParameterSetName = 'config',
            ValueFromPipeline,
            ValueFromPipelineByPropertyName
        )]
        [BotConfiguration[]]$Configuration,

        [parameter(Mandatory)]
        [Backend]$Backend
    )

    begin {
        $here = $PSScriptRoot
    }

    process {
        if ($PSCmdlet.ParameterSetName -eq 'path' -or $PSCmdlet.ParameterSetName -eq 'LiteralPath') {
            # Resolve path(s)
            if ($PSCmdlet.ParameterSetName -eq 'Path') {
                $paths = Resolve-Path -Path $Path | Select-Object -ExpandProperty Path
            } elseif ($PSCmdlet.ParameterSetName -eq 'LiteralPath') {
                $paths = Resolve-Path -LiteralPath $LiteralPath | Select-Object -ExpandProperty Path
            }

            $Configuration = @()
            foreach ($item in $paths) {
                if (Test-Path $item) {
                    if ( (Get-Item -Path $item).Extension -eq '.psd1') {
                        $Configuration += Get-PoshBotConfiguration -Path $item
                    } else {
                        Throw 'Path must be to a valid .psd1 file'
                    }
                } else {
                    Write-Error -Message "Path [$item] is not valid."
                }
            }
        }

        foreach ($config in $Configuration) {
            Write-Verbose -Message "Creating bot instance with name [$($config.Name)]"
            [Bot]::new($Backend, $here, $config)
        }
    }
}

Export-ModuleMember -Function 'New-PoshBotInstance'


function New-PoshBotMiddlewareHook {
    <#
    .SYNOPSIS
        Creates a PoshBot middleware hook object.
    .DESCRIPTION
        PoshBot can execute custom scripts during various stages of the command processing lifecycle. These scripts
        are defined using New-PoshBotMiddlewareHook and added to the bot configuration object under the MiddlewareConfiguration section.
        Hooks are added to the PreReceive, PostReceive, PreExecute, PostExecute, PreResponse, and PostResponse properties.
        Middleware gets executed in the order in which it is added under each property.
    .PARAMETER Name
        The name of the middleware hook. Must be unique in each middleware lifecycle stage.
    .PARAMETER Path
        The file path the the PowerShell script to execute as a middleware hook.
    .EXAMPLE
        PS C:\> $userDropHook = New-PoshBotMiddlewareHook -Name 'dropuser' -Path 'c:/poshbot/middleware/dropuser.ps1'
        PS C:\> $config.MiddlewareConfiguration.Add($userDropHook, 'PreReceive')

        Creates a middleware hook called 'dropuser' and adds it to the 'PreReceive' middleware lifecycle stage.
    .OUTPUTS
        MiddlewareHook
    #>
    [cmdletbinding()]
    param(
        [parameter(mandatory)]
        [string]$Name,

        [parameter(mandatory)]
        [ValidateScript({
            if (-not (Test-Path -Path $_)) {
                throw 'Invalid script path'
            } else {
                $true
            }
        })]
        [string]$Path
    )

    [MiddlewareHook]::new($Name, $Path)
}

Export-ModuleMember -Function 'New-PoshBotMiddlewareHook'


function New-PoshBotScheduledTask {
    <#
    .SYNOPSIS
        Creates a new scheduled task to run PoshBot in the background.
    .DESCRIPTION
        Creates a new scheduled task to run PoshBot in the background. The scheduled task will always be configured
        to run on startup and to not stop after any time period.
    .PARAMETER Name
        The name for the scheduled task
    .PARAMETER Description
        The description for the scheduled task
    .PARAMETER Path
        The path to the PoshBot configuration file to load and execute
    .PARAMETER Credential
        The credential to run the scheduled task under.
    .PARAMETER PassThru
        Return the newly created scheduled task object
    .PARAMETER Force
        Overwrite a previously created scheduled task
    .EXAMPLE
        PS C:\> $cred = Get-Credential
        PS C:\> New-PoshBotScheduledTask -Name PoshBot -Path C:\PoshBot\myconfig.psd1 -Credential $cred

        Creates a new scheduled task to start PoshBot using the configuration file located at C:\PoshBot\myconfig.psd1
        and the specified credential.
    .EXAMPLE
        PS C:\> $cred = Get-Credential
        PC C:\> $params = @{
            Name = 'PoshBot'
            Path = 'C:\PoshBot\myconfig.psd1'
            Credential = $cred
            Description = 'Awesome ChatOps bot'
            PassThru = $true
        }
        PS C:\> $task = New-PoshBotScheduledTask @params
        PS C:\> $task | Start-ScheduledTask

        Creates a new scheduled task to start PoshBot using the configuration file located at C:\PoshBot\myconfig.psd1
        and the specified credential then starts the task.
    .OUTPUTS
        Microsoft.Management.Infrastructure.CimInstance#root/Microsoft/Windows/TaskScheduler/MSFT_ScheduledTask
    .LINK
        Get-PoshBotConfiguration
    .LINK
        New-PoshBotConfiguration
    .LINK
        Save-PoshBotConfiguration
    .LINK
        Start-PoshBot
    #>
    [cmdletbinding(SupportsShouldProcess)]
    param(
        [string]$Name = 'PoshBot',

        [string]$Description = 'Start PoshBot',

        [parameter(Mandatory)]
        [ValidateScript({
            if (Test-Path -Path $_) {
                if ( (Get-Item -Path $_).Extension -eq '.psd1') {
                    $true
                } else {
                    Throw 'Path must be to a valid .psd1 file'
                }
            } else {
                Throw 'Path is not valid'
            }
        })]
        [string]$Path,

        [parameter(Mandatory)]
        [pscredential]
        [System.Management.Automation.CredentialAttribute()]
        $Credential,

        [switch]$PassThru,

        [switch]$Force
    )

    if ($Force -or (-not (Get-ScheduledTask -TaskName $Name -ErrorAction SilentlyContinue))) {
        if ($PSCmdlet.ShouldProcess($Name, 'Created PoshBot scheduled task')) {

            $taskParams = @{
                Description = $Description
            }

            # Determine path to scheduled task script
            # Not adding '..\' to -ChildPath parameter because during module build
            # this script will get merged into PoshBot.psm1 and \Task folder will be
            # a direct child of $PSScriptRoot
            $startScript = Resolve-Path -LiteralPath (Join-Path -Path $PSScriptRoot -ChildPath 'Task\StartPoshBot.ps1')

            # Scheduled task action
            $arg = "& '$startScript' -Path '$Path'"
            $actionParams = @{
                Execute = "$($env:SystemDrive)\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
                Argument = '-ExecutionPolicy Bypass -NonInteractive -Command "' + $arg + '"'
                WorkingDirectory = $PSScriptRoot
            }
            $taskParams.Action = New-ScheduledTaskAction @actionParams

            # Scheduled task at logon trigger
            $taskParams.Trigger = New-ScheduledTaskTrigger -AtStartup

            # Scheduled task settings
            $settingsParams = @{
                AllowStartIfOnBatteries = $true
                DontStopIfGoingOnBatteries = $true
                ExecutionTimeLimit = 0
                RestartCount = 999
                RestartInterval = (New-TimeSpan -Minutes 1)
            }
            $taskParams.Settings = New-ScheduledTaskSettingsSet @settingsParams

            # Create / register the task
            $registerParams = @{
                TaskName = $Name
                Force = $true
            }
            # Scheduled task principal
            $registerParams.User = $Credential.UserName
            $registerParams.Password = $Credential.GetNetworkCredential().Password
            $task = New-ScheduledTask @taskParams
            $newTask = Register-ScheduledTask -InputObject $task @registerParams
            if ($PassThru) {
                $newTask
            }
        }
    } else {
        Write-Error -Message "Existing task named [$Name] found. To overwrite, use the -Force"
    }
}

Export-ModuleMember -Function 'New-PoshBotScheduledTask'


function New-PoshBotTextResponse {
    <#
    .SYNOPSIS
        Tells PoshBot to handle the text response from a command in a special way.
    .DESCRIPTION
        Responses from PoshBot commands can be sent back to the channel they were posted from (default) or redirected to a DM channel with the
        calling user. This could be useful if the contents the bot command returns are sensitive and should not be visible to all users
        in the channel.
    .PARAMETER Text
        The text response from the command.
    .PARAMETER AsCode
        Format the text in a code block if the backend supports it.
    .PARAMETER DM
        Tell PoshBot to redirect the response to a DM channel.
    .EXAMPLE
        function Get-Foo {
            [cmdletbinding()]
            param(
                [parameter(mandatory)]
                [string]$MyParam
            )

            New-PoshBotTextResponse -Text $MyParam -DM
        }

        When Get-Foo is executed by PoshBot, the text response will be sent back to the calling user as a DM rather than back in the channel the
        command was called from. This could be useful if the contents the bot command returns are sensitive and should not be visible to all users
        in the channel.
    .INPUTS
        String
    .OUTPUTS
        PSCustomObject
    .LINK
        New-PoshBotCardResponse
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='*')]
    [cmdletbinding()]
    param(
        [parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string[]]$Text,

        [switch]$AsCode,

        [switch]$DM
    )

    process {
        foreach ($item in $text) {
            [pscustomobject][ordered]@{
                PSTypeName = 'PoshBot.Text.Response'
                Text = $item.Trim()
                AsCode = $PSBoundParameters.ContainsKey('AsCode')
                DM = $PSBoundParameters.ContainsKey('DM')
            }
        }
    }
}

Export-ModuleMember -Function 'New-PoshBotTextResponse'


function Remove-PoshBotStatefulData {
    <#
    .SYNOPSIS
        Remove existing stateful data
    .DESCRIPTION
        Remove existing stateful data
    .PARAMETER Name
        Property to remove from the stateful data file
    .PARAMETER Scope
        Sets the scope of stateful data to remove:
            Module: Remove stateful data from the current module's data
            Global: Remove stateful data from the global PoshBot data
    .PARAMETER Depth
        Specifies how many levels of contained objects are included in the XML representation. The default value is 2
    .EXAMPLE
        PS C:\> Remove-PoshBotStatefulData -Name 'ToUse'

        Removes the 'ToUse' property from stateful data for the PoshBot plugin you are currently running this from.
    .EXAMPLE
        PS C:\> Remove-PoshBotStatefulData -Name 'Something' -Scope Global

        Removes the 'Something' property from PoshBot's global stateful data
    .LINK
        Get-PoshBotStatefulData
    .LINK
        Set-PoshBotStatefulData
    .LINK
        Start-PoshBot
    #>
    [cmdletbinding(SupportsShouldProcess)]
    param(
        [parameter(Mandatory)]
        [string[]]$Name,

        [validateset('Global','Module')]
        [string]$Scope = 'Module',

        [int]$Depth = 2
    )
    process {
        if($Scope -eq 'Module') {
            $FileName = "$($global:PoshBotContext.Plugin).state"
        } else {
            $FileName = "PoshbotGlobal.state"
        }
        $Path = Join-Path $global:PoshBotContext.ConfigurationDirectory $FileName


        if(-not (Test-Path $Path)) {
            return
        } else {
            $ToWrite = Import-Clixml -Path $Path | Select-Object * -ExcludeProperty $Name
        }

        if ($PSCmdlet.ShouldProcess($Name, 'Remove stateful data')) {
            Export-Clixml -Path $Path -InputObject $ToWrite -Depth $Depth -Force
            Write-Verbose -Message "Stateful data [$Name] removed from [$Path]"
        }
    }
}

Export-ModuleMember -Function 'Remove-PoshBotStatefulData'


function Save-PoshBotConfiguration {
    <#
    .SYNOPSIS
        Saves a PoshBot configuration object to the filesystem in the form of a PowerShell data (.psd1) file.
    .DESCRIPTION
        PoshBot configurations can be stored on the filesytem in PowerShell data (.psd1) files.
        This function will save a previously created configuration object to the filesystem.
    .PARAMETER InputObject
        The bot configuration object to save to the filesystem.
    .PARAMETER Path
        The path to a PowerShell data (.psd1) file to save the configuration to.
    .PARAMETER Force
        Overwrites an existing configuration file.
    .PARAMETER PassThru
        Returns the configuration file path.
    .EXAMPLE
        PS C:\> Save-PoshBotConfiguration -InputObject $botConfig

        Saves the PoshBot configuration. If now -Path is specified, the configuration will be saved to $env:USERPROFILE\.poshbot\PoshBot.psd1.
    .EXAMPLE
        PS C:\> $botConfig | Save-PoshBotConfig -Path c:\mybot\mybot.psd1

        Saves the PoshBot configuration to [c:\mybot\mybot.psd1].
    .EXAMPLE
        PS C:\> $configFile = $botConfig | Save-PoshBotConfig -Path c:\mybot\mybot.psd1 -Force -PassThru

        Saves the PoshBot configuration to [c:\mybot\mybot.psd1] and Overwrites existing file. The new file will be returned.
    .INPUTS
        BotConfiguration
    .OUTPUTS
        System.IO.FileInfo
    .LINK
        Get-PoshBotConfiguration
    .LINK
        Start-PoshBot
    #>
    [cmdletbinding(SupportsShouldProcess)]
    param(
        [parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Configuration')]
        [BotConfiguration]$InputObject,

        [string]$Path = (Join-Path -Path $script:defaultPoshBotDir -ChildPath 'PoshBot.psd1'),

        [switch]$Force,

        [switch]$PassThru
    )

    process {
        if ($PSCmdlet.ShouldProcess($Path, 'Save PoshBot configuration')) {
            $hash = @{}
            foreach ($prop in ($InputObject | Get-Member -MemberType Property)) {
                switch ($prop.Name) {
                    # Serialize ChannelRules, ApprovalConfiguration, and MiddlewareConfiguration propertes differently as
                    # ConvertTo-Metadata won't know how to do it since they're custom PoshBot classes
                    'ChannelRules' {
                        $hash.Add($prop.Name, $InputObject.($prop.Name).ToHash())
                        break
                    }
                    'ApprovalConfiguration' {
                        $hash.Add($prop.Name, $InputObject.($prop.Name).ToHash())
                        break
                    }
                    'MiddlewareConfiguration' {
                        $hash.Add($prop.Name, $InputObject.($prop.Name).ToHash())
                        break
                    }
                    Default {
                        $hash.Add($prop.Name, $InputObject.($prop.Name))
                        break
                    }
                }
            }

            $meta = $hash | ConvertTo-Metadata -WarningAction SilentlyContinue
            if (-not (Test-Path -Path $Path) -or $Force) {
                New-Item -Path $Path -ItemType File -Force | Out-Null

                $meta | Out-file -FilePath $Path -Force -Encoding utf8
                Write-Verbose -Message "PoshBot configuration saved to [$Path]"

                if ($PassThru) {
                    Get-Item -Path $Path | Select-Object -First 1
                }
            } else {
                Write-Error -Message 'File already exists. Use the -Force switch to overwrite the file.'
            }
        }
    }
}

Export-ModuleMember -Function 'Save-PoshBotConfiguration'


function Set-PoshBotStatefulData {
    <#
    .SYNOPSIS
        Save stateful data to use in another PoshBot command
    .DESCRIPTION
        Save stateful data to use in another PoshBot command

        Stores data in clixml format, in the PoshBot ConfigurationDirectory.

        If <Name> property exists in current stateful data file, it is overwritten
    .PARAMETER Name
        Property to add to the stateful data file
    .PARAMETER Value
        Value to set for the Name property in the stateful data file
    .PARAMETER Scope
        Sets the scope of stateful data to set:
            Module: Allow only this plugin to access the stateful data you save
            Global: Allow any plugin to access the stateful data you save
    .PARAMETER Depth
        Specifies how many levels of contained objects are included in the XML representation. The default value is 2
    .EXAMPLE
        PS C:\> Set-PoshBotStatefulData -Name 'ToUse' -Value 'Later'

        Adds a 'ToUse' property to the stateful data for the PoshBot plugin you are currently running this from.
    .EXAMPLE
        PS C:\> $Anything | Set-PoshBotStatefulData -Name 'Something' -Scope Global

        Adds a 'Something' property to PoshBot's global stateful data, with the value of $Anything
    .LINK
        Get-PoshBotStatefulData
    .LINK
        Remove-PoshBotStatefulData
    .LINK
        Start-PoshBot
    #>
    [cmdletbinding(SupportsShouldProcess)]
    param(
        [parameter(Mandatory)]
        [string]$Name,

        [parameter(ValueFromPipeline,
                   Mandatory)]
        [object[]]$Value,

        [validateset('Global','Module')]
        [string]$Scope = 'Module',

        [int]$Depth = 2
    )

    end {
        if ($Value.Count -eq 1) {
            $Value = $Value[0]
        }

        if($Scope -eq 'Module') {
            $FileName = "$($global:PoshBotContext.Plugin).state"
        } else {
            $FileName = "PoshbotGlobal.state"
        }
        $Path = Join-Path $global:PoshBotContext.ConfigurationDirectory $FileName

        if(-not (Test-Path $Path)) {
            $ToWrite = [pscustomobject]@{
                $Name = $Value
            }
        } else {
            $Existing = Import-Clixml -Path $Path
            # TODO: Consider handling for -Force?
            If($Existing.PSObject.Properties.Name -contains $Name) {
                Write-Verbose "Overwriting [$Name]`nCurrent value: [$($Existing.$Name | Out-String)])`nNew Value: [$($Value | Out-String)]"
            }
            Add-Member -InputObject $Existing -MemberType NoteProperty -Name $Name -Value $Value -Force
            $ToWrite = $Existing
        }

        if ($PSCmdlet.ShouldProcess($Name, 'Set stateful data')) {
            Export-Clixml -Path $Path -InputObject $ToWrite -Depth $Depth -Force
            Write-Verbose -Message "Stateful data [$Name] saved to [$Path]"
        }
    }
}

Export-ModuleMember -Function 'Set-PoshBotStatefulData'


function Start-PoshBot {
    <#
    .SYNOPSIS
        Starts a new instance of PoshBot interactively or in a job.
    .DESCRIPTION
        Starts a new instance of PoshBot interactively or in a job.
    .PARAMETER InputObject
        An existing PoshBot instance to start.
    .PARAMETER Configuration
        A PoshBot configuration object to use to start the bot instance.
    .PARAMETER Path
        The path to a PoshBot configuration file.
        A new instance of PoshBot will be created from this file.
    .PARAMETER AsJob
        Run the PoshBot instance in a background job.
    .PARAMETER PassThru
        Return the PoshBot instance Id that is running as a job.
    .EXAMPLE
        PS C:\> Start-PoshBot -Bot $bot

        Runs an instance of PoshBot that has already been created interactively in the shell.
    .EXAMPLE
        PS C:\> $bot | Start-PoshBot -Verbose

        Runs an instance of PoshBot that has already been created interactively in the shell.
    .EXAMPLE
        PS C:\> $config = Get-PoshBotConfiguration -Path (Join-Path -Path $env:USERPROFILE -ChildPath '.poshbot\MyPoshBot.psd1')
        PS C:\> Start-PoshBot -Config $config

        Gets a PoshBot configuration from file and starts the bot interactively.
    .EXAMPLE
        PS C:\> Get-PoshBot -Id 100

        Id         : 100
        Name       : PoshBot_eab96f2ad147489b9f90e110e02ad805
        State      : Running
        InstanceId : eab96f2ad147489b9f90e110e02ad805
        Config     : BotConfiguration

        Gets the PoshBot job instance with ID 100.
    .INPUTS
        Bot
    .INPUTS
        BotConfiguration
    .INPUTS
        String
    .OUTPUTS
        PSCustomObject
    .LINK
        Start-PoshBot
    .LINK
        Stop-PoshBot
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='*')]
    [cmdletbinding(DefaultParameterSetName = 'bot')]
    param(
        [parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'bot')]
        [Alias('Bot')]
        [Bot]$InputObject,

        [parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'config')]
        [BotConfiguration]$Configuration,

        [parameter(Mandatory, ParameterSetName = 'path')]
        [string]$Path,

        [switch]$AsJob,

        [switch]$PassThru
    )

    process {
        try {
            switch ($PSCmdlet.ParameterSetName) {
                'bot' {
                    $bot = $InputObject
                    $Configuration = $bot.Configuration
                }
                'config' {
                    $backend = New-PoshBotSlackBackend -Configuration $Configuration.BackendConfiguration
                    $bot = New-PoshBotInstance -Backend $backend -Configuration $Configuration
                }
                'path' {
                    $Configuration = Get-PoshBotConfiguration -Path $Path
                    $backend = New-PoshBotSlackBackend -Configuration $Configuration.BackendConfiguration
                    $bot = New-PoshBotInstance -Backend $backend -Configuration $Configuration
                }
            }

            if ($AsJob) {
                $sb = {
                    param(
                        [parameter(Mandatory)]
                        [hashtable]$Configuration,
                        [string]$PoshBotManifestPath
                    )

                    Import-Module $PoshBotManifestPath -ErrorAction Stop

                    try {
                        $tempConfig = New-PoshBotConfiguration
                        $realConfig = $tempConfig.Serialize($Configuration)

                        while ($true) {
                            try {
                                if ($realConfig.BackendConfiguration.Name -in @('Slack', 'SlackBackend')) {
                                    $backend = New-PoshBotSlackBackend -Configuration $realConfig.BackendConfiguration
                                } elseIf ($realConfig.BackendConfiguration.Name -in @('Teams', 'TeamsBackend')) {
                                    $backend = New-PoshBotTeamsBackend -Configuration $realConfig.BackendConfiguration
                                } else {
                                    Write-Error "Unable to determine backend type. Name property in BackendConfiguration should have a value of 'Slack', 'SlackBackend', 'Teams', or 'TeamsBackend'"
                                    break
                                }

                                $bot = New-PoshBotInstance -Backend $backend -Configuration $realConfig
                                $bot.Start()
                            } catch {
                                Write-Error $_
                                Write-Error 'PoshBot crashed :( Restarting...'
                                Start-Sleep -Seconds 5
                            }
                        }
                    } catch {
                        throw $_
                    }
                }

                $instanceId = (New-Guid).ToString().Replace('-', '')
                $jobName = "PoshBot_$instanceId"
                $poshBotManifestPath = (Join-Path -Path $PSScriptRoot -ChildPath "PoshBot.psd1")

                $job = Start-Job -ScriptBlock $sb -Name $jobName -ArgumentList $Configuration.ToHash(),$poshBotManifestPath

                # Track the bot instance
                $botTracker = @{
                    JobId = $job.Id
                    Name = $jobName
                    InstanceId = $instanceId
                    Config = $Configuration
                }
                $script:botTracker.Add($job.Id, $botTracker)

                if ($PSBoundParameters.ContainsKey('PassThru')) {
                    Get-PoshBot -Id $job.Id
                }
            } else {
                $bot.Start()
            }
        } catch {
            throw $_
        }
        finally {
            if (-not $AsJob) {
                # We're here because CTRL+C was entered.
                # Make sure to disconnect the bot from the backend chat network
                if ($bot) {
                    Write-Verbose -Message 'Stopping PoshBot'
                    $bot.Disconnect()
                }
            }
        }
    }
}

Export-ModuleMember -Function 'Start-Poshbot'


function Stop-Poshbot {
    <#
    .SYNOPSIS
        Stop a currently running PoshBot instance that is running as a background job.
    .DESCRIPTION
        PoshBot can be run in the background with PowerShell jobs. This function stops
        a currently running PoshBot instance.
    .PARAMETER Id
        The job Id of the bot to stop.
    .PARAMETER Force
        Stop PoshBot instance without prompt
    .EXAMPLE
        Stop-PoshBot -Id 101

        Stop the bot instance with Id 101.
    .EXAMPLE
        Get-PoshBot | Stop-PoshBot

        Gets all running PoshBot instances and stops them.
    .INPUTS
        System.Int32
    .LINK
        Get-PoshBot
    .LINK
        Start-PoshBot
    #>
    [cmdletbinding(SupportsShouldProcess, ConfirmImpact = 'high')]
    param(
        [parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [int[]]$Id,

        [switch]$Force
    )

    begin {
        $remove = @()
    }

    process {
        foreach ($jobId in $Id) {
            if ($Force -or $PSCmdlet.ShouldProcess($jobId, 'Stop PoshBot')) {
                $bot = $script:botTracker[$jobId]
                if ($bot) {
                    Write-Verbose -Message "Stopping PoshBot Id: $jobId"
                    Stop-Job -Id $jobId -Verbose:$false
                    Remove-Job -Id $JobId -Verbose:$false
                    $remove += $jobId
                } else {
                    throw "Unable to find PoshBot instance with Id [$Id]"
                }
            }
        }
    }

    end {
        # Remove this bot from tracking
        $remove | ForEach-Object {
            $script:botTracker.Remove($_)
        }
    }
}

Export-ModuleMember -Function 'Stop-Poshbot'


[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '', Scope='Class', Target='*')]
class SlackBackend : Backend {

    # The types of message that we care about from Slack
    # All othere will be ignored
    [string[]]$MessageTypes = @(
        'channel_rename'
        'member_joined_channel'
        'member_left_channel'
        'message'
        'pin_added'
        'pin_removed'
        'presence_change'
        'reaction_added'
        'reaction_removed'
        'star_added'
        'star_removed'
    )

    [int]$MaxMessageLength = 3900

    # Import some color defs.
    hidden [hashtable]$_PSSlackColorMap = @{
        aliceblue = "#F0F8FF"
        antiquewhite = "#FAEBD7"
        aqua = "#00FFFF"
        aquamarine = "#7FFFD4"
        azure = "#F0FFFF"
        beige = "#F5F5DC"
        bisque = "#FFE4C4"
        black = "#000000"
        blanchedalmond = "#FFEBCD"
        blue = "#0000FF"
        blueviolet = "#8A2BE2"
        brown = "#A52A2A"
        burlywood = "#DEB887"
        cadetblue = "#5F9EA0"
        chartreuse = "#7FFF00"
        chocolate = "#D2691E"
        coral = "#FF7F50"
        cornflowerblue = "#6495ED"
        cornsilk = "#FFF8DC"
        crimson = "#DC143C"
        darkblue = "#00008B"
        darkcyan = "#008B8B"
        darkgoldenrod = "#B8860B"
        darkgray = "#A9A9A9"
        darkgreen = "#006400"
        darkkhaki = "#BDB76B"
        darkmagenta = "#8B008B"
        darkolivegreen = "#556B2F"
        darkorange = "#FF8C00"
        darkorchid = "#9932CC"
        darkred = "#8B0000"
        darksalmon = "#E9967A"
        darkseagreen = "#8FBC8F"
        darkslateblue = "#483D8B"
        darkslategray = "#2F4F4F"
        darkturquoise = "#00CED1"
        darkviolet = "#9400D3"
        deeppink = "#FF1493"
        deepskyblue = "#00BFFF"
        dimgray = "#696969"
        dodgerblue = "#1E90FF"
        firebrick = "#B22222"
        floralwhite = "#FFFAF0"
        forestgreen = "#228B22"
        fuchsia = "#FF00FF"
        gainsboro = "#DCDCDC"
        ghostwhite = "#F8F8FF"
        gold = "#FFD700"
        goldenrod = "#DAA520"
        gray = "#808080"
        green = "#008000"
        greenyellow = "#ADFF2F"
        honeydew = "#F0FFF0"
        hotpink = "#FF69B4"
        indianred = "#CD5C5C"
        indigo = "#4B0082"
        ivory = "#FFFFF0"
        khaki = "#F0E68C"
        lavender = "#E6E6FA"
        lavenderblush = "#FFF0F5"
        lawngreen = "#7CFC00"
        lemonchiffon = "#FFFACD"
        lightblue = "#ADD8E6"
        lightcoral = "#F08080"
        lightcyan = "#E0FFFF"
        lightgoldenrodyellow = "#FAFAD2"
        lightgreen = "#90EE90"
        lightgrey = "#D3D3D3"
        lightpink = "#FFB6C1"
        lightsalmon = "#FFA07A"
        lightseagreen = "#20B2AA"
        lightskyblue = "#87CEFA"
        lightslategray = "#778899"
        lightsteelblue = "#B0C4DE"
        lightyellow = "#FFFFE0"
        lime = "#00FF00"
        limegreen = "#32CD32"
        linen = "#FAF0E6"
        maroon = "#800000"
        mediumaquamarine = "#66CDAA"
        mediumblue = "#0000CD"
        mediumorchid = "#BA55D3"
        mediumpurple = "#9370DB"
        mediumseagreen = "#3CB371"
        mediumslateblue = "#7B68EE"
        mediumspringgreen = "#00FA9A"
        mediumturquoise = "#48D1CC"
        mediumvioletred = "#C71585"
        midnightblue = "#191970"
        mintcream = "#F5FFFA"
        mistyrose = "#FFE4E1"
        moccasin = "#FFE4B5"
        navajowhite = "#FFDEAD"
        navy = "#000080"
        oldlace = "#FDF5E6"
        olive = "#808000"
        olivedrab = "#6B8E23"
        orange = "#FFA500"
        orangered = "#FF4500"
        orchid = "#DA70D6"
        palegoldenrod = "#EEE8AA"
        palegreen = "#98FB98"
        paleturquoise = "#AFEEEE"
        palevioletred = "#DB7093"
        papayawhip = "#FFEFD5"
        peachpuff = "#FFDAB9"
        peru = "#CD853F"
        pink = "#FFC0CB"
        plum = "#DDA0DD"
        powderblue = "#B0E0E6"
        purple = "#800080"
        red = "#FF0000"
        rosybrown = "#BC8F8F"
        royalblue = "#4169E1"
        saddlebrown = "#8B4513"
        salmon = "#FA8072"
        sandybrown = "#F4A460"
        seagreen = "#2E8B57"
        seashell = "#FFF5EE"
        sienna = "#A0522D"
        silver = "#C0C0C0"
        skyblue = "#87CEEB"
        slateblue = "#6A5ACD"
        slategray = "#708090"
        snow = "#FFFAFA"
        springgreen = "#00FF7F"
        steelblue = "#4682B4"
        tan = "#D2B48C"
        teal = "#008080"
        thistle = "#D8BFD8"
        tomato = "#FF6347"
        turquoise = "#40E0D0"
        violet = "#EE82EE"
        wheat = "#F5DEB3"
        white = "#FFFFFF"
        whitesmoke = "#F5F5F5"
        yellow = "#FFFF00"
        yellowgreen = "#9ACD32"
    }

    SlackBackend ([string]$Token) {
        Import-Module PSSlack -Verbose:$false -ErrorAction Stop

        $config = [ConnectionConfig]::new()
        $secToken = $Token | ConvertTo-SecureString -AsPlainText -Force
        $config.Credential = New-Object System.Management.Automation.PSCredential('asdf', $secToken)
        $conn = [SlackConnection]::New()
        $conn.Config = $config
        $this.Connection = $conn
    }

    # Connect to Slack
    [void]Connect() {
        $this.LogInfo('Connecting to backend')
        $this.LogInfo('Listening for the following message types. All others will be ignored', $this.MessageTypes)
        $this.Connection.Connect()
        $this.BotId = $this.GetBotIdentity()
        $this.LoadUsers()
        $this.LoadRooms()
    }

    # Receive a message from the websocket
    [Message[]]ReceiveMessage() {
        $messages = New-Object -TypeName System.Collections.ArrayList
        try {
            # Read the output stream from the receive job and get any messages since our last read
            [string[]]$jsonResults = $this.Connection.ReadReceiveJob()

            foreach ($jsonResult in $jsonResults) {
                if ($null -ne $jsonResult -and $jsonResult -ne [string]::Empty) {
                    #Write-Debug -Message "[SlackBackend:ReceiveMessage] Received `n$jsonResult"
                    $this.LogDebug('Received message', $jsonResult)

                    # Strip out Slack's URI formatting
                    $jsonResult = $this._SanitizeURIs($jsonResult)

                    $slackMessage = @($jsonResult | ConvertFrom-Json)

                    # Slack will sometimes send back ephemeral messages from user [SlackBot]. Ignore these
                    # These are messages like notifing that a message won't be unfurled because it's already
                    # in the channel in the last hour. Helpful message for some, but not for us.
                    if ($slackMessage.subtype -eq 'bot_message') {
                        $this.LogDebug('SubType is [bot_message]. Ignoring')
                        continue
                    }

                    # Ignore "message_replied" subtypes
                    # These are message Slack sends to update the client that the original message has a new reply.
                    # That reply is sent is another message.
                    # We do this because if the original message that this reply is to is a bot command, the command
                    # will be executed again so we....need to not do that :)
                    if ($slackMessage.subtype -eq 'message_replied') {
                        $this.LogDebug('SubType is [message_replied]. Ignoring')
                        continue
                    }

                    # We only care about certain message types from Slack
                    if ($slackMessage.Type -in $this.MessageTypes) {
                        $msg = [Message]::new()

                        # Set the message type and optionally the subtype
                        #$msg.Type = $slackMessage.type
                        switch ($slackMessage.type) {
                            'channel_rename' {
                                $msg.Type = [MessageType]::ChannelRenamed
                            }
                            'member_joined_channel' {
                                $msg.Type = [MessageType]::Message
                                $msg.SubType = [MessageSubtype]::ChannelJoined
                            }
                            'member_left_channel' {
                                $msg.Type = [MessageType]::Message
                                $msg.SubType = [MessageSubtype]::ChannelLeft
                            }
                            'message' {
                                $msg.Type = [MessageType]::Message
                            }
                            'pin_added' {
                                $msg.Type = [MessageType]::PinAdded
                            }
                            'pin_removed' {
                                $msg.Type = [MessageType]::PinRemoved
                            }
                            'presence_change' {
                                $msg.Type = [MessageType]::PresenceChange
                            }
                            'reaction_added' {
                                $msg.Type = [MessageType]::ReactionAdded
                            }
                            'reaction_removed' {
                                $msg.Type = [MessageType]::ReactionRemoved
                            }
                            'star_added' {
                                $msg.Type = [MessageType]::StarAdded
                            }
                            'star_removed' {
                                $msg.Type = [MessageType]::StarRemoved
                            }
                        }

                        # The channel the message occured in is sometimes
                        # nested in an 'item' property
                        if ($slackMessage.item -and ($slackMessage.item.channel)) {
                            $msg.To = $slackMessage.item.channel
                        }

                        if ($slackMessage.subtype) {
                            switch ($slackMessage.subtype) {
                                'channel_join' {
                                    $msg.Subtype = [MessageSubtype]::ChannelJoined
                                }
                                'channel_leave' {
                                    $msg.Subtype = [MessageSubtype]::ChannelLeft
                                }
                                'channel_name' {
                                    $msg.Subtype = [MessageSubtype]::ChannelRenamed
                                }
                                'channel_purpose' {
                                    $msg.Subtype = [MessageSubtype]::ChannelPurposeChanged
                                }
                                'channel_topic' {
                                    $msg.Subtype = [MessageSubtype]::ChannelTopicChanged
                                }
                            }
                        }
                        $this.LogDebug("Message type is [$($msg.Type)`:$($msg.Subtype)]")

                        $msg.RawMessage = $slackMessage
                        $this.LogDebug('Raw message', $slackMessage)
                        if ($slackMessage.text)    { $msg.Text = $slackMessage.text }
                        if ($slackMessage.channel) { $msg.To   = $slackMessage.channel }
                        if ($slackMessage.user)    { $msg.From = $slackMessage.user }

                        # Resolve From name
                        if ($msg.From) {
                            $msg.FromName = $this.UserIdToUsername($msg.From)
                        }

                        # Resolve channel name
                        # Skip DM channels, they won't have names
                        if ($msg.To -and $msg.To -notmatch '^D') {
                            $msg.ToName = $this.ChannelIdToName($msg.To)
                        }

                        # Mark as DM
                        if ($msg.To -match '^D') {
                            $msg.IsDM = $true
                        }

                        # Get time of message
                        $unixEpoch = [datetime]'1970-01-01'
                        if ($slackMessage.ts) {
                            $msg.Time = $unixEpoch.AddSeconds($slackMessage.ts)
                        } elseIf ($slackMessage.event_ts) {
                            $msg.Time = $unixEpoch.AddSeconds($slackMessage.event_ts)
                        } else {
                            $msg.Time = (Get-Date).ToUniversalTime()
                        }

                        # Sometimes the message is nested in a 'message' subproperty. This could be
                        # if the message contained a link that was unfurled.  We would receive a
                        # 'message_changed' message and need to look in the 'message' subproperty
                        # to see who the message was from.  Slack is weird
                        # https://api.slack.com/events/message/message_changed
                        if ($slackMessage.message) {
                            if ($slackMessage.message.user) {
                                $msg.From = $slackMessage.message.user
                            }
                            if ($slackMessage.message.text) {
                                $msg.Text = $slackMessage.message.text
                            }
                        }

                        # Slack displays @mentions like '@devblackops' but internally in the message
                        # it is <@U4AM3SYI8>
                        # Fix that so we actually see the @username
                        $processed = $this._ProcessMentions($msg.Text)
                        $msg.Text = $processed

                        # ** Important safety tip, don't cross the streams **
                        # Only return messages that didn't come from the bot
                        # else we'd cause a feedback loop with the bot processing
                        # it's own responses
                        if (-not $this.MsgFromBot($msg.From)) {
                            $messages.Add($msg) > $null
                        }
                    } else {
                        $this.LogDebug("Message type is [$($slackMessage.Type)]. Ignoring")
                    }

                }
            }
        } catch {
            Write-Error $_
        }

        return $messages
    }

    # Send a Slack ping
    [void]Ping() {
        # $msg = @{
        #     id = 1
        #     type = 'ping'
        #     time = [System.Math]::Truncate((Get-Date -Date (Get-Date) -UFormat %s))
        # }
        # $json = $msg | ConvertTo-Json
        # $bytes = ([System.Text.Encoding]::UTF8).GetBytes($json)
        # Write-Debug -Message '[SlackBackend:Ping]: One ping only Vasili'
        # $cts = New-Object System.Threading.CancellationTokenSource -ArgumentList 5000

        # $task = $this.Connection.WebSocket.SendAsync($bytes, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $cts.Token)
        # do { Start-Sleep -Milliseconds 100 }
        # until ($task.IsCompleted)
        #$result = $this.Connection.WebSocket.SendAsync($bytes, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $cts.Token).GetAwaiter().GetResult()
    }

    # Send a message back to Slack
    [void]SendMessage([Response]$Response) {
        # Process any custom responses
        $this.LogDebug("[$($Response.Data.Count)] custom responses")
        foreach ($customResponse in $Response.Data) {

            [string]$sendTo = $Response.To
            if ($customResponse.DM) {
                $sendTo = "@$($this.UserIdToUsername($Response.MessageFrom))"
            }

            switch -Regex ($customResponse.PSObject.TypeNames[0]) {
                '(.*?)PoshBot\.Card\.Response' {
                    $this.LogDebug('Custom response is [PoshBot.Card.Response]')
                    $chunks = $this._ChunkString($customResponse.Text)
                    $x = 0
                    foreach ($chunk in $chunks) {
                        $attParams = @{
                            MarkdownFields = 'text'
                            Color = $customResponse.Color
                        }
                        $fbText = 'no data'
                        if (-not [string]::IsNullOrEmpty($chunk.Text)) {
                            $this.LogDebug("Response size [$($chunk.Text.Length)]")
                            $fbText = $chunk.Text
                        }
                        $attParams.Fallback = $fbText
                        if ($customResponse.Title) {

                            # If we chunked up the response, only display the title on the first one
                            if ($x -eq 0) {
                                $attParams.Title = $customResponse.Title
                            }
                        }
                        if ($customResponse.ImageUrl) {
                            $attParams.ImageURL = $customResponse.ImageUrl
                        }
                        if ($customResponse.ThumbnailUrl) {
                            $attParams.ThumbURL = $customResponse.ThumbnailUrl
                        }
                        if ($customResponse.LinkUrl) {
                            $attParams.TitleLink = $customResponse.LinkUrl
                        }
                        if ($customResponse.Fields) {
                            $arr = New-Object System.Collections.ArrayList
                            foreach ($key in $customResponse.Fields.Keys) {
                                $arr.Add(
                                    @{
                                        title = $key;
                                        value = $customResponse.Fields[$key];
                                        short = $true
                                    }
                                )
                            }
                            $attParams.Fields = $arr
                        }

                        if (-not [string]::IsNullOrEmpty($chunk)) {
                            $attParams.Text = '```' + $chunk + '```'
                        } else {
                            $attParams.Text = [string]::Empty
                        }
                        $att = New-SlackMessageAttachment @attParams
                        $msg = $att | New-SlackMessage -Channel $sendTo -AsUser
                        $this.LogDebug("Sending card response back to Slack channel [$sendTo]", $att)
                        $slackResponse = $msg | Send-SlackMessage -Token $this.Connection.Config.Credential.GetNetworkCredential().Password -Verbose:$false
                    }
                    break
                }
                '(.*?)PoshBot\.Text\.Response' {
                    $this.LogDebug('Custom response is [PoshBot.Text.Response]')
                    $chunks = $this._ChunkString($customResponse.Text)
                    foreach ($chunk in $chunks) {
                        if ($customResponse.AsCode) {
                            $t = '```' + $chunk + '```'
                        } else {
                            $t = $chunk
                        }
                        $this.LogDebug("Sending text response back to Slack channel [$sendTo]", $t)
                        $slackResponse = Send-SlackMessage -Token $this.Connection.Config.Credential.GetNetworkCredential().Password -Channel $sendTo -Text $t -Verbose:$false -AsUser
                    }
                    break
                }
                '(.*?)PoshBot\.File\.Upload' {
                    $this.LogDebug('Custom response is [PoshBot.File.Upload]')

                    $uploadParams = @{
                        Token = $this.Connection.Config.Credential.GetNetworkCredential().Password
                        Channel = $sendTo
                    }

                    if ([string]::IsNullOrEmpty($customResponse.Path) -and (-not [string]::IsNullOrEmpty($customResponse.Content))) {
                        $uploadParams.Content = $customResponse.Content
                        if (-not [string]::IsNullOrEmpty($customResponse.FileType)) {
                            $uploadParams.FileType = $customResponse.FileType
                        }
                        if (-not [string]::IsNullOrEmpty($customResponse.FileName)) {
                            $uploadParams.FileName = $customResponse.FileName
                        }
                    } else {
                        # Test if file exists and send error response if not found
                        if (-not (Test-Path -Path $customResponse.Path -ErrorAction SilentlyContinue)) {
                            # Mark command as failed since we could't find the file to upload
                            $this.RemoveReaction($Response.OriginalMessage, [ReactionType]::Success)
                            $this.AddReaction($Response.OriginalMessage, [ReactionType]::Failure)
                            $att = New-SlackMessageAttachment -Color '#FF0000' -Title 'Rut row' -Text "File [$($uploadParams.Path)] not found" -Fallback 'Rut row'
                            $msg = $att | New-SlackMessage -Channel $sendTo -AsUser
                            $this.LogDebug("Sending card response back to Slack channel [$sendTo]", $att)
                            $null = $msg | Send-SlackMessage -Token $this.Connection.Config.Credential.GetNetworkCredential().Password -Verbose:$false
                            break
                        }

                        $this.LogDebug("Uploading [$($customResponse.Path)] to Slack channel [$sendTo]")
                        $uploadParams.Path = $customResponse.Path
                        $uploadParams.Title = Split-Path -Path $customResponse.Path -Leaf
                    }

                    if (-not [string]::IsNullOrEmpty($customResponse.Title)) {
                        $uploadParams.Title = $customResponse.Title
                    }

                    Send-SlackFile @uploadParams -Verbose:$false
                    if (-not $customResponse.KeepFile -and -not [string]::IsNullOrEmpty($customResponse.Path)) {
                        Remove-Item -LiteralPath $customResponse.Path -Force
                    }
                    break
                }
            }
        }

        if ($Response.Text.Count -gt 0) {
            foreach ($t in $Response.Text) {
                $this.LogDebug("Sending response back to Slack channel [$($Response.To)]", $t)
                $slackResponse = Send-SlackMessage -Token $this.Connection.Config.Credential.GetNetworkCredential().Password -Channel $Response.To -Text $t -Verbose:$false -AsUser
            }
        }
    }

    # Add a reaction to an existing chat message
    [void]AddReaction([Message]$Message, [ReactionType]$Type, [string]$Reaction) {
        if ($Message.RawMessage.ts) {
            if ($Type -eq [ReactionType]::Custom) {
                $emoji = $Reaction
            } else {
                $emoji = $this._ResolveEmoji($Type)
            }

            $body = @{
                name = $emoji
                channel = $Message.To
                timestamp = $Message.RawMessage.ts
            }
            $this.LogDebug("Adding reaction [$emoji] to message Id [$($Message.RawMessage.ts)]")
            $resp = Send-SlackApi -Token $this.Connection.Config.Credential.GetNetworkCredential().Password -Method 'reactions.add' -Body $body -Verbose:$false
            if (-not $resp.ok) {
                $this.LogInfo([LogSeverity]::Error, 'Error adding reaction to message', $resp)
            }
        }
    }

    # Remove a reaction from an existing chat message
    [void]RemoveReaction([Message]$Message, [ReactionType]$Type, [string]$Reaction) {
        if ($Message.RawMessage.ts) {
            if ($Type -eq [ReactionType]::Custom) {
                $emoji = $Reaction
            } else {
                $emoji = $this._ResolveEmoji($Type)
            }

            $body = @{
                name = $emoji
                channel = $Message.To
                timestamp = $Message.RawMessage.ts
            }
            $this.LogDebug("Removing reaction [$emoji] from message Id [$($Message.RawMessage.ts)]")
            $resp = Send-SlackApi -Token $this.Connection.Config.Credential.GetNetworkCredential().Password -Method 'reactions.remove' -Body $body -Verbose:$false
            if (-not $resp.ok) {
                $this.LogInfo([LogSeverity]::Error, 'Error removing reaction from message', $resp)
            }
        }
    }

    # Resolve a channel name to an Id
    [string]ResolveChannelId([string]$ChannelName) {
        if ($ChannelName -match '^#') {
            $ChannelName = $ChannelName.TrimStart('#')
        }
        $channelId = ($this.Connection.LoginData.channels | Where-Object name -eq $ChannelName).id
        if (-not $ChannelId) {
            $channelId = ($this.Connection.LoginData.channels | Where-Object id -eq $ChannelName).id
        }
        $this.LogDebug("Resolved channel [$ChannelName] to [$channelId]")
        return $channelId
    }

    # Populate the list of users the Slack team
    [void]LoadUsers() {
        $this.LogDebug('Getting Slack users')
        $allUsers = Get-Slackuser -Token $this.Connection.Config.Credential.GetNetworkCredential().Password -Verbose:$false
        $this.LogDebug("[$($allUsers.Count)] users returned")
        $allUsers | ForEach-Object {
            $user = [SlackPerson]::new()
            $user.Id = $_.ID
            $user.Nickname = $_.Name
            $user.FullName = $_.RealName
            $user.FirstName = $_.FirstName
            $user.LastName = $_.LastName
            $user.Email = $_.Email
            $user.Phone = $_.Phone
            $user.Skype = $_.Skype
            $user.IsBot = $_.IsBot
            $user.IsAdmin = $_.IsAdmin
            $user.IsOwner = $_.IsOwner
            $user.IsPrimaryOwner = $_.IsPrimaryOwner
            $user.IsUltraRestricted = $_.IsUltraRestricted
            $user.Status = $_.Status
            $user.TimeZoneLabel = $_.TimeZoneLabel
            $user.TimeZone = $_.TimeZone
            $user.Presence = $_.Presence
            $user.Deleted = $_.Deleted
            if (-not $this.Users.ContainsKey($_.ID)) {
                $this.LogDebug("Adding user [$($_.ID):$($_.Name)]")
                $this.Users[$_.ID] =  $user
            }
        }

        foreach ($key in $this.Users.Keys) {
            if ($key -notin $allUsers.ID) {
                $this.LogDebug("Removing outdated user [$key]")
                $this.Users.Remove($key)
            }
        }
    }

    # Populate the list of channels in the Slack team
    [void]LoadRooms() {
        $this.LogDebug('Getting Slack channels')
        $getChannelParams = @{
            Token           = $this.Connection.Config.Credential.GetNetworkCredential().Password
            ExcludeArchived = $true
            Verbose         = $false
            Paging          = $true
        }
        $allChannels = Get-SlackChannel @getChannelParams
        $this.LogDebug("[$($allChannels.Count)] channels returned")

        $allChannels.ForEach({
            $channel = [SlackChannel]::new()
            $channel.Id          = $_.ID
            $channel.Name        = $_.Name
            $channel.Topic       = $_.Topic
            $channel.Purpose     = $_.Purpose
            $channel.Created     = $_.Created
            $channel.Creator     = $_.Creator
            $channel.IsArchived  = $_.IsArchived
            $channel.IsGeneral   = $_.IsGeneral
            $channel.MemberCount = $_.MemberCount
            foreach ($member in $_.Members) {
                $channel.Members.Add($member, $null)
            }
            $this.LogDebug("Adding channel: $($_.ID):$($_.Name)")
            $this.Rooms[$_.ID] = $channel
        })

        foreach ($key in $this.Rooms.Keys) {
            if ($key -notin $allChannels.ID) {
                $this.LogDebug("Removing outdated channel [$key]")
                $this.Rooms.Remove($key)
            }
        }
    }

    # Get the bot identity Id
    [string]GetBotIdentity() {
        $id = $this.Connection.LoginData.self.id
        $this.LogVerbose("Bot identity is [$id]")
        return $id
    }

    # Determine if incoming message was from the bot
    [bool]MsgFromBot([string]$From) {
        $frombot = ($this.BotId -eq $From)
        if ($fromBot) {
            $this.LogDebug("Message is from bot [From: $From == Bot: $($this.BotId)]. Ignoring")
        } else {
            $this.LogDebug("Message is not from bot [From: $From <> Bot: $($this.BotId)]")
        }
        return $fromBot
    }

    # Get a user by their Id
    [SlackPerson]GetUser([string]$UserId) {
        $user = $this.Users[$UserId]
        if (-not $user) {
            $this.LogDebug([LogSeverity]::Warning, "User [$UserId] not found. Refreshing users")
            $this.LoadUsers()
            $user = $this.Users[$UserId]
        }

        if ($user) {
            $this.LogDebug("Resolved user [$UserId]", $user)
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve user [$UserId]")
        }
        return $user
    }

    # Get a user Id by their name
    [string]UsernameToUserId([string]$Username) {
        $Username = $Username.TrimStart('@')
        $user = $this.Users.Values | Where-Object {$_.Nickname -eq $Username}
        $id = $null
        if ($user) {
            $id = $user.Id
        } else {
            # User each doesn't exist or is not in the local cache
            # Refresh it and try again
            $this.LogDebug([LogSeverity]::Warning, "User [$Username] not found. Refreshing users")
            $this.LoadUsers()
            $user = $this.Users.Values | Where-Object {$_.Nickname -eq $Username}
            if (-not $user) {
                $id = $null
            } else {
                $id = $user.Id
            }
        }
        if ($id) {
            $this.LogDebug("Resolved [$Username] to [$id]")
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve user [$Username]")
        }
        return $id
    }

    # Get a user name by their Id
    [string]UserIdToUsername([string]$UserId) {
        $name = $null
        if ($this.Users.ContainsKey($UserId)) {
            $name = $this.Users[$UserId].Nickname
        } else {
            $this.LogDebug([LogSeverity]::Warning, "User [$UserId] not found. Refreshing users")
            $this.LoadUsers()
            $name = $this.Users[$UserId].Nickname
        }
        if ($name) {
            $this.LogDebug("Resolved [$UserId] to [$name]")
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve user [$UserId]")
        }
        return $name
    }

    # Get the channel name by Id
    [string]ChannelIdToName([string]$ChannelId) {
        $name = $null
        if ($this.Rooms.ContainsKey($ChannelId)) {
            $name = $this.Rooms[$ChannelId].Name
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Channel [$ChannelId] not found. Refreshing channels")
            $this.LoadRooms()
            $name = $this.Rooms[$ChannelId].Name
        }
        if ($name) {
            $this.LogDebug("Resolved [$ChannelId] to [$name]")
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve channel [$ChannelId]")
        }
        return $name
    }

    # Get all user info by their ID
    [hashtable]GetUserInfo([string]$UserId) {
        $user = $null
        if ($this.Users.ContainsKey($UserId)) {
            $user = $this.Users[$UserId]
        } else {
            $this.LogDebug([LogSeverity]::Warning, "User [$UserId] not found. Refreshing users")
            $this.LoadUsers()
            $user = $this.Users[$UserId]
        }

        if ($user) {
            $this.LogDebug("Resolved [$UserId] to [$($user.Nickname)]")
            return $user.ToHash()
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve channel [$UserId]")
            return $null
        }
    }

    # Remove extra characters that Slack decorates urls with
    hidden [string] _SanitizeURIs([string]$Text) {
        $sanitizedText = $Text -replace '<([^\|>]+)\|([^\|>]+)>', '$2'
        $sanitizedText = $sanitizedText -replace '<(http([^>]+))>', '$1'
        return $sanitizedText
    }

    # Break apart a string by number of characters
    # This isn't a very efficient method but it splits the message cleanly on
    # whole lines and produces better output
    hidden [Collections.Generic.List[string]] _ChunkString([string]$Text) {

        # Don't bother chunking an empty string
        if ([string]::IsNullOrEmpty($Text)) {
            return $text
        }

        $chunks             = [Collections.Generic.List[string]]::new()
        $currentChunkLength = 0
        $currentChunk       = ''
        $array              = $Text -split [Environment]::NewLine

        foreach ($line in $array) {
            if (($currentChunkLength + $line.Length) -lt $this.MaxMessageLength) {
                $currentChunkLength += $line.Length
                $currentChunk += ($line + [Environment]::NewLine)
            } else {
                $chunks.Add($currentChunk + [Environment]::NewLine)
                $currentChunk = ($line + [Environment]::NewLine)
                $currentChunkLength = $line.Length
            }
        }
        $chunks.Add($currentChunk)

        return $chunks
    }

    # Resolve a reaction type to an emoji
    hidden [string]_ResolveEmoji([ReactionType]$Type) {
        $emoji = [string]::Empty
        Switch ($Type) {
            'Success'        { return 'white_check_mark' }
            'Failure'        { return 'exclamation' }
            'Processing'     { return 'gear' }
            'Warning'        { return 'warning' }
            'ApprovalNeeded' { return 'closed_lock_with_key'}
            'Cancelled'      { return 'no_entry_sign'}
            'Denied'         { return 'x'}
        }
        return $emoji
    }

    # Translate formatted @mentions like <@U4AM3SYI8> into @devblackops
    hidden [string]_ProcessMentions([string]$Text) {
        $processed = $Text

        $mentions = $processed | Select-String -Pattern '(?<name><@[^>]*>*)' -AllMatches | ForEach-Object {
            $_.Matches | ForEach-Object {
                [pscustomobject]@{
                    FormattedId = $_.Value
                    UnformattedId = $_.Value.TrimStart('<@').TrimEnd('>')
                }
            }
        }
        $mentions | ForEach-Object {
            if ($name = $this.UserIdToUsername($_.UnformattedId)) {
                $processed = $processed -replace $_.FormattedId, "@$name"
                $this.LogDebug($processed)
            } else {
                $this.LogDebug([LogSeverity]::Warning, "Unable to translate @mention [$($_.FormattedId)] into a username")
            }
        }

        return $processed
    }
}

function New-PoshBotSlackBackend {
    <#
    .SYNOPSIS
        Create a new instance of a Slack backend
    .DESCRIPTION
        Create a new instance of a Slack backend
    .PARAMETER Configuration
        The hashtable containing backend-specific properties on how to create the Slack backend instance.
    .EXAMPLE
        PS C:\> $backendConfig = @{Name = 'SlackBackend'; Token = '<SLACK-API-TOKEN>'}
        PS C:\> $backend = New-PoshBotSlackBackend -Configuration $backendConfig

        Create a Slack backend using the specified API token
    .INPUTS
        Hashtable
    .OUTPUTS
        SlackBackend
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='*')]
    [cmdletbinding()]
    param(
        [parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('BackendConfiguration')]
        [hashtable[]]$Configuration
    )

    process {
        foreach ($item in $Configuration) {
            if (-not $item.Token) {
                throw 'Configuration is missing [Token] parameter'
            } else {
                Write-Verbose 'Creating new Slack backend instance'
                $backend = [SlackBackend]::new($item.Token)
                if ($item.Name) {
                    $backend.Name = $item.Name
                }
                $backend
            }
        }
    }
}

Export-ModuleMember -Function 'New-PoshBotSlackBackend'


class SlackChannel : Room {
    [datetime]$Created
    [string]$Creator
    [bool]$IsArchived
    [bool]$IsGeneral
    [int]$MemberCount
    [string]$Purpose
}

class SlackConnection : Connection {

    [System.Net.WebSockets.ClientWebSocket]$WebSocket
    [pscustomobject]$LoginData
    [string]$UserName
    [string]$Domain
    [string]$WebSocketUrl
    [bool]$Connected
    [object]$ReceiveJob = $null

    SlackConnection() {
        $this.WebSocket = New-Object System.Net.WebSockets.ClientWebSocket
        $this.WebSocket.Options.KeepAliveInterval = 5
    }

    # Connect to Slack and start receiving messages
    [void]Connect() {
        if ($null -eq $this.ReceiveJob -or $this.ReceiveJob.State -ne 'Running') {
            $this.LogDebug('Connecting to Slack Real Time API')
            $this.RtmConnect()
            $this.StartReceiveJob()
        } else {
            $this.LogDebug([LogSeverity]::Warning, 'Receive job is already running')
        }
    }

    # Log in to Slack with the bot token and get a URL to connect to via websockets
    [void]RtmConnect() {
        $token = $this.Config.Credential.GetNetworkCredential().Password
        $url = "https://slack.com/api/rtm.start?token=$($token)&pretty=1"
        try {
            $r = Invoke-RestMethod -Uri $url -Method Get -Verbose:$false
            $this.LoginData = $r
            if ($r.ok) {
                $this.LogInfo('Successfully authenticated to Slack Real Time API')
                $this.WebSocketUrl = $r.url
                $this.Domain = $r.team.domain
                $this.UserName = $r.self.name
            } else {
                throw $r
            }
        } catch {
            $this.LogInfo([LogSeverity]::Error, 'Error connecting to Slack Real Time API', [ExceptionFormatter]::Summarize($_))
        }
    }

    # Setup the websocket receive job
    [void]StartReceiveJob() {
        $recv = {
            [cmdletbinding()]
            param(
                [parameter(mandatory)]
                $url
            )

            # Connect to websocket
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Write-Verbose "[SlackBackend:ReceiveJob] Connecting to websocket at [$($url)]"
            [System.Net.WebSockets.ClientWebSocket]$webSocket = New-Object System.Net.WebSockets.ClientWebSocket
            $cts = New-Object System.Threading.CancellationTokenSource
            $task = $webSocket.ConnectAsync($url, $cts.Token)
            do { Start-Sleep -Milliseconds 100 }
            until ($task.IsCompleted)

            # Receive messages and put on output stream so the backend can read them
            [ArraySegment[byte]]$buffer = [byte[]]::new(4096)
            $ct = New-Object System.Threading.CancellationToken
            $taskResult = $null
            while ($webSocket.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
                do {
                    $taskResult = $webSocket.ReceiveAsync($buffer, $ct)
                    while (-not $taskResult.IsCompleted) {
                        Start-Sleep -Milliseconds 100
                    }
                } until (
                    $taskResult.Result.Count -lt 4096
                )
                $jsonResult = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $taskResult.Result.Count)

                if (-not [string]::IsNullOrEmpty($jsonResult)) {
                    $jsonResult
                }
            }
            $socketStatus = [pscustomobject]@{
                State = $webSocket.State
                CloseStatus = $webSocket.CloseStatus
                CloseStatusDescription = $webSocket.CloseStatusDescription
            }
            $socketStatusStr = ($socketStatus | Format-List | Out-String).Trim()
            Write-Warning -Message "Websocket state is [$($webSocket.State.ToString())].`n$socketStatusStr"
        }

        try {
            $this.ReceiveJob = Start-Job -Name ReceiveRtmMessages -ScriptBlock $recv -ArgumentList $this.WebSocketUrl -ErrorAction Stop -Verbose
            $this.Connected = $true
            $this.Status = [ConnectionStatus]::Connected
            $this.LogInfo("Started websocket receive job [$($this.ReceiveJob.Id)]")
        } catch {
            $this.LogInfo([LogSeverity]::Error, "$($_.Exception.Message)", [ExceptionFormatter]::Summarize($_))
        }
    }

    # Read all available data from the job
    [string[]]ReadReceiveJob() {
        # Read stream info from the job so we can log them
        $infoStream = $this.ReceiveJob.ChildJobs[0].Information.ReadAll()
        $warningStream = $this.ReceiveJob.ChildJobs[0].Warning.ReadAll()
        $errStream = $this.ReceiveJob.ChildJobs[0].Error.ReadAll()
        $verboseStream = $this.ReceiveJob.ChildJobs[0].Verbose.ReadAll()
        $debugStream = $this.ReceiveJob.ChildJobs[0].Debug.ReadAll()
        foreach ($item in $infoStream) {
            $this.LogInfo($item.ToString())
        }
        foreach ($item in $warningStream) {
            $this.LogInfo([LogSeverity]::Warning, $item.ToString())
        }
        foreach ($item in $errStream) {
            $this.LogInfo([LogSeverity]::Error, $item.ToString())
        }
        foreach ($item in $verboseStream) {
            $this.LogVerbose($item.ToString())
        }
        foreach ($item in $debugStream) {
            $this.LogVerbose($item.ToString())
        }

        # The receive job stopped for some reason. Reestablish the connection if the job isn't running
        if ($this.ReceiveJob.State -ne 'Running') {
            $this.LogInfo([LogSeverity]::Warning, "Receive job state is [$($this.ReceiveJob.State)]. Attempting to reconnect...")
            Start-Sleep -Seconds 5
            $this.Connect()
        }

        if ($this.ReceiveJob.HasMoreData) {
            [string[]]$jobResult = $this.ReceiveJob.ChildJobs[0].Output.ReadAll()
            return $jobResult
        } else {
            return $null
        }
    }

    # Stop the receive job
    [void]Disconnect() {
        $this.LogInfo('Closing websocket')
        if ($this.ReceiveJob) {
            $this.LogInfo("Stopping receive job [$($this.ReceiveJob.Id)]")
            $this.ReceiveJob | Stop-Job -Confirm:$false -PassThru | Remove-Job -Force -ErrorAction SilentlyContinue
        }
        $this.Connected = $false
        $this.Status = [ConnectionStatus]::Disconnected
    }
}

enum SlackMessageType {
    Normal
    Error
    Warning
}

class SlackMessage : Message {

    [SlackMessageType]$MessageType = [SlackMessageType]::Normal

    SlackMessage(
        [string]$To,
        [string]$From,
        [string]$Body = [string]::Empty
    ) {
        $this.To = $To
        $this.From = $From
        $this.Body = $Body
    }

}


class SlackPerson : Person {
    [string]$Email
    [string]$Phone
    [string]$Skype
    [bool]$IsBot
    [bool]$IsAdmin
    [bool]$IsOwner
    [bool]$IsPrimaryOwner
    [bool]$IsRestricted
    [bool]$IsUltraRestricted
    [string]$Status
    [string]$TimeZoneLabel
    [string]$TimeZone
    [string]$Presence
    [bool]$Deleted
}


function New-PoshBotTeamsBackend {
    <#
    .SYNOPSIS
        Create a new instance of a Microsoft Teams backend
    .DESCRIPTION
        Create a new instance of a Microsoft Teams backend
    .PARAMETER Configuration
        The hashtable containing backend-specific properties on how to create the instance.
    .EXAMPLE
        PS C:\> $backendConfig = @{
            Name = 'TeamsBackend'
            Credential = [pscredential]::new(
                '<BOT-ID>',
                ('<BOT-PASSWORD>' | ConvertTo-SecureString -AsPlainText -Force)
            )
            ServiceBusNamespace = '<SERVICEBUS-NAMESPACE>'
            QueueName           = '<QUEUE-NAME>'
            AccessKeyName       = '<KEY-NAME>'
            AccessKey           = '<SECRET>' | ConvertTo-SecureString -AsPlainText -Force
        }
        PS C:\> $$backend = New-PoshBotTeamsBackend -Configuration $backendConfig

        Create a Microsoft Teams backend using the specified Bot Framework credentials and Service Bus information
    .INPUTS
        Hashtable
    .OUTPUTS
        TeamsBackend
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='*')]
    [cmdletbinding()]
    param(
        [parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('BackendConfiguration')]
        [hashtable[]]$Configuration
    )

    begin {
        $requiredProperties = @(
            'BotName', 'TeamId', 'Credential', 'ServiceBusNamespace', 'QueueName', 'AccessKeyName', 'AccessKey'
        )
    }

    process {
        foreach ($item in $Configuration) {

            # Validate required hashtable properties
            if ($missingProperties = $requiredProperties.Where({$item.Keys -notcontains $_})) {
                throw "The following required backend properties are not defined: $($missingProperties -join ', ')"
            }
            Write-Verbose 'Creating new Teams backend instance'

            $connectionConfig = [TeamsConnectionConfig]::new()
            $connectionConfig.BotName             = $item.BotName
            $connectionConfig.TeamId              = $item.TeamId
            $connectionConfig.Credential          = $item.Credential
            $connectionConfig.ServiceBusNamespace = $item.ServiceBusNamespace
            $connectionConfig.QueueName           = $item.QueueName
            $connectionConfig.AccessKeyName       = $item.AccessKeyName
            $connectionConfig.AccessKey           = $item.AccessKey

            $backend = [TeamsBackend]::new($connectionConfig)
            if ($item.Name) {
                $backend.Name = $item.Name
            }
            $backend
        }
    }
}

Export-ModuleMember -Function 'New-PoshBotTeamsBackend'


class TeamsBackend : Backend {

    [bool]$LazyLoadUsers = $true

    # The types of message that we care about from Teams
    # All othere will be ignored
    [string[]]$MessageTypes = @(
        'message'
    )

    [string]$TeamId     = $null
    [string]$ServiceUrl = $null
    [string]$BotId      = $null
    [string]$BotName    = $null
    [string]$TenantId   = $null

    [hashtable]$DMConverations = @{}

    [hashtable]$FileUploadTracker = @{}

    TeamsBackend([TeamsConnectionConfig]$Config) {
        $conn = [TeamsConnection]::new($Config)
        $this.TeamId = $Config.TeamId
        $this.Connection = $conn
    }

    # Connect to Teams
    [void]Connect() {
        $this.LogInfo('Connecting to backend')
        $this.Connection.Connect()
    }

    [Message[]]ReceiveMessage() {
        $messages = New-Object -TypeName System.Collections.ArrayList
        try {
            # Read the output stream from the receive thread and get any messages since our last read
            $jsonResults = $this.Connection.ReadReceiveThread()

            if (-not [string]::IsNullOrEmpty($jsonResults)) {

                foreach ($jsonResult in $jsonResults) {

                    $this.LogDebug('Received message', $jsonResult)

                    $teamsMessages = @($jsonResult | ConvertFrom-Json)

                    foreach ($teamsMessage in $teamsMessages) {

                        $this.DelayedInit($teamsMessage)

                        # We only care about certain message types from Teams
                        if ($teamsMessage.type -in $this.MessageTypes) {
                            $msg = [Message]::new()

                            switch ($teamsMessage.type) {
                                'message' {
                                    $msg.Type = [MessageType]::Message
                                    break
                                }
                            }
                            $this.LogDebug("Message type is [$($msg.Type)]")

                            $msg.Id = $teamsMessage.id
                            if ($teamsMessage.recipient) {
                                $msg.To = $teamsMessage.recipient.id
                            }

                            $msg.RawMessage = $teamsMessage
                            $this.LogDebug('Raw message', $teamsMessage)

                            # When commands are directed to PoshBot, the bot must be "at" mentioned.
                            # This will show up in the text of the message received. We don't need it
                            # so strip it out.
                            if ($teamsMessage.text)    {
                                $msg.Text = $teamsMessage.text.Replace("<at>$($this.Connection.Config.BotName)</at> ", '') -Replace '\n', ''
                            }

                            if ($teamsMessage.from) {
                                $msg.From     = $teamsMessage.from.id
                                $msg.FromName = $teamsMessage.from.name
                            }

                            # Mark as DM
                            # 'team' data is not passed in channel conversations
                            # so we can use it to determine if message is in personal chat
                            # https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/bots/bot-conversations/bots-conversations#teams-channel-data
                            if (-not $teamsMessage.channelData.team) {
                                $msg.IsDM = $true
                                $msg.ToName = $this.Connection.Config.BotName
                            } else {
                                if ($msg.To) {
                                    $msg.ToName = $this.UserIdToUsername($msg.To)
                                }
                            }

                            # Resolve channel name
                            # Skip DM channels, they won't have names
                            if (($teamsMessage.channelData.teamsChannelId) -and (-not $msg.IsDM)) {
                                $msg.ToName = $this.ChannelIdToName($teamsMessage.channelData.teamsChannelId)
                            }

                            # Get time of message
                            $msg.Time = [datetime]$teamsMessage.timestamp

                            $messages.Add($msg) > $null
                        } else {
                            $this.LogDebug("Message type is [$($teamsMessage.type)]. Ignoring")
                        }
                    }
                }
            }
        } catch {
            $this.LogInfo([LogSeverity]::Error, 'Error authenticating to Teams', [ExceptionFormatter]::Summarize($_))
        }

        return $messages
    }

    [void]Ping() {}

    # Send a message
    [void]SendMessage([Response]$Response) {

        $baseUrl        = $Response.OriginalMessage.RawMessage.serviceUrl
        $fromId         = $Response.OriginalMessage.RawMessage.from.id
        $fromName       = $Response.OriginalMessage.RawMessage.from.name
        $recipientId    = $Response.OriginalMessage.RawMessage.recipient.id
        $recipientName  = $Response.OriginalMessage.RawMessage.recipient.name
        $conversationId = $Response.OriginalMessage.RawMessage.conversation.id
        $activityId     = $Response.OriginalMessage.RawMessage.id
        $responseUrl    = "$($baseUrl)v3/conversations/$conversationId/activities/$activityId"
        $channelId      = $Response.OriginalMessage.RawMessage.channelData.teamsChannelId
        $headers = @{
            Authorization = "Bearer $($this.Connection._AccessTokenInfo.access_token)"
        }

        # Process any custom responses
        $this.LogDebug("[$($Response.Data.Count)] custom responses")
        foreach ($customResponse in $Response.Data) {

            if ($customResponse.Text) {
                #$customResponse.Text = $this._RepairText($customResponse.Text)
            }

            # Redirect response to DM channel if told to
            if ($customResponse.DM) {
                $conversationId = $this._CreateDMConversation($Response.OriginalMessage.RawMessage.from.id)
                $activityId = $conversationId
                $responseUrl = "$($baseUrl)v3/conversations/$conversationId/activities/"
            }

            switch -Regex ($customResponse.PSObject.TypeNames[0]) {
                '(.*?)PoshBot\.Card\.Response' {
                    $this.LogDebug('Custom response is [PoshBot.Card.Response]')

                    $cardBody = @{
                        type = 'message'
                        from = @{
                            id   = $fromId
                            name = $fromName
                        }
                        conversation = @{
                            id = $conversationId
                        }
                        recipient = @{
                            id = $recipientId
                            name = $recipientName
                        }
                        attachments = @(
                            @{
                                contentType = 'application/vnd.microsoft.teams.card.o365connector'
                                content = @{
                                    "@type" = 'MessageCard'
                                    "@context" = 'http://schema.org/extensions'
                                    themeColor = $customResponse.Color -replace '#', ''
                                    sections = @(
                                        @{

                                        }
                                    )
                                }
                            }
                        )
                        replyToId = $activityId
                    }

                    # Thumbnail
                    if ($customResponse.ThumbnailUrl) {
                        $cardBody.attachments[0].content.sections[0].activityImageType = 'article'
                        $cardBody.attachments[0].content.sections[0].activityImage = $customResponse.ThumbnailUrl
                    }

                    # Title
                    if ($customResponse.Title) {
                        $cardBody.attachments[0].content.summary = $customResponse.Title
                        if ($customResponse.LinkUrl) {
                            $cardBody.attachments[0].content.title = "[$($customResponse.Title)]($($customResponse.LinkUrl))"
                        } else {
                            $cardBody.attachments[0].content.title = $customResponse.Title
                        }
                    }

                    # TextBlock
                    if ($customResponse.Text) {
                        $cardBody.attachments[0].content.sections[0].text = '<pre>' + $customResponse.Text + '</pre>'
                        $cardBody.attachments[0].content.sections[0].textFormat = 'markdown'
                    }

                    # Facts
                    if ($customResponse.Fields.Count -gt 0) {
                        $cardBody.attachments[0].content.sections[0].facts = @()
                        foreach ($field in $customResponse.Fields.GetEnumerator()) {
                            $cardBody.attachments[0].content.sections[0].facts += @{
                                name = $field.Name
                                value = $field.Value.ToString()
                            }
                        }
                    }

                    # Prepend image if needed
                    if ($customResponse.ImageUrl) {
                        $cardBody.attachments[0].content.sections = @(
                            @{
                                images = @(
                                    @{
                                        image = $customResponse.ImageUrl
                                    }
                                )
                            }
                        ) + $cardBody.attachments[0].content.sections
                    }

                    $body = $cardBody | ConvertTo-Json -Depth 20
                    Write-Verbose $body
                    # $body | Out-File -FilePath "$script:moduleBase/responses.json" -Append
                    $this.LogDebug("Sending response back to Teams conversation [$conversationId]", $body)
                    try {
                        $responseParams = @{
                            Uri         = $responseUrl
                            Method      = 'Post'
                            Body        = $body
                            ContentType = 'application/json'
                            Headers     = $headers
                        }
                        $teamsResponse = Invoke-RestMethod @responseParams
                    } catch {
                        $this.LogInfo([LogSeverity]::Error, "$($_.Exception.Message)", [ExceptionFormatter]::Summarize($_))
                    }

                    break
                }
                '(.*?)PoshBot\.Text\.Response' {
                    $this.LogDebug('Custom response is [PoshBot.Text.Response]')

                    $textFormat = 'plain'
                    $cardText = $customResponse.Text
                    if ($customResponse.AsCode) {
                        $textFormat = 'markdown'
                        $cardText = '<pre>' + $cardText + '</pre>'
                    }

                    $cardBody = @{
                        type = 'message'
                        from = @{
                            id   = $fromId
                            name = $fromName
                        }
                        conversation = @{
                            id = $conversationId
                        }
                        recipient = @{
                            id = $recipientId
                            name = $recipientName
                        }
                        text = $cardText
                        textFormat = $textFormat
                        # attachments = @(
                        #     @{
                        #         contentType = 'application/vnd.microsoft.teams.card.o365connector'
                        #         content = @{
                        #             "@type" = 'MessageCard'
                        #             "@context" = 'http://schema.org/extensions'
                        #             text = $cardText
                        #             textFormat = $textFormat
                        #         }
                        #     }
                        # )
                        replyToId = $activityId
                    }

                    $body = $cardBody | ConvertTo-Json -Depth 15
                    Write-Verbose $body
                    # $body | Out-File -FilePath "$script:moduleBase/responses.json" -Append
                    $this.LogDebug("Sending response back to Teams channel [$conversationId]", $body)
                    try {
                        $responseParams = @{
                            Uri         = $responseUrl
                            Method      = 'Post'
                            Body        = $body
                            ContentType = 'application/json'
                            Headers     = $headers
                        }
                        $teamsResponse = Invoke-RestMethod @responseParams
                    } catch {
                        $this.LogInfo([LogSeverity]::Error, "$($_.Exception.Message)", [ExceptionFormatter]::Summarize($_))
                    }

                    break
                }
                '(.*?)PoshBot\.File\.Upload' {
                    # Teams documentation: https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/bots/bots-files
                    $this.LogDebug('Custom response is [PoshBot.File.Upload]')

                    # Teams doesn't support generic file uploads yet :(
                    # Send a message informing the user of this sad fact
                    $jsonResponse = @{
                        type = 'message'
                        from = @{
                            id = $recipientId
                            name = $recipientName
                        }
                        conversation = @{
                            id = $conversationId
                            name = ''
                        }
                        recipient = @{
                            id = $fromId
                            name = $fromName
                        }
                        text = "I don't know how to upload files to Teams yet but I'm learning."
                        replyToId = $activityId
                    } | ConvertTo-Json

                    # $jsonResponse | Out-File -FilePath "$script:moduleBase/responses.json" -Append
                    $this.LogDebug("Sending response back to Teams conversation [$conversationId]")
                    try {
                        $responseParams = @{
                            Uri         = $responseUrl
                            Method      = 'Post'
                            Body        = $jsonResponse
                            ContentType = 'application/json'
                            Headers     = $headers
                        }
                        $teamsResponse = Invoke-RestMethod @responseParams
                    } catch {
                        $this.LogInfo([LogSeverity]::Error, "$($_.Exception.Message)", [ExceptionFormatter]::Summarize($_))
                    }

                    # # Get details about file to upload
                    # $fileToUpload = @{
                    #     Path      = $customResponse.Path
                    #     Name      = Split-Path -Path $customResponse.Path -Leaf
                    #     Size      = (Get-Item -Path $customResponse.Path).Length
                    #     ConsentId = [guid]::NewGuid().ToString()
                    # }
                    # if (-not [string]::IsNullOrEmpty($customResponse.Title)) {
                    #     $fileToUpload.Description = $customResponse.Title
                    # } else {
                    #     $fileToUpload.Description = $fileToUpload.Name
                    # }

                    # Teams doesn't support file uploads to group channels (lame)
                    # Setup a private DM session with the user so we can send the
                    # file consent card
                    # $conversationId = $this._CreateDMConversation($Response.OriginalMessage.RawMessage.from.id)
                    # $responseUrl = "$($baseUrl)v3/conversations/$conversationId/activities/"

                    # $fileConsentRequest = @{
                    #     type = 'message'
                    #     from = @{
                    #         id = $recipientId
                    #         name = $recipientName
                    #     }
                    #     conversation = @{
                    #         id = $conversationId
                    #         name = ''
                    #     }
                    #     recipient = @{
                    #         id = $fromId
                    #         name = $fromName
                    #     }
                    #     replyToId = $activityId
                    #     attachments = @(
                    #         @{
                    #             contentType = 'application/vnd.microsoft.teams.card.file.consent'
                    #             name = $fileToUpload.Name
                    #             content = @{
                    #                 description = $fileToUpload.Description
                    #                 sizeInBytes = $fileToUpload.Size
                    #                 acceptContext = @{
                    #                     consentId = $fileToUpload.ConsentId
                    #                 }
                    #                 declineContext = @{
                    #                     consentId = $fileToUpload.ConsentId
                    #                 }
                    #             }
                    #         }
                    #     )
                    # } | ConvertTo-Json -Depth 15

                    # $fileConsentRequest | Out-File -FilePath "$script:moduleBase/file-requests.json" -Append
                    # $this.LogDebug("Sending file upload request [$($fileToUpload.ConsentId)] to Teams conversation [$conversationId]")
                    # try {
                    #     $responseParams = @{
                    #         Uri         = $responseUrl
                    #         Method      = 'Post'
                    #         Body        = $fileConsentRequest
                    #         ContentType = 'application/json'
                    #         Headers     = $headers
                    #     }
                    #     $teamsResponse = Invoke-RestMethod @responseParams

                    #     $this.FileUploadTracker.Add($fileToUpload.ConsentId, $fileToUpload)

                    # } catch {
                    #     $this.LogInfo([LogSeverity]::Error, "$($_.Exception.Message)", [ExceptionFormatter]::Summarize($_))
                    # }

                    # $contentType = 'application/octet-stream'
                    # if (($null -eq $global:IsWindows) -or $global:IsWindows) {

                    # } else {
                    #     if (Get-Command -Name file -CommandType Application) {
                    #         $contentType =  & file --mime-type -b $customResponse.Path
                    #     }
                    # }

                    # $uploadParams = @{
                    #     type           = $contentType
                    #     name           = $customResponse.Title
                    # }

                    # if ((Test-Path $customResponse.Path -ErrorAction SilentlyContinue)) {
                    #     $bytes = [System.Text.Encoding]::UTF8.GetBytes($customResponse.Path)
                    #     $uploadParams.originalBase64  = [System.Convert]::ToBase64String($bytes)
                    #     $uploadParams.thumbnailBase64 = [System.Convert]::ToBase64String($bytes)
                    #     $this.LogDebug("Uploading [$($customResponse.Path)] to Teams conversation [$conversationId]")
                    #     $payLoad = $uploadParams | ConvertTo-Json
                    #     $this.LogDebug('JSON payload', $payLoad)
                    #     $attachmentUrl = "$($baseUrl)v3/conversations/$conversationId/attachments"

                    #     $responseParams = @{
                    #         Uri         = $attachmentUrl
                    #         Method      = 'Post'
                    #         Body        = $payLoad
                    #         ContentType = 'application/json'
                    #         Headers     = $headers
                    #     }
                    #     $teamsResponse = Invoke-RestMethod @responseParams
                    # }

                    break
                }
            }
        }

        # Normal responses
        if ($Response.Text.Count -gt 0) {
            $this.LogDebug("Sending response back to Teams channel [$($Response.To)]")
            $this.SendTeamsMessaage($Response)
        }
    }

    # Add a reaction to an existing chat message
    # Currently only supports sending a 'typing' indicator to DMConverations
    [void]AddReaction([Message]$Message, [ReactionType]$Type, [string]$Reaction) {

        $baseUrl = $Message.RawMessage.serviceUrl
        $fromId = $Message.rawmessage.from.id
        $fromName = $Message.RawMessage.from.name
        $recipientId = $Message.RawMessage.recipient.id
        $recipientName = $Message.RawMessage.recipient.name
        $conversationId = $Message.RawMessage.conversation.id
        $activityId = $Message.RawMessage.id
        $responseUrl = "$($baseUrl)v3/conversations/$conversationId/activities/$activityId"
        $channelId = $Message.RawMessage.channelData.teamsChannelId

        $headers = @{
            Authorization = "Bearer $($this.Connection._AccessTokenInfo.access_token)"
        }

        $conversationType = $Message.RawMessage.Conversation.ConversationType
        $isGroup = $Message.RawMessage.Conversation.isGroup


        # Currently only DMs work, but the documentation doesn't indicate that
        if ($Message.isDM) {
            $conversationId = $this._CreateDMConversation($fromId)
            $activityId = $conversationId
            $responseUrl = "$($baseUrl)v3/conversations/$conversationId/activities/"
        }

        if ($Type -eq [ReactionType]::Processing) {

            $cardBody = @{
                type         = 'typing'
                from         = @{
                    id   = $fromId
                    name = $fromName
                }
                conversation = @{
                    id   = $conversationId
                    name = ''
                }
                recipient    = @{
                    id   = $recipientId
                    name = $recipientName
                }
                replyToId    = $activityId
                channelId    = $channelId
            }

            $body = $cardBody | ConvertTo-Json -Depth 15
            Write-Verbose $body

            $this.LogDebug("Sending typing indicator to Teams conversation [$conversationId]", $body)

            try {
                $responseParams = @{
                    Uri         = $responseUrl
                    Method      = 'Post'
                    Body        = $body
                    ContentType = 'application/json'
                    Headers     = $headers
                }
                $teamsResponse = Invoke-RestMethod @responseParams
            } catch {
                $this.LogInfo([LogSeverity]::Error, "$($_.Exception.Message)", [ExceptionFormatter]::Summarize($_))
            }
        }
    }

    # Remove a reaction from an existing chat message
    [void]RemoveReaction([Message]$Message, [ReactionType]$Type, [string]$Reaction) {
        # NOT IMPLEMENTED YET
    }

    # Populate the list of users the team
    [void]LoadUsers() {
        if (-not [string]::IsNullOrEmpty($this.ServiceUrl)) {
            $this.LogDebug('Getting Teams users')

            $uri = "$($this.ServiceUrl)v3/conversations/$($this.TeamId)/members/"
            $headers = @{
                Authorization = "Bearer $($this.Connection._AccessTokenInfo.access_token)"
            }
            $members = Invoke-RestMethod -Uri $uri -Headers $headers
            $this.LogDebug('Finished getting Teams users')

            $members | Foreach-Object {
                $user = [TeamsPerson]::new()
                $user.Id                = $_.id
                $user.FirstName         = $_.givenName
                $user.LastName          = $_.surname
                $user.NickName          = $_.userPrincipalName
                $user.FullName          = "$($_.givenName) $($_.surname)"
                $user.Email             = $_.email
                $user.UserPrincipalName = $_.userPrincipalName

                if (-not $this.Users.ContainsKey($_.ID)) {
                    $this.LogDebug("Adding user [$($_.ID):$($_.Name)]")
                    $this.Users[$_.ID] =  $user
                }
            }

            foreach ($key in $this.Users.Keys) {
                if ($key -notin $members.ID) {
                    $this.LogDebug("Removing outdated user [$key]")
                    $this.Users.Remove($key)
                }
            }
        }
    }

    # Populate the list of channels in the team
    [void]LoadRooms() {
        #if (-not [string]::IsNullOrEmpty($this.TeamId)) {
            $this.LogDebug('Getting Teams channels')

            $uri = "$($this.ServiceUrl)v3/teams/$($this.TeamId)/conversations"
            $headers = @{
                Authorization = "Bearer $($this.Connection._AccessTokenInfo.access_token)"
            }
            $channels = Invoke-RestMethod -Uri $uri -Headers $headers

            if ($channels.conversations) {
                $channels.conversations | ForEach-Object {
                    $channel = [TeamsChannel]::new()
                    $channel.Id = $_.id
                    $channel.Name = $_.name
                    $this.LogDebug("Adding channel: $($_.id):$($_.name)")
                    $this.Rooms[$_.id] = $channel
                }

                foreach ($key in $this.Rooms.Keys) {
                    if ($key -notin $channels.conversations.ID) {
                        $this.LogDebug("Removing outdated channel [$key]")
                        $this.Rooms.Remove($key)
                    }
                }
            }
        #}
    }

    [bool]MsgFromBot([string]$From) {
        return $false
    }

    # Get a user by their Id
    [TeamsPerson]GetUser([string]$UserId) {
        $user = $this.Users[$UserId]
        if (-not $user) {
            $this.LogDebug([LogSeverity]::Warning, "User [$UserId] not found. Refreshing users")
            $this.LoadUsers()
            $user = $this.Users[$UserId]
        }

        if ($user) {
            $this.LogDebug("Resolved user [$UserId]", $user)
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve user [$UserId]")
        }
        return $user
    }

    # Get a user Id by their name
    [string]UsernameToUserId([string]$Username) {
        $Username = $Username.TrimStart('@')
        $user = $this.Users.Values | Where-Object {$_.Nickname -eq $Username}
        $id = $null
        if ($user) {
            $id = $user.Id
        } else {
            # User each doesn't exist or is not in the local cache
            # Refresh it and try again
            $this.LogDebug([LogSeverity]::Warning, "User [$Username] not found. Refreshing users")
            $this.LoadUsers()
            $user = $this.Users.Values | Where-Object {$_.Nickname -eq $Username}
            if (-not $user) {
                $id = $null
            } else {
                $id = $user.Id
            }
        }
        if ($id) {
            $this.LogDebug("Resolved [$Username] to [$id]")
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve user [$Username]")
        }
        return $id
    }

    # Get a user name by their Id
    [string]UserIdToUsername([string]$UserId) {
        $name = $null
        if ($this.Users.ContainsKey($UserId)) {
            $name = $this.Users[$UserId].Nickname
        } else {
            $this.LogDebug([LogSeverity]::Warning, "User [$UserId] not found. Refreshing users")
            $this.LoadUsers()
            $name = $this.Users[$UserId].Nickname
        }
        if ($name) {
            $this.LogDebug("Resolved [$UserId] to [$name]")
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve user [$UserId]")
        }
        return $name
    }

    # Get the channel name by Id
    [string]ChannelIdToName([string]$ChannelId) {
        $name = $null
        if ($this.Rooms.ContainsKey($ChannelId)) {
            $name = $this.Rooms[$ChannelId].Name
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Channel [$ChannelId] not found. Refreshing channels")
            $this.LoadRooms()
            $name = $this.Rooms[$ChannelId].Name
        }
        if ($name) {
            $this.LogDebug("Resolved [$ChannelId] to [$name]")
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve channel [$ChannelId]")
        }
        return $name
    }

    # Get all user info by their ID
    [hashtable]GetUserInfo([string]$UserId) {
        $user = $null
        if ($this.Users.ContainsKey($UserId)) {
            $user = $this.Users[$UserId]
        } else {
            $this.LogDebug([LogSeverity]::Warning, "User [$UserId] not found. Refreshing users")
            $this.LoadUsers()
            $user = $this.Users[$UserId]
        }

        if ($user) {
            $this.LogDebug("Resolved [$UserId] to [$($user.Nickname)]")
            return $user.ToHash()
        } else {
            $this.LogDebug([LogSeverity]::Warning, "Could not resolve channel [$UserId]")
            return $null
        }
    }

    hidden [void]DelayedInit([pscustomobject]$Message) {
        if ([string]::IsNullOrEmpty($this.ServiceUrl)) {
            $this.ServiceUrl = $Message.serviceUrl
            $this.LoadUsers()
            $this.LoadRooms()
        }

        if ([string]::IsNullOrEmpty($this.BotId)) {
            if ($Message.recipient) {
                $this.BotId   = $Message.recipient.Id
                $this.BotName = $Message.recipient.name
            }
        }

        if ([string]::IsNullOrEmpty($this.TenantId)) {
            if ($Message.channelData.tenant.id) {
                $this.TenantId = $Message.channelData.tenant.id
            }
        }
    }

    hidden [void]SendTeamsMessaage([Response]$Response) {
        $baseUrl        = $Response.OriginalMessage.RawMessage.serviceUrl
        $conversationId = $Response.OriginalMessage.RawMessage.conversation.id
        $activityId     = $Response.OriginalMessage.RawMessage.id
        $responseUrl    = "$($baseUrl)v3/conversations/$conversationId/activities/$activityId"
        $channelId      = $Response.OriginalMessage.RawMessage.channelData.teamsChannelId
        $headers = @{
            Authorization = "Bearer $($this.Connection._AccessTokenInfo.access_token)"
        }

        if ($Response.Text.Count -gt 0) {
            foreach ($text in $Response.Text) {
                $jsonResponse = @{
                    type = 'message'
                    from = @{
                        id = $Response.OriginalMessage.RawMessage.recipient.id
                        name = $Response.OriginalMessage.RawMessage.recipient.name
                    }
                    conversation = @{
                        id = $Response.OriginalMessage.RawMessage.conversation.id
                        name = ''
                    }
                    recipient = @{
                        id = $Response.OriginalMessage.RawMessage.from.id
                        name = $Response.OriginalMessage.RawMessage.from.name
                    }
                    text = $text
                    replyToId = $activityId
                } | ConvertTo-Json

                # $jsonResponse | Out-File -FilePath "$script:moduleBase/responses.json" -Append
                $this.LogDebug("Sending response back to Teams conversation [$conversationId]")
                try {
                    $responseParams = @{
                        Uri         = $responseUrl
                        Method      = 'Post'
                        Body        = $jsonResponse
                        ContentType = 'application/json'
                        Headers     = $headers
                    }
                    $teamsResponse = Invoke-RestMethod @responseParams
                } catch {
                    $this.LogInfo([LogSeverity]::Error, "$($_.Exception.Message)", [ExceptionFormatter]::Summarize($_))
                }
            }
        }
    }

    # Create a new DM conversation and return the converation ID
    # If there is an existing conversation, return that ID
    hidden [string]_CreateDMConversation([string]$UserId) {
        if ($this.DMConverations.ContainsKey($userId)) {
            return $this.DMConverations[$UserId]
        } else {
            $newConversationUrl = "$($this.ServiceUrl)v3/conversations"
            $headers = @{
                Authorization = "Bearer $($this.Connection._AccessTokenInfo.access_token)"
            }

            $conversationParams = @{
                bot = @{
                    id = $this.BotId
                    name = $this.BotName
                }
                members = @(
                    @{
                        id = $UserId
                    }
                )
                channelData = @{
                    tenant = @{
                        id = $this.TenantId
                    }
                }
            }

            $body = $conversationParams | ConvertTo-Json
            #$body | Out-File -FilePath "$script:moduleBase/create-dm.json" -Append
            $params = @{
                Uri         = $newConversationUrl
                Method      = 'Post'
                Body        = $body
                ContentType = 'application/json'
                Headers     = $headers
            }
            $conversation = Invoke-RestMethod @params
            if ($conversation) {
                $this.LogDebug("Created DM conversation [$($conversation.id)] with user [$UserId]")
                return $conversation.id
            } else {
                $this.LogInfo([LogSeverity]::Error, "$($_.Exception.Message)", [ExceptionFormatter]::Summarize($_))
                return $null
            }
        }
    }

    hidden [hashtable]_GetCardStub() {
        return @{
            type = 'message'
            from = @{
                id   = $null
                name = $null
            }
            conversation = @{
                id = $null
                #name = ''
            }
            recipient = @{
                id = $null
                name = $null
            }
            attachments = @(
                @{
                    contentType = 'application/vnd.microsoft.card.adaptive'
                    content = @{
                        type = 'AdaptiveCard'
                        version = '1.0'
                        fallbackText = $null
                        body = @(
                            @{
                                type = 'Container'
                                spacing = 'none'
                                items = @(
                                    # # Title & Thumbnail row
                                    @{
                                        type = 'ColumnSet'
                                        spacing = 'none'
                                        columns = @()
                                    }
                                    # Text & image row
                                    @{
                                        type = 'ColumnSet'
                                        spacing = 'none'
                                        columns = @()
                                    }
                                    # Facts row
                                    @{
                                        type = 'FactSet'
                                        facts = @()
                                    }
                                )
                            }
                        )
                    }
                }
            )
            replyToId = $null
        }
    }

    hidden [string]_RepairText([string]$Text) {
        if (-not [string]::IsNullOrEmpty($Text)) {
            $fixed = $Text.Replace('"', '\"').Replace('\', '\\').Replace("`n", '\n\n').Replace("`r", '').Replace("`t", '\t')
            $fixed = [System.Text.RegularExpressions.Regex]::Unescape($Text)
        } else {
            $fixed = ' '
        }

        return $fixed
    }

}

class TeamsChannel : Room {

}


class TeamsConnection : Connection {

    [object]$ReceiveJob = $null

    [System.Management.Automation.PowerShell]$PowerShell

    # To control the background thread
    [System.Collections.Concurrent.ConcurrentDictionary[string,psobject]]$ReceiverControl = [System.Collections.Concurrent.ConcurrentDictionary[string,psobject]]@{}

    # Shared queue between the class and the background thread to receive messages with
    [System.Collections.Concurrent.ConcurrentQueue[string]]$ReceiverMessages = [System.Collections.Concurrent.ConcurrentQueue[string]]@{}

    [object]$Handler = $null

    hidden [pscustomobject]$_AccessTokenInfo

    hidden [datetime]$_AccessTokenExpiration

    [bool]$Connected

    TeamsConnection([TeamsConnectionConfig]$Config) {
        $this.Config = $Config
    }

    # Setup runspace for the receiver thread to run in
    [void]Initialize() {
        $runspacePool = [RunspaceFactory]::CreateRunspacePool(1, 1)
        $runspacePool.Open()
        $this.PowerShell = [PowerShell]::Create()
        $this.PowerShell.RunspacePool = $runspacePool
        $this.ReceiverControl['ShouldRun'] = $true
    }

    # Connect to Teams and start receiving messages
    [void]Connect() {
        #if ($null -eq $this.ReceiveJob -or $this.ReceiveJob.State -ne 'Running') {
        if ($this.PowerShell.InvocationStateInfo.State -ne 'Running') {
            $this.Initialize()
            $this.Authenticate()
            $this.StartReceiveThread()
        } else {
            $this.LogDebug([LogSeverity]::Warning, 'Receive thread is already running')
        }
    }

    # Authenticate with Teams and get token
    [void]Authenticate() {
        try {
            $this.LogDebug('Getting Bot Framework access token')
            $authUrl = 'https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token'
            $payload = @{
                grant_type    = 'client_credentials'
                client_id     = $this.Config.Credential.Username
                client_secret = $this.Config.Credential.GetNetworkCredential().password
                scope         = 'https://api.botframework.com/.default'
            }
            $response = Invoke-RestMethod -Uri $authUrl -Method Post -Body $payload -Verbose:$false
            $this._AccessTokenExpiration = ([datetime]::Now).AddSeconds($response.expires_in)
            $this._AccessTokenInfo = $response
        } catch {
            $this.LogInfo([LogSeverity]::Error, 'Error authenticating to Teams', [ExceptionFormatter]::Summarize($_))
            throw $_
        }
    }

    [void]StartReceiveThread() {

        # Service Bus receive script
        $recv = {
            [cmdletbinding()]
            param(
                [parameter(Mandatory)]
                [System.Collections.Concurrent.ConcurrentDictionary[string,psobject]]$ReceiverControl,

                [parameter(Mandatory)]
                [System.Collections.Concurrent.ConcurrentQueue[string]]$ReceiverMessages,

                [parameter(Mandatory)]
                [string]$ModulePath,

                [parameter(Mandatory)]
                [string]$ServiceBusNamespace,

                [parameter(Mandatory)]
                [string]$QueueName,

                [parameter(Mandatory)]
                [string]$AccessKeyName,

                [parameter(Mandatory)]
                [string]$AccessKey
            )

            $connectionString = "Endpoint=sb://{0}.servicebus.windows.net/;SharedAccessKeyName={1};SharedAccessKey={2}" -f $ServiceBusNamespace, $AccessKeyName, $AccessKey
            $receiveTimeout = [timespan]::new(0, 0, 0, 5)

            # Honestly this is a pretty hacky way to go about using these
            # Service Bus DLLs but we can only implement one method or the
            # other without PSScriptAnalyzer freaking out about missing classes
            if ($PSVersionTable.PSEdition -eq 'Desktop') {
                . "$ModulePath/lib/windows/ServiceBusReceiver_net45.ps1"
            } else {
                . "$ModulePath/lib/linux/ServiceBusReceiver_netstandard.ps1"
            }
        }

        try {
            $cred = [pscredential]::new($this.Config.AccessKeyName, $this.Config.AccessKey)
            $runspaceParams = @{
                ReceiverControl     = $this.ReceiverControl
                ReceiverMessages    = $this.ReceiverMessages
                ModulePath          = $script:moduleBase
                ServiceBusNamespace = $this.Config.ServiceBusNamespace
                QueueName           = $this.Config.QueueName
                AccessKeyName       = $this.Config.AccessKeyName
                AccessKey           = $cred.GetNetworkCredential().password
            }

            $this.PowerShell.AddScript($recv)
            $this.PowerShell.AddParameters($runspaceParams) > $null
            $this.Handler = $this.PowerShell.BeginInvoke()
            $this.Connected = $true
            $this.Status = [ConnectionStatus]::Connected
            $this.LogInfo('Started Teams Service Bus background thread')
        } catch {
            $this.LogInfo([LogSeverity]::Error, "$($_.Exception.Message)", [ExceptionFormatter]::Summarize($_))
            $this.PowerShell.EndInvoke($this.Handler)
            $this.PowerShell.Dispose()
            $this.Connected = $false
            $this.Status = [ConnectionStatus]::Disconnected
        }
    }

    [string[]]ReadReceiveThread() {
        # # Read stream info from the job so we can log them
        # $infoStream    = $this.ReceiveJob.ChildJobs[0].Information.ReadAll()
        # $warningStream = $this.ReceiveJob.ChildJobs[0].Warning.ReadAll()
        # $errStream     = $this.ReceiveJob.ChildJobs[0].Error.ReadAll()
        # $verboseStream = $this.ReceiveJob.ChildJobs[0].Verbose.ReadAll()
        # $debugStream   = $this.ReceiveJob.ChildJobs[0].Debug.ReadAll()

        # foreach ($item in $infoStream) {
        #     $this.LogInfo($item.ToString())
        # }
        # foreach ($item in $warningStream) {
        #     $this.LogInfo([LogSeverity]::Warning, $item.ToString())
        # }
        # foreach ($item in $errStream) {
        #     $this.LogInfo([LogSeverity]::Error, $item.ToString())
        # }
        # foreach ($item in $verboseStream) {
        #     $this.LogVerbose($item.ToString())
        # }
        # foreach ($item in $debugStream) {
        #     $this.LogVerbose($item.ToString())
        # }

        # TODO
        # Read all the streams from the thread

        # Validate access token is still current and refresh
        # if expiration is less than half the token lifetime
        if (($this._AccessTokenExpiration - [datetime]::Now).TotalSeconds -lt 1800) {
            $this.LogDebug('Teams access token is expiring soon. Refreshing...')
            $this.Authenticate()
        }

        # The receive thread stopped for some reason. Reestablish the connection if it isn't running
        if ($this.PowerShell.InvocationStateInfo.State -ne 'Running') {

            # Log any errors from the background thread
            if ($this.PowerShell.Streams.Error.Count -gt 0) {
                $this.PowerShell.Streams.Error.Foreach({
                    $this.LogInfo([LogSeverity]::Error, "$($_.Exception.Message)", [ExceptionFormatter]::Summarize($_))
                })
            }
            $this.PowerShell.Streams.ClearStreams()

            $this.LogInfo([LogSeverity]::Warning, "Receive thread is [$($this.PowerShell.InvocationStateInfo.State)]. Attempting to reconnect...")
            Start-Sleep -Seconds 5
            $this.Connect()
        }

        # Dequeue messages from receiver thread
        if ($this.ReceiverMessages.Count -gt 0) {
            $dequeuedMessages = $null
            $messages = [System.Collections.Generic.LinkedList[string]]::new()
            while($this.ReceiverMessages.TryDequeue([ref]$dequeuedMessages)) {
                foreach ($m in $dequeuedMessages) {
                    $messages.Add($m) > $null
                }
            }
            return $messages
        } else {
            return $null
        }
    }

    # Stop the Teams listener
    [void]Disconnect() {
        $this.LogInfo('Stopping Service Bus receiver')
        $this.ReceiverControl.ShouldRun = $false
        $result = $this.PowerShell.EndInvoke($this.Handler)
        $this.PowerShell.Dispose()
        $this.Connected = $false
        $this.Status = [ConnectionStatus]::Disconnected
    }
}


class TeamsConnectionConfig : ConnectionConfig {
    [string]$BotName
    [string]$TeamId
    [string]$ServiceBusNamespace
    [string]$QueueName
    [string]$AccessKeyName
    [securestring]$AccessKey
}

class TeamsPerson : Person {
    [string]$Email
    [string]$UserPrincipalName
}

