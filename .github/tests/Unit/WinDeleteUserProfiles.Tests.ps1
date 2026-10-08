# WinToolkit CI/CD V4.2.0
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
Regression coverage for issue #189 without deleting Windows profiles.
Loads individual function definitions from the AST, never the tool or toolkit scripts.
Deletion commands are mocked; folder removal tests use empty TestDrive directories.
#>
param([string]$CompiledScriptPath)

Describe 'Issue #189 safe regressions' {
    BeforeAll {
        $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
        $toolPath = Join-Path $script:RepoRoot 'tools\WinDeleteUserProfiles.ps1'
        $localizationPath = Join-Path $script:RepoRoot 'wintoolkit-modules\10-Module.Localization.ps1'
        $loggingPath = Join-Path $script:RepoRoot 'wintoolkit-modules\30-Module.Logging.ps1'
        $processesPath = Join-Path $script:RepoRoot 'wintoolkit-modules\60-Module.Processes.ps1'
        $uiPath = Join-Path $script:RepoRoot 'wintoolkit-modules\20-Module.UI.ps1'

        $definitions = @(
            @{ Path = $toolPath; Name = 'New-ProfileRemovalSessionState' },
            @{ Path = $toolPath; Name = 'Receive-ProfileRemovalResult' },
            @{ Path = $toolPath; Name = 'Close-ProfileRemovalPowerShell' },
            @{ Path = $toolPath; Name = 'New-ProfileDirectoryEnumerator' },
            @{ Path = $toolPath; Name = 'Get-ProfileCleanupSubtrees' },
            @{ Path = $toolPath; Name = 'New-ProtectedNameSet' },
            @{ Path = $toolPath; Name = 'Get-RegisteredProfilePathSet' },
            @{ Path = $toolPath; Name = 'Get-ProfileCleanupPathAttribute' },
            @{ Path = $toolPath; Name = 'Test-ProfileCleanupDirectory' },
            @{ Path = $toolPath; Name = 'Assert-ProfileCleanupProfileState' },
            @{ Path = $toolPath; Name = 'Get-ResidualUserFolders' },
            @{ Path = $toolPath; Name = 'Remove-ProfileRegistryEntries' },
            @{ Path = $toolPath; Name = 'Remove-ResidualUserFolder' },
            @{ Path = $toolPath; Name = 'Remove-ResidualUserFolders' },
            @{ Path = $localizationPath; Name = 'Get-SourceTextLoc' },
            @{ Path = $loggingPath; Name = 'Write-ToolkitLog' },
            @{ Path = $toolPath; Name = 'Invoke-ProfileCleanupCommand' },
            @{ Path = $processesPath; Name = 'Remove-ItemSafely' },
            @{ Path = $uiPath; Name = 'Get-SpinnerChar' },
            @{ Path = $uiPath; Name = 'Write-StyledMessage' },
            @{ Path = $uiPath; Name = 'Clear-ProgressLine' },
            @{ Path = $uiPath; Name = 'Write-ProgressUpdate' }
        )
        foreach ($definition in $definitions) {
            $path = if ($CompiledScriptPath) { $CompiledScriptPath } else { $definition.Path }
            $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$errors)
            if ($errors) { throw "Invalid PowerShell syntax in $path" }
            $function = $ast.Find({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $definition.Name
            }, $true)
            if (-not $function) { throw "Missing helper $($definition.Name) in $path" }
            . ([scriptblock]::Create($function.Extent.Text))
        }

        $workerSourcePath = if ($CompiledScriptPath) { $CompiledScriptPath } else { $toolPath }
        $toolAst = [System.Management.Automation.Language.Parser]::ParseFile($workerSourcePath, [ref]$null, [ref]$null)
        $batchFunction = $toolAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-ProfileRemovalBatch'
        }, $true)
        $workerAssignment = $batchFunction.Body.Find({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$scriptBlock'
        }, $true)
        $workerExpression = $workerAssignment.Right.Find({
            param($node)
            $node -is [System.Management.Automation.Language.ScriptBlockExpressionAst]
        }, $true)
        $script:ProfileRemovalWorker = $workerExpression.ScriptBlock.GetScriptBlock()
        $script:ProfileRemovalBatchDefinition = $batchFunction.Extent.Text
        $script:ProfileRemovalBatchInvoker = $batchFunction.Body.GetScriptBlock()
        $script:ProfileRemovalWorkerOffset = $workerAssignment.Extent.StartOffset - $batchFunction.Extent.StartOffset
        $script:ProfileRemovalWorkerLength = $workerAssignment.Extent.EndOffset - $workerAssignment.Extent.StartOffset
        $script:ProfileBatchResultReceiver = (Get-Command Receive-ProfileRemovalResult).ScriptBlock
        $script:ResidualSessionStateFactory = (Get-Command New-ProfileRemovalSessionState).ScriptBlock
        $script:RegisteredProfilePathSetReader = (Get-Command Get-RegisteredProfilePathSet).ScriptBlock
        $residualBatchFunction = $toolAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Remove-ResidualUserFolders'
        }, $true)
        $residualWorkerAssignment = $residualBatchFunction.Body.Find({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$scriptBlock'
        }, $true)
        $script:ResidualRemovalBatchDefinition = $residualBatchFunction.Extent.Text
        $script:ResidualRemovalBatchInvoker = (Get-Command Remove-ResidualUserFolders).ScriptBlock
        $script:ResidualRemovalWorkerOffset = $residualWorkerAssignment.Extent.StartOffset - $residualBatchFunction.Extent.StartOffset
        $script:ResidualRemovalWorkerLength = $residualWorkerAssignment.Extent.EndOffset - $residualWorkerAssignment.Extent.StartOffset

        Import-LocalizedData -BindingVariable italianData -BaseDirectory (Join-Path $script:RepoRoot 'languages') -FileName 'WinToolkit.psd1' -UICulture 'it-IT'
        Import-LocalizedData -BindingVariable englishData -BaseDirectory (Join-Path $script:RepoRoot 'languages') -FileName 'WinToolkit.psd1' -UICulture 'en-US'
        $script:ItalianData = $italianData
        $script:EnglishData = $englishData
        $script:SavedGlobals = @{}
        foreach ($name in @(
            'SourceTextLanguageData', 'SourceTextDefaultLanguageData', 'SourceTextKeyAliases',
            'CurrentLogFile', 'CurrentToolName', 'Spinners', 'GuiSessionActive'
        )) {
            $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
            $script:SavedGlobals[$name] = @{ Exists = $null -ne $variable; Value = if ($variable) { $variable.Value } else { $null } }
        }

        function Invoke-ProfileWorkerProbe {
            param([scriptblock]$Probe, [int]$Count = 1)

            $sessionState = New-ProfileRemovalSessionState
            $pool = [RunspaceFactory]::CreateRunspacePool(1, 2, $sessionState, $Host)
            $jobs = @()
            try {
                $pool.Open()
                for ($index = 0; $index -lt $Count; $index++) {
                    $powerShell = [PowerShell]::Create()
                    $powerShell.RunspacePool = $pool
                    [void]$powerShell.AddScript($Probe.ToString())
                    $jobs += @{ PowerShell = $powerShell; Handle = $powerShell.BeginInvoke() }
                }
                foreach ($job in $jobs) {
                    $job.PowerShell.EndInvoke($job.Handle)
                }
            }
            finally {
                foreach ($job in $jobs) { $job.PowerShell.Dispose() }
                $pool.Close()
                $pool.Dispose()
            }
        }
    }

    BeforeEach {
        $Global:SourceTextLanguageData = $script:ItalianData.Clone()
        $Global:SourceTextDefaultLanguageData = $script:EnglishData.Clone()
        $Global:SourceTextKeyAliases = @{}
        $Global:CurrentLogFile = Join-Path $TestDrive 'profile-worker.log'
        $Global:CurrentToolName = 'Issue189Probe'
        $Global:Spinners = '|/-\'.ToCharArray()
        $Global:GuiSessionActive = $false
        Set-Content -LiteralPath $Global:CurrentLogFile -Value '' -Encoding UTF8
    }

    AfterAll {
        foreach ($name in $script:SavedGlobals.Keys) {
            $saved = $script:SavedGlobals[$name]
            if ($saved.Exists) {
                Set-Variable -Name $name -Scope Global -Value $saved.Value
            }
            else {
                Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'WinDeleteUserProfiles worker initialization without deletion' {
        It 'provides localization, logging, and every worker helper to pooled runspaces' {
            $results = @(Invoke-ProfileWorkerProbe -Count 4 -Probe {
                $message = Get-SourceTextLoc 'toolText.registeredProfilesRemoved0' -Args @(0)
                Write-ToolkitLog -Level INFO -Message $message -Context @{ Tool = $Global:CurrentToolName }
                [pscustomobject]@{
                    Message = $message
                    LogPath = $Global:CurrentLogFile
                    ToolName = $Global:CurrentToolName
                    Helpers = @(
                        'Get-SourceTextLoc', 'Write-ToolkitLog', 'Invoke-ProfileCleanupCommand',
                        'Get-ProfileCleanupPathAttribute',
                        'Test-ProfileCleanupDirectory', 'Assert-ProfileCleanupProfileState',
                        'Remove-ItemSafely', 'Remove-ProfileRegistryEntries', 'Remove-ResidualUserFolder', 'Get-SpinnerChar',
                        'Clear-ProgressLine', 'Write-ProgressUpdate'
                    ) | ForEach-Object { (Get-Command -Name $_ -CommandType Function -ErrorAction Stop).Name }
                }
            })

            $results.Count | Should -Be 4
            foreach ($result in $results) {
                $result.Message | Should -Be 'Profili registrati rimossi: 0'
                $result.LogPath | Should -Be $Global:CurrentLogFile
                $result.ToolName | Should -Be 'Issue189Probe'
                $result.Helpers.Count | Should -Be 12
            }
            $logLines = @(Get-Content -LiteralPath $Global:CurrentLogFile | Where-Object { $_ -match '\[INFO\]' })
            $logLines.Count | Should -Be 4
            $logLines | Should -Match 'Profili registrati rimossi: 0'
            $logLines | Should -Match 'Issue189Probe'
        }

        It 'keeps worker progress helpers silent while the parent owns the console' {
            $result = Invoke-ProfileWorkerProbe -Probe {
                Clear-ProgressLine
                Write-ProgressUpdate -Activity 'Probe' -Percent 50
                [pscustomobject]@{
                    Quiet = $Global:GuiSessionActive
                    ClearDefinition = (Get-Command Clear-ProgressLine).Definition
                    ProgressDefinition = (Get-Command Write-ProgressUpdate).Definition
                }
            }
            $result.Quiet | Should -BeTrue
            $result.ClearDefinition | Should -BeNullOrEmpty
            $result.ProgressDefinition | Should -BeNullOrEmpty
            $Global:GuiSessionActive | Should -BeFalse
        }

        It 'passes aliases and the English fallback to workers' {
            $Global:SourceTextKeyAliases['test.alias'] = 'toolText.registeredProfilesRemoved0'
            $Global:SourceTextDefaultLanguageData['test.fallback'] = 'Fallback: {0}'
            $result = Invoke-ProfileWorkerProbe -Probe {
                [pscustomobject]@{
                    Alias = Get-SourceTextLoc 'test.alias' -Args @(0)
                    Fallback = Get-SourceTextLoc 'test.fallback' -Args @(0)
                }
            }
            $result.Alias | Should -Be 'Profili registrati rimossi: 0'
            $result.Fallback | Should -Be 'Fallback: 0'
        }

        It 'fails initialization when a required helper is unavailable' {
            Mock Get-Command { throw 'Required helper unavailable' } -ParameterFilter { $Name -eq 'Remove-ItemSafely' }
            { New-ProfileRemovalSessionState } | Should -Throw '*Required helper unavailable*'
        }
    }

    Context 'Profile result collection without deletion' {
        It 'disposes a pipeline even when clearing its command references throws' {
            $commands = [pscustomobject]@{}
            $commands | Add-Member -MemberType ScriptMethod -Name Clear -Value { throw 'Synthetic clear failure' }
            $powerShell = [pscustomobject]@{ Commands = $commands; Disposed = $false }
            $powerShell | Add-Member -MemberType ScriptMethod -Name EndInvoke -Value { param($Handle) [pscustomobject]@{ Type = 'Profile'; Success = $true } }
            $powerShell | Add-Member -MemberType ScriptMethod -Name Dispose -Value { $this.Disposed = $true }
            $job = [pscustomobject]@{ PowerShell = $powerShell; Handle = $null }

            { Receive-ProfileRemovalResult -Job $job } | Should -Throw '*Synthetic clear failure*'

            $powerShell.Disposed | Should -BeTrue
        }
        It 'records a worker exception as one failed profile with its identity' {
            $powerShell = [PowerShell]::Create()
            [void]$powerShell.AddScript("throw 'Synthetic runspace failure'")
            $job = [pscustomobject]@{
                PowerShell = $powerShell
                Handle = $powerShell.BeginInvoke()
                Profile = [pscustomobject]@{ LocalPath = 'C:\Users\Issue189Synthetic'; SID = 'S-1-5-21-189-1' }
            }
            $results = @(Receive-ProfileRemovalResult -Job $job)

            $results.Count | Should -Be 1
            $results[0].Type | Should -Be 'Profile'
            $results[0].UserName | Should -Be 'Issue189Synthetic'
            $results[0].Path | Should -Be $job.Profile.LocalPath
            $results[0].Sid | Should -Be $job.Profile.SID
            $results[0].Success | Should -BeFalse
            @($results | Where-Object { -not $_.Success }).Count | Should -Be 1
            Get-Content -LiteralPath $Global:CurrentLogFile -Raw | Should -Match '\[ERROR\].*Synthetic runspace failure'
        }

        It 'preserves a structured <Success> worker result without adding other output' -ForEach @(
            @{ Success = $true }, @{ Success = $false }
        ) {
            $powerShell = [PowerShell]::Create()
            [void]$powerShell.AddScript('param($success) [pscustomobject]@{ Type = "Profile"; Success = $success }').AddArgument($Success)
            $job = [pscustomobject]@{ PowerShell = $powerShell; Handle = $powerShell.BeginInvoke(); Profile = $null }
            $results = @(Receive-ProfileRemovalResult -Job $job)
            $results.Count | Should -Be 1
            $results[0].Type | Should -Be 'Profile'
            $results[0].Success | Should -Be $Success
        }
    }

    Context 'Bounded native command capture with harmless child processes' {
        BeforeAll {
            $script:ProbeExecutable = (Get-Process -Id $PID).Path
            function ConvertTo-ProfileProbeArguments {
                param([string]$Code)
                @('-NoLogo', '-NoProfile', '-NonInteractive', '-OutputFormat', 'Text', '-EncodedCommand',
                    [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes('$ProgressPreference = ''SilentlyContinue''; ' + $Code)))
            }
        }

        It 'drains megabytes on both pipes without retaining more than the diagnostic snippets' {
            $arguments = ConvertTo-ProfileProbeArguments '[Console]::Out.Write((''o'' * 4194304)); [Console]::Error.Write((''e'' * 4194304))'
            $result = Invoke-ProfileCleanupCommand -Command $script:ProbeExecutable -Arguments $arguments

            $result.Success | Should -BeTrue
            $result.StdOut | Should -Be (('o' * 8000) + "`n[...output truncated...]")
            # A restricted child host may prepend provider initialization diagnostics to stderr.
            $result.StdErr.Length | Should -BeLessOrEqual 8025
            $result.StdErr | Should -Match ('e' * 4096)
            $result.StdErr.EndsWith("`n[...stderr truncated...]") | Should -BeTrue
        }

        It 'preserves short output and a nonzero exit code' {
            $arguments = ConvertTo-ProfileProbeArguments '[Console]::Out.Write(''short output''); [Console]::Error.Write(''short error''); exit 3'
            $result = Invoke-ProfileCleanupCommand -Command $script:ProbeExecutable -Arguments $arguments

            $result.Success | Should -BeFalse
            $result.ExitCode | Should -Be 3
            $result.StdOut | Should -Be 'short output'
            $result.StdErr | Should -Match 'short error$'
        }

        It 'reports a start failure' {
            $result = Invoke-ProfileCleanupCommand -Command (Join-Path $TestDrive 'missing-command.exe')

            $result.Success | Should -BeFalse
            $result.ExitCode | Should -Be -1
            $result.StdOut | Should -BeNullOrEmpty
            $result.StdErr | Should -BeNullOrEmpty
        }

        It 'allows a finite child to finish without a command timeout' {
            $arguments = ConvertTo-ProfileProbeArguments '[Threading.Thread]::Sleep(2200); [Console]::Out.Write(''completed'')'
            $result = Invoke-ProfileCleanupCommand -Command $script:ProbeExecutable -Arguments $arguments

            $result.Success | Should -BeTrue
            $result.ExitCode | Should -Be 0
            $result.StdOut | Should -Be 'completed'
        }

        It 'cancels a running child promptly and closes its async wait handle' {
            $marker = Join-Path $TestDrive 'cancel-child.pid'
            $code = '[IO.File]::WriteAllText(''{0}'', $PID.ToString()); [Threading.Thread]::Sleep(10000)' -f $marker.Replace("'", "''")
            $state = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
            $state.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Invoke-ProfileCleanupCommand', (Get-Command Invoke-ProfileCleanupCommand).Definition))
            $state.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Write-ToolkitLog', 'param($Level, $Message, $Context)'))
            $state.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Get-SourceTextLoc', 'param($Key, $Args) $Key'))
            $worker = [PowerShell]::Create($state)
            $handle = $null
            $waitHandle = $null
            try {
                [void]$worker.AddScript('param($Executable, $Arguments) Invoke-ProfileCleanupCommand -Command $Executable -Arguments $Arguments', $true).
                    AddArgument($script:ProbeExecutable).AddArgument((ConvertTo-ProfileProbeArguments $code))
                $handle = $worker.BeginInvoke()
                $waitHandle = $handle.AsyncWaitHandle
                $startup = [Diagnostics.Stopwatch]::StartNew()
                while (-not [IO.File]::Exists($marker) -and $startup.Elapsed.TotalSeconds -lt 8) { [Threading.Thread]::Sleep(20) }
                [IO.File]::Exists($marker) | Should -BeTrue
                $timer = [Diagnostics.Stopwatch]::StartNew()
                Close-ProfileRemovalPowerShell -PowerShell $worker -Handle $handle -WaitHandle $waitHandle
                $timer.Stop()

                $timer.Elapsed.TotalSeconds | Should -BeLessThan 4
                $waitHandle.SafeWaitHandle.IsClosed | Should -BeTrue
                Get-Process -Id ([int][IO.File]::ReadAllText($marker)) -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
            }
            finally { Close-ProfileRemovalPowerShell -PowerShell $worker -Handle $handle -WaitHandle $waitHandle }
        }
    }

    Context 'Directory absence proof without profile deletion' {
        It 'recognizes a directory with literal wildcard characters' {
            $path = Join-Path $TestDrive 'Inspection[Literal]'
            [void][System.IO.Directory]::CreateDirectory($path)

            Test-ProfileCleanupDirectory -Path $path | Should -BeTrue
        }

        It 'recognizes a confirmed missing directory' {
            Test-ProfileCleanupDirectory -Path (Join-Path $TestDrive 'Missing[Literal]') | Should -BeFalse
        }

        It 'does not turn <Failure> into confirmed absence' -ForEach @(
            @{ Failure = 'Access' }, @{ Failure = 'IO' }
        ) {
            Mock Get-ProfileCleanupPathAttribute {
                if ($Failure -eq 'Access') { throw [System.UnauthorizedAccessException]::new('Synthetic access failure') }
                throw [System.IO.IOException]::new('Synthetic I/O failure')
            }

            { Test-ProfileCleanupDirectory -Path 'C:\ReviewMock\Unknown' } | Should -Throw
        }

        It 'does not mistake an existing file for a deleted directory' {
            $path = Join-Path $TestDrive 'UnexpectedFile.txt'
            Set-Content -LiteralPath $path -Value 'Owned test fixture'

            { Test-ProfileCleanupDirectory -Path $path } | Should -Throw '*not a directory*'
        }

        It 'does not report absence when <MissingError> also hides an unavailable volume' -ForEach @(
            @{ MissingError = 'File' }, @{ MissingError = 'Directory' }
        ) {
            Mock Get-ProfileCleanupPathAttribute {
                if ($Path -eq 'C:\') { throw [System.IO.IOException]::new('Synthetic volume unavailable') }
                if ($MissingError -eq 'File') { throw [System.IO.FileNotFoundException]::new('Synthetic missing path') }
                throw [System.IO.DirectoryNotFoundException]::new('Synthetic missing path')
            }

            { Test-ProfileCleanupDirectory -Path 'C:\ReviewMock\Missing' } | Should -Throw '*volume unavailable*'
        }
    }

    Context 'Folder removal failures with mocked Windows operations' {
        BeforeEach {
            $script:ProbeFolderPath = Join-Path $TestDrive 'Issue189[Probe]'
            [void][System.IO.Directory]::CreateDirectory($script:ProbeFolderPath)
            # Preserve the cmdlet's input type so Pester observes CIM calls instead of binding failures.
            $script:ProbeProfile = [Microsoft.Management.Infrastructure.CimInstance]::new('Win32_UserProfile')
            foreach ($property in @(
                    @{ Name = 'LocalPath'; Value = $script:ProbeFolderPath },
                    @{ Name = 'SID'; Value = 'S-1-5-21-189-1' }
                )) {
                $script:ProbeProfile.CimInstanceProperties.Add(
                    [Microsoft.Management.Infrastructure.CimProperty]::Create($property.Name, $property.Value,
                        [Microsoft.Management.Infrastructure.CimType]::String, [Microsoft.Management.Infrastructure.CimFlags]::None)
                )
            }
            $script:RemovalAttempts = 0
            $script:FreshProfileLoaded = $false
            Mock Get-CimInstance {
                [pscustomobject]@{
                    SID = $script:ProbeProfile.SID; LocalPath = $script:ProbeFolderPath
                    Loaded = $script:FreshProfileLoaded; Special = $false
                }
            }
            Mock Remove-CimInstance { throw 'Synthetic CIM removal failure' }
            Mock Remove-ProfileRegistryEntries { $true }
            Mock Invoke-ProfileCleanupCommand { [pscustomobject]@{ Success = $true; ExitCode = 0 } }
            Mock Remove-ItemSafely { $false }
            Mock Remove-Item { throw 'Unconfigured deletion probe' }
            Mock Test-Path { $true } -ParameterFilter { $Path -eq (Join-Path $env:TEMP 'EmptyFolder') }
            Mock Clear-ProgressLine {}
            Mock Write-ProgressUpdate {}
        }

        AfterEach { $script:ProbeProfile.Dispose() }

        It 'recovers with ACL reset when the first literal folder removal fails' {
            Mock Remove-Item {
                $script:RemovalAttempts++
                if ($script:RemovalAttempts -eq 1) { throw 'Synthetic access denied' }
                if ($LiteralPath -ne $script:ProbeFolderPath) { throw 'Unexpected deletion target' }
                [System.IO.Directory]::Delete($LiteralPath)
            }

            $result = & $script:ProfileRemovalWorker $script:ProbeProfile

            $result.Success | Should -BeTrue
            [System.IO.Directory]::Exists($script:ProbeFolderPath) | Should -BeFalse
            Should -Invoke Remove-Item -Times 2 -Exactly -ParameterFilter {
                $LiteralPath -eq $script:ProbeFolderPath -and $Recurse -and $Force -and $ErrorAction -eq 'Stop'
            }
            Should -Invoke Invoke-ProfileCleanupCommand -Times 1 -Exactly -ParameterFilter { $Command -eq 'takeown.exe' }
            Should -Invoke Invoke-ProfileCleanupCommand -Times 1 -Exactly -ParameterFilter {
                $Command -eq 'robocopy.exe' -and $Arguments -contains '/XJ'
            }
            Should -Invoke Invoke-ProfileCleanupCommand -Times 1 -Exactly -ParameterFilter {
                $Command -eq 'icacls.exe' -and $Arguments -contains '*S-1-5-32-544:F' -and $Arguments -contains '/L'
            }
            Should -Invoke Invoke-ProfileCleanupCommand -Times 1 -Exactly -ParameterFilter {
                $Command -eq 'takeown.exe' -and $Arguments -contains '/SKIPSL'
            }
            Should -Invoke Remove-ProfileRegistryEntries -Times 1 -Exactly -ParameterFilter { $Sid -eq $script:ProbeProfile.SID }
            Should -Invoke Remove-ItemSafely -Times 0 -Exactly
        }

        It 'reports failure and preserves the registry in <Phase> when both folder removal attempts fail' -ForEach @(
            @{ Phase = 'Complete' }, @{ Phase = 'Finalize' }
        ) {
            Mock Remove-Item { throw 'Synthetic access denied' }

            $result = & $script:ProfileRemovalWorker $script:ProbeProfile $Phase

            $result.Success | Should -BeFalse
            [System.IO.Directory]::Exists($script:ProbeFolderPath) | Should -BeTrue
            Should -Invoke Remove-Item -Times 2 -Exactly
            Should -Invoke Remove-ProfileRegistryEntries -Times 0 -Exactly
            Get-Content -LiteralPath $Global:CurrentLogFile -Raw | Should -Match '\[ERROR\].*Synthetic access denied'
        }

        It 'prepares remaining profile contents without starting filesystem or registry cleanup' {
            $result = & $script:ProfileRemovalWorker $script:ProbeProfile 'Prepare'

            $result.Type | Should -Be 'ProfilePreparation'
            $result.Start | Should -BeOfType ([datetime])
            Should -Invoke Remove-CimInstance -Times 1 -Exactly
            Should -Invoke Remove-Item -Times 0 -Exactly
            Should -Invoke Invoke-ProfileCleanupCommand -Times 0 -Exactly
            Should -Invoke Remove-ProfileRegistryEntries -Times 0 -Exactly
        }

        It 'finishes in preparation when CIM has already removed the folder' {
            Mock Remove-CimInstance { [System.IO.Directory]::Delete($script:ProbeFolderPath) }

            $result = & $script:ProfileRemovalWorker $script:ProbeProfile 'Prepare'

            $result.Type | Should -Be 'Profile'
            $result.Success | Should -BeTrue
            Should -Invoke Remove-Item -Times 0 -Exactly
            Should -Invoke Remove-ProfileRegistryEntries -Times 1 -Exactly
        }

        It 'removes a child subtree without repeating CIM or touching profile registration' {
            Mock Remove-Item {
                if ($LiteralPath -ne $script:ProbeFolderPath) { throw 'Unexpected deletion target' }
                [System.IO.Directory]::Delete($LiteralPath)
            }

            $result = & $script:ProfileRemovalWorker $script:ProbeProfile 'Folder'

            $result.Type | Should -Be 'ResidualFolder'
            $result.Path | Should -Be $script:ProbeFolderPath
            $result.Success | Should -BeTrue
            Should -Invoke Remove-CimInstance -Times 0 -Exactly
            Should -Invoke Remove-ProfileRegistryEntries -Times 0 -Exactly
            Should -Invoke Remove-Item -Times 1 -Exactly -ParameterFilter { $LiteralPath -eq $script:ProbeFolderPath }
        }

        It 'skips a child that has become a reparse point after partitioning' {
            Mock Get-ProfileCleanupPathAttribute { [System.IO.FileAttributes]::ReparsePoint }

            { & $script:ProfileRemovalWorker $script:ProbeProfile 'Folder' } | Should -Throw '*reparse point*'
            Should -Invoke Remove-Item -Times 0 -Exactly
            Should -Invoke Invoke-ProfileCleanupCommand -Times 0 -Exactly
            Should -Invoke Remove-ProfileRegistryEntries -Times 0 -Exactly
        }

        It 'rejects a reparse root before any cleanup in <Phase>' -ForEach @(
            @{ Phase = 'Prepare' }, @{ Phase = 'Complete' }, @{ Phase = 'Finalize' }
        ) {
            Mock Get-ProfileCleanupPathAttribute { [System.IO.FileAttributes]::ReparsePoint }

            { & $script:ProfileRemovalWorker $script:ProbeProfile $Phase } | Should -Throw '*reparse point*'

            Should -Invoke Remove-CimInstance -Times 0 -Exactly
            Should -Invoke Remove-Item -Times 0 -Exactly
            Should -Invoke Invoke-ProfileCleanupCommand -Times 0 -Exactly
            Should -Invoke Remove-ProfileRegistryEntries -Times 0 -Exactly
        }

        It 'does not bypass a failed CIM removal when fresh state is <State>' -ForEach @(
            @{ State = 'Loaded' }, @{ State = 'Special' }, @{ State = 'Missing' }, @{ State = 'Unknown' }, @{ State = 'ChangedPath' }
        ) {
            Mock Get-CimInstance {
                switch ($State) {
                    'Unknown' { throw 'Synthetic profile-state detection failure' }
                    'Missing' { return }
                    default {
                        [pscustomobject]@{
                            SID = $script:ProbeProfile.SID
                            LocalPath = $(if ($State -eq 'ChangedPath') { $script:ProbeFolderPath + '-Changed' } else { $script:ProbeFolderPath })
                            Loaded = $State -eq 'Loaded'; Special = $State -eq 'Special'
                        }
                    }
                }
            }

            { & $script:ProfileRemovalWorker $script:ProbeProfile 'Prepare' } | Should -Throw

            Should -Invoke Remove-CimInstance -Times 1 -Exactly
            Should -Invoke Get-CimInstance -Times 1 -Exactly -ParameterFilter {
                $ClassName -eq 'Win32_UserProfile' -and $Filter -eq "SID='$($script:ProbeProfile.SID)'" -and $ErrorAction -eq 'Stop'
            }
            Should -Invoke Remove-Item -Times 0 -Exactly
            Should -Invoke Invoke-ProfileCleanupCommand -Times 0 -Exactly
            Should -Invoke Remove-ProfileRegistryEntries -Times 0 -Exactly
        }

        It 'rechecks the parent profile in <Phase> after it becomes loaded' -ForEach @(
            @{ Phase = 'Folder' }, @{ Phase = 'Finalize' }
        ) {
            $preparation = & $script:ProfileRemovalWorker $script:ProbeProfile 'Prepare'
            $preparation.CheckProfileState | Should -BeTrue
            $script:FreshProfileLoaded = $true

            { & $script:ProfileRemovalWorker $script:ProbeProfile $Phase $preparation.Start $script:ProbeProfile $preparation.CheckProfileState } | Should -Throw '*safely verified*'

            Should -Invoke Remove-Item -Times 0 -Exactly
            Should -Invoke Invoke-ProfileCleanupCommand -Times 0 -Exactly
            Should -Invoke Remove-ProfileRegistryEntries -Times 0 -Exactly
        }

        It 'does not requery removed registration after successful CIM cleanup leaves a folder' {
            Mock Remove-CimInstance {}

            $preparation = & $script:ProfileRemovalWorker $script:ProbeProfile 'Prepare'

            $preparation.CheckProfileState | Should -BeFalse
            Should -Invoke Get-CimInstance -Times 0 -Exactly
        }

        It 'stops deletion and ACL recovery when robocopy is followed by a replaced link' {
            $script:TargetBecameLink = $false
            Mock Get-ProfileCleanupPathAttribute {
                if ($script:TargetBecameLink) { [System.IO.FileAttributes]::ReparsePoint }
                else { [System.IO.FileAttributes]::Directory }
            }
            Mock Invoke-ProfileCleanupCommand { $script:TargetBecameLink = $true }

            { & $script:ProfileRemovalWorker $script:ProbeProfile 'Finalize' } | Should -Throw '*reparse point*'

            Should -Invoke Invoke-ProfileCleanupCommand -Times 1 -Exactly -ParameterFilter { $Command -eq 'robocopy.exe' }
            Should -Invoke Invoke-ProfileCleanupCommand -Times 0 -Exactly -ParameterFilter { $Command -in @('takeown.exe', 'icacls.exe') }
            Should -Invoke Remove-Item -Times 0 -Exactly
            Should -Invoke Remove-ProfileRegistryEntries -Times 0 -Exactly
        }

        It 'keeps registration when directory inspection fails before cleanup' {
            Mock Get-ProfileCleanupPathAttribute { throw [System.UnauthorizedAccessException]::new('Synthetic directory denial') }

            { & $script:ProfileRemovalWorker $script:ProbeProfile 'Finalize' } | Should -Throw '*Synthetic directory denial*'

            Should -Invoke Remove-Item -Times 0 -Exactly
            Should -Invoke Invoke-ProfileCleanupCommand -Times 0 -Exactly
            Should -Invoke Remove-ProfileRegistryEntries -Times 0 -Exactly
        }

        It 'keeps registration when post-deletion inspection fails' {
            $script:InspectionFailed = $false
            Mock Get-ProfileCleanupPathAttribute {
                if ($script:InspectionFailed) { throw [System.IO.IOException]::new('Synthetic verification I/O failure') }
                [System.IO.FileAttributes]::Directory
            }
            # The mocked command changes inspection state and never deletes the fixture.
            Mock Remove-Item { $script:InspectionFailed = $true }

            { & $script:ProfileRemovalWorker $script:ProbeProfile 'Finalize' } | Should -Throw '*Synthetic verification I/O failure*'

            Should -Invoke Remove-Item -Times 1 -Exactly
            Should -Invoke Remove-ProfileRegistryEntries -Times 0 -Exactly
            [System.IO.Directory]::Exists($script:ProbeFolderPath) | Should -BeTrue
        }

        It 'rejects a residual candidate replaced by a link before its worker starts' {
            Mock Get-ProfileCleanupPathAttribute { [System.IO.FileAttributes]::ReparsePoint }
            $folder = [pscustomobject]@{ Name = 'ReplacedLink'; Path = $script:ProbeFolderPath }

            (Remove-ResidualUserFolder -Folder $folder).Success | Should -BeFalse

            Should -Invoke Remove-Item -Times 0 -Exactly
            Should -Invoke Invoke-ProfileCleanupCommand -Times 0 -Exactly
        }

        It 'finalizes the root with the original start time and without repeating CIM removal' {
            Mock Remove-Item {
                if ($LiteralPath -ne $script:ProbeFolderPath) { throw 'Unexpected deletion target' }
                [System.IO.Directory]::Delete($LiteralPath)
            }
            $start = (Get-Date).AddSeconds(-10)

            $result = & $script:ProfileRemovalWorker $script:ProbeProfile 'Finalize' $start

            $result.Type | Should -Be 'Profile'
            $result.Sid | Should -Be $script:ProbeProfile.SID
            $result.Success | Should -BeTrue
            $result.Duration.TotalSeconds | Should -BeGreaterOrEqual 10
            Should -Invoke Remove-CimInstance -Times 0 -Exactly
            Should -Invoke Remove-ProfileRegistryEntries -Times 1 -Exactly
        }

        It 'checks the remaining registry entry even when CIM reported success' {
            Mock Remove-CimInstance {}
            Mock Remove-ProfileRegistryEntries { $false }
            Mock Remove-Item {
                if ($LiteralPath -ne $script:ProbeFolderPath) { throw 'Unexpected deletion target' }
                [System.IO.Directory]::Delete($LiteralPath)
            }

            $result = & $script:ProfileRemovalWorker $script:ProbeProfile

            $result.Success | Should -BeFalse
            [System.IO.Directory]::Exists($script:ProbeFolderPath) | Should -BeFalse
            Should -Invoke Remove-ProfileRegistryEntries -Times 1 -Exactly
        }

        It 'retries a residual folder using its literal path after ACL reset' {
            Mock Remove-Item {
                $script:RemovalAttempts++
                if ($script:RemovalAttempts -eq 1) { throw 'Synthetic access denied' }
                if ($LiteralPath -ne $script:ProbeFolderPath) { throw 'Unexpected deletion target' }
                [System.IO.Directory]::Delete($LiteralPath)
            }
            $folder = [pscustomobject]@{ Name = 'Issue189[Probe]'; Path = $script:ProbeFolderPath }

            $result = Remove-ResidualUserFolder -Folder $folder

            $result.Success | Should -BeTrue
            Should -Invoke Remove-Item -Times 2 -Exactly -ParameterFilter {
                $LiteralPath -eq $script:ProbeFolderPath -and $Recurse -and $Force -and $ErrorAction -eq 'Stop'
            }
            Should -Invoke Invoke-ProfileCleanupCommand -Times 1 -Exactly -ParameterFilter {
                $Command -eq 'icacls.exe' -and $Arguments -contains '*S-1-5-32-544:F'
            }
        }

        It 'reports a residual folder failure when ACL recovery cannot remove it' {
            Mock Remove-Item { throw 'Synthetic access denied' }
            $folder = [pscustomobject]@{ Name = 'Issue189[Probe]'; Path = $script:ProbeFolderPath }

            $result = Remove-ResidualUserFolder -Folder $folder

            $result.Success | Should -BeFalse
            [System.IO.Directory]::Exists($script:ProbeFolderPath) | Should -BeTrue
            Should -Invoke Remove-Item -Times 2 -Exactly
            Get-Content -LiteralPath $Global:CurrentLogFile -Raw | Should -Match '\[ERROR\].*Synthetic access denied'
        }
    }

    Context 'Localized arguments with zero or empty values' {
        It 'formats the <Name> argument' -ForEach @(
            @{ Name = 'zero'; Value = 0; Expected = 'Value: 0' },
            @{ Name = 'false'; Value = $false; Expected = 'Value: False' },
            @{ Name = 'empty string'; Value = ''; Expected = 'Value: ' },
            @{ Name = 'null'; Value = $null; Expected = 'Value: ' },
            @{ Name = 'nonzero'; Value = 2; Expected = 'Value: 2' }
        ) {
            $Global:SourceTextLanguageData['test.argument'] = 'Value: {0}'
            Get-SourceTextLoc 'test.argument' -Args @($Value) | Should -Be $Expected
        }

        It 'leaves the template intact when no arguments are supplied' {
            $Global:SourceTextLanguageData['test.argument'] = 'Value: {0}'
            Get-SourceTextLoc 'test.argument' | Should -Be 'Value: {0}'
        }

        It 'prints both Italian zero counts in the cleanup summary' {
            Get-SourceTextLoc 'toolText.registeredProfilesRemoved0' -Args @(0) | Should -Be 'Profili registrati rimossi: 0'
            Get-SourceTextLoc 'toolText.residualFoldersRemoved0' -Args @(0) | Should -Be 'Cartelle residue rimosse: 0'
        }
    }

    Context 'Profile subtree planning without deletion' {
        BeforeAll {
            if (-not ('ProfileCleanupEnumeratorProbe' -as [type])) {
                Add-Type -TypeDefinition @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
public sealed class ProfileCleanupEnumeratorProbe : IEnumerator<string> {
    private readonly string prefix;
    private readonly int count;
    private int index = -1;
    public int Reads;
    public bool Disposed;
    public int FailAfter = -1;
    public ProfileCleanupEnumeratorProbe(string prefix, int count) { this.prefix = prefix; this.count = count; }
    public string Current { get { return prefix + index.ToString(); } }
    object IEnumerator.Current { get { return Current; } }
    public bool MoveNext() {
        Reads++;
        if (FailAfter >= 0 && index >= FailAfter) { throw new IOException("Synthetic partial enumeration failure"); }
        index++;
        return index < count;
    }
    public void Reset() { throw new NotSupportedException(); }
    public void Dispose() { Disposed = true; }
}
'@
            }
        }
        BeforeEach {
            $script:PlanRoot = 'C:\Users\Plan[Profile]'
            $script:PlanTree = @{
                $script:PlanRoot = @(
                    [pscustomobject]@{ FullName = "$script:PlanRoot\AppData"; Attributes = [System.IO.FileAttributes]::Directory },
                    [pscustomobject]@{ FullName = "$script:PlanRoot\Documents"; Attributes = [System.IO.FileAttributes]::Directory },
                    [pscustomobject]@{ FullName = "$script:PlanRoot\Junction"; Attributes = [System.IO.FileAttributes]::ReparsePoint }
                )
                "$script:PlanRoot\AppData" = @(
                    [pscustomobject]@{ FullName = "$script:PlanRoot\AppData\Local"; Attributes = [System.IO.FileAttributes]::Directory },
                    [pscustomobject]@{ FullName = "$script:PlanRoot\AppData\Roaming"; Attributes = [System.IO.FileAttributes]::Directory },
                    [pscustomobject]@{ FullName = "$script:PlanRoot\AppData\Link"; Attributes = [System.IO.FileAttributes]::ReparsePoint }
                )
                "$script:PlanRoot\Documents" = @()
            }
            Mock Get-Item {
                foreach ($entries in $script:PlanTree.Values) {
                    foreach ($entry in $entries) { if ($entry.FullName -eq $LiteralPath) { return $entry } }
                }
                [pscustomobject]@{ Attributes = [System.IO.FileAttributes]::Directory }
            }
            Mock New-ProfileDirectoryEnumerator {
                return , (@($script:PlanTree[$Path] | ForEach-Object { $_.FullName }).GetEnumerator())
            }
        }

        It 'returns disjoint child paths and never traverses a reparse point or enumerates files' {
            $paths = @(Get-ProfileCleanupSubtrees -Path $script:PlanRoot)

            $paths.Count | Should -Be 3
            $paths[0] | Should -Be "$script:PlanRoot\AppData\Local"
            $paths[1] | Should -Be "$script:PlanRoot\AppData\Roaming"
            $paths[2] | Should -Be "$script:PlanRoot\Documents"
            $paths | Should -Not -Contain "$script:PlanRoot\AppData"
            Should -Invoke New-ProfileDirectoryEnumerator -Times 3 -Exactly
            Should -Invoke New-ProfileDirectoryEnumerator -Times 0 -Exactly -ParameterFilter { $Path -like '*\Link' -or $Path -like '*\Junction' }
        }

        It 'leaves a reparse profile root to the original final cleanup' {
            Mock Get-Item { [pscustomobject]@{ Attributes = [System.IO.FileAttributes]::ReparsePoint } }

            @(Get-ProfileCleanupSubtrees -Path $script:PlanRoot).Count | Should -Be 0

            Should -Invoke New-ProfileDirectoryEnumerator -Times 0 -Exactly
        }

        It 'falls back to whole-profile cleanup when root enumeration fails' {
            Mock New-ProfileDirectoryEnumerator { throw 'Synthetic root enumeration failure' }

            @(Get-ProfileCleanupSubtrees -Path $script:PlanRoot).Count | Should -Be 0

            Get-Content -LiteralPath $Global:CurrentLogFile -Raw | Should -Match 'Synthetic root enumeration failure'
        }

        It 'uses the inaccessible child as one leaf without losing its siblings' {
            Mock New-ProfileDirectoryEnumerator { throw 'Synthetic child enumeration failure' } -ParameterFilter { $Path -like '*\AppData' }

            $paths = @(Get-ProfileCleanupSubtrees -Path $script:PlanRoot)

            $paths.Count | Should -Be 2
            $paths | Should -Contain "$script:PlanRoot\AppData"
            $paths | Should -Contain "$script:PlanRoot\Documents"
            Get-Content -LiteralPath $Global:CurrentLogFile -Raw | Should -Match 'Synthetic child enumeration failure'
        }

        It 'reads only requested paths from a million-entry branch and disposes both enumerators' {
            $script:RootIterator = [ProfileCleanupEnumeratorProbe]::new("$script:PlanRoot\Branch", 1)
            $script:ChildIterator = [ProfileCleanupEnumeratorProbe]::new("$script:PlanRoot\Branch0\Child", 1000000)
            Mock New-ProfileDirectoryEnumerator {
                if ($Path -eq $script:PlanRoot) { return , $script:RootIterator }
                return , $script:ChildIterator
            }
            $cursor = Get-ProfileCleanupSubtrees -Path $script:PlanRoot -AsEnumerator
            try {
                $script:RootIterator.Reads | Should -Be 0
                $cursor.MoveNext() | Should -BeTrue
                $cursor.Current | Should -Be "$script:PlanRoot\Branch0\Child0"
                $cursor.MoveNext() | Should -BeTrue
                $script:ChildIterator.Reads | Should -Be 2
            }
            finally { $cursor.Dispose() }
            $script:RootIterator.Disposed | Should -BeTrue
            $script:ChildIterator.Disposed | Should -BeTrue
        }

        It 'never yields an overlapping parent when child enumeration fails after a yielded leaf' {
            $script:RootIterator = [ProfileCleanupEnumeratorProbe]::new("$script:PlanRoot\Branch", 1)
            $script:ChildIterator = [ProfileCleanupEnumeratorProbe]::new("$script:PlanRoot\Branch0\Child", 3)
            $script:ChildIterator.FailAfter = 0
            Mock New-ProfileDirectoryEnumerator {
                if ($Path -eq $script:PlanRoot) { return , $script:RootIterator }
                return , $script:ChildIterator
            }

            $paths = @(Get-ProfileCleanupSubtrees -Path $script:PlanRoot)

            $paths.Count | Should -Be 1
            $paths[0] | Should -Be "$script:PlanRoot\Branch0\Child0"
            $paths | Should -Not -Contain "$script:PlanRoot\Branch0"
            $script:RootIterator.Disposed | Should -BeTrue
            $script:ChildIterator.Disposed | Should -BeTrue
            Get-Content -LiteralPath $Global:CurrentLogFile -Raw | Should -Match 'Synthetic partial enumeration failure'
        }
    }

    Context 'Bounded profile scheduling with synthetic workers' {
        BeforeAll {
            $syntheticWorker = @'
$scriptBlock = {
    param($ProfileItem, $Phase = 'Complete', $StartTime = [datetime]::MinValue)
    $ProfileSchedulingEvents.Enqueue("${Phase}:$($ProfileItem.LocalPath)")
    if ($Phase -eq 'Folder') {
        if ($ProfileSchedulingGate) {
            [void]$ProfileSchedulingGate.Signal()
            if (-not $ProfileSchedulingGate.Wait(5000)) { throw 'Synthetic children did not run concurrently' }
        }
        Start-Sleep -Milliseconds 5
        $ProfileSchedulingEvents.Enqueue("Folder-End:$($ProfileItem.LocalPath)")
        if ($ProfileItem.LocalPath.EndsWith('\Fail')) { throw 'Synthetic subtree failure' }
        return [pscustomobject]@{ Type = 'ResidualFolder'; Success = $true }
    }
    Start-Sleep -Milliseconds $ProfileItem.DelayMilliseconds
    if ($ProfileItem.Fail) { throw 'Synthetic bounded-worker failure' }
    if ($Phase -eq 'Prepare' -and $ProfileItem.Remaining) {
        return [pscustomobject]@{ Type = 'ProfilePreparation'; Start = (Get-Date).AddSeconds(-1) }
    }
    [pscustomobject]@{
        Type = 'Profile'
        UserName = $ProfileItem.Name
        Path = $ProfileItem.LocalPath
        Sid = $ProfileItem.SID
        Success = -not $ProfileItem.FinalFail
        Started = $StartTime
        Duration = [TimeSpan]::Zero
    }
}
'@
            $probeSource = $script:ProfileRemovalBatchDefinition.Remove(
                $script:ProfileRemovalWorkerOffset, $script:ProfileRemovalWorkerLength
            ).Insert($script:ProfileRemovalWorkerOffset, $syntheticWorker)
            $probeSource = $probeSource.Replace('$ps = [PowerShell]::Create()', '$ps = New-TrackedProfilePowerShell')
            $probeSource = $probeSource.Replace('$handle = $ps.BeginInvoke()', '$handle = Invoke-ProfileProbeBeginInvoke -PowerShell $ps')
            if ($probeSource -match '\b(Remove-CimInstance|Remove-Item|Invoke-ProfileCleanupCommand)\b') {
                throw 'A destructive command remained in the scheduling probe.'
            }
            . ([scriptblock]::Create($probeSource))

            function New-TrackedProfilePowerShell {
                $outstanding = $script:BatchCreatedPowerShells.Count - $script:BatchCollectedCount + 1
                $script:BatchPeakOutstanding = [Math]::Max($script:BatchPeakOutstanding, $outstanding)
                $powerShell = [PowerShell]::Create()
                $script:BatchCreatedPowerShells.Add($powerShell)
                return $powerShell
            }

            function Invoke-ProfileProbeBeginInvoke {
                param($PowerShell)
                $PowerShell.BeginInvoke()
            }

            function Assert-ProfileProbePowerShellDisposed {
                param($PowerShell)
                $disposed = $false
                try { [void]$PowerShell.AddScript('$null') }
                catch { $disposed = $_.Exception.InnerException -is [System.ObjectDisposedException] }
                $disposed | Should -BeTrue
            }
        }

        BeforeEach {
            $script:BatchCreatedPowerShells = [System.Collections.Generic.List[object]]::new()
            $script:BatchCollectedCount = 0
            $script:BatchPeakOutstanding = 0
            $script:BatchReceivedNames = [System.Collections.Generic.List[string]]::new()
            $script:BatchProgress = [System.Collections.Generic.List[int]]::new()
            $script:BatchEvents = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
            $script:BatchGate = $null
            $script:BatchSubtrees = @{}
            Mock New-ProfileRemovalSessionState {
                $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
                foreach ($entry in @(
                        @{ Name = 'ProfileSchedulingEvents'; Value = $script:BatchEvents },
                        @{ Name = 'ProfileSchedulingGate'; Value = $script:BatchGate }
                    )) {
                    $sessionState.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new($entry.Name, $entry.Value, ''))
                }
                return $sessionState
            }
            Mock Get-ProfileCleanupSubtrees {
                $cursor = [pscustomobject]@{ Paths = @($script:BatchSubtrees[$Path]).GetEnumerator(); Current = $null }
                $cursor | Add-Member -MemberType ScriptMethod -Name MoveNext -Value {
                    if ($this.Paths.MoveNext()) { $this.Current = $this.Paths.Current; return $true }
                    return $false
                }
                $cursor | Add-Member -MemberType ScriptMethod -Name Dispose -Value {}
                return $cursor
            }
            Mock Write-ProgressUpdate { $script:BatchProgress.Add($Percent) }
            Mock Clear-ProgressLine {}
            Mock Receive-ProfileRemovalResult {
                $script:BatchReceivedNames.Add([System.IO.Path]::GetFileName($Job.Profile.LocalPath))
                $result = & $script:ProfileBatchResultReceiver -Job $Job -ResultType $ResultType
                $script:BatchCollectedCount++
                $result
            }
        }

        AfterEach {
            foreach ($powerShell in $script:BatchCreatedPowerShells) { $powerShell.Dispose() }
        }

        It 'completes an empty profile batch without allocating workers or a pool' {
            $maxThreadsEffective = 4

            @(Invoke-ProfileRemovalBatch -Profiles @()).Count | Should -Be 0

            $script:BatchCreatedPowerShells.Count | Should -Be 0
            Should -Invoke New-ProfileRemovalSessionState -Times 0 -Exactly
            Should -Invoke Clear-ProgressLine -Times 1 -Exactly
        }

        It 'runs two children of one profile concurrently and finalizes after both finish' {
            $maxThreadsEffective = 2
            $script:BatchGate = [System.Threading.CountdownEvent]::new(2)
            $profile = [pscustomobject]@{
                Name = 'Split'; LocalPath = 'C:\Users\Split'; SID = 'split'; DelayMilliseconds = 1; Remaining = $true
            }
            $script:BatchSubtrees[$profile.LocalPath] = @('C:\Users\Split\Left', 'C:\Users\Split\Right')
            try {
                $results = @(Invoke-ProfileRemovalBatch -Profiles @($profile))

                $script:BatchGate.CurrentCount | Should -Be 0
                $script:BatchPeakOutstanding | Should -Be 2
                $script:BatchCreatedPowerShells.Count | Should -Be 4
                $results.Count | Should -Be 1
                $results[0].Success | Should -BeTrue
                $results[0].Path | Should -Be $profile.LocalPath
                $results[0].Sid | Should -Be $profile.SID
                $results[0].Started | Should -BeLessThan (Get-Date).AddMilliseconds(-500)
                $events = $script:BatchEvents.ToArray()
                $events[-1] | Should -Be 'Finalize:C:\Users\Split'
                @($events | Where-Object { $_ -like 'Folder-End:*' }).Count | Should -Be 2
                Should -Invoke New-ProfileRemovalSessionState -Times 1 -Exactly
                foreach ($powerShell in $script:BatchCreatedPowerShells) { Assert-ProfileProbePowerShellDisposed -PowerShell $powerShell }
            }
            finally { $script:BatchGate.Dispose() }
        }

        It 'dispatches real worker phases and removes registration only after the empty fixture root is gone' {
            $maxThreadsEffective = 2
            $root = Join-Path $TestDrive 'Phased[Profile]'
            $children = @((Join-Path $root 'Left'), (Join-Path $root 'Right'))
            foreach ($path in $children) { [void][System.IO.Directory]::CreateDirectory($path) }
            $profile = [pscustomobject]@{ LocalPath = $root; SID = 'S-1-5-21-189-777' }
            $script:BatchSubtrees[$root] = $children
            $targets = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($path in @($root) + $children) { [void]$targets.Add($path) }
            Mock New-ProfileRemovalSessionState {
                $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
                $sessionState.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('FixtureTargets', $targets, ''))
                $sessionState.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('FixtureRoot', $root, ''))
                $sessionState.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('FixtureEvents', $script:BatchEvents, ''))
                $stubs = @{
                    'Get-ProfileCleanupPathAttribute' = @'
param($Path)
if (-not $FixtureTargets.Contains($Path) -and $Path -ne [System.IO.Path]::GetPathRoot($FixtureRoot)) {
    throw 'Unexpected fixture inspection target'
}
[System.IO.File]::GetAttributes($Path)
'@
                    'Test-ProfileCleanupDirectory' = (Get-Command Test-ProfileCleanupDirectory).Definition
                    'Assert-ProfileCleanupProfileState' = (Get-Command Assert-ProfileCleanupProfileState).Definition
                    'Get-CimInstance' = @'
[CmdletBinding()] param($ClassName, $Filter)
[pscustomobject]@{ SID = 'S-1-5-21-189-777'; LocalPath = $FixtureRoot; Loaded = $false; Special = $false }
'@
                    'Remove-CimInstance' = @'
[CmdletBinding()] param($InputObject, [switch]$Confirm)
$FixtureEvents.Enqueue("CIM:$($InputObject.LocalPath)")
throw 'Synthetic CIM leftovers'
'@
                    'Remove-Item' = @'
[CmdletBinding()] param([string]$LiteralPath, [switch]$Recurse, [switch]$Force, [switch]$Confirm)
if (-not $FixtureTargets.Contains($LiteralPath)) { throw 'Unexpected fixture deletion target' }
# The non-recursive API can delete only the explicitly allowed, empty fixture directories.
[System.IO.Directory]::Delete($LiteralPath)
$FixtureEvents.Enqueue("Removed:$LiteralPath")
'@
                    'Remove-ProfileRegistryEntries' = @'
param($Sid, $UserName)
if ([System.IO.Directory]::Exists($FixtureRoot)) { throw 'Registry cleanup preceded parent removal' }
$FixtureEvents.Enqueue("Registry:$Sid")
$true
'@
                    'Invoke-ProfileCleanupCommand' = 'param($Command, $Arguments, $LogContextKey)'
                    'Write-ToolkitLog' = 'param($Level, $Message)'
                    'Get-SourceTextLoc' = 'param($Key, $Args) $Key'
                    'Test-Path' = 'param($Path) $true'
                    'Get-Item' = @'
[CmdletBinding()] param($LiteralPath, [switch]$Force)
if (-not $FixtureTargets.Contains($LiteralPath)) { throw 'Unexpected fixture inspection target' }
if (-not [System.IO.Directory]::Exists($LiteralPath)) {
    throw [System.Management.Automation.ItemNotFoundException]::new('Owned fixture directory is gone')
}
[pscustomobject]@{ PSIsContainer = $true; Attributes = [System.IO.FileAttributes]::Directory }
'@
                }
                foreach ($name in $stubs.Keys) {
                    $sessionState.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($name, $stubs[$name]))
                }
                return $sessionState
            }

            $results = @(& $script:ProfileRemovalBatchInvoker -Profiles @($profile))

            $results.Count | Should -Be 1
            $results[0].Success | Should -BeTrue
            $results[0].Path | Should -Be $root
            $results[0].Sid | Should -Be 'S-1-5-21-189-777'
            [System.IO.Directory]::Exists($root) | Should -BeFalse
            $events = $script:BatchEvents.ToArray()
            @($events | Where-Object { $_ -like 'CIM:*' }).Count | Should -Be 1
            @($events | Where-Object { $_ -like 'Removed:*' }).Count | Should -Be 3
            $events[-2] | Should -Be "Removed:$root"
            $events[-1] | Should -Be 'Registry:S-1-5-21-189-777'
        }

        It 'uses a global cap of <Threads> across odd child counts and combines profiles in input order' -ForEach @(
            @{ Threads = 1 }, @{ Threads = 4 }
        ) {
            $maxThreadsEffective = $Threads
            $profiles = @(for ($index = 0; $index -lt 3; $index++) {
                $profile = [pscustomobject]@{
                    Name = "Split$index"; LocalPath = "C:\Users\Split$index"; SID = "split$index"; DelayMilliseconds = 1; Remaining = $true
                }
                $script:BatchSubtrees[$profile.LocalPath] = @(for ($child = 0; $child -lt 5; $child++) { "$($profile.LocalPath)\Child$child" })
                $profile
            })

            $results = @(Invoke-ProfileRemovalBatch -Profiles $profiles)

            $script:BatchPeakOutstanding | Should -BeLessOrEqual $Threads
            $script:BatchCreatedPowerShells.Count | Should -Be 21
            $script:BatchCollectedCount | Should -Be 21
            $results.Count | Should -Be 3
            for ($index = 0; $index -lt 3; $index++) {
                $results[$index].Path | Should -Be $profiles[$index].LocalPath
                $results[$index].Sid | Should -Be $profiles[$index].SID
            }
            $events = $script:BatchEvents.ToArray()
            @($events | Where-Object { $_ -like 'Folder:*' } | Sort-Object -Unique).Count | Should -Be 15
            foreach ($profile in $profiles) {
                $finalIndex = [array]::IndexOf($events, "Finalize:$($profile.LocalPath)")
                foreach ($path in $script:BatchSubtrees[$profile.LocalPath]) {
                    [array]::IndexOf($events, "Folder-End:$path") | Should -BeLessThan $finalIndex
                }
            }
            Should -Invoke New-ProfileRemovalSessionState -Times 1 -Exactly
        }

        It 'attempts root recovery after a failed child and counts only the final profile result' {
            $maxThreadsEffective = 2
            $profile = [pscustomobject]@{
                Name = 'Recovery'; LocalPath = 'C:\Users\Recovery'; SID = 'recovery'; DelayMilliseconds = 1; Remaining = $true
            }
            $script:BatchSubtrees[$profile.LocalPath] = @('C:\Users\Recovery\Fail', 'C:\Users\Recovery\Next')

            $results = @(Invoke-ProfileRemovalBatch -Profiles @($profile))

            $results.Count | Should -Be 1
            $results[0].Success | Should -BeTrue
            $script:BatchCollectedCount | Should -Be 4
            $script:BatchEvents.ToArray()[-1] | Should -Be 'Finalize:C:\Users\Recovery'
            Get-Content -LiteralPath $Global:CurrentLogFile -Raw | Should -Match '\[ERROR\].*Synthetic subtree failure'
            $script:BatchProgress[0] | Should -Be 0
            $script:BatchProgress[-1] | Should -Be 100
        }

        It 'disposes all workers if a child fails to start after profile preparation succeeded' {
            $maxThreadsEffective = 2
            $profile = [pscustomobject]@{
                Name = 'ChildStartup'; LocalPath = 'C:\Users\ChildStartup'; SID = 'childStartup'; DelayMilliseconds = 1; Remaining = $true
            }
            $script:BatchSubtrees[$profile.LocalPath] = @('C:\Users\ChildStartup\Left', 'C:\Users\ChildStartup\Right')
            Mock Invoke-ProfileProbeBeginInvoke {
                if ($script:BatchCreatedPowerShells.Count -eq 3) { throw 'Synthetic child startup failure' }
                $PowerShell.BeginInvoke()
            }

            { Invoke-ProfileRemovalBatch -Profiles @($profile) } | Should -Throw '*Synthetic child startup failure*'

            $script:BatchCreatedPowerShells.Count | Should -Be 3
            $script:BatchCollectedCount | Should -Be 1
            foreach ($powerShell in $script:BatchCreatedPowerShells) { Assert-ProfileProbePowerShellDisposed -PowerShell $powerShell }
        }

        It 'uses whole-root finalization without child pipelines for <Count> planned subtrees' -ForEach @(
            @{ Count = 0 }, @{ Count = 1 }
        ) {
            $maxThreadsEffective = 2
            $profile = [pscustomobject]@{
                Name = 'Small'; LocalPath = 'C:\Users\Small'; SID = 'small'; DelayMilliseconds = 1; Remaining = $true; FinalFail = $true
            }
            $script:BatchSubtrees[$profile.LocalPath] = @(for ($child = 0; $child -lt $Count; $child++) { "C:\Users\Small\Child$child" })

            $results = @(Invoke-ProfileRemovalBatch -Profiles @($profile))

            $script:BatchCreatedPowerShells.Count | Should -Be 2
            $results.Count | Should -Be 1
            $results[0].Success | Should -BeFalse
            @($script:BatchEvents.ToArray() | Where-Object { $_ -like 'Folder:*' }).Count | Should -Be 0
        }

        It 'limits <Count> profiles to <Threads> outstanding jobs and preserves every result' -ForEach @(
            @{ Count = 3; Threads = 1 },
            @{ Count = 32; Threads = 2 },
            @{ Count = 256; Threads = 4 }
        ) {
            $maxThreadsEffective = $Threads
            $profiles = @(for ($index = 0; $index -lt $Count; $index++) {
                [pscustomobject]@{
                    Name = "Synthetic$index"
                    LocalPath = "C:\Users\Synthetic$index"
                    SID = "S-1-5-21-189-$index"
                    DelayMilliseconds = 1
                    Fail = $false
                }
            })

            $results = @(Invoke-ProfileRemovalBatch -Profiles $profiles)

            $script:BatchPeakOutstanding | Should -BeLessOrEqual $Threads
            $script:BatchCreatedPowerShells.Count | Should -Be $Count
            $script:BatchCollectedCount | Should -Be $Count
            $results.Count | Should -Be $Count
            @($results | Where-Object { -not $_.Success }).Count | Should -Be 0
            for ($index = 0; $index -lt $Count; $index++) {
                $results[$index].UserName | Should -Be $profiles[$index].Name
            }
            $script:BatchProgress[-1] | Should -Be 100
            foreach ($powerShell in $script:BatchCreatedPowerShells) {
                Assert-ProfileProbePowerShellDisposed -PowerShell $powerShell
            }
        }

        It 'returns input order when a later worker finishes first' {
            $maxThreadsEffective = 2
            $profiles = @(
                [pscustomobject]@{ Name = 'Slow'; LocalPath = 'C:\Users\Slow'; SID = 'slow'; DelayMilliseconds = 250; Fail = $false },
                [pscustomobject]@{ Name = 'Fast'; LocalPath = 'C:\Users\Fast'; SID = 'fast'; DelayMilliseconds = 1; Fail = $false }
            )

            $results = @(Invoke-ProfileRemovalBatch -Profiles $profiles)

            $script:BatchReceivedNames[0] | Should -Be 'Fast'
            $results[0].UserName | Should -Be 'Slow'
            $results[1].UserName | Should -Be 'Fast'
        }

        It 'retains one failed result while scheduling replacements' {
            $maxThreadsEffective = 1
            $profiles = @(
                [pscustomobject]@{ Name = 'Failed'; LocalPath = 'C:\Users\Failed'; SID = 'failed'; DelayMilliseconds = 1; Fail = $true },
                [pscustomobject]@{ Name = 'Next'; LocalPath = 'C:\Users\Next'; SID = 'next'; DelayMilliseconds = 1; Fail = $false }
            )

            $results = @(Invoke-ProfileRemovalBatch -Profiles $profiles)

            $script:BatchPeakOutstanding | Should -Be 1
            $results.Count | Should -Be 2
            $results[0].UserName | Should -Be 'Failed'
            $results[0].Success | Should -BeFalse
            $results[1].UserName | Should -Be 'Next'
            $results[1].Success | Should -BeTrue
            $script:BatchCollectedCount | Should -Be 2
        }

        It 'disposes every outstanding instance when progress reporting fails' {
            $maxThreadsEffective = 2
            $profiles = @(
                [pscustomobject]@{ Name = 'First'; LocalPath = 'C:\Users\First'; SID = 'first'; DelayMilliseconds = 1000; Fail = $false },
                [pscustomobject]@{ Name = 'Second'; LocalPath = 'C:\Users\Second'; SID = 'second'; DelayMilliseconds = 1000; Fail = $false }
            )
            Mock Write-ProgressUpdate { throw 'Synthetic progress failure' }

            { Invoke-ProfileRemovalBatch -Profiles $profiles } | Should -Throw '*Synthetic progress failure*'

            $script:BatchCreatedPowerShells.Count | Should -Be 2
            foreach ($powerShell in $script:BatchCreatedPowerShells) {
                Assert-ProfileProbePowerShellDisposed -PowerShell $powerShell
            }
        }

        It 'exposes no native command timeout parameter' {
            foreach ($command in @('Invoke-ProfileCleanupCommand', 'Remove-ResidualUserFolder')) {
                (Get-Command $command).Parameters.ContainsKey('TimeoutSeconds') | Should -BeFalse
            }
            $sourcePath = if ($CompiledScriptPath) { $CompiledScriptPath } else { Join-Path $script:RepoRoot 'tools\WinDeleteUserProfiles.ps1' }
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($sourcePath, [ref]$null, [ref]$null)
            $tool = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'WinDeleteUserProfiles' }, $true)
            $tool.Body.ParamBlock.Parameters.Name.VariablePath.UserPath | Should -Not -Contain 'ExternalCommandTimeoutSeconds'
        }

        It 'disposes an unfinished subtree cursor when collection fails' {
            $maxThreadsEffective = 2
            $profile = [pscustomobject]@{ Name = 'Cursor'; LocalPath = 'C:\Users\Cursor'; SID = 'cursor'; DelayMilliseconds = 1; Remaining = $true }
            $script:AbandonedCursor = [pscustomobject]@{ Index = -1; Current = $null; Disposed = $false }
            $script:AbandonedCursor | Add-Member -MemberType ScriptMethod -Name MoveNext -Value {
                $this.Index++
                $this.Current = "C:\Users\Cursor\Child$($this.Index)"
                return $this.Index -lt 1000000
            }
            $script:AbandonedCursor | Add-Member -MemberType ScriptMethod -Name Dispose -Value { $this.Disposed = $true }
            Mock Get-ProfileCleanupSubtrees { $script:AbandonedCursor }
            Mock Receive-ProfileRemovalResult {
                $result = & $script:ProfileBatchResultReceiver -Job $Job -ResultType $ResultType
                $script:BatchCollectedCount++
                if ($Job.Phase -eq 'Folder') { throw 'Synthetic cursor collection failure' }
                return $result
            }

            { Invoke-ProfileRemovalBatch -Profiles @($profile) } | Should -Throw '*Synthetic cursor collection failure*'

            $script:AbandonedCursor.Disposed | Should -BeTrue
            $script:AbandonedCursor.Index | Should -BeLessThan 10
            foreach ($powerShell in $script:BatchCreatedPowerShells) { Assert-ProfileProbePowerShellDisposed -PowerShell $powerShell }
        }

        It 'disposes an instance when startup fails before job registration' {
            $maxThreadsEffective = 1
            $profiles = @([pscustomobject]@{
                Name = 'Startup'; LocalPath = 'C:\Users\Startup'; SID = 'startup'; DelayMilliseconds = 1; Fail = $false
            })
            Mock Invoke-ProfileProbeBeginInvoke { throw 'Synthetic startup failure' }

            { Invoke-ProfileRemovalBatch -Profiles $profiles } | Should -Throw '*Synthetic startup failure*'

            $script:BatchCreatedPowerShells.Count | Should -Be 1
            Assert-ProfileProbePowerShellDisposed -PowerShell $script:BatchCreatedPowerShells[0]
        }

        It 'disposes remaining instances when collection fails after one instance was disposed' {
            $maxThreadsEffective = 2
            $profiles = @(
                [pscustomobject]@{ Name = 'Collected'; LocalPath = 'C:\Users\Collected'; SID = 'collected'; DelayMilliseconds = 1; Fail = $false },
                [pscustomobject]@{ Name = 'Remaining'; LocalPath = 'C:\Users\Remaining'; SID = 'remaining'; DelayMilliseconds = 1000; Fail = $false }
            )
            Mock Receive-ProfileRemovalResult {
                & $script:ProfileBatchResultReceiver -Job $Job | Out-Null
                $script:BatchCollectedCount++
                throw 'Synthetic collection failure'
            }

            { Invoke-ProfileRemovalBatch -Profiles $profiles } | Should -Throw '*Synthetic collection failure*'

            $script:BatchCollectedCount | Should -Be 1
            foreach ($powerShell in $script:BatchCreatedPowerShells) {
                Assert-ProfileProbePowerShellDisposed -PowerShell $powerShell
            }
        }

        It 'still disposes an instance when stopping it fails' {
            $powerShell = [pscustomobject]@{ Disposed = $false }
            $powerShell | Add-Member -MemberType ScriptMethod -Name Stop -Value { throw 'Synthetic stop failure' }
            $powerShell | Add-Member -MemberType ScriptMethod -Name Dispose -Value { $this.Disposed = $true }
            Mock Write-Warning {}

            { Close-ProfileRemovalPowerShell -PowerShell $powerShell } | Should -Not -Throw

            $powerShell.Disposed | Should -BeTrue
            Should -Invoke Write-Warning -Times 1 -Exactly -ParameterFilter {
                $Message -like '*Synthetic stop failure*' -and $WarningAction -eq 'Continue'
            }
        }

        It 'continues cleanup after a disposal failure and reports the failure' {
            $failedPowerShell = [pscustomobject]@{}
            $failedPowerShell | Add-Member -MemberType ScriptMethod -Name Stop -Value {}
            $failedPowerShell | Add-Member -MemberType ScriptMethod -Name Dispose -Value { throw 'Synthetic dispose failure' }
            $nextPowerShell = [PowerShell]::Create()
            $script:BatchCreatedPowerShells.Add($nextPowerShell)
            Mock Write-Warning {}

            {
                foreach ($powerShell in @($failedPowerShell, $nextPowerShell)) {
                    Close-ProfileRemovalPowerShell -PowerShell $powerShell
                }
            } | Should -Not -Throw

            Assert-ProfileProbePowerShellDisposed -PowerShell $nextPowerShell
            Should -Invoke Write-Warning -Times 1 -Exactly -ParameterFilter {
                $Message -like '*Synthetic dispose failure*' -and $WarningAction -eq 'Continue'
            }
        }
    }

    Context 'Registry cleanup leaves with mocked operations' {
        BeforeEach {
            Mock Test-Path { $true }
            Mock Remove-Item {}
            Mock Write-ToolkitLog {}
        }

        It 'removes only the selected SID key' {
            Remove-ProfileRegistryEntries -Sid 'S-1-5-21-189-1' -UserName 'Probe' | Should -BeTrue

            Should -Invoke Remove-Item -Times 1 -Exactly -ParameterFilter {
                $LiteralPath -eq 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\S-1-5-21-189-1' -and $Recurse -and $Force -and $ErrorAction -eq 'Stop'
            }
        }

        It 'treats an already absent key as successful' {
            Mock Test-Path { $false }

            Remove-ProfileRegistryEntries -Sid 'S-1-5-21-189-1' -UserName 'Probe' | Should -BeTrue

            Should -Invoke Remove-Item -Times 0 -Exactly
        }

        It 'retains registry failures in the profile outcome' {
            Mock Remove-Item { throw 'Synthetic registry denial' }

            Remove-ProfileRegistryEntries -Sid 'S-1-5-21-189-1' -UserName 'Probe' | Should -BeFalse

            Should -Invoke Write-ToolkitLog -Times 1 -Exactly -ParameterFilter { $Level -eq 'WARNING' }
        }

        It 'rejects the placeholder SID without accessing the registry' {
            Remove-ProfileRegistryEntries -Sid 'NULL' -UserName 'Probe' | Should -BeFalse

            Should -Invoke Test-Path -Times 0 -Exactly
            Should -Invoke Remove-Item -Times 0 -Exactly
        }
    }

    Context 'Residual discovery range partitioning without deletion' {
        BeforeEach {
            $usersRoot = 'C:\Users\Scan[Root]\'
            $script:ScanRoot = $usersRoot
            $script:ScanFolders = @()
            $script:ScanSids = @()
            $script:ScanExcluded = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $script:ScanRegistered = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            Mock New-ProtectedNameSet { return , $script:ScanExcluded }
            Mock Get-RegisteredProfilePathSet { return , $script:ScanRegistered }
            Mock Get-ChildItem {
                if ($LiteralPath -eq $script:ScanRoot) { foreach ($folder in $script:ScanFolders) { $folder } }
                elseif ($LiteralPath -eq 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList') {
                    foreach ($sid in $script:ScanSids) { [pscustomobject]@{ PSChildName = $sid; PSPath = "HKLM:\ProfileReviewMock\$sid" } }
                }
                else { throw 'Unexpected discovery target' }
            }
            Mock Write-StyledMessage {}
            Mock Write-ToolkitLog {}
            Mock Get-ItemProperty { [pscustomobject]@{ ProfileImagePath = 'C:\OutsideScanRoot\Profile' } }
            Mock New-ProfileRemovalSessionState { throw 'Discovery must not allocate a worker pool' }
        }

        It 'stops residual discovery when CIM detection fails' {
            Mock Get-RegisteredProfilePathSet { & $script:RegisteredProfilePathSetReader }
            Mock Get-CimInstance { throw 'Synthetic CIM detection failure' }

            { Get-ResidualUserFolders } | Should -Throw '*Synthetic CIM detection failure*'

            Should -Invoke Get-ChildItem -Times 0 -Exactly
        }

        It 'protects ordinary username folders mapped only by registry ProfileImagePath' {
            $script:ScanFolders = @([pscustomobject]@{
                Name = 'Alice'; FullName = $script:ScanRoot + 'Alice'; Attributes = [System.IO.FileAttributes]::Directory
            })
            $script:ScanSids = @('S-1-5-21-189-1001')
            Mock Get-ItemProperty { [pscustomobject]@{ ProfileImagePath = $script:ScanRoot.ToUpperInvariant() + 'ALICE\' } }

            @(Get-ResidualUserFolders).Count | Should -Be 0

            Should -Invoke Get-ItemProperty -Times 1 -Exactly -ParameterFilter { $Name -eq 'ProfileImagePath' -and $ErrorAction -eq 'Stop' }
        }

        It 'stops discovery on <Failure> instead of accepting incomplete registry protection' -ForEach @(
            @{ Failure = 'Enumeration' }, @{ Failure = 'PathRead' }, @{ Failure = 'EmptyPath' }
        ) {
            $script:ScanSids = @('S-1-5-21-189-1001')
            $script:ScanFolders = @([pscustomobject]@{
                Name = 'Alice'; FullName = $script:ScanRoot + 'Alice'; Attributes = [System.IO.FileAttributes]::Directory
            })
            if ($Failure -eq 'Enumeration') {
                Mock Get-ChildItem { throw 'Synthetic registry enumeration failure' } -ParameterFilter { $LiteralPath -like 'HKLM:*' }
            }
            elseif ($Failure -eq 'PathRead') {
                Mock Get-ItemProperty { throw 'Synthetic registry path read failure' }
            }
            else { Mock Get-ItemProperty { [pscustomobject]@{ ProfileImagePath = '' } } }

            { Get-ResidualUserFolders } | Should -Throw

            Should -Invoke Get-ChildItem -Times 1 -Exactly -ParameterFilter { $LiteralPath -like 'HKLM:*' -and $ErrorAction -eq 'Stop' }
        }

        It 'covers every leaf of <Count> folders and preserves discovery order' -ForEach @(
            @{ Count = 0 }, @{ Count = 1 }, @{ Count = 255 }, @{ Count = 256 },
            @{ Count = 257 }, @{ Count = 513 }, @{ Count = 4097 }
        ) {
            $script:ScanFolders = @(for ($index = 0; $index -lt $Count; $index++) {
                [pscustomobject]@{
                    Name = "Candidate[$index]"; FullName = "$($script:ScanRoot)Candidate[$index]"
                    Attributes = [System.IO.FileAttributes]::Directory
                }
            })

            $results = @(Get-ResidualUserFolders)

            $results.Count | Should -Be $Count
            ($results.Name -join "`n") | Should -Be ($script:ScanFolders.Name -join "`n")
            ($results.Path -join "`n") | Should -Be ($script:ScanFolders.FullName -join "`n")
            Should -Invoke New-ProfileRemovalSessionState -Times 0 -Exactly
            Should -Invoke Get-RegisteredProfilePathSet -Times 1 -Exactly
            Should -Invoke Get-ChildItem -Times 2 -Exactly
        }

        It 'preserves protected names, CIM paths, SID names, and link exclusions across partitions' {
            $script:ScanFolders = @(for ($index = 0; $index -lt 513; $index++) {
                [pscustomobject]@{
                    Name = "Candidate[$index]"; FullName = "$($script:ScanRoot)Candidate[$index]"
                    Attributes = [System.IO.FileAttributes]::Directory
                }
            })
            [void]$script:ScanExcluded.Add('CANDIDATE[0]')
            [void]$script:ScanRegistered.Add($script:ScanFolders[256].FullName.ToUpperInvariant())
            $script:ScanSids = @('CANDIDATE[257]', 'CANDIDATE[257]')
            $script:ScanFolders[512].Attributes = [System.IO.FileAttributes]::ReparsePoint

            $results = @(Get-ResidualUserFolders)

            $results.Count | Should -Be 509
            $results.Name | Should -Not -Contain 'Candidate[0]'
            $results.Name | Should -Not -Contain 'Candidate[256]'
            $results.Name | Should -Not -Contain 'Candidate[257]'
            $results.Name | Should -Not -Contain 'Candidate[512]'
            $results[0].Name | Should -Be 'Candidate[1]'
            $results[-1].Name | Should -Be 'Candidate[511]'
            Should -Invoke Write-ToolkitLog -Times 4 -Exactly -ParameterFilter { $Level -eq 'INFO' }
            Should -Invoke Get-ChildItem -Times 1 -Exactly -ParameterFilter {
                $LiteralPath -eq $script:ScanRoot -and $Directory -and $Force
            }
        }
    }

    Context 'Divide-and-conquer residual scheduling with synthetic workers' {
        BeforeAll {
            $syntheticWorker = @'
$scriptBlock = {
    param($Folder)
    Clear-ProgressLine
    Write-ProgressUpdate -Activity 'Synthetic worker' -Percent 50
    if ($Folder.Gate) {
        [void]$Folder.Gate.Signal()
        if (-not $Folder.Gate.Wait(5000)) { throw 'Synthetic workers did not run concurrently' }
    }
    Start-Sleep -Milliseconds $Folder.DelayMilliseconds
    if ($Folder.Fail) { throw 'Synthetic residual-worker failure' }
    [pscustomobject]@{
        Type = 'ResidualFolder'
        UserName = $Folder.Name
        Path = $Folder.Path
        Success = $true
        Duration = [TimeSpan]::Zero
    }
}
'@
            $probeSource = $script:ResidualRemovalBatchDefinition.Remove(
                $script:ResidualRemovalWorkerOffset, $script:ResidualRemovalWorkerLength
            ).Insert($script:ResidualRemovalWorkerOffset, $syntheticWorker)
            $probeSource = $probeSource.Replace('$ps = [PowerShell]::Create()', '$ps = New-TrackedResidualPowerShell')
            $probeSource = $probeSource.Replace('$handle = $ps.BeginInvoke()', '$handle = Invoke-ResidualProbeBeginInvoke -PowerShell $ps')
            if ($probeSource -match '\b(Remove-CimInstance|Remove-Item|Invoke-ProfileCleanupCommand|Remove-ResidualUserFolder)\b') {
                throw 'A destructive command remained in the residual scheduling probe.'
            }
            . ([scriptblock]::Create($probeSource))

            function New-TrackedResidualPowerShell {
                $outstanding = $script:ResidualCreatedPowerShells.Count - $script:ResidualCollectedCount + 1
                $script:ResidualPeakOutstanding = [Math]::Max($script:ResidualPeakOutstanding, $outstanding)
                $powerShell = [PowerShell]::Create()
                $script:ResidualCreatedPowerShells.Add($powerShell)
                return $powerShell
            }

            function Invoke-ResidualProbeBeginInvoke {
                param($PowerShell)
                $PowerShell.BeginInvoke()
            }

            function Assert-ResidualProbePowerShellDisposed {
                param($PowerShell)
                $disposed = $false
                try { [void]$PowerShell.AddScript('$null') }
                catch { $disposed = $_.Exception.InnerException -is [System.ObjectDisposedException] }
                $disposed | Should -BeTrue
            }
        }

        BeforeEach {
            $script:ResidualCreatedPowerShells = [System.Collections.Generic.List[object]]::new()
            $script:ResidualCollectedCount = 0
            $script:ResidualPeakOutstanding = 0
            $script:ResidualReceivedNames = [System.Collections.Generic.List[string]]::new()
            $script:ResidualProgress = [System.Collections.Generic.List[object]]::new()
            Mock New-ProfileRemovalSessionState { & $script:ResidualSessionStateFactory }
            Mock Write-ProgressUpdate {
                $script:ResidualProgress.Add([pscustomobject]@{ Percent = $Percent; Completed = $script:ResidualCollectedCount })
            }
            Mock Clear-ProgressLine {}
            Mock Receive-ProfileRemovalResult {
                $script:ResidualReceivedNames.Add($Job.Folder.Name)
                $result = & $script:ProfileBatchResultReceiver -Job $Job -ResultType $ResultType
                $script:ResidualCollectedCount++
                $result
            }
        }

        AfterEach {
            foreach ($powerShell in $script:ResidualCreatedPowerShells) { $powerShell.Dispose() }
        }

        It 'covers every leaf of <Count> folders with at most <Threads> live pipelines' -ForEach @(
            @{ Count = 1; Threads = 1 },
            @{ Count = 3; Threads = 2 },
            @{ Count = 17; Threads = 4 },
            @{ Count = 256; Threads = 4 }
        ) {
            $maxThreadsEffective = $Threads
            $folders = @(for ($index = 0; $index -lt $Count; $index++) {
                [pscustomobject]@{
                    Name = "Residual[$index]"
                    Path = "C:\Users\Residual[$index]"
                    DelayMilliseconds = 1
                    Fail = $false
                    Gate = $null
                }
            })

            $results = @(Remove-ResidualUserFolders -Folders $folders)

            $script:ResidualPeakOutstanding | Should -BeLessOrEqual $Threads
            $script:ResidualCreatedPowerShells.Count | Should -Be $Count
            $script:ResidualCollectedCount | Should -Be $Count
            $results.Count | Should -Be $Count
            @($script:ResidualReceivedNames | Sort-Object -Unique).Count | Should -Be $Count
            for ($index = 0; $index -lt $Count; $index++) {
                $results[$index].Type | Should -Be 'ResidualFolder'
                $results[$index].UserName | Should -Be $folders[$index].Name
                $results[$index].Path | Should -Be $folders[$index].Path
                $results[$index].Success | Should -BeTrue
            }
            $script:ResidualProgress[0].Percent | Should -Be 0
            $script:ResidualProgress[-1].Percent | Should -Be 100
            foreach ($progress in $script:ResidualProgress) {
                $progress.Percent | Should -Be ([math]::Floor(($progress.Completed / $Count) * 100))
            }
            foreach ($powerShell in $script:ResidualCreatedPowerShells) {
                Assert-ResidualProbePowerShellDisposed -PowerShell $powerShell
            }
        }

        It 'completes an empty input without creating a pool or a worker' {
            $maxThreadsEffective = 4

            $results = @(Remove-ResidualUserFolders -Folders @())

            $results.Count | Should -Be 0
            $script:ResidualCreatedPowerShells.Count | Should -Be 0
            $script:ResidualProgress[-1].Percent | Should -Be 100
            Should -Invoke New-ProfileRemovalSessionState -Times 0 -Exactly
            Should -Invoke Clear-ProgressLine -Times 1 -Exactly
        }

        It 'runs independent leaves concurrently in the same pool' {
            $maxThreadsEffective = 2
            $gate = [System.Threading.CountdownEvent]::new(2)
            try {
                $folders = @(for ($index = 0; $index -lt 2; $index++) {
                    [pscustomobject]@{
                        Name = "Concurrent$index"; Path = "C:\Users\Concurrent$index"
                        DelayMilliseconds = 1; Fail = $false; Gate = $gate
                    }
                })

                $results = @(Remove-ResidualUserFolders -Folders $folders)

                $results.Count | Should -Be 2
                @($results | Where-Object { -not $_.Success }).Count | Should -Be 0
                $gate.CurrentCount | Should -Be 0
            }
            finally { $gate.Dispose() }
        }

        It 'passes the selected folder to the real worker script without running deletion commands' {
            $maxThreadsEffective = 1
            Mock New-ProfileRemovalSessionState {
                $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
                $leafStub = @'
param($Folder)
[pscustomobject]@{
    Type = 'ResidualFolder'
    UserName = $Folder.Name
    Path = $Folder.Path
    Success = $true
    Duration = [TimeSpan]::Zero
}
'@
                $sessionState.Commands.Add(
                    [System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Remove-ResidualUserFolder', $leafStub)
                )
                return $sessionState
            }
            $folder = [pscustomobject]@{ Name = 'Bridge[Leaf]'; Path = 'C:\Users\Bridge[Leaf]' }

            $results = @(& $script:ResidualRemovalBatchInvoker -Folders @($folder))

            $results.Count | Should -Be 1
            $results[0].Type | Should -Be 'ResidualFolder'
            $results[0].UserName | Should -Be $folder.Name
            $results[0].Path | Should -Be $folder.Path
            $results[0].Success | Should -BeTrue
        }

        It 'starts another leaf before a slow leaf finishes and combines results in input order' {
            $maxThreadsEffective = 2
            $folders = @(
                [pscustomobject]@{ Name = 'Slow'; Path = 'C:\Users\Slow'; DelayMilliseconds = 700; Fail = $false; Gate = $null },
                [pscustomobject]@{ Name = 'Fast'; Path = 'C:\Users\Fast'; DelayMilliseconds = 1; Fail = $false; Gate = $null },
                [pscustomobject]@{ Name = 'Next'; Path = 'C:\Users\Next'; DelayMilliseconds = 1; Fail = $false; Gate = $null }
            )

            $results = @(Remove-ResidualUserFolders -Folders $folders)

            $script:ResidualReceivedNames[0] | Should -Be 'Fast'
            $script:ResidualReceivedNames[1] | Should -Be 'Next'
            $results[0].UserName | Should -Be 'Slow'
            $results[1].UserName | Should -Be 'Fast'
            $results[2].UserName | Should -Be 'Next'
        }

        It 'retains a failed residual leaf without losing its identity or the next leaf' {
            $maxThreadsEffective = 1
            $folders = @(
                [pscustomobject]@{ Name = 'Failed[Leaf]'; Path = 'C:\Users\Failed[Leaf]'; DelayMilliseconds = 1; Fail = $true; Gate = $null },
                [pscustomobject]@{ Name = 'Next'; Path = 'C:\Users\Next'; DelayMilliseconds = 1; Fail = $false; Gate = $null }
            )

            $results = @(Remove-ResidualUserFolders -Folders $folders)

            $results.Count | Should -Be 2
            $results[0].Type | Should -Be 'ResidualFolder'
            $results[0].UserName | Should -Be $folders[0].Name
            $results[0].Path | Should -Be $folders[0].Path
            $results[0].Success | Should -BeFalse
            $results[1].Success | Should -BeTrue
            $script:ResidualCollectedCount | Should -Be 2
            Get-Content -LiteralPath $Global:CurrentLogFile -Raw | Should -Match '\[ERROR\].*Synthetic residual-worker failure'
        }

        It 'disposes every residual instance when coordinator progress fails' {
            $maxThreadsEffective = 2
            $folders = @(for ($index = 0; $index -lt 2; $index++) {
                [pscustomobject]@{
                    Name = "Progress$index"; Path = "C:\Users\Progress$index"
                    DelayMilliseconds = 1000; Fail = $false; Gate = $null
                }
            })
            Mock Write-ProgressUpdate { throw 'Synthetic residual progress failure' }

            { Remove-ResidualUserFolders -Folders $folders } | Should -Throw '*Synthetic residual progress failure*'

            $script:ResidualCreatedPowerShells.Count | Should -Be 2
            foreach ($powerShell in $script:ResidualCreatedPowerShells) {
                Assert-ResidualProbePowerShellDisposed -PowerShell $powerShell
            }
        }

        It 'disposes a residual instance when startup fails before registration' {
            $maxThreadsEffective = 1
            $folders = @([pscustomobject]@{
                Name = 'Startup'; Path = 'C:\Users\Startup'; DelayMilliseconds = 1; Fail = $false; Gate = $null
            })
            Mock Invoke-ResidualProbeBeginInvoke { throw 'Synthetic residual startup failure' }

            { Remove-ResidualUserFolders -Folders $folders } | Should -Throw '*Synthetic residual startup failure*'

            $script:ResidualCreatedPowerShells.Count | Should -Be 1
            Assert-ResidualProbePowerShellDisposed -PowerShell $script:ResidualCreatedPowerShells[0]
        }

        It 'cleans up remaining residual instances after collection throws' {
            $maxThreadsEffective = 2
            $folders = @(
                [pscustomobject]@{ Name = 'Collected'; Path = 'C:\Users\Collected'; DelayMilliseconds = 1; Fail = $false; Gate = $null },
                [pscustomobject]@{ Name = 'Remaining'; Path = 'C:\Users\Remaining'; DelayMilliseconds = 1000; Fail = $false; Gate = $null }
            )
            Mock Receive-ProfileRemovalResult {
                & $script:ProfileBatchResultReceiver -Job $Job -ResultType $ResultType | Out-Null
                $script:ResidualCollectedCount++
                throw 'Synthetic residual collection failure'
            }

            { Remove-ResidualUserFolders -Folders $folders } | Should -Throw '*Synthetic residual collection failure*'

            $script:ResidualCollectedCount | Should -Be 1
            foreach ($powerShell in $script:ResidualCreatedPowerShells) {
                Assert-ResidualProbePowerShellDisposed -PowerShell $powerShell
            }
        }
    }
}
