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
            @{ Path = $toolPath; Name = 'Remove-ProfileRegistryEntries' },
            @{ Path = $toolPath; Name = 'Remove-ResidualUserFolder' },
            @{ Path = $toolPath; Name = 'Remove-ResidualUserFolders' },
            @{ Path = $localizationPath; Name = 'Get-SourceTextLoc' },
            @{ Path = $loggingPath; Name = 'Write-ToolkitLog' },
            @{ Path = $processesPath; Name = 'Invoke-ExternalCommandWithLog' },
            @{ Path = $processesPath; Name = 'Remove-ItemSafely' },
            @{ Path = $uiPath; Name = 'Get-SpinnerChar' },
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
        $script:ProfileRemovalWorkerOffset = $workerAssignment.Extent.StartOffset - $batchFunction.Extent.StartOffset
        $script:ProfileRemovalWorkerLength = $workerAssignment.Extent.EndOffset - $workerAssignment.Extent.StartOffset
        $script:ProfileBatchResultReceiver = (Get-Command Receive-ProfileRemovalResult).ScriptBlock
        $script:ResidualSessionStateFactory = (Get-Command New-ProfileRemovalSessionState).ScriptBlock
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
                        'Get-SourceTextLoc', 'Write-ToolkitLog', 'Invoke-ExternalCommandWithLog',
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
                $result.Helpers.Count | Should -Be 9
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

    Context 'Folder removal failures with mocked Windows operations' {
        BeforeEach {
            $script:ProbeFolderPath = Join-Path $TestDrive 'Issue189[Probe]'
            [void][System.IO.Directory]::CreateDirectory($script:ProbeFolderPath)
            $script:ProbeProfile = [pscustomobject]@{
                LocalPath = $script:ProbeFolderPath
                SID = 'S-1-5-21-189-1'
            }
            $script:RemovalAttempts = 0
            Mock Remove-CimInstance { throw 'Synthetic CIM removal failure' }
            Mock Remove-ProfileRegistryEntries { $true }
            Mock Invoke-ExternalCommandWithLog { [pscustomobject]@{ Success = $true; ExitCode = 0 } }
            Mock Remove-ItemSafely { $false }
            Mock Test-Path { $true } -ParameterFilter { $Path -eq (Join-Path $env:TEMP 'EmptyFolder') }
            Mock Clear-ProgressLine {}
            Mock Write-ProgressUpdate {}
        }

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
            Should -Invoke Invoke-ExternalCommandWithLog -Times 1 -Exactly -ParameterFilter { $Command -eq 'takeown.exe' }
            Should -Invoke Invoke-ExternalCommandWithLog -Times 1 -Exactly -ParameterFilter {
                $Command -eq 'robocopy.exe' -and $Arguments -contains '/XJ'
            }
            Should -Invoke Invoke-ExternalCommandWithLog -Times 1 -Exactly -ParameterFilter {
                $Command -eq 'icacls.exe' -and $Arguments -contains '*S-1-5-32-544:F'
            }
            Should -Invoke Remove-ProfileRegistryEntries -Times 1 -Exactly -ParameterFilter { $Sid -eq $script:ProbeProfile.SID }
            Should -Invoke Remove-ItemSafely -Times 0 -Exactly
        }

        It 'reports failure and preserves the registry when both folder removal attempts fail' {
            Mock Remove-Item { throw 'Synthetic access denied' }

            $result = & $script:ProfileRemovalWorker $script:ProbeProfile

            $result.Success | Should -BeFalse
            [System.IO.Directory]::Exists($script:ProbeFolderPath) | Should -BeTrue
            Should -Invoke Remove-Item -Times 2 -Exactly
            Should -Invoke Remove-ProfileRegistryEntries -Times 0 -Exactly
            Get-Content -LiteralPath $Global:CurrentLogFile -Raw | Should -Match '\[ERROR\].*Synthetic access denied'
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
            Should -Invoke Invoke-ExternalCommandWithLog -Times 1 -Exactly -ParameterFilter {
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

    Context 'Bounded profile scheduling with synthetic workers' {
        BeforeAll {
            $syntheticWorker = @'
$scriptBlock = {
    param($ProfileItem)
    Start-Sleep -Milliseconds $ProfileItem.DelayMilliseconds
    if ($ProfileItem.Fail) { throw 'Synthetic bounded-worker failure' }
    [pscustomobject]@{
        Type = 'Profile'
        UserName = $ProfileItem.Name
        Success = $true
        Duration = [TimeSpan]::Zero
    }
}
'@
            $probeSource = $script:ProfileRemovalBatchDefinition.Remove(
                $script:ProfileRemovalWorkerOffset, $script:ProfileRemovalWorkerLength
            ).Insert($script:ProfileRemovalWorkerOffset, $syntheticWorker)
            $probeSource = $probeSource.Replace('$ps = [PowerShell]::Create()', '$ps = New-TrackedProfilePowerShell')
            $probeSource = $probeSource.Replace('$handle = $ps.BeginInvoke()', '$handle = Invoke-ProfileProbeBeginInvoke -PowerShell $ps')
            if ($probeSource -match '\b(Remove-CimInstance|Remove-Item|Invoke-ExternalCommandWithLog)\b') {
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
            Mock New-ProfileRemovalSessionState { [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault() }
            Mock Write-ProgressUpdate { $script:BatchProgress.Add($Percent) }
            Mock Clear-ProgressLine {}
            Mock Receive-ProfileRemovalResult {
                $script:BatchReceivedNames.Add([System.IO.Path]::GetFileName($Job.Profile.LocalPath))
                $result = & $script:ProfileBatchResultReceiver -Job $Job
                $script:BatchCollectedCount++
                $result
            }
        }

        AfterEach {
            foreach ($powerShell in $script:BatchCreatedPowerShells) { $powerShell.Dispose() }
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
            if ($probeSource -match '\b(Remove-CimInstance|Remove-Item|Invoke-ExternalCommandWithLog|Remove-ResidualUserFolder)\b') {
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
