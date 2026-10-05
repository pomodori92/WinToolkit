# WinToolkit CI/CD V4.2.0
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
Regression coverage for issue #189 without running profile deletion.
Loads individual function definitions from the AST, never the tool or toolkit scripts.
Only initialization, localization, logging, and result collection are invoked.
Removal helpers are registered as text and inspected, never called.
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
            @{ Path = $toolPath; Name = 'Remove-ProfileRegistryEntries' },
            @{ Path = $localizationPath; Name = 'Get-SourceTextLoc' },
            @{ Path = $loggingPath; Name = 'Write-ToolkitLog' },
            @{ Path = $processesPath; Name = 'Invoke-ExternalCommandWithLog' },
            @{ Path = $processesPath; Name = 'Remove-ItemSafely' },
            @{ Path = $uiPath; Name = 'Get-SpinnerChar' }
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
                        'Remove-ItemSafely', 'Remove-ProfileRegistryEntries', 'Get-SpinnerChar',
                        'Clear-ProgressLine', 'Write-ProgressUpdate'
                    ) | ForEach-Object { (Get-Command -Name $_ -CommandType Function -ErrorAction Stop).Name }
                }
            })

            $results.Count | Should -Be 4
            foreach ($result in $results) {
                $result.Message | Should -Be 'Profili registrati rimossi: 0'
                $result.LogPath | Should -Be $Global:CurrentLogFile
                $result.ToolName | Should -Be 'Issue189Probe'
                $result.Helpers.Count | Should -Be 8
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
}
