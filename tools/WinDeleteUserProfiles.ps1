# Coordinate registered-profile and residual-folder cleanup with shared protections and result reporting.
function WinDeleteUserProfiles {
    <#
    .SYNOPSIS
        Safely removes unloaded local user profiles and residual folders from C:\Users.

    .DESCRIPTION
        Performs a controlled cleanup of local profiles in C:\Users using Win32_UserProfile.
        Excludes special and loaded profiles, system accounts, the current user profile, and protected names.
        After deleting registered profiles, checks the users folder and removes residual directories that are no
        longer associated with profiles in the registry or CIM, while preserving all protected exclusions.
        Residual folders are divided into independent deletion tasks and processed with bounded parallelism.

        The script does not request interactive confirmation before deletion.

    .PARAMETER MaxThreads
        Maximum number of parallel runspaces for each cleanup phase. Automatically limited to 4.

    .PARAMETER CountdownSeconds
        Number of seconds in the countdown before a recommended restart.

    .PARAMETER SuppressIndividualReboot
        Suppresses the individual restart and delegates the final restart to the toolkit.

    .PARAMETER UsersRoot
        Root path of the local user profiles.

    .PARAMETER MinimumProfileAgeDays
        Minimum profile age in days since its last use. The default value 0 preserves the original behavior.

    .PARAMETER SkipResidualFolderCleanup
        Skips the final cleanup of residual folders in C:\Users.

    .EXAMPLE
        WinDeleteUserProfiles

    .EXAMPLE
        WinDeleteUserProfiles -MinimumProfileAgeDays 30
    #>
    [CmdletBinding()]
    param(
        [ValidateRange(1, 4)]
        [int]$MaxThreads = [Math]::Min(2, [Environment]::ProcessorCount),

        [ValidateRange(0, 3600)]
        [int]$CountdownSeconds = 30,

        [switch]$SuppressIndividualReboot,

        [ValidateNotNullOrEmpty()]
        [string]$UsersRoot = 'C:\Users',

        [ValidateRange(0, 3650)]
        [int]$MinimumProfileAgeDays = 0,

        [switch]$SkipResidualFolderCleanup
    )

    $script:ToolName = 'WinDeleteUserProfiles'
    # Keep a trailing separator so root checks cannot match sibling paths with the same prefix.
    $usersRoot = [System.IO.Path]::GetFullPath($UsersRoot.TrimEnd('\') + '\')
    $currentUser = $env:USERNAME
    $computerName = $env:COMPUTERNAME
    # Resolve the current profile by SID because its folder name can differ from the account name.
    $currentUserSid = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
    $currentUserProfilePath = (Get-CimInstance -ClassName Win32_UserProfile -ErrorAction SilentlyContinue | Where-Object { $_.SID -eq $currentUserSid } | Select-Object -First 1 -ExpandProperty LocalPath -ErrorAction SilentlyContinue)
    if ($currentUserProfilePath -and $currentUserProfilePath.StartsWith($usersRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        $currentUserFolder = [System.IO.Path]::GetFileName($currentUserProfilePath)
    }
    else {
        $currentUserFolder = $currentUser
    }
    $minimumLastUseDate = if ($MinimumProfileAgeDays -gt 0) { (Get-Date).AddDays(-$MinimumProfileAgeDays) } else { $null }
    $rebootRecommended = $false
    # Cap simultaneous CIM operations and the resources held by the worker batch.
    $maxThreadsEffective = [Math]::Min($MaxThreads, 4)

    # Preserve shared/system folders and both names that can identify the current user's profile.
    $protectedProfileNames = @(
        'Public',
        'Pubblica',
        'Default',
        'Default User',
        'All Users',
        'defaultuser0',
        'WDAGUtilityAccount',
        'Administrator',
        'Guest',
        $currentUser,
        $currentUserFolder
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    # Apply the same case-insensitive name exclusions to registered profiles and residual folders.
    function New-ProtectedNameSet {
        $excluded = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $protectedProfileNames | ForEach-Object { [void]$excluded.Add($_) }
        # The unary comma returns the set itself instead of enumerating its entries into the pipeline.
        return , $excluded
    }

    # Remove leftover ProfileList state after the profile directory is gone, reporting registry failure separately.
    # A single SID is a cleanup leaf; different profiles already reach it through the bounded removal pool.
    function Remove-ProfileRegistryEntries {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Sid,

            [Parameter(Mandatory = $true)]
            [string]$UserName
        )

        # Never construct a destructive registry path from a missing or placeholder SID.
        if ([string]::IsNullOrWhiteSpace($Sid) -or $Sid -eq 'NULL') { 
            return $false
        }

        $profileRegKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\' + $Sid

        # CIM removal can already delete this key; an absent entry counts as successful cleanup.
        if (-not (Test-Path -LiteralPath $profileRegKey)) {
            return $true
        }

        try {
            Remove-Item -LiteralPath $profileRegKey -Recurse -Force -ErrorAction Stop
            Write-ToolkitLog -Level 'SUCCESS' -Message (Get-SourceTextLoc 'toolText.registryEntryRemovedForSid01' -Args @($UserName, $Sid))
            
            return $true
        }
        catch {
            Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.failedToRemoveRegistryEntryForSid01' -Args @($UserName, $($_.Exception.Message)))
            
            return $false
        }
    }

    # Collect normalized registered paths so the residual scan preserves folders still associated with CIM profiles.
    function Get-RegisteredProfilePathSet {
        $pathSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        try {
            $cimProfiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop
        }
        catch {
            Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.cimProfileDetectionFailed0' -Args @($($_.Exception.Message)))
            
            # Unknown registration must not turn registered folders into residual deletion candidates.
            throw
        }

        $cimProfiles |
        Where-Object {
            $_.LocalPath -and
            $_.LocalPath.StartsWith($usersRoot, [System.StringComparison]::OrdinalIgnoreCase)
        } |
        ForEach-Object {
            try {
                [void]$pathSet.Add([System.IO.Path]::GetFullPath($_.LocalPath).TrimEnd('\'))
            }
            catch {
                Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.failedToNormalizeRegisteredProfileLocalpath0' -Args @($($_.LocalPath)))
                throw
            }
        }

        return , $pathSet
    }

    # Native metadata reads preserve literal paths and expose errors that Directory.Exists suppresses.
    function Get-ProfileCleanupPathAttribute {
        param([Parameter(Mandatory = $true)][string]$Path)
        return [System.IO.File]::GetAttributes($Path)
    }

    # Only a path-not-found error proves absence; access and I/O failures must preserve profile registration.
    function Test-ProfileCleanupDirectory {
        param([Parameter(Mandatory = $true)][string]$Path)

        try {
            $attributes = Get-ProfileCleanupPathAttribute -Path $Path
        }
        catch [System.IO.FileNotFoundException] {
            # A missing volume or network share must not masquerade as a deleted profile directory.
            [void](Get-ProfileCleanupPathAttribute -Path ([System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($Path))))
            return $false
        }
        catch [System.IO.DirectoryNotFoundException] {
            [void](Get-ProfileCleanupPathAttribute -Path ([System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($Path))))
            return $false
        }

        if ($attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            throw [System.IO.IOException]::new("Profile cleanup cannot follow a reparse point: '$Path'.")
        }
        if (-not ($attributes -band [System.IO.FileAttributes]::Directory)) {
            throw [System.IO.IOException]::new("Profile cleanup target is not a directory: '$Path'.")
        }
        return $true
    }

    # CIM refusal can mean that a session loaded the profile after discovery; verify before filesystem fallback.
    function Assert-ProfileCleanupProfileState {
        param([Parameter(Mandatory = $true)][object]$Profile)

        $sid = $Profile.SID
        if ($sid -notmatch '^S-\d+(?:-\d+)+$') {
            throw [System.InvalidOperationException]::new('Profile cleanup requires a valid profile SID.')
        }
        $current = @(Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$sid'" -ErrorAction Stop)
        if ($current.Count -ne 1 -or $current[0].Loaded -ne $false -or $current[0].Special -ne $false) {
            throw [System.InvalidOperationException]::new("Profile '$sid' is loaded, special, or cannot be safely verified.")
        }
        if ([string]::IsNullOrWhiteSpace($current[0].LocalPath) -or
            -not [string]::Equals([System.IO.Path]::GetFullPath($current[0].LocalPath).TrimEnd('\'),
                [System.IO.Path]::GetFullPath($Profile.LocalPath).TrimEnd('\'), [System.StringComparison]::OrdinalIgnoreCase)) {
            throw [System.InvalidOperationException]::new("Profile '$sid' no longer maps to the selected cleanup path.")
        }
    }

    # Select unloaded, non-special profiles under the configured root before scheduling destructive work.
    function Get-RemovableUserProfiles {
        $excluded = New-ProtectedNameSet

        Write-StyledMessage -Type 'Info' -Text ("🔍 " + (Get-SourceTextLoc 'toolText.scanningRegisteredLocalProfiles'))
        Write-ToolkitLog -Level 'INFO' -Message (Get-SourceTextLoc 'toolText.scanningRegisteredLocalProfiles')

        try {
            $profiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop
        }
        catch {
            Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.cimProfileDetectionFailed0' -Args @($($_.Exception.Message)))
            Write-StyledMessage -Type 'Warning' -Text ((Get-SourceTextLoc 'toolText.cimProfileDetectionFailed0' -Args @($($_.Exception.Message))))
            return
        }

        # Active sessions and Windows-managed special profiles must keep their profile data.
        $profiles = $profiles | Where-Object {
            -not $_.Special -and
            -not $_.Loaded -and
            $_.LocalPath -and
            $_.LocalPath.StartsWith($usersRoot, [System.StringComparison]::OrdinalIgnoreCase)
        }

        foreach ($profileItem in $profiles) {
            $profileName = [System.IO.Path]::GetFileName($profileItem.LocalPath)

            if ($excluded.Contains($profileName)) {
                Write-ToolkitLog -Level 'INFO' -Message (Get-SourceTextLoc 'toolText.excludedProfile01' -Args @($profileName, $($profileItem.LocalPath)))
                continue
            }

            # Protect the current account independently of folder-name exclusions.
            if ($profileItem.SID -eq $currentUserSid) {
                Write-ToolkitLog -Level 'INFO' -Message (Get-SourceTextLoc 'toolText.excludedProfileBecauseSidMatchesCurrentUser0' -Args @($profileName, $profileItem.SID))
                continue
            }

            # Apply the optional age cutoff only when Windows provides a last-use timestamp.
            if ($minimumLastUseDate -and $profileItem.LastUseTime) {
                $lastUse = $profileItem.LastUseTime
                if ($lastUse -gt $minimumLastUseDate) {
                    Write-ToolkitLog -Level 'INFO' -Message (Get-SourceTextLoc 'toolText.profileExcludedDueToTimeThreshold0LastUse1' -Args @($profileName, $lastUse))
                    continue
                }
            }

            $profileItem
        }
    }

    # Cleanup commands need bounded capture: recursive ACL tools can print one line for every file.
    function Invoke-ProfileCleanupCommand {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Command,
            [string[]]$Arguments = @(),
            [string]$LogContextKey = ''
        )

        Write-ToolkitLog -Level 'INFO' -Message (Get-SourceTextLoc 'uiText.commandContext') -Context @{
            Command = $Command; Arguments = $Arguments; ContextKey = $LogContextKey
        }
        $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
        $startInfo.FileName = $Command
        $startInfo.Arguments = $Arguments -join ' '
        $startInfo.UseShellExecute = $false
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $startInfo.CreateNoWindow = $true
        $process = [System.Diagnostics.Process]::new()
        $process.StartInfo = $startInfo
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        $started = $false
        $exitCode = -1
        $streams = @()

        try {
            $started = $process.Start()
            if (-not $started) { throw (Get-SourceTextLoc 'uiText.unableToStartExternalProcess') }
            foreach ($reader in @($process.StandardOutput, $process.StandardError)) {
                $buffer = [char[]]::new(4096)
                $streams += [PSCustomObject]@{
                    Reader = $reader; Buffer = $buffer; Text = [System.Text.StringBuilder]::new(8000)
                    Read = $reader.ReadAsync($buffer, 0, $buffer.Length); Ended = $false; Truncated = $false
                }
            }

            # Drain both pipes even after their snippets are full, so a noisy process cannot block on output.
            while ($true) {
                $readAvailable = $false
                foreach ($stream in $streams) {
                    if ($stream.Ended -or -not $stream.Read.IsCompleted) { continue }
                    $readAvailable = $true
                    $count = $stream.Read.GetAwaiter().GetResult()
                    if ($count -eq 0) {
                        $stream.Ended = $true
                        $stream.Read = $null
                        continue
                    }
                    $retain = [Math]::Min($count, 8000 - $stream.Text.Length)
                    if ($retain -gt 0) { [void]$stream.Text.Append($stream.Buffer, 0, $retain) }
                    if ($retain -lt $count) { $stream.Truncated = $true }
                    $stream.Read = $stream.Reader.ReadAsync($stream.Buffer, 0, $stream.Buffer.Length)
                }
                if ($process.HasExited -and $streams[0].Ended -and $streams[1].Ended) {
                    $exitCode = $process.ExitCode
                    break
                }
                # Short waits let PowerShell cancellation reach finally instead of blocking in WaitForExit().
                if (-not $readAvailable) {
                    $pendingReads = [System.Threading.Tasks.Task[]]@($streams | Where-Object { -not $_.Ended } | ForEach-Object { $_.Read })
                    if ($pendingReads.Count -gt 0) { [void][System.Threading.Tasks.Task]::WaitAny($pendingReads, 20) }
                    else { [System.Threading.Thread]::Sleep(20) }
                }
            }
        }
        catch [System.Management.Automation.PipelineStoppedException] { throw }
        catch {
            Write-ToolkitLog -Level 'ERROR' -Message (Get-SourceTextLoc 'uiText.exceptionWhileRunningExternalCommand') -Context @{
                Command = $Command; Arguments = $Arguments
                ContextKey = $LogContextKey; Exception = $_.Exception.Message; Stack = $_.ScriptStackTrace
            }
        }
        finally {
            # Cancellation can bypass catch; always terminate this invocation's child before releasing its pipes.
            try {
                if ($started -and -not $process.HasExited) {
                    try { $process.Kill() }
                    catch { Write-Warning -Message $_.Exception.Message -WarningAction Continue }
                    [void]$process.WaitForExit(1000)
                }
            }
            finally {
                try { foreach ($stream in $streams) { $stream.Reader.Dispose() } }
                finally { $process.Dispose(); $timer.Stop() }
            }
        }

        $outText = if ($streams.Count -gt 0) { $streams[0].Text.ToString() } else { '' }
        $errText = if ($streams.Count -gt 1) { $streams[1].Text.ToString() } else { '' }
        if ($streams.Count -gt 0 -and $streams[0].Truncated) { $outText += "`n[...output truncated...]" }
        if ($streams.Count -gt 1 -and $streams[1].Truncated) { $errText += "`n[...stderr truncated...]" }
        $success = $exitCode -eq 0
        $status = if ($success) { Get-SourceTextLoc 'sourceText.completedSuccessfully' } else { Get-SourceTextLoc 'sourceText.completedWithErrors' }
        Write-ToolkitLog -Level 'INFO' -Message (Get-SourceTextLoc 'uiText.command0ExitCode1Duration2' -Args @($status, $exitCode, $timer.Elapsed.ToString('hh\:mm\:ss')))
        Write-ToolkitLog -Level 'DEBUG' -Message (Get-SourceTextLoc 'uiText.commandOutput0' -Args @($Command)) -Context @{
            ContextKey = $LogContextKey; StdOutSnippet = $outText; StdErrSnippet = $errText
        }
        [PSCustomObject]@{
            Success = $success; ExitCode = $exitCode; StdOut = $outText; StdErr = $errText
            Elapsed = $timer.Elapsed
        }
    }

    # Supply worker dependencies explicitly: runspaces do not inherit toolkit functions or language/log context.
    # Build this small, fixed dependency set once per pool; there is no bulk work to partition here.
    function New-ProfileRemovalSessionState {
        $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

        foreach ($functionName in @(
                'Get-SourceTextLoc',
                'Write-ToolkitLog',
                'Invoke-ProfileCleanupCommand',
                'Get-ProfileCleanupPathAttribute',
                'Test-ProfileCleanupDirectory',
                'Assert-ProfileCleanupProfileState',
                'Remove-ItemSafely',
                'Remove-ProfileRegistryEntries',
                'Remove-ResidualUserFolder',
                'Get-SpinnerChar'
            )) {
            $command = Get-Command -Name $functionName -CommandType Function -ErrorAction Stop
            $sessionState.Commands.Add(
                [System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($functionName, $command.Definition)
            )
        }

        # Only the parent runspace may update the shared console progress line.
        foreach ($functionName in @('Clear-ProgressLine', 'Write-ProgressUpdate')) {
            $sessionState.Commands.Add(
                [System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($functionName, '')
            )
        }

        foreach ($variableName in @(
                'SourceTextLanguageData',
                'SourceTextDefaultLanguageData',
                'SourceTextKeyAliases',
                'CurrentLogFile',
                'CurrentToolName',
                'Spinners'
            )) {
            $value = Get-Variable -Name $variableName -Scope Global -ValueOnly -ErrorAction Stop
            $sessionState.Variables.Add(
                [System.Management.Automation.Runspaces.SessionStateVariableEntry]::new($variableName, $value, '')
            )
        }
        # Shared helpers must leave console rendering to the parent even inside worker sessions.
        $sessionState.Variables.Add(
            [System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('GuiSessionActive', $true, '')
        )

        return $sessionState
    }

    # Collect a finished deletion worker's result and release its pipeline before scheduling replacement work.
    function Receive-ProfileRemovalResult {
        param(
            [Parameter(Mandatory = $true)]
            [object]$Job,

            [ValidateSet('Profile', 'ResidualFolder')]
            [string]$ResultType = 'Profile'
        )

        try {
            $Job.PowerShell.EndInvoke($Job.Handle)
        }
        catch {
            Write-ToolkitLog -Level 'ERROR' -Message (Get-SourceTextLoc 'toolText.runspaceError0' -Args @($($_.Exception.Message)))
            # Preserve identity in a failed result so worker exceptions remain visible in summary counts.
            if ($ResultType -eq 'ResidualFolder') {
                [PSCustomObject]@{
                    Type     = 'ResidualFolder'
                    UserName = $Job.Folder.Name
                    Path     = $Job.Folder.Path
                    Success  = $false
                    Duration = [TimeSpan]::Zero
                }
            }
            else {
                [PSCustomObject]@{
                    Type     = 'Profile'
                    UserName = [System.IO.Path]::GetFileName($Job.Profile.LocalPath)
                    Path     = $Job.Profile.LocalPath
                    Sid      = $Job.Profile.SID
                    Success  = $false
                    Duration = [TimeSpan]::Zero
                }
            }
        }
        finally {
            # Release command references and pipeline resources as soon as collection finishes.
            try { $Job.PowerShell.Commands.Clear() }
            finally {
                try { if ($Job.WaitHandle) { $Job.WaitHandle.Dispose() } }
                finally { $Job.PowerShell.Dispose() }
            }
        }
    }

    # Clean up workers abandoned by a batch failure while keeping individual cleanup errors non-terminating.
    # Stop, EndInvoke, and Dispose form one ordered leaf operation before the owning pool closes.
    function Close-ProfileRemovalPowerShell {
        param(
            [Parameter(Mandatory = $true)]
            [object]$PowerShell,
            [System.IAsyncResult]$Handle,
            [System.Threading.WaitHandle]$WaitHandle
        )

        try {
            $PowerShell.Stop()
            if ($Handle) {
                # EndInvoke releases the async result's resources even when stopping made it throw.
                try { [void]$PowerShell.EndInvoke($Handle) }
                catch [System.Management.Automation.PipelineStoppedException] {}
                catch [System.ObjectDisposedException] {}
            }
        }
        # Collection may have disposed the instance before a later batch operation failed.
        catch [System.ObjectDisposedException] {}
        catch {
            # Explicit Continue prevents the caller's warning preference from interrupting cleanup.
            Write-Warning -Message $_.Exception.Message -WarningAction Continue
        }
        finally {
            # Dispose must still run when Stop fails.
            try {
                try { if ($WaitHandle) { $WaitHandle.Dispose() } }
                finally { $PowerShell.Dispose() }
            }
            catch {
                Write-Warning -Message $_.Exception.Message -WarningAction Continue
            }
        }
    }

    # Native directory iterators avoid Get-ChildItem materializing a wide directory before yielding paths.
    function New-ProfileDirectoryEnumerator {
        param([Parameter(Mandatory = $true)][string]$Path)
        return , ([System.IO.Directory]::EnumerateDirectories($Path).GetEnumerator())
    }

    # Partition leftovers into disjoint directory subtrees, keeping only the current branch in memory.
    function Get-ProfileCleanupSubtrees {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path,
            [switch]$AsEnumerator
        )

        try {
            $root = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
            if ($root.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { return }
            $directories = New-ProfileDirectoryEnumerator -Path $Path
        }
        catch {
            # Enumeration failure leaves the original whole-profile cleanup responsible for recovery.
            Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.standardFolderCleanupFailed01' -Args @($Path, $($_.Exception.Message)))
            return
        }

        $cursor = [PSCustomObject]@{
            Path = $Path; Directories = $directories; Children = $null
            Branch = $null; BranchEmitted = $false; Current = $null; Disposed = $false
        }
        $cursor | Add-Member -MemberType ScriptMethod -Name Dispose -Value {
            try { if ($this.Children -is [System.IDisposable]) { $this.Children.Dispose() } }
            finally {
                try { if ($this.Directories -is [System.IDisposable]) { $this.Directories.Dispose() } }
                finally {
                    $this.Children = $null; $this.Directories = $null; $this.Current = $null
                    $this.Branch = $null; $this.Disposed = $true
                }
            }
        }
        $cursor | Add-Member -MemberType ScriptMethod -Name MoveNext -Value {
            if ($this.Disposed) { return $false }
            while ($true) {
                if ($null -ne $this.Children) {
                    try {
                        while ($this.Children.MoveNext()) {
                            $childPath = [string]$this.Children.Current
                            $child = Get-Item -LiteralPath $childPath -Force -ErrorAction Stop
                            if ($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                            $this.BranchEmitted = $true
                            $this.Current = $childPath
                            return $true
                        }
                    }
                    catch {
                        Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.standardFolderCleanupFailed01' -Args @($this.Branch, $_.Exception.Message))
                    }
                    if ($this.Children -is [System.IDisposable]) { $this.Children.Dispose() }
                    $this.Children = $null
                    # Once any child was yielded, its parent must wait for root finalization to avoid overlap.
                    if (-not $this.BranchEmitted) {
                        $this.Current = $this.Branch
                        return $true
                    }
                }

                try {
                    if (-not $this.Directories.MoveNext()) {
                        $this.Dispose()
                        return $false
                    }
                    $this.Branch = [string]$this.Directories.Current
                }
                catch {
                    try {
                        Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.standardFolderCleanupFailed01' -Args @($this.Path, $_.Exception.Message))
                    }
                    finally { $this.Dispose() }
                    return $false
                }
                $this.BranchEmitted = $false
                try {
                    $directory = Get-Item -LiteralPath $this.Branch -Force -ErrorAction Stop
                    if ($directory.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                    # A second level separates AppData branches; files in their parents remain for finalization.
                    $this.Children = New-ProfileDirectoryEnumerator -Path $this.Branch
                }
                catch {
                    Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.standardFolderCleanupFailed01' -Args @($this.Branch, $_.Exception.Message))
                    $this.Current = $this.Branch
                    return $true
                }
            }
        }

        if ($AsEnumerator) { return $cursor }
        # Preserve the path-producing interface for callers that do not need incremental scheduling.
        try { while ($cursor.MoveNext()) { $cursor.Current } }
        finally { $cursor.Dispose() }
    }

    # Divide leftover subtrees in one shared pool, then combine them into one result per original profile.
    function Invoke-ProfileRemovalBatch {
        param(
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [array]$Profiles
        )

        # Retain only outstanding jobs so pipeline memory stays bounded by the concurrency limit.
        $jobs = [System.Collections.Generic.List[object]]::new()
        # Keep ownership of a new instance until startup succeeds and it joins the tracked job list.
        $pendingPowerShell = $null
        $pendingHandle = $null
        $pendingWaitHandle = $null
        $pool = $null
        $activePlans = [System.Collections.Generic.HashSet[object]]::new()
        $total = $Profiles.Count
        $results = [object[]]::new($total)
        $completed = 0
        $lastPercent = -1

        # Store index ranges, splitting lazily until a free worker can take one leaf.
        $ranges = [System.Collections.Generic.Stack[object]]::new()
        if ($total -gt 0) {
            $ranges.Push([PSCustomObject]@{ Phase = 'Prepare'; Start = 0; Count = $total; Context = $null })
        }

        $scriptBlock = {
            param(
                $ProfileItem,
                [ValidateSet('Complete', 'Prepare', 'Folder', 'Finalize')]
                [string]$Phase = 'Complete',
                [datetime]$StartTime = [datetime]::MinValue,
                $ParentProfile = $null,
                [bool]$CheckProfileState = $true
            )

            # Terminating errors route failed deletion steps through their fallback cleanup paths.
            $ErrorActionPreference = 'Stop'

            $userPath = $ProfileItem.LocalPath
            $userName = [System.IO.Path]::GetFileName($userPath)
            $userSid = $ProfileItem.SID
            $start = if ($StartTime -eq [datetime]::MinValue) { Get-Date } else { $StartTime }
            $profileToCheck = if ($ParentProfile) { $ParentProfile } else { $ProfileItem }

            # Guard roots as well as children before CIM, robocopy, deletion, or ACL recovery can touch a link.
            [void](Test-ProfileCleanupDirectory -Path $userPath)
            if ($ParentProfile) { [void](Test-ProfileCleanupDirectory -Path $ParentProfile.LocalPath) }

            if ($Phase -ne 'Finalize') {
                Write-ToolkitLog -Level 'INFO' -Message (Get-SourceTextLoc 'toolText.startResidualFolder01' -Args @($userName, $userPath))
            }

            # Let Windows remove the registered profile first; filesystem cleanup handles any leftovers.
            if ($Phase -in @('Complete', 'Prepare')) {
                if ($ProfileItem.Loaded -eq $true -or $ProfileItem.Special -eq $true) {
                    throw [System.InvalidOperationException]::new("Profile '$userSid' is loaded or special.")
                }
                $CheckProfileState = $false
                try {
                    Remove-CimInstance -InputObject $ProfileItem -ErrorAction Stop -Confirm:$false
                    Write-ToolkitLog -Level 'SUCCESS' -Message (Get-SourceTextLoc 'toolText.cimProfileRemoved0' -Args @($userName))
                }
                catch {
                    Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.cimRemoveFailed01' -Args @($userName, $($_.Exception.Message)))
                    $CheckProfileState = $true
                }
            }

            # Only failed CIM removal leaves a registered profile requiring fresh checks in subsequent workers.
            if ($CheckProfileState) { Assert-ProfileCleanupProfileState -Profile $profileToCheck }
            $directoryExists = Test-ProfileCleanupDirectory -Path $userPath
            if ($Phase -eq 'Prepare' -and $directoryExists) {
                # Return control to the coordinator before any child or parent filesystem deletion starts.
                return [PSCustomObject]@{ Type = 'ProfilePreparation'; Start = $start; CheckProfileState = $CheckProfileState }
            }

            if ($directoryExists) {
                try {
                    # An empty source lets /MIR remove leftover contents; /XJ excludes junctions.
                    $tempEmpty = Join-Path $env:TEMP "EmptyFolder"

                    if (-not (Test-Path $tempEmpty)) {
                        # Concurrent subtree workers can create the shared empty directory safely.
                        [void][System.IO.Directory]::CreateDirectory($tempEmpty)
                    }

                    [void](Test-ProfileCleanupDirectory -Path $userPath)
                    Invoke-ProfileCleanupCommand -Command 'robocopy.exe' `
                        -Arguments @("`"$tempEmpty`"", "`"$userPath`"", '/MIR', '/XJ', '/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS', '/NP') `
                        -LogContextKey "ProfileCleanup-Robocopy-$userName" | Out-Null

                    if (Test-ProfileCleanupDirectory -Path $userPath) {
                        Remove-Item -LiteralPath $userPath -Recurse -Force -ErrorAction Stop -Confirm:$false
                    }

                    Write-ToolkitLog -Level 'SUCCESS' -Message (Get-SourceTextLoc 'toolText.folderRemoved0' -Args @($userName))
                }
                catch {
                    Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.standardFolderCleanupFailed01' -Args @($userName, $($_.Exception.Message)))

                    # Repair ownership and access if permissions prevented standard folder removal.
                    try {
                        if ($CheckProfileState) { Assert-ProfileCleanupProfileState -Profile $profileToCheck }
                        if (Test-ProfileCleanupDirectory -Path $userPath) {
                            Invoke-ProfileCleanupCommand -Command 'takeown.exe' -Arguments @('/F', "`"$userPath`"", '/R', '/D', 'Y', '/SKIPSL') -LogContextKey "ProfileCleanup-TakeOwn-$userName" | Out-Null
                        }
                        # S-1-5-32-544 - Fixed SID across Windows languages; avoids depending on localized group names.
                        if (Test-ProfileCleanupDirectory -Path $userPath) {
                            Invoke-ProfileCleanupCommand -Command 'icacls.exe' -Arguments @("`"$userPath`"", '/grant', '*S-1-5-32-544:F', '/T', '/C', '/L') -LogContextKey "ProfileCleanup-Icacls-$userName" | Out-Null
                        }
                        if (Test-ProfileCleanupDirectory -Path $userPath) {
                            Remove-Item -LiteralPath $userPath -Recurse -Force -ErrorAction Stop -Confirm:$false
                        }
                        Write-ToolkitLog -Level 'SUCCESS' -Message (Get-SourceTextLoc 'toolText.folderRemovedAfterAclReset0' -Args @($userName))
                    }
                    catch {
                        Write-ToolkitLog -Level 'ERROR' -Message (Get-SourceTextLoc 'toolText.cleanupFailed01' -Args @($userName, $($_.Exception.Message)))
                    }
                }
            }

            $folderGone = -not (Test-ProfileCleanupDirectory -Path $userPath)
            if ($Phase -eq 'Folder') {
                # Child results never touch profile registration or count as completed profiles.
                return [PSCustomObject]@{
                    Type = 'ResidualFolder'; UserName = $userName; Path = $userPath
                    Success = $folderGone; Duration = (New-TimeSpan -Start $start -End (Get-Date))
                }
            }

            # Remove any remaining registry entry only after the directory is gone; require both outcomes for success.
            $registrySuccess = $false
            if ($folderGone -and $userSid) {
                $registrySuccess = Remove-ProfileRegistryEntries -Sid $userSid -UserName $userName
            }

            $success = $folderGone -and $registrySuccess
            $duration = New-TimeSpan -Start $start -End (Get-Date)

            if ($success) {
                Write-ToolkitLog -Level 'SUCCESS' -Message (Get-SourceTextLoc 'toolText.completedProfile01' -Args @($userName, $duration.ToString()))
            }
            else {
                Write-ToolkitLog -Level 'ERROR' -Message (Get-SourceTextLoc 'toolText.failedProfile0' -Args @($userName))
            }

            return [PSCustomObject]@{
                Type     = 'Profile'
                UserName = $userName
                Path     = $userPath
                Sid      = $userSid
                Success  = $success
                Duration = $duration
            }
        }

        try {
            if ($total -gt 0) {
                $sessionState = New-ProfileRemovalSessionState
                $pool = [RunspaceFactory]::CreateRunspacePool(1, $maxThreadsEffective, $sessionState, $Host)
                $pool.Open()
            }

            while ($ranges.Count -gt 0 -or $jobs.Count -gt 0) {
                # Walk backwards so removing completed jobs cannot shift an unvisited entry.
                for ($jobIndex = $jobs.Count - 1; $jobIndex -ge 0; $jobIndex--) {
                    $job = $jobs[$jobIndex]
                    if (-not $job.Handle.IsCompleted) { continue }

                    $resultType = if ($job.Phase -eq 'Folder') { 'ResidualFolder' } else { 'Profile' }
                    $result = Receive-ProfileRemovalResult -Job $job -ResultType $resultType
                    $jobs.RemoveAt($jobIndex)

                    if ($job.Phase -eq 'Folder') {
                        $job.Context.Remaining--
                        if ($job.Context.Remaining -eq 0 -and $job.Context.EnumerationComplete) {
                            # Finalize this parent only after every child has been collected, including failures.
                            $ranges.Push([PSCustomObject]@{ Phase = 'Finalize'; Start = 0; Count = 1; Context = $job.Context })
                        }
                    }
                    elseif ($result.Type -eq 'ProfilePreparation') {
                        $plan = Get-ProfileCleanupSubtrees -Path $job.Profile.LocalPath -AsEnumerator
                        if ($plan) { [void]$activePlans.Add($plan) }
                        $context = [PSCustomObject]@{
                            Profile = $job.Profile; Index = $job.Index; Start = $result.Start
                            CheckProfileState = $result.CheckProfileState
                            Plan = $plan; PendingPaths = [System.Collections.Generic.Queue[string]]::new()
                            Remaining = 0; EnumerationComplete = $false
                        }
                        # Two-path lookahead preserves whole-root cleanup for small trees without retaining a manifest.
                        if ($plan -and $plan.MoveNext()) { $context.PendingPaths.Enqueue($plan.Current) }
                        if ($context.PendingPaths.Count -gt 0 -and $plan.MoveNext()) {
                            $context.PendingPaths.Enqueue($plan.Current)
                            $ranges.Push([PSCustomObject]@{ Phase = 'Enumerate'; Start = 0; Count = 1; Context = $context })
                        }
                        else {
                            if ($plan) { $plan.Dispose(); [void]$activePlans.Remove($plan) }
                            $context.Plan = $null
                            $context.PendingPaths.Clear()
                            $context.EnumerationComplete = $true
                            $ranges.Push([PSCustomObject]@{ Phase = 'Finalize'; Start = 0; Count = 1; Context = $context })
                        }
                    }
                    else {
                        # Workers finish out of order; only final profile results occupy the original index.
                        $results[$job.Index] = $result
                        $completed++
                    }
                }

                # Create replacements only after collection frees a slot, bounding queued and running instances.
                while ($ranges.Count -gt 0 -and $jobs.Count -lt $maxThreadsEffective) {
                    $range = $ranges.Pop()
                    while ($range.Count -gt 1) {
                        $leftCount = [int][math]::Floor($range.Count / 2)
                        $ranges.Push([PSCustomObject]@{
                                Phase = $range.Phase; Start = $range.Start + $leftCount
                                Count = $range.Count - $leftCount; Context = $range.Context
                            })
                        $range = [PSCustomObject]@{
                            Phase = $range.Phase; Start = $range.Start; Count = $leftCount; Context = $range.Context
                        }
                    }

                    $folder = $null
                    $startTime = [datetime]::MinValue
                    if ($range.Phase -eq 'Enumerate') {
                        $context = $range.Context
                        if ($context.PendingPaths.Count -gt 0) {
                            $path = $context.PendingPaths.Dequeue()
                        }
                        elseif ($context.Plan.MoveNext()) {
                            $path = $context.Plan.Current
                        }
                        else {
                            $context.Plan.Dispose()
                            [void]$activePlans.Remove($context.Plan)
                            $context.Plan = $null
                            $context.EnumerationComplete = $true
                            if ($context.Remaining -gt 0) { continue }
                            $range.Phase = 'Finalize'
                        }
                        if (-not $context.EnumerationComplete) {
                            # Retain one continuation per profile and admit only the next disjoint leaf.
                            $ranges.Push($range)
                            $context.Remaining++
                            $range = [PSCustomObject]@{ Phase = 'Folder'; Start = 0; Count = 1; Context = $context }
                        }
                    }
                    if ($range.Phase -eq 'Prepare') {
                        $profileItem = $Profiles[$range.Start]
                        $profileIndex = $range.Start
                    }
                    else {
                        $profileIndex = $range.Context.Index
                        $startTime = $range.Context.Start
                        if ($range.Phase -eq 'Folder') {
                            $profileItem = [PSCustomObject]@{ LocalPath = $path; SID = $null }
                            $folder = [PSCustomObject]@{ Name = [System.IO.Path]::GetFileName($path); Path = $path }
                        }
                        else {
                            $profileItem = $range.Context.Profile
                        }
                    }

                    $ps = [PowerShell]::Create()
                    $pendingPowerShell = $ps
                    $ps.RunspacePool = $pool

                    [void]$ps.AddScript($scriptBlock, $true).
                    AddArgument($profileItem).
                    AddArgument($range.Phase).
                    AddArgument($startTime).
                    AddArgument($(if ($range.Context) { $range.Context.Profile } else { $null })).
                    AddArgument($(if ($range.Context) { [bool]$range.Context.CheckProfileState } else { $true }))

                    $handle = $ps.BeginInvoke()
                    $pendingHandle = $handle
                    $pendingWaitHandle = $handle.AsyncWaitHandle

                    $jobs.Add([PSCustomObject]@{
                            PowerShell = $ps
                            Handle     = $handle
                            WaitHandle = $pendingWaitHandle
                            Profile    = $profileItem
                            Folder     = $folder
                            Index      = $profileIndex
                            Phase      = $range.Phase
                            Context    = $range.Context
                        })
                    # Transfer cleanup ownership only after the new job has been registered successfully.
                    $pendingPowerShell = $null
                    $pendingHandle = $null
                    $pendingWaitHandle = $null
                }

                $percent = if ($total -gt 0) { [math]::Floor(($completed / $total) * 100) } else { 100 }

                # Update the console only when the whole-number percentage changes.
                if ($percent -ne $lastPercent) {
                    $lastPercent = $percent
                    Write-ProgressUpdate -Activity (Get-SourceTextLoc 'toolText.extra.removingRegisteredProfiles') -Status (Get-SourceTextLoc 'toolText.extra.01Completed' -Args @($completed, $total)) -Percent $percent -Icon '🗑️'
                }

                if ($jobs.Count -gt 0) {
                    # Wait briefly for worker completion instead of continuously polling.
                    $waitHandles = [System.Threading.WaitHandle[]]@($jobs | ForEach-Object { $_.WaitHandle })
                    [void][System.Threading.WaitHandle]::WaitAny($waitHandles, 500)
                    $waitHandles = $null
                }
            }

            Write-ProgressUpdate -Activity (Get-SourceTextLoc 'toolText.extra.removingRegisteredProfiles') -Status (Get-SourceTextLoc 'uiText.completed') -Percent 100 -Icon '✅'
            Clear-ProgressLine

            return $results
        }
        finally {
            # Failure paths can bypass collection, leaving pending or tracked instances for this cleanup.
            try {
                if ($pendingPowerShell) {
                    Close-ProfileRemovalPowerShell -PowerShell $pendingPowerShell -Handle $pendingHandle -WaitHandle $pendingWaitHandle
                }
                foreach ($job in $jobs) {
                    Close-ProfileRemovalPowerShell -PowerShell $job.PowerShell -Handle $job.Handle -WaitHandle $job.WaitHandle
                }
                $jobs.Clear()
            }
            finally {
                try {
                    foreach ($plan in $activePlans) {
                        try { $plan.Dispose() }
                        catch { Write-Warning -Message $_.Exception.Message -WarningAction Continue }
                    }
                }
                finally {
                    $activePlans.Clear()
                    # Release the pool even after instance or iterator cleanup fails.
                    if ($pool) {
                        try { $pool.Close() }
                        finally { $pool.Dispose() }
                    }
                }
            }
        }
    }

    # Select leftovers from balanced index ranges while preserving name, CIM-path, SID, and link exclusions.
    function Get-ResidualUserFolders {
        $excluded = New-ProtectedNameSet
        # Refresh registration after profile removal so this scan reflects the state Windows now reports.
        $registeredProfilePaths = Get-RegisteredProfilePathSet

        Write-StyledMessage -Type 'Info' -Text ("🔎 " + (Get-SourceTextLoc 'toolText.checkResidualFoldersInTheUsersDirectory'))
        Write-ToolkitLog -Level 'INFO' -Message (Get-SourceTextLoc 'toolText.checkResidualFoldersInTheUsersDirectory')

        # Exclude reparse points so residual cleanup does not select links into other locations.
        $folders = @(Get-ChildItem -LiteralPath $usersRoot -Directory -Force |
        Where-Object {
            -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint)
        })

        # Registry paths protect ordinary username folders even when CIM omits a registered profile.
        $profileListKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
        $registeredSids = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        try {
            # Constant-time lookups avoid scanning the entire SID list for every folder.
            Get-ChildItem -LiteralPath $profileListKey -ErrorAction Stop |
            ForEach-Object {
                [void]$registeredSids.Add($_.PSChildName)
                $registeredPath = (Get-ItemProperty -LiteralPath $_.PSPath -Name 'ProfileImagePath' -ErrorAction Stop).ProfileImagePath
                if ([string]::IsNullOrWhiteSpace($registeredPath)) {
                    throw [System.InvalidOperationException]::new("Profile registration '$($_.PSChildName)' has no profile path.")
                }
                $registeredPath = [System.IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($registeredPath)).TrimEnd('\')
                if ($registeredPath.StartsWith($usersRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
                    [void]$registeredProfilePaths.Add($registeredPath)
                }
            }
        }
        catch {
            Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.couldNotEnumerateRegistryProfileList0' -Args @($($_.Exception.Message)))
            throw
        }

        # Divide index ranges without copying folders; the cheap checks stay in this runspace to avoid pool overhead.
        $ranges = [System.Collections.Generic.Stack[object]]::new()
        if ($folders.Count -gt 0) {
            $ranges.Push([PSCustomObject]@{ Start = 0; Count = $folders.Count })
        }
        while ($ranges.Count -gt 0) {
            $range = $ranges.Pop()
            if ($range.Count -gt 256) {
                $leftCount = [int][math]::Floor($range.Count / 2)
                # Push right first so leaves combine in the same order as filesystem discovery.
                $ranges.Push([PSCustomObject]@{ Start = $range.Start + $leftCount; Count = $range.Count - $leftCount })
                $ranges.Push([PSCustomObject]@{ Start = $range.Start; Count = $leftCount })
                continue
            }

            for ($index = $range.Start; $index -lt $range.Start + $range.Count; $index++) {
                $folder = $folders[$index]
                $folderName = $folder.Name
                $folderPath = [System.IO.Path]::GetFullPath($folder.FullName).TrimEnd('\')

                if ($excluded.Contains($folderName)) {
                    Write-ToolkitLog -Level 'INFO' -Message (Get-SourceTextLoc 'toolText.residualFolderExcludedForProtectedName01' -Args @($folderName, $folderPath))
                    continue
                }

                if ($registeredProfilePaths.Contains($folderPath)) {
                    Write-ToolkitLog -Level 'INFO' -Message (Get-SourceTextLoc 'toolText.residualFolderExcludedBecauseItIsStillAssociatedWithWin32Userprofile01' -Args @($folderName, $folderPath))
                    continue
                }

                if ($registeredSids.Contains($folderName)) {
                    Write-ToolkitLog -Level 'INFO' -Message (Get-SourceTextLoc 'toolText.residualFolderExcludedBecauseSidStillRegistered01' -Args @($folderName, $folderPath))
                    continue
                }

                # Keep the link exclusion at candidate selection as well as during initial enumeration.
                if ($folder.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                    Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.residualFolderExcludedBecauseReparsePointSymlink01' -Args @($folderName, $folderPath))
                    continue
                }

                [PSCustomObject]@{
                    Name = $folderName
                    Path = $folderPath
                }
            }
        }
    }

    # Delete one residual folder as an independent leaf task, preserving literal paths and ACL recovery.
    function Remove-ResidualUserFolder {
        param(
            [Parameter(Mandatory = $true)]
            [object]$Folder
        )

        $start = Get-Date
        $success = $false

        Write-ToolkitLog -Level 'INFO' -Message (Get-SourceTextLoc 'toolText.startResidualFolder01' -Args @($($folder.Name), $($folder.Path)))

        $folderPath = $folder.Path

        try {
            # Literal paths preserve folder names containing wildcard characters during deletion.
            if (Test-ProfileCleanupDirectory -Path $folderPath) {
                Remove-Item -LiteralPath $folderPath -Force -Recurse -ErrorAction Stop -Confirm:$false
            }
            # A completed command is insufficient; verify that the directory is actually gone.
            $success = -not (Test-ProfileCleanupDirectory -Path $folderPath)
        }
        catch {
            Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.standardResidualFolderRemovalFailed01' -Args @($folderPath, $($_.Exception.Message)))

            # Retry with ownership and access repaired when normal filesystem deletion fails.
            try {
                if (Test-ProfileCleanupDirectory -Path $folderPath) {
                    Invoke-ProfileCleanupCommand -Command 'takeown.exe' -Arguments @('/F', "`"$folderPath`"", '/R', '/D', 'Y', '/SKIPSL') -LogContextKey "ResidualTakeOwn-$($folder.Name)" | Out-Null
                }
                # S-1-5-32-544 - Fixed SID across Windows languages; avoids depending on localized group names.
                if (Test-ProfileCleanupDirectory -Path $folderPath) {
                    Invoke-ProfileCleanupCommand -Command 'icacls.exe' -Arguments @("`"$folderPath`"", '/grant', '*S-1-5-32-544:F', '/T', '/C', '/L') -LogContextKey "ResidualIcacls-$($folder.Name)" | Out-Null
                }
                if (Test-ProfileCleanupDirectory -Path $folderPath) {
                    Remove-Item -LiteralPath $folderPath -Recurse -Force -ErrorAction Stop -Confirm:$false
                }
                $success = -not (Test-ProfileCleanupDirectory -Path $folderPath)
            }
            catch {
                Write-ToolkitLog -Level 'ERROR' -Message (Get-SourceTextLoc 'toolText.remnantFolderRemovalFailed01' -Args @($folderPath, $($_.Exception.Message)))
                $success = $false
            }
        }

        $duration = New-TimeSpan -Start $start -End (Get-Date)

        if ($success) {
            Write-ToolkitLog -Level 'SUCCESS' -Message (Get-SourceTextLoc 'toolText.completedResidualFolder01' -Args @($($folder.Name), $duration.ToString()))
        }
        else {
            Write-ToolkitLog -Level 'ERROR' -Message (Get-SourceTextLoc 'toolText.failedResidualFolder0' -Args @($($folder.Name)))
        }

        return [PSCustomObject]@{
            Type     = 'ResidualFolder'
            UserName = $folder.Name
            Path     = $folder.Path
            Success  = $success
            Duration = $duration
        }
    }

    # Divide the folder list into independent leaves, conquer them in one bounded pool, then combine ordered results.
    function Remove-ResidualUserFolders {
        param(
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [array]$Folders
        )

        $total = $Folders.Count
        $results = [object[]]::new($total)
        $jobs = [System.Collections.Generic.List[object]]::new()
        $pendingPowerShell = $null
        $pendingHandle = $null
        $pendingWaitHandle = $null
        $pool = $null
        $completed = 0
        $lastPercent = -1

        # Index ranges avoid copying folder arrays or creating a pipeline for every folder upfront.
        # Depth-first bisection retains only a logarithmic number of pending ranges.
        $ranges = [System.Collections.Generic.Stack[object]]::new()
        if ($total -gt 0) {
            $ranges.Push([PSCustomObject]@{ Start = 0; Count = $total })
        }

        $scriptBlock = {
            param($Folder)
            $ErrorActionPreference = 'Stop'
            Remove-ResidualUserFolder -Folder $Folder
        }

        try {
            if ($total -gt 0) {
                $sessionState = New-ProfileRemovalSessionState
                $pool = [RunspaceFactory]::CreateRunspacePool(1, $maxThreadsEffective, $sessionState, $Host)
                $pool.Open()
            }

            while ($ranges.Count -gt 0 -or $jobs.Count -gt 0) {
                # Reclaim completed pipelines before admitting another leaf task.
                for ($jobIndex = $jobs.Count - 1; $jobIndex -ge 0; $jobIndex--) {
                    $job = $jobs[$jobIndex]
                    if (-not $job.Handle.IsCompleted) { continue }

                    $results[$job.Index] = Receive-ProfileRemovalResult -Job $job -ResultType 'ResidualFolder'
                    $jobs.RemoveAt($jobIndex)
                    $completed++
                }

                while ($ranges.Count -gt 0 -and $jobs.Count -lt $maxThreadsEffective) {
                    $range = $ranges.Pop()
                    while ($range.Count -gt 1) {
                        $leftCount = [int][math]::Floor($range.Count / 2)
                        $ranges.Push([PSCustomObject]@{
                                Start = $range.Start + $leftCount
                                Count = $range.Count - $leftCount
                            })
                        $range = [PSCustomObject]@{ Start = $range.Start; Count = $leftCount }
                    }

                    # Each leaf owns one selected folder; completed leaves free slots immediately.
                    $folder = $Folders[$range.Start]
                    $ps = [PowerShell]::Create()
                    $pendingPowerShell = $ps
                    $ps.RunspacePool = $pool
                    [void]$ps.AddScript($scriptBlock, $true).AddArgument($folder)
                    $handle = $ps.BeginInvoke()
                    $pendingHandle = $handle
                    $pendingWaitHandle = $handle.AsyncWaitHandle
                    $jobs.Add([PSCustomObject]@{
                            PowerShell = $ps
                            Handle     = $handle
                            WaitHandle = $pendingWaitHandle
                            Folder     = $folder
                            Index      = $range.Start
                        })
                    $pendingPowerShell = $null
                    $pendingHandle = $null
                    $pendingWaitHandle = $null
                }

                # Only the coordinator renders progress, using completed work rather than scheduled work.
                $percent = [math]::Floor(($completed / $total) * 100)
                if ($percent -ne $lastPercent) {
                    $lastPercent = $percent
                    Write-ProgressUpdate -Activity (Get-SourceTextLoc 'toolText.extra.removingResidualFoldersInCUsers') -Status ("{0} / {1}" -f $completed, $total) -Percent $percent -Icon '🗑️'
                }

                if ($jobs.Count -gt 0) {
                    $waitHandles = [System.Threading.WaitHandle[]]@($jobs | ForEach-Object { $_.WaitHandle })
                    [void][System.Threading.WaitHandle]::WaitAny($waitHandles, 500)
                    $waitHandles = $null
                }
            }

            Write-ProgressUpdate -Activity (Get-SourceTextLoc 'toolText.extra.removingResidualFoldersInCUsers') -Status (Get-SourceTextLoc 'uiText.completed') -Percent 100 -Icon '✅'
            Clear-ProgressLine

            # Original indices combine the independently completed leaves into a stable result sequence.
            return $results
        }
        finally {
            try {
                if ($pendingPowerShell) {
                    Close-ProfileRemovalPowerShell -PowerShell $pendingPowerShell -Handle $pendingHandle -WaitHandle $pendingWaitHandle
                }
                foreach ($job in $jobs) {
                    Close-ProfileRemovalPowerShell -PowerShell $job.PowerShell -Handle $job.Handle -WaitHandle $job.WaitHandle
                }
                $jobs.Clear()
            }
            finally {
                if ($pool) {
                    try { $pool.Close() }
                    finally { $pool.Dispose() }
                }
            }
        }
    }

    # Both cleanup phases share the toolkit's progress, localization, and log context.
    Start-ToolkitSession -ToolName $script:ToolName -SubTitle (Get-SourceTextLoc 'script.WinDeleteUserProfiles')

    try {
        # Validate the configured directory before either cleanup phase can delete anything.
        if (-not (Test-Path -LiteralPath $usersRoot -PathType Container)) {
            throw (Get-SourceTextLoc 'toolText.extra.profilePathDoesNotExist0' -Args @($usersRoot))
        }

        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
        if ($os -and $os.Caption -notmatch 'Windows 11') {
            Write-StyledMessage -Type 'Warning' -Text ((Get-SourceTextLoc 'toolText.systemDetected0TheScriptIsDesignedForWindows11' -Args @($($os.Caption))))
        }

        Write-StyledMessage -Type 'Info' -Text ("🖥️ " + (Get-SourceTextLoc 'toolText.computer0' -Args @($computerName)))
        Write-StyledMessage -Type 'Info' -Text ("👤 " + (Get-SourceTextLoc 'toolText.currentUserProtected0' -Args @($currentUser)))
        Write-StyledMessage -Type 'Info' -Text ("📁 " + (Get-SourceTextLoc 'toolText.profilePath0' -Args @($usersRoot)))
        Write-StyledMessage -Type 'Info' -Text ("🧵 " + (Get-SourceTextLoc 'toolText.threadsConfigured0' -Args @($maxThreadsEffective)))
        Write-StyledMessage -Type 'Warning' -Text (Get-SourceTextLoc 'toolText.nonInteractiveModeNoConfirmationWillBeRequestedBeforeCancellations')

        if ($minimumLastUseDate) {
            Write-StyledMessage -Type 'Info' -Text ("📅 " + (Get-SourceTextLoc 'toolText.lastActivityThresholdProfilesNotUsedForAtLeast0Days' -Args @($MinimumProfileAgeDays)))
        }

        Write-ToolkitLog -Level 'INFO' -Message (Get-SourceTextLoc 'toolText.sessionStartedOn0' -Args @($computerName))

        $profileResults = @()
        $residualResults = @()

        $targets = @(Get-RemovableUserProfiles)

        if ($targets -and $targets.Count -gt 0) {
            Write-Host ''
            Write-StyledMessage -Type 'Warning' -Text (Get-SourceTextLoc 'toolText.registeredProfilesSelectedForAutomaticRemoval')
            Write-Host ''

            $targets |
            Select-Object @{Name = 'User'; Expression = { [System.IO.Path]::GetFileName($_.LocalPath) } },
            @{Name = 'Loaded'; Expression = { $_.Loaded } },
            @{Name = 'LastUseTime'; Expression = { $_.LastUseTime } },
            @{Name = 'Path'; Expression = { $_.LocalPath } } |
            Format-Table -AutoSize

            Write-Host ''
            Write-StyledMessage -Type 'Info' -Text ("🚀 " + (Get-SourceTextLoc 'toolText.startAutomaticRemovalOf0RegisteredProfiles' -Args @($targets.Count)))
            Write-Host ''

            $profileResults = @(Invoke-ProfileRemovalBatch -Profiles $targets)
        }
        else {
            Write-Host ''
            Write-StyledMessage -Type 'Success' -Text ((Get-SourceTextLoc 'toolText.noRemovableRegisteredProfilesFound'))
        }

        # Scan leftovers after registered-profile removal using the updated registration state.
        if (-not $SkipResidualFolderCleanup) {
            $residualFolders = @(Get-ResidualUserFolders)

            if ($residualFolders -and $residualFolders.Count -gt 0) {
                Write-Host ''
                Write-StyledMessage -Type 'Warning' -Text (Get-SourceTextLoc 'toolText.residualFoldersSelectedForAutomaticRemoval0' -Args @($residualFolders.Count))
                $residualFolders | Select-Object Name, Path | Format-Table -AutoSize

                Write-Host ''
                Write-StyledMessage -Type 'Info' -Text ("🧹 " + (Get-SourceTextLoc 'toolText.startingRemovalOfResidualFolders'))
                Write-Host ''

                $residualResults = @(Remove-ResidualUserFolders -Folders $residualFolders)
            }
            else {
                Write-StyledMessage -Type 'Success' -Text ((Get-SourceTextLoc 'toolText.noRemovableResidualFolderFoundInCUsers'))
            }
        }
        else {
            Write-StyledMessage -Type 'Warning' -Text (Get-SourceTextLoc 'toolText.residualFolderCleanupSkippedForSkipresidualfoldercleanupParameter')
        }

        # Report successes separately for each phase, but base restart advice on failures from either phase.
        $allResults = @($profileResults) + @($residualResults)
        $failedCount = @($allResults | Where-Object { -not $_.Success }).Count
        $profileSuccessCount = @($profileResults | Where-Object { $_.Success }).Count
        $residualSuccessCount = @($residualResults | Where-Object { $_.Success }).Count

        # Open sessions or handles may prevent cleanup; recommend a restart only when items failed.
        if ($failedCount -gt 0) {
            $rebootRecommended = $true
            Write-StyledMessage -Type 'Warning' -Text ((Get-SourceTextLoc 'toolText.extra.restartRecommendedAfterProfileCleanup'))
            Write-ToolkitLog -Level 'WARNING' -Message (Get-SourceTextLoc 'toolText.recommendedReboot0' -Args @((Get-SourceTextLoc 'toolText.0ItemsNotRemovedMayBeBlockedByOpenSessionsOrHandles' -Args @($failedCount))))
        }

        Write-StyledMessage -Type 'Success' -Text ((Get-SourceTextLoc 'toolText.registeredProfilesRemoved0' -Args @($profileSuccessCount)))
        Write-StyledMessage -Type 'Success' -Text ((Get-SourceTextLoc 'toolText.residualFoldersRemoved0' -Args @($residualSuccessCount)))
        if ($failedCount -gt 0) {
            Write-StyledMessage -Type 'Warning' -Text ((Get-SourceTextLoc 'toolText.itemsNotRemoved0' -Args @($failedCount)))
        }

        if ($rebootRecommended) {
            # Forward suppression so a toolkit batch can manage its final restart centrally.
            Invoke-ToolkitReboot -Message (Get-SourceTextLoc 'toolText.extra.restartRecommendedAfterProfileCleanup') -Seconds $CountdownSeconds -SuppressIndividualReboot:$SuppressIndividualReboot
        }
    }
    catch {
        Write-ToolkitError -Record $_ -ToolName $script:ToolName
        # Preserve the original exception so the caller can detect an unsuccessful run.
        throw
    }
    finally {
        # Report that the session ended even when a cleanup phase throws.
        Write-StyledMessage -Type 'Info' -Text ("♻️ " + (Get-SourceTextLoc 'toolText.winDeleteUserProfilesSessionEnded'))
        Write-ToolkitLog -Level INFO -Message (Get-SourceTextLoc 'toolText.winDeleteUserProfilesSessionEnded')
    }
}
