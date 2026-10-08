# WinToolkit CI/CD V4.2.0
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
Compile inert tools in TestDrive and invoke their entry points as the menu does.
Never load or invoke the real toolkit or Windows profile cleanup implementation.
#>

Describe 'Compiler tool entry points' {
    BeforeAll {
        $repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
        $buildRoot = Join-Path $TestDrive 'compiler-entry-points'
        $toolRoot = Join-Path $buildRoot 'tools'
        $moduleRoot = Join-Path $buildRoot 'wintoolkit-modules'
        New-Item -ItemType Directory -Path $toolRoot, $moduleRoot | Out-Null
        Copy-Item -LiteralPath (Join-Path $repoRoot 'compiler.ps1') -Destination $buildRoot
        Copy-Item -LiteralPath (Join-Path $repoRoot 'languages') -Destination $buildRoot -Recurse

        $toolSources = @{
            WinDeleteUserProfiles = @'
# A leading comment must not hide the public entry point from the compiler.
function WinDeleteUserProfiles {
    [CmdletBinding()]
    param([string]$Marker = 'default', [switch]$SuppressIndividualReboot)
    function Get-ProbeMarker { param([string]$Value) return $Value }
    Start-ToolkitSession -ToolName 'WinDeleteUserProfiles'
    Get-ProbeMarker -Value $Marker
}
WinDeleteUserProfiles
'@
            SyntheticBlockTool = @'
<#
A block comment may include braces { } and function-like text.
#>
function SyntheticBlockTool
{
    param([string]$Marker = 'default', [switch]$SuppressIndividualReboot)
    Start-ToolkitSession -ToolName 'SyntheticBlockTool'
    $Marker
}
# Keep comments following the function as well.
'@
            SyntheticInlineTool = @'
function SyntheticInlineTool { param([string]$Marker = '{default}', [switch]$SuppressIndividualReboot) Start-ToolkitSession -ToolName 'SyntheticInlineTool'; $Marker }
# A closing brace need not occupy the last nonempty line.
'@
            SyntheticLegacyTool = @'
function SyntheticLegacyTool {
    param([string]$Marker = 'default', [switch]$SuppressIndividualReboot)
    Start-ToolkitSession -ToolName 'SyntheticLegacyTool'
    $Marker
}
SyntheticLegacyTool
'@
            SyntheticBodyTool = @'
param([string]$Marker = 'default', [switch]$SuppressIndividualReboot)
Start-ToolkitSession -ToolName 'SyntheticBodyTool'
$Marker
'@
        }
        foreach ($entry in $toolSources.GetEnumerator()) {
            Set-Content -LiteralPath (Join-Path $toolRoot ($entry.Key + '.ps1')) -Value $entry.Value -Encoding UTF8
        }
        Set-Content -LiteralPath (Join-Path $moduleRoot '00-Probe.ps1') -Encoding UTF8 -Value @'
function Start-ToolkitSession {
    param([string]$ToolName)
    $script:ProbeSessionCalls++
}
'@
        $placeholders = $toolSources.Keys | Sort-Object | ForEach-Object { 'function ' + $_ + ' {}' }
        Set-Content -LiteralPath (Join-Path $moduleRoot '87-Placeholders.ps1') -Value $placeholders -Encoding UTF8

        $script:CompiledSources = @{}
        foreach ($variant in @('full', 'minified')) {
            & (Join-Path $buildRoot 'compiler.ps1') -Minify:($variant -eq 'minified') *> (Join-Path $buildRoot ($variant + '.log'))
            if ($LASTEXITCODE -ne 0) { throw "Synthetic $variant compilation failed; see $buildRoot." }
            $script:CompiledSources[$variant] = Get-Content -Raw -LiteralPath (Join-Path $buildRoot 'WinToolkit.ps1') -Encoding UTF8
        }
    }

    It '<Variant>: <FunctionName> retains one parameterized entry point' -ForEach @(
        foreach ($variant in @('full', 'minified')) {
            foreach ($name in @('WinDeleteUserProfiles', 'SyntheticBlockTool', 'SyntheticInlineTool', 'SyntheticLegacyTool', 'SyntheticBodyTool')) {
                @{ Variant = $variant; FunctionName = $name }
            }
        }
    ) {
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($script:CompiledSources[$Variant], [ref]$null, [ref]$errors)
        $errors.Count | Should -Be 0
        $entries = @($ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $FunctionName
        }, $true))
        $entries.Count | Should -Be 1 -Because 'a nested function definition returns without executing the tool'
        $entries[0].Body.ParamBlock | Should -Not -BeNullOrEmpty
    }

    It '<Variant>: menu dispatch executes <FunctionName> once with its arguments' -ForEach @(
        foreach ($variant in @('full', 'minified')) {
            foreach ($name in @('WinDeleteUserProfiles', 'SyntheticBlockTool', 'SyntheticInlineTool', 'SyntheticLegacyTool', 'SyntheticBodyTool')) {
                @{ Variant = $variant; FunctionName = $name }
            }
        }
    ) {
        $probe = & {
            param([string]$Source, [string]$Name)
            # These definitions come only from the inert fixtures above.
            . ([scriptblock]::Create($Source))
            $script:ProbeSessionCalls = 0
            $result = & $ExecutionContext.InvokeCommand.GetCommand($Name, 'Function') -Marker 'probe {literal}' -SuppressIndividualReboot
            [pscustomobject]@{ Result = $result; SessionCalls = $script:ProbeSessionCalls }
        } $script:CompiledSources[$Variant] $FunctionName
        $probe.Result | Should -Be 'probe {literal}'
        $probe.SessionCalls | Should -Be 1
    }
}
