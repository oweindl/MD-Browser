$root = Join-Path $env:TEMP 'md-browser-reveal-selection'
$localAppData = Join-Path $env:TEMP 'md-browser-reveal-selection-appdata'
$harnessPath = Join-Path $env:TEMP 'md-browser-reveal-selection-harness.ps1'

try {
    Remove-Item $root, $localAppData -Recurse -Force -ErrorAction SilentlyContinue
    New-Item $root, $localAppData -ItemType Directory | Out-Null
    $file = Join-Path $root 'first.md'
    Set-Content -LiteralPath $file -Value '# First' -Encoding UTF8

    $source = Get-Content (Join-Path $PSScriptRoot 'MD-Browser.ps1') -Raw
    $testBody = @'
$script:ExplorerLaunches = New-Object System.Collections.Generic.List[object]
function Start-Process {
    param([string]$FilePath, [string]$ArgumentList)
    $script:ExplorerLaunches.Add([pscustomobject]@{ FilePath = $FilePath; ArgumentList = $ArgumentList }) | Out-Null
}

$Tree = [pscustomobject]@{ SelectedItem = [pscustomobject]@{ Tag = @{ Type = 'dir'; Path = $script:RootPath } } }
Show-SelectedTreeItemInExplorer

$selectedFile = Join-Path $script:RootPath 'first.md'
$Tree.SelectedItem = [pscustomobject]@{ Tag = @{ Type = 'file'; Path = $selectedFile } }
Show-SelectedTreeItemInExplorer

if ($script:ExplorerLaunches.Count -ne 2) { throw "Unexpected launch count: $($script:ExplorerLaunches.Count)" }
if ($script:ExplorerLaunches[0].FilePath -ne 'explorer.exe' -or $script:ExplorerLaunches[0].ArgumentList -ne "`"$script:RootPath`"") {
    throw "Unexpected folder launch: $($script:ExplorerLaunches[0] | Out-String)"
}
if ($script:ExplorerLaunches[1].FilePath -ne 'explorer.exe' -or $script:ExplorerLaunches[1].ArgumentList -ne "/select,`"$selectedFile`"") {
    throw "Unexpected file launch: $($script:ExplorerLaunches[1] | Out-String)"
}
'@
    [System.IO.File]::WriteAllText($harnessPath, $source.Replace('$win.ShowDialog() | Out-Null', $testBody))

    $oldLocalAppData = $env:LOCALAPPDATA
    $env:LOCALAPPDATA = $localAppData
    try {
        & powershell.exe -STA -NoProfile -File $harnessPath -Path $root
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    } finally {
        $env:LOCALAPPDATA = $oldLocalAppData
    }

    'Reveal selection test PASS'
} finally {
    Remove-Item -LiteralPath $harnessPath -Force -ErrorAction SilentlyContinue
    Remove-Item $root, $localAppData -Recurse -Force -ErrorAction SilentlyContinue
}