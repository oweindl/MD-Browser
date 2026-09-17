$rootA = Join-Path $env:TEMP 'md-browser-remembered-a'
$rootB = Join-Path $env:TEMP 'md-browser-remembered-b'
$localAppData = Join-Path $env:TEMP 'md-browser-remembered-appdata'
$harnessPath = Join-Path $env:TEMP 'md-browser-remembered-path-harness.ps1'

try {
    Remove-Item $rootA, $rootB, $localAppData -Recurse -Force -ErrorAction SilentlyContinue
    New-Item $rootA, $rootB, $localAppData -ItemType Directory | Out-Null

    $source = Get-Content (Join-Path $PSScriptRoot 'MD-Browser.ps1') -Raw
    $harness = @'
$expected = (Resolve-Path -LiteralPath $env:MD_BROWSER_EXPECTED_ROOT).Path
if ($script:RootPath -ne $expected) { throw "Unexpected root: $script:RootPath (expected $expected)" }
$settings = Get-Content -LiteralPath $script:SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($settings.LastRoot -ne $expected) { throw "Unexpected remembered root: $($settings.LastRoot)" }
'@
    [System.IO.File]::WriteAllText($harnessPath, $source.Replace('$win.ShowDialog() | Out-Null', $harness))

    $oldLocalAppData = $env:LOCALAPPDATA
    $oldExpectedRoot = $env:MD_BROWSER_EXPECTED_ROOT
    $env:LOCALAPPDATA = $localAppData
    try {
        $env:MD_BROWSER_EXPECTED_ROOT = $rootA
        & powershell.exe -STA -NoProfile -File $harnessPath -Path $rootA
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

        & powershell.exe -STA -NoProfile -File $harnessPath
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

        $env:MD_BROWSER_EXPECTED_ROOT = $rootB
        & powershell.exe -STA -NoProfile -File $harnessPath -Path $rootB
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

        & powershell.exe -STA -NoProfile -File $harnessPath
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    } finally {
        $env:LOCALAPPDATA = $oldLocalAppData
        $env:MD_BROWSER_EXPECTED_ROOT = $oldExpectedRoot
    }

    'Remembered path test PASS'
} finally {
    Remove-Item -LiteralPath $harnessPath -Force -ErrorAction SilentlyContinue
    Remove-Item $rootA, $rootB, $localAppData -Recurse -Force -ErrorAction SilentlyContinue
}