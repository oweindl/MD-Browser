$root = Join-Path $env:TEMP 'md-browser-search-navigation'
$localAppData = Join-Path $env:TEMP 'md-browser-search-navigation-appdata'
$harness = Join-Path $env:TEMP 'md-browser-search-navigation-harness.ps1'

try {
    Remove-Item $root, $localAppData -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $root, $localAppData | Out-Null
    $lines = @('# Search test', 'needle near the top')
    $lines += 1..160 | ForEach-Object { "Filler line $_" }
    $lines += 'needle near the bottom'
    [System.IO.File]::WriteAllLines((Join-Path $root 'first.md'), $lines)

    $source = Get-Content (Join-Path $PSScriptRoot 'MD-Browser.ps1') -Raw
    $testBody = @'
$script:TestStage = 0
$script:TestResult = $null
$script:TestBusyWorked = $false
$script:TestCheckTimer = $null
$script:TestTimeout = New-Object System.Windows.Threading.DispatcherTimer
$script:TestTimeout.Interval = [TimeSpan]::FromSeconds(8)
$script:TestTimeout.Add_Tick({
    $script:TestResult = 'Timed out waiting for preview navigation'
    $script:TestTimeout.Stop()
    $win.Close()
})
$script:TestTimeout.Start()

$Preview.Add_LoadCompleted({
    if ($script:TestStage -eq 0) {
        $script:TestStage = 1
        $TxtFilter.Text = 'needle'
        $script:FilterTimer.Stop()
        Invoke-TreeFilter
        return
    }
    if ($script:TestStage -ne 1) { return }

    $script:TestStage = 2
    Move-Hit 1
    Move-Hit 1
    Start-SearchBusy -Term 'needle'
    $firstFrame = $SearchBusyPulse.Text
    $script:SearchLastYield = [datetime]::MinValue
    Update-SearchBusy
    $script:TestBusyWorked = $SearchBusy.Visibility -eq 'Visible' -and $SearchBusyPulse.Text -ne $firstFrame
    Stop-SearchBusy

    $script:TestCheckTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:TestCheckTimer.Interval = [TimeSpan]::FromMilliseconds(250)
    $script:TestCheckTimer.Add_Tick({
        $script:TestCheckTimer.Stop()
        $element = $Preview.Document.getElementById('mdhit1')
        $rect = $element.getBoundingClientRect()
        $viewportHeight = $Preview.Document.documentElement.clientHeight
        $inView = $rect.top -ge 0 -and $rect.top -lt $viewportHeight
        $hidden = $SearchBusy.Visibility -eq 'Collapsed'
        if ($script:HitIndex -ne 2 -or -not $inView) {
            $script:TestResult = "Second hit was not visible (index=$script:HitIndex, top=$($rect.top), viewport=$viewportHeight)"
        } elseif (-not $script:TestBusyWorked -or -not $hidden) {
            $script:TestResult = 'Search busy indicator did not animate or hide correctly'
        } else {
            $script:TestResult = 'PASS'
        }
        $script:TestTimeout.Stop()
        $win.Close()
    })
    $script:TestCheckTimer.Start()
})

$win.ShowDialog() | Out-Null
if ($script:TestResult -ne 'PASS') { throw $script:TestResult }
'@

    [System.IO.File]::WriteAllText($harness, $source.Replace('$win.ShowDialog() | Out-Null', $testBody))
    $oldLocalAppData = $env:LOCALAPPDATA
    $env:LOCALAPPDATA = $localAppData
    try {
        & powershell.exe -STA -NoProfile -File $harness -Path $root
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    } finally {
        $env:LOCALAPPDATA = $oldLocalAppData
    }
    'Search navigation test PASS'
} finally {
    Remove-Item -LiteralPath $harness -Force -ErrorAction SilentlyContinue
    Remove-Item $root, $localAppData -Recurse -Force -ErrorAction SilentlyContinue
}