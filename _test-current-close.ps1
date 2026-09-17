$src = Get-Content "$PSScriptRoot\MD-Browser.ps1" -Raw
$replacement = '$win.Add_Closing({"CLOSING"|Out-Host});$x=$win.ShowDialog();"RESULT=$x"'
$harness = "$PSScriptRoot\_close-harness.ps1"
$root = Join-Path $env:TEMP 'md-browser-close'
$localAppData = Join-Path $env:TEMP 'md-browser-close-appdata'

try {
	Remove-Item $root, $localAppData -Recurse -Force -ErrorAction SilentlyContinue
	New-Item $root, $localAppData -ItemType Directory | Out-Null
	Set-Content $harness ($src.Replace('$win.ShowDialog() | Out-Null', $replacement))
	Set-Content "$root\first.md" '# First'

	$oldLocalAppData = $env:LOCALAPPDATA
	$env:LOCALAPPDATA = $localAppData
	try {
		powershell.exe -STA -NoProfile -File $harness -Path $root
	} finally {
		$env:LOCALAPPDATA = $oldLocalAppData
	}
} finally {
	Remove-Item $harness -Force -ErrorAction SilentlyContinue
	Remove-Item $root, $localAppData -Recurse -Force -ErrorAction SilentlyContinue
}