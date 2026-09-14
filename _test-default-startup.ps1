$root = Join-Path $env:TEMP 'md-browser-default-startup'
$oneDrive = Join-Path $env:TEMP 'md-browser-default-onedrive'
$harnessPath = Join-Path $env:TEMP 'md-browser-default-startup-harness.ps1'

try {
	Remove-Item $root, $oneDrive -Recurse -Force -ErrorAction SilentlyContinue
	New-Item $root, $oneDrive -ItemType Directory | Out-Null
	Set-Content (Join-Path $root 'first.md') '# First' -Encoding UTF8

	$source = Get-Content (Join-Path $PSScriptRoot 'MD-Browser.ps1') -Raw
	$harness = @'
if ($script:RootPath -ne (Resolve-Path $HOME).Path) { throw "Unexpected root: $script:RootPath" }
if ($Tree.Items.Count -ne 1) { throw "Unexpected tree root count: $($Tree.Items.Count)" }
"Default startup test PASS"
'@
	[System.IO.File]::WriteAllText($harnessPath, $source.Replace('$win.ShowDialog() | Out-Null', $harness))

	$oldHome = $env:HOME
	$oldOneDrive = $env:OneDriveCommercial
	$env:HOME = $root
	$env:OneDriveCommercial = $oneDrive
	try {
		& powershell.exe -STA -NoProfile -File $harnessPath
		if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
	} finally {
		$env:HOME = $oldHome
		$env:OneDriveCommercial = $oldOneDrive
	}
} finally {
	Remove-Item $harnessPath -Force -ErrorAction SilentlyContinue
	Remove-Item $root, $oneDrive -Recurse -Force -ErrorAction SilentlyContinue
}
