$output = & (Join-Path $PSScriptRoot 'MD-Browser.ps1') -Version
if ($output -ne 'MD-Browser 1.0.0') {
    throw "Unexpected version output: $output"
}

'Version test PASS'