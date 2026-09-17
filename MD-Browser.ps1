<#
.SYNOPSIS
    MD-Browser - a WPF GUI to browse, read and edit a folder structure of Markdown files.

.DESCRIPTION
    Left pane  : tree view of the selected folder structure (sub folders + .md files) with a filter box.
    Right pane : rendered preview of the selected .md file, switchable to a plain text editor (Ctrl+S saves).
    Links      : .md links that resolve inside the root folder select the target in the tree.
                 Local files outside the root are added to a "Standalone files" branch (full path in brackets).
                 Everything else (http/https/mailto) opens in the default application.

.NOTES
    No external dependencies. Requires Windows PowerShell 5.1 or PowerShell 7+ on Windows.
#>

[CmdletBinding()]
param(
    [string]$Path
)

# WPF needs a single threaded apartment - relaunch if we are not in one.
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $shell = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $shell)) { $shell = (Get-Process -Id $PID).Path }
    $argList = @('-STA', '-NoProfile', '-File', $PSCommandPath)
    if ($Path) { $argList += @('-Path', $Path) }
    Start-Process -FilePath $shell -ArgumentList $argList -WorkingDirectory (Split-Path -Parent $PSCommandPath) -WindowStyle Normal
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Web

#region ---------------------------------------------------------------- state

$script:RootPath        = $null
$script:CurrentFile     = $null
$script:CurrentDir      = $null
$script:IsDirty         = $false
$script:IsEditMode      = $false
$script:Loading         = $false
$script:Standalone      = New-Object System.Collections.Generic.List[string]
$script:SuppressSelect  = $false
$script:ContentCache    = @{}
$script:MatchCount      = 0
$script:FilterTimer     = $null
$script:PendingLink     = $null
$script:HitTotal        = 0
$script:HitIndex        = 0
$script:JumpPending     = $false
$script:History         = New-Object System.Collections.Generic.List[string]
$script:PendingAnchor   = $null
$script:AnchorTimer     = $null
$script:AnchorAttempts  = 0
$script:JumpDebug       = $false
$script:JumpDebugLines  = New-Object System.Collections.Generic.List[string]
$script:PendingJumpUrl   = $null
$script:PendingJumpPath  = $null
$script:PendingToggle   = $null
$script:IsSearching     = $false
$script:SearchPulse     = 0
$script:SearchLastYield = [datetime]::MinValue
$script:PreviewLoading  = $false
$script:PendingHitId    = $null
$script:HitScrollTimer  = $null
$script:HitScrollAttempts = 0

$script:MarkdownExt     = @('.md', '.markdown', '.mdown', '.mkd')
$script:LinkHost        = 'http://md-browser.local/'
$script:DefaultFolder   = 'Brain'
$script:ContentMinChars = 3
$script:SettingsPath    = Join-Path (Join-Path $env:LOCALAPPDATA 'MD-Browser') 'settings.json'

#endregion

#region ------------------------------------------------------------ markdown

function ConvertFrom-HtmlEntity {
    param([string]$Text)
    $Text -replace '&lt;', '<' -replace '&gt;', '>' -replace '&quot;', '"' -replace '&amp;', '&'
}

function Get-HeadingSlug {
    param([string]$Text)
    $s = ($Text -replace '<[^>]+>', '').ToLowerInvariant()
    $s = $s -replace '[^a-z0-9 _-]', ''
    ($s.Trim() -replace '\s+', '-')
}

function Resolve-MarkdownTarget {
    <#
        Classifies a raw link target from a markdown document.
        Returns Kind = anchor | external | local, plus the resolved absolute path for local targets.
    #>
    param(
        [string]$Target,
        [string]$BaseDir
    )

    $raw = ConvertFrom-HtmlEntity $Target
    $raw = $raw.Trim().Trim('<', '>')

    if ([string]::IsNullOrWhiteSpace($raw)) {
        return [pscustomobject]@{ Kind = 'external'; Raw = $raw; FullPath = $null; Fragment = $null }
    }
    if ($raw.StartsWith('#')) {
        return [pscustomobject]@{ Kind = 'anchor'; Raw = $raw; FullPath = $null; Fragment = $raw.Substring(1) }
    }
    if ($raw -match '^[a-zA-Z][a-zA-Z0-9+.-]*:' -and $raw -notmatch '^[a-zA-Z]:[\\/]') {
        if ($raw -match '^file:') {
            try {
                $u = [uri]$raw
                return [pscustomobject]@{ Kind = 'local'; Raw = $raw; FullPath = $u.LocalPath; Fragment = $u.Fragment.TrimStart('#') }
            } catch { }
        }
        return [pscustomobject]@{ Kind = 'external'; Raw = $raw; FullPath = $null; Fragment = $null }
    }

    $fragment = $null
    $pathPart = $raw
    $hash = $raw.IndexOf('#')
    if ($hash -ge 0) {
        $fragment = $raw.Substring($hash + 1)
        $pathPart = $raw.Substring(0, $hash)
    }
    if ([string]::IsNullOrWhiteSpace($pathPart)) {
        return [pscustomobject]@{ Kind = 'anchor'; Raw = $raw; FullPath = $null; Fragment = $fragment }
    }

    try { $pathPart = [uri]::UnescapeDataString($pathPart) } catch { }
    $pathPart = $pathPart -replace '/', '\'

    try {
        $full = if ([System.IO.Path]::IsPathRooted($pathPart)) { $pathPart }
                else { Join-Path $BaseDir $pathPart }
        $full = [System.IO.Path]::GetFullPath($full)
    } catch {
        return [pscustomobject]@{ Kind = 'external'; Raw = $raw; FullPath = $null; Fragment = $fragment }
    }

    [pscustomobject]@{ Kind = 'local'; Raw = $raw; FullPath = $full; Fragment = $fragment }
}

function Convert-MarkdownInline {
    param(
        [string]$Text,
        [string]$BaseDir
    )

    $marker = [char]1
    $codes  = New-Object System.Collections.Generic.List[string]

    # 1. protect inline code spans
    $Text = [regex]::Replace($Text, '(`+)([^`]|[^`][\s\S]*?[^`])\1(?!`)',
        [System.Text.RegularExpressions.MatchEvaluator] {
            param($m)
            $codes.Add($m.Groups[2].Value.Trim())
            "$marker$($codes.Count - 1)$marker"
        })

    # 2. images
    $Text = [regex]::Replace($Text, '!\[([^\]]*)\]\(\s*([^)\s]+)(?:\s+"[^"]*")?\s*\)',
        [System.Text.RegularExpressions.MatchEvaluator] {
            param($m)
            $alt = $m.Groups[1].Value
            $res = Resolve-MarkdownTarget -Target $m.Groups[2].Value -BaseDir $BaseDir
            $src = if ($res.Kind -eq 'local' -and $res.FullPath) {
                       ([uri]$res.FullPath).AbsoluteUri
                   } else { $res.Raw }
            '<img src="{0}" alt="{1}" />' -f $src, $alt
        })

    # 3. inline links -> routed through our pseudo host so we can intercept them
    $Text = [regex]::Replace($Text, '\[([^\]]*)\]\(\s*([^)\s]+)(?:\s+"[^"]*")?\s*\)',
        [System.Text.RegularExpressions.MatchEvaluator] {
            param($m)
            $label  = $m.Groups[1].Value
            $target = ConvertFrom-HtmlEntity $m.Groups[2].Value
            $tip = [System.Web.HttpUtility]::HtmlAttributeEncode($target)
            if ($target.StartsWith('#')) {
                return '<a href="{0}" title="{1}">{2}</a>' -f $tip, $tip, $label
            }
            if ($target.StartsWith($script:LinkHost, [StringComparison]::OrdinalIgnoreCase)) {
                return '<a href="{0}" title="{1}">{2}</a>' -f $tip, $tip, $label
            }
            '<a href="{0}?t={1}" title="{2}">{3}</a>' -f $script:LinkHost, [uri]::EscapeDataString($target), $tip, $label
        })

    # 4. autolinks  <https://...>  (angle brackets are already html escaped at this point)
    $Text = [regex]::Replace($Text, '&lt;((?:https?|ftp|mailto):[^\s&]+)&gt;',
        [System.Text.RegularExpressions.MatchEvaluator] {
            param($m)
            $u = $m.Groups[1].Value
            $tip = [System.Web.HttpUtility]::HtmlAttributeEncode($u)
            '<a href="{0}?t={1}" title="{2}">{3}</a>' -f $script:LinkHost, [uri]::EscapeDataString($u), $tip, $u
        })

    # 5. emphasis
    $Text = $Text -replace '\*\*(?=\S)([\s\S]+?)(?<=\S)\*\*', '<strong>$1</strong>'
    $Text = $Text -replace '(?<![\w])__(?=\S)([\s\S]+?)(?<=\S)__(?![\w])', '<strong>$1</strong>'
    $Text = $Text -replace '(?<![\*\w])\*(?=\S)([^\*]+?)(?<=\S)\*(?!\*)', '<em>$1</em>'
    $Text = $Text -replace '(?<![_\w])_(?=\S)([^_]+?)(?<=\S)_(?![_\w])', '<em>$1</em>'
    $Text = $Text -replace '~~(?=\S)([\s\S]+?)(?<=\S)~~', '<del>$1</del>'

    # 6. restore code spans
    for ($i = 0; $i -lt $codes.Count; $i++) {
        $Text = $Text.Replace("$marker$i$marker", '<code>' + $codes[$i] + '</code>')
    }

    $Text
}

function Add-HtmlHighlight {
    <# Wraps every occurrence of $Term in the text nodes of $Html and records the hit count. #>
    param([string]$Html, [string]$Term)

    $script:HitTotal = 0
    if ([string]::IsNullOrEmpty($Term)) { return $Html }

    $needle = $Term -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
    $sb = New-Object System.Text.StringBuilder
    $n  = 0

    foreach ($part in [regex]::Split($Html, '(<[^>]*>)')) {
        if ($part.Length -eq 0) { continue }
        if ($part[0] -eq '<') { [void]$sb.Append($part); continue }

        $i = 0
        while ($true) {
            $j = $part.IndexOf($needle, $i, [StringComparison]::OrdinalIgnoreCase)
            if ($j -lt 0) { [void]$sb.Append($part.Substring($i)); break }
            [void]$sb.Append($part.Substring($i, $j - $i))
            [void]$sb.Append(('<span class="mdhit" id="mdhit{0}">{1}</span>' -f $n, $part.Substring($j, $needle.Length)))
            $n++
            $i = $j + $needle.Length
        }
    }

    $script:HitTotal = $n
    $sb.ToString()
}

function Convert-MarkdownToHtml {
    param(
        [string]$Markdown,
        [string]$BaseDir,
        [string]$Highlight,
        [string]$TargetAnchor
    )

    # html-escape once up front, everything we emit afterwards is trusted markup
    $text = $Markdown -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
    $text = [regex]::Replace($text, '&lt;a\s+id="([^"]+)"\s*/?&gt;&lt;/a&gt;', '<a id="$1"></a>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $lines = $text -split "\r?\n"

    $out        = New-Object System.Collections.Generic.List[string]
    $para       = New-Object System.Collections.Generic.List[string]
    $listStack  = New-Object System.Collections.Generic.List[object]   # @{ Type; Indent }
    $inFence    = $false
    $fenceMark  = ''
    $fenceBuf   = New-Object System.Collections.Generic.List[string]
    $inQuote    = $false

    function Close-Lists {
        while ($listStack.Count -gt 0) {
            $top = $listStack[$listStack.Count - 1]
            $out.Add("</$($top.Type)>")
            if ($top.InLi) { $out.Add('</li>') }
            $listStack.RemoveAt($listStack.Count - 1)
        }
    }
    function Close-Para {
        if ($para.Count -gt 0) {
            $out.Add('<p>' + (Convert-MarkdownInline -Text ($para -join "`n") -BaseDir $BaseDir) + '</p>')
            $para.Clear()
        }
    }

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]

        # ---- fenced code blocks ------------------------------------------
        if ($inFence) {
            if ($line -match "^\s*$fenceMark\s*$") {
                $out.Add('<pre><code>' + ($fenceBuf -join "`n") + '</code></pre>')
                $fenceBuf.Clear()
                $inFence = $false
            } else {
                $fenceBuf.Add($line)
            }
            continue
        }
        if ($line -match '^\s*(`{3,}|~{3,})\s*\S*\s*$') {
            Close-Para; Close-Lists
            if ($inQuote) { $out.Add('</blockquote>'); $inQuote = $false }
            $fenceMark = $Matches[1]
            $inFence = $true
            continue
        }

        # ---- blank line ---------------------------------------------------
        if ($line -match '^\s*$') {
            Close-Para; Close-Lists
            if ($inQuote) { $out.Add('</blockquote>'); $inQuote = $false }
            continue
        }

        # ---- block quote --------------------------------------------------
        if ($line -match '^\s*&gt;\s?(.*)$') {
            if (-not $inQuote) { Close-Para; Close-Lists; $out.Add('<blockquote>'); $inQuote = $true }
            $para.Add($Matches[1])
            continue
        } elseif ($inQuote -and $para.Count -eq 0) {
            $out.Add('</blockquote>'); $inQuote = $false
        }

        # ---- atx heading ----------------------------------------------------
        if ($line -match '^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$') {
            Close-Para; Close-Lists
            $level = $Matches[1].Length
            $body  = Convert-MarkdownInline -Text $Matches[2] -BaseDir $BaseDir
            $out.Add(('<h{0} id="{1}">{2}</h{0}>' -f $level, (Get-HeadingSlug $Matches[2]), $body))
            continue
        }

        # ---- setext heading (must win over the horizontal rule below) ---------
        if ($para.Count -gt 0 -and $listStack.Count -eq 0 -and $line -match '^\s*(=+|-{2,})\s*$') {
            $level = if ($Matches[1].StartsWith('=')) { 1 } else { 2 }
            $body  = Convert-MarkdownInline -Text ($para -join ' ') -BaseDir $BaseDir
            $para.Clear()
            $out.Add(('<h{0} id="{1}">{2}</h{0}>' -f $level, (Get-HeadingSlug $body), $body))
            continue
        }

        # ---- horizontal rule ----------------------------------------------
        if ($line -match '^\s*([-*_])\s*(\1\s*){2,}$') {
            Close-Para; Close-Lists
            $out.Add('<hr />')
            continue
        }

        # ---- table --------------------------------------------------------
        if ($line -match '\|' -and $i + 1 -lt $lines.Count -and
            $lines[$i + 1] -match '^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$') {

            Close-Para; Close-Lists
            $splitRow = {
                param($r)
                ($r.Trim() -replace '^\|', '' -replace '\|$', '') -split '(?<!\\)\|' | ForEach-Object { $_.Trim() }
            }
            $headers = & $splitRow $line
            $aligns  = (& $splitRow $lines[$i + 1]) | ForEach-Object {
                if ($_ -match '^:.*:$') { 'center' } elseif ($_ -match ':$') { 'right' } elseif ($_ -match '^:') { 'left' } else { '' }
            }
            $out.Add('<table><thead><tr>')
            for ($c = 0; $c -lt $headers.Count; $c++) {
                $st = if ($aligns[$c]) { ' style="text-align:' + $aligns[$c] + '"' } else { '' }
                $out.Add(('<th{0}>{1}</th>' -f $st, (Convert-MarkdownInline -Text $headers[$c] -BaseDir $BaseDir)))
            }
            $out.Add('</tr></thead><tbody>')
            $i += 2
            while ($i -lt $lines.Count -and $lines[$i] -match '\|' -and $lines[$i] -notmatch '^\s*$') {
                $cells = & $splitRow $lines[$i]
                $out.Add('<tr>')
                for ($c = 0; $c -lt $headers.Count; $c++) {
                    $val = if ($c -lt $cells.Count) { $cells[$c] } else { '' }
                    $st  = if ($aligns[$c]) { ' style="text-align:' + $aligns[$c] + '"' } else { '' }
                    $out.Add(('<td{0}>{1}</td>' -f $st, (Convert-MarkdownInline -Text $val -BaseDir $BaseDir)))
                }
                $out.Add('</tr>')
                $i++
            }
            $out.Add('</tbody></table>')
            $i--
            continue
        }

        # ---- lists --------------------------------------------------------
        if ($line -match '^(\s*)([-*+]|\d+[.)])\s+(.*)$') {
            $indent  = $Matches[1].Replace("`t", '    ').Length
            $bullet  = $Matches[2]
            $content = $Matches[3]
            $type    = if ($bullet -match '^\d') { 'ol' } else { 'ul' }

            Close-Para

            while ($listStack.Count -gt 0 -and $indent -lt $listStack[$listStack.Count - 1].Indent) {
                $top = $listStack[$listStack.Count - 1]
                $out.Add("</$($top.Type)>")
                if ($top.InLi) { $out.Add('</li>') }
                $listStack.RemoveAt($listStack.Count - 1)
            }
            if ($listStack.Count -eq 0 -or $indent -gt $listStack[$listStack.Count - 1].Indent) {
                # a nested list belongs inside the preceding <li>
                $inLi = $false
                if ($listStack.Count -gt 0 -and $out.Count -gt 0 -and $out[$out.Count - 1].EndsWith('</li>')) {
                    $out[$out.Count - 1] = $out[$out.Count - 1].Substring(0, $out[$out.Count - 1].Length - 5)
                    $inLi = $true
                }
                $out.Add("<$type>")
                $listStack.Add(@{ Type = $type; Indent = $indent; InLi = $inLi })
            } elseif ($listStack[$listStack.Count - 1].Type -ne $type) {
                $top = $listStack[$listStack.Count - 1]
                $out.Add("</$($top.Type)>")
                if ($top.InLi) { $out.Add('</li>') }
                $listStack.RemoveAt($listStack.Count - 1)
                $out.Add("<$type>")
                $listStack.Add(@{ Type = $type; Indent = $indent; InLi = $false })
            }

            if ($content -match '^\[( |x|X)\]\s*(.*)$') {
                $chk = if ($Matches[1] -eq ' ') { '' } else { ' checked="checked"' }
                $nextChecked = if ($Matches[1] -eq ' ') { '1' } else { '0' }
                $toggleUrl = '{0}?toggle=1&line={1}&checked={2}' -f $script:LinkHost, $i, $nextChecked
                $toggleLink = '<a class="task-toggle" href="{0}"><input type="checkbox" disabled="disabled"{1} /></a>' -f $toggleUrl, $chk
                $out.Add(('<li class="task">{0} {1}</li>' -f `
                         $toggleLink, (Convert-MarkdownInline -Text $Matches[2] -BaseDir $BaseDir)))
            } else {
                $out.Add('<li>' + (Convert-MarkdownInline -Text $content -BaseDir $BaseDir) + '</li>')
            }
            continue
        }

        # ---- indented code block -------------------------------------------
        if ($listStack.Count -eq 0 -and $para.Count -eq 0 -and $line -match '^(\t|    )(.*)$') {
            $buf = New-Object System.Collections.Generic.List[string]
            while ($i -lt $lines.Count -and ($lines[$i] -match '^(\t|    )(.*)$' -or $lines[$i] -match '^\s*$')) {
                $buf.Add(($lines[$i] -replace '^(\t|    )', ''))
                $i++
            }
            $i--
            $out.Add('<pre><code>' + (($buf -join "`n").TrimEnd()) + '</code></pre>')
            continue
        }

        # ---- plain paragraph text -------------------------------------------
        if ($listStack.Count -gt 0) {
            # continuation of the previous list item
            $out.Add(' ' + (Convert-MarkdownInline -Text $line.Trim() -BaseDir $BaseDir))
            continue
        }
        $para.Add($line.Trim())
    }

    if ($inFence -and $fenceBuf.Count -gt 0) { $out.Add('<pre><code>' + ($fenceBuf -join "`n") + '</code></pre>') }
    Close-Para
    Close-Lists
    if ($inQuote) { $out.Add('</blockquote>') }

    $body = $out -join "`n"
    $body = Add-HtmlHighlight -Html $body -Term $Highlight
    if ($TargetAnchor) {
        $anchorId = Get-HeadingSlug ([uri]::UnescapeDataString($TargetAnchor))
        $body = [regex]::Replace($body, ('(<h[1-6]) id="{0}[^"]*"' -f [regex]::Escape($anchorId)), ('$1 id="' + $anchorId + '"'), 1)
    }

    @"
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta charset="utf-8" />
<style>
$script:PreviewCss
</style>
</head>
<body>
$body
</body>
</html>
"@
}

$script:PreviewCss = @'
body { font-family: Segoe UI, Verdana, sans-serif; font-size: 10.5pt; line-height: 1.55;
       color: #24292f; background: #ffffff; margin: 16px 22px 40px 22px; }
h1, h2, h3, h4, h5, h6 { font-weight: 600; line-height: 1.25; margin: 22px 0 10px 0; }
h1 { font-size: 1.9em; border-bottom: 1px solid #d8dee4; padding-bottom: 6px; }
h2 { font-size: 1.5em; border-bottom: 1px solid #d8dee4; padding-bottom: 5px; }
h3 { font-size: 1.25em; }
h4 { font-size: 1.05em; }
p  { margin: 0 0 12px 0; }
a  { color: #0969da; text-decoration: none; }
a:hover { text-decoration: underline; }
code { font-family: Consolas, Courier New, monospace; font-size: 0.92em;
       background: #f2f3f5; padding: 1px 5px; border-radius: 3px; }
pre  { background: #f6f8fa; border: 1px solid #e3e6ea; border-radius: 4px;
       padding: 12px 14px; overflow: auto; }
pre code { background: transparent; padding: 0; }
blockquote { margin: 0 0 12px 0; padding: 2px 14px; color: #57606a;
             border-left: 4px solid #d0d7de; background: #fafbfc; }
table { border-collapse: collapse; margin: 0 0 14px 0; }
th, td { border: 1px solid #d0d7de; padding: 6px 12px; }
th { background: #f6f8fa; font-weight: 600; }
tr:nth-child(even) td { background: #fbfcfd; }
ul, ol { margin: 0 0 12px 0; padding-left: 28px; }
li { margin: 3px 0; }
li.task { list-style: none; margin-left: -20px; }
li.task a.task-toggle { color: inherit; text-decoration: none; }
img { max-width: 100%; }
hr { border: 0; border-top: 1px solid #d8dee4; margin: 20px 0; }
del { color: #8c959f; }
span.mdhit { background: #ffe95e; color: #24292f; }
'@

#endregion

#region ---------------------------------------------------------------- xaml

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MD-Browser" Height="820" Width="1280" WindowStartupLocation="CenterScreen">
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto" />
      <RowDefinition Height="*" />
      <RowDefinition Height="Auto" />
    </Grid.RowDefinitions>

    <ToolBarTray Grid.Row="0">
      <ToolBar>
        <Button x:Name="BtnOpen"    Padding="8,3" ToolTip="Choose another root folder">Open folder...</Button>
        <Button x:Name="BtnRefresh" Padding="8,3" ToolTip="Rescan the folder structure (F5)">Refresh</Button>
        <ToggleButton x:Name="BtnExpandAll" Padding="8,3" ToolTip="Expand or collapse the whole tree">Expand all</ToggleButton>
        <Separator />
        <ToggleButton x:Name="BtnEdit" Padding="8,3" ToolTip="Switch between preview and editor (F4)">Edit</ToggleButton>
        <Button x:Name="BtnSave" Padding="8,3" ToolTip="Save the current file (Ctrl+S)" IsEnabled="False">Save</Button>
        <Separator />
        <Button x:Name="BtnReveal" Padding="8,3" ToolTip="Show the selected tree item in Explorer">Show in Explorer</Button>
      </ToolBar>
    </ToolBarTray>

    <Grid Grid.Row="1">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="320" MinWidth="180" />
        <ColumnDefinition Width="5" />
        <ColumnDefinition Width="*" MinWidth="300" />
      </Grid.ColumnDefinitions>

      <DockPanel Grid.Column="0" Margin="6,6,0,6">
        <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6" LastChildFill="True">
          <Button x:Name="BtnHitDown" DockPanel.Dock="Right" Width="24" Height="24" Margin="2,0,0,0"
                  IsEnabled="False" ToolTip="Next occurrence (F3)">&#x25BC;</Button>
          <Button x:Name="BtnHitUp" DockPanel.Dock="Right" Width="24" Height="24" Margin="2,0,0,0"
                  IsEnabled="False" ToolTip="Previous occurrence (Shift+F3)">&#x25B2;</Button>
          <TextBlock x:Name="LblHits" DockPanel.Dock="Right" MinWidth="48" Margin="6,0,2,0"
                     VerticalAlignment="Center" TextAlignment="Center" Foreground="#57606A" />
          <Button x:Name="BtnClear" DockPanel.Dock="Right" Width="24" Height="24" Margin="4,0,0,0"
                  ToolTip="Clear the search (Esc)">&#x2715;</Button>
          <TextBox x:Name="TxtFilter" Height="24" VerticalContentAlignment="Center"
                   ToolTip="Search: matches file and folder names, and from 3 characters on also the file content. Hits are highlighted in the displayed file." />
        </DockPanel>
        <TreeView x:Name="Tree" BorderBrush="#C8CDD3" />
      </DockPanel>

      <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch" Background="#EDEFF2" />

      <DockPanel Grid.Column="2" Margin="0,6,6,6">
        <Button x:Name="BtnBack" DockPanel.Dock="Top" HorizontalAlignment="Left" Padding="8,3"
                Margin="0,0,0,6" Visibility="Collapsed"
                ToolTip="Back to the previously visited file (Alt+Left)">Back</Button>
        <Grid>
          <WebBrowser x:Name="Preview" />
          <TextBox x:Name="Editor" Visibility="Collapsed"
                   FontFamily="Consolas" FontSize="13" AcceptsReturn="True" AcceptsTab="True"
                   TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto"
                   HorizontalScrollBarVisibility="Auto" BorderBrush="#C8CDD3" />
        </Grid>
      </DockPanel>

            <Border x:Name="SearchBusy" Grid.Column="2" Panel.ZIndex="10"
                            HorizontalAlignment="Right" VerticalAlignment="Bottom" Visibility="Collapsed"
                            Margin="0,0,18,18" Padding="11,7" CornerRadius="4"
                            Background="#F6F8FA" BorderBrush="#C8CDD3" BorderThickness="1">
                <StackPanel Orientation="Horizontal">
                    <TextBlock x:Name="SearchBusyText" VerticalAlignment="Center" Foreground="#24292F" />
                    <TextBlock x:Name="SearchBusyPulse" Width="18" Margin="8,0,0,0"
                                         VerticalAlignment="Center" TextAlignment="Center" FontFamily="Consolas"
                                         FontWeight="Bold" Foreground="#0969DA" />
                </StackPanel>
            </Border>
    </Grid>

    <StatusBar Grid.Row="2">
      <StatusBarItem><TextBlock x:Name="LblStatus" Text="Ready" /></StatusBarItem>
    </StatusBar>
  </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$win    = [Windows.Markup.XamlReader]::Load($reader)

$Tree       = $win.FindName('Tree')
$TxtFilter  = $win.FindName('TxtFilter')
$BtnClear   = $win.FindName('BtnClear')
$LblHits    = $win.FindName('LblHits')
$BtnHitUp   = $win.FindName('BtnHitUp')
$BtnHitDown = $win.FindName('BtnHitDown')
$Preview    = $win.FindName('Preview')
$Editor     = $win.FindName('Editor')
$BtnBack    = $win.FindName('BtnBack')
$BtnOpen    = $win.FindName('BtnOpen')
$BtnRefresh = $win.FindName('BtnRefresh')
$BtnExpand  = $win.FindName('BtnExpandAll')
$BtnEdit    = $win.FindName('BtnEdit')
$BtnSave    = $win.FindName('BtnSave')
$BtnReveal  = $win.FindName('BtnReveal')
$LblStatus  = $win.FindName('LblStatus')
$SearchBusy = $win.FindName('SearchBusy')
$SearchBusyText = $win.FindName('SearchBusyText')
$SearchBusyPulse = $win.FindName('SearchBusyPulse')

function Start-SearchBusy {
    param([string]$Term)
    if ([string]::IsNullOrWhiteSpace($Term)) { return }
    $script:IsSearching = $true
    $script:SearchPulse = 0
    $script:SearchLastYield = [datetime]::MinValue
    $SearchBusyText.Text = "Searching for '$Term'"
    $SearchBusyPulse.Text = '|'
    $SearchBusy.Visibility = 'Visible'
    $win.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [action]{})
}

function Update-SearchBusy {
    if (-not $script:IsSearching -or ([datetime]::UtcNow - $script:SearchLastYield).TotalMilliseconds -lt 80) { return }
    $frames = @('|', '/', '-', '\')
    $script:SearchPulse = ($script:SearchPulse + 1) % $frames.Count
    $SearchBusyPulse.Text = $frames[$script:SearchPulse]
    $script:SearchLastYield = [datetime]::UtcNow
    $win.Dispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
}

function Stop-SearchBusy {
    $script:IsSearching = $false
    $SearchBusy.Visibility = 'Collapsed'
}

function Set-PreviewSilent {
    try {
        $flags = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
        $browser = $null
        $field = $Preview.GetType().GetField('axIWebBrowser2', $flags)
        if ($field) { $browser = $field.GetValue($Preview) }
        if (-not $browser) {
            $property = $Preview.GetType().GetProperty('ActiveXInstance', $flags)
            if ($property) { $browser = $property.GetValue($Preview, $null) }
        }
        if ($browser) { $browser.Silent = $true }
    } catch { }
}

$Preview.Add_Loaded({ Set-PreviewSilent })

#endregion

#region ---------------------------------------------------------------- tree

function Test-MarkdownFile {
    param([string]$FilePath)
    $script:MarkdownExt -contains ([System.IO.Path]::GetExtension($FilePath)).ToLowerInvariant()
}

function New-TreeNode {
    param(
        [string]$Header,
        [hashtable]$Tag,
        [string]$ToolTip,
        [string]$Foreground
    )
    $node = New-Object System.Windows.Controls.TreeViewItem
    $node.Header  = $Header
    $node.Tag     = $Tag
    $node.Padding = '2,1,2,1'
    if ($ToolTip)    { $node.ToolTip = $ToolTip }
    if ($Foreground) { $node.Foreground = $Foreground }
    $node
}

function Get-CachedFileText {
    <# Reads a file once and keeps it in memory until its timestamp changes. #>
    param([System.IO.FileInfo]$File)

    $entry = $script:ContentCache[$File.FullName]
    if ($entry -and $entry.Stamp -eq $File.LastWriteTimeUtc.Ticks) { return $entry.Text }

    $text = ''
    try { $text = [System.IO.File]::ReadAllText($File.FullName) } catch { }
    $script:ContentCache[$File.FullName] = @{ Stamp = $File.LastWriteTimeUtc.Ticks; Text = $text }
    $text
}

function Measure-TextHit {
    param([string]$Text, [string]$Needle)
    if ([string]::IsNullOrEmpty($Text) -or [string]::IsNullOrEmpty($Needle)) { return 0 }
    $count = 0
    $i = $Text.IndexOf($Needle, [StringComparison]::OrdinalIgnoreCase)
    while ($i -ge 0) {
        $count++
        $i = $Text.IndexOf($Needle, $i + $Needle.Length, [StringComparison]::OrdinalIgnoreCase)
    }
    $count
}

function Add-FolderNodes {
    <# Populates $Parent with the sub folders and markdown files of $FolderPath. Returns $true if anything was added. #>
    param(
        [string]$FolderPath,
        $Parent,
        [string]$Filter,
        [bool]$SearchContent
    )

    $added = $false
    Update-SearchBusy

    $dirs = @(Get-ChildItem -LiteralPath $FolderPath -Directory -ErrorAction SilentlyContinue |
              Where-Object { -not $_.Attributes.HasFlag([System.IO.FileAttributes]::Hidden) } |
              Sort-Object Name)

    foreach ($d in $dirs) {
        Update-SearchBusy
        $node = New-TreeNode -Header $d.Name -Tag @{ Type = 'dir'; Path = $d.FullName } -ToolTip $d.FullName
        $hasChildren = Add-FolderNodes -FolderPath $d.FullName -Parent $node -Filter $Filter -SearchContent $SearchContent
        if ($hasChildren) {
            if ($Filter) { $node.IsExpanded = $true }
            $Parent.Items.Add($node) | Out-Null
            $added = $true
        }
    }

    $files = @(Get-ChildItem -LiteralPath $FolderPath -File -ErrorAction SilentlyContinue |
               Where-Object { Test-MarkdownFile $_.FullName } |
               Sort-Object Name)

    foreach ($f in $files) {
        Update-SearchBusy
        $hits = 0
        if ($Filter) {
            $nameHit = $f.Name.IndexOf($Filter, [StringComparison]::OrdinalIgnoreCase) -ge 0
            if ($SearchContent) { $hits = Measure-TextHit -Text (Get-CachedFileText $f) -Needle $Filter }
            if (-not $nameHit -and $hits -eq 0) { continue }
        }
        $header = if ($hits -gt 0) { '{0}  ({1})' -f $f.Name, $hits } else { $f.Name }
        $Parent.Items.Add((New-TreeNode -Header $header -Tag @{ Type = 'file'; Path = $f.FullName } -ToolTip $f.FullName)) | Out-Null
        $script:MatchCount++
        $added = $true
    }

    $added
}

function Update-Tree {
    param([string]$Filter)

    if (-not $script:RootPath) { return }

    $script:SuppressSelect = $true
    $script:MatchCount     = 0
    $searchContent = [bool]($Filter -and $Filter.Length -ge $script:ContentMinChars)

    $Tree.Items.Clear()

    $root = New-TreeNode -Header ([System.IO.Path]::GetFileName($script:RootPath.TrimEnd('\'))) `
                         -Tag @{ Type = 'dir'; Path = $script:RootPath } -ToolTip $script:RootPath
    $root.FontWeight = 'Bold'
    $root.IsExpanded = $true
    Add-FolderNodes -FolderPath $script:RootPath -Parent $root -Filter $Filter -SearchContent $searchContent | Out-Null
    $Tree.Items.Add($root) | Out-Null

    Add-StandaloneBranch
    $script:SuppressSelect = $false

    if ($Filter) {
        $scope = if ($searchContent) { 'name + content' } else { "name only - type $script:ContentMinChars characters to search the content" }
        Set-Status ("{0} matching file(s) for '{1}' ({2})" -f $script:MatchCount, $Filter, $scope)
    }
}

function Set-TreeExpansion {
    param([bool]$Expanded)

    $stack = New-Object System.Collections.Generic.Stack[object]
    foreach ($i in $Tree.Items) { $stack.Push($i) }
    while ($stack.Count -gt 0) {
        $item = $stack.Pop()
        $item.IsExpanded = $Expanded
        foreach ($c in $item.Items) { $stack.Push($c) }
    }
    $BtnExpand.IsChecked = $Expanded
    $label = if ($Expanded) { 'Collapse all' } else { 'Expand all' }
    $BtnExpand.Content = $label
}

function Add-StandaloneBranch {
    if ($script:Standalone.Count -eq 0) { return }

    $branch = New-TreeNode -Header 'Standalone files' -Tag @{ Type = 'group' }
    $branch.FontWeight = 'Bold'
    $branch.IsExpanded = $true

    foreach ($p in ($script:Standalone | Sort-Object)) {
        $label = '{0}  [{1}]' -f ([System.IO.Path]::GetFileName($p)), $p
        $fg    = if (Test-Path -LiteralPath $p) { '#1F6F3C' } else { '#B4232A' }
        $branch.Items.Add((New-TreeNode -Header $label -Tag @{ Type = 'file'; Path = $p; Standalone = $true } -ToolTip $p -Foreground $fg)) | Out-Null
    }
    $Tree.Items.Add($branch) | Out-Null
}

function Select-TreeFile {
    <# Expands the tree down to $FilePath and selects that node. Returns $true on success. #>
    param([string]$FilePath)

    $stack = New-Object System.Collections.Generic.Stack[object]
    foreach ($i in $Tree.Items) { $stack.Push($i) }

    while ($stack.Count -gt 0) {
        $item = $stack.Pop()
        $tag  = $item.Tag
        if ($tag -and $tag.Type -eq 'file' -and $tag.Path -eq $FilePath) {
            $p = $item.Parent
            while ($p -is [System.Windows.Controls.TreeViewItem]) { $p.IsExpanded = $true; $p = $p.Parent }
            $item.IsSelected = $true
            $item.BringIntoView()
            return $true
        }
        foreach ($c in $item.Items) { $stack.Push($c) }
    }
    $false
}

function Get-FirstTreeFile {
    param($Items)

    foreach ($item in $Items) {
        if ($item.Tag -and $item.Tag.Type -eq 'file' -and (Test-MarkdownFile $item.Tag.Path)) {
            return $item.Tag.Path
        }
        $found = Get-FirstTreeFile -Items $item.Items
        if ($found) { return $found }
    }
    $null
}

#endregion

#region --------------------------------------------------------------- files

function Set-Status {
    param([string]$Text)
    $LblStatus.Text = $Text
}

function Update-Title {
    $name = if ($script:CurrentFile) { $script:CurrentFile } else { '(no file)' }
    $mark = if ($script:IsDirty) { ' *' } else { '' }
    $win.Title = "MD-Browser - $name$mark"
    $BtnSave.IsEnabled = $script:IsDirty
}

function Save-CurrentFile {
    if (-not $script:CurrentFile -or -not $script:IsDirty) { return }
    try {
        [System.IO.File]::WriteAllText($script:CurrentFile, $Editor.Text, (New-Object System.Text.UTF8Encoding($false)))
        $script:IsDirty = $false
        Update-Title
        Set-Status "Saved $($script:CurrentFile)"
    } catch {
        [System.Windows.MessageBox]::Show("Could not save the file:`n$($_.Exception.Message)", 'MD-Browser',
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
}

function Confirm-PendingChanges {
    if (-not $script:IsDirty) { return $true }
    $r = [System.Windows.MessageBox]::Show(
            "'$([System.IO.Path]::GetFileName($script:CurrentFile))' has unsaved changes. Save them now?",
            'MD-Browser', [System.Windows.MessageBoxButton]::YesNoCancel, [System.Windows.MessageBoxImage]::Question)
    switch ($r) {
        'Yes'    { Save-CurrentFile; return $true }
        'No'     { $script:IsDirty = $false; return $true }
        default  { return $false }
    }
}

function Update-StandaloneList {
    <# Rebuilds the standalone list from the links of the currently opened document. #>
    param(
        [string]$Markdown,
        [string]$BaseDir
    )

    $script:Standalone.Clear()
    $rootFull = [System.IO.Path]::GetFullPath($script:RootPath).TrimEnd('\') + '\'

    foreach ($m in [regex]::Matches($Markdown, '(?<!\!)\[[^\]]*\]\(\s*([^)\s]+)(?:\s+"[^"]*")?\s*\)')) {
        $res = Resolve-MarkdownTarget -Target $m.Groups[1].Value -BaseDir $BaseDir
        if ($res.Kind -ne 'local' -or -not $res.FullPath) { continue }
        if ($res.FullPath.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) { continue }
        if (-not $script:Standalone.Contains($res.FullPath)) { $script:Standalone.Add($res.FullPath) | Out-Null }
    }
}

function Get-SearchTerm {
    $TxtFilter.Text.Trim()
}

function Update-HitUi {
    $term  = Get-SearchTerm
    if ($script:IsEditMode -and $script:CurrentFile) {
        $script:HitTotal = Measure-TextHit -Text $Editor.Text -Needle $term
    }
    if ($script:HitIndex -gt $script:HitTotal) { $script:HitIndex = 0 }

    $label = if ($term -and $script:CurrentFile) { '{0} / {1}' -f $script:HitIndex, $script:HitTotal } else { '' }
    $LblHits.Text        = $label
    $BtnHitUp.IsEnabled   = $script:HitTotal -gt 0
    $BtnHitDown.IsEnabled = $script:HitTotal -gt 0
}

function Update-Preview {
    <# Renders the editor content into the preview, highlighting the current search term. #>
    if (-not $script:CurrentFile) { return }
    try {
        $script:PreviewLoading = $true
        $Preview.NavigateToString((Convert-MarkdownToHtml -Markdown $Editor.Text -BaseDir $script:CurrentDir -Highlight (Get-SearchTerm)))
    } catch {
        Set-Status "Render error: $($_.Exception.Message)"
    }
    Update-HitUi
}

function Move-Hit {
    param([int]$Delta)

    Update-HitUi
    if ($script:HitTotal -le 0) { return }

    $script:HitIndex += $Delta
    if ($script:HitIndex -lt 1)                { $script:HitIndex = $script:HitTotal }
    if ($script:HitIndex -gt $script:HitTotal) { $script:HitIndex = 1 }

    if ($script:IsEditMode) {
        $term = Get-SearchTerm
        $pos  = -1
        $from = 0
        for ($k = 0; $k -lt $script:HitIndex; $k++) {
            $pos = $Editor.Text.IndexOf($term, $from, [StringComparison]::OrdinalIgnoreCase)
            if ($pos -lt 0) { break }
            $from = $pos + $term.Length
        }
        if ($pos -ge 0) {
            $Editor.Focus() | Out-Null
            $Editor.Select($pos, $term.Length)
            $Editor.ScrollToLine($Editor.GetLineIndexFromCharacterIndex($pos))
        }
    } else {
        Start-HitScroll -ElementId "mdhit$($script:HitIndex - 1)"
    }

    Update-HitUi
}

function Open-MarkdownFile {
    param([string]$FilePath)

    if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
        Set-Status "File not found: $FilePath"
        $Preview.NavigateToString("<html><body style='font-family:Segoe UI'><p>File not found:</p><p><b>$FilePath</b></p></body></html>")
        return
    }

    $script:CurrentFile = $FilePath
    $script:CurrentDir  = Split-Path -Parent $FilePath
    $script:IsDirty     = $false

    $md = Get-Content -LiteralPath $FilePath -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
    if ($null -eq $md) { $md = '' }

    $script:Loading = $true
    $Editor.Text    = $md
    $script:Loading = $false
    $script:IsDirty = $false

    $before = ($script:Standalone -join '|')
    Update-StandaloneList -Markdown $md -BaseDir $script:CurrentDir
    if (($script:Standalone -join '|') -ne $before) {
        $sel = $script:CurrentFile
        Update-Tree -Filter $TxtFilter.Text
        $script:SuppressSelect = $true
        Select-TreeFile -FilePath $sel | Out-Null
        $script:SuppressSelect = $false
    }

    try {
        $script:PreviewLoading = $true
        $Preview.NavigateToString((Convert-MarkdownToHtml -Markdown $md -BaseDir $script:CurrentDir `
            -Highlight (Get-SearchTerm) -TargetAnchor $script:PendingAnchor))
    } catch {
        Set-Status "Render error: $($_.Exception.Message)"
    }

    $script:HitIndex = 0
    Update-HitUi
    Update-Title
    Set-Status $FilePath
}

function ConvertTo-HtmlText {
    param([object]$Value)
    [System.Web.HttpUtility]::HtmlEncode([string]$Value)
}

function Open-FolderInfo {
    param([string]$FolderPath)

    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) { return }

    $script:CurrentFile = $null
    $script:CurrentDir  = $FolderPath
    $script:IsDirty     = $false
    $script:HitIndex    = 0
    $script:HitTotal    = 0
    $script:Loading     = $true
    $Editor.Text        = ''
    $script:Loading     = $false
    $Editor.Visibility  = 'Collapsed'
    $Preview.Visibility = 'Visible'
    $BtnEdit.IsChecked  = $false
    $BtnEdit.Content    = 'Edit'
    Update-HitUi
    Update-Title

    $folder = Get-Item -LiteralPath $FolderPath
    $children = @(Get-ChildItem -LiteralPath $FolderPath -Force -ErrorAction SilentlyContinue |
                  Where-Object { -not $_.Attributes.HasFlag([System.IO.FileAttributes]::Hidden) })
    $folders = @($children | Where-Object { $_.PSIsContainer } | Sort-Object Name)
    $files = @($children | Where-Object { -not $_.PSIsContainer } | Sort-Object Name)
    $mdFiles = @($files | Where-Object { Test-MarkdownFile $_.FullName })

    $rows = New-Object System.Collections.Generic.List[string]
    foreach ($sub in $folders) {
        $subItems = @(Get-ChildItem -LiteralPath $sub.FullName -Force -ErrorAction SilentlyContinue |
                      Where-Object { -not $_.Attributes.HasFlag([System.IO.FileAttributes]::Hidden) })
        $subMd = @($subItems | Where-Object { -not $_.PSIsContainer -and (Test-MarkdownFile $_.FullName) }).Count
        $rows.Add(('<tr><td>{0}</td><td>Folder</td><td>{1} direct item(s), {2} Markdown file(s)</td></tr>' -f `
            (ConvertTo-HtmlText $sub.Name), $subItems.Count, $subMd))
    }
    foreach ($file in $mdFiles) {
        $text = Get-CachedFileText $file
        $lines = if ($text) { @($text -split "`r?`n").Count } else { 0 }
        $headings = ([regex]::Matches($text, '(?m)^\s{0,3}#{1,6}\s+')).Count
        $rows.Add(('<tr><td>{0}</td><td>Markdown</td><td>{1} KB, {2} line(s), {3} heading(s), modified {4}</td></tr>' -f `
            (ConvertTo-HtmlText $file.Name), [math]::Round($file.Length / 1KB, 1), $lines, $headings, $file.LastWriteTime))
    }

    $table = if ($rows.Count -gt 0) { $rows -join "`n" } else { '<tr><td colspan="3">No visible subfolders or Markdown files.</td></tr>' }
    $html = @"
<!DOCTYPE html><html><head><meta http-equiv="X-UA-Compatible" content="IE=edge" /><style>$script:PreviewCss
body { max-width: 980px; } .summary { color: #57606a; } table { width: 100%; } td:first-child { font-weight: 600; }
</style></head><body>
<h1>$(ConvertTo-HtmlText $folder.Name)</h1>
<p class="summary">$(ConvertTo-HtmlText $folder.FullName)</p>
<p><b>$($folders.Count)</b> subfolder(s), <b>$($mdFiles.Count)</b> Markdown file(s), <b>$($files.Count)</b> file(s) total.</p>
<table><thead><tr><th>Name</th><th>Type</th><th>Details</th></tr></thead><tbody>$table</tbody></table>
</body></html>
"@
    $Preview.NavigateToString($html)
    Set-Status "Folder: $FolderPath"
}

function Invoke-ExternalTarget {
    param([string]$Target)
    if ([string]::IsNullOrWhiteSpace($Target)) { Set-Status 'Empty link target - nothing to open'; return }
    try { Start-Process $Target } catch {
        [System.Windows.MessageBox]::Show("Could not open:`n$Target`n`n$($_.Exception.Message)", 'MD-Browser',
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
    }
}

function Update-BackUi {
    if ($script:History.Count -gt 0) {
        $prev = $script:History[$script:History.Count - 1]
        $BtnBack.Content    = '{0} Back to {1}' -f ([char]0x25C0), ([System.IO.Path]::GetFileName($prev))
        $BtnBack.ToolTip    = "Back to $prev  (Alt+Left)"
        $BtnBack.Visibility = 'Visible'
    } else {
        $BtnBack.Visibility = 'Collapsed'
    }
}

function Clear-History {
    $script:History.Clear()
    Update-BackUi
}

function Invoke-PreviewScroll {
    <# Scrolls the preview to an element id, or to the first heading whose slug starts with it. #>
    param([string]$ElementId)

    if ([string]::IsNullOrWhiteSpace($ElementId)) { return $false }
    $id = $ElementId -replace '[^a-zA-Z0-9_-]', ''
    try {
        $doc = $Preview.Document
        if (-not $doc) { return $false }

        if ($script:JumpDebug) {
            $script:JumpDebugLines.Add("Rendered headings/anchors:")
            foreach ($tagName in @('h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'a')) {
                try {
                    $elements = $doc.getElementsByTagName($tagName)
                    for ($j = 0; $j -lt $elements.length -and $j -lt 50; $j++) {
                        $candidate = $elements.item($j)
                        $idText = [string]$candidate.id
                        $text = ([string]$candidate.innerText).Trim()
                        if ($idText -or $tagName -ne 'a') {
                            $script:JumpDebugLines.Add("<$tagName> id='$idText' text='$text'")
                        }
                    }
                } catch { $script:JumpDebugLines.Add("<$tagName> read error: $($_.Exception.Message)") }
            }
        }

        $element = $null
        try { $element = $doc.getElementById($id) } catch { }
        if (-not $element) {
            foreach ($tagName in @('h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'a')) {
                try {
                    $elements = $doc.getElementsByTagName($tagName)
                    for ($i = 0; $i -lt $elements.length; $i++) {
                        $candidate = $elements.item($i)
                        if ($tagName -eq 'a' -and ([string]$candidate.id) -eq $id) {
                            $element = $candidate
                            break
                        }
                        if ($tagName -ne 'a') {
                            $headingSlug = Get-HeadingSlug ([string]$candidate.innerText)
                            if ($headingSlug.StartsWith($id, [StringComparison]::OrdinalIgnoreCase)) {
                                $element = $candidate
                                break
                            }
                        }
                    }
                } catch { }
                if ($element) { break }
            }
        }

        if ($element) {
            $element.scrollIntoView($true)
            return $true
        }
    } catch { }
    $false
}

function Start-HitScroll {
    param([string]$ElementId)

    $script:PendingHitId = $ElementId
    $script:HitScrollAttempts = 0
    if ($script:HitScrollTimer) { $script:HitScrollTimer.Stop() }

    if (-not $script:PreviewLoading -and (Invoke-PreviewScroll -ElementId $ElementId)) {
        $script:PendingHitId = $null
        return
    }

    $script:HitScrollTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:HitScrollTimer.Interval = [TimeSpan]::FromMilliseconds(75)
    $script:HitScrollTimer.Add_Tick({
        if ($script:PreviewLoading) { return }
        $script:HitScrollAttempts++
        $hitId = $script:PendingHitId
        if (-not $hitId -or (Invoke-PreviewScroll -ElementId $hitId)) {
            $script:PendingHitId = $null
            $script:HitScrollTimer.Stop()
        } elseif ($script:HitScrollAttempts -ge 20) {
            $script:PendingHitId = $null
            $script:HitScrollTimer.Stop()
        }
    })
    $script:HitScrollTimer.Start()
}

function Show-JumpDebug {
    param([string]$Result)

    if (-not $script:JumpDebug) { return }
    $details = @($script:JumpDebugLines)
    $details += "Result: $Result"
    [System.Windows.MessageBox]::Show(($details -join "`n"), 'MD-Browser section jump diagnostics',
        [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
}

function Save-CheckboxState {
    param(
        [int]$LineIndex,
        [bool]$Checked
    )

    if (-not $script:CurrentFile -or -not (Test-Path -LiteralPath $script:CurrentFile -PathType Leaf)) { return }

    try {
        $text = [System.IO.File]::ReadAllText($script:CurrentFile)
        $newline = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }
        $lines = $text -split "`r?`n", -1
        if ($LineIndex -lt 0 -or $LineIndex -ge $lines.Count) { throw "Checkbox line is outside the file" }

        $marker = if ($Checked) { 'x' } else { ' ' }
        $updated = $lines[$LineIndex] -replace '^([\s]*(?:[-*+]|\d+[.)])\s+)\[(?: |x|X)\]', ('$1[' + $marker + ']')
        if ($updated -eq $lines[$LineIndex]) { throw "Checkbox marker was not found on line $($LineIndex + 1)" }

        $lines[$LineIndex] = $updated
        [System.IO.File]::WriteAllText($script:CurrentFile, ($lines -join $newline), (New-Object System.Text.UTF8Encoding($false)))
        Open-MarkdownFile -FilePath $script:CurrentFile
        Set-Status "Saved checkbox change in $([System.IO.Path]::GetFileName($script:CurrentFile))"
    } catch {
        Set-Status "Could not save checkbox: $($_.Exception.Message)"
    }
}

function Invoke-AnchorJump {
    param([string]$Fragment)

    if ([string]::IsNullOrWhiteSpace($Fragment)) { return }
    $slug = Get-HeadingSlug ([uri]::UnescapeDataString($Fragment))
    $script:JumpDebugLines.Clear()
    if ($script:PendingJumpUrl) { $script:JumpDebugLines.Add("Link URL: $script:PendingJumpUrl") }
    if ($script:PendingJumpPath) { $script:JumpDebugLines.Add("Resolved target: $script:PendingJumpPath (exists=$(Test-Path -LiteralPath $script:PendingJumpPath))") }
    $script:JumpDebugLines.Add("Fragment: #$Fragment")
    $script:JumpDebugLines.Add("Slug: $slug")
    $script:JumpDebugLines.Add("Current file: $script:CurrentFile")

    $found = Invoke-PreviewScroll -ElementId $slug
    if (-not $found) {
        Set-Status ("Section '#{0}' not found in {1}" -f $Fragment, [System.IO.Path]::GetFileName($script:CurrentFile))
    }
    Show-JumpDebug -Result $(if ($found) { 'MATCH FOUND' } else { 'NO MATCH' })
    $script:PendingJumpUrl = $null
    $script:PendingJumpPath = $null
}

function Start-PendingAnchorJump {
    if (-not $script:PendingAnchor) { return }

    if ($script:AnchorTimer) { $script:AnchorTimer.Stop() }
    $script:AnchorAttempts = 0
    $script:AnchorTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:AnchorTimer.Interval = [TimeSpan]::FromMilliseconds(100)
    $script:AnchorTimer.Add_Tick({
        $script:AnchorAttempts++
        $fragment = $script:PendingAnchor
        if ($fragment -and (Invoke-PreviewScroll -ElementId (Get-HeadingSlug ([uri]::UnescapeDataString($fragment))))) {
            $script:PendingAnchor = $null
            $script:AnchorTimer.Stop()
        } elseif ($script:AnchorAttempts -ge 30) {
            $script:PendingAnchor = $null
            $script:AnchorTimer.Stop()
            if ($fragment) {
                Set-Status ("Section '#{0}' not found in {1}" -f $fragment, [System.IO.Path]::GetFileName($script:CurrentFile))
            }
        }
    })
    $script:AnchorTimer.Start()
}

function Invoke-HistoryBack {
    if ($script:History.Count -eq 0) { return }
    if (-not (Confirm-PendingChanges)) { return }

    $prev = $script:History[$script:History.Count - 1]
    $script:History.RemoveAt($script:History.Count - 1)

    $script:SuppressSelect = $true
    $found = Select-TreeFile -FilePath $prev
    $script:SuppressSelect = $false

    Open-MarkdownFile -FilePath $prev
    if (-not $found) {
        $script:SuppressSelect = $true
        Select-TreeFile -FilePath $prev | Out-Null
        $script:SuppressSelect = $false
    }
    Update-BackUi
}

function Invoke-LinkTarget {
    param([string]$Target)

    if ([string]::IsNullOrWhiteSpace($Target)) { Set-Status 'Empty link target - nothing to open'; return }

    if ($Target.StartsWith($script:LinkHost, [StringComparison]::OrdinalIgnoreCase)) {
        try {
            $nested = [System.Web.HttpUtility]::ParseQueryString(([uri]$Target).Query)['t']
            if ($nested) { $Target = $nested }
        } catch { }
    }

    $res = Resolve-MarkdownTarget -Target $Target -BaseDir $script:CurrentDir

    switch ($res.Kind) {
        'anchor'   { Invoke-AnchorJump -Fragment $res.Fragment; return }
        'external' { Invoke-ExternalTarget $res.Raw; return }
    }

    $full = $res.FullPath

    if (Test-Path -LiteralPath $full -PathType Container) {
        Invoke-ExternalTarget $full
        return
    }
    if (-not (Test-MarkdownFile $full)) {
        if (Test-Path -LiteralPath $full -PathType Leaf) { Invoke-ExternalTarget $full }
        else { Set-Status "Target does not exist: $full" }
        return
    }

    if ($full -eq $script:CurrentFile) {
        $script:PendingJumpPath = $full
        Invoke-AnchorJump -Fragment $res.Fragment
        return
    }
    if (-not (Confirm-PendingChanges)) { return }

    if ($script:CurrentFile) { $script:History.Add($script:CurrentFile) | Out-Null }
    $script:PendingAnchor = $res.Fragment
    $script:PendingJumpPath = $full

    $script:SuppressSelect = $true
    $found = Select-TreeFile -FilePath $full
    $script:SuppressSelect = $false

    Open-MarkdownFile -FilePath $full
    if (-not $found) {
        $script:SuppressSelect = $true
        Select-TreeFile -FilePath $full | Out-Null
        $script:SuppressSelect = $false
    }
    Update-BackUi
}

function Set-EditMode {
    param([bool]$Enabled)

    $script:IsEditMode = $Enabled
    if ($Enabled) {
        $Editor.Visibility  = 'Visible'
        $Preview.Visibility = 'Collapsed'
        $BtnEdit.Content    = 'Preview'
        $Editor.Focus() | Out-Null
        Set-Status "Editing $($script:CurrentFile)"
    } else {
        $Editor.Visibility  = 'Collapsed'
        $Preview.Visibility = 'Visible'
        $BtnEdit.Content    = 'Edit'
        Update-Preview
        Set-Status $(if ($script:CurrentFile) { $script:CurrentFile } else { 'Ready' })
    }
    $script:HitIndex = 0
    Update-HitUi
    if ($BtnEdit.IsChecked -ne $Enabled) { $BtnEdit.IsChecked = $Enabled }
}

function Get-DefaultRoot {
    <# Uses <OneDrive for Business>\customer-brain when that folder exists locally. #>
    if (-not $env:OneDriveCommercial) { return $null }
    $p = Join-Path $env:OneDriveCommercial $script:DefaultFolder
    if (Test-Path -LiteralPath $p -PathType Container) { return (Resolve-Path -LiteralPath $p).Path }
    $null
}

function Get-UserHomeRoot {
    if ($HOME -and (Test-Path -LiteralPath $HOME -PathType Container)) {
        return (Resolve-Path -LiteralPath $HOME).Path
    }
    $profilePath = [Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)
    if ($profilePath -and (Test-Path -LiteralPath $profilePath -PathType Container)) {
        return (Resolve-Path -LiteralPath $profilePath).Path
    }
    $null
}

function Get-RememberedRoot {
    if (-not (Test-Path -LiteralPath $script:SettingsPath -PathType Leaf)) { return $null }
    try {
        $settings = Get-Content -LiteralPath $script:SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($settings.LastRoot -and (Test-Path -LiteralPath $settings.LastRoot -PathType Container)) {
            return (Resolve-Path -LiteralPath $settings.LastRoot).Path
        }
    } catch { }
    $null
}

function Save-RememberedRoot {
    param([string]$FolderPath)

    try {
        $settingsFolder = Split-Path -Parent $script:SettingsPath
        if (-not (Test-Path -LiteralPath $settingsFolder -PathType Container)) {
            New-Item -ItemType Directory -Path $settingsFolder -Force | Out-Null
        }
        @{ LastRoot = $FolderPath } | ConvertTo-Json | Set-Content -LiteralPath $script:SettingsPath -Encoding UTF8
    } catch { }
}

function Select-RootFolder {
    param([string]$Initial)

    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Select the folder that contains your Markdown files'
    $dlg.ShowNewFolderButton = $false
    if ($Initial -and (Test-Path -LiteralPath $Initial)) { $dlg.SelectedPath = $Initial }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath }
    $null
}

#endregion

#region -------------------------------------------------------------- events

$Tree.Add_SelectedItemChanged({
    if ($script:SuppressSelect) { return }
    $item = $Tree.SelectedItem
    if (-not $item -or -not $item.Tag) { return }
    $tag = $item.Tag

    if (-not (Confirm-PendingChanges)) { return }

    Clear-History
    if ($tag.Type -eq 'dir') {
        Open-FolderInfo -FolderPath $tag.Path
    } elseif ($tag.Type -eq 'file' -and $tag.Path -ne $script:CurrentFile) {
        if (Test-MarkdownFile $tag.Path) { Open-MarkdownFile -FilePath $tag.Path }
        else { Invoke-ExternalTarget $tag.Path }
    }
})

$Preview.Add_Navigating({
    param($sender, $e)

    if (-not $e.Uri) { return }
    $u = $e.Uri.AbsoluteUri

    if ($u -like 'about:*') { return }                       # initial load / in page anchors

    if ($u.StartsWith($script:LinkHost, [StringComparison]::OrdinalIgnoreCase)) {
        $e.Cancel = $true
        $query = [System.Web.HttpUtility]::ParseQueryString($e.Uri.Query)
        if ($query['toggle'] -eq '1') {
            $lineIndex = 0
            $script:PendingToggle = @{
                Line    = $null
                Checked = $query['checked'] -eq '1'
            }
            if (-not [int]::TryParse($query['line'], [ref]$lineIndex)) {
                Set-Status 'Could not read the checkbox line number'
                $script:PendingToggle = $null
                return
            }
            $script:PendingToggle.Line = $lineIndex
            $win.Dispatcher.BeginInvoke(
                [System.Windows.Threading.DispatcherPriority]::Background,
                [action] {
                    $toggle = $script:PendingToggle
                    $script:PendingToggle = $null
                    if ($toggle) { Save-CheckboxState -LineIndex $toggle.Line -Checked $toggle.Checked }
                }) | Out-Null
            return
        }
        # a deferred scriptblock does not close over locals, so the target is handed over in script scope
        $script:PendingJumpUrl = $u
        $script:PendingLink = $query['t']
        if ([string]::IsNullOrWhiteSpace($script:PendingLink)) {
            Set-Status "Could not read the link target from '$u'"
            return
        }
        # deferred so the navigation event is finished before the preview is replaced
        $win.Dispatcher.BeginInvoke(
            [System.Windows.Threading.DispatcherPriority]::ApplicationIdle,
            [action] {
                $t = $script:PendingLink
                $script:PendingLink = $null
                try { Invoke-LinkTarget -Target $t }
                catch { Set-Status "Link error: $($_.Exception.Message)" }
            }) | Out-Null
        return
    }

    $e.Cancel = $true
    Invoke-ExternalTarget $u
})

$Editor.Add_TextChanged({
    if ($script:Loading) { return }
    if ($script:CurrentFile -and -not $script:IsDirty -and $script:IsEditMode) {
        $script:IsDirty = $true
        Update-Title
    }
})

function Invoke-TreeFilter {
    $sel = $script:CurrentFile
    $filter = $TxtFilter.Text.Trim()
    Start-SearchBusy -Term $filter
    try {
        Update-Tree -Filter $filter
        if ($sel) { $script:SuppressSelect = $true; Select-TreeFile -FilePath $sel | Out-Null; $script:SuppressSelect = $false }
        if (-not $filter) {
            Set-Status $(if ($script:CurrentFile) { $script:CurrentFile } else { "Root: $script:RootPath" })
        }
        $script:HitIndex = 0
        if ($script:IsEditMode) { Update-HitUi } else { Update-Preview }
    } finally {
        Stop-SearchBusy
    }
}

$script:FilterTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:FilterTimer.Interval = [TimeSpan]::FromMilliseconds(300)
$script:FilterTimer.Add_Tick({
    if ($script:IsSearching) { return }
    $script:FilterTimer.Stop()
    Invoke-TreeFilter
})

$TxtFilter.Add_TextChanged({
    $script:FilterTimer.Stop()
    $script:FilterTimer.Start()
})

$TxtFilter.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq 'Return') {
        if ($script:FilterTimer.IsEnabled) {
            # the tree/preview still have to be rebuilt - jump once the new document is loaded
            $script:FilterTimer.Stop()
            $script:JumpPending = $true
            Invoke-TreeFilter
        } else {
            Move-Hit 1
        }
        $e.Handled = $true
    } elseif ($e.Key -eq 'Escape') {
        $TxtFilter.Text = ''
        $e.Handled = $true
    }
})

$Preview.Add_LoadCompleted({
    $script:PreviewLoading = $false
    if ($script:PendingAnchor) {
        Start-PendingAnchorJump
    }
    if ($script:JumpPending) {
        $script:JumpPending = $false
        Move-Hit 1
    }
    if ($script:PendingHitId) {
        Start-HitScroll -ElementId $script:PendingHitId
    }
})

$BtnBack.Add_Click({ Invoke-HistoryBack })

$BtnClear.Add_Click({ $TxtFilter.Text = ''; $TxtFilter.Focus() | Out-Null })
$BtnHitUp.Add_Click({ Move-Hit -1 })
$BtnHitDown.Add_Click({ Move-Hit 1 })

$BtnOpen.Add_Click({
    if (-not (Confirm-PendingChanges)) { return }
    $p = Select-RootFolder -Initial $script:RootPath
    if ($p) {
        $script:RootPath = (Resolve-Path -LiteralPath $p).Path
        Save-RememberedRoot -FolderPath $script:RootPath
        $script:Standalone.Clear()
        $script:ContentCache.Clear()
        Clear-History
        $TxtFilter.Text = ''
        Update-Tree -Filter ''
        Set-Status "Root: $script:RootPath"
    }
})

function Invoke-TreeRefresh {
    $script:ContentCache.Clear()
    $sel = $script:CurrentFile
    Update-Tree -Filter $TxtFilter.Text.Trim()
    if ($sel) { $script:SuppressSelect = $true; Select-TreeFile -FilePath $sel | Out-Null; $script:SuppressSelect = $false }
    if (-not $TxtFilter.Text.Trim()) { Set-Status 'Tree refreshed' }
}

$BtnRefresh.Add_Click({ Invoke-TreeRefresh })

$BtnExpand.Add_Click({ Set-TreeExpansion ([bool]$BtnExpand.IsChecked) })

$BtnEdit.Add_Click({
    if (-not $script:CurrentFile) { $BtnEdit.IsChecked = $false; return }
    Set-EditMode ([bool]$BtnEdit.IsChecked)
})

$BtnSave.Add_Click({ Save-CurrentFile })

function Show-SelectedTreeItemInExplorer {
    $item = $Tree.SelectedItem
    if (-not $item -or -not $item.Tag -or -not $item.Tag.Path) { return }

    $selectedPath = [string]$item.Tag.Path
    if (Test-Path -LiteralPath $selectedPath -PathType Container) {
        Start-Process explorer.exe "`"$selectedPath`""
    } elseif (Test-Path -LiteralPath $selectedPath -PathType Leaf) {
        Start-Process explorer.exe "/select,`"$selectedPath`""
    } else {
        Set-Status "Path not found: $selectedPath"
    }
}

$BtnReveal.Add_Click({ Show-SelectedTreeItemInExplorer })

$win.Add_KeyDown({
    param($sender, $e)
    $ctrl = [System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control
    $alt  = [System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Alt
    if ($alt -and ($e.SystemKey -eq 'Left' -or $e.Key -eq 'Left')) { Invoke-HistoryBack; $e.Handled = $true; return }
    if ($ctrl -and $e.Key -eq 'S')  { Save-CurrentFile; $e.Handled = $true; return }
    if ($e.Key -eq 'F3') {
        $shift = [System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift
        Move-Hit $(if ($shift) { -1 } else { 1 })
        $e.Handled = $true
        return
    }
    if ($e.Key -eq 'F5')            { Invoke-TreeRefresh; $e.Handled = $true; return }
    if ($e.Key -eq 'F4' -and $script:CurrentFile) { Set-EditMode (-not $script:IsEditMode); $e.Handled = $true; return }
})

$win.Add_Closing({
    param($sender, $e)
    if (-not (Confirm-PendingChanges)) { $e.Cancel = $true }
})

#endregion

#region --------------------------------------------------------------- start

$startupNote = $null
$pathWasSpecified = $PSBoundParameters.ContainsKey('Path') -and -not [string]::IsNullOrWhiteSpace($Path)
if ($pathWasSpecified) {
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        $invalidPath = $Path
        $Path = Get-UserHomeRoot
        $startupNote = "The requested folder '$invalidPath' was not found. Loaded the user home directory instead."
    }
} else {
    $Path = Get-RememberedRoot
    if ($Path) { $startupNote = 'Loaded the last used folder.' }
    else {
        $Path = Get-DefaultRoot
        if ($Path) { $startupNote = "Loaded the default folder '$script:DefaultFolder' from OneDrive for Business." }
        else {
            $Path = Get-UserHomeRoot
            $startupNote = "The default folder '$script:DefaultFolder' was not found. Loaded the user home directory instead."
        }
    }
}
if (-not $Path) { return }

$script:RootPath = (Resolve-Path -LiteralPath $Path).Path
Save-RememberedRoot -FolderPath $script:RootPath
Update-Tree -Filter ''
Set-Status $(if ($startupNote) { "$startupNote  Root: $script:RootPath" } else { "Root: $script:RootPath" })
Update-Title
$Preview.NavigateToString(@"
<html><head><meta http-equiv="X-UA-Compatible" content="IE=edge" /><style>$script:PreviewCss</style></head>
<body><h1>MD-Browser</h1>
<p>Select a Markdown file in the tree on the left.</p>
<ul>
  <li><b>Search box</b> filters the tree by file and folder name; from $script:ContentMinChars characters on the file content is searched too and the number of hits is shown behind the file name.</li>
  <li>Hits in the displayed file are highlighted in yellow. The counter next to the search box shows <i>current / total</i>, the arrows (or <b>F3 / Shift+F3</b>) jump from hit to hit and <b>&#x2715;</b> or <b>Esc</b> clears the search.</li>
  <li><b>Expand all / Collapse all</b> opens or closes the complete tree.</li>
  <li><b>Edit / F4</b> switches between the rendered preview and the text editor.</li>
  <li><b>Ctrl+S</b> saves the current file.</li>
  <li><b>F5</b> rescans the folder structure.</li>
  <li>Links to local files outside of the root folder appear under <b>Standalone files</b>.</li>
  <li>Following a link records a history - the <b>Back</b> button above the document (or <b>Alt+Left</b>) returns to the previous file. Selecting a file in the tree starts a new history.</li>
  <li>Links like <code>file.md#section</code> or <code>#section</code> scroll straight to that heading.</li>
</ul>
<p>Root folder: <code>$script:RootPath</code></p>
</body></html>
"@)

$rootNode = $Tree.Items[0]
if ($rootNode -and $rootNode.Items.Count -gt 0) {
    $firstNode = $rootNode.Items[0]
    $rootNode.IsExpanded = $true
    $firstNode.IsExpanded = $true

    if ($firstNode.Tag.Type -eq 'file') {
        $script:SuppressSelect = $true
        $firstNode.IsSelected = $true
        $script:SuppressSelect = $false
        Open-MarkdownFile -FilePath $firstNode.Tag.Path
    }
}

$win.Visibility = [System.Windows.Visibility]::Visible
$win.ShowActivated = $true
$win.Add_Loaded({
    $win.Activate() | Out-Null
    $win.Topmost = $true
    $win.Topmost = $false
})
$win.ShowDialog() | Out-Null

#endregion
