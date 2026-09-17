# MD-Browser

A lightweight Windows desktop application for browsing, searching, previewing, and editing folders of Markdown files.

MD-Browser is a companion viewer/editor for [SecondBrain](https://github.com/oweindl/SecondBrain), an agent-agnostic Markdown knowledge-space pattern. It can browse any Markdown folder, and it works especially well with SecondBrain structures such as:

```text
<RootFolder>\SecondBrain\<BrainName>\context.md
<RootFolder>\SecondBrain\<BrainName>\startup.md
<RootFolder>\SecondBrain\index.md
```

## Features

- Folder tree with file-name and content search
- Rendered Markdown preview with highlighted search matches
- Previous and next match navigation
- Plain-text editing and saving
- Local Markdown link navigation and section jumps
- External-link support
- GitHub-style task checkbox toggling
- Remembers the last opened folder between launches
- Checks GitHub for newer versions and offers to install them
- No external runtime dependencies

## SecondBrain compatibility

MD-Browser does not require a SecondBrain folder, but it recognizes the pattern naturally because SecondBrain uses plain Markdown files and local links.

Recommended optional frontmatter for `context.md`:

```yaml
---
brainName: Example Brain
schema: secondbrain-v1
browser: md-browser
---
```

This metadata is optional; MD-Browser remains a general-purpose Markdown folder browser.

## Requirements

- Windows
- Windows PowerShell 5.1 or PowerShell 7+

## Run

From PowerShell:

```powershell
.\MD-Browser.ps1
```

To open a specific folder:

```powershell
.\MD-Browser.ps1 -Path "C:\path\to\markdown"
```

To print the installed version:

```powershell
.\MD-Browser.ps1 -Version
```

Without `-Path`, MD-Browser reopens the last used folder. If no valid folder was remembered, it uses the configured default folder and then the current user's home directory as fallbacks. An explicit `-Path` always takes precedence and becomes the remembered folder.
