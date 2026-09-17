# MD-Browser

A lightweight Windows desktop application for browsing, searching, previewing, and editing folders of Markdown files.

## Features

- Folder tree with file-name and content search
- Rendered Markdown preview with highlighted search matches
- Previous and next match navigation
- Plain-text editing and saving
- Local Markdown link navigation and section jumps
- External-link support
- GitHub-style task checkbox toggling
- Remembers the last opened folder between launches
- No external runtime dependencies

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
