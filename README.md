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

If the requested folder or configured default folder does not exist, MD-Browser starts in the current user's home directory.

## Tests

Run the available PowerShell harnesses from the project folder:

```powershell
.\_test-default-startup.ps1
.\_test-search-navigation.ps1
```

The close-behavior harness is interactive:

```powershell
.\_test-current-close.ps1
```
