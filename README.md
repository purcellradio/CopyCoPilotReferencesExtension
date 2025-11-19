# Copy CoPilot References Extension

This Visual Studio extension simplifies referencing files in GitHub Copilot Chat by copying their relative paths to the clipboard, formatted for immediate use.

![CopyCoPilotReferencesExtension](Resources/README/CopyCoPilotReferencesExtension.png)

## How to Use

1.  In the **Solution Explorer**, select one or more files.
2.  Right-click the selected files, or selcted folders in a single project, or a single project, to open the context menu (see caveats below).
3.  Click on **"Copy CoPilot references"**.
4.  The formatted references (e.g., `#file:'YourProject/YourFile.cs'`) are now on your clipboard.
5.  Paste the text directly into the GitHub Copilot Chat window.

## Caveats
- You cannot make selections of folders AND files within a project
- You cannot make selections across multiple projects

## Features

*   **Single & Multi-file Selection**: Copy references for one or multiple files at once.
*   **Correct Formatting**: Automatically formats the paths with the `#file:'...'` syntax required by Copilot Chat.
*   **Relative Paths**: Generates paths relative to the solution directory for cleaner references.
*   **Easy Access**: Integrates directly into the Solution Explorer item context menu for a seamless workflow.

## Revision History

### 2025-11-19 V1.2
- Enhanced copying to include contents of selected folders and an entire project
- Added icon to context menu item to enhance recognition in a busy menu