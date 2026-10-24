CLAIMANT DROP — WINDOWS

INSTALL ONCE
1. Extract this ZIP file.
2. Double-click "Install Claimant Drop.cmd".
3. The installer copies the tool into your user app folder and automatically adds "Claimant Drop" to Send To and your desktop.
4. You can delete the extracted download after installation. Run the installer again to update the installed tool.

USE IT
1. Right-click one .xlsx report, choose Send to, then Claimant Drop.
   If needed, choose Show more options first to find Send to.
   Alternatively, drag the report onto the desktop shortcut.
2. Wait for the confirmation box: True Full or True Partial (or Needs review for unrecognized answers).
3. The deduplicated table is already copied. Press Ctrl+V wherever you want it.
4. A CSV copy is also saved beside the original report.

FILTER RULES
- Column E accepts YES, Yes, Y, and (Y), ignoring capitalization and spaces.
- All Yes = True Full; any No (including N and (N)) = True Partial; blanks are ignored.
- One report represents one transfer application and is expected to contain at least one Yes.
- Clipboard and CSV headers: Admin, Admin P#, Publisher, Publisher P#.
- Source columns: R, Q, N, M. Member fields are excluded.
- To use a different report layout, put a claimant-columns.json file next to the installer before installing.
  See claimant-columns.example.json in the repository.
- Duplicate four-column claimant rows are removed.
- Existing Send To shortcuts can continue to call the launcher with the report path.

REQUIREMENTS
- Windows with desktop Excel installed.
- If your employer blocks PowerShell scripts, ask IT to approve the extracted files.
