Google Drive Permissions Auditor
1. Overview

The Google Drive Permissions Auditor is an automated tool designed to scan, analyze, and manage Google Drive folder permissions via Google Sheets. It addresses the limitations of standard Google Apps Script execution time by utilizing a recursive, state-preserving mechanism. This allows the tool to process large folder structures without interruption.

The application serves two primary functions:

    Auditing: It generates a comprehensive report of all files and folders within a target directory, listing every user with access and their specific permission role.
    Management: It allows administrators to perform granular or bulk modifications to permissions directly within the spreadsheet and synchronize those changes back to Google Drive.

2. Technical Features

    Recursive Scanning with State Persistence: The script uses a queue-based system to track scanning progress. When the Google Apps Script execution limit approaches, the script saves its state and schedules a trigger to resume automatically.
    Bulk Processing Logic: The tool is designed to handle mass updates to the "Adjusted Permission" column without triggering data validation errors, specifically when handling the removal of users.
    Integrity-Safe Row Insertion: When adding users to multiple files simultaneously, the script processes the selection in descending order (bottom-up). This ensures that row indices remain accurate as new rows are inserted.
    Security Guardrails: The script includes logic to prevent the modification or removal of file "Owners." It also validates inputs to ensure only recognized roles (Editor, Viewer, Commenter) are applied.

3. Installation

    Create a new Google Sheet.
    Navigate to Extensions > Apps Script.
    Remove any existing code in the default Code.gs file.
    Create the following script files within the editor and paste the corresponding code into each:

        config.gs
        menu.gs
        audit_main.gs
        Audit_helpers.gs
        sync_main.gs
        user_main.gs
        utilities.gs

    Save the project and reload the Google Sheet. A menu labeled "Drive Audit" will appear in the toolbar.

4. Usage Instructions
4.1. Initiating an Audit

    Select Drive Audit > 1. Start New Audit.
    Input the URL of the target Google Drive folder.
    The script will begin scanning. Progress is maintained via background triggers; the browser tab may be closed during this process.

4.2. Bulk Editing by Selection

To update multiple specific rows at once:

    Highlight the desired rows or cells (use Ctrl/Cmd or Shift for multi-selection).
    Select Drive Audit > 6. Bulk Edit Permissions (Selected Rows).
    Enter the corresponding letter for the new role:

        E: Editor
        V: Viewer
        C: Commenter
        R: Remove Access

4.3. Bulk Editing by Email (Find and Replace)

To change permissions for a specific user across the entire audit:

    Select Drive Audit > 7. Bulk Edit by Email (Find & Replace).
    Enter the user's exact email address.
    Enter the letter for the new role (E, V, C, or R).
    The script will update every instance of that user found in the sheet, excluding files they own.

4.4. Adding a User to Multiple Files

    Select the rows representing the files you wish to share.
    Select Drive Audit > 4. Add New User to Selected File.
    Enter the email address of the new user.
    Enter the letter for their initial permission level (E, V, or C).
    New rows (highlighted yellow) will be inserted beneath each selected file with the specified settings.

4.5. Synchronizing Changes

Updates made in the spreadsheet are not applied to Google Drive until synchronized.
    Review the Adjusted Permission column to ensure accuracy.
    Select Drive Audit > 5. Sync Changes (Update Drive).
    The script will process the updates and change the "Current Role" status to reflect the successful modification.

5. Troubleshooting

    Data Validation Errors: If a cell displays a validation error, it is likely because a role was entered that is not supported by the dropdown menu. Ensure only "Editor," "Viewer," "Commenter," or "Remove Access" are present in Column G.
    Script Interruption: If the scan halts, select Drive Audit > 2. Force Resume.
    Manual Halt: To stop an active audit and clear all background triggers, select Drive Audit > 3. Stop/Cancel Auto-Audit.

Written by Chad Jacks
Chad.Jacks@k21schools.org
2026-02-13
