Google Drive Permissions Auditor
1. Overview

The Google Drive Permissions Auditor is an automated tool designed to scan, analyze, and manage Google Drive folder permissions via Google Sheets. It addresses the limitations of standard Google Apps Script execution time by utilizing a recursive, state-preserving mechanism. This allows the tool to process large folder structures without interruption.

The application serves two primary functions:

    Auditing: It generates a comprehensive report of all files and folders within a target directory, listing every user with access and their specific permission role.
    Management: It allows administrators to modify permissions (add, remove, or change roles) directly within the spreadsheet and synchronize those changes back to Google Drive.

2. Technical Features

    Recursive Scanning with State Persistence: The script uses a queue-based system to track scanning progress. When the Google Apps Script 6-minute execution limit approaches, the script saves its state to a hidden sheet and schedules a trigger to resume automatically.
    Batched Write Operations: To optimize performance, data is buffered in memory and written to the spreadsheet in batches, significantly reducing the time required to populate the report.
    Permission Synchronization: The tool compares the "Current Role" against the "Adjusted Permission" column. If discrepancies are found, it updates the Google Drive permissions to match the spreadsheet configuration.
    Safe Permission Handling: The script includes logic to handle permission downgrades (e.g., changing an Editor to a Viewer) by first removing the higher-level role. It also prevents the accidental removal of file owners.

3. Installation

    Create a new Google Sheet.
    Navigate to Extensions > Apps Script.
    Remove any existing code in the default Code.gs file.

    Create the following script files within the editor and paste the corresponding code into each:

        Config.gs
        Menu.gs
        Audit_Engine.gs
        Audit_Helpers.gs
        Sync_Manager.gs
        User_Tools.gs
        Utilities.gs
    Save the project.
    Reload the Google Sheet. A new menu item labeled "Drive Audit" will appear in the toolbar.

4. Usage Instructions
4.1. Initiating an Audit

    Select Drive Audit > 1. Start New Audit from the menu.
    Input the URL of the target Google Drive folder when prompted.
    The script will begin scanning. Progress is maintained via background triggers; the browser tab may be closed without interrupting the process.

4.2. Reviewing Results

Upon completion, the Audit Results sheet will populate with the following columns:

    Type: Indicates if the item is a File or Folder.
    Path/Folder: The directory path of the item.
    Item Name: The name of the file or folder.
    User Email: The account holding access rights.
    Current Role: The access level currently assigned (Owner, Editor, Viewer, or Commenter).

4.3. Modifying Permissions

To alter access rights:

    Locate the Adjusted Permission column (Column G).
    Select the desired role from the dropdown menu (e.g., Viewer, Remove Access).
    Once all adjustments are set in the spreadsheet, select Drive Audit > 5. Sync Changes.
    The script will process the changes and update Google Drive accordingly.

4.4. Adding New Users

To grant access to a user not currently listed:

    Select the row containing the target file.
    Navigate to Drive Audit > 4. Add New User to Selected File.
    Enter the email address of the new user.
    A new row will be generated. Set the desired permission in Column G.
    Run Sync Changes to apply the new permission.

5. Troubleshooting

    Script Termination: If the script halts unexpectedly due to API errors, select Drive Audit > 2. Force Resume. This reads the hidden queue sheet and resumes scanning from the last recorded position.
    Triggers: The script generates temporary time-based triggers to manage the execution loop. These are automatically removed upon completion. To manually halt the process and clear triggers, select Drive Audit > 3. Stop/Cancel Auto-Audit.

6. Disclaimer

This software modifies permissions on live Google Drive files. While safety mechanisms are implemented to prevent the removal of owners, users should exercise caution when performing bulk synchronization operations. Verify all "Adjusted Permission" values before syncing.


Written by Chad Jacks
Chad.Jacks@k21schools.org
2026-02-13
