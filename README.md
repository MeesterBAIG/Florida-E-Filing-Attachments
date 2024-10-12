# Florida E-Filing E-Mail Download

## Overview

This VBA script automates the process of downloading court documents from hyperlinks embedded in selected Outlook emails. It creates a folder structure based on the email subject and organizes the downloaded files accordingly. The script:

1. **Extracts hyperlinks** from the email body that match a specific URL pattern.
2. **Downloads the linked files** into folders named after the email subject (with a portion of the subject removed).
3. **Appends the email received date** to the downloaded file names.
4. **Creates a summary** at the end showing all files that were downloaded.

## Features

- Automatically **creates a base folder** (`C:\CourtDocuments\`) if it doesn't exist.
- Skips the **first hyperlink** in the email, as requested.
- Saves files with the format: `yyyy-mm-dd_HyperlinkText.pdf`.
- **Creates folders** based on the email subject by removing `"SERVICE OF COURT DOCUMENT CASE NUMBER "` from the subject.
- **Final summary popup** displays all downloaded files at the end of the process.
- Handles multiple selected emails, processing each one individually.

## Folder Structure

The downloaded files are organized as follows:

C:\CourtDocuments
|-- [Modified Email Subject 1]
|-- yyyy-mm-dd_HyperlinkText1.pdf |-- yyyy-mm-dd_HyperlinkText2.pdf |-- [Modified Email Subject 2]
|-- yyyy-mm-dd_HyperlinkText1.pdf


The `[Modified Email Subject]` is the email subject with `"SERVICE OF COURT DOCUMENT CASE NUMBER "` removed.

## Prerequisites

- **Microsoft Outlook**: This script works within Microsoft Outlook and interacts with selected emails.
- **VBA Editor**: The script must be placed inside Outlook's VBA editor.

## How to Use

1. **Open Microsoft Outlook**.
2. **Select emails** in your inbox or another folder that contain court document links.
3. **Press `Alt + F11`** to open the VBA editor.
4. **Insert the script** into a new or existing module.
5. **Run the script**:
    - Press `Alt + F8` in Outlook.
    - Select `DownloadFilesToSubjectFolder` from the macro list and click "Run".

The script will then:
- Check if the `C:\CourtDocuments\` folder exists and create it if it doesn't.
- Process each selected email, download the court documents, and store them in folders named after the email subject.
- Display a summary of the downloaded files.

## Configuration

- **Base Folder**: By default, the script saves files in `C:\CourtDocuments\`. You can change this path by modifying the `baseFolder` variable in the script.
  
    ```vba
    baseFolder = "C:\YourCustomPath\"
    ```

- **File Name Format**: The downloaded files are saved using the format `yyyy-mm-dd_HyperlinkText.pdf`. You can adjust the file extension or name format if needed by modifying the `fileName` variable.

## Customization

- **Skipping the First Link**: The script is configured to skip the first hyperlink found in the email. This behavior can be modified or removed by adjusting the `linkCounter` logic.
  
    ```vba
    If linkCounter > 0 Then
        ' Process only if it's not the first link
    End If
    ```

- **Subject Modification**: The folder names are created based on the email subject, with `"SERVICE OF COURT DOCUMENT CASE NUMBER "` removed. You can change this to modify or handle other subject formats.

## Summary

After running the script, a popup will display a summary of all the files that were successfully downloaded. If no files were downloaded, a message will indicate this.

## Troubleshooting

- **No Files Downloaded**: Ensure that the emails contain hyperlinks that match the expected pattern (`https://url.avanan.click/v2/r01/___https://www.myflcourtaccess.com/nefdocuments/document.nefdd?nai=`).
- **Errors**: If there are issues with folder permissions or invalid filenames, ensure that the `baseFolder` path is valid and that no invalid characters are present in the email subject or hyperlink text.

