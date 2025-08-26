# 📂 VBA File & Email Automation

This module automates the process of creating subfolders, moving/renaming files, managing grouped emails, and sending files via Outlook. It also includes cleanup routines (archiving, moving to DONE, deleting subfolders).

## 🚀 Features
Folder & File Operations
- CreateSubFolder → Creates subfolders based on cell values from the Data sheet.
- FSOMoveFile → Moves files from original location to designated folder.
- FSOReverseMoveFile → Moves files back to original location (undo).
- MovePDFsToAnotherFolder → Moves .xlsx files without email info into a No Email folder.
- Move_Folder → Copies all files and subfolders to a final DONE folder.
- DeleteSubfolders → Deletes all subfolders inside a target directory.

## Data Preparation
- CopytoGroupEmail → Copies raw file list (Data sheet) into GroupEmail sheet.
- Remove_Duplicates_Folder → Removes duplicates from folder records.
- DeleteErrorCountry → Removes rows with "00 NOT FOUND" errors.

## Email Automation (Outlook Required)
- SendtoEmail
- Sends emails per customer with their respective attachments.
- Reads Email Address (Col J), Subject (Col K), and Body (Col L) from GroupEmail sheet.
- Automatically attaches all .xlsx files from the specified folder.
- Prompts user to confirm verification and Outlook readiness before sending.

## 📊 Input / Output Table
| Sheet / Cell          | Purpose                              |
| ---------------------- | ------------------------------------ |
| Dashboard!C21          | Source folder path                   |
| Dashboard!C22          | Loop limit for folder/file creation  |
| Dashboard!C23          | Loop limit for email send            |
| Dashboard!C24          | Destination folder for "DONE" move   |
| Dashboard!C16          | Project/Batch folder name            |
| Data!C:F               | Source → Destination file paths      |
| GroupEmail!J:K:L       | Email To, Subject, Body              |

## 🛠 Requirements
- Excel with VBA enabled.
- Microsoft Outlook installed and open.
- File system access to paths defined in Dashboard.

## ⚡ Example Workflow
1. Run CreateSubFolder → Generates subfolders and moves files.
2. Run SendtoEmail → Sends grouped emails with attached files.
3. Run Move_Folder → Moves processed files into the DONE archive.
4. Run DeleteSubfolders → Cleans up temporary folders.
