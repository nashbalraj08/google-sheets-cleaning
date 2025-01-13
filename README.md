# Google Sheets Cleaning Scripts

This repository contains Google Apps Script code snippets for cleaning and processing data in Google Sheets.

## Clean Truncated Emails

This spreadsheet organizes and refines raw data, particularly focusing on truncated emails. The sheet tab named **"Clean Truncated Emails"** contains emails that need to be corrected, such as `nashbalraj@g.....`. The purpose of the macro is to:

- Clean and correct the data (e.g., convert `nashbalraj@g.....` to `nashbalraj@gmail.com`).
- Eliminate duplicate emails.
- Remove unnecessary white spaces to ensure data accuracy.

### Steps to Use the Macro

1. Open the spreadsheet using the following link:  
   [Google Sheets Link](https://docs.google.com/spreadsheets/d/1QSPQi9cmtxP3WHHuoNMSGs2gHV2pMApplq4PIb9DWyQ/edit?usp=sharing).

2. Log in with your Gmail account to access the spreadsheet.

3. Familiarize yourself with the sheets in the spreadsheet. It contains the following tabs:
   - **Clean Truncated Emails**  
     This is the main sheet where the dummy data is displayed. Ensure you are on this sheet when running the macro.
   - **Dummy Data Truncated Emails**  
     If you want to rerun the macro with the dummy data, this sheet contains the original unprocessed data for testing.
   - **Truncated Emails**  
     For emails that are difficult to fix, this sheet contains predefined corrections. For example:
       - Support requests that provide multiple versions of emails.
       - Cases where users provide corrected emails themselves.

4. Navigate to **Extensions** in the toolbar, then select **Macros** and click on the macro named **"cleanUp"**.

5. If you do not see the macro **"cleanUp"** in the list:
   - Go to **Extensions** → **Apps Script**.
   - Open the Apps Script editor and run the macro directly from there.

6. When prompted, a pop-up will appear asking for permission to execute the macro. Ensure you grant the necessary permissions.

7. Return to the **Clean Truncated Emails** tab to view the cleaned and processed data.

---

### Key Features of the Macro

- Corrects truncated emails to display the full address.
- Removes duplicate entries to maintain a clean dataset.
- Trims unnecessary white spaces from all cells.

By following the above steps, you can ensure that your spreadsheet data is accurate, consistent, and ready for use.
