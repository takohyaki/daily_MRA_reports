# Automated Daily MRA Reports

## Example Workflow
1. User uploads email files through the GUI.
2. The application parses the emails and stores the data in the SQL database.
3. The user can then export this data to Excel with the click of a button.

##  Command Line
```
python your_script_name.py path_to_email_file.eml path_to_excel_file.xlsx
```
For example,
```
python email_parser.py "sample_emails/FW_ Updates required_ Claim 2302MFKX-CL01 on hold.eml" "sample_file/MRA_daily_updates.xlsx"
``````
