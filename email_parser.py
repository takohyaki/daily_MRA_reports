import os
import argparse
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from datetime import datetime
import re
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

def parse_email(file_path):
    with open(file_path, 'rb') as file:
        email = BytesParser(policy=policy.default).parse(file)

    # Extract Application/Claim from the subject
    subject = email['Subject']
    if "Claim" in subject:
        application_claim = "Claim"
    elif "Application" in subject:
        application_claim = "Application"
    else:
        application_claim = 'N/A'

    # Use email's date header for Date Received and parse it
    date_received_header = email['Date']
    date_received = parsedate_to_datetime(date_received_header).replace(tzinfo=None) if date_received_header else None

    # Calculate days since received
    today = datetime.today()
    days_since_received = 'N/A'
    if date_received:
        days_since_received = (today - date_received).days

    # Extract other details as before
    body = email.get_body(preferencelist=('plain',)).get_content()

    # Extract specific details
    company_name_match = re.search(r'Company Name\s*\n\s*(.+)', body)
    company_name = company_name_match.group(1).strip() if company_name_match else 'N/A'

    reference_id_match = re.search(r'Ref ID\s*\n\s*(\S+)', body)
    reference_id = reference_id_match.group(1).strip() if reference_id_match else 'N/A'

    project_title_match = re.search(r'Project Title\s*\n\s*(.+)', body)
    project_title = project_title_match.group(1).strip() if project_title_match else 'N/A'

    deadline_match = re.search(r'We would appreciate a response by (\d+ \w+ \d+)', body)
    deadline = deadline_match.group(1) if deadline_match else 'N/A'

    # Extract Queries
    queries_start_match = body.find("Please clarify the reason why.")
    queries_end_match = body.find("Thanks. We would appreciate a response", queries_start_match)
    queries = body[queries_start_match:queries_end_match].strip() if queries_start_match != -1 and queries_end_match != -1 else 'N/A'

    return {
        "Company Name": company_name,
        "Reference ID": reference_id,
        "Project Title": project_title,
        "Application/Claim": application_claim,
        "Date Received": date_received_header,
        "# of Days Since Received": days_since_received,
        "Deadline": deadline,
        "Queries": queries
    }

def save_to_excel(extracted_data, excel_file_path):
    # Create a DataFrame with the specified column order
    column_order = ["Company Name", "Reference ID", "Project Title", "Application/Claim", 
                    "Date Received", "Deadline", "Status", "# of Days Since Received", "Remarks"]
    df = pd.DataFrame([extracted_data], columns=column_order)

    # Add empty columns for 'Status' and 'Remarks' if not already present
    for col in ['Status', 'Remarks']:
        if col not in df.columns:
            df[col] = ''

    # Format current date as string for the sheet name
    current_date_str = datetime.now().strftime("%Y-%m-%d")
    sheet_name = current_date_str

    # Save to Excel file
    with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Adjust column widths
        workbook = writer.book
        worksheet = workbook[sheet_name]
        for column in worksheet.columns:
            max_length = max(len(str(cell.value)) for cell in column)
            adjusted_width = (max_length + 2) * 1.2  # Adjust the multiplier as needed
            worksheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Process emails and save data to an Excel file.')
    parser.add_argument('email_file', help='Relative path to the email file from the script directory (e.g., sample_emails/email.eml)')
    parser.add_argument('excel_file', help='Relative path to the Excel file from the script directory (e.g., output_data/excel_file.xlsx)')

    # Parse arguments
    args = parser.parse_args()

    # Get the script's directory
    script_dir = os.path.dirname(os.path.realpath(__file__))

    # Construct full paths for the email and Excel files
    full_email_path = os.path.join(script_dir, args.email_file)
    full_excel_path = os.path.join(script_dir, args.excel_file)

    # Run the main functionality
    extracted_data = parse_email(full_email_path)
    save_to_excel(extracted_data, full_excel_path)
    print(f"Data from {full_email_path} has been saved to {full_excel_path}")

if __name__ == "__main__":
    main()

# Example usage
# file_path = "/sample_emails/FW_ Updates required_ Claim 2302MFKX-CL01 on hold.eml"
# extracted_data = parse_email(file_path)
# print(extracted_data)
# excel_file_path = "/sample_file/MRA_daily_updates.xlsx"
# save_to_excel(extracted_data, excel_file_path)


