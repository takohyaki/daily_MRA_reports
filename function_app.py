import logging
import azure.functions as func
import re
from datetime import datetime, timedelta
from auth import get_access_token, tenant_id, client_id, client_secret
from graph_api import get_outlook_emails
from email_parser import parse_email

user_email = "b7efe171-fa97-4766-9917-a2848f0fea3f"  # Daphne Toh's email

def email_filter(subject, patterns):
    return any(re.search(pattern, subject) for pattern in patterns)

app = func.FunctionApp()

@app.schedule(schedule="0 0 14 * * *", arg_name="myTimer", use_monitor=False)
def timer_trigger(myTimer: func.TimerRequest) -> None:
    utc_now = datetime.utcnow()
    yesterday = utc_now - timedelta(days=1)
    last_run = yesterday.replace(hour=22, minute=0, second=0, microsecond=0)

    if myTimer.past_due:
        logging.info('The timer is past due!')

    access_token = get_access_token(tenant_id, client_id, client_secret)
    emails_dict = get_outlook_emails(access_token, user_email)
    emails_list = emails_dict.get('value', [])

    patterns = [
        r"Updates required: Grant application.*on hold",
        r"Updates required: Claim.*on hold",
        r"Updates required: Change request.*on hold"
    ]

    filtered_emails = [email for email in emails_list if email_filter(email.get('subject', ''), patterns)]
    for email in filtered_emails:
        extracted_data = parse_email(email)
        logging.info(extracted_data)
