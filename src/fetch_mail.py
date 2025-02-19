import os
import win32com.client
from datetime import datetime, timedelta
import pytz

def fetch_outlook_emails():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        sent_items = outlook.GetDefaultFolder(5)  # 5 represents the Sent Items folder

        # Get the local timezone
        local_tz = pytz.timezone('Etc/UTC')  # Replace with your local timezone if needed

        one_week_ago = datetime.now(local_tz) - timedelta(days=7)

        script_dir = os.path.dirname(os.path.abspath(__file__))
        docs_dir = os.path.join(script_dir, 'docs')
        if not os.path.exists(docs_dir):
            os.makedirs(docs_dir)

        items = sent_items.Items
        items.Sort("[SentOn]", True)

        emails_saved = 0
        for item in items:
            if item.Class == 43:  # 43 indicates a Mail Item
                try:
                    sent_date = item.SentOn.replace(tzinfo=None)
                    sent_date = local_tz.localize(sent_date)
                    if sent_date < one_week_ago:
                        break  # Stop processing older emails

                    subject = item.Subject if item.Subject else "No_Subject"
                    safe_subject = "".join([c for c in subject if c.isalnum() or c in (' ', '-', '_')]).rstrip()
                    sent_date_str = sent_date.strftime("%Y%m%d_%H%M%S")
                    filename = f"{sent_date_str}_{safe_subject[:50]}"

                    # Extract email content and format it into HTML
                    html_content = f"""
                    <html>
                    <head>
                    </head>
                    <body>
                    <h2>Subject: {subject}</h2>
                    <p>From: {item.SenderName}</p>
                    <p>Sent: {sent_date_str}</p>
                    <hr>
                    {item.HTMLBody if item.HTMLBody else item.Body}
                    </body>
                    </html>
                    """

                    html_path = os.path.join(docs_dir, f"{filename}.html")
                    with open(html_path, 'w', encoding='utf-8') as html_file:
                        html_file.write(html_content)

                    print(f"Saved: {filename}.html")
                    emails_saved += 1
                except Exception as save_error:
                    print(f"Failed to save email: {save_error}")

        if emails_saved == 0:
            print("No emails were saved. Check if there are any sent emails within the last 7 days.")
        else:
            print(f"Total emails saved: {emails_saved}")

    except Exception as e:
        print(f"An error occurred: {e}")

    print("Email fetching process completed.")

if __name__ == "__main__":
    fetch_outlook_emails()
