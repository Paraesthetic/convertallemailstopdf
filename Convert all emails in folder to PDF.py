import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog

# Function to install packages
def install(package):
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])

# Install necessary packages if not already installed
try:
    import extract_msg
except ImportError:
    install('extract-msg')
    import extract_msg

try:
    from xhtml2pdf import pisa
except ImportError:
    install('xhtml2pdf')
    from xhtml2pdf import pisa

try:
    from bs4 import BeautifulSoup
except ImportError:
    install('beautifulsoup4')
    from bs4 import BeautifulSoup

def process_msg_files(folder_path, pdf_folder):
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.msg'):
            filepath = os.path.join(folder_path, filename)
            try:
                msg = extract_msg.Message(filepath)
                msg_sender = msg.sender
                msg_to = msg.to
                msg_date = msg.date
                msg_subject = msg.subject
                msg_body = msg.body

                # Clean the body text
                msg_body = msg_body.replace('\r\n', '<br>')

                # Build an HTML representation
                html = f"""
                <html>
                <head>
                <meta charset="utf-8">
                <style>
                    body {{
                        font-family: Calibri, sans-serif;
                        font-size: 11pt;
                        margin: 1in;
                    }}
                    .header {{
                        margin-bottom: 1em;
                    }}
                    .header p {{
                        margin: 0;
                    }}
                    hr {{
                        border: none;
                        border-top: 1px solid #000;
                        margin: 1em 0;
                    }}
                    .body-content {{
                        white-space: pre-wrap;
                        font-size: 11pt;
                    }}
                </style>
                </head>
                <body>
                    <div class="header">
                        <p><b>From:</b> {msg_sender}</p>
                        <p><b>To:</b> {msg_to}</p>
                        <p><b>Sent:</b> {msg_date}</p>
                        <p><b>Subject:</b> {msg_subject}</p>
                    </div>
                    <hr>
                    <div class="body-content">
                        {msg_body}
                    </div>
                </body>
                </html>
                """

                # Output PDF file path
                pdf_filename = os.path.splitext(filename)[0] + '.pdf'
                pdf_filepath = os.path.join(pdf_folder, pdf_filename)

                # Convert HTML to PDF
                with open(pdf_filepath, 'wb') as pdf_file:
                    pisa_status = pisa.CreatePDF(html, dest=pdf_file)

                if pisa_status.err:
                    print(f"Failed to create PDF for {filename}")
                else:
                    print(f"Created PDF for {filename}")

            except Exception as e:
                print(f"Failed to process {filename}: {e}")

def process_eml_files(folder_path, pdf_folder):
    import email
    from email import policy
    from email.parser import BytesParser

    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.eml'):
            filepath = os.path.join(folder_path, filename)
            try:
                with open(filepath, 'rb') as fp:
                    msg = BytesParser(policy=policy.default).parse(fp)
                msg_from = msg['from']
                msg_to = msg['to']
                msg_date = msg['date']
                msg_subject = msg['subject']

                # Get the email body
                if msg.is_multipart():
                    body = ''
                    for part in msg.walk():
                        ctype = part.get_content_type()
                        cdispo = str(part.get('Content-Disposition'))

                        # Skip any attachments
                        if ctype == 'text/plain' and 'attachment' not in cdispo:
                            body = part.get_content()
                            break
                else:
                    body = msg.get_payload(decode=True).decode('utf-8', errors='replace')

                # Clean the body text
                body = body.replace('\r\n', '<br>')

                # Build an HTML representation
                html = f"""
                <html>
                <head>
                <meta charset="utf-8">
                <style>
                    body {{
                        font-family: Calibri, sans-serif;
                        font-size: 11pt;
                        margin: 1in;
                    }}
                    .header {{
                        margin-bottom: 1em;
                    }}
                    .header p {{
                        margin: 0;
                    }}
                    hr {{
                        border: none;
                        border-top: 1px solid #000;
                        margin: 1em 0;
                    }}
                    .body-content {{
                        white-space: pre-wrap;
                        font-size: 11pt;
                    }}
                </style>
                </head>
                <body>
                    <div class="header">
                        <p><b>From:</b> {msg_from}</p>
                        <p><b>To:</b> {msg_to}</p>
                        <p><b>Sent:</b> {msg_date}</p>
                        <p><b>Subject:</b> {msg_subject}</p>
                    </div>
                    <hr>
                    <div class="body-content">
                        {body}
                    </div>
                </body>
                </html>
                """

                # Output PDF file path
                pdf_filename = os.path.splitext(filename)[0] + '.pdf'
                pdf_filepath = os.path.join(pdf_folder, pdf_filename)

                # Convert HTML to PDF
                with open(pdf_filepath, 'wb') as pdf_file:
                    pisa_status = pisa.CreatePDF(html, dest=pdf_file)

                if pisa_status.err:
                    print(f"Failed to create PDF for {filename}")
                else:
                    print(f"Created PDF for {filename}")

            except Exception as e:
                print(f"Failed to process {filename}: {e}")

def main():
    # Hide the root window
    root = tk.Tk()
    root.withdraw()

    # Ask the user to select a folder
    folder_path = filedialog.askdirectory(title='Select Folder Containing Emails')

    if not folder_path:
        print("No folder selected.")
        return

    # Create 'email pdf' subfolder if it doesn't exist
    pdf_folder = os.path.join(folder_path, 'email pdf')
    os.makedirs(pdf_folder, exist_ok=True)

    print(f"Processing emails in folder: {folder_path}")
    print(f"PDFs will be saved in: {pdf_folder}")

    process_msg_files(folder_path, pdf_folder)
    process_eml_files(folder_path, pdf_folder)

if __name__ == "__main__":
    main()
