import subprocess
import os
import sys
import yaml
import xml.etree.ElementTree as ET

def send_mail(from_address, to_address, subject, body, cc=None, bcc=None, attachments=None):
    def escape_quotes(text):
        return text.replace('"', '\\"').replace('\\', '\\\\')

    to_address = escape_quotes(to_address)
    from_address = escape_quotes(from_address)
    subject = escape_quotes(subject)
    body = escape_quotes(body)

    if not from_address:
        print("Provide from_adress in .yaml!")

    elif from_address:
        applescript = f'''
        tell application "Mail"
            set newMessage to make new outgoing message with properties {{subject:"{subject}", content:"{body}", visible:false}}

            tell newMessage
                set sender to "{from_address}"
                make new to recipient at end of to recipients with properties {{address:"{to_address}"}}
        '''

    if cc:
        cc = escape_quotes(cc)
        applescript += f'''
            make new cc recipient at end of cc recipients with properties {{address:"{cc}"}}
    '''

    if bcc:
        bcc = escape_quotes(bcc)
        applescript += f'''
            make new bcc recipient at end of bcc recipients with properties {{address:"{bcc}"}}
    '''

    if attachments:
        for attachment_path in attachments:
            abs_path = os.path.abspath(attachment_path)

            if not os.path.exists(abs_path):
                print(f"Warning: Attachment not found: {abs_path}")
                continue

            escaped_path = escape_quotes(abs_path)
            applescript += f'''
            make new attachment with properties {{file name:POSIX file "{escaped_path}"}} at after the last paragraph
    '''

    applescript += '''
            send
        end tell
    end tell
    '''

    try:
        subprocess.run(['osascript', '-e', applescript],
                       capture_output=True,
                       text=True,
                       check=True,
                       )
        return True

    except subprocess.CalledProcessError as e:
        print(f'''Error while sending mail to {to_address}:/n
              Error: {e}''')
        return False

    except Exception as e:
        print(f'Unexpected error: {e}')
        return False

def load_email_config(yaml_file):
    try:
        with open(yaml_file, 'r') as f:
            config = yaml.safe_load(f)
        return config

    except FileNotFoundError:
        print(f"Error: Configuration file '{yaml_file}' not found.")
        sys.exit(1)

    except yaml.YAMLError as e:
        print(f"Error parsing YAML file: {e}")
        sys.exit(1)

def load_recipients_from_xml(xml_file):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        recipients = []
        for recipient in root.findall('recipient'):
            email_elem = recipient.find('email')
            attachment_elem = recipient.find('attachment')

            if email_elem is not None and email_elem.text:
                recipient_data = {
                    'email': email_elem.text.strip(),
                    'attachment': attachment_elem.text.strip() if attachment_elem is not None and attachment_elem.text else None
                }
                recipients.append(recipient_data)

        return recipients

    except FileNotFoundError:
        print(f"Error: XML file '{xml_file}' not found.")
        sys.exit(1)

    except ET.ParseError as e:
        print(f"Error parsing XML file: {e}")
        sys.exit(1)

def send_bulk_emails(yaml_config_file, xml_recipients_file):
    config = load_email_config(yaml_config_file)

    from_address = config.get('from_address')
    subject = config.get('subject', '')
    body = config.get('body', '')
    cc = config.get('cc')
    bcc = config.get('bcc')

    if not from_address:
        print("Error: 'from' field is required in the YAML configuration.")
        sys.exit(1)

    recipients = load_recipients_from_xml(xml_recipients_file)

    if not recipients:
        print("Error: No recipients found in XML file.")
        sys.exit(1)

    print(f"Preparing to send {len(recipients)} email(s)...")

    success_count = 0
    fail_count = 0

    for idx, recipient in enumerate(recipients, 1):
        to_address = recipient['email'] + '@myubt.de'
        attachment_path = 'exams/' + recipient['attachment']

        print(f"\n[{idx}/{len(recipients)}] Sending to: {to_address}")
        attachments = [attachment_path] if attachment_path else []

        success = send_mail(
            to_address=to_address,
            from_address=from_address,
            subject=subject,
            body=body,
            # cc=cc,
            # bcc=bcc,
            attachments=attachments
        )

        if success:
            print(f"✓ Email sent successfully to {to_address} with attachment {attachments}")
            success_count += 1
        else:
            print(f"✗ Failed to send email to {to_address}")
            fail_count += 1

    print(f"\n{'=' * 50}")
    print(f"Summary:")
    print(f"  Total: {len(recipients)}")
    print(f"  Successful: {success_count}")
    print(f"  Failed: {fail_count}")
    print(f"{'=' * 50}")

def main():
    yaml_config_file = "content.yml"
    xml_recipients_file = "mail.xml"

    send_bulk_emails(yaml_config_file, xml_recipients_file)

if __name__ == '__main__':
    main()