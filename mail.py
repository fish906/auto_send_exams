import subprocess
import os
import sys
import yaml

def send_mail(from_address, to_address, subject, body, cc=None, bcc=None, attachment=None):
    """
    Args:
        from_address (str): Address of sender
        to_adress (str or list): Recipient email address(es)
        subject (str): Email subject
        body (str): Email body text
        cc (str or list, optional): CC recipient(s)
        bcc (str or list, optional): BCC recipient(s)
        attachment (str): Attached exam
    """

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
        
    if attachment:
        abs_path = os.path.abspath(attachment)

        if not os.path.exists(abs_path):
            print(f'Warning: Attachment not found: {abs_path}')

        else:
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
        print(f"Email sent successfully to {to_address}!")
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
    
if __name__ == '__main__':
    config_file = "content.yml"
    config = load_email_config(config_file)

    send_mail(
        from_address=config.get('from_address'),
        to_address=config.get('to_address'),
        subject=config.get('subject', ''),
        body=config.get('body', ''),
        attachment=config.get('attachment')
    )

    """
    if not to_address or not from_address:
        print("Error: 'to' and 'from' fields are required in the YAML configuration.")
        sys.exit(1)
    """