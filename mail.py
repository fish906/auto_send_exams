import subprocess
import sys

def send_mail(from_address, to_address, subject, body, cc=None, bcc=None):
    """
    Args:
        from_address (str): Address of sender
        to_adress (str or list): Recipient email address(es)
        subject (str): Email subject
        body (str): Email body text
        cc (str or list, optional): CC recipient(s)
        bcc (str or list, optional): BCC recipient(s)
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
    
if __name__ == '__main__':
    send_mail(
        from_address="bt477259@myubt.de",
        to_address="fabian.netz@outlook.de",
        subject="Test",
        body="This is a test"
    )