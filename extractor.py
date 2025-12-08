import os
import re
import xml.etree.ElementTree as ET
from xml.dom import minidom

def extract_userid_from_filename(filename):
    pattern = r'bt\d{6}'
    match = re.search(pattern, filename)
    return match.group(0) if match else None

def create_xml (pdf_files, output_file='mail.xml'):
    root = ET.Element('recipients')
    for pdf_file in pdf_files:
        filename_without_ext = os.path.splitext(pdf_file)[0]
        userid = extract_userid_from_filename(pdf_file)

        if userid:
            recipient = ET.SubElement(root, 'recipient')
            email = ET.SubElement(recipient, 'email')
            email.text = userid
            attachment = ET.SubElement(recipient, 'attachment')
            attachment.text = filename_without_ext

    xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="  ")

    with open(output_file, 'w', encoding='UTF-8') as f:
        f.write(xml_str)

    print(f"XML file created: {output_file}")
    print(f"Empfänger gesamt: {len(root.findall('recipient'))}")

if __name__ == "__main__":
    pdf_directory = "attachments"
    pdf_files = [f for f in os.listdir(pdf_directory) if f.endswith('.pdf')]

    if not pdf_files:
        print('Keine PDFs in Ordner gefunden')

    elif pdf_files:
        print(f'{len(pdf_files)} PDFs gefunden')

    create_xml(pdf_files)

    # ggf. wieder entfernen
    print('Datentest:')
    for file in pdf_files[:5]:
        userid = extract_userid_from_filename(file)
        print(f'    {file} -> UserID: {userid}')