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
    error_userid_list = []

    if not pdf_files:
        print('Keine PDFs in Ordner gefunden')

    if pdf_files:
        print(f'{len(pdf_files)} PDFs gefunden')

    for pdf_file in pdf_files:
        userid = extract_userid_from_filename(pdf_file)

        if userid is not None:
            recipient = ET.SubElement(root, 'recipient')
            email = ET.SubElement(recipient, 'email')
            email.text = userid
            attachment = ET.SubElement(recipient, 'attachment')
            attachment.text = pdf_file

        if userid is None:
            error_userid_list.append(pdf_file)

    xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="  ")

    with open(output_file, 'w', encoding='UTF-8') as f:
        f.write(xml_str)

    total_recipients = len(root.findall('recipient'))
    print(f"Empfänger gesamt: {total_recipients}")

    if total_recipients != len(pdf_files):
        print("\nWARNING: Not all files were properly processed")
        print('    The following file(s) were not processed:')
        for error in error_userid_list:
            print(f'        - {error}')

    print(f"\nXML file created: {output_file}")

def main():
    pdf_directory = "exams"
    pdf_files = [f for f in os.listdir(pdf_directory) if f.endswith('.pdf')]

    create_xml(pdf_files)

if __name__ == "__main__":
    main()