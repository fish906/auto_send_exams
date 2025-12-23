import os
import re
import xml.etree.ElementTree as ET
from xml.dom import minidom

def extract_userid_from_filename(filename):
    pattern = r'bt\d{6}'
    match = re.search(pattern, filename)
    return match.group(0) if match else None

def create_xml (dictionary, output_file='mail.xml'):
    root = ET.Element('recipients')
    error_userid_list = []

    if not dictionary:
        print('Keine PDFs in Ordner gefunden')

    if dictionary:
        print(f'{len(dictionary.values())} PDS gefunden')

    for path in dictionary.keys():
        userid = extract_userid_from_filename(dictionary[path])

        if userid is not None:
            recipient = ET.SubElement(root, 'recipient')
            email = ET.SubElement(recipient, 'email')
            email.text = userid
            attachment = ET.SubElement(recipient, 'attachment')
            attachment.text = dictionary[path]
            path_to_attachment = ET.SubElement(recipient, 'path_to_attachment')
            path_to_attachment.text = path

        if userid is None:
            error_userid_list.append(dictionary[path])

    xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="  ")

    with open(output_file, 'w', encoding='UTF-8') as f:
        f.write(xml_str)

    total_recipients = len(root.findall('recipient'))
    print(f"Empfänger gesamt: {total_recipients}")

    if total_recipients != len(dictionary):
        print("\nWARNING: Not all files were properly processed")
        print('    The following file(s) were not processed:')
        for error in error_userid_list:
            print(f'        - {error}')

    print(f"\nXML file created: {output_file}")

def get_file_path(submissions_directory):
    sub_directory = [f for f in os.listdir(submissions_directory) if f.endswith('_file')]
    file_paths = {}

    for directory in sub_directory:
        full_path = os.path.join(submissions_directory, directory)
        test = os.listdir(full_path)
        for item in test:
            file_paths[full_path] = item

    return file_paths

def main():
    submissions_directory = "attachments"
    dictionary = get_file_path(submissions_directory)
    create_xml(dictionary)

if __name__ == "__main__":
    main()