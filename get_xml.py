import zipfile
from lxml import etree
import pandas as pd

FILE_PATH = r'F:\WORK\DIGITAL\fos2fos\Docs\fos/'
FILE_NAME = 'ФОС_Феномен факта в современных масс-медиа_ДО.docx'


class Application:
    def __init__(self, file_path, file_name):
        xml_content = self.get_word_xml(f'{file_path}/{file_name}')
        self.xml_tree = self.get_xml_tree(xml_content)

    def get_word_xml(self, docx_filename):
        with open(docx_filename, 'rb') as f:
            zip = zipfile.ZipFile(f)
            xml_content = zip.read('word/document.xml')
        return xml_content

    def get_xml_tree(self, xml_string):
        return (etree.fromstring(xml_string))


a = Application(FILE_PATH, FILE_NAME)

print(a.xml_tree)
