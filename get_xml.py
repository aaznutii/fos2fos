import zipfile
from lxml import etree
from Functions import get_file_path, get_file_name


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


a = Application(get_file_path(), get_file_name())

print(a.xml_tree)
