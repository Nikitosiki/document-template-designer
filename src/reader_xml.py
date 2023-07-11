from entity.usestring import UseString
import xml.etree.ElementTree as ET


# Парсим xml в dictionary
def xml_to_dict(file_path):
    dictionary = {}

    tree = ET.parse(file_path)
    root = tree.getroot()

    for item in root.findall('item'):
        key = item.get('key')
        value = item.get('text')
        dictionary[key] = UseString(value)


    return dictionary