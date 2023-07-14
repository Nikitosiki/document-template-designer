from entity.usestring import UseString
import xml.etree.ElementTree as ET


# Парсим xml в dictionary
def xml_to_dict(file_path):
    dictionary = {}

    tree = ET.parse(file_path)
    root = tree.getroot()

    for item in root.findall('item'):
        key = item.get('key')
        text = item.get('text')
        image_path = item.get('image_path')

        if (key == None):
            raise Exception(f"Error. In file: {file_path}. \nKey not found!")

        if (text != None and image_path != None):
            raise Exception(f"Error. In file: {file_path}. \nIn key: {key}. \nCannot contain both text and image!")

        if (text == None and image_path == None):
            raise Exception(f"Error. In file: {file_path}. \nIn key: {key}. \nDoes not contain text or image!")

        if (text != None):
            dictionary[key] = UseString(text, False)
        elif (image_path != None):
            image_width = item.get('image_width_Inch')
            image_height = item.get('image_height_Inch')
            dictionary[key] = UseString(image_path,
                                        True,
                                        float(image_width) if image_width else 0,
                                        float(image_height) if image_height else 0)
        else:
            raise Exception(f"Error. In file: {file_path}. \nIn key: {key}.")


    return dictionary