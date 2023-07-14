from entity.usestring import UseString


# Парсим строку в dictionary
def str_to_dict(dictionary_str):
    dictionary = {}
    items = dictionary_str.strip('{}').split(', ')
    for item in items:
        key, value = item.split(': ')
        key = key.strip("'")
        value = value.strip("'")
        dictionary[key] = UseString(value, False)
    return dictionary