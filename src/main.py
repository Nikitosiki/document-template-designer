from handler import main
from parsers.reader_xml import xml_to_dict
from parsers.reader_str import str_to_dict
from helpers.docxtopdf import LibreOfficeError
from helpers.errormes import*
from helpers.fileos import*

import argparse


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Replacing keywords in a .docx file, clearing blank lines in tables, and saving to .docx and/or .pdf. \nIf --save_docx or --save_pdf is not selected then --save_pdf is used, you can select both save modes at the same time.")
    parser.add_argument('input_file', type=str, help="Input .docx file with path")
    parser.add_argument('output_file', type=str, help="Path to the output file named")

    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--replacements_string', type=str, help="Dictionary of replacements as a string")
    group.add_argument('--replacements_file', type=str, help="Path to the XML file with replacements")

    parser.add_argument('--remove_replacements_file', action='store_true', help="Removes the replacement file after use")
    parser.add_argument('--all_replacements_used_check', action='store_true', help="When used, displays an error if at least one word from the replacements is not used")
    parser.add_argument('--clear_table', action='store_true', help="Add this flag to clear the table")
    parser.add_argument('--save_docx', action='store_true', help="Add this flag to save to the .docx")
    parser.add_argument('--save_pdf', action='store_true', help="Add this flag to save to the .pdf")

    args = parser.parse_args()
    if args.replacements_string and args.replacements_file:
        show_error("Error. Only one of --replacements_string or --replacements_file should be specified.")

    try:
        if args.replacements_string:
            replacements = str_to_dict(args.replacements_string)
        elif args.replacements_file:
            replacements = xml_to_dict(args.replacements_file)
            delete_file(args.remove_replacements_file, args.replacements_file)

    except Exception as error_mess:
        show_error(f"Error. Reading replacements. \n\n{error_mess}")

    try:
        main(args.input_file, args.output_file, args.clear_table, args.all_replacements_used_check, args.save_docx, args.save_pdf, replacements)
        print("Done!")
    except LibreOfficeError:
        show_error("LibreOffice not found or returned an error! Requires installation to save the .pdf file.")


# ______________________________________
# ____________Usage example_____________
# ______________________________________

# Console:
# PDFhandler.exe "I:\\NEEEET (XAI)\\projects\\DesignerPDFApp\\test_template-1.docx" "I:\\NEEEET (XAI)\\projects\\DesignerPDFApp\\out_template-1.pdf" --clear_table "{'protocol_number': '234-123-1', 'outside_temp': '23°C', 'calibration_date': '01.07.2023', 'atmospheric_pressure': '78', 'sensor_name': 'датчик (Надлишковий тиск)', 'sensor_model': '2140 мод.ряд САФІР', 'sensor_number': '1', 'sensor_upper_limit': '0 / 250 kPa', 't_pressure_1': 'ывавы', 't_reverse_5': 'fsdfsdfsd fsdf sdf sfsdfdsfsdfd', 'max_inaccuracy': '-125,000 % ', 'performed_position': 'пров.інженер з електроніки', 'performed_name': 'ЖУРАВЕЛЬ Д.Ю.', 'certificate': 'Сертифікат відсутній'}"

# ______________________________________
# ______________________________________
# ______________________________________


# def main_test1():
#     # Засекаем начальное время
#     start_time = time.time()


#     input_file = 'I:\\NEEEET (XAI)\\projects\\DesignerPDFApp\\test_template-1'
#     output_file = 'I:\\NEEEET (XAI)\\projects\\DesignerPDFApp\\out_template-1'
#     replacements = {
#         'protocol_number':      UseString('234-123-1'),
#         'outside_temp':         UseString('23°C'),
#         'calibration_date':     UseString('01.07.2023'),
#         'atmospheric_pressure': UseString('78'),
#         'sensor_name':          UseString('датчик (Надлишковий тиск)'),
#         'sensor_model':         UseString('2140 мод.ряд САФІР'),
#         'sensor_number':        UseString('1'),
#         'sensor_upper_limit':   UseString('0 / 250 kPa'),
#         't_pressure_1':         UseString('ывавы'),
#         't_reverse_5':          UseString('fsdfsdfsd fsdf sdf sfsdfdsfsdfd'),

#         'max_inaccuracy':       UseString('-125,000 % '),
#         'performed_position':   UseString('пров.інженер з електроніки'),
#         'performed_name':       UseString('ЖУРАВЕЛЬ Д.Ю.'),
#         'certificate':          UseString('Сертифікат відсутній')
#         }

#     # main(input_file + ".docx", output_file + ".pdf", True, replacements)

#     print(f"Время заполнения: {time.time() - start_time} секунд")
#     start_time = time.time()


#     doc = Document(input_file + '.docx')
#     modify_docx(doc, replacements)
#     remove_empty_table_rows( doc )
#     doc.save(output_file + '.docx')


#     print(f"Время выполнения функции 1: {time.time() - start_time} секунд")
#     start_time = time.time()

#     convert(output_file + '.docx', output_file + '.pdf')

#     # convert_to_pdf(output_file + '.docx', output_file + '.pdf')


#     print(f"Время выполнения функции 2: {time.time() - start_time} секунд")
#     print(f"Файл {output_file} успешно сохранен.")


# main_test1()

# main(
#     "I:\\NEEEET (XAI)\\projects\\DesignerPDFApp\\test_template-1.docx",
#     "I:\\NEEEET (XAI)\\projects\\DesignerPDFApp\\out_template-1.pdf",
#     True,
#     False,
#     True,
#     string_to_dict("{'protocol_number': '234-123-1', 'outside_temp': '23°C', 'calibration_date': '01.07.2023', 'atmospheric_pressure': '78', 'sensor_name': 'датчик (Надлишковий тиск)', 'sensor_model': '2140 мод.ряд САФІР', 'sensor_number': '1', 'sensor_upper_limit': '0 / 250 kPa', 't_pressure_1': 'ывавы', 't_reverse_5': 'fsdfsdfsd fsdf sdf sfsdfdsfsdfd', 'max_inaccuracy': '-125,000 % ', 'performed_position': 'пров.інженер з електроніки', 'performed_name': 'ЖУРАВЕЛЬ Д.Ю.', 'certificate': 'Сертифікат відсутній'}")
#     )