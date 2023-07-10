import sys
import subprocess
import re
import os
import sys


def convert_to(file_docx, out_dir, timeout=None):
    file_docx = file_docx.replace("\\\\", "\\")
    out_dir   = out_dir.replace("\\\\", "\\")


    # args = ["start /D \"" + read_path_libreoffice() + "\" soffice", '--headless', '--convert-to', 'pdf', file_docx, '--outdir', out_dir]
    argsStr = "start /D \"" + read_path_libreoffice() + '\" soffice --headless --convert-to pdf ' + file_docx + ' --outdir ' + out_dir
    # print("args:  " + str(args))
    # process = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)
    print("argsStr:  " + argsStr)
    process = subprocess.run(argsStr, shell=True)
    # print(".stdout.decode():  " + str(process.stdout.decode()))
    # filename = re.search('-> (.*?) using filter', process.stdout.decode())
    # print("filename: " + str(filename))

    # if filename is None:
    #     raise LibreOfficeError(process.stdout.decode())
    # else:
    #     return filename.group(1)


# def libreoffice_exec():
#     # TODO: Provide support for more platforms
#     if sys.platform == 'darwin':
#         return '/Applications/LibreOffice.app/Contents/MacOS/soffice'
#     return 'libreoffice'


def read_path_libreoffice():
    # Получаем путь к текущей директории
    current_dir = os.path.dirname(sys.executable)

    # Формируем путь к файлу, расположенному рядом с программой
    file_path = os.path.join(current_dir, 'LibreOfficePath.txt')

    # Проверяем существование файла
    if not os.path.exists(file_path):
        # Если файл не существует, создаем его и записываем строку
        with open(file_path, 'w') as file:
            file.write('C:\\Program Files\\LibreOffice\\program\\')

    # Открываем файл на чтение
    with open(file_path, 'r') as file:
        # Читаем строку из файла
        line = file.readline()
    
    # Возвращаем прочитанную строку
    return line


# class LibreOfficeError(Exception):
#     def __init__(self, output):
#         self.output = output


# if __name__ == '__main__':
#     print('Converted to ' + convert_to(sys.argv[1], sys.argv[2]))