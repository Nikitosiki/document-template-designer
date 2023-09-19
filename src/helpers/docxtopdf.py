import sys
import subprocess
import re
import os
import sys


class LibreOfficeError(Exception):
    def __init__(self, output):
        self.output = output

# def convert_to(file_docx, out_dir, timeout=None):
#     file_docx = file_docx.replace("\\\\", "\\")
#     out_dir   = out_dir.replace("\\\\", "\\")

#     # libreoffice_path = read_path_libreoffice()+'soffice.exe'
#     # args = ['--headless', '--convert-to pdf', file_docx, '--outdir', out_dir]

#     argsStr = "start /min /D \"" + read_path_libreoffice() + '\" soffice --headless --convert-to pdf ' + file_docx + ' --outdir ' + out_dir
#     subprocess.run(argsStr, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)

#     if os.path.exists(os.path.splitext(file_docx)[0] + '.pdf'):
#         raise LibreOfficeError('')


def convert_to(file_docx, out_dir, timeout=None):
    file_docx = file_docx.replace("\\\\", "\\")
    out_dir   = out_dir.replace("\\\\", "\\")

    libreoffice_path = read_path_libreoffice()
    # libreoffice_path = "libreOffice"
    args = ['--headless', '--convert-to', 'pdf', file_docx, '--outdir', out_dir]
    process = subprocess.run([libreoffice_path] + args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)
    filename = re.search('-> (.*?) using filter', process.stdout.decode('cp1251'))

    if filename is None:
        raise LibreOfficeError(process.stdout.decode('cp1251'))
    else:
        return filename.group(1)


def read_path_libreoffice():
    try:
        # Получаем путь к текущей директории
        current_dir = os.path.dirname(sys.executable)

        # Формируем путь к файлу, расположенному рядом с программой
        file_path = os.path.join(current_dir, 'libreofficepath.txt')

        # Проверяем существование файла
        if not os.path.exists(file_path):
            # Если файл не существует, создаем его и записываем строку
            with open(file_path, 'w') as file:
                file.write('C:\\Program Files\\LibreOffice\\program\\')

        # Открываем файл на чтение
        with open(file_path, 'r') as file:
            line = file.readline() + 'soffice.exe'

        if not os.path.exists(line):
            raise LibreOfficeError("Error reading path to LibreOffice: {}".format(str(e)))

    except Exception as e:
        raise LibreOfficeError("Error reading path to LibreOffice: {}".format(str(e)))
    
    return line