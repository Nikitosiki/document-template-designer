# Document Template designer

This project was created to generate documents in .pdf and .docx format. As a basis, you provide a .docx document, on the basis of which the output file is formed, the words {key} can also be replaced with text or an image, and it is also possible to remove empty lines in tables.

Этот проект создан для формирования документов в формате .pdf и .docx. За основу вы предоставляет документ .docx, по основе которого формируется виходной файл, также могут быть заменены слова {key} на текст или картинку, а также есть возможность удалить пустие строки в таблицах.

## Installation / Установка

### Requirements / Требования

- Python 3.8 or higher / Python 3.8 или выше

### Setup / Настройка

1. Clone the repository / Клонируйте репозиторий:

```bash
git clone <repository-url>
cd document-template-designer
```

2. Install dependencies / Установите зависимости:

```bash
pip install -e .
```

This command installs the project in development mode with all required dependencies.
Эта команда устанавливает проект в режиме разработки со всеми необходимыми зависимостями.

## Usage / Использование

Run the script with the following command / Запустите скрипт с помощью следующей команды:

```bash
cd ./src
python main.py -h
python main.py <input_file> <output_file> [options]
```

### Arguments / Аргументы

- `input_file` - Path to input .docx file / Путь к входному файлу .docx
- `output_file` - Path to output file / Путь к выходному файлу
- `--replacements_string` - Dictionary of replacements as a string / Словарь замен в виде строки
- `--replacements_file` - Path to XML file with replacements / Путь к XML файлу с заменами
- `--remove_replacements_file` - Remove the replacement file after use / Удалить файл замен после использования
- `--all_replacements_used_check` - Check if all replacements were used / Проверить, использованы ли все замены
- `--clear_table` - Clear blank lines in tables / Очистить пустые строки в таблицах
- `--save_docx` - Save as .docx file / Сохранить как .docx файл
- `--save_pdf` - Save as .pdf file / Сохранить как .pdf файл

### Example / Пример

```bash
cd ./src
python main.py template.docx output --replacements_file replacements.xml --save_docx --save_pdf
```
