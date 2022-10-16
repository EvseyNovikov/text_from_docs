import requests
import os
import re
import shutil
import patoolib
import re
import xlrd


def create_folder(folder_name: str, directory_path: str) -> None:
    """Создает папку для скаченного файла."""

    dirs = os.listdir(directory_path)

    if 'files' not in dirs:
        os.mkdir('files')

    dirs = os.listdir(f'{directory_path}/files/')
    if folder_name not in dirs:
        os.mkdir(f'files/{folder_name}')
    return


def delete_folder(file_name: str, folder_name: str, directory_path: str) -> None:
    """Удаляет папку со скаченым файлов по завершению."""
    dirs = os.listdir(f'{directory_path}/files')

    if folder_name in dirs:
        shutil.rmtree(f'{directory_path}/files/{folder_name}')
    return


def download_file(link: str, file_name: str, folder_name: str, directory_path) -> bool:
    """Скачивает файл в созданную папку и возвращает имя файла"""
    try:
        download_file = requests.get(link)
    except UnicodeDecodeError:
        return False

    with open(f'{directory_path}/files/{folder_name}/{file_name}', 'wb') as file:
        file.write(download_file.content)

    return True


def unpacking_file(file_name: str, folder_name: str, file_extension:str, directory_path: str) -> bool:
    """
        Распаковывает файлы в нужную папку
        Использует пакет unzip
        sudo apt-get install unzip
        sudo apt-get install unrar
    """
    if file_extension == 'zip':
        # Распаковываем zip архив
        os.system(f'unzip {directory_path}/files/{folder_name}/{file_name} -d {directory_path}/files/{folder_name}/')
        os.remove(f'{directory_path}/files/{folder_name}/{file_name}')
    elif file_extension == 'rar':
        # Распаковываем rar архив
        try:
            patoolib.extract_archive(f'{directory_path}/files/{folder_name}/{file_name}', outdir=f'{directory_path}/files/{folder_name}')
        except patoolib.util.PatoolError:
            return False
        os.remove(f'{directory_path}/files/{folder_name}/{file_name}')
    return True


def all_paths_files_in_folder(folder_name: str, directory_path: str) -> list:
    """Возвращает все пути к файлам в директории"""
    files = [i for i in os.walk(f'{directory_path}/files/{folder_name}')]

    all_paths = []

    for f in files:
        for i in f[2]:
            if '~$' in i:
                continue
            if i.split('.')[-1] in ['doc', 'docx', 'html', 'xls', 'xlsx', 'rtf']:
                all_paths.append(f'{f[0]}/{i}')
    return all_paths


def convert_to_txt_file(path: str) -> str:
    """
        Конвертирует файл в txt файл
        Возвращает путь к файлу
        sudo apt-get install libreoffice
    """
    # Директория в которой лежит файл для конвертации
    directory_path = '/'.join(path.split('/')[0:-1])
    print(directory_path, 'directory_path')
    # Конвертируем в txt файл rtf файл
    cmd = f'cd "{directory_path}"; lowriter --headless --convert-to txt "' + path + '"'
    os.system(cmd)
    # Получаем имя файла сконвертированное
    file_name = '.'.join(path.split('/')[-1].split('.')[0:-1]) + '.txt'

    file_path = f'{directory_path}/{file_name}'

    return file_path


def get_text_from_file(path: str) -> str:
    """Возвращает текст из файла."""
    with open(path, 'r') as file:
        text = file.read()
    return text


def doc2text(path: str) -> str:
    """
        Возвращает текст из файлов с расширение doc, docx
    """
    txt_file_path = convert_to_txt_file(path)

    text = get_text_from_file(txt_file_path)
    return text


def html2text(path: str) -> str:
    """
        Возвращает текст из html документа.
    """
    try:
        with open(path, 'r') as file:
            text = file.read()

        text = text.split('</head>')[-1]
        comp = re.compile(r'<.*?>')
        # Удаляет из текста все тэги
        text = comp.sub('', text)
        text = text.replace('\n', ' ').replace('\t', '')
    except UnicodeDecodeError:
        text = ''
    return text

def rtf2text(path: str) -> str:
    """
        Возвращает текст из rtf документа
    """

    txt_file_path = convert_to_txt_file(path)
    text = get_text_from_file(txt_file_path)
    return text


def xlsx2text(path: str) -> str:
    """Возвращает текст из xlsx документа"""
    # Открываем эксель файл
    try:
        wb = xlrd.open_workbook(path)
    except xlrd.biffh.XLRDError:
        return ''
    # Получаем все страницы файла
    sheets = wb.sheets()

    text = ''

    for sheet in sheets:
        #получаем список значений из всех записей
        vals = [' '.join(map(str, sheet.row_values(rownum))) for rownum in range(sheet.nrows)] 
        text += '\n'.join(vals)

    return text


def get_text_in_file(path: str) -> str:
    """Возвращает текст из файла"""
    extension_file = path.split('.')[-1]
    # Если расширение doc то конвертируем в docx
    if extension_file in ['doc', 'docx']:
        return doc2text(path)      
    elif extension_file == 'html':
        return html2text(path)
    elif extension_file == 'rtf':
        return rtf2text(path)
    elif extension_file in ['xls', 'xlsx']:
        return xlsx2text(path)
    else:
        return ''


def get_all_texts(files_paths: list) -> list:
    """
        Отдает текста из всех файлов
    """
    all_texts = []

    for path in files_paths:
        all_texts.append(get_text_in_file(path))
    return all_texts


def get_all_inn(texts: list) -> list:
    """
        Возвращает список ИНН во всех найденных текстах
    """
    inn = []
    for text in texts:
        INN = re.findall('ИНН\x20*(\d{10})', text)
        if not INN:
            continue
        else:
            for i in INN:
                inn.append(i)
    return inn


def get_text_and_inn(link: str) -> tuple:
    """
    Возвращает кортеж с массивом найденных ИНН 
        и текстами файлов
    """
    directory_path = os.getcwd()

    all_texts = []
    all_inn = []

    file_name = link.split('/')[-1]
    folder_name = link.split('/')[-1].split('.')[0]
    file_extension = file_name.split('.')[-1]

    # Создаем папку для скачеваемого файла
    create_folder(folder_name, directory_path)
    
    # Скачиваем файл
    download = download_file(link, file_name, folder_name, directory_path)
    # Если файл скачалася, то идем дальше
    if download:
        # Распаковываем файл если это архив
        unpack = unpacking_file(file_name, folder_name, file_extension, directory_path)

        if unpack:
            # Ищем все файлы в папке
            all_paths = all_paths_files_in_folder(folder_name, directory_path)

            # Собираем все текста из файлов
            all_texts = get_all_texts(all_paths)

            # Собираем все ИНН в текстах
            all_inn = get_all_inn(all_texts)

    # Удаляем скаченные файлы
    delete_folder(file_name, folder_name, directory_path)
    return all_texts, all_inn
