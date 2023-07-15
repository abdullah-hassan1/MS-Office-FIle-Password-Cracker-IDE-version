from constants import *
from ziptools import update_archive, \
    extract_archive, \
    extract_archive_all
import re, os, shutil, time, sys

exit = sys.exit


def timeit(func):
    '''
    Decorator for measuring void functions time
    '''

    def wrapper(*args, **kwargs):
        start = time.time()
        func(*args, **kwargs)
        print(
            f'{os.path.basename(args[0])} Unlocked in {time.time() - start:.4f}s'
        )

    return wrapper


def decide(file):
    '''
    Detect the extension of input file

    :return: Settings related to this extension
    '''
    ext = file.split('.')[-1].lower()
    assert ext in EXTS, ext + ' files are not supported!'
    return EXTS[ext]


def omit_protection_code(file, regex):
    '''
    Remove protection strings from in office xml files
    '''
    locked_xml = re.sub(
        regex,
        '',
        open(file, encoding='utf-8').read()
    )
    open(file, 'w', encoding='utf-8').write(locked_xml)


def handle_xlsx_sheets(file, regex):
    '''
    Remove read-only protection for sheets only
    '''
    extract_archive_all(file, 'xl/worksheets/')
    for sheet in filter(lambda i: i.startswith('sheet'), os.listdir()):
        omit_protection_code(sheet, regex)
        update_archive(file, sheet, 'xl/worksheets/' + sheet)
        os.remove(sheet)


@timeit
def unlock(file, tmp='tmp'):
    '''
    Unlocking read-only protection from office files
    '''
    _type = decide(file)

    file = shutil.copy2(
        file,
        os.path.join(
            os.path.dirname(file),
            '[UNOFFICE] - ' + os.path.basename(file)
        )
    )

    handle_xlsx_sheets(file, _type[1]) if _type == XLSX else None
    extract_archive(file, _type[0], tmp)
    omit_protection_code(tmp, _type[1])
    update_archive(file, tmp, _type[0])
    os.remove(tmp)


def argparse():
    '''
    Argument parser for command-line options
    '''
    num_files = int(input("Enter the number of files to unlock: "))
    assert num_files > 0, 'Must enter at least 1 file'

    for _ in range(num_files):
        file_path = input("Enter the path to the file e.g E:\\123.docx or E:\\123.pptx or any with its extension like "
                          ".docx etc: ")
        unlock(file_path)

    print("Your File is Unlocked, Check the Folder Where the Original File is. Thanks")


if __name__ == '__main__':
    argparse()
