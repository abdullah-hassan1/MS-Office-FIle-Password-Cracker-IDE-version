from unoffice import unlock
import os

os.chdir(os.path.dirname(os.path.realpath(__file__)))

for ext in ['docx', 'pptx', 'xlsx']:
    file_path = input(f"Enter the path to the {ext} file: ")
    unlock(os.path.join('tests', 'test.' + ext))
