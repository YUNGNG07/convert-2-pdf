"""
Command line program to convert Office documents and images to PDF formats
"""

from pdf_merger import word_to_pdf, excel_to_pdf, ppt_to_pdf, img_to_pdf
import sys
import os

word_formats = ['.doc', '.docx']
excel_formats = ['.xls', '.xlsx', '.csv']
ppt_formats = ['.ppt', '.pptx']
img_formats = ['.jpg', '.png', '.gif', '.webp']
format_dictionary = {
    'WORD': word_formats,
    'EXCEL': excel_formats,
    'PPT': ppt_formats,
    'IMG': img_formats,
}

def convert():
    """
    Convert Office documents and images to PDF format according to the file extension specified in the command-line argument
    """
    cmd = sys.argv[1:]
    formats = []
    # Gets program name cli.py if no formats are specified
    if len(cmd) == 0:
        formats = word_formats + excel_formats + ppt_formats
    # len = | cli.py | -f | <format> |
    # cmd = | -f | <format> |
    elif len(cmd) == 2 and cmd[0] == '-f':
        if cmd[1] == 'word':
            formats = word_formats
        elif cmd[1] == 'ppt':
            formats = ppt_formats
        elif cmd[1] == 'excel':
            formats = excel_formats
        elif cmd[1] == 'img':
            formats = img_formats
        elif cmd[1] == '*':
            formats = ppt_formats + word_formats + excel_formats + img_formats
        else:
            print("Invalid format.\nUse: python -f <word/ppt/excel/img/*>")
    else:
        print("Invalid format.\nUse: python -f <word/ppt/excel/img/*> or python merge")

    out_path = os.path.abspath('PDF')
    files = os.listdir()
    files.sort()

    for i in files:
        # Find the last occurence of '.' to determine file extension type
        pos = i.rfind('.')
        # rfind() returns -1 if '.' is not found
        if pos != -1:
            file, extension = out_path + r'\\' + i[:pos], i[pos:]
            if extension in formats:
                if i.startswith('~$') and i[2:] in files:
                    continue
                if extension in format_dictionary['WORD']:
                    word_to_pdf(os.path.abspath(i), file)
                elif extension in format_dictionary['EXCEL']:
                    excel_to_pdf(os.path.abspath(i), file)
                elif extension in format_dictionary['PPT']:
                    ppt_to_pdf(os.path.abspath(i), file)
                elif extension in format_dictionary['IMG']:
                    img_to_pdf(os.path.abspath(i), file)
                print(i, ": CONVERTED")

if __name__ == '__main__':
    if 'PDF' not in os.listdir():
        os.mkdir('PDF')
    convert()
