import os
from comtypes.client import CreateObject
from PyPDF2 import PdfWriter

input_directory = r'C:\\Users/yungng07/Downloads/'
output_directory = r'C:\\Users/yungng07/Downloads/'
i = 1

def word_to_pdf(input_filename, output_filename, input_directory, output_directory, multiple=False, formatType=17):
    """
    Convert Word document(s) to PDF format
    """
    word = CreateObject('Word.Application')
    word.Visible = False
    output_filename = output_filename + '.pdf'
    file = word.Documents.Open(input_directory + input_filename)
    print('Opened')
    file.SaveAs(output_directory + output_filename, formatType)
    print('Saved')
    file.Close()
    print('Closed')
    word.Quit()
    print('Quit')

    if multiple:
        for word_file in os.listdir(input_directory):
            # print('Converting
            if word_file.endswith('.doc') or word_file.endswith('.docx'):
                print(word_file)
                word_to_pdf(word_file, 'pdf' + str(i) + '.pdf', input_directory, output_directory)
                i += 1

def excel_to_pdf(input_filename, output_filename, input_directory, output_directory, multiple=False, formatType=17):
    """
    Convert Excel document(s) to PDF format
    """
    word = CreateObject('Excel.Application')
    word.Visible = False
    output_filename = output_filename + '.pdf'
    file = word.Documents.Open(input_directory + input_filename)
    print('Opened')
    file.SaveAs(output_directory + output_filename, formatType)
    print('Saved')
    file.Close()
    print('Closed')
    word.Quit()
    print('Quit')

    if multiple:
        word_file_list = []
        for word_file in os.listdir(input_directory):
            # TODO: Get a string of all Word files
            word_file_list.append(word_file)
            # Get Word document filenames
            filename = word_file.split('.')[0]
            print('Converting ' + filename)
            if word_file.endswith('.doc') or word_file.endswith('.docx'):
                print(word_file)
                word_to_pdf(word_file, 'pdf' + str(i) + '.pdf', input_directory, output_directory)
                i += 1

def ppt_to_pdf(input_filename, output_filename, input_directory, output_directory, multiple=False, formatType=17):
    """
    Convert PowerPoint document(s) to PDF format
    """
    word = CreateObject('Powerpoint.Application')
    word.Visible = False
    output_filename = output_filename + '.pdf'
    file = word.Documents.Open(input_directory + input_filename)
    print('Opened')
    file.SaveAs(output_directory + output_filename, formatType)
    print('Saved')
    file.Close()
    print('Closed')
    word.Quit()
    print('Quit')

    if multiple:
        for word_file in os.listdir(input_directory):
            # print('Converting
            if word_file.endswith('.doc') or word_file.endswith('.docx'):
                print(word_file)
                word_to_pdf(word_file, 'pdf' + str(i) + '.pdf', input_directory, output_directory)
                i += 1

def merge_pdf(filename, file_directory, output_directory):
    """
    Merge PDFs
    """
    merger = PdfWriter()

    if file_directory:
        for pdf in os.listdir(file_directory):
            if pdf.endswith('.pdf'):
                print(pdf)
                # Reference: https://stackoverflow.com/questions/65162124/python3-filenotfounderror-errno-2-no-such-file-or-directory-first-filename
                # os.listdir() returns relative paths, need to reconstruct absolute path to open the files
                pdf_directory = os.path.join(file_directory, pdf)
                merger.append(pdf_directory)

    output_destination = os.path.join(output_directory, filename)
    merger.write(output_destination)
    merger.close()

