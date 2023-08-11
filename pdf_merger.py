"""
Python module for converting Office documents and images into PDF format
"""

from comtypes.client import CreateObject
from img2pdf import convert

def word_to_pdf(input_filename, output_filename, formatType=17):
    """
    Convert Word document(s) to PDF format.

    Parameters:
        input_filename : Absolute file path to the specified documents
        output_filename: Name of converted PDF file
        formatType     : PDF wdExportFormat value

    Notes:
        WdExportFormat value: https://learn.microsoft.com/en-us/office/vba/api/word.wdexportformat
        Documents.Open.SaveAs method: https://learn.microsoft.com/en-us/office/vba/api/word.saveas2
    """
    word = CreateObject('Word.Application')
    word.Visible = False
    output_filename = output_filename + '.pdf'
    file = word.Documents.Open(input_filename)
    print('Opened')
    file.SaveAs(output_filename, formatType)
    print('Saved')
    file.Close()
    print('Closed')
    word.Quit()
    print('Quit')

def excel_to_pdf(input_filename, output_filename):
    """
    Convert Excel document(s) to PDF format.

    Parameters:
        input_filename : Absolute file path to the specified documents
        output_filename: Name of converted PDF file

    Notes:
        XlFixedFormatType: xlTypePDF = 0
        XlFixedFormatType value: https://learn.microsoft.com/en-us/office/vba/api/excel.xlfixedformattype

        XlFileFormat value: https://learn.microsoft.com/en-us/office/vba/api/excel.xlfileformat

        XlFixedFormatQuality value: https://learn.microsoft.com/en-us/office/vba/api/excel.xlfixedformatquality
        ------------------------------------------------------
        |        Name         |  Value  |    Description     |
        ------------------------------------------------------
        |  xlQualityMinimum   |    1    |  Minimum quality   |
        ------------------------------------------------------
        |  xlQualityStandard  |    0    |  Standard quality  |
        ------------------------------------------------------

        ExportAsFixedFormat method: https://learn.microsoft.com/en-us/office/vba/api/excel.chart.exportasfixedformat
    """
    excel = CreateObject('Excel.Application')
    excel.Visible = False
    output_filename = output_filename + '.pdf'
    file = excel.Workbooks.Open(input_filename)
    print('Opened')
    # fileType, FileName, Quality , IncludeDocProperties (do not include)
    file.ExportAsFixedFormat(0, output_filename, 1, 0)
    print('Saved')
    file.Close()
    print('Closed')
    excel.Quit()
    print('Quit')

def ppt_to_pdf(input_filename, output_filename, formatType=32):
    """
    Convert PowerPoint document(s) to PDF format.

    Parameters:
        input_filename : Absolute file path to the specified documents
        output_filename: Name of converted PDF file
        formatType     : PDF PpSaveAsFileType value

    Notes:
        PpSaveAsFileType value: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype

        Alternate way using Constants(): https://stackoverflow.com/questions/52258446/using-file-format-constants-when-saving-powerpoint-presentation-with-comtypes
    """
    ppt = CreateObject('Powerpoint.Application')
    ppt.Visible = False
    output_filename = output_filename + '.pdf'
    file = ppt.Documents.Open(input_filename)
    print('Opened')
    file.SaveAs(output_filename, formatType)
    print('Saved')
    file.Close()
    print('Closed')
    ppt.Quit()
    print('Quit')

def img_to_pdf(input_filename, output_filename):
    """
    Convert images to PDF format.

    Parameters:
        input_filename : Absolute file path to the specified documents
        output_filename: Name of converted PDF file
    """
    with open(f'{output_filename}.pdf', 'wb') as f:
        f.write(convert(input_filename))
