from comtypes.client import CreateObject

filename = 'myInnovation Form-myInnovation FY22 - Automate Consolidating Quality Issue for TBird 1.5_1.6 rev2.docx'
filename_pdf = filename + '.pdf'

word = CreateObject('Word.Application')
doc_com = word.Documents.Open(r'C:\\Users/yungng07/Downloads/' + filename)
# PDF format: 17
filename_pdf = filename + '.pdf'
doc_com.SaveAs(r'C:/Users/yungng07/Documents/pdf-generator/report/' + filename_pdf, FileFormat=17)
doc_com.Close()
word.Quit()

