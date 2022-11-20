__author__="Disha Arya"
__credits__="Disha Arya"
__date__="19-Nov-2022"
__email__="ajaydisha2019@gmail.com"
import win32com.client
import os
word=win32com.client.Dispatch("word.Application")
word.visible=0
doc_pdf="LAB3(1).pdf"
input_file=os.path.abspath(doc_pdf)
wb=word.Documents.Open(input_file)
output_file=os.path.abspath(doc_pdf[0:4]+"docx".format())
wb.SaveAs2(output_file,FileFormat=16)
print("pdf to docx is Completed")
wb.Close()
word.Quit()