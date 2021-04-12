import PyPDF2
import win32com.client
book = open("Test.pdf.pdf","rb")
pdf = PyPDF2.PdfFileReader(book)
pages = pdf.getPage(0)
print(pages.extractText())
a=win32com.client.Dispatch("SAPI.SpVoice")
a.Speak(pages.extractText())
