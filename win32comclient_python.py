
import win32com.client

# Inicializar o Word
word = win32com.client.Dispatch('Word.Application')
doc = word.Documents.Add()
doc.Content.Text = "Texto no Word"
doc.SaveAs('documento.docx')
word.Quit()
