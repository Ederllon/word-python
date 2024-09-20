import win32com.client

# Cria uma instância do Word
word = win32com.client.Dispatch("Word.Application")
word.Visible = True  # Mostra o Word

# Adiciona um novo documento
document = word.Documents.Add()

# Adiciona texto
addOnDocument = document.Content
addOnDocument.Text = "texto"

# função de tabela no word
def tableFunction():
    table = document.Tables.Add(document.Range(0, 0) , 3, 2)
    table.Cell(1, 1).Range.Text = "Célula 1"
    table.Cell(1, 2).Range.Text = "Célula 2"
    table.Cell(2, 1).Range.Text = "Célula 3"
    table.Cell(2, 2).Range.Text = "Célula 4"

tableFunction()

# Salva o documento
document.SaveAs("./document.docx")

# Fecha o Word
# word.Quit()
