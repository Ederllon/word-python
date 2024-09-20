import win32com.client

# Cria uma instância do Word
word = win32com.client.Dispatch("Word.Application")
word.Visible = True  # Mostra o Word

# Adiciona um novo documento
documento = word.Documents.Add()

# Adiciona texto
parágrafo = documento.Content
parágrafo.Text = "Olá, Word!"

# Salva o documento
documento.SaveAs("C:\\caminho\\para\\seu\\arquivo.docx")

# Fecha o Word
word.Quit()
