# word-python

# DOCX

python-docx: É uma das bibliotecas mais utilizadas para criar, modificar e ler arquivos do Microsoft Word (.docx). Você pode adicionar texto, tabelas, imagens, e muito mais.

# WIN32COM.CLIENT

win32com.client: Se você estiver usando o Windows e tiver o Microsoft Word instalado, pode usar esta biblioteca para automação, permitindo que você controle o Word através do Python
O win32com.client é um módulo do pacote pywin32 que permite a automação de aplicativos do Windows via COM (Component Object Model). Aqui estão os passos básicos para usá-lo:

Instalação do pywin32: Primeiro, você precisa instalar o pacote pywin32. Você pode fazer isso usando o pip:

programa: 
fazer um import do client.

Criar uma instância do Word:
word = win32com.client.Dispatch("Word.Application")
word.Visible = True  # Para mostrar o Word

Adicionar um novo documento:
documento = word.Documents.Add()

Você pode adicionar texto ao documento da seguinte forma:
parágrafo = documento.Content
parágrafo.Text = "Olá, Word!"

Para salvar e fechar o Word, use:
documento.SaveAs("C:\\caminho\\para\\seu\\arquivo.docx")
word.Quit()


