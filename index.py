import win32com.client as initW

print('Press ENTER to proceed...')
num_line = int(input("Insert the number line: "))
num_column = int(input("Insert the number column: "))



word = initW.Dispatch("Word.Application")
wordOpenCheck = initW.gencache.EnsureDispatch("Word.Application")
word.Visible = True
wordOpenCheck.Visible = True 

document = word.Documents.Add()

table = document.Tables.Add(document.Range(0, 0), num_line, num_column)

for i in range(num_line):
    for j in range(num_column):
        table.Cell(i + 1, j + 1).Range.Text =''


table.Borders.Enable = True 
