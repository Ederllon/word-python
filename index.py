import win32com.client as initW
# test --------------------------------------------------------------------
# print()
# word = initW.Dispatch("Word.Application")
# wordOpenCheck = initW.gencache.EnsureDispatch("Word.Application")
# word.Visible = True
# wordOpenCheck.Visible = True 
# localrepo = input(str('Insert repo: '))
# localvalue = r'{}' .format(localrepo)
# word.Documents.Open(localvalue)
# raise SystemExit("TEST FINISHED!")
# test  --------------------------------------------------------------------

print('Press ENTER to proceed...')

q1 = input(str('Arquive exist? [Y/N] ')).upper()



word = initW.Dispatch("Word.Application")
wordOpenCheck = initW.gencache.EnsureDispatch("Word.Application")





if 'N'in q1 : 
    document = word.Documents.Add()
    num_line = int(input("Insert the number line: "))
    num_column = int(input("Insert the number column: "))
    table = document.Tables.Add(document.Range(0, 0), num_line, num_column)    

else:
    localrepo = input(str('Insert repo: '))
    localvalue = r'{}' .format(localrepo)
    document = word.Documents.Open(localvalue)
    num_line = int(input("Insert the number line: "))
    num_column = int(input("Insert the number column: "))
    table = document.Tables.Add(document.Range(0, 0), num_line, num_column)    
    




for i in range(num_line):
    for j in range(num_column):
        table.Cell(i + 1, j + 1).Range.Text =''

word.Visible = True
wordOpenCheck.Visible = True 
table.Borders.Enable = True  


