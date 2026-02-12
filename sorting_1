import win32com.client

file_name = r'C:\Users\chika\2021切花.xlsx'

def excelSort(file_name):

    #Excelファイル操作のための準備
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(Filename=file_name)

    # ワークシートの指定
    ws = wb.Worksheets(1)
    ws.Activate()

    #Sort 数値の割り振り
    xlAscending = 1
    xlYes = 1
    xlUp = -4162
    lastrow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    lastcolumn = ws.Cells(ws.Columns.Count, 2).End(xlUp).Row

    #A列で昇順
    ws.Range(ws.Range("A2"), ws.Cells(lastrow, lastcolumn)).Sort(Key1=ws.Range("A2"), Order1=xlAscending, Header=xlYes)

    #Excelファイルを保存
    wb.Save()

    #Excelを閉じる
    excel.Quit()

    return

excelSort(file_name)
