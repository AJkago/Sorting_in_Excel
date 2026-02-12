import pandas as pd

df = pd.read_excel(r"C:\Users\chika\210517出荷データ(令和2年度まで).xlsx",sheet_name='Sheet3', index_col=0)


        else:
            #リストをループ
            for list in datecolumns:
                #セル値がすでにリストに含まれていたら何もしない
                if list == cell.value:
                    j = j + 1
                    break
            if j == 0:
                #セル値がリストに含まれていなければ配列へ追加
                if cell.value is not None:
                    datecolumns.append(cell.value)
                    num = num + 1
            k = 1
            for k in range(1,lastrow+1):
                filled_cell = ws3.cell(row = k, column = i).value
                if filled_cell is not None:
                    ws4.cell(row = k, column = num).value = filled_cell
                    k = k + 1
            i = i + 1
