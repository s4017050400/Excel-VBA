'主程式
Public Function idf(ByVal row As Integer, ByVal col As Integer) 'ByVal避免變數跑掉
    Dim myarray() As String
    Dim x As String
    Dim cnt, i As Integer
    i = 0
    cnt = count(row, col) 'counting the the count of row
    ReDim myarray(cnt) 'defined the size of array, by total number of row
    myarray(0) = Cells(row, col).Value '存在第一個陣列中
    Do While Cells(row, col).Value <> ""
        If myarray(i) = Cells(row, col).Value Then 'if 第一個參數相等 不做任何事
        Else
            i = i + 1 'if 兩數不相等，把數存在第二個陣列中
            myarray(i) = Cells(row, col).Value
        End If
        row = row + 1
    Loop
    'MsgBox myarray(0) & myarray(1) & myarray(2) & myarray(3)
    'reslut 各種結果存進對應的矩陣中
End Function
_____________________________________________________________________________________

"副程式 負責計數來決定陣列多大
Public Function count(ByVal row As Integer, ByVal col As Integer) 'ByVal避免變數跑掉
    'MsgBox Cells(row, col).Value
    Do While Cells(row, col).Value <> ""
        row = row + 1
    Loop
    count = row - 1
End Function
