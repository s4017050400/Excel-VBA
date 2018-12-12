Option Explicit

Public Sub main()   'ok
    Dim row As Range
    For Each row In Range("A:A")
        If row.Value = "                    +---------+---------+---------+---------+---------+---------+---------+---------+---------+---------+---------+---------+---------+---------+---------+---------+---------+---------+---------+" Then
            row.EntireRow.Delete
        End If
    Next
End Sub



Public Sub main2()  'ok
    Dim r As Range
    Dim ary(11), x, i, z As Integer '存ABcontition的矩陣
    x = 1
    i = 1
    z = 1
    For Each r In Range("A:A")  'loop in A column
        If r.Value = "=========== UAA077 ===========" Or r.Value = "ZQ CALIBRATION INITIAL" Then              'condition A next B
            ary(x) = r.row
            x = x + 1       '1-11 for5*2=10
        End If
    Next    'get value of array(7),odd=a,even=b
    For z = 1 To 5 'loop 5time (11-1)/2=5
        Range(Rows(ary(2 * z - 1)), Rows(ary(2 * z))).Select 'rowA to rowB to be selected
        Selection.Copy
        Sheets("Sheet1").Activate
        Cells(i, 1).Select
        ActiveSheet.Paste
        i = i + ary(2 * z) - ary(2 * z - 1) + 4 'i=i+四格
        Sheets("W").Activate
    Next z
        Sheets("Sheet1").Activate
End Sub



Public Sub main3()  'ok
    Dim r As Range
    For Each r In Range("A:A")
        If r.Value = "ZQ CALIBRATION INITIAL" Then
            r.EntireRow.Delete
        End If
    Next
End Sub




Public Sub main4()  'ok
    Dim r As Range
    Dim ary(15), x As Integer   '5*3=15
    Dim i, sht2c, sht2r As Integer
    x = 1
    sht2r = 1 'index of sheets2 row
    For Each r In Range("A:A")
        Select Case r   'parameter to be asign
            Case "-50 IDD RESULT (400Mbps)"
                ary(1) = r.row  '第一個位置
                Cells(r.row, 1).Select
                Selection.End(xlDown).Select
                Selection.End(xlToRight).Select
                ary(2) = Selection.row  '末的位置
                ary(3) = Selection.Column   'column, i+3
                'MsgBox 1
            Case "-37 IDD RESULT (533Mbps)"
                ary(4) = r.row  '第一個位置
                Cells(r.row, 1).Select
                Selection.End(xlDown).Select
                Selection.End(xlToRight).Select
                ary(5) = Selection.row  '末的位置
                ary(6) = Selection.Column
                'MsgBox 2
            Case "-30 IDD RESULT (667Mbps)"
                ary(7) = r.row  '第一個位置
                Cells(r.row, 1).Select
                Selection.End(xlDown).Select
                Selection.End(xlToRight).Select
                ary(8) = Selection.row  '末的位置
                ary(9) = Selection.Column
               'MsgBox 3
            Case "-25 IDD RESULT (800Mbps)"
                ary(10) = r.row  '第一個位置
                Cells(r.row, 1).Select
                Selection.End(xlDown).Select
                Selection.End(xlToRight).Select
                ary(11) = Selection.row  '末的位置
                ary(12) = Selection.Column
                'MsgBox 4
            Case "-18 IDD RESULT (1066Mbps)"
                ary(13) = r.row  '第一個位置
                Cells(r.row, 1).Select
                Selection.End(xlDown).Select
                Selection.End(xlToRight).Select
                ary(14) = Selection.row  '末的位置
                ary(15) = Selection.Column
                'MsgBox 5
                
                sht2c = 1 'index of sheets2 column,且必須重置才擺這邊
                For i = 1 To 5 'divide upon case to 5
                    'MsgBox 6
                    Range( _
                        Cells(ary(3 * (i - 1) + 1), 1), _
                            Cells(ary(3 * (i - 1) + 2), ary(3 * (i - 1) + 3)) _
                                ).Select
                    'ary(3 * (i - 1) + 1) =1    ary(3 * (i - 1) + 2) =2     ary(3 * (i - 1) + 3) =3
                    Selection.Copy
                    Sheets("Sheet2").Activate
                    Cells(sht2r, sht2c).Select
                    ActiveSheet.Paste
                    sht2c = sht2c + ary(3 * (i - 1) + 3) - 1 + 4    'i=i+四格 [ary(3 * (i - 1) + 3)-1]=length column
                    Sheets("Sheet1").Activate
                Next
                    sht2r = sht2r + ary(2) - ary(1) + 4 'i=i+四格 [ary(2) - ary(1)]=length row
        End Select
    Next
End Sub
