'5五次連續貼
Public Sub paste()
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("iddx_SAMPLE_5VDD_used (Recovered).xlsm").Activate
    Selection.End(xlToLeft).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Range("B1").Select
    ActiveSheet.paste
    ThisWorkbook.Activate
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("iddx_SAMPLE_5VDD_used (Recovered).xlsm").Activate
    Selection.End(xlToLeft).Select
    Selection.End(xlDown).Select
    Range("B128").Select
    ActiveSheet.paste
    ThisWorkbook.Activate
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("iddx_SAMPLE_5VDD_used (Recovered).xlsm").Activate
    Selection.End(xlToLeft).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B256").Select
    ActiveSheet.paste
    ThisWorkbook.Activate
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("iddx_SAMPLE_5VDD_used (Recovered).xlsm").Activate
    Selection.End(xlToLeft).Select
    Range("A257").Select
    Selection.End(xlDown).Select
    Range("B384").Select
    ActiveSheet.paste
    ThisWorkbook.Activate
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("iddx_SAMPLE_5VDD_used (Recovered).xlsm").Activate
    Selection.End(xlToLeft).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B512").Select
    ActiveSheet.paste

End Sub



'標題貼
Public Sub main()
    Application.ScreenUpdating = False  '不要在螢幕上顯示過程
    For Row = 2 To 4 'for row
        For Count = 1 To 160 'for col
            Sheets("-40C_AC").Activate
            x = Cells(Row, 3 * (Count - 1) + 13).Value
            Sheets("-40C_DC_400MHz").Activate
            Cells(Row, 5 * (Count - 1) + 8).Value = x
        Next Count
    Next Row
    Application.ScreenUpdating = True
End Sub
