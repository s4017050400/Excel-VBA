Option Explicit

Public Sub test1()
    Dim myRange As Range
    
    For Each myRange In Range("D2:F6")
        myRange.Value = "人"
    Next myRange
End Sub



