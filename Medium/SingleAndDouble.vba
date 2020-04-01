Sub effectnumber()
    Dim sin1 As Single
    Dim dou1 As Double
    
        sin1 = 10 / 3
        dou1 = 10 / 3
    
        [A1] = "Single"
        [B1] = sin1
        [A2] = "Double"
        [B2] = dou1
        [B1:B2].NumberFormat = "0.000000000000000000"
End Sub
