Sub Biking()
    Dim Ww As Double
    Ww = ActiveSheet.Cells(3, 3).Value

ActiveSheet.Cells(4, 6) = 60*Ww*10^(ActiveSheet.Cells(4,5)/25 - 1.85)
ActiveSheet.Cells(5, 6) = 60*Ww*10^(ActiveSheet.Cells(5,5)/25 - 1.85)
ActiveSheet.Cells(6, 6) = 60*Ww*10^(ActiveSheet.Cells(6,5)/25 - 1.85)

End sub    