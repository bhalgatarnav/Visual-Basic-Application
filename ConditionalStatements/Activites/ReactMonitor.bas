Sub ReactorMonitor()

    Dim Temp As Double
    Dim Pressure As Double

    ' Take in the value from the worksheet.
    Temp = ActiveSheet.Cells(4, 3).Value
    Pressure = ActiveSheet.Cells(5, 3).Value

If (Temp > 355 Or Pressure > 0.1) Then
        ActiveSheet.Cells(3, 4).Value = "Melt Down"

ElseIf (Temp > 345 Or Pressure > 0.095) Then
        ActiveSheet.Cells(3, 4).Value = "Very Severe"

ElseIf (Temp > 335 Or Pressure > 0.09) Then
        ActiveSheet.Cells(3, 4).Value = "Severe"
    
ElseIf (Temp > 325 Or Pressure > 0.085) Then
        ActiveSheet.Cells(3, 4).Value = "Moderate"

Else
        ActiveSheet.Cells(3, 4).Value = "Normal"
    End If
    
End Sub
    