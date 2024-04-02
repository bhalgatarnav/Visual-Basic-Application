Sub WeatherProcessing()

    Dim Temp As Double
    Dim Rain As Double
    Dim Wind As Double

    ' Take in the value from the worksheet.
    Temp = ActiveSheet.Cells(4, 3).Value
    Rain = ActiveSheet.Cells(5, 3).Value
    Wind = ActiveSheet.Cells(6, 3).Value

    If (Temp > 25 And Rain < 0.1 And Wind < 0.1) Then
        ActiveSheet.Cells(3, 4).Value = "Sunny"
    ElseIf (Temp > 20 And Rain < 0.2 And Wind < 0.2) Then
        ActiveSheet.Cells(3, 4).Value = "Cloudy"
    ElseIf (Temp > 15 And Rain < 0.3 And Wind < 0.3) Then
        ActiveSheet.Cells(3, 4).Value = "Rainy"
    Else
        ActiveSheet.Cells(3, 4).Value = "Stormy"
    End If

End Sub
