Sub Safety_Factor()
    Dim yeildStrength As Double
    Dim designLoad As Double
    Dim area As Double
    Dim safetyFactor As Double

    ' Emptying the cells for restarting the program
    ActiveSheet.Cells(4,2).Value = ""
    ActiveSheet.Cells(4,3).Value = ""
    ActiveSheet.Cells(4,4).Value = ""

    ' Assigning the variables to the values in the excel sheet
    yeildStrength = ActiveSheet.Cells(1,2).Value
    yeildStrength = yeildStrength
    designLoad = ActiveSheet.Cells(2,2).Value
    area = ActiveSheet.Cells(3,2).Value

    ' Calculating the safety factor
    safetyFactor = (yeildStrength/(designLoad/area))

    ' Printing our the message corresponding to the safety factor
    Dim m As String
    
    If safetyFactor<1 then
        m = "The design is in danger"
        ActiveSheet.Cells(4,2).Value = m
    ElseIf safetyFactor=1 then
        m = "The design is safe"
        ActiveSheet.Cells(4,3).Value = m
    Else
        m = "The design is safe but is Over-Engineered "
        ActiveSheet.Cells(4,4).Value = m
    End If



End Sub
