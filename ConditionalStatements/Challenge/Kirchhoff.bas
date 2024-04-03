Sub Kirchhoff()

    ' Declare the variables
    Dim r1, r2, r3, r4, r5, rtotal, vr As Double
    Dim Vr1, Vr2, Vr3, Vr4, Vr5, vtotal, v As Double
    Dim Ir1, Ir2, Ir3, Ir4, Ir5, itotal, i As Double
    Dim Count As As Integer
    Dim combination As String

    ' Clear the Output Area
    ActiveSheet.Cells(4, 11) = ""
    ActiveSheet.Range("K9:K13").ClearContents
    ActiveSheet.Range("L9:L13").ClearContents

    ' Checking the number of Resistances entered.
    Count = WorksheetFunction.Count(Range("D4:D8"))
    vtotal = ActiveSheet.Cells(11,4)
    itotal = ActiveSheet.Cells(14,4)
    combination = ActiveSheet.Cells(10,7)

    If (Count = 2 or Count = 5) Then 
    ' If the number of resistances entered is 2 or 5, then calculate the Resistance
        If (Count = 2) Then
        r1 = ActiveSheet.Cells(4,4)
        r2 = ActiveSheet.Cells(5,4)

            If (combination="parallel") Then 
            ' If the resistances are parallel, then calculate the Resistance
                rtotal = 1/((1/r1)+(1/r2))
                ActiveSheet.Cells(4,11) = rtotal

                ActiveSheet.Range("K9:K10").Value = vtotal

                ActiveSheet.Cells(9,12) = vtotal/r1
                ActiveSheet.Cells(10,12) = vtotal/r2

            Else
            ' If the resistances are in series, then calculate the Resistance
                rtotal = r1 + r2
                ActiveSheet.Cells(4,11) = rtotal

                ActiveSheet.Range("L9:L10").Value = itotal

                ActiveSheet.Cells(9,12) = vtotal*r1/rtotal
                ActiveSheet.Cells(10,12) = vtotal*r2/rtotal

 
            End If
    
        Else
            r1 = ActiveSheet.Cells(4,4)
            r2 = ActiveSheet.Cells(5,4)
            r3 = ActiveSheet.Cells(6,4)
            r4 = ActiveSheet.Cells(7,4)
            r5 = ActiveSheet.Cells(8,4)


    
        End If

    Else
    MsgBox "Please enter values for either 2 or 5 Resistors." 
    End If
End Sub