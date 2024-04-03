Sub Kirchhoff()

    ' Declare the variables
    Dim r1, r2, r3, r4, r5, rtotal, vr As Double
    Dim Vr1, Vr2, Vr3, Vr4, Vr5, vtotal, v As Double
    Dim Ir1, Ir2, Ir3, Ir4, Ir5, itotal, i As Double
    Dim Count As Integer
    Dim combination As String

    ' Clear the Output Area
    ActiveSheet.Cells(4, 11) = ""
    ActiveSheet.Cells(14, 4) = ""
    ActiveSheet.Range("K9:K13").ClearContents
    ActiveSheet.Range("L9:L13").ClearContents

    ' Checking the number of Resistances entered.
    Count = WorksheetFunction.Count(Range("D4:D8"))
    vtotal = ActiveSheet.Cells(11,4)

    combination = ActiveSheet.Cells(10,7)

    If (Count = 2 or Count = 5) Then 
    ' If the number of resistances entered is 2 or 5, then calculate the Resistance
        If (Count = 2) Then
        r1 = ActiveSheet.Cells(4,4)
        r2 = ActiveSheet.Cells(5,4)

            If (combination="Parallel") Then 
            ' If the resistances are parallel, then calculate the Resistance
                rtotal = 1/((1/r1)+(1/r2))
                ActiveSheet.Cells(4,11) = rtotal
                
                itotal = vtotal/rtotal

                ActiveSheet.Range("K9:K10").Value = vtotal
                ActiveSheet.Cells(14,4).Value = itotal

                ActiveSheet.Cells(9,12) = vtotal/r1
                ActiveSheet.Cells(10,12) = vtotal/r2

            Else
            ' If the resistances are in series, then calculate the Resistance
                rtotal = r1 + r2
                ActiveSheet.Cells(4,11) = rtotal

                itotal = vtotal/rtotal
                ActiveSheet.Cells(14,4) = itotal

                ActiveSheet.Range("L9:L10").Value = itotal

                ActiveSheet.Cells(9,11) = itotal * r1
                ActiveSheet.Cells(10,11) = itotal * r2

 
            End If
    
        Else
            r1 = ActiveSheet.Cells(4,4)
            r2 = ActiveSheet.Cells(5,4)
            r3 = ActiveSheet.Cells(6,4)
            r4 = ActiveSheet.Cells(7,4)
            r5 = ActiveSheet.Cells(8,4)

            If (combination="Parallel") Then 
            ' If the resistances are parallel, then calculate the Resistance
                rtotal = 1/((1/r1)+(1/r2)+(1/r3)+(1/r4)+(1/r5))
                ActiveSheet.Cells(4,11) = rtotal

                itotal = vtotal/rtotal
                ActiveSheet.Cells(14,4).Value = itotal
                ActiveSheet.Range("K9:K13").Value = vtotal

                ActiveSheet.Cells(9,12) = vtotal/r1
                ActiveSheet.Cells(10,12) = vtotal/r2
                ActiveSheet.Cells(11,12) = vtotal/r3
                ActiveSheet.Cells(12,12) = vtotal/r4
                ActiveSheet.Cells(13,12) = vtotal/r5

            Else
            ' If the resistances are in series, then calculate the Resistance
                rtotal = r1 + r2 + r3 + r4 + r5
                ActiveSheet.Cells(4,11) = rtotal

                itotal = vtotal/rtotal
                ActiveSheet.Cells(14,4) = itotal

                ActiveSheet.Range("L9:L13").Value = itotal

                ActiveSheet.Cells(9,11) = itotal * r1
                ActiveSheet.Cells(10,11) = itotal * r2
                ActiveSheet.Cells(11,11) = itotal * r3
                ActiveSheet.Cells(12,11) = itotal * r4
                ActiveSheet.Cells(13,11) = itotal * r5

 
            End If

    
        End If

    Else
    MsgBox "Please enter values for either 2 or 5 Resistors." 
    End If
End Sub