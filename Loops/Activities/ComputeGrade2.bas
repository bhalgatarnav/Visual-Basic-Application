Sub ComputeGrade2()

    Dim hw1, hw2, hw3, Ahw As Double
    Dim iniRow, cA, cB, cC, cD, cF As Integer
    Dim grade As Double
    Dim letter As String
        cA = 0
        cB = 0
        cC = 0
        cD = 0
        cF = 0
    
    For iniRow = 3 To 22 Step 1
    
        grade = 0
        Ahw = 0
    

    
        hw1 = ActiveSheet.Cells(iniRow, 7)
        hw2 = ActiveSheet.Cells(iniRow, 8)
        hw3 = ActiveSheet.Cells(iniRow, 9)
        
        Ahw = (hw1 + hw2 + hw3) / 3
    
        ActiveSheet.Cells(iniRow, 10) = Ahw
        
        grade = 0.2 * Ahw + 0.25 * ActiveSheet.Cells(iniRow, 11) + 0.35 * ActiveSheet.Cells(iniRow, 12) + 0.2 * ActiveSheet.Cells(iniRow, 13)
    
        ActiveSheet.Cells(iniRow, 14) = grade
    
        If grade >= 90 Then
            letter = "A"
            cA = cA + 1
            ActiveSheet.Cells(iniRow, 15) = letter
    
        ElseIf grade >= 80 Then
            letter = "B"
            cB = cB + 1
            ActiveSheet.Cells(iniRow, 15) = letter
    
        ElseIf grade >= 70 Then
            letter = "C"
            cC = cC + 1
            ActiveSheet.Cells(iniRow, 15) = letter
    
        ElseIf grade >= 60 Then
            letter = "D"
            cD = cD + 1
            ActiveSheet.Cells(iniRow, 15) = letter
    
        Else
            letter = "F"
            cF = cF + 1
            ActiveSheet.Cells(iniRow, 15) = letter
        
        End If
    
    Next
    
    ActiveSheet.Cells(3, 3) = cA
    ActiveSheet.Cells(4, 3) = cB
    ActiveSheet.Cells(5, 3) = cC
    ActiveSheet.Cells(6, 3) = cD
    ActiveSheet.Cells(7, 3) = cF
    
End Sub
