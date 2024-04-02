Sub Game()
    ' Initialise the variables
    Dim N As Integer
    Dim Max As Double
    Dim num As Integer
    Dim Value As Integer

    ' Giving Variables the initial values.
    num = 1
    Max = 0
    iniRow = 6
    iniCol = 3

    
    ' Assigning the values
    N = ActiveSheet.Cells(2, 4).Value

    ActiveSheet.Cells(6, 3).Value = N
    
    ' Loops
    While N <> 1
        If N Mod 2 = 0 Then
            N = N / 2
        Else
            N = 3 * N + 1
        End If
        num = num + 1
    
        If N > Max Then
            Max = N
        End If
        iniCol = iniCol + 1
        ActiveSheet.Cells(6, iniCol).Value = N

    Wend

   
    If num > 100 Then
        MsgBox "Too many iterations"
    Else
    ' Setting the values
     ActiveSheet.Cells(3, 4).Value = Max
     ActiveSheet.Cells(4, 4).Value = num
    
    End If
    
    End Sub
        
    