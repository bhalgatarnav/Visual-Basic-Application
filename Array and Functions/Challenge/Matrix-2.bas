' Function that calculated the score of the product based on the popularity, profit margin and affordability.
    Function Calculate(p As Double, po As Double, A As Double) As Double
        Dim s As Double
        
        s = (0.4 * p) + (0.3 * po) + (0.3 * A)
        
        Calculate = s
    End Function
    
    Sub Decision_Matrix()
    'Cleaning the values.
    ActiveSheet.Range("G2:G68").ClearContents
    ActiveSheet.Range("H2:H68").ClearContents
    ActiveSheet.Cells(4, 11).ClearContents
    
    ' Removing the colour codes.
    ActiveSheet.Range("H2:H68").Interior.Color = xlNone
    

     ' Initialising the arrays.
    Dim Popularity() As Double
    Dim ProfitMargin() As Double
    Dim Affordability() As Double
    Dim Score() As Double

    Dim row, col, count As Integer
    Dim sc As Double
    count = 0
    row = 2
    col = 2

        
        Do While (ActiveSheet.Cells(row, col).Value <> "")
            ReDim Preserve Popularity(0 To row)
            ReDim Preserve ProfitMargin(0 To row)
            ReDim Preserve Affordability(0 To row)
            ReDim Preserve Score(0 To row)
            
            Popularity(count) = ActiveSheet.Cells(row, 2).Value
            'ActiveSheet.Cells(row, 12).Value = Popularity(count)
            ProfitMargin(count) = ActiveSheet.Cells(row, 3).Value
            'ActiveSheet.Cells(row, 13).Value = ProfitMargin(count)
            Affordability(count) = ActiveSheet.Cells(row, 4).Value
            'ActiveSheet.Cells(row, 14).Value = Affordability(count)
            
            sc = Calculate(Popularity(count), ProfitMargin(count), Affordability(count))
            Score(count) = sc
            ActiveSheet.Cells(row, 15).Value = sc

            row = row + 1
            count = count + 1
        Loop
    
     
    ' Finding the Median and then marking it
        Dim Median As Double
        Median = Application.WorksheetFunction.Median(Score())
        ActiveSheet.Cells(4, 11).Value = Median
    
        For i = 2 To row Step 1
            ActiveSheet.Cells(i, 7).Value = Score(i - 2)
            If Score(i - 2) < Median Then
                ActiveSheet.Cells(i, 8).Value = "Retire"
                ActiveSheet.Cells(i, 8).Interior.Color = RGB(255, 0, 0)
            Else
                ActiveSheet.Cells(i, 8).Value = "Keep"
                ActiveSheet.Cells(i, 8).Interior.Color = RGB(0, 255, 0)
            End If
        Next
    
    
    End Sub


