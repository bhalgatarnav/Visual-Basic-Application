' Function that calculated the score of the product based on the popularity, profit margin and affordability.
Function CalculateScore(Popularity As Double, ProfitMargin As Double, Affordability As Double) As Double
    CalculateScore = 0.4 * Popularity + 0.3 * ProfitMargin + 0.3 * Affordability
End Function

Sub Decision_Matrix()
'Cleaning the values.
ActiveSheet.Range("G2:K68").ClearContents
ActiveSheet.Range("H2:M68").ClearContents
ActiveSheet.Cells(4,11).ClearContents

' Removing the colour codes.
ActiveSheet.Range("H2:H68").Interior.Color = xlNone

    ' Initialising the arrays.
    Dim Popularity(), ProfitMargin(), Affordability() As Double
    Dim row, col, count As Integer
    count = 0
    row = 2
    col = 2
    Do While (ActiveSheet.Cells(row,col).Value <> "")
        count = count + 1
        Popularity(count, 1) = ActiveSheet.Cells(row, 2).Value
        ProfitMargin(count, 1) = ActiveSheet.Cells(row, 3).Value
        Affordability(count, 1) = ActiveSheet.Cells(row, 4).Value
        row = row + 1
    Loop

 

' Calculating the score of the products and then storing it in an array.
    Dim i As Integer
    Dim Score(), pop, pro, aff As Double
    ReDim Score(1 To UBound(Popularity, 1), 1 To 1)

    For i = 1 To 67 Step 1
        pop = Popularity(i, 1)
        pro = ProfitMargin(i, 1)
        aff = Affordability(i, 1)
        ' There is an type mismatch error in this statement.
        Score(i, 1) = CalculateScore(pop, pro, aff)
    Next 

' Finding the Median and then marking it 
    Dim Median As Double
    Median = Application.WorksheetFunction.Median(Scores())
    ActiveSheet.Cells(4,11).Value = Median

    For i = 1 To 67 Step 1
        ActiveSheet.Cells(i + 1, 7).Value = Score(i, 1)
        If Score(i, 1) < Median Then
            ActiveSheet.Cells(i + 1, 8).Value = "Retire"
            ActiveSheet.Cells(i + 1, 8).Interior.Color = RGB(0, 255, 0)
        Else
            ActiveSheet.Cells(i + 1, 8).Value = "Keep"
            ActiveSheet.Cells(i + 1, 8).Interior.Color = RGB(255, 0, 0)
        End If
    Next 


End Sub
