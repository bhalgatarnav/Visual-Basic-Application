Sub StandardDev()

    ' Declaring the Variables:

    Dim x1 As Double
    Dim x2 As Double
    Dim x3 As Double
    Dim x4 As Double
    Dim x5 As Double

    ' Reading the values of the cells into the variables just created.
    x1 = ActiveSheet.Cells(6, 4).Value
    x2 = ActiveSheet.Cells(7, 4).Value
    x3 = ActiveSheet.Cells(8, 4).Value
    x4 = ActiveSheet.Cells(9, 4).Value
    x5 = ActiveSheet.Cells(10, 4).Value

    ' Calculating the mean of the values.
    Dim mean As Double
    mean = WorksheetFunction.Average(Range("D6:D10"))

    ' Calculating the sum of squares of the difference between the values and the mean.
    Dim sumOfSquares As Double
    sumOfSquares = (x1 - mean) ^ 2 + (x2 - mean) ^ 2 + (x3 - mean) ^ 2 + (x4 - mean) ^ 2 + (x5 - mean) ^ 2

    ' Calculating the standard deviation.
    Dim standardDeviation As Double
    standardDeviation = Sqr(sumOfSquares / 4)
    ' Displaying the result.
    ActiveSheet.Cells(7, 7).Value = standardDeviation



End Sub