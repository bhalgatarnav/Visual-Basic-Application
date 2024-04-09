    Function RSeries(result() As Double) As Double
    RSeries = (result(0) + result(1) + result(2) + result(3) + result(4))
    End Function
    
    Function RParallel(result() As Double) As Double
    Dim xparallel As Double
    xparallel = ((1 / result(0)) + (1 / result(1)) + (1 / result(2)) + (1 / result(3)) + (1 / result(4)))
    RParallel = 1 / xparallel
    End Function
    
    
    Sub Reistors()
    Dim r(5) As Double
    r(0) = ActiveSheet.Cells(4, 3).Value
    r(1) = ActiveSheet.Cells(5, 3).Value
    r(2) = ActiveSheet.Cells(6, 3).Value
    r(3) = ActiveSheet.Cells(7, 3).Value
    r(4) = ActiveSheet.Cells(8, 3).Value
    ActiveSheet.Cells(2, 3).Value = RSeries(r)
    ActiveSheet.Cells(3, 3).Value = RParallel(r)
    End Sub
    