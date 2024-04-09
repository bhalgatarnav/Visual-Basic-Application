    ' Calculated the net resistance of the resistors connected in series
    Function RSeries(result As Range) As Double
    Dim i As Range
    Dim req As Double
    
    req = 0
    For Each i In result
        req = req + i
    Next
    RSeries = req
    End Function
    
    ' Calculated the net resistance of the resistors connected in parallel
    Function RParallel(result As Range) As Double
    Dim te As Double
    Dim i As Range
    
    te = 0
    For Each i In result
        te = te + (1 / i)
    Next
    
    RParallel = 1 / te
    End Function
    
    