Sub Einstein()

    ActiveSheet.Cells(5,3).Value= 3*(10)^8
    
    'Initialising the Variables
    Dim time_OnE As Double
    Dim velocity As String
    Dim c As Double
    

    ' Take the necessary inputs;
    time_OnE = ActiveSheet.Cells(3,3).Value
    velocity = ActiveSheet.Cells(4,3)
    c = ActiveSheet.Cells(5,3).Value
    
    'Making the changes to the string.    
    velocity = Left(velocity, Len(velocity) - 1)
    velocity = CDbl(velocity)
     
    Dim Ans  As As Double
    Ans = time_OnE*(sqr(1-((velocity)^2)))
    ActiveSheet.Cells(9,3).Value= round(Ans,2)



End Sub
