' Function "Side" that will calculate the side length by taking 
' the radius of circumcircle as the input parameter
Function Side(r As Double) As Double
    Side = 2 * r * Sin(PI / 3)
End Function

' A function called “Semi” to find the semi-perimeter of the equilateral triangle.
' semi-perimeter is the sum of all the sides divided by 2.
Function Semi(s As Double) As Double
    Semi = (3 * s) / 2
End Function

' A function called “Area” to find the area of the triangle.
Function Area(side As Double, s As Double) As Double
    Area = Sqr(s * (s-side) * (s-side) * (s-side))

Sub Heron()
 ' Calculating the Area of Triangle Using the above functions
 Dim TriSide, TriSemi, TriArea As Double
 
 TriSide = Side(ActiveSheet.Cells(5,4).Value)
 TriSemi = Semi(TriSide)
 TriArea = Area(TriSide, TriSemi)
 
 ActiveSheet.Cells(5,7).Value = TriSide
 ActiveSheet.Cells(5,9).Value = TriArea
    
End Sub
