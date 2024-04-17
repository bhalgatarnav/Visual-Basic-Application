' Function "Side" that will calculate the side length by taking
' the radius of circumcircle as the input parameter
Function side(r As Double) As Double
    side = 2 * r * Sin(Application.WorksheetFunction.Pi() / 3)
End Function

' A function called “Semi” to find the semi-perimeter of the equilateral triangle.
' semi-perimeter is the sum of all the sides divided by 2.
Function semi(si As Double) As Double
    semi = (3 * si) / 2
End Function

' A function called “Area” to find the area of the triangle.

Function Area(sd As Double, s As Double) As Double
    Area = Sqr(s * (s - sd) * (s - sd) * (s - sd))
End Function

Sub Heron()
 ' Calculating the Area of Triangle Using the above functions
 Dim TriSide, TriSemi, TriArea, In1 As Double
    
 In1 = ActiveSheet.Cells(5, 4).Value
 TriSide = side(In1)

 ActiveSheet.Cells(5, 7).Value = TriSide
 ActiveSheet.Cells(5, 9).Value = Area(CDbl(TriSide), semi(CDbl(TriSide)))
 

End Sub

