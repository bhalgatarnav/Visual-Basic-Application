' Function "Side" that will calculate the side length by taking 
' the radius of circumcircle as the input parameter
Function Side(r As Double) As Double
    Side = 2 * r * Sin(3.14159 / 3)
End Function

' A function called “Semi” to find the semi-perimeter of the equilateral triangle.
' semi-perimeter is the sum of all the sides divided by 2.
Function Semi(s As Double) As Double
    Semi = (3 * s) / 2
End Function

' A function called “Area” to find the area of the triangle.
Function Area(sd As Double, s As Double) As Double
    Area = Sqr(s * (s-side) * (s-side) * (s-side))
End Function

' Using all of the above funcitons in the excel file to calculate the answers.
        