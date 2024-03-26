Sub Moment()
    ' Agenda for the VBA:
    ' 1. Check if Mass is positve, if not assign the mass value to 1 and use MsgBox to tell the users that the mass has to be positive.
    ' 2. Check if the radius is positive, if not assign the radius value to 1 and use MsgBox to tell the users that the radius has to be positive.
    ' Use a conditional structure to calculate the inertia based on the object description.
    
    
        Dim Mass As Double
        Dim Radius As Double
        Dim Inertia As Double
        Dim ObjectType As String
        
        ActiveSheet.Cells(2, 5).Value = ""
        
       Mass = ActiveSheet.Cells(2, 1).Value
       Radius = ActiveSheet.Cells(2, 2).Value
       ObjectType = ActiveSheet.Cells(2, 3).Value
    
       If Mass < 0 Then
         MsgBox "Mass has to be positive"
         Mass = 1
        
       ElseIf Radius < 0 Then
            MsgBox "Radius has to be positive"
            Radius = 1
            
       ElseIf ObjectType = "Sphere" Then
            Inertia = 0.4 * Mass * (Radius) ^ 2
       ElseIf ObjectType = "Cylinder" Then
            Inertia = 0.5 * Mass * (Radius) ^ 2
       ElseIf ObjectType = "Hoop" Then
            Inertia = Mass * (Radius) ^ 2
       Else
            MsgBox "Object type not recognized"
       End If
       ActiveSheet.Cells(2, 5).Value = Inertia
    End Sub
    