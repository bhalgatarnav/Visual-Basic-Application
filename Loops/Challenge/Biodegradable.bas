Sub Biodegradable()
' Initialising the Variables to use
Dim cTeam1, cTeam2, cTeam3, count As Integer
Dim row, col As Integer

' Clearing the contents
ActiveSheet.Range("K6:K20").ClearContents
ActiveSheet.Range("L6:L20").ClearContents
ActiveSheet.Range("M6:M20").ClearContents

' Clearing the colour coding
ActiveSheet.Range("B6:B25").Interior.Color = xlNone
ActiveSheet.Range("C6:C25").Interior.Color = xlNone
ActiveSheet.Range("D6:D25").Interior.Color = xlNone

' Count the number of enteries the team had entered:
cTeam1 = WorksheetFunction.Count(Range("B6:B25"))
cTeam2 = WorksheetFunction.Count(Range("C6:C25"))
cTeam3 = WorksheetFunction.Count(Range("D6:D25"))

If (cTeam1<11) Then 
    ActiveSheet.Cells(6,11).Value = "NMT"
Else
    Do While (ActiveSheet.Cells(row, col) <> "" And count<16)
        
    Loop

End If 


If (cTeam2<11) Then 
    ActiveSheet.Cells(6,12).Value = "NMT"
Else

End If 


If (cTeam3<11) Then 
    ActiveSheet.Cells(6,13).Value = "NMT"
Else

End If 

End Sub
