Sub Biodegradable()
' Initialising the Variables to use
Dim cTeam1, cTeam2, cTeam3, count As Integer
Dim row, col As Integer
Dim SF As Double

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

row = 6
col = 1
for col = 1 to 4 Step 1

    Do While (ActiveSheet.Cells(row, col) <> "" And count<16)
        SF = ActiveSheet.Cells(row, col).Value/ActiveSheet.Cells(7,7).Value
        ActiveSheet.Cells(row, col+10).Value = SF
        If (SF>1.2) Then 
        ActiveSheet.Cells(row, col+10).Interior.Color = RGB(255, 0, 0)
        
        Else If (SF<1) Then
        ActiveSheet.Cells(row, col+10).Interior.Color = RGB(255, 255, 153)
        
        Else
        ActiveSheet.Cells(row, col+10).Interior.Color = RGB(0, 255, 0)

        End If

        count = count + 1
        row = row + 1
    Loop

End If  

Next

' Looping through the teams
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
