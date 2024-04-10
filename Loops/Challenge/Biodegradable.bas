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
ActiveSheet.Range("K6:K25").Interior.Color = xlNone
ActiveSheet.Range("L6:L25").Interior.Color = xlNone
ActiveSheet.Range("M6:M25").Interior.Color = xlNone

' Count the number of enteries the team had entered:
cTeam1 = WorksheetFunction.count(Range("B6:B25"))
cTeam2 = WorksheetFunction.count(Range("C6:C25"))
cTeam3 = WorksheetFunction.count(Range("D6:D25"))
count = 0


For col = 2 To 4 Step 1
count = 0
row = 6

    Do While (ActiveSheet.Cells(row, col).Value <> "" And count <= 14)
        count = count + 1
        SF = 0
        SF = ActiveSheet.Cells(row, col).Value / ActiveSheet.Cells(7, 7).Value
        SF = round(SF, 2) 
        ActiveSheet.Cells(row, col + 9).Value = SF

        If (SF > 1.2) Then
        ActiveSheet.Cells(row, col + 9).Interior.Color = RGB(255, 0, 0)
        
        ElseIf (SF < 1) Then
        ActiveSheet.Cells(row, col + 9).Interior.Color = RGB(255, 255, 153)
        
        Else
        ActiveSheet.Cells(row, col + 9).Interior.Color = RGB(0, 255, 0)

        End If


        row = row + 1
    Loop

    If (count < 11) Then
        ActiveSheet.Range(Cells(6, col + 9), Cells(20, col + 9)).ClearContents
        ActiveSheet.Range(Cells(6, col + 9), Cells(20, col + 9)).Interior.Color = xlNone
        ActiveSheet.Cells(6, col + 9).Value = "NMT"
    End If

Next


End Sub
