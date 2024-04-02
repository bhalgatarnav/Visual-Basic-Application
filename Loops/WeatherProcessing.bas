Sub WeatherProcessing()
' Variables
    Dim NumInvalid As Double
    Dim MaxSnw, snw, snwd As Double
    Dim MaxDPreci As Double
    Dim Hot As Double
    Dim AHot As Double
    Dim Cold As Double
    Dim ACold As Double
    
' Loop Structure

Dim yr As Integer
Dim iniRow As Integer
Dim iniCol As Integer

iniRow = 11
iniCol = 1

' Loop
yr = ActiveSheet.Cells(iniRow,iniCol);
While yr<2019
    snw = ActiveSheet.Cells(iniRow, 5).Value
    snwd = ActiveSheet.Cells(iniRow, 6).Value
    If (snw==-9999 or snwd==-9999) Then
        NumInvalid = NumInvalid + 1
    
    Else
        if (MaxSnw<snw) then
            MaxSnw = snw
        End if
    End If

    
    


    wend 

End Sub
