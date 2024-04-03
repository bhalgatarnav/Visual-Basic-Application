Sub WeatherProcessing()

    ' Variables
        Dim NumInvalid As Double
        Dim MaxSnw, snw, snwd As Double
        Dim MaxDPreci As Double
        Dim preci As Double
        Dim Hot, vhot As Double
        Dim AHot, totalh As Double
        Dim Cold, vcold As Double
        Dim ACold, totalc As Double
        Dim counter As Integer
        
    ' Loop Structure
    
    Dim yr As Integer
    Dim iniRow As Integer
    Dim iniCol As Integer
    
    iniRow = 11
    iniCol = 1
    snw = 0
    preci = 0
    totalh = 0
    totalc = 0
    counter = 0
    vcold = 100
    NumInvalid = 1
    
    
    
    
    
     yr = ActiveSheet.Cells(iniRow, iniCol)
    
    For iniRow = 12 To 25212 Step 1
    
        snw = ActiveSheet.Cells(iniRow, 5).Value
        snwd = ActiveSheet.Cells(iniRow, 6).Value
         If (snw = -9999) Then
            NumInvalid = NumInvalid + 1
        End If
        
        If (snwd = -9999) Then
         NumInvalid = NumInvalid + 1
        End If
        
        
            If (MaxSnw < snw) Then
                MaxSnw = snw
            End If
        
    
        preci = ActiveSheet.Cells(iniRow, 4).Value
    
        If (MaxDPreci < preci) Then
            MaxDPreci = preci
        End If
    
        vhot = ActiveSheet.Cells(iniRow, 7).Value
        vcold = ActiveSheet.Cells(iniRow, 8).Value
        totalh = totalh + vhot
        totalc = totalc + vcold
        
        If (Hot < vhot) Then
            Hot = vhot
        End If
    
        If (Cold > vcold) Then
            Cold = vcold
        End If
        
         iniRow = iniRow + 1
         counter = counter + 1
    
        yr = ActiveSheet.Cells(iniRow, iniCol)
    
    Next
    
     AHot = totalh / counter
     ACold = totalc / counter
     
     ActiveSheet.Cells(2, 2).Value = NumInvalid
     ActiveSheet.Cells(3, 2).Value = MaxSnw
     ActiveSheet.Cells(4, 2).Value = MaxDPreci
     ActiveSheet.Cells(5, 2).Value = Hot
     ActiveSheet.Cells(6, 2).Value = Cold
     ActiveSheet.Cells(7, 2).Value = AHot
     ActiveSheet.Cells(8, 2).Value = ACold
    
    
    End Sub
    
    
    