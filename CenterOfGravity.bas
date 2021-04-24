Attribute VB_Name = "Module1"
Public Function Haversine(Lat1 As Variant, Lon1 As Variant, Lat2 As Variant, Lon2 As Variant)
Dim R As Integer, dlon As Variant, dlat As Variant, Rad1 As Variant
Dim a As Variant, c As Variant, d As Variant, Rad2 As Variant

R = 6371
dlon = Excel.WorksheetFunction.Radians(Lon2 - Lon1)
dlat = Excel.WorksheetFunction.Radians(Lat2 - Lat1)
Rad1 = Excel.WorksheetFunction.Radians(Lat1)
Rad2 = Excel.WorksheetFunction.Radians(Lat2)
a = Sin(dlat / 2) * Sin(dlat / 2) + Cos(Rad1) * Cos(Rad2) * Sin(dlon / 2) * Sin(dlon / 2)
c = 2 * Excel.WorksheetFunction.Atan2(Sqr(1 - a), Sqr(a))
d = R * c
Haversine = d

End Function

Sub cog()
Sheet1.Activate
Dim n As Double
Dim myValue As Double

n = InputBox("Insert the number of warehouse to open :")
myValue = Sheet1.Range("A" & Rows.Count).End(xlUp).Row - 1

Dim data() As Variant
ReDim data(myValue, 3)
For i = 1 To myValue
    data(i, 1) = Sheet1.Cells(i + 1, 1)
    data(i, 2) = Sheet1.Cells(i + 1, 2)
    data(i, 3) = Sheet1.Cells(i + 1, 3)
Next i

Dim ploc() As Double
Dim test() As Double
ReDim ploc(n, 2) As Double
ReDim test(n, 2) As Double

For i = 1 To n
    ploc(i, 1) = 48.2822 + i * 0.18
    ploc(i, 2) = 7.4583 + i * 0.17
Next i


For tg = 1 To 5
    Dim compmatrix() As Variant
    Dim assigned() As Variant
    ReDim compmatrix(myValue, n) As Variant
    ReDim assigned(myValue, 4) As Variant

    For i = 1 To myValue
        For j = 1 To n
                compmatrix(i, j) = Haversine(data(i, 1), data(i, 2), ploc(j, 1), ploc(j, 2))
        Next j
            
        im = 9999999
        Index = 1
        For k = 1 To n
            If compmatrix(i, k) < im Then
                im = compmatrix(i, k)
                Index = k
            End If
        Next k
        
        assigned(i, 1) = data(i, 1)
        assigned(i, 2) = data(i, 2)
        assigned(i, 3) = data(i, 3)
        assigned(i, 4) = Index
        
    Next i
    
    For i = 1 To myValue
        Sheet1.Cells(i + 2, 14) = assigned(i, 1)
        Sheet1.Cells(i + 2, 15) = assigned(i, 2)
        Sheet1.Cells(i + 2, 16) = assigned(i, 3)
        Sheet1.Cells(i + 2, 17) = assigned(i, 4)
        Next i
    
    test = ploc
    
    For x = 1 To n
    
        Dim latarray() As Double
        Dim lonarray() As Double
        Dim weightarray() As Double
        ReDim latarray(myValue) As Double
        ReDim lonarray(myValue) As Double
        ReDim weightarray(myValue) As Double
        t = 1
        For y = 1 To myValue
            If assigned(y, 4) = x Then
                latarray(t) = assigned(y, 1)
                lonarray(t) = assigned(y, 2)
                weightarray(t) = assigned(y, 3)
                t = t + 1
            End If
        Next y
                
        Dim d() As Variant
        Dim topx, downx, topy, downy As Double
        Dim ag, bg As Double
        
        ag = 45
        bg = 4
        
        For k = 1 To 50
            topx = 0
            downx = 0
            topy = 0
            downy = 0
            TC = 0
            For i = 1 To t - 1
                ReDim d(i)
                d(i) = Haversine(latarray(i), lonarray(i), ag, bg)
                If d(i) = 0 Then GoTo Line Else GoTo Layn
Layn:               topx = topx + (weightarray(i) * latarray(i)) / d(i)
                    downx = downx + (weightarray(i) / d(i))
                    topy = topy + (weightarray(i) * lonarray(i)) / d(i)
                    downy = downy + (weightarray(i) / d(i))
                    TC = TC + weightarray(i) * d(i)
             
Line:
            Next i
        
        ag = topx / downx
        bg = topy / downy
        Next k
        
        ploc(x, 1) = ag
        ploc(x, 2) = bg
              
    Next x
Next tg

Dim ptage() As Variant
ReDim ptage(n, 1) As Variant
Dim toplam As Double
toplam = 0
    For i = 1 To n
        For j = 1 To myValue
            If Sheet1.Cells(j + 2, 17) = i Then
                ptage(i, 1) = ptage(i, 1) + Sheet1.Cells(j + 2, 16)
            End If
        Next j
        toplam = ptage(i, 1) + toplam
    Next i
            
For t = 1 To n
    Sheet1.Cells(15 + t, 6) = t
    Sheet1.Cells(15 + t, 7) = ploc(t, 1)
    Sheet1.Cells(15 + t, 8) = ploc(t, 2)
    Sheet1.Cells(15 + t, 9) = ptage(t, 1) / toplam
Next t
End Sub
