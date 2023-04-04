Sub boton()
Call Sheet1.prueba
Call Sheet2.prueba
Call Sheet3.prueba


End Sub

Sub prueba()
Dim tic, ticnext As String
i = 2
j = 2
'tic and ticnext store the ticker symbol
ticnext = Cells(3, 1).Value
tic = Cells(2, 1).Value


While Not (tic = "")

If tic = ticnext Then
    'sumavol stores the volume of stock
   sumavol = Cells(i, 7).Value + sumavol
    'x helps identifying the position of the opening price
   x = x + 1


Else
'prints the table

  Cells(j, 9).Value = tic
  Cells(j, 12).Value = sumavol + Cells(i, 7).Value

  Cells(j, 10).Value = Cells(i, 6).Value - Cells(i - x, 3).Value

  'conditional for the color on column J
     If Cells(j, 10).Value > 0 Then
      Cells(j, 10).Interior.Color = RGB(0, 255, 0)
     ElseIf Cells(j, 10).Value < 0 Then
      Cells(j, 10).Interior.Color = RGB(255, 0, 0)
      Else
      Cells(j, 10).Interior.Color = RGB(255, 255, 255)
      End If

  Cells(j, 11).Value = FormatPercent(Cells(j, 10).Value / Cells(i - x, 3).Value)


  sumavol = 0
  x = 0


  j = j + 1

End If




tic = Cells(i + 1, 1).Value
ticnext = Cells(i + 2, 1).Value

i = i + 1


Wend


'Greatest table

k = 2
y = 2
per = Cells(2, 11).Value
volu = Cells(2, 12).Value
maxin = 0
maxde = 0
maxvol = 0
mx = 0
mn = 0
vl = 0


While Not (per = "")

If per > 0 Then
 'greatest increase
 If per > maxin Then
 maxin = per
 mx = y
 End If
 
Else
 'Greatest decrease
 If per < maxde Then
 maxde = per
 mn = y
 End If
 
End If

'Greatest volume
If volu > maxvol Then
 maxvol = volu
 vl = y
End If


 
 y = y + 1
per = Cells(y, 11).Value
volu = Cells(y, 12).Value

Wend

'prints greatest tickers
Range("P2").Value = Cells(mx, 9).Value
Range("P3").Value = Cells(mn, 9).Value
Range("P4").Value = Cells(vl, 9).Value

'prints greatest values
Range("Q2").Value = FormatPercent(maxin)
Range("Q3").Value = FormatPercent(maxde)
Range("Q4").Value = maxvol

End Sub


