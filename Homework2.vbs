Sub WorksheetLoop()

Dim Current As Worksheet

For Each Current In Worksheets

  Dim Ticker As String
  Dim O_pen, C_lose, Stockvol As Double
  Stockvol = 0
  Dim Counter As Integer
  Counter = 2
  Dim rowdelimiter As Long
  
  rowdelimiter = Cells(Rows.Count, 1).End(xlUp).Row
  
  For I = 2 To rowdelimiter

    
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

      
      Ticker = Cells(I, 1).Value
      O_pen = Cells(I, 3).Value
      C_lose = Cells(I, 6).Value
      
      Stockvol = Stockvol + Cells(I, 7).Value

      Range("I" & Counter).Value = Ticker
      Range("J" & Counter).Value = (O_pen - C_lose)
      Range("K" & Counter).Value = (O_pen / C_lose) - 1
     
      Columns("K:K").Select
      Selection.Style = "Percent"
      Selection.NumberFormat = "0.00%"
           
      Range("L" & Counter).Value = Stockvol
      
      Counter = Counter + 1
      
      Stockvol = 0

    Else

      Stockvol = Stockvol + Cells(I, 7).Value

    End If

  Next I
  
  MsgBox Current.Name
  
  Next

End Sub