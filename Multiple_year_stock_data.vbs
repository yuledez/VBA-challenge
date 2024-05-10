
Sub Main():
  ' Declare Variables
  Dim i As Long
  Dim next_empty_row As Integer
  Dim lead_ticker As String
  Dim next_ticker As String
  Dim ticker_Open As Double
  Dim ticker_Close As Double
  Dim first_flag As Boolean
  Dim total As Double ' Double allows higher numbers than Long
  
  ' Initialize Variables
  next_empty_row = 2
  lead_ticker = "none"
  ticker_Open = 0
  ticker_Close = 0
  first_flag = True
  total = 0
  
  ' Set New Column headers
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Quarterly Change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Stock Volume"
  
  ' Loop through rows - TODO: LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To 93001
    
        ' Iterate through the rows placing a value of 1 throughout
        next_ticker = Cells(i, 1).Value
        total = total + Cells(i, 7).Value
        
        If lead_ticker <> next_ticker Then
            ' It means this is the first row with new Ticker
            Cells(next_empty_row, 9).Value = next_ticker
            
            If first_flag = True Then
                first_flag = False
            Else
                ticker_Close = Cells(i - 1, 6).Value
                Cells(next_empty_row - 1, 10).Value = ticker_Close - ticker_Open
                Cells(next_empty_row - 1, 11).Value = (ticker_Close - ticker_Open) / ticker_Open
                ' falta forzar la columna anterior a type = %
                Cells(next_empty_row - 1, 12).Value = total - Cells(i, 7).Value
                total = Cells(i, 7).Value
            End If
            
            ticker_Open = Cells(i, 3).Value
            next_empty_row = next_empty_row + 1
            lead_ticker = next_ticker
        Else
            ' Is this the last row?
            If i = 93001 Then
                ticker_Close = Cells(i - 1, 6).Value
                Cells(next_empty_row - 1, 10).Value = ticker_Close - ticker_Open
                Cells(next_empty_row - 1, 11).Value = (ticker_Close - ticker_Open) / ticker_Open
                ' falta forzar la columna anterior a type = %
                Cells(next_empty_row - 1, 12).Value = total
            End If
            
        End If

    ' Call the next iteration
    Next i
  
End Sub



