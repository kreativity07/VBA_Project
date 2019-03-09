# VBA_Project

Sub WorksheetLoop()

         Dim WS_Count As Integer
         Dim i As Integer

         ' Set WS_Count equal to number of worksheets in the active
         ' workbook.
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         For i = 1 To WS_Count

            ' This shows how to reference a sheet within the loop and displays worksheet name in dialog box.
            MsgBox ActiveWorkbook.Worksheets(i).Name

         Next i

      End Sub
      
Sub StockAnalysis()

  Dim ws As Worksheet
  For Each ws In Worksheets
  ws.Activate
  Dim Ticker As String

  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0

  ' Track the Location of each row/line in the Summary Table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  ws.Range("i1").Value = "Ticker"
  ws.Range("j1").Value = "Total Stock Volume"
  
  ' Define the Last Row
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  For i = 2 To LastRow

    ' Check to see if we are within the same value
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker Value
      Ticker = Cells(i, 1).Value

      ' Add to Total_Stock_Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

      ' Print Ticker in Summary Table
      ws.Range("i" & Summary_Table_Row).Value = Ticker

      ' Print Total_Stock_Volume to Summary Table
      ws.Range("j" & Summary_Table_Row).Value = Total_Stock_Volume

      ' Add 1 to Summary Table Row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset Total_Stock_Volume
      Total_Stock_Volume = 0

    Else

      ' Add to Total Stock Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    End If
      
  Next i

 Next ws

End Sub


