Attribute VB_Name = "Module1"
'Calculate total stock volume
Sub TotalStockVolume():

'loop through worksheets
 For Each ws In Worksheets

'variables
 Dim i, j As Long
 Dim total As Double
 Dim ticker As String
 Dim lastrow As Long
 Dim ychange As Double
 Dim pchange As Double
 Dim oprice As Double
 Dim cprice As Double
 Dim oprice_row As Long

 'header
 ws.Range("H1").Value = "Ticker"
 ws.Range("I1").Value = "Yearly Change"
 ws.Range("J1").Value = "Percent Change"
 ws.Range("K1").Value = "Total Stock Value"

 total = 0
 j = 2
 oprice_row = 2

 'find last row
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

 'Loop Through Each Year of Stock Data
 For i = 2 To lastrow
     
     'Compare Each Ticker
     If ws.Range("A" & i + 1).Value = ws.Range("A" & i).Value Then

         'Calculate Total Volume for Each Ticker If Tickers are same
         total = total + ws.Range("G" & i).Value

     Else
         'Grab Ticker when it change
         ticker = ws.Range("A" & i).Value

         'Calculate Yearly Change and Percent Change
         oprice = ws.Range("C" & oprice_row)
         cprice = ws.Range("F" & i)
         ychange = cprice - oprice

         'Calculate Percent Change
         If oprice = 0 Then
            pchange = 0
         Else
            pchng = ychng / oprice
         End If

         'Insert Grabbed Ticker,Total Volume,Yearly Change and Percent Change into Display Cells
         ws.Range("I" & j).Value = ticker
         ws.Range("L" & j).Value = total + ws.Range("G" & i).Value
         ws.Range("J" & j).Value = ychange
         ws.Range("K" & j).Value = pchng
         ws.Range("K" & j).NumberFormat = "0.00%"
         
         'Conditional Formating Yearly Change, Positive Green/ Negative Red
         If ws.Range("J" & j).Value > 0 Then
            ws.Range("J" & j).Interior.ColorIndex = 4
         Else
            ws.Range("J" & j).Interior.ColorIndex = 3
         End If

         'Add a New Row itno Display Cells for Next Ticker, Set New open rice row and Reset Total
         j = j + 1
         total = 0
         oprice_row = i + 1
         
     End If
 Next i
 Next ws
End Sub
