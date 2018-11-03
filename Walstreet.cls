VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub WallStreet_PtI()

   For Each ws In Worksheets

       ' Created a Variable to Hold File Name & Last Row
       Dim WorksheetName As String

       ' Determine the Last Row
       LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

       ' Grabbed the WorksheetName
       WorksheetName = ws.Name
       MsgBox WorksheetName

       ' Add headers to table
       ws.Cells(1, 10).Value = "Ticker Symbol"
       ws.Cells(1, 11).Value = "Total Volume"
       ws.Cells(1, 12).Value = "Open Price"
       ws.Cells(1, 13).Value = "End Price"
       ws.Cells(1, 14).Value = "Change in Price"
       ws.Cells(1, 15).Value = "Percent Change"

       'Set an initial variable for holding the ticker symbol
       Dim Ticker_Name As String

       'Set an initial variable for holding the total volume per ticker symbol
       Dim Ticker_Total As Double
       Ticker_Total = 0
       
       'Set an initial variable for holding the opening, closing, difference & % change
       Dim Start_Price As Double
       Dim End_Price As Double
       Dim Diff_Price As Double
       Dim Percent_Change As Double

       'Keep track of the location for each ticker symbol in the summary table
       Dim Summary_Table_Row As Integer
       Summary_Table_Row = 2

 'Loop through all stock transactions
 For i = 2 To LastRow

 If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
     Start_Price = ws.Cells(i, 3).Value
 End If
   
  'Check if we are still within the same ticker symbol, if it is not...
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

  ' Set the Ticker Symbol & End Price
  Ticker_Name = ws.Cells(i, 1).Value
  End_Price = ws.Cells(i, 6).Value

  'Add to the Ticker Volume Total
 Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

  ' Print the Ticker Symbol, Open Price, End Price, Differen in Prices, and % Change in the Summary Table
 ws.Range("J" & Summary_Table_Row).Value = Ticker_Name
 ws.Range("M" & Summary_Table_Row).Value = End_Price
 ws.Range("L" & Summary_Table_Row).Value = Start_Price
 ws.Range("N" & Summary_Table_Row).Value = End_Price - Start_Price
 If Start_Price <> 0 Then
 ws.Range("O" & Summary_Table_Row).Value = CDbl((End_Price - Start_Price) / Start_Price)
 Else: ws.Range("L" & Summary_Table_Row).Value = "N/A"
 End If
 ws.Range("O" & Summary_Table_Row).NumberFormat = "0.00%"

 ' Print the Ticker Symbol's total volume to the Summary Table
 ws.Range("K" & Summary_Table_Row).Value = Ticker_Total

 ' Add one to the summary table row
  Summary_Table_Row = Summary_Table_Row + 1

  'Reset the Ticker Total
   Ticker_Total = 0

   'If the cell immediately following a row is the same ticker symbol...
   Else

     'Add to the Ticker Total
     Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

   End If

 Next i

   Next ws

   MsgBox ("Fixes Complete")


End Sub


