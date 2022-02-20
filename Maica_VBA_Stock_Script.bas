Attribute VB_Name = "Module1"
Sub Stock_analysis()


'Loop the entire workseet

Dim ws As Worksheet

For Each ws In ThisWorkbook.Sheets
 
WorksheetName = ws.Name

Debug.Print ws.Name
 
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Name the headers

ws.Range("J1").Value = " Yearly Change"

ws.Range("I1").Value = "Ticker"

ws.Range("K1").Value = "Percent Change"

ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"

ws.Range("Q1") = "Value"


'Set an initial variable for holding the ticker name


Dim Open_Price As Double

Open_Price = 0

Dim Close_Price As Double

Close_Price = 0

Dim Price_Change As Double

Price_Change = 0

Dim Price_Change_Percent As Double

Price_Change_Percent = 0

Dim Max_Ticker_Name As String

Max_Ticker_Name = " "

Dim Min_Ticker_Name As String

Min_Ticker_Name = " "

Dim Max_Percent As Double

Max_Percent = 0

Dim Min_Percent As Double

Min_Percent = 0

Dim Max_Volume_Ticker_Name As String

Max_Volume_Ticker_Name = " "

Dim Max_Volume As Double

Max_Volume = 0


Dim Ticker_Total As Double

Ticker_Total = 0

Dim Summary_Table_Row As Long

Summary_Table_Row = 2

Dim Ticker_Name As String

Ticker_Name = " "

Open_Price = ws.Cells(2, 3).Value



'Loop through all ticker

For i = 2 To LastRow

'Keep tracking of the location for each ticker in the summary table

    'Check if we are still within the same ticker, if we are not

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

         'Set the ticker name

         Ticker_Name = ws.Cells(i, 1).Value

        'Calculate Yearly Change and Percentage Change
         
         Close_Price = ws.Cells(i, 6).Value

         Price_Change = Close_Price - Open_Price

        'Check Division by 0 condition

        If Open_Price <> 0 Then

         Price_Change_Percent = (Price_Change / Open_Price) * 100
  
         End If

    'Add to the Ticker_Total

    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

    'Print the Ticker Name in the summary table

    ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

    'Print the Ticker change to the Summary Table

    ws.Range("J" & Summary_Table_Row).Value = Price_Change
    
    
    'Colour price change: Green is positive, Red is Negative
    
      If (Price_Change > 0) Then

        'Green is 4
    
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

         ElseIf (Price_Change <= 0) Then

        'Red is 3

        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        
         End If
    
    
    'Print the Price Change as a percent in the summary table

        ws.Range("K" & Summary_Table_Row).Value = (CStr(Price_Change_Percent) & "%")

    'Print the Ticker Amount to the Summary Table

        ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
        
    'Add one to the summary table row

         Summary_Table_Row = Summary_Table_Row + 1
         

        'Count the next open_price

        Open_Price = ws.Cells(i + 1, 3).Value
    

        'Get the max and min of Price_Change_Percent

          If (Price_Change_Percent > Max_Percent) Then

             Max_Percent = Price_Change_Percent

             Max_Ticker_Name = Ticker_Name


            ElseIf (Price_Change_Percent < Min_Percent) Then

             Min_Percent = Price_Change_Percent

             Min_Ticker_Name = Ticker_Name

            End If


            If Ticker_Total > Max_Volume Then

            Max_Volume = Ticker_Total

            Max_Volume_Ticker_Name = Ticker_Name
            
            

            End If
            
            'Reset values
            
            Price_Change_Percent = 0
            Ticker_Total = 0
            

'If the cells immediately follwoing a row is the same ticker

Else

Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value


End If


Next i

'Print values in assigned cells

ws.Range("Q2").Value = (CStr(Max_Percent) & "%")
ws.Range("Q3").Value = (CStr(Min_Percent) & "%")
ws.Range("P2").Value = Max_Ticker_Name
ws.Range("P3").Value = Min_Ticker_Name
ws.Range("Q4").Value = Max_Volume
ws.Range("P4").Value = Max_Volume_Ticker
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P4").Value = Max_Volume_Ticker_Name

Next ws



End Sub


