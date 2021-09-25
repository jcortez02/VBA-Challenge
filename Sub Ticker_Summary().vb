Sub Ticker_Summary()
    'Create summary table symbol header
    Cells(1, 9).Value = "Ticker"
    
    'Create summary table Yearly Change in Price header
    Cells(1, 10).Value = "Yearly Change"
    
    'Create summary table Yearly Precent Change in Price header
    Cells(1, 11).Value = "Precent Change"
    
    'Create summary table total volume header
    Cells(1, 12).Value = "Total Volume"
    
    'Define the range of column1
    Dim LR As Long
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Define a variable for symbol ticker
    Dim Symbol_Ticker_Name As String
    Symbol_Ticker_Name = Cells(2, 1)
    
    'Define a variable for symbol ticker total per symbol
    Dim Symbol_Ticker_Total As Double
    Symbol_Ticker_Total = 0
    
    'Summary table for running totals of symbol tickers
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Define a variable for symbol ticker opening date value
    Dim Opening_Date_Value As Double

    'Define a variable for symbol ticker closing date value
    Dim Closing_Date_Value As Double

    'Subtract Open Date Value - Closing Date Value
    Dim Value_Change_Total As Double

    'Find Change Total Value Percentage
    Dim Value_Change_Total_Percentage As Double

    'Find the opening date value
    Opening_Date_Value = Cells(2, 3).Value

    'Loop through all ticker symbol names
    For i = 2 To LR
    
    'Ticker Open and Close
    
    'Check to see if the ticker symbol is the same
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        'Add to the Symbol Ticker Total
        Symbol_Ticker_Total = Symbol_Ticker_Total + Cells(i, 7).Value
    
         'Find the closing date value
         Closing_Date_Value = Cells(i, 6).Value

         'Find Change Total Value
         Value_Change_Total = Opening_Date_Value - Closing_Date_Value

         'Find Change Total Value Percentage
         Value_Change_Total_Percentage = Opening_Date_Value/Closing_Date_Value

         'Change Format in Column 11
         Cells(i,11).NumberFormat = "0.00%"

         'Print the Symbol Ticker Change in Summary Table
         Range("I" & Summary_Table_Row).Value = Symbol_Ticker_Name

         'Print the Value Change in Summary Table
         Range("J" & Summary_Table_Row).Value = Value_Change_Total

         'Print the Symbol Ticker Change % in Summary Table
         Range("K" & Summary_Table_Row).Value = Value_Change_Total_Percentage

         'Print the Symbol Ticker Total in Summary Table
         Range("L" & Summary_Table_Row).Value = Symbol_Ticker_Total

         'Set the Symbol Ticker Name
         Symbol_Ticker_Name = Cells(i + 1, 1).Value

         'Add one row to the summary table row
         Summary_Table_Row = Summary_Table_Row + 1
    
         'Reset the Brand Total
         Symbol_Ticker_Total = 0

         'Rest the opening value of ticker
         Opening_Date_Value = Cells(i + 1, 3)
   
    
    'If the cell immediatley following a row is the same brand....
    Else
    
        'Add to the Symbol Ticker Total
        Symbol_Ticker_Total = Symbol_Ticker_Total + Cells(i, 7).Value
            
    End If
    
    Next i
    
    End Sub
    
