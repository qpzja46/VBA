Attribute VB_Name = "Module1"
Sub asd()
'Loop through all Sheets
 For Each ws In Worksheets

    'Add the word Ticker to Column 9 and 16 Header
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 16).Value = "Ticker"
    
    'Add the word Yearly Change to Column 10 Header
    ws.Cells(1, 10).Value = "Yearly Change"
    
    'Add the word Percentage Change to Column 11 Header
    ws.Cells(1, 11).Value = "Percentage Change"
    
    'Add the word Total Stock Volume to Column 12 Header
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Add the word Value to Column 17 Header
    ws.Cells(1, 17).Value = "Value"
    
    'Add the word greatest % increase, decrease, and total volume to Column 15s
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'Declare variables calculate yearly, percentage, and total volume by ticker
    Dim i As Long
    Dim Ticker As String
    Dim Total_Volume As Double
    Total_Volume = 0
    Dim Yearly_Change As Double
    Yearly_Change = 0
    Dim Next_Ticker_Open_Price As Double
    Next_Ticker_Open_Price = ws.Cells(2, 3).Value
    Dim Current_Ticker_Open_Price As Double
    Dim Percentage_Change As Double
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    Dim Last_Row As Long
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Add loop
For i = 2 To Last_Row

'Add conditional

    'Check if we are still within the same ticker, if it is not...
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'Set the ticker
     Ticker = ws.Cells(i, 1).Value
    
    'Print the ticker in the summary table
     ws.Range("i" & Summary_Table_Row).Value = Ticker
    
    'Find out yearly change in stock price using opening and closing stock price
     Current_Ticker_Open_Price = Next_Ticker_Open_Price
     Current_Ticker_Close_Price = ws.Cells(i, 6).Value
     Yearly_Change = Current_Ticker_Close_Price - Current_Ticker_Open_Price
     
    'Print the yealy change in the sumamry table
     ws.Range("j" & Summary_Table_Row).Value = Yearly_Change
     
    
    'Add to the total volume
     Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
    'Print the total volume to the summary table
     ws.Range("L" & Summary_Table_Row).Value = Total_Volume
     
        'Calculate Yearly Change Percentage
        If Current_Ticker_Open_Price = 0 Then
        ws.Cells(Summary_Table_Row, 10) = "NA"
     
        Else
     
        'Find out percentage change
        Percentage_Change = Yearly_Change / Current_Ticker_Open_Price
        ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
        ws.Range("K" & Summary_Table_Row).Style = "Percent"
    
        End If
    
    'Reset the total volume
     Summary_Table_Row = Summary_Table_Row + 1
     Total_Volume = 0
     Yearly_Change = 0
    
    'Gra
    Next_Ticker_Open_Price = ws.Cells(i + 1, 3).Value
    
    'If the cell immdiately following a row is the same ticker...
    Else
    
    'Add the the total volume
    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
   End If
   
 Next i
 
    'Declare variables for cell formatting
    Dim yearLastRow As Long
    yearLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    'Add Loop for cell formatting
    For i = 2 To yearLastRow
    
    'Add conditional for cell formatting
    If ws.Cells(i, 10).Value >= 0 Then
       ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
       ws.Cells(i, 10).Interior.ColorIndex = 3
       
    End If
   
   Next i
   
    'Declare variable for finding maximum and minimum
    Dim Percentage_Last_Row As Long
    Percentage_Last_Row = ws.Cells(Rows.Count, 11).End(xlUp).Row
    Dim Percentage_Max As Double
    Percentage_Max = 0
    Dim Percentage_Min As Double
    Percentage_Min = 0
    
    'Add Loop for finding maximum and minumum
    For i = 2 To Percentage_Last_Row
    
    'Add conditionals for maximum and minimum
    If Percentage_Max < ws.Cells(i, 11).Value Then
       Percentage_Max = ws.Cells(i, 11).Value
       ws.Cells(2, 17).Value = Percentage_Max
       ws.Cells(2, 17).Value = Format(Percentage_Max, "0.00%")
       ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
       
    End If
    
    If Percentage_Min > ws.Cells(i, 11).Value Then
       Percentage_Min = ws.Cells(i, 11).Value
           ws.Cells(3, 17).Value = Percentage_Min
           ws.Cells(3, 17).Value = Format(Percentage_Min, "0.00%")
           ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
           
    End If
   
      'Declare variable for greatest total volume
    Dim Total_Volume_Row As Long
    Total_Volume_Row = ws.Cells(Rows.Count, 12).End(xlUp).Row
    Dim Total_Volume_Max As Double
    Total_Volume_Max = 0
    
    'Add conditional for greatest total volume
    If Total_Volume_Max < ws.Cells(i, 12).Value Then
       Total_Volume_Max = ws.Cells(i, 12).Value
       ws.Cells(4, 17).Value = Total_Volume_Max
       ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    
    Else
    Total_Volume_Max = Total_Volume_Max + ws.Cells(i, 7).Value
    
    End If
   
    Next i
    
 Next ws
    
MsgBox ("Complete")

End Sub
