Attribute VB_Name = "Module1"
Sub StockAnalyzer():

'Loop through the tabs
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets


'Add headers for both analysis tables and adjust the columns
ws.Range("G1").EntireColumn.NumberFormat = "0"
ws.Range("L1").EntireColumn.NumberFormat = "0"
ws.Range("I1,P1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Defining Variables
Dim Ticker_name As String

Dim Ticker_tally As LongLong
Ticker_tally = 0

LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

Dim YrChg As Double
Dim PercChg As Double

Dim FirstOpen As Double
Dim LastClose As Double

FirstOpen = 0

'track the row of tickers in table 1
Dim Table1_row As Integer
Table1_row = 2

'Loop the stock data
For t = 2 To LastRow

    'check and summarize tickers
    If ws.Cells(t + 1, 1).Value <> ws.Cells(t, 1).Value Then
    
    Ticker_name = ws.Cells(t, 1).Value
    
    Ticker_tally = Ticker_tally + ws.Cells(t, 7).Value
    
    'Obtain last closing amount for the ticker
    LastClose = ws.Cells(t, 6).Value

    'update the 1st table with ticker names and amounts
    ws.Range("I" & Table1_row).Value = Ticker_name
    ws.Range("L" & Table1_row).Value = Ticker_tally
    
    
    'update 1st Table with yearly change and percent change amounts
    ws.Range("J" & Table1_row).Value = LastClose - FirstOpen
    ws.Range("K" & Table1_row).Value = (LastClose - FirstOpen) / FirstOpen
       
    'add a row for the next ticker
    Table1_row = Table1_row + 1
    
    'reset ticker count for next
    Ticker_tally = 0
    FirstOpen = 0
    
    Else
    
    'Add the row to the total stock volume for that ticker
    Ticker_tally = Ticker_tally + ws.Cells(t, 7).Value
    
        'obtain first open amount per ticker
        If FirstOpen = 0 Then
        FirstOpen = ws.Cells(t, 3).Value
        
        Else
        
        End If
           
    
    End If
    
Next t
    
    'Updating formats
    'Add conditional formatting to yearly change
    For f = 2 To LastRow
    
    If ws.Range("J" & f).Value > 0 Then
    ws.Range("J" & f).Interior.ColorIndex = 4
    
    ElseIf ws.Range("J" & f).Value < 0 Then
    ws.Range("J" & f).Interior.ColorIndex = 3
    
    Else
    
    End If
    
    Next f
    
    'Update percent change column format
    ws.Range("K1").EntireColumn.NumberFormat = "0.00%"
    ws.Range("Q2", "Q3").NumberFormat = "0.00%"

'Table 2 added functionality
Dim LastRow2 As Double
Dim Vol_Count As LongLong
Dim MaxPerc_Count As Double
Dim LowPerc_Count As Double

Dim MaxPerc As Double
Dim LowPerc As Double
Dim MaxVol As LongLong

Dim MaxPerc_Row As Double
Dim LowPerc_Row As Double
Dim MaxVol_Row As LongLong

'determine last row
LastRow2 = ws.Cells(Rows.Count, "I").End(xlUp).Row

'Starting rows
MaxVol = ws.Cells(2, "L").Value
MaxPerc = ws.Cells(2, "K").Value
LowPerc = ws.Cells(2, "K").Value


    'Loop to determine the high and low values
    For m = 2 To LastRow2
    Vol_Count = ws.Cells(m, "L").Value
    MaxPerc_Count = ws.Cells(m, "K").Value
    LowPerc_Count = ws.Cells(m, "K").Value

        'tally up max volumes
        If Vol_Count >= MaxVol Then
        MaxVol = Vol_Count
        MaxVol_Row = m
                
        'populate table 2 with the outputs
        ws.Range("Q4").Value = Vol_Count
        ws.Range("P4").Value = ws.Cells(MaxVol_Row, "I").Value
        
        Else
        End If
    
            'find the greatest % increase
            If MaxPerc_Count >= MaxPerc Then
            MaxPerc = MaxPerc_Count
            MaxPerc_Row = m
                
            'populate table 2 with the outputs
            ws.Range("Q2").Value = MaxPerc_Count
            ws.Range("P2").Value = ws.Cells(MaxPerc_Row, "I").Value
                
            Else
            End If
    
                'find the greatest % decrease
                If LowPerc_Count <= LowPerc Then
                LowPerc = LowPerc_Count
                LowPerc_Row = m
                
                'populate table 2 with the outputs
                ws.Range("Q3").Value = LowPerc_Count
                ws.Range("P3").Value = ws.Cells(LowPerc_Row, "I").Value
        
                Else
                End If
    
Next m

'align the columns
ws.Range("I1:Q1").EntireColumn.AutoFit

ws.Range("Q4").NumberFormat = "0"
  
'finish tab loops
Next ws



End Sub

