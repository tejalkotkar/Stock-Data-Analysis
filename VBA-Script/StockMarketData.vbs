'--------------------------------------------------------------------------------
'                                   VBA CHALLENGE
'--------------------------------------------------------------------------------
'## Instructions
'
'* Create a script that will loop through all the stocks for one year and output the following information.
'
'  * The ticker symbol.
'
'  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'
'  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'
'  * The total stock volume of the stock.
'
'* You should also have conditional formatting that will highlight positive change in green and negative change in red.
'### CHALLENGES
'
'1. Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:
'
'2. Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.


Sub StockMarketData()

' Define variables
Dim i, LRow As Long
Dim Total As LongLong
Dim Start_Open, End_Close, yearly_change, percent_change As Double
Dim column, flag As Integer
Dim ticker_symbol As String

'Loop through each sheet

For Each ws In Worksheets

    'Define directory called dict where ticker symbol is the unique key
    Dim dict As Object
    Set dict = CreateObject("scripting.dictionary")

    ' Get the last row of the sheet
    LRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Set values of TotalVolume & flag as 0 & column as 1
    Total = 0
    flag = 0
    column = 1

    'Loop Through all rows to fine Yearly Change, Percent Change and Total Volume
    For i = 2 To LRow
    
        If flag = 0 Then
            Start_Open = ws.Cells(i, 3).Value  'Get the opening price at the start of the year for every ticker
            flag = 1
        End If

        'Searches for when value of next cell is differnt value of current cell
        If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
     
            'Get the ticker symbol from cell value and store it in variable
            ticker_symbol = ws.Cells(i, column).Value
        
            'Get the closing price at the end of the year
            End_Close = ws.Cells(i, 6).Value
        
            'Calculating Yearly change
            yearly_change = End_Close - Start_Open
        
            'calculating Percent Change
            If yearly_change = 0 Or Start_Open = 0 Then
                percent_change = 0
            Else
                percent_change = yearly_change / Start_Open
            End If
                
            'Calculating Total volume
            Total = Total + ws.Cells(i, 7).Value
            'Debug.Print Total
    
            'Add Yearly Change, Percent Change and Total to Dict
            dict.Add ticker_symbol, Array(yearly_change, percent_change, Total)
    
            'Reset Total & flag to 0
            Total = 0
            flag = 0
    
        'If current and next cell is same just add up Total in else block
        Else
            Total = Total + ws.Cells(i, 7).Value
            'Debug.Print Total
        End If
        
    Next i

    ' Create a header for summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("I1:L1").Interior.ColorIndex = 37
    ws.Range("I1:L1").Font.Bold = True


    ' Initialise row count to print dictory to worksheet
    Dim summary_row, Summary_column, s_column As Integer
    summary_row = 2
    Summary_column = 10


    ' Print the directory structure with all values in summary table
    ' loop through each key in dictonary  and print key & value to sheet

    Dim key As Variant
    For Each key In dict.Keys
        ws.Cells(summary_row, 9).Value = key
        For j = 0 To 2
            s_column = Summary_column + j
            ws.Cells(summary_row, s_column).Value = dict.Item(key)(j)
        Next j
        summary_row = summary_row + 1
    Next
    

    ' Format Percent_Change column to percentage with 2 decimal points
    ws.Range("K2:K" & summary_row).NumberFormat = "0.00%"

    'Autofit data
    ws.Columns("I:L").AutoFit

    'Perform Conditional formatting on yearly change column
    'Loop Through Summary Table and change color of cell according to value in cell

    For j = 2 To summary_row - 1
        If ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
        ElseIf ws.Cells(j, 10).Value >= 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
        End If
    Next j

'--------------------------------------------------------------------------------
                        'CHALLENGE I
'--------------------------------------------------------------------------------
    ' Find out greatest % Increase, decrease and total

    Dim max_total As LongPtr
    Dim max_increase, max_decrease As Double
    Dim total_ticker, increase_ticker, decrease_ticker As String

    ' Store initial values in the variables
    max_total = ws.Cells(2, 12).Value
    max_increase = ws.Cells(2, 11).Value
    max_decrease = ws.Cells(2, 11).Value
    total_ticker = ws.Cells(2, 9).Value
    increase_ticker = ws.Cells(2, 9).Value
    decrease_ticker = ws.Cells(2, 9).Value

    ' Loop through summary table to calculate max_increase, max_decrease & max_total
    For i = 2 To summary_row - 1
        If ws.Cells(i, 11).Value < max_decrease Then
            max_decrease = ws.Cells(i, 11).Value
            decrease_ticker = ws.Cells(i, 9).Value
        End If

        If ws.Cells(i, 11).Value > max_increase Then
            max_increase = ws.Cells(i, 11).Value
            increase_ticker = ws.Cells(i, 9).Value
        End If

        If ws.Cells(i, 12).Value > max_total Then
            max_total = ws.Cells(i, 12).Value
            total_ticker = ws.Cells(i, 9)
        End If
    Next i

    'Adding all calculated values to worksheet
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"

    ws.Range("O2").Value = increase_ticker
    ws.Range("P2").Value = max_increase
    ws.Range("O3").Value = decrease_ticker
    ws.Range("P3").Value = max_decrease
    ws.Range("O4").Value = total_ticker
    ws.Range("P4").Value = max_total

    'Format the values back to percentage
    ws.Range("P2:P3").NumberFormat = "0.00%"
    
    'Autofit data
    ws.Columns("N:P").AutoFit
    
Next ws
        
End Sub


