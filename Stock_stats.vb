'The Stock Stats program is intended to be run as a macro within an Excel worksheet. The program
'takes an Excel worksheet of stock ticker data as input and outputs ticker stats in a formatted table.
'The table displays each ticker symbol, its yearly change, its percent yearly change, and its total
'stock volume for the year. Positive change is highlighted in green and negative change in red.
'The spreadsheet contains the following fields: ticker symbol, date, opening price, closing price,
'and stock volume. The spreadsheet is assumed to be sorted first by ticker symbol, then sorted
'ascending by date so that the data is arranged in 'ticker blocks'.

Sub Stock_stats()

    Dim Previous_ticker As String
    Dim Current_ticker As String
    Dim Current_stock_volume As Long
    Dim Total_stock_volume As Double
    Dim Opening_price As Double
    Dim Closing_price As Double
    Dim Yearly_change As Double
    Dim Percent_yearly_change As Double
    Dim i As Long
    Dim Results_row As Integer
    Dim Last_row As Long
    
    'Loops through all worksheets
    For Each worksheet In Worksheets

        'Computes and initiates variables needed before for loop starts.
        Row_after_last_data_row = worksheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
        Results_row = 2
        Total_stock_volume = 0
        Opening_price = worksheet.Cells(2, 3).Value

        'Loops through all records, beginning with row 3 in order to operate on current row and previous row.
        For i = 3 To Row_after_last_data_row
            'Resets current ticker, previous ticker, and current stock
            Current_ticker = worksheet.Cells(i, 1).Value
            Previous_ticker = worksheet.Cells(i-1, 1).Value
            Current_stock_volume = worksheet.Cells(i-1, 7)
            
            'Evaluates whether or not the end of a ticker block has been reached by comparing
            'current ticker value with previous ticker value. Until the end of a ticker block is
            'reached, adds stock volume to running total. Once the end of a ticker block is reached, computes
            'ticker statistics and prints to results table.
            If Current_ticker = Previous_ticker Then
                'Adds current stock volume to running total.
                Total_stock_volume = Total_stock_volume + Current_stock_volume
            Else 
                'Once the end of a ticker block is reached, stores the closing price of the last
                'ticker block to a variable and computes Yearly Change and % Yearly Change.
                Closing_price = worksheet.Cells(i-1, 6)
                Yearly_change = Closing_price - Opening_price
                'Check for zeros to avoid divide by zero error.
                If Opening_price = 0 Then
                    Percent_yearly_change = 0
                Else
                    Percent_yearly_change = Round((Yearly_change / Opening_price * 100), 2)
                End if
                'Prints results to table in Excel worksheet.
                worksheet.Cells(1, 9).Value = "Ticker"
                worksheet.Cells(1, 10).Value = "Yearly Change"
                worksheet.Cells(1, 11).Value = "Percent Change"
                worksheet.Cells(1, 12).Value = "Total Stock Volume"
                worksheet.Cells(Results_row, 9).Value = Previous_ticker
                worksheet.Cells(Results_row, 10).Value = Yearly_change
                worksheet.Cells(Results_row, 11).Value = Percent_yearly_change
                worksheet.Cells(Results_row, 12).Value = Total_stock_volume
                
                'Applies red and green highlights for gains and losses.
                If Yearly_change > 0 Then
                    worksheet.Cells(Results_row, 10).Interior.ColorIndex = 4
                Elseif Yearly_change < 0 Then
                    worksheet.Cells(Results_row, 10).Interior.ColorIndex = 3
                End if
                
                'Resets stock volume accummulator variable and stores the opening price of the next ticker block.
                Total_stock_volume = 0
                Opening_price = worksheet.Cells(i, 3).Value
                Results_row = Results_row + 1
            End if
        Next i 'Ends loop within a worksheet.
    Next worksheet 'Ends loop through all worksheets.
End Sub