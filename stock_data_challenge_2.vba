Sub alphabeticaltesting()
 
    ' Declare variables
    Dim ws As Worksheet
    Dim tickername As String
    Dim totalvolume As Double
    Dim summary_table_row As Integer
    Dim lastrow As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim tickervolume As Double
    Dim i As Long, j As Long
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As Double
    Dim lastrow_summary As Long
 
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Reset variables for each worksheet
        totalvolume = 0
        summary_table_row = 2
 
        ' Set headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
 
        ' Find the last row of data in the current worksheet
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
 
        ' Set initial open price (assumes first row is header and data starts from row 2)
        open_price = ws.Cells(2, 3).Value
 
        ' Main loop: Iterate through all rows of data
        For i = 2 To lastrow
            ' Check if we're still within the same ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' We've reached the last row for the current ticker
 
                ' Get the ticker name
                tickername = ws.Cells(i, 1).Value
 
                ' Add the last volume to the total volume
                tickervolume = tickervolume + ws.Cells(i, 7).Value
 
                ' Output ticker name and total volume to summary table
                ws.Range("I" & summary_table_row).Value = tickername
                ws.Range("L" & summary_table_row).Value = tickervolume
 
                ' Get the closing price
                close_price = ws.Cells(i, 6).Value
 
                ' Calculate and output yearly change
                yearly_change = (close_price - open_price)
                ws.Range("J" & summary_table_row).Value = yearly_change
 
                ' Calculate percent change, avoiding division by zero
                If open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / open_price
                End If
 
                ' Output percent change and format as percentage
                ws.Range("K" & summary_table_row).Value = percent_change
                ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
 
                ' Move to next row in summary table
                summary_table_row = summary_table_row + 1
 
                ' Reset volume for next ticker
                tickervolume = 0
 
                ' Set open price for next ticker
                open_price = ws.Cells(i + 1, 3)
            Else
                ' Still within the same ticker, accumulate volume
                tickervolume = tickervolume + ws.Cells(i, 7).Value
            End If
        Next i
 
        ' Bonus calculations: Find greatest increase, decrease, and volume
        lastrow_summary = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
 
        ' Initialize with values from the first data row
        greatest_increase = ws.Cells(2, 11).Value
        greatest_decrease = ws.Cells(2, 11).Value
        greatest_volume = ws.Cells(2, 12).Value
 
        ' Loop through summary table to find greatest values
        For j = 2 To lastrow_summary
            ' Check for greatest percentage increase
            If ws.Cells(j, 11).Value > greatest_increase Then
                greatest_increase = ws.Cells(j, 11).Value
                ws.Cells(2, 18).Value = greatest_increase
                ws.Cells(2, 18).NumberFormat = "0.00%"
                ws.Cells(2, 17).Value = ws.Cells(j, 9).Value
            End If
 
            ' Check for greatest percentage decrease
            If ws.Cells(j, 11).Value < greatest_decrease Then
                greatest_decrease = ws.Cells(j, 11).Value
                ws.Cells(3, 18).Value = greatest_decrease
                ws.Cells(3, 18).NumberFormat = "0.00%"
                ws.Cells(3, 17).Value = ws.Cells(j, 9).Value
            End If
 
            ' Check for greatest total volume
            If ws.Cells(j, 12).Value > greatest_volume Then
                greatest_volume = ws.Cells(j, 12).Value
                ws.Cells(4, 18).Value = greatest_volume
                ws.Cells(4, 17).Value = ws.Cells(j, 9).Value
            End If
        Next j
 
        ' Apply conditional formatting to yearly change column
        For i = 2 To lastrow_summary
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
 
        ' Autofit columns for better readability
        ws.Columns("I:R").AutoFit
    Next ws
 
    ' Display completion message
    MsgBox "Analysis complete for all worksheets!", vbInformation
End Sub
