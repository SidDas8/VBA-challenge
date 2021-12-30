Attribute VB_Name = "Module1"
Sub FinalVBA():

    Dim ws As Worksheet
    
    For Each ws In Worksheets

        ' Create variable to find the last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Create variable to hold ticker
        Dim ticker As String
        
        ' Create variable to hold yearly change
        Dim yearly_change As Double
        yearly_change = 0
        
        ' Create variable to hold opening price
        Dim opening_price As Double
        opening_price = 0
        
        ' Create variable to hold closing price
        Dim closing_price As Double
        closing_price = 0
        
        ' Create variable to hold percent change
        Dim percent_change As Double
        percent_change = 0
        
        ' Create variable to hold total stock volume
        Dim total_stock_volume As Double
        total_stock_volume = 0
        
        ' Create variable to track tablerow
        Dim tablerow As Double
        tablerow = 2
        
        ' Create headings for report
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        
        
' Create a script that will loop through all the stocks for one year and output the following information.
        For i = 2 To lastrow
            
                ' If opening price is zero (from previous row)
                If opening_price = 0 Then
                
                    ' Set opening price from current row
                    opening_price = ws.Cells(i, 3).Value
                
                End If
                
                ' Add stock volume
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            ' If the next cell is different from the current cell...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Print the ticker symbol to the table
                ticker = ws.Cells(i, 1).Value
                ws.Range("I" & tablerow).Value = ticker
                
                
                
' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
                
                closing_price = ws.Cells(i, 6).Value
                yearly_change = closing_price - opening_price
                ws.Range("J" & tablerow).Value = yearly_change
                
                
                
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
                
                ' If final opening price = 0
                If opening_price = 0 Then

                    ' Set percent change to zero because we cannot divide by zero
                    percent_change = 0
                    
                Else
                    
                    ' Calculate percent change when opening price is not zero
                    ' Round to 2 decimals for percentage
                    percent_change = Round((yearly_change / opening_price) * 100, 2)
                
                End If
                
                ' Format column with "%"
                ws.Range("K" & tablerow).Value = "%" & percent_change
                
                
                
' You should also have conditional formatting that will highlight positive change in green and negative change in red.
                If yearly_change < 0 Then
                
                    ' If percent change is negative, highlight red
                    ws.Range("J" & tablerow).Interior.ColorIndex = 3
                
                ElseIf yearly_change > 0 Then
                
                    ' If percent change is positive, highlight green
                    ws.Range("J" & tablerow).Interior.ColorIndex = 4
                    
                Else:
                    ' Highlight blank
                    ws.Range("J" & tablerow).Interior.ColorIndex = 0
                
                End If
                
                ' Post total stock volume
                ws.Range("L" & tablerow).Value = total_stock_volume
                
                
                
' For the next ticker...
                
                ' Add one to the summary table row
                tablerow = tablerow + 1
                
                ' Set the opening price
                opening_price = 0
                
                ' Reset the total_stock_volume
                total_stock_volume = 0
        
            End If
        
        Next i
        
        
        
        ' Bonus
        
        ' Create headings for bonus
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Create variable to find the last row of posted table
        smallrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Create variable for greatest increase
        Dim increase As Double
        increase = ws.Range("K2").Value
        
        ' Create variable for greatest decrease
        Dim decrease As Double
        decrease = ws.Range("K2").Value
        
        ' Create variable for greatest total volume stock
        Dim total As Double
        total = ws.Range("L2").Value
        
        ' Insert ticker into P2:P4
        ws.Range("P2:P4").Value = ws.Range("I2").Value
        
        ' Loop through posted table percent change column
        For n = 3 To smallrow
        
            ' If current cell is greater than greatest increase
            If ws.Cells(n, 11).Value > increase Then
            
                ' Place current cell value into greatest increase variable
                increase = ws.Cells(n, 11).Value
                
                ' Post greatest increase and associated ticker
                ws.Range("Q2").Value = "%" & increase * 100
                
                ws.Range("P2").Value = ws.Cells(n, 9).Value
                
            ' If current cell is less than greatest decrease
            ElseIf ws.Cells(n, 11).Value < decrease Then
            
                ' Place current cell value into greatest decrease variable
                decrease = ws.Cells(n, 11).Value
                
                ' Post greatest decrease and associated ticker
                ws.Range("Q3").Value = "%" & decrease * 100
                
                ws.Range("P3").Value = ws.Cells(n, 9).Value
                
            End If
            
            ' If current cell is greater than greatest total
            If ws.Cells(n, 12).Value > total Then
            
                ' Place current cell value into greatest total variable
                total = ws.Cells(n, 12).Value
                
                ' Post greatest total and associated ticker
                ws.Range("Q4").Value = total
                
                ws.Range("P4").Value = ws.Cells(n, 9).Value
            
            End If
        
        Next n
    
        ' Add autofit line
        ws.Columns("A:Q").AutoFit
        
    Next ws
    
End Sub


