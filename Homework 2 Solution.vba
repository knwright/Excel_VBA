Sub alpha_testing()

    'For each worksheet in this workbook, establish the following variables and run subsequent loop.    
    For Each ws In Worksheets

        'Assign variables to hold the range for conditional formatting; it is important that these 
        'variables are set first to ensure the same formatting does not apply to the cells containing
        'the final ticker values for each ticker type.
        Dim rg As Range
        
        'Cond1 will hold the greater than or equal to zero values, 
        'cond2 will hold the less than zero values.
        Dim cond1 As FormatCondition, cond2 As FormatCondition

        'Assign Variables to hold remaining values for remaining calculations.
        Dim Total_Stock_Volume As Double
        
        Dim Ticker_Type As String
        
        Dim TickerCount As Integer
        TickerCount = 1
        
        Dim Yearly_Change As Double
        
        Dim Year_Open As Double
        'Assign the starting point for the year open stock value.
        Year_Open = ws.Cells(2, 3)
            
        Dim Year_Close As Double
        
        Dim Percent_Change As Double

        Dim Greatest_Percent_Increase As Double

        Dim Greatest_Percent_Decrease As Double

        Dim Greatest_Total_Volume As Double

        Dim Greatest_Incr_Tckr As String

        Dim Grt_Dcr_Tckr As String

        Dim Grt_TotVol_Tckr As String       

        'Establish column and row headers for where gathered data will be stored.
        ws.Range("I1").Value = "Ticker Type"
        
        ws.Range("J1").Value = "Yearly Change"
        
        ws.Range("K1").Value = "Percent Change"
        
        ws.Range("L1").Value = "Total Stock Volume"

        ws.Range("O2").Value = "Greatest % Increase"

        ws.Range("O3").Value = "Greatest % Decrease"

        ws.Range("O4").Value = "Greatest Total Volume"

        ws.Range("P1") = "Ticker"

        ws.Range("Q1") = "Value"
            
        'Determine the last row and column for each sheet.
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        'For each row starting in row 2 through the last row of data, run the following loop:
        For r = 2 To LastRow
            
                    'If the ticker type in the first column does not match the ticker type preceeding it then:
                    If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
                
                        'Establish what the new ticker name is.   
                        Ticker_Type = ws.Cells(r, 1).Value
                    
                        'Add a row and the new ticker type (name) to column 9.    
                        ws.Cells(TickerCount + 1, 9).Value = Ticker_Type
                    
                        'Establish the stock value at year close for that ticker type.
                        Year_Close = ws.Cells(r, 6).Value
                    
                        'Determine the yearly change by subtracting the stock open value from the stock closing value
                        'for that ticker type.
                        Yearly_Change = (Year_Close) - (Year_Open)
                    
                        'If the year close value equals the year open value then set percent change to zero.
                        If Year_Close = Year_Open Then

                            Percent_Change = 0
                        
                        'If the year open value is zero then set the percent change to 100%.   
                        ElseIf Year_Open = 0 Then

                            Percent_Change = 1
                        
                        'Otherwise, calculate the percent change by dividing the previously calculated yearly change by the
                        'year open value.
                        Else
                    
                            Percent_Change = (Yearly_Change) / (Year_Open)

                        End If
                    
                        'Calculate the total stock volume for each ticker type.
                        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(r, 7).Value
                        
                        'Place the yearly change for each ticker type in subsequent rows in column 10.
                        ws.Cells(TickerCount + 1, 10).Value = Yearly_Change
                    
                        'Place the percent change for each ticker type in subsequent rows in column 11.
                        ws.Cells(TickerCount + 1, 11).Value = Percent_Change

                        'If the percent change is greater than the greatest percent increase or the ticker count is one then
                        'assign the greastest percent increase and corresponding ticker type.
                        If Percent_Change > Greatest_Percent_Increase Or TickerCount = 1 Then

                            Greatest_Percent_Increase = Percent_Change

                            Greatest_Incr_Tckr = Ticker_Type

                        End If

                        'If percent change is less than the greatest percent decrease or the ticker count is one then
                        'assign the greatest percent decrease and corresponding ticker type.
                        If Percent_Change < Greatest_Percent_Decrease Or TickerCount = 1 Then

                            Greatest_Percent_Decrease = Percent_Change

                            Grt_Dcr_Tckr = Ticker_Type

                        End If

                        'If total stock volume is greater than the greatest total volume or the ticker count is one then
                        'assign the greatest total volume and corresponding ticker type.
                        If Total_Stock_Volume > Greatest_Total_Volume Or TickerCount = 1 Then

                            Greatest_Total_Volume = Total_Stock_Volume

                            Grt_TotVol_Tckr = Ticker_Type

                        End If

                        'Store the running tally of total stock volumes for each ticker type in column 12.
                        ws.Cells(TickerCount + 1, 12).Value = Total_Stock_Volume
                    
                        'Reset the total stock volume to zero for each new ticker type.
                        Total_Stock_Volume = 0

                        'Add a new row to the ticker count.
                        TickerCount = TickerCount + 1

                        'Assign the year open value for each new ticker type.
                        Year_Open = ws.Cells(r + 1, 3).Value
                    
                    Else

                        'If the ticker type is the same, continue to sum the stock volumes.
                        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(r, 7).Value
                    
                                    
                    End If
                
        Next

        'Assign where the greatest percent increase, decrease and greatest total volume values will go.
        ws.Range("P2") = Greatest_Incr_Tckr

        ws.Range("Q2") = Greatest_Percent_Increase

        ws.Range("P3") = Grt_Dcr_Tckr

        ws.Range("Q3") = Greatest_Percent_Decrease

        ws.Range("P4") = Grt_TotVol_Tckr

        ws.Range("Q4") = Greatest_Total_Volume
                       
        'Apply conditional formatting to the yearly change column so that positive values are green
        'and negative values are red, only apply this formatting to the end of the data. 
        Set rg = ws.Range("K2", ws.Range("K2").End(xlDown))
            
        Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreaterEqual, "0")
            
        Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "0")
            
        With cond1
            
            .Interior.Color = vbGreen
            
        End With
            
        With cond2
            
            .Interior.Color = vbRed
                
        End With
        
    Next

End Sub
