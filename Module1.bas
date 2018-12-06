Attribute VB_Name = "Module1"
Sub stock()

'This macro analyzes stock market data from 3 years (2014, 2015 and 2016)
'The workbook includes three sheets with information includes price, price changes and total volume.

'----------------
'I. Easy and Moderate assignment
'----------------
'A. Define variables and set inital values as needed
'----------------
    Dim Ticker_name As String
    Dim Ticker_Total As Double
    Ticker_Total = 0
    Dim Open_Price As Double
    Open_Price = 0
    Dim Close_Price As Double
    Dim Pr_Ch_Yr As Double
    Dim c_open As Double
    
    
    Dim lastRow As Double
        
    'Ratio of opening price from begining of year and end of year close price
     Dim Pt_PrCh_yr As Double
           
   ' Variable to assign the row in the summary section for each sheet
     Dim I_Summary_Tab_Row As Integer
      
  
'----------------
' B. For all worksheets in the workbook, use a loop to summarize information by <ticker>
'----------------
    ' BEGIN LOOP:WORKBOOK go through each worksheet in the workbook
    For Each ws In Worksheets
        
    ' Set the summary section row counter at the start of each sheet
      I_Summary_Tab_Row = 3
      
      ws.Range("M2").Value = ws.Range("C2")
      
        ' Find last row in this worksheet and store that row in lastRow
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                                       
      'BEGIN LOOP#1:EACH ROW IN SHEET go through each row in a single worksheet
            For k = 3 To lastRow
                         
              ' Check if we are on the last row of a particular ticker
                If ws.Cells(k + 1, 1).Value <> ws.Cells(k, 1).Value Then
              
              ' Grab the open price in the first row of the ticker on the tab and print to test
                Open_Price = ws.Range("C" & k + 1).Value
                ws.Range("M" & I_Summary_Tab_Row).Value = Open_Price
              ' Increment the Summary section Counter
                I_Summary_Tab_Row = I_Summary_Tab_Row + 1
              
                Else
              End If
            Next k
                
        'BEGIN LOOP#2:EACH ROW IN SHEET go through each row in a single worksheet
             
             ' Set the summary section row counter at the start of each sheet
               I_Summary_Tab_Row = 2
             
             For j = 2 To lastRow
                         
              ' Check if we are on the last row of a particular ticker
                If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
                
                '----------------
                ' B1. If YES - last row for ticker name then do below
                '----------------
                            ' Set the ticker name
                              Ticker_name = ws.Cells(j, 1).Value
                            ' Print the Ticker Name and Open Price in the Summary Area of Sheet
                              ws.Range("I" & I_Summary_Tab_Row).Value = Ticker_name
                            ' Capture opening price for calc
                              c_open = ws.Range("M" & I_Summary_Tab_Row).Value
                            ' Add last row for this ticker to the Ticker Total Volume
                              Ticker_Total = Ticker_Total + ws.Cells(j, 7).Value
                            ' Print the Ticker Total Volume in  Summary Area of Sheet
                              ws.Range("L" & I_Summary_Tab_Row).Value = Ticker_Total
                              
                            ' Grab the close_price for the ticker for the year and test print
                              Close_Price = ws.Cells(j, 6).Value
                              ws.Range("N" & I_Summary_Tab_Row).Value = Close_Price
                              
                            ' Calculate the difference between open and close for the year
                              Pr_Ch_Yr = ws.Range("M" & I_Summary_Tab_Row) - Close_Price
                            
                            ' Print the Yearly Change in the Summary Area of Sheet
                              ws.Range("J" & I_Summary_Tab_Row).Value = Pr_Ch_Yr
                                    If Pr_Ch_Yr > 0 Then
                                        ws.Range("J" & I_Summary_Tab_Row).Interior.ColorIndex = 4
                                    ElseIf Pr_Ch_Yr < 0 Then
                                        ws.Range("J" & I_Summary_Tab_Row).Interior.ColorIndex = 3
                                    End If
                                             
                            ' Print the calculated Percent Change in the Summary Area of Sheet if not zero
                                  If Close_Price > 0 Then
                                    Pt_PrCh_yr = c_open / Close_Price
                                    ws.Range("K" & I_Summary_Tab_Row).Value = Pt_PrCh_yr
                                    ws.Range("K" & I_Summary_Tab_Row).NumberFormat = "0.00%"
                                  End If
                               
                            ' Increment the Summary section Counter
                              I_Summary_Tab_Row = I_Summary_Tab_Row + 1
                            
                            ' Reset total for next ticker
                              Ticker_Total = 0
                              Close_Price = 0
                              Pr_PrCh_yr = 0
                    Else
                            ' Still looping through same ticker so adding to the Ticker Total Volume
                              Ticker_Total = Ticker_Total + ws.Cells(j, 7).Value

                    End If
                    
            Next j
            
 '----------------
 ' II. Evaluate each row in the new summary section of each tab (Loop)
 '----------------
                
            '----------------
            'A. Define variables needed
            Dim HPr_T_Name As String
            Dim H_Pr_Ch_Yr As Double
            '----------------
            'B. Define Variables needed
            Dim LPr_T_Name As String
            Dim L_Pr_Ch_Yr As Double
            '----------------
            'C. Define variables needed
            Dim H_TT_Name As String
            Dim H_Ticker_Total_Yr As Double
                                  
            HPr_T_Name = ws.Cells(2, 9).Value
            H_Pr_Ch_Yr = ws.Cells(2, 10).Value
            
            LPr_T_Name = ws.Cells(2, 9).Value
            L_Pr_Ch_Yr = ws.Cells(2, 10).Value
            
            H_TT_Name = ws.Cells(2, 9).Value
            H_Ticker_Total_Yr = ws.Cells(2, 12).Value
            
            '----------------
            'Begin loop "m", greatest price increase
            '----------------
            
            For m = 2 To I_Summary_Tab_Row - 1
                                                              
                 If H_Pr_Ch_Yr <= ws.Cells(m, 10).Value Then
                 
                 H_Pr_Ch_Yr = ws.Cells(m, 10).Value
                 HPr_T_Name = ws.Cells(m, 9).Value
                                                
                 
                 End If
               
            Next m
                                
            '----------------
            'End loop "m", Begin loop "n", greatest priced drop
            '----------------
            
            For n = 2 To I_Summary_Tab_Row - 1
                                                               
                 If L_Pr_Ch_Yr >= ws.Cells(n, 10).Value Then
                 
                 L_Pr_Ch_Yr = ws.Cells(n, 10).Value
                 LPr_T_Name = ws.Cells(n, 9).Value
                                                
                 
                 End If
               
            Next n
            
            '----------------
            'End loop "n", Begin loop "p" Max Volume
            '----------------
            
            For p = 2 To I_Summary_Tab_Row - 1
                               
                 If H_Ticker_Total_Yr <= ws.Cells(p, 12).Value Then
                 
                 H_Ticker_Total_Yr = ws.Cells(p, 12).Value
                 H_TT_Name = ws.Cells(p, 9).Value
                                                
                 
                 End If
               
            Next p
            
            '----------------
            'Print summary highlights
            '----------------
            'A. Higest Price Increase
            ws.Range("P2").Value = HPr_T_Name
            ws.Range("Q2").Value = H_Pr_Ch_Yr
            
            'B.Greatestgest Price Decrease
             ws.Range("P3").Value = LPr_T_Name
             ws.Range("Q3").Value = L_Pr_Ch_Yr
                       
            'C.Highest Volume
             ws.Range("P4").Value = H_TT_Name
             ws.Range("Q4").Value = H_Ticker_Total_Yr
             
       
        
'----------------
' III. Format each sheet and create headers "summary" section
'----------------
             ws.Range("I1").Value = "Ticker"
             ws.Range("J1").Value = "Yearly Change"
             ws.Range("K1").Value = "Percent Change"
             ws.Range("L1").Value = "Total Stock Volume"
             ws.Range("M1").Value = "Open_Price"
             ws.Range("N1").Value = "Close_Price"
             
    '----------------
    ' Create headers for the "highlights" section
    '----------------
    
             ws.Range("P1").Value = "Ticker"
             ws.Range("Q1").Value = "Value"
             
             ws.Cells(2, 15).Value = "Greatest % Increase"
             ws.Cells(3, 15).Value = "Greatest % Decrease"
             ws.Cells(4, 15).Value = "Greatest Total Volume"
             
             ' Autofit to display data
             ws.Columns("A:Q").AutoFit
        
    Next ws

End Sub
