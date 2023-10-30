Attribute VB_Name = "Module1"
Sub module_challenge()

For Each ws In Worksheets

 Worksheets(ws.Name).Activate
 
Dim Ticker_Name As String

Dim select_row As Double

Dim Percent_Change As Double


Dim TSV As Double
TSV = 0

  ' I keep track of the location for each ticker in Ticker column
  
Dim price_row As Integer

price_row = 2


Dim opening As Double
Dim closing As Double


   'I assign the headings

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

   'I find Last row

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

   'I Loop through all tickers code

For i = 2 To LastRow

   'Checking If I am still within the same Ticker's code
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

       Ticker_Name = ws.Cells(i, 1).Value

   'I set the total volume
       TSV = TSV + ws.Cells(i, 7).Value
       
    'I print the tickers code in Ticker Column
       ws.Range("I" & price_row).Value = Ticker_Name
       
    'I print the sum of volume in the total volume column
       ws.Range("L" & price_row).Value = TSV
    
    ' I add one to the ticker and total volume rows
       price_row = price_row + 1
       
    ' I reset the total volume
      TSV = 0
   
   Else

  TSV = TSV + ws.Cells(i, 7).Value


  End If

 Next i
 
   'I keep track of the location for each value in column 10 and 11
 
select_row = 2
 
 
   'I Loop through all opening and closing columns
       For i = 2 To LastRow
            
             'Checking If I am still within the same Ticker's code
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                 closing = ws.Cells(i, 6).Value
                 
             'Or If I am not within the same Ticker's code
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                   opening = ws.Cells(i, 3).Value
                   
                
            End If
            
            'I check the value of the yearly opening and closing values If they are > then 0
            If opening > 0 And closing > 0 Then
            
            'I defined the increase value
                increase = closing - opening
                
            'I defined the percent change value
                Percent_Change = increase / opening
            
                
                ws.Cells(select_row, 10).Value = increase
                
                ws.Cells(select_row, 11).Value = FormatPercent(Percent_Change)
            
            'I reset the yearly opening and closing values
                closing = 0
                
                opening = 0
                
            'I add one to column 10 and 11 rows
               select_row = select_row + 1
                
             End If
             
         Next i
       
        'I refer to the currently active worksheet in the active workbook and find Max,Min values of greatest incresa or decrease rates
        greatest_increase = WorksheetFunction.Max(ActiveSheet.Columns("k"))
       greatest_decrease = WorksheetFunction.Min(ActiveSheet.Columns("k"))
        greatest_vol = WorksheetFunction.Max(ActiveSheet.Columns("l"))
        
        'I assign the percent format to relate ranges
        Range("Q2").Value = FormatPercent(greatest_increase)
        Range("Q3").Value = FormatPercent(greatest_decrease)
        Range("Q4").Value = greatest_vol
        
          'I Loop through all  columns of yearly and percent changes
         For i = 2 To LastRow
                  
            If greatest_increase = Cells(i, 11).Value Then
                Range("P2").Value = Cells(i, 9).Value
                
            ElseIf greatest_decrease = Cells(i, 11).Value Then
                Range("P3").Value = Cells(i, 9).Value
                
            ElseIf greatest_vol = Cells(i, 12).Value Then
                Range("P4").Value = Cells(i, 9).Value
            
            End If
        Next i
        
      'I loop through the column of yearly change and find If the valur of the row is bigger then 0 or not.
     For i = 2 To LastRow
 
       If ws.Cells(i, 10).Value > 0 Then
 
         ws.Cells(i, 10).Interior.ColorIndex = 4
 
    Else
 
         ws.Cells(i, 10).Interior.ColorIndex = 3
 
    End If
 
    Next i
    
    Next ws
    
End Sub
