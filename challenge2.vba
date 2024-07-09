Attribute VB_Name = "Module1"
Sub stockforquarter():

For Each ws In ThisWorkbook.Worksheets



 
Dim Ticker As String
Dim QyarterlyChange As Double
Dim PercentChange As Double
Dim TotalVolume As Double
Dim LastRow As Long
Dim SummaryRow As Long
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestVolume As Double
Dim increaseTicker As String
Dim decreaseTicker As String
Dim volumeTicker As String
        

  WorksheetName = ws.Name
        

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 15).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
        
      
        TickCount = 2
        
      
        j = 2
        
       
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
        
  
        For i = 2 To LastRow
            
         
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               
    Ticker = ws.Cells(i, 1).Value
                
                
                openPrice = ws.Cells(i, 3).Value
                closePrice = ws.Cells(i, 5).Value
                quarterlyChange = ClosingPrice - OpeningPrice
                percentageChange = (quarterlyChange / openPrice) * 100
                

        
                
                    End If
                    
            TickCount = TickCount + 1
                
            TotalVolume = TotalVolume + ws.Cells(i, 6).Value
            
           j = i + 1
           
           TotalVolume = 0
           
           
                
                
            
            Next i
      
        LastRowI = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
   
       prevQuarter = 0
         
            For i = 2 To LastRowI
            
            Ticker = ws.Cells(i, 1).Value
            
                
                If ws.Cells(i, 12).Value > greatestIncrease Then
                greatestIncrease = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
               
                If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                
                End If
                
               
                If ws.Cells(i, 11).Value < greatestdec Then
                greatestdec = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
          
                End If
                
   
        
            Next i
            
        
            
    Next ws
        
End Sub
