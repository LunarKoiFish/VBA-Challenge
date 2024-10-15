Attribute VB_Name = "Module1"
Sub stockdata()

    Dim ws As Worksheet
    Dim lastRow As Long
    
    
    Dim r As Long
    Dim ticker As String
    Dim quarterlychange As Double
    Dim closingprice As Double
    Dim openingprice As Double
    Dim percentchange As Single
    Dim stock_vl As LongLong
    Dim outrow As Long
    Dim maxpercent As Single
    Dim minpercent As Single
    Dim maxvolume As Single
    
    
    



    For Each ws In ThisWorkbook.Worksheets
    
        ws.Cells(1, "I").Value = "Tickers"
        ws.Cells(1, "J").Value = "Quarterly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volue"
        ws.Cells(1, "P").Value = "Ticker"
        ws.Cells(1, "Q").Value = "Value"
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"
    
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        outrow = 2
        stock_vl = 0
        openingprice = ws.Cells(2, 3).Value
        maxpercent = 0
        minpercent = 0
        maxvolume = 0
    

            For r = 2 To lastRow
                
              
                
    
                If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1) Then
                    ws.Cells(outrow, "I").Value = ws.Cells(r, 1).Value
                    ws.Cells(outrow, "L").Value = stock_vl + ws.Cells(r, 7).Value
            
            
            
                    closingprice = ws.Cells(r, 6).Value
                
                    quarterlychange = closingprice - openingprice
                    
                    percentchange = (closingprice / openingprice) - 1
            
                    ws.Cells(outrow, "j").Value = quarterlychange
                    ws.Cells(outrow, "k").Value = percentchange
                    
                    
                    If ws.Cells(outrow, "j").Value > "0" Then
                        ws.Cells(outrow, "j").Interior.ColorIndex = 4
                    
                    ElseIf ws.Cells(outrow, "j").Value < "0" Then
                        ws.Cells(outrow, "j").Interior.ColorIndex = 3
                
                    Else
                        ws.Cells(outrow, "j").Interior.ColorIndex = 2
                
                
                    End If
                    
                    
                    
                    If ws.Cells(outrow, "k").Value > maxpercent Then
                        maxpercent = ws.Cells(outrow, "k").Value
                        ws.Cells(2, "q").Value = maxpercent
                        ws.Cells(2, "P").Value = ws.Cells(outrow, "i").Value
                    
                    End If
                
                    If ws.Cells(outrow, "k").Value < minpercent Then
                        minpercent = ws.Cells(outrow, "k").Value
                        ws.Cells(3, "q").Value = minpercent
                        ws.Cells(3, "P").Value = ws.Cells(outrow, "i").Value
                    
                    End If
                
                    If ws.Cells(outrow, "L").Value > maxvolume Then
                        maxvolume = ws.Cells(outrow, "L").Value
                        ws.Cells(4, "q").Value = maxvolume
                        ws.Cells(4, "P").Value = ws.Cells(outrow, "i").Value
                    
                    End If
                    
            
                    outrow = outrow + 1
                    stock_vl = 0
        
                    openingprice = ws.Cells(r + 1, 3).Value
        
                Else
                    stock_vl = stock_vl + ws.Cells(r, 7).Value
                    
    
        
                End If
                
                
                
                
                ws.Cells(r, "j").NumberFormat = "0.00"
                ws.Cells(r, "k").NumberFormat = "0.00%"
                ws.Cells(r, "l").NumberFormat = "0"
                ws.Cells(2, "Q").NumberFormat = "0.00%"
                ws.Cells(3, "Q").NumberFormat = "0.00%"
                
                
                
            
            Next r
        
    Next ws

    
    
    
End Sub





