Attribute VB_Name = "Module1"
Sub AlphabeticalTest()
    Dim ws As Worksheet
    
        For Each ws In ThisWorkbook.Worksheets
        
        ws.Range("K1").Value = "Ticker"
        ws.Range("L1").Value = "Yearly Change"
        ws.Range("M1").Value = "Percent Change"
        ws.Range("N1").Value = "Total Stock Volume"
        ws.Range("R1").Value = "Ticker"
        ws.Range("S1").Value = "Value"
        ws.Range("Q2").Value = "Greatest % Increase"
        ws.Range("Q3").Value = "Greatest % Decrease"
        ws.Range("Q4").Value = "Greatest Total Volume"
        
        Dim TickerName As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim TotalVolume As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim RowCounter As Integer
        
        TotalVolume = 0
        YearlyChange = 0
        PercentChange = 0
        RowCounter = 2
        YearlyChangeTable = 2
        PercentChangeTable = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To LastRow
                If i = 2 Then
                
                TickerName = ws.Cells(i, 1).Value
                OpenPrice = ws.Cells(i, 3).Value
                ClosePrice = ws.Cells(i, 6).Value
                TotalVolume = ws.Cells(i, 7).Value
                
                Else
                
                    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    OpenPrice = ws.Cells(i, 3).Value
                    TickerName = ws.Cells(i, 1).Value
                    ClosePrice = ws.Cells(i, 6).Value
                    
                    End If
                    
                        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                        TickerName = ws.Cells(i, 1).Value
                        ClosePrice = ws.Cells(i, 6).Value
                     
                        YearlyChange = ClosePrice - OpenPrice
                   
                        
                            If OpenPrice <> 0 Then
                            PercentChange = (ClosePrice - OpenPrice) / OpenPrice
                            Else
                            PercentChange = 0
                            End If
                        
     
                        TotalVolume = ws.Cells(i, 7).Value + TotalVolume
                        ws.Range("K" & RowCounter).Value = TickerName
                        ws.Range("L" & RowCounter).Value = YearlyChange
                        ws.Range("M" & RowCounter).Value = PercentChange
                        ws.Range("M" & RowCounter).Value = FormatPercent(PercentChange, 2)
                        ws.Range("N" & RowCounter).Value = TotalVolume
                        RowCounter = RowCounter + 1
               
                        YearlyChange = 0
                        PercentChange = 0
                        TotalVolume = 0
                        Else
                        
                        ClosePrice = ws.Cells(i, 6).Value
                        TotalVolume = ws.Cells(i, 7).Value + TotalVolume
                        
                        End If
                    
                 End If
                
             Next i
             
        Dim NewLastRow As Integer
        Dim GrPerIncrease As Double
        Dim GrPerDecrease As Double
        Dim GrTotVolume As Double
        Dim FndRng As Range
        Dim lRow As Long
        
        NewLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        GrPerIncrease = WorksheetFunction.Max(ws.Range("M2:M" & NewLastRow)) * 100
              
        Set FndRng = ws.Range("M2:M" & NewLastRow).Find(what:=GrPerIncrease)
        lRow = FndRng.Row
        ws.Range("R2").Value = ws.Cells(lRow, 11)
      
        GrPerDecrease = WorksheetFunction.Min(ws.Range("M2:M" & NewLastRow)) * 100
        
        Set FndRng = ws.Range("M2:M" & NewLastRow).Find(what:=GrPerDecrease)
        lRow = FndRng.Row
        ws.Range("R3").Value = ws.Cells(lRow, 11)
        
        
        GrTotVolume = WorksheetFunction.Max(ws.Range("N2:N" & NewLastRow))
        
        Set FndRng = ws.Range("N2:N" & NewLastRow).Find(what:=GrTotVolume)
        lRow = FndRng.Row
        ws.Range("R4").Value = ws.Cells(lRow, 11)
                       
        ws.Range("S2").Value = GrPerIncrease
        ws.Range("S3").Value = GrPerDecrease
        ws.Range("S4").Value = GrTotVolume
        
               
        For i = 2 To NewLastRow
            
        ws.Range("S2").Value = "%" & WorksheetFunction.Max(ws.Range("M2:M" & NewLastRow)) * 100
        ws.Range("S3").Value = "%" & WorksheetFunction.Min(ws.Range("M2:M" & NewLastRow)) * 100
        ws.Range("S4").Value = WorksheetFunction.Max(ws.Range("N2:N" & NewLastRow))
                
            
        Next i
        
    Next ws

End Sub
