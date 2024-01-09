Sub StockData()

    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets

        Dim Ticker As String

        Dim Volume As Double
        Volume = 0

        Dim OpenPrice As Double
        OpenPrice = Cells(2, 3).Value
        Dim ClosePrice As Double
    
        Dim YearlyChange As Double
        Dim PercentChange As Double

        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
        Dim TempValue As Double
        Dim Value As Double

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                Ticker = ws.Cells(i, 1).Value

                Volume = Volume + ws.Cells(i, 7).Value
            
                ClosePrice = ws.Cells(i, 6).Value
            
                YearlyChange = (ClosePrice - OpenPrice)
            
                If (OpenPrice = 0) Then
            
                    PercentChange = 0
            
                Else
            
                    PercentChange = (YearlyChange / OpenPrice)
            
                End If

                ws.Range("I" & Summary_Table_Row).Value = Ticker
            
                ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            
                ws.Range("K" & Summary_Table_Row).Value = PercentChange
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

                ws.Range("L" & Summary_Table_Row).Value = Volume

                Summary_Table_Row = Summary_Table_Row + 1

                Volume = 0
            
                OpenPrice = ws.Cells(i + 1, 3).Value
    
            Else
    
                Volume = Volume + ws.Cells(i, 7).Value
    
            End If

        Next i
    
        LastRow_Summary_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        For i = 2 To LastRow_Summary_Table
    
                If ws.Cells(i, 10).Value > 0 Then
            
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                
                Else
            
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                
                End If
            
        Next i
    
        ws.Range("I:L").Sort Key1:=ws.Range("K1"), Order1:=xlDescending, Header:=xlYes
    
        ws.Range("P2").Value = ws.Range("I2").Value
        ws.Range("Q2").Value = ws.Range("K2").Value
        ws.Range("Q2").NumberFormat = "0.00%"
    
        ws.Range("I:L").Sort Key1:=ws.Range("K1"), Order1:=xlAscending, Header:=xlYes
    
        ws.Range("P3").Value = ws.Range("I2").Value
        ws.Range("Q3").Value = ws.Range("K2").Value
        ws.Range("Q3").NumberFormat = "0.00%"
    
        ws.Range("I:L").Sort Key1:=ws.Range("L1"), Order1:=xlDescending, Header:=xlYes
    
        ws.Range("P4").Value = ws.Range("I2").Value
        ws.Range("Q4").Value = ws.Range("L2").Value
    
        ws.Range("I:L").Sort Key1:=ws.Range("I1"), Order1:=xlAscending, Header:=xlYes
    
    Next ws
    
End Sub
