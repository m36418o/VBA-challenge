Attribute VB_Name = "Module1"
Sub StockAnalysis():
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        'MsgBox ws.Name
        ws.Activate

        'Var init
        Dim ticker As String 'To store ticker symbol
        Dim maxticker As String
        Dim minticker As String
        Dim tolticker As String
        Dim volume As Double 'To store total volume of each stock
        Dim lastrow As Long
        Dim tablerow As Integer
        Dim openprice As Double
        Dim closeprice As Double
        Dim max As Double
        Dim min As Double
        Dim total As Double
                
        tablerow = 2
        volume = 0
        total = 0
        max = 0
        min = 0
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        openprice = Cells(2, 3).Value
        
        'Setting headers
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percentage Change"
        Cells(1, 12) = "Total Stock volume"
        Cells(1, 15) = "Ticker"
        Cells(1, 16) = "Value"
        Cells(2, 14) = "Greatest % Increase"
        Cells(3, 14) = "Greatest % Decrease"
        Cells(4, 14) = "Greatest Total Volume"
      
        ' Begin the loop.
        For I = 2 To lastrow:
            volume = volume + Cells(I, 7).Value
            'Conditional to check for ticker change
            If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
                'Grabbing the values necessary for the table
                ticker = Cells(I, 1).Value
                closeprice = Cells(I, 6).Value
                
                'Conditional to capture tickers with opening price of zero
                If openprice = 0 Then
                    Cells(tablerow, 9).Value = ticker 'Sets ticker
                    Cells(tablerow, 10).Value = 0 'Yearly change
                    Cells(tablerow, 11).Value = 0 'Percentage change
                    Cells(tablerow, 12).Value = volume 'Total trade volume
                Else
                    'Set the values to corresponding cell in table
                    Cells(tablerow, 9).Value = ticker 'Sets ticker
                    Cells(tablerow, 10).Value = closeprice - openprice 'Yearly change
                    Cells(tablerow, 11).Value = (closeprice - openprice) / openprice 'Percentage change
                    Cells(tablerow, 12).Value = volume 'Total trade volume
                End If
                
                'Bonus conditionals for % increase/decrease
                If Cells(tablerow, 11).Value > max Then
                    maxticker = Cells(tablerow, 9).Value
                    max = Cells(tablerow, 11).Value
                ElseIf Cells(tablerow, 11).Value < min Then
                    minticker = Cells(tablerow, 9).Value
                    min = Cells(tablerow, 11).Value
                End If
                'Bonus conditionals for max volume
                If Cells(tablerow, 12).Value > total Then
                    tolticker = Cells(tablerow, 9).Value
                    total = Cells(tablerow, 12).Value
                End If
                
                'Conditional for styling and formating
                If Cells(tablerow, 10).Value > 0 Then
                    Cells(tablerow, 10).Interior.ColorIndex = 4
                ElseIf Cells(tablerow, 10).Value < 0 Then
                    Cells(tablerow, 10).Interior.ColorIndex = 3
                ElseIf Cells(tablerow, 10).Value = 0 Then
                    Cells(tablerow, 10).Interior.ColorIndex = 0
                End If
                
                Cells(tablerow, 11).Style = "percent" 'sets cell type to percentage
                'Resetting variables
                openprice = Cells(I + 1, 3)
                tablerow = tablerow + 1 'Adds one increment
                volume = 0 'Volume reset
            End If
        Next I
        
        'Max increase
        Cells(2, 15) = maxticker
        Cells(2, 16) = max
        Cells(2, 16).Style = "percent" 'sets cell type to percentage
        'Max decrease
        Cells(3, 15) = minticker
        Cells(3, 16) = min
        Cells(3, 16).Style = "percent" 'sets cell type to percentage
        'Max volume
        Cells(4, 15) = tolticker
        Cells(4, 16) = total
        
    Next ws
    
End Sub


