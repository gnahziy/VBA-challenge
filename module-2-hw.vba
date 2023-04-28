Sub Testing():
    'declare j and worksheet count
    Dim J As Integer
    Dim worksheet_count As Integer
    'count my worksheet
    worksheet_count = ActiveWorkbook.Worksheets.Count
    'looping three all three shees
    For J = 1 To worksheet_count
        'activate all worksheets
        Worksheet = Worksheets(J).Name
            Worksheets(Worksheet).Activate
        
        'Headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volumn"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volumn"

        'set an variable for i
        Dim i As Long
        'set an variable for holding ticker
        Dim ticker As String
        'set a variable for holding total volumn per ticker
        Dim volumn As LongLong
        volumn = 0
        'put individual ticker name on the sheet
        Dim table As Integer
        table = 2
        'set variable for yearly change
        Dim yearly As Long
        Dim open_row As Long
        open_row = 2
        'Last Row
        Dim RowCount As Long
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
   
        'looping ticker name
        For i = 2 To RowCount
            'check if we're still within the same ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'set ticker name
                ticker = Cells(i, 1).Value
            
                'this section is for calculating yearly change
                close_price = Cells(i, 6).Value
                open_price = Cells(open_row, 3).Value
                Change = close_price - open_price
                
                 'this section is for percent change
                 Percent = Change / open_price
                
                'add to ticker volumn
                volumn = volumn + Cells(i, 7).Value
                
                'print ticker
                Range("I" & table).Value = ticker
                'print volumn
                Range("L" & table).Value = volumn
                'print yearly change
                Range("J" & table).Value = Change
                'print percent change
                Range("K" & table).Value = Percent
                Range("K1:K91").NumberFormat = "0.00%"
                'print greatest % increase
                Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & RowCount)) * 100
                'print greatest % dncrease
                Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & RowCount)) * 100
                'print greatest total volumn
                Range("Q4") = WorksheetFunction.Max(Range("L2:L" & RowCount))
                
                'find greatest % increase ticker
                Increase_Index = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
                'find greatest % decrease ticker
                Decrease_Index = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
                'find greatest total volumn ticker
                Volumn_Index = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & RowCount)), Range("L2:L" & RowCount), 0)
             
                'print greatest % increase ticker
                Range("P2") = Cells(Increase_Index + 1, 9)
                'print greatest % dncrease ticker
                Range("P3") = Cells(Decrease_Index + 1, 9)
                'print greatest total volumn ticker
                Range("P4") = Cells(Volumn_Index + 1, 9)
                
                'Conditional Formatting year
                Select Case Change
                    Case Is > 0
                        Range("J" & table).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & table).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & table).Interior.ColorIndex = 0
                                
                End Select
                
                'Conditional Formatting percent
                Select Case Percent
                    Case Is > 0
                        Range("K" & table).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("K" & table).Interior.ColorIndex = 3
                    Case Else
                        Range("K" & table).Interior.ColorIndex = 0
                                
                End Select
            
                'add additional lines to the tabel
                table = table + 1
                'reset volumn
                volumn = 0
                'calculate yearly change
                open_row = i + 1
                
                Else
                'add to ticker volumn
                volumn = volumn + Cells(i, 7).Value
            
            End If
        Next i
    Next J
End Sub

