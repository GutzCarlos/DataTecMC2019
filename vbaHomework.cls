Sub hw()
Dim i, j, nr, nt, op, cl, inc, dcr, tv, vol As Double
Dim ninc, ndcr, ntv As String
Dim ws As Worksheet

    For Each ws In Worksheets
        Worksheets(ws.Name).Activate
        nr = Cells(Rows.Count, 1).End(xlUp).Row
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volum"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        For i = 2 To nr
            nt = Cells(Rows.Count, 9).End(xlUp).Row
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                Cells(nt + 1, 9).Value = Cells(i, 1).Value
                If i <> 2 Then
                    cl = Cells(i - 1, 6).Value
                    Cells(nt, 10).Value = cl - op
                    Cells(nt, 11).Value = ((cl - op) / op) * 100
                End If
                op = Cells(i, 3).Value
            End If
            If i = nr Then
                cl = Cells(i - 1, 6).Value
                Cells(nt, 10).Value = cl - op
                Cells(nt, 11).Value = ((cl - op) / op) * 100
            End If
        Next i
        MsgBox (nt & " Tickers")
        
        For i = 2 To nt
            vol = 0
            For j = 2 To nr
                If Cells(j, 1).Value = Cells(i, 9).Value Then
                vol = vol + Cells(j, 7).Value
                End If
            Next j
            Cells(i, 12).Value = vol
            If Cells(i, 10).Value < 0 Then
                Cells(i, 10).Interior.Color = RGB(120, 3, 3)
            Else
                Cells(i, 10).Interior.Color = RGB(131, 131, 0)
            End If
        Next i
        inc = Cells(2, 11)
        dcr = Cells(2, 11)
        tv = Cells(2, 12)
        For i = 3 To nt
            If Cells(i, 11).Value > inc Then
                inc = Cells(i, 11).Value
                ninc = Cells(i, 9).Value
            End If
            If Cells(i, 11).Value < dcr Then
                dcr = Cells(i, 11).Value
                ndcr = Cells(i, 9).Value
            End If
            If Cells(i, 12).Value > tv Then
                tv = Cells(i, 12).Value
                ntv = Cells(i, 9).Value
            End If
        Next i
        Cells(2, 16).Value = ninc
        Cells(2, 17).Value = inc
        Cells(3, 16).Value = ndcr
        Cells(3, 17).Value = dcr
        Cells(4, 16).Value = ntv
        Cells(4, 17).Value = tv
    Next ws
End Sub


