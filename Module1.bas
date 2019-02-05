Attribute VB_Name = "Module1"
Sub StkAdds()
Dim ticker As String
Dim lookup_range As Range
Dim total_volume As Long
Dim volume As Long
Dim summary_table_row As Long
Dim sheet As String
Dim ticker_name_header As String
Dim ticker_volume_header As String

    summary_table_row = 2
    volume = 0
    total_volume = 0
    Column = 1
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"

For i = 2 To 798000
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    ticker_name = Cells(i, 1).Value
    ticker_total = ticker_total + Cells(i, 7).Value
    Range("I" & summary_table_row).Value = ticker_name
    Range("j" & summary_table_row).Value = ticker_total
    Range("I1:J1").Columns.AutoFit
    
    summary_table_row = summary_table_row + 1
    
    ticker_total = 0

Else
    ticker_name = Cells(i + 1, 1).Value
    ticker_total = ticker_total + Cells(i, 7).Value
    
    End If
    
Next i

End Sub
