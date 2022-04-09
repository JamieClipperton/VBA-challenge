Attribute VB_Name = "Module1"
Sub Market_Analysis_Report():

For Each ws In Worksheets

ws.Range("I1").Value = "Ticker ID"
ws.Range("J1").Value = "Annual Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Annual Stock Volume"
        
Dim TickerID As String
Dim LastRow As Long
Dim AnnualVolume As Double
    AnnualVolume = 0
Dim SummaryTableRow As Long
    SummaryTableRow = 2
Dim AnnualOpen As Double
Dim AnnualClose As Double
Dim AnnualChange As Double
Dim PreviousAmount As Long
    PreviousAmount = 2
Dim PercentChange As Double
        

LastRow = ws.cells(Rows.Count, 1).End(xlUp).Row
        
For i = 2 To LastRow


AnnualVolume = AnnualVolume + ws.cells(i, 7).Value

If ws.cells(i + 1, 1).Value <> ws.cells(i, 1).Value Then

TickerID = ws.cells(i, 1).Value

ws.Range("I" & SummaryTableRow).Value = TickerID

ws.Range("L" & SummaryTableRow).Value = AnnualVolume
    AnnualVolume = 0

AnnualOpen = ws.Range("C" & PreviousAmount)
AnnualClose = ws.Range("F" & i)
                AnnualChange = AnnualClose - AnnualOpen
                ws.Range("J" & SummaryTableRow).Value = AnnualChange

                If AnnualOpen = 0 Then
                    PercentChange = 0
                Else
                    AnnualOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = AnnualChange / AnnualOpen
                End If

                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange

                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
            
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                End If
            Next i

        ws.Columns("I:L").AutoFit

    Next ws

End Sub
