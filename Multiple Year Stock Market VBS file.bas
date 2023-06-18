Attribute VB_Name = "Module1"
Sub AlphabeticalTesting():

    'Declare variables
    Dim Ticker As String
    Dim Opening As Double
    Dim Closing As Double
    Dim Volume As Double
    Dim yrChange As Double
    Dim pctChange As Double
    Dim lastrow As Long
    Dim rowCount As Integer
    Dim resultRow As Integer
    Dim maxPct As Double
    Dim maxTick As String
    Dim minPct As Double
    Dim minTick As String
    Dim maxVol As Double
    Dim volTick As String
    Dim r As Range
    Dim rVol As Range
    Dim ws As Worksheet

    
For Each ws In Worksheets
    'set default values
    resultRow = 2
    
    rowCount = 1
    
    'find the last row of the sheet
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'insert row headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % increase"
    ws.Range("N3").Value = "Greatest % decrease"
    ws.Range("N4").Value = "Greatest total volume"
    
    
    For i = 2 To lastrow
    
        'assign ticker and open value
        If (rowCount = 1) Then
            Ticker = ws.Cells(i, 1).Value
            Opening = ws.Cells(i, 3).Value
            rowCount = rowCount + 1

        'test if ticker value is the same as the next cell
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
            'set closing value
            Closing = ws.Cells(i, 6).Value

            'determine yearly and percent change
            yrChange = Closing - Opening
            pctChange = (yrChange / Opening) * 100


            'print information
            ws.Cells(resultRow, 9).Value = Ticker
            ws.Cells(resultRow, 10).Value = yrChange
            ws.Cells(resultRow, 11).Value = pctChange & "%"
            ws.Cells(resultRow, 12).Value = Volume
            
            'Reset rowcount and volume
            rowCount = 1
            Volume = 0
            
            'Move results to next row
            resultRow = resultRow + 1

        Else
            'keep track of total stock volume
            Volume = Volume + ws.Cells(i, 7).Value
            'update count of rows
            
        End If


    Next i


    'Reset result row for greatest values
    resultRow = 2
    
    Set r = ws.Range("k2:k3001")
    Set rVol = ws.Range("l2:l3001")
    
    For i = 2 To 3001
        'color code yearly change depending if change is positive or negative
        If (ws.Cells(i, 10).Value > 0) Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
        
        'Determine Greatest % increase
        If (ws.Cells(i, 11) = ws.Application.WorksheetFunction.Max(r)) Then
            maxPct = ws.Cells(i, 11).Value
            maxTick = ws.Cells(i, 9).Value
        'Determine Greatest % decrease
        ElseIf (ws.Cells(i, 11) = ws.Application.WorksheetFunction.Min(r)) Then
            minPct = ws.Cells(i, 11).Value
            minTick = ws.Cells(i, 9).Value
        'determine greatest volume
        ElseIf (ws.Cells(i, 12) = ws.Application.WorksheetFunction.Max(rVol)) Then
            maxVol = ws.Cells(i, 12).Value
            volTick = ws.Cells(i, 9).Value
        End If
    Next i

    'print greatest % increase
    ws.Range("O2").Value = maxTick
    ws.Range("P2").Value = (maxPct * 100) & "%"
    'print greatest % decrease
    ws.Range("O3").Value = minTick
    ws.Range("P3").Value = (minPct * 100) & "%"
    'print greatest volume
    ws.Range("O4").Value = volTick
    ws.Range("P4").Value = maxVol
    
    ws.Columns("A:Z").AutoFit

Next ws
End Sub

