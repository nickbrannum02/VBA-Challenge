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

    
For Each ws In Worksheets
    'set default values
    resultRow = 2
    
    rowCount = 1
       
    Columns("A:Z").AutoFit
    
    'find the last row of the sheet
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'insert row headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % increase"
    Range("N3").Value = "Greatest % decrease"
    Range("N4").Value = "Greatest total volume"
    
    
    For i = 2 To lastrow
    
        'assign ticker and open value
        If (rowCount = 1) Then
            Ticker = Cells(i, 1).Value
            Opening = Cells(i, 3).Value
            rowCount = rowCount + 1

        'test if ticker value is the same as the next cell
        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1) Then
            'set closing value
            Closing = Cells(i, 6).Value

            'determine yearly and percent change
            yrChange = Opening - Closing
            pctChange = (yrChange / Opening) * 100


            'print information
            Cells(resultRow, 9).Value = Ticker
            Cells(resultRow, 10).Value = yrChange
            Cells(resultRow, 11).Value = pctChange & "%"
            Cells(resultRow, 12).Value = Volume
            
            'Reset rowcount and volume
            rowCount = 1
            Volume = 0
            
            'Move results to next row
            resultRow = resultRow + 1

        Else
            'keep track of total stock volume
            Volume = Volume + Cells(i, 7).Value
            'update count of rows
            
        End If


    Next i


    'Reset result row for greatest values
    resultRow = 2
    
    Set r = Range("k2:k91")
    Set rVol = Range("l2:l91")
    
    For i = 2 To 91
        'color code yearly change depending if change is positive or negative
        If (Cells(i, 10).Value > 0) Then
            Cells(i, 10).Interior.ColorIndex = 4
        Else
            Cells(i, 10).Interior.ColorIndex = 3
        End If
        
        'Determine Greatest % increase
        If (Cells(i, 11) = Application.WorksheetFunction.Max(r)) Then
            maxPct = Cells(i, 11).Value
            maxTick = Cells(i, 9).Value
        'Determine Greatest % decrease
        ElseIf (Cells(i, 11) = Application.WorksheetFunction.Min(r)) Then
            minPct = Cells(i, 11).Value
            minTick = Cells(i, 9).Value
        'determine greatest volume
        ElseIf (Cells(i, 12) = Application.WorksheetFunction.Max(rVol)) Then
            maxVol = Cells(i, 12).Value
            volTick = Cells(i, 9).Value
        End If
    Next i

    'print greatest % increase
    Range("O2").Value = maxTick
    Range("P2").Value = (maxPct * 100) & "%"
    'print greatest % decrease
    Range("O3").Value = minTick
    Range("P3").Value = (minPct * 100) & "%"
    'print greatest volume
    Range("O4").Value = volTick
    Range("P4").Value = maxVol

Next ws
End Sub
