Attribute VB_Name = "Module1"
Sub stocks()


Dim ticker As String
Dim volume As Double
volume = 0
Dim summarytablerow As Integer
summarytablerow = 2
Dim change As Double
Dim counter As Long
counter = 0
Dim percent As Double
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

ticker = Cells(i, 1).Value

volume = volume + Cells(i, 7).Value

change = Cells(i, 6).Value - Cells(i - counter, 3).Value

percent = (change / Cells(i - counter, 3).Value) * 100

Range("I" & summarytablerow).Value = ticker

Range("J" & summarytablerow).Value = volume

Range("K" & summarytablerow).Value = change

Range("L" & summarytablerow).Value = percent

summarytablerow = summarytablerow + 1

volume = 0

Else

volume = volume + Cells(i, 7).Value

counter = counter + 1

End If


Next i

lastrow1 = Cells(Rows.Count, 12).End(xlUp).Row

For i = 2 To lastrow1

If Cells(i, 12).Value > 0 Then

Cells(i, 12).Interior.ColorIndex = 4

Else

Cells(i, 12).Interior.ColorIndex = 3

End If

Next i


Dim worst As Double
Dim best As Double
Dim bestvol As Double
Dim rng As Range
Dim rngvol As Range
Dim rngtic As Range
Dim tickerb As String
Dim tickerc As String


lastrow1 = Cells(Rows.Count, 12).End(xlUp).Row
lastrowvol = Cells(Rows.Count, 10).End(xlUp).Row
lastrowt = Cells(Rows.Count, 9).End(xlUp).Row

Set rng = Range("L2:L290")


Set rngvol = Range("J2:J290")

Set rngtic = Range("I2:I290")


best = Application.WorksheetFunction.Max(rng)

worst = Application.WorksheetFunction.Min(rng)

bestvol = Application.WorksheetFunction.Max(rngvol)


Range("O2").Value = best

Range("O3").Value = worst

Range("O4").Value = bestvol





End Sub

