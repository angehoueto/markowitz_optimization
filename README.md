# markowitz_optimization
'declare our variables

Dim sp As Worksheet
Dim sr As Worksheet
Dim si As Worksheet
Dim res As Worksheet
Dim cel As Range
Dim cel1 As Range

Set sp = Sheets("Stocks Prices")
Set sr = Sheets("Stocks Returns")
Set si = Sheets("Infos")
Set res = Sheets("Results")

'Fill our first non numeric data
sr.Range("A1").Value = "Dates/Stocks"

Sheets(1).Range(Sheets(1).Range("A8"), Sheets(1).Range("A8").End(xlDown)).Copy
sr.Range("B1").PasteSpecial Transpose:=True

'Fill our dates
For Each cel In sp.Range(sp.Range("B4"), sp.Range("B4").End(xlDown))
    Cells((cel.Row - 2), 1) = cel
Next cel

Range(Range("A2"), Range("A2").End(xlDown)).NumberFormat = "mm/dd/yyyy"
