  MARKOWITZ OPTIMIZATION IN VBA
  
  Here is our new project, we will try to make a markowitz portofolio optimization. We will make our optimization considering the fact that a weight in our portofolio could be a negative weight. It will mean that we gotta sell that stock. You could click [here](https://www.investopedia.com/terms/p/portfolio-weight.asp#:~:text=If%20you%20do%20this%20for,values%20and%20carry%20negative%20weights.) for more information.
  
  Download your data
  
  As usual our first step will be to get our data. You could directly get it from the code or you could follow those steps.
  1-Open thomson reuters.
  2-Type your indice in the search bar.
  3-Go to the content of the indice.
  4-Export your tickers to excel.
  You could also enter your tickers by hands or by a screen. Whatever the way you choose. You will have to download your prices after and the code will do everything else for you. If you alrady have your tickers list. Just go to the thomson reuters tab in excel and select fromula builder, then you will click on time periods data and download your data. Click  [here](https://training.refinitiv.com/docs/attachments/shared/eikon_office_quick_start_guide.pdf) for more information.
  Notice that you could also get your data in the way you want just be sure to fill the first column of the first sheets with names because the code run by counting how many stock you have. You will also have to rename your sheets if you don't have the same name as us.
  
  Process your data
  
  We will first declare the variables that we will use in our code
  
```{r}
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

```
 
  We will now fiill our returns page. We will first copy our stocks names, we will copy our dates, and now we will fill our returns stocks by stocks and date by date.
  
```{r}

'Fill our first non numeric data
sr.Range("A1").Value = "Dates/Stocks"

Sheets(1).Range(Sheets(1).Range("A8"), Sheets(1).Range("A8").End(xlDown)).Copy
sr.Range("B1").PasteSpecial Transpose:=True

'Fill our dates
For Each cel In sp.Range(sp.Range("B4"), sp.Range("B4").End(xlDown))
    Cells((cel.Row - 2), 1) = cel
Next cel

Range(Range("A2"), Range("A2").End(xlDown)).NumberFormat = "mm/dd/yyyy"
      
'Fill our stock returns by a loop who will iterate itself firstly on each stock and secondly on each dates for each stocks

For Each cel1 In sp.Range(sp.Range("C4"), sp.Range("C4").End(xlToRight))
    For Each cel In sp.Range(cel1, cel1.End(xlDown))
        Cells((cel.Row - 2), (cel1.Column - 1)) = cel.Value / cel.Offset(-1, 0).Value - 1
    Next cel
Next cel1
```
  
  Some of our stocks can stop to be priced for any reason. Thomson return a blank cell on the dates where our stocks don't have data. So when we compute our return we also have a blank cells. It will be a problem in our covariance matrix. 
  So we made the choice to replace those blank cells with zero. Because if one day you don't trade a stock. Your return that day will simply be zero, we think that our assumption is good. But as usual the way to treat unavailable data depend of each person. So just modify the code bellow if you want to modify it but be sure to modify it in the right way before going further.
  
```{r}
lign = sr.Range("A1").End(xlDown).Row
col = sr.Range("A1").End(xlToRight).Column

For Each cellule In sr.Range(Cells(1, 1), Cells(lign, col))
    If IsEmpty(cellule) = True Then
        cellule.Value = 0
    End If
Next cellule
 
```
  
  Our next step will be to compute the mean return and the risk (sd) of each stocks.
  
```{r}
sr.Range("A2").EntireRow.Insert
sr.Range("A2").EntireRow.Insert

sr.Range("A2").Value = "mean_return"
sr.Range("A3").Value = "sd"
    'mean computation
    For Each cel In sr.Range(sr.Range("B4"), sr.Range("B4").End(xlToRight))
        cel.Offset(-2, 0) = Application.WorksheetFunction.Average(Range(cel, cel.End(xlDown)))
    Next cel
    
    'sd computaion
    For Each cel In sr.Range(sr.Range("B4"), sr.Range("B4").End(xlToRight))
        cel.Offset(-1, 0) = VBA.Sqr(Application.WorksheetFunction.Var((Range(cel, cel.End(xlDown)))))
    Next cel
```
  
  Now we choose to treat our data ten by ten to facilitate and made our code faster. VBA code is so low so every loop not necessary, every lign of code that can be avoid should be. Here we will take our data ten by ten because the computing power of a covariance matrix of 500 by 500 will be hard to calcul. 
  We also made the choice to take our data ten by ten, because an usual way to optimize with markowitz in Excel is to use the solver, and the solver can't run if you have too much data.
  Here we made the choice to get our optimal weights by a simple algebra treatment. In Markowitz optimization we try to minimize the risk of our portofolio subject to the constraint that all the weights should sum to 100%. This is a Lagrangien problem.
  A way to resolve Lagrangian is to use matrix. You can click on [this](https://faculty.washington.edu/ezivot/econ424/portfolioTheoryMatrix.pdf) article very well explained to know more,go to the seventh page. By clicking on it and reading all the paper, you will also have a way to optimize by Markowitz in Rstudio.
  So let us go we willl copy our data then by then to the next sheets.
  
```{r}
res.Cells.Clear
max_iteration = Sheets(1).Range("B8").End(xlDown).Row / 10

For copy_num = 1 To 2
 step = (copy_num - 1) * 10

    Range("B1:K3").Offset(0, step).Copy Sheets("Infos").Range("B2")
```
  
 Now we are in the heart of oup main loop all the code that will follow will be in that loop. The line of code "Next copy_num" will be the end of our loop.
 We will now compute our covariance matrix on each iteration on the loop.
 
```{r}
    sr.Range("A2").EntireRow.Insert
    
    For i = 1 To Sheets(1).Range("B8").End(xlDown).Row
     Cells(2, i) = i
    Next i
    
    col = Application.WorksheetFunction.HLookup(si.Range("B2"), sr.Range("A1:XFD2"), 2, False)
    
    sr.Rows(2).EntireRow.Delete
    
    For j = 0 To 9
        For i = 0 To 9
            x1 = col + i
            x2 = col + j
            si.Cells(j + 10, i + 2) = Application.WorksheetFunction.Covar(sr.Range(Cells(4, x1), Cells(4, x1).End(xlDown)), sr.Range(Cells(4, x2), Cells(4, x2).End(xlDown)))
        Next i
    Next j
```
 
 We will compute our optimal weights by matrix inverse as explained on the link above. When we will got our results, we will copy them on the results sheets to see each loop iteration, and that with the goal to read each stocks weights for any reason even if we will got exactly the number of stocks we want in our portofolio at the end.
 
```{r}
    
        si.Range("B20:K20").Value = 1
        si.Range("L10:L19").Value = 1
        si.Range("L20").Value = 0
          
        si.Range("P10:P19").Value = 0
        si.Range("P20").Value = 1
           
        si.Range("B4:K4").Copy
        si.Range("N10:N19").PasteSpecial Transpose:=True
           
        si.Range("N20").Value = "Lambda"
            
        si.Range("S10:S19") = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MInverse(si.Range("B10:L20")), si.Range("P10:P20"))
         
        si.Range("S10:S19").Copy
        si.Range("B27:K27").PasteSpecial Transpose:=True
    
    
    If copy_num = 1 Then
    GoTo copy_num_one
    Else
    GoTo copy_num_dif_one
    End If
    
copy_num_one:
  
        si.Range("A2:K2").Copy res.Range("A1")
        si.Range("A27:K27").Copy res.Range("A2")
        si.Range("A29:B29").Copy res.Range("A3")
        
copy_num_dif_one:
        Position = res.Range("A1").End(xlDown).Row
        si.Range("A2:K2").Copy res.Range("A" & (Position + 1))
        si.Range("A27:K27").Copy res.Range("A" & (Position + 2))
        si.Range("A29:B29").Copy res.Range("A" & (Position + 3))
```
 
 We will now copy all the datas, copmute the absolute value and sort them to return the first top hits.
 
```{r}

'copy our results to find what's the stocks that will meet tour criteria
For i = 1 To ((res.Range("A1").End(xlDown).Row / 3) + 1)
    step = (i - 1) * 3
    step1 = (i - 1) * 10
    res.Range("B4:K4").Offset(step, 0).Copy
    res.Range("M1").Offset(step1, 0).PasteSpecial Transpose:=True
Next i

For i = 1 To ((res.Range("A1").End(xlDown).Row / 3) + 1)
    step = (i - 1) * 3
    step1 = (i - 1) * 10
    res.Range("B5:K5").Offset(step, 0).Copy
    res.Range("N1").Offset(step1, 0).PasteSpecial Transpose:=True
Next i
For Each cel In res.Range(res.Range("N1"), res.Range("N1").End(xlDown))
    cel.Offset(0, 1) = Abs(cel)
Next cel

'sort our stocks
res.Range(res.Range("M1"), res.Range("M1").End(xlDown).Offset(0, 2)).Sort Key1:=res.Range("O1"), Order1:=xlDescending, Header:=xlNo


```
 
 The user will be asked how many stocks he want in his portofolio in an inputbox and the code below will return those stocks to him.
 
```{r}
res.Range("R1:S1").merge
res.Range("R1").Value = "Here are your stocks"
res.Range("R1").VerticalAlignment = xlCenter
res.Range("R1").HorizontalAlignment = xlCenter
res.Range("R1").Interior.ColorIndex = 5
res.Range("R1").Font.Size = 14

res.Range("R2").Value = "Stocks"
res.Range("S2").Value = "Weights"

res.Range("R4:S25000").Cells.ClearContents

stock_number = InputBox("Please, telll us how many sotcks do you want to have in your potofolio?")

res.Range(res.Range("M1"), res.Range("N" & stock_number)).Copy res.Range("R4")
res.Range("R:T").Interior.ColorIndex = 0

res.Range("R4:S25000").VerticalAlignment = xlCenter
res.Range("R4:S25000").HorizontalAlignment = xlCenter

```
 
 
  
