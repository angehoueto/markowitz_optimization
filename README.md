  #MARKOWITZ OPTIMIZATION IN VBA
  Here is our new project, we will try to make a markowitz portofolio optimization. We will make our optimization considering the fact that a weight in our portofolio could be a negative weight. It will mean that we gotta sell that stock. You could click [here]https://www.investopedia.com/terms/p/portfolio-weight.asp#:~:text=If%20you%20do%20this%20for,values%20and%20carry%20negative%20weights.) for more information.
  ##Download your data
  As usual our first step will be to get our data. You could directly get it from the code or you could follow those steps.
  1-Open thomson reuters.
  2-Type your indice in the search bar.
  3-Go to the content of the indice.
  4-Export your tickers to excel.
  You could also enter your tickers by hands or by a screen. Whatever the way you choose. You will have to download your prices after and the code will do everything else for you. If you alrady have your tickers list. Just go to the thomson reuters tab in excel and select fromula builder, then you will click on time periods data and download your data. Click  [here] (https://training.refinitiv.com/docs/attachments/shared/eikon_office_quick_start_guide.pdf) for more information.
  Notice that you could also get your data in the way you want just be sure to fill the first column of the first sheets with names because the code run by counting how many stock you have. You will also have to rename your sheets if you don't have the same name as us.
  ##Process your data
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
