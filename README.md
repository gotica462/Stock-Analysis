# VBA Challenge
 
## Project Overview

The purpose of this analyisis is to refractor our module 2 VBA code for Stocks Analysis to  collect the same information that we did for all our stocks and determine if our refracted code was succesfull and more effcient. In simple terms, we'll see if our VBA script ran faster while giving us the same results.

## Results

After reviewing the data for the stocks in  2017 and 2018 in Steve's workbook we can see that all stocks in 2017 had a positive return except one. On the contrary, in 2018 most of our stocks gave a negative return except those from "ENPH" and "RUN". We can conclude thar the stocks were more profitable in 2017 and if you're are going to choose a stock to invest should only be on the two that have a positive return in both years.

![image](https://user-images.githubusercontent.com/99451833/155895442-3a7ae766-f774-438b-aee5-11b65caa8f3f.png) ![image](https://user-images.githubusercontent.com/99451833/155895897-d6e96e70-2e25-477f-8e33-f8cad862b44c.png)

When we refactored the code, we created output arrays for each ticker's volume, starting price, and ending price. 



We use the arrays so We don’t have to declare one type of variable multiple times if each variable store’s different values. We then use the ticketIndex variable we created to obtain our results. (See code below)


           










![image](https://github.com/gotica462/Stock-Analysis/blob/main/VBA_Challenge_2017.png)
