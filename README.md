# Estimating Changes in Stock Volume for 2018 - 2020 using VBA Script

This section describes briefly various sections of the VBA script attached as part of this changelle. ***Comments*** are also provided in the code for better understanding.

## Declaration of variables and assigning values to the variables
    
![image](https://github.com/pkrachakonda/VBA-challenge/assets/20739237/7b7e9db9-76dd-4ff8-88a2-58a922c388be)
    

In the code, Lines 4 - 8 defines the nature of variables used in the code. Line 10 ask the code to loop each worksheet in the spreadsheet (excel) i.e. to loop through 2018, 2019, 2020 and to perform the below code.
Lines 13 and 14 assign an initial values to the two variables. Values Open_Price and Last row for each worksheet are obtained through lines 16 and 17 of the code respectively.

## Assigning Cell Names
Lines 21 to Lines 29 in the code assign names to various cells as ***Column Headers*** as listed in the challenge. It also helps in sorting the data.

![image](https://github.com/pkrachakonda/VBA-challenge/assets/20739237/e4d229b5-1edf-4c95-8980-7d8c597d0de9)

## Extracting Tinker symbol and estimation of yearly changes

![image](https://github.com/pkrachakonda/VBA-challenge/assets/20739237/e2e0dc90-61ab-4ed1-ae83-dd309daea313)

### Extraction Tinker(Stock) symbol
The first line of the code after the For loop, check whether the Tinker(stock) name in the current row is same as in the next row, if it is same it will sum the stock volume and store that value to *"Stock_Volume"* variable and then it moves to next row. If there is a name mismatch, the Tinker symbol (locatd in Column A), total stock volume till that row and close price of the tinker are copied to *respective variables* as shown in the code.
Based on these *respective variables*, yearly change = (Open_Price - Close_Price) and percent change = (yearly change/open_price) are estimated and entered to there respective columns (Lines 39 and 40), as well as total stock volume (line 41).

### Conditional Formatting 
Line 44 checks whether the yearly and percent change values are positive or negative. If the values are postive, the cells are formatted with *Green* colour otherwise with *Red*.  

Once the Tinker symbol, Yearly and Percent Change values are recorded in Columns J, K ,and L; codes assign a new values to variables: *Open_Price*, *Stock_Volumes* *Summary_Row* (Lines 52, 53, 55)

## Assigning Number Format to Columns

![image](https://github.com/pkrachakonda/VBA-challenge/assets/20739237/9f9a377c-3b15-4e69-a5cd-8aa219323468)

Lines 65 - 69 assign numbering style/format to various columns in each worksheet of the spreadsheet. For each year highest and lowest percent yearly increase as well as greatest total stock volume are estimated using vba builtin functions (lines 62 -64) and are stored in Column Q2 to Q4 in each worksheet.

![image](https://github.com/pkrachakonda/VBA-challenge/assets/20739237/2673425f-4958-4ef4-bb90-04c74f8364ee)

Lines 72 - 80 loops over the values stored in Columns K (Percent Change) and L (Total Stock Volume) and checks whether the values match with those stored in Column Q2: Q4. If values match, corresponding Tinker value stored in the Column I is copied to their respective row in Column in P. After extracting all the values the worksheet is autofit to properly display the values in the columns and rows in each worksheet.



