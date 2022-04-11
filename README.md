# stocks-analysis

## Project Overview
A client requested to see a refactored analysis of green energy stocks using VBA. The purpose of the refactoring the VBA code is to improve execution time for future datasets that could be significantly larger. 

## Resources
- Software: Microsoft Excel v2203, Microsoft Visual Basic for Applications 7.1

## Results
![alt text](https://github.com/thehatch4815162342/stocks-analysis/blob/main/Images/all_stocks_refactored_2017.png?raw=true) ![alt text](https://github.com/thehatch4815162342/stocks-analysis/blob/main/Images/all_stocks_refactored_2018.png?raw=true)
Here is the stock analysis for both before and after the refactoring of the code. 
A few takeaways here include:
-2017 saw a better return for nearly all stocks besides TERP (with a -7.2% return) in comparison to 2018
-DQ saw the highest return in 2017 at 199.4%, however in 2018 it saw the lowest return at -62.6%
-Only ENPH (81.9%) and RUN (84.0%) saw positive return in 2018
 
### Execution Time Before
 ![alt text](https://github.com/thehatch4815162342/stocks-analysis/blob/main/Resources/original_2017_runtime.png?raw=true)![alt text](https://github.com/thehatch4815162342/stocks-analysis/blob/main/Resources/original_2018_runtime.png.png?raw=true)
 
### Execution Time After
 ![alt text](https://github.com/thehatch4815162342/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=true) ![alt text](https://github.com/thehatch4815162342/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.png?raw=true)
 
As we can see above, refactoring the code has speed up execution time, but in this case the time differences are not drastically different. However, with bigger datasets, I believe there could be an actual significant difference.
 
## Summary
It appears, in this case, refactoring the code was beneficial. The code is now faster and more efficient just like the client desired. However, I can see one disadvantage of refactoring code is that it can be time consuming.
