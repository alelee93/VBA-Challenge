#VBA Refactoring Challenge

## Overview of Project
The main purpose of this project is to refactor VBA code to make it more efficient. More specifically, the refactored VBA code should achieve the same results output as the original code, have equal or better readability, and have a shorter running time. A secondary objective of this project is to analyze the data obtained through the VBA code. More specifically understand the stock performance between years for the provided dataset. 

## Results
In general, 2017 returns were higher than those of 2018.  The only exceptions were the RUN (+78% higher return) and TERP (2% higher return) stocks, which performed better in 2018 than in 2017. Something to note is that while TERP stock’s return increased, it is still negative in 2018 (it went from -7.2% to -5%). On the other hand, the RUN stock went from a 5% return in 2017 to an 81.0% return in 2018. Another stock to note is the ENPH, while the return of this stock decreased by 48% between 2017 and 2018, its 2018 returns are still positive at 81.9%.

The execution times of the refactored code is significantly faster than that of the original code. This is likely a result of using arrays in the code to store information.  

### Run times from refactored code
![VBA_Challenge_2018](https://user-images.githubusercontent.com/61717854/152666046-14cc4915-e0d9-4349-a3e4-3fdd69153470.PNG)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/61717854/152666068-3743cde4-a29a-49d1-8621-6b03906b3214.PNG)


## Summary
If done correctly, there should not be any disadvantages to refactoring code as it should be cleaner (more readable), more efficient (faster), without sacrificing results (archives the same outcomes as the original code). If done incorrectly, a possible disadvantage of refactoring code could be to sacrifice readability for efficiency (i.e. speed) or to inadvertently change the output of the code (the refactored code yields incorrect results).  

The run time and readability of the refactored code is better than that of the original code. Additionally, the output is the same as the one from the original code. As such, refactoring the original VBA script helped us have a cleaner and more efficient code (no disadvantages).  
