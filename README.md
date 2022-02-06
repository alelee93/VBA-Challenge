#VBA Refactoring Challenge

## Overview of Project
The main goal of this project is to refactor a VBA script to make it more efficient. The provided VBA script calculates and displays the total daily volumes and returns for stocks using  data from an Excel spreadsheet. The refactored VBA code should achieve the same results / yield the same output as the original code, but have better readability, and a shorter running time. A secondary objective of this project is to analyze the data obtained through the refactored VBA code to understand the performance of the stocks in the provided dataset. 

## Results

### Stock Analysis
In general, 2017 returns were higher than those of 2018.  The only exceptions were the RUN (+78% higher return) and TERP (2% higher return) stocks, which performed better in 2018 than in 2017. Something to note is that while TERP’s return increased, it is still negative in 2018 (it went from -7.2% to -5%). On the other hand, the RUN stock went from a 5.5% return in 2017 to an 84% return in 2018. Another stock to note is the ENPH, while the return of this stock decreased by 48% between 2017 and 2018, its 2018 returns are still positive at 81.9%

### Refactored Script Performance

The execution times of the refactored code are significantly faster than those of the original script. This is a result of using [arrays](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-arrays) in the code to store and retrieve information. More specifically, the original code used "Double" type variables to store the starting price, ending price, and volume of each ticker. Through a nested loop, the variables would be assigned the correct values for each ticker and then used to output the correct values in an Excel spreadsheet. While the results that the code yielded were correct, they came at a very costly "run-time" as the inner loop with if statement calculations had to run for each ticker. In the refactored code, we used arrays to store the starting price, ending price, and volume of each ticker. This allowed us to use a simple for loop to store the starting price, ending price, and volume of each ticker in the relevant arrays. To output the values, we simply created another for loop, retrieiving the data from the arrays.

### Original nested for loop

![original for loop](https://user-images.githubusercontent.com/61717854/152667132-4c9c3822-57fc-4d2c-84d3-e6bad7b1fba3.PNG)

### Refactored code uses arrays to store and retrieve data

![Refactored loop](https://user-images.githubusercontent.com/61717854/152667140-42cafb14-33ef-4a4d-be0f-219b3301655a.PNG)


### Run times from refactored code
![VBA_Challenge_2018](https://user-images.githubusercontent.com/61717854/152666046-14cc4915-e0d9-4349-a3e4-3fdd69153470.PNG)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/61717854/152666068-3743cde4-a29a-49d1-8621-6b03906b3214.PNG)

### Run times from original script

![original_runtime_2018](https://user-images.githubusercontent.com/61717854/152666509-a7922abe-f5a3-4e03-b51a-aa0a2fa12bdd.PNG)

![original_runtime_2017](https://user-images.githubusercontent.com/61717854/152666512-da6bc328-f34b-41f5-b3b5-9c71f3e18490.PNG)


## Summary
If done correctly, there should not be any disadvantages to refactoring code as it should be cleaner (more readable) and more efficient (faster), without sacrificing results (archives the same outcomes as the original code). If done incorrectly, a possible disadvantage of refactoring code could be to sacrifice readability for efficiency (i.e. faster but hard to follow) or to inadvertently change the logic of the code (the refactored code yields incorrect results).  

The run time and readability of the refactored code is better than that of the original script. Additionally, the output/result is the same as the one from the original code. As such, refactoring the original VBA script helped us have a cleaner and more efficient code (no disadvantages).  
