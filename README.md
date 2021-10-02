# Stock Analysis (VBA Challenge) 

## Overview of Project

- Purpose of this challenge was to help Steve determine which stocks were performing well. The challenge was to create a code through VBA that would calculate the results for us quicker than performing it manually. 

## Results

- In general, stocks in ![2017](https://github.com/tutran90/stock_analysis/blob/main/2017_Original_VBA_Code_TIme.png)    had a better return compared to stocks in ![2018](https://github.com/tutran90/stock_analysis/blob/main/2018_Original_VBA_Code_Time.png). 

- The purpose of this project was to refactor the code to see if it would analyze through a set of worksheets quicker. The time for the original code for both the years ![2017](https://github.com/tutran90/stock_analysis/blob/main/2017_Original_VBA_Code_TIme.png) and ![2018](https://github.com/tutran90/stock_analysis/blob/main/2018_Original_VBA_Code_Time.png) resulted in a negative time. While, the refactored time for ![2017](https://github.com/tutran90/stock_analysis/blob/main/2017_Refactored.png) and for ![2018](https://github.com/tutran90/stock_analysis/blob/main/2018_Refactored.png)resulted in milliseconds. This means that the original code was running longer, which could be a result of where the "startTime" and "endTime" was placed in the code. In the original code it was last macro of Module 1. The results would populate based on what was selected and the call function was used to pull results from each year (which was created on a different macro on the same module). In other terms, this means the clock was already running before the actual code ran the conditions. Itook longer since it had to "call" on other results. Where as the refactored code did not "call" on to different macros in a module. It was all included in 1 subroutine. 

- In summary, both codes came up with the same results. The refactored code ran quicker than the original code.

    

## Summary 

**A. Refactoring A Code in General**

1. Advantages

   - The only advantage is that you are not having to work from scratch. 

2. Disadvantages

    - Style: When having to refactor a given code, one of the disadvantages is working through the creators style. I.e one may be writing the code in VBA and another in VS Code. There may be indentations and maybe not. This can cause confusion when trying to loop through the codes and finding where it ends. 
    - Clear Comments. This is also dependent on the creators style. If the comments are short and not descriptive it will be hard to figure out what they were trying to do. 

**B. Refactoring VBA Script v. Original** 

1. Advantages

    - The only advantage (in my opinion) to refactoring the VBA Script, was that it was similar and went a long with what was discussed during the module. 

2. Disadvantages

    - Debugging the errors and finding the solution. It could have been as simple as case sensitive or forgetting a "next i", which I now see why building and editing in VS Code is so valuable. 
    - As I have learned, mapping out what you are trying to create first will help with building the code. 
   








