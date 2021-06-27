
# stock-analysis
# VBA of Wall Street: Challenge 2 by Kieran Persaud
## Overview of Project

Steve wanted to prepare a macro-enabled workbook that analyzes the performance of green energy stocks for 2 years worth of data. The VBA code originally written provides this information with a click of a button. The original code provided a formatted table of volume and returns for 12 stocks based on the year the user entered. 
### Purpose

Steve likes the original workbook and code that was provided, but worries that it may not be able handle calculations for thousands of stocks in a timely manner. Thus, the purpose of this project is to refactor the code to make it run faster.
## Results

### Stock Performances

The output of both Macros show that Green Energy stocks performed significantly better in 2017 than in 2018. All stocks except ENPH and RUN saw losses in 2018. Steve might suggest purchasing those stocks, as they saw gains in both years of his analysis.

### Comparison of Outputs from Original and Refactored Macros

The pictures below show that both Macros outputted the same results. The workbook has a _Clear Worksheet_ Macro that was used to clear the worksheet and allow both codes to run fully.

**2017 Original**

<img src="https://user-images.githubusercontent.com/84286467/123561395-19915600-d776-11eb-939d-7473b3e2b628.PNG" width="675" height="300">

**2017 Refactored**

<img src="https://user-images.githubusercontent.com/84286467/123561415-2ca42600-d776-11eb-90df-0fc1fe55165b.PNG" width="675" height="300">

**2018 Original**

<img src="https://user-images.githubusercontent.com/84286467/123561430-45acd700-d776-11eb-8387-5aaabba5401b.PNG" width="675" height="300">

**2018 Refactored**

<img src="https://user-images.githubusercontent.com/84286467/123561435-50676c00-d776-11eb-97c7-8aabe1e17247.PNG" width="675" height="300">

### Comparison of Run Times 

In both cases, the Refactored Macro ran faster than the original Macro. This was due to utilization of arrays to store the calculations of Starting Price, Ending Price, and Volume, and the use of _tickerIndex_ to call those results quickly. The Refactored Macro ran 2017 Data 0.14 seconds faster and ran 2018 Data 0.16 seconds faster.

#### 2017 Comparison

<table>
  <tr>
    <td>Original</td>
     <td>Refactored</td>
  </tr>
  <tr>
    <td><img src="https://user-images.githubusercontent.com/84286467/123562172-ef8e6280-d77a-11eb-9840-dcca4c5d69f8.PNG" width=500 height= 250></td>
    <td><img src="https://user-images.githubusercontent.com/84286467/123562190-19478980-d77b-11eb-846a-2420bf3cc2e1.PNG" width=500 height=250></td>
  </tr>
 </table>

#### 2018 Comparison

<table>
  <tr>
    <td>Original</td>
     <td>Refactored</td>
  </tr>
  <tr>
    <td><img src="https://user-images.githubusercontent.com/84286467/123562249-8529f200-d77b-11eb-8090-5531f29c786e.PNG" width=500 height= 250></td>
    <td><img src="https://user-images.githubusercontent.com/84286467/123562262-9672fe80-d77b-11eb-996a-bda317922637.PNG" width=500 height=250></td>
  </tr>
 </table>

## Summary
### Advantages and Disadvantages of Refactoring Code
Advantages of refactoring code include making it more robust and easier to maintain. It shortens long methods, eliminates duplicate code, decreases the run time, and allows inheritors of the code to parse it better.
The main disadvantage of code refactoring is the time it takes to refactor. This may not be desired if a project is under tight time constraints, and the project manager simply wants code that works. Another disadvantage is that bugs may be introduced to the code, and methods may have to be changed completely to fix those bugs.

### Advantages and Disadvantages of the Original and Refactored VBA script
The main advantage of refactoring this particular VBA script was a cutdown on processing time. In theory, this will make the code more robust and able to handle larger amounts of data if Steve were to pull historical data of those same 12 stocks. Again, the time it took to refactor the code was a disadvantage. My introduction of new arrays and variables resulted in several bugs, and I took the time to sort out those issues. A disadvantage of both codes though is that it requires manual changes of the tickers if there are new stocks Steve wants to analyze.
