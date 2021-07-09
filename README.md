# stock-analysis

# VBA of Wall Street Unit 2 Challenge Project
## Student:  Christopher Mastrangelo

Please refer to the following files in the repository "<a href="https://github.com/gcmastra/stock-analysis.git">gcmastra/stock-analysis</a>"

<a href="https://github.com/gcmastra/stock-analysis/blob/6c6f8a946b4440422f29e8ca3c13e9816947c9d8/VBA_challenge.vbs">VBA_challenge.vbs</a><br>
<a href="https://github.com/gcmastra/stock-analysis/blob/f6e6bfc0f5ae14809889e4ad29c53a8694e8bc14/VBA_challenge.xlsm">VBA_challenge.xlsm</a><br>

Graphics that show the final versions of the data tables

<a href="https://github.com/gcmastra/stock-analysis/blob/main/resources/VBA_challenge_2017_extra.PNG">VBA_challenge_2017_extra.PNG</a><br> 

Graphics in /resource folder that show the elapsed time to run the original code and refactored code for both years (2017, 2018)

<a href="https://github.com/gcmastra/stock-analysis/">file_1</a><br>
<a href="https://github.com/gcmastra/stock-analysis/">file_2</a><br>
<a href="https://github.com/gcmastra/stock-analysis/">file_3</a><br>
<a href="https://github.com/gcmastra/stock-analysis/">file_4</a><br>


# Overview of Project

1. Given a spreadsheet containing stock market ticker data for 12 stocks, use VBA code to summarize the data using conditional formatting.
2. Solve the problem for a single stock, then generalize to a list of 12 stocks, making 12 separate loops through the full data table.
3. Finally, optimize (refactor) the VBA code to run more efficiently by using a single loop and storing data in arrays with a row for each stock ticker.
4. Demonstrate that the data table is the same using either the nested loop approach or the single loop approach
5. Demonstrate the code ran in less time after the refactoring.

# Technical Challenges
1. Write pseudocode using comments first and make sure to write connents that are descriptive enough to follow the logic flow 

For example [image]

![image](https://user-images.githubusercontent.com/86205000/125141518-3ec18500-e0e3-11eb-8aae-75c3a080631a.png)

3. Create buttons in Excel that run subroutines to perform various actions.  This makes the finished spreadsheet more user friendly. This is also useful while designing and debugging to make it easier to run the code multiple times 

# Extra Credit?
in Module 1 of the VBA code (before the refactroing exercise), I wrote a subroutine "Sub BuildTickerColumn(year$)" which returns the list of ticker values for a specific year based on the data in the spreadsheet rather than hard coding the values.  [See Image]

![image](https://user-images.githubusercontent.com/86205000/125141894-459cc780-e0e4-11eb-9f53-91dc7af84a40.png)

Also I added an option to display the start/end prices to vaidate that the percentages were correct

![image](https://user-images.githubusercontent.com/86205000/125142903-eab89f80-e0e6-11eb-824e-f9e151006e68.png)



# Results - the timing before and after refactoring

Before 2017 + 2018 . . .

![image](https://user-images.githubusercontent.com/86205000/125142324-75989a80-e0e5-11eb-93d6-51c9cc9fa1b6.png)

After refactoring . . . 

![image](https://user-images.githubusercontent.com/86205000/125142627-31f26080-e0e6-11eb-88df-77d59f1f56a4.png)

## Thank you for reading
