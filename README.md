# VBA-Challenge
 Week 2 Challenge Assignment
Started Aug 27th Due Date: Aug 31st
#1 Inserted new column headers: Ticker, Yearly Change, Percent Change, Total Stock Volume on all sheets
#2 Output populating the new columns with corresponding values
I had help from a Learning Assistant to identify that my code was wrong, and I needed to make some corrections:
    Total = 0
    Row = 2
    Start = 2
    
    OpenPrice = ws.Cells(Start, 3).Value
    ClosePrice = ws.Cells(i, 6).Value

    Start = i + 1

#3 Formatted the Yearly Change column to show 'red' for negative and 'green' for positive changes
    I learnt how to format the percent change column to show red for negative and green for positive  from my personal study using the resources provided: https://www.homeandlearn.org/excel_vba_practice1.html 

#4 Output the Greatest % Increase, Greatest % Decrease and Greatest Total Volume
