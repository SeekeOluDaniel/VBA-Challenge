# VBA-Challenge
 Week 2 Challenge Assignment

#1 Inserted new column headers: Ticker, Yearly Change, Percent Change, Total Stock Volume on all sheets
#2 Output populating the new columns with corresponding values
I had help from a Learning Assistant to identify that my code was wrong, and I needed to make some corrections:
    Total = 0
    Row = 2
    Start = 2
    
    OpenPrice = ws.Cells(Start, 3).Value
    ClosePrice = ws.Cells(i, 6).Value

    YearlyChange = ClosePrice - OpenPrice        
    PercentChange = YearlyChange / OpenPrice
    YearlyChange = ws.Range("J" & Row).Value
                
    PercentChange = ws.Range("K" & Row).Value

    Start = i + 1

#3 Formatted the Yearly Change column to show 'red' for negative and 'green' for positive changes
    I learnt how to format the percent change column to show red for negative and green for positive  from my personal study using the resources provided: https://www.homeandlearn.org/excel_vba_practice1.html 

#4 Output the Greatest % Increase, Greatest % Decrease and Greatest Total Volume
    I reached out to a Learning Assistant when my code was not working. It turned out that I needed to type "Set" before Set Rng = ws.Range("K2:K" & lastRow) and Set PercentIncrease = ws.Range("Q2")

#5 To output the corresponding ticker symbol for #4
    I had a tutoring session to clarify what further steps I needed to do to ensure that all assignment requirements were met. I had to declare new variables for the summary table and reference the accurate cells using the new variables to return summary values and corresponding ticker symbols.
