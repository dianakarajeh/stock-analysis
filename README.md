# Stock-Analysis
Perform a refactor of a specific VBA code to collect specific stock information from the years 2017 and 2018 in order to decide if any of the stocks are worth investing.
## Project Overview
#### Steve has taken on his first real gig post graduation, and he has decided to help his parents out with investing in new stock- specifcally green energy stock.  Steve's parents had previously notified him that even though they did not do much research, that they will be invested ALL their money into DAQO New Energy Corp.  Before making a final decision, Steve took it upon himself to research and analyze thousands of other green energy stocks to present to his parents in order to make the best investment.  A thorough analysis of all the green energy stocks from 2017-2018 were done using VBA in order to determine which stocks would be the best investments. 
## Results
#### Prior to refactoring the code, the initial code was manually done by creating mutliple different for loops to incorporate the necessary data from 2017 and 2018 that Steve pulled.  Ticker, Total Daily Volume, and Return were chosen as categories to analyze most efficiently to get the answer we needed for Steve's parents. An array of all tickers (12) was initialized before beginning to loop all the data that was needed, as well as three different output arrays: tickerVolumes(12) As Long, tickerStartingPrices(12) As Single, and tickerEndingPrices(12) As Single.
#### The first loop created was _For i = 0 to 11_, and this was done to initialize our three output arrays as 0.  All rows were looped over in order to have the code apply to all sets of data.  Below is a snippet of the final code and how all the for loops added were able to efficiently analyze all the data that was given:
<img width="740" alt="Screen Shot 2022-03-06 at 6 52 12 PM" src="https://user-images.githubusercontent.com/99656224/156947756-7c5e4cea-7fa3-492f-9a55-5297a600f785.png">

#### To get the data from the "All Stocks Analysis" worksheet that Steve compiled, I had to activate the sheet and then loop through all the arrays to output the Ticker, Daily Volume, and Return. Below is how I was able to conditionally format the data to change colors if the Return was negative (red) or positive (green) for both years:
<img width="683" alt="Screen Shot 2022-03-06 at 6 54 35 PM" src="https://user-images.githubusercontent.com/99656224/156947908-3ff52697-7598-42ec-a03b-97b7e60795d6.png">

#### Regardless of refactoring or not, both methods used should have produced the same table.  
<img width="337" alt="Screen Shot 2022-03-06 at 5 23 58 PM" src="https://user-images.githubusercontent.com/99656224/156947971-689c9a0b-4276-4332-b8e4-eda9fb4b1a63.png">
<img width="338" alt="Screen Shot 2022-03-06 at 5 24 08 PM" src="https://user-images.githubusercontent.com/99656224/156947986-213d9a90-21cd-4579-bd3c-0d49fc3160d2.png">

#### As observed from both figures above, 2017 and 2018 both had extremely different return years than DAQO had.
### Before and After Refactoring
#### Before refactoring, the original script displayed different execution times for both years.
<img width="256" alt="Original script 2017" src="https://user-images.githubusercontent.com/99656224/156948476-df50b96d-27fa-4771-9fd3-9df484473044.png">
<img width="256" alt="Original Script 2018" src="https://user-images.githubusercontent.com/99656224/156948480-e59dd1c3-b70b-4e21-b359-73e0c9a62771.png">

#### Although these execution times seem great, they improve after refactoring of the original script is done. 
### Summary
### Pros and Cons of Refactoring Code
#### The pros of refactoring code are that it simply makes everything much more organzied and neat.  As seen above with the difference between execution times, refactoring allows for faster programming to occur as well as an improvement in readability.  Readability becomes very important when customers such as Steve's parents have no prior experience with analyzing stocks. A con of refactoring code is that it could get confusing and difficult if the correct methods are not used.  
### Pros and Cons of Refactoring Original VBA Script?
#### The most important thing that was gained by refactoring the original VBA Script was much faster macro run time.  In order to analyze the data most efficiently, run time is very important.  The original run times were around ~3.04 seconds for both years while the refactored run times were around ~0.08 seconds for both years.
