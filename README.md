# Stock Analysis
We will refactor VBA code to loop through Stock results more efficiently
## Overview of Project
Steve had asked for our help to analyze different stocks so that his parents can make the best decision about what stock to invest in.
### Purpose
The purpose is to refactor the VBA code so that we can loop through only once and having a more efficient run time.
## Results
To begin refactoring the code we created output arrays for each of our variables. The goal is to make our analysis more efficient and by reducing the number of times the code must loop through different variables, we are able to do so. We kept the array from the original code for tickers. We created output arrays tickerVolumes(), tickerStartingPrices(), and tickerEndingPrices().  We created the variable tickerIndex to refer to the tickers in our initial array. We set this variable equal to zero and prepared to run our loop through our datasheet <br>
<br>
![Create Output Arrays](Resources/screenshot1.png)
<br>
To run analysis through the spreadsheet we created a for loop to run through the dataset, just like the original code. We first created a ticker to increase volume as the code ran through each cell. We set each of our tickerVolumes(i) equal to zero in a previous For Loop. In our current For loop we ran through each cell and added each volume until we reached the end of the ticker by referencing tickerIndex as seen in the code below. <br>
<br>
![Find Total Volumes](Resources/screenshot2.png)
<br>
We then needed to find our ending prices and starting prices. Instead of creating a nested For loop and running a new loop for each ticker to find volume and then looping through rows to find starting/ending prices , we can use our variable tickerIndex to loop through more efficiently. <br>
<br>
![Find Starting and Ending Prices](Resources/screenshot3.png)
<br>
Lastly we created a new datasheet and were able to display the results for each variable using our arrays for each ticker. . <br>
<br>
![Display Refactored Results](Resources/screenshot4.png)
<br>

### Advantages and Disadvantages of Refactoring Code
Refactoring can have great advantages because it allows the collaborative nature of coding to bloom. New eyes can bring a new take and add different ways of coding. There is always room for improvement and this allows coders a chance to improve their code. It each code to be made more effective and improve readability. 
Refactoring code can be disadvantageous if there is already broken code or if the code is built around a code that is not as effective. It might be more beneficial for some code to start fresh.
### Advantages and Disadvantages of Refactored VBA script
Some advantages of refactored VBA script is we can see the efficiency of our refactoring by comparing coding time. As seen below with our refactored code, we have significantly improved out code performance. The original codes ran for .496 seconds for 2017 and .484 seconds in 2018. <br>
![VBA Challenge 2017](Resources/VBA_Challenge_2017.png)
<br>
![VBA Challenge 2018](Resources/VBA_Challenge_2018.png)
<br>
 A disadvantage is that is can be difficult to test our refactored code if there is broken code. We cannot see the efficiency if everything is dependent on each other and we are unable to rerun our code without having a solution.