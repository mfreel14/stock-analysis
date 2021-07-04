# Stock Analysis

## Overview of Project

Steve wants to expand the dataset to include the entire stock market over the last few years. Although our code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

Through refactoring the code to loop through all the data one time, we'll determine whether refactoring the code successfully made the VBA script run faster. 

## Results

When comparing stock performance between 2017 and 2018, we can see those overall stocks for 2017 performed better.  

For 2017, 92% of the stocks had a positive return with highs of 199.4% for DQ and 184.5% for SEDG.  The lows for 2017 were TERP (-7.2%) and RUN (5.5%).  The average return across all stocks was 67.32%.

![2017 Performance](https://user-images.githubusercontent.com/691355/124395628-f99fec00-dcb9-11eb-8ddd-daf2b229e358.png)

Comparatively for 2018, 16% of stocks had a positive return.  The high performing stocks were ENPH (81.9%) and RUN (84%).  The lowest performing stocks were DQ (-81.9%) and JKS (-60.5%).  The average return for 2018 was -8.51%.
 
![2018 Performance](https://user-images.githubusercontent.com/691355/124395902-9c0c9f00-dcbb-11eb-9caa-4c861968abfc.png)

For this project a comparison was made between our original macro and a refactored macro.  

The runtime of our original macro was .765625 seconds for 2017 and .7734375 seconds for 2018.

<img width="252" alt="Analysis - 2017" src="https://user-images.githubusercontent.com/691355/124395992-01609000-dcbc-11eb-8945-b41bd141636e.png">
<img width="248" alt="Analysis - 2018" src="https://user-images.githubusercontent.com/691355/124395995-032a5380-dcbc-11eb-8887-3b7739f3482f.png">

Comparatively our refactored macro ran 2017 at .1640625 seconds and .1640625 seconds for 2018.


<img width="255" alt="Refactored - 2017" src="https://user-images.githubusercontent.com/691355/124396029-2f45d480-dcbc-11eb-8a27-d5deeda594c6.png">
<img width="256" alt="Refactored - 2018" src="https://user-images.githubusercontent.com/691355/124396031-30770180-dcbc-11eb-961f-8d66b91f5a00.png">

We can see that running the macro in the refactored version gives us faster results.  

To achieve this, we modified our original code and added the following to our script.

![Refactored Code](https://user-images.githubusercontent.com/691355/124396129-ada27680-dcbc-11eb-9d0e-1e23109fb884.png)
![Refactored Code 2](https://user-images.githubusercontent.com/691355/124396152-c9a61800-dcbc-11eb-8621-f5c86a14ca36.png)

What are the advantages or disadvantages of refactoring code?

From this project we can visually see the benefit to running refactored code.  The added code made our macro faster with a reduction in execution time of .6015625 seconds for 2017 and .609375 for 2018.  The code, with added notes is also easier to understand and can make future updates easier.  

The disadvantage of refactoring is the amount of time it takes to repurpose the code to run faster.  Since the code may be someone else's work there can be situations where it would be easier and less time consuming to start the project from scratch.  

How do these pros and cons apply to refactoring the original VBA script?

For this project we were able to see an immediate benefit from the faster run time of our macro.  However, the amount of time spent to add the code may be looked at as a disadvantage.  Although we were able to produce faster run times, the amount of time we save may not be worth the amount of time spent.  If Steve were paying us to refactor this code our time spent may not be worth the amount of time saved.  The situation may be different if the original macro ran for 10 mins, and we were able to reduce that to 1 min.  The reduction in time for this is far greater than the .6015625 seconds we're saving between our original and refactored code.
