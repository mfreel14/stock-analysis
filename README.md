# Stock Analysis

## Overview of Project

Steve wants to expand the dataset to include the entire stock market over the last few years. Although our code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

Through refactoring the code to loop through all the data one time, we'll determine whether refactoring the code successfully made the VBA script run faster. 

## Results
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

When comparing stock performance between 2017 and 2018, we can see that overall stocks for 2017 performed better.  

For 2017, 92% of the stocks had a positive return with highs of 199.4% for DQ and 184.5% for SEDG.  The lows for 2017 were TERP (-7.2%) and RUN (5.5%).  The average return across all stocks was 67.32%.

![2017 Performance](https://user-images.githubusercontent.com/691355/124395628-f99fec00-dcb9-11eb-8ddd-daf2b229e358.png)

Comparatively for 2018, 16% of stocks had a positive return.  The high performing stocks were ENPH (81.9%) and RUN(84%).  The lowest performing sotcks were DQ (-81.9%) and JKS (-60.5%).  The average return for 2018 was -8.51%.  

![2018 Performance](https://user-images.githubusercontent.com/691355/124395902-9c0c9f00-dcbb-11eb-9caa-4c861968abfc.png)

For this project a comparison was made between our original script and a refactored script.  

THe runtime of our original script was .765625 seconds for 2017 and .7734375 seconds for 2018.

<img width="252" alt="Analysis - 2017" src="https://user-images.githubusercontent.com/691355/124395992-01609000-dcbc-11eb-8945-b41bd141636e.png">
<img width="248" alt="Analysis - 2018" src="https://user-images.githubusercontent.com/691355/124395995-032a5380-dcbc-11eb-8887-3b7739f3482f.png">

Comparatively our refactored script ran 2017 at .1640625 seconds and .1640625 seconds for 2018.

<img width="255" alt="Refactored - 2017" src="https://user-images.githubusercontent.com/691355/124396029-2f45d480-dcbc-11eb-8a27-d5deeda594c6.png">
<img width="256" alt="Refactored - 2018" src="https://user-images.githubusercontent.com/691355/124396031-30770180-dcbc-11eb-961f-8d66b91f5a00.png">



Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
