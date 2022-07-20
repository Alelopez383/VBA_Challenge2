# **VBA_Challenge2**
We are refactoring a code we have done previously to loop through all the data one time in order to collect the same information as before but faster and trying to make lesser steps.


# Overview of the Project  
The purpose of the analysis performed is to help Steve´s parents to make the best investmente decision for the green stocks market, for that, we are going to use Visual Basic for Applications, or VBA, to enhance the analytical power of Excel. In particular, we seek to write an efficient script that allows us to automate the analysis of stocks across different years. 

We are facing a database that includes information on different green stock shares and their performance during the years 2017 and 2018. However, we are looking for an efficient and automated way to review the data and be able to analyze stocks to make optimal decisions. In this case, we want to help Steve's family improve their investments by telling them which stocks had a better return on investment and which ones had large losses, so it would be recommended not to invest in them.

# **Results:** 
Looking at the results, we can see that out of the 12 stocks in 2017, AY and TERP had a negative return. Instead, DQ, SEDG, ENPH and FSLR were the stocks with the best results, exceeding 100% of return. In particular, the shares of DQ and SEDG showed the best returns, 188.45 and 184.5% respectively. Therefore, Steve´s parente where in the right way to invest in DQ stocks in the year 2017.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/43974872/179879016-ec35c0ad-0315-458d-b7e3-adabe543023c.png)

Also, refactoring the code showed a small improvement in the code time, as the code initially took 1.041992 seconds to run the analysis for the year 2017. Refactoring the code, took 1.019531 seconds for the same year, so in the end, it was faster.


However, the analysis of 2018 shows that stocks lost mostly. In fact, only ENPH and RUN stocks were the only ones with positive results for 2018. Instead, DQ had one of the worst falls, it was the stock with the most losses compared to the others, it lost 62.6%.

![VBA_Challenge_2018](https://user-images.githubusercontent.com/43974872/179879027-853e96b7-f7ad-4db5-b032-d7bf3f9d5eb1.png)

In this case too, refactoring the code made faster the code. In the beginning without refactoring, the analysis of the year 2018 lasted 1.050049 seconds; now, the code ran in 1.007813 seconds for that year.

In general, the code and the refactoring code showed the same results, but not the same as the correct answer.

![VBA_2017 my results](https://user-images.githubusercontent.com/43974872/179880737-daf0464b-c62b-4c2e-ab6a-d2487a57142a.png)

There are slightly differences for the year 2017, althought, for the year 2018, are the same.

![VBA_2018 my results](https://user-images.githubusercontent.com/43974872/179880748-fb22f7c2-df6c-4c3f-b509-de6786941604.png)

# **Summary:** 

First of all, refactoring the code was a bit complicated since, although they were the same steps just extrapolating them, the syntax was a problem. The purpose was to repeat the information obtained previously, only with a much shorter and more efficient script. 

To avoid mistakes, I kept certain steps:

1. Always check which active sheet we want to see the values
2. Always define the type of variable that we are going to use
3. Always define in the command, which variable does what. Define what the variables are goint to do. Example: 'i' will be the indication for rows and 'j' will be indication for columns.
4. When using nested loops, the indication must always end with the second variable (inner loop) and at the end, the first (outer loop).
5. Have a 'clear worksheet' buttom is definitely a must, so we can check easily the results.
6. Research evrywhere for hints when the code doesnt work. Sometimes the hints doesnt work because we have differente versions (in my case, I assume) 

For example: the hint in canvas for the Step 3a didnt work for me, so I have to search and play a little bit with the syntax, so the code could run.

>To increase the volume of the current tickerVolumes by using the tickerIndex variable as the index, use the following code: tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value<

So instead I used: 
 tickerVolumes = tickerVolumes + Cells(j, 8).Value
 
 In the end, to refactoring the code could be simple, if we have very clearly what are we doing when we create variables and run the code.
 
 Maybe I have some mistakes when running the code, because the results are not quite identical, but there are the same results before refactoring the code.
 
 ![2018 before refactor code](https://user-images.githubusercontent.com/43974872/179884589-80ff692e-8038-4dab-a91a-6cde337af5fa.png)

   
 In the end, Steve´s parent should get ENPH stocks, becuase they presented the best returns for both years. In 2017, the stock had a return of 129.5%, and in the year 2018, was of 81.9%.
            
 
            
