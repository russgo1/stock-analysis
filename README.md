# Stock Analysis with VBA

## Project Overview

The purpose of this analysis is to help a client visualize and compare data related to a set of publicly traded equities so that he can advise his parents as to which investments will both fit into their area of interest and offer the highest likelihood of a great return on their investment. This analysis goes a step further by creating automation that allows stock data to be quickly compiled based on year. 

## Results

### Stock Performance

Below are Total Daily Volumes and Return (calendar-year price performance) for the 12 stocks that Steve is examining for his parents for the years 2017 and 2018. 

<img width="265" alt="Returns_2017" src="https://user-images.githubusercontent.com/114126935/195967116-7ec8dcdd-495c-45b2-9bd2-da2c3d97bf68.png">     <img width="265" alt="Returns_2018" src="https://user-images.githubusercontent.com/114126935/195967126-bd26ccaf-94e2-4aa4-980b-9acb232917eb.png">

#### Return
Thanks to our bright red and green coloring, it is obvious that 2017 saw much better performance for these stocks than did 2018. The only stocks that did better in 2018 than in 2017 were RUN and TERP, with TERP still declining in 2018, just not as badly as in 2017. The only stocks to return positively for both years were RUN and ENPH. Two wins for RUN.

#### Volume
Steve’s parents also expressed that they believe trading volume to be an indicator of price-accuracy. Let’s take a look at our two double-year winners. 

##### ENPH
ENPH traded around 220 million shares in 2017, putting it around the middle of the group for trading volume, and it returned an incredible 129.5%. In 2018, ENPH had the highest trading volume of our group, around 600 million while maintaining an impressive 81.9% return. 

##### RUN
In 2017, RUN had a comparable trading volume to ENPH, with around 260 million shares, though it lagged in return (about 5%). In 2018, however, RUN was number 2 for trading volume (right behind ENPH) with around 500 million shared. RUN soared in 2018 to a return of 84%, the best of the group. 

### Conclusion

##### DQ
The favorite of Steve’s parents, DQ, had an astounding 2017 with a nearly 200% return, but that was with the lowest trading volume of the group, not topping a mere 36 million shares. In 2018, DQ’s trading volume increased to over 100 million shares, but its return was the worst of the group with a loss of about 62%. This did not erase the gains made in 2017, but both pieces of data together suggest that the 2017 price performance may have been overly jubilant. 

	
##### RUN and ENPH
These may be better suggestions for Steve’s parents, based off of the data that we have. They were the only two stocks that returned positive for both years, and they both saw a major increase in trading volume. As returns for the group overall were so high in 2017 and so low in 2018, it may be worth considering that 2017 saw too much enthusiasm placed in these equities  and that 2018 saw a correction to more reasonable prices.


### Program Performance
The original script ran the above analyses in 0.5625 seconds. 

<img width="260" alt="First_Runtime_2017" src="https://user-images.githubusercontent.com/114126935/195967192-9f8e9065-f35a-4864-81f3-346ed36db057.png">     <img width="260" alt="First_Runtime_2018" src="https://user-images.githubusercontent.com/114126935/195967198-d3f543e2-2348-4f09-ae9e-923fd0081684.png">

That’s pretty quick, but what if this were an analysis of 1,200 stocks instead of just 12? That might take an irritatingly long time. By refactoring the code, we were able to reduce the runtime to a mere 0.086 seconds (2017) and 0.082 seconds (2018). 

<img width="258" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/114126935/195967157-6170b320-9337-404d-9651-5ad866437240.png">     <img width="258" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/114126935/195967158-79fd9699-5e03-4bfc-8725-4f470012e262.png">

This is a major reduction that might make a difference if Steve ever wants to increase the scope of his analysis.

#### Additional Refactoring
The directions for refactoring for this challenge inspired me to do some additional refactoring. I liked the idea of using a `For` loop to assign values to an array. I decided to use this method to assign values of `0` to the `tickerStartingPrices` and `tickerEndingPrices` arrays, as well as the `tickerVolumes` array, as specified in the directions. This had little-to-no affect on run time. Despite this, I thought it looked nicer and made the code easier to understand, so I kept it. The following screenshots show before and after. 

#### Before
<img width="312" alt="Long_Arrays" src="https://user-images.githubusercontent.com/114126935/195967232-b02d8a43-3e20-4490-af73-48b0b13f2ba4.png">

#### After
<img width="465" alt="Short_Arrays_And_Loops" src="https://user-images.githubusercontent.com/114126935/195967238-05218163-1763-4c31-aebf-c3002738e3b8.png">


## Summary

### Pros and Cons of Refactoring
Refactoring can help to make code function better and look better. It helps to keep code clean, up-to-date, and understandable for developers who may work with it in the future. These are all positiver things. The downside is that refactoring can be time consuming. We may call to mind the old refrain, “If it’s not broken, don’t fix it.” It’s tempting to leave something be if it’s already working, but refactoring can be a good way to improve and maintain work that’s already been done. 

### Refactoring the Original VBA Script
By refactoring this script, we were able to make it run much faster, which may come in handy for larger datasets. We were also able to make it more clear and understandable with both better syntax and annotation that describes to any future reader (including ourselves) what is going on. This insight can be very helpful after weeks, months, or years.  
