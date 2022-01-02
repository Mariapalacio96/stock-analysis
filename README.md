# stock-analysis

## Overview of the project
In this analysis, we had as a goal to help Steve to analyze wich is the best way his parents could make their investment. We were given a dataset with different stocks. In this dataset we were given the information necessary to be able to analyse these stocks. Steve's parents had though about one specific stock that they would like to invest in, DQ, we will help steve to find out if this would be a good investment or if there is a better way that we can diversify their portfolio.

This analys will get done by using VBA where we will write a code to analyse these stocks. At the begining we used different subroutines to make the code but after we are done with the analysis, we are going to refactor the code to make it more simple and short but it will be showing the same results. By refatoring the code we will see if it actually makes the code run faster or not. 

### DQ'S analysis
Using the information we were given at the begining to see how DQ performed in the year 2018. As our first step we will loop through all the stocks to calculate the total daily volume at wich the stock was traded. The daily volume will teel us how much the stock was actively traded in this specific year. After this, our next stop was to find the total return that the stock had wich means how much the price increased or decreased from the begining of the year to the end. Having this two values, we can now say that there may be better investments in this group of stocks since its yearly return is negative wich means that this stock drop more than 50% of its price.

<img width="627" alt="Screen Shot 2022-01-02 at 5 36 23 PM" src="https://user-images.githubusercontent.com/95391094/147891230-e3df9641-d205-4543-9a44-e17a694ef378.png">

<img width="424" alt="Screen Shot 2022-01-02 at 4 49 39 PM" src="https://user-images.githubusercontent.com/95391094/147890707-b2e6c5c7-dc1b-4438-bcea-a6348887816a.png">

### Analysis of other stocks
After we already made an analysis in DQ, we can use part of this code to analyse some more stocks. We will be based in 12 stocks for this part of the analysis, we will be find the same values as we did for DQ. Since this time we will be analysing more than one stock, in order to find these values we will need to create an array that will hold each stock.

<img width="288" alt="Screen Shot 2022-01-02 at 5 24 52 PM" src="https://user-images.githubusercontent.com/95391094/147891012-8f678bee-b35c-4b7c-bff5-06fae64c8691.png">

After we wrote the code, we can make the results look more readable by putting ccolors on the yearly return. We will use red for those negative returns and green to show those positive return. 

![image](https://user-images.githubusercontent.com/95391094/147891128-bed0c7c6-42f9-4cfd-af54-dff2ff8c47f7.png)

### Analysis for each year
After having the code to analyse yhe year 2018, we can modify the code to be able to analyse any year of the worksheet 

<img width="881" alt="Screen Shot 2022-01-02 at 5 43 00 PM" src="https://user-images.githubusercontent.com/95391094/147891367-426520e6-f56d-4300-8fe6-67dfbe147c63.png">

## Results

### Results on stocks 
According to our analysis, we can conclude that the best choice for steve's parents is to invest in ENPH and RUN wich returns aree of more than 80%. I would void investing in DQ because of how much its price has droped in 2018. As we are helping Steve to make this analysis, we also made a buttom to make it easier for Steve to clean the analysis and to run it again.

![image](https://user-images.githubusercontent.com/95391094/147891586-6e94c313-feaf-4852-82b8-9ecd5fa74e26.png)

### Results in refactoring the code
At begining of this project, we started developing the code by creating different subroutines wich i think were so helpful to understand more each step we were taking though writting the code. Even though it was a lot of help to have the code in different subroutine, it made a few of the pieces in the code so repetitive wich make the code to run slower.

![image](https://user-images.githubusercontent.com/95391094/147891786-251f6a61-1908-4a4a-a597-8b7830af3860.png)


![image](https://user-images.githubusercontent.com/95391094/147891772-7d1ffd3c-e6d0-4f3f-a8f0-bb74d5f5c0fe.png)

After refactoring the code, we can see how short the code can actually be since we are not repeatinf a few steps. Another pro of refactoring is how faster the code runs after we edited.

![image](https://user-images.githubusercontent.com/95391094/147891838-1648fcc0-be24-43e9-ba38-7e7982509a4f.png)


![image](https://user-images.githubusercontent.com/95391094/147891828-445f7679-4ae0-4e6d-9943-ad737dbaa3ae.png)



