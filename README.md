# An Analysis of a Refactored VBA Script for Stock Analysis

## Overview of Project

### Purpose

The purpose of this analysis was to take an original VBA script that was used to analyze select stock performances in 2017 & 2018 and successfully refactor the VBA code to make the VBA script run faster.

## Results

### Comparison of Stock Performance between 2017 and 2018
  
2017 Stock Performance            |  2018 Stock Performance
:-------------------------:|:-------------------------:
<img src="https://user-images.githubusercontent.com/92001105/139770940-9dd76976-7c3e-4df8-a015-0f3a3219c548.png" width="150%"></img>   | <img src="https://user-images.githubusercontent.com/92001105/139771111-d0bb4d19-b558-4a7b-bc1f-a8717c7dc50b.png" width="150%"></img> 


In terms of the stock performance that is being shown above in 2017 almost all of the stocks had a positive return, measuring the ending stock price and comparing the growth compared to the starting stock price. The only stock in 2017 that did not yield a positive return was TERP. However, in 2018 this was almost completely opposite will almost all the stocks returning a negative return for the year, with only ENPH and RUN yielded positive returns.

### Comparison of Execution Times between Original and Refactored Scripts

#### *2017 Comparison*

Original Script            |  Refactored Script
:-------------------------:|:-------------------------:
![](https://user-images.githubusercontent.com/92001105/139763231-4aa6d6cd-ffcb-47c0-ab8d-aa39bec6cb2b.png)  |  ![](https://user-images.githubusercontent.com/92001105/139763314-535e684c-7ace-45ac-9ae7-2144c9cfb6e2.png)

When running the scripts on the 2017 data when the original script was run, which had nested loops that significantly increased the time to run, the time was just over 0.89 seconds. After the script was refactored and an index was created with arrays instead of running through multiple loops the time to run decreased to 0.156 seconds - an 80%+ improvement!


#### *2018 Comparison*

Original Script            |  Refactored Script
:-------------------------:|:-------------------------:
![](https://user-images.githubusercontent.com/92001105/139771297-0422f1e1-1847-4211-8148-8f2a350e0ba5.png)  |  ![](https://user-images.githubusercontent.com/92001105/139771324-c1a17cdd-734c-4d62-8b6b-6d548c37af47.png)

When running the scripts on the 2018 data the original script ran in 1.0 seconds. After the script was refactored the time dropped to 0.156 seconds - an 84% improvement!



## Summary

### Advantages and Disadvantages of Refactoring Code

#### *Advantages*

1. Does not require you to start from scratch and can save time by reusing existing code and logic.

2. Can utilize named variables and recreate into arrays and can save time on writing the script.

#### *Disadvantages*

1. Can sometimes be time consuming to have to change the code and reword compared to just starting fresh in some sections.

2. Need to watch out for syntax errors and ensure all changes are complete, which can become inefficient at times

### How did these Pros and Cons apply to Refactoring the Original VBA script

I found that it was quite efficient to utilize a lot of the original script and just make specific edits throughout. The downside was I did have some syntax errors and was hard to catch specific changes to the code namely only having one loop compared to a nested loop.
