# VBA Refactoring Analysis

## Overview

I have a friend named Steve. And while I've already built a button-loaded worksheet for my friend to analyze the performance of green energy stocks during 2018-2019, still he wants more. Steve wants to analyze the performance of all the stocks from 2018-2019. The goal is to refactor the code so that the script runs faster and can handle an expanded dataset. Secondary goal: convince Steve to compensate me for all this work--I mean how close are we? 

### Results

To speed up the code, I created a three utput arrays to loop through istead of relying on referencing the data sheet itself each time.

Performance speed for original code:

![2017](https://github.com/watsonlarry/stock-analysis/blob/main/Resources/Stocks%202017.png)
![2018](https://github.com/watsonlarry/stock-analysis/blob/main/Resources/Stocks_2018.png)

Performance speed for refactored code:

![2017r](https://github.com/watsonlarry/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![2018r](https://github.com/watsonlarry/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

The refactored code cut -1.20sec from the 2017 analysis time and -1.32sec from the 2018 analysis time.

### Summary
1. Refactoring code can have many advantages. By simplifying the code you make it easier to reuse, modify, or even return to (if some time passes between you and a project). By making the code more elegant and easier to read, your code becomes more welcoming to others. If you're planning on passing off your code to a partner or giving it to others to use, readability is helpful. There are risks though. Anytime you attempt to modify already working code, you risk breaking the original code. Though as long as you've saved the functional code in a separate location I don't see the risk being that great. 

2. The primary goal was a success--the refactored VBA script runs faster. Thus, one advantage of refactoring is immediately apparent: you can increase the processing speed. With regard to the simplification of the script post-refactoring, I think the positives are less certain. Working through the module I thought that writing the script with the given instructions was clear enough. As I built the multiple new arrays and for-loops for this challenge, I found that the syntax becomes more complex and harder to read. While it seems that the end result was worth the refactoring, I think the alterations made the VBA syntax harder to parse.
