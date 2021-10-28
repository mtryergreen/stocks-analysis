# stocks-analysis
# Overview of Project: 
The purpose of this project is to refactor the code that analyzes returns on a variety of stocks, and report whether the refactored version is in fact more effective and efficient. 
The stocks were initially observed to determine if DQ was a good investment, and they were evidently not. Now, we analyze more stocks by looking at their volumes traded, starting prices and ending prices. 

# Results: 
Overall, the returns of all stocks were almost entirely at a gain (green) in 2017 (see attached image):

<img width="226" alt="All Stocks 2017" src="https://user-images.githubusercontent.com/89936913/139163820-1d672484-f973-47fd-9c09-9b8a1997b6b3.png">
... and the returns in 2018 were almost entirely at a loss (red) in 2018 (see attached image):

<img width="225" alt="All Stocks 2018" src="https://user-images.githubusercontent.com/89936913/139163988-deef776a-6658-44b5-84d4-9fbc2b76a8ca.png">

The refactored code for 2017 ran in 0.14 seconds, while the prior code ran just barely under 0.3 seconds, indicating that the refactored code indeed sped up the process
Here is a comparison via the images with the prior version first and refactored version last: 

![2017 unfactored](https://user-images.githubusercontent.com/89936913/139165728-6af2ad77-c9a8-452d-a184-9525986295da.png) ![VBA_Challenge_2017](https://user-images.githubusercontent.com/89936913/139281814-9b8f4755-8e6d-445c-9b38-401199133687.png)


A similar conclusion comes from the 2018 stock analysis data, for which the prior code ran just over 0.3 seconds and the refactored code ran just a hair over 0.1 seconds. 
Here is a comparison of the images again: 

![2018 prior](https://user-images.githubusercontent.com/89936913/139166068-7962571e-e503-42bc-8544-4465182d6b70.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/89936913/139281898-27d69378-8786-4240-aaf6-543438406948.png)


# Summary: In a summary statement, address the following questions.

## What are the advantages or disadvantages of refactoring code?
The advantages to refactoring code are most notably: you can shave a good chunk of a second off the loading time, which means the code is easier for the computer to process and will take up less space and energy; additionally, refactored code is often less text and is easier to read. 
The disadvantage to refactoring code is that it requires a pretty decent understanding of how the code works on its own before it is refactored. As a new coder who still find VBA much less intuitive than basic Excel or even Python and RStudio, this was a big challenge for me. Looking up advice and tutorials on Google and StackOverflow was kind of a saving grace when it came to learning how to create both the pre-refactored code and the more efficient version. 
## How do these pros and cons apply to refactoring the original VBA script?
The original VBA script features a lot more white space with more short lines of code spread apart- the refactored version features less of this space between code chunks, but a few longer lines of code. 
Again, knowing how to write these longer lines and condense the code was very tricky without consulting the module and some StackOverflow walkthroughs. 
Overall, a significant challenge for a new coder, probably a solid day's work, if even that, for a more experienced coder. 

# Code
The refactored code had several components which were specified in the module:

## Ticker Index

![Code 1](https://user-images.githubusercontent.com/89936913/139280413-795735af-d748-41dd-84cf-2f0f531ca4e6.png)

My ticker index was set to zero.
This is important because all arrays need to have a base of zero; and each array consisted of the ticker volumes, starting prices and ending prices. Ticker volumes was set as a long array in order to hold numeric values while starting and ending prices were single arrays because these values could include numbers with decimals. 


