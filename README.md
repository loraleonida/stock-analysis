# Stock Analysis with VBA
*** Worked on challenge with study group: Ryan Knauff, Marshall Miley, Cayli Swartz & Shawn Logan ***

## Overview of Project

### Purpose
In this project we used VBA to loop through a data set of stocks in order to analyze twelve different stocks and determine the total daily volume and return for each stock over the course of two years, 2017 and 2018. We then refactored this code to be more efficient at analyzing a larger data set of stocks. We did this using by applying the skills acquired from this module, including creating a VBA macro, For loops, If Then conditionals, arrays, nested For loops, conditional formatting, pattern recognition, measuring code performance, and debugging.

### Analysis
The refactored code analysis of the 12 different stocks is more efficient and has faster performance times than the original code. I believe this is because the refactored code is cleaner and unnecessary code has been removed, particularly in the If Then conditional statements. Along with removing unnecessary code, the structure of the code was changed to make the refactored code run faster. The original code uses nested For loops, while the refactored code uses three separate For loops to perform the same analysis. These separate For loops were accomplished by adding code for three new arrays and a ticker index. It seems that in regards to looping through larger data sets, the separate For loops are faster than the nested For loops. The following pictures show the cleaned up If Then conditionals, as well as the addition of the code for the three arrays and ticker index and the changed For loop structure.


This image shows the original code and the nested For loops.

<img width="552" alt="Original Code" src="https://user-images.githubusercontent.com/101693004/161404563-d7b08651-0dea-4c4b-beec-90d4b4ed88d5.png">



This image shows the first part of the refactored code with the addition of the ticker index and the three new arrays. It also shows the first separate For loop, and the beginning of the second separate For loop.

<img width="591" alt="Refactored Code 1" src="https://user-images.githubusercontent.com/101693004/161404569-06b9e2c3-3ade-4f04-bb32-e4983940c3b6.png">



This image shows the second part of the refactored code with the end of the second separate For loop and the third separate For loop.

<img width="749" alt="Refactored Code 2" src="https://user-images.githubusercontent.com/101693004/161404578-802c01a6-331a-4beb-8708-4badeb5778bf.png">




## Results
According to the performance times, the refactored code for the stock analysis in 2017 and 2018 ran faster than the original code. The refactored code for 2017 ran 5.15 times faster than the original code, and the refactored code for 2018 ran 5.19 times faster than the original code. The following images show the performance times for original code and refactored code for each year.


<img width="935" alt="2017" src="https://user-images.githubusercontent.com/101693004/161404628-bf28ca1e-5273-4fe2-ae26-49a03f80d9b2.png">


<img width="908" alt="2017 Refactored" src="https://user-images.githubusercontent.com/101693004/161404640-6ec3b549-5dd2-460f-8864-48e3fb56c328.png">



<img width="952" alt="2018" src="https://user-images.githubusercontent.com/101693004/161404647-81b5e8d3-f2b3-4e39-bb3d-3602ba54be23.png">


<img width="908" alt="2018 Refactored" src="https://user-images.githubusercontent.com/101693004/161404652-030488cf-a747-4ef0-9774-dc6b70fbdb51.png">




## Summary

- What are the advantages or disadvantages of refactoring code?

Some of the advantages of refactoring code include making the code more efficient by taking fewer steps, and using less memory to run the code, therefore making the code run faster. Also, when refactoring code it can make the logic of the code easier to follow and for others to read.
Some of the disadvantages of refactoring code are that it's time consuming to do, you may not always know how to make the code faster so it might take a lot of research, and as you refactor and rewrite code you might introduce more bugs to the code during the process (Zhala et al,. "What Are The Advantages and Disadvantages of Refactoring Code Smell in Software Quality").

- How do these pros and cons apply to refactoring the original VBA script?

Most of these pros and cons applied to the process of refactoring the original VBA script. The refactored code is more efficient, it takes fewer steps, it runs faster, and it's easier to read. At the same time, the process of refactoring was time consuming especially when I wasn't sure how to accomplish some of it. I also introduced more bugs and had to work through debugging the new code when the original code was already running cleanly.


Zhala SarkawtZhala Sarkawt 1111 silver badge11 bronze badge, et al. “What Are the Advantages and Disadvantages of Refactoring Code Smell in Software Quality?” Stack Overflow, 15 May 2017, https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software. 
