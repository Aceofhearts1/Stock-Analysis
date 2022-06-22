# Stock-Analysis
## VBA Macros Overview
In this analysis we wanted to follow the trends of the stocks. The dataset was pretty big. Since our data came in an excel sheet, we had an oppurtuinty to use VBA macros to make our jobs a little easier. In excel, we could have just repeatedly typed our formualas into the cells and eventually received the same nummbers we had after running our macros. However, that would have taken quite a bit of time. As well as, would've been difficult to keep up with what we were doing if we did not form all of our data in one day. Our goal was to show which stocks were safer to pick. To show which ones had a positive trend going between to years. Once we were able to code the macros to show the trends, our next goal was to format the results. This would then help people who are looking at the visuals to understand what is happening.

**The Visuals:**
![The Visuals](https://github.com/Aceofhearts1/Stock-Analysis/blob/main/Resources/VBA_Challenge_Not_Refactored_2017.png)
![The Visuals](https://github.com/Aceofhearts1/Stock-Analysis/blob/main/Resources/VBA_Challenge_Not_refactored_2018.png)

## Refactoring
Refactoring code is a way to speed up and lessen the work of the computer to achieve the same results. "Don't Repeat Yourself," is a big rule in coding. During the refactoring process I also took the liberty to test out a few coding ideas. I thought to myself, what if the data I receive is not in order. I know how to place it in order but that was not my point. So I based my search on comparing the dates to one another. To find the earliest date, I set a date far into the fututre and coded it to switch to the new date if it was earlier. Once the program ran through all of the dates for the specific stock, we would have our starting date. Same for the ending date.

**Our Results:**
![The Results from 2017](https://github.com/Aceofhearts1/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)
![The Results from 2018](https://github.com/Aceofhearts1/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)

## The Pros and Cons
The pros of refactoring code are that you can get your macros to run faster and it usually means it is more organized with less actual characters involved. The loops I created saved me a ton of time from having to set the array varaibles to 0 and the dates to whatever I placed them at. Creating a loop to do it for me saved me at least 30 minutes. Who doesn't love extra time to do other things in life. I truly don't see too many cons with refactoring. Howver, it could probably leave you in a spot where you might mess up the code you have already written and cause you to lose some time and headaches. There are options oput there to ensure that your code doesn't break so that can be prevented.

**How that ties in with our code:**
Refactoring did cause me break my script a few times. I encountered error after error but that is a part of the growing pains. There will always be a code that does not work the way you intended. This was great practice on the ups and downs of breaking your code. That was a con but the pros outweigh them in this. The refactoring that I did allows me to save time with the actual processing of the script. It saved me time in not having to repeat lines of code repeatedly. Now the code is organized and if the format of the data were to ever change, theer are very few changes I would have to make to get my script to run again.
