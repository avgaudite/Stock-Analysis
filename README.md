# Stock Analysis

## Overview
![](https://i.imgur.com/7jWNx6P.png)
![](https://i.imgur.com/k2xA8XG.png)


### Purpose 
To help Steve determine which green energy stocks his parents should invest in besides DQ, where currently, all of their money is invested.

### Results
The only two stocks with positive returns were "ENPH" and "RUN". Our refactored code returned these results in 0.17 seconds. Our original code returned these results in .97 seconds. It is almost five times faster.

The original code, with its back and forth jumps between sheets before loops, is slower. The refactored code does not make these switches; in its loop, all data is captured and stored in arrays.

Creating a ticker index and creating three output arrays I believe were instrumental in getting the runtime down to a fifth of the original.

> tickerIndex = 0
> Dim tickerVolumes(12) As Long
> Dim tickerStartingPrices(12) As Single
> Dim tickerEndingPrices(12) As Single 

#### Advantages and Drawbacks of Refactoring Code

In refactoring code, it becomes more readable (for us and the computer). The result is more efficiency in terms of runtime.

I qualified the previous statement because even though the runtime is faster, to have to rewrite the code, especially when the person is a novice, is much slower and any time gained is quickly lost.

The original VBA script, though longer, I found easier to follow since every step is shown. The refactored code is harder to follow for the novice. If not for the steps detailed throughout the module, refactoring the code, an exercise already difficult, would have much more so.
