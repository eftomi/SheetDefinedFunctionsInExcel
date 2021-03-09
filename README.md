# SDF Add-In for MS Excel

## Sheet Defined Functions in MS Excel

#### This add-in enables the use of a part of an Excel workbook as a function. 

If you developed a model in Excel spreadsheet and would like to expose it as a function to be used in other parts of the same workbook or in other workbooks, you can use this add-in. Instead of coding a VBA function or constructing a complex lambda expression, you can develop your function in traditional way as a spreadsheet model and define inputs and outputs for it. After that, your model can be regarded as a "black box" - once developed and exposed as a sheet defined function, you have no concerns about its inner workings and internal structure, you just use its results. 

Sheet defined function can take input parameters, recalculate itself and return the results. Your model, once exposed as sheet defined function, can take one or more input parameters as individual cells or ranges, and return one or more results as individual values, ranges or arrays. SDF add-in also supports spill ranges and array functions.

In this way, you can define and use many sheet defined functions simultaneously. They can be defined in one workbook and used in same or in other workbooks. 

For example, let's develop a model which estimates the position (destination and altitude) of a projectile which is lounched with a certain speed at a certain angle (source: https://en.wikipedia.org/wiki/Projectile_motion).

![Projectile model](/images/projectile1.png)

bla bla