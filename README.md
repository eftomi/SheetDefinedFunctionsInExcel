# SDF Add-In

### Sheet Defined Functions in MS Excel

#### If you developed a model in Excel spreadsheet and would like to expose it as function to be used in other parts of the same workbook or in other workbooks, this add-in can help you. 

Instead of coding a VBA function or constructing a complex lambda expression, you can develop your function as a spreadsheet model in a usual way by placing expressions in cells. After your model is developed with all the complexities, this add-in allows you to specify its inputs and outputs. By doing that, your model can be regarded as a "black box" - once developed and exposed as a sheet defined function, you will have no concerns about its inner workings and internal structure, you just use the results that are returned from it. In this way, you won't have to use any direct references into the model structure, nor the model should reference any cells outside of it. It will become an independent module.

Sheet defined function can take input parameters, recalculate itself and return the results. Your model, once exposed as sheet defined function can take one or more input arguments as individual cells or ranges, and return one or more results as individual values, ranges or arrays. Spill ranges and array formulas can be used as input arguments and results, too.

In this way, you can create and use many sheet defined functions simultaneously. They can be defined in one workbook and used in the same or in other workbooks. 

* [How to use SDF Add-In](/docs/Usage.md)


