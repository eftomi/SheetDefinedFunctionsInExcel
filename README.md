# SDF Add-In

## Sheet Defined Functions in MS Excel

### If you developed a model in Excel spreadsheet and would like to expose it as a function to be used in other parts of the same workbook or in other workbooks, this add-in can help you. 

Instead of coding a VBA function or constructing a complex lambda expression, you can develop your function as a spreadsheet model in a traditional way by putting expressions in cells. After your model is developed, you define its inputs and outputs. By doing that, your model can be regarded as a "black box" - once developed and exposed as a sheet defined function, you will have no concerns about its inner workings and internal structure, you just use its results. 

Sheet defined function can take input parameters, recalculate itself and return the results. Your model, once exposed as sheet defined function, can take one or more input arguments as individual cells or ranges, and return one or more results as individual values, ranges or arrays. SDF add-in also supports spill ranges and array functions.

In this way, you can define and use many sheet defined functions simultaneously. They can be defined in one workbook and used in the same or in other workbooks. 

## Basic use

For example, let's develop a model which estimates the position (destination and altitude) of a projectile after it has been lounched with a certain speed at a certain angle (Source: https://en.wikipedia.org/wiki/Projectile_motion).

![Projectile model](/images/projectile1.png)

We would like to use this model in other workbooks, but just its calculation "service" and not the structure itself. For that, we expose it as a sheet defined function with two special worksheet functions: ModuleInput() and ModuleOutput().

### How to define module inputs

The model will take four arguments - initial speed, angle and altitude, and the time after lounch for which we are estimating the position. We define each of the four inputs with  function ModuleInput() which looks like:

`=ModuleInput(module_name, module_range, input_name, input_initial_value)`

- `module_name` is a name for our module. Since we can have many modules (many sheet defined functions), it is essential that we distinguish them by unique names. 
- `module_range` is a range of cells where the module is defined. 
- `input_name` is the name of the input argument.
- `input_initial_value` is the initial value for the input argument. This value will be used by our model while we are developing its inner structure and formulas. When the model is used, this value is ignored.

In our case, we can define four inputs in cells C2 to C5 as:

```
=ModuleInput("Projectile", A2:C8, "Initial speed", 130)
=ModuleInput("Projectile", A2:C8, "Initial angle", 25)
=ModuleInput("Projectile", A2:C8, "Initial altitude", 0)
=ModuleInput("Projectile", A2:C8, "Time", 0.5)
```

Our module will thus be called "Projectile", its structure and formulas are defined in the range $A$2:$C$8. The names of the inputs are descriptive, and we'll refer to them when we will use the model. The initial values will be displayed in cells with ModuleInput() functions as their result; we can conveniently use them while constructing the module. 

![Projectile model](/images/projectile2.png)

As it can be seen in the picture above, we used absolute references for module range $A$2:$C$8 to simplify copying the formula from cell C2 to cells C3..C5. Since we have module argument names (initial speed, initial angle ...) already nicely written in cells A2..A5, we used references to these cells as the third argument to ModuleInput() functions. Similarly, we took values from cells B2..B5 to be initial input values for our module.

In this way, we defined module "Projectile" with four input "slots" - in other words, cells C2 to C5 will take input parameters when the model will be used. 

### How to define module outputs

For the outputs (result values) we use function ModuleOutput() - one function for a given output:

`=ModuleOutput(module_name, output_name, output_value)`

- `module_name` is a name for our module
- `output_name` is the name of the output
- `output_value` is the value which will be returned to the caller

For our projectile, we can define two outputs (distance and altitude) in cells C7 and C8 like

```
=ModuleOutput("Projectile", "Distance", B7)
=ModuleOutput("Projectile", "Altitude", B8)
``` 

In this way, when we will call the module to get the "Distance" as a result, it will return the estimated value from cell B7, where the formula for distance is entered. The "Altitude" will work in the same way.

![Projectile model](/images/projectile3.png)

Similar as above, we can use other cells to define names (e.g. in cells A7 we have names of model ouputs).

**Please note that we changed the model expressions in cells B7 and B8 to use input parameters from cells C2, C3, C4 and C5. These cells are regarded as input "slots" into our "Projectile" module, and expressions in cells B7 and B8 will take these parameters to calculate the results. We defined these input slots by using function ModuleInput() in them.**

### How to use the module

bla bla