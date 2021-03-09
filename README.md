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

### Module inputs

The model will take four arguments - initial speed, angle and altitude, and the time after lounch for which we are estimating the position. We define each of the four inputs with  function ModuleInput() which looks like:

`=ModuleInput(module_name, module_range, input_name, input_initial_value)`

- `module_name` is a name for our module. Since we can have many modules (many sheet defined functions), it is essential that we distinguish them by unique names. 
- `module_range` is a range of cells where the module is defined. 
- `input_name` is the name of the input argument.
- `input_initial_value` is the initial value for the input argument. This value will be used by our model while we are developing its inner structure and formulas. When the model is used, this value is ignored.

In our case, we can redefine four inputs in cells B2 to B5 as:

```
=ModuleInput("Projectile", A2:C8, "Initial speed", 130)
=ModuleInput("Projectile", A2:C8, "Initial angle", 25)
=ModuleInput("Projectile", A2:C8, "Initial altitude", 0)
=ModuleInput("Projectile", A2:C8, "Time", 0.5)
```

Our module will thus be called "Projectile", its structure and formulas are defined in the range A2:C8. The names of the inputs are descriptive, and we'll refer to them when we will use the model. The initial values will be displayed in cells B2 to B5 with ModuleInput() functions as their value; we can conveniently use them while constructing and updating the "body" of our module: 

![Projectile model](/images/projectile2.png)

As it can be seen from the picture above, we used absolute references for module range $A$2:$C$8 to simplify copying the formula from cell B2 to cells B3..B5. Since we have module argument names (initial speed, initial angle ...) already nicely written in cells A2..A5, we used references to these cells as the third argument to ModuleInput() functions. Similarly, we took values from cells B2..B5 to be initial input values for our module.

**In this way, we defined module "Projectile" with four input "slots" - in other words, cells B2 to B5 will take input parameters when the model will be used. This is important since the body of our model (formulas which calculate distance in altitude) should use these input values.**

### Module outputs

For the outputs (result values) we use function ModuleOutput():

`=ModuleOutput(module_name, output_name, output_value)`

- `module_name` is a name for our module
- `output_name` is the name of the output
- `output_value` is the value which will be returned to the caller

For our projectile, we define two outputs (distance and altitude), so we need two ModuleOutput() functions. Let's put them in cells C7 and C8 like:

```
=ModuleOutput("Projectile", "Distance", B7)
=ModuleOutput("Projectile", "Altitude", B8)
```

In this way, when we will call the module to get the "Distance" as a result, the ModuleOutput() in C7 will return the estimated value from cell B7, where the formula for distance is entered. The "Altitude" will work in the same way.

![Projectile model](/images/projectile3.png)

Similar as above, we can use other cells to define names (e.g. in cells A7 we have references to cells A7 and A8 with names of model ouputs).


### Module use

The module "Projectile" is prepared - let's try it out! Create new worksheet and prepare the cells with input values like this:

![Projectile model](/images/projectile4.png)

We can call our module with ModuleUse() function like this:

`=ModuleUse("Projectile", "Distance", "Initial speed", A2, "Initial angle", B2, "Initial altitude", C2, "Time", D2)`

