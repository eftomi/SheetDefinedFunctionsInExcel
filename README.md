# SDF Add-In

## Sheet Defined Functions in MS Excel

### If you developed a model in Excel spreadsheet and would like to expose it as a function to be used in other parts of the same workbook or in other workbooks, this add-in can help you. 

Instead of coding a VBA function or constructing a complex lambda expression, you can develop your function as a spreadsheet model in a traditional way by putting expressions in cells. After your model is developed with all the complexities, this add-in allows you to specify its inputs and outputs. By doing that, your model can be regarded as a "black box" - once developed and exposed as a sheet defined function, you will have no concerns about its inner workings and internal structure, you just use the results that are returned from it. 

Sheet defined function can take input parameters, recalculate itself and return the results. Your model, once exposed as sheet defined function can take one or more input arguments as individual cells or ranges, and return one or more results as individual values, ranges or arrays. Spill ranges and array formulas can be used as input arguments and results, too.

In this way, you can create and use many sheet defined functions simultaneously. They can be defined in one workbook and used in the same or in other workbooks. 

## Basic use

Let's develop a model which estimates the position (destination and altitude) of a projectile after it has been lounched with a certain speed at a certain angle (Source: https://en.wikipedia.org/wiki/Projectile_motion). Besides speed and angle, we also need time elapsed after lounch, and the initial altitude.

![Projectile model](/images/projectile1.png)

We have input values in cells B2 to B5, and projectile motion formulas in cells B7 and B8. We would like to use this model in other workbooks, but just its calculation "service" and not the actual structure by itself. In this way, we won't have to use any direct references into model structure, nor the model should reference any cells from outside of it. To do that, we can expose the model as a sheet defined function with two special worksheet functions: ModuleInput() and ModuleOutput() which are a part of SDF add-in functionality. We use these two functions within the model that should be exposed as sheet defined function.

When used in a spreadsheet cell, each ModuleInput() function creates an "input slot" for the module, i.e. a cell that will take input values from module calls. Besides this, it specifies the name of the module (for our example we'll use the name "Projectile"), the range of cells where the structure of the module is defined (ours is defined in range A2:B8 for now), the name of the input (e.g. "Initial speed"), and input's initial value (e.g. 130 m/s).

ModuleOutput() function declares the name of the output (e.g. "Distance"), the value that should be considered as the result (formula or cell reference, e.g. B7), and the name of the module ("Projectile").

For each module, we can use one or more inputs and one or more outputs. Our module "Projectile" will have four inputs and two outputs.

Module names in ModuleInput() and ModuleOutput() functions serve two purposes: (1) they uniquely define a module, and (2) they tie named inputs and outputs to this module.

### Module inputs

Our "projectile" model takes four arguments - initial speed, angle and altitude, and the time after lounch for which we are estimating the position of the projectile. We declare each of the four inputs with function ModuleInput() which looks like:

`=ModuleInput(module_name, module_range, input_name, input_initial_value)`

- `module_name` is the name for our module. Since we can have many modules (many sheet defined functions), it is essential that we distinguish them by unique names. 
- `module_range` is the range of cells where the module structure and formulas are defined. 
- `input_name` is the name of the input argument.
- `input_initial_value` is the initial value for this input argument. This value will be used by our model while we are developing its inner structure and formulas. When the model will be used from "outside", this value will be overriden by actual input values.

In our case, we can rewrite four input parameters in cells B2 to B5 to:

```
=ModuleInput("Projectile", A2:C8, "Initial speed", 130)
=ModuleInput("Projectile", A2:C8, "Initial angle", 25)
=ModuleInput("Projectile", A2:C8, "Initial altitude", 0)
=ModuleInput("Projectile", A2:C8, "Time", 0.5)
```

Our module will thus be called "Projectile", its structure and formulas are defined in range A2:C8. The names of the inputs are descriptive, and we'll refer to them when we will use the model. When entering the ModuleInput() functions into cells B2 to B5, the initial values 130, 25, 0 and 0.5 will be displayed in these cells; we can conveniently use them while constructing and updating the "body" of our module: 

![Projectile model](/images/projectile2.png)

As it can be seen from the picture above, we used absolute references for module range $A$2:$C$8 to simplify copying the formula from cell B2 to cells B3..B5. Since we have module argument names (initial speed, initial angle ...) already nicely written in cells A2..A5, we used references to these cells as the third argument to ModuleInput() functions. 

**In this way, we defined module "Projectile" with four input "slots" - in other words, cells B2 to B5 will take input parameters when the model will be used. This is important since the body of our model (formulas which calculate distance in altitude) should use (reference) these input values.**

### Module outputs

To declare outputs (results of the module) we use function ModuleOutput():

`=ModuleOutput(module_name, output_name, output_value)`

- `module_name` is the name for our module
- `output_name` is the name of the output that we are declaring
- `output_value` is the value which will be returned to the caller

For our projectile, we declare two outputs (distance and altitude), so we need two ModuleOutput() functions. Let's put them in cells C7 and C8 like:

```
=ModuleOutput("Projectile", "Distance", B7)
=ModuleOutput("Projectile", "Altitude", B8)
```

In this way, when we will call the module to get the "Distance" as a result, the ModuleOutput() in C7 will return the estimated value from cell B7, where the formula for distance is entered. The "Altitude" will work in the same way.

![Projectile model](/images/projectile3.png)

Similar as above, we can use other cells to specify names (e.g. in cells A7 we have references to cells A7 and A8 with names of model ouputs).


### Module use

The module "Projectile" is prepared - let's try it out! We create new worksheet and prepare the cells with input values like this:

![Projectile model](/images/projectile4.png)

We can call our module with ModuleUse() function: 

`=ModuleUse(module_name, output_name, input_name_1, input_value_1, [input_name_2, input_value_2], ... )`

- `module_name` is a name of module that we would like to use
- `output_name` is the name of the output that we need from module
- `input_name_1` is the name of the first input argument
- `input_value_1` is the value of the first input argument

We can use as many input arguments as needed by our module.

In our case, we need distance and altitude, so we have two ModuleUse() functions. We can enter them in cells F2 and G2:

`=ModuleUse("Projectile", "Distance", "Initial speed", A2, "Initial angle", B2, "Initial altitude", C2, "Time", D2)`
`=ModuleUse("Projectile", "Altitude", "Initial speed", A2, "Initial angle", B2, "Initial altitude", C2, "Time", D2)`

Since we wrote input and output names in cells A1..D1 and F1..G1 exactly as defined by module, we can use cell references in ModuleUse() like this:

![Projectile model](/images/projectile5.png)

For the calculation to be performed, we click on *Calculate SDFs* button on *Sheet Defined Functions* ribbon.

![Projectile model](/images/calculateSDFs.png)

![Projectile model](/images/projectile6.png)

We can use the module many times. Suppose that we need to estimate projectile trajectory for the first 5 seconds, in 0.5 seconds intervals. For simplicity, we can copy input values to cells below, where we increment the time parameter by 0.5. Formulas with module calls can be copied from cells F1 and G1, too. After copying and clicking on *Calculate SDFs*, we get:

![Projectile model](/images/projectile7.png)

If we order input parameters and their names in rows or columns, we can instead use ModuleUseRangeInputs() function to call the module, which is more convenient:

`=ModuleUseRangeInputs(module_name, output_name, input_names, input_values)`

- `module_name` is a name of module that we would like to use
- `output_name` is the name of the output that we need from module
- `input_names` is the range of input names
- `input_values` is the range of input values

Needless to say, the order of input names and corresponding input values has to be the same in both ranges.

![Projectile model](/images/projectile8.png)