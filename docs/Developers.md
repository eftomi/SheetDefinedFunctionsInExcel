## Information for developers

Let's develop a model which estimates the position (destination and altitude) of a projectile after it has been lounched with a certain speed at a certain angle (Source: https://en.wikipedia.org/wiki/Projectile_motion). Besides speed and angle, we also need time elapsed after lounch, and the initial altitude.

![Projectile model](/images/projectile1.png)

We have input values in cells B2 to B5, and projectile motion formulas in cells B7 and B8. We would like to use this model in other workbooks, but just its calculation "service" and not the actual structure by itself. In this way, we won't have to use any direct references into the model structure, nor the model should reference any cells outside of it. To do that, we can expose the model as a sheet defined function with two special worksheet functions which are a part of SDF add-in functionality: ModuleInput() and ModuleOutput(). We use these two functions within the model that should be exposed as a sheet defined function.

When used in a spreadsheet cell, each ModuleInput() function creates one "input slot" for the module, i.e. a cell that will take input values from module calls. Besides this, it specifies the name of the module (for our example we'll use the module name "Projectile"), the range of cells where the structure of the module is defined (ours is defined in range A2:B8 for now), the name of the input (e.g. "Initial speed"), and input's initial value (e.g. 130 m/s).

ModuleOutput() function declares the name of the output (e.g. "Distance"), the value that should be considered as the result (formula or cell reference, e.g. B7). We also have to reference the proper module by its name ("Projectile").

Module names in ModuleInput() and ModuleOutput() functions serve two purposes: (1) they uniquely define a module, and (2) they tie named inputs and outputs to this module.

For each module, we can use one or more inputs (i.e. ModuleInput() functions) and one or more outputs (i.e. ModuleOutput() functions). Our module "Projectile" will thus have four inputs and two outputs.

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

In this way, when we will call the module "Projectile" to get the "Distance" as a result, the ModuleOutput() in C7 will tell it to return the value from cell B7, where the formula for distance is entered. The "Altitude" will work in the same way.

![Projectile model](/images/projectile3.png)

Similar as above, we can use other cells to specify names (e.g. in cells A7 we already have names of model ouputs that can be used as output names).


### Module use

The module "Projectile" is prepared - let's try it out! We can create new worksheet and prepare the cells with input values a bit differently this time, for intance in a row like this:

![Projectile model](/images/projectile4.png)

We call our module with ModuleUse() function: 

`=ModuleUse(module_name, output_name, input_name_1, input_value_1, [input_name_2, input_value_2], ... )`

- `module_name` is the name of the module that we would like to use
- `output_name` is the name of the output that we need from module
- `input_name_1` is the name of the first input argument
- `input_value_1` is the value of the first input argument

We can use as many input arguments as needed by our module.

In our case, we of course need two results, distance and altitude, so we have two ModuleUse() functions. We can enter them in cells F2 and G2:

```
=ModuleUse("Projectile", "Distance", 
 "Initial speed", A2, "Initial angle", B2, "Initial altitude", C2, "Time", D2)
=ModuleUse("Projectile", "Altitude", 
 "Initial speed", A2, "Initial angle", B2, "Initial altitude", C2, "Time", D2)
```

Since we wrote input and output names in cells A1..D1 and F1..G1 exactly as defined by module, we can use cell references in ModuleUse() like this:

![Projectile model](/images/projectile5.png)

For the calculation to be performed, we click on *Calculate SDFs* button on *Sheet Defined Functions* ribbon.

![Projectile model](/images/calculateSDFs.png)

![Projectile model](/images/projectile6.png)

We can use the module many times, in other words it can be referenced by many ModuleUse() functions. Suppose that we need to estimate projectile trajectory for the first 5 seconds, in 0.5 seconds intervals. For simplicity, we can copy input values to cells below the first row, where we increment the time parameter by 0.5. Formulas with module calls can be copied from cells F1 and G1, too. After copying and clicking on *Calculate SDFs*, we get:

![Projectile model](/images/projectile7.png)

If we order input parameters and their names in rows or columns, we can instead use ModuleUseRangeInputs() function to call the module, which is more convenient:

`=ModuleUseRangeInputs(module_name, output_name, input_names, input_values)`

- `module_name` is a name of module that we would like to use
- `output_name` is the name of the output that we need from module
- `input_names` is the range of input names
- `input_values` is the range of input values

Needless to say, the order of input names and corresponding input values has to be the same in both ranges.

![Projectile model](/images/projectile8.png)

### Spill ranges as inputs and outputs

SDF add-in supports spill ranges. Module inputs and outputs can be arrays, not just single values. As a simple example, let's create a module that calculates the sum of arbitrary number of input values.

In a new worksheet, we'll arrange values that have to be summed up in a column, starting from cell A5 and down. This column will represent the module input, so we declare it with the formula in cell A5 like:

`=ModuleInput("Stats", A1:B20, "Values", {1, 2, 3})`

"Stats" is our new module, A1:B20 is the range of cells with the module structure, "Values" is the name of this input, and array {1, 2, 3} is the initial array to be summed up. After entering the formula, the array {1, 2, 3} should be spilled in cells A5..A7.

Let's put the sum of input values in cell B1. Since this is also the output from our model, we can declare it with a formula like:

`=ModuleOutput("Stats", "Sum", SUM(A5#))`

"Sum" is the name of the output, and return value will be the sum of spill range which begins by the A5 cell. Our module structure is not complicated at all:

![Projectile model](/images/projectile9.png)

The module "Stats" is prepared, and we can use it from some other worksheet. In the worksheet in the picture below, we have entered input values in cells A3..A11. 

![Projectile model](/images/projectile10.png)

We call the module with formula in cell B2:

`=ModuleUse("Stats", "Sum", "Values", A3:A11)`

After clicking on *Calculate SDFs* button, the module returns the sum. It behaves as expected - it updates the sum correctly if we change the size of the range of input values.

Outputs can be spilled ranges as well. Let's suppose we would like to have another output from this module - for each input value, the module should return a difference between specific value and the average of all input values. 

Firstly, we need a formula to calculate average of all input values:

`=AVERAGE(A5#)`

We can put this formula in cell B2 of our "Stats" module. Secondly, for each of the input values we need to calculate the difference. Since the input values are nicely ordered in spill range which begins at cell A5, we can simply write

`=A5#-B2`

in cell B5. B2 keeps the average of all the values, and A5 is the first value from our spill range. 

To safe the space, we can directly expose this formula as an output by wrapping it with ModuleOutput() function like in:

`=ModuleOutput("Stats", "Differences", A5#-B2)`

Finally, our module looks like this:

![Projectile model](/images/projectile11.png)

We can later refer to "Differences" output range with the formula like:

`=ModuleUse("Stats", "Differences", "Values", A3:A11)`

Our use of "Stats" module finally looks like:

![Projectile model](/images/projectile12.png)

![Projectile model](/images/StatsMovie.gif)

Besides spill ranges, we can use array formulas, too. In this case, we have to predict in advance the maximum size of the input and output arrays, and enter the ModuleInput() and ModuleOutput() functions as array formulas.