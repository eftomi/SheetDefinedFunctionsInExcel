# SDF Add-In

## How to use SDF Add-In

Let's develop a model which estimates the position (destination and altitude) of a projectile after it has been lounched with a certain speed at a certain angle (Source: https://en.wikipedia.org/wiki/Projectile_motion). Besides speed and angle, we will also need the time that elapsed after lounch, and the initial altitude. 

![Projectile model](/images/projectile1.png)

Let's arrange input values in cells B2 to B5, and projectile motion formulas in cells B7 and B8 as above. We would like to use this model in other workbooks, but just its calculation "service" and not the actual structure by itself, so that we could write a formula to do the calculation in an easy way, just by calling one function instead of copying the whole calculation block from cells A2 to B8. In this way, we won't have to use any direct references into the model structure (its outputs), nor the model should reference any cells outside of it. SDFs enable the modularization of workbooks, that is the separation of the functionalities into independent, interchangeable modules, such that each contains everything necessary to compute a particular aspect of the desired functionality. 

As we see, the inputs of our model are represented by cells B2 to B5 (input parameter values), and outputs (results) of our module are the cells B7 and B8.

To expose the model as a sheet-defined function (SDF), we use two special worksheet functions that are a part of SDF add-in functionality: ModuleInput() and ModuleOutput(). 

Each ModuleInput() function creates one "input slot" for the SDF. The cell with this function (within the model that we would like to expose as SDF) will take and display input values that we send to SDF when it will be called from outside. Besides this, ModuleInput() specifies the name of the SDF (for our example we'll use the name "Projectile"), the name of the input (e.g. "Initial speed"), and input's initial value (e.g. 130 m/s) - like this:

`=ModuleInput("Projectile", "Initial speed", 130)`

ModuleOutput() function declares the value that should be considered as the result (formula or cell reference, e.g. B7) and the output name (e.g. "Distance"), because a model can expose two or more outputs. We also have to declare the SDF that the output belongs to ("Projectile"):

`=ModuleOutput("Projectile", B7, "Distance")`

SDF names (like "Projectile") in ModuleInput() and ModuleOutput() functions serve two purposes: (1) they uniquely define a SDF, and (2) they tie named inputs and outputs to this SDF.

For each SDF, we can declare zero, one or more inputs (with ModuleInput() functions) and one or more outputs (with ModuleOutput() functions). Our SDF "Projectile" will thus have four inputs and two outputs. 

### More about SDF inputs

Our "projectile" model takes four arguments - initial speed, angle and altitude, and the time after lounch for which we are estimating the position of the projectile. We declare each of the four inputs with function ModuleInput() which looks like:

`=ModuleInput(module_name, [input_name], [input_initial_value], [enforce_my_input_values])`

- `module_name` is the name for our SDF. Since we can have many distinct SDFs, it is essential that we distinguish them by unique names. 
- `input_name` is the name of the input argument. If SDF has only one input, this parameter can be ommited, because there is no need to set the name of the input.
- `input_initial_value` is the initial value for this input argument. This value will be used by our SDF while we are developing its inner structure and formulas. When the SDF will be used from "outside", this value will be overriden by actual input values. If ommited, the value of 0 will be used.
- `enforce_my_input_values` if TRUE, the `input_initial_value` will always be used for calculations, which comes handy for debugging purposes (FALSE if ommited).

It may happen that the SDF does not need any inputs - in this case, there is no need to use ModuleInput() function.

In our case, we can rewrite four input parameters in cells B2 to B5 to:

```
=ModuleInput("Projectile", "Initial speed", 130)
=ModuleInput("Projectile", "Initial angle", 25)
=ModuleInput("Projectile", "Initial altitude", 0)
=ModuleInput("Projectile", "Time", 0.5)
```

Our SDF will thus be called "Projectile". The names of the inputs are descriptive, and we'll refer to them when we will use the SDF. When entering the ModuleInput() functions into cells B2 to B5, the initial values 130, 25, 0 and 0.5 will be displayed in these cells; we can conveniently use them while constructing and updating the "body" of our SDF: 

![Projectile model](/images/projectile2.png)

As it can be seen from the picture above, we have SDF argument names (initial speed, initial angle ...) already nicely written in cells A2..A5, so we used references to these cells as the second argument to ModuleInput() functions. 

In this way, we defined SDF "Projectile" with four input "slots" - in other words, cells B2 to B5 will take input parameters when the SDF will be used. This is important since the body of our SDF (formulas which calculate distance in altitude) should use (reference) these input values.

### SDF outputs

To declare outputs (results of the module) we use function ModuleOutput():

`=ModuleOutput(module_name, output_value, [output_name])`

- `module_name` is the name for our SDF.
- `output_value` is the value which will be returned to the caller. 
- `output_name` is the name of the output that we are declaring. If SDF has only one output, this parameter can be ommited, because there is no need to set the name of the output.

To define SDF, we need at least one ModuleOutput() function.

For our projectile, we declare two outputs (distance and altitude), so we need two ModuleOutput() functions. Let's put them in cells C7 and C8 like:

```
=ModuleOutput("Projectile", B7, "Distance")
=ModuleOutput("Projectile", B8, "Altitude")
```

In this way, when we will call the SDF "Projectile" to get the "Distance" as a result, the ModuleOutput() in C7 will tell it to return the value from cell B7, where the formula for distance is entered. The "Altitude" will work in the same way.

![Projectile model](/images/projectile3.png)

Similar as above, we can use other cells to specify names (e.g. in cells A7 we already have names of model ouputs that can be used as output names).


### SDF use

The "Projectile" SDF is prepared - let's try it out! We can create new worksheet and prepare the cells with input values a bit differently this time, for intance in a row like this:

![Projectile model](/images/projectile4.png)

We call our SDF with ModuleUse() function: 

`=ModuleUse(module_name, [output_name_and_inputs (...)] )`

- `module_name` is the name of the SDF that we would like to use.
- `output_name_and_inputs` declares the name of output that should be returned. If the module has only one output, its name can be ommited. Subsequent parameters are input names & input values, given in pairs. If module has one input, provide just its value without name.

We can use as many input arguments as needed by our SDF.

In our case, we of course need two results, distance and altitude, so we have two ModuleUse() functions. We can enter them in cells F2 and G2:

```
=ModuleUse("Projectile", "Distance", 
 "Initial speed", A2, "Initial angle", B2, "Initial altitude", C2, "Time", D2)
=ModuleUse("Projectile", "Altitude", 
 "Initial speed", A2, "Initial angle", B2, "Initial altitude", C2, "Time", D2)
```

Because "Projectile" has four inputs, we also used four input name & input value pairs.

Since we wrote input and output names in cells A1..D1 and F1..G1 exactly as defined by SDF, we can use cell references in ModuleUse() like this:

![Projectile model](/images/projectile5.png)

For the calculation to be performed, we click on *Calculate SDFs* button on *Sheet Defined Functions* ribbon.

![Projectile model](/images/calculateSDFs.png)

![Projectile model](/images/projectile6.png)
 
We can use the SDF many times, in other words it can be referenced by many ModuleUse() functions. Suppose that we need to estimate projectile trajectory for the first 5 seconds, in 0.5 seconds intervals. For simplicity, we can copy input values to cells below the first row, where we increment the time parameter by 0.5. Formulas with SDF calls can be copied from cells F1 and G1, too. After copying and clicking on *Calculate SDFs*, we get:

![Projectile model](/images/projectile7.png)

To get the overview of all SDFs and their arguments, we can click Use SDFs button on the toolbar, which brings a dialog box with which we can construct the ModuleUse() formula:

![Projectile model](/images/dialogBox.png)

Sometimes it is convenient to arrange input parameters and their names in rows or columns. In such a case, we can use ModuleUseRangeInputs() function to call the SDF, which is more convenient:

`=ModuleUseRangeInputs(module_name, output_name, input_names, input_values)`

- `module_name` is a name of SDF that we would like to use
- `output_name` is the name of the output that we need from SDF
- `input_names` is the range of input names
- `input_values` is the range of input values

Needless to say, the order of input names and corresponding input values has to be the same in both ranges.

![Projectile model](/images/projectile8.png)

### Spill ranges as inputs and outputs

SDF add-in supports spill ranges. SDF inputs and outputs can be arrays, not just single values. As a simple example, let's create a module that calculates the sum of arbitrary number of input values.

In a new worksheet, we'll arrange values that have to be summed up in a column, starting from cell A5 and down. This column will represent the SDF input, so we declare it with the formula in cell A5 like:

`=ModuleInput("Stats", "Values", {1, 2, 3})`

"Stats" is our new SDF, "Values" is the name of this input, and array {1, 2, 3} is the initial array to be summed up. After entering the formula, the array {1, 2, 3} should be spilled in cells A5..A7.

Let's put the sum of input values in cell B1. This is also the output from our SDF, so we can declare it with a formula like:

`=ModuleOutput("Stats", SUM(A5#), "Sum")`

"Sum" is the name of the output, and return value will be the sum of spill range which begins by the A5 cell. Our SDF structure is not complicated at all:

![Projectile model](/images/projectile9.png)

The SDF "Stats" is prepared, and we can use it from some other worksheet. In the worksheet in the picture below, we have entered input values in cells A3..A11. 

![Projectile model](/images/projectile10.png)

We call the SDF with formula in cell B2:

`=ModuleUse("Stats", "Sum", "Values", A3:A11)`

After clicking on *Calculate SDFs* button, the SDF returns the sum. It behaves as expected - it updates the sum correctly if we change the size of the range of input values.

Because the "Stats" SDF has only one input and only one output, they can be "anonymous", without names. Thus, instead of the above functions, we could declare:

`=ModuleInput("Stats",, {1, 2, 3})`

and

`=ModuleOutput("Stats", SUM(A5#))`

We would then call the function as in:

`=ModuleUse("Stats",, A3:A11)`

Outputs can be spilled ranges as well. Let's suppose we would like to have another output from this SDF - for each input value, the module should return a difference between specific value and the average of all input values. 

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

Our use of "Stats" SDF finally looks like:

![Projectile model](/images/projectile12.png)

Besides spill ranges, we can use array formulas, too. In this case, we have to predict in advance the maximum size of the input and output arrays, and enter the ModuleInput() and ModuleOutput() functions as array formulas.
