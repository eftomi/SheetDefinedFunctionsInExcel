## Information for developers

#### Motivation

In spreadsheets, calculations are typically defined as a group of cells with formulas that reference other cells in the same group. If the user wishes to reuse this functionality, the formulas can be copied to other spreadsheet locations, however this takes time, the user is usually concerned by inner structure of the model that is being copied, and the process might introduce errors and inconsistencies. Jones et al. (2003) proposed that a function can be defined by an ordinary worksheet with specially-identified input and output cells. Sheet-defined functions (SDFs) were firstly implemented by Sestof and Sørensen (2013) in Funcalc spreadsheet platform. 

The most natural implementation of this idea in Excel is the introduction of special worksheet functions which would allow the end user to define SDF together with its inputs and outputs. That kind of worksheet functions can be developed relatively easy by Visual Basic for Applications (VBA). However, if the SDF should remain active in the worksheet after we define it, in the sense that Excel recalculates it each time the SDF is being called (a call to SDF can be used in many formulas, and issued in each spreadsheet recalculation if necessary), then the biggest obstacle is the Excel calculation chain. 

Namely, for the purposes of cell recalculation, Excel constructs a dependency tree based on references in cell formulas ("Excel Recalculation," 2018). From this, a calculation chain is constructed, which lists all the cells that contain formulas in the order in which they should be calculated. The construction of calculation tree and chain depends from cell references. The recalculation mechanism cannot provide for many calls to the same SDF in such a way that it would return the output from SDF for each of the calls. Additionally, since SDFs should be callable by their name and not by reference, we cannot directly rely on Excel recalculation. If we call SDF by name and not by reference, then it will not be recalculated, since it is not included in the calculation chain. Thus, the execution (recalculation) of SDFs should be controlled independently from Excel usual recalculation.

Lately, Microsoft intruduced LAMBDA function in Excel, which deals with the obstacle of calculation chain. Together with the help of named ranges the resulting function can be called by its name and recalculated from elsewhere. However, the lambda function has to be defined in one formula (one cell), which is not suitable for complex models where the calculation is defined by formulas in two or more cells. According to "LAMBDA: The ultimate Excel worksheet function" (2021), Microsoft plans to introduce sheet-defined functions in future Excel versions. The SDF Add-In is a "proof of concept" prototype of sheet-defined functions in Excel.

#### Implementation

SDF add-in is implemented by three VBA modules:
* WSFunction module exposes user-defined functions ModuleInput(), ModuleOutput(), ModuleUse(), and ModuleUseRangeInputs()
* Recalculation module provides functionality for SDF recalculation
* Ribbon module implements ribbon user interface hooks for the add-in

Additionally, add-in defines four classes and four collections (eight class modules):
* objModule class that represents a particular SDF
* objInput class represents SDF's input
* objOutput class represents SDF's output
* objModuleUse class represents the use (a call) to a module

objModule object has public properties which describe module name, a collection of module inputs, a collection of module outputs, a collection of module uses (calls), and the range of cells of the module for recalculation purposes. Module WSFunction has a public variable AllModules which is a collection of all modules (SDFs).

Functions ModuleInput(), ModuleOutput(), ModuleUse(), and ModuleUseRangeInputs() are being executed each time the recalculation (automatic or manual) of a worksheet is executed. Depending from workbook complexity, this execution can happen several times, not only once. Each of these functions firstly looks for the module definition in AllModules collection; if the module is not found, it is created. Then, function ModuleInput() checks for the existence of the named input that it defines, and creates one if necessary. It also sets the initial return value, so that the end user can use this value as a cell result for prototyping and module construction purposes. ModuleInput() function defines the range of the module, so that the module can be recalculated by add-in.

Function ModuleOutput() creates the named output for a given module, including a property with a formula (typically with a reference) that should provide a return value of this output. 

A particular call to SDF is made by functions ModuleUse() and ModuleUseRangeInputs(), which set the input values for this call and equip the called objModule with the information of the cell from which the call was made (object objModuleUse). A collection of objModuleUse objects is used to recalculate SDF outputs for each SDF call.

When the end user initiates SDF recalculation from the ribbon, the RecalculateModules() method is called. Firstly, all the information of SDFs is refreshed by setting AllModules variable to an empty collection, and full recalculation of worksheet is done, so that ModuleInput(), ModuleOutput(), ModuleUse(), and ModuleUseRangeInputs() functions enumerate all modules, inputs, outputs and module uses by their name. After that, for each of the modules, cells with ModuleUse() and ModuleUseRangeInputs() functions (module calls) are evaluated in such a way that:
* the cell with the call is recalculated, so that it enumerates all input values for this module and use
* the module range is recalculated (i. e. the SDF performs its calculation for this particular call)
* the cell with the call is recalculated once more, so that it reads the result of the module calculation, and displays the result


#### References

* Jones, S.P., Blackwell, A., Burnett, M., 2003. A User-Centred Approach to Functions in Excel. SIGPLAN Not. 38, 165–176. https://doi.org/10.1145/944746.944721
* Sestoft, P., Sørensen, J.Z., 2013. Sheet-Defined Functions: Implementation and Initial Evaluation, in: Dittrich, Y., Burnett, M., Mørch, A., Redmiles, D. (Eds.), End-User Development. Springer Berlin Heidelberg, Berlin, Heidelberg, pp. 88–103.
* Excel Recalculation, 2018. https://docs.microsoft.com/en-us/office/client-developer/excel/excel-recalculation
* LAMBDA: The ultimate Excel worksheet function, 2021. https://www.microsoft.com/en-us/research/blog/lambda-the-ultimatae-excel-worksheet-function/
