## Information for developers

#### Motivation

In spreadsheets, calculations are usually defined as a group of cells with formulas that reference other cells in the same group. If the user wishes to reuse this functionality, the formulas can be copied to other spreadsheet locations, however this takes time, the user is usually concerned by inner structure of the model that is being copied, and the process might introduce errors and inconsistencies. Jones et al. (2003) proposed that a function can be defined by an ordinary worksheet with specially-identified input and output cells. Sheet-defined functions (SDFs) were firstly implemented by Sestof and Sørensen (2013) in Funcalc spreadsheet platform. 

The most natural implementation of this idea in Excel would be the introduction of special worksheet functions which would allow the user to define SDF together with its inputs and outputs. That kind of worksheet functions can be developed relatively easy by Visual Basic for Applications. However, if the SDF should remain active in the worksheet after we define it, in the sense that Excel recalculates it each time the SDF is being used (a call to SDF can be used in many formulas, and issued in each spreadsheet recalculation if necessary), then the biggest obstacle is the Excel calculation chain. 

Namely, for the purposes of cell recalculation, Excel constructs a dependency tree based on references in cell formulas ("Excel Recalculation," 2018). From this, a calculation chain is constructed, which lists all the cells that contain formulas in the order in which they should be calculated. The construction of calculation tree and chain depends from cell references. The recalculation mechanism cannot provide for many calls to the same SDF in such a way that it would return the output from SDF for each of the calls. Additionally, since SDFs should be callable by their name and not by reference, we cannot directly rely on Excel recalculation. Namely, if we call SDF by name and not by reference, then it will not be included in recalculation, since it is not included in the calculation chain. Thus, the execution (recalculation) of SDFs should be controlled independently from Excel usual recalculation.

#### Implementation

SDF add-in is implemented by three VBA modules:
* WSFunction module exposes user-defined functions ModuleInput(), ModuleOutput(), ModuleUse(), and ModuleUseRangeInputs()
* Recalculation module provides functionality for SDF recalculation
* Ribbon module implements the add-in's ribbon user interface hooks

Additionally, add-in defines four classes:
* objModule class that represents a particular SDF
* objInput class represents SDF's input
* objOutput class represents SDF's output
* objModuleUse class represents the use (a call) to a module

objModule object has public properties which describe its name, a collection of module inputs, a collection of module outputs, a collection of module uses, and the range of cells of the module for recalculation purposes. Module WSFunction has a public variable which is a collection of all modules.

#### References

Excel Recalculation, 2018. https://docs.microsoft.com/en-us/office/client-developer/excel/excel-recalculation
Jones, S.P., Blackwell, A., Burnett, M., 2003. A User-Centred Approach to Functions in Excel. SIGPLAN Not. 38, 165–176. https://doi.org/10.1145/944746.944721
Sestoft, P., Sørensen, J.Z., 2013. Sheet-Defined Functions: Implementation and Initial Evaluation, in: Dittrich, Y., Burnett, M., Mørch, A., Redmiles, D. (Eds.), End-User Development. Springer Berlin Heidelberg, Berlin, Heidelberg, pp. 88–103.