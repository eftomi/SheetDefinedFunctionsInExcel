VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_AddinInstall()
    RegisterMyFunctions
End Sub

Private Sub Workbook_AddinUninstall()

End Sub

Private Sub RegisterMyFunctions()
    Application.MacroOptions _
        Macro:="ModuleInput", _
        Description:="Used to define one module input - a cell with this function will " _
            & "receive an argument value from a module call during model execution. " _
            & "You must declare input's name if a module has two or more inputs.", _
        Category:="SDFunctions", _
        ArgumentDescriptions:=Array( _
            "uniquely identifies a module (ties this input to the module definition).", _
            "declares the name of the input. Optional, if the module has only one input.", _
            "sets the initial input's value which is displayed in the cell with ModuleInput() function to be used during module body development (optional).", _
            "- if TRUE, the result of this function will be equal to InitialValue (to be used during module debugging; FALSE if ommited).")

    Application.MacroOptions _
        Macro:="ModuleOutput", _
        Description:="Used to define one module output - a cell with this function will " _
            & "send the result of the module to the module call during model execution. " _
            & "You must declare output's name if a module has two or more outputs.", _
        Category:="SDFunctions", _
        ArgumentDescriptions:=Array( _
            "uniquely identifies a module (ties this output to the module definition).", _
            "defines the output's value, normally by reference or by formula.", _
            "declares the name of the output. Optional, if the module has only one output.")

    Application.MacroOptions _
        Macro:="ModuleUse", _
        Description:="Performs a module call. Module body should be defined beforehand by " _
            & "at least one ModuleOutput() function, and one or more ModuleInput() if it accepts input arguments.", _
        Category:="SDFunctions", _
        ArgumentDescriptions:=Array( _
            "name of the module that is called.", _
            "declares the name of output that should be returned. If the module has only one output, its name can be ommited. " _
            & "Subsequent parameters are input names & input values, given in pairs. " _
            & "If module has one input, provide just its value without name.")

    Application.MacroOptions _
        Macro:="ModuleUseRangeInputs", _
        Description:="Performs a module call. Module body should be defined beforehand by " _
            & "at least one ModuleOutput() function, and one or more ModuleInput() if it accepts input arguments.", _
        Category:="SDFunctions", _
        ArgumentDescriptions:=Array( _
            "name of the module that is called.", _
            "declares the name of output that should be returned.", _
            "an array of input names.", _
            "an array of input values, in the same order as input names.")

End Sub
