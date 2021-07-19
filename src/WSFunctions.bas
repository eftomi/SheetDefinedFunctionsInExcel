Attribute VB_Name = "WSFunctions"
Option Explicit

Public AllModules As New objModules         'placeholder for all module objects

Public Function ModuleInput(ModuleName As String, Optional InputName As String = "_default", Optional InitialValue As Variant = 0, Optional EnforceMyInputValues = False) As Variant
Attribute ModuleInput.VB_Description = "Used to define one module input - a cell with this function will receive an argument value from a module call during model execution. You must declare input's name if a module has two or more inputs."
Attribute ModuleInput.VB_ProcData.VB_Invoke_Func = " \n20"
'Used to define specific module input.

    Dim mdl As objModule
    Dim inpt As objInput
    
    Set mdl = GetOrCreateModule(ModuleName)
    
    Set inpt = mdl.ModuleInputs.item(InputName)
    If inpt Is Nothing Then
        Set inpt = mdl.ModuleInputs.Add(InputName)
        inpt.Value = InitialValue
    End If
    
    If TypeName(Application.Caller) = "Range" Then
        Set inpt.ModuleRangeInput = Application.Caller
    Else
        ModuleInput = CVErr(xlErrNA)
    End If
    
    If EnforceMyInputValues Then
        inpt.Value = InitialValue
        ModuleInput = inpt.Value
    Else
        ModuleInput = inpt.Value
    End If
End Function

Public Function ModuleOutput(ModuleName As String, FormulaDefinition As Variant, Optional OutputName As String = "_default") As Variant
Attribute ModuleOutput.VB_Description = "Used to define one module output - a cell with this function will send the result of the module to the module call during model execution. You must declare output's name if a module has two or more outputs."
Attribute ModuleOutput.VB_ProcData.VB_Invoke_Func = " \n20"
'Used to define specific module output.
    Dim mdl As objModule
    Dim outp As objOutput
    
    Set mdl = GetOrCreateModule(ModuleName)
    
    Set outp = mdl.ModuleOutputs.item(OutputName)
    If outp Is Nothing Then
        Set outp = mdl.ModuleOutputs.Add(OutputName)
    End If
    
    outp.Value = FormulaDefinition
    If TypeName(Application.Caller) = "Range" Then
        Set outp.ModuleRangeOutput = Application.Caller
    Else
        ModuleOutput = CVErr(xlErrNA)
    End If
    ModuleOutput = outp.Value
End Function


Public Function ModuleUse(ModuleName As String, ParamArray OutputNameAndInputs() As Variant) As Variant
Attribute ModuleUse.VB_Description = "Performs a module call. Module body should be defined beforehand by at least one ModuleOutput() function, and one or more ModuleInput() if it accepts input arguments."
Attribute ModuleUse.VB_ProcData.VB_Invoke_Func = " \n20"
'Used to issue a module call, requesting the named output, and setting all the required module inputs
'Each inputs is a pair of name, value optional parameters, as required by module
    Dim mdl As objModule
    Dim inpt As objInput
    Dim inptI As Integer
    Dim InputName As Variant
    Dim InputValue As Variant
    Dim OutputName As Variant
    
    Dim outp As objOutput
    
    Dim mdlu As objModuleUse
    Dim mdluID As String
    Dim ourCaller As Range
    
    Set mdl = GetOrCreateModule(ModuleName)
    
    If UBound(OutputNameAndInputs) = -1 Then            'call with =ModuleUse("myModule")
        OutputName = "_default"
    ElseIf UBound(OutputNameAndInputs) = 0 Then         'call with =ModuleUse("myModule", "myOutput")
        OutputName = OutputNameAndInputs(0)
    ElseIf UBound(OutputNameAndInputs) = 1 Then
        If IsMissing(OutputNameAndInputs(0)) Then
            OutputName = "_default"
        Else
            OutputName = OutputNameAndInputs(0)
        End If
        
        InputName = "_default"
        InputValue = OutputNameAndInputs(1)
        
        Set inpt = mdl.ModuleInputs.item(InputName)
        If Not inpt Is Nothing Then
            inpt.Value = InputValue
        End If
    ElseIf UBound(OutputNameAndInputs) = 2 And IsMissing(OutputNameAndInputs(1)) Then   'call with =ModuleUse("myModule", , , A1:A4)
        If IsMissing(OutputNameAndInputs(0)) Then
            OutputName = "_default"
        Else
            OutputName = OutputNameAndInputs(0)
        End If
        
        InputName = "_default"
        InputValue = OutputNameAndInputs(2)
        
        Set inpt = mdl.ModuleInputs.item(InputName)
        If Not inpt Is Nothing Then
            inpt.Value = InputValue
        End If
    ElseIf UBound(OutputNameAndInputs) Mod 2 = 0 Then   'call with =ModuleUse("myModule", , "InputData1", A1:A4, "InputData2", B1:B4)
        If IsMissing(OutputNameAndInputs(0)) Then
            OutputName = "_default"
        Else
            OutputName = OutputNameAndInputs(0)
        End If
        For inptI = 1 To UBound(OutputNameAndInputs) Step 2
            If VarType(OutputNameAndInputs(inptI)) = vbString Then
                InputName = OutputNameAndInputs(inptI)
                InputValue = OutputNameAndInputs(inptI + 1)
            
                Set inpt = mdl.ModuleInputs.item(InputName)
                If Not inpt Is Nothing Then
                    inpt.Value = InputValue
                End If
            Else
                ModuleUse = CVErr(xlErrValue)
                Exit Function
            End If
        Next
    Else
        ModuleUse = CVErr(xlErrValue)
        Exit Function
    End If
    
    If TypeName(Application.Caller) = "Range" Then
        Set ourCaller = Application.Caller
        'mdluID = "'" & ourCaller.Parent.Name & "'!" & ourCaller.address(External:=False)
        mdluID = ourCaller.address(, , , True)
        
        Set mdlu = mdl.ModuleUses.item(mdluID)
        If mdlu Is Nothing Then
            Set mdlu = mdl.ModuleUses.Add(mdluID)
        End If
        Set mdlu.SourceCells = ourCaller
    End If
    
    Set outp = mdl.ModuleOutputs.item(OutputName)
    If Not outp Is Nothing Then
        ModuleUse = outp.Value
    Else
        ModuleUse = CVErr(xlErrNA)
    End If
End Function

Public Function ModuleUseRangeInputs(ModuleName As String, OutputName As String, InputNames As Range, InputValues As Range) As Variant
Attribute ModuleUseRangeInputs.VB_Description = "Performs a module call. Module body should be defined beforehand by at least one ModuleOutput() function, and one or more ModuleInput() if it accepts input arguments."
Attribute ModuleUseRangeInputs.VB_ProcData.VB_Invoke_Func = " \n20"
'Used to issue a module call, requesting the named output, and setting all the required module inputs
'Inputs are defined as two ranges - a range of names and a range of values
    Dim mdl As objModule
    Dim inpt As objInput
    Dim inptI As Integer
    Dim InputName As Variant
    Dim InputValue As Variant
    
    Dim outp As objOutput
    
    Dim mdlu As objModuleUse
    Dim mdluID As String
    Dim ourCaller As Range
    
    Set mdl = GetOrCreateModule(ModuleName)
    
    For inptI = 1 To InputNames.Count
        If VarType(InputNames(inptI)) = vbString Then
            InputName = InputNames(inptI)
            InputValue = InputValues(inptI)
            
            Set inpt = mdl.ModuleInputs.item(InputName)
            If Not inpt Is Nothing Then
                inpt.Value = InputValue
            End If
        Else
            ModuleUseRangeInputs = CVErr(xlErrValue)
            Exit Function
        End If
    Next
    
    If TypeName(Application.Caller) = "Range" Then
        Set ourCaller = Application.Caller
        'mdluID = "'" & ourCaller.Parent.Name & "'!" & ourCaller.address(External:=False)
        mdluID = ourCaller.address(, , , True)
        
        Set mdlu = mdl.ModuleUses.item(mdluID)
        If mdlu Is Nothing Then
            Set mdlu = mdl.ModuleUses.Add(mdluID)
        End If
        Set mdlu.SourceCells = ourCaller
    End If
        
    Set outp = mdl.ModuleOutputs.item(OutputName)
    If Not outp Is Nothing Then
        ModuleUseRangeInputs = outp.Value
    Else
        ModuleUseRangeInputs = CVErr(xlErrNA)
    End If
End Function


Private Function GetOrCreateModule(ModuleName As String) As objModule
'Used to get an existing module from AllModules placeholder or create a new one
    Dim mdl As objModule

    Set mdl = AllModules.item(ModuleName)
    If mdl Is Nothing Then
        Set mdl = AllModules.Add(ModuleName)
    End If
    
    Set GetOrCreateModule = mdl
End Function


