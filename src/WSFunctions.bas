Attribute VB_Name = "WSFunctions"
Option Explicit

Public AllModules As New objModules         'placeholder for all module objects

Public Function ModuleInput(ModuleName As String, ModuleRange As Range, inputName As String, InitialValue As Variant, Optional UseMyInputValues = False) As Variant
'Used to define specific module input.
'ModuleRange is used for partial evaluation/recalculation of particular module, without influencing other parts of workbook, including other uses of the same module

    Dim mdl As objModule
    Dim inpt As objInput
    
    Set mdl = GetOrCreateModule(ModuleName)
    
    Set mdl.ModuleRange = ModuleRange
    
    Set inpt = mdl.ModuleInputs.item(inputName)
    If inpt Is Nothing Then
        Set inpt = mdl.ModuleInputs.Add(inputName)
        inpt.Value = InitialValue
    End If
    
    If UseMyInputValues Then
        inpt.Value = InitialValue
        ModuleInput = inpt.Value
    Else
        ModuleInput = inpt.Value
    End If
End Function

Public Function ModuleOutput(ModuleName As String, OutputName As String, FormulaDefinition As Variant) As Variant
'Used to define specific module output.
    Dim mdl As objModule
    Dim outp As objOutput
    
    Set mdl = GetOrCreateModule(ModuleName)
    
    Set outp = mdl.ModuleOutputs.item(OutputName)
    If outp Is Nothing Then
        Set outp = mdl.ModuleOutputs.Add(OutputName)
    End If
    
    outp.Value = FormulaDefinition
    ModuleOutput = outp.Value
End Function

Public Function ModuleUse(ModuleName As String, OutputName As String, ParamArray Inputs() As Variant) As Variant
'Used to issue a module call, requesting the named output, and setting all the required module inputs
'Each inputs is a pair of name, value optional parameters, as required by module
    Dim mdl As objModule
    Dim inpt As objInput
    Dim inptI As Integer
    Dim inputName As Variant
    Dim inputValue As Variant
    
    Dim outp As objOutput
    
    Dim mdlu As objModuleUse
    Dim mdluID As String
    Dim ourCaller As Range
    
    Set mdl = GetOrCreateModule(ModuleName)
    
    For inptI = LBound(Inputs) To UBound(Inputs) Step 2
        If VarType(Inputs(inptI)) = vbString Then
            inputName = Inputs(inptI)
            inputValue = Inputs(inptI + 1)
            
            Set inpt = mdl.ModuleInputs.item(inputName)
            If Not inpt Is Nothing Then
                inpt.Value = inputValue
            End If
        Else
            ModuleUse = CVErr(xlErrValue)
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
        ModuleUse = outp.Value
    Else
        ModuleUse = CVErr(xlErrNA)
    End If
End Function


Public Function ModuleUseRangeInputs(ModuleName As String, OutputName As String, InputNames As Range, InputValues As Range) As Variant
'Used to issue a module call, requesting the named output, and setting all the required module inputs
'Inputs are defined as two ranges - a range of names and a range of values
    Dim mdl As objModule
    Dim inpt As objInput
    Dim inptI As Integer
    Dim inputName As Variant
    Dim inputValue As Variant
    
    Dim outp As objOutput
    
    Dim mdlu As objModuleUse
    Dim mdluID As String
    Dim ourCaller As Range
    
    Set mdl = GetOrCreateModule(ModuleName)
    
    For inptI = 1 To InputNames.Count
        If VarType(InputNames(inptI)) = vbString Then
            inputName = InputNames(inptI)
            inputValue = InputValues(inptI)
            
            Set inpt = mdl.ModuleInputs.item(inputName)
            If Not inpt Is Nothing Then
                inpt.Value = inputValue
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


