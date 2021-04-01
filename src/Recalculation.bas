Attribute VB_Name = "Recalculation"
Public Sub RecalculateModules()
    Dim mdl As objModule
    Dim mdlu As objModuleUse
    Dim calcMode As Variant
    
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    
    Set AllModules = New objModules
    Application.CalculateFull
    For Each mdl In WSFunctions.AllModules.Collection
        For Each mdlu In mdl.ModuleUses.Collection
            RecalculateModule mdlu.SourceCells  'execute ModuleUse() to enumerate all inputs
            RecalculateModule mdl.ModuleRange   'execute module
            RecalculateModule mdlu.SourceCells  'execute ModuleUse() to read the result
        Next
    Next
    
    Application.ScreenUpdating = True
    Application.Calculation = calcMode
End Sub

Public Sub RecalculateModule(ModuleRange As Range)
    Dim cell As Range
    
    Application.Calculation = xlManual      'switch manual calculation mode
    ModuleRange.Calculate                   'recalculate just the module range
    Application.Calculation = xlAutomatic   'switch back to previous calculation mode
End Sub

Public Sub ListAllSDFs()
    Dim mdl As objModule
    Dim inpt As objInput
    Dim outp As objOutput
    Dim currCell As Range
    Dim offst As Integer
    
    offst = 0
    Set currCell = ActiveCell
    Set AllModules = New objModules
    Application.CalculateFull
    For Each mdl In WSFunctions.AllModules.Collection
        currCell.offset(offst).Formula = "SDF:"
        currCell.offset(offst, 1).Formula = mdl.Name
        offst = offst + 1
        For Each inpt In mdl.ModuleInputs.Collection
            currCell.offset(offst).Formula = "input:"
            currCell.offset(offst, 1).Formula = inpt.Name
            offst = offst + 1
        Next
        For Each outp In mdl.ModuleOutputs.Collection
            currCell.offset(offst).Formula = "output:"
            currCell.offset(offst, 1).Formula = outp.Name
            offst = offst + 1
        Next
    Next
End Sub

