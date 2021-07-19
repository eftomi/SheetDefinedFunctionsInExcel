Attribute VB_Name = "Recalculation"
Option Explicit

Public Sub RecalculateModules()
    Dim mdl As objModule
    Dim mdlu As objModuleUse
    Dim mdli As objInput
    Dim mdlo As objOutput
    Dim calcMode As Variant
    
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    
    Set AllModules = Nothing
    Set AllModules = New objModules
    Application.CalculateFull
    For Each mdl In WSFunctions.AllModules.Collection
        For Each mdlu In mdl.ModuleUses.Collection
            RecalculateModule mdlu.SourceCells  'execute ModuleUse() to enumerate all inputs
            If mdl.ModuleInputs.Count > 0 Then
                For Each mdli In mdl.ModuleInputs.Collection
                    RecalculateModule mdli.ModuleRangeInput
                Next
            Else
                For Each mdlo In mdl.ModuleOutputs.Collection
                    RecalculateModule mdlo.ModuleRangeOutput
                Next
            End If
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


