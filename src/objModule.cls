VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Module class
Option Explicit

Public Name As String
Public ModuleInputs As objInputs
Public ModuleOutputs As objOutputs
Public ModuleUses As objModuleUses
Public ModuleRange As Range

Private Sub Class_Initialize()
    Set ModuleInputs = New objInputs
    Set ModuleOutputs = New objOutputs
    Set ModuleUses = New objModuleUses
End Sub

