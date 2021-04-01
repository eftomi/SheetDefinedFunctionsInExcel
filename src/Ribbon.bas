Attribute VB_Name = "Ribbon"
Option Explicit

Sub rxCalculateSDFs_onAction(control As IRibbonControl)
    Recalculation.RecalculateModules
End Sub

Sub rxListAllSDFs_onAction(control As IRibbonControl)
    ListAllSDFs
End Sub

