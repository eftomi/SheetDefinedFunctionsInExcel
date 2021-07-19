Attribute VB_Name = "Ribbon"
Option Explicit

Sub rxCalculateSDFs_onAction(control As IRibbonControl)
    Recalculation.RecalculateModules
End Sub

Sub rxUseSDFs_onAction(control As IRibbonControl)
    InsertSDF
End Sub

Private Sub InsertSDF()
    Dim frmISDF As frmInsertSDF
    
    Set frmISDF = New frmInsertSDF
    frmISDF.Show
End Sub

