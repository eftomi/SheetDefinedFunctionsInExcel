VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInsertSDF 
   Caption         =   "Insert Sheet Defined Function"
   ClientHeight    =   7275
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9780
   OleObjectBlob   =   "frmInsertSDF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInsertSDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblInputValue01_Click()
    CheckInput Me.lblInputValue01
End Sub
Private Sub lblInputValue02_Click()
    CheckInput Me.lblInputValue02
End Sub
Private Sub lblInputValue03_Click()
    CheckInput Me.lblInputValue03
End Sub
Private Sub lblInputValue04_Click()
    CheckInput Me.lblInputValue04
End Sub
Private Sub lblInputValue05_Click()
    CheckInput Me.lblInputValue05
End Sub
Private Sub lblInputValue06_Click()
    CheckInput Me.lblInputValue06
End Sub
Private Sub lblInputValue07_Click()
    CheckInput Me.lblInputValue07
End Sub

Private Sub CheckInput(liv As MSForms.Label)
    Dim inpVal As Variant
    
    inpVal = Application.InputBox("Set this input to be:", "Enter Input Value", lblInputValue01.Caption, , , , , 0)
    If (VarType(inpVal) <> vbBoolean) Then
        inpVal = Application.ConvertFormula(inpVal, xlR1C1, xlA1)
        If Left(inpVal, 1) = "=" Then
            inpVal = Right(inpVal, Len(inpVal) - 1)
        End If
        liv.Caption = inpVal
    End If
End Sub

Private Sub lblInputValue01_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblInputValue01.BackColor = &H80000003
    lblInputValue01.BorderColor = &HA9A9A9
End Sub

Private Sub lblInputValue02_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblInputValue02.BackColor = &H80000003
    lblInputValue02.BorderColor = &HA9A9A9
End Sub
Private Sub lblInputValue03_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblInputValue03.BackColor = &H80000003
    lblInputValue03.BorderColor = &HA9A9A9
End Sub
Private Sub lblInputValue04_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblInputValue04.BackColor = &H80000003
    lblInputValue04.BorderColor = &HA9A9A9
End Sub
Private Sub lblInputValue05_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblInputValue05.BackColor = &H80000003
    lblInputValue05.BorderColor = &HA9A9A9
End Sub
Private Sub lblInputValue06_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblInputValue06.BackColor = &H80000003
    lblInputValue06.BorderColor = &HA9A9A9
End Sub
Private Sub lblInputValue07_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblInputValue07.BackColor = &H80000003
    lblInputValue07.BorderColor = &HA9A9A9
End Sub

Private Sub UserForm_Activate()
    Dim mdl As objModule
    Dim inpt As objInput
    Dim outp As objOutput
    Dim i As Integer
    
    Set AllModules = Nothing
    Set AllModules = New objModules
    Application.CalculateFull
    For Each mdl In WSFunctions.AllModules.Collection
        Me.lstSDFs.AddItem mdl.Name
    Next
    
End Sub

Private Sub btnOK_Click()
    Dim actCell As Range
    Dim frmla As String
    Dim livs(7) As Object
    Dim i As Integer

    Set actCell = Application.ActiveCell
    frmla = "=ModuleUse("""
    
    frmla = frmla & Me.lstSDFs.Value & """"
    
    If Me.cmbOutputs.Value <> "_default" Then
        frmla = frmla & ",""" & Me.cmbOutputs.Value & ""","
    Else
        frmla = frmla & ",,"
    End If
    
    GetLivs livs
    For i = 1 To UBound(livs)
        If livs(i).Visible Then
            If livs(i).Tag <> "_default" Then
                frmla = frmla & """" & livs(i).Tag & ""","
            End If
            frmla = frmla & livs(i).Caption & ","
        End If
    Next
    frmla = Left(frmla, Len(frmla) - 1) & ")"
    
    actCell.Formula = frmla
    
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    btnCancel.BackColor = &H80000003
    btnCancel.BorderColor = &HA9A9A9
End Sub


Private Sub btnOK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    btnOK.BackColor = &H80000003
    btnOK.BorderColor = &HA9A9A9
End Sub

Private Sub lstSDFs_Click()
    Dim mdl As objModule
    Dim inpt As objInput
    Dim outp As objOutput
    Dim first As String

    Dim lbls(7) As Object
    Dim livs(7) As Object
    Dim i As Integer
    
    GetLbls lbls
    GetLivs livs
    HideAllLbls lbls
    HideAllLivs livs
    
    Me.cmbOutputs.Clear

    Set mdl = WSFunctions.AllModules.item(Me.lstSDFs.Value)
    For Each outp In mdl.ModuleOutputs.Collection
        Me.cmbOutputs.AddItem outp.Name
        If first = "" Then
            first = outp.Name
        End If
    Next
    Me.cmbOutputs.Value = first
    
    For Each inpt In mdl.ModuleInputs.Collection
        i = i + 1
        lbls(i).Caption = inpt.Name
        lbls(i).Visible = True
        livs(i).Tag = inpt.Name
        livs(i).Visible = True
    Next
    If i > 0 Then Me.lblSetTheInputs.Visible = True
End Sub



Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    btnOK.BackColor = &HE0E0E0
    btnOK.BorderColor = &H808080
    btnCancel.BackColor = &HE0E0E0
    btnCancel.BorderColor = &H808080
    
    lblInputValue01.BackColor = &HE0E0E0
    lblInputValue01.BorderColor = &H808080
    lblInputValue02.BackColor = &HE0E0E0
    lblInputValue02.BorderColor = &H808080
    lblInputValue03.BackColor = &HE0E0E0
    lblInputValue03.BorderColor = &H808080
    lblInputValue04.BackColor = &HE0E0E0
    lblInputValue04.BorderColor = &H808080
    lblInputValue05.BackColor = &HE0E0E0
    lblInputValue05.BorderColor = &H808080
    lblInputValue06.BackColor = &HE0E0E0
    lblInputValue06.BorderColor = &H808080
    lblInputValue07.BackColor = &HE0E0E0
    lblInputValue07.BorderColor = &H808080
End Sub

Private Sub GetLbls(ByRef lbls)
    Set lbls(1) = Me.lblInputName01
    Set lbls(2) = Me.lblInputName02
    Set lbls(3) = Me.lblInputName03
    Set lbls(4) = Me.lblInputName04
    Set lbls(5) = Me.lblInputName05
    Set lbls(6) = Me.lblInputName06
    Set lbls(7) = Me.lblInputName07
End Sub

Private Sub GetLivs(ByRef livs)
    Set livs(1) = Me.lblInputValue01
    Set livs(2) = Me.lblInputValue02
    Set livs(3) = Me.lblInputValue03
    Set livs(4) = Me.lblInputValue04
    Set livs(5) = Me.lblInputValue05
    Set livs(6) = Me.lblInputValue06
    Set livs(7) = Me.lblInputValue07
End Sub
Private Sub HideAllLbls(ByRef lbls)
    Dim i As Integer
    
    For i = 1 To UBound(lbls)
        lbls(i).Visible = False
    Next
    Me.lblSetTheInputs.Visible = False
End Sub
Private Sub HideAllLivs(ByRef livs)
    Dim i As Integer
    
    For i = 1 To UBound(livs)
        livs(i).Visible = False
    Next
End Sub

