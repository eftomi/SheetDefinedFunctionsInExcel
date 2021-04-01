VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objModuleUses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ModuleUses collection
Option Explicit

Private colModuleUses As Collection

Private Sub Class_Initialize()
    Set colModuleUses = New Collection
End Sub

Public Function Add(Name As String) As objModuleUse
    Dim item As New objModuleUse
    
    item.Name = Name
    colModuleUses.Add item, Name
    
    Set Add = item
End Function

Public Property Get Count() As Long
    Count = colModuleUses.Count
End Property

Public Sub Remove(IDOrName As Variant)
    On Error GoTo DoesntExist
    colModuleUses.Remove IDOrName
    Exit Sub
DoesntExist:
End Sub

Public Property Get item(IDOrName As Variant) As objModuleUse
    On Error GoTo DoesntExist
    Set item = colModuleUses.item(IDOrName)
    
    Exit Property
DoesntExist:
    Set item = Nothing
End Property

Public Property Get Collection() As Collection
    'kako se to naredi lepše:
    'http://www.utteraccess.com/forum/custom-collection-VBA-t1750880.html
    Set Collection = colModuleUses
End Property



