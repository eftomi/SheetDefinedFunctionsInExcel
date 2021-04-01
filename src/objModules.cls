VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Collection of modules
Option Explicit

Private colModules As Collection

Private Sub Class_Initialize()
    Set colModules = New Collection
End Sub

Public Function Add(Name As String) As objModule
    Dim item As New objModule
    
    item.Name = Name
    colModules.Add item, Name
    
    Set Add = item
End Function

Public Property Get Count() As Long
    Count = colModules.Count
End Property

Public Sub Remove(IDOrName As Variant)
    On Error GoTo DoesntExist
    colModules.Remove IDOrName
    Exit Sub
DoesntExist:
End Sub

Public Property Get item(IDOrName As Variant) As objModule
    On Error GoTo DoesntExist
    Set item = colModules.item(IDOrName)
    
    Exit Property
DoesntExist:
    Set item = Nothing
End Property

Public Property Get Collection() As Collection
    'kako se to naredi lepše:
    'http://www.utteraccess.com/forum/custom-collection-VBA-t1750880.html
    Set Collection = colModules
End Property


