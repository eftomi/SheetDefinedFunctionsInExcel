VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Inputs collection of objects
Option Explicit

Private colInputs As Collection

Private Sub Class_Initialize()
    Set colInputs = New Collection
End Sub

Public Function Add(Name As String) As objInput
    Dim item As New objInput
    
    item.Name = Name
    colInputs.Add item, Name
    
    Set Add = item
End Function

Public Property Get Count() As Long
    Count = colInputs.Count
End Property

Public Sub Remove(IDOrName As Variant)
    On Error GoTo DoesntExist
    colInputs.Remove IDOrName
    Exit Sub
DoesntExist:
End Sub

Public Property Get item(IDOrName As Variant) As objInput
    On Error GoTo DoesntExist
    Set item = colInputs.item(IDOrName)
    
    Exit Property
DoesntExist:
    Set item = Nothing
End Property

Public Property Get Collection() As Collection
    'kako se to naredi lepše:
    'http://www.utteraccess.com/forum/custom-collection-VBA-t1750880.html
    Set Collection = colInputs
End Property


