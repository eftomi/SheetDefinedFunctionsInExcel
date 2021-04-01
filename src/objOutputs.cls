VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objOutputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Outputs collection
Option Explicit

Private colOutputs As Collection

Private Sub Class_Initialize()
    Set colOutputs = New Collection
End Sub

Public Function Add(Name As String) As objOutput
    Dim item As New objOutput
    
    item.Name = Name
    colOutputs.Add item, Name
    
    Set Add = item
End Function

Public Property Get Count() As Long
    Count = colOutputs.Count
End Property

Public Sub Remove(IDOrName As Variant)
    On Error GoTo DoesntExist
    colOutputs.Remove IDOrName
    Exit Sub
DoesntExist:
End Sub

Public Property Get item(IDOrName As Variant) As objOutput
    On Error GoTo DoesntExist
    Set item = colOutputs.item(IDOrName)
    
    Exit Property
DoesntExist:
    Set item = Nothing
End Property

Public Property Get Collection() As Collection
    'kako se to naredi lepše:
    'http://www.utteraccess.com/forum/custom-collection-VBA-t1750880.html
    Set Collection = colOutputs
End Property

