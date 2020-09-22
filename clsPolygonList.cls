VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPolygonList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Plg() As clsPolygon

Public Sub Add(ByRef aPlg As clsPolygon)
Set Plg(UBound(Plg)) = aPlg
End Sub

Private Sub Class_Initialize()
Init
End Sub

Public Sub Init()
Erase Plg
ReDim Plg(0)
End Sub

Public Sub Remove(Index As Long)
If Index < 0 Or Index >= Count Then
   Err.Raise 9
   Exit Sub
End If

If Count = 1 Then
   Init
   Exit Sub
End If

For I = Index To Count - 2
   Set Plg(I) = Plg(I + 1)
Next I

Set Plg(I) = Nothing
ReDim Preserve Plg(UBound(Plg) - 1)
End Sub

Public Property Get Item(Index As Long) As clsPolygon
Set Item = Plg(Index)
End Property

Public Property Let Item(Index As Long, ByVal vNewValue As clsPolygon)

End Property
