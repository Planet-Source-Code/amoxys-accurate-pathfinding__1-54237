VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPolygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'A polygon is not valid if it has less than 3 points.
'A polygon is not valid if it has self-intersections.

Option Explicit

Dim Coords() As POINTAPI
Dim Pnt() As clsPoint
Dim RgnPtr As Long

Private Sub Class_Initialize()
Init
End Sub

Public Sub Init()
Erase Pnt
ReDim Pnt(0)
Erase Coords
ReDim Coords(0)
deleteRegion
End Sub

Public Sub Add(X As Long, Y As Long)
Dim I As Long

For I = 0 To Count - 3
   If Coords(I).X = X And Coords(I).Y = Y Then Exit Sub
   If areLinesCrossing(Coords(Count - 1).X, Coords(Count - 1).Y, X, Y, Coords(I).X, Coords(I).Y, Coords(I + 1).X, Coords(I + 1).Y) Then Exit Sub
Next I
For I = I To Count - 1
   If Coords(I).X = X And Coords(I).Y = Y Then Exit Sub
Next I

Coords(Count).X = X
Coords(Count).Y = Y
Set Pnt(Count) = New clsPoint
Set Pnt(Count).Parent = Me
Pnt(Count).X = X
Pnt(Count).Y = Y
ReDim Preserve Pnt(Count + 1)
ReDim Preserve Coords(Count)

'polygons must have a surface
If Count < 3 Then Exit Sub

'delete region if polygon is not valid
For I = 1 To Count - 3
   If areLinesCrossing(X, Y, Coords(0).X, Coords(0).Y, Coords(I).X, Coords(I).Y, Coords(I + 1).X, Coords(I + 1).Y) Then
      deleteRegion
      Exit Sub
   End If
Next I

'create new region if polygon is valid
refreshRegion
End Sub

Public Sub Remove()
Dim I As Long

If Count = 0 Then Exit Sub

If Count = 1 Then
   'free ressources (don't know if it really works)
   Init
   Exit Sub
End If

ReDim Preserve Pnt(Count - 1)
Set Pnt(Count) = Nothing
ReDim Preserve Coords(Count)

'delete region if polygon is not valid
For I = 1 To Count - 3
   If areLinesCrossing(Coords(Count - 1).X, Coords(Count - 1).Y, Coords(0).X, Coords(0).Y, Coords(I).X, Coords(I).Y, Coords(I + 1).X, Coords(I + 1).Y) Then
      deleteRegion
      Exit Sub
   End If
Next I

'create new region if polygon is valid
refreshRegion
End Sub

Public Function Count() As Long
Count = UBound(Pnt)
End Function

Public Property Get Vtx(ByVal Index As Long) As clsPoint
If Index < 0 Or Index >= Count Then
   Err.Raise 9
   Exit Property
End If
Set Vtx = Pnt(Index)
End Property

Public Function modulateIndex(Index As Long) As Long
'very neat
modulateIndex = Index Mod Count
If modulateIndex < 0 Then modulateIndex = modulateIndex + Count
End Function

Public Sub Draw(DstHDC As Long)
Polygon DstHDC, Coords(0), Count
End Sub

Public Function HitTest(X As Long, Y As Long) As Boolean
'Calculates whether the polygon is at the position

If RgnPtr = 0 Then
   HitTest = False
Else
   HitTest = PtInRegion(RgnPtr, X, Y)
End If
End Function

Private Sub refreshRegion()
deleteRegion
If Count > 2 Then RgnPtr = CreatePolygonRgn(Coords(0), Count, ALTERNATE)
End Sub

Private Sub Class_Terminate()
Erase Pnt
Erase Coords
deleteRegion
End Sub

Public Property Get Region() As Long
Region = RgnPtr
End Property

Public Sub Reverse()
'Reverses the arrangement of points.
Dim I As Long, lastPnt As Long, tmpPnt As clsPoint

lastPnt = Count - 1
For I = 0 To Int(lastPnt / 2)
   Set tmpPnt = Pnt(I)
   Set Pnt(I) = Pnt(lastPnt - I)
   Set Pnt(lastPnt - I) = tmpPnt
   SwitchVar Coords(I).X, Coords(lastPnt - I).X
   SwitchVar Coords(I).Y, Coords(lastPnt - I).Y
Next I
'don't need to refreshRegion
End Sub

Private Sub deleteRegion()
If RgnPtr <> 0 Then
   DeleteObject RgnPtr
   RgnPtr = 0
End If
End Sub

Public Property Get Valid() As Boolean
Valid = (RgnPtr <> 0)
End Property

Public Function getIndex(ByRef aPnt As clsPoint) As Long
For getIndex = 0 To Count - 1
   If Pnt(getIndex) Is aPnt Then Exit Function
Next getIndex

Err.Raise vbObjectError
End Function
