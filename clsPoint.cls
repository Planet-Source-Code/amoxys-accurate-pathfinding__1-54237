VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Coords As POINTAPI
Dim ParentPlg As clsPolygon

Dim BestPathLen As Double
Dim bestPdc As clsPoint

Dim Distance() As Double
Dim Cnt() As clsPoint

Public Property Get X() As Long
X = Coords.X
End Property

Public Property Let X(ByVal vNewValue As Long)
Coords.X = vNewValue
End Property

Public Property Get Y() As Long
Y = Coords.Y
End Property

Public Property Let Y(ByVal vNewValue As Long)
Coords.Y = vNewValue
End Property

Public Sub Init()
Erase Distance
Erase Cnt
ReDim Distance(0)
ReDim Cnt(0)
End Sub

Public Sub Connect(ByRef aPnt As clsPoint)
Dim I As Long

If aPnt Is Me Then Exit Sub

For I = 0 To Count - 1
   If Cnt(I) Is aPnt Then Exit Sub
Next I

Set Cnt(Count) = aPnt
Distance(Count) = getDistance(Coords.X, Coords.Y, aPnt.X, aPnt.Y)
ReDim Preserve Cnt(Count + 1)
ReDim Preserve Distance(Count)
End Sub

Public Function Count() As Long
Count = UBound(Cnt)
End Function

Public Sub DisConnectPlg(ByRef aPlg As clsPolygon)
Dim I As Long, numFound As Long

For I = 0 To Count - 1
   If Cnt(I).Parent Is aPlg Then
      numFound = numFound + 1
   ElseIf numFound > 0 Then
      Set Cnt(I - numFound) = Cnt(I)
      Distance(I - numFound) = Distance(I)
   End If
Next I

ReDim Cnt(UBound(Cnt) - numFound)
ReDim Distance(Count)

Set Cnt(Count - 1) = Nothing
End Sub

Public Sub DisConnectIndex(Index As Long)
Dim I As Long

If Index < 0 Or Index >= Count Then
   Err.Raise 9
   Exit Sub
End If

For I = Index To Count - 2
   Set Cnt(I) = Cnt(I + 1)
   Distance(I) = Distance(I + 1)
Next I

ReDim Preserve Cnt(I)
ReDim Preserve Distance(I)

Set Cnt(Count) = Nothing
End Sub

Public Sub DisConnectPnt(ByRef aPnt As clsPoint)
Dim I As Long

For I = 0 To Count - 1
   If Cnt(I) Is aPnt Then
      DisConnectIndex (I)
      Exit Sub
   End If
Next I
End Sub

Private Sub Class_Initialize()
Init
End Sub

Public Property Get Parent() As clsPolygon
Set Parent = ParentPlg
End Property

Public Property Set Parent(ByRef vNewValue As clsPolygon)
Set ParentPlg = vNewValue
End Property

Public Property Get Connection(Index As Long) As clsPoint
If Index < 0 Or Index >= Count Then
   Err.Raise 9
   Exit Property
End If

Set Connection = Cnt(Index)
End Property

Public Property Get PathLen() As Double
PathLen = BestPathLen
End Property

Public Property Let PathLen(ByVal vNewValue As Double)
BestPathLen = vNewValue
End Property

Public Property Get Pdc() As clsPoint
Set Pdc = bestPdc
End Property

Public Property Let Pdc(ByVal vNewValue As clsPoint)
bestPdc = vNewValue
End Property

Public Property Set Pdc(ByRef vNewValue As clsPoint)
Set bestPdc = vNewValue
End Property

Public Property Get Dist(Index As Long) As Double
Dist = Distance(Index)

'Dist = getDistance(Coords.X, Coords.Y, Cnt(Index).X, Cnt(Index).Y)
End Property
