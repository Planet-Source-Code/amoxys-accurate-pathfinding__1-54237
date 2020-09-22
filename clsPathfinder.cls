VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPathfinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Sometimes POINTAPI is used to define Obstacle:Vertex
'and sometimes just coords.


Dim StartPos As POINTAPI 'Obs(-2)
Dim EndPos As POINTAPI 'Obs(-3)

Dim Obs() As clsPolygon
Dim BestPath() As POINTAPI 'coords of the best path
Dim BestPathLen As Double
Dim PathExists As Boolean
Dim PathCalculated As Boolean
Dim CalcTime As Long

Public Sub Init()
ReDim Obs(0)
ReDim BestPath(0)
PathCalculated = False
End Sub

Private Sub Class_Initialize()
Init
End Sub

Public Sub DrawPlgs(hdc As Long)
Dim I As Long

For I = 0 To Count - 1
   Obs(I).Draw hdc
Next I
End Sub

Public Sub DrawPath(hdc As Long)
If Not PathExists Or Not PathCalculated Then Exit Sub
Polyline hdc, BestPath(0), UBound(BestPath)
End Sub

Public Function blockedBySides(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Boolean
'Calculates whether the line crosses the boundaries of any obstacle

Dim I As Long, I2 As Long

For I = 0 To UBound(Obs) - 1
   For I2 = 0 To Obs(I).Count - 2
      If areLinesCrossing(X1, Y1, X2, Y2, Obs(I).Vtx(I2).X, Obs(I).Vtx(I2).Y, Obs(I).Vtx(I2 + 1).X, Obs(I).Vtx(I2 + 1).Y) Then
         blockedBySides = True
         Exit Function
      End If
   Next I2
   If areLinesCrossing(X1, Y1, X2, Y2, Obs(I).Vtx(I2).X, Obs(I).Vtx(I2).Y, Obs(I).Vtx(0).X, Obs(I).Vtx(0).Y) Then
      blockedBySides = True
      Exit Function
   End If
Next I

blockedBySides = False
End Function

Public Function blockedBySolid(OrgPlg As clsPolygon, OrgPntIndex As Long, DstX As Long, DstY As Long) As Boolean
'Calculates whether the line goes through the obstacle
'This function can only be used if if the line starts at a corner from this obstacle

Dim Angle1 As Double, Angle2 As Double, Angle3 As Double

Angle1 = getAngle(OrgPlg.Vtx(OrgPlg.modulateIndex(OrgPntIndex - 1)).X - OrgPlg.Vtx(OrgPntIndex).X, OrgPlg.Vtx(OrgPlg.modulateIndex(OrgPntIndex - 1)).Y - OrgPlg.Vtx(OrgPntIndex).Y)
Angle2 = getAngle(OrgPlg.Vtx(OrgPlg.modulateIndex(OrgPntIndex + 1)).X - OrgPlg.Vtx(OrgPntIndex).X, OrgPlg.Vtx(OrgPlg.modulateIndex(OrgPntIndex + 1)).Y - OrgPlg.Vtx(OrgPntIndex).Y)
Angle3 = getAngle(DstX - OrgPlg.Vtx(OrgPntIndex).X, DstY - OrgPlg.Vtx(OrgPntIndex).Y)

If Angle3 < Angle2 Then
   If Angle1 < Angle2 Then
      If Angle3 > Angle1 Then
         blockedBySolid = True
         Exit Function
      End If
   Else
      blockedBySolid = True
      Exit Function
   End If
Else
   If Angle1 > Angle2 Then
      If Angle1 < Angle3 Then
         blockedBySolid = True
         Exit Function
      End If
   End If
End If

blockedBySolid = False
End Function

Public Sub findPath()
'Searching from EndPos to StartPos
'so the "predecessor trace" starts at StartPos

Dim I As Long, I2 As Long
Dim ToDo() As clsPoint
Dim CurPathLen As Double
Dim bestPdc As clsPoint 'best predecessor of StartPos
Dim CurPnt As clsPoint
CalcTime = GetTickCount

Erase BestPath
BestPathLen = -1
PathCalculated = True

For I = 0 To UBound(Obs) - 1
   If Obs(I).HitTest(StartPos.X, StartPos.Y) Then
      'StartPos IN a polygon
      PathExists = False
      CalcTime = GetTickCount - CalcTime
      Exit Sub
   End If
   If Obs(I).HitTest(EndPos.X, EndPos.Y) Then
      'EndPos IN a polygon
      PathExists = False
      CalcTime = GetTickCount - CalcTime
      Exit Sub
   End If
Next I

If Not blockedBySides(StartPos.X, StartPos.Y, EndPos.X, EndPos.Y) Then
   'free way
   PathExists = True
   ReDim BestPath(2)
   BestPath(0) = StartPos
   BestPath(1) = EndPos
   BestPathLen = getDistance(StartPos.X, StartPos.Y, EndPos.X, EndPos.Y)
   CalcTime = GetTickCount - CalcTime
   Exit Sub
End If

ReDim ToDo(0)

'note vertices which are accessible from EndPos
For I = 0 To UBound(Obs) - 1
   For I2 = 0 To Obs(I).Count - 1
      Set Obs(I).Vtx(I2).Pdc = Nothing
      If blockedBySides(EndPos.X, EndPos.Y, Obs(I).Vtx(I2).X, Obs(I).Vtx(I2).Y) Then
         Obs(I).Vtx(I2).PathLen = -1
      Else
         Obs(I).Vtx(I2).PathLen = getDistance(EndPos.X, EndPos.Y, Obs(I).Vtx(I2).X, Obs(I).Vtx(I2).Y)
         Set ToDo(UBound(ToDo)) = Obs(I).Vtx(I2)
         ReDim Preserve ToDo(UBound(ToDo) + 1)
      End If
   Next I2
Next I

Do Until UBound(ToDo) = 0
   'test whether StartPos is accessible
   If blockedBySolid(ToDo(0).Parent, ToDo(0).Parent.getIndex(ToDo(0)), StartPos.X, StartPos.Y) Then
   ElseIf Not blockedBySides(ToDo(0).X, ToDo(0).Y, StartPos.X, StartPos.Y) Then
      CurPathLen = ToDo(0).PathLen + getDistance(ToDo(0).X, ToDo(0).Y, StartPos.X, StartPos.Y)
      If (CurPathLen < BestPathLen) Or (BestPathLen = -1) Then
         BestPathLen = CurPathLen
         Set bestPdc = ToDo(0)
      Else
         GoTo tryNext
      End If
   End If
            
   'note connected vertices which are not too far away
   For I = 0 To ToDo(0).Count - 1
      CurPathLen = ToDo(0).PathLen + ToDo(0).Dist(I)
      If (BestPathLen < CurPathLen) And (BestPathLen > -1) Then
      ElseIf (ToDo(0).Connection(I).PathLen > CurPathLen) Or (ToDo(0).Connection(I).PathLen = -1) Then
         ToDo(0).Connection(I).PathLen = CurPathLen
         Set ToDo(0).Connection(I).Pdc = ToDo(0)
         Set ToDo(UBound(ToDo)) = ToDo(0).Connection(I)
         ReDim Preserve ToDo(UBound(ToDo) + 1)
      End If
   Next I
   
tryNext:
   For I = 0 To UBound(ToDo) - 2
      Set ToDo(I) = ToDo(I + 1)
   Next I
   ReDim Preserve ToDo(UBound(ToDo) - 1)
Loop


If BestPathLen = -1 Then
   'no path found
   PathExists = False
   ReDim BestPath(0)
   CalcTime = GetTickCount - CalcTime
   Exit Sub
End If

'trace it back (from StartPos to EndPos)
ReDim BestPath(1)
BestPath(0) = StartPos
Set CurPnt = bestPdc
Do Until CurPnt Is Nothing
   BestPath(UBound(BestPath)).X = CurPnt.X
   BestPath(UBound(BestPath)).Y = CurPnt.Y
   ReDim Preserve BestPath(UBound(BestPath) + 1)
   Set CurPnt = CurPnt.Pdc
Loop
BestPath(UBound(BestPath)) = EndPos
ReDim Preserve BestPath(UBound(BestPath) + 1)

PathExists = True
CalcTime = GetTickCount - CalcTime
End Sub

Public Sub Add(ByVal aPlg As clsPolygon)
Dim I As Long, I2 As Long, I3 As Long, I4 As Long

If Not aPlg.Valid Then Exit Sub

If Not isClockwise(aPlg) Then aPlg.Reverse

'Remove destroyed connections
For I = 0 To Count - 1
   For I2 = 0 To Obs(I).Count - 1
      I3 = 0
      Do Until I3 = Obs(I).Vtx(I2).Count
         For I4 = 0 To aPlg.Count - 2
            If areLinesCrossing(aPlg.Vtx(I4).X, aPlg.Vtx(I4).Y, aPlg.Vtx(I4 + 1).X, aPlg.Vtx(I4 + 1).Y, Obs(I).Vtx(I2).X, Obs(I).Vtx(I2).Y, Obs(I).Vtx(I2).Connection(I3).X, Obs(I).Vtx(I2).Connection(I3).Y) Then
               Obs(I).Vtx(I2).Connection(I3).DisConnectPnt Obs(I).Vtx(I2)
               Obs(I).Vtx(I2).DisConnectIndex I3
               GoTo testSame
            End If
         Next I4
         If areLinesCrossing(aPlg.Vtx(I4).X, aPlg.Vtx(I4).Y, aPlg.Vtx(0).X, aPlg.Vtx(0).Y, Obs(I).Vtx(I2).X, Obs(I).Vtx(I2).Y, Obs(I).Vtx(I2).Connection(I3).X, Obs(I).Vtx(I2).Connection(I3).Y) Then
            Obs(I).Vtx(I2).Connection(I3).DisConnectPnt Obs(I).Vtx(I2)
            Obs(I).Vtx(I2).DisConnectIndex I3
         Else
            I3 = I3 + 1
         End If
testSame:
      Loop
   Next I2
Next I
         
'Add obstacle
Set Obs(UBound(Obs)) = aPlg
ReDim Preserve Obs(UBound(Obs) + 1)

'Make new connections
For I = 0 To aPlg.Count - 1
   For I2 = 0 To Count - 1
      For I3 = 0 To Obs(I2).Count - 1
         'If aPlg.Vtx(I) Is Obs(I2).Vtx(I3) Then
         If blockedBySolid(Obs(Count - 1), I, Obs(I2).Vtx(I3).X, Obs(I2).Vtx(I3).Y) Then
         ElseIf blockedBySolid(Obs(I2), I3, aPlg.Vtx(I).X, aPlg.Vtx(I).Y) Then
         ElseIf Not blockedBySides(aPlg.Vtx(I).X, aPlg.Vtx(I).Y, Obs(I2).Vtx(I3).X, Obs(I2).Vtx(I3).Y) Then
            aPlg.Vtx(I).Connect Obs(I2).Vtx(I3)
            Obs(I2).Vtx(I3).Connect aPlg.Vtx(I)
         End If
      Next I3
   Next I2
Next I

PathCalculated = False
End Sub

Public Property Get Time() As Long
If PathCalculated Then Time = CalcTime
End Property

Public Function Count() As Long
Count = UBound(Obs)
End Function

Public Sub Remove(Index As Long)
Dim I As Long, I2 As Long, I3 As Long, I4 As Long
Dim PntCnt() As clsPoint

If Index < 0 Or Index >= Count Then
   Err.Raise 9
   Exit Sub
End If

If Count = 1 Then
   Init
   Exit Sub
End If

'Remove destroyed connections
ReDim PntCnt(0)
For I = 0 To Obs(Index).Count - 1
   For I2 = 0 To Obs(Index).Vtx(I).Count - 1
      For I3 = 0 To UBound(PntCnt) - 1
         If PntCnt(I3) Is Obs(Index).Vtx(I).Connection(I2) Then Exit For
      Next I3
      If I3 < UBound(PntCnt) Then
         Set PntCnt(UBound(PntCnt)) = Obs(Index).Vtx(I).Connection(I2)
         ReDim Preserve PntCnt(UBound(PntCnt) + 1)
      End If
   Next I2
Next I
For I = 0 To UBound(PntCnt) - 1
   PntCnt(I).DisConnectPlg Obs(Index)
Next I

'Remove obstacle
For I = Index To Count - 2
   Set Obs(I) = Obs(I + 1)
Next I
ReDim Preserve Obs(Count - 1)

'Make new connections
For I = 0 To Count - 1
   For I2 = 0 To Obs(I).Count - 1
      For I3 = I To Count - 1
         For I4 = IIf(I = I3, I2 + 1, 0) To Obs(I3).Count - 1
            If blockedBySolid(Obs(I), I2, Obs(I3).Vtx(I4).X, Obs(I3).Vtx(I4).Y) Then
            ElseIf blockedBySolid(Obs(I3), I4, Obs(I).Vtx(I2).X, Obs(I).Vtx(I2).Y) Then
            ElseIf Not blockedBySides(Obs(I).Vtx(I2).X, Obs(I).Vtx(I2).Y, Obs(I3).Vtx(I4).X, Obs(I3).Vtx(I4).Y) Then
               Obs(I).Vtx(I2).Connect Obs(I3).Vtx(I4)
               Obs(I3).Vtx(I4).Connect Obs(I).Vtx(I2)
            End If
         Next I4
      Next I3
   Next I2
Next I

PathCalculated = False
End Sub

Public Function HitTest(Index As Long, X As Long, Y As Long) As Boolean
'Calculates which obstacle is at the position

If Index < 0 Or Index >= Count Then
   Err.Raise 9
   Exit Function
End If
HitTest = Obs(Index).HitTest(X, Y)
End Function

Public Property Get StartPosX() As Variant
StartPosX = StartPos.X
End Property

Public Property Let StartPosX(ByVal vNewValue As Variant)
StartPos.X = vNewValue
PathCalculated = False
End Property

Public Property Get StartPosY() As Variant
StartPosY = StartPos.Y
End Property

Public Property Let StartPosY(ByVal vNewValue As Variant)
StartPos.Y = vNewValue
PathCalculated = False
End Property

Public Property Get EndPosX() As Variant
EndPosX = EndPos.X
End Property

Public Property Let EndPosX(ByVal vNewValue As Variant)
EndPos.X = vNewValue
PathCalculated = False
End Property

Public Property Get EndPosY() As Variant
EndPosY = EndPos.Y
End Property

Public Property Let EndPosY(ByVal vNewValue As Variant)
EndPos.Y = vNewValue
PathCalculated = False
End Property

Public Property Get PathLen() As Double
PathLen = BestPathLen
End Property
