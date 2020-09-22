Attribute VB_Name = "modMain"
Option Explicit

Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public PI As Double        ' Calculated at Form_Load

Public Const ALTERNATE = 1 ' ALTERNATE and WINDING are
Public Const WINDING = 2   ' constants for FillMode.

Public Function getDistance(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Double
getDistance = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
End Function

Public Function isClockwise(Plg As clsPolygon) As Boolean
Dim I As Long, AngleA As Double, AngleB As Double, AngleSum As Double

If Plg.Count < 3 Then Exit Function

AngleA = getAngle(Plg.Vtx(1).X - Plg.Vtx(0).X, Plg.Vtx(1).Y - Plg.Vtx(0).Y)
For I = 1 To Plg.Count
   AngleB = getAngle(Plg.Vtx(Plg.modulateIndex(I + 1)).X - Plg.Vtx(Plg.modulateIndex(I)).X, Plg.Vtx(Plg.modulateIndex(I + 1)).Y - Plg.Vtx(Plg.modulateIndex(I)).Y)
   AngleA = AngleB - AngleA
   If Abs(AngleA) > PI Then AngleA = AngleA - Sgn(AngleA) * PI
   AngleSum = AngleSum + AngleA
   AngleA = AngleB
Next I

isClockwise = (AngleSum < 0)
End Function

Public Function getAngle(X As Double, Y As Double) As Double
If X = 0 And Y = 0 Then Exit Function

getAngle = aSin(X / Sqr(X ^ 2 + Y ^ 2))
If Y < 0 Then getAngle = PI - getAngle

If getAngle < 0 Then getAngle = getAngle + 2 * PI
End Function

Public Function aSin(X As Double) As Double
If Abs(X) = 1 Then
   aSin = Sgn(X) * PI / 2
   Exit Function
End If

aSin = Atn(X / Sqr(-X * X + 1))
End Function


