Attribute VB_Name = "Bauersfeld"
Option Explicit

'The original version of this module was not written by me.
'It's not much "thinking" but "writing", so I looked for it on Google.
'Source: http://www.vb-fun.de/cgi-bin/loadframe.pl?ID=vb/tipps/tip0294.shtml
'Later I had to edit it a lot, to make it compatible.

Public Function areLinesCrossing(ByVal L1_X1 As Single, ByVal L1_Y1 As Single, ByVal L1_X2 As Single, ByVal L1_Y2 As Single, ByVal L2_X1 As Single, ByVal L2_Y1 As Single, ByVal L2_X2 As Single, ByVal L2_Y2 As Single) As Boolean
'Gradients
Dim L1_m As Single, L2_m As Single
'Intersections with the Y-axis
Dim L1_n As Single, L2_n As Single
'Intersection of the lines
Dim s    As Single
'Artificial variables for X and Y
Dim h_Y  As Single, h_X  As Single

'BoxCollision Y?
If (L1_Y1 <= L2_Y1 And L1_Y1 <= L2_Y2) And (L1_Y2 <= L2_Y1 And L1_Y2 <= L2_Y2) Then
   areLinesCrossing = False
   Exit Function
End If
If (L1_Y1 >= L2_Y1 And L1_Y1 >= L2_Y2) And (L1_Y2 >= L2_Y1 And L1_Y2 >= L2_Y2) Then
   areLinesCrossing = False
   Exit Function
End If

'Lines have same origin?
If (L1_X1 = L2_X1 And L1_Y1 = L2_Y1) Or (L1_X1 = L2_X2 And L1_Y1 = L2_Y2) Or (L1_X2 = L2_X1 And L1_Y2 = L2_Y1) Or (L1_X2 = L2_X2 And L1_Y2 = L2_Y2) Then
   areLinesCrossing = False
   Exit Function
End If

'Rotate if line is from right to left
If L1_X2 < L1_X1 Then
   SwitchVar L1_X1, L1_X2
   SwitchVar L1_Y1, L1_Y2
End If
If L2_X2 < L2_X1 Then
   SwitchVar L2_X1, L2_X2
   SwitchVar L2_Y1, L2_Y2
End If

'BoxCollision X?
If L1_X2 <= L2_X1 Or L2_X2 <= L1_X1 Then
   areLinesCrossing = False
   Exit Function
End If

'Rotate if line is vertically and from bottom to top
If L1_Y2 < L1_Y1 And L1_X1 = L1_X2 Then
   SwitchVar L1_X1, L1_X2
   SwitchVar L1_Y1, L1_Y2
End If
If L2_Y2 < L2_Y1 And L2_X1 = L2_X2 Then
   SwitchVar L2_X1, L2_X2
   SwitchVar L2_Y1, L2_Y2
End If
  
'Both lines vertically?
If L1_X1 = L1_X2 And L2_X1 = L2_X2 Then
   areLinesCrossing = False
   Exit Function
End If

'Both lines horizontal?
If L1_Y1 = L1_Y2 And L2_Y1 = L2_Y2 Then
   areLinesCrossing = False
   Exit Function
End If

'One line horizontal, one line vertically?
If L1_X1 = L1_X2 And L2_Y1 = L2_Y2 Then
   areLinesCrossing = (L1_X1 >= L2_X1 And L1_X1 <= L2_X2 And L2_Y1 >= L1_Y1 And L2_Y1 <= L1_Y2)
   Exit Function
ElseIf L1_Y1 = L1_Y2 And L2_X1 = L2_X2 Then
   areLinesCrossing = (L1_Y1 >= L2_Y1 And L1_Y1 <= L2_Y2 And L2_X1 >= L1_X1 And L2_X1 <= L1_X2)
   Exit Function
End If

'One line vertically, one line diagonally?
If L1_X1 = L1_X2 Then
   'Gradient and intersection with the Y-axis for Line2
   L2_m = (L2_Y2 - L2_Y1) / (L2_X2 - L2_X1)
   L2_n = L2_Y2 - L2_m * L2_X2

   h_Y = L2_m * L1_X1 + L2_n

   areLinesCrossing = (L1_X1 >= L2_X1 And L1_X1 <= L2_X2 And h_Y >= L1_Y1 And h_Y <= L1_Y2)
   Exit Function
ElseIf L2_X1 = L2_X2 Then
   'Gradient and intersection with the Y-axis for Line1
   L1_m = (L1_Y2 - L1_Y1) / (L1_X2 - L1_X1)
   L1_n = L1_Y2 - L1_m * L1_X2

   h_Y = L1_m * L2_X1 + L1_n

   areLinesCrossing = (L2_X1 >= L1_X1 And L2_X1 <= L1_X2 And h_Y >= L2_Y1 And h_Y <= L2_Y2)
   Exit Function
End If

'One line horizontal, one line diagonally?
If L1_Y1 = L1_Y2 Then
   'Gradient and intersection with the Y-axis for Line2
   L2_m = (L2_Y2 - L2_Y1) / (L2_X2 - L2_X1)
   L2_n = L2_Y2 - L2_m * L2_X2

   h_X = (L1_Y1 - L2_n) / L2_m

   areLinesCrossing = (h_X >= L1_X1 And h_X <= L1_X2 And h_X >= L2_X1 And h_X <= L2_X2)
   Exit Function
ElseIf L2_Y1 = L2_Y2 Then
   'Gradient and intersection with the Y-axis for Line1
   L1_m = (L1_Y2 - L1_Y1) / (L1_X2 - L1_X1)
   L1_n = L1_Y2 - L1_m * L1_X2

   h_X = (L2_Y1 - L1_n) / L1_m

   areLinesCrossing = (h_X >= L2_X1 And h_X <= L2_X2 And h_X >= L1_X1 And h_X <= L1_X2)
   Exit Function
End If

'Both lines are diagonally!
  
'Gradient and intersection with the Y-axis for Line1
L1_m = (L1_Y2 - L1_Y1) / (L1_X2 - L1_X1)
L1_n = L1_Y2 - L1_m * L1_X2

'Gradient and intersection with the Y-axis for Line2
L2_m = (L2_Y2 - L2_Y1) / (L2_X2 - L2_X1)
L2_n = L2_Y2 - L2_m * L2_X2

'Lines are parallel?
If L2_m = L1_m Then
   areLinesCrossing = False
   Exit Function
Else
   'X-value of the intersection
   s = (L1_n - L2_n) / (L2_m - L1_m)
End If

'Lines are at intersection?
areLinesCrossing = (s >= L1_X1 And s <= L1_X2 And s >= L2_X1 And s <= L2_X2)
End Function

'Use it to swap two variables
Public Sub SwitchVar(ByRef a As Variant, ByRef b As Variant)
  Dim dummy As Variant

  dummy = a
  a = b
  b = dummy
End Sub
