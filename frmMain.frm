VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Pathfinder"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   7485
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.OptionButton optMode 
      BackColor       =   &H0000C000&
      Caption         =   "Set End"
      Height          =   195
      Index           =   2
      Left            =   3795
      TabIndex        =   3
      Top             =   30
      Width           =   870
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H0000FF00&
      Caption         =   "Set Start"
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   2
      Top             =   30
      Width           =   915
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H0080FF80&
      Caption         =   "Delete Polygons"
      Height          =   195
      Index           =   3
      Left            =   1425
      TabIndex        =   6
      Top             =   30
      Width           =   1455
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Make Polygons"
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      FillColor       =   &H00000080&
      Height          =   5415
      Left            =   0
      MousePointer    =   1  'Pfeil
      ScaleHeight     =   357
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   497
      TabIndex        =   0
      Top             =   300
      Width           =   7515
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "@Gungsuh"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   480
         Left            =   4500
         TabIndex        =   5
         Top             =   2550
         Width           =   330
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "@Gungsuh"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   480
         Left            =   1650
         TabIndex        =   4
         Top             =   2550
         Width           =   315
      End
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "Time: 0ms; PathLen: -1px"
      Height          =   195
      Left            =   4665
      TabIndex        =   7
      Top             =   30
      Width           =   1800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Any question? Ask robert_kaltenbach@hotmail.com

'I know that it's slow. Any suggestions?
'I think it would be much faster
'if I use fewer classes.

'Another setback is that it's quite impossible
'to integrate into another application.

'But the advantage is:
'It's extremely accurate! (imho)

'I'm working on a new version without polygons
'but have lost the interest.

'Frequently occurring abbreviations:
'Obs  -> obstacle
'Plg  -> polygon
'Pnt  -> point
'Pdc  -> predecessor
'Pos  -> position
'Rgn  -> region
'Ptr  -> pointer
'Len  -> length
'Tmp  -> temporarly
'Cur  -> current
'Num  -> number
'Org  -> origin
'Dst  -> destination
'Vtx  -> vertex
'Cnt  -> connection


'Hint:
'_________________________________
'If Not Function1 Then
'ElseIf Function2 Then
'   commands
'End If
'_________________________________
'and
'_________________________________
'If Function1 And Function2 Then
'   commands
'End If
'_________________________________
'isn't the same!!!
'
'For example:
'_________________________________
'If (X <> 0) And (Y / X > 0) Then
'   commands
'End If
'_________________________________
'raises an error if "X = 0" (Division by Zero).
'Even though VB didn't need to divide,
'because the first argument was "False".

Option Explicit

Const BorderColor As Long = vbBlack
Const PathColor As Long = vbGreen

Dim MousePos As POINTAPI
Dim NewObs As New clsPolygon
Dim Pathfinder As New clsPathfinder

Private Sub Form_Load()
Show

PI = Atn(1) * 4

Pathfinder.StartPosX = picMap.ScaleWidth / 4
Pathfinder.StartPosY = picMap.ScaleHeight / 2
Pathfinder.EndPosX = picMap.ScaleWidth / 4 * 3
Pathfinder.EndPosY = picMap.ScaleHeight / 2

Pathfinder.findPath
lblTime.Caption = "Time: " & Pathfinder.Time & "ms; PathLen: " & CInt(Pathfinder.PathLen) & "px"

lblStart.Move Pathfinder.StartPosX - lblStart.Width / 2, Pathfinder.StartPosY - lblStart.Height / 2
lblEnd.Move Pathfinder.EndPosX - lblEnd.Width / 2, Pathfinder.EndPosY - lblEnd.Height / 2

DrawThem

MsgBox "Beginners Guide:" & vbNewLine & _
   "Click anywhere with left button three times." & vbNewLine & _
   "Click anywhere with right button." & vbNewLine & _
   "The green line shows you the shortest way from ""S"" to ""E""."

'Experts Guide:
'-
'1   Unneeded points:
'1.1 Don't double-click.
'1.2 You don't have to finish a polygon by yourself.
'-
'2   The boundaries of a polygon mustn't cross.
'-
'3   Polygons can overlap.
'-
'4   At variance with the opinions of many ppl:
'    You can resize the form!!!
End Sub

Private Sub Form_Resize()
On Error Resume Next
picMap.Width = ScaleWidth
picMap.Height = ScaleHeight - picMap.Top
End Sub

Private Sub lblEnd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap_MouseUp Button, Shift, X / (picMap.Width / picMap.ScaleWidth) + lblEnd.Left, Y / (picMap.Height / picMap.ScaleHeight) + lblEnd.Top
End Sub

Private Sub lblStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap_MouseMove Button, Shift, X / (picMap.Width / picMap.ScaleWidth) + lblStart.Left, Y / (picMap.Height / picMap.ScaleHeight) + lblStart.Top
End Sub

Private Sub lblEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap_MouseMove Button, Shift, X / (picMap.Width / picMap.ScaleWidth) + lblEnd.Left, Y / (picMap.Height / picMap.ScaleHeight) + lblEnd.Top
End Sub

Private Sub lblStart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMap_MouseUp Button, Shift, X / (picMap.Width / picMap.ScaleWidth) + lblStart.Left, Y / (picMap.Height / picMap.ScaleHeight) + lblStart.Top
End Sub

Private Sub optMode_Click(Index As Integer)
NewObs.Init
DrawThem
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePos.X = X
MousePos.Y = Y
If NewObs.Count > 0 Then
   picMap.Cls
   picMap.ForeColor = BorderColor
   picMap.Line (NewObs.Vtx(NewObs.Count - 1).X, NewObs.Vtx(NewObs.Count - 1).Y)-(MousePos.X, MousePos.Y)
End If
End Sub

Public Sub DrawThem()
Dim I As Long, I2 As Long

picMap.AutoRedraw = True
picMap.Cls
picMap.ForeColor = BorderColor

picMap.FillStyle = 0
Pathfinder.DrawPlgs picMap.hdc
picMap.FillStyle = 1
Pathfinder.DrawPlgs picMap.hdc

If NewObs.Count > 0 Then
   For I2 = 0 To NewObs.Count - 2
      picMap.Line (NewObs.Vtx(I2).X, NewObs.Vtx(I2).Y)-(NewObs.Vtx(I2 + 1).X, NewObs.Vtx(I2 + 1).Y)
   Next I2
End If

picMap.ForeColor = PathColor
Pathfinder.DrawPath picMap.hdc
picMap.AutoRedraw = False

End Sub

Private Sub picMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Long, I2 As Long

If optMode(0).Value Then
   If Button = 1 Then
      NewObs.Add MousePos.X, MousePos.Y
   Else
      If Not NewObs.Valid Then Exit Sub
      Pathfinder.Add NewObs
      Set NewObs = New clsPolygon
      
      Pathfinder.findPath
      lblTime.Caption = "Time: " & Pathfinder.Time & "ms; PathLen: " & CInt(Pathfinder.PathLen) & "px"
   End If

ElseIf optMode(1).Value Then
   Pathfinder.StartPosX = MousePos.X
   Pathfinder.StartPosY = MousePos.Y
   lblStart.Move Pathfinder.StartPosX - lblStart.Width / 2, Pathfinder.StartPosY - lblStart.Height / 2
   Pathfinder.findPath
   lblTime.Caption = "Time: " & Pathfinder.Time & "ms; PathLen: " & CInt(Pathfinder.PathLen) & "px"
   
ElseIf optMode(2).Value Then
   Pathfinder.EndPosX = MousePos.X
   Pathfinder.EndPosY = MousePos.Y
   lblEnd.Move Pathfinder.EndPosX - lblEnd.Width / 2, Pathfinder.EndPosY - lblEnd.Height / 2
   Pathfinder.findPath
   lblTime.Caption = "Time: " & Pathfinder.Time & "ms; PathLen: " & CInt(Pathfinder.PathLen) & "px"
   
ElseIf optMode(3).Value Then
   I = 0
   Do Until I >= Pathfinder.Count
      If Pathfinder.HitTest(I, MousePos.X, MousePos.Y) Then
         Pathfinder.Remove (I)
      Else
         I = I + 1
      End If
   Loop
   Pathfinder.findPath
   lblTime.Caption = "Time: " & Pathfinder.Time & "ms; PathLen: " & CInt(Pathfinder.PathLen) & "px"
End If

DrawThem
End Sub
