VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbDepth 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1170
      List            =   "Form1.frx":0019
      TabIndex        =   1
      Text            =   "1"
      ToolTipText     =   "Depth of iteneration"
      Top             =   0
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type points
  x As Double
  y As Double
End Type

Private Sub draw(s As points, e As points, i As Integer)
Dim a1 As points, a2 As points, a3 As points
Dim a4 As points, a5 As points
Dim x, y As Double
  a1 = s: a5 = e
  x = sina(s, e) * s.y + cosa(s, e) * s.x
  y = cosa(s, e) * s.y - sina(s, e) * s.x
  a2.x = CDbl((x + dis(s, e) / 3) * cosa(s, e) - y * sina(s, e))
  a2.y = CDbl(y * cosa(s, e) + (x + dis(s, e) / 3) * sina(s, e))
  a3.x = CDbl((x + dis(s, e) / 2) * cosa(s, e) - (y + dis(s, e) * Sqr(3) / 6) * sina(s, e))
  a3.y = CDbl((y + dis(s, e) * Sqr(3) / 6) * cosa(s, e) + (x + dis(s, e) / 2) * sina(s, e))
  a4.x = CDbl((x + dis(s, e) * 2 / 3) * cosa(s, e) - y * sina(s, e))
  a4.y = CDbl(y * cosa(s, e) + (x + dis(s, e) * 2 / 3) * sina(s, e))
  If (i = CInt(cmbDepth.Text - 1)) Then
    Line (CInt(a1.x), CInt(a1.y))-(CInt(a2.x), CInt(a2.y)), RGB(0, 255, 0)
    Line (CInt(a2.x), CInt(a2.y))-(CInt(a3.x), CInt(a3.y)), RGB(0, 255, 0)
    Line (CInt(a3.x), CInt(a3.y))-(CInt(a4.x), CInt(a4.y)), RGB(0, 255, 0)
    Line (CInt(a4.x), CInt(a4.y))-(CInt(a5.x), CInt(a5.y)), RGB(0, 255, 0)
  Else
    Call draw(a1, a2, i + 1)
    Call draw(a2, a3, i + 1)
    Call draw(a3, a4, i + 1)
    Call draw(a4, a5, i + 1)
  End If
End Sub

Private Function dis(s As points, e As points) As Double
  dis = Sqr(CDbl(e.x - s.x) * CDbl(e.x - s.x) + CDbl(e.y - s.y) * CDbl(e.y - s.y))
End Function

Private Function sina(s As points, e As points) As Double
  If (e.x <> s.x) Then
    If (e.x > s.x) Then
      sina = Sin(Atn((e.y - s.y) / (e.x - s.x)))
    Else
      sina = -Sin(Atn((e.y - s.y) / (e.x - s.x)))
    End If
  Else
    If (e.y > s.y) Then
      sina = 1
    Else
      sina = -1
    End If
  End If
End Function

Private Function cosa(s As points, e As points) As Double
  If (e.x <> s.x) Then
    If (e.x > s.x) Then
      cosa = Cos(Atn((e.y - s.y) / (e.x - s.x)))
    Else
      cosa = -Cos(Atn((e.y - s.y) / (e.x - s.x)))
    End If
  Else
    cosa = 0
  End If
End Function

Private Sub Command1_Click()
  Dim s As points: Dim e As points
  Dim i As Integer
  'ScaleMode = vbPixels
  s.x = ScaleWidth: s.y = 4100
  e.x = 0: e.y = 4100
  i = 0
  Me.Cls
  Call draw(s, e, i)
End Sub

