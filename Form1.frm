VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D Cube"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   ForeColor       =   &H000000FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "&Color"
      Height          =   375
      Left            =   8400
      TabIndex        =   7
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Zoom --"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Zoom +"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "RIGHT"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "LEFT"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DOWN"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UP"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   6360
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   6975
      Left            =   0
      ScaleHeight     =   461
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   661
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   960
         Top             =   3600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ww As Integer
Dim Ixy_angle, Iz_angle, dYYshift, dXXshift, csx, csy As Integer
Dim cosa, cosb, sina, sinb, coscosba, cossinba, sincosba, sinsinba, zoom, pi180 As Double



Private Sub posxy(x1 As Double, y1 As Double, z1 As Double)
    Dim Yy, Xx As Double
    Yy = zoom / (10# - (z1 * cosb + y1 * sinsinba - x1 * sincosba))
    Xx = 100# * (1# + (y1 * cosa + x1 * sina) * Yy)
    csx = Int(dXXshift) + Int(Xx)
    Xx = 100# * (1# + (y1 * cossinba - x1 * coscosba - z1 * sinb) * Yy)
    csy = Int(dYYshift) + Int(Xx)
End Sub


Sub rollup()
    Iz_angle = (Iz_angle + 5)
    cosb = Cos(Iz_angle * pi180)
    sinb = Sin(Iz_angle * pi180)
    sinsinba = sinb * sina
    sincosba = sinb * cosa
    cossinba = sina * cosb
    coscosba = cosb * cosa


    Picture1.Cls
        NewPaint
    End Sub


Sub rolldown()
    Iz_angle = (Iz_angle - 5)
    cosb = Cos(Iz_angle * pi180)
    sinb = Sin(Iz_angle * pi180)
    sinsinba = sinb * sina
    sincosba = sinb * cosa
    cossinba = sina * cosb
    coscosba = cosb * cosa


   Picture1.Cls
        NewPaint
    End Sub


Sub rollright()
    Ixy_angle = (Ixy_angle - 5)
    cosa = Cos(Ixy_angle * pi180)
    sina = Sin(Ixy_angle * pi180)
    sinsinba = sinb * sina
    sincosba = sinb * cosa
    cossinba = sina * cosb
    coscosba = cosb * cosa


   Picture1.Cls
        NewPaint
    End Sub


Sub rollleft()
    Ixy_angle = (Ixy_angle + 5)
    cosa = Cos(Ixy_angle * pi180)
    sina = Sin(Ixy_angle * pi180)
    sinsinba = sinb * sina
    sincosba = sinb * cosa
    cossinba = sina * cosb
    coscosba = cosb * cosa


  Picture1.Cls
        NewPaint
    End Sub

Private Sub Command1_Click()
Picture1.SetFocus
Call picture1_KeyDown(vbKeyUp, 0)
End Sub

Private Sub Command2_Click()
Call picture1_KeyDown(vbKeyDown, 0)
Picture1.SetFocus
End Sub

Private Sub Command3_Click()
Picture1.SetFocus
Call picture1_KeyDown(vbKeyLeft, 0)
End Sub

Private Sub Command4_Click()
Picture1.SetFocus
Call picture1_KeyDown(vbKeyRight, 0)
End Sub

Private Sub Command5_Click()
Call picture1_KeyDown(vbKeySpace, 0)
Picture1.SetFocus
End Sub

Private Sub Command6_Click()
Call picture1_KeyPress(13)
Picture1.SetFocus
End Sub

Private Sub Command7_Click()
CommonDialog1.Color = Picture1.ForeColor
CommonDialog1.Flags = 1
CommonDialog1.ShowColor
Picture1.ForeColor = CommonDialog1.Color
Picture1.SetFocus
End Sub

Private Sub picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyLeft
ww = 1
Case vbKeyRight
ww = 2
Case vbKeyUp
ww = 3
Case vbKeyDown
ww = 4
Case vbKeySpace
ww = 5
Case vbKeyEscape
End
End Select
End Sub

Private Sub picture1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ww = 6
End If
End Sub

Private Sub Form_Load()
'   Picture1.ScaleMode = 3
    pi180 = 0.01745392
    Ixy_angle = 270
    Iz_angle = 85
    cosa = Cos(Ixy_angle * pi180)
    sina = Sin(Ixy_angle * pi180)
    cosb = Cos(Iz_angle * pi180)
    sinb = Sin(Iz_angle * pi180)
    sinsinba = sinb * sina
    sincosba = sinb * cosa
    cossinba = sina * cosb
    coscosba = cosb * cosa
    dYYshift = 80
    dXXshift = 80
    zoom = 5#
    NewPaint
End Sub


Sub NewPaint()
  
    posxy -1, -1, -1: xxx = csx: yyy = csy:
    posxy -1, 1, -1: Picture1.Line (xxx, yyy)-(csx, csy): x = csx: y = csy
    posxy -1, 1, 1: Picture1.Line (x, y)-(csx, csy): x = csx: y = csy
    posxy -1, -1, 1: Picture1.Line (x, y)-(csx, csy): Picture1.Line (csx, csy)-(xxx, yyy)
    posxy 1, -1, -1: xxx = csx: yyy = csy:
    posxy 1, 1, -1: Picture1.Line (xxx, yyy)-(csx, csy): x = csx: y = csy
    posxy 1, 1, 1: Picture1.Line (x, y)-(csx, csy): x = csx: y = csy
    posxy 1, -1, 1: Picture1.Line (x, y)-(csx, csy): Picture1.Line (csx, csy)-(xxx, yyy)
    
    posxy 1, -1, -1: x = csx: y = csy: posxy -1, -1, -1: Picture1.Line (x, y)-(csx, csy)
    posxy 1, -1, 1: x = csx: y = csy: posxy -1, -1, 1: Picture1.Line (x, y)-(csx, csy)
    posxy 1, 1, 1: x = csx: y = csy: posxy -1, 1, 1: Picture1.Line (x, y)-(csx, csy)
    posxy 1, 1, -1: x = csx: y = csy: posxy -1, 1, -1: Picture1.Line (x, y)-(csx, csy)
End Sub

Private Sub Text1_Change()
Timer1.Interval = Text1.Text
End Sub

Private Sub Timer1_Timer()
Select Case ww
Case 1
rollleft
Case 2
rollright
Case 3
rollup
Case 4
rolldown
Case 5
zoom = zoom * 1.01
Picture1.Cls
NewPaint
Case 6
zoom = zoom * 0.98
Picture1.Cls
NewPaint
End Select
End Sub

