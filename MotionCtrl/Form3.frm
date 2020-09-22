VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.MotionCtrl MotionCtrl1 
      Height          =   1065
      Left            =   1920
      TabIndex        =   6
      Top             =   5760
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   1879
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5880
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   6240
      Top             =   4440
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This is a motion recognition system that accepts the user's hand movements as the system's main input."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   -3000
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1005
      TabIndex        =   4
      Top             =   4680
      Width           =   2265
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   495
      Left            =   840
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1005
      TabIndex        =   3
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Logoff Computer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1005
      TabIndex        =   2
      Top             =   2640
      Width           =   2925
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows Media Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1005
      TabIndex        =   1
      Top             =   1680
      Width           =   4005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Explorer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1005
      TabIndex        =   0
      Top             =   720
      Width           =   2925
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This test program will demonstrate the application of motion recognition in a menu interface
'Note that most of the functions are not working properly as the purpose of the test program is to test the component

'Each of Forms will perform different tests, please change the Startup Object from
'Project>Project1 Properties...>Startup Object
'Choose any Forms that you like to test.
'Thanks.

Option Explicit
Dim blueValue As Integer
Dim EWin As New ExitWindows
Dim TFlag As Byte

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
MotionCtrl1.Visible = False
TFlag = 0
Me.Top = 0
Me.Left = 0
Me.Height = Screen.Height
Me.Width = Screen.Width
Label3.Top = Label2.Top + 960
Label4.Top = Label3.Top + 960
Label5.Top = Label4.Top + 960
MotionCtrl1.AutoAdjust = True
MotionCtrl1.MovThresholdX = 7
MotionCtrl1.MovThresholdY = 2
MotionCtrl1.EnableWebcam True
End Sub

Private Sub MotionCtrl1_MotionDetectX(Value As Integer)
Dim h As Byte
h = getPos()
If Value < 0 And TFlag = 1 Then 'going left
    TFlag = 2
    Timer2.Enabled = True
ElseIf Value > 0 Then
    
    Select Case h
    Case 1
    'Shell "C:\Program Files\Internet Explorer\iexplore.exe", vbNormalFocus
    Case 2
    'Shell "C:\Program Files\Windows Media Player\wmplayer.exe", vbNormalFocus
    Case 3
    'EWin.LogOff
    Case 4
    TFlag = 1
    Label6.Visible = True
    Timer2.Enabled = True
    Case 5
    Unload Me
    End Select
    
End If
End Sub

Private Sub MotionCtrl1_MotionDetectY(Value As Integer)
If Value < 0 And TFlag = 0 Then 'going up
    setPos getPos - 1
ElseIf Value > 0 And TFlag = 0 Then
    setPos getPos + 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MotionCtrl1.EnableWebcam False
End
End Sub

Private Sub Timer1_Timer()
If blueValue < 255 Then
blueValue = blueValue + 20
Else
blueValue = 0
End If
Shape1.BorderColor = RGB(0, 0, blueValue)
End Sub

Private Function getPos() As Byte
If Shape1.Top = Label1.Top Then
getPos = 1
ElseIf Shape1.Top = Label2.Top Then
getPos = 2
ElseIf Shape1.Top = Label3.Top Then
getPos = 3
ElseIf Shape1.Top = Label4.Top Then
getPos = 4
ElseIf Shape1.Top = Label5.Top Then
getPos = 5
End If
End Function

Private Sub setPos(pos As Byte)
Select Case pos
Case 1
    Shape1.Top = Label1.Top
Case 2
    Shape1.Top = Label2.Top
Case 3
    Shape1.Top = Label3.Top
Case 4
    Shape1.Top = Label4.Top
Case 5
    Shape1.Top = Label5.Top
End Select
End Sub

Private Sub Timer2_Timer()
If Label6.Left < 5280 And TFlag = 1 Then
Label6.Left = Label6.Left + 250
ElseIf Label6.Left >= 5280 And TFlag = 1 Then
Timer2.Enabled = False
ElseIf Label6.Left > -3000 And TFlag = 2 Then
Label6.Left = Label6.Left - 250
ElseIf Label6.Left <= -3000 And TFlag = 2 Then
Label6.Visible = False
Timer2.Enabled = False
TFlag = 0
End If
End Sub
