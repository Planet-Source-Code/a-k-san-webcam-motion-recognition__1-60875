VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin Project1.MotionCtrl MotionCtrl1 
      Height          =   1065
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   1879
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1320
      Top             =   3120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is just a test program to demonstrate the control of mouse cursor with motion recognition.

'Each of Forms will perform different tests, please change the Startup Object from
'Project>Project1 Properties...>Startup Object
'Choose any Forms that you like to test.
'Thanks.

Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Dim XY As POINTAPI

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
MotionCtrl1.Visible = False
MotionCtrl1.RetValueModX = 10 'Screen.TwipsPerPixelX - 5
MotionCtrl1.RetValueModY = 10 'Screen.TwipsPerPixelY - 5
MotionCtrl1.RetValueMulX = 1
MotionCtrl1.RetValueMulY = 1
MotionCtrl1.AutoAdjust = True
MotionCtrl1.MovThresholdX = 1
MotionCtrl1.MovThresholdY = 1
MotionCtrl1.EnableWebcam True
End Sub

Private Sub Form_Unload(Cancel As Integer)
MotionCtrl1.EnableWebcam False
End
End Sub

Private Sub MotionCtrl1_MotionDetectX(Value As Integer)
GetCursorPos XY
SetCursorPos XY.X + Value, XY.Y
End Sub

Private Sub MotionCtrl1_MotionDetectY(Value As Integer)
GetCursorPos XY
SetCursorPos XY.X, XY.Y + Value
End Sub

Private Sub MotionCtrl1_UsableState(CurrentState As Byte)
Select Case CurrentState
Case 1
Me.Caption = "Poor visibility"
Case 2
Me.Caption = "May not work well"
Case 3
Me.Caption = "Ready"
Case 4
Me.Caption = "Too bright"
End Select
End Sub
