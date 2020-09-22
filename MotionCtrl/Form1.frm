VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin Project1.MotionCtrl MotionCtrl1 
      Height          =   1065
      Left            =   1200
      TabIndex        =   1
      Top             =   4320
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   1879
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is only a test program for the Motion Recognition component.
'The component can be implemented in many other types of applications like games, 3D object manipulation....

'Each of Forms will perform different tests, please change the Startup Object from
'Project>Project1 Properties...>Startup Object
'Choose any Forms that you like to test.
'Thanks.

Private Sub Form_Load()
With MotionCtrl1
.Visible = False 'make the control invisible at runtime
.AutoAdjust = True 'turn on the AutoAdjust feature
'must be more than 0
'larger number will make it less sensitive
.MovThresholdX = 1 'sensitivity level for X axis
.MovThresholdY = 1 'sensitivity level for Y axis
.RetValueModX = 100 'return value modifier for X axis
.RetValueModY = 100 'renturn value modifier for Y axis
.RetValueMulX = 1 'return value multiplier for X axis
.RetValueMulY = 1 'return value multiplier for Y axis
.EnableWebcam True 'turn on the webcam
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'REMEMBER: Must turn the webcam off before ending the program.
'Else it will crash.
MotionCtrl1.EnableWebcam False
End
End Sub

Private Sub MotionCtrl1_MotionDetectX(Value As Integer)
'When motion on X axis detected
Shape1.Left = Shape1.Left + Value
Me.Caption = MotionCtrl1.ToleranceLevel & " : " & MotionCtrl1.LightThreshold
End Sub

Private Sub MotionCtrl1_MotionDetectY(Value As Integer)
'when motion on Y axis detected
Shape1.Top = Shape1.Top + Value
Me.Caption = MotionCtrl1.ToleranceLevel & " : " & MotionCtrl1.LightThreshold
End Sub

Private Sub MotionCtrl1_UsableState(CurrentState As Byte)
'when the state changes
'display the state
Select Case CurrentState
Case 1
Label1.Caption = "The environment is too dark."
Case 2
Label1.Caption = "Poor visibility."
Case 3
Label1.Caption = "Ready."
Case 4
Label1.Caption = "Too bright."
End Select
End Sub
