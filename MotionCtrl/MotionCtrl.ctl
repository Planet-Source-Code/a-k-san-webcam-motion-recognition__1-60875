VERSION 5.00
Begin VB.UserControl MotionCtrl 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   ScaleHeight     =   1065
   ScaleWidth      =   5160
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   33
      Left            =   2160
      Top             =   3000
   End
   Begin VB.PictureBox currentFrame 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Motion Recognition Control Created by A.K.San Copyright(c) 2005 A.K.San Version 2.1 (Luminance)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   0
      Picture         =   "MotionCtrl.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "MotionCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   *******************************************************
'   * Title: Custom control to detect directional motions *
'   * Project code name: Moonshine                        *
'   * Coded by A.K.San                                    *
'   * Created on 25 April 2005                            *
'   * Modified on 31 May 2005                             *
'   * Copyright(c) 2005 A.K.San                           *
'   *******************************************************

'NOTE: Something must be fixed. There is a run-time error when terminating the component.
'      After checking, there is nothing wrong with the test program or webcam's driver.
'      The problem is surely not with the webcam's driver or anything low level.
'      Finally the error found on the ActiveX control's InvisibleAtRuntime property.
'      The error will occur when the property's value is set to True.
'      Solution: Set the InvisibleAtRuntime property to False and set the
'                Visible property to False before running it.

'force declaration
Option Explicit

'=== Variable for public use ===
Public isEnabled As Boolean '(is the webcam enabled?)
Public AutoAdjust As Boolean '(set to true if the auto adjust feature is needed)
Public RetValueModX As Long '(the return value modifier for X axis)
Public RetValueModY As Long '(the return value modifier for Y axis)
Public RetValueMulX As Byte '(the return value multiplier for X axis)
Public RetValueMulY As Byte '(the return value multiplier for Y axis)
Public MovThresholdX As Byte '(the movement threshold for X axis before triggering the MotionDetect events)
Public MovThresholdY As Byte '(the movement threshold for Y axis before triggering the MotionDetect events)
Public ToleranceLevel As Integer '(Level of tolerance for colour difference)
Public LightThreshold As Integer '(The threshold to filter the background noise)

'=== Event for public use ===
Public Event MotionDetectX(Value As Integer) 'occurs when motion detected on X axis
Public Event MotionDetectY(Value As Integer) 'occurs when motion detected on Y axis
Public Event UsableState(CurrentState As Byte) 'occurs when the usable state changes

'=== API constants' declarations ===
Private Const CONNECT As Long = 1034 '(call code for connecting the capturing device)
Private Const DISCONNECT As Long = 1035 '(call code for disconnecting the capturing device)
Private Const GET_FRAME As Long = 1084 '(call code for getting a frame from the capturin device)
Private Const COPY As Long = 1054 '(call code for copying the frame captured to the clipboard)

'=== API functions' declarations ===

'Use to send messages to the capturing device
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Use to create the handler to the capturing device
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long

'=== Custom data type's declaration ===

'A data type for colour information in terms of Hue, Saturation and Luminance
Private Type HSL
    Hue As Long
    Saturation As Long
    Luminance As Long
End Type

'=== Variables for private use ===
Dim WebcamH As Long '(Handler to the webcam)
Dim TPPX As Single '(Twips per pixel on X axis)
Dim TPPY As Single '(Twips per pixel on Y axis)
Dim SensitivityLevel As Integer '(Level of sensitivity on number of pixels involved)
Dim PColour() As Long '(Array to store the colour information on both axis)
Dim PChgCol1() As Boolean '(Array to mark the pixels that changed colour on both axis)
Dim PChgCol2() As Boolean '(Same as PChgCol1 just that it buffers the information for comparison)
Dim PChgCol3() As Boolean '(Same as above - This one will make it more accurate)
Dim colourHSL As HSL '(Store the colour information in Hue, Saturation and Luminance)
'Since Hue and Saturation must be ingored to reduce background noice, only Luminance is require for comparing the changes
Dim Luminance1 As Single '(First luminance value that will be compared with the second)
Dim Luminance2 As Single '(Second luminance value that will be compared with the first)
Dim TotalLum As Double '(The total level of luminance required by the AutoAdjust feature)
Dim ColourBuffer As Long '(A buffer that will temporary hold a single pixel's colour value)
Dim LoopCounter As Integer '(Multipurpose counter for use in any loop)
Dim LoopCounter2 As Integer '(Second multipurpose counter for nested loop)
Dim PMoveX As Integer '(Same as moveX but this one is for private use)
Dim PMoveY As Integer '(Same as moveY but this one is for private use)
Dim State As Byte '(The current usable state based on the brightness level)

'=== Function for public use ===

'Use to enable the webcam
'Usage: Pass in True as the argument to enable and False to disable
'Return: False for no error and true if error occur
Public Function EnableWebcam(Enabled As Boolean) As Boolean
On Error GoTo ext
Select Case Enabled
Case True
WebcamH = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 640, 480, currentFrame.hwnd, 0)
DoEvents: SendMessage WebcamH, CONNECT, 0, 0
Timer1.Enabled = True
isEnabled = True
Case False
DoEvents: SendMessage WebcamH, DISCONNECT, 0, 0
Timer1.Enabled = False
isEnabled = False
End Select

EnableWebcam = False 'return false if no error
Exit Function
ext:
EnableWebcam = True 'return true on error
End Function

'=== Functions for private use ===

'Initialization of the control
Private Sub UserControl_Initialize()
'Set all the required variables to initial value
TPPX = Screen.TwipsPerPixelX
TPPY = Screen.TwipsPerPixelY
State = 0
currentFrame.Width = 640 * TPPX
currentFrame.Height = 480 * TPPY
SensitivityLevel = 10
ToleranceLevel = 40
LightThreshold = ToleranceLevel
AutoAdjust = False
isEnabled = False
RetValueModX = 0
RetValueModY = 0
RetValueMulX = 1
RetValueMulY = 1
MovThresholdX = 3
MovThresholdY = 3
'Redefine the arrays' size
ReDim PColour(640 / SensitivityLevel, 480 / SensitivityLevel)
ReDim PChgCol1(640 / SensitivityLevel, 480 / SensitivityLevel)
ReDim PChgCol2(640 / SensitivityLevel, 480 / SensitivityLevel)
ReDim PChgCol3(640 / SensitivityLevel, 480 / SensitivityLevel)
'Setup the initial size of the component
UserControl.Width = 5160
UserControl.Height = 1065
End Sub

Private Sub Timer1_Timer()
'reset the movement counters
PMoveX = 0
PMoveY = 0
'reset the total luminance level
TotalLum = 0

GetPic 'get a picture
ProcessColour 'process its colour
ProcessDirection 'process the direction
SaveColourInfo 'save the colour information
AdjustToleranceLevel 'adjusting the tolerance level for next capture
CheckState 'check the current lightning condition
End Sub

'Getting a snapshot from the webcam
Private Sub GetPic()
SendMessage WebcamH, GET_FRAME, 0, 0
SendMessage WebcamH, COPY, 0, 0
currentFrame.Picture = Clipboard.GetData
Clipboard.Clear
End Sub

'Process the raw colour data got from the webcam
Private Sub ProcessColour()
'Unlike the previous version, this one will process all the pixels based on the sensitivity level
    For LoopCounter = 0 To 640 / SensitivityLevel - 1
    For LoopCounter2 = 0 To 480 / SensitivityLevel - 1
    
        'getting the value of luminance for the first variable
        ColourBuffer = currentFrame.Point(LoopCounter * SensitivityLevel * TPPX, LoopCounter2 * SensitivityLevel * TPPY)
        colourHSL = RGBToHSL(ColourBuffer)
        Luminance1 = colourHSL.Luminance
        
        'getting the second value of luminance
        ColourBuffer = PColour(LoopCounter, LoopCounter2)
        colourHSL = RGBToHSL(ColourBuffer)
        Luminance2 = colourHSL.Luminance
        
        'accumulating the total luminance level
        TotalLum = TotalLum + Luminance1
        
        'compare and determine the motion
        If Abs(Luminance1 - Luminance2) > ToleranceLevel And Luminance1 > LightThreshold Then
        PChgCol1(LoopCounter, LoopCounter2) = True
        Else
        PChgCol1(LoopCounter, LoopCounter2) = False
        End If
        
        'saving the pixel's colour information
        PColour(LoopCounter, LoopCounter2) = currentFrame.Point(LoopCounter * SensitivityLevel * TPPX, LoopCounter2 * SensitivityLevel * TPPY)
        
    Next LoopCounter2
    Next LoopCounter

End Sub

'Process directional data
Private Sub ProcessDirection()
'This loop will process the first stage of directional detection
    For LoopCounter = 4 To 640 / SensitivityLevel - 5
    For LoopCounter2 = 4 To 480 / SensitivityLevel - 5
        If PChgCol3(LoopCounter, LoopCounter2) = True And PChgCol2(LoopCounter + 3, LoopCounter2) = True And PChgCol2(LoopCounter - 3, LoopCounter2) = False And PChgCol1(LoopCounter + 4, LoopCounter2) = True And PChgCol1(LoopCounter - 4, LoopCounter2) = False Then
        PMoveX = PMoveX - 1
        ElseIf PChgCol3(LoopCounter, LoopCounter2) = True And PChgCol2(LoopCounter - 3, LoopCounter2) = True And PChgCol2(LoopCounter + 3, LoopCounter2) = False And PChgCol1(LoopCounter - 4, LoopCounter2) = True And PChgCol1(LoopCounter + 4, LoopCounter2) = False Then
        PMoveX = PMoveX + 1
        End If
        
        If PChgCol3(LoopCounter, LoopCounter2) = True And PChgCol2(LoopCounter, LoopCounter2 + 3) = True And PChgCol2(LoopCounter, LoopCounter2 - 3) = False And PChgCol1(LoopCounter, LoopCounter2 + 4) = True And PChgCol1(LoopCounter, LoopCounter2 - 4) = False Then
        PMoveY = PMoveY + 1
        ElseIf PChgCol3(LoopCounter, LoopCounter2) = True And PChgCol2(LoopCounter, LoopCounter2 - 3) = True And PChgCol2(LoopCounter, LoopCounter2 + 3) = False And PChgCol1(LoopCounter, LoopCounter2 - 4) = True And PChgCol1(LoopCounter, LoopCounter2 + 4) = False Then
        PMoveY = PMoveY - 1
        End If
    Next LoopCounter2
    Next LoopCounter

'Filtering or refining the directional detection for X axis
    If PMoveX > MovThresholdX Then
    RaiseEvent MotionDetectX((PMoveX * RetValueMulX) + RetValueModX) 'motion detected on X axis
    ElseIf PMoveX < (-1 * MovThresholdX) Then
    RaiseEvent MotionDetectX((PMoveX * RetValueMulX) - RetValueModX)
    End If
    
'Filtering or refining the directional detection for Y axis
    If PMoveY > MovThresholdY Then
    RaiseEvent MotionDetectY((PMoveY * RetValueMulY) + RetValueModY) 'motion detected on Y axis
    ElseIf PMoveY < (-1 * MovThresholdY) Then
    RaiseEvent MotionDetectY((PMoveY * RetValueMulY) - RetValueModY)
    End If
    
End Sub

'Saving the colour information for later use
Private Sub SaveColourInfo()
    For LoopCounter = 0 To 640 / SensitivityLevel - 1
    For LoopCounter2 = 0 To 480 / SensitivityLevel - 1
        PChgCol3(LoopCounter, LoopCounter2) = PChgCol2(LoopCounter, LoopCounter2)
        PChgCol2(LoopCounter, LoopCounter2) = PChgCol1(LoopCounter, LoopCounter2)
    Next LoopCounter2
    Next LoopCounter
End Sub

'Adjusting the tolerance level based on the environment's lightning
Private Sub AdjustToleranceLevel()
If AutoAdjust = True Then 'adjust the level of tolerance only if required
    ToleranceLevel = Int(TotalLum / 10000)
    LightThreshold = Int(TotalLum / 3072)
End If
End Sub

'Check the usable state of the webcam
'since this sub only runs when the AutoAdjust function is enabled
'the AutoAdjust function will adjust the tolerance level based on the environment's lightning condition
'so this function will check on the ToleranceLevel to see if the environment is bright enough
Private Sub CheckState()
If AutoAdjust = True And ToleranceLevel < 6 And State <> 1 Then
    State = 1 'the system can't process properly in this state
    RaiseEvent UsableState(State)
ElseIf AutoAdjust = True And ToleranceLevel >= 6 And ToleranceLevel < 12 And State <> 2 Then
    State = 2 'the motion recognition's accuracy will be low
    RaiseEvent UsableState(State)
ElseIf AutoAdjust = True And ToleranceLevel >= 12 And ToleranceLevel < 20 And State <> 3 Then
    State = 3 'the best state
    RaiseEvent UsableState(State)
ElseIf AutoAdjust = True And ToleranceLevel >= 20 And State <> 4 Then
    State = 4 'the environment maybe too bright (not tested yet) - can't find the environment
    RaiseEvent UsableState(State)
End If
End Sub

'Got this function from the Internet
Private Function RGBToHSL(ByVal RGBValue As Long) As HSL
  ' by Paul - wpsjr1@syix.com, 20011120
  Dim R As Long, G As Long, B As Long
  Dim lMax As Long, lMin As Long
  Dim q As Single
  Dim lDifference As Long
  Static Lum(255) As Long
  Static QTab(255) As Single
  Static init As Long
  
  If init = 0 Then
    For init = 2 To 255 ' 0 and 1 are both 0
      Lum(init) = init * 100 / 255
    Next
    For init = 1 To 255
      QTab(init) = 60 / init
    Next init
  End If

  R = RGBValue And &HFF
  G = (RGBValue And &HFF00&) \ &H100&
  B = (RGBValue And &HFF0000) \ &H10000

  If R > G Then
    lMax = R: lMin = G
  Else
    lMax = G: lMin = R
  End If
  If B > lMax Then
    lMax = B
  ElseIf B < lMin Then
    lMin = B
  End If

  RGBToHSL.Luminance = Lum(lMax)
  
  lDifference = lMax - lMin
  If lDifference Then
    ' do a 65K 2D lookup table here for more speed if needed
    RGBToHSL.Saturation = (lDifference) * 100 / lMax
    q = QTab(lDifference)
    Select Case lMax
    Case R
      If B > G Then
        RGBToHSL.Hue = q * (G - B) + 360
      Else
        RGBToHSL.Hue = q * (G - B)
      End If
    Case G
      RGBToHSL.Hue = q * (B - R) + 120
    Case B
      RGBToHSL.Hue = q * (R - G) + 240
    End Select
  End If
End Function

Private Sub UserControl_Resize()
'These are the fixed dimensions for the component
UserControl.Width = 5160
UserControl.Height = 1065
End Sub
