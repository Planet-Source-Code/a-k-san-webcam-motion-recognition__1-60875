VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExitWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Const EWX_FORCE As Long = 4

Private Type LUID
UsedPart As Long
IgnoredForNowHigh32BitPart As Long
End Type

Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
TheLuid As LUID
Attributes As Long
End Type

Private Enum EnumExitWindows
WE_LOGOFF = 0
WE_SHUTDOWN = 1
WE_REBOOT = 2
WE_POWEROFF = 8
End Enum
Private Sub AdjustToken()
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle
' Get the LUID for shutdown privilege.
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1    ' One privilege to set
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
' Enable the shutdown privilege in the access token of this process.
AdjustTokenPrivileges hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub

Public Sub ShutOff()
AdjustToken
ExitWindowsEx (WE_POWEROFF Or EWX_FORCE), 0
End Sub

Public Sub LogOff()
AdjustToken
ExitWindowsEx (WE_LOGOFF Or EWX_FORCE), 0
End Sub

Public Sub Reset()
AdjustToken
ExitWindowsEx (WE_REBOOT Or EWX_FORCE), 0
End Sub

Public Sub ShutDown()
AdjustToken
ExitWindowsEx (WE_SHUTDOWN Or EWX_FORCE), 0
End Sub
