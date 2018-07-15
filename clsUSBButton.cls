VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUSBButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function PacInitialize Lib "PacDrive32" () As Long
Private Declare Sub PacShutdown Lib "PacDrive32" ()
Private Declare Function PacSetLEDStates Lib "PacDrive32" (ByVal id As Long, ByVal Data As Integer) As Long
Private Declare Function PacSetLEDState Lib "PacDrive32" (ByVal id As Long, ByVal port As Long, ByVal state As Long) As Long
Private Declare Function PacGetDeviceType Lib "PacDrive32" (ByVal id As Long) As Long
Option Explicit

Private Declare Function PacGetVersionNumber Lib "PacDrive32" (ByVal id As Long) As Long

Private Declare Function Pac64SetLEDIntensities Lib "PacDrive32" (ByVal id As Long, Data As Byte) As Long
Private Declare Function Pac64SetLEDIntensity Lib "PacDrive32" (ByVal id As Long, ByVal port As Long, ByVal intensity As Byte) As Long

Private Declare Function Pac64SetLEDStates Lib "PacDrive32" (ByVal id As Long, ByVal group As Long, ByVal Data As Byte) As Long
Private Declare Function Pac64SetLEDState Lib "PacDrive32" (ByVal id As Long, ByVal group As Long, ByVal port As Long, ByVal state As Long) As Long

Private Declare Function Pac64SetLEDStatesRandom Lib "PacDrive32" (ByVal id As Long) As Long
Private Declare Function Pac64SetLEDFadeTime Lib "PacDrive32" (ByVal id As Long, ByVal fadeTime As Byte) As Long
Private Declare Function Pac64SetScriptStepDelay Lib "PacDrive32" (ByVal id As Long, ByVal stepDelay As Byte) As Long

Private Declare Function Pac64StartScriptRecording Lib "PacDrive32" (ByVal id As Long) As Long
Private Declare Function Pac64StopScriptRecording Lib "PacDrive32" (ByVal id As Long) As Long
Private Declare Function Pac64RunScript Lib "PacDrive32" (ByVal id As Long) As Long
Private Declare Function Pac64ClearFlash Lib "PacDrive32" (ByVal id As Long) As Long

Private Declare Function Pac64SetDeviceId Lib "PacDrive32" (ByVal id As Long, ByVal newId As Long) As Long

Private Declare Function PacProgramUHid Lib "PacDrive32" (ByVal id As Long, ByVal sFilePath As String) As Long

'USBButtonConfigureColor(INT id, BYTE red, BYTE green, BYTE blue)
Private Declare Function USBButtonConfigureColor Lib "PacDrive32" (ByVal id As Long, Optional ByVal red As Long, Optional ByVal green As Byte, Optional ByVal blue As Byte) As Long
'bool USBButtonGetState(int id, bool *state)
Private Declare Function USBButtonGetState Lib "PacDrive32" (ByVal id As Long, arry As Variant) As Long

Dim Ret As Long

Private Sub Class_Initialize()
Dim numdevices As Integer
    
    numdevices = PacInitialize
    If numdevices > 0 Then
        
'PacDrive32.dll must be in YachtSource

'Appears to be 2 devices for USBButton ? Output & Input
        Ret = PacGetDeviceType(0)   'Returns 8 if button
        Ret = PacGetDeviceType(1)   'Returns 8 if button
        
ReDim arry(16)
        Ret = USBButtonGetState(0, arry)    'returns 1=true
        Ret = USBButtonGetState(1, arry)    'returns 0=false

        Ret = USBButtonConfigureColor(0, 0, 0, 0)   'turn off
        Ret = USBButtonConfigureColor(0, 255, 0, 0) 'Red
        Ret = USBButtonConfigureColor(0, 0, 255, 0) 'Green
        Ret = USBButtonConfigureColor(0, 0, 0, 255) 'Blue
        Ret = USBButtonConfigureColor(0, 0, 255, 255) '
         Ret = USBButtonConfigureColor(0, 255, 0, 255) 'Magenta
       Ret = USBButtonConfigureColor(0, 255, 255, 0) 'Yellow

    End If

End Sub

Public Function Off()
        Ret = USBButtonConfigureColor(0, 0, 0, 0)   'turn off
End Function

Public Function Yellow()
       Ret = USBButtonConfigureColor(0, 255, 255, 0) 'Yellow
End Function

Private Sub Class_Terminate()
    Ret = USBButtonConfigureColor(0, 0, 0, 0)   'turn off
    PacSetLEDStates 0, 0
    PacShutdown
End Sub
