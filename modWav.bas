Attribute VB_Name = "modWav"
Option Explicit
'See http://msdn.microsoft.com/en-us/library/ms712587.aspx
'see http://www.vbforfree.com/mci-multimedia-command-string-tutorial-a-step-by-step-guide/
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Private WavFile As String
Private Command As String
Private Success As Boolean
Private retVal As Long
Private returnData As String

      
Public SoundFilePath As String
Private WavMem As String

Public Function OpenWav(ByVal FileName As String)

    If WavFile <> "" Then
        Call CloseWav
    End If
    
    WavFile = Chr(34) & SoundFilePath & FileName & Chr(34)
'make the buffer 128 characters
    returnData = String(128, 0) 'Space(128)
    Command = "open " & WavFile & " type waveaudio alias mysound" ' buffer 6"
    retVal = mciSendString(Command, 0, 0, 0)
    Success = mciGetErrorString(retVal, returnData, Len(returnData))
    If Success = False Then
        WavFile = ""
        frmMain.StatusBar1.Panels(5).Picture = Nothing
    Else
        frmMain.StatusBar1.Panels(5).Picture = LoadPicture(SignalImageFilePath & "speaker.gif")
    End If
End Function

Public Function PlayWav()
    
'May be a controller but no Sound Card
    If WavFile = "" Then Exit Function
    
    Command = "play mysound from 0"    ' & " from 0 to 500"
    retVal = mciSendString(Command, 0, 0, 0)
'    Debug.Print "play " & retVal
    Success = mciGetErrorString(retVal, returnData, Len(returnData))
    If Success = False Then MsgBox Trim(returnData), , Command
End Function

Public Function PauseWav()
    
'May be a controller but no Sound Card
    If WavFile = "" Then Exit Function
    
    retVal = mciSendString("pause mysound", 0, 0, 0)
'    Debug.Print "pause " & retVal
    Success = mciGetErrorString(retVal, returnData, Len(returnData))
    If Success = False Then MsgBox Trim(returnData), , Command
End Function

Public Function CloseWav()
    
'May be a controller but no Sound Card
    If WavFile = "" Then Exit Function
    
    Command = "close mysound"
    retVal = mciSendString(Command, 0, 0, 0)
'    Debug.Print "close " & retVal
    Success = mciGetErrorString(retVal, returnData, 128)
    If Success = False Then MsgBox Trim(returnData), , Command
End Function


