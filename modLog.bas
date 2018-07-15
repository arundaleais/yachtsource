Attribute VB_Name = "modLog"
Option Explicit

'===================================
'These are for detecting 64 bit processor
Private Declare Function GetProcAddress Lib "KERNEL32" _
    (ByVal hModule As Long, _
    ByVal lpProcName As String) As Long
    
Private Declare Function GetModuleHandle Lib "KERNEL32" _
    Alias "GetModuleHandleA" _
    (ByVal lpModuleName As String) As Long
    
Private Declare Function GetCurrentProcess Lib "KERNEL32" _
    () As Long

Private Declare Function IsWow64Process Lib "KERNEL32" _
    (ByVal hProc As Long, _
    bWow64Process As Boolean) As Long
'=====================================
'http://vbcity.com/forums/p/99944/422558.aspx
Private Declare Function GetVersionExA Lib "KERNEL32" _
               (lpVersionInformation As OSVERSIONINFO) As Integer
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
'===================================
'=================================
'http://www.developerfusion.com/code/1607/counting-lines-in-a-multiline-textbox/
Private Declare Function SendMessageAsLong Lib "user32" _
     Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, ByVal lParam As Long) As Long
Const EM_GETLINECOUNT = 186
'================================
Public LogFileCh As Integer

Private LogFileName As String
Private TempPath As String

Public Function WriteLog(kb As String, Optional Idx As Long)
Static PreviousTime As String
Dim LogOk As Boolean

    If LogFileCh = -1 Then  'First time its opened
        LogFileCh = FreeFile
        TempPath = LongFileName(Environ("TEMP") & "\")
'        LogFileName = TempPath & App.EXEName & "Log.log"
        LogFileName = TempPath & App.EXEName & "_" & Format$(Date, "yyyymmdd_") & Format$(Time, "hhnnss") & ".log"
        Open LogFileName For Output As #LogFileCh
        WriteLog "Open Event Log [" & LogFileName & "]"
    End If
    If LogFileCh = 0 Then   'Re-open in append
        LogFileCh = FreeFile
        Open LogFileName For Append As #LogFileCh
    End If
    
'Check if a Try (supress)   'Used for Repeat Count (see NmeaRouter)
    LogOk = True
        
    If LogOk Then
        If Now() <> PreviousTime Then
            Print #LogFileCh, Now() & vbTab & kb
        Else
            Print #LogFileCh, Space$(Len(PreviousTime)) & vbTab & kb
        End If
        PreviousTime = Now()
    End If
    Close LogFileCh
    LogFileCh = 0

End Function


Public Function AddFileToLog(FileName As String)
Dim ch As Long
Dim nextline As String

    ch = FreeFile
    On Error GoTo nofil
    Open FileName For Input As #ch
    Do Until EOF(ch)
        Line Input #ch, nextline
        Call WriteLog(nextline)
    Loop
    Close #ch
'com0com log adds incremantally to the con0com log file
'so we need to delete it when we transfer the details to my event log
    Kill FileName
nofil:
End Function

#If False Then
Public Function CloseLog(Optional Flush As Boolean)
Dim Idx As Long

    For Idx = 1 To UBound(Sockets)
        With Sockets(Idx)
            If .TryCount > 0 Then
                WriteLog .TryCount & " Tries Opening " & .DevName
            End If
        End With
    Next Idx

    Close #LogFileCh
    If Flush Then
        Open LogFileName For Append As #LogFileCh
    Else
        LogFileCh = 0
    End If
End Function
#End If

Public Function LongFileName(ByVal short_name As String) As _
    String
Dim pos As Integer
Dim result As String
Dim long_name As String

    ' Start after the drive letter if any.
    If Mid$(short_name, 2, 1) = ":" Then
        result = Left$(short_name, 2)
        pos = 3
    Else
        result = ""
        pos = 1
    End If

    ' Consider each section in the file name.
    Do While pos > 0
        ' Find the next \
        pos = InStr(pos + 1, short_name, "\")

        ' Get the next piece of the path.
        If pos = 0 Then
'            long_name = Dir$(short_name, vbNormal + _
'                vbHidden + vbSystem + vbDirectory)
'a blank name (above returns a ".")
            long_name = ""
        Else
            long_name = Dir$(Left$(short_name, pos - 1), _
                vbNormal + vbHidden + vbSystem + _
                vbDirectory)
        End If
        result = result & "\" & long_name
    Loop

    LongFileName = result
End Function

'http://www.freevbcode.com/ShowCode.asp?ID=9043
Public Function Is64bit() As Boolean
    Dim handle As Long, bolFunc As Boolean

    ' Assume initially that this is not a Wow64 process
    bolFunc = False

    ' Now check to see if IsWow64Process function exists
    handle = GetProcAddress(GetModuleHandle("kernel32"), _
                   "IsWow64Process")

    If handle > 0 Then ' IsWow64Process function exists
        ' Now use the function to determine if
        ' we are running under Wow64
        IsWow64Process GetCurrentProcess(), bolFunc
    End If

    Is64bit = bolFunc

End Function

Public Function GetVersion1() As String
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    With osinfo
        Select Case .dwPlatformId
        Case 1
            Select Case .dwMinorVersion
            Case 0
                GetVersion1 = "Windows 95"
            Case 10
                GetVersion1 = "Windows 98"
            Case 90
                GetVersion1 = "Windows Millenium"
            Case Else
                GetVersion1 = "Unknown"
            End Select
        Case 2
            Select Case .dwMajorVersion
            Case 3
                GetVersion1 = "Windows NT 3.51"
            Case 4
                    GetVersion1 = "Windows NT 4.0"
            Case 5
                Select Case .dwMinorVersion
                Case 0
                    GetVersion1 = "Windows 2000"
                Case 1
                    GetVersion1 = "Windows XP"
                Case 2
                    GetVersion1 = "Windows Server 2003"
                Case Else
                    GetVersion1 = "Unknown"
                End Select
            Case 6
                Select Case .dwMinorVersion
                Case 0
                    GetVersion1 = "Windows Vista"
                Case 1
                    GetVersion1 = "Windows 7"
                Case 2
                    GetVersion1 = "Windows 8"
                Case Else
                    GetVersion1 = "Unknown"
                End Select
            Case Else
                GetVersion1 = "Unknown"
            End Select
        Case Else
            GetVersion1 = "Unknown"
        End Select
    End With
End Function

Public Function WriteState(Caller As String)
Dim KeyNo As Long
Dim kb As String
    Call DisplayState(Caller & " " & state.NextEventTime)

    Call DisplayState(vbTab & "NextEventTime=" & frmMain.aSecToElapsed(state.NextEventTime))
    Call DisplayState(vbTab & "Program=" & aProgramState(state.Program))
    Call DisplayState(vbTab & "Sequence=" & aSequenceState(state.Sequence))
    Call DisplayState(vbTab & "Recalls=" & aRecallsState(state.Recalls))
    If IsClassesInitialised(Classes) = True Then
        Call DisplayState(vbTab & "NextClassToStart=" & Classes(state.StartClass.Next).Name)
        Call DisplayState(vbTab & "PreviousClassStarted=" & Classes(state.StartClass.Previous).Name)
    End If
    If IsKeysInitialised(Keys) = True Then
        For KeyNo = 1 To UBound(Keys)
            Call DisplayState(vbTab & "Key=" & Keys(KeyNo).KeyName _
            & "," & aKeyState(Keys(KeyNo).state) _
            & "," & Keys(KeyNo).Cancel)
        Next KeyNo
    End If
    Call DisplayState("")
End Function

Private Function DisplayState(Data As String) As Long
'    Call WriteLog(kb)
    If frmMain.txtLog.Enabled = True Then
        frmMain.txtLog.SelStart = Len(frmMain.txtLog.Text)
        frmMain.txtLog.SelText = Replace(Data, vbTab, "") & vbCrLf
         If Len(frmMain.txtLog.Text) > 4096 Then
            frmMain.txtLog.Text = Right$(frmMain.txtLog.Text, 2048)
        End If
    End If
    DisplayState = SendMessageAsLong(frmMain.txtLog.hwnd, EM_GETLINECOUNT, 0, 0)
End Function
