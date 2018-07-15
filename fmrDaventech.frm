VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmDaventech 
   Caption         =   "Controller"
   ClientHeight    =   3216
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   4872
   ControlBox      =   0   'False
   Icon            =   "fmrDaventech.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3216
   ScaleWidth      =   4872
   Begin VB.Frame fraState 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   15
      Top             =   2160
      Width           =   2775
      Begin VB.Label lblState 
         Caption         =   "Program"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblProgram 
         Caption         =   "Program"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblSequence 
         Caption         =   "Sequence"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblRecalls 
         Caption         =   "Recalls"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblState 
         Caption         =   "Seq"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblState 
         Caption         =   "Recalls"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame fraContainer 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1935
      Begin VB.CheckBox Relay 
         Caption         =   "Relay 8"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox Relay 
         Caption         =   "Relay 7"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox Relay 
         Caption         =   "Relay 6"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox Relay 
         Caption         =   "Relay 5"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox Relay 
         Caption         =   "Relay 4"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox Relay 
         Caption         =   "Relay 3"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox Relay 
         Caption         =   "Relay 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Relay 
         Caption         =   "Relay 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame fraBoard 
      Caption         =   "Board"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton Option1 
         Caption         =   "ETHRLY16"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ETH008"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Timer ReconnectTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4200
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3960
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4200
      Top             =   120
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.Label lblVoltage 
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblWinsock 
      Height          =   1815
      Left            =   2160
      TabIndex        =   5
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmDaventech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim DataSnd As String   'Keep a copy of the CommandString in case we need to retry
Dim Controller As Long
Dim Command As Long
Dim OnTime As Long
Private Closing As Boolean  'Stops reconnect timer
Private ReplyWait As Boolean    'Must wait for a reply after data is sent
Dim ReturnStates As Boolean     'True if we are checking the states, suppresses Click event
                                'gernerated if the tick box is changed programaticcally
Dim RcvBytes() As Byte

'Ad hoc test
Private Sub Command1_Click()
Dim State As Long
Dim arry(2) As String
Dim ControlString As String
'Debug.Print "======="
    If Command = 0 Then
        Call OpenAndSend("")    'reset
        Command = 32
        Exit Sub
    End If
    If Controller = 0 Then Controller = 1
    If Controller = 9 Then
        Controller = 1
        If Command = 32 Then
            Command = 33
        Else
            Command = 32
        End If
    End If
    arry(0) = Command
    arry(1) = Controller
    arry(2) = OnTime
    ControlString = Join(arry, ",")

    State = OpenAndSend(ControlString)
    Controller = Controller + 1
End Sub

'Opens Winsock if closed Returns the state of Winsock
Public Function OpenAndSend(ControlString As String) As Long

WriteLog "OpenAndSend " & ControlString & " (" & Timer & ")"
    
    If ControlString = "" Then
    'Reset all off
        DataSnd = "35,0,0"  'eth008
    Else
        DataSnd = ControlString
    End If
    With Winsock1
        Select Case .State
        Case Is = sckClosed
            Call CreateWinsock
        Case Is = sckConnected
            Call WinsockOutput
        End Select
    OpenAndSend = .State
    End With
End Function

Private Sub CreateWinsock()
Dim lWaitUntil As Long
    With Winsock1
        If .State = 0 Then
            .Protocol = sckTCPProtocol

            If Option1(0).Value = True Then
                .RemoteHost = "eth008"  'SYC
            Else
                .RemoteHost = "ethrly16"
                Option1(1).Enabled = True
            End If
'Debug.Print .RemoteHost
            .RemotePort = 17494
            lWaitUntil = TimeToQuit(1)
WriteLog "CreateWinsock " & .RemoteHost & ":" & .RemotePort
            .Connect
'Debug.Print "connecting"
'            Do Until Winsock1.State = sckConnected Or Timer > lWaitUntil
                 DoEvents
'            Loop
'Debug.Print "End"
            If Winsock1.State = sckConnected Then
'Debug.Print "Connection Successful"
WriteLog "Connection Successful" & " (" & Timer & ")"
                DataSnd = "35,0,0"   'clear all eth008
                Call WinsockOutput
                Call GetVoltage 'also sets BatteryVoltage
            Else
'Debug.Print "Connection TimedOut"
'Call DisplayWinsock
            End If
        End If
    End With
WriteLog "WinsockState = " & aState(Winsock1.State)

'Terminate ReconnectTimer when unloading frmDaventech
    ReconnectTimer.Enabled = False
ReconnectTimer.Interval = 5000
    ReconnectTimer.Enabled = Not Closing

End Sub

Public Sub CloseWinsock()
Dim lWaitUntil As Long

    On Error GoTo Winsock_err
    With Winsock1
'Clear all controller relays
        If .State = sckConnected Then
            DataSnd = "35,0,0"   'clear all ethoo8
            Call WinsockOutput
            lblVoltage = "0.0"
        End If
        lWaitUntil = TimeToQuit(5)
        .Close
        Do Until Winsock1.State = sckClosed Or Timer > lWaitUntil
            DoEvents
        Loop
        If Winsock1.State = sckClosed Then
'Debug.Print "Connection Closed"
        Else
'Debug.Print "Close TimedOut"
        End If
    End With
Exit Sub
Winsock_err:
    On Error GoTo 0
    MsgBox "CloseWinsock Error " & Err.Number & " " & Err.Description & vbCrLf, , "Close Winsock"
End Sub

Private Sub Form_Load()
    Top = frmMain.Top + frmMain.Height - Height
    Left = frmMain.Width - Width
If GetComputerName = "ADMIN-PC" Then
    Option1(1).Value = True
    Option1(0).Enabled = Option1(0).Value
    Option1(1).Enabled = Option1(1).Value
End If
    Call CreateWinsock
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Debug.Print "frmDaventech.unload " & Winsock1.State
WriteLog "frmDaventech.unload " & Winsock1.State & " (" & Timer & ")"
    Closing = True  'stop reconnecttimer restarting
    ReconnectTimer.Enabled = False
    If Winsock1.State = sckConnected Then
        DataSnd = "35,0,0"  'eth008
        Call WinsockOutput
        lblVoltage = "0.0"
        Call CloseWinsock
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    Call CloseWinsock
    ReconnectTimer.Enabled = False
    Call CreateWinsock
End Sub

Private Sub ReconnectTimer_Timer()
'Debug.Print "Reconnect " & Winsock1.State
        
    With Winsock1
        Select Case .State
        Case Is = sckClosed, sckClosing ', sckError
            Call CreateWinsock
        Case Is = sckConnecting, sckError 'Client
            Call CloseWinsock
'ReconnectTimer.Interval = ReconnectTimer.Interval * 2
            Call CreateWinsock
        Case Is = sckConnectionPending  'Server Only
'            Call CloseHandler(Idx)
'            Call OpenHandler(Idx)
        Case Is = sckConnected
'ReconnectTimer.Interval = 500
'            ReconnectTimer.Enabled = False
        End Select
        If .State = sckConnected Then
            frmMain.StatusBar1.Panels(4).Picture = LoadPicture(SignalImageFilePath & "connected.gif")
'            Call SendControllers
        Else
            frmMain.StatusBar1.Panels(4).Picture = LoadPicture(SignalImageFilePath & "notconnected.gif")
        End If
    End With
'ReconnectTimer.Interval = 500
'Debug.Print ReconnectTimer.Interval
End Sub


Private Sub Relay_Click(Index As Integer)
Debug.Print "Click " & Relay(Index).Value & ", " & ReturnStates
WriteLog "Click " & Relay(Index).Value & ", " & ReturnStates & " (" & Timer & ")"
    
    If Relay(Index).Value = vbChecked Then
        DataSnd = "32," & Index + 1 & ",0"  'Perm on eth008
    Else
        DataSnd = "33," & Index + 1 & ",0"  'Perm off eth008
    End If
'    Call WinsockOutput
'Dont action the event a second time if returning the states
    If ReturnStates = False Then Call OpenAndSend(DataSnd)
End Sub

Private Sub Winsock1_Connect()
    If Winsock1.State = sckConnected Then
Debug.Print "Connected " & Winsock1.State
WriteLog "Connected " & Winsock1.State & " (" & Timer & ")"
        DataSnd = "35,0,0"   'clear all eth008
        Call WinsockOutput
        lblVoltage = GetVoltage 'also sets BatteryVoltage
    End If
    Call DisplayWinsock
End Sub

Sub WinsockOutput()
Dim lWaitUntil As Long
Dim b() As Byte
Dim arry() As String
Dim i As Long
Dim MyRelay As OptionButton
Dim kb As String

'the Winsock Control element may not have had time to be created
'by the time it starts
'the Port & Socket may have been closed by the user while there'
'were unsent Sentences in the buffer
        
    ReturnStates = False    'Assume not returning the states of the relays
    With Winsock1
        If .State = sckConnected Then
            If DataSnd <> "" Then
'Format is always eth008 in files
                arry = Split(DataSnd, ",")
'These are the command for which we return a state (otherwise we have to wait for a timeout)
                Select Case arry(0)
                Case Is = 32, 33, 35
                    ReturnStates = True
                End Select
                
                Select Case .RemoteHost
                Case Is = "eth008"  'SYC
Debug.Print "eth008=" & DataSnd
WriteLog "eth008=" & DataSnd & " (" & Timer & ")"
                    ReplyWait = True            'All eth008 commands return a least 1 byte
                Case Is = "ethrly16"            'convert eth008 to ethrly16
'91 is to force a DataArrival event  which will set ReplyWait=false terminating the Do events loop
                    Select Case arry(0)
                    Case Is = 32    'on
                        DataSnd = 100 + arry(1)
                    Case Is = 33    'off
                        DataSnd = 110 + arry(1)
                    Case Is = 35    'all off
                        DataSnd = "92,0"
                    Case Is = 120
                        DataSnd = "93"      'Get battery voltage 1 byte
                        ReplyWait = True    'Wait for reply
                    End Select
Debug.Print "ethrly16=" & DataSnd
WriteLog "ethrly16=" & DataSnd & " (" & Timer & ")"
                    arry = Split(DataSnd, ",")
                End Select
                
                ReDim b(UBound(arry))
                                
                For i = 0 To UBound(arry)
                    If IsNumeric(arry(i)) Then
                        b(i) = arry(i)
                    End If
                Next i
'Debug.Print "DataSnd=" & DataSnd
                lWaitUntil = TimeToQuit(5)
                .SendData b
'if eth008  the controller will reply here with a byte from here and a second from the returned states
'If setting a relay state, then also return new state
                If ReturnStates = True Then
                    ReDim b(0)
                    Select Case .RemoteHost
                    Case Is = "eth008"  'SYC
                        b(0) = 36
                    Case Is = "ethrly16"            'convert eth008 to ethrly16
                        b(0) = 91
                    End Select
                    ReplyWait = True        'Wait for reply after requestiog states
                    .SendData b
                End If
Debug.Print "WaitForReply=" & ReplyWait
WriteLog "WaitForReply=" & ReplyWait & " (" & Timer & ")"
                
                Do Until ReplyWait = False Or Timer > lWaitUntil
'Wait until a reply is received (if a reply is expected)
WriteLog "DoEvents " & " (" & Timer & ")"
                    Sleep 10    'This allows time for the response before the next linked command
                                'Otherwise in the log there are 2 commands then 2 responses
                    DoEvents
                Loop
Debug.Print "ReplyWait=" & ReplyWait
WriteLog "ReplyWait=" & ReplyWait & " (" & Timer & ")"
'there should only be one byte output for the relay states
'this rejects any other response from containing more than 1 byte
                If UBound(RcvBytes) <> -1 Then
                    If ReplyWait = False Then
                        For i = 0 To 7
                            If RcvBytes(UBound(RcvBytes)) And 2 ^ i Then
                                If ReturnStates Then
                                    Relay(i).Value = vbChecked
                                End If
                                kb = kb & "1"
                            Else
                                If ReturnStates Then
                                    Relay(i).Value = vbUnchecked
                                End If
                                kb = kb & "0"
                            End If
                        Next i
Debug.Print "Replied Bytes=" & kb
WriteLog "RepliedBytes " & kb & " (" & Timer & ")"
                    Else
Debug.Print "Timeout"
WriteLog "Timeout" & " (" & Timer & ")"
                    End If
                Else
Debug.Print "ReplyRejected"     'Too many bytes
WriteLog "ReplyRejected" & " (" & Timer & ")"
                End If
'                Set MyRelay = Nothing
            End If
        End If
    End With
    Call DisplayWinsock
    ReturnStates = False    'Assume not returning the states of the relays
Exit Sub
SendData_err:
'    MsgBox "Send Data Error " & Str(Err.Number) & " " & Err.Description & vbCrLf
WriteLog "Send Data Error " & Str(Err.Number) & " " & Err.Description & " (" & Timer & ")"
    Call DisplayWinsock
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim DataRcv As String
Dim PeekBytes() As Byte
Dim kb As String
Dim i As Long

    On Error GoTo DataArrival_err
    With Winsock1
        .GetData RcvBytes, vbArray + vbByte       'Must receive bytes not string
        For i = 0 To UBound(RcvBytes)
            kb = kb & RcvBytes(i) & " "
         Next i
Debug.Print "DataArrival=" & kb
WriteLog "DataArrival=" & kb & " (" & Timer & ")"
'        .PeekData PeekBytes, vbArray + vbByte
'If UBound(PeekBytes) = -1 Then
        ReplyWait = False
'End If
    End With
Exit Sub

DataArrival_err:
    Select Case Err.Number
    Case Is = sckBadState
    Case Is = sckMsgTooBig
    Case Is = sckConnectionReset
    Case Else
        MsgBox "UDP/TCP DataArrival Error " & Str(Err.Number) & " " & Err.Description
    End Select
End Sub

'If we cant connect - just close the connection (it will be retried by the reconnect timer)
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'frmMain.StatusBar1.Panels(1).Text = Description
Dim kb As String

    kb = "Error:" & Number & vbCrLf & "Message:" & Description
'Call DisplayWinsock(kb)
    Select Case Number
    Case Is = 11001         'No Host Found
        Call CloseWinsock
'        Call CreateWinsock
    Case Is = 10065         'No route to host
        Call CloseWinsock
'        Call CreateWinsock
'Stop
    Case Else
        Call CloseWinsock
'Stop
    End Select
    Call DisplayWinsock
End Sub

Public Function TimeToQuit(TimeToWait As Long) As Long
'http://www.freevbcode.com/ShowCode.asp?ID=1977
'PURPOSE:  Returns a TimeOut value, in metric
'Seconds from Midnight (i.e., return value of Timer function)
'taking into account that Midnight may occur within the elapsed time

'PARAMETER: TimeToWait: Number of Seconds to Wait

'RETURN VALUE:  When to TimeOut, in Seconds From Midnight

'EXAMPLE: Implements a 30 second timeout before giving
'up on a winsock Connection

'Dim lWaitUntil as Long
'lWaitUntil = TimeToQuit(30)
'Winsock1.Connect
'Do Until Winsock1.State = sckConnected or Timer > _
'   lWaitUntil
'     DoEvents
'Loop
'If Winsock1.State = sckConnected Then
'     MsgBox "Connection Successful"
'Else
'    MsgBox "Connection TimedOut
'End If
'*************************************************************
Dim lStart As Long
Dim lTimeToQuit As Long
Dim lTimeToWait As Long

lStart = Timer
lTimeToWait = TimeToWait

If lStart + TimeToWait < 86400 Then
        lTimeToQuit = lStart + lTimeToWait
    Else
        lTimeToQuit = (lStart - 86400) + TimeToWait
    End If

TimeToQuit = lTimeToQuit

End Function

Public Sub DisplayWinsock(Optional ErrorMessage As String)
Dim kb As String
    With Winsock1
            lblVoltage = "Battery Voltage = " & BatteryVoltage
            kb = kb & "Local IP=" & .LocalIP & vbCrLf
            kb = kb & "Local Port=" & .LocalPort & vbCrLf
            kb = kb & "Protocol=" & aProtocol(.Protocol) & vbCrLf
            kb = kb & "Remote Host=" & .RemoteHost & vbCrLf
            kb = kb & "Remote Host IP=" & .RemoteHostIP & vbCrLf
            kb = kb & "Remote Port=" & .RemotePort & vbCrLf
            kb = kb & "State=" & aState(.State) & vbCrLf
            If ErrorMessage <> "" Then
                kb = kb & ErrorMessage & vbCrLf
            End If
    End With
'    MsgBox kb, , "TCP/IP Sockets"
    lblWinsock = kb
    
End Sub

Public Function aProtocol(Protocol As Integer)
    Select Case Protocol
    Case Is = sckTCPProtocol
        aProtocol = "TCP"
    Case Is = sckUDPProtocol
        aProtocol = "UDP"
    End Select
End Function

Public Function aState(State As Integer) As String
    Select Case State
    Case Is = -1
        aState = "Nothing"
    Case Is = 0
        aState = "Closed"
    Case Is = 1
        aState = "Open"
    Case Is = 2
        aState = "Listening"
    Case Is = 3
        aState = "Connection pending"
    Case Is = 4
        aState = "Resolving host"
    Case Is = 5
        aState = "Host resolved"
    Case Is = 6
        aState = "Connecting"
    Case Is = 7
        aState = "Connected"
    Case Is = 8
        aState = "Peer is closing connection"
    Case Is = 9
        aState = "Error"
    Case Is = 11
        aState = "Opening"
    Case Is = 18
        aState = "Closing"
    Case Is = 21
        aState = "Data loss"
    Case Is = 22
        aState = "Data in buffer"
    Case Else
        aState = "Invalid"
    End Select
End Function

Private Function GetVoltage()
    Call OpenAndSend("120")
    If UBound(RcvBytes) = -1 Then
        BatteryVoltage = CSng("0.00")
    Else
        BatteryVoltage = CSng(Format$(RcvBytes(0) / 10, "00.0"))
    End If
    lblVoltage = "Battery Voltage = " & GetVoltage
WriteLog lblVoltage & " (" & Timer & ")"
End Function
