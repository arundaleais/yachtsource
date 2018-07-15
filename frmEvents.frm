VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFlxGd.ocx"
Begin VB.Form frmEvents 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Events"
   ClientHeight    =   3096
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   6072
   ControlBox      =   0   'False
   Icon            =   "frmEvents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3096
   ScaleWidth      =   6072
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshEvents 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10181
      _ExtentY        =   2561
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup Menu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuCopyToClipBoard 
         Caption         =   "Copy to ClipBoard as CSV"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub Form_Load()
    Height = Screen.Height - 1000
    Width = 6800
    Left = Screen.Width - Width
    Top = 100
    
'    Top = Screen.Top
    With mshEvents
        .Top = ScaleTop
        .Left = ScaleLeft
        .Width = ScaleWidth
        .Height = ScaleHeight
        .FormatString = "^Event|<Time|Eidx|<Signal|<Action|<Class"
        .ColWidth(1) = 800  'Time
        .ColWidth(2) = 400  'Eidx
        .ColWidth(3) = 2000 'Signal
        .ColWidth(4) = 2000 'Action
        .ColWidth(5) = 1000
'        For i = 1 To 20
'            .Rows = i + 1
'            .TextMatrix(i, 0) = i
'        Next i
'        .TextMatrix(1, 1) = "13:22:45"
    End With
    
End Sub

Public Function ListEvents()
Dim Row As Long
Dim Eidx As Long
Dim Sidx As Long
Dim Bidx As Long
Dim Idx As Long     'Keeps the Previous Signal or button (to combine Action on 1 line)
Dim kb As String
Dim i As Long
Dim Class As String

    If UdtArrayExists = False Then 'Occurs when form is loading and recall raised
        Exit Function
    End If
    
'    .Visible = True
'    frmEvents.SetFocus
    With mshEvents
        Do While .Rows > 2
            .RemoveItem .Rows - 1
        Loop
        For i = 0 To .Cols - 1
            .TextMatrix(1, i) = ""  'blank remaining rows
            .BackColor = vbWhite
        Next i
        Row = 0
'        If Not UBound(Evts) Is Nothing Then
'        End If
        
        For Eidx = 0 To UBound(Evts)
            Row = Row + 1

'created with first row ""
            If .TextMatrix(1, 0) <> "" Then
                .AddItem Row, Row
            End If
            .TextMatrix(Row, 0) = Row
            .TextMatrix(Row, 1) = aMins(Evts(Eidx).ElapsedTime + Classes(Evts(Eidx).Class).Offset)
            .TextMatrix(Row, 1) = Evts(Eidx).ElapsedTime + Classes(Evts(Eidx).Class).Offset
            .TextMatrix(Row, 2) = Eidx
'This will be overwritten if there is a Sidx or Bidx
            .TextMatrix(Row, 3) = Evts(Eidx).Message
            If Evts(Eidx).Focus >= 0 Then
                Call AddMessage(Row, "Focus-" & frmMain.Commands(Evts(Eidx).Focus).Caption)
            End If
            If Evts(Eidx).Signal > 0 Then
                .TextMatrix(Row, 5) = SignalAttributes(Evts(Eidx).Signal).Name
            End If
            If IsSignalsInitialised(Evts(Eidx).Signals) Then
                For Sidx = 0 To UBound(Evts(Eidx).Signals)
'                   If Evts(Eidx).Signals(Sidx).Signal <> Idx Then
'Add another Row after First Eidx for this time
'                        If Sidx > 0 Then
                            Row = Row + 1
                            .AddItem Row, Row
'                        End If
'                    End If
                    Idx = Evts(Eidx).Signals(Sidx).Signal
'                    .TextMatrix(Row, 2) = Eidx
                    .TextMatrix(Row, 3) = SignalAttributes(Idx).Name
                    kb = Evts(Eidx).Signals(Sidx).Raise
                    If kb <> "" Then
                        If kb = "True" Then
                            Call AddMessage(Row, "Up")
                        Else
                            Call AddMessage(Row, "Down")
                        End If
                    End If
                    kb = Evts(Eidx).Signals(Sidx).Silent
                    If kb <> "" Then
                        If kb = "True" Then
                            Call AddMessage(Row, "Silent")
                        Else
                            Call AddMessage(Row, "Sound")
                        End If
                    End If
'                    If Evts(Eidx).Signals(Sidx).signal > 0 Then
'                        .TextMatrix(Row, 5) = SignalAttributes(Evts(Eidx).Signals(Sidx).signal).Name
'                    End If
                Next Sidx
            End If
            If IsButtonsInitialised(Evts(Eidx).Buttons) Then
                For Bidx = 0 To UBound(Evts(Eidx).Buttons)
'Add another Row after First Eidx for this time
'                    If Evts(Eidx).Buttons(Bidx).Button <> Idx Then
'Add another Row after First Eidx for this time
'                        If Sidx > 0 Or Bidx > 0 Then
                            Row = Row + 1
                            .AddItem Row, Row
'                        End If
'                    End If
                    Idx = Evts(Eidx).Buttons(Bidx).Button
'                    .TextMatrix(Row, 2) = Eidx
                    .TextMatrix(Row, 3) = frmMain.Commands(Idx).Caption
                    kb = Evts(Eidx).Buttons(Bidx).Enabled
                    If kb <> "" Then
                        If kb = "True" Then
                            Call AddMessage(Row, "Enabled")
                        Else
                            Call AddMessage(Row, "Disabled")
                        End If
                    End If
'                    .TextMatrix(Row, 5) = SignalAttributes(Evts(Eidx).Buttons(Bidx).signal).Name
                Next Bidx
            End If
            Idx = 0
'Me.Visible = True
        Next Eidx
    Me.Visible = True
    .Col = 0
    .Row = 0
    .FocusRect = flexFocusNone ' (The selected cell changes)
    End With
'Me.Visible = False
frmEvents.WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
frmEvents.Refresh
frmEvents.Visible = True
'If Events are after LoadingProfile, you need to set the focus back display
'the cursor in the FirstStartTime box
    frmMain.SetFocus
    Call CopyToLog
    Call CopyToClipBoard
End Function

Private Function AddMessage(Row, kb)
    With mshEvents
        If .TextMatrix(Row, 4) = "" Then
            .TextMatrix(Row, 4) = kb
        Else
            .TextMatrix(Row, 4) = .TextMatrix(Row, 4) & "," & kb
        End If
    End With
End Function

Private Sub Form_Resize()
    mshEvents.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub CopyToClipBoard()
Dim i As Long
Dim j As Long
Dim kb As String

With mshEvents
    For i = 1 To .Rows - 1
        If Not (.TextMatrix(i, 0) = "" And .TextMatrix(i, 2) = "") Then
            For j = 1 To .Cols - 1
            kb = kb + QuotedString(.TextMatrix(i, j), ",") & ","
            Next j
            kb = kb + vbCrLf
        End If
    Next i
End With
Clipboard.Clear
Clipboard.SetText kb
End Sub

Private Sub CopyToLog()
Dim i As Long
Dim j As Long
Dim kb As String

kb = "mshEvents"
WriteLog kb
With mshEvents
    For i = 0 To .Rows - 1
        If Not (.TextMatrix(i, 0) = "" And .TextMatrix(i, 2) = "") Then
            kb = ""
            For j = 1 To .Cols - 1
            kb = kb + QuotedString(.TextMatrix(i, j), ",") & ","
            Next j
            WriteLog kb
        End If
    Next i
End With
End Sub

Private Sub mshEvents_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If state.Program = 2 Then   'Occurs when first event is manually clicked
        Call SetProgramState(3)
        Call frmMain.DisplayStartTimes
    End If
    Call ActionEvent
frmMain.SetFocus
End Sub

Private Sub mshEvents_SelChange()
'    Call ActionEvent
End Sub

Private Function ActionEvent()
Dim Idx As Long
Dim Index As Long
Dim arry() As String
Dim Minus As Boolean
Dim Sign As Integer
        
    With mshEvents
        If .Row < 1 Then .Row = 1 'Miss Header Row
            If .Row = 1 Then
'                Call frmMain.StartTimeIsSet
frmMain.RaceTimer.Enabled = False
            End If
'            Index = .TextMatrix(.Row, 0)
'                If .Col = 0 Then
'                    Call frmMain.DoEvent(Index)
'                End If
            If .Col = 1 Then
                If .TextMatrix(.Row, .Col) <> "" Then
                    arry = Split(.TextMatrix(.Row, .Col), ":")
                    If UBound(arry) > 0 Then
                        Sign = Sgn(Replace(.TextMatrix(.Row, .Col), ":", ""))
                        state.NextEventTime = CLng(arry(0)) * 60 + CLng(arry(1)) * Sign
                    Else
                        state.NextEventTime = CLng(arry(0))
                    End If
                    
'Update the elapsed time for the event clicked
'Should be done by the event                    Call frmMain.DisplayElapsedTimes
                    Call frmMain.DoTimerEvents
If .TextMatrix(.Row, 3) = "Finish Enabled" Then frmMain.RaceTimer.Enabled = True
                End If
            End If
        If .Row = .Rows - 1 Then
            .Row = 0
            .TopRow = 1
        End If
    End With
    
End Function

