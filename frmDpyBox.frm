VERSION 5.00
Begin VB.Form frmDpyBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   Icon            =   "frmDpyBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer RefreshTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Timer HideMeTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmDpyBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MIN_TEXT_HEIGHT = 780#
Const MIN_TEXT_WIDTH = 4680#
Const TEXTBOX_PADDING = 50# 'otherwise text runs into form border

'http://vbnet.mvps.org/index.html?code/textapi/txscroll.htm
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Private Declare Function PutFocus Lib "user32" _
   Alias "SetFocus" _
  (ByVal hwnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal _
    hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hwnd As Long) As Long

 Private Const EM_LINESCROLL = &HB6
Const EM_GETLINECOUNT = 186

'Dim MaxFrmHeight As Single
'Dim MaxFrmWidth As Single
Dim FrmBorderWidth As Single    'External=Internal size
Dim FrmBorderHeight As Single
Dim MaxTextHeight As Single
Dim MaxTextWidth As Single
'Dim FrmWidth As Single  'Used to calculate the Frm Size, before setting at the end
'Dim FrmHeight As Single
Dim MsgBuffer As String

Public Sub DpyBox(Message As String, Optional DisplaySecs As Long, Optional strCaption As String)
    Dim SavedWnd As Long

   'save the window handle of the control that currently has focus
    On Error Resume Next    'may be no window
    SavedWnd = Screen.ActiveControl.hwnd
    On Error GoTo 0
'Buffer the message while the display timer is enabled
'Cant put into Text1 as it would force a refresh
    If strCaption <> "" Then
        Caption = strCaption
    Else
        Caption = "Message"
    End If
    If DisplaySecs = 0 Then DisplaySecs = 5
    MsgBuffer = MsgBuffer + Message
    HideMeTimer.Interval = DisplaySecs * 1000
    HideMeTimer.Enabled = True   'restart timer
    If RefreshTimer.Enabled = True Then Exit Sub
    Text1.SelStart = Len(Text1.Text)    'causes less flicker than above
    Text1.SelText = MsgBuffer
    MsgBuffer = ""
    Call RefreshDisplay
'To slow it down to debug RefreshTimer.Interval = 4000
    RefreshTimer.Enabled = True 'comment out to debug
'    Call SetForegroundWindow(Me.hWnd)
    On Error Resume Next    'may not be a window
    Call PutFocus(SavedWnd)
    On Error GoTo 0
End Sub

Private Function RefreshDisplay()
Dim LineCount As Long
Dim ret As Long

    On Error GoTo Error_RefreshDisplay
'    Text1 = Text1 & Message

'CANT SHOW if a MODAL form is currently displayed
    Show

'MsgBox TextWidth(Text1)
    Select Case TextWidth(Text1)
    Case Is < MIN_TEXT_WIDTH
        Text1.Width = MIN_TEXT_WIDTH
    Case Is > MaxTextWidth
        Text1.Width = MaxTextWidth
    Case Else
        Text1.Width = TextWidth(Text1)
    End Select

'Make the textbox the maximum size so we can calculate the no of lines
    Width = Text1.Width + FrmBorderWidth + TEXTBOX_PADDING * 2
    Height = Text1.Height + FrmBorderHeight + TEXTBOX_PADDING * 2
    Top = Screen.Height - Height
    Left = Screen.Width - Width
'Remove all lines over 30
    LineCount = GetLineCount(Text1)
    Do Until LineCount < 30
        Text1 = Mid$(Text1, InStr(1, Text1, vbCrLf) + 2)
'        Height = Text1.Height + FrmBorderHeight + 100    '50 each side
        LineCount = GetLineCount(Text1)
    Loop
'We now have all the text we wish to display in the textbox
    Select Case TextHeight(Text1)
    Case Is < MIN_TEXT_HEIGHT
        Text1.Height = MIN_TEXT_HEIGHT
    Case Is > MaxTextHeight
        Text1.Height = MaxTextHeight
    Case Else
        Text1.Height = TextHeight(Text1)
    End Select
    Height = Text1.Height + FrmBorderHeight + TEXTBOX_PADDING * 2
    Top = Screen.Height - Height
    ret = SendMessage(Text1.hwnd, EM_LINESCROLL, 0, 100)
    Exit Function
Error_RefreshDisplay:
    Select Case Err.Number
    Case Is = 401   'cant show nonmodal when modal displayed
                    'retry until modal form is closed
    Case Is = 6     'overflow (text1.text too big)
        MsgBuffer = ""
        Text1 = ""
        Text1 = "RefreshDisplay Error " & Str(Err.Number) & " " & Err.Description & vbCrLf
    Case Else
        Text1 = Text1 & "RefreshDisplay Error " & Str(Err.Number) & " " & Err.Description & vbCrLf
    End Select
End Function

#If False Then
' Make the TextBox fit its contents.
'http://www.vb-helper.com/howto_size_textbox.html
Private Sub FitTextBoxContents(ByVal txt As Textbox)
    Font = Text1.Font
    txt.Width = TextWidth(txt.Text) + 120
    txt.Height = TextHeight(txt.Text) + 120
End Sub

'http://vbnet.mvps.org/index.html?code/textapi/txscroll.htm
Function ScrollText(Textbox As Control, vLines As Integer) As Long
Dim Success As Long
Dim SavedWnd As Long
Dim moveLines As Long
   'save the window handle of the control that currently has focus
    SavedWnd = Screen.ActiveControl.hwnd
    moveLines = vLines
   'Set the focus to the passed control (text control)
    Textbox.SetFocus
   'Scroll the lines.
    Success = SendMessage(Textbox.hwnd, EM_LINESCROLL, 0, ByVal moveLines)
   'Restore the focus to the original control
    Call PutFocus(SavedWnd)
   'Return the number of lines actually scrolled (INCORRECT)
    ScrollText = Success
End Function
#End If

Function GetLineCount(Textbox As Control) As Long
Dim lCount As Long
'The EM_GETLINECOUNT message retrieves the total number of text
'lines, not just the number of lines that are currently visible.
'If the Wordwrap feature is enabled, the number of lines can
'change when the dimensions of the editing window change.

        lCount = SendMessage(Textbox.hwnd, EM_GETLINECOUNT, 0, 0)
    GetLineCount = lCount
End Function
    
Private Sub Form_Load()
'You must NOT set scale explicitly as we are calculationg the height
    
'The icon needs setting up from a file. If you try loading it
'at design time - it just keeps the file location. This will
'be different when a user tries.
'    Me.Icon = LoadPicture(NmeaRouterIcon)
    
    FrmBorderWidth = Width - ScaleWidth
    FrmBorderHeight = Height - ScaleHeight
    Text1.Top = TEXTBOX_PADDING
    Text1.Left = TEXTBOX_PADDING
    MaxTextHeight = Screen.Height / 2 - FrmBorderHeight - TEXTBOX_PADDING * 2
    MaxTextWidth = Screen.Width / 2 - FrmBorderWidth - TEXTBOX_PADDING * 2
    If MaxTextWidth < MIN_TEXT_WIDTH Then MaxTextWidth = MIN_TEXT_WIDTH
    Me.BackColor = vbWhite
    Text1.BorderStyle = vbBSNone    'must be set at design time
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Call HideMe
    End If
End Sub

Private Sub RefreshTimer_Timer()
Dim SavedWnd As Long
    
   'save the window handle of the control that currently has focus
    SavedWnd = Screen.ActiveControl.hwnd
    RefreshTimer.Enabled = False
    HideMeTimer.Enabled = True   'restart
    Call RefreshDisplay
    Call PutFocus(SavedWnd)
End Sub

Private Sub HideMeTimer_Timer()
    HideMeTimer.Enabled = False
'Dont hide if were just waiting for the next update
    If RefreshTimer.Enabled = False Then
        Call HideMe
    End If
End Sub

Private Sub HideMe()
    On Error GoTo Modal_Error
    Me.Hide
    MsgBuffer = ""
    Text1 = ""
    Exit Sub

Modal_Error:
    HideMeTimer.Enabled = True  'retry
End Sub


