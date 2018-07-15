Attribute VB_Name = "modMyMsgBox"
Option Explicit

Dim Ret As Long
Dim DisplayScreenHeight As Long
Dim LineCount As Long

Const TEXTBOX_PADDING = 50# 'otherwise text runs into form border

Public Function MyMsgBox(Prompt As String, Optional Buttons As Long = 0, Optional Title As String) As Long
Dim ButtonValue As Long     '1st 4 bits
Dim IconValue As Long       'Bits 5 6 7
Dim Mask As Long
Dim kb As String

'MsgBox Prompt, Buttons, Title  'test

With frmMyMsgBox
'extract 1st 4 bits from RHS
    Mask = (2 ^ 4 - 1)  '(1111) = 15
    ButtonValue = Buttons And Mask
'Output 0-4 Extract next 3 bits and SHR 4 bits
'    IconValue = Buttons \ (2 ^ 4) 'SHR 4 bits 1110000 to 111 (div by 16 dropping remainder)
    Select Case ButtonValue
    Case Is = 0
        Debug.Print "OK only"
        .cmdOK.Visible = True
    Case Is = 1
        Debug.Print "OK and Cancel"
        .cmdOK.Visible = True
        .cmdCancel.Visible = True
    Case Is = 2
        Debug.Print "Abort, Retry, Cancel"
    Case Is = 3
        Debug.Print "Yes, No, Cancel"
    Case Is = 4
        Debug.Print "Yes and No"
    Case Is = 5
        Debug.Print "Retry and Cancel"
    Case Else
        Debug.Print "Invalid"
    End Select
    
    Mask = (2 ^ 7 - 1) - (2 ^ 4 - 1) '(1111) = 15
    IconValue = Buttons And Mask
    Select Case IconValue
    Case Is = 0
        .Text1.BackColor = vbWhite     '&H80000002
    Case Is = 16
        Debug.Print "Critical Message"
        .Text1.BackColor = vbRed
    Case Is = 32
        Debug.Print "Warning Query"
        .Text1.BackColor = RGB(255, 102, 0)
Case Is = 48
        Debug.Print "Warning Message"
    Case Is = 64
        Debug.Print "Information Message"
        .Text1.BackColor = vbGreen
    Case Else
        Debug.Print "invalid"
    End Select
    
    .Text1.Text = Prompt
    .Visible = True 'test only

'Scale the Msg Form to 1/3 screen height
'when OK do as part of the form load
'    DisplayScreenHeight = Screen.Height / 3     ' 11520 'Twips of Acer (768 Px)
'   Call ScaleForm(frmMyMsgBox, DisplayScreenHeight / frmMyMsgBox.Height)

    Call ScaleControlToForm(.Text1)

    frmMyMsgBox.Hide
    frmMyMsgBox.Show vbModal

    MyMsgBox = Ret
End With
End Function

Public Function ClickedButton(Arg As Long)
    Ret = Arg
End Function

Public Sub ScaleForm(MyForm As Form, ScreenScale As Single)
Dim ctrl As Control
Dim kb As String
Dim ScalePos As Boolean
Dim ScaleFont As Boolean
Dim ExceptHeight As Boolean
Dim ExceptWidth As Boolean

'    For Each MyForm In Forms
kb = MyForm.Name
        With MyForm
            .Height = .Height * ScreenScale
            .Width = .Width * ScreenScale
            .Top = .Top / ScreenScale
            .Left = .Left / ScreenScale
'centre on screen
            .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2

'The form font must be scaled because .Text1.TextHeight and .Text1.TextWidth uses the size of the font on the form
'not on the text box
            
Debug.Print MyForm.Name & " FontSize=" & MyForm.Font.Size
            .Font.Size = .Font.Size * ScreenScale
Debug.Print MyForm.Name & " FontSize=" & MyForm.Font.Size
        End With
'    Next MyForm

    For Each ctrl In MyForm
kb = ctrl.Name
        ScalePos = False
        ExceptHeight = False
        ExceptWidth = False
        If TypeOf ctrl Is Frame Then ScalePos = True
        If TypeOf ctrl Is CommandButton Then ScalePos = True
        If TypeOf ctrl Is Label Then ScalePos = True
        If TypeOf ctrl Is MSHFlexGrid Then ScalePos = True
        If TypeOf ctrl Is ComboBox Then
            ScalePos = True
            ExceptHeight = True
        End If
        If TypeOf ctrl Is Textbox Then
kb = ctrl.Name
            ScalePos = True
'determined by the text in frmMyMsgBox but not in SetStartTime
            If ctrl.Text <> "" Then 'If not blank dont scale size
                ExceptHeight = True
                ExceptWidth = True
            End If
        End If
        If TypeOf ctrl Is PictureBox Then ScalePos = True
        If TypeOf ctrl Is Image Then ScalePos = True
        If ctrl.Name = "fraState" Then
            ExceptHeight = True
            ExceptWidth = True
        End If
        If ctrl.Name = "lblState" Then ScalePos = False
        If ctrl.Name = "frmdaventech.lblProgram" Then ScalePos = False
        If ctrl.Name = "frmdaventech.lblsequence" Then ScalePos = False
        If ctrl.Name = "frmdaventech.lblrecalls" Then ScalePos = False
        If ScalePos Then
            With ctrl
'Combo box Height determined by text size
                If Not ExceptHeight Then .Height = .Height * ScreenScale
                If Not ExceptWidth Then .Width = .Width * ScreenScale
                .Top = .Top * ScreenScale
                .Left = .Left * ScreenScale
            End With
        End If
 MyForm.Refresh
       
       ScaleFont = False
        If TypeOf ctrl Is Frame Then ScaleFont = True
        If TypeOf ctrl Is CommandButton Then ScaleFont = True
        If TypeOf ctrl Is Label Then ScaleFont = True
        If TypeOf ctrl Is MSHFlexGrid Then ScaleFont = True
        If TypeOf ctrl Is ComboBox Then ScaleFont = True
        If TypeOf ctrl Is Textbox Then ScaleFont = True
        If TypeOf ctrl Is PictureBox Then ScaleFont = True
        If ctrl.Name = "fraState" Then ScaleFont = False
        If ctrl.Name = "lblState" Then ScaleFont = False
        If ctrl.Name = "frmdaventech.lblProgram" Then ScaleFont = False
        If ctrl.Name = "frmdaventech.lblsequence" Then ScaleFont = False
        If ctrl.Name = "frmdaventech.lblrecalls" Then ScaleFont = False
         If ScaleFont Then
           With ctrl
 Debug.Print .Name & " Font.Size=" & .Font.Size
                .Font.Size = .Font.Size * ScreenScale
 Debug.Print .Name & " Font.Size=" & .Font.Size
            End With
        End If

        If TypeOf ctrl Is MSHFlexGrid Then
            With ctrl
                .ColWidth(0) = .ColWidth(0) * ScreenScale 'Position
                .ColWidth(1) = .ColWidth(0) * ScreenScale 'Position
            End With
        End If
        
        If TypeOf ctrl Is Textbox Then
'            Call ScaleControlToForm(ctrl)  'Dont if set start time on racing signals
        End If
        
'        If TypeOf ctrl Is Timer Then
'            If ctrl.Enabled = True Then
'                kb = kb & ctrl.Name
'                On Error Resume Next
'                kb = kb & ctrl.Index
'                On Error GoTo 0
'                kb = kb & "."
'            End If
'        End If
    Next ctrl
End Sub

Public Sub ScaleControlToForm(ctrl As Control)

'set the form font to be same as text box because textwidth & textheight use the form font not textbox
    ctrl.Parent.Font.Size = ctrl.Font.Size
    ctrl.Parent.Font.Bold = ctrl.Font.Bold
    
'Set the TextBox height to include all lines
    LineCount = GetLineCount(ctrl)
    ctrl.Height = ctrl.Parent.TextHeight(ctrl.Text) + LineCount * TEXTBOX_PADDING

'Set TextBox width to include all characters + white space padding eother side of text
    ctrl.Width = ctrl.Parent.TextWidth(ctrl.Text) + 2 * TEXTBOX_PADDING '18 is padding either side of text  + ctrl.Parent.TextWidth("W")
    
'Set form width maintaining gap from LHS at design time at both sides of text box
    ctrl.Parent.Width = ctrl.Width + ctrl.Left * 2

'Set the form height
'centre on screen
    ctrl.Parent.Move (Screen.Width - ctrl.Parent.Width) / 2, (Screen.Height - ctrl.Parent.Height) / 2

End Sub

Function GetLineCount(Textbox As Control) As Long
Dim lCount As Long
    lCount = UBound(Split(Textbox.Text, vbCrLf)) + 1
    GetLineCount = lCount
End Function


