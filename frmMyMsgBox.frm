VERSION 5.00
Begin VB.Form frmMyMsgBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMyMsgBox"
   ClientHeight    =   2868
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   3972
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2868
   ScaleWidth      =   3972
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmMyMsgBox.frx":0000
      Top             =   480
      Width           =   2652
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1596
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   780
   End
End
Attribute VB_Name = "frmMyMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DrawText _
 Lib "user32.dll" Alias "DrawTextA" ( _
 ByVal hdc As Long, _
 ByVal lpStr As String, _
 ByVal nCount As Long, _
 ByRef lpRect As RECT, _
 ByVal wFormat As Long) As Long
 
Private Const DT_CALCRECT As Long = &H400
 
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Sub Form_Click_unused()
   Dim HalfWidth, HalfHeight, Msg   ' Declare variable.
   AutoRedraw = -1   ' Turn on AutoRedraw.
   BackColor = QBColor(4)   ' Set background color.
   ForeColor = QBColor(15)   ' Set foreground color.
   Msg = "Visual Basic"   ' Create message.
   FontSize = 48   ' Set font size.
   HalfWidth = TextWidth(Msg) / 2     ' Calculate one-half width.
   HalfHeight = TextHeight(Msg) / 2   ' Calculate one-half height.  object.TextHeight(string)
   CurrentX = ScaleWidth / 2 - HalfWidth   ' Set X.
   CurrentY = ScaleHeight / 2 - HalfHeight   ' Set Y.
kb = frmMyMsgBox.CurrentX
   frmMyMsgBox.Print Msg   ' Print message.
End Sub

Private Sub cmdOK_Click()
    ClickedButton vbOK
    Visible = False
End Sub

Private Sub cmdCancel_Click()
    ClickedButton vbCancel
    Visible = False

End Sub

Private Sub Form_Load()
    Appearance = 1  '3D
'    BorderStyle = 0 '=None  (Must be set at design time)
    Text1.BorderStyle = vbsnone '=0 removes padding line
    Text1.Appearance = 0 'not 3D
'    text1.MultiLine=true    (Must be set at design time)
    WindowState = vbNormal
    Visible = False     'Must be shown modal
'To test do after loaded so you can view each operation
'Scale the Msg Form to 1/3 screen height
    DisplayScreenHeight = Screen.Height / 3     ' 11520 'Twips of Acer (768 Px)
    Call ScaleForm(Me, DisplayScreenHeight / Height)

'centre on screen
'   Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
'Style must be Graphical for colours
    cmdOK.BackColor = vbYellow
    cmdCancel.BackColor = vbYellow
    
    
End Sub

Public Sub AdjustHeight_unused(txtBox As Textbox)
    Dim sText As String
    Dim r As RECT
    Dim nHeight As Long
    
    'adjust the scale mode for easier calculations
    Me.ScaleMode = vbPixels
    r.Right = txtBox.Width - 4 ' -4 px for the border, assumes 3d style is used
    sText = txtBox.Text
    nHeight = DrawText(Me.hdc, sText, Len(sText), r, DT_CALCRECT)
    If nHeight Then
        txtBox.Height = nHeight + 6
    End If
End Sub

