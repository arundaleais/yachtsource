VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFlxGd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   6240
   ClientLeft      =   192
   ClientTop       =   816
   ClientWidth     =   12876
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   12876
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.TextBox txtLog 
      Height          =   5652
      Left            =   9000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   35
      Top             =   480
      Visible         =   0   'False
      Width           =   3852
   End
   Begin VB.Timer UnloadTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7440
      Top             =   960
   End
   Begin VB.CommandButton Commands 
      Caption         =   "Commands"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Commands 
      Caption         =   "Commands"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1695
      Left            =   5400
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton cmdPause 
         Caption         =   "Pause"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdEvents 
         Caption         =   "Show Events"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   25
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Commands 
         Caption         =   "Commands"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblElapsedTime 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   27
         Top             =   1080
         Width           =   960
      End
   End
   Begin VB.Timer LoadProfileTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6960
      Top             =   0
   End
   Begin VB.Timer FinishTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   720
   End
   Begin VB.Timer ClearFlagsTimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7320
      Top             =   480
   End
   Begin VB.Timer HoistTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6720
      Top             =   960
   End
   Begin VB.Frame fraMain 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4370
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   7935
      Begin VB.Frame fraNextStart 
         Caption         =   "Next Start"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   16
         Top             =   2200
         Width           =   2655
         Begin VB.CommandButton Commands 
            Caption         =   "Commands"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   2
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1080
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Countdown"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   840
         End
         Begin VB.Label lblNextStartName 
            Caption         =   "None"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblTimeToNextStart 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            Caption         =   "00:00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1100
            TabIndex        =   17
            Top             =   540
            Width           =   960
         End
      End
      Begin VB.Frame fraPreviousStart 
         Caption         =   "Previous Start"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   3000
         TabIndex        =   15
         Top             =   1470
         Width           =   2655
         Begin VB.CommandButton Commands 
            Caption         =   "Commands"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   960
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.CommandButton Commands 
            Caption         =   "Commands"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   3
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1800
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Elapsed"
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
            Left            =   120
            TabIndex        =   22
            Top             =   555
            Width           =   735
         End
         Begin VB.Label lblTimeFromPreviousStart 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            Caption         =   "00:00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   840
            TabIndex        =   21
            Top             =   480
            Width           =   960
         End
         Begin VB.Label lblPreviousStartName 
            Caption         =   "None"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame fraStartFinish 
         Caption         =   "Start/Finish Times"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4120
         Left            =   5760
         TabIndex        =   10
         Top             =   120
         Width           =   2055
         Begin VB.CommandButton Commands 
            Caption         =   "Commands"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   5
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   3150
            Visible         =   0   'False
            Width           =   1815
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshFinish 
            Height          =   2745
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1695
            _ExtentX        =   2985
            _ExtentY        =   4847
            _Version        =   393216
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraTime 
         Caption         =   "Current Time"
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
         Left            =   3000
         TabIndex        =   8
         Top             =   120
         Width           =   2655
         Begin VB.Label lblCurrTime 
            AutoSize        =   -1  'True
            BackColor       =   &H80000014&
            Caption         =   "00:00:00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   2460
         End
      End
      Begin VB.Frame fraPreparatory 
         Caption         =   "Preparatory"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1970
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2655
         Begin VB.Frame fraNoOfStarts 
            Caption         =   "Starts"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   120
            TabIndex        =   34
            Top             =   1000
            Width           =   852
            Begin VB.TextBox txtNoOfStarts 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   13.8
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   468
               Left            =   120
               TabIndex        =   2
               Top             =   240
               Width           =   612
            End
         End
         Begin VB.Frame fraStartSequence 
            Caption         =   "Start Sequence"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   2415
            Begin VB.ComboBox cboProfile 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   0
               ToolTipText     =   "Select a Start Sequence"
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.Frame fraFirstStartTime 
            Caption         =   "First Start Time"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   1080
            TabIndex        =   7
            Top             =   1000
            Width           =   1452
            Begin VB.TextBox txtFirstStartTime 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   468
               Left            =   120
               MaxLength       =   4
               TabIndex        =   3
               ToolTipText     =   "Enter Start Time as 24 Hour eg 1230"
               Top             =   240
               Width           =   1215
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1572
      ScaleWidth      =   6012
      TabIndex        =   4
      Tag             =   "2"
      Top             =   0
      Width           =   6015
      Begin VB.Label lblStartTime 
         BackColor       =   &H80000009&
         Caption         =   "Please Set a Start Time"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   2595
      End
      Begin VB.Label lblStartSequence 
         BackColor       =   &H80000009&
         Caption         =   "Please select a Start Sequence"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1320
         TabIndex        =   13
         Top             =   360
         Width           =   3255
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   40
         Left            =   5400
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   39
         Left            =   4800
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   38
         Left            =   4200
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   37
         Left            =   3600
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   36
         Left            =   3000
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   35
         Left            =   2400
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   34
         Left            =   1800
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   33
         Left            =   1200
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   32
         Left            =   600
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   31
         Left            =   0
         Top             =   1080
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   30
         Left            =   5400
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   29
         Left            =   4800
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   28
         Left            =   4200
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   27
         Left            =   3600
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   26
         Left            =   3000
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   25
         Left            =   2400
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   24
         Left            =   1800
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   23
         Left            =   1200
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   22
         Left            =   600
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   21
         Left            =   0
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   20
         Left            =   5400
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   19
         Left            =   4800
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   18
         Left            =   4200
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   17
         Left            =   3600
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   16
         Left            =   3000
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   15
         Left            =   2400
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   14
         Left            =   1800
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   13
         Left            =   1200
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   12
         Left            =   600
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   11
         Left            =   0
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   10
         Left            =   5400
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   9
         Left            =   4800
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   8
         Left            =   4200
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   7
         Left            =   3600
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   6
         Left            =   3000
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   5
         Left            =   2400
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   4
         Left            =   1800
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   3
         Left            =   1200
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   2
         Left            =   600
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Flags 
         Height          =   375
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   615
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      _Version        =   393216
   End
   Begin VB.Timer SignalTimer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   6600
      Top             =   360
   End
   Begin VB.Timer RaceTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   1
      Top             =   5988
      Width           =   12876
      _ExtentX        =   22712
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11959
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Settings 
      Height          =   492
      Left            =   7200
      ToolTipText     =   "Change Settings"
      Top             =   0
      Width           =   720
   End
   Begin VB.Image ClubBurgee 
      Height          =   1116
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1656
   End
   Begin VB.Image Flags 
      Height          =   375
      Index           =   0
      Left            =   6600
      Top             =   600
      Width           =   615
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuEvents 
         Caption         =   "&Events"
      End
      Begin VB.Menu mnuController 
         Caption         =   "&Controller"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "&Log"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSpeed 
         Caption         =   "&Speed"
         Begin VB.Menu mnux10 
            Caption         =   "Times 10"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Cancel As Boolean
Private reply As Long
Private KeyIsPressed As Boolean

Private Type defhms
    Hour As Long
    Min As Long
    Sec As Long
End Type
Private Type defCols
    group As String    'Dynamic, Index is ColCount
    Items As Long        'Dynamic
End Type
Private DisplayScreenHeight As Long
Private ScreenScale As Single

Private BaseWidth As Single 'Form Width with no Commands
Private FixedCommandIdx As Long   'No of initial commands on the form
Private Cols() As defCols
Private NextFreeCol As Long
Private NextCommandTop As Single    'To set Command button positions in LoadProfile
Public FirstStartTime As Date  'To calculate elapsed time (MUST NOT BE CHANGED)
                                'because it will upset the Offset calculation and
                                'cause the catchup to not work
Private StartTimeValid As Boolean
Private PreviousTimeOutput As Date  'Used for catch-up
'Private PausedSecs As Long
Private CmdQ(8) As Integer     'Idx of next signal (if timer)
Private PreviousHoist As String  'Group of Previous flag hoisted (Timer Suppresses sound signal)
Private IsClassStartEvent As Boolean   'True if there's been a Class Start at this Elapsed Time
                                'To suppress the sound signal if a Class Warning is
                                'Raised after a class start at the same elapsed second
Private FinishCount As Long     'No of Finishers Clocked
Private FinishSignalCount As Long   'No of FinishSignals made
'Private PostponeIdx As Long 'The Current Postpone Class - changes at the start
'Private PostponeClass As Long   'The First Class that will be postponed, if the
                                'Postpone Flag is raised.
                                'It is the next class to start
Private RecallIdx As Long   'The Recall Class Flag Set when
'Private EventTime As Long   'The time used for the current event we are processing
                            'This is passed to DoTimerEvents for EVERY second
                            'In between events (when a Command is clicked)
                            'this is the PreviousEventTime
Private PostponeCountDown As Long   'Seconds before Postpone Flag will be dropped
Private USBButton As New clsUSBButton

'Private Paused As Boolean

Private Sub cboProfile_Click()
    
    WriteLog "Profile_Click " & cboProfile
    If cboProfile = "Terminate" Then
'Cannot terminate directly from a call in a Combo Box
'Unable to unload within this context (Error 365)
        UnloadTimer.Enabled = True
        Exit Sub
    End If
    
    If Cancel = True Then   'Called by the UnloadTimer to Cancel Terminate
        Cancel = False
        Exit Sub    'Because the Program State remains the same
    End If
    
    If cboProfile.ListIndex >= 0 Then
        CurrentProfile = cboProfile.List(cboProfile.ListIndex)
        Call SetProgramState(1) 'Load Profile
    Else
        Call SetProgramState(0) 'Loading Profile
    End If
End Sub

'Only used to clear all the flags off the display 3 secs after loading the profile
Private Sub ClearFlagsTimer_Timer_old()
    Loading = True
    Call DefaultsPreStartTimeSet
    Loading = False
'Must do after defaults to stop PreviousStart flag being set
    ClearFlagsTimer.Enabled = False
End Sub

Private Sub cboProfile_GotFocus()
Debug.Print "cboProfile"
End Sub


'Used to clear all the flags off the display 3 secs after loading the profile
'Sets the default settings when the Splash screen timer has finished and before
'a valid start time has been entered
Private Sub ClearFlagsTimer_Timer()
Dim Idx As Long
Dim kb As String

Debug.Print "ClearFlags_Timer"
WriteLog "ClearFlags_Timer Fired"
    RecallSignalIdx = 0  'Remove so that logic for recall is not applied when
                        'the flags are lowered
    For Idx = 1 To UBound(SignalAttributes)
        With SignalAttributes(Idx)
            If SignalAttributes(Idx).Flag.pos <> 0 Then
                Call LowerFlag(Idx)
            End If
        End With
    Next Idx
'Reset
    RecallSignalIdx = SignalFromName("Recall Class")
'removed    Call DisplayStartTimes
'Dont allow start time to be changed until splash screen has been removed
'    txtFirstStartTime.Enabled = True
'Position cursor at RHS
'    txtFirstStartTime.SelStart = Len(txtFirstStartTime)
'    txtFirstStartTime.SetFocus
'txtFirstStartTime.SelStart = 2
    lblStartTime.Visible = True
    HoistTimer.Enabled = False
    PreviousHoist = ""  'Group
'    Call ResetRecall
    Call ResetCols

    Loading = False
'Must do after defaults to stop PreviousStart flag being set
    ClearFlagsTimer.Enabled = False
'Enable USB Button
    Set USBButton = New clsUSBButton

      If frmDaventech.Winsock1.state = sckConnected Then
        kb = "Battery Voltage = " & BatteryVoltage
        Select Case BatteryVoltage
        Case Is > "12.5"
            reply = MyMsgBox(kb, vbOKOnly + vbInformation)  'green
        Case Is > "12.0"
            kb = kb & vbCrLf & vbCrLf & "This is low, please check battery"
            reply = MyMsgBox(kb, vbOKOnly + vbQuestion)      'orange
        Case Is > "0.0"
            kb = kb & vbCrLf & vbCrLf & "This is very low, please check battery." _
            & vbCrLf & "Horn and/or Lights may not work." _
            & vbCrLf & "Prepare to use VHF and manual Horn"
            reply = MyMsgBox(kb, vbOKOnly + vbCritical)      'red
        Case Else       'No battery connected -  Controller will not work
            reply = MyMsgBox("Battery is not connected." & vbCrLf _
            & "No signals (eg Lights/Horn) connected to controller will work." & vbCrLf _
            & "You will have to use VHF and manual Horn", vbOKOnly)
        End Select
     Else
        reply = MyMsgBox("Controller is not connected to " & GetComputerName & "." & vbCrLf & vbCrLf _
        & "No signals (eg Lights/Horn) connected to the controller will work." & vbCrLf _
        & "You will have to use VHF and manual Horn", vbOKOnly)
      End If
  
'    Call SetState("Program", 2)
'Dont set the FirstStart Class until we have a start time
'Dont allow events to be manually called until program state is 2
'Otherwise Clcicking the first event will not display the Elapsed time
#If jnasetup = True Then
    Call frmEvents.ListEvents
    frmDaventech.Visible = True
#End If
    Call SetProgramState(2) 'Profile Loaded & flags cleared
End Sub


Private Sub ClubBurgee_Click()
'    With frmDaventech
'        If .Visible = False Then
'        .Visible = True
'        Else
'            .Visible = False
'        End If
'    End With
End Sub

Private Sub cmdEvents_Click()
    If frmEvents.Visible Then
        frmEvents.Visible = False
    Else
        Call frmEvents.ListEvents
        frmEvents.Visible = True
    End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

'Only used for testing (jnasetup)
Private Sub cmdPause_Click()
    If RaceTimer.Enabled = True Then
    With cmdPause
        Select Case .BackColor
        Case Is = cbEnabled     'Idle
            .BackColor = vbGreen    'Commence Pause on next event
        Case Is = vbCyan        'Paused
            .BackColor = vbGreen     'Remove pause on next whole minute
        End Select
    End With
    Else
        Call DoTimerEvents  'Already incremented to next event
    End If
End Sub

Private Sub Commands_Click(Index As Integer)
Dim Position As Long
Dim NextCommand As Long

Debug.Print "--- " & Commands(Index).Caption & " ---"
WriteLog "Command " & Commands(Index).Caption & " (" & Timer & ")"

'Clear the message user may be playing with the flags
    
    lblStartSequence.Visible = False
    With SignalAttributes(Index)
'If this command is queued then just remove it (same as clicking when up)
'This must be done in the Click event because the user is making the request
'You cannot do it in RaiseRequest or LowerRequest because all queued events
'would get removed.
        If Commands(Index).BackColor = vbCyan Then
            NextCommand = DequeCmd(.group)
            Exit Sub
        End If
'If we have another commandButton queued in this group, remove this before
'actioning a raise request so we dont have 2 flags in same group queued
'This is important with Recall & General Recall
        If .Flag.pos = 0 Then
            NextCommand = DequeCmd(.group)
'We must set the Recall States using Go...
            If state.Program >= 4 Then  '4=Start Sequence Running
                Select Case Commands(Index).Caption
                Case Is = "Recall"
                    Call GoRaiseRecall
                Case Is = "General Recall"
                    Call GoRaiseGeneralRecall
                Case Else
                    Call RaiseRequest(CLng(Index))
                End Select
            Else
                Call RaiseRequest(CLng(Index))
            End If
            If Commands(Index).Caption = "Postpone" Then
'Causes StartTime to be validated if one has not yet been entered
'Which then causes Events to be reloaded and hence Postponed start time will be set
                If StartTimeValid = False Then
                    StartTimeValid = True
                End If
            End If
        Else
'We must set the Recall States using Go...
            If state.Program >= 4 Then  '4=Start Sequence Running
                Select Case Commands(Index).Caption
                Case Is = "Recall"
                    Call RecallTimeout
'                    Call GoLowerRecall
                Case Is = "General Recall"
                    Call RecallTimeout
'                    Call GoLowerGeneralRecall
                Case Else
                    Call LowerRequest(CLng(Index))
                End Select
            Else
                Call LowerRequest(CLng(Index))
            End If
        End If
    End With
End Sub

Private Sub Flags_Click(Index As Integer)
'MsgBox Flags(Index).Picture.Handle
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim FinishIdx As Long
Debug.Print Time & " Down"
    If KeyIsPressed = False Then
        KeyIsPressed = True
        FinishIdx = CommandFromCaption("Finish")
        If KeyCode = 119 Then   'F8
            If Commands(FinishIdx).Enabled = True Then
                Commands_Click (FinishIdx) 'Generate Click event
            End If
        End If
    End If
End Sub


'Call Load/unload functions if Other actions required dependant on state
'or Unload conditional and can be called from Evts
Private Sub Form_KeyDown_notused(KeyCode As Integer, Shift As Integer)
Dim kb As String
Dim Idx As Long
Dim KeyNo As Long
Dim NextCommand As String
    
    For KeyNo = 1 To UBound(Keys)
        If KeyCode = Keys(KeyNo).Code Then Exit For
    Next KeyNo
        
    If KeyNo > UBound(Keys) Then
        If KeyCode > 111 Then       'Function key
MsgBox "Function Key <F" & KeyCode - 111 & "> is not in use", vbInformation, "KeyDown"
        Else
#If jnasetup = True Then
'If StartTime is set, but Sequence not started we can Postpone and Change the Start Time
            If state.Program <> 3 Then
                MsgBox "Undefined Key", , "Form_KeyDown"
            End If
#End If
        End If
        Exit Sub    'ignore KeyPreview
    End If

'0=No Profile,1=Loading Profile,2=Profile Loaded
'3=Start Time Set,4=Start Sequence Running,5=Start Sequence Finished

    Select Case aKeyState(Keys(KeyNo).state)
    Case Is = "Postpone"
'syc dont use postpone
'        If Keys(KeyNo).Cancel = False Then
'            Call GoRaisePostpone
'        Else
'            Call GoLowerPostpone
'        End If
'        Call GoToggleKey(KeyNo)
    Case Is = "Recall"
        If Keys(KeyNo).Cancel = False Then
            Call GoRaiseRecall
        Else
'syc            Call GoLowerRecall
            Call RecallTimeout  'Must use to set Finish if Previous sequence
        End If
        Call GoToggleKey(KeyNo)
    Case Is = "General Recall"
        If Keys(KeyNo).Cancel = False Then
            Call GoRaiseGeneralRecall
        Else
            Call RecallTimeout  'Must use to set Finish if Previous sequence
'            Call GoLowerGeneralRecall
        End If
        Call GoToggleKey(KeyNo)
    Case Is = "Finish"
        Idx = SignalFromName("Finish")
'This MUST be done by clicking the command button., otherwise duplicates the time (sometimes)
'You Click the Command Button by setting Value=True
        Commands(Idx).Value = True
    Case Else
        If Keys(KeyNo).KeyName = "F12" Then
            Idx = SignalFromName("Horn Short")
            Commands(Idx).Value = True
        End If
    End Select
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim FinishIdx As Long
Debug.Print Time & " Up"
    KeyIsPressed = False
End Sub

Private Sub Form_Load()
Dim i As Long
Dim url As String
Dim Major As Long
Dim Minor As Long
Dim Revision As Long
Dim NewVersion As Boolean
Dim cmd As CommandButton
Dim kb As String

'    Width = Screen.Width
'    Height = Screen.Height
    Top = 0
    Left = 0
    
    For Each cmd In Commands
        ReDim Preserve StaticCommands(cmd.Index)
        StaticCommands(cmd.Index) = True
    Next cmd
    
#If jnasetup = False Then
    Commands(0).Visible = False
    cmdEvents.Visible = False
    cmdPause.Visible = False  'Only used for testing (jnasetup)
    lblElapsedTime.Visible = False
    StatusBar1.Panels(1).Bevel = sbrNoBevel
    StatusBar1.Panels(2).Bevel = sbrNoBevel
    StatusBar1.Panels(3).Bevel = sbrNoBevel
    StatusBar1.Panels(4).Bevel = sbrNoBevel
    StatusBar1.Panels(5).Bevel = sbrNoBevel
    txtLog.Visible = False
    txtLog.Enabled = False

#End If

    Caption = App.EXEName & " [" & App.Major & "." & App.Minor & "." _
    & App.Revision & "] "
    WriteLog Caption
    
    Caption = Replace(Caption, ".exe", "")
'Check if a later version exists
    url = "http://arundale.com/docs/ais/" & App.EXEName _
    & "_setup_"
    WriteLog "Checking " & url & " for later version"
    
    Major = App.Major
    Do
        If HTTPFileExists(url & Major & ".0.0.exe") = False Then Exit Do
        Major = Major + 1
    Loop
    If Major > 0 Then Major = Major - 1   'Highest major that exists
    
    url = url & Major & "."
    If Major = App.Major Then
        Minor = App.Minor
    Else
        Minor = 0
    End If
    Do
        If HTTPFileExists(url & Minor & ".0.exe") = False Then Exit Do
        Minor = Minor + 1
    Loop
    If Minor > 0 Then Minor = Minor - 1

    WriteLog "Latest released version is " & url & Minor & ".0"
    url = url & Minor & "."
    If Not (Major = App.Major And Minor = App.Minor) Then
        NewVersion = True
    End If
'Only let a user get next revision if he is using a revision
'of his current version. Otherwise he goes up to the next minor version
    If NewVersion = False And App.Revision > 0 Then
        WriteLog "User Revision is " & App.Revision & " (>0 so check web revision)"
        Revision = App.Revision
        Do
            If HTTPFileExists(url & Revision & ".exe") = False Then Exit Do
            Revision = Revision + 1
        Loop
        If Revision > 0 Then Revision = Revision - 1
        WriteLog "Web Revision is " & Revision
        If Revision < App.Revision Then
            WriteLog "Web revision > User Revision, user can be updated"
            NewVersion = True
        End If
    Else
        WriteLog "User Revision is " & App.Revision & " (=0 so don't check web revision)"
    End If
    url = url & Revision & ".exe"
    
'If we are working on a higher version in VBE, don't try for newversion
    If App.Major * 2 ^ 8 + App.Minor * 2 ^ 4 + App.Revision > _
    Major * 2 ^ 8 + Minor * 2 ^ 4 + Revision Then
        NewVersion = False
    End If
    If NewVersion = True Then
        WriteLog ("A new update is available")
        Call frmDpyBox.DpyBox("A new update is available", 10, "New Version")
'Check we have internet access
        If HTTPFileExists(url) Then
            Call HttpSpawn(url)
        End If
    Else
        WriteLog ("A new update is not available")
    End If

'Load the available Profiles
    With cboProfile
        kb = Dir$(SequencesFilePath & "*.ini")
        Do While kb > ""
'Dont allow *.ini_old
            If Right$(kb, 4) = ".ini" Then
'Remove .ini so it's not displayed
                i = InStrRev(kb, ".ini")
                If i > 0 Then
                    kb = Left$(kb, i - 1)
                    .AddItem kb
WriteLog "Profile (" & .ListCount & ") " & kb & " added"
                End If
            End If
            kb = Dir$
        Loop
'If none exit sub & program (in .main)
        If .ListCount = 0 Then
            WriteLog "No Start Sequences available, Exiting Program, No Start Sequence"
            MsgBox "No Start Sequences available in" & vbCrLf & SequencesFilePath & vbCrLf & "Exiting Program", vbCritical, "No Start Sequence"
            Exit Sub
        End If

'Add an exit
        .AddItem "Terminate"

'Stop the blank box being set in Blue
        .ListIndex = -1
    .BackColor = cbEnabled
    End With
    
'Must set the date  to prevent overflow on datediff
    FirstStartTime = Date
WriteLog "Position Window"
    WindowState = vbNormal
'    Me.Left = NulToZero(GetSetting(App.Title, "Settings", "MainLeft"))
'    Me.Top = NulToZero(GetSetting(App.Title, "Settings", "MainTop"))
'    Me.Width = GetSetting(App.Title, "Settings", "MainWidth")
'    Me.Height = GetSetting(App.Title, "Settings", "MainHeight")
    Visible = True
    
WriteLog "Set Finish Display Size"
'Set the size of the finish display
    With mshFinish
'Remove dotted focus rectangle (but causes Blue)
        .FocusRect = flexFocusNone
        .Width = fraStartFinish.Width - 250
        .FormatString = "<No|<Time"
        .ColWidth(0) = 500 'Position
        .ColWidth(1) = fraStartFinish.Width ' - 500 - 30   '1295  'Time
'        For i = 1 To 20
'            .Rows = i + 1
'            .TextMatrix(i, 0) = i
'        Next i
'        .TextMatrix(1, 1) = "13:22:45"
    End With
    
    Call mshFinish_SelChange    'Remove blue focus rectangle

WriteLog "Set Col array"
'Set the Cols array
'Flags(0) exists - but not used
    RowCount = FlagRow(Flags.Count - 1)
    ColCount = FlagCol(Flags.Count - 1)
    ColCountFree = ColCount 'Reduces by number of Fixed cols
    ReDim Cols(1 To ColCount)

    StatusBar1.Panels(4).Width = 200
    StatusBar1.Panels(5).Width = 200
    
    lblStartSequence.BackStyle = vbTransparent
    lblStartSequence.Visible = True
    lblStartTime.BackStyle = vbTransparent
    lblStartTime.Visible = False
    
    WindowState = vbNormal
    Visible = True

'Programmers Guide P617
'    ScaleMode = vbTwips
'    Height = Height - ScaleHeight + ScaleY(768, vbPixels)
'    Width = Width - ScaleWidth + ScaleX(1024, vbPixels)
'    ScaleMode = vbPixels
'Scale the Display Screen
    DisplayScreenHeight = 11520 'Twips of Acer (768 Px)
    Call ScaleForm(frmMain, DisplayScreenHeight / Height)
    Width = Screen.Width    'SYC fill to full width
     BaseWidth = Width
   
'Resize to get finish tomes in (after scaling)
    With mshFinish
        .Width = fraStartFinish.Width - 350
'        .FormatString = "<No|<Time"
'        .ColWidth(0) = 500  'Position
        .ColWidth(1) = .Width - .ColWidth(0) 'Time
    End With
      
 #If False Then
    ScreenScale = DisplayScreenHeight / Height
    Height = Height * ScreenScale
    Width = Width * ScreenScale
    With fraMain
        .Height = .Height * ScreenScale
        .Width = .Width * ScreenScale
        .Top = .Top * ScreenScale
        .Left = .Left * ScreenScale
    End With
    With fraState
        .Height = .Height * ScreenScale
        .Width = .Width * ScreenScale
        .Top = .Top * ScreenScale
        .Left = .Left * ScreenScale
    End With
    With fraStartFinish
        .Height = .Height * ScreenScale
        .Width = .Width * ScreenScale
        .Top = .Top * ScreenScale
        .Left = .Left * ScreenScale
    End With
        With mshFinish
            .Height = .Height * ScreenScale
            .Width = .Width * ScreenScale
            .Top = .Top * ScreenScale
            .Left = .Left * ScreenScale
            .Font.Size = .Font.Size * ScreenScale
            .ColWidth(0) = .ColWidth(0) * ScreenScale 'Position
            .ColWidth(1) = .ColWidth(0) * ScreenScale 'Position
        End With
    With fraPreparatory
        .Height = .Height * ScreenScale
        .Width = .Width * ScreenScale
        .Top = .Top * ScreenScale
        .Left = .Left * ScreenScale
    End With
        With fraStartSequence
            .Height = .Height * ScreenScale
            .Width = .Width * ScreenScale
            .Top = .Top * ScreenScale
            .Left = .Left * ScreenScale
        End With
            With cboProfile
'                .Height = .Height * ScreenScale
'With combo box Font size determines height
                .Width = .Width * ScreenScale
                .Top = .Top * ScreenScale
                .Left = .Left * ScreenScale
                .Font.Size = .Font.Size * ScreenScale
            End With
        With fraFirstStartTime
            .Height = .Height * ScreenScale
            .Width = .Width * ScreenScale
            .Top = .Top * ScreenScale
            .Left = .Left * ScreenScale
        End With
            With txtFirstStartTime
                .Height = .Height * ScreenScale
                .Width = .Width * ScreenScale
                .Top = .Top * ScreenScale
                .Left = .Left * ScreenScale
                .Font.Size = .Font.Size * ScreenScale
            End With
    With fraPreviousStart
        .Height = .Height * ScreenScale
        .Width = .Width * ScreenScale
        .Top = .Top * ScreenScale
        .Left = .Left * ScreenScale
    End With
    With fraNextStart
        .Height = .Height * ScreenScale
        .Width = .Width * ScreenScale
        .Top = .Top * ScreenScale
        .Left = .Left * ScreenScale
    End With
    With fraTime
        .Height = .Height * ScreenScale
        .Width = .Width * ScreenScale
        .Top = .Top * ScreenScale
        .Left = .Left * ScreenScale
    End With
        With lblCurrTime
            .Height = .Height * ScreenScale
            .Width = .Width * ScreenScale
            .Top = .Top * ScreenScale
            .Left = .Left * ScreenScale
            .Font.Size = .Font.Size * ScreenScale
        End With
    With Commands(0)
        .Height = .Height * ScreenScale
        .Width = .Width * ScreenScale
        .Top = .Top * ScreenScale
        .Left = .Left * ScreenScale
    End With
#End If
'Reset the Profile variables
'    Call ClearProfile  'Called when loaded
'Set up initial start time, LoadEvents not called
'    FirstStartTime = Date & " " _
'    & Format$(NulToZero(txtFirstStartTime), "00:00") & ":00"
    txtFirstStartTime.Enabled = False
    txtFirstStartTime.BackColor = cbDisabled
    txtNoOfStarts.Enabled = False
    txtNoOfStarts.BackColor = cbDisabled
Debug.Print Format$(FirstStartTime, "dd-mmm-yyyy")
Debug.Print Format$(FirstStartTime, "hh:mm:ss")

'    Call LoadSequence
WriteLog "frmMain loaded"
End Sub



Private Sub Form_Resize()
'Caption = Height & " x " & Width & " scrn=" & Screen.Height & " x " & Screen.Width

'Caption = ""    'SYC only to not display Tiltle bar

End Sub


Private Sub lblStartTime_Click()
MsgBox "lblstarttime_click"
End Sub

'Must use to Load the profile
Private Sub LoadProfileTimer_Timer()
WriteLog "LoadProfileTimer_Fired"
    LoadProfileTimer.Enabled = False
    Call LoadProfile
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer
Dim ch As Long

 'Stop any events generated by DoTimerEvents interupting the Unloading
    RaceTimer.Enabled = False
'Clear any Controller Relays on when program is terminated (if connected)
    Unload frmDaventech
    Call EncryptFiles(SequencesFilePath, ".txt", ".ini")
    
'MsgBox "Save Results"
    
                
    If state.Program >= 3 Then      'start sequence started
Dim kb As String
Dim Row As Long
Dim Col As Long
'If full pathe not used .log is OK but .csv fails to write out to default directory
'MsgBox ResultsFileName
        ResultsFileCh = FreeFile
        Open ResultsFileName For Output As #ResultsFileCh
        kb = ""
        For Row = 0 To mshFinish.Rows - 1
            mshFinish.Row = Row
            For Col = 0 To mshFinish.Cols - 1
                mshFinish.Col = Col
                If Col = 0 Then
                    kb = kb & mshFinish.TextMatrix(Row, Col)
                Else
                    kb = kb & "," & mshFinish.TextMatrix(Row, Col)
                End If
            Next Col
            Print #ResultsFileCh, kb
            kb = ""
        Next Row
        Close ResultsFileCh
'        Name ResultsFileName As Replace(ResultsFileName, ".log", ".csv")
    End If
    Me.WindowState = vbNormal
    SaveSetting App.Title, "Settings", "MainLeft", Me.Left
    SaveSetting App.Title, "Settings", "MainTop", Me.Top
    SaveSetting App.Title, "Settings", "MainWidth", Me.Width
    SaveSetting App.Title, "Settings", "MainHeight", Me.Height
'Must NOT reference frmDaventech as this will cause a reload and open of winsock
'Debug.Print "frmain.Unload " & frmDaventech.Winsock1.State
    'close all sub forms
    For i = Forms.Count - 1 To 0 Step -1
        Unload Forms(i)
    Next
    End     'terminate program
End Sub

Private Sub HoistTimer_Timer()
    HoistTimer.Enabled = False
'Set Previous hoist to Blank so any subsequent hoist will action the sound signal
    PreviousHoist = ""
'Debug.Print "HoistTimer disabled"
End Sub


Private Sub mnuController_Click()
    With mnuController
        .Checked = Not .Checked
        frmDaventech.Visible = .Checked 'Loads
    End With
    If frmDaventech.Visible = False Then
        Unload frmDaventech
    End If
End Sub

Private Sub mnuEvents_Click()
    With mnuEvents
        .Checked = Not .Checked
        frmEvents.Visible = .Checked    'Loads
        If .Checked = True Then
            Call frmEvents.ListEvents
        End If
    End With
    
End Sub

Private Sub mnuLog_Click()
    With mnuLog
        .Checked = Not .Checked
        txtLog.Visible = .Checked
        txtLog.Enabled = .Checked
    End With

End Sub

Private Sub mnux10_Click()
    With mnux10
        .Checked = Not .Checked
    End With

End Sub

Private Sub mshFinish_SelChange()
'This is to remove the blue on the selected cell
'see http://vbcity.com/forums/t/30486.aspx
   With mshFinish
      .Col = .ColSel
      .Row = .RowSel
      .BackColorSel = .CellBackColor
      .ForeColorSel = .CellForeColor
   End With
'
End Sub

'Runs every 1 second
Private Sub RaceTimer_Timer()
Dim CurrTime As Date    'May be speeded up from Now() time for testing
Dim SecsSinceOutput As Long
Dim TimeToOutput As Date
Dim SecsToAdd As Long
Dim SecsSinceFirstStart As Long

    CurrTime = Now()
    lblCurrTime = Format$(CurrTime, "hh:mm:ss")
'WriteLog Timer  'elapsed secs since midnight
'Sets PreviousTimeOutput to to Current Time - 1 sec
    If PreviousTimeOutput = "00:00:00" Then
        Call ResetOutput(CurrTime)
    End If
    
    SecsSinceOutput = DateDiff("s", PreviousTimeOutput, CurrTime)
'No Output due yet
    If SecsSinceOutput = 0 Then
'Debug.Print "Skip"
        Exit Sub
    End If
                
'Goes round this loop once for each second between PreviousTimeOutput and CurrentTime
    Do
        
'Only used for testing (jnasetup)Backcolor will be cbDefault
        Select Case cmdPause.BackColor
            Case Is = vbGreen   'Commence Pause (next second)
                cmdPause.BackColor = vbCyan 'Paused
            Case Is = vbCyan    'Until pause is removed
                Exit Sub
        End Select
        
        TimeToOutput = DateAdd("s", 1, PreviousTimeOutput)
        If TimerOutput(TimeToOutput) = True Then PreviousTimeOutput = TimeToOutput
        SecsSinceOutput = DateDiff("s", PreviousTimeOutput, CurrTime)
        SecsSinceFirstStart = DateDiff("s", FirstStartTime, CurrTime)
'        EventTime = SecsSinceFirstStart - SecsSinceOutput
        state.NextEventTime = SecsSinceFirstStart - SecsSinceOutput
'Debug.Print ElapsedTime - SecsSinceOutput

        If StartTimeValid Then Call DoTimerEvents   '(EventTime)
'        lblElapsedTime = aSecToElapsed(ElapsedTime)
'display of ElapsedTime has catch-up secs taken off and PausedSecs
'        lblElapsedTime = aSecToElapsed(ElapsedTime - SecsSinceOutput) ' - PausedSecs)
If SecsSinceOutput > 0 Then Debug.Print "Catch-up " & SecsSinceOutput
    Loop Until SecsSinceOutput = 0  'Always execute once
End Sub


Private Sub ResetOutput(StartTime As Date)
        PreviousTimeOutput = DateAdd("s", -1, StartTime)
End Sub

Public Function aSecToElapsed(ByVal Secs As Long) As String
Dim hms As defhms
Dim Sign As Long
Dim aSign As String

'Secs = 3600& * 100&
    Sign = Sgn(Secs)    '-1 = -ve, 0 = 0 , +1 = +ve
    If Sign = -1 Then
        Secs = Secs * Sign 'force +ve
        aSign = "-"
    Else
        aSign = " "
    End If
    hms.Hour = Int(Secs / 3600&)
    Secs = Secs - hms.Hour * 3600&
    hms.Min = Int(Secs / 60&)
    Secs = Secs - hms.Min * 60&
    hms.Sec = Secs
    aSecToElapsed = aSign & Format$(hms.Hour, "###")
    If Abs(hms.Hour) >= 1 Then aSecToElapsed = aSecToElapsed & ":"
    aSecToElapsed = aSecToElapsed & Format$(hms.Min, "00") _
    & ":" & Format$(hms.Sec, "00")
End Function


Private Sub SignalTimer_Timer(Index As Integer)
Dim FlagIdx  As Long
Dim kb As String
Dim CyclesCompleted As Long
Dim LinkedFlagPos As Long
    
    With SignalAttributes(Index)
kb = SignalTimer(Index).Enabled
'Debug.Print Flags(FlagIdx).Visible
'A cycle is completed every time a flag is turned off AFTER it has been on
               
        If .Flag.pos Then
            If Flags(.Flag.pos).Visible = True Then
                .OnCycles = .OnCycles + 1
                SignalTimer(Index).Interval = .TTD
            Else
                SignalTimer(Index).Interval = .TTL
                CyclesCompleted = .OnCycles
                            
            End If
        Else
            .OnCycles = .OnCycles + 1
'Terminate Timer & Lower flag
            CyclesCompleted = .OnCycles
'MsgBox "Signal(" & Index & ")." & .Name & " has no associated Flag", vbCritical, "SignalTimer_Timer"
        End If
'Debug.Print CyclesCompleted & "(" & Index & ")"
        
'Continuous
        If .CyclesRequired = 0 Then
            If Loading = False Then CyclesCompleted = -1
        End If
        
        If Loading And CyclesCompleted > 5 Then CyclesCompleted = .CyclesRequired
        Select Case CyclesCompleted
'The timer has started but we do not want the Signal Off
'In fact we should not have started it in the first place
        Case Is >= .CyclesRequired
'This only occurs when the flag is about to be made invisible
'Turn off Signal, before disabling the timer
'Otherwise MakeSignals will start it again
'Click the command button (set to True) to put the flag down
'Only disable if not Continuous
            SignalTimer(Index).Enabled = False
'Must be after timer is disabled
            .OnCycles = 0
            Call LowerRequest(Index)
'        Commands(Index).Value = True
'Click the command button
'kb = SignalTimer(Index).Enabled    'Must be turned off
'Do this Previous, so if the timer is called again
'another off will be generated, and the timer will
'not re-start
'Remove this from the queue and re-enable with next signal (if any)
'        Call DequeTimer(Index)
        Case Is > .CyclesRequired
'Continuous
        Case Is < .CyclesRequired
'Reverse the Visibility of this flag and do another cycle
'No linked Flags are activated
            Call FlagVisibility(Index, Not Flags(.Flag.pos).Visible)

'Change the Visibility of any Linked flag UP Position only
'Because if it is the Horn that is linked we do not want to keep cycling it
''            If .Linkup(lidx).Flag > 0 Then
'Linked Flag must be raised as well (Pos > 0)
''                If SignalAttributes(.Linkup(lidx).Flag).Flag.Pos > 0 Then
''                    Call FlagVisibility(.Linkup(lidx).Flag, Flags(.Flag.Pos).Visible)
''                End If
''            End If
'Keep the timer running
        End Select
    End With
End Sub

Private Sub txtFirstStartTime_Change()
'Manually set
    If ValidateStartTime = True Then
'0=No Profile,1=Loading Profile,2=Profile Loaded
'3=Start Time Set,4=Start Sequence Running,5=Start Sequence Finished
        If state.Program = 2 Then
            Call SetProgramState(3)
        End If
        lblTimeToNextStart.Visible = True
        Call DisplayStartTimes
    End If
End Sub

Private Sub txtFirstStartTime_GotFocus()
Debug.Print "txtFirstStartTime"
End Sub

Private Function ValidateStartTime() As Boolean
Dim MyElapsedTime As Long
Dim Ret As Long
Dim kb As String

    On Error GoTo ValidateStartTime_error
    If txtFirstStartTime = "" Then
        txtFirstStartTime.BackColor = vbRed
    Else
        txtFirstStartTime.BackColor = cbEnabled
    End If
    
    If Len(txtFirstStartTime) <> 4 Then GoTo ValidateStartTime_error
    If IsNumeric(txtFirstStartTime) = False Then GoTo ValidateStartTime_error
    If CLng(txtFirstStartTime) < 1 Then GoTo ValidateStartTime_error
    If CLng(txtFirstStartTime) > 2400 Then GoTo ValidateStartTime_error
    
    If Len(txtFirstStartTime) = 4 _
        And CLng(NulToZero(txtFirstStartTime)) >= 0 _
        And CLng(NulToZero(txtFirstStartTime)) <= 2400 _
        And IsNumeric(NulToZero(txtFirstStartTime)) = True Then
            FirstStartTime = Date & " " _
            & Format$(NulToZero(txtFirstStartTime), "00:00") & ":00"
            On Error GoTo 0
            MyElapsedTime = DateDiff("s", FirstStartTime, Now())
            If IsEvtsInitialised(Evts) Then
                If MyElapsedTime > Evts(0).ElapsedTime Then
                    FirstStartTime = DateAdd("d", 1, FirstStartTime)
kb = "The Start Sequence will commence at "
kb = kb & DateAdd("s", Evts(0).ElapsedTime, FirstStartTime) & vbCrLf
kb = kb & aMins(-Evts(0).ElapsedTime) & " mins:secs, before the First Start" & vbCrLf
kb = kb & "This is after the current time " & Now & vbCrLf
kb = kb & "Do you wish the start to be tomorrow"
Ret = MsgBox(kb, vbOKCancel + vbDefaultButton2, "Validate Start Time")
                    If Ret <> vbOK Then
                        GoTo ValidateStartTime_error
                    End If
                MyElapsedTime = DateDiff("s", FirstStartTime, Now())
                End If
            End If
            txtFirstStartTime.ForeColor = vbBlack
'                StatusBar1.Panels(1).Text = ""
'Must not only reset the flags because once the start sequence
'has commenced the whole profile should be reloaded
'        Call ResetFlags
            ValidateStartTime = True
            StartTimeValid = True
            Exit Function
    End If
ValidateStartTime_error:
'    StatusBar1.Panels(1).Text = "Start time invalid"
    StartTimeValid = False  'Suppress Time Events
    txtFirstStartTime.ForeColor = vbRed
'dont allow the user to leave without a valid start time
    txtFirstStartTime.Enabled = True
    txtFirstStartTime.BackColor = cbEnabled
    txtFirstStartTime.SetFocus
End Function

Public Function DebugFlagsCheck()
Dim MyImage As Image
Dim Idx As Long
Dim Count As Long
Dim NoImageCount As Long

        For Idx = 1 To UBound(SignalAttributes)
            If SignalAttributes(Idx).Flag.pos > 0 Then
'RecallClass may not have been set when check is called
                If Not SignalAttributes(Idx).Image Is Nothing Then
                    
                    If SignalAttributes(Idx).Image <> Flags(SignalAttributes(Idx).Flag.pos).Picture Then

'                    Stop
                    End If
                End If
            End If
        Next Idx
    
Exit Function
    For Each MyImage In frmMain.Flags
        Count = 0
        For Idx = 1 To UBound(SignalAttributes)
            If SignalAttributes(Idx).Flag.pos = MyImage.Index Then
                Count = Count + 1
            End If
        Next Idx
        If MyImage.Picture.handle > 0 Then
            If Count <> 1 Then Stop
        Else
            NoImageCount = Count
        End If
    Next MyImage

End Function

Private Function ResetFlags()  'Not used for SYC
Dim MyImage As Image

    For Each MyImage In frmMain.Flags
        MyImage.Picture = Nothing
    Next
'Must set so that when loading profile the Queue Check for Recall does not fail
'With error
    RecallSignalIdx = 0
'    PreviousStart = 0
End Function

Private Function ResetCommands()  'Not used for SYC
Dim MyCommand As CommandButton
    Width = BaseWidth
    For Each MyCommand In Commands
        If MyCommand.Index <> 0 Then
            MyCommand.Enabled = True
            MyCommand.Visible = True
            MyCommand.BackColor = cbEnabled
        End If
    Next MyCommand
    NextCommandTop = 0
'    txtFirstStartTime = "0000"
    txtFirstStartTime = ""
    txtFirstStartTime.ForeColor = vbBlack
    txtFirstStartTime.Enabled = False
    txtFirstStartTime.BackColor = cbDisabled
    txtNoOfStarts = 1
    txtNoOfStarts.ForeColor = vbBlack
    txtNoOfStarts.Enabled = False
    txtNoOfStarts.BackColor = cbDisabled
    
End Function

Private Function ResetFinish()  'Not used for SYC
Dim Row As Long
Dim Col As Long

    FinishCount = 0
    FinishSignalCount = 0
    With mshFinish
'Clear rows (except 1)
        For Row = 2 To .Rows - .FixedRows
            .RemoveItem 1
        Next Row
'Clear Row 1
        For Col = 0 To .Cols - .FixedCols
            .TextMatrix(1, Col) = ""
        Next Col
    End With
End Function

Private Function ResetSignalTimers()  'Not used for SYC
Dim MySignalTimer As Timer
Dim i As Long
    For Each MySignalTimer In frmMain.SignalTimer
        If MySignalTimer.Index > 0 Then  'Dont delete SignalTimer(0)
            Unload MySignalTimer
        End If
    Next
    HoistTimer.Enabled = False
    PreviousHoist = ""
'Debug.Print "HoistTimer disabled"
End Function

'Sets the default settings when the Splash screen timer has finished and before
'a valid start time has been entered
Private Function DefaultsPreStartTimeSet()  'Not used for SYC
Dim Idx As Long
    RecallSignalIdx = 0  'Remove so that logic for recall is not applied when
                        'the flags are lowered
    For Idx = 1 To UBound(SignalAttributes)
        With SignalAttributes(Idx)
            If SignalAttributes(Idx).Flag.pos <> 0 Then
                Call LowerFlag(Idx)
            End If
        End With
    Next Idx
'Reset
    RecallSignalIdx = SignalFromName("Recall Class")
    Call DisplayStartTimes
    lblStartSequence.Caption = "Please Set a Start Time"
    lblStartSequence.Visible = True
    HoistTimer.Enabled = False
    PreviousHoist = ""  'Group
'    Call ResetRecall
    Call ResetCols

End Function

'Clears the form of all Programatic changes made as a result of the profile
'Called when frmMain is first loaded
'Called when a new proflie is loaded by LoadProfile
Public Function ClearProfile()
Dim MyCommand As CommandButton
Dim MyImage As Image
Dim MySignalTimer As Timer
Dim i As Long
Dim Row As Long
Dim Col As Long
    
WriteLog "ClearProfile"

'Start a fresh Profile
    If mnux10.Checked = True Then
        Multiplier = 10
    Else
        Multiplier = 1
    End If
    ClassSeparation = 0
    RaceTimer.Enabled = False
    StartTimeValid = False
    Caption = App.EXEName & " [" & App.Major & "." & App.Minor & "." _
        & App.Revision & "] " & CurrentProfile
'    StatusBar1.Panels(1).Text = ""
'    StatusBar1.Panels(2).Text = ""
    lblStartSequence.Visible = False
    lblStartTime.Visible = False    'Overlays lblStartSequence
    lblTimeToNextStart.Visible = False
    lblTimeFromPreviousStart.Visible = False
'remove functionkeys    Call DisplayKeys    'reset the status bar
    cmdPause.BackColor = cbEnabled  'Only used for testing (jnasetup)
'Clear Finish Display
    FinishCount = 0
    FinishSignalCount = 0
    With mshFinish
'Clear rows (except 1)
        For Row = 2 To .Rows - .FixedRows
            .RemoveItem 1
        Next Row
'Clear Row 1
        For Col = 0 To .Cols - .FixedCols
            .TextMatrix(1, Col) = ""
        Next Col
    End With
    

'Must set so that when loading profile the Queue Check for Recall does not fail
'With error
    RecallSignalIdx = 0
    RecallIdx = 0
    
'Must clear this before disabling as it causes a Validate, which sets the focus
'which will cause eror if the txtFirstStartTime is disabled
    txtFirstStartTime = ""
    txtFirstStartTime.Enabled = False
    txtFirstStartTime.BackColor = cbDisabled
    txtFirstStartTime.ForeColor = vbBlack
    txtNoOfStarts.ForeColor = vbBlack
    txtNoOfStarts.Enabled = False
    txtNoOfStarts.BackColor = cbDisabled
    state.NextEventTime = 0
    lblElapsedTime = aSecToElapsed(0)
    lblTimeToNextStart = aSecToElapsed(0)
    PostponeCountDown = 0
    If WindowState = vbNormal Then
        Width = BaseWidth
    End If
'Reset Command Buttons
    For Each MyCommand In Commands
        If StaticCommands(MyCommand.Index) = True Then
'            MyCommand.Enabled = False
            MyCommand.BackColor = cbdefault(MyCommand.Index)
        Else
'Make the base index invisible as it is not used
            Unload MyCommand
'            MyCommand.Enabled = False
'            MyCommand.Visible = False
        End If
    Next MyCommand
    NextCommandTop = 0

'Remove all Signal Timers
    For Each MySignalTimer In frmMain.SignalTimer
        If MySignalTimer.Index > 0 Then  'Dont delete SignalTimer(0)
            Unload MySignalTimer
        End If
    Next
    HoistTimer.Enabled = False
    PreviousHoist = ""
    
'Clear Flag images
    For Each MyImage In frmMain.Flags
        MyImage.Picture = Nothing
'MyImage.Stretch = True 'Fills image to size of container
    Next MyImage
WriteLog "Profile Cleared"

End Function

'Called from SetProgramState(3)
Public Function StartTimeIsSet() As Boolean
    
    lblStartTime.Visible = False
    With Commands(CommandFromCaption("Postpone"))
'*        .BackColor = vbGreen
        .Enabled = True
        .Enabled = False 'syc postpone not allowed after start time is set
        .BackColor = cbdefault(CommandFromCaption("Postpone")) 'syc postpone not allowed after start time is set
'*        .SetFocus
        End With
    With Commands(CommandFromCaption("Recall"))
        .Enabled = False
'.Enabled = True 'syc temp now set by DisplayRecalls
        .BackColor = cbdefault(CommandFromCaption("Recall"))
    End With
    With Commands(CommandFromCaption("General Recall"))
        .Enabled = False
        .BackColor = cbdefault(CommandFromCaption("GeneralRecall"))
    End With
    With Commands(CommandFromCaption("Finish"))
        .Enabled = False
'        .Visible = False
        .BackColor = cbdefault(CommandFromCaption("Finish"))
'StartTimeIsSet is called by SetProgramState(3)
'        USBButton.Off
    End With

End Function

'Called by DoTimerEvents immediately before first event is triggered
'And when Postpone is Raised
Public Function DefaultsFirstEvent()  'Not used for SYC
Dim Eidx As Long
'When the first event is carried out disable the start time
    
    txtFirstStartTime.Enabled = False
    txtFirstStartTime.BackColor = cbDisabled
'    Call StartTimeIsSet
    Call SetProgramState(3)
    
Exit Function
    frmMain.Commands(0).Visible = True
    frmMain.Commands(0).Enabled = True
    frmMain.Commands(0).BackColor = cbdefault(0)
    frmMain.Commands(0).SetFocus
    frmMain.Commands(0).Visible = False

End Function
'Requires MSINET.OCX
'See http://officeone.mvps.org/vba/http_file_exists.html
Public Function HTTPFileExists(ByVal url As String) As Boolean
    Dim S As String
    Dim Exists As Boolean
    On Error GoTo Inet1_Error
    With Inet1
        .RequestTimeout = 5
        .Protocol = icHTTP
        .url = url
        .Execute
'see http://support.microsoft.com/kb/182152 =True doesnt work
        Do While .StillExecuting <> False
            DoEvents
        Loop
        S = UCase(.GetHeader())
        Exists = (InStr(1, S, "200 OK") > 0)
        .Cancel 'close therequest
    End With
    HTTPFileExists = Exists
    Exit Function
Inet1_Error:
    Select Case Err.Number
    Case Is = icConnectFailed 'No internet connection
    End Select
    Inet1.Cancel
End Function

Public Function HttpSpawn(url As String)
Dim r As Long
Dim Command As String

If Environ("windir") <> "" Then
    r = ShellExecute(0, "open", url, 0, 0, 1)
Else
'try for linux compatibility
    Command = "winebrowser " & url & " ""%1"""

    Shell (Command)
End If
End Function

Public Function PositionCommand(Idx As Long)
'You dont need these unless testing this module in VBE
'If you have a break set frmMain is minimised and
'the Scale values will be 0
'Dont leave a blank gap
Dim BaseTop As Single   'Top of first Command

    BaseTop = 0 'fraMain.Top
    With Commands(Idx)
        .Caption = .Caption & "(" & Idx & ")"
        If .Visible = True Then
'This will be overwritten with the Name from SignalAttributes
'Align first command with top of main frame
            If .Width > ScaleWidth - fraMain.Width Then
                Width = Width + .Width
            End If
            WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
            .Top = ScaleTop + BaseTop + NextCommandTop
'            If .Top + .Height > BaseTop + fraMain.Height Then
            If .Top + .Height > StatusBar1.Top Then
                NextCommandTop = 0
                Width = Width + .Width
                WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
                .Top = ScaleTop + BaseTop + NextCommandTop
            End If
            WindowState = vbNormal  'Scale will be 0 in VBE (window is minimized)
            .Left = ScaleWidth - .Width
            NextCommandTop = NextCommandTop + .Height
        End If
    End With
End Function

'Called by the a Command to Raise a flag
'Must called by the Link (Sound may be clicked, with Sound still running)
'Queues if fixed position and Fixed position in use)
'Queues the the command if HoistTimer is running for this Group
'Queues Recall if ClassFlag is UP
'Actions Linked Flag by calling LinkRequest (If not Queued)
'Starts HoistTimer for this Group, if not Queueable (Flags.Queue=False)

Public Function RaiseRequest(ByVal Idx As Long)
Dim SoundEnabled As Boolean
Dim pos As Long
Dim QueueSignal As Long
Dim NextCmd As Long
Dim MyLink As defLink
Dim ClassIdx As Long
Dim MyImage As Image
Dim i As Long
Dim PostponeIdx As Long

'Dim PreparatoryIdx As Long
'If Idx = 3 Then Stop
    If Idx > UBound(SignalAttributes) Then
        Exit Function
    End If
    
'Check if Request requires Queueing or actioning
'If Fixed position and Position is in use
    With SignalAttributes(Idx)

        If Loading = False Then
            Select Case .Name
            Case Is = "Finish"
'A Finish Requires Completely different handling, Only Raise the Link UP event
'Do not RaiseRequest to put Finish Flag up (would toggle finish command actions)
                If .Name = "Finish" Then
'A Finish always clocks the time and must give correct no Linked signals
                    Call FinishTime
                    Exit Function
                End If

'Used After the start if the Recall is changed to General Recall
'I now don't think the user should be allowed to change this as it would
'Cause confusing signals (Recall followed by General Recall)
'It is required before the start to Deque the previous request if the user
'changes the type of recall
'            Case Is = "Recall", "General Recall"
'               Call LowerGroup(.Group)
'Dont exit, we need to action this signal
'Now I think if a recall is called the Other Recall should be disabled
'Providing it is not queued so it is moved to the end
            End Select
        End If
'Debug.Print "RaiseReq " & .Name
        pos = RC(.Flag.FixedRow, .Flag.FixedCol)
'If the Flag has a fixed position, check if any flag is already in this position
        If pos > 0 And .Flag.Queue = True Then
            If Flags(pos).Picture.handle <> 0 Then
                QueueSignal = Idx
'Debug.Print "Q Flag is UP"
            End If
        End If

'Queues the the command if HoistTimer is running for this Group
'So linked sound signal not made as another flag will be raised on same col
        If .group = PreviousHoist And .Flag.Queue Then
            QueueSignal = Idx
'Debug.Print "Q Timer On"
        End If
        
#If False Then
'Set up the Recall Class 2 flag hoist
        If .Name = "Recall" Then
'Stop
            i = SignalFromName("Recall Class")
            RecallIdx = Classes(PreviousClassStart(EventTime)).Signal
'            If RecallIdx > 0 Then   '0 if no start yet
'                Set SignalAttributes(i).Image = SignalAttributes(RecallIdx).Image
'            End If
            If state.StartClass.Previous > 0 Then
                Set SignalAttributes(i).Image = SignalAttributes(Classes(state.StartClass.Previous).Signal).Image
            End If
        End If
#End If
        If Loading = True Then
            QueueSignal = 0
        End If
        
        If QueueSignal > 0 Then
                Call QueueCmd(QueueSignal)
        Else

'Put the Flag up (if not up)
            If .Flag.pos = 0 Then
                
                Call RaiseFlag(Idx)

'Actions Linked Flag by calling LinkRequest (If not Queued)
                Call LinkRequest(Idx)
                
'Start HoistTimer for this Group, if not Queueable (Flags.Queue=False)
'So we dont Create a Second Sound signal
                If .Flag.Queue = False Then
                    HoistTimer.Enabled = False
                    HoistTimer.Enabled = True
                    PreviousHoist = .group
'Debug.Print "HoistTimer Enabled"
                End If
        
            End If  'Flag was not already up
        End If  'Not Queued
        
    End With

End Function

'Called by RaiseRequest and to action the UP link
Public Function RaiseFlag(ByVal Idx As Long)
Dim Col As Long
Dim Row As Long

'Load Profile-Linked Signals with a higher idx will not have been created
'Debug.Print "Raise " & SignalAttributes(Idx).Name
'Action Command now
'Display Image first (if there is one for this Signal)
    With SignalAttributes(Idx)
                
        If Not .Image Is Nothing Then
            Call NextFreeGroupFlagPos(Idx)
            If .Flag.pos = 0 And Loading = False Then
MsgBox "No free Flag positions", vbCritical, "RaiseFlag"
            End If
        End If

'If we have a flag position then create it (not set if no Image)
        If .Flag.pos > 0 Then
            Flags(.Flag.pos).Picture = .Image
'You have to set it to False becuase FlagVisibility only reacts to a change
            Flags(.Flag.pos).Visible = False
'Must use flagvisibility to create controller event
            Call FlagVisibility(Idx, True)
            .Flag.Changed = True

            Commands(Idx).BackColor = vbGreen
'May still be a timer even if no image to display
            If .TTL > 0 Then
                SignalTimer(Idx).Interval = .TTL
                SignalTimer(Idx).Enabled = True
            End If
        End If
    End With
        
    Call ResetCols  'Resets Cols().Group & .Items from SignalAttributes

End Function

'Called by command to Lower as Flag or when SignalTimer terminates
'Lowers This Flag
'Actions Linked Flag by calling LinkRequest
'Calls LowerFlag to Lower any subservient flags WITHOUT actioning any link
'(Only action the Link of the TOP flag)
'Dequeues Recall when Class Lowered by calling RaiseRequest
'Dequeues any Commands in the same Group Calling RaiseRequest
'Never called by the Link
'Never Queues the Command
Public Function LowerRequest(ByVal Idx As Long)
Dim NextCmd As Integer
Dim i As Long
Dim pos As Long     '> 0 if flag was up

'Load Profile-Linked Signals with a higher idx will not have been created
    If Idx > UBound(SignalAttributes) Then
        Exit Function
    End If

    With SignalAttributes(Idx)
'Keep whether flag was actually up. because this can ve called (by Evts Recall Down)
'even when it was not up
        pos = .Flag.pos
'Debug.Print "LowerReq " & SignalAttributes(Idx).Name
             
#If False Then
For i = 0 To UBound(CmdQ)
    If CmdQ(CInt(i)) <> 0 Then
'Debug.Print "Queued(" & i & ")=" & CmdQ(CInt(i))
    End If
Next i
#End If
'We cant ResetClassStart here beacause it is called even when the flag is not up
'Reset the ElapsedTime and remove any class flags
 
 'Lower Flag (if Up) then any below the Flag
        If SignalAttributes(Idx).Flag.pos > 0 Then

'This appears to be a ISAF one off. If Abandon is a 2 flag hoist then dont sound the LowerSignal
            If SignalAttributes(Idx).group = "Abandon" And Cols(SignalAttributes(Idx).Flag.Col).Items > 1 Then
                Call LowerFlag(Idx)
            Else
                Call LowerFlag(Idx)
                Call LinkRequest(Idx)
            End If
        End If
                                
'Dequeues any Commands in the same Group Calling RaiseRequest
        NextCmd = DequeCmd(.group)
        If NextCmd <> 0 Then
            Call RaiseRequest(NextCmd)
        End If
     
    End With
End Function

'Called by LowerRequest
'Lowers the Flag and any subservient flags
'Does not action any links
Private Function LowerFlag(ByVal Idx As Long)
Dim StartCol As Long
Dim StartRow As Long
Dim group As String
Dim i As Long
Dim Remove As Boolean
Dim PostponeIdx As Long
Dim PostponeClass As Long

'Debug.Print "LowerFlag " & SignalAttributes(Idx).Name
    With SignalAttributes(Idx)
        StartCol = .Flag.Col
        StartRow = .Flag.Row
        group = .group
    End With
'Calls LowerFlag to Lower any subservient flags WITHOUT actioning any link
'(Only the Link of the TOP flag is actioned by LowerRequest)
    For i = 1 To UBound(SignalAttributes)
'If i = 36 Then Stop
        With SignalAttributes(i)
            If .group = group Or (group = "Class" And .group = "Preparatory") Then
'If in different col or lower row in same col remove
                If .Flag.Col = StartCol And .Flag.Row >= StartRow Then

'Stop first. otherwise Timer will fail when it calls FlagVisibility
                    If SignalTimer(i).Enabled = True Then
                        SignalTimer(i).Enabled = False
                    End If
                        
'Clear the flag (if it exists)
                    If Flags(.Flag.pos).Picture.handle <> 0 Then
'If .Flag.Pos=0, FlagVisibility reports an error so must do first
                        Call FlagVisibility(i, False)
                        Flags(.Flag.pos).Picture = Nothing
                    End If
                    .Flag.pos = 0
                    .Flag.Col = 0
                    .Flag.Row = 0
                    Commands(i).BackColor = cbdefault(i)
'Must call the link to remove any Lights linked to non class flags (Postpone)
                    .Silent = True
                    Call LinkRequest(i)
                    .Silent = False
                End If
            End If
        End With
    Next i


'Stop Hoist Timer if Previous Flag up in this Group
    If group = PreviousHoist Then
        HoistTimer.Enabled = False
        PreviousHoist = ""
'Debug.Print "HoistTimer disabled"
    End If
    
    Call ResetCols  'Resets Cols().Group & .Items from SignalAttributes

End Function

'Calling Flag must be positioned (Up or Down) before LinkRequest is Called
'If HoistTimer for this Group is running (PreviousHoist = IdxGroup) dont action Link
'If Queueable (Flags.Queue=True) there should not be a link

Private Function LinkRequest(ByVal Idx As Long)
Dim Lidx As Long
Dim MyLink As defLink
Dim Suppress As Boolean

    With SignalAttributes(Idx)
        If IsLinksInitialised(.Links) Then
            For Lidx = 0 To UBound(.Links)
                MyLink = .Links(Lidx)
                If MyLink.Flag > 0 Then
                    Suppress = False
'If MyLink.Flag = 4 Then Stop
                    If .Flag.pos > 0 And MyLink.Type = "UpLink" Then
                        Call LinkExecute(Idx, MyLink)
                    End If
                    If .Flag.pos = 0 And MyLink.Type = "DownLink" Then
                        If SignalAttributes(MyLink.Flag).group = "Sound" _
                        And .Name = "Finish" And FinishCount > 1 _
                        And SoundOnAllFinishers = False Then
                            Suppress = True
                        End If
                        If Suppress = True Then
Debug.Print SignalAttributes(MyLink.Flag).Name & " suppressed"
                        Else
                            Call LinkExecute(Idx, MyLink)
                        End If
                    End If
                End If
'Stop 'Link execute can delete a links index which causes a subscript error
'Change for to a loop with mo0re checking
            Next Lidx
        End If
    End With
End Function

'IDx is the Signal containing the Link to Link
Private Function LinkExecute(Idx As Long, Link As defLink)
Dim LinkRejected As String
Dim Silent As Boolean
        
        With Link  'Raising Signal

'On ProfileLoad the linked flag may not have been created yet
'.Name is cleared when the Hoist Timer has finished its cycle (5 secs)
            If .Flag <> 0 And .Flag <= UBound(SignalAttributes) Then
Debug.Print "Link " & SignalAttributes(Idx).Name & " > " & SignalAttributes(.Flag).Name
'Only on raise because we have to action the downlink (White) when postpone is dropped
'within 10 secs
                If SignalAttributes(Idx).group = PreviousHoist And .Raise = True Then
                    LinkRejected = "Suppressed, PreviousHoist(" & PreviousHoist & ")"
                End If
                If SignalAttributes(Idx).Silent = True And SignalAttributes(.Flag).group = "Sound" Then
                    LinkRejected = "Silenced"
                End If
                If LinkRejected = "" Then
                    If .Raise = True Then   'Raise Linked flag
                        Call RaiseRequest(.Flag)
                    Else
                        Call LowerRequest(.Flag)   'Lower Linked flag
                    End If
                Else
Debug.Print LinkRejected
                End If
            Else
'There are no Linked Flags to this Flag
'Debug.Print "Link " & SignalAttributes(Idx).Name & " > none"
            End If
        End With
End Function

Private Function RC(ByVal Row As Long, ByVal Col As Long) As Long
'Both must be valid as a pair
    If Row > 0 And Col > 0 Then
        RC = (Row - 1) * 10 + Col
    End If
End Function

Private Function FlagRow(ByVal pos As Long) As Long
    If pos > 0 Then
        FlagRow = (pos - 1) \ 10 + 1
    End If
End Function
    
Private Function FlagCol(ByVal pos As Long) As Long
    If pos > 0 Then
        FlagCol = pos - (FlagRow(pos) - 1) * 10
    End If
End Function

'Called when Raising Flag, SignalAttributes Col & Row = 0 if no Position available
Private Function NextFreeGroupFlagPos(ByVal Idx As Long)
Dim Col As Long
Dim Row As Long
Dim pos As Long
Dim ClassIdx As Long

'If we do not have a set position see if this flag has a parent
'ie a 2 flag hoist and the parent flag is up
    
'    Call ResetCols
'If Idx = 9 Then Stop
   With SignalAttributes(Idx).Flag
'Get the Column first
        If .FixedCol > 0 Then
            .Col = .FixedCol
        End If

'See if this flag wants placing in same col as the first Class Flag
'DONT REMOVE may want to use it later
            If .Col = 0 Then
            Select Case SignalAttributes(Idx).group
            Case Is = "Preparatory", "Shortened"
'Not Recall as next Class flag may be up
                   ClassIdx = GroupIdx("Class")
                    If ClassIdx > 0 Then
'Put flag in same col
Debug.Print "Top Row"
                        .Col = SignalAttributes(ClassIdx).Flag.Col
                        .Row = Cols(.Col).Items + 1   '1st free row
'                        Call ShiftDown(.Row, .Col)
                    End If
            End Select
        End If
        
        If .Col = 0 Then
'See if we have a flag Raised in this Group with a spare Row available
            If Left$(SignalAttributes(Idx).Name, 6) <> "Class " Then
'Class Flags are always in separate cols (Keep in the same group)
                For Col = 1 To ColCountFree
                    If Cols(Col).group = SignalAttributes(Idx).group Then
                        If Cols(Col).Items < RowCount Then
                            .Col = Col
                            Exit For
                        End If
                    End If
                Next Col
            End If
        End If

'If no Col Group found, get First free col
        If .Col = 0 Then
            For Col = 1 To ColCountFree
                If Cols(Col).Items = 0 Then
'.Group is created by ResetCols
                    .Col = Col
                    Exit For
                End If
            Next Col
        End If
            
'If a Class flag see if we can place it in a free column but lower row
'Should only happen on initial load
        If .Col = 0 Then
            For Col = 1 To ColCountFree
                If Cols(Col).Items < RowCount Then  'This Col is full
                    If Cols(Col).group = SignalAttributes(Idx).group Then
                        For Row = Cols(Col).Items + 1 To RowCount
                            .Col = Col
                            .Row = Row
                            Exit For
                        Next Row
                    End If
                If .Col > 0 Then Exit For
                End If
            If .Col > 0 Then Exit For
            Next Col
        End If
            
'On initial load place in any free slot
        If .Col = 0 Then
            For Row = 1 To RowCount
                For Col = 1 To ColCount
                    If Cols(Col).Items < RowCount Then
                        .Col = Col
                        If Row < RowCount Then
                            .Row = Cols(.Col).Items + 1
                            Exit For
                        Else
MsgBox "No free Rows", vbCritical, "NextFreeGroupFlagPos"
                        End If
                    End If
                If .Col > 0 Then Exit For
                Next Col
            If .Col > 0 Then Exit For
            Next Row
        End If
                    
        If .Col = 0 Then
MsgBox "No free Cols", vbCritical, "NextFreeGroupFlagPos"
            Exit Function
        End If
        
        If .Row = 0 Then
            If .FixedRow > 0 Then
                .Row = .FixedRow
            Else
                If Row < RowCount Then
                    .Row = Cols(.Col).Items + 1
                Else
MsgBox "No free Rows", vbCritical, "NextFreeGroupFlagPos"
                End If
            End If
        End If
        
'Check position is actually free
        pos = RC(.Row, .Col)
        If Flags(pos).Picture = 0 Then
'Debug check (before .Pos is Set)
            Call DebugFlagsCheck
            .pos = pos
        Else
'This will happen when SplashScreen is loaded (multiple flags in fixed Positions)
            If Loading = False Then
MsgBox "Signal(" & Idx & ") " & SignalAttributes(Idx).Name & vbCrLf & "Flags(" & pos & ") not empty", vbCritical, "NextFreeGroupFlagPos"
            End If
            .Col = 0
            .Row = 0
        End If
    
    End With
'Debug.Print "NextPos=" & NextFreeGroupFlagPos & " (" & Row & "," & Col & ")"
End Function

Private Function DequeCmd(Optional group As String) As Integer
Dim i As Long
    For i = 0 To UBound(CmdQ)
        If CmdQ(i) <> 0 Then
            If group = "" Or SignalAttributes(CmdQ(i)).group = group Then
                If DequeCmd = 0 Then
                    DequeCmd = CmdQ(i)
Debug.Print "Deque " & SignalAttributes(CmdQ(i)).Name & " (" & group & ")"
                    Commands(CmdQ(i)).BackColor = cbdefault(i)
#If False Then
'When a recall and the command is cancelled, put focus & colour on the other recall command
'Must be done here because only deque is called not Lower Flag
                    Select Case SignalAttributes(CmdQ(i)).Name
                    Case Is = "Recall"
'                        Commands(CommandFromCaption("General Recall")).BackColor = vbGreen
'                        Commands(CommandFromCaption("General Recall")).SetFocus
                    Case Is = "General Recall"
'                        Commands(CommandFromCaption("Recall")).BackColor = vbGreen
'                        Commands(CommandFromCaption("Recall")).SetFocus
                    End Select
#End If
                    CmdQ(i) = 0
                End If
            End If
        End If
'Shift remaining commands up the queue
        If DequeCmd <> 0 Then
            If i = UBound(CmdQ) Then
                CmdQ(i) = 0
            Else
                CmdQ(i) = CmdQ(i + 1)
            End If
        End If
    Next i
    
'Stop
End Function

Private Function QueueCmd(Idx As Long)
Dim i As Long
    
    For i = 0 To UBound(CmdQ)
        If CmdQ(i) = 0 Then
            CmdQ(i) = Idx
            Commands(Idx).BackColor = vbCyan
'Debug.Print "Queue " & SignalAttributes(CmdQ(i)).Name
            Exit Function
        Else
'Only q the same command once (must not queue Recall more than once)
            If CmdQ(i) = Idx Then Exit Function
        End If
    Next i
'MsgBox "Command Queue is full (" & UBound(CmdQ) & ") maximum"
End Function

Public Function DisplayStartTimes()
Dim Csidx As Long
Dim FirstStartSecs As Long
Dim kb As String
Dim PreviousCsidx As Long

'Nothing to update
    If state.StartClass.Next = 0 Then
        Exit Function
    End If
'Total Secs at first start time
    FirstStartSecs = DateDiff("s", Date, FirstStartTime)
    If FirstStartSecs >= 86400 Then FirstStartSecs = FirstStartSecs - 86400
'    FirstStartSecs = FirstStartSecs - EventTime
    
#If jnasetup Then
    PreviousCsidx = UBound(Classes)
#Else
    PreviousCsidx = state.StartClass.Next
#End If
        
    With mshFinish
        For Csidx = 1 To UBound(Classes)
            If Csidx <= Csidx Then
                If Csidx > .Rows - .FixedRows Then
                    .Rows = Csidx + .FixedRows
                End If
                .TextMatrix(Csidx, 0) = "C" & Csidx
                .TextMatrix(Csidx, 1) = Trim$(aSecToElapsed(FirstStartSecs + Classes(Csidx).start + Classes(Csidx).Offset))
            Else
'Remove any class starts after next class start - happens with General recall
                If .Rows > 2 And .Rows - .FixedRows >= Csidx Then
                    .RemoveItem (.Rows - 1)
'Stop
                End If
            End If
        Next Csidx
    End With

End Function
Private Function StartTime_notused(ByVal Class As Long)
    With mshFinish
'not the first (blank) row
        If Class > .Rows - .FixedRows Then
            .Rows = Class + .FixedRows
        End If
        .TextMatrix(Class, 0) = "C" & Class
        .TextMatrix(Class, 1) = lblCurrTime.Caption
'Scroll to bottom
        .TopRow = .Rows - 1
    End With
End Function

'The finish time must be taken immediately
Private Function FinishTime()
Debug.Print "FinishCount=" & FinishCount
    With mshFinish
'not the first (blank) row
        FinishCount = FinishCount + 1
        If .TextMatrix(.Rows - .FixedRows, 0) <> "" Then
            .Rows = .Rows + .FixedRows
        End If
        .TextMatrix(.Rows - .FixedRows, 0) = FinishCount
        .TextMatrix(.Rows - .FixedRows, 1) = lblCurrTime.Caption
'Scroll to bottom
        .TopRow = .Rows - .FixedRows
    End With
'Check is there is a linked signal still visible
    Call FinishSignalRequest
End Function

'A Finish signal must be made for each finisher, so they must
'be queued, if the previous signal has not yet finished
Private Function FinishSignalRequest()
Dim Idx As Long
    
    Idx = SignalFromName("Finish")
    If LinkedSignalVisible(Idx) = False Then
'Make the linked signals immediately
        FinishSignalCount = FinishSignalCount + 1
        Call LinkRequest(Idx)
    End If
    
    If FinishSignalCount >= FinishCount Then
'all outstanding signals have been made
        FinishTimer.Enabled = False
    Else
'Make the signal later - try again in 1 sec
        FinishTimer.Enabled = True
    End If
    
End Function

'Check if there is a finish signal in progress
Private Function LinkedSignalVisible(ByVal Idx As Long) As Boolean
Dim Lidx As Long
Dim MyLink As defLink
    
    With SignalAttributes(Idx)
        If IsLinksInitialised(.Links) Then
            For Lidx = 0 To UBound(.Links)
                MyLink = .Links(Lidx)
                If MyLink.Flag > 0 Then
                    If SignalAttributes(MyLink.Flag).Flag.pos > 0 Then
                        LinkedSignalVisible = True
                        Exit Function
                    End If
                End If
            Next Lidx
        End If
    End With
End Function

'Keeps running until no outstanding finish signals to make
Private Sub FinishTimer_Timer()
    Call FinishSignalRequest
End Sub

Private Function FlagVisibility(ByVal Idx As Long, Visible As Boolean)
Dim pos As Long
Dim Cidx As Long
    pos = SignalAttributes(Idx).Flag.pos
'See if visiblility has changed (To generate Controller event)
    If pos > 0 Then
        If Flags(pos).Visible <> Visible Then
            Flags(pos).Visible = Visible
            Cidx = SignalAttributes(Idx).Controller
            If Cidx <> -1 Then
                With Controllers(Cidx)
                    If Visible Then
'Debug.Print .Connection & "(" & Cidx & ")" & .On
                        If .Sound <> "" Then Call PlayWav
                        If .On <> "" Then
                            Call frmDaventech.OpenAndSend(.On)
                            .state = True
                        End If
                    Else
'Debug.Print .Connection & "(" & Cidx & ")" & .Off
                        If .Sound <> "" Then Call PauseWav
                        If .Off <> "" Then
                            Call frmDaventech.OpenAndSend(.Off)
                            .state = False
                        End If
                    End If
                End With
            End If
        End If
    Else
        MsgBox "Flag " & SignalAttributes(Idx).Name & " not Raised", vbCritical, "FlagVisibility"
    End If
End Function

Private Function ResetCols()
Dim Idx As Long
Dim Col As Long
    ReDim Cols(ColCount)
    For Idx = 1 To UBound(SignalAttributes)
        With SignalAttributes(Idx)
            If .Flag.Col > 0 And .Flag.Row = 1 Then
                Cols(.Flag.Col).group = .group
            End If
            If .Flag.FixedCol > 0 Then
                Cols(.Flag.FixedCol).group = .group
            End If
            If SignalAttributes(Idx).Flag.Col > 0 Then
                Cols(.Flag.Col).Items = Cols(.Flag.Col).Items + 1
            End If
        End With
    Next Idx
End Function

'Used to Check if a Class Flag is up when Recall is asked for
'If 2 Class flags are up it will select the lowest class (Idx is in class order)
Private Function GroupIdx(ByVal group As String) As Long
Dim Idx As Long
    For Idx = 1 To UBound(SignalAttributes)
        With SignalAttributes(Idx)
            If .group = group And .Flag.pos > 0 Then
                 GroupIdx = Idx
                Exit For
            End If
        End With
    Next Idx
End Function

'Return the Command Button IDX, as we should find it within the first 6 buttons (Fixed)
'Command buttons may not have contiguous indexes (on start up in particular)
Public Function CommandFromCaption(ByVal CbName As String) As Integer
Dim MyCommand As CommandButton
WriteLog "CommandFromCaption " & CbName
    For Each MyCommand In frmMain.Commands
        If MyCommand.Caption = CbName Then
            CommandFromCaption = MyCommand.Index
            Exit Function
        End If
    Next MyCommand
End Function

'The EventTime has the PausedTime taken off
'Called once a second
Public Function DoTimerEvents() '(ByVal EventTime As Long)
Dim Eidx As Long
Dim Sidx As Long
Dim Bidx As Long
Dim Csidx As Long
Dim Pause As Boolean
Dim kb As String
Dim MyTime As Date
Dim TimeToFirstWarning As Long

'Timer is enabled to show the time while splash screen is displayed
    If ClearFlagsTimer.Enabled = True Then
        Exit Function
    End If

'If start time is set & within 1 min of start time, reset the start time
    If state.Program = 3 Then   'Postpone
'Up to first start time
        If state.NextEventTime <= Classes(1).start Then
'After 1 min before first warning
            If state.NextEventTime > Classes(1).Warning - 60 / Multiplier Then
                If SignalAttributes(SignalFromName("Postpone")).Flag.pos > 0 Then
                    MyTime = TimeSerial(Left(txtFirstStartTime, 2), Right(txtFirstStartTime, 2), 0)
                    MyTime = DateAdd("n", 1, MyTime)
                    txtFirstStartTime = Format$(MyTime, "hhmm")
                    Call ValidateStartTime
'                State.NextEventTime = State.NextEventTime + 60
                    Call DisplayStartTimes
                End If
            End If
        End If
    End If

'First Event - stop user changing start time
    If state.NextEventTime = Evts(0).ElapsedTime Then
'        Call SetState("Program", 4)
'        Call ClearKeyStates     'Must be reset to let the sequence alter the keys
                                'without causing an error, because postpone is initially
                                'set before the start sequence
        Call SetProgramState(4) 'Start Sequence Started
'These are only set on the First Start to Change All to the actual class being started
'Otherwise the F5 class is changed when Recalls are dropped
        Call frmMain.DisplayPreviousNextStartClass
'        Call SetKeyClass(StateToKeyIdx(State.Sequence), 1)
'remove functionkeys        Call frmMain.DisplayKeys    'Update the Status bar
     End If
            
'Previous Event
    If state.NextEventTime = Evts(UBound(Evts)).ElapsedTime Then
        Call SetProgramState(5) 'Start Sequence Finished
    End If
    
'To first start from FirstStartTime
    lblElapsedTime = aSecToElapsed(state.NextEventTime)

'We are postponing or recalling the function class
    If state.EventPause <> 0 Then
        Call AddOffset(state.StartClass.Next, 1)
    End If
    
'Call the events for each class in turn
    For Csidx = 0 To UBound(Classes)
        For Eidx = 0 To UBound(Evts)
            If Evts(Eidx).Class = Csidx Then
                If state.NextEventTime = Evts(Eidx).ElapsedTime + Classes(Evts(Eidx).Class).Offset Then
                    Call ProcessEvent(Eidx)
                End If
            End If
        Next Eidx
    Next Csidx
    

    If SyncEventsToNoOfStarts = True Then
MsgBox "Terminate" & vbCrLf & Classes(state.StartClass.Next).Name & " to " & Classes(state.StartClass.Last).Name
'   Call SynchroniseEvents
'Reset to next class start for testing
    NoOfStarts = state.StartClass.Next
    frmMain.txtNoOfStarts = NoOfStarts
    SyncEventsToNoOfStarts = False
    End If
    
    If state.EventPause <> 0 Then
        Call DisplayStartTimes
    End If
    If state.EventPause > 0 Then
        state.EventPause = state.EventPause - 1
    End If
    
    
'To next Class start after events
    lblTimeToNextStart = aSecToElapsed(state.NextEventTime - Classes(state.StartClass.Next).start - Classes(state.StartClass.Next).Offset)
    If state.Recalls <> 3 Then
        lblTimeFromPreviousStart = aSecToElapsed(state.NextEventTime - Classes(state.StartClass.Previous).start - Classes(state.StartClass.Previous).Offset)
    Else
        lblTimeFromPreviousStart = aSecToElapsed(state.NextEventTime - state.GrStartTime)
    End If
    
    state.NextEventTime = state.NextEventTime + 1

End Function

'This processes 1 event
Private Function ProcessEvent(ByVal Eidx As Long)
Dim Sidx As Long
Dim Bidx As Long
Dim Fcidx As Long
Dim Idx As Long

            If Left$(Evts(Eidx).Message, 1) <> "~" Then
'                StatusBar1.Panels(1).Text = Evts(Eidx).Message
            End If
            If IsSignalsInitialised(Evts(Eidx).Signals) Then
                For Sidx = 0 To UBound(Evts(Eidx).Signals)
                    With Evts(Eidx).Signals(Sidx)
'Silent on SignalAttributes is only used by LinkRquest and is set temporarily
'for this call only
                        If .Silent = "True" Then
                            SignalAttributes(.Signal).Silent = True
                        End If
'Dont make sound signal if another class flag is raised at the same
'time as a Class is started
                        If SignalAttributes(.Signal).group = "Class" And IsClassStartEvent = True Then
                            SignalAttributes(.Signal).Silent = True
                        End If
'Must be explictly asked for
                        If .Raise = "True" Then
                            If SignalAttributes(.Signal).Flag.pos = 0 Then
                                Call frmMain.RaiseRequest(.Signal)
                            End If
                        End If
'Must be explictly asked for
                        If .Raise = "False" Then
'All events associated with a class start must be actioned, even if flag not up
                            If SignalAttributes(.Signal).Flag.pos > 0 Then
                                Call frmMain.LowerRequest(.Signal)
'Reset (may have been set to silent temporarily for this event above)
                            End If
                        End If
                        SignalAttributes(.Signal).Silent = False
                    End With
                Next Sidx
            End If
            
            If IsFunctionCallsInitialised(Evts(Eidx).FunctionCalls) Then
                For Fcidx = 0 To UBound(Evts(Eidx).FunctionCalls)
                    With Evts(Eidx).FunctionCalls(Fcidx)
                        Select Case .Name
                        Case Is = "QueryRecallTimeout"  'Conditional (30 secs after start)
                            Call QueryRecallTimeout
                        Case Is = "RecallTimeout"   'Unconditional (4 Mins after start)
                                                    'NOT General Recall
                            Call RecallTimeout
                        Case Is = "GoRaisePostpone"
                        Case Is = "GoLowerPostpone"
                        Case Is = "LoadStart"   'Next Class Start
                            Idx = Classes(Evts(Eidx).Class).Signal
                            Call LowerRequest(Idx)
'                            Call SetProgramState(4)
'                            Call GoPreviousClassState(Evts(Eidx).Class)
                            Call GoNextClassStarted
'                            Call LoadStart(Evts(Eidx).Class) called by GoPreviousClassState
                        Case Else
                        End Select
                    End With
                Next Fcidx
            End If
            
            If IsButtonsInitialised(Evts(Eidx).Buttons) Then
                For Bidx = 0 To UBound(Evts(Eidx).Buttons)
                    With Evts(Eidx).Buttons(Bidx)
'Stop
'The Command button properties can be set immediately if any
                        If .Enabled <> "" Then
                                
                            If Commands(.Button).Enabled <> AtoBool(.Enabled) Then

'If the flag is up, you may not disable the command button because the user
'must be able to manually drop the flag
'This is important with recalls and probably postpone
                                If SignalAttributes(.Button).Flag.pos = 0 Or Commands(.Button).Enabled = False Then
                                    Commands(.Button).Enabled = AtoBool(.Enabled)
                                    Commands(.Button).BackColor = cbdefault(.Button)
                                End If
                 
'If we disable any command we must clear the colour
                                If Commands(.Button).Enabled = False Then
                                    Commands(.Button).BackColor = cbdefault(.Button)
                                End If
                            End If
                        End If
                    End With
                Next Bidx
            End If
'            If Evts(Eidx).Focus > 0 Then
'Must be enabled & visible to put focus on it
'                Commands(Evts(Eidx).Focus).Enabled = True
'                Commands(Evts(Eidx).Focus).Visible = True
'*                Commands(Evts(Eidx).Focus).SetFocus
'*                Commands(Evts(Eidx).Focus).BackColor = vbGreen
'            End If
'This is the Commands(0) button
'            If Evts(Eidx).Focus = 0 Then
'                Commands(Evts(Eidx).Focus).Enabled = True
'                Commands(Evts(Eidx).Focus).Visible = True
'*                Commands(Evts(Eidx).Focus).SetFocus
'                Commands(Evts(Eidx).Focus).Enabled = False
'                Commands(Evts(Eidx).Focus).Visible = False
'            End If
End Function

'Called when Postpone or GeneralRecall is raised
'Removes all raised ClassFlags (any any below the class flag)
Public Function ClassRestart(ByVal NextClassStart As Long)  'Not used for SYC
Dim Idx As Long
Dim TimeToNextWarning As Long

Debug.Print "ClassClear"
    For Idx = 1 To UBound(SignalAttributes)
        With SignalAttributes(Idx)
            If .group = "Class" And .Flag.pos > 0 Then
                    .Silent = True
                    Call LowerFlag(Idx)
                    Call LinkRequest(Idx)
                    .Silent = False
            End If
        End With
    Next Idx
'    TimeToNextWarning = State.NextEventTime - Classes(State.startclass.Next).Warning - Classes(State.startclass.Next).Offset - 1
    TimeToNextWarning = state.NextEventTime - Classes(NextClassStart).Warning - Classes(NextClassStart).Offset - 1
'    Call AddOffset(-TimeToNextStart)
    Call AddOffset(NextClassStart, 60 / Multiplier + TimeToNextWarning) '1 min delay before start sequence
'    Call SetState("ClassRestart", Class)
'    Call DisplayStartTimes
End Function

'Sets the ClassFlagImage on the RecallClassIdx
'Sets the RecallIdx
'Called when the Recall Buttons are enabled
'Also called when Recall is clicked before Start Sequence and there is no recallIdx
'to help testing)
Private Function RecallSetSignal()  'Not used for SYC
Dim i As Long
Dim Idx As Long
        
    i = SignalFromName("Recall Class")
    If i > 0 Then   'When Loading the Recall Class is not set
        RecallIdx = GetPostponeIdx  'The Class we will recall
'If there is no Class Flag up There will be no Postpone Idx so we
'cannot set a Class to recall
        If RecallIdx > 0 Then
            Set SignalAttributes(i).Image = SignalAttributes(RecallIdx).Image
    Commands(CommandFromCaption("Recall")).Enabled = True
    Commands(CommandFromCaption("General Recall")).Enabled = True
        End If
    End If
End Function

'Called when a Recall Flag is actually raised, which can be after the start
'Idx is the Recall Flag we are raising
Private Function RecallChange(ByVal Idx As Long)  'Not used for SYC
Dim OtherIdx As Long
Dim SaveRecallIdx As Long
Dim i As Long

'*    Call DequeCmd("Recall") 'May be queued and this request is to Deque it
    Select Case SignalAttributes(Idx).Name
    Case Is = "Recall"
        OtherIdx = SignalFromName("General Recall")
    Case Is = "General Recall"
        OtherIdx = SignalFromName("Recall")
    Case Else
'This is the RecallClass flag being raised
'Stop
        Exit Function
    End Select
    
'Check if other recall flag is up
    If SignalAttributes(OtherIdx).Flag.pos > 0 Then
'We need to save the class flag Idx because LowerFlag will remove it
        SaveRecallIdx = RecallIdx
        Call LowerFlag(OtherIdx)
        Call LinkRequest(OtherIdx)  'Drop the Linked signal (Fl White)
        RecallIdx = SaveRecallIdx
        i = SignalFromName("Recall Class")
        Set SignalAttributes(i).Image = SignalAttributes(RecallIdx).Image
        Commands(Idx).Enabled = True
        Commands(OtherIdx).Enabled = True
    End If
    
'Clear Green off Recall if General Recall is clicked before the start
'as well as when the Recall is changed
    Commands(OtherIdx).BackColor = cbdefault(OtherIdx)
    
End Function

'Returns the Lowest Class Flag currently UP. With SYC there can be 2 Class Flags
'UP at the same time, The Lowest will be the next one to start
'Returns 0 if no class flag up
Private Function GetPostponeIdx() As Long  'Not used for SYC
Dim Csidx As Long
Dim Idx As Long

    For Csidx = 1 To UBound(Classes)
        If SignalAttributes(Classes(Csidx).Signal).Flag.pos > 0 Then
            Idx = Classes(Csidx).Signal
            Exit For
        End If
    Next Csidx
    GetPostponeIdx = Idx
End Function

'Only called by ClassRestart
Private Function FirstClassEventTime(ByVal Class As Long) As Long  'Not used for SYC
Dim Eidx As Long
Dim Sidx As Long
Dim Silence As Boolean
Dim i As Long

'If no valid date has yet been entered the the next class must be 1
'The EventTime that will have been passed will be the no of secs since midnight
'If class 0 was returned the next DoEvents would think all classes had started
'    If txtFirstStartTime.Enabled = True Then
'Stop
'        Exit Function
'    End If
    For Eidx = 0 To UBound(Evts)
        If Evts(Eidx).Class = Class Then
            FirstClassEventTime = Evts(Eidx).ElapsedTime ' + Classes(Class).Offset
            Exit Function
        End If
    Next Eidx
End Function


'Only called by evts
'Unload Recalls if no Recall Flag has been raised (after 30 secs from start)
Private Function QueryRecallTimeout()
Dim Idx As Long
WriteLog "QueryRecallTimeout " & state.NextEventTime
'If Recall or General Recall not selected after 30 secs then Change to Postpone
'Otherwise do not action the Timeout
    If state.Recalls = 1 Then   'Both so neither selected (no recall)
        If state.StartClass.Previous < UBound(Classes) Then
            Call SetSequenceState(2) 'OK to Postpone next class,  Previous class has not started
        Else
            Call SetSequenceState(4)  'OK to Finish    Previous class has now started and cannot be recalled
        End If
    Else
            'Either Postpone or General Postpone is happening
    End If

End Function

'Unload Recall Flags, then reset state to Postpone or Finish
Private Function RecallTimeout()
Dim Idx As Long

WriteLog "RecallTimeout " & state.NextEventTime
'0=None,1=Both,2=Recall in progress ,3=General Recall in progress
'State.recalls=0 'If state is Finish
    If state.Recalls <> 0 Then
        If state.Recalls = 2 Then   'Recall in progress
            Call GoLowerRecall
        End If
        If state.Recalls = 3 Then   'Recall in progress
            Call GoLowerGeneralRecall
        End If
                
        If state.StartClass.Previous < UBound(Classes) Then
            Call SetSequenceState(2) 'OK to Postpone next class,  Last class has not started
        Else
            Call SetSequenceState(4)  'OK to Finish    Last class has now started
        End If
   End If

End Function

'Lowers the Class Flag
'Only Called by GoPreviousClassState each time a class is started
Public Function LoadStart_notused(ByVal Class As Long)
Dim Idx As Long

    If state.StartClass.Previous <> Class Then
'Stop
    End If
    Idx = Classes(Class).Signal
    Call LowerRequest(Idx)
End Function

'Adds an offset to each Class that will be restarted
'ie Postponed or General Recall
Private Function AddOffset(ByVal Class, Secs As Long)
Dim Csidx As Long
'Call the events for each class in turn
    For Csidx = Class To UBound(Classes)
        Classes(Csidx).Offset = Classes(Csidx).Offset + Secs
    Next Csidx
End Function


Public Function DisplayPreviousNextStartClass()
    If state.StartClass.Next > 0 Then
        lblNextStartName = Classes(state.StartClass.Next).Name
    Else
        lblNextStartName = "None"
'Turned on when starttimeisset
        lblTimeToNextStart.Visible = False
    End If
    If state.StartClass.Previous > 0 Then
        If state.Recalls <> 3 Then
            lblPreviousStartName = Classes(state.StartClass.Previous).Name
        Else
            lblPreviousStartName = Classes(state.StartClass.Previous).Name & " (recalled)"
        End If
        lblTimeFromPreviousStart.Visible = True
    Else
        lblPreviousStartName = "None"
    End If

End Function

'This actually enables/disables the text box
Public Function DisplayNoOfStarts()
    Select Case state.Program
    Case Is = 2, 3  'Profile Loaded but start sequence not started
        txtNoOfStarts.Enabled = True
        txtNoOfStarts.BackColor = cbEnabled
'Position cursor at RHS
        txtNoOfStarts.SelStart = Len(txtFirstStartTime)
        txtNoOfStarts.SetFocus
    Case Else
        txtNoOfStarts.Enabled = False
        txtNoOfStarts.BackColor = cbDisabled
    End Select
End Function

Public Function DisplayProgram()
    frmDaventech.lblProgram = aProgramState(state.Program)
End Function

'Update the State > Sequence Label
'Update the Finish Command Button Visibility
'Update the Postpone Command Button Visibility
Public Function DisplaySequence()
Dim FinishIdx As Long
Dim PostponeIdx As Long

WriteLog "DisplaySequence"

    frmDaventech.lblSequence = aSequenceState(state.Sequence)
    
'Command buttons may not have been defined
    FinishIdx = CommandFromCaption("Finish")
    If FinishIdx > 0 Then
        Select Case state.Sequence
        Case Is = 0, 4  '0=Disabled,1=NotStarted,2=Postpone,3=Recalls,4=Finish
            Commands(FinishIdx).Enabled = True
            USBButton.Yellow
        Case Else
            Commands(FinishIdx).Enabled = False
            USBButton.Off
        End Select
        Commands(FinishIdx).BackColor = cbdefault(FinishIdx)    'Sets back colour (Yellow if enabled)
    End If
    
    PostponeIdx = CommandFromCaption("Postpone")
    If PostponeIdx > 0 Then
        Select Case state.Sequence
        Case Is = 0, 1 '0=Disabled,1=NotStarted,2=Postpone,3=Recalls,4=Finish
            Commands(PostponeIdx).Enabled = True
        Case Is = 2
'syc Postpone after start sequence not used
'            Commands(PostponeIdx).Enabled = True
        Case Else
            Commands(PostponeIdx).Enabled = False
        End Select
        Commands(PostponeIdx).BackColor = cbdefault(PostponeIdx)
    End If
End Function

'Update the State > Recalls Label
'Update the Recall & General Recall Command Button Visibility
Public Function DisplayRecalls()
Dim RecallIdx As Long
Dim GeneralRecallIdx As Long

    frmDaventech.lblRecalls = aRecallsState(state.Recalls)
    
'Command buttons may not have been defined
    RecallIdx = CommandFromCaption("Recall")
    GeneralRecallIdx = CommandFromCaption("General Recall")
    If RecallIdx > 0 And GeneralRecallIdx > 0 Then
'Set Recall command button visibility
        Select Case state.Recalls
        Case Is = 0 'None
            Commands(RecallIdx).Enabled = False
            Commands(GeneralRecallIdx).Enabled = False
        Case Is = 1 'Both
            Commands(RecallIdx).Enabled = True
            Commands(GeneralRecallIdx).Enabled = True
        Case Is = 2 'Recall
            Commands(RecallIdx).Enabled = True
            Commands(GeneralRecallIdx).Enabled = False
        Case Is = 3 'General Recall
            Commands(RecallIdx).Enabled = False
            Commands(GeneralRecallIdx).Enabled = True
        End Select
        Commands(RecallIdx).BackColor = cbdefault(RecallIdx)
        Commands(GeneralRecallIdx).BackColor = cbdefault(GeneralRecallIdx)
    End If
End Function

'Update the Finish Command Button Visibility
'Public Function DisplayFinish()
'Stop
'End Function

Public Function DisplayKeys()  'Not used for SYC
Dim PanelNo As Long
Dim KeyNo As Long
    PanelNo = 1
    If IsKeysInitialised(Keys) = True Then
        For KeyNo = 1 To UBound(Keys)
'syc            If Keys(KeyNo).State > 0 Then
            If Keys(KeyNo).state > 1 Then
                With StatusBar1.Panels(PanelNo)
                    .Text = Keys(KeyNo).KeyName
                    If Keys(KeyNo).Cancel = True Then
                        .Text = .Text & " Cancel"
                    End If
                        .Text = .Text & " " & aKeyState(Keys(KeyNo).state)
                    Select Case Keys(KeyNo).state
                    Case Is = 1     'Postpone
                        .Text = .Text & " " & Classes(state.StartClass.Next).Name
                    Case Is = 2, 3   'Recall
                        .Text = .Text & " " & Classes(state.StartClass.Previous).Name
                    Case Is = 4         'Finish
'                        Call DisplayFinish  'enable the finish command button
                        .Text = .Text & " " & Classes(0).Name
                    End Select
                End With
                PanelNo = PanelNo + 1
            End If
        Next KeyNo
    End If
'Clear unused Panels
    For PanelNo = PanelNo To 2
        StatusBar1.Panels(PanelNo).Text = ""
    Next PanelNo
'    StatusBar1.Panels(3).Text = "F12 Horn Short"

End Function


Private Sub txtNoOfStarts_Change()
Stop
End Sub

Private Sub txtNoOfStarts_Validate(Cancel As Boolean)
Dim Idx As Long

Stop
    If IsNumeric(txtNoOfStarts.Text) Then
        If txtNoOfStarts.Text < 1 Then
            txtNoOfStarts.SetFocus
            Cancel = True
        Else
            NoOfStarts = txtNoOfStarts.Text
            If NoOfStarts > UBound(Classes) Then
                For Idx = UBound(Classes) To NoOfStarts
'            Call LoadClassEvents(Classes(Idx).ClassStartElapsedTime / Multiplier)
                Next Idx
            End If
        End If
    Else
    txtNoOfStarts.Text = CStr(NoOfStarts)
    txtNoOfStarts.SetFocus
  End If
End Sub

Private Sub UnloadTimer_Timer()
Dim reply
Dim i As Long
Dim kb As String

    If state.Program >= 3 Then  'Start Time set
        kb = "Do you wish to Terminate Racing Signals ?" & vbCrLf & vbCrLf
'        kb = kb & "If OK the program will be terminated and all times will be cleared" & vbCrLf
'        reply = MsgBox(kb, vbOKCancel + vbDefaultButton2, "Terminate")
        kb = kb & "If OK the program will be terminated" & vbCrLf
        ResultsFileName = Environ("appdata") & "\Arundale\RacingSignals\Results_" & Format$(FirstStartTime, "yyyymmdd_") & Format$(FirstStartTime, "hhnnss") & ".csv"
        kb = kb & "Start/Finish times will be saved in " & vbCrLf & Environ("appdata") & "\Arundale\RacingSignals\" & vbCrLf & Format$(FirstStartTime, "yyyymmdd_") & Format$(FirstStartTime, "hhnnss") & ".csv"
        reply = MyMsgBox(kb, vbOKCancel)
    
    Else
        reply = vbOK
    End If
        
    Select Case reply
    Case Is = vbCancel
        Cancel = True
    Case Is = vbOK
        WriteLog "Unloading " & Me.Name
        Set USBButton = Nothing
        Unload Me
        Unload frmDaventech
    Case Else
        MsgBox "Invalid", , "CboProfile_Click"
    End Select
    
    UnloadTimer.Enabled = False
    
    If Cancel = True Then
        If CurrentProfile <> "" Then
            For i = 0 To cboProfile.ListCount
            If cboProfile.List(i) = CurrentProfile Then
                cboProfile.ListIndex = i
'Caused Cbo Click Event with Cancel=true
                Exit For
            End If
        Next i
        Else
'No Current profile so dont select anything
            cboProfile.ListIndex = -1
        End If
    End If
End Sub
