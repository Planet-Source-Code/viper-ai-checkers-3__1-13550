VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Checkers V3.1"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12630
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":08CA
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   842
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   375
      Left            =   11880
      TabIndex        =   107
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<"
      Height          =   375
      Left            =   10320
      TabIndex        =   106
      Top             =   6840
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2280
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   113
      ImageHeight     =   140
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A00
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ADE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D2A0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEditRules 
      Caption         =   "&Edit Rules"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   104
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox ComPicBox 
      Height          =   2100
      Left            =   10560
      ScaleHeight     =   2040
      ScaleWidth      =   1635
      TabIndex        =   103
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "AI Options"
      Height          =   1575
      Left            =   10320
      TabIndex        =   98
      Top             =   1560
      Width           =   2175
      Begin VB.TextBox txtThinkTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   99
         ToolTipText     =   "Click to change"
         Top             =   480
         Width           =   1935
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   120
         TabIndex        =   102
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   3
         Max             =   5
         TickStyle       =   3
      End
      Begin VB.Label Labels 
         Alignment       =   2  'Center
         Caption         =   "Move Speed"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   101
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Labels 
         Alignment       =   2  'Center
         Caption         =   "Thinking Time"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   1
      Left            =   3120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdAdvanced 
      Caption         =   "&Advanced -->"
      Height          =   375
      Left            =   120
      TabIndex        =   95
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Start"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdName 
      Caption         =   "&Name Players"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10320
      TabIndex        =   29
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Load"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   6120
      Width           =   735
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   28
      Top             =   8130
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   529
      Style           =   1
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Frame Frame6 
      Caption         =   "Statistics"
      Height          =   855
      Left            =   10320
      TabIndex        =   20
      Top             =   3240
      Width           =   2175
      Begin VB.Label Labels 
         Caption         =   "Movelist:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Labels 
         Caption         =   "Cutoffs:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblMoves 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblCutoffs 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Game Info"
      Height          =   1335
      Left            =   10320
      TabIndex        =   16
      Top             =   120
      Width           =   2175
      Begin VB.Label Labels 
         Caption         =   "Total Turns:"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Labels 
         Caption         =   "P2 Time:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Labels 
         Caption         =   "P1 Time:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblTurns 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblP2Time 
         Alignment       =   2  'Center
         Caption         =   "0 Min 0 Sec"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblP1Time 
         Alignment       =   2  'Center
         Caption         =   "0 Min 0 Sec"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options"
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   1575
      Begin VB.CheckBox CheckMsgs 
         Caption         =   "Show Help"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox CheckCheat 
         Caption         =   "Cheat"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox CheckAutoSwitch 
         Caption         =   "Auto Switch"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Points"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
      Begin VB.Label lblP2Points 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblP1Points 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gameplay Type"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1575
      Begin VB.OptionButton Option1 
         Caption         =   "1 Player"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "2 Player"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Skip &Go"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdReverse 
      Caption         =   "&Reverse Board"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   7560
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2400
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   65
      ImageHeight     =   65
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1062C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13844
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CE8C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   0
      Left            =   2160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   2
      Left            =   4080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   3
      Left            =   5040
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   4
      Left            =   6000
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   5
      Left            =   6960
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   6
      Left            =   7920
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   7
      Left            =   8880
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   8
      Left            =   2160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   9
      Left            =   3120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   10
      Left            =   4080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   11
      Left            =   5040
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   12
      Left            =   6000
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   13
      Left            =   6960
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   14
      Left            =   7920
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   15
      Left            =   8880
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   16
      Left            =   2160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   17
      Left            =   3120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   18
      Left            =   4080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   19
      Left            =   5040
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   20
      Left            =   6000
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   21
      Left            =   6960
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   22
      Left            =   7920
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   23
      Left            =   8880
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   24
      Left            =   2160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   25
      Left            =   3120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   26
      Left            =   4080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   27
      Left            =   5040
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   28
      Left            =   6000
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   29
      Left            =   6960
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   30
      Left            =   7920
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   31
      Left            =   8880
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   32
      Left            =   2160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   33
      Left            =   3120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   34
      Left            =   4080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   35
      Left            =   5040
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   36
      Left            =   6000
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   37
      Left            =   6960
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   38
      Left            =   7920
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   39
      Left            =   8880
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   40
      Left            =   2160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   5040
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   41
      Left            =   3120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   5040
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   42
      Left            =   4080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   5040
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   43
      Left            =   5040
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   5040
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   44
      Left            =   6000
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   5040
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   45
      Left            =   6960
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   5040
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   46
      Left            =   7920
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   5040
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   47
      Left            =   8880
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   5040
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   48
      Left            =   2160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   49
      Left            =   3120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   50
      Left            =   4080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   51
      Left            =   5040
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   52
      Left            =   6000
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   53
      Left            =   6960
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   54
      Left            =   7920
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   55
      Left            =   8880
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   6000
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   56
      Left            =   2160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   6960
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   57
      Left            =   3120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   6960
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   58
      Left            =   4080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   6960
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   59
      Left            =   5040
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   6960
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   60
      Left            =   6000
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   6960
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   61
      Left            =   6960
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   6960
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   62
      Left            =   7920
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   6960
      Width           =   975
   End
   Begin VB.PictureBox Shape1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   63
      Left            =   8880
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label lblComPicNum 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   11040
      TabIndex        =   108
      Top             =   6900
      Width           =   735
   End
   Begin VB.Label lblComStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   105
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   672
      X2              =   672
      Y1              =   16
      Y2              =   528
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      Caption         =   "Player Turn:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblTurn 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   128
      X2              =   128
      Y1              =   16
      Y2              =   528
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------'
'    This Program and the code contained within is Copyright of Infostrategy Ltd. 2000
'    Any attempt to copy this program or any portion of it without first consulting the
'    Company will be met with deadly force and/or a lawsuit (particularly if copied with
'    Comercial gain in mind) -Viper
'-------------------------------------------------------------------------------------------'

Option Explicit
Dim StartTime As Long

Private Sub CheckAutoSwitch_Click()
  AutoSwitch = CheckAutoSwitch
  If AutoSwitch = 1 Then
    cmdReverse.Enabled = False
  ElseIf GameStarted Then
    cmdReverse.Enabled = True
  End If
  
  Select Case CurrentBoard.Turn
    Case 1
      Reversed = True: RefreshBoard CurrentBoard
    Case 2
      Reversed = False: RefreshBoard CurrentBoard
  End Select
  
  General.SaveSettings
End Sub

Private Sub CheckAutoSwitch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to automatically switch the board orientation to that of the current players' view (only in 2 player mode)"
End Sub

Private Sub CheckCheat_Click()
  CheatSwitch = CheckCheat
  MoveListChanged = True
  General.SaveSettings
End Sub

Private Sub CheckCheat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Used as debug, allow you to move any piece anywhere, moving a piece onto another deletes the original piece"
End Sub

Private Sub CheckMsgs_Click()
  ShowHelp = CheckMsgs
End Sub

Private Sub CheckMsgs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Toggles whether a message box is shown whenever the user generates an error as opposed to ignoring it"
End Sub

Private Sub cmdAdvanced_Click()
  If IsAdvanced Then
    IsAdvanced = False
    cmdAdvanced.Caption = "&Advanced -->"
    frmMain.Width = 10320
  Else
    IsAdvanced = True
    cmdAdvanced.Caption = "&Advanced <--"
    frmMain.Width = 12720
  End If
  General.SaveSettings
End Sub

Private Sub cmdAdvanced_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Toggles display of the advanced section"
End Sub

Private Sub cmdBack_Click()
Dim MovesBack As Long, Lng1 As Long

If PlayType = 1 Then MovesBack = 2 Else MovesBack = 1

If UBound(BoardHistory) - MovesBack < 1 Then Exit Sub

CurrentBoard = BoardHistory(UBound(BoardHistory) - MovesBack)
RefreshBoard CurrentBoard
VTurns = VTurns - MovesBack
RefreshDisplay
MoveListChanged = True

For Lng1 = 1 To MovesBack

  If UBound(BoardHistory) = 1 Then Exit For
  
  With BoardHistory(UBound(BoardHistory))
    Erase .Fields
    Erase .Pieces
    .Score = 0
    .MovesListFrom = 0
    .MovesListTo = 0
  End With
  
  ReDim Preserve BoardHistory(1 To UBound(BoardHistory) - 1)

Next

If UBound(BoardHistory) < 2 Then cmdBack.Enabled = False

End Sub

Private Sub cmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to revert back to board state before last move"
End Sub

Private Sub cmdDebug_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to toggle Automatic Debug Mode (the computer plays itself until an error occurs)"
End Sub

Private Sub cmdEditRules_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to edit the rules of the game (for both the AI and human player)"
End Sub

Private Sub cmdExit_Click()
  Unload Me
  End
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to exit game"
End Sub

Private Sub cmdName_Click()
Dim Response As String

Response = InputBox("Enter Player 1 Name", "Player 1", Names(1))
If Response <> "" Then Names(1) = Response Else Exit Sub

Response = InputBox("Enter Player 2 Name", "Player 2", Names(2))
If Response <> "" Then Names(2) = Response Else Exit Sub

SaveSettings
If GameStarted Then RefreshDisplay

End Sub

Private Sub cmdName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to name players"
End Sub

Private Sub cmdNext_Click()
  ComputerGraphic = ComputerGraphic + 1
  If ComputerGraphic = NumComImages + 1 Then ComputerGraphic = 1
  lblComPicNum = ComputerGraphic
  SaveSettings
End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to change computer player graphic to next in image list"
End Sub

Private Sub cmdOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to load a saved game"
End Sub

Private Sub cmdPrevious_Click()
  ComputerGraphic = ComputerGraphic - 1
  If ComputerGraphic = 0 Then ComputerGraphic = NumComImages
  lblComPicNum = ComputerGraphic
  SaveSettings
End Sub

Private Sub cmdPrevious_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to change computer player graphic to previous in image list"
End Sub

Private Sub cmdReset_Click()
  If GameStarted = True Then
    If MsgBox("Are you sure you want to quit current game?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
  End If
  ResetGame
  StartTime = Timer
  GameStarted = True
  cmdReverse.Enabled = True
  cmdSave.Enabled = True
  cmdSwitch.Enabled = True
End Sub

Private Sub cmdReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Starts a new game"
End Sub

Private Sub cmdReverse_Click()
  Reversed = Not Reversed
  RefreshBoard CurrentBoard
End Sub

Private Sub cmdReverse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Reverses board orientation"
End Sub

Private Sub cmdOpen_Click()
Dim File As String, Response As Long, Buffer As String, Gametype As Long
  If GameStarted Then
    Response = MsgBox("Erase current game?", vbExclamation + vbYesNo)
    If Response = vbNo Then Exit Sub
  End If
  
  CD1.DialogTitle = "Load Board"
  CD1.Filter = "Board File (*.brd)|*.brd"
  CD1.Flags = cdlOFNHideReadOnly
  CD1.InitDir = CurDir
  CD1.FileName = ""
  CD1.ShowOpen
  File = CD1.FileName
  If File = "" Then Exit Sub
  
  If Dir(File, vbHidden Or vbSystem Or vbNormal Or vbReadOnly) = "" Then
    MsgBox CD1.FileTitle & " does not exist", vbExclamation
    Exit Sub
  End If
  
  Open File For Binary Access Read As #1
    Get #1, , CurrentBoard
    'Get #1, , CurrentBoard.Turn
    Get #1, , Gametype
    Get #1, , VTurns
    Get #1, , VP1Time
    Get #1, , VP2Time
  Close #1
  
  GameStarted = True
  
  If Gametype Then Option1 = True Else Option2 = True
  
  StartTime = Timer
  GameStarted = True
  cmdReverse.Enabled = True
  cmdSave.Enabled = True
  cmdSwitch.Enabled = True
  
  MoveListChanged = True
  
  General.RefreshDisplay
  General.RefreshBoard CurrentBoard
  
End Sub

Private Sub cmdSave_Click()
Dim File As String, Response As Long, Buffer As String

  CD1.DialogTitle = "Save Board"
  CD1.Filter = "Board File (*.brd)|*.brd"
  CD1.Flags = cdlOFNHideReadOnly
  CD1.InitDir = CurDir
  CD1.FileName = ""
  CD1.ShowSave
  File = CD1.FileName
  If File = "" Then Exit Sub
  If Dir(File, vbHidden Or vbSystem Or vbNormal Or vbReadOnly) <> "" Then
    Response = MsgBox(CD1.FileTitle & " already exists, overwrite?", vbExclamation + vbYesNo)
    If Response = vbNo Then Exit Sub
  End If
  
  Open File For Binary As #1
    Put #1, , CurrentBoard
    'Put #1, , CurrentBoard.Turn
    Put #1, , CLng(Option1)
    Put #1, , VTurns
    Put #1, , VP1Time
    Put #1, , VP2Time
  Close #1
  
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to save the current game"
End Sub

Private Sub cmdSwitch_Click()
  Select Case CurrentBoard.Turn
    Case 1
      CurrentBoard.Turn = 2
      If Option1 = True Then Call AIMove
    Case 2
      CurrentBoard.Turn = 1
  End Select
  MoveListChanged = True
  RefreshDisplay
End Sub

Private Sub cmdSwitch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to skip the current players go"
End Sub

Private Sub ComPicBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ComPicBox.Picture = ImageList2.ListImages(ComputerGraphic).Picture
End Sub

Private Sub ComPicBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click (and hold) to display the current computer player graphic"
End Sub

Private Sub ComPicBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ComPicBox.Picture = Nothing
End Sub

Private Sub Form_Load()
Dim Lng1 As Long, Lng2 As Long

  General.GetSettings
  
  If Names(1) = "" Then
    Names(1) = "Player 1"
    Names(2) = "Nemesis"
    PlayType = 1
    IsAdvanced = 1
    TimeLimit = 5
    MoveSpeed = 300
    General.SaveSettings
  End If
  
  IndexMoves(1) = 11
  IndexMoves(2) = 9
  IndexMoves(3) = -9
  IndexMoves(4) = -11
  
  Reversed = True
  
  'For Lng1 = 0 To 63
  '  Shape1(Lng1).MouseIcon = LoadPicture("C:\windows\cursors\h_beam.cur")
  '  Shape1(Lng1).MousePointer = vbCrosshair
  'Next
  
  If ComputerGraphic = 0 Then ComputerGraphic = 1
  lblComPicNum = ComputerGraphic
  If PlayType = 1 Then
    Option1 = True
    CheckAutoSwitch.Enabled = False
  Else
    Option2 = True
    CheckAutoSwitch.Enabled = True
  End If
  CheckAutoSwitch = AutoSwitch
  CheckCheat = CheatSwitch
  txtThinkTime = TimeLimit & " Sec"
  If IsAdvanced Then
    cmdAdvanced.Caption = "&Advanced <--"
    frmMain.Width = 12720
  Else
    cmdAdvanced.Caption = "&Advanced -->"
    frmMain.Width = 10320
  End If
  
  Select Case MoveSpeed
    Case 50
      Slider1 = 5
    Case 100
      Slider1 = 4
    Case 200
      Slider1 = 3
    Case 300
      Slider1 = 2
    Case 400
      Slider1 = 1
    Case 500
      Slider1 = 0
  End Select
  
  'General.ResetGame
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = ""
End Sub

Private Sub lblMMatrixSize_Click()
  StatusBar1.SimpleText = "Displays the amount of moves the AI engine has generated so far"
End Sub

Private Sub lblComPicNum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays the number of the computer player graphic in use at present"
End Sub

Private Sub lblCutoffs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays the number of alpha-beta cutoffs the AI engine made last turn"
End Sub

Private Sub lblMoves_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays the number of moves the AI engine generated last turn"
End Sub

Private Sub lblP1Time_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays total time taken for " & Names(1) & " this game"
End Sub

Private Sub lblP2Time_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays total time taken for " & Names(2) & " this game"
End Sub

Private Sub lblPlyDepth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays current depth that the AI engine is thinking at"
End Sub

Private Sub lblTotalTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays total time taken for this game"
End Sub

Private Sub lblTurn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays whos turn it is"
End Sub

Private Sub lblTurns_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Displays the total number of turns this game"
End Sub

Private Sub Option1_Click()
  If CurrentBoard.Turn = 2 Then Call AIMove
  CheckAutoSwitch.Enabled = False
  PlayType = 1
  General.SaveSettings
End Sub

Private Sub Option2_Click()
  If GameStarted Then cmdReverse.Enabled = True
  CheckAutoSwitch.Enabled = True
  PlayType = 2
  General.SaveSettings
End Sub

Private Sub Shape1_Click(Index As Integer)
Static FromField As Long, ToField As Long, Captured As Long, MovesSoFar As String
Dim Direction As Long, Result As Long

If GameStarted = False Then Exit Sub

If CheatSwitch Then
  If FromField <> 0 Then
    ToField = IndexTranslation(Index, , True)
    If FromField = ToField Then
      Shape1(IndexTranslation(FromField, , False)).Picture = ImageList1.ListImages(PieceImageNum(CurrentBoard.Fields(FromField))).Picture
      FromField = 0
      ToField = 0
      Exit Sub
    End If
    If CurrentBoard.Fields(ToField) <> 0 Then CurrentBoard.Pieces(BothSides - ((CurrentBoard.Fields(ToField) And PlayerMask) / 16), CurrentBoard.Fields(ToField) And PieceNumMask) = 0
    CurrentBoard.Pieces(BothSides - ((CurrentBoard.Fields(FromField) And PlayerMask) / 16), CurrentBoard.Fields(FromField) And PieceNumMask) = ToField Or (DoubleMask And CurrentBoard.Fields(FromField))
    CurrentBoard.Fields(ToField) = CurrentBoard.Fields(FromField)
    CurrentBoard.Fields(FromField) = 0
    Shape1(Index).Picture = ImageList1.ListImages(PieceImageNum(CurrentBoard.Fields(ToField))).Picture
    Shape1(IndexTranslation(FromField, , False)).Picture = Nothing
    FromField = 0
    ToField = 0
    ReDim Preserve BoardHistory(1 To UBound(BoardHistory) + 1)
    BoardHistory(UBound(BoardHistory)) = CurrentBoard
    cmdBack.Enabled = True
  ElseIf CurrentBoard.Fields(IndexTranslation(Index, , True)) <> 0 Then
    FromField = IndexTranslation(Index, , True)
    Shape1(Index).Picture = ImageList1.ListImages(5).Picture
  End If
  Exit Sub
End If

If FromField <> 0 Then
  ToField = IndexTranslation(Index, , True)
  If FromField = ToField Then
    MovesSoFar = ""
    Shape1(IndexTranslation(FromField, , False)).Picture = ImageList1.ListImages(PieceImageNum(CurrentBoard.Fields(FromField))).Picture
    FromField = 0
    ToField = 0
    Exit Sub
  ElseIf CurrentBoard.Fields(FromField) And ((BothSides - CurrentBoard.Turn) * 16) Then
    If MovePiece(Chr$(IndexTranslation(Index, , True))) Then
      Shape1(IndexTranslation(FromField, , False)).Picture = ImageList1.ListImages(PieceImageNum(CurrentBoard.Fields(FromField))).Picture
      Shape1(Index).Picture = ImageList1.ListImages(5).Picture
      FromField = IndexTranslation(Index, , True)
    End If
  End If
  If (ToField - FromField) < 0 Then
    If (ToField - FromField) Mod 11 = 0 Then
      Direction = 4
    Else
      Direction = 3
    End If
  Else
    If (ToField - FromField) Mod 11 = 0 Then
      Direction = 1
    Else
      Direction = 2
    End If
  End If
  If Abs(ToField - FromField) > 11 And CurrentBoard.Fields(ToField - IndexMoves(Direction)) <> 0 Then
    Captured = ToField - IndexMoves(Direction)
  ElseIf CurrentBoard.Fields(ToField) <> 0 And CurrentBoard.Fields(ToField + IndexMoves(Direction)) = 0 Then
    Captured = ToField
    ToField = ToField + IndexMoves(Direction)
  End If
  Result = MovePiece(MovesSoFar & Chr(FromField) & Chr(ToField) & IIf(Captured, Chr(Captured), ""))
  If Result And MoveCorrect Then
    If Captured Then Shape1(IndexTranslation(Captured, , False)).Picture = Nothing
    Shape1(IndexTranslation(FromField, , False)).Picture = Nothing
    If Result And MoveCompleted Then
      Shape1(IndexTranslation(ToField, , False)).Picture = ImageList1.ListImages(PieceImageNum(CurrentBoard.Fields(ToField))).Picture
      MovesSoFar = ""
      FromField = 0
      If CurrentBoard.Turn = 1 Then
        VP2Time = VP2Time + (Timer - StartTime)
      Else
        VP1Time = VP1Time + (Timer - StartTime)
      End If
      VTurns = VTurns + 1
      MoveListChanged = True
      StartTime = Timer
      If AutoSwitch And PlayType = 2 Then
        If CurrentBoard.Turn = 1 Then Reversed = True Else Reversed = False
        Sleep 250
        RefreshBoard CurrentBoard
      End If
      RefreshDisplay
      DoEvents
      If CurrentBoard.Turn = 2 And frmMain.Option1 Then AIMove
    Else
      Shape1(IndexTranslation(ToField, , False)).Picture = ImageList1.ListImages(5).Picture
      MovesSoFar = MovesSoFar & Chr(FromField) & Chr(ToField) & Chr(Captured)
      FromField = ToField
    End If
    Captured = 0
    RefreshDisplay
  Else
    ToField = 0
    Captured = 0
  End If
Else
  If MovePiece(Chr$(IndexTranslation(Index, , True))) Then
    Shape1(Index).Picture = ImageList1.ListImages(5).Picture
    FromField = IndexTranslation(Index, , True)
  ElseIf ShowHelp Then
    MsgBox "This piece cannot be moved at present", vbExclamation
  End If
End If

End Sub

Private Sub Shape1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X1 As Long, Y1 As Long, Msg As String, BigIndex As Long
  BigIndex = IndexTranslation(CLng(Index), , True)
  Index = IndexTranslation(CLng(Index), True)
  XYConvert CLng(Index), X1, Y1
  
  If CurrentBoard.Fields(BigIndex) <> 0 Then
    If CurrentBoard.Fields(BigIndex) And P1Mask Then Msg = "     " & Names(1) Else Msg = "     " & Names(2)
    If CurrentBoard.Fields(BigIndex) And DoubleMask Then Msg = Msg & " Double Piece" Else Msg = Msg & " Single Piece"
    If MovePiece(Chr$(BigIndex)) = 0 Then Msg = Msg & "     (Cannot move)"
  End If
  
  StatusBar1.SimpleText = "Index = " & BigIndex & "     X = " & X1 & "     Y = " & Y1 & Msg
  
End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Adjusts the speed at which the computer makes its move (this is unrelated to thinking time)"
End Sub

Private Sub Slider1_Scroll()
  Select Case Slider1
      Case 0
        Slider1.Text = "Slowest"
        MoveSpeed = 500
      Case 1
        Slider1.Text = "Slow"
        MoveSpeed = 400
      Case 2
        Slider1.Text = "Normal"
        MoveSpeed = 300
      Case 3
        Slider1.Text = "Fast"
        MoveSpeed = 200
      Case 4
        Slider1.Text = "Fastest"
        MoveSpeed = 100
      Case 5
        Slider1.Text = "Insane"
        MoveSpeed = 50
    End Select
    General.SaveSettings
End Sub

Private Sub txtthinktime_Click()
If txtThinkTime.BorderStyle <> 1 Then
  txtThinkTime = Left(txtThinkTime, Len(txtThinkTime) - 4)
  txtThinkTime.Alignment = 0
  txtThinkTime.BackColor = &H8000000E
  txtThinkTime.BorderStyle = 1
  txtThinkTime.SelStart = 0
  txtThinkTime.SelLength = Len(txtThinkTime)
End If
End Sub

Private Sub txtthinktime_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cmdReset.SetFocus
End Sub

Private Sub txtthinktime_LostFocus()
  txtThinkTime.Alignment = 2
  txtThinkTime.BackColor = &H8000000F
  txtThinkTime.BorderStyle = 0
  If IsNumeric(txtThinkTime) Then
    TimeLimit = CLng(txtThinkTime)
    General.SaveSettings
  Else
    MsgBox "The maximum thought time must be numeric", vbExclamation
    txtThinkTime = TimeLimit
    txtThinkTime.SelStart = 0
    txtThinkTime.SelLength = Len(txtThinkTime)
  End If
  txtThinkTime = txtThinkTime & " Sec"
End Sub

Private Sub txtthinktime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusBar1.SimpleText = "Click to set time limit for the computer player"
End Sub
