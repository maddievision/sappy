VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sappy v1.6"
   ClientHeight    =   7260
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6975
   Icon            =   "frmsapp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameshit 
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   6135
      Left            =   7440
      TabIndex        =   275
      Top             =   240
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CheckBox cM 
         Caption         =   "Show note numbers"
         Height          =   255
         Left            =   120
         TabIndex        =   299
         Top             =   6000
         Width           =   1815
      End
      Begin VB.OptionButton auto1 
         Caption         =   "Song No. AutoADV"
         Height          =   255
         Left            =   0
         TabIndex        =   298
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton auto2 
         Caption         =   "List No. AutoADV"
         Height          =   255
         Left            =   1800
         TabIndex        =   297
         Top             =   600
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Supported ROMs"
         Height          =   255
         Left            =   360
         TabIndex        =   296
         Top             =   0
         Width           =   1575
      End
      Begin VB.CheckBox cAuto 
         Caption         =   "AutoNext"
         Height          =   255
         Left            =   120
         TabIndex        =   295
         Top             =   4320
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox ttable 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   294
         Text            =   "&H455C8C"
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox cRand 
         Caption         =   "Shuffle"
         Height          =   255
         Left            =   120
         TabIndex        =   293
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox tout 
         Enabled         =   0   'False
         Height          =   1215
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   292
         Top             =   4320
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "LIST"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   291
         Top             =   2040
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox tmidi 
         Enabled         =   0   'False
         Height          =   285
         Left            =   0
         TabIndex        =   290
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2640
         TabIndex        =   289
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   288
         Text            =   "c:\vba\lmf.txt"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Load ROM . . ."
         Height          =   255
         Left            =   2160
         TabIndex        =   287
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   360
         Top             =   840
      End
      Begin VB.TextBox tLoops 
         Height          =   285
         Left            =   120
         TabIndex        =   286
         Text            =   "2"
         Top             =   5400
         Width           =   495
      End
      Begin VB.TextBox tSt2 
         Height          =   285
         Left            =   240
         TabIndex        =   285
         Text            =   "&H15E"
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox tend 
         Height          =   285
         Left            =   240
         TabIndex        =   284
         Text            =   "&H1D4"
         Top             =   2280
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox cD 
         Caption         =   "Debug"
         Height          =   255
         Left            =   840
         TabIndex        =   283
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Play"
         Height          =   255
         Left            =   2760
         TabIndex        =   282
         Top             =   2760
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Export Song to .MID"
         Height          =   255
         Left            =   840
         TabIndex        =   281
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox cIMAP 
         Caption         =   "Ignore MIDI MAP"
         Height          =   255
         Left            =   120
         TabIndex        =   280
         Top             =   5760
         Width           =   1575
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Export InstBnk to .DLS"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   279
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CheckBox cGBMode 
         Caption         =   """GB Mode"""
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   278
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox cDirectMusic 
         Caption         =   "DirectMusic"
         Height          =   255
         Left            =   120
         TabIndex        =   277
         Top             =   4800
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox tdls 
         Enabled         =   0   'False
         Height          =   285
         Left            =   0
         TabIndex        =   276
         Top             =   1440
         Width           =   2055
      End
      Begin MSComDlg.CommonDialog cdl 
         Left            =   360
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Load GBA Rom"
         Filter          =   $"frmsapp.frx":12FA
      End
      Begin MSComDlg.CommonDialog cdl2 
         Left            =   360
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Export to MIDI File"
         Filter          =   "MIDI Files (*.mid)|*.mid|All Files (*.*)|*.*"
      End
      Begin VB.Label Label18 
         Caption         =   "Loop limit (0 = infinite)"
         Height          =   255
         Left            =   120
         TabIndex        =   300
         Top             =   5160
         Width           =   1575
      End
   End
   Begin VB.TextBox tFile 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "To open a ROM use File -> Load ROM . . . (Ctrl+L)"
      Top             =   240
      Width           =   6975
   End
   Begin VB.TextBox Label23 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   274
      Text            =   "No ROM Loaded"
      Top             =   0
      Width           =   6975
   End
   Begin VB.CommandButton Command12 
      Caption         =   ">"
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "<"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   28
      Top             =   6600
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   27
      Top             =   6360
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   26
      Top             =   6120
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   25
      Top             =   5880
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      Caption         =   ">"
      Height          =   195
      Left            =   2760
      TabIndex        =   12
      Top             =   2520
      Width           =   135
   End
   Begin VB.ComboBox listSongs 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   5535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "+"
      Height          =   195
      Left            =   2640
      TabIndex        =   11
      Top             =   2520
      Width           =   135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "-"
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   2520
      Width           =   135
   End
   Begin VB.CommandButton sD 
      Caption         =   "-"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   135
   End
   Begin VB.CommandButton sI 
      Caption         =   "+"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   1440
      Width           =   135
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   24
      Top             =   5640
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   23
      Top             =   5400
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   70
      Top             =   8280
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   5160
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   4920
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   4200
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox tstart 
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "<"
      Height          =   195
      Left            =   2400
      TabIndex        =   9
      Top             =   2520
      Width           =   135
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Play"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   495
   End
   Begin VB.ComboBox listGroups 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   273
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   15
      Left            =   6480
      TabIndex        =   272
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   14
      Left            =   6480
      TabIndex        =   271
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   13
      Left            =   6480
      TabIndex        =   270
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   12
      Left            =   6480
      TabIndex        =   269
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   11
      Left            =   6480
      TabIndex        =   268
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   10
      Left            =   6480
      TabIndex        =   267
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   9
      Left            =   6360
      TabIndex        =   266
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   8
      Left            =   6480
      TabIndex        =   265
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   7
      Left            =   6480
      TabIndex        =   264
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   6
      Left            =   6480
      TabIndex        =   263
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   5
      Left            =   6480
      TabIndex        =   262
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   261
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   260
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   259
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   258
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label LabelC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   257
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Chan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   22
      Left            =   6480
      TabIndex        =   256
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   15
      Left            =   3000
      TabIndex        =   255
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   15
      Left            =   5520
      TabIndex        =   254
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   15
      Left            =   5760
      TabIndex        =   253
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   15
      Left            =   6120
      TabIndex        =   252
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   15
      Left            =   3840
      TabIndex        =   251
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   15
      Left            =   4680
      TabIndex        =   250
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   15
      Left            =   2160
      TabIndex        =   249
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   15
      Left            =   2520
      TabIndex        =   248
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   15
      Left            =   1440
      TabIndex        =   247
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   15
      Left            =   720
      TabIndex        =   246
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   15
      Left            =   1800
      TabIndex        =   245
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   15
      Left            =   360
      TabIndex        =   244
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   243
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   13
      Left            =   360
      TabIndex        =   242
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   241
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   240
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   239
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   238
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   237
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   236
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   235
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   234
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   233
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   232
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   231
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   230
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   229
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Trk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Index           =   21
      Left            =   360
      TabIndex        =   228
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   14
      Left            =   1800
      TabIndex        =   227
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   14
      Left            =   720
      TabIndex        =   226
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   14
      Left            =   1440
      TabIndex        =   225
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   14
      Left            =   2520
      TabIndex        =   224
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   14
      Left            =   2160
      TabIndex        =   223
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   14
      Left            =   4680
      TabIndex        =   222
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   14
      Left            =   3840
      TabIndex        =   221
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   14
      Left            =   6120
      TabIndex        =   220
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   14
      Left            =   5760
      TabIndex        =   219
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   14
      Left            =   5520
      TabIndex        =   218
      Top             =   6360
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   14
      Left            =   3000
      TabIndex        =   217
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   13
      Left            =   1800
      TabIndex        =   216
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   13
      Left            =   720
      TabIndex        =   215
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   13
      Left            =   1440
      TabIndex        =   214
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   13
      Left            =   2520
      TabIndex        =   213
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   13
      Left            =   2160
      TabIndex        =   212
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   13
      Left            =   4680
      TabIndex        =   211
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   13
      Left            =   3840
      TabIndex        =   210
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   13
      Left            =   6120
      TabIndex        =   209
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   13
      Left            =   5760
      TabIndex        =   208
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   13
      Left            =   5520
      TabIndex        =   207
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   13
      Left            =   3000
      TabIndex        =   206
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   12
      Left            =   1800
      TabIndex        =   205
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   12
      Left            =   720
      TabIndex        =   204
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   12
      Left            =   1440
      TabIndex        =   203
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   12
      Left            =   2520
      TabIndex        =   202
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   12
      Left            =   2160
      TabIndex        =   201
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   12
      Left            =   4680
      TabIndex        =   200
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   12
      Left            =   3840
      TabIndex        =   199
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   12
      Left            =   6120
      TabIndex        =   198
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   12
      Left            =   5760
      TabIndex        =   197
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   12
      Left            =   5520
      TabIndex        =   196
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   12
      Left            =   3000
      TabIndex        =   195
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   194
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Frame"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   4560
      TabIndex        =   193
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   192
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   4560
      TabIndex        =   191
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   720
      TabIndex        =   190
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Priority"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   189
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   188
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   187
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   186
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   185
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   184
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   183
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   182
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   181
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   180
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   9
      Left            =   2880
      TabIndex        =   179
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   10
      Left            =   3000
      TabIndex        =   178
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Label5"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   177
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   176
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Mod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   6
      Left            =   5400
      TabIndex        =   175
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   174
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   173
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   172
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   171
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   170
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   169
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   6
      Left            =   5520
      TabIndex        =   168
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   7
      Left            =   5520
      TabIndex        =   167
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   8
      Left            =   5520
      TabIndex        =   166
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   9
      Left            =   5400
      TabIndex        =   165
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   10
      Left            =   5520
      TabIndex        =   164
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Caption         =   "Label9"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Index           =   11
      Left            =   5520
      TabIndex        =   163
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Pitc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   12
      Left            =   5760
      TabIndex        =   162
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   0
      Left            =   5760
      TabIndex        =   161
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   160
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   159
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   158
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   157
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   156
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   155
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   7
      Left            =   5760
      TabIndex        =   154
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   8
      Left            =   5760
      TabIndex        =   153
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   9
      Left            =   5640
      TabIndex        =   152
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   10
      Left            =   5760
      TabIndex        =   151
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label16 
      BackColor       =   &H00800080&
      Caption         =   "Label16"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   11
      Left            =   5760
      TabIndex        =   150
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Pan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   13
      Left            =   6120
      TabIndex        =   149
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   0
      Left            =   6120
      TabIndex        =   148
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   147
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   2
      Left            =   6120
      TabIndex        =   146
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   145
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   4
      Left            =   6120
      TabIndex        =   144
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   5
      Left            =   6120
      TabIndex        =   143
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   6
      Left            =   6120
      TabIndex        =   142
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   7
      Left            =   6120
      TabIndex        =   141
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   8
      Left            =   6120
      TabIndex        =   140
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   9
      Left            =   6000
      TabIndex        =   139
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   10
      Left            =   6120
      TabIndex        =   138
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00004080&
      Caption         =   "Label17"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   11
      Left            =   6120
      TabIndex        =   137
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Velocity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   16
      Left            =   3840
      TabIndex        =   136
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   135
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   134
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   133
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   132
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   131
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   130
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   6
      Left            =   3840
      TabIndex        =   129
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   7
      Left            =   3840
      TabIndex        =   128
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   8
      Left            =   3840
      TabIndex        =   127
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   9
      Left            =   3720
      TabIndex        =   126
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   10
      Left            =   3840
      TabIndex        =   125
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label20 
      BackColor       =   &H00800080&
      Caption         =   "Label20"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   11
      Left            =   3840
      TabIndex        =   124
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   17
      Left            =   4680
      TabIndex        =   123
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   122
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   121
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   120
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   119
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   118
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   5
      Left            =   4680
      TabIndex        =   117
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   116
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   115
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   8
      Left            =   4680
      TabIndex        =   114
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   9
      Left            =   4560
      TabIndex        =   113
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   112
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008080&
      Caption         =   "Label21"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   111
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   110
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   109
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   108
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   107
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   106
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   105
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   6
      Left            =   2160
      TabIndex        =   104
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   7
      Left            =   2160
      TabIndex        =   103
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   102
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   9
      Left            =   2040
      TabIndex        =   101
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   10
      Left            =   2160
      TabIndex        =   100
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   11
      Left            =   2160
      TabIndex        =   99
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   98
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   97
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   96
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   95
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   94
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   93
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   92
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   91
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   8
      Left            =   2520
      TabIndex        =   90
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   9
      Left            =   2400
      TabIndex        =   89
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   10
      Left            =   2520
      TabIndex        =   88
      Top             =   5400
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   11
      Left            =   2520
      TabIndex        =   87
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Voic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   86
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Vol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   85
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   1800
      TabIndex        =   84
      Top             =   6960
      Width           =   4455
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Special Event"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   240
      TabIndex        =   83
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   720
      TabIndex        =   82
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Inst"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   81
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2040
      TabIndex        =   80
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3720
      TabIndex        =   79
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Beat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   3120
      TabIndex        =   78
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   0
      TabIndex        =   77
      Top             =   1800
      Width           =   6975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3720
      TabIndex        =   76
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2040
      TabIndex        =   75
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Loop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   3120
      TabIndex        =   74
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Spd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   1560
      TabIndex        =   73
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Loc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   1560
      TabIndex        =   72
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Def"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   71
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   11
      Left            =   1440
      TabIndex        =   68
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   10
      Left            =   1440
      TabIndex        =   67
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   9
      Left            =   1320
      TabIndex        =   66
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   8
      Left            =   1440
      TabIndex        =   65
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   1440
      TabIndex        =   64
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   63
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   62
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   61
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   60
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   59
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   58
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   57
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Rel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   56
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Len"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   55
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lx 
      BackStyle       =   0  'Transparent
      Caption         =   "Loc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   54
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   720
      TabIndex        =   53
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   11
      Left            =   720
      TabIndex        =   52
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   10
      Left            =   720
      TabIndex        =   51
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   600
      TabIndex        =   50
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   8
      Left            =   720
      TabIndex        =   49
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   48
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   47
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   46
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   45
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   44
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   43
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   42
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   41
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   1800
      TabIndex        =   40
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   1800
      TabIndex        =   39
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   1680
      TabIndex        =   38
      Top             =   8280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   37
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   36
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   35
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   34
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   33
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   32
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   31
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   30
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   29
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   5895
      Left            =   0
      TabIndex        =   69
      Top             =   1800
      Width           =   7095
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mLoadROM 
         Caption         =   "&Load ROM . . ."
         Shortcut        =   ^L
      End
      Begin VB.Menu mExportSong 
         Caption         =   "&Export Song to .MID"
         Shortcut        =   ^E
      End
      Begin VB.Menu mRefreshRom 
         Caption         =   "&Refresh ROM List"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mS3 
         Caption         =   "-"
      End
      Begin VB.Menu mR 
         Caption         =   "&1"
         Index           =   0
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mR 
         Caption         =   "&2"
         Index           =   1
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mR 
         Caption         =   "&3"
         Index           =   2
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mR 
         Caption         =   "&4"
         Index           =   3
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mS 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mControl 
      Caption         =   "&Control"
      Begin VB.Menu mPlay 
         Caption         =   "&Play"
         Shortcut        =   ^A
      End
      Begin VB.Menu mStop 
         Caption         =   "&Stop"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mSP 
         Caption         =   "Song -1"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mSN 
         Caption         =   "Song +1"
         Shortcut        =   ^W
      End
      Begin VB.Menu mSLP 
         Caption         =   "Prev Song"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mSLN 
         Caption         =   "Next Song"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mOptions 
      Caption         =   "&Options"
      Begin VB.Menu mIgnoreMap 
         Caption         =   "&Ignore General MIDI Instrument Map"
         Shortcut        =   ^I
      End
      Begin VB.Menu mMIDINumbers 
         Caption         =   "Show &Notes in MIDI Note Numbers"
         Shortcut        =   ^N
      End
      Begin VB.Menu mS2 
         Caption         =   "-"
      End
      Begin VB.Menu mLoopLimit 
         Caption         =   "&Set Loop Limit . . ."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mAutoAdvance 
         Caption         =   "&Auto Advance to Next Track"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
      Begin VB.Menu mSongAutoAdv 
         Caption         =   "Use &Song Lists for Auto Advance"
         Checked         =   -1  'True
         Shortcut        =   {F4}
      End
      Begin VB.Menu mShuffle 
         Caption         =   "Shuffle &Tracks"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&Help"
      Begin VB.Menu mSupported 
         Caption         =   "Supported &ROMs . . ."
         Shortcut        =   {F9}
      End
      Begin VB.Menu mAbout 
         Caption         =   "&About . . ."
         Shortcut        =   {F11}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const apptitle = "Sappy v1.6d"

Private Type romsetting
 romheader As String * 4
 romname As String
 romtype As String
 songlist As String
 maplist As String
 tableoffs As Long
 songstart As Long
 songend As Long
End Type
Private roms(0 To &HFF) As romsetting
Private romcount As Long
Private songnames(0 To 65535) As String
Private songbanks(0 To 15) As String
Private songlistindex(0 To 65536) As Long
Private romheader As String * 4
Private gamename As String
Private romid As Long
Private recentfiles(0 To 3) As String
Private cIni As New CDS_CIni


Private Sub loadgrouplist(ByVal FileName As String)
Dim x As String
Dim y As Long
listGroups.Clear
For i = 0 To 15
songbanks(i) = ""
Next i

Open IIf(cD.value = vbChecked, debugpath, App.Path & "\") & "data\" & FileName For Input As #2
 
m = 0
Do
If EOF(2) Then
MsgBox "Unexpected End of File", vbCritical
Exit Sub
End If

Line Input #2, x

If x <> "ENDFILE" Then
songbanks(m) = x
listGroups.AddItem (m + 1) & ": " & songbanks(m)
j = 0

 Do
  If EOF(2) Then
  MsgBox "Unexpected End of File", vbCritical
  Exit Sub
  End If
 Line Input #2, x
 x = Trim(x)
 If x <> "END" Then
 j = j + 1
 End If
 Loop Until Trim(x) = "END"
m = m + 1
End If
Loop Until Trim(x) = "ENDFILE"
Close #2
If listGroups.ListCount > 0 Then listGroups.ListIndex = 0


End Sub
Private Sub loadsonglist(ByVal FileName As String, ByVal PlaylistGroupID)
Dim x As String
Dim y As Long
listSongs.Clear
For i = 0 To 65535
songnames(i) = ""
songlistindex(i) = -1
Next i
Open IIf(cD.value = vbChecked, debugpath, App.Path & "\") & "data\" & FileName For Input As #2


m = 0
Do
If EOF(2) Then
MsgBox "Unexpected End of File", vbCritical
Exit Sub
End If

Line Input #2, x

If x <> "ENDFILE" Then
songbanks(m) = x
j = 0

 Do
  If EOF(2) Then
  MsgBox "Unexpected End of File", vbCritical
  Exit Sub
  End If
 Line Input #2, x
 x = Trim(x)
 If x <> "END" Then
If m = PlaylistGroupID Then
 y = Val("&H" & Mid(x, 1, 4))
 x = Trim(Mid(x, 6))
 songnames(y) = x
 If songlistindex(y) < 0 Then songlistindex(y) = j
 listSongs.AddItem Hex2(y, 4) & ": " & x
End If
 j = j + 1
 End If
 Loop Until Trim(x) = "END"
m = m + 1
End If
Loop Until Trim(x) = "ENDFILE"
Close #2


If listSongs.ListCount > 0 Then listSongs.ListIndex = 0

End Sub
Private Sub setinstmap(ByVal FileName As String)
custommap = IIf(cD.value = vbChecked, debugpath, App.Path & "\") & "data\" & FileName
End Sub

Private Sub Check1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
For i = 0 To 15
If i = Index Then
 Check1(i).value = vbChecked
Else
 Check1(i).value = vbUnchecked
End If
Next i
Check2.value = vbUnchecked
End If
End Sub

Private Sub Check2_Click()
For i = 0 To 15
Check1(i).value = Check2.value
Next i
End Sub

Private Sub Check3_Click()

End Sub

Private Sub Command1_Click()
tout = ""
Dim ox As Long
Dim oy As Long
Dim a(0 To 3) As Byte
Open tFile For Binary As #1
For i = tstart To tend
ox = ttable + (i * 8)
Get #1, ox + 1, a(0)
Get #1, ox + 2, a(1)
Get #1, ox + 3, a(2)
Get #1, ox + 4, a(3)
oy = (CLng(a(2)) * CLng(&H10000)) + (CLng(a(1)) * CLng(&H100)) + a(0)
tout = tout & Hex(ox) & " " & Hex(i) & " [" & Hex(oy) & "] -" & vbCrLf
Next i
Close #1
End Sub

Private Sub Command10_Click()
x = "These ROMs are currently supported by Sappy:" & vbCrLf & vbCrLf
For i = 0 To romcount - 1
If roms(i).romtype = "sapphire" Then
x = x & roms(i).romheader & ":" & Chr(9) & roms(i).romname & " (GBA)" & vbCrLf
ElseIf roms(i).romtype = "pkmngbc" Then
x = x & roms(i).romheader & ":" & Chr(9) & roms(i).romname & " (GB)" & vbCrLf
End If
Next i
MsgBox x, vbInformation, "Supported ROMs"
End Sub

Private Sub Command11_Click()
fullstop = True
If listSongs.ListIndex > 0 And listSongs.ListCount > 0 Then
listSongs.ListIndex = listSongs.ListIndex - 1
If songplay = True Then Command2_Click
End If
End Sub

Private Sub Command12_Click()
fullstop = True
If listSongs.ListIndex < (listSongs.ListCount - 1) And listSongs.ListCount > 0 Then
listSongs.ListIndex = listSongs.ListIndex + 1
If songplay = True Then Command2_Click
End If
If cRand.value = vbChecked Then
Randomize Timer
With Form1.listSongs
If .ListCount > 1 Then
 .ListIndex = Fix(Rnd * (.ListCount))
End If
If .ListCount > 0 And songplay = True Then Command2_Click
End With
End If
End Sub

Private Sub Command13_Click()
fullstop = True
Command2_Click
End Sub

Private Sub Command14_Click()

datasize = 0
eventpass = 0
lastcmd = 0
If tFile = "" Or tFile = "To open a ROM use File -> Load ROM . . . (Ctrl+L)" Then
MsgBox "Please Load A ROM!", vbInformation, "No ROM Loaded"
Exit Sub
End If
cdl2.FileName = ""
cdl2.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames
cdl2.ShowSave
If cdl2.FileName <> "" Then
tmidi = cdl2.FileName
cIni.WriteValue "Sappy", "lastpathmidi", GetFilePath(cdl2.FileName)
Else
Exit Sub
End If
If Val(tLoops) < 1 Then
MsgBox "A loop limit MUST be set when exporting to .MID!", vbInformation, "No Loop Limit Set"
Exit Sub
End If
tstart = "&H" & Hex(Val(tstart))
listSongs.ListIndex = songlistindex(tstart)
Command10.Enabled = False
Command5.Enabled = False
Command3.Enabled = True
Form1.mStop.Enabled = True

'Do While songplay = True
'DoEvents
'xstop = True
'Loop
'Command2.Enabled = False
Label13 = "Exporting Song " & Hex(tstart) & vbCrLf & Replace(gamename, "&", "&&") & " - " & Replace(songnames(tstart), "&", "&&")
Open IIf(cD.value = vbChecked, debugpath, App.Path & "\") & "sappy.stt" For Output As #3
Print #3, App.Major & "." & App.Minor & App.Revision & "|" & tFile & "|" & roms(romid).romheader & "|" & Hex(Val(tstart)) & "|" & Trim(gamename) & "|" & Trim(songnames(tstart))
Close #3

For i = 0 To 14

 Label2(i) = ""
 Label8(i) = ""
 Label1(i) = ""
  Label16(i) = ""
   Label20(i) = ""
Label20(i).BackStyle = 0
Label21(i) = ""
Label21(i).BackStyle = 0
 Label17(i) = ""
 Label5(i) = ""
Label5(i).BackStyle = 0
 Label9(i) = ""
  LabelC(i) = ""

 Label9(i).BackStyle = 0
 Label6(i) = ""
 Label16(i).BackStyle = 0
 Label7(i) = ""
 Label17(i).BackStyle = 0
Next i
Label19 = ""
FileName = tFile
midiNewFile tmidi, midiSingleTrack, 48, 1
newMidiTrack Replace(gamename, "&", "&&") & " - " & Replace(songnames(tstart), "&", "&&")
looplimit = tLoops
tLoops.Enabled = False
fileout = True
Select Case roms(romid).romtype
 Case "sapphire": callback
 Case "pkmngbc": callback_pkmngbc
End Select
Exit Sub
haltall:
MidiClose
Err.Raise Err.Number
End
End Sub

Private Sub Command2_Click()
If tFile = "" Or tFile = "To open a ROM use File -> Load ROM . . . (Ctrl+L)" Then
MsgBox "Please Load A ROM!", vbInformation, "No ROM Loaded"
Exit Sub
End If

tstart = "&H" & Hex(Val(tstart))
listSongs.ListIndex = songlistindex(tstart)
Command10.Enabled = False
Command5.Enabled = False

If songplay = True Then
Timer1.Enabled = True
xstop = True
Exit Sub
End If
Timer1.Enabled = False

'On Error GoTo haltall
Command3.Enabled = True
Form1.mStop.Enabled = True

'Do While songplay = True
'DoEvents
'xstop = True
'Loop
'Command2.Enabled = False
Label13 = "Playing Song " & Hex(tstart) & vbCrLf & Replace(gamename, "&", "&&") & " - " & Replace(songnames(tstart), "&", "&&")
Open IIf(cD.value = vbChecked, debugpath, App.Path & "\") & "sappy.stt" For Output As #3
Print #3, App.Major & "." & App.Minor & App.Revision & "|" & tFile & "|" & roms(romid).romheader & "|" & Hex(Val(tstart)) & "|" & Trim(gamename) & "|" & Trim(songnames(tstart))
Close #3

For i = 0 To 14

 Label2(i) = ""
 Label8(i) = ""
 Label1(i) = ""
  Label16(i) = ""
   Label20(i) = ""
Label20(i).BackStyle = 0
Label21(i) = ""
Label21(i).BackStyle = 0
 Label17(i) = ""
 Label5(i) = ""
Label5(i).BackStyle = 0
 Label9(i) = ""
  LabelC(i) = ""

 Label9(i).BackStyle = 0
 Label6(i) = ""
 Label16(i).BackStyle = 0
 Label7(i) = ""
 Label17(i).BackStyle = 0
Next i
Label19 = ""
FileName = tFile
looplimit = tLoops
tLoops.Enabled = False
Select Case roms(romid).romtype
 Case "sapphire": callback
 Case "pkmngbc": callback_pkmngbc
End Select
Exit Sub
haltall:
MidiClose
Err.Raise Err.Number
End

End Sub

Private Sub Command3_Click()
'Do While songplay = True
'DoEvents
fullstop = True
xstop = True
'Loop
End Sub

Private Sub Command4_Click()
Open Text1 For Input As #1
For i = 0 To &H3F
Line Input #1, x
notelen(i) = Val("&H" & Mid(x, 4, 2))
Next i
Close #1
End Sub

Private Sub Command5_Click()
'On Error GoTo exitsub
cdl.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames
cdl.ShowOpen
If cdl.FileName <> "" Then


cdl.InitDir = GetFilePath(cdl.FileName)
cIni.WriteValue "Sappy", "lastpath", GetFilePath(cdl.FileName)

tFile = cdl.FileName
notarom = False
notarom = FileLen(tFile) < &HB0
If notarom = False Then
Open tFile For Binary As #92
Select Case getext(tFile)
 Case "gba": co = &HAD
 Case "bin": co = &HAD
 Case "gb": co = &H140
 Case "gbc": co = &H140
 Case Else: co = &HAD
End Select
Get #92, co, romheader
Close #92
romid = getromid(romheader)
If romid = 0 Then
MsgBox "This ROM [" & romheader & "] is unsupported by this version of Sappy.", vbExclamation, "ROM Unsupported [" & romheader & "]"
tFile = "To open a ROM use ""Load ROM . . ."" -->"
Else
If roms(romid).romtype = "sapphire" Or roms(romid).romtype = "pkmngbc" Then
loadromsettings romid
Me.Caption = apptitle & " - " & roms(romid).romname

zIndex = 4
For i = 0 To 3
If tFile = recentfiles(i) Then
 zIndex = i
 Exit For
End If
Next i
Dim bxp(0 To 2) As String


If zIndex < 4 Then

If zIndex > 0 Then

sf = recentfiles(zIndex)
m = 0
For ij = 0 To 3
If ij <> zIndex Then
bxp(m) = recentfiles(ij)
m = m + 1
End If
Next ij
recentfiles(0) = sf
For ij = 1 To 3
recentfiles(ij) = bxp(ij - 1)
Next ij
End If
Else

For ij = 0 To 2
bxp(ij) = recentfiles(ij)
Next ij
For i = 1 To 3
recentfiles(i) = bxp(i - 1)
Next i
recentfiles(0) = tFile
End If


For ij = 0 To 3
mR(ij).Caption = "&" & CStr(ij + 1) & " " & GetFilename(recentfiles(ij))
cIni.WriteValue "RecentFiles", CStr(ij), recentfiles(ij)
If Trim(recentfiles(ij)) <> "" Then
 mR(ij).Visible = True
Else
 mR(ij).Visible = False
End If
Next ij









ElseIf roms(romid).romtype = "error" Then
MsgBox roms(romid).songlist, vbInformation, roms(romid).romname & " [" & roms(romid).romheader & "]"
tFile = "To open a ROM use ""Load ROM . . ."" -->"
Else
MsgBox "The music engine used by this ROM [" & roms(romid).romheader & "] is unsupported by Sappy at this time."
tFile = "To open a ROM use ""Load ROM . . ."" -->"
End If
End If
Else
MsgBox "Could not read ROM header. This is not a valid ROM.", vbCritical, "Invalid ROM"
tFile = "To open a ROM use ""Load ROM . . ."" -->"
End If

Else
Exit Sub
End If
Exit Sub
exitsub:
MsgBox "Couldn't open dialog. Free up some memory!", vbCritical, "Insufficient Memory"
End Sub

Private Sub Command6_Click()
If speed > 1 Then speed = speed - 1
Select Case speed
 Case Is > 255: Label11.ForeColor = RGB(255, 0, 0)
 Case Else: Label11.ForeColor = RGB(192, 192, 192)
End Select
End Sub

Private Sub Command7_Click()
If speed < 511 Then speed = speed + 1
Select Case speed
 Case Is > 255: Label11.ForeColor = RGB(255, 0, 0)
 Case Else: Label11.ForeColor = RGB(192, 192, 192)
End Select

End Sub

Private Sub Command8_Click()
If speed > 1 Then speed = speed \ 2
Select Case speed
 Case Is > 255: Label11.ForeColor = RGB(255, 0, 0)
 Case Else: Label11.ForeColor = RGB(192, 192, 192)
End Select
End Sub

Private Sub Command9_Click()
If (speed * 2) < 511 Then
speed = speed * 2
Else
speed = 511
End If
Select Case speed
 Case Is > 255: Label11.ForeColor = RGB(255, 0, 0)
 Case Else: Label11.ForeColor = RGB(192, 192, 192)
End Select
End Sub

Private Sub cRand_Click()
If cRand.value = vbChecked Then
sD.Enabled = False
Command11.Enabled = False
Else
sD.Enabled = True
Command11.Enabled = True
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("<") Then If Command8.Enabled = True Then Command8_Click
If KeyAscii = Asc("-") Then If Command6.Enabled = True Then Command6_Click
If KeyAscii = Asc("+") Then If Command7.Enabled = True Then Command7_Click
If KeyAscii = Asc(">") Then If Command9.Enabled = True Then Command9_Click
If KeyAscii = Asc(",") Then If Command8.Enabled = True Then Command8_Click
If KeyAscii = Asc("_") Then If Command6.Enabled = True Then Command6_Click
If KeyAscii = Asc("=") Then If Command7.Enabled = True Then Command7_Click
If KeyAscii = Asc(".") Then If Command9.Enabled = True Then Command9_Click
End Sub

Private Sub Form_Load()
Me.Caption = apptitle
If App.EXEName = "prjsapp" Then cD.value = vbChecked
debugpath = "C:\sappy\"
cIni.FileName = IIf(cD.value = vbChecked, debugpath, App.Path) & "\sappy.ini"

Dim bz As String
Dim ba As Boolean
ba = cIni.ReadValue("Sappy", "lastpath", bz)
If ba = False Then
cdl.InitDir = IIf(cD.value = vbChecked, debugpath, App.Path)
cIni.WriteValue "Sappy", "lastpath", IIf(cD.value = vbChecked, debugpath, App.Path)
Else
cdl.InitDir = bz
End If

ba = cIni.ReadValue("Sappy", "lastpathmidi", bz)
If ba = False Then
cdl2.InitDir = IIf(cD.value = vbChecked, debugpath, App.Path)
cIni.WriteValue "Sappy", "lastpathmidi", IIf(cD.value = vbChecked, debugpath, App.Path)
Else
cdl2.InitDir = bz
End If

ba = cIni.ReadValue("Sappy", "LoopLimit", bz)
If ba = False Then
 cIni.WriteValue "Sappy", "LoopLimit", "2"
 tLoops = 2
Else
 tLoops = IIf(Val(bz) > 255 Or Val(bz) < 0, 2, Val(bz))
End If

ba = cIni.ReadValue("Sappy", "IgnoreMap", bz)
If ba = False Then
 cIni.WriteValue "Sappy", "IgnoreMap", CStr(False)
 mIgnoreMap.Checked = False
 cIMAP.value = vbUnchecked
Else
 mIgnoreMap.Checked = CBool(bz)
 cIMAP.value = IIf(bz = True, vbChecked, vbUnchecked)
End If


ba = cIni.ReadValue("Sappy", "MIDINumbers", bz)
If ba = False Then
 cIni.WriteValue "Sappy", "MIDINumbers", CStr(False)
 mMIDINumbers.Checked = False
 cM = vbUnchecked
Else
 mMIDINumbers.Checked = CBool(bz)
 cM.value = IIf(bz = True, vbChecked, vbUnchecked)
End If

ba = cIni.ReadValue("Sappy", "AutoAdvance", bz)
If ba = False Then
 cIni.WriteValue "Sappy", "AutoAdvance", CStr(True)
 mAutoAdvance.Checked = True
 cAuto = vbUnchecked
Else
 mAutoAdvance.Checked = CBool(bz)
 cAuto.value = IIf(bz = True, vbChecked, vbUnchecked)
End If

ba = cIni.ReadValue("Sappy", "SongAutoAdv", bz)
If ba = False Then
 cIni.WriteValue "Sappy", "SongAutoAdv", CStr(True)
 mSongAutoAdv.Checked = True
 auto2.value = True
Else
 mSongAutoAdv.Checked = CBool(bz)
 If bz = True Then
  auto2.value = True
 Else
  auto1.value = True
 End If
End If

ba = cIni.ReadValue("Sappy", "Shuffle", bz)
If ba = False Then
 cIni.WriteValue "Sappy", "Shuffle", CStr(False)
 mShuffle.Checked = False
 cRand = vbUnchecked
Else
 mShuffle.Checked = CBool(bz)
 cRand.value = IIf(bz = True, vbChecked, vbUnchecked)
End If

For i = 0 To 3
ba = cIni.ReadValue("RecentFiles", CStr(i), bz)

If ba = False Then
recentfiles(i) = ""
mR(i).Visible = False
mR(i).Caption = "&" & CStr(i + 1)
cIni.WriteValue "RecentFiles", CStr(i), ""
Else
recentfiles(i) = Trim(bz)
If Trim(bz) <> "" Then
mR(i).Enabled = True
mR(i).Caption = "&" & CStr(i + 1) & " " & GetFilename(bz)
End If
End If
Next i


'2815967
loadromlist
For i = 0 To 15
Label26(i) = i + IIf(i > 9, 0, 1)
 Label2(i) = ""
 Label8(i) = ""
 Label1(i) = ""
 LabelC(i) = ""
  Label16(i) = ""
   Label20(i) = ""
Label20(i).BackStyle = 0
Label21(i) = ""
Label21(i).BackStyle = 0
 Label17(i) = ""
 Label5(i) = ""
Label5(i).BackStyle = 0
 Label9(i) = ""
 Label9(i).BackStyle = 0
 Label6(i) = ""
 Label16(i).BackStyle = 0
 Label7(i) = ""
 Label17(i).BackStyle = 0
Next i
'Command4_Click

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
i = MsgBox("Are you sure you want to close Sappy?", vbYesNo + vbDefaultButton2, "Confirm Close")
If i = vbNo Then Cancel = 1
End Sub

Private Sub Form_Terminate()
fullstop = True
xstop = True
exiting = True

MidiClose
End Sub

Private Sub Form_Unload(Cancel As Integer)
fullstop = True
xstop = True
exiting = True
MidiClose
End Sub

Private Sub listGroups_Click()
If listGroups.ListCount = 0 Or listSongs.ListCount < 0 Then Exit Sub
loadsonglist roms(romid).songlist & ".lst", listGroups.ListIndex
End Sub

Private Sub listSongs_Click()
If listSongs.ListCount = 0 Or listSongs.ListIndex < 0 Then Exit Sub
tstart = "&H" & Hex(Val("&H" & Mid(listSongs.List(listSongs.ListIndex), 1, 4)))
End Sub



Private Sub mAbout_Click()
MsgBox "Sappy 1.6kls by DJ Bouche (Andrew Lim)" & vbCrLf & "GBA MIDI Music Engine to MIDI converter" & vbCrLf & "Released: July 23rd, 2003", vbInformation

End Sub

Private Sub mAutoAdvance_Click()
mAutoAdvance.Checked = Not mAutoAdvance.Checked
cAuto.value = IIf(mAutoAdvance.Checked = True, vbChecked, vbUnchecked)
cIni.WriteValue "Sappy", "AutoAdvance", CStr(mAutoAdvance.Checked)
End Sub

Private Sub mExit_Click()
Unload Form1
End Sub

Private Sub mExportSong_Click()
Command14_Click
End Sub

Private Sub mIgnoreMap_Click()
mIgnoreMap.Checked = Not mIgnoreMap.Checked
cIMAP.value = IIf(mIgnoreMap.Checked = True, vbChecked, vbUnchecked)
cIni.WriteValue "Sappy", "IgnoreMap", CStr(mIgnoreMap.Checked)
End Sub

Private Sub mLoadROM_Click()
Command5_Click
End Sub

Private Sub mLoopLimit_Click()
j = InputBox("Enter Loop Limit (1 to 256) (Type in ""0"" for infinite)", "Set Loop Limit", tLoops)
If j = "" Then Exit Sub
j = Val(j)
If j > 256 Then
MsgBox "Too many loops, setting limit to default of 2.", vbCritical
j = 2
End If
tLoops = j
cIni.WriteValue "Sappy", "LoopLimit", CStr(tLoops)
End Sub

Private Sub mMIDINumbers_Click()
mMIDINumbers.Checked = Not mMIDINumbers.Checked
cM.value = IIf(mMIDINumbers.Checked = True, vbChecked, vbUnchecked)
cIni.WriteValue "Sappy", "MIDINumbers", CStr(mIgnoreMap.Checked)
End Sub

Private Sub mPlay_Click()
Command13_Click
End Sub

Private Sub mR_Click(Index As Integer)
cdl.FileName = recentfiles(Index)
If cdl.FileName <> "" Then


cdl.InitDir = GetFilePath(cdl.FileName)
cIni.WriteValue "Sappy", "lastpath", GetFilePath(cdl.FileName)

tFile = cdl.FileName
notarom = False
notarom = FileLen(tFile) < &HB0
If notarom = False Then
Open tFile For Binary As #92
Select Case getext(tFile)
 Case "gba": co = &HAD
 Case "bin": co = &HAD
 Case "gb": co = &H140
 Case "gbc": co = &H140
 Case Else: co = &HAD
End Select
Get #92, co, romheader
Close #92
romid = getromid(romheader)
If romid = 0 Then
MsgBox "This ROM [" & romheader & "] is unsupported by this version of Sappy.", vbExclamation, "ROM Unsupported [" & romheader & "]"
tFile = "To open a ROM use ""Load ROM . . ."" -->"
Else
If roms(romid).romtype = "sapphire" Or roms(romid).romtype = "pkmngbc" Then
loadromsettings romid
Me.Caption = apptitle & " - " & roms(romid).romname

Dim bxp(0 To 2) As String
If Index > 0 Then

sf = recentfiles(Index)
m = 0
For i = 0 To 3
If i <> Index Then
bxp(m) = recentfiles(i)
m = m + 1
End If
Next i
recentfiles(0) = sf
For i = 1 To 3
recentfiles(i) = bxp(i - 1)
Next i
End If
For i = 0 To 3
mR(i).Caption = "&" & CStr(i + 1) & " " & GetFilename(recentfiles(i))
cIni.WriteValue "RecentFiles", CStr(i), recentfiles(i)
If Trim(recentfiles(i)) <> "" Then
 mR(i).Visible = True
Else
 mR(i).Visible = False
End If
Next i



ElseIf roms(romid).romtype = "error" Then
MsgBox roms(romid).songlist, vbInformation, roms(romid).romname & " [" & roms(romid).romheader & "]"
tFile = "To open a ROM use ""Load ROM . . ."" -->"
Else
MsgBox "The music engine used by this ROM [" & roms(romid).romheader & "] is unsupported by Sappy at this time."
tFile = "To open a ROM use ""Load ROM . . ."" -->"
End If
End If
Else
MsgBox "Could not read ROM header. This is not a valid ROM.", vbCritical, "Invalid ROM"
tFile = "To open a ROM use ""Load ROM . . ."" -->"
End If

Else
Exit Sub
End If
Exit Sub
exitsub:

End Sub

Private Sub mRefreshRom_Click()
loadromlist
MsgBox "Reloaded 'sappy.lst'.", vbInformation

End Sub

Private Sub mShuffle_Click()
mShuffle.Checked = Not mShuffle.Checked
cRand.value = IIf(mShuffle.Checked = True, vbChecked, vbUnchecked)
cIni.WriteValue "Sappy", "Shuffle", CStr(mShuffle.Checked)
End Sub

Private Sub mSLN_Click()
Command12_Click
End Sub

Private Sub mSLP_Click()
Command11_Click
End Sub

Private Sub mSN_Click()
sI_Click
End Sub

Private Sub mSongAutoAdv_Click()
mSongAutoAdv.Checked = Not mSongAutoAdv.Checked
If mSongAutoAdv.Checked = True Then
 auto2.value = True
Else
 auto1.value = True
End If
cIni.WriteValue "Sappy", "SongAutoAdv", CStr(mSongAutoAdv.Checked)
End Sub

Private Sub mSP_Click()
sD_Click
End Sub

Private Sub mStop_Click()
Command3_Click
End Sub

Private Sub mSupported_Click()
Command10_Click
End Sub

Private Sub sD_Click()
fullstop = True
tstart = "&H" & Hex(Val(tstart))
If tstart > 0 Then
tstart = "&H" & Hex(tstart - 1)
If songplay = True Then Command2_Click
End If
End Sub

Public Sub sI_Click()
fullstop = True
tstart = "&H" & Hex(Val(tstart))
If tstart < &HFFF Then
tstart = "&H" & Hex(tstart + 1)
If songplay = True Then Command2_Click
End If
If cRand.value = vbChecked Then
tstart = "&H" & Hex(Fix(Rnd * &HFFF))
If songplay = True Then Command2_Click
End If
End Sub

Private Sub Timer1_Timer()
Command2_Click
End Sub
Private Function Hex2(ByVal value As Long, Optional ByVal length As Byte = 2)
x = Hex(value)
Do While Len(x) < length
x = "0" & x
Loop
Hex2 = x
End Function

Private Sub loadromsettings(ByVal lromid As Long)
With roms(lromid)
clearinstmap
ttable = "&H" & Hex(.tableoffs)
tstart = "&H" & Hex(.songstart)
tSt2 = "&H" & Hex(.songstart)
tend = "&H" & Hex(.songend)
ttable.Enabled = False
loadgrouplist .songlist & ".lst"
loadsonglist .songlist & ".lst", 0
setinstmap .maplist & ".map"
gamename = .romname
Label13 = .romname
Label23 = .romname & " (" & .romheader & ")"
End With
End Sub

Private Function getromid(ByVal romheader As String) As Long
match = 0
For i = 1 To romcount
testhead = roms(i).romheader
If romheader = testhead Or (Right(testhead, 1) = Chr(255) And Left(romheader, 3) = Left(testhead, 3)) Then
match = i
Exit For
End If
Next i
getromid = match
End Function
Private Sub loadromlist()
Open IIf(cD.value = vbChecked, debugpath, App.Path & "\") & "data\sappy.lst" For Input As #91
i = 1
Do
Input #91, rhead
If Trim(UCase(rhead)) = "ENDFILE" Then Exit Do
Input #91, rname
Input #91, rtype
Input #91, rsongs
Input #91, rmaps
Input #91, roffs
Input #91, rstart
Input #91, rend
With roms(i)
.romheader = rhead
.romname = rname
.romtype = rtype
.songlist = rsongs
.maplist = rmaps
.tableoffs = roffs
.songstart = rstart
.songend = rend
End With
i = i + 1
Loop
romcount = i
Close #91
End Sub

 Public Function getext(ByVal filenam As String) As String
posx = InStrRev(filenam, ".")
If posx = 0 Then
getext = ""
Else
getext = Mid(filenam, posx + 1, Len(filenam) - posx)
End If
End Function
