VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cbm_main 
   BackColor       =   &H00000000&
   Caption         =   "Cerebellum Simulation 2019"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   13185
   ClientWidth     =   7830
   DrawMode        =   8  'Xor Pen
   DrawStyle       =   6  'Inside Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   -257.254
   ScaleMode       =   0  'User
   ScaleTop        =   101
   ScaleWidth      =   187.703
   Begin VB.CheckBox Check4 
      BackColor       =   &H00000000&
      Caption         =   "PC Form BW"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   145
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton WeightChangeButton 
      Caption         =   "gr BC Up"
      Height          =   495
      Index           =   5
      Left            =   5160
      TabIndex        =   144
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton WeightChangeButton 
      Caption         =   "gr BC Down"
      Height          =   495
      Index           =   4
      Left            =   6000
      TabIndex        =   143
      Top             =   7800
      Width           =   735
   End
   Begin VB.CheckBox MLIplasticityCheck 
      BackColor       =   &H00000000&
      Caption         =   "MLI plasticity"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   142
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CheckBox PC_record 
      BackColor       =   &H00000000&
      Caption         =   "PC Packed "
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   141
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton CompetingStimulusButton 
      Caption         =   "Change!"
      Height          =   375
      Left            =   5520
      TabIndex        =   140
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox CompetingStimulusText 
      Height          =   375
      Left            =   4800
      TabIndex        =   139
      Text            =   "0"
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox CompetingCheck 
      BackColor       =   &H00000000&
      Caption         =   "Competing Stimulus"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   138
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton GGButton 
      Caption         =   "Golgi Histos"
      Height          =   495
      Left            =   6360
      TabIndex        =   137
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00000000&
      Caption         =   "Record"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   136
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton ConductanceWinButton 
      Caption         =   "Conductance Win"
      Height          =   375
      Left            =   3240
      TabIndex        =   135
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox MFCollateralsCheck 
      BackColor       =   &H00000000&
      Caption         =   "MF Collaterals"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   134
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Frame MFCFrame 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2160
      TabIndex        =   129
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "5.3%"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   133
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "4.0%"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   132
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "2.6%"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   131
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "1.3%"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   130
         Top             =   840
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CheckBox GolToGolCheck 
      BackColor       =   &H00000000&
      Caption         =   "Golgi Golgi Inhibition"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   128
      Top             =   4920
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   127
      Text            =   "Text2"
      Top             =   10080
      Width           =   975
   End
   Begin VB.CheckBox PCtoPCSynapsesCheck 
      BackColor       =   &H00000000&
      Caption         =   "PC to Basket synapses"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   126
      Top             =   4680
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox PCNucCFCheck 
      BackColor       =   &H00000000&
      Caption         =   "AutoSave PC NUC CF"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   125
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Frame MFframe 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2160
      TabIndex        =   120
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "6%"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   124
         Top             =   840
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "12%"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   123
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "25%"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   122
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "50%"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   121
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.CheckBox MFbyNuc 
      BackColor       =   &H00000000&
      Caption         =   "Response MFs"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   119
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton RecordCommand 
      Caption         =   "Record Form"
      Height          =   375
      Left            =   6240
      TabIndex        =   118
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox RecordCheck 
      BackColor       =   &H00000000&
      Caption         =   "Record Mode"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4080
      TabIndex        =   117
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton AmpCommand 
      Caption         =   "dn"
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   116
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton AmpCommand 
      Caption         =   "up"
      Height          =   375
      Index           =   0
      Left            =   6000
      TabIndex        =   115
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox AmpCheck 
      BackColor       =   &H80000008&
      Caption         =   "Amplitude Mode"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4080
      TabIndex        =   112
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CheckBox CFCouplingBox 
      BackColor       =   &H00000000&
      Caption         =   "CF Oscillations"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   111
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CheckBox HomeoCheck 
      BackColor       =   &H00000000&
      Caption         =   "CF precludes STP"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   7
      Left            =   960
      TabIndex        =   110
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CheckBox HomeoCheck 
      BackColor       =   &H00000000&
      Caption         =   "Gol > gr heterogeneous"
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   109
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   108
      Text            =   "65"
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   107
      Text            =   "65"
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox HomeoCheck 
      BackColor       =   &H00000000&
      Caption         =   "Nuc > Cf Asynchronous "
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   106
      Top             =   3480
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox BranchSpecificLTPCheck 
      BackColor       =   &H00000000&
      Caption         =   "Branch Specific LTP"
      Enabled         =   0   'False
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   120
      TabIndex        =   104
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CheckBox BranchSpecificLTDCheck 
      BackColor       =   &H00000000&
      Caption         =   "Branch Specific LTD"
      Enabled         =   0   'False
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   120
      TabIndex        =   103
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton USUpDownButton 
      Caption         =   "dn"
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   102
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton USUpDownButton 
      Caption         =   "up"
      Height          =   375
      Index           =   0
      Left            =   7080
      TabIndex        =   101
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton DemoMode 
      Caption         =   "Demo Mode"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3600
      TabIndex        =   99
      Top             =   8400
      Width           =   735
   End
   Begin VB.CommandButton EyelidButton 
      Caption         =   "View Eyelid Response"
      Height          =   495
      Left            =   3840
      TabIndex        =   98
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox ChimeCheck 
      BackColor       =   &H00000000&
      Caption         =   "Chime mode"
      ForeColor       =   &H00FF80FF&
      Height          =   375
      Left            =   3600
      TabIndex        =   97
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox RepeatLastNumber 
      Height          =   285
      Left            =   5520
      TabIndex        =   94
      Text            =   "32000"
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Activity History"
      Height          =   375
      Left            =   3960
      TabIndex        =   52
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stats window"
      Height          =   375
      Left            =   4680
      TabIndex        =   39
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Weights"
      Height          =   375
      Left            =   3840
      TabIndex        =   50
      Top             =   11520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Activity Window"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CheckBox RepeatModeCheck 
      BackColor       =   &H00000000&
      Caption         =   "Repeat Mode"
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   4320
      TabIndex        =   93
      Top             =   600
      Width           =   1455
   End
   Begin VB.CheckBox CouplingCheck 
      BackColor       =   &H00000000&
      Caption         =   "CF Electrical coupling"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   92
      Top             =   3960
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog OutputCD 
      Left            =   5640
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.dat|*.dat"
   End
   Begin VB.Frame ContextFrame 
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   360
      TabIndex        =   89
      Top             =   10200
      Width           =   1335
      Begin VB.OptionButton ContextOption 
         BackColor       =   &H00000000&
         Caption         =   "Context 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   91
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton ContextOption 
         BackColor       =   &H00000000&
         Caption         =   "Context 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   90
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog RunCD 
      Left            =   6360
      Top             =   9360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.exp|*.exp"
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   8
      Left            =   1440
      TabIndex        =   85
      Top             =   10200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   84
      Top             =   10200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   9
      Left            =   1440
      TabIndex        =   79
      Top             =   10560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   78
      Top             =   10560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   8
      Left            =   840
      TabIndex        =   87
      Top             =   10200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change US1 onset"
      Height          =   255
      Index           =   8
      Left            =   1440
      TabIndex        =   86
      Top             =   10200
      Width           =   1695
   End
   Begin VB.CheckBox US2 
      BackColor       =   &H00000000&
      Caption         =   "US 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   83
      Top             =   10560
      Width           =   735
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   9
      Left            =   840
      TabIndex        =   81
      Top             =   10560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change US2 onset"
      Height          =   255
      Index           =   9
      Left            =   1440
      TabIndex        =   80
      Top             =   10560
      Width           =   1695
   End
   Begin VB.CheckBox CS4 
      BackColor       =   &H00000000&
      Caption         =   "CS 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   77
      Top             =   9360
      Width           =   735
   End
   Begin VB.CheckBox CS3 
      BackColor       =   &H00000000&
      Caption         =   "CS 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   76
      Top             =   8520
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   7
      Left            =   2160
      TabIndex        =   71
      Top             =   9720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   7
      Left            =   1440
      TabIndex        =   70
      Top             =   9720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   6
      Left            =   2160
      TabIndex        =   67
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   66
      Top             =   9360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS4 duration"
      Height          =   255
      Index           =   7
      Left            =   1440
      TabIndex        =   73
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS4 onset"
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   72
      Top             =   9360
      Width           =   1695
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   7
      Left            =   840
      TabIndex        =   69
      Top             =   9720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   6
      Left            =   840
      TabIndex        =   68
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   65
      Top             =   6150
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4856
            MinWidth        =   4856
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3090
            MinWidth        =   3090
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2796
            MinWidth        =   2796
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4194
            MinWidth        =   4194
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "gr>PC plasticity"
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   3240
      TabIndex        =   60
      Top             =   1440
      Width           =   1335
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Abbott"
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   64
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "binary"
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   63
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "graded"
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   62
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "gr>PC Plasticity"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.CheckBox HomeoCheck 
      BackColor       =   &H00000000&
      Caption         =   "STP"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   59
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton WeightChangeButton 
      Caption         =   "Reset Weights"
      Height          =   495
      Index           =   3
      Left            =   6840
      TabIndex        =   58
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton WeightChangeButton 
      Caption         =   "Average Weights"
      Height          =   495
      Index           =   2
      Left            =   4440
      TabIndex        =   57
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton WeightChangeButton 
      Caption         =   "Weights Down"
      Height          =   495
      Index           =   1
      Left            =   6000
      TabIndex        =   56
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton WeightChangeButton 
      Caption         =   "Weights Up"
      Height          =   495
      Index           =   0
      Left            =   5160
      TabIndex        =   55
      Top             =   6600
      Width           =   735
   End
   Begin VB.CheckBox HomeoCheck 
      BackColor       =   &H00000000&
      Caption         =   "PC Synaptic Scaling"
      ForeColor       =   &H00FF80FF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   54
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Save Activity Data"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   53
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Save activity each trial"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   51
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CheckBox HomeoCheck 
      BackColor       =   &H00000000&
      Caption         =   "Nuc Pre"
      ForeColor       =   &H00FF80FF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   49
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox HomeoCheck 
      BackColor       =   &H00000000&
      Caption         =   "PC Pre"
      ForeColor       =   &H00FF80FF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   48
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   47
      Top             =   8880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Number_of_trials_per_session 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   45
      Text            =   "1000"
      Top             =   720
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "MF gr Gol only"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   44
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CheckBox Weights_CHECK 
      BackColor       =   &H00000000&
      Caption         =   "Save weights each trial"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   43
      Top             =   1650
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Close all Wins"
      Height          =   420
      Left            =   6240
      TabIndex        =   42
      Top             =   4080
      Width           =   1260
   End
   Begin VB.CheckBox Scaling_gr_Gol 
      BackColor       =   &H00000000&
      Caption         =   "Granule to Golgi synaptic Scaling"
      Enabled         =   0   'False
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CheckBox Scaling_MF_Gol 
      BackColor       =   &H00000000&
      Caption         =   "MF to Golgi synaptic Scaling"
      Enabled         =   0   'False
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   1800
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog Histo_Dialog 
      Left            =   6960
      Top             =   9840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Purkinje window"
      Height          =   330
      Left            =   3360
      TabIndex        =   38
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Oscilloscope"
      Height          =   375
      Left            =   5160
      TabIndex        =   37
      Top             =   8760
      Width           =   1335
   End
   Begin VB.CheckBox MF_NUC_toggle 
      BackColor       =   &H00000000&
      Caption         =   "MF>NUC plasticity"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   840
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox gr_Purk_toggle 
      BackColor       =   &H00000000&
      Caption         =   "granule>Purkinje plasticity"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   600
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   32
      Top             =   8880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   5
      Left            =   840
      TabIndex        =   31
      Top             =   8880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   4
      Left            =   840
      TabIndex        =   28
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   27
      Top             =   8520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   26
      Top             =   8520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   3
      Left            =   840
      TabIndex        =   23
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   22
      Top             =   8040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   21
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   19
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   18
      Top             =   7680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   2
      Left            =   840
      TabIndex        =   17
      Top             =   7680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS2 onset"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   16
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   13
      Top             =   7200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   12
      Top             =   7200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   11
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   10
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   0
      Left            =   1455
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS1 onset"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   7
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CheckBox Scaling 
      BackColor       =   &H00000000&
      Caption         =   "Granule synaptic Scaling"
      Enabled         =   0   'False
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CheckBox CS2 
      BackColor       =   &H00000000&
      Caption         =   "CS 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   7680
      Width           =   735
   End
   Begin VB.CheckBox CS1 
      BackColor       =   &H00000000&
      Caption         =   "CS 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox US1 
      BackColor       =   &H00000000&
      Caption         =   "US 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   10200
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Rasters Window"
      Height          =   420
      Left            =   3720
      TabIndex        =   1
      Top             =   12120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   6960
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog raster_dialog 
      Left            =   6960
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.srf | *.srf"
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6960
      Top             =   9360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.cbm | *.cbm"
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS2 duration"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   24
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS1 duration"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   14
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS3 onset"
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   29
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS3 duration"
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   33
      Top             =   8880
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   45
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cbm_main 7.frx":0000
            Key             =   "Open"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cbm_main 7.frx":005E
            Key             =   "Save"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cbm_main 7.frx":00BC
            Key             =   "Play"
            Object.Tag             =   "Play"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cbm_main 7.frx":011A
            Key             =   "Pause"
            Object.Tag             =   "Pause"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cbm_main 7.frx":0178
            Key             =   "Purk"
            Object.Tag             =   "Purk"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cbm_main 7.frx":01D6
            Key             =   "Raster"
            Object.Tag             =   "Raster"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cbm_main 7.frx":0234
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox RepeatForCheck 
      BackColor       =   &H00000000&
      Caption         =   "Repeat For                  Experiments"
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   4320
      TabIndex        =   95
      Top             =   960
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CheckBox HomeoCheck 
      BackColor       =   &H00000000&
      Caption         =   "PC Intrinsic"
      ForeColor       =   &H00FF80FF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   96
      Top             =   2760
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog ExcelOutDialog 
      Left            =   6960
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.txt | *.txt"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   105
      Top             =   0
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   953
      ButtonWidth     =   1376
      ButtonHeight    =   794
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Saved Simulation"
            Object.Tag             =   "Open"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Simulation State"
            Object.Tag             =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Go"
            Object.Tag             =   "Go"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pause"
            Object.Tag             =   "Pause"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Weights"
            Object.Tag             =   "Weights"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Purk"
            Object.Tag             =   "Purk"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Raster"
            Object.Tag             =   "Raster"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Label AmpLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "6.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5640
      TabIndex        =   114
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "US"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   113
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label USLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "US"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6600
      TabIndex        =   100
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "1500"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   9
      Left            =   840
      TabIndex        =   88
      Top             =   10560
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "1500"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   8
      Left            =   840
      TabIndex        =   82
      Top             =   10200
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "1000"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   840
      TabIndex        =   75
      Top             =   9360
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "500"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   840
      TabIndex        =   74
      Top             =   9720
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Max Trials"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   46
      Top             =   720
      Width           =   735
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "500"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   34
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "1000"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   30
      Top             =   8520
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "500"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   25
      Top             =   8040
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "1000"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   20
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "550"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   15
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "1000"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   6
      Top             =   6840
      Width           =   495
   End
   Begin VB.Menu f_menu 
      Caption         =   "&File"
      Begin VB.Menu synaptogenesis_menu 
         Caption         =   "Build New Simulation"
         Index           =   0
      End
      Begin VB.Menu file_menu 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin VB.Menu file_menu 
         Caption         =   "&Save"
         Index           =   2
      End
      Begin VB.Menu raster_save_menu 
         Caption         =   "&Raster_save_PC_CF_N"
      End
      Begin VB.Menu VOR_menu 
         Caption         =   "New VOR simulation"
      End
      Begin VB.Menu SaveSpikesMenu 
         Caption         =   "Save Spikes"
      End
      Begin VB.Menu SavePCPackedMenu 
         Caption         =   "Save PC Packed"
      End
      Begin VB.Menu Exit_menu 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu ExpMenu 
      Caption         =   "&Exp"
   End
   Begin VB.Menu SecondExperimentMenu 
      Caption         =   "&2nd Exp"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu UseSecondMenu 
         Caption         =   "Alternating with first"
         Index           =   1
      End
      Begin VB.Menu UseSecondMenu 
         Caption         =   "Repeat Second"
         Index           =   2
      End
   End
   Begin VB.Menu OutputMenu 
      Caption         =   "&Output"
      Visible         =   0   'False
   End
   Begin VB.Menu Builditmenu 
      Caption         =   "&Build"
      Begin VB.Menu BuildMenu 
         Caption         =   "Build &Trial"
         Index           =   1
      End
      Begin VB.Menu BuildMenu 
         Caption         =   "Build &Block"
         Index           =   2
      End
      Begin VB.Menu BuildMenu 
         Caption         =   "Build &Session"
         Index           =   3
      End
      Begin VB.Menu BuildMenu 
         Caption         =   "Build &Experiment"
         Index           =   4
      End
   End
   Begin VB.Menu MossyFiberMenu 
      Caption         =   "Mossy Fibers"
      Begin VB.Menu AddMFMenu 
         Caption         =   "Add Recorded MFs"
      End
      Begin VB.Menu AlterMFMenu 
         Caption         =   "Alter Mossy fibers"
      End
   End
   Begin VB.Menu diag_menu 
      Caption         =   "&Connectivity"
      Begin VB.Menu diagnostics_menu 
         Caption         =   "Glomerulus to granule and Golgi"
         Index           =   1
      End
      Begin VB.Menu diagnostics_menu 
         Caption         =   "Golgi to granule"
         Index           =   2
      End
      Begin VB.Menu diagnostics_menu 
         Caption         =   "Granule to Golgi"
         Index           =   3
      End
      Begin VB.Menu diagnostics_menu 
         Caption         =   "cell by cell"
         Index           =   4
      End
      Begin VB.Menu diagnostics_menu 
         Caption         =   "CS mossy fibers"
         Index           =   5
      End
      Begin VB.Menu diagnostics_menu 
         Caption         =   "Real Time"
         Index           =   6
      End
   End
   Begin VB.Menu speed_menu 
      Caption         =   "go"
   End
   Begin VB.Menu pause_menu 
      Caption         =   "&Pause"
   End
   Begin VB.Menu resume_menu 
      Caption         =   "&Resume"
   End
   Begin VB.Menu SimpsonMainMenu 
      Caption         =   "Simpson"
      Begin VB.Menu SimpsonMenu 
         Caption         =   "Granule cells"
         Index           =   0
      End
      Begin VB.Menu SimpsonMenu 
         Caption         =   "Golgi cells"
         Index           =   1
      End
      Begin VB.Menu SimpsonMenu 
         Caption         =   "Stellate"
         Index           =   2
      End
      Begin VB.Menu SimpsonMenu 
         Caption         =   "Basket"
         Index           =   3
      End
   End
   Begin VB.Menu analysis_menu 
      Caption         =   "Analysis"
      Begin VB.Menu Rasters_histos_toggle_menu 
         Caption         =   "Rasters_Histos_engaged"
         Checked         =   -1  'True
      End
      Begin VB.Menu histo_menu 
         Caption         =   "Histos"
      End
      Begin VB.Menu StatsWindowMenu 
         Caption         =   "Stats Window"
      End
      Begin VB.Menu Big_Rasters_Menu 
         Caption         =   "Record Mode"
      End
      Begin VB.Menu clear_rasters_menu 
         Caption         =   "Clear Rasters"
         Enabled         =   0   'False
      End
      Begin VB.Menu StimMenu 
         Caption         =   "Stimulation"
         Enabled         =   0   'False
      End
      Begin VB.Menu CompressedDatFile_Menu 
         Caption         =   "Compressed .dat file"
      End
   End
End
Attribute VB_Name = "cbm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddMFMenu_Click()
    AddMossyFiberForm.Visible = True
End Sub

Private Sub AlterMFMenu_Click()
Dim i As Integer
Dim j As Integer
Dim CSFreq(4) As Single

Dim NumCS(4) As Single

    For i = 1 To 2
        For j = 1 To MFNUMBER
            mfsBackup(i, j) = MFS(i, j)
        Next j
    Next i
    
    For i = 1 To 4
        CSFreq(i) = 0
        NumCS(i) = 0
    Next i
    
    
    For j = 1 To 4
        For i = 1 To MFNUMBER
            If MFS(1, i).CStype = j Then
                CSFreq(j) = CSFreq(j) + (MFS(1, i).CSFreq * 1000)
                NumCS(j) = NumCS(j) + 1
            End If
        Next i
        
        If NumCS(j) <> 0 Then
            CSFreq(j) = CSFreq(j) / NumCS(j)
            i = Int(CSFreq(j))
            AlterMFForm.Label2(j - 1).Caption = i
        End If
        AlterMFForm.Label2(j + 3).Caption = NumCS(j)
    Next j

    AlterMFForm.Visible = True
    AlterMFForm.Command5.Enabled = False
End Sub

Private Sub AmpCheck_Click()
    If AmpCheck.Value = Checked Then
        AmpCommand(0).Visible = True
        AmpCommand(1).Visible = True
        AmpLabel.Visible = True
        AmpLabel.Caption = Format(Str(AmpModeAmp), "#.#")
        AmpMode = 1
    Else
        AmpCommand(0).Visible = False
        AmpCommand(1).Visible = False
        AmpLabel.Visible = False
        AmpMode = 0
    End If
End Sub

Private Sub AmpCommand_Click(Index As Integer)

    If Index = 0 Then
        AmpModeAmp = AmpModeAmp + 0.1
    Else
        AmpModeAmp = AmpModeAmp - 0.1
    End If
    AmpLabel.Caption = Format(Str(AmpModeAmp), "#.#")
End Sub

Private Sub Big_Rasters_Menu_Click()
Dim i As Integer
    If Do_Big_Rasters = 0 Then
        Do_Big_Rasters = 1
        RecordCheck.Enabled = False
        RecordCommand.Visible = True
       
        RecordForm.Visible = True
        RecordForm.ScaleLeft = -50
        RecordForm.ScaleWidth = 1050
        RecordForm.ScaleTop = 1200
        RecordForm.ScaleHeight = -1200
        RecordForm.DrawWidth = 2
        RecordForm.BackColor = vbBlack
        Command4.Visible = True
       
        For i = 1 To 4
            RRCellType(i) = 4
            RRCellNum(i) = i
        Next i
        For i = 5 To 12
            RRCellType(i) = 3
            RRCellNum(i) = i - 4
        Next i
        For i = 13 To 36
            RRCellType(i) = 2
            RRCellNum(i) = i - 12
        Next i
    Else
        Do_Big_Rasters = 0
        RecordCheck.Enabled = True
    End If
    If Big_Rasters_Menu.Checked = True Then Big_Rasters_Menu.Checked = False Else Big_Rasters_Menu.Checked = True
End Sub

Private Sub BuildMenu_Click(Index As Integer)
    Select Case Index
        Case 1
            TrialsForm.Visible = True
            TrialsForm.Command1(3).Value = True
        Case 2
            Block_edit_form.Visible = True
            Block_edit_form.Command1(3).Value = True
        Case 3
            session_edit_form.Visible = True
            session_edit_form.Command1(3).Value = True
        Case 4
            Experiment_edit_form.Visible = True
            Experiment_edit_form.Command1(3).Value = True
    End Select
End Sub

Private Sub Cancel_buttons_Click(Index As Integer)
    change_text(Index).Visible = False
    Change_button(Index).Visible = False
    Cancel_buttons(Index).Visible = False
End Sub

Private Sub Change_button_Click(Index As Integer)
    If Int(change_text(Index).Text) > 0 And Int(change_text(Index).Text) < 4000 Then
        OnsetLabels(Index).Caption = change_text(Index).Text
        Select Case Index
        Case 0
            cs_onset(1) = Int(OnsetLabels(0))
        Case 1
            cs_duration(1) = Int(OnsetLabels(1))
        Case 2
            cs_onset(2) = Int(OnsetLabels(2))
        Case 3
            cs_duration(2) = Int(OnsetLabels(3))
        Case 4
            cs_onset(3) = Int(OnsetLabels(4))
        Case 5
            cs_duration(3) = Int(OnsetLabels(5))
        Case 6
            cs_onset(4) = Int(OnsetLabels(6))
        Case 7
            cs_duration(4) = Int(OnsetLabels(7))
        Case 8
            US_onset(1) = Int(OnsetLabels(8))
        Case 9
            US_onset(2) = Int(OnsetLabels(9))
        End Select
    End If
    change_text(Index).Visible = False
    Change_button(Index).Visible = False
    Cancel_buttons(Index).Visible = False
End Sub

Private Sub Change_stimuli_Click(Index As Integer)
    change_text(Index).Text = OnsetLabels(Index).Caption
    change_text(Index).Visible = True
    Change_button(Index).Visible = True
    Cancel_buttons(Index).Visible = True
End Sub

Private Sub Check1_Click()
    If MFgrGolOnly = 0 Then MFgrGolOnly = 1 Else MFgrGolOnly = 0
End Sub

Private Sub Check2_Click()
    If Check2.Value = Checked Then
        cbm_main.Command9.Enabled = True
    Else
        cbm_main.Command9.Enabled = False
    End If
End Sub

Private Sub Check3_Click()
If Check3.Value = Checked Then
    DoSpecialRecord = 1
Else
    DoSpecialRecord = 0
End If
End Sub

Private Sub Check4_Click()
    If Check4.Value = Checked Then
        PCFormColors = 0
        PC_form.BackColor = vbWhite
        PCcolor = vbBlack
        NCcolor = vbBlack
        CFcolor = vbBlack
    Else
        PCFormColors = 1
        PC_form.BackColor = vbBlack
        PCcolor = &HFF8080
        NCcolor = vbWhite
        CFcolor = RGB(255, 0, 0)
    End If
    PC_form.Cls
End Sub

Private Sub Command1_Click()
    If activity_window.Visible = False Then
        activity_window.Visible = True
    Else
        activity_window.Visible = False
    End If
End Sub




Private Sub Command2_Click()
    If raster_form.Visible = False Then
        raster_form.WindowState = 0
        raster_form.Visible = True
        DoEvents
        raster_form.Top = 0
        raster_form.Left = SysInfo1.WorkAreaWidth - PC_form.Width
        raster_form.Width = PC_form.Width
        raster_form.Height = SysInfo1.WorkAreaHeight
    Else
        raster_form.Visible = False
    End If
End Sub

Private Sub Command3_Click()
    If OScope.Visible = True Then OScope.Visible = False Else OScope.Visible = True
End Sub

Private Sub Command4_Click()
    If PC_form.Visible = True Then PC_form.Visible = False Else PC_form.Visible = True
End Sub

Private Sub Command5_Click()
If Stats_form.Visible = True Then Stats_form.Visible = False Else Stats_form.Visible = True
Stats_form.Cell_option(0).Value = True
End Sub

Private Sub Command6_Click()
    Stats_form.Visible = False
    OScope.Visible = False
    raster_form.Visible = False
    activity_window.Visible = False
    PC_form.Visible = False
    PM_Form.Visible = False
    plasticity.Visible = False
    WeightHistory.Visible = False
    ActivityHistoryForm.Visible = False
  
    ConductanceForm.Visible = False
    RecordForm.Visible = False
End Sub

Private Sub Command7_Click()
Dim X As Integer
    
    If plasticity.Visible = False Then
        PM_Form.Visible = True
        plasticity.Visible = True
        'WeightHistory.Visible = True
    Else
        PM_Form.Visible = False
        plasticity.Visible = False
        WeightHistory.Visible = False
    End If
    
    plasticity.ScaleWidth = SYNUMBER
    'plasticity.ScaleTop = 1
    'plasticity.ScaleHeight = -1
    plasticity.Cls
    
End Sub

Private Sub Command8_Click()
Dim i As Integer
Dim j As Integer
Dim p As Single
Dim n As Single
Dim PCavg(1000) As Single
Dim TTT As Integer


    ActivityHistoryForm.Visible = True
    ActivityHistoryForm.ScaleWidth = 1000
    ActivityHistoryForm.ScaleHeight = -101
    ActivityHistoryForm.ScaleTop = 100
    ActivityHistoryForm.Cls
    ActivityHistoryForm.DrawWidth = 1
    For i = 10 To 100 Step 10
        ActivityHistoryForm.Line (1, i)-(1000, i), vbBlack
    Next i
    ActivityHistoryForm.DrawWidth = 2
    
    
    If Trials_this_time = 1 Then
        TTT = 1000
    Else
        TTT = Trials_this_time
    End If
    
    For j = 1 To 20
        PCavg(1) = PCavg(1) + PurkinjeActivity(j, 1)
        If ActivityHistoryForm.SMenu(j).Checked = True Then
            ActivityHistoryForm.PSet (1, PurkinjeActivity(j, 1) / 5), vbBlack
            
            For i = 2 To TTT - 1
                ActivityHistoryForm.Line -(i, PurkinjeActivity(j, i) / 5), vbBlack
                PCavg(i) = PCavg(i) + PurkinjeActivity(j, i)
            Next i
        End If
    Next j
    
    If ActivityHistoryForm.PCMenu(2).Checked = True Then
        ActivityHistoryForm.PSet (1, PCavg(1) / 100), vbRed
        For i = 1 To TTT - 1
            ActivityHistoryForm.Line -(i, PCavg(i) / 100), vbRed
        Next i
    End If
        
    For j = 1 To 6
        ActivityHistoryForm.PSet (1, NucleusActivity(j, 1) / 5), vbBlue
        For i = 1 To TTT - 1
            ActivityHistoryForm.Line -(i, NucleusActivity(j, i) / 5), vbBlue
        Next i
    Next j

End Sub

Private Sub Command9_Click()
    SaveActivityData
End Sub

Private Sub CompetingCheck_Click()
    If CompetingCheck.Value = Checked Then
        CompetingStimulusButton.Visible = True
        CompetingStimulusText.Visible = True
    Else
        CompetingStimulusButton.Visible = False
        CompetingStimulusText.Visible = False
        CompetingStimulusText.Text = 0
        CompetingStimulusNumber = 0
    End If
End Sub

Private Sub CompetingStimulusButton_Click()
    CompetingStimulusNumber = Val(CompetingStimulusText.Text)
End Sub

Private Sub CompressedDatFile_Menu_Click()
    If CompressedDatFile_Menu.Checked = False Then CompressedDatFile_Menu.Checked = True Else CompressedDatFile_Menu.Checked = False
    cbm_main.Caption = "Cerebellum Simulation 2010"
End Sub

Private Sub ConductanceWinButton_Click()
    If ConductanceForm.Visible = True Then
        ConductanceForm.Visible = False
    Else
        ConductanceForm.Visible = True
        Call ConductanceForm.Refreshg
        'Call ConductanceForm.RefreshOriginal
    End If
End Sub

Private Sub ContextOption_Click(Index As Integer)
    CurrentContext = Index + 1
End Sub

Private Sub CouplingCheck_Click()
    If CouplingCheck.Value = Checked Then
        CFCoupled = 1
    Else
        CFCoupled = 0
    End If
End Sub

Private Sub cs_mossy_fiber_menu_Click()
Dim X As Integer
Dim y As Integer
Dim gran As Integer
Dim c As Integer
Dim c2 As Integer
Dim dend As Integer

    gran = 0
    For X = 1 To GrX
        For y = 1 To GrY
            c = 0
            c2 = 0
            gran = gran + 1
            For dend = 1 To 4
                If (MFS(1, Gr(gran).MF(dend)).CStype = 1) Then c = c + 1
                If (MFS(1, Gr(gran).MF(dend)).CStype = 2) Then c2 = c2 + 1
            Next dend
            If c * c2 <> 0 Then
                cbm_main.PSet (X, y), &HFFFF&
            End If
        Next y
    Next X
End Sub

Private Sub CS1_Click()
    If CS1.Value = Checked Then
        CS_ON(1) = 1
    Else
        CS_ON(1) = 0
    End If
End Sub

Private Sub CS2_Click()
    If CS2.Value = Checked Then
        CS_ON(2) = 1
    Else
        CS_ON(2) = 0
    End If
End Sub

Private Sub CS3_Click()
    If CS3.Value = Checked Then
        CS_ON(3) = 1
    Else
        CS_ON(3) = 0
    End If
End Sub

Private Sub CS4_Click()
    If CS4.Value = Checked Then
        CS_ON(4) = 1
    Else
        CS_ON(4) = 0
    End If
End Sub

Private Sub DemoMode_Click()
Dim i As Integer
    
    cbm_main.Command2.Value = True
    raster_form.Visible = True
    raster_form.Top = 0
    raster_form.Left = 2800
     
    GWin.Top = 0
    GWin.Left = raster_form.Width + raster_form.Left
    GWin.Width = 5500
    GWin.Height = cbm_main.Top
    GWin.AutoRedraw = True
    GWin.Visible = True
    'Set PCImage = LoadPicture("c:\mike\cbm sims\purkinjelearning2.bmp")
    PCImageX = 0
    GWin.PaintPicture PCImage, PCImageX, 0, 12000, 16000
    
    PC_form.Visible = True
    PC_form.Top = 0
    PC_form.Left = GWin.Left + GWin.Width
    
    PC_form.SetFocus
    raster_form.SetFocus
    GWin.SetFocus
    
    For i = 0 To 6
        raster_form.cell_menu2(i).Checked = False
    Next i
    raster_form.cell_menu2(6).Checked = True
    raster_Cell_type = 6
   
End Sub

Private Sub diagnostics_menu_Click(Index As Integer)
Dim X As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim posxx As Integer
Dim posyy As Integer
Dim gc As Integer
  
    diagnostic_mode = 0
    GWin.BackColor = vbBlack
    GWin.Left = 1000
    GWin.Top = 1000
    GWin.Height = 10000
    GWin.Width = 12000
    GWin.ScaleTop = 101
    GWin.ScaleHeight = -101
    GWin.ScaleLeft = 0
    GWin.ScaleWidth = 121
    GWin.Visible = True
    GWin.Cls
    GWin.Caption = "Simulation connectivity"
  
  
  If Index = 1 Then
  For X = 1 To GlX * GlY
    GWin.Cls
    posxx = X Mod GlX
    If posxx = 0 Then posxx = 30
    posyy = Int((X - 1) / GlX) + 1
    
    GWin.Line (posxx * GrGlScaleX, posyy * GrGlScaleY)-(posxx * GrGlScaleX + 1, posyy * GrGlScaleY + 1), RGB(0, 255, 0), BF
    For i = 1 To GrX
    For j = 1 To GrY
      gc = ((j - 1) * GrX) + i
      For k = 1 To Gr(gc).numdend
        If prex(k) = X Then
          GWin.PSet (i, j), RGB(255, 255, 255)
        End If
      Next k
      
    Next j
    Next i
    For i = 1 To GoX
      For j = 1 To GoY
        gc = ((j - 1) * GoX) + i
        For k = 1 To numGoGlDend
          If Gol(gc).preGl(k) = X Then GWin.Line (i * GrGoScaleX, j * GrGoScaleY)-(i * GrGoScaleX + 0.5, j * GrGoScaleY + 0.5), RGB(0, 0, 255), BF
        Next k
      Next j
    Next i
    DoEvents
  Next X
  ElseIf Index = 2 Then
    GWin.Cls
    For X = 1 To GoX * GoY
      GWin.Cls
      posxx = X Mod GoX
      If posxx = 0 Then posxx = GoX
      posyy = Int((X - 1) / GoX) + 1
      
      GWin.Line (posxx * GrGoScaleX, posyy * GrGoScaleY)-(posxx * GrGoScaleX + 1, posyy * GrGoScaleY + 1), RGB(255, 0, 0), BF
      For i = 1 To GrX
      For j = 1 To GrY
        gc = ((j - 1) * GrX) + i
        For k = 1 To Gr(gc).numdend
          If Gr(gc).Gol(k) = X Then GWin.PSet (i, j), RGB(255, 255, 255)
        Next k
      Next j
      Next i
      DoEvents
    Next X
  ElseIf Index = 3 Then
    GWin.Cls
    For X = 1 To GrX * GrY
      GWin.Cls
      posxx = X Mod GrX
      If posxx = 0 Then posxx = GrX
      posyy = Int((X - 1) / GrX) + 1
      GWin.PSet (posxx, posyy), RGB(255, 255, 255)
      
      For i = 1 To GoX
      For j = 1 To GoY
        gc = ((j - 1) * GoX) + i
        For k = 1 To numGoGrDend
          If Gol(gc).preGr(k) = X Then GWin.Line (i * GrGoScaleX, j * GrGoScaleY)-(i * GrGoScaleX + 1, j * GrGoScaleY + 1), RGB(0, 0, 255), BF
        Next k
      Next j
      Next i
      DoEvents
    Next X
  ElseIf Index = 4 Then
    diagnostic_mode = 1
    For i = 1 To GrX
      For j = 1 To GrY
        GWin.PSet (i, j), RGB(150, 150, 150)
      Next j
    Next i
    For i = 1 To GoX
      For j = 1 To GoY
          GWin.Line (i * GrGoScaleX, j * GrGoScaleY)-(i * GrGoScaleX + 0.5, j * GrGoScaleY + 0.5), RGB(50, 50, 150), BF
          GWin.Line (i * GrGlScaleX + 0.5, j * GrGlScaleY + 0.5)-(i * GrGlScaleX + 0.5 + 0.5, j * GrGlScaleY + 0.5 + 0.5), RGB(50, 150, 50), BF
      Next j
    Next i
  ElseIf Index = 6 Then
    If ShowRealTime = 0 Then
        ShowRealTime = 1
        GWin.Height = 10000
        GWin.Width = 12000
        GWin.Visible = True
        
    Else
        ShowRealTime = 0
        GWin.Visible = False
    End If
  End If
End Sub

Private Sub example_cells_menu_Click(Index As Integer)
    
End Sub

Private Sub DisplayModeButton_Click()

End Sub

Private Sub Exit_menu_Click()
    QuitUserInterface
End Sub

Private Sub ExpMenu_Click()
Dim SessionFilename As String
Dim BlockFilename(25) As String
Dim TrialFilename(25) As String

Dim f As String
Dim t As Integer
Dim b As Integer
Dim s As Integer
Dim i As Integer

Dim TrialInfo As TrialStruct
Dim counter As Integer
Dim NTrials As Integer
Dim NBlocks As Integer

    TCounter = 0
    If CommandLine = 1 Then
        RunCD.filename = ExpFileName
    Else
        RunCD.filename = ""
        RunCD.ShowOpen
    End If
    If RunCD.filename <> "" Then
        'TrialsThisRun = 0
        Close #22
        Open RunCD.filename For Input As #22
        counter = 0
        While Not (EOF(22))
            counter = counter + 1
            Input #22, SessionNames(counter)
        Wend
        NumberofSessions = counter
        'Debug.Print "Sessions: "; NumberofSessions
        Close #22
        For s = 1 To NumberofSessions
            Open SessionNames(s) For Input As #22
            counter = 0
            While Not (EOF(22))
                counter = counter + 1
                Input #22, BlockFilename(counter)
                If BlockFilename(counter) = "yes" Or BlockFilename(counter) = "no" Then
                    If BlockFilename(counter) = "yes" Then
                        STPFadeSession = 1
                    Else
                        STPFadeSession = 0
                    End If
                    STPFades(s) = STPFadeSession
                    counter = counter - 1
                    Input #22, CurrentContext
                    SessionContexts(s) = CurrentContext
                    Debug.Print CurrentContext
                End If
            Wend
            NBlocks = counter
            'Debug.Print "Blocks"; NBlocks
            Close 22
            For b = 1 To NBlocks
                Open BlockFilename(b) For Input As #22
                counter = 0
                While Not (EOF(22))
                    counter = counter + 1
                    Input #22, TrialFilename(counter)
                Wend
                NTrials = counter
                'Debug.Print BlockFilename(b), NTrials
                Close 22
                For t = 1 To NTrials
                    Open TrialFilename(t) For Binary As 22
                    Get 22, , TrialInfo
                    Close #22
                    TrialsPerSession(s, 1) = TrialsPerSession(s, 1) + 1
                    TCounter = TCounter + 1
                    MasterTrialNames(TCounter) = TrialFilename(t)
                    RunData(TCounter) = TrialInfo
                Next t
            Next b
            'Debug.Print s, TrialsPerSession(s, 1)
        Next s
        SessionsThisExp = 1
        Close #22
        
' Checks to see whether there is a trial that exceeds the 2500 point limit for bunvis and prompts to select compressed file mode
        For t = 1 To TCounter
            'look for CS durations greater than 1500
            If RunData(t).CSduration(1) > 2000 Or RunData(t).CSduration(2) > 2000 Or RunData(t).USonset(1) > 3000 Or RunData(t).USonset(2) > 3000 Then
                t = TCounter + 1
                cbm_main.Caption = "Select compressed data file mode!!!!!!!!!!!!!!!"
            End If
        Next t
        
        LoadTrial (1)
        
        cbm_main.Number_of_trials_per_session.Text = TCounter + TrialCounter
        NO_OF_TRIALS_GIVEN_BY_USER = Int(cbm_main.Number_of_trials_per_session.Text)
        
        
        For t = 1 To TCounter
            f = ""
            i = Len(MasterTrialNames(t)) + 1
            While f <> "\"
                i = i - 1
                f = Mid$(MasterTrialNames(t), i, 1)
            Wend
            MasterTrialNames(t) = Mid$(MasterTrialNames(t), i + 1, Len(MasterTrialNames(t)) - i - 4)
        Next t
    End If
    cbm_main.StatusBar1.Panels(3).Text = MasterTrialNames(1)
    TrialThatEndsSession = TrialsPerSession(1, 1)
    TrialsPerSession(0, 1) = 1  'this stores the session number
    CurrentContext = SessionContexts(1)
    STPFadeSession = STPFades(1)
    If CurrentContext = 1 Then
        cbm_main.ContextOption(0).Value = True
    Else
        cbm_main.ContextOption(1).Value = True
    End If
    
    TrialsThisSession = 1
    RepeatExperimentsCounter = 1
    OutputMenu_Click
End Sub

Private Sub file_menu_Click(Index As Integer)

If Index = 1 Then
    If CommandLine = 1 Then
        CD1.filename = SimFileName
    Else
    
        CD1.filename = ""
        CD1.ShowOpen
    End If
    If CD1.filename <> "" Then
        Open CD1.filename For Binary As #1
        GetSim
        
        
        cbm_main.CS1.Value = CS_ON(1)
        cbm_main.CS2.Value = CS_ON(2)
        cbm_main.US1.Value = US1_ON
        cbm_main.US2.Value = US2_ON
'        cbm_main.Change_text(0).Text = cs1_onset
'        cbm_main.Change_text(1).Text = Str$(cs_duration)
'        cbm_main.Change_text(4).Text = Str$(US_onset)


        Close #1
        resume_menu.Enabled = True
        HistoDivisor = 1
        If NumCF = 1 Then
            MAXDRUS = 17.5
        Else
            MAXDRUS = 4
        End If
        cbm_main.USLabel.Caption = MAXDRUS
    
    End If
    
ElseIf Index = 2 Then
    CD1.filename = ""
    CD1.ShowSave
    If CD1.filename <> "" Then
        Open CD1.filename For Binary As #1
        
        SaveSim
    End If
End If
End Sub



Public Sub RemoveCancelMenuItem(frm As Form)
'        Dim hSysMenu As Long
'
'        'get the system menu for this form
'        hSysMenu = GetSystemMenu(frm.hWnd, 0)
'
'        'remove the close item
'        Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
'
'        'remove the separator that was over the close item
'        Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub




Private Sub Form_Load()
Dim imgX As ListImage
Dim btnX As Button

Dim i As Integer
Dim j As Integer
Dim filename As String

'Dim offset_arg As Integer

'Dim a_strArgs() As String

Dim iPri As Long
'iPri = ProcessPrioritySet(Priority:=ppBelowNormal)

'cbm_main.Caption = Command$
RRScale = 1
AmpMode = 0
AmpModeAmp = 6#
Erase Rasters
ConductanceForm.Visible = False
ConductanceWinButton.Visible = True
CFgDisplayMultiplier = 1

UsePCtoBasketSynapses = 1

MLIPlasticity = 0

If GolToGolCheck.Value = Checked Then DoGG = 1 Else DoGG = 0


MFsAdded = 0
MFsAddedTotal = 0
Erase MFsAddedFilenames
Erase MFsAddedSequence
Erase MFsAddedSpikes
Erase MFsAddedNumber

DoSpecialRecord = 0
DoPCPackedRecord = 0

ResponseDrivenMFs = 0
DoMFCollaterals = 0
DrivenMFsStepper = 16
DoEvents

RecordFormBackColor = vbBlack
RecordFormDotColor = vbWhite
RecordFormHistoColor = vbWhite
RecordFormCSColor = vbBlue
RecordFormDotSize = 1
RecordFormScaleRows = 1200

DoPurkClassic = 0
DoPurkPre = 0
DoNucPre = 0
DoPurkSynapticScaling = 0
DoLTPShortTermPlasticity = 0
DoPurkIntrinsic = 0
RemoveCancelMenuItem Me

PCFormColors = 1
PCcolor = &HFF8080
NCcolor = vbWhite
CFcolor = RGB(255, 0, 0)

RepeatMode = 0
RepeatExperimentsCounter = 1
RepeatExperimentsGoal = 32000

DoGrAudio = 0
DoGoAudio = 0
DoMFAudio = 0
DoBSAudio = 0
DoPCAudio = 0
DoNCAudio = 0
DoCFAudio = 0

DoSimpson = 0

STPblockedbyCF = 0

USLabel.Caption = MAXDRUS

CFCoupled = 1
SessionCounter = 1

SEED_FOR_RANDOM_NUMBER_GENERATOR = 123  ' default value
NO_OF_TRIALS_GIVEN_BY_USER = 1000

LTDURATION = 100
LTD_OFFSET = 60 ' was 100, but can't learn isi 150.. see wade's work...
DELMINUS = DELPLUS - (DELPLUS / (LTDURATION * 0.001))


grPCPlasticityType = 0      'graded

''LTD_OFFSET = 100
''LTDURATION = 100
For j = 1 To NumCF
    CF_spike_counter(j) = 0
Next j

  GrGlScaleX = (GrX * 1#) / (GlX * 1#)
  GrGoScaleX = (GrX * 1#) / (GoX * 1#)

  GrGlScaleY = (GrY * 1#) / (GlY * 1#)
  GrGoScaleY = (GrY * 1#) / (GoY * 1#)

  
  cbm_main.Visible = True

  'PC_form.Visible = False
  'cbm_main.Visible = False

  speed_menu.Enabled = False
  pause_menu.Enabled = False
  resume_menu.Enabled = False
  plasticity.Visible = False
  histo_switch = 1
  Trials_this_time = 1
  HistoDivisor = 1
  
  raster_mode = 0
  raster_Cell_type = 1
  Gr_weights_denom = 1
  MF_weights_denom = 0.7

  For i = 1 To 1000
    Raster_list(i) = i
  Next i
  raster_start = 1
  For i = 1 To NumCF
    ltdtime(i) = 1000
  Next i
  Do_Big_Rasters = 0
  CS_ON(1) = 1
  CS_ON(2) = 0
  CS_ON(3) = 0
  CS_ON(4) = 0
  US1_ON = 0
  US2_ON = 0
  
'  cs_onset(1) = Int(OnsetLabels(0))
'  cs_onset(2) = Int(OnsetLabels(2))
'  cs_onset(3) = Int(OnsetLabels(4))
'  cs_onset(4) = Int(OnsetLabels(6))
'  cs_duration(1) = Int(OnsetLabels(1))
'  cs_duration(2) = Int(OnsetLabels(3))
'  cs_duration(3) = Int(OnsetLabels(5))
'  cs_duration(4) = Int(OnsetLabels(7))
'  US_onset(1) = Int(OnsetLabels(8))
'  US_onset(2) = Int(OnsetLabels(9))
  CHANGEgrWEIGHTS = 1
  CHANGEmfWEIGHTS = 1
  
  Raster_Histos_ON = 1
  MFtoGr = 1
  MFtoGo = 1
  GRtoGo = 1
  GOtoGr = 1
  MFgrGolOnly = 0
  WEIGHTS_filename = ""
  
  network_filename = ""
  
  SaveTrialbyTrialDataFileName = "TrialbyTrialData.dat"
  NumRepeats = 1

  'Set PuffImage = LoadPicture("c:\mike\cbm sims\puff.bmp")
'*************** Default output filenames so it's not necessary to specify one *********************

        OutputBaseFilename = "SimOutput"
        CFoutputBaseFilename = "CFTimes"
        OutputFilename = OutputBaseFilename + "0001.txt"
        CFoutputFilename = CFoutputBaseFilename + "0001.txt"




    SaveTrialbyTrialDataFileName = strFilenameRoot + "-TrialbyTrialData.dat"




Close #18




repeated_savings_number = 0
   
   
   
   



''''''''' END   ----  Code added by HVoicu
For i = 0 To 6
    NeuralynxRecordingCells(i) = 0
Next i
CFAsynchronouse = 1
ConductanceFormMode = 0


      MFBGROUNDFREQMIN_CS = 1#
      MFBGROUNDFREQMAX_CS = 5#  '10  2007DEC

      MFCONTEXTFREQMIN = 60
      MFCONTEXTFREQMAX = 120

      MFCONTEXTFREQMIN2 = 60  '**mm2012  was 30
      MFCONTEXTFREQMAX2 = 120  '**mm2012  was 60

    'these are used as the background rates for all other cells
      MFBGROUNDFREQMIN = 1#
      MFBGROUNDFREQMAX = 10 '2007DEC

      MFTONICFREQ_INCREMENT = 80
      MFTONICFREQ_INCREMENT2 = 80 ' 80 FEB2008
      MFTONICFREQ_INCREMENT3 = 80
      MFTONICFREQ_INCREMENT4 = 80

      MFPHASICFREQ_INCREMENT = 100
      MFPHASICFREQ_INCREMENT2 = 100
      MFPHASICFREQ_INCREMENT3 = 100
      MFPHASICFREQ_INCREMENT4 = 100

    GolGrHeterogenous = 0
    ShowRealTime = 0
    
    If Command$ <> "" Then   ' launched from command line
        GetCommandLine
        'Debug.Print ArgArray(0), ArgArray(1)
        SimFileName = "C:\mike\cbm sims\Trial info\data\" + ArgArray(1) + ".cbm"
        cbm_main.Caption = SimFileName
        DoEvents
        ExpFileName = "C:\mike\cbm sims\Trial info\" + ArgArray(2) + ".exp"
        OutFileName = "C:\mike\cbm sims\Trial info\data\" + ArgArray(3) + ".dat"
        CommandLine = 1
        Raster_Histos_ON = 0
        Rasters_histos_toggle_menu.Checked = False
        For i = 4 To NumArgs
            Select Case ArgArray(i)
                Case "groff"
                    gr_Purk_toggle.Value = Unchecked
                    CHANGEgrWEIGHTS = 0
                Case "mfoff"
                    MF_NUC_toggle.Value = Unchecked
                    CHANGEmfWEIGHTS = 0
                Case "stpOff"
                    cbm_main.HomeoCheck(4).Value = Unchecked
                    DoLTPShortTermPlasticity = 0
                Case "stp2"
                    cbm_main.HomeoCheck(4).Value = Checked
                    DoLTPShortTermPlasticity = 1
                    cbm_main.HomeoCheck(7).Value = Checked
                    STPblockedbyCF = 1
                Case "asynchoff"
                    cbm_main.HomeoCheck(5).Value = Unchecked
                    CFAsynchronouse = 0
                Case "hetero"
                    cbm_main.HomeoCheck(6).Value = Checked
                    GolGrHeterogenous = 1
                Case "CoupleOff"
                    cbm_main.CouplingCheck.Value = Unchecked
                    CFCoupled = 0
                Case "rasters"
                    Raster_Histos_ON = 1
                    Rasters_histos_toggle_menu.Checked = True
                Case "mfgrgolonly"
                    MFgrGolOnly = 1
                    Check1.Value = Checked
                Case "saverasters"
                    SaveRastersON = 1
                Case "saveweights"
                    SaveWeightsON = 1
            End Select
        Next i
        file_menu_Click (1)
        ExpMenu_Click
        OutputMenu_Click
        PC_form.Visible = False
        speed_menu_Click
    Else                                    ' launched by double clicking icon in windows
        CommandLine = 0
'        filename = "Default.sim"
'        GetInputFile filename
        PC_form.Visible = True
    End If
    CollateralStepper = 1
    CompetingStimulusNumber = 0
    CompetingStimulusTotal = 0
    
End Sub


Public Sub QuitUserInterface()

'' Unload forms before end
  Dim f As Form
  For Each f In Forms
   Unload f
   Set f = Nothing
  Next

  End
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
  Dim gc As Integer
  Dim golgi As Integer
  Dim glom As Integer
  Dim i As Integer
  Dim k As Integer
  Dim j As Integer
  
  If diagnostic_mode = 1 Then
    cbm_main.Cls
    If Button = 1 And Shift = 0 Then
      PSet (Int(X), Int(y)), RGB(255, 255, 255)
      gc = ((Int(y) - 1) * GrX) + Int(X)
      'Debug.Print Int(x), Int(Y), gc
      For i = 1 To GoX
        For j = 1 To GoY
          X = ((j - 1) * GoX) + i
          
          For k = 1 To numGoGrDend
            If Gol(X).preGr(k) = gc Then Line (i * GrGoScaleX, j * GrGoScaleY)-(i * GrGoScaleX + 0.5, j * GrGoScaleY + 0.5), RGB(0, 0, 255), BF
          Next k
        Next j
      Next i
    ElseIf Button = 2 And Shift = 0 Then
        X = Int(X / GrGoScaleX)
        y = Int(y / GrGoScaleY)
        Line (X * GrGoScaleX, y * GrGoScaleY)-(X * GrGoScaleX + 0.5, y * GrGoScaleY + 0.5), RGB(0, 0, 255), BF
        golgi = ((Int(y) - 1) * GoX) + Int(X)
        
        For i = 1 To GrX
          For j = 1 To GrY
            gc = ((j - 1) * GrX) + i
        
            For k = 1 To Gr(gc).numdend
              If Gr(gc).Gol(k) = golgi Then PSet (i, j), RGB(255, 255, 255)
            Next k
          Next j
        Next i
    ElseIf Button = 1 And Shift = 1 Then
      X = Int(X / GrGlScaleX)
      y = Int(y / GrGlScaleY)
      Line (X * GrGlScaleX, y * GrGlScaleY)-(X * GrGlScaleX + 0.5, y * GrGlScaleY + 0.5), RGB(0, 255, 0), BF
      glom = ((Int(y) - 1) * GlX) + Int(X)
      'Debug.Print Int(X), Int(Y), gc
      For i = 1 To GrX
        For j = 1 To GrY
          gc = ((j - 1) * GrX) + i
          For k = 1 To Gr(gc).numdend
            If prex(k) = glom Then PSet (i, j), RGB(255, 255, 255)
          Next k
        Next j
      Next i
      For i = 1 To GoX
        For j = 1 To GoY
          golgi = ((j - 1) * GoX) + i
          For k = 1 To numGoGlDend
            
            If Gol(golgi).preGl(k) = glom Then Line (i * GrGlScaleX + 0.5, j * GrGlScaleY + 0.5)-(i * GrGlScaleX + 0.5 + 0.5, j * GrGlScaleY + 0.5 + 0.5), RGB(0, 0, 255), BF
          Next k
        Next j
      Next i
    End If

  End If
End Sub

Private Sub Form_Terminate()
    keepgoing = 0
'    Unload Progress
    Unload PC_form
    Unload Histo_form
    Unload plasticity
    Unload activity_window
    Unload histo_info_form
    Unload OScope
    Unload VORform
    Unload raster_form
    Unload Raster_adjust
    Unload GWin
    Unload Stats_form
    Unload PM_Form
    Unload WeightHistory
    Unload ActivityHistoryForm
    Unload TrialsForm
    Unload Block_edit_form
'    Unload StimulationForm
    Unload SimOptionsForm
   
    Unload AddMossyFiberForm
  
    Unload AlterMFForm
    Unload ConductanceForm
    Unload RecordForm
    Unload Experiment_edit_form
    Unload session_edit_form
    Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    keepgoing = 0
'    Unload Progress
    Unload PC_form
    Unload Histo_form
    Unload plasticity
    Unload activity_window
    Unload histo_info_form
    Unload OScope
    Unload VORform
    Unload raster_form
    Unload Raster_adjust
    Unload GWin
    Unload Stats_form
    Unload PM_Form
    Unload WeightHistory
    Unload ActivityHistoryForm
    Unload TrialsForm
    Unload Block_edit_form
'    Unload StimulationForm
    Unload SimOptionsForm
    
    Unload AddMossyFiberForm
   
    Unload AlterMFForm
    Unload ConductanceForm
    Unload RecordForm
    Unload Experiment_edit_form
    Unload session_edit_form
End Sub

Private Sub genesis_menu_Click()
  
End Sub

Private Sub histograms_menu_Click(Index As Integer)
'    Open "granule cell histograms from NT simulation" For Binary As #2
'    Put #2, , gran_histo
'    Close #2
End Sub

Private Sub GGButton_Click()
    Histo_form.Visible = True
    Histo_form.HM_menu_Click (2)
End Sub

Private Sub GolToGolCheck_Click()
    If DoGG = 1 Then
        DoGG = 0
        
    Else
        DoGG = 1
        
    End If
End Sub

Private Sub gr_Purk_toggle_Click()
    If CHANGEgrWEIGHTS = 1 Then CHANGEgrWEIGHTS = 0 Else CHANGEgrWEIGHTS = 1
End Sub




Private Sub histo_menu_Click()
    If Histo_form.Visible = True Then Histo_form.Visible = False Else Histo_form.Visible = True
    If histo_info_form.Visible = True Then histo_info_form.Visible = False Else histo_info_form.Visible = True
End Sub

Private Sub HomeoCheck_Click(Index As Integer)
    
    Select Case Index
        Case 0
            If DoPurkIntrinsic = 0 Then DoPurkIntrinsic = 1 Else DoPurkIntrinsic = 0
        Case 1
            If DoPurkPre = 0 Then DoPurkPre = 1 Else DoPurkPre = 0
        Case 2
            If DoNucPre = 0 Then DoNucPre = 1 Else DoNucPre = 0
        Case 3
            If DoPurkSynapticScaling = 0 Then DoPurkSynapticScaling = 1 Else DoPurkSynapticScaling = 0
        Case 4
            If DoLTPShortTermPlasticity = 0 Then
                DoLTPShortTermPlasticity = 1
                HomeoCheck(7).Visible = True
            Else
                DoLTPShortTermPlasticity = 0
                HomeoCheck(7).Visible = False
            End If
        Case 5
            If CFAsynchronouse = 0 Then CFAsynchronouse = 1 Else CFAsynchronouse = 0
        Case 6
            If GolGrHeterogenous = 0 Then GolGrHeterogenous = 1 Else GolGrHeterogenous = 0
        Case 7
            If STPblockedbyCF = 0 Then STPblockedbyCF = 1 Else STPblockedbyCF = 0
    End Select
        If DoPurkPre + DoPurkSynapticScaling + DoPurkIntrinsic > 0 Then Text1(0).Visible = True Else Text1(0).Visible = False
        If DoNucPre = 1 Then Text1(1).Visible = True Else Text1(1).Visible = False
End Sub

Private Sub ISIAnalysisCheck_Click()

End Sub

Private Sub MF_NUC_toggle_Click()
    If CHANGEmfWEIGHTS = 1 Then CHANGEmfWEIGHTS = 0 Else CHANGEmfWEIGHTS = 1
End Sub

Private Sub MFbyNuc_Click()
    If MFbyNuc.Value = Checked Then
        ResponseDrivenMFs = 1
        MFframe.Visible = True
    Else
        MFframe.Visible = False
        ResponseDrivenMFs = 0
    End If
End Sub

Private Sub MFCollateralsCheck_Click()
    If cbm_main.MFCollateralsCheck.Value = Checked Then
        DoMFCollaterals = 1
        MFCFrame.Visible = True
    Else
        DoMFCollaterals = 0
        MFCFrame.Visible = False
    End If
End Sub

Private Sub MLIplasticityCheck_Click()
    If MLIPlasticity = 1 Then
        MLIPlasticity = 0
    Else
        MLIPlasticity = 1
    End If
End Sub

Private Sub Number_of_trials_per_session_Change()
If Number_of_trials_per_session <> "" And Command$ = "" Then
   NO_OF_TRIALS_GIVEN_BY_USER = CLng(Number_of_trials_per_session)
End If
End Sub

Private Sub Option1_Click(Index As Integer)
    grPCPlasticityType = Index
End Sub

Private Sub Option2_Click(Index As Integer)
    Select Case Index
        Case 0
            CollateralStepper = 4
        Case 1
            CollateralStepper = 3
        Case 2
            CollateralStepper = 2
        Case 3
            CollateralStepper = 1
    End Select
End Sub

Private Sub OutputMenu_Click()
Dim FileTemp As String
Dim lentemp As Integer

    If CommandLine = 1 Then
        OutputCD.filename = OutFileName
    Else
        OutputCD.filename = ""
        OutputCD.ShowOpen
    End If
    If OutputCD.filename <> "" Then
        Close #24
        cbm_main.Caption = OutputCD.filename
        OpenDatFile
        FileTemp = OutputCD.filename
        lentemp = Len(FileTemp)
        OutputBaseFilename = Mid$(FileTemp, 1, lentemp - 4)
        CFoutputBaseFilename = OutputBaseFilename + "CFTimes"
        OutputFilename = OutputBaseFilename + "0001.txt"
        CFoutputFilename = CFoutputBaseFilename + "0001.txt"
    End If
End Sub

Public Sub pause_menu_Click()
    keepgoing = 0
    pause_menu.Enabled = False
    resume_menu.Enabled = True
    
     cbm_main.ExpMenu.Enabled = True
    cbm_main.SecondExperimentMenu.Enabled = True
    cbm_main.OutputMenu.Enabled = True
    cbm_main.Builditmenu.Enabled = True
    cbm_main.f_menu.Enabled = True
    raster_form.ShowATCMenu.Enabled = True
End Sub

Private Sub PC_record_Click()
    If PC_record.Value = Checked Then
        DoPCPackedRecord = 1
    Else
        DoPCPackedRecord = 0
    End If
End Sub

Private Sub PCtoPCSynapsesCheck_Click()
    If UsePCtoBasketSynapses = 1 Then UsePCtoBasketSynapses = 0 Else UsePCtoBasketSynapses = 1
End Sub

Private Sub raster_save_menu_Click()
    raster_dialog.filename = ""
    raster_dialog.ShowSave
    If raster_dialog.filename <> "" Then
        Open raster_dialog.filename For Binary As #4
        Put #4, , Rasters
        Close #4
    End If
End Sub

Private Sub Rasters_histos_toggle_menu_Click()
    If Raster_Histos_ON = 1 Then Raster_Histos_ON = 0 Else Raster_Histos_ON = 1
    If Rasters_histos_toggle_menu.Checked = True Then Rasters_histos_toggle_menu.Checked = False Else Rasters_histos_toggle_menu.Checked = True
End Sub

Private Sub rasters_menu_Click()

End Sub

Private Sub Reciprocal_Click()
    
End Sub

Private Sub RecordCheck_Click()
    If RecordCheck.Value = Checked Then
        RecordCommand.Visible = True
        DoRasters = 1
        Big_Rasters_Menu.Enabled = False
        RecordForm.Visible = True
    Else
        RecordCommand.Visible = False
        DoRasters = 0
        Big_Rasters_Menu.Enabled = True
        RecordForm.Visible = False
    End If
    
End Sub

Private Sub RecordCommand_Click()
    If RecordForm.Visible = False Then RecordForm.Visible = True Else RecordForm.Visible = False
End Sub

Private Sub RepeatLastNumber_Change()
    RepeatExperimentsGoal = Val(RepeatLastNumber.Text)
End Sub

Private Sub RepeatModeCheck_Click()
    If RepeatMode = 0 Then
        RepeatMode = 1
        RepeatLastNumber.Visible = True
        RepeatForCheck.Visible = True
    Else
        RepeatMode = 0
        RepeatLastNumber.Visible = False
        RepeatForCheck.Visible = False
    End If
End Sub



Public Sub resume_menu_Click()
Dim f As String
Dim i As Integer

    cbm_main.ExpMenu.Enabled = False
    cbm_main.SecondExperimentMenu.Enabled = False
    cbm_main.OutputMenu.Enabled = False
    cbm_main.Builditmenu.Enabled = False
    'cbm_main.f_menu.Enabled = False
    
    keepgoing = 1
    resume_menu.Enabled = False
    speed_menu.Enabled = False
    pause_menu.Enabled = True
    
    'Calculate_Time_Dependent_variables
    Close #5
    ''''''''''''''''''Open "CRs" For Output As #19
    raster_form.ShowATCMenu.Enabled = False
    While keepgoing = 1
        MFinput 0
        DoWork GrX * GrY, GoX * GoY, STELLATENUMBER, PCNUMBER
        DoEvents
    Wend

endlabel02:
End Sub

Private Sub Show_PC_Menu_Click()

End Sub


Private Sub SimpsonButton_Click()
    
End Sub

Private Sub SavePCPackedMenu_Click()
    CD1.filename = ""
    CD1.ShowSave
    If CD1.filename <> "" Then
        Close 11
        Open CD1.filename For Binary As #11
        Put #11, , PCPacked
        Close 11
    End If
End Sub

Private Sub SaveSpikesMenu_Click()
    
    Close #11
    Open "PCspikes" For Binary As 11
    Put #11, , PCSpikes
    Close 11
    Open "BCspikes" For Binary As 11
    Put #11, , BCSpikes
    Close 11
    Open "SCspikes" For Binary As 11
    Put #11, , SCSpikes
    Close 11
End Sub

Private Sub SimpsonMenu_Click(Index As Integer)
    StartSimpson Index
End Sub

Private Sub speed_menu_Click()
Dim f As String
Dim i As Integer

    cbm_main.ExpMenu.Enabled = False
    cbm_main.SecondExperimentMenu.Enabled = False
    cbm_main.OutputMenu.Enabled = False
    cbm_main.Builditmenu.Enabled = False
    'cbm_main.f_menu.Enabled = False
    
    TrialCounter = 1
    keepgoing = 1
    
    HistoDivisor = 1
    
    'cbm_main.Caption = "Cerebellum Simulation" + CommandLine
   
    Bincounter = 0
    bincounter5 = 0
    
    If Time_step_size = 1 Then
       bincounter5_temp = -1
    Else
       bincounter5_temp = 0
    End If
    
    speed_menu.Enabled = False
    pause_menu.Enabled = True
    
    Calculate_Time_Dependent_variables
    start_time = Timer
     
    While keepgoing = 1
        MFinput 0
        DoWork GrX * GrY, GoX * GoY, STELLATENUMBER, PCNUMBER
        DoEvents
    Wend
  
endlabel01:
End Sub



Private Sub ts_menu_Click(Index As Integer)
    
End Sub


Private Sub StatsWindowMenu_Click()
Command5.Value = True
End Sub

Private Sub StimMenu_Click()
'    If StimulationForm.Visible = False Then
'        StimulationForm.Visible = True
'    Else
'        StimulationForm.Visible = False
'    End If
End Sub

Private Sub synaptogenesis_menu_Click(Index As Integer)
    SimOptionsForm.Visible = True
End Sub

Private Sub Trial_Freq_button_Click(Index As Integer)
    
        CS_counter = Index
        If Index = 0 Then CS_counter = 0
        If Index = 5 Then CS_counter = 9
End Sub

Private Sub ToExcelMenu_Click()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Go"
            If speed_menu.Enabled = True Then
                speed_menu_Click
            ElseIf resume_menu.Enabled = True Then
                resume_menu_Click
            End If
        Case "Pause"
            If pause_menu.Enabled = True Then pause_menu_Click
        Case "Save"
            file_menu_Click (2)
        Case "Open"
            file_menu_Click (1)
            Calculate_Time_Dependent_variables
        Case "Raster"
            Command2.Value = True
        Case "Purk"
            Command4.Value = True
        Case "Weights"
            Command7.Value = True
    End Select
    
End Sub

Private Sub US1_Click()
    If US1.Value = Checked Then US1_ON = 1 Else US1_ON = 0
End Sub

Private Sub US2_Click()
    If US2.Value = Checked Then US2_ON = 1 Else US2_ON = 0
End Sub

Private Sub UseSecondMenu_Click(Index As Integer)
Dim SessionFilename As String
Dim BlockFilename(25) As String
Dim TrialFilename(25) As String

Dim f As String
Dim t As Integer
Dim b As Integer
Dim s As Integer
Dim i As Integer

Dim TrialInfo As TrialStruct
Dim counter As Integer
Dim NTrials As Integer
Dim NBlocks As Integer

    TCounter = 0
    RunCD.filename = ""
    RunCD.ShowOpen
    If RunCD.filename <> "" Then
        'TrialsThisRun = 0
        Close #22
        Open RunCD.filename For Input As #22
        counter = 0
        While Not (EOF(22))
            counter = counter + 1
            Input #22, SessionNames2(counter)
        Wend
        NumberofSessions2 = counter
        'Debug.Print "Sessions: "; NumberofSessions2
        Close #22
        For s = 1 To NumberofSessions2
            Open SessionNames2(s) For Input As #22
            counter = 0
            While Not (EOF(22))
                counter = counter + 1
                Input #22, BlockFilename(counter)
                If BlockFilename(counter) = "yes" Or BlockFilename(counter) = "no" Then
                    If BlockFilename(counter) = "yes" Then
                        STPFadeSession = 1
                    Else
                        STPFadeSession = 0
                    End If
                    STPFades2(s) = STPFadeSession
                    counter = counter - 1
                    Input #22, CurrentContext
                    SessionContexts2(s) = CurrentContext
                End If
            Wend
            NBlocks = counter
            'Debug.Print "Blocks"; NBlocks
            Close 22
            For b = 1 To NBlocks
                Open BlockFilename(b) For Input As #22
                counter = 0
                While Not (EOF(22))
                    counter = counter + 1
                    Input #22, TrialFilename(counter)
                Wend
                NTrials = counter
                'Debug.Print BlockFilename(b), NTrials
                Close 22
                For t = 1 To NTrials
                    Open TrialFilename(t) For Binary As 22
                    Get 22, , TrialInfo
                    Close #22
                    TrialsPerSession(s, 2) = TrialsPerSession(s, 2) + 1
                    TCounter = TCounter + 1
                    MasterTrialNames2(TCounter) = TrialFilename(t)
                    RunData2(TCounter) = TrialInfo
                Next t
            Next b
            'Debug.Print s, TrialsPerSession(s, 1)
        Next s
        SessionsThisExp = 1
        Close #22
        
        LoadTrial (1)
        
        'cbm_main.Number_of_trials_per_session.Text = TCounter + TrialCounter
        'NO_OF_TRIALS_GIVEN_BY_USER = Int(cbm_main.Number_of_trials_per_session.Text)
         
        For t = 1 To TCounter
            f = ""
            i = Len(MasterTrialNames2(t)) + 1
            While f <> "\"
                i = i - 1
                f = Mid$(MasterTrialNames2(t), i, 1)
            Wend
            MasterTrialNames2(t) = Mid$(MasterTrialNames2(t), i + 1, Len(MasterTrialNames2(t)) - i - 4)
        Next t
    End If
    'cbm_main.StatusBar1.Panels(3).Text = MasterTrialNames2(1)
    'TrialThatEndsSession = TrialsPerSession(1, 1)
    'TrialsPerSession(0, 1) = 1  'this stores the session number
    'CurrentContext = SessionContexts(1)
    'STPFadeSession = STPFades(1)
    If CurrentContext = 1 Then
        cbm_main.ContextOption(0).Value = True
    Else
        cbm_main.ContextOption(1).Value = True
    End If
    
    TrialsThisSession = 1
    'RepeatMode = 1
    RepeatModeCheck.Value = Checked
    RepeatForCheck.Visible = True
    RepeatForCheck.Caption = "Repeat 2nd               times"
    RepeatLastNumber.Visible = True
    If Index = 1 Then
        RepeatSecondExperimentCounter = 1
    Else
        RepeatSecondExperimentCounter = 0
    End If
    UseSecondExperiment = Index
End Sub

Private Sub USUpDownButton_Click(Index As Integer)
    If Index = 0 Then
        MAXDRUS = MAXDRUS + 0.1
    Else
        MAXDRUS = MAXDRUS - 0.1
    End If
    cbm_main.USLabel.Caption = MAXDRUS
End Sub

Private Sub VOR_menu_Click()
    Time_step_size = 1
    Calculate_Time_Dependent_variables
    SynaptoGenesis 2, 1
    Init_stuff
    init_VOR
    
    VORform.Visible = True
End Sub

Private Sub WeightChangeButton_Click(Index As Integer)
Dim i As Integer
Dim avg As Double
Dim X As Integer
Dim count As Integer

    avg = 0
    Select Case Index
        Case 3
            For X = 1 To SYNUMBER
                grWeight(X) = 1
            Next X
        Case 2
           For i = 1 To SYNUMBER
                avg = avg + grWeight(i)
           Next i
            avg = avg / SYNUMBER
        Case 0
            For i = 1 To SYNUMBER
                grWeight(i) = grWeight(i) * 1.05
                If grWeight(i) > 1 Then grWeight(i) = 1
            Next i
        Case 1
            For i = 1 To SYNUMBER
                grWeight(i) = grWeight(i) / 1.05
            Next i
        Case 5  ' gr BC weights up
            For X = 1 To BasketNUMBER
                For count = 1 To PFBasketsynUMBER
                    BCells(X).grW(count) = BCells(X).grW(count) * 1.05
                Next count
            Next X
        Case 6  ' gr BC weights down
            For X = 1 To BasketNUMBER
                For count = 1 To PFBasketsynUMBER
                    BCells(X).grW(count) = BCells(X).grW(count) / 1.05
                Next count
            Next X
    End Select
End Sub

Private Sub weights_menu_Click()

End Sub


