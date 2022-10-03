VERSION 5.00
Begin VB.Form SimOptionsForm 
   BackColor       =   &H00400000&
   Caption         =   "Simulation Options"
   ClientHeight    =   8865
   ClientLeft      =   750
   ClientTop       =   1725
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   5610
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   27
      Left            =   11280
      TabIndex        =   74
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   26
      Left            =   11280
      TabIndex        =   73
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   25
      Left            =   11280
      TabIndex        =   72
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   24
      Left            =   11280
      TabIndex        =   71
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   23
      Left            =   9480
      TabIndex        =   70
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   22
      Left            =   9480
      TabIndex        =   69
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   21
      Left            =   9480
      TabIndex        =   68
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   20
      Left            =   9480
      TabIndex        =   67
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   19
      Left            =   7680
      TabIndex        =   66
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   18
      Left            =   7680
      TabIndex        =   65
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   17
      Left            =   7680
      TabIndex        =   64
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   16
      Left            =   7680
      TabIndex        =   63
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel MF changes"
      Height          =   495
      Left            =   3840
      TabIndex        =   62
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "UPDATE MFs"
      Height          =   495
      Left            =   3840
      TabIndex        =   61
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   15
      Left            =   2400
      TabIndex        =   60
      Top             =   8280
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   14
      Left            =   2400
      TabIndex        =   59
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   13
      Left            =   2400
      TabIndex        =   58
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   12
      Left            =   2400
      TabIndex        =   57
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   11
      Left            =   1200
      TabIndex        =   56
      Top             =   8280
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   10
      Left            =   1200
      TabIndex        =   55
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   9
      Left            =   1200
      TabIndex        =   54
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   8
      Left            =   1200
      TabIndex        =   53
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   7
      Left            =   3960
      TabIndex        =   46
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   6
      Left            =   3960
      TabIndex        =   45
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   5
      Left            =   3960
      TabIndex        =   41
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   40
      Top             =   3840
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   35
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   34
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   29
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   28
      Top             =   3840
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00400000&
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   24
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00400000&
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   23
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   4560
      TabIndex        =   19
      Text            =   "65"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   4560
      TabIndex        =   18
      Text            =   "10"
      Top             =   2520
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   3360
      TabIndex        =   14
      Top             =   720
      Width           =   1695
      Begin VB.OptionButton UniformityOption 
         BackColor       =   &H00400000&
         Caption         =   "Non-uniform"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton UniformityOption 
         BackColor       =   &H00400000&
         Caption         =   "Uniform"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PC to Nuc, Nuc to CF"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Frame ConnectivityFrame 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H000000C0&
      Height          =   1095
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
      Begin VB.OptionButton ConnectOptions 
         BackColor       =   &H00400000&
         Caption         =   "Loops"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton ConnectOptions 
         BackColor       =   &H00400000&
         Caption         =   "Random"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "PC>NC>CF Connect"
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Frame UBCFrame 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   1800
      TabIndex        =   6
      Top             =   720
      Width           =   1215
      Begin VB.OptionButton UBCOptions 
         BackColor       =   &H00400000&
         Caption         =   "None"
         ForeColor       =   &H00FF80FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton UBCOptions 
         BackColor       =   &H00400000&
         Caption         =   "2 %"
         ForeColor       =   &H00FF80FF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "UBCs"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame CFframe 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H000000C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
      Begin VB.OptionButton CFOptions 
         BackColor       =   &H00400000&
         Caption         =   "12 CF"
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   93
         Top             =   840
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton CFOptions 
         BackColor       =   &H00400000&
         Caption         =   "4 CF"
         Enabled         =   0   'False
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton CFOptions 
         BackColor       =   &H00400000&
         Caption         =   "1 CF"
         Enabled         =   0   'False
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Climbing Fibers"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CommandButton SimOptionsButton 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton SimOptionsButton 
      Caption         =   "Build"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Context 7:"
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
      Height          =   375
      Index           =   37
      Left            =   10320
      TabIndex        =   92
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Context 8:"
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
      Height          =   375
      Index           =   36
      Left            =   10320
      TabIndex        =   91
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   35
      Left            =   10440
      TabIndex        =   90
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   34
      Left            =   10440
      TabIndex        =   89
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   33
      Left            =   10440
      TabIndex        =   88
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   32
      Left            =   10440
      TabIndex        =   87
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Context 5:"
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
      Height          =   375
      Index           =   31
      Left            =   8520
      TabIndex        =   86
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Context 6:"
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
      Height          =   375
      Index           =   30
      Left            =   8520
      TabIndex        =   85
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   29
      Left            =   8640
      TabIndex        =   84
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   28
      Left            =   8640
      TabIndex        =   83
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   27
      Left            =   8640
      TabIndex        =   82
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   26
      Left            =   8640
      TabIndex        =   81
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Context 3:"
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
      Height          =   375
      Index           =   25
      Left            =   6720
      TabIndex        =   80
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Context 4:"
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
      Height          =   375
      Index           =   24
      Left            =   6720
      TabIndex        =   79
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   23
      Left            =   6840
      TabIndex        =   78
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   22
      Left            =   6840
      TabIndex        =   77
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   21
      Left            =   6840
      TabIndex        =   76
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   20
      Left            =   6840
      TabIndex        =   75
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tonic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   19
      Left            =   2400
      TabIndex        =   52
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Phasic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   18
      Left            =   1200
      TabIndex        =   51
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CS 4:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   17
      Left            =   120
      TabIndex        =   50
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CS 3:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   16
      Left            =   120
      TabIndex        =   49
      Top             =   7920
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CS 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   15
      Left            =   120
      TabIndex        =   48
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CS 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   14
      Left            =   120
      TabIndex        =   47
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   3120
      TabIndex        =   44
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   3120
      TabIndex        =   43
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Context 2:"
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
      Height          =   375
      Index           =   11
      Left            =   2880
      TabIndex        =   42
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CS:"
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
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   39
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   9
      Left            =   360
      TabIndex        =   38
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   8
      Left            =   360
      TabIndex        =   37
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Non-CS:"
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
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   36
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   3120
      TabIndex        =   33
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   3120
      TabIndex        =   32
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Context 1:"
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
      Height          =   375
      Index           =   4
      Left            =   2880
      TabIndex        =   31
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Mossy fiber context frequencies (Hz)"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   30
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   27
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   26
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MF background frequencies (Hz)"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   25
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Non-Uniform"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   22
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "NC Homeostatic SetPoint"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   21
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PC Homeostatic SetPoint"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   20
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "SimOptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CFOptions_Click(Index As Integer)
    Select Case Index
        Case 0
            NumCF = 1
            ConnectivityFrame.Visible = False
        Case 1
            NumCF = 4
            ConnectivityFrame.Visible = True
        Case 2
            NumCF = 12
            ConnectivityFrame.Visible = True
    End Select
    
End Sub

Private Sub Command1_Click()

    MFBGROUNDFREQMIN = Val(MFText(0).Text)
    MFBGROUNDFREQMAX = Val(MFText(1).Text)
    
    MFBGROUNDFREQMIN_CS = Val(MFText(2).Text)
    MFBGROUNDFREQMAX_CS = Val(MFText(3).Text)
    
    MFCONTEXTFREQMIN = Val(MFText(4).Text)
    MFCONTEXTFREQMAX = Val(MFText(5).Text)
    
    MFCONTEXTFREQMIN2 = Val(MFText(6).Text)
    MFCONTEXTFREQMAX2 = Val(MFText(7).Text)
    
    MFPHASICFREQ_INCREMENT = Val(MFText(8).Text)
    MFPHASICFREQ_INCREMENT2 = Val(MFText(9).Text)
    MFPHASICFREQ_INCREMENT3 = Val(MFText(10).Text)
    MFPHASICFREQ_INCREMENT4 = Val(MFText(11).Text)
    
    MFTONICFREQ_INCREMENT = Val(MFText(12).Text)
    MFTONICFREQ_INCREMENT2 = Val(MFText(13).Text)
    MFTONICFREQ_INCREMENT3 = Val(MFText(14).Text)
    MFTONICFREQ_INCREMENT4 = Val(MFText(15).Text)
    
    Command1.Enabled = False
    Command2.Enabled = False
End Sub

Private Sub Command2_Click()
    MFText(0).Text = MFBGROUNDFREQMIN
    MFText(1).Text = MFBGROUNDFREQMAX
    
    MFText(2).Text = MFBGROUNDFREQMIN_CS
    MFText(3).Text = MFBGROUNDFREQMAX_CS
    
    MFText(4).Text = MFCONTEXTFREQMIN
    MFText(5).Text = MFCONTEXTFREQMAX
    
    MFText(6).Text = MFCONTEXTFREQMIN2
    MFText(7).Text = MFCONTEXTFREQMAX2
    
    MFText(8).Text = MFPHASICFREQ_INCREMENT
    MFText(9).Text = MFPHASICFREQ_INCREMENT2
    MFText(10).Text = MFPHASICFREQ_INCREMENT3
    MFText(11).Text = MFPHASICFREQ_INCREMENT4
    
    MFText(12).Text = MFTONICFREQ_INCREMENT
    MFText(13).Text = MFTONICFREQ_INCREMENT2
    MFText(14).Text = MFTONICFREQ_INCREMENT3
    MFText(15).Text = MFTONICFREQ_INCREMENT4
    
    Command1.Enabled = False
    Command2.Enabled = False
End Sub

Private Sub Form_Load()
    NumCF = 12
'    UseUBCs = 0
    
    MFText(0).Text = MFBGROUNDFREQMIN
    MFText(1).Text = MFBGROUNDFREQMAX
    
    MFText(2).Text = MFBGROUNDFREQMIN_CS
    MFText(3).Text = MFBGROUNDFREQMAX_CS
    
    MFText(4).Text = MFCONTEXTFREQMIN
    MFText(5).Text = MFCONTEXTFREQMAX
    
    MFText(6).Text = MFCONTEXTFREQMIN2
    MFText(7).Text = MFCONTEXTFREQMAX2
    
    MFText(8).Text = MFPHASICFREQ_INCREMENT
    MFText(9).Text = MFPHASICFREQ_INCREMENT2
    MFText(10).Text = MFPHASICFREQ_INCREMENT3
    MFText(11).Text = MFPHASICFREQ_INCREMENT4
    
    MFText(12).Text = MFTONICFREQ_INCREMENT
    MFText(13).Text = MFTONICFREQ_INCREMENT2
    MFText(14).Text = MFTONICFREQ_INCREMENT3
    MFText(15).Text = MFTONICFREQ_INCREMENT4
    
    Command1.Enabled = False
    Command2.Enabled = False
End Sub

Private Sub MFText_Change(Index As Integer)
    Command1.Enabled = True
    Command2.Enabled = True
End Sub

Private Sub SimOptionsButton_Click(Index As Integer)
Dim i As Integer
Dim X As Single

    If Index = 0 Then
        If NumCF = 1 Then
            Index = 1
            MAXDRUS = 15
        ElseIf NumCF = 4 Then
            If ConnectOptions(1).Value = True Then
                Index = 2
            Else
                Index = 1
            End If
            MAXDRUS = 6
        Else
            Index = 3
            MAXDRUS = 3
        End If
        cbm_main.USLabel.Caption = MAXDRUS
        
        Time_step_size = 1
        Calculate_Time_Dependent_variables
        SynaptoGenesis Index, UseUBCs
        Init_stuff
'        Unload Progress
        cbm_main.speed_menu.Enabled = True
    End If
    
    SimOptionsForm.Visible = False
    For i = 1 To PCNUMBER
        PCHomeoValue(i) = Val(Text1(0).Text)
    Next i
    cbm_main.Text1(0).Text = Text1(0).Text
    If Check1(0).Value = Checked Then
        cbm_main.Text1(0).Visible = True
        For i = 1 To PCNUMBER
            X = (Rnd() - 0.5) * (0.1 * PCHomeoValue(i))
            PCHomeoValue(i) = PCHomeoValue(i) + X
'            Debug.Print PCHomeoValue(i)
        Next i
    Else
        cbm_main.Text1(0).Visible = False
    End If
    For i = 1 To NCNUMBER
        NCHomeoValue(i) = Val(Text1(1).Text)
    Next i
    cbm_main.Text1(1).Text = Text1(1).Text
    If Check1(1).Value = Checked Then
        cbm_main.Text1(1).Visible = True
        For i = 1 To NCNUMBER
            X = (Rnd() - 0.5) * (0.1 * NCHomeoValue(i))
            NCHomeoValue(i) = NCHomeoValue(i) + X
'            Debug.Print NCHomeoValue(i)
        Next i
    Else
        cbm_main.Text1(1).Visible = False
    End If
End Sub

Private Sub UBCOptions_Click(Index As Integer)
    UseUBCs = Index
    If UseUBCs = 1 Then
        cbm_main.change_text(6).Visible = False
        cbm_main.change_text(7).Visible = False
        cbm_main.Change_button(6).Visible = False
        cbm_main.Change_button(7).Visible = False
        cbm_main.Cancel_buttons(6).Visible = False
        cbm_main.Cancel_buttons(7).Visible = False
        cbm_main.Change_stimuli(6).Visible = False
        cbm_main.Change_stimuli(7).Visible = False
        cbm_main.OnsetLabels(6).Visible = False
        cbm_main.OnsetLabels(7).Visible = False
        cbm_main.CS4.Visible = False
    Else
        cbm_main.OnsetLabels(6).Visible = True
        cbm_main.OnsetLabels(7).Visible = True
'        cbm_main.change_text(6).Visible = True
'        cbm_main.change_text(7).Visible = True
        cbm_main.Change_button(6).Visible = True
        cbm_main.Change_button(7).Visible = True
        cbm_main.Cancel_buttons(6).Visible = True
        cbm_main.Cancel_buttons(7).Visible = True
'        cbm_main.Change_stimuli(6).Visible = True
'        cbm_main.Change_stimuli(7).Visible = True
        
        cbm_main.CS4.Visible = True
    End If
End Sub

