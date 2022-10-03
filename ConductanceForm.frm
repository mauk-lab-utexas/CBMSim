VERSION 5.00
Begin VB.Form ConductanceForm 
   BackColor       =   &H00000000&
   Caption         =   "Conductances"
   ClientHeight    =   7620
   ClientLeft      =   150
   ClientTop       =   4965
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   7515
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   36
      Left            =   960
      TabIndex        =   97
      Text            =   "Text1"
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   35
      Left            =   960
      TabIndex        =   94
      Text            =   "Text1"
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   315
      Index           =   23
      Left            =   1920
      TabIndex        =   93
      Top             =   6240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   315
      Index           =   22
      Left            =   2520
      TabIndex        =   92
      Top             =   6240
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   34
      Left            =   6120
      TabIndex        =   91
      Text            =   "Text1"
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   33
      Left            =   3240
      TabIndex        =   90
      Text            =   "0.001"
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   32
      Left            =   960
      TabIndex        =   87
      Text            =   "Text1"
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   315
      Index           =   21
      Left            =   1920
      TabIndex        =   86
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   315
      Index           =   20
      Left            =   2520
      TabIndex        =   85
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   31
      Left            =   6120
      TabIndex        =   84
      Text            =   "Text1"
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   30
      Left            =   3240
      TabIndex        =   83
      Text            =   "0.001"
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton ShowActivityButton 
      Caption         =   "Show Activity"
      Height          =   255
      Left            =   4800
      TabIndex        =   82
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   29
      Left            =   3240
      TabIndex        =   74
      Text            =   "0.001"
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   28
      Left            =   3240
      TabIndex        =   73
      Text            =   "0.0001"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   27
      Left            =   3240
      TabIndex        =   72
      Text            =   "0.001"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   3240
      TabIndex        =   71
      Text            =   "0.0002"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   25
      Left            =   3240
      TabIndex        =   70
      Text            =   "0.001"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   24
      Left            =   3240
      TabIndex        =   69
      Text            =   "0.001"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   3240
      TabIndex        =   68
      Text            =   "0.001"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   3240
      TabIndex        =   67
      Text            =   "0.0002"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   21
      Left            =   3240
      TabIndex        =   66
      Text            =   "0.0002"
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   3240
      TabIndex        =   65
      Text            =   "0.001"
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh Original values"
      Height          =   495
      Index           =   3
      Left            =   4200
      TabIndex        =   64
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore to Original values"
      Height          =   495
      Index           =   2
      Left            =   5880
      TabIndex        =   63
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   6120
      TabIndex        =   51
      Text            =   "0.001"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   6120
      TabIndex        =   50
      Text            =   "."
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   6120
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   6120
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   6120
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   6120
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   6120
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   6120
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   6120
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   6120
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   41
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      Enabled         =   0   'False
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   40
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   315
      Index           =   19
      Left            =   2520
      TabIndex        =   39
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   315
      Index           =   18
      Left            =   1920
      TabIndex        =   38
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   960
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   315
      Index           =   17
      Left            =   2520
      TabIndex        =   36
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   315
      Index           =   16
      Left            =   1920
      TabIndex        =   35
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   960
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   315
      Index           =   15
      Left            =   2520
      TabIndex        =   33
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   315
      Index           =   14
      Left            =   1920
      TabIndex        =   32
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   960
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   315
      Index           =   13
      Left            =   2520
      TabIndex        =   30
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   315
      Index           =   12
      Left            =   1920
      TabIndex        =   29
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   960
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   315
      Index           =   11
      Left            =   2520
      TabIndex        =   27
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   315
      Index           =   10
      Left            =   1920
      TabIndex        =   26
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   960
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   315
      Index           =   9
      Left            =   2520
      TabIndex        =   24
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   315
      Index           =   8
      Left            =   1920
      TabIndex        =   23
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   315
      Index           =   7
      Left            =   2520
      TabIndex        =   21
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   315
      Index           =   6
      Left            =   1920
      TabIndex        =   20
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   315
      Index           =   5
      Left            =   2520
      TabIndex        =   18
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   315
      Index           =   4
      Left            =   1920
      TabIndex        =   17
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   315
      Index           =   3
      Left            =   2520
      TabIndex        =   15
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   315
      Index           =   2
      Left            =   1920
      TabIndex        =   14
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   315
      Index           =   1
      Left            =   2520
      TabIndex        =   12
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   11
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "BC Tonic"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   24
      Left            =   120
      TabIndex        =   98
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "LTDuration"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   96
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "LTDuration"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   22
      Left            =   5280
      TabIndex        =   95
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "LTD offset"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   89
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "LTD offset"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   20
      Left            =   5280
      TabIndex        =   88
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   81
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Golgi activity"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   80
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   79
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "granule activity"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   78
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   77
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Purkinje activity"
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   76
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label GotoGoLabel 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   75
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Original Values"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   62
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "MF to gr"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   19
      Left            =   5280
      TabIndex        =   61
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "MF to Gol"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   18
      Left            =   5280
      TabIndex        =   60
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "gr to Gol"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   17
      Left            =   5280
      TabIndex        =   59
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Gol to gr"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   16
      Left            =   5280
      TabIndex        =   58
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Gol to Gol"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   15
      Left            =   5280
      TabIndex        =   57
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "gr to BC"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   14
      Left            =   5280
      TabIndex        =   56
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "gr to SC"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   13
      Left            =   5280
      TabIndex        =   55
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "BC to PC"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   12
      Left            =   5280
      TabIndex        =   54
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "SC to PC"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   11
      Left            =   5280
      TabIndex        =   53
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "PC to BC"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   10
      Left            =   5280
      TabIndex        =   52
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "PC to BC"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "SC to PC"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "BC to PC"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "gr to SC"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "gr to BC"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Gol to Gol"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Gol to gr"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "gr to Gol"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "MF to Gol"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "MF to gr"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Menu CFViewMenu 
      Caption         =   "View"
      Begin VB.Menu CFormModeMenu 
         Caption         =   "None"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu CFormModeMenu 
         Caption         =   "CF inputs"
         Index           =   1
      End
      Begin VB.Menu CFormModeMenu 
         Caption         =   "Total STP"
         Index           =   2
      End
      Begin VB.Menu CFormModeMenu 
         Caption         =   "Average STP of active synapses"
         Index           =   3
      End
   End
End
Attribute VB_Name = "ConductanceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CFormModeMenu_Click(Index As Integer)
Dim i As Integer
    ConductanceFormMode = Index
    For i = 0 To 3
        CFormModeMenu(i).Checked = False
    Next i
    CFormModeMenu(Index).Checked = True
End Sub

Private Sub Command2_Click(Index As Integer)
    If Index = 3 Then
        Call RefreshOriginal
    ElseIf Index = 2 Then
        Call RestoreOriginal
    ElseIf Index = 1 Then
        Call Updateg
    ElseIf Index = 0 Then
        Call Refreshg
    End If
End Sub

Private Sub Form_Load()
    ConductanceForm.ScaleLeft = 0
    ConductanceForm.ScaleWidth = 5000
    ConductanceForm.ScaleTop = 4#
    ConductanceForm.ScaleHeight = -4#
    Call Refreshg
    Call RefreshOriginal
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    If Button = 1 Then
        CFgDisplayMultiplier = CFgDisplayMultiplier * 2#
    Else
        CFgDisplayMultiplier = CFgDisplayMultiplier * 0.5
    End If
End Sub

Public Sub Refreshg()

    Text1(0) = gEconstGr
    Text1(1) = gEconstGoMF
    Text1(2) = gEconstGoGr
    Text1(3) = gIconstGr
    Text1(4) = GGconst
    
    Text1(5) = GrBCWeights
    Text1(6) = WEIGHTSTELL
    Text1(7) = GCONSTBCPC
    Text1(8) = GCONSTStellPC
    Text1(9) = PCBCWeights
    
    Text1(32) = LTD_OFFSET
    Text1(35) = LTDURATION
    Text1(36) = TonicgI_Basket
End Sub
Public Sub Updateg()
Dim X As Integer

    gEconstGr = Text1(0)
    gEconstGoMF = Text1(1)
    gEconstGoGr = Text1(2)
    
    For X = 1 To GoX * GoY
        Gol(X).g_varMF = gEconstGoMF * MFtoGo
        Gol(X).g_varGr = gEconstGoGr * GRtoGo
    Next X
    
    gIconstGr = Text1(3)
    GGconst = Text1(4)
    
    
    GrBCWeights = Text1(5)
    WEIGHTSTELL = Text1(6)
    GCONSTBCPC = Text1(7)
    GCONSTStellPC = Text1(8)
    PCBCWeights = Text1(9)
    
    TonicgI_Basket = Text1(36)
    
    
    LTD_OFFSET = Text1(32)
    LTDURATION = Text1(35)
    DELMINUS = DELPLUS - (DELPLUS / (LTDURATION * 0.001))
End Sub
Public Sub RefreshOriginal()
Dim i As Integer

    For i = 0 To 9
        Text1(i + 10) = Text1(i)
    Next i
    Text1(31) = Text1(32)
    Text1(34) = Text1(35)
    
End Sub
Public Sub RestoreOriginal()
Dim i As Integer

    For i = 0 To 9
        Text1(i) = Text1(i + 10)
    Next i
    Text1(32) = Text1(31)
    Text1(35) = Text1(34)
    
End Sub

