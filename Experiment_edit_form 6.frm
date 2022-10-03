VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Experiment_edit_form 
   BackColor       =   &H00004000&
   Caption         =   "Experiment Editor"
   ClientHeight    =   9630
   ClientLeft      =   7305
   ClientTop       =   3450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   7740
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Index           =   4
      Left            =   3360
      TabIndex        =   53
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save As"
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   27
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   26
      Top             =   8760
      Width           =   1215
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   25
      Top             =   120
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   24
      Top             =   480
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   23
      Top             =   840
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   22
      Top             =   1200
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   21
      Top             =   1560
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   20
      Top             =   1920
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   6
      Left            =   600
      TabIndex        =   19
      Top             =   2280
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   7
      Left            =   600
      TabIndex        =   18
      Top             =   2640
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   8
      Left            =   600
      TabIndex        =   17
      Top             =   3000
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   9
      Left            =   600
      TabIndex        =   16
      Top             =   3360
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   10
      Left            =   600
      TabIndex        =   15
      Top             =   3720
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   11
      Left            =   600
      TabIndex        =   14
      Top             =   4080
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   12
      Left            =   600
      TabIndex        =   13
      Top             =   4440
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   13
      Left            =   600
      TabIndex        =   12
      Top             =   4800
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   14
      Left            =   600
      TabIndex        =   11
      Top             =   5160
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   15
      Left            =   600
      TabIndex        =   10
      Top             =   5520
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   16
      Left            =   600
      TabIndex        =   9
      Top             =   5880
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   17
      Left            =   600
      TabIndex        =   8
      Top             =   6240
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   18
      Left            =   600
      TabIndex        =   7
      Top             =   6600
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   19
      Left            =   600
      TabIndex        =   6
      Top             =   6960
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   20
      Left            =   600
      TabIndex        =   5
      Top             =   7320
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   21
      Left            =   600
      TabIndex        =   4
      Top             =   7680
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   22
      Left            =   600
      TabIndex        =   3
      Top             =   8040
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00004000&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   23
      Left            =   600
      TabIndex        =   2
      Top             =   8400
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   1
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   375
      Index           =   3
      Left            =   600
      TabIndex        =   0
      Top             =   8760
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   52
      Top             =   9240
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   51
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   50
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   49
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   48
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   47
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   46
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   45
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   44
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   43
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   42
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   41
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   40
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   39
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   38
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      TabIndex        =   37
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   36
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   35
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   34
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      TabIndex        =   33
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   32
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   31
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   30
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   120
      TabIndex        =   29
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   28
      Top             =   8400
      Width           =   375
   End
End
Attribute VB_Name = "Experiment_edit_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ExperimentSave(filename As String)
Dim i As Integer
Dim V As Integer
    Close #22
    Open filename For Output As #22
    For i = 0 To 23
        If Experiment_edit_form.session(i).Text <> "" Then Write #22, Experiment_edit_form.session(i).Text Else i = 23
    Next i
    Close #2
End Sub
Private Sub ExperimentLoad(filename As String)
Dim i As Integer
Dim s$
    Close #22
    i = 0
    Open filename For Input As #22
    While Not EOF(22)
        Input #22, s$
        Experiment_edit_form.session(i).Text = s$
        i = i + 1
    Wend
    Experiment_edit_form.Visible = True
    Close #22
End Sub
Private Sub Command1_Click(Index As Integer)
Dim i As Integer
    Select Case Index
        Case 0
            CD1.filename = ""
            CD1.Filter = "(*.exp)|*.exp"
            CD1.ShowOpen
            If CD1.filename <> "" Then
                ExperimentSave (CD1.filename)
                Label2.Caption = CD1.filename
            End If
        Case 1
            Experiment_edit_form.Visible = False
        Case 2
            CD1.filename = ""
            CD1.Filter = "(*.exp)|*.exp"
            CD1.ShowOpen
            If CD1.filename <> "" Then
                ExperimentLoad (CD1.filename)
                Label2.Caption = CD1.filename
            End If
        Case 3
            For i = 0 To 23
                Experiment_edit_form.session(i).Text = ""
            Next i
            Label2.Caption = ""
        Case 4
            If Label2.Caption <> "" Then
                ExperimentSave (Label2.Caption)
            End If
    End Select
End Sub

Private Sub Form_Load()
Dim i As Integer
    For i = 0 To 23
        Label1(i).Caption = i + 1
    Next i
End Sub

Private Sub Label1_Click(Index As Integer)
    If Index > 0 Then
        session(Index) = session(Index - 1)
    End If
End Sub

Private Sub session_DblClick(Index As Integer)
CD1.filename = ""
    CD1.Filter = "(*.sse)|*.sse"
    CD1.ShowOpen
    If CD1.filename <> "" Then
        Experiment_edit_form.session(Index).Text = CD1.filename
    End If
End Sub
