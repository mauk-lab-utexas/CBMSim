VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form session_edit_form 
   BackColor       =   &H00000040&
   Caption         =   "Session Editor"
   ClientHeight    =   9780
   ClientLeft      =   7290
   ClientTop       =   3000
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   7665
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Index           =   4
      Left            =   2640
      TabIndex        =   57
      Top             =   8760
      Width           =   975
   End
   Begin VB.CheckBox STPFades 
      BackColor       =   &H00000040&
      Caption         =   "STP reverses"
      ForeColor       =   &H00FF80FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   52
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000040&
      Height          =   735
      Left            =   6120
      TabIndex        =   53
      Top             =   9000
      Width           =   1455
      Begin VB.OptionButton ContextOption 
         BackColor       =   &H00000040&
         Caption         =   "Context 2"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton ContextOption 
         BackColor       =   &H00000040&
         Caption         =   "Context 1"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   51
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   50
      Top             =   8760
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   23
      Left            =   600
      TabIndex        =   49
      Top             =   8400
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   22
      Left            =   600
      TabIndex        =   48
      Top             =   8040
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   21
      Left            =   600
      TabIndex        =   47
      Top             =   7680
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   20
      Left            =   600
      TabIndex        =   46
      Top             =   7320
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   19
      Left            =   600
      TabIndex        =   45
      Top             =   6960
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   18
      Left            =   600
      TabIndex        =   44
      Top             =   6600
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   17
      Left            =   600
      TabIndex        =   43
      Top             =   6240
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   16
      Left            =   600
      TabIndex        =   42
      Top             =   5880
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   15
      Left            =   600
      TabIndex        =   41
      Top             =   5520
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   14
      Left            =   600
      TabIndex        =   40
      Top             =   5160
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   13
      Left            =   600
      TabIndex        =   39
      Top             =   4800
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   12
      Left            =   600
      TabIndex        =   38
      Top             =   4440
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   11
      Left            =   600
      TabIndex        =   37
      Top             =   4080
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   10
      Left            =   600
      TabIndex        =   36
      Top             =   3720
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   9
      Left            =   600
      TabIndex        =   35
      Top             =   3360
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   8
      Left            =   600
      TabIndex        =   34
      Top             =   3000
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   7
      Left            =   600
      TabIndex        =   33
      Top             =   2640
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   6
      Left            =   600
      TabIndex        =   32
      Top             =   2280
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   31
      Top             =   1920
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   30
      Top             =   1560
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   29
      Top             =   1200
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   28
      Top             =   840
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   27
      Top             =   480
      Width           =   6975
   End
   Begin VB.TextBox session 
      BackColor       =   &H00000040&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   1
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save As"
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   0
      Top             =   8760
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   56
      Top             =   9240
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   26
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   25
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   24
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   23
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   22
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   21
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   20
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   19
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   18
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   17
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   16
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   15
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   14
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   13
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   12
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   11
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   10
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   9
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   8
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   7
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   6
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
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
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "session_edit_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub SessionSave(filename As String)
Dim i As Integer
Dim v As Integer
    Close #22
    Open filename For Output As #22
    For i = 0 To 23
        If session_edit_form.session(i).Text <> "" Then Write #22, session_edit_form.session(i).Text Else i = 23
    Next i
    If STPFades.Value = Checked Then
        Write #22, "yes"
    Else
        Write #22, "no"
    End If
    Write #22, CurrentContext
    Close #22
End Sub
Private Sub SessionLoad(filename As String)
Dim i As Integer
Dim s$
    Close #22
    i = 0
    Open filename For Input As #22
    While Not EOF(22)
        Input #22, s$
        If s$ = "yes" Then
            STPFades.Value = Checked
            Input #22, CurrentContext
        ElseIf s$ = "no" Then
            STPFades.Value = Unchecked
            Input #22, CurrentContext
        Else
            session_edit_form.session(i).Text = s$
        End If
        i = i + 1
    Wend
    session_edit_form.Visible = True
    If CurrentContext = 1 Then
        ContextOption(0).Value = True
    Else
        ContextOption(1).Value = True
    End If
    Close #22
End Sub
Private Sub block_DblClick(Index As Integer)

End Sub

Private Sub Command1_Click(Index As Integer)
Dim i As Integer
    Select Case Index
        Case 0
            CD1.filename = ""
            CD1.Filter = "(*.sse)|*.sse"
            CD1.ShowOpen
            If CD1.filename <> "" Then
                SessionSave (CD1.filename)
                Label2.Caption = CD1.filename
            End If
        Case 1
            session_edit_form.Visible = False
        Case 2
            CD1.filename = ""
            CD1.Filter = "(*.sse)|*.sse"
            CD1.ShowOpen
            If CD1.filename <> "" Then
                SessionLoad (CD1.filename)
                Label2.Caption = CD1.filename
            End If
        Case 3
            For i = 0 To 23
                session_edit_form.session(i).Text = ""
            Next i
            Label2.Caption = ""
        Case 4
            If Label2.Caption <> "" Then
                SessionSave (Label2.Caption)
            End If
    End Select
    
End Sub

Private Sub ContextOption_Click(Index As Integer)
    CurrentContext = Index + 1
End Sub

Private Sub Form_Load()
Dim i As Integer
    For i = 0 To 23
        Label1(i).Caption = i + 1
    Next i
    CurrentContext = 1
End Sub

Private Sub Label1_Click(Index As Integer)
    If Index > 0 Then
        session(Index) = session(Index - 1)
    End If
End Sub

Private Sub session_DblClick(Index As Integer)
    CD1.filename = ""
    CD1.Filter = "(*.sbl)|*.sbl"
    CD1.ShowOpen
    If CD1.filename <> "" Then
        session_edit_form.session(Index).Text = CD1.filename
    End If
End Sub

