VERSION 5.00
Begin VB.Form histo_info_form 
   BackColor       =   &H00000080&
   Caption         =   "Histogram information"
   ClientHeight    =   2790
   ClientLeft      =   18465
   ClientTop       =   4380
   ClientWidth     =   5415
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   5415
   Visible         =   0   'False
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000080&
      Caption         =   "Show diagnostics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3885
      TabIndex        =   33
      Top             =   2310
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   105
      TabIndex        =   32
      Text            =   "1000"
      Top             =   2415
      Width           =   645
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   2115
      Left            =   3255
      TabIndex        =   22
      Top             =   735
      Width           =   540
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000080&
         Caption         =   "Option1"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   30
         Top             =   420
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000080&
         Caption         =   "Option1"
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   29
         Top             =   630
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000080&
         Caption         =   "Option1"
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   28
         Top             =   840
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000080&
         Caption         =   "Option1"
         Height          =   225
         Index           =   4
         Left            =   210
         TabIndex        =   27
         Top             =   1050
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000080&
         Caption         =   "Option1"
         Height          =   225
         Index           =   5
         Left            =   210
         TabIndex        =   26
         Top             =   1260
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000080&
         Caption         =   "Option1"
         Height          =   225
         Index           =   6
         Left            =   210
         TabIndex        =   25
         Top             =   1470
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000080&
         Caption         =   "Option1"
         Height          =   225
         Index           =   7
         Left            =   210
         TabIndex        =   24
         Top             =   1680
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000080&
         Caption         =   "Option1"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   23
         Top             =   210
         Width           =   225
      End
   End
   Begin VB.TextBox cell_num_text 
      Height          =   330
      Left            =   3465
      TabIndex        =   12
      Text            =   "1"
      Top             =   210
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   105
      TabIndex        =   11
      Text            =   "10000"
      Top             =   105
      Width           =   645
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   435
      Index           =   1
      Left            =   4410
      TabIndex        =   9
      Top             =   630
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   435
      Index           =   0
      Left            =   4410
      TabIndex        =   8
      Top             =   105
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   105
      TabIndex        =   7
      Text            =   "110"
      Top             =   1680
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   105
      TabIndex        =   6
      Text            =   "400"
      Top             =   1365
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Text            =   "1000"
      Top             =   945
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Text            =   "1"
      Top             =   630
      Width           =   645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Threshold"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   6
      Left            =   945
      TabIndex        =   31
      Top             =   2415
      Width           =   1380
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "resp"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   7
      Left            =   2940
      TabIndex        =   21
      Top             =   2415
      Width           =   330
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "bsk"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   6
      Left            =   2940
      TabIndex        =   20
      Top             =   2205
      Width           =   330
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pkj"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   2940
      TabIndex        =   19
      Top             =   1995
      Width           =   330
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nuc"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   2940
      TabIndex        =   18
      Top             =   1785
      Width           =   330
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "cf"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   2940
      TabIndex        =   17
      Top             =   1575
      Width           =   330
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "MF"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   2940
      TabIndex        =   16
      Top             =   1365
      Width           =   330
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Go"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   2940
      TabIndex        =   15
      Top             =   1155
      Width           =   330
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "gr"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   2940
      TabIndex        =   14
      Top             =   945
      Width           =   330
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1st cell"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   2625
      TabIndex        =   13
      Top             =   315
      Width           =   1380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of cells"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   840
      TabIndex        =   10
      Top             =   105
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CS duration "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   1680
      Width           =   1380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CS onset "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   1365
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stop time bin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   945
      Width           =   1380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start time  bin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   630
      Width           =   1485
   End
End
Attribute VB_Name = "histo_info_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Dim i As Integer
    If Index = 0 Then
        If Histo_form.show_histo_menu(0).Checked = True Then
            draw_histos 0, Val(histo_info_form.cell_num_text), 50, Val(histo_info_form.Text1(4)), 10, 5, Val(histo_info_form.Text1(0)), Val(histo_info_form.Text1(1)), 1000, Val(histo_info_form.Text1(2)), Val(histo_info_form.Text1(3)), max
        End If
    ElseIf Index = 1 Then
        
    End If
    
End Sub

