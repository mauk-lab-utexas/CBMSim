VERSION 5.00
Begin VB.Form activity_window 
   BackColor       =   &H00004000&
   Caption         =   "Total activity window"
   ClientHeight    =   11310
   ClientLeft      =   75
   ClientTop       =   6675
   ClientWidth     =   15015
   ForeColor       =   &H00008080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8
   ScaleLeft       =   1
   ScaleMode       =   0  'User
   ScaleWidth      =   5002
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "gNuc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   330
      Index           =   7
      Left            =   0
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "gPurk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Top             =   5775
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "basket"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   4830
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "mossy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   330
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   3990
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Golgi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   330
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   2940
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "nucleus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   330
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   2205
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purkinje"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   1470
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "granule"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   735
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu show_activity_menu 
      Caption         =   "&granule cells"
      Index           =   0
   End
   Begin VB.Menu show_activity_menu 
      Caption         =   "&Purkinje cells"
      Index           =   1
   End
   Begin VB.Menu show_activity_menu 
      Caption         =   "&Nucleus cells"
      Index           =   2
   End
   Begin VB.Menu show_activity_menu 
      Caption         =   "G&olgi cells"
      Index           =   3
   End
   Begin VB.Menu show_activity_menu 
      Caption         =   "mossy fibers"
      Index           =   4
   End
   Begin VB.Menu show_activity_menu 
      Caption         =   "basket cells"
      Index           =   5
   End
   Begin VB.Menu show_activity_menu 
      Caption         =   "Purkinje &conductances"
      Index           =   6
   End
   Begin VB.Menu show_activity_menu 
      Caption         =   "Nucleus conductances"
      Index           =   7
   End
   Begin VB.Menu change_colors_menus 
      Caption         =   "change colors"
   End
End
Attribute VB_Name = "activity_window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim i As Integer
    activity_window.AutoRedraw = True
    For i = 1 To 8
        order_of_labels(i) = i - 1
        activity_window.Line (1, i - 1)-(5000, i - 1)
    Next i
    DoEvents
    activity_window.AutoRedraw = False
End Sub

Private Sub show_activity_menu_Click(Index As Integer)
    If Label1(Index).Visible = True Then
        Label1(Index).Visible = False
    Else
        Label1(Index).Top = order_of_labels(Index + 1)
        Label1(Index).Visible = True
    End If
    
End Sub
