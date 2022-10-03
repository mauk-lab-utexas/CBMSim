VERSION 5.00
Begin VB.Form Raster_adjust 
   BackColor       =   &H00000040&
   Caption         =   "Raster adjust form"
   ClientHeight    =   2925
   ClientLeft      =   255
   ClientTop       =   6180
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   7845
   Begin VB.HScrollBar MF_Scroll 
      Height          =   375
      Left            =   120
      Max             =   70
      Min             =   1
      TabIndex        =   2
      Top             =   1200
      Value           =   1
      Width           =   7575
   End
   Begin VB.HScrollBar Weights_scroll 
      Height          =   375
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   0
      Top             =   360
      Value           =   1
      Width           =   7575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Color for MF-Nuc synaptic weights scale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Color for gr-PURK synaptic weights scale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Raster_adjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Weights_scroll.Value = Gr_weights_denom * 10
MF_Scroll.Value = MF_weights_denom * 100
End Sub

Private Sub MF_Scroll_Change()
MF_weights_denom = MF_Scroll.Value / 100
End Sub

Private Sub Weights_scroll_Change()
Gr_weights_denom = Weights_scroll.Value / 10
End Sub
