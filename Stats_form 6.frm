VERSION 5.00
Begin VB.Form Stats_form 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Statistics window"
   ClientHeight    =   8070
   ClientLeft      =   3270
   ClientTop       =   4830
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   12735
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2892
      Left            =   11400
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      Begin VB.OptionButton Cell_option 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nuc to CF"
         Height          =   195
         Index           =   9
         Left            =   0
         TabIndex        =   27
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton Cell_option 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Purk to Nuc"
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   25
         Top             =   1920
         Width           =   1575
      End
      Begin VB.OptionButton Cell_option 
         BackColor       =   &H00FFFFFF&
         Caption         =   "None"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Cell_option 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Climbing fibers"
         Height          =   195
         Index           =   7
         Left            =   0
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton Cell_option 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Purkinje cells"
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton Cell_option 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nucleus cells"
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   5
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Cell_option 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gogli cells"
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Cell_option 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Basket cells"
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Cell_option 
         BackColor       =   &H00FFFFFF&
         Caption         =   "granule cells"
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Cell_option 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mossy fibers"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   252
      Left            =   0
      TabIndex        =   28
      Top             =   240
      Width           =   732
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   252
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   732
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   15
      Left            =   10200
      TabIndex        =   24
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   14
      Left            =   10200
      TabIndex        =   23
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   13
      Left            =   10200
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   12
      Left            =   10200
      TabIndex        =   21
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   11
      Left            =   10200
      TabIndex        =   20
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   10
      Left            =   10200
      TabIndex        =   19
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   9
      Left            =   10200
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   8
      Left            =   10200
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   7
      Left            =   10200
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   6
      Left            =   10200
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   5
      Left            =   10200
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   4
      Left            =   10200
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   3
      Left            =   10200
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   2
      Left            =   10200
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   1
      Left            =   10200
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Fire_label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      Height          =   255
      Index           =   0
      Left            =   10200
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Stats_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Stats_cell_menu_Click(Index As Integer)

End Sub

Private Sub Cell_option_Click(Index As Integer)
Dim x As Integer
Dim cell As Integer
Dim num As Integer
Dim max_freq As Integer
Dim max_fire As Single
Dim avg_fire As Single
Dim ss As Double
Dim s As Double
Dim stats(SYNUMBER) As Single
Dim f As String
Dim bin As Integer
Dim bin_size As Integer
Dim c As Integer
Dim bins(200) As Integer
Dim sd As Double
Dim max As Single
Dim avg As Single
Dim i, j As Integer

Erase stats
ss = 0
s = 0
avg_fire = 0

    Cls
    Select Case Index
        Case 0
            num = 0
        Case 1
            num = MFNUMBER
            For cell = 1 To num
                For x = 1 To 1000
                    stats(cell) = stats(cell) + MF_histo(cell, x)
                Next x
                stats(cell) = stats(cell) / (5 * HistoDivisor)
                s = s + stats(cell)
                ss = ss + (stats(cell) * stats(cell))
                avg_fire = avg_fire + stats(cell)
                If stats(cell) > max_fire Then max_fire = stats(cell)
            Next cell
            avg_fire = avg_fire / num
            f = "Granule cells  " + Str$(avg_fire) + Str$(max_fire)
            
        Case 2
            num = GrX * GrY
            For cell = 1 To num
                For x = 1 To 1000
                    stats(cell) = stats(cell) + gr_histo(cell, x)
                Next x
                stats(cell) = stats(cell) / (5 * HistoDivisor)
                s = s + stats(cell)
                ss = ss + (stats(cell) * stats(cell))
                avg_fire = avg_fire + stats(cell)
                If stats(cell) > max_fire Then max_fire = stats(cell)
            Next cell
            avg_fire = avg_fire / num
            f = "Granule cells  " + Str$(avg_fire) + Str$(max_fire)
            
            
        Case 3
            num = GoX * GoY
            For cell = 1 To num
                For x = 1 To 1000
                    stats(cell) = stats(cell) + Go_histo(cell, x)
                Next x
                stats(cell) = stats(cell) / (5 * HistoDivisor)
                s = s + stats(cell)
                ss = ss + (stats(cell) * stats(cell))
                avg_fire = avg_fire + stats(cell)
                If stats(cell) > max_fire Then max_fire = stats(cell)
            Next cell
            avg_fire = avg_fire / num
            f = "Golgi cells  " + Str$(avg_fire) + Str$(max_fire)
            
        Case 4
            num = STELLATENUMBER
            For cell = 1 To num
                For x = 1 To 1000
                    stats(cell) = stats(cell) + Stellate_histo(cell, x)
                Next x
                stats(cell) = stats(cell) / (5 * HistoDivisor)
                s = s + stats(cell)
                ss = ss + (stats(cell) * stats(cell))
                avg_fire = avg_fire + stats(cell)
                If stats(cell) > max_fire Then max_fire = stats(cell)
            Next cell
            avg_fire = avg_fire / num
            f = "Basket cells  " + Str$(avg_fire) + Str$(max_fire)
            
        Case 5
            num = PCNUMBER
            For cell = 1 To num
                For x = 1 To 1000
                    stats(cell) = stats(cell) + Purk_histo(cell, x)
                Next x
                stats(cell) = stats(cell) / (5 * HistoDivisor)
                s = s + stats(cell)
                ss = ss + (stats(cell) * stats(cell))
                avg_fire = avg_fire + stats(cell)
                If stats(cell) > max_fire Then max_fire = stats(cell)
            Next cell
            avg_fire = avg_fire / num
            f = "Purkine cells  " + Str$(avg_fire) + Str$(max_fire)
            
        Case 6
            num = NCNUMBER
            For cell = 1 To num
                For x = 1 To 1000
                    stats(cell) = stats(cell) + Nuc_histo(cell, x)
                Next x
                stats(cell) = stats(cell) / (5 * HistoDivisor)
                s = s + stats(cell)
                ss = ss + (stats(cell) * stats(cell))
                avg_fire = avg_fire + stats(cell)
                If stats(cell) > max_fire Then max_fire = stats(cell)
            Next cell
            avg_fire = avg_fire / num
            f = "Nucleus cells  " + Str$(avg_fire) + Str$(max_fire)
            
        Case 7
            num = 1
        Case 8
            Stats_form.ScaleWidth = PCNUMBER + 4
            Stats_form.DrawWidth = 8
            Stats_form.Cls
            max = 0
            avg = 0
'            For i = 1 To PCNUMBER
'                If Pc(i).gPURKtoNUC > max Then max = Pc(i).gPURKtoNUC
'                avg = avg + Pc(i).gPURKtoNUC
'            Next i
            avg = avg / PCNUMBER
            
            Stats_form.ScaleTop = max * 1.2
            Stats_form.ScaleHeight = max * -1.2
            Stats_form.Label1.Caption = max
            Stats_form.Label2.Caption = avg
'            For i = 1 To PCNUMBER
'                Stats_form.Line (i, 0)-(i, Pc(i).gPURKtoNUC)
'            Next i
            Stats_form.DrawWidth = 1
       Case 9
            Stats_form.ScaleWidth = NCNUMBER + 4
            Stats_form.DrawWidth = 8
            Stats_form.Cls
            max = 0
            avg = 0
            For i = 1 To NCNUMBER
                For j = 1 To NumCF
                    If Nc(i).gNUCtoCF(j) > max Then max = Nc(i).gNUCtoCF(j)
                    avg = avg + Nc(i).gNUCtoCF(j)
                Next j
            Next i
            avg = avg / (NCNUMBER * NumCF)
            
            Stats_form.ScaleTop = max * 1.2
            Stats_form.ScaleHeight = max * -1.2
            Stats_form.Label1.Caption = max
            Stats_form.Label2.Caption = avg
            For i = 1 To NCNUMBER
                Stats_form.Line (i, 0)-(i, Nc(i).gNUCtoCF(1))
            Next i
            Stats_form.DrawWidth = 1
    End Select
    
    If Index > 0 And Index < 8 Then
     sd = Sqr((ss - ((s * s) / num)) / (num - 1))
     f = f + "+/-  " + Str$(sd)
     Stats_form.Caption = f
     Stats_form.ScaleLeft = -10
     Stats_form.ScaleWidth = 110
    
     bin_size = 1
     c = 0
     For bin = 0 To 100 Step bin_size
         c = c + 1
         bins(c) = 0
         For x = 1 To num
             If stats(x) <= bin And stats(x) > bin - bin_size Then bins(c) = bins(c) + 1
         Next x
         If bins(c) > max_freq Then max_freq = bins(c)
     Next bin
     Stats_form.ScaleTop = max_freq
     Stats_form.ScaleHeight = -max_freq - (max_freq * 0.1)
     c = 0
     For bin = 0 To 100 Step bin_size
         c = c + 1
         Stats_form.Line (bin, 0)-(bin + (bin_size * 0.9), bins(c)), vbBlack, BF
         If bin Mod 10 = 0 Then
             Fire_label(c / 10).Visible = True
             Fire_label(c / 10).Top = -0.02 * max_freq
             Fire_label(c / 10).Left = bin
             Fire_label(c / 10).Caption = bin
         End If
     Next bin
    End If
End Sub

