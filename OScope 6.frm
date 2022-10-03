VERSION 5.00
Begin VB.Form OScope 
   BackColor       =   &H00404040&
   Caption         =   "Oscilloscope"
   ClientHeight    =   9555
   ClientLeft      =   735
   ClientTop       =   7950
   ClientWidth     =   14490
   LinkTopic       =   "Form1"
   ScaleHeight     =   -80
   ScaleMode       =   0  'User
   ScaleWidth      =   14490
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Screen"
      Height          =   495
      Left            =   3420
      TabIndex        =   20
      Top             =   6315
      Width           =   1140
   End
   Begin VB.Frame CellFrame 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Mode"
      ForeColor       =   &H00800000&
      Height          =   4485
      Left            =   2430
      TabIndex        =   7
      Top             =   780
      Width           =   1905
      Begin VB.OptionButton CellOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nucleus cell"
         Height          =   240
         HelpContextID   =   1
         Index           =   11
         Left            =   120
         TabIndex        =   24
         Top             =   3540
         Width           =   1545
      End
      Begin VB.OptionButton CellOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Purkinje cell"
         Height          =   240
         HelpContextID   =   1
         Index           =   10
         Left            =   150
         TabIndex        =   23
         Top             =   3255
         Width           =   1545
      End
      Begin VB.OptionButton CellOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Basket cell"
         Height          =   240
         HelpContextID   =   1
         Index           =   9
         Left            =   150
         TabIndex        =   22
         Top             =   2940
         Width           =   1545
      End
      Begin VB.OptionButton CellOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Golgi cell"
         Height          =   240
         HelpContextID   =   1
         Index           =   8
         Left            =   135
         TabIndex        =   21
         Top             =   2640
         Width           =   1545
      End
      Begin VB.OptionButton CellOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PC to Nucleus"
         Height          =   240
         HelpContextID   =   1
         Index           =   7
         Left            =   135
         TabIndex        =   18
         Top             =   2340
         Width           =   1545
      End
      Begin VB.OptionButton CellOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MF to Nucleus"
         Height          =   240
         HelpContextID   =   1
         Index           =   6
         Left            =   135
         TabIndex        =   17
         Top             =   2040
         Width           =   1545
      End
      Begin VB.OptionButton CellOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nucleus to CF"
         Height          =   240
         HelpContextID   =   1
         Index           =   5
         Left            =   135
         TabIndex        =   16
         Top             =   1755
         Width           =   1545
      End
      Begin VB.OptionButton CellOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "gran to Golgi"
         Height          =   240
         HelpContextID   =   1
         Index           =   4
         Left            =   135
         TabIndex        =   15
         Top             =   1470
         Width           =   1545
      End
      Begin VB.OptionButton CellOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MF to Golgi"
         Height          =   240
         HelpContextID   =   1
         Index           =   3
         Left            =   135
         TabIndex        =   14
         Top             =   1170
         Width           =   1545
      End
      Begin VB.OptionButton CellOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Golgi to granule"
         Height          =   240
         HelpContextID   =   1
         Index           =   2
         Left            =   150
         TabIndex        =   13
         Top             =   870
         Width           =   1545
      End
      Begin VB.OptionButton CellOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "granule cell"
         Height          =   240
         HelpContextID   =   1
         Index           =   1
         Left            =   150
         TabIndex        =   12
         Top             =   585
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.TextBox CellNumber 
         Height          =   300
         Left            =   1065
         TabIndex        =   11
         Text            =   "1"
         Top             =   4095
         Width           =   795
      End
      Begin VB.OptionButton CellOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mossy fiber"
         Height          =   240
         HelpContextID   =   1
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Width           =   1560
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   "Cell Type"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   495
         TabIndex        =   9
         Top             =   45
         Width           =   750
      End
   End
   Begin VB.CommandButton HoldNext 
      Caption         =   "NEXT-HOLD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   2
      Left            =   3060
      TabIndex        =   6
      Top             =   150
      Width           =   1410
   End
   Begin VB.CommandButton HoldNext 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   1
      Left            =   1620
      TabIndex        =   5
      Top             =   150
      Width           =   1395
   End
   Begin VB.CommandButton HoldNext 
      Caption         =   "HOLD"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   0
      Left            =   165
      TabIndex        =   4
      Top             =   135
      Width           =   1380
   End
   Begin VB.Frame ModeFrame 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Mode"
      ForeColor       =   &H00800000&
      Height          =   2625
      Left            =   210
      TabIndex        =   1
      Top             =   810
      Width           =   1905
      Begin VB.OptionButton ModeOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Voltage ramp"
         Height          =   240
         HelpContextID   =   1
         Index           =   2
         Left            =   180
         TabIndex        =   19
         Top             =   885
         Width           =   1620
      End
      Begin VB.OptionButton ModeOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Current histogram"
         Height          =   240
         HelpContextID   =   1
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   585
         Width           =   1620
      End
      Begin VB.OptionButton ModeOption 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cell one shot"
         Height          =   240
         HelpContextID   =   1
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1620
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "Mode"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   645
         TabIndex        =   2
         Top             =   45
         Width           =   465
      End
   End
   Begin VB.PictureBox Tube 
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFF80&
      Height          =   6855
      Left            =   4605
      ScaleHeight     =   6795
      ScaleMode       =   0  'User
      ScaleWidth      =   1000
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.Line Hgrid 
         BorderColor     =   &H00800000&
         Index           =   8
         X1              =   710.567
         X2              =   906.585
         Y1              =   6240
         Y2              =   6600
      End
      Begin VB.Line Hgrid 
         BorderColor     =   &H00800000&
         Index           =   7
         X1              =   698.315
         X2              =   894.334
         Y1              =   5880
         Y2              =   6240
      End
      Begin VB.Line Hgrid 
         BorderColor     =   &H00800000&
         Index           =   6
         X1              =   686.064
         X2              =   882.083
         Y1              =   5520
         Y2              =   5880
      End
      Begin VB.Line Hgrid 
         BorderColor     =   &H00800000&
         Index           =   5
         X1              =   698.315
         X2              =   894.334
         Y1              =   5160
         Y2              =   5520
      End
      Begin VB.Line Hgrid 
         BorderColor     =   &H00C00000&
         Index           =   4
         X1              =   686.064
         X2              =   882.083
         Y1              =   4800
         Y2              =   5160
      End
      Begin VB.Line Hgrid 
         BorderColor     =   &H00800000&
         Index           =   3
         X1              =   698.315
         X2              =   894.334
         Y1              =   4560
         Y2              =   4920
      End
      Begin VB.Line Hgrid 
         BorderColor     =   &H00800000&
         Index           =   2
         X1              =   686.064
         X2              =   882.083
         Y1              =   4200
         Y2              =   4560
      End
      Begin VB.Line Hgrid 
         BorderColor     =   &H00800000&
         Index           =   1
         X1              =   686.064
         X2              =   882.083
         Y1              =   3960
         Y2              =   4320
      End
      Begin VB.Line Hgrid 
         BorderColor     =   &H00800000&
         Index           =   0
         X1              =   673.813
         X2              =   869.832
         Y1              =   3600
         Y2              =   3960
      End
      Begin VB.Line Vgrid 
         BorderColor     =   &H00800000&
         Index           =   8
         X1              =   514.548
         X2              =   551.302
         Y1              =   1320
         Y2              =   3360
      End
      Begin VB.Line Vgrid 
         BorderColor     =   &H00800000&
         Index           =   7
         X1              =   453.292
         X2              =   514.548
         Y1              =   1320
         Y2              =   3360
      End
      Begin VB.Line Vgrid 
         BorderColor     =   &H00800000&
         Index           =   6
         X1              =   416.539
         X2              =   477.795
         Y1              =   1320
         Y2              =   3360
      End
      Begin VB.Line Vgrid 
         BorderColor     =   &H00800000&
         Index           =   5
         X1              =   367.534
         X2              =   428.79
         Y1              =   1320
         Y2              =   3360
      End
      Begin VB.Line Vgrid 
         BorderColor     =   &H00800000&
         Index           =   4
         X1              =   330.781
         X2              =   392.037
         Y1              =   1320
         Y2              =   3360
      End
      Begin VB.Line Vgrid 
         BorderColor     =   &H00800000&
         Index           =   3
         X1              =   281.776
         X2              =   343.032
         Y1              =   1440
         Y2              =   3480
      End
      Begin VB.Line Vgrid 
         BorderColor     =   &H00800000&
         Index           =   2
         X1              =   245.023
         X2              =   306.279
         Y1              =   1440
         Y2              =   3480
      End
      Begin VB.Line Vgrid 
         BorderColor     =   &H00800000&
         Index           =   1
         X1              =   208.27
         X2              =   208.27
         Y1              =   1080
         Y2              =   3120
      End
      Begin VB.Line Vgrid 
         BorderColor     =   &H00800000&
         Index           =   0
         X1              =   122.511
         X2              =   122.511
         Y1              =   1080
         Y2              =   3120
      End
   End
End
Attribute VB_Name = "OScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SM(15, 1000) As Single
Dim SM_pointer As Integer



Private Sub Command1_Click()
    Tube.Cls
End Sub

Private Sub Form_Load()
Dim X As Integer
    For X = 0 To 8
        Vgrid(X).X1 = ((X + 1) * 0.1) * Tube.ScaleWidth
        Vgrid(X).X2 = Vgrid(X).X1
        Vgrid(X).y1 = Tube.ScaleTop
        Vgrid(X).Y2 = Tube.ScaleTop + Tube.ScaleHeight
        Hgrid(X).y1 = Tube.ScaleTop + (0.1 * (X + 1) * Tube.ScaleHeight)
        Hgrid(X).Y2 = Hgrid(X).y1
        Hgrid(X).X1 = 0
        Hgrid(X).X2 = Tube.ScaleWidth
    Next X
    SM_pointer = 0
End Sub

Private Sub HoldNext_Click(Index As Integer)
Dim i As Integer
    For i = 0 To 2
        HoldNext(i).Enabled = True
        If Index = 2 Then
            HoldNext(1).Enabled = False
        End If
    Next i
    HoldNext(Index).Enabled = False
    
    If Index = 1 Then
        If ModeOption(0).Value = True Then  'One shot
            If CellOption(0).Value = True Then 'MF
                do_oneshot 0
            ElseIf CellOption(1).Value = True Then 'gr
                do_oneshot 1
            ElseIf CellOption(2).Value = True Then 'Golgi to granule
                do_oneshot 2
            ElseIf CellOption(3).Value = True Then  'MF to Golgi
                do_oneshot 3
            ElseIf CellOption(4).Value = True Then 'gran to Golgi
                do_oneshot 4
            ElseIf CellOption(5).Value = True Then 'Nucleus to CF
                do_oneshot 5
            ElseIf CellOption(6).Value = True Then 'MF to Nucleus
                do_oneshot 6
            ElseIf CellOption(7).Value = True Then 'Purkinje to Nucleus
                do_oneshot 7
            End If
        ElseIf ModeOption(1).Value = True Then 'current histogram
            If CellOption(0).Value = True Then 'MF
                do_CurrentHisto 0
            ElseIf CellOption(1).Value = True Then  'granule
                do_CurrentHisto 1
            ElseIf CellOption(2).Value = True Then  'Golgi
                do_CurrentHisto 2
            End If
        ElseIf ModeOption(2).Value = True Then  'Voltage ramp
            If CellOption(1).Value = True Then  'granule
                do_Ramp 1
            ElseIf CellOption(8).Value = True Then  'Golgi
                do_Ramp 2
            ElseIf CellOption(9).Value = True Then  'Basket
                do_Ramp 3
            ElseIf CellOption(10).Value = True Then  'Purkinje
                do_Ramp 4
            ElseIf CellOption(11).Value = True Then  'Nucleus
                do_Ramp 5
            End If
        End If
        HoldNext(1).Enabled = True
        HoldNext(0).Enabled = False
    ElseIf Index = 2 Then
        If ModeOption(1).Value = True Then 'current histogram
            If CellOption(0).Value = True Then 'MF
                CellNumber.Text = Val(CellNumber.Text) + 1
                Tube.Cls
                If Val(CellNumber.Text) > 600 Then CellNumber.Text = 1
                do_CurrentHisto 0
            ElseIf CellOption(1).Value = True Then  'granule
                CellNumber.Text = Val(CellNumber.Text) + 1
                Tube.Cls
                If Val(CellNumber.Text) > 10000 Then CellNumber.Text = 1
                do_CurrentHisto 1
            ElseIf CellOption(2).Value = True Then  'Golgi
                CellNumber.Text = Val(CellNumber.Text) + 1
                Tube.Cls
                If Val(CellNumber.Text) > 900 Then CellNumber.Text = 1
                do_CurrentHisto 2
            End If
            
        End If
        HoldNext(2).Enabled = True
        HoldNext(1).Enabled = True
        HoldNext(0).Enabled = False
    End If
End Sub
Private Sub do_oneshot(Index As Integer)
Dim X As Integer
Dim y As Integer
Dim cell As Integer

    cell = Val(CellNumber.Text)
    If Index = 0 Then
        Tube.ScaleTop = 1
        Tube.ScaleHeight = -1
        Tube.ScaleLeft = 0
        Tube.ScaleWidth = 200
        For X = 1 To 200
            If X Mod Time_step_size = 0 Then
                If X = 20 Then
                    MFinput 2
                Else
                    MFinput 1
                End If
                Tube.PSet (X, (MFS(cell, 1).bfreq + (0 * MFS(cell, 1).CSFreq)) * MFS(cell, 1).Thr)
            End If
        Next X
    ElseIf Index = 1 Then
        Tube.ScaleTop = 20
        Tube.ScaleHeight = -100
        Tube.ScaleLeft = 0
        Tube.ScaleWidth = 400
        For X = 0 To 400
            If X Mod Time_step_size = 0 Then
                If X = 40 Or X = 45 Or X = 50 Then
                    MF(Gr(1).MF(1)) = 1
                    MF(Gr(1).MF(2)) = 1
                    MF(Gr(1).MF(3)) = 1
                    MF(Gr(1).MF(4)) = 1
                    MF(Gr(1).MF(5)) = 1
                    MF(Gr(1).MF(6)) = 1
                Else
                    MF(Gr(1).MF(1)) = 0
                    MF(Gr(1).MF(2)) = 0
                    MF(Gr(1).MF(3)) = 0
                    MF(Gr(1).MF(4)) = 0
                    MF(Gr(1).MF(5)) = 0
                    MF(Gr(1).MF(6)) = 0
                End If
                DoWork 1, 0, 0, 0
                Tube.PSet (X, Gr(1).Thr), vbRed
                Tube.PSet (X, Gr(1).v)
                If Gr(1).act = 1 Then Tube.Line -(X, 0)
                
                'Tube.PSet (X, 19 + (-1 * (Gr(1).gE_NMDA) * 50))
            End If
        Next X
    ElseIf Index = 2 Then 'Golgi to granule
        Tube.ScaleTop = 20
        Tube.ScaleHeight = -100
        Tube.ScaleLeft = 0
        Tube.ScaleWidth = 400
        For X = 0 To 400
            If X Mod Time_step_size = 0 Then
                If X = 40 Or X = 45 Or X = 50 Then
                    Gol(Gr(1).Gol(1)).act = 1
                    Gol(Gr(1).Gol(2)).act = 1
                    Gol(Gr(1).Gol(3)).act = 1
                   
                Else
                    Gol(Gr(1).Gol(1)).act = 0
                    Gol(Gr(1).Gol(2)).act = 0
                    Gol(Gr(1).Gol(3)).act = 0
                End If
                DoWork 1, 0, 0, 0
                
                Tube.PSet (X, Gr(1).Thr), vbRed
                Tube.PSet (X, Gr(1).v)
                If Gr(1).act = 1 Then Tube.Line -(X, 0)
                
                'Tube.PSet (X, 19 + (-1 * (Gr(1).gi(1) * 30)))
                
            End If
        Next X
    ElseIf Index = 3 Then  '  MF to Gol
        Tube.ScaleTop = 10
        Tube.ScaleHeight = -100
        Tube.ScaleLeft = 0
        Tube.ScaleWidth = 400
        For X = 0 To 400
            If X Mod Time_step_size = 0 Then
                If X = 40 Or X = 45 Or X = 50 Then
                    MF(Gol(1).MF(1)) = 1
                    MF(Gol(1).MF(2)) = 1
                    MF(Gol(1).MF(3)) = 1
                    MF(Gol(1).MF(4)) = 1
                   
                Else
                    MF(Gol(1).MF(1)) = 0
                    MF(Gol(1).MF(2)) = 0
                    MF(Gol(1).MF(3)) = 0
                    MF(Gol(1).MF(4)) = 0
                  
                End If
                
                DoWork 0, 1, 0, 0
                Tube.PSet (X, Gol(1).Thr), vbRed
                Tube.PSet (X, Gol(1).v)
                If Gol(1).act = 1 Then Tube.Line -(X, 0)
                
                Tube.PSet (X, -9 + (-1 * (Gol(1).gMF * 30)))
            End If
        Next X
    ElseIf Index = 4 Then  '  gr to Gol
        Tube.ScaleTop = 20
        Tube.ScaleHeight = -100
        Tube.ScaleLeft = 0
        Tube.ScaleWidth = 400
        For X = 0 To 400
            If X Mod Time_step_size = 0 Then
                If X = 40 Or X = 45 Or X = 50 Then
                    Gr(Gol(1).preGr(1)).act = 1
                    Gr(Gol(1).preGr(2)).act = 1
                    Gr(Gol(1).preGr(3)).act = 1
                    Gr(Gol(1).preGr(4)).act = 1
                    Gr(Gol(1).preGr(5)).act = 1
                    Gr(Gol(1).preGr(6)).act = 1
                    Gr(Gol(1).preGr(7)).act = 1
                    Gr(Gol(1).preGr(8)).act = 1
                    Gr(Gol(1).preGr(9)).act = 1
                    Gr(Gol(1).preGr(11)).act = 1
                    Gr(Gol(1).preGr(10)).act = 1
                   
                Else
                    Gr(Gol(1).preGr(1)).act = 0
                    Gr(Gol(1).preGr(2)).act = 0
                    Gr(Gol(1).preGr(3)).act = 0
                    Gr(Gol(1).preGr(4)).act = 0
                    Gr(Gol(1).preGr(5)).act = 0
                    Gr(Gol(1).preGr(6)).act = 0
                    Gr(Gol(1).preGr(7)).act = 0
                    Gr(Gol(1).preGr(8)).act = 0
                    Gr(Gol(1).preGr(9)).act = 0
                    Gr(Gol(1).preGr(11)).act = 0
                    Gr(Gol(1).preGr(10)).act = 0
                  
                End If
                DoWork 0, 1, 0, 0
                Tube.PSet (X, Gol(1).Thr), vbRed
                Tube.PSet (X, Gol(1).v)
                If Gol(1).act = 1 Then Tube.Line -(X, 0)
                
                Tube.PSet (X, 19 + (-1 * (Gol(1).GGr * 30)))
            End If
        Next X
    ElseIf Index = 5 Then  '  Nucleus to CF
        Tube.ScaleTop = 20
        Tube.ScaleHeight = -100
        Tube.ScaleLeft = 0
        Tube.ScaleWidth = 400
        bincounter = 1
        For X = 0 To 400
            If X Mod Time_step_size = 0 Then
                If X = 40 Or X = 45 Or X = 50 Then
                    Nc(1).act = 1
                    Nc(2).act = 1
                    Nc(3).act = 1
                Else
                    Nc(1).act = 0
                    Nc(2).act = 0
                    Nc(3).act = 0
                End If
                CF(1).GNc = 0
'                For y = 1 To 6
'                    NcCfG(y) = (Nc(y).act * GCONSTNCCF * (1 - NcCfG(y))) + ((1 - Nc(y).act) * NcCfG(y) * GDecayNCCF)
'                    CF(1).GNc = CF(1).GNc + NcCfG(y)
'                Next y
                CF(1).v = CF(1).v + ((CF(1).GLeak * (ELEAKCF - CF(1).v)) + (CF(1).GNc * (VNCCF - CF(1).v)))
                If CF(1).v > CF(1).Thr Then
                    CF(1).act = 1
                    CF(1).Thr = THRMAXCF
                Else
                    CF(1).act = 0
                    CF(1).Thr = CF(1).Thr + (ThrDecayCF * (THRBASECF - CF(1).Thr))
                End If
                Tube.PSet (X, CF(1).Thr), vbRed
                Tube.PSet (X, CF(1).v)
                If CF(1).act = 1 Then Tube.Line -(X, 0)
                
                Tube.PSet (X, 19 + (-1 * (CF(1).GNc * 30)))
            End If
        Next X
    ElseIf Index = 6 Then  '  MF to Nucleus
        Tube.ScaleTop = 0
        Tube.ScaleHeight = -1
        Tube.ScaleLeft = 0
        Tube.ScaleWidth = 250
        mfweight(1) = 0.6
        mfweight(2) = 0.6
        mfweight(3) = 0.6
        For X = 0 To 250
            If X Mod Time_step_size = 0 Then
                If X = 25 Then
                    MF(1) = 1
                    MF(2) = 1
                    MF(3) = 1
                Else
                    MF(1) = 0
                    MF(2) = 0
                    MF(3) = 0
                End If
                DoWork 0, 0, 0, 0
                
                Tube.PSet (X, Nc(1).Thr / 80), vbRed
                Tube.PSet (X, Nc(1).v / 80)
                If Nc(1).act = 1 Then Tube.Line -(X, -0.2)
                
                'Tube.PSet (x, -1 * Nc(1).NMDABind(1))
                Tube.PSet (X, -1 * (Nc(1).gNMDA(1))), vbRed
                Tube.PSet (X, -1 * I_NMDA(1)), vbWhite
                Tube.PSet (X, -1 * Nc(1).gMF2), vbGreen
            End If
        Next X
    ElseIf Index = 7 Then  '  PC to Nucleus
        Tube.ScaleTop = -30
        Tube.ScaleHeight = -50
        Tube.ScaleLeft = 0
        Tube.ScaleWidth = 400
        
        For X = 0 To 400
            If X Mod Time_step_size = 0 Then
                If X = 45 Then
                    For y = 1 To 20
                        Pc(y).act = 1
                    Next y
                Else
                    For y = 1 To 20
                        Pc(y).act = 0
                    Next y
                End If
                
                DoWork 0, 0, 0, 0
                
                Tube.PSet (X, Nc(1).Thr), vbRed
                Tube.PSet (X, -75 + (Nc(1).gPc * 30)), vbRed
                
                Tube.PSet (X, Nc(1).v)
                If Nc(1).act = 1 Then Tube.Line -(X, 0)
            End If
        Next X
    Else
    
    End If
End Sub
Private Sub do_CurrentHisto(Index As Integer)
Dim X As Integer
Dim cell As Integer

    cell = Val(CellNumber.Text)

    Tube.ScaleLeft = 1
    Tube.ScaleWidth = 1000
    Tube.ScaleTop = 195
    Tube.ScaleHeight = -200
    If Index = 0 Then  'MF
        For X = 1 To 1000
            Tube.PSet (X, 200 * (MF_histo(cell, X) / HistoDivisor))
        Next X
    ElseIf Index = 1 Then  'gr
        For X = 1 To 1000
            Tube.PSet (X, 200 * (GR_histo(cell, X) / HistoDivisor))
        Next X
    ElseIf Index = 2 Then  'Golgi
        For X = 1 To 1000
            Tube.PSet (X, 200 * (Go_histo(cell, X) / HistoDivisor))
        Next X
    End If
End Sub
Private Sub do_Ramp(Index As Integer)
Dim X As Integer
Dim cell As Integer

    cell = Val(CellNumber.Text)
    Tube.ScaleLeft = 14
    Tube.ScaleWidth = 1000
    Tube.ScaleTop = 20
    Tube.ScaleHeight = -100
    If Index = 1 Then       'Granule
        Gr(1).v = ThrBasegr - 2
        'Gr(1).gE_NMDA = 0.02
        For X = 1 To 1000 Step Time_step_size
            Gr(1).v = Gr(1).v + 0.02 * Time_step_size
            'Gr(1).gE_NMDA = Gr(1).gE_NMDA + 0.00002 * Time_step_size
            'Gr(1).V = Gr(1).V + (((gLGr) * (ELeakgr - Gr(1).V) - Gr(1).gE_NMDA * Gr(1).V + Gr(1).gi * (EGABAgr - Gr(1).V)))
            
            Tube.PSet (X, Gr(1).v)
            Gr(1).Thr = Gr(1).Thr + (ThrDecayGr * (ThrBasegr - Gr(1).Thr))
            If Gr(1).v > Gr(1).Thr Then
                Gr(1).Thr = ThrmaxGr
                Tube.Line -(X, 0)
            Else
                'Gr(1).Thr = Gr(1).Thr + (ThrDecayGr * (ThrBasegr - Gr(1).Thr))
            End If
            Tube.PSet (X, Gr(1).Thr), vbRed
        Next X
    ElseIf Index = 2 Then   'Golgi
        Gol(1).v = ThrBaseGo - 2
        For X = 1 To 1000 Step Time_step_size
            Gol(1).v = Gol(1).v + 0.025 * Time_step_size
            Tube.PSet (X, Gol(1).v)
            Gol(1).Thr = Gol(1).Thr + (ThrDecayGo * (ThrBaseGo - Gol(1).Thr))
            If Gol(1).v > Gol(1).Thr Then
              Gol(1).Thr = ThrmaxGo
              Tube.Line -(X, 0)
            End If
            Tube.PSet (X, Gol(1).Thr), vbRed
        Next X
    ElseIf Index = 3 Then               'Basket
        Bk(1).v = THRBASEStell - 2
        For X = 1 To 1000 Step Time_step_size
            Bk(1).v = Bk(1).v + 0.02 * Time_step_size
            Tube.PSet (X, Bk(1).v)
            Bk(1).Thr = Bk(1).Thr + (ThrDecayStell * (THRBASEStell - Bk(1).Thr))
            If Bk(1).v > Bk(1).Thr Then
                Bk(1).Thr = THRMAXStell
                Tube.Line -(X, 0)
            End If
            Tube.PSet (X, Bk(1).Thr), vbRed
        Next X
    ElseIf Index = 4 Then            'Purkinje
        Pc(1).v = THRBASEPC - 2
        For X = 1 To 1000 Step Time_step_size
            Pc(1).v = Pc(1).v + 0.02 * Time_step_size
            Tube.PSet (X, Pc(1).v)
            Pc(1).Thr = Pc(1).Thr + (ThrDecayPC * (THRBASEPC - Pc(1).Thr))
            If Pc(1).v > Pc(1).Thr Then
                Pc(1).Thr = THRMAXPC
                Tube.Line -(X, 0)
            End If
            Tube.PSet (X, Pc(1).Thr), vbRed
        Next X
    
    ElseIf Index = 5 Then             'Nucleus
        Nc(1).v = THRBASENC - 2
        For X = 1 To 1000 Step Time_step_size
            Nc(1).v = Nc(1).v + 0.02 * Time_step_size
            Tube.PSet (X, Nc(1).v)
            Nc(1).Thr = Nc(1).Thr + (THRdecayNC * (THRBASENC - Nc(1).Thr))
            If Nc(1).v > Nc(1).Thr Then
                Nc(1).Thr = THRMAXNC
                Tube.Line -(X, 0)
            End If
            Tube.PSet (X, Nc(1).Thr), vbRed
        Next X
    
    End If

End Sub

