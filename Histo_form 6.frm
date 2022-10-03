VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Histo_form 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Histograms"
   ClientHeight    =   13275
   ClientLeft      =   3885
   ClientTop       =   750
   ClientWidth     =   14520
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   -19940.51
   ScaleMode       =   0  'User
   ScaleTop        =   13000
   ScaleWidth      =   14435.79
   Begin MSComDlg.CommonDialog histo_dialog 
      Left            =   7665
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label excel_label 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   16065
      TabIndex        =   50
      Top             =   420
      Width           =   645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   225
      Index           =   49
      Left            =   13335
      TabIndex        =   49
      Top             =   6615
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   225
      Index           =   48
      Left            =   13335
      TabIndex        =   48
      Top             =   6300
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   225
      Index           =   47
      Left            =   13440
      TabIndex        =   47
      Top             =   5775
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   46
      Left            =   13440
      TabIndex        =   46
      Top             =   5460
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   45
      Left            =   13440
      TabIndex        =   45
      Top             =   5145
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   44
      Left            =   13440
      TabIndex        =   44
      Top             =   4830
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   43
      Left            =   13440
      TabIndex        =   43
      Top             =   4515
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   42
      Left            =   13440
      TabIndex        =   42
      Top             =   4200
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   41
      Left            =   13440
      TabIndex        =   41
      Top             =   3885
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   40
      Left            =   13440
      TabIndex        =   40
      Top             =   3570
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   39
      Left            =   13545
      TabIndex        =   39
      Top             =   3360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   38
      Left            =   13545
      TabIndex        =   38
      Top             =   3045
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   37
      Left            =   13545
      TabIndex        =   37
      Top             =   2730
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   36
      Left            =   13545
      TabIndex        =   36
      Top             =   2415
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   35
      Left            =   13545
      TabIndex        =   35
      Top             =   2100
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   34
      Left            =   13545
      TabIndex        =   34
      Top             =   1785
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   33
      Left            =   13545
      TabIndex        =   33
      Top             =   1470
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   32
      Left            =   13545
      TabIndex        =   32
      Top             =   1155
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   31
      Left            =   12390
      TabIndex        =   31
      Top             =   11235
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   30
      Left            =   12390
      TabIndex        =   30
      Top             =   10920
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   29
      Left            =   12390
      TabIndex        =   29
      Top             =   10605
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   28
      Left            =   12390
      TabIndex        =   28
      Top             =   10290
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   27
      Left            =   12390
      TabIndex        =   27
      Top             =   9975
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   26
      Left            =   12390
      TabIndex        =   26
      Top             =   9660
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   25
      Left            =   12390
      TabIndex        =   25
      Top             =   9345
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   24
      Left            =   12390
      TabIndex        =   24
      Top             =   9030
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   23
      Left            =   12390
      TabIndex        =   23
      Top             =   8610
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   22
      Left            =   12390
      TabIndex        =   22
      Top             =   8295
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   21
      Left            =   12390
      TabIndex        =   21
      Top             =   7980
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   20
      Left            =   12390
      TabIndex        =   20
      Top             =   7665
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   19
      Left            =   12390
      TabIndex        =   19
      Top             =   7350
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   18
      Left            =   12390
      TabIndex        =   18
      Top             =   7035
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   17
      Left            =   12390
      TabIndex        =   17
      Top             =   6720
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   16
      Left            =   12390
      TabIndex        =   16
      Top             =   6405
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   15
      Left            =   12390
      TabIndex        =   15
      Top             =   5985
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   14
      Left            =   12390
      TabIndex        =   14
      Top             =   5670
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   13
      Left            =   12390
      TabIndex        =   13
      Top             =   5355
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   12
      Left            =   12390
      TabIndex        =   12
      Top             =   5040
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   11
      Left            =   12390
      TabIndex        =   11
      Top             =   4725
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   10
      Left            =   12390
      TabIndex        =   10
      Top             =   4410
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   9
      Left            =   12390
      TabIndex        =   9
      Top             =   4095
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   8
      Left            =   12390
      TabIndex        =   8
      Top             =   3780
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   7
      Left            =   12390
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   6
      Left            =   12390
      TabIndex        =   6
      Top             =   3045
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   5
      Left            =   12390
      TabIndex        =   5
      Top             =   2730
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   4
      Left            =   12390
      TabIndex        =   4
      Top             =   2415
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   3
      Left            =   12390
      TabIndex        =   3
      Top             =   2100
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   2
      Left            =   12390
      TabIndex        =   2
      Top             =   1785
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   1
      Left            =   12390
      TabIndex        =   1
      Top             =   1470
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   225
      Index           =   0
      Left            =   12390
      TabIndex        =   0
      Top             =   1155
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Menu file_m 
      Caption         =   "&File"
      Begin VB.Menu open_histo_m 
         Caption         =   "&Open"
         Begin VB.Menu open_histo_menu 
            Caption         =   "granule"
            Index           =   0
         End
         Begin VB.Menu open_histo_menu 
            Caption         =   "Golgi"
            Index           =   1
         End
         Begin VB.Menu open_histo_menu 
            Caption         =   "Mossy fiber"
            Index           =   2
         End
         Begin VB.Menu open_histo_menu 
            Caption         =   "climbing fiber"
            Index           =   3
         End
         Begin VB.Menu open_histo_menu 
            Caption         =   "Nucleus"
            Index           =   4
         End
         Begin VB.Menu open_histo_menu 
            Caption         =   "Purkinje"
            Index           =   5
         End
         Begin VB.Menu open_histo_menu 
            Caption         =   "Stellate"
            Index           =   6
         End
         Begin VB.Menu open_histo_menu 
            Caption         =   "Basket"
            Index           =   7
         End
         Begin VB.Menu open_histo_menu 
            Caption         =   "Responses"
            Index           =   8
         End
      End
      Begin VB.Menu save_histo_m 
         Caption         =   "&Save"
         Begin VB.Menu save_histo_menu 
            Caption         =   "granule"
            Index           =   0
         End
         Begin VB.Menu save_histo_menu 
            Caption         =   "Golgi"
            Index           =   1
         End
         Begin VB.Menu save_histo_menu 
            Caption         =   "Mossy fiber"
            Index           =   2
         End
         Begin VB.Menu save_histo_menu 
            Caption         =   "climbing fiber"
            Index           =   3
         End
         Begin VB.Menu save_histo_menu 
            Caption         =   "Nucleus"
            Index           =   4
         End
         Begin VB.Menu save_histo_menu 
            Caption         =   "Purkinje"
            Index           =   5
         End
         Begin VB.Menu save_histo_menu 
            Caption         =   "Stellate"
            Index           =   6
         End
         Begin VB.Menu save_histo_menu 
            Caption         =   "Basket"
            Index           =   7
         End
         Begin VB.Menu save_histo_menu 
            Caption         =   "Responses"
            Index           =   8
         End
         Begin VB.Menu save_histo_menu 
            Caption         =   "Save all histos and rasters"
            Index           =   9
         End
         Begin VB.Menu save_histo_menu 
            Caption         =   "Granule and Golgi"
            Index           =   10
         End
      End
   End
   Begin VB.Menu show_histo 
      Caption         =   "&Show"
      Begin VB.Menu show_histo_menu 
         Caption         =   "granule"
         Index           =   0
      End
      Begin VB.Menu show_histo_menu 
         Caption         =   "Golgi"
         Index           =   1
      End
      Begin VB.Menu show_histo_menu 
         Caption         =   "Mossy fiber"
         Index           =   2
      End
      Begin VB.Menu show_histo_menu 
         Caption         =   "climbing fiber"
         Index           =   3
      End
      Begin VB.Menu show_histo_menu 
         Caption         =   "Nucleus"
         Index           =   4
      End
      Begin VB.Menu show_histo_menu 
         Caption         =   "Purkinje"
         Index           =   5
      End
      Begin VB.Menu show_histo_menu 
         Caption         =   "Stellate"
         Index           =   6
      End
      Begin VB.Menu show_histo_menu 
         Caption         =   "Basket"
         Index           =   7
      End
      Begin VB.Menu show_histo_menu 
         Caption         =   "Responses"
         Index           =   8
      End
   End
   Begin VB.Menu collect_histo 
      Caption         =   "&Collect"
      Begin VB.Menu collect_histo_menu 
         Caption         =   "granule"
         Index           =   0
      End
      Begin VB.Menu collect_histo_menu 
         Caption         =   "Gogli"
         Index           =   1
      End
      Begin VB.Menu collect_histo_menu 
         Caption         =   "Mossy fiber"
         Index           =   2
      End
      Begin VB.Menu collect_histo_menu 
         Caption         =   "climbing fiber"
         Index           =   3
      End
      Begin VB.Menu collect_histo_menu 
         Caption         =   "Nucleus"
         Index           =   4
      End
      Begin VB.Menu collect_histo_menu 
         Caption         =   "Purkinje"
         Index           =   5
      End
      Begin VB.Menu collect_histo_menu 
         Caption         =   "Basket"
         Index           =   6
      End
      Begin VB.Menu collect_histo_menu 
         Caption         =   "Responses"
         Index           =   7
      End
   End
   Begin VB.Menu HeatMaps_menu 
      Caption         =   "HeatMaps"
      Begin VB.Menu HM_menu 
         Caption         =   "Golgi Raw"
         Index           =   1
      End
      Begin VB.Menu HM_menu 
         Caption         =   "Golgi normalized"
         Index           =   2
      End
      Begin VB.Menu HM_menu 
         Caption         =   "granule Raw"
         Index           =   3
      End
      Begin VB.Menu HM_menu 
         Caption         =   "granule normalized"
         Index           =   4
      End
   End
   Begin VB.Menu reset_histo 
      Caption         =   "&Reset"
      Begin VB.Menu reset_histo_menu 
         Caption         =   "granule"
         Index           =   0
      End
      Begin VB.Menu reset_histo_menu 
         Caption         =   "Golgi"
         Index           =   1
      End
      Begin VB.Menu reset_histo_menu 
         Caption         =   "Mossy fiber"
         Index           =   2
      End
      Begin VB.Menu reset_histo_menu 
         Caption         =   "climbing fiber"
         Index           =   3
      End
      Begin VB.Menu reset_histo_menu 
         Caption         =   "Nucleus"
         Index           =   4
      End
      Begin VB.Menu reset_histo_menu 
         Caption         =   "Purkinje"
         Index           =   5
      End
      Begin VB.Menu reset_histo_menu 
         Caption         =   "basket"
         Index           =   6
      End
      Begin VB.Menu reset_histo_menu 
         Caption         =   "Responses"
         Index           =   7
      End
   End
   Begin VB.Menu analyze_histo 
      Caption         =   "Analyze"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Histo_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempHistos(12001, 1002) As Integer
Dim ScaleMax As Integer
Dim StartHisto As Integer
Dim StopHisto As Integer
Dim MaxHistos As Integer
Dim TempHistoDivisor As Integer

Private Sub analyze_histo_Click()
Dim cells As Integer
    Dim histo_data(SYNUMBER, 10) As Single
    Dim max_cells As Integer
    Dim back_total As Single
    Dim CS_total As Single
    Dim total As Single
    Dim pre_CS_total As Single
    Dim threshold As Single
    Dim bin As Integer
    Dim start_bin As Integer
    Dim stop_bin As Integer
    Dim cs_onset As Integer
    Dim cs_duration As Integer
    Dim latency_to_onset As Integer
    Dim latency_to_peak As Integer
    Dim pre_CS_std_dev As Double
    Dim pre_CS_sum As Double
    Dim pre_CS_sum_sqr As Double
    Dim up_or_down As Integer
    Dim max As Single
    Dim max_bin As Integer
    Dim min_bin As Integer
    Dim min As Single
        
    Dim X As Integer
    Dim y As Integer
    Dim times As Integer
    Dim cap As String
    Dim x_spacer As Single
    Dim is_zero As Single
    Dim current_sweeps(50) As Integer
    Dim i As Integer
    Dim j As Integer
    Dim smooth(1000) As Single
    Dim smooth2(1000) As Single
    Dim smoother As Single
    Dim baseline As Single
    Dim base_counter As Single
    Dim is_a_max As Integer
    Dim baseline2 As Single
    Dim integrate As Single
    Dim integrate2 As Single
    Dim onset_temp(4) As Integer
        
        
    histo_data_form.Visible = True
    threshold = Val(histo_info_form.Text1(5).Text)
    max_cells = Val(histo_info_form.Text1(4).Text)
    start_bin = Val(histo_info_form.Text1(0).Text)
    stop_bin = Val(histo_info_form.Text1(1).Text)
    cs_onset = Val(histo_info_form.Text1(2).Text)
    cs_duration = Val(histo_info_form.Text1(3).Text)
    
    total = 0
    max = 0
    max_bin = 0
    min_bin = 0
    min = 32000
    
    Histo_form.MousePointer = 11
    For cells = 1 To max_cells
        For bin = start_bin To stop_bin
            total = total + GR_histo(cells, bin)
            If GR_histo(cells, bin) > max Then
                max = GR_histo(cells, bin)
                max_bin = bin
            End If
            If GR_histo(cells, bin) < min Then
                min = GR_histo(cells, bin)
                min_bin = bin
            End If
        Next bin
        histo_data(cells, 3) = total / (stop_bin - (start_bin - 1))
        histo_data(cells, 6) = max
        histo_data(cells, 7) = max_bin
        histo_data(cells, 8) = min
        histo_data(cells, 9) = min_bin
        
        If total > threshold Then histo_data(cells, 1) = 1 Else histo_data(cells, 1) = 0
        
        pre_CS_total = 0
        pre_CS_sum_sqr = 0
        For bin = start_bin To cs_onset - 1
            pre_CS_total = pre_CS_total + GR_histo(cells, bin)
            pre_CS_sum = GR_histo(cells, bin)
            pre_CS_sum_sqr = pre_CS_sum_sqr + (pre_CS_sum * pre_CS_sum)
        Next bin
        histo_data(cells, 4) = pre_CS_total / (cs_onset - start_bin)
        
        CS_total = 0
        For bin = cs_onset To cs_onset + cs_duration - 1
            CS_total = CS_total + GR_histo(cells, bin)
        Next bin
        histo_data(cells, 5) = CS_total / cs_duration
        
        If histo_data(cells, 5) > histo_data(cells, 4) Then
            up_or_down = 1
        ElseIf histo_data(cells, 5) = histo_data(cells, 4) Then
            up_or_down = 0
        Else
            up_or_down = -1
        End If
        histo_data(cells, 2) = up_or_down
        If pre_CS_total > 0 Then
            pre_CS_std_dev = (pre_CS_total * pre_CS_total) / (cs_onset - start_bin + 1)
            pre_CS_std_dev = pre_CS_sum_sqr / pre_CS_std_dev
            pre_CS_std_dev = pre_CS_std_dev / (cs_onset - start_bin + 1)
        Else
            pre_CS_std_dev = 0
        End If
        
        latency_to_onset = 0
        If up_or_down = 1 Then
            For bin = cs_onset To stop_bin
                If GR_histo(cells, bin) > pre_CS_total + (25 * pre_CS_std_dev) Then
                    latency_to_onset = bin
                    bin = stop_bin
                End If
            Next bin
        
        ElseIf up_or_down = -1 Then
            For bin = cs_onset To stop_bin
                If GR_histo(cells, bin) < pre_CS_total + (25 * pre_CS_std_dev) Then
                    latency_to_onset = bin
                    bin = stop_bin
                End If
            Next bin
        End If
        
        
'***************************** PEAK DETECTION *****************************************

starter = Val(histo_info_form.Text1(0))
stopper = Val(histo_info_form.Text1(1))
counter = cells

For times = starter + 2 To stopper - 2
            smoother = 0
            For i = times - 2 To times + 2
                smoother = smoother + GR_histo(counter, i)
            Next i
            smoother = smoother / 5#
            smooth(times) = smoother
        Next times
        
        For j = 1 To 10
                        
            For times = starter + 2 To stopper - 2
                smoother = 0
                For i = times - 2 To times + 2
                    smoother = smoother + smooth(i)
                Next i
                smoother = smoother / 5#
                smooth(times) = smoother
                
            Next times
            DoEvents
        Next j
        
        baseline = 0
        base_counter = 0
        
        For times = starter To cs_onset - 1
            baseline = baseline + smooth(times)
            base_counter = base_counter + 1
        Next times
        baseline = baseline / base_counter
        
        
        For i = 1 To 4
            onset_temp(i) = 0
        Next i
        For times = cs_onset To stopper - 20
                is_a_max = 1
                integrate = 0
                
            For i = times - 20 To times + 20
                If (smooth(i) > smooth(times)) Then
                    is_a_max = 0
                    i = times + 20
                End If
                integrate = integrate + smooth(i)
            Next i
            
            If is_a_max = 1 Then
                integrate = integrate / 41
               If (smooth(times) < integrate + 4) Then
                    is_a_max = 0
               End If
            End If
            
            If is_a_max = 1 Then 'And smooth(times) > baseline + 20 Then
                If onset_temp(1) = 0 Then
                    onset_temp(1) = times
                ElseIf onset_temp(2) = 0 Then
                    onset_temp(2) = times
                ElseIf onset_temp(3) = 0 Then
                    onset_temp(3) = times
                Else
                    onset_temp(4) = times
                End If
                
             '   Histo_form.Line ((x * x_spacer) + times - starter, y + 0.2 + ((smooth(times) - baseline) / max))-((x * x_spacer) + times - starter, y + 0.3 + ((smooth(times) - baseline) / max)), RGB(255, 0, 0)
            End If
        Next times
        
        
        If baseline > 30 Then
        For times = starter To stopper
            smooth2(times) = -1 * smooth(times)
        Next times
        basline = -1 * baseline
        onset_temp(0) = 0
        For times = cs_onset To stopper - 40
                is_a_max = 1
                integrate = 0
                integrate2 = 0
                
            For i = times - 40 To times + 40
                If i > times - 20 And i < times + 40 Then
                    If (smooth2(i) > smooth2(times)) Then
                        is_a_max = 0
                        i = times + 40
                    End If
                End If
                
                If i < times Then
                    integrate = integrate + smooth2(i)
                ElseIf i > times Then
                    integrate2 = integrate2 + smooth2(i)
                End If
            Next i
            
            If is_a_max = 1 Then
                
                integrate = integrate / 39
                integrate2 = integrate2 / 39
               If smooth2(times) < integrate + 5 Then is_a_max = 0
               
               If smooth2(times) < integrate2 + 5 Then is_a_max = 0
               
            End If
            If is_a_max = 1 And smooth2(times) > (baseline * -1) + 100 Then
                onset_temp(0) = times
            '    Histo_form.Line ((x * x_spacer) + times - starter, y + 0.2 + ((smooth(times) - baseline) / max))-((x * x_spacer) + times - starter, y + 0.3 + ((smooth(times) - baseline) / max)), RGB(0, 0, 255)
                
            End If
            
        Next times
        For i = 0 To 4
                If onset_temp(i) <> 0 Then
                    
                    onsets(onset_temp(i), i) = onsets(onset_temp(i), i) + 1
                    If i = 1 Then X = 4
                    
                End If
            Next i
        End If
        
        DoEvents
    Next cells
    Histo_form.MousePointer = 0
    histo_data_form.Cls
        For times = 400 To 1000
            histo_data_form.Line (times, 0)-(times, onsets(times, 1)), RGB(255, 255, 255)
            histo_data_form.Line (times, onsets(times, 1))-(times, onsets(times, 1) + onsets(times, 2)), RGB(255, 255, 0)
            histo_data_form.Line (times, onsets(times, 1) + onsets(times, 2))-(times, onsets(times, 1) + onsets(times, 2) + onsets(times, 3)), RGB(255, 0, 0)
        Next times
    DoEvents
    'If filecount + starting_sheet - 1 < 10 Then
    '    sheet = sheet + Right$(Str(filecount + starting_sheet - 1), 1)
    'Else
    '    sheet = sheet + Right$(Str(filecount + starting_sheet - 1), 2)
    'End If
    
    
'    histo_info_form.Check1.Value = True
    draw_histos 0, Val(histo_info_form.cell_num_text), 50, SYNUMBER, 10, 5, Val(histo_info_form.Text1(0)), Val(histo_info_form.Text1(1)), 1000, Val(histo_info_form.Text1(2)), Val(histo_info_form.Text1(3)), max
    'DoEvents
End Sub

Private Sub collect_histo_menu_Click(Index As Integer)
If collect_histo_menu(Index).Checked = False Then collect_histo_menu(Index).Checked = True Else collect_histo_menu(Index).Checked = False
End Sub

Private Sub file_histo_menu_Click(Index As Integer)
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Debug.Print KeyAscii
    
    If KeyAscii = 61 Then  'scale up
        ScaleMax = ScaleMax + 50
    ElseIf KeyAscii = 45 Then  ' scale down
        ScaleMax = ScaleMax - 50
        If ScaleMax = 0 Then ScaleMax = 50
    ElseIf KeyAscii = 46 Then  ' advance
        StartHisto = StartHisto + 10
        If StartHisto > MaxHistos - 10 Then StartHisto = MaxHistos - 10
    ElseIf KeyAscii = 44 Then ' retreat
        StartHisto = StartHisto - 10
        If StartHisto < 1 Then StartHisto = 1
    End If
    Histo_form.Caption = "Scale max = " + Str$(ScaleMax) + "Starthisto =" + Str$(StartHisto)
    DrawHistos
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then
        histo_info_form.cell_num_text.Text = Histo_form.Label1(49).Caption
        histo_info_form.Command1(0).Value = True
    
    ElseIf Button = 1 Then
'        histo_info_form.cell_num_text.Text = Histo_form.label1(0).Caption - 50
'        histo_info_form.command1(0).Value = True
    End If
End Sub

Public Sub HM_menu_Click(Index As Integer)
Dim Htemp(900, 1000) As Single
Dim Htemp2(900, 1000) As Single
Dim MaxFire(900) As Single
Dim MaxFireTime(900) As Integer
Dim Mx As Single
Dim r, g, b As Integer
Dim holder As Integer

Dim i, j As Integer

    If Index < 3 Then 'Doing Golgi histos Raw
        For i = 1 To 900
            For j = 1 To 1000
                Htemp2(i, j) = Go_histo(i, j)
            Next j
        Next i
        
        Erase MaxFire
        Erase MaxFireTime
        Mx = 0
        
        For i = 1 To 900
            For j = 1 To 1000
                If Htemp2(i, j) > MaxFire(i) Then
                    MaxFire(i) = Htemp2(i, j)
                    MaxFireTime(i) = j
                End If
                If Htemp2(i, j) > Mx Then Mx = Htemp2(i, j)
            Next j
            
            If Index = 2 Then    ' normalize each cell to its own max rate
                For j = 1 To 1000
                    If MaxFire(i) > 0 Then Htemp2(i, j) = Htemp2(i, j) / MaxFire(i)
                Next j
            End If
        Next i
        If Index = 1 Then       'normalize to the maximum for all cells
            For i = 1 To 900
                For j = 1 To 1000
                    Htemp2(i, j) = Htemp2(i, j) / Mx
                Next j
            Next i
        End If
        
    ElseIf Index < 5 Then                ' Doing granule histos
    
    
    End If

    For i = 1 To 900
        Mx = 1001
        For j = 1 To 900
            If MaxFireTime(j) < Mx Then
                Mx = MaxFireTime(j)
                holder = j
            End If
        Next j
        For j = 1 To 1000
            Htemp(i, j) = Htemp2(holder, j)
            MaxFireTime(holder) = 10000
        Next j
    Next i
    
    
    '  Now display Htemp
'    Histo_form.ScaleMode = vbPixels
'    Histo_form.Width = 1020
'    Histo_form.Height = 950
'    Histo_form.ScaleMode = vbUser
    
    Histo_form.ScaleLeft = 0
    Histo_form.ScaleWidth = 1000
    Histo_form.ScaleTop = 0
    Histo_form.ScaleHeight = 900
    Histo_form.BackColor = vbBlack
    Histo_form.Cls
    Histo_form.DrawWidth = 1
    For i = 1 To 900
        For j = 1 To 1000
            r = ColorMap(Htemp(i, j), 1)
            g = ColorMap(Htemp(i, j), 2)
            b = ColorMap(Htemp(i, j), 3)
            Histo_form.PSet (j, i), RGB(r, g, b)
        Next j
    Next i
    
End Sub

Private Sub open_histo_menu_Click(Index As Integer)
    
    Dim i As Integer
    Dim j As Integer
    
    histo_dialog.filename = ""
    
    histo_dialog.ShowOpen
    
    histo_info_form.Option1(0).Value = True
    
    If histo_dialog.filename <> "" Then
        Close #5
        Open histo_dialog.filename For Binary As #5
        Select Case Index
            Case 0
                Get #5, , GR_histo
                Close #5
                max = 0
                
                For i = 1 To SYNUMBER
                For j = 1 To 1000
                    If GR_histo(i, j) > max Then max = GR_histo(i, j)
                Next j
                Next i
                For i = 0 To 7
                   show_histo_menu(i).Checked = False
                Next i
                show_histo_menu(0).Checked = True
                histo_info_form.Visible = True
                draw_histos 0, Val(histo_info_form.cell_num_text), 50, SYNUMBER, 10, 5, Val(histo_info_form.Text1(0)), Val(histo_info_form.Text1(1)), 1000, Val(histo_info_form.Text1(2)), Val(histo_info_form.Text1(3)), max
                
            Case 1
                Get #5, , Go_histo
                TempHistoDivisor = Go_histo(0, 0)
            Case 2
                Get #5, , MF_histo
                TempHistoDivisor = MF_histo(0, 0)
            Case 3
                Get #5, , CF_histo
                TempHistoDivisor = CF_histo(0, 0)
            Case 4
                Get #5, , Nuc_histo
                TempHistoDivisor = Nuc_histo(0, 0)
            Case 5
                Get #5, , Purk_histo
                TempHistoDivisor = Purk_histo(0, 0)
            Case 6
                Get #5, , Stellate_histo
                TempHistoDivisor = Stellate_histo(0, 0)
            Case 7
                Get #5, , Basket_histo
                TempHistoDivisor = Basket_histo(0, 0)
            Case 7
                Get #5, , response_histo
                TempHistoDivisor = response_histo(0)
        End Select
        show_histo.Enabled = True
    End If
    
End Sub

Private Sub reset_histo_menu_Click(Index As Integer)
Dim bin As Integer
Dim cells As Integer

For bin = 1 To 1002
    Select Case Index
    Case 0
    
    For cells = 1 To SYNUMBER
        GR_histo(cells, bin) = 0
    Next cells
    
    Case 1
    
    For cells = 1 To 900
        Go_histo(cells, bin) = 0
    Next cells
    
    Case 2
    
    For cells = 1 To 20
        Purk_histo(cells, bin) = 0
    Next cells
    
    Case 3
    
    For cells = 1 To 6
        Nuc_histo(cells, bin) = 0
    Next cells
    
    Case 4
    
    
    For cells = 1 To 600
        MF_histo(cells, bin) = 0
    Next cells
    
    Case 5
    
    For cells = 1 To 1
        CF_histo(cells, bin) = 0
    Next cells
    
    Case 6
    
    For cells = 1 To 200
        Stellate_histo(cells, bin) = 0
    Next cells
    
    Case 7
    
    response_histo(bin) = 0
    
    
    End Select
    HistoDivisor = 1
Next bin
End Sub

Public Sub save_histo_menu_Click(Index As Integer)
Dim save_name As String

    If Index <> 8 Then
        histo_dialog.filename = ""
        histo_dialog.ShowOpen
        If histo_dialog.filename <> "" Then
            Close #5
            Open histo_dialog.filename For Binary As #5
            Select Case Index
                Case 0
                    GR_histo(0, 0) = (HistoDivisor)
                    Put #5, , GR_histo
                Case 1
                    Go_histo(0, 0) = (HistoDivisor)
                    Put #5, , Go_histo
                Case 2
                    MF_histo(0, 0) = (HistoDivisor)
                    Put #5, , MF_histo
                Case 3
                    CF_histo(0, 0) = (HistoDivisor)
                    Put #5, , CF_histo
                Case 4
                    Nuc_histo(0, 0) = (HistoDivisor)
                    Put #5, , Nuc_histo
                Case 5
                    Purk_histo(0, 0) = (HistoDivisor)
                    Put #5, , Purk_histo
                Case 6
                    Stellate_histo(0, 0) = (HistoDivisor)
                    Put #5, , Stellate_histo
                Case 7
                    Basket_histo(0, 0) = (HistoDivisor)
                    Put #5, , Basket_histo
                Case 8
                    response_histo(0) = (HistoDivisor)
                    Put #5, , response_histo
                Case 10
                    GR_histo(0, 0) = (HistoDivisor)
                    Go_histo(0, 0) = (HistoDivisor)
                    Put #5, , GR_histo
                    Put #5, , Go_histo
            End Select
        End If
    Else
                
                Close #5
                '''save_name = strFilenameRoot + "-gr_histo " + Format(Now, "mm-dd-yy  hh-mm AM/PM")
                save_name = strFilenameRoot + "-gr_histo-" + Left(CommandLine, 3) + "-" + Format(TrialCounter, "00000")
                
                Open save_name For Binary As #5
                Put #5, , GR_histo
                Close #5
''''''''                save_name = strFilenameRoot + "-Go_histo " + Format(Now, "mm-dd-yy  hh-mm AM/PM")
''''''''                Open save_name For Binary As #5
''''''''                Put #5, , Go_histo
''''''''                Close #5
''''''''                save_name = strFilenameRoot + "-MF_histo " + Format(Now, "mm-dd-yy  hh-mm AM/PM")
''''''''                Open save_name For Binary As #5
''''''''                Put #5, , MF_histo
''''''''                Close #5
''''''''                save_name = strFilenameRoot + "-CF_histo " + Format(Now, "mm-dd-yy  hh-mm AM/PM")
''''''''                Open save_name For Binary As #5
''''''''                Put #5, , CF_histo
''''''''                Close #5
''''''''                save_name = strFilenameRoot + "-Nuc_histo " + Format(Now, "mm-dd-yy  hh-mm AM/PM")
''''''''                Open save_name For Binary As #5
''''''''                Put #5, , Nuc_histo
''''''''                Close #5
''''''''                save_name = strFilenameRoot + "-Purk_histo " + Format(Now, "mm-dd-yy  hh-mm AM/PM")
''''''''                Open save_name For Binary As #5
''''''''                Put #5, , Purk_histo
''''''''                Close #5
''''''''                save_name = strFilenameRoot + "-response_histo " + Format(Now, "mm-dd-yy  hh-mm AM/PM")
''''''''                Open save_name For Binary As #5
''''''''                Put #5, , response_histo
''''''''                Close #5
''''''''                save_name = strFilenameRoot + "-rasters " + Format(Now, "mm-dd-yy  hh-mm AM/PM")
''''''''                Open save_name For Binary As #5
''''''''                Put #5, , rasters
''''''''                Close #5
''''''''                save_name = strFilenameRoot + "-Gran_weights " + Format(Now, "mm-dd-yy  hh-mm AM/PM")
''''''''                Open save_name For Binary As #5
''''''''                Put #5, , gr_PC_weights_rasters
''''''''                Close #5
''''''''                save_name = strFilenameRoot + "-MF_weights " + Format(Now, "mm-dd-yy  hh-mm AM/PM")
''''''''                Open save_name For Binary As #5
''''''''                Put #5, , mf_nuc_weights_rasters
''''''''                Close #5
    End If
End Sub

Private Sub show_histo_menu_Click(Index As Integer)
Dim i As Integer
Dim j As Integer
Dim X As Integer

Dim HistoPointer As Integer

    HistoPointer = 1
    ScaleMax = 100
    StartHisto = 1
    Histo_form.ScaleLeft = -20
    Histo_form.ScaleWidth = 1040
    Histo_form.ScaleHeight = -10.4
    Histo_form.ScaleTop = 10.2
    Histo_form.ForeColor = vbBlack
    Erase TempHistos
    Select Case Index
        Case 0      ' granule
            MaxHistos = GrX * GrY
            For i = 1 To MaxHistos
                For j = 1 To 1000
                    TempHistos(i, j) = GR_histo(i, j)
                Next j
            Next i
        Case 1      ' Golgi
            MaxHistos = GoX * GoY
            For i = 1 To MaxHistos
                For j = 1 To 1000
                    TempHistos(i, j) = Go_histo(i, j)
                Next j
            Next i
        Case 2      ' mossy
            MaxHistos = MFNUMBER
            For i = 1 To MaxHistos
                For j = 1 To 1000
                    TempHistos(i, j) = MF_histo(i, j)
                Next j
            Next i
        Case 3      ' climbing
            MaxHistos = 4
            For i = 1 To MaxHistos
                For j = 1 To 1000
                    TempHistos(i, j) = CF_histo(i, j)
                Next j
            Next i
        Case 4      ' nucleus
            MaxHistos = NCNUMBER
            For i = 1 To MaxHistos
                For j = 1 To 1000
                    TempHistos(i, j) = Nuc_histo(i, j)
                Next j
            Next i
        Case 5      ' Purkinje
            MaxHistos = PCNUMBER
            For i = 1 To MaxHistos
                For j = 1 To 1000
                    TempHistos(i, j) = Purk_histo(i, j)
                Next j
            Next i
        Case 6      ' Stellate
            MaxHistos = STELLATENUMBER
            For i = 1 To MaxHistos
                For j = 1 To 1000
                    TempHistos(i, j) = Stellate_histo(i, j)
                Next j
            Next i
        Case 7      ' basket
            MaxHistos = BasketNUMBER
            For i = 1 To MaxHistos
                For j = 1 To 1000
                    TempHistos(i, j) = Basket_histo(i, j)
                Next j
            Next i
    End Select
    DrawHistos
    
End Sub
Private Sub DrawHistos()
Dim X As Integer
Dim i As Integer
Dim d As Single
Dim n As Single

Dim y As Single

    Histo_form.Cls
    StopHisto = StartHisto + 9
    If StopHisto > MaxHistos Then StopHisto = MaxHistos
    Debug.Print MaxHistos, StartHisto, StopHisto
    y = 10
    For i = StartHisto To StopHisto
        y = y - 1
        For X = 1 To 1000
            n = TempHistos(i, X)
            n = n * 200
            d = TempHistoDivisor
            d = ScaleMax * d
            Histo_form.Line (X, y)-(X, y + (n / d))
        Next X
    Next i
    DoEvents


End Sub
