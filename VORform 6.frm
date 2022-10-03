VERSION 5.00
Begin VB.Form VORform 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Eye movement window"
   ClientHeight    =   11955
   ClientLeft      =   210
   ClientTop       =   2400
   ClientWidth     =   15060
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11955
   ScaleLeft       =   -250
   ScaleMode       =   0  'User
   ScaleWidth      =   15060
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   972
      Left            =   13440
      TabIndex        =   5
      Top             =   0
      Width           =   1572
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Examples"
         Height          =   492
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   972
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rasters"
         Height          =   372
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   120
         Width           =   972
      End
   End
   Begin VB.Label VORlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Image"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   336
      Index           =   5
      Left            =   0
      TabIndex        =   8
      Top             =   1560
      Width           =   852
   End
   Begin VB.Label VORlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Afferents"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   336
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   5760
      Width           =   1212
   End
   Begin VB.Label VORlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "PVPs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   336
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   3720
      Width           =   612
   End
   Begin VB.Label VORlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Eyes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   1890
      Width           =   615
   End
   Begin VB.Label VORlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Head"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   945
      Width           =   615
   End
   Begin VB.Label VORlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Target"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu VOR_diagnostics_menu 
      Caption         =   "&Diagnostics"
      Begin VB.Menu VOR_tests_menu 
         Caption         =   ".2 Hz Head stimulus"
         Index           =   0
      End
      Begin VB.Menu VOR_tests_menu 
         Caption         =   ".5 Hz Head stimulus"
         Index           =   1
      End
      Begin VB.Menu VOR_tests_menu 
         Caption         =   "1 Hz Head stimulus"
         Index           =   2
      End
      Begin VB.Menu VOR_tests_menu 
         Caption         =   "2 Hz Head stimulus"
         Index           =   3
      End
      Begin VB.Menu VOR_tests_menu 
         Caption         =   "4 Hz Head stimulus"
         Index           =   4
      End
      Begin VB.Menu VOR_tests_menu 
         Caption         =   "5 Hz Head stimulus"
         Index           =   5
      End
      Begin VB.Menu VOR_tests_menu 
         Caption         =   "10 Hz Head stimulus"
         Index           =   6
      End
      Begin VB.Menu VOR_tests_menu 
         Caption         =   "Lisberger and Pavelko step stimulus"
         Index           =   7
      End
      Begin VB.Menu VOR_tests_menu 
         Caption         =   "Lisberger and Pavelko CV test"
         Index           =   8
      End
   End
End
Attribute VB_Name = "VORform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isi_min(numVest_Afferents) As Single
Dim tonic(numVest_Afferents) As Single
Dim tonic_num(numVest_Afferents) As Single

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Debug.Print x, y
End Sub

Private Sub Form_Resize()
Dim i As Integer
VORform.ScaleTop = 10
VORform.ScaleHeight = -10
VORform.ScaleLeft = 1

VORform.ScaleWidth = 5250
VORform.Left = 1

For i = 0 To 2
    VORform.VORlabel(i).Top = 9.5 - i
Next i
VORform.VORlabel(3).Top = 5.7
VORform.VORlabel(4).Top = 2.1
VORform.VORlabel(5).Top = 7.7
End Sub

Private Sub VOR_tests_menu_Click(Index As Integer)
Dim x As Single
Dim pi As Single
Dim i As Single
Dim j As Integer

Dim head_p As Single
Dim head_v As Single
Dim head_a As Single
Dim head_p_last As Single
Dim head_v_last As Single

Dim eye_p_hold(5000) As Single
Dim eye_p As Single
Dim eye_p_last As Single
Dim eye_v As Single
Dim eye_p_f As Single

Dim image_v As Single

Dim mn_inE As Single
Dim mn_inI As Single

Dim eye_average(100) As Single
Dim eye_counter As Single
Dim spike_counter As Single
Dim temp_p As Single


Dim firing_rateL As Single
Dim firing_rateR As Single

Dim spike_timeL As Single
Dim spike_timeR As Single

Dim dynamic_index(numVest_Afferents) As Double
Dim CV(numVest_Afferents) As Single

Dim target_p As Single
Dim target_p_last As Single
Dim target_v As Single
Dim target_v_last As Single
Dim target_a As Single

Dim amplitude As Single

Dim max_p As Single
Dim max_v As Single
Dim max_a As Single

Dim Freq As Single
Dim v_cells As Integer

Dim PVP_cells As Integer
Dim PVP_counter As Integer
Dim PVP_in As Single
Dim temp_v As Single

Dim ISIs(numVest_Afferents) As Double
Dim ISIs_squared(numVest_Afferents) As Double
Dim ISI_num(numVest_Afferents) As Single

Dim vest_in_V As Single
Dim vest_in_a As Single

'VestibularForm.Visible = True
pi = 3.1412

Select Case Index
Case 0
    Freq = 0.2
Case 1
    Freq = 0.5
Case 2
    Freq = 1
Case 3
    Freq = 2
Case 4
    Freq = 4
Case 5
    Freq = 5
Case 6
    Freq = 10
End Select

amplitude = (Freq * ((0.5 / 0.565) / 19.91151)) * 5.6541 ' just kludged numbers to equate velocities

'Randomize Timer

VORform.Cls

target_p = 0
target_v = 0
target_a = 0
max_p = 0
max_v = 0
max_a = 0
temp_v = 0


For i = 1 To numVest_Afferents
    ISIs(i) = 0
    ISIs_squared(i) = 0
    ISI_num(i) = 0
    Vest_L(i).last_spike = 0
    Vest_R(i).last_spike = 0
    If Index = 7 Then isi_min(i) = 5000
Next i
eye_counter = 0

    'cbm_main.Label1(1).Caption = Timer - start_time
    start_time = Timer
   
    bincounter = Time_step_size
    bincounter5 = 1
    bincounter5_temp = 0
    
For i = 1 To 5000
    bincounter = i
    head_p_last = head_p
    
    If Index < 7 Then    ' sinusoidal stimuli
        head_p = (Cos((i / (500 / Freq)) * pi)) / amplitude
        
    ElseIf Index = 7 Then  '  Lisberger and Palveko step stimuli
        If i > 3000 Then
            If i <= 3300 Then
                If i <= 3050 Then
                    temp_v = temp_v + 0.0006
                ElseIf i > 3250 Then
                    temp_v = temp_v - 0.0006
                End If
                head_p = head_p + temp_v
            End If
            
            If i > 4000 Then
                If i <= 4300 Then
                    If i <= 4050 Then
                        temp_v = temp_v - 0.0006
                    ElseIf i > 4250 Then
                        temp_v = temp_v + 0.0006
                    End If
                End If
                head_p = head_p + temp_v
            End If
        End If
    Else
        head_p = 0
    End If
    
    If head_p > max_p Then max_p = head_p
    
        
    If i > 1 Then
        head_v_last = head_v
        head_v = head_p - head_p_last
        If head_v > max_v Then max_v = head_v
        
        
    End If
    
    
    If i > 2 Then
        head_a = head_v - head_v_last
        If head_a > max_a Then max_a = head_a
        VORform.PSet (i + 250, 8.5 + (head_a * 330)), RGB(255, 0, 255)
    End If
    
    
    
    '***************************************
    'Integrate and fire vestibular afferents
    '***************************************
    
    For v_cells = 1 To numVest_Afferents
      
      ' Left sensitive vestibular afferents
      
      'If head_v > 0 Then vest_in_V = head_v Else vest_in_V = 0
      'If head_a > 0 Then vest_in_a = head_a * 150 Else vest_in_a = 0
         
      If head_v > 0 Then vest_in_V = head_v Else vest_in_V = head_v * 0.5
      If head_a > 0 Then vest_in_a = head_a * 150 Else vest_in_a = head_a * 0.5
         
      Vest_L(v_cells).gVel = (vest_in_V * Vest_L(v_cells).gVelConst + (Vest_L(v_cells).gVelNoise * Rnd()) - Vest_L(v_cells).gVelNoise * Rnd()) * (1# - Vest_L(v_cells).gVel) + Vest_L(v_cells).gVel * Vest_L(v_cells).gVelDecayConst
      'If Vest_L(vcells).gVel < 0 Then Vest_L(v_cells).gVel = 0
      
      Vest_L(v_cells).gAcc = (vest_in_a * Vest_L(v_cells).gAccConst + (Vest_L(v_cells).gAccNoise * Rnd()) - Vest_L(v_cells).gAccNoise * Rnd()) * (1# - Vest_L(v_cells).gAcc) + Vest_L(v_cells).gAcc * Vest_L(v_cells).gAccDecayConst
      'If Vest_L(vcells).gAcc < 0 Then Vest_L(v_cells).gAcc = 0
      
      ' Right sensitive vestibular afferents
      
      If head_v < 0 Then vest_in_V = -head_v Else vest_in_V = -head_v * 0.5
      If head_a < 0 Then vest_in_a = -head_a * 150 Else vest_in_a = -head_a * 0.5
      
      Vest_R(v_cells).gVel = (vest_in_V * Vest_R(v_cells).gVelConst + (Vest_R(v_cells).gVelNoise * Rnd()) - Vest_R(v_cells).gVelNoise * Rnd()) * (1# - Vest_R(v_cells).gVel) + Vest_R(v_cells).gVel * Vest_R(v_cells).gVelDecayConst
      'If Vest_R(vcells).gVel < 0 Then Vest_R(v_cells).gVel = 0
      
      Vest_R(v_cells).gAcc = (vest_in_a * Vest_R(v_cells).gAccConst + (Vest_R(v_cells).gAccNoise * Rnd()) - Vest_R(v_cells).gAccNoise * Rnd()) * (1# - Vest_R(v_cells).gAcc) + Vest_R(v_cells).gAcc * Vest_R(v_cells).gAccDecayConst
      'If Vest_R(vcells).gAcc < 0 Then Vest_R(v_cells).gAcc = 0
      
    
      Vest_L(v_cells).V = Vest_L(v_cells).V + (GLeakVest * (ELeakVest - Vest_L(v_cells).V) - (Vest_L(v_cells).gVel * Vest_L(v_cells).V) - (Vest_L(v_cells).gAcc * Vest_L(v_cells).V))
      Vest_R(v_cells).V = Vest_R(v_cells).V + (GLeakVest * (ELeakVest - Vest_R(v_cells).V) - (Vest_R(v_cells).gVel * Vest_R(v_cells).V) - (Vest_R(v_cells).gAcc * Vest_R(v_cells).V))
      
      Vest_L(v_cells).act = Int(1 - ((Vest_L(v_cells).Thr - Vest_L(v_cells).V) * 0.001))
      Vest_R(v_cells).act = Int(1 - ((Vest_R(v_cells).Thr - Vest_R(v_cells).V) * 0.001))
      
      Vest_L(v_cells).Thr = (Vest_L(v_cells).act * ThrmaxVest) + ((1 - Vest_L(v_cells).act) * (Vest_L(v_cells).Thr + (ThrDecayVest * (ThrBaseVest - Vest_L(v_cells).Thr))))
      Vest_R(v_cells).Thr = (Vest_R(v_cells).act * ThrmaxVest) + ((1 - Vest_R(v_cells).act) * (Vest_R(v_cells).Thr + (ThrDecayVest * (ThrBaseVest - Vest_R(v_cells).Thr))))
      
    Next v_cells
    
    
    x = 0
    For v_cells = 1 To numVest_Afferents Step 10
        x = x + 1
        
            MF(x) = Vest_L(v_cells).act
        
        x = x + 1
        MF(x) = Vest_R(v_cells).act
    Next v_cells
    'dowork1
'**************************************************************************************
' Integrate and Fire PVP neurons ******************************************************
'**************************************************************************************
    For PVP_cells = 1 To numPVP
        PVP_in = 0
        
        For v_cells = PVP_cells To PVP_cells + 450 Step 9
            If v_cells > numVest_Afferents Then PVP_counter = v_cells - numVest_Afferents Else PVP_counter = v_cells
            PVP_in = PVP_in + Vest_R(PVP_counter).act
        Next v_cells
        
            PVP_in = PVP_in / 50#
            PVP_L(PVP_cells).gE = PVP_L(PVP_cells).gE + ((PVP_in * gEConstPVP * (1 - PVP_L(PVP_cells).gE)) - PVP_L(PVP_cells).gE * gEDecayPVP)
            PVP_L(PVP_cells).V = PVP_L(PVP_cells).V + (gLeakPVP * (ELeakPVP - PVP_L(PVP_cells).V)) - (PVP_L(PVP_cells).gE * (PVP_L(PVP_cells).V))
            PVP_L(PVP_cells).act = Int(1 - ((PVP_L(PVP_cells).Thr - PVP_L(PVP_cells).V) * 0.001))
            PVP_L(PVP_cells).Thr = (PVP_L(PVP_cells).act * ThrMaxPVP) + ((1 - PVP_L(PVP_cells).act) * (PVP_L(PVP_cells).Thr + (ThrDecayPVP * (ThrBasePVP - PVP_L(PVP_cells).Thr))))
        
        PVP_in = 0
        
        For v_cells = PVP_cells To PVP_cells + 450 Step 9
            If v_cells > numVest_Afferents Then PVP_counter = v_cells - numVest_Afferents Else PVP_counter = v_cells
            PVP_in = PVP_in + Vest_L(PVP_counter).act
        Next v_cells
        
            PVP_in = PVP_in / 50#
            PVP_R(PVP_cells).gE = PVP_R(PVP_cells).gE + ((PVP_in * gEConstPVP * (1 - PVP_R(PVP_cells).gE)) - PVP_R(PVP_cells).gE * gEDecayPVP)
            PVP_R(PVP_cells).V = PVP_R(PVP_cells).V + (gLeakPVP * (ELeakPVP - PVP_R(PVP_cells).V)) - (PVP_R(PVP_cells).gE * (PVP_R(PVP_cells).V))
            PVP_R(PVP_cells).act = Int(1 - ((PVP_R(PVP_cells).Thr - PVP_R(PVP_cells).V) * 0.001))
            PVP_R(PVP_cells).Thr = (PVP_R(PVP_cells).act * ThrMaxPVP) + ((1 - PVP_R(PVP_cells).act) * (PVP_R(PVP_cells).Thr + (ThrDecayPVP * (ThrBasePVP - PVP_R(PVP_cells).Thr))))
    Next PVP_cells

'***********************************
' Quasi-integrate and fire motor neurons
'***********************************
    mn_inE = 0
    mn_inI = 0
    For PVP_cells = 1 To numPVP
        mn_inE = mn_inE + (PVP_L(PVP_cells).act)
        mn_inI = mn_inI + (PVP_R(PVP_cells).act)
    Next PVP_cells
   
    'MN_L.ge = MN_L.ge + (mn_inE * 0.002 * (1 - MN_L.ge) - (MN_L.ge * 0.002))
    'MN_L.gi = MN_L.gi + (mn_inI * 0.002 * (1 - MN_L.gi) - (MN_L.gi * 0.02))
    'MN_L.v = MN_L.v + (0.2 * (0.002 * (-70 - MN_L.v)) - (MN_L.ge * MN_L.v) + (MN_L.gi * (-70 - MN_L.v)))
    
    spike_counter = 0
    
    For PVP_cells = 1 To numPVP
        spike_counter = spike_counter + PVP_L(PVP_cells).act - PVP_R(PVP_cells).act
        eye_p = eye_p + PVP_L(PVP_cells).act - PVP_R(PVP_cells).act
    Next PVP_cells
        eye_p_f = eye_p_f + 0.1 * (eye_p - eye_p_f)
        
        eye_p_hold(i) = eye_p_f
        
        If i > 30 Then
            For j = 0 To 4
                eye_v = eye_p_hold(i - j) - eye_p_hold(i - j - 20)
            Next j
        End If
        
        image_v = -(head_v * 15) - (eye_v / numVest_Afferents)
        
'****** PLOTS *****************************

    'VORform.PSet (i + 250, (eye_p / (numVest_Afferents * 10)) + 7.5), RGB(255, 0, 0)
    VORform.PSet (i + 250, (eye_p_f / (numVest_Afferents * 10)) + 7.5), RGB(255, 255, 255)
    VORform.PSet (i + 250, (eye_v / numVest_Afferents) + 7.5), RGB(0, 0, 255)
    
    VORform.PSet (i + 250, 7.5 - image_v), RGB(255, 0, 0)
    
    VORform.PSet (i + 250, ((head_p / 2) / 12.5) + 8.5), RGB(255, 255, 255)
    VORform.PSet (i + 250, 8.5 + (head_v * 15)), RGB(0, 0, 255)
    VORform.PSet (i + 250, 9.5 + target_p), RGB(0, 0, 0)
        
        
'***************************************************************
'Mostly calculation of firing rates of the Vestibular afferents
    
    For v_cells = 1 To numVest_Afferents
        If Vest_L(v_cells).act = 1 Then
            'VestibularForm.Line (i, (v_cells - 1) * 2)-(i, 1 + ((v_cells - 1) * 2)), RGB(0, 0, 0)
            If Index = 8 Then
                If Vest_L(v_cells).last_spike = 0 Then
                    Vest_L(v_cells).last_spike = i
                Else
                    ISIs(v_cells) = ISIs(v_cells) + (i - Vest_L(v_cells).last_spike)
                    ISI_num(v_cells) = ISI_num(v_cells) + 1
                    ISIs_squared(v_cells) = ISIs_squared(v_cells) + ((i - Vest_L(v_cells).last_spike) * (i - Vest_L(v_cells).last_spike))
                    Vest_L(v_cells).last_spike = i
                End If
            ElseIf Index = 7 Then
                If Vest_L(v_cells).last_spike = 0 Then
                    Vest_L(v_cells).last_spike = i
                Else
                    If i > 3000 And i < 3250 Then
                        If i - Vest_L(v_cells).last_spike < isi_min(v_cells) Then isi_min(v_cells) = i - Vest_L(v_cells).last_spike
                    End If
                    If i > 3150 And i < 3250 Then
                        tonic(v_cells) = tonic(v_cells) + i - Vest_L(v_cells).last_spike
                        tonic_num(v_cells) = tonic_num(v_cells) + 1
                    End If
                    Vest_L(v_cells).last_spike = i
                End If
            End If
        End If
    Next v_cells
    
    ' draw_examples
    If Option2.Value = True Then
        VORform.PSet (i + 250, (Vest_L(1).V + 65) / 50), RGB(0, 0, 0)
            If Vest_L(1).act = 1 Then
                VORform.Line (i + 250, (Vest_L(1).V + 65) / 50)-(i + 250, (Vest_L(1).V + 105) / 50), RGB(0, 0, 0)
                'If spike_timeL = 0 Then
                '    spike_timeL = i
                'Else
                '    firing_rateL = 1000 / (i - spike_timeL)
                '    spike_timeL = i
                'End If
            End If
        'If i > 50 Then VORform.PSet (i + 250, 0.7 + (firing_rateL / 150#)), RGB(0, 0, 0)
        
        VORform.PSet (i + 250, (Vest_R(1).V + 115) / 50), RGB(0, 0, 255)
        If Vest_R(1).act = 1 Then
            VORform.Line (i + 250, (Vest_R(1).V + 115) / 50)-(i + 250, (Vest_R(1).V + 155) / 50), RGB(0, 0, 255)
            'If spike_timeR = 0 Then
            '    spike_timeR = i
            'Else
            '    firing_rateR = 1000 / (i - spike_timeR)
            '    spike_timeR = i
            'End If
        End If
        'If i > 50 Then VORform.PSet (i + 250, 2.1 + (firing_rateR / 150#)), RGB(0, 0, 255)
        
        VORform.PSet (i + 250, (Vest_L(numVest_Afferents).V + 165) / 50), RGB(0, 0, 0)
            If Vest_L(numVest_Afferents).act = 1 Then
                VORform.Line (i + 250, (Vest_L(numVest_Afferents).V + 165) / 50)-(i + 250, (Vest_L(numVest_Afferents).V + 205) / 50), RGB(0, 0, 0)
            End If
        
        VORform.PSet (i + 250, (Vest_R(numVest_Afferents).V + 215) / 50), RGB(0, 0, 255)
        If Vest_R(numVest_Afferents).act = 1 Then
            VORform.Line (i + 250, (Vest_R(numVest_Afferents).V + 215) / 50)-(i + 250, (Vest_R(numVest_Afferents).V + 255) / 50), RGB(0, 0, 255)
        End If
        
        
    ElseIf Option1(0).Value = True Then
        For v_cells = 1 To 100
            If Vest_L(v_cells).act = 1 Then VORform.PSet (i + 250, v_cells * 0.025), RGB(0, 0, 0)
            If Vest_R(v_cells).act = 1 Then VORform.PSet (i + 250, 2.55 + v_cells * 0.025), RGB(0, 0, 255)
        Next v_cells
        
    End If
    
    VORform.PSet (i + 250, (PVP_L(1).V + 290) / 50), RGB(0, 0, 255)
    If PVP_L(1).act = 1 Then
        VORform.Line (i + 250, (PVP_L(1).V + 290) / 50)-(i + 250, (PVP_L(1).V + 330) / 50), RGB(0, 0, 255)
    End If
    
    VORform.PSet (i + 250, (PVP_R(1).V + 340) / 50), RGB(0, 0, 0)
    If PVP_R(1).act = 1 Then
        VORform.Line (i + 250, (PVP_R(1).V + 340) / 50)-(i + 250, (PVP_R(1).V + 380) / 50), RGB(0, 0, 0)
    End If
    
    'VORform.PSet (i + 250, (MN_L.v + 400) / 50), RGB(0, 0, 0)
    DoEvents
    
Next i
    If Index = 8 Then
        GWin.Left = 1125
        GWin.Top = 4965
        GWin.Height = 7500
        GWin.Width = 8205
        GWin.ScaleLeft = 0
        GWin.ScaleTop = 0
        GWin.ScaleHeight = 7095
        GWin.ScaleWidth = 8085
        GWin.BackColor = vbWhite
        GWin.Cls
        GWin.Caption = "Lisberger & Pavelko 1986 Figure 6"
        GWin.ScaleWidth = 1
        GWin.ScaleHeight = -8
        GWin.ScaleTop = 8
        GWin.DrawWidth = 3
        GWin.Line (0.02, 1)-(1, 1), RGB(0, 0, 0)
        GWin.Line (0.02, 1)-(0.02, 8), RGB(0, 0, 0)
        GWin.DrawWidth = 1
        For i = 1 To 8
            GWin.Line (0.02, i)-(1, i), RGB(0, 0, 0)
        Next i
        For i = 0.02 To 1# Step 0.2
            GWin.Line (i, 1)-(i, 8), RGB(0, 0, 0)
        Next i
        GWin.Visible = True
        
        For v_cells = 1 To numVest_Afferents
            dynamic_index(v_cells) = 1000 / isi_min(v_cells)
            dynamic_index(v_cells) = dynamic_index(v_cells) - (1000 / (ISIs(v_cells) / ISI_num(v_cells)))
            tonic(v_cells) = 1000 / ((tonic(v_cells) / tonic_num(v_cells)))
            dynamic_index(v_cells) = dynamic_index(v_cells) / (tonic(v_cells) - (1000 / (ISIs(v_cells) / ISI_num(v_cells))))
        
            CV(v_cells) = Sqr(((ISI_num(v_cells) * ISIs_squared(v_cells)) - (ISIs(v_cells) * ISIs(v_cells))) / (ISI_num(v_cells) * (ISI_num(v_cells) - 1))) / (ISIs(v_cells) / ISI_num(v_cells))
            'Debug.Print v_cells, 1000 / (ISIs(v_cells) / ISI_num(v_cells)), 1000 / isi_min(v_cells), tonic(v_cells), CV(v_cells), dynamic_index(v_cells)
            GWin.Line (CV(v_cells), dynamic_index(v_cells))-(CV(v_cells) + 0.01, 0.1 + dynamic_index(v_cells)), RGB(0, 0, 0), BF
        Next v_cells
    End If
End Sub


