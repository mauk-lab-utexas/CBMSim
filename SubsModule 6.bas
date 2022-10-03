Attribute VB_Name = "SubsModule"
Public ArgArray(25) As String
Public NumArgs As Integer

Public Sub Calculate_Time_Dependent_variables()
    gr_TDV
    Gol_TDV
    MF_TDV
    Stellate_TDV
    BC_TDV
    PC_TDV
    NUC_TDV
    CF_TDV
    Vest_TDV
    RN_TDV
    STP_TDV
End Sub
Public Sub LoadTrial(trial As Integer)
Dim i As Integer
    For i = 1 To 4
        CS_ON(i) = RunData(trial).CSon(i)
        cs_duration(i) = RunData(trial).CSduration(i)
        cs_onset(i) = RunData(trial).CSonset(i)
    Next i
    Debug.Print cs_onset(3)
    US1_ON = RunData(trial).USon(1)
    US2_ON = RunData(trial).USon(2)
    For i = 1 To 2
        US_onset(i) = RunData(trial).USonset(i)
    Next i
    
    If CS_ON(1) = 1 Then cbm_main.CS1.Value = Checked Else cbm_main.CS1.Value = Unchecked
    If CS_ON(2) = 1 Then cbm_main.CS2.Value = Checked Else cbm_main.CS2.Value = Unchecked
    If CS_ON(3) = 1 Then cbm_main.CS3.Value = Checked Else cbm_main.CS3.Value = Unchecked
    If CS_ON(4) = 1 Then cbm_main.CS4.Value = Checked Else cbm_main.CS4.Value = Unchecked
   
    If US1_ON = 1 Then cbm_main.US1.Value = Checked Else cbm_main.US1.Value = Unchecked
    If US2_ON = 1 Then cbm_main.US2.Value = Checked Else cbm_main.US2.Value = Unchecked
        
    cbm_main.OnsetLabels(0) = Str(cs_onset(1))
    cbm_main.OnsetLabels(2) = Str(cs_onset(2))
    cbm_main.OnsetLabels(4) = Str(cs_onset(3))
    cbm_main.OnsetLabels(6) = Str(cs_onset(4))
    
    cbm_main.OnsetLabels(1) = Str(cs_duration(1))
    cbm_main.OnsetLabels(3) = Str(cs_duration(2))
    cbm_main.OnsetLabels(5) = Str(cs_duration(3))
    cbm_main.OnsetLabels(7) = Str(cs_duration(4))
    
    cbm_main.OnsetLabels(8) = Str(US_onset(1))
    cbm_main.OnsetLabels(9) = Str(US_onset(2))
    

End Sub
Public Sub ShowWeights()
Dim n As Integer
Dim p As Integer
Dim q As Integer
Dim count As Single

    PM_Form.ScaleTop = 1#
    PM_Form.ScaleHeight = -1#
    PM_Form.ScaleLeft = 0
    PM_Form.ScaleWidth = 220
    PM_Form.Line (0, 1)-(190, 0), vbBlack, BF
    PM_Form.Line1.BorderColor = vbWhite
    PM_Form.Line2.BorderColor = vbRed
    PM_Form.Line1.X1 = 0
    PM_Form.Line1.X2 = 80
    PM_Form.Line1.y1 = gPurktoNucBeginAverage
    PM_Form.Line1.Y2 = gPurktoNucBeginAverage
    PM_Form.Line2.X1 = 80
    PM_Form.Line2.X2 = 90
    PM_Form.Line2.y1 = gNuctoCFBeginAverage
    PM_Form.Line2.Y2 = gNuctoCFBeginAverage
    PM_Form.Line (90, gPurktoNucBeginAverage)-(220, gPurktoNucBeginAverage), vbWhite
    
    For n = 1 To NCNUMBER
        For p = 1 To PCNCSYNUMBER
            PM_Form.Line ((13 * (n - 1)) + p - 0.7, 0)-((13 * (n - 1)) + p, gPURKtoNUCLEUS(Nc(n).PCsyn(p), n)), vbRed, BF
        Next p
    Next n
'    For n = 1 To NumCF
'        For p = 1 To 3
'            PM_Form.Line (80 + p - 0.7, 0)-(80 + p, Nc(p).gNUCtoCF(n) * 10), vbWhite, BF
'            'Debug.Print n, Nc(n).gNUCtoCF
'        Next p
'    Next n
'
'    count = 0
'    For q = 1 To PCNUMBER
'        For n = 1 To NCNUMBER
'            For p = 1 To PCNCSYNUMBER
'                If Nc(n).PCsyn(p) = q Then
'                    PM_Form.Line (90.7 + (count * 1.5), 0)-(90 + (count * 1.5), gPURKtoNUCLEUS(q, n)), vbGreen, BF
'                    count = count + 0.85
'                End If
'            Next p
'        Next n
'        count = count + 1
'    Next q
End Sub
Public Sub MakeBVIFile()
Dim f As String
Dim i As Integer
    f = Format(Date, "mm/dd/yyyy")
    Print #25, f
    Print #25, SessionNames(TrialsPerSession(0, 1))
    Print #25, 1
    Print #25, TrialsPerSession(TrialsPerSession(0, 1), 1)
    Print #25, 0
    Print #25, "NONE"
    Print #25, 0
    Print #25, "NONE"
    Print #25, 0
    Print #25, "NONE"
    Print #25, 0
    Print #25, "NONE"
    For i = TrialThatEndsSession - TrialsPerSession(TrialsPerSession(0, 1), 1) + 1 To TrialThatEndsSession
        Print #25, MasterTrialNames(i)
        
        If cbm_main.CompressedDatFile_Menu.Checked = False Then    ' normal data file mode
            Print #25, 2500, 0, RunData(i).CSon(1), 200, 200 + RunData(i).CSduration(1), 1000, RunData(i).CSon(2), 200, 200 + RunData(i).CSduration(2), 1000, RunData(i).USon(1),
            If RunData(i).USon(1) = 0 Then Print #25, 2000, 2100, Else Print #25, RunData(i).USonset(1) - RunData(i).CSonset(1) + 200, RunData(i).USonset(1) - RunData(i).CSonset(1) + 210,
            Print #25, RunData(i).USon(2),
            If RunData(i).USon(2) = 0 Then Print #25, 2000, 2000 Else Print #25, RunData(i).USonset(2), RunData(i).USonset(2) - RunData(i).CSonset(1) + 200, RunData(i).USonset(2) - RunData(i).CSonset(1) + 210
        Else
            Print #25, 2500, 0, RunData(i).CSon(1), 200, 200 + (RunData(i).CSduration(1) / 2), 1000, RunData(i).CSon(2), 200, 200 + (RunData(i).CSduration(2) / 2), 1000, RunData(i).USon(1), ((RunData(i).USonset(1) - RunData(i).CSonset(1)) / 2) + 200, ((RunData(i).USonset(1) - RunData(i).CSonset(1)) / 2) + 210, RunData(i).USon(2), RunData(i).USonset(2), ((RunData(i).USonset(2) - RunData(i).CSonset(1)) / 2) + 200, ((RunData(i).USonset(2) - RunData(i).CSonset(1)) / 2) + 210
        End If

    Next i
    Close #25
End Sub
Public Sub ReadyToRepeat()
Dim tempS As String
Dim tempLength As Integer
Dim tempN As Integer
cbm_main.Number_of_trials_per_session.Text = TCounter + TrialCounter
    NO_OF_TRIALS_GIVEN_BY_USER = Int(cbm_main.Number_of_trials_per_session.Text)
    cbm_main.StatusBar1.Panels(3).Text = MasterTrialNames(1)
    TrialThatEndsSession = TrialsPerSession(1, 1)
    TrialsPerSession(0, 1) = 1  'this stores the session number
    SessionsThisExp = 1
    CurrentContext = SessionContexts(1)
    STPFadeSession = STPFades(1)
'    cbm_main.speed_menu.Enabled = False
'    cbm_main.pause_menu.Enabled = False
'    cbm_main.resume_menu.Enabled = False
    plasticity.Visible = False
    histo_switch = 1
    
'    cbm_main.resume_menu.Enabled = True
    HistoDivisor = 1
    Erase GR_histo
    Erase Go_histo
    Erase Purk_histo
    Erase Nuc_histo
    Erase CF_histo
    Erase MF_histo
    Erase Stellate_histo
    Erase Basket_histo
    Erase response_histo
    Erase RN_Histo
    Erase jav_Mfinputaverage
    Erase Rasters
    Erase gr_PC_weights_rasters
    Erase mf_nuc_weights_rasters
    Erase Raster_list
    Erase Raster_plasticity
    Erase rasters_MF_plasticity
    PC_form.Cls
    OpenDatFile
    TrialsThisSession = 1
    tempS = CStr(NumRepeats)
    If cbm_main.Check2.Value = Checked Then
        
        If NumRepeats < 10 Then
            OutputFilename = OutputBaseFilename + "000" + tempS + ".txt"
            CFoutputFilename = CFoutputBaseFilename + "000" + tempS + ".txt"
            SIMoutputFilename = OutputBaseFilename + "000" + tempS + ".cbm"
        ElseIf NumRepeats < 100 Then
            OutputFilename = OutputBaseFilename + "00" + tempS + ".txt"
            CFoutputFilename = CFoutputBaseFilename + "00" + tempS + ".txt"
            SIMoutputFilename = OutputBaseFilename + "00" + tempS + ".cbm"
        ElseIf NumRepeats < 1000 Then
            OutputFilename = OutputBaseFilename + "0" + tempS + ".txt"
            CFoutputFilename = CFoutputBaseFilename + "0" + tempS + ".txt"
            SIMoutputFilename = OutputBaseFilename + "0" + tempS + ".cbm"
        Else
            OutputFilename = OutputBaseFilename + tempS + ".txt"
            CFoutputFilename = CFoutputBaseFilename + tempS + ".txt"
            SIMoutputFilename = OutputBaseFilename + tempS + ".cbm"
        End If
        SaveActivityData
    End If
    Trials_this_time = 1
    If NumRepeats < 10 Then
            SIMoutputFilename = OutputBaseFilename + "000" + tempS + ".cbm"
        ElseIf NumRepeats < 100 Then
            SIMoutputFilename = OutputBaseFilename + "00" + tempS + ".cbm"
        ElseIf NumRepeats < 1000 Then
            SIMoutputFilename = OutputBaseFilename + "0" + tempS + ".cbm"
        Else
            SIMoutputFilename = OutputBaseFilename + tempS + ".cbm"
        End If
    NumRepeats = NumRepeats + 1
    Open SIMoutputFilename For Binary As #1
        SaveSim
    RepeatExperimentsCounter = RepeatExperimentsCounter + 1
'    LoadTrial (1)
'    cbm_main.resume_menu_Click
End Sub
Public Sub SaveWeights()
Dim p As Integer
Dim i As Long
Dim j As Integer
Dim avg As Single
Dim count As Single

'PC to nuc weights arranged in rows by PC (24 rows, 6 columns)
    For p = 1 To PCNUMBER
        avg = 0
        count = 0
        For i = 1 To NCNUMBER
            For j = 1 To PCNCSYNUMBER
                If Nc(i).PCsyn(j) = p Then
                    avg = avg + gPURKtoNUCLEUS(p, i)
                    count = count + 1
                End If
            Next j
        Next i
        PNWeightsBYPURK(p, Trials_this_time) = avg / count
    Next p
    
    'PC to Nuc weights arranged in rows by nucleus cell (6 rows, 10 columns)
    For i = 1 To NCNUMBER
        avg = 0
        count = 0
        For p = 1 To PCNUMBER
            For j = 1 To PCNCSYNUMBER
                If Nc(i).PCsyn(j) = p Then
                    avg = avg + gPURKtoNUCLEUS(p, i)
                    count = count + 1
                End If
            Next j
        Next p
        PNWeightsBYNUC(i, Trials_this_time) = avg / count
    Next i
    
    For i = 1 To SYNUMBER
        grPCWeights(Trials_this_time) = grPCWeights(Trials_this_time) + grWeight(i)
    Next i
End Sub
Public Sub WeightHistoryDraw()
Dim avg As Single
Dim i As Long
Dim j As Integer
Dim counter As Integer
    avg = 0
    For i = 1 To SYNUMBER
        avg = avg + grWeight(i)
    Next i
    avg = avg / 20000
    WeightHistory.PSet (Trials_this_time, avg), vbWhite
    avg = 0
    counter = 0
    For i = 1 To NCNUMBER
        For j = 1 To PCNCSYNUMBER
            counter = counter + 1
            avg = avg + gPURKtoNUCLEUS(Nc(i).PCsyn(j), i)
        Next j
    Next i
    avg = avg / counter
    WeightHistory.PSet (Trials_this_time, avg), vbRed
    avg = 0
    For i = 1 To NCNUMBER
        For j = 1 To NumCF
            avg = avg + Nc(i).gNUCtoCF(j)
        Next j
    Next i
    avg = avg / NCNUMBER
    WeightHistory.PSet (Trials_this_time, avg * 10), vbYellow
End Sub
Public Sub SaveGrWeights()
    If grWEIGHTS_filename = "" Then
        grWEIGHTS_filename = "gr" + Str$(Timer) + ".wts"
        Close #13
        Open grWEIGHTS_filename For Binary As #13
        mfWEIGHTS_filename = "mf" + Str$(Timer) + ".wts"
        Close #14
        Open mfWEIGHTS_filename For Binary As #14
    End If
    Put #13, , grWeight
    Put #14, , mfweight
End Sub
Public Sub ResetTimers()
    Debug_gran = 0
    Debug_Gol = 0
    Debug_BS = 0
    
    TIME_MF = 0
    TIME_GR = 0
    TIME_GOL = 0
    TIME_BK = 0
    TIME_PK = 0
    TIME_NUC = 0
    TIME_Plasticity = 0
End Sub

Public Sub SaveActivityData()
Dim i As Integer
Dim j As Integer
Dim p As Integer
Dim n As Integer
Dim q As Integer

Dim avg As Single

    Close #13
    Open OutputFilename For Append As #13
    
    For j = 1 To Trials_this_time
        Write #13, j,
        
        For i = 1 To PCNUMBER
            Write #13, PurkinjeActivity(i, j),
        Next i
        
        For i = 1 To NCNUMBER
            Write #13, NucleusActivity(i, j),
        Next i
        
        For i = 1 To NumCF
            Write #13, ClimbingFiberActivity(j, i),
        Next i
        
        Write #13, MFActivity(j),
        Write #13, GranuleActivity(j),
        Write #13, GolgiActivity(j),
        
        For q = 1 To PCNUMBER
            For n = 1 To NCNUMBER
                For p = 1 To PCNCSYNUMBER
                    If Nc(n).PCsyn(p) = q Then
                        Write #13, gPURKtoNUCLEUS(q, n),
                    End If
                Next p
            Next n
        Next q
        
        Write #13, grPCWeights(j),
        Write #13,
        
    Next j
    Close #13
    Open CFoutputFilename For Append As #13
        For j = 1 To 1000
            For i = 1 To NumCF
                For p = 1 To 16
                    If CFactivityTimes(j, i, p) <> 0 Then
                        Write #13, CFactivityTimes(j, i, p),
                    Else
                        Write #13, -1,
                        p = 16
                    End If
                Next p
            Next i
            Write #13,
        Next j
        
    Close #13
End Sub

Public Sub SwapExperiments()
Dim i As Integer

    For i = 0 To 1000
        RunDataTEMP(i) = RunData(i)
        RunData(i) = RunData2(i)
        RunData2(i) = RunDataTEMP(i)
        
        MasterTrialNamesTEMP(i) = MasterTrialNames(i)
        MasterTrialNames(i) = MasterTrialNames2(i)
        MasterTrialNames2(i) = MasterTrialNamesTEMP(i)
    Next i
    For i = 0 To 25
        TrialsPerSessionTEMP(i, 1) = TrialsPerSession(i, 1)
        TrialsPerSession(i, 1) = TrialsPerSession(i, 2)
        TrialsPerSession(i, 2) = TrialsPerSessionTEMP(i, 1)
        
        SessionContextsTEMP(i) = SessionContexts(i)
        SessionContexts(i) = SessionContexts2(i)
        SessionContexts2(i) = SessionContextsTEMP(i)
        
        STPFadesTEMP(i) = STPFades(i)
        STPFades(i) = STPFades2(i)
        STPFades2(i) = STPFadesTEMP(i)
        
        SessionNamesTEMP(i) = SessionNames(i)
        SessionNames(i) = SessionNames2(i)
        SessionNames2(i) = SessionNamesTEMP(i)
    Next i
    i = NumberofSessions
    NumberofSessions = NumberofSessions2
    NumberofSessions2 = i

End Sub
Public Sub GetCommandLine()
    Dim c, CmdLineLen, InArg
    Dim CmdLine As String
    Dim i As Integer
    
    CmdLine = Command$
    
    CmdLineLen = Len(CmdLine)
    
    For i = 1 To CmdLineLen
        c = (Mid(CmdLine, i, 1))
        If c <> " " Then
            If Not InArg Then
                NumArgs = NumArgs + 1
                InArg = True
            End If
            ArgArray(NumArgs) = ArgArray(NumArgs) & c
        Else
            InArg = False
        End If
    Next i
    
End Sub
Public Sub SaveSim()
        Put #1, , MF
        Put #1, , Gr
        Put #1, , Gol
        Put #1, , MFS
        Put #1, , Pc
        Put #1, , Bk
        Put #1, , BCells
        Put #1, , Nc
        Put #1, , CF
        Put #1, , NumCF
        Put #1, , grWeight
        Put #1, , mfweight
        Put #1, , gPURKtoNUCLEUS
        Put #1, , StellPcG
        
        Put #1, , MfElig
        Put #1, , NcCfG
        Put #1, , runavpc
        Put #1, , runavpc2
        Put #1, , gr_elig
        
        Put #1, , response_counter
        Put #1, , TrialCounter
        
        Put #1, , Bincounter
        Put #1, , HistoDivisor
        Put #1, , CSisPresent
        Put #1, , csnumber
        Put #1, , ltdtime
        Put #1, , ltdperiod
        Put #1, , ltpperiod
        Put #1, , mfltptime
        Put #1, , mfltpperiod
        Put #1, , mfltdperiod
        Put #1, , pcrunavcounter
        Put #1, , hyper_switch
        Put #1, , pcspknow
        Put #1, , pcrunav
        Put #1, , Time_step_size
        Put #1, , cs_onset
        Put #1, , cs_duration
        Put #1, , US_onset
        Put #1, , CS_ON
        
        Put #1, , US1_ON
        Put #1, , US2_ON
        Put #1, , UseUBCs
        
        Put #1, , PCHomeoValue
        Put #1, , NCHomeoValue
        
        Put #1, , PCNCSYNUMBER
        Put #1, , grPreElig
        Put #1, , grSTP
        Put #1, , grPreEligStatus
        Put #1, , grPreEligTimeout
        
        Put #1, , MFBGROUNDFREQMIN_CS
        Put #1, , MFBGROUNDFREQMAX_CS
        
        Put #1, , MFCONTEXTFREQMIN
        Put #1, , MFCONTEXTFREQMAX
        
        Put #1, , MFCONTEXTFREQMIN2
        Put #1, , MFCONTEXTFREQMAX2
        
        Put #1, , MFBGROUNDFREQMIN
        Put #1, , MFBGROUNDFREQMAX
        
        Put #1, , MFTONICFREQ_INCREMENT
        Put #1, , MFTONICFREQ_INCREMENT2
        Put #1, , MFTONICFREQ_INCREMENT3
        Put #1, , MFTONICFREQ_INCREMENT4
        
        Put #1, , MFPHASICFREQ_INCREMENT
        Put #1, , MFPHASICFREQ_INCREMENT2
        Put #1, , MFPHASICFREQ_INCREMENT3
        Put #1, , MFPHASICFREQ_INCREMENT4
        Put #1, , GG
        Put #1, , gGG
        Put #1, , wGG
        
        Close #1
End Sub
Public Sub GetSim()

        Get #1, , MF
        Get #1, , Gr
        Get #1, , Gol
        Get #1, , MFS
        Get #1, , Pc
        Get #1, , Bk
        Get #1, , BCells
        Get #1, , Nc
        Get #1, , CF
        Get #1, , NumCF
        Get #1, , grWeight
        Get #1, , mfweight
        Get #1, , gPURKtoNUCLEUS
        Get #1, , StellPcG
        
        Get #1, , MfElig
        Get #1, , NcCfG
        Get #1, , runavpc
        Get #1, , runavpc2
        Get #1, , gr_elig
        
        Get #1, , response_counter
        Get #1, , TrialCounter
        
        Get #1, , Bincounter
        Get #1, , HistoDivisor
        Get #1, , CSisPresent
        Get #1, , csnumber
        Get #1, , ltdtime
        Get #1, , ltdperiod
        Get #1, , ltpperiod
        Get #1, , mfltptime
        Get #1, , mfltpperiod
        Get #1, , mfltdperiod
        Get #1, , pcrunavcounter
        Get #1, , hyper_switch
        Get #1, , pcspknow
        Get #1, , pcrunav
        Get #1, , Time_step_size
        Get #1, , cs_onset
        Get #1, , cs_duration
        Get #1, , US_onset
        Get #1, , CS_ON
        Get #1, , US1_ON
        Get #1, , US2_ON
        Get #1, , UseUBCs
        
        Get #1, , PCHomeoValue
        Get #1, , NCHomeoValue
        
        Get #1, , PCNCSYNUMBER
        Get #1, , grPreElig
        Get #1, , grSTP
        Get #1, , grPreEligStatus
        Get #1, , grPreEligTimeout
        
        Get #1, , MFBGROUNDFREQMIN_CS
        Get #1, , MFBGROUNDFREQMAX_CS
        
        Get #1, , MFCONTEXTFREQMIN
        Get #1, , MFCONTEXTFREQMAX
        
        Get #1, , MFCONTEXTFREQMIN2
        Get #1, , MFCONTEXTFREQMAX2
        
        Get #1, , MFBGROUNDFREQMIN
        Get #1, , MFBGROUNDFREQMAX
        
        Get #1, , MFTONICFREQ_INCREMENT
        Get #1, , MFTONICFREQ_INCREMENT2
        Get #1, , MFTONICFREQ_INCREMENT3
        Get #1, , MFTONICFREQ_INCREMENT4
        
        Get #1, , MFPHASICFREQ_INCREMENT
        Get #1, , MFPHASICFREQ_INCREMENT2
        Get #1, , MFPHASICFREQ_INCREMENT3
        Get #1, , MFPHASICFREQ_INCREMENT4
        Get #1, , GG
        Get #1, , gGG
        Get #1, , wGG
        
'        Get #1, , MF
'        Get #1, , Gr
'        Get #1, , Gol
'        Get #1, , MFS
'        Get #1, , Pc
'        Get #1, , Bk
'        Get #1, , BCells
'        Get #1, , Nc
'        Get #1, , CF
'        Get #1, , NumCF
'        Get #1, , grweight
'        Get #1, , mfweight
'        Get #1, , gPURKtoNUCLEUS
'        Get #1, , StellPcG
'
'        Get #1, , MfElig
'        Get #1, , NcCfG
'        Get #1, , runavpc
'        Get #1, , runavpc2
'        Get #1, , gr_elig
'
'        Get #1, , response_counter
'        Get #1, , TrialCounter
'
'        Get #1, , bincounter
'        Get #1, , HistoDivisor
'        Get #1, , CSisPresent
'        Get #1, , csnumber
'        Get #1, , ltdtime
'        Get #1, , ltdperiod
'        Get #1, , ltpperiod
'        Get #1, , mfltptime
'        Get #1, , mfltpperiod
'        Get #1, , mfltdperiod
'        Get #1, , pcrunavcounter
'        Get #1, , hyper_switch
'        Get #1, , pcspknow
'        Get #1, , pcrunav
'        Get #1, , Time_step_size
'        Get #1, , cs_onset
'        Get #1, , cs_duration
'        Get #1, , US_onset
'        Get #1, , CS_ON
'
'        Get #1, , US1_ON
'        Get #1, , US2_ON
'        Get #1, , UseUBCs
'
'        Get #1, , PCHomeoValue
'        Get #1, , NCHomeoValue
'
'        Get #1, , PCNCSYNUMBER
'        Get #1, , grPreElig
'        Get #1, , grSTP
'        Get #1, , grPreEligStatus
'        Get #1, , grPreEligTimeout
'
'        Get #1, , MFBGROUNDFREQMIN_CS
'        Get #1, , MFBGROUNDFREQMAX_CS
'
'        Get #1, , MFCONTEXTFREQMIN
'        Get #1, , MFCONTEXTFREQMAX
'
'        Get #1, , MFCONTEXTFREQMIN2
'        Get #1, , MFCONTEXTFREQMAX2
'
'        Get #1, , MFBGROUNDFREQMIN
'        Get #1, , MFBGROUNDFREQMAX
'
'        Get #1, , MFTONICFREQ_INCREMENT
'        Get #1, , MFTONICFREQ_INCREMENT2
'        Get #1, , MFTONICFREQ_INCREMENT3
'        Get #1, , MFTONICFREQ_INCREMENT4
'
'        Get #1, , MFPHASICFREQ_INCREMENT
'        Get #1, , MFPHASICFREQ_INCREMENT2
'        Get #1, , MFPHASICFREQ_INCREMENT3
'        Get #1, , MFPHASICFREQ_INCREMENT4
End Sub
