Attribute VB_Name = "cbm_mainModule"

Option Explicit


Public peakcr As Single
Public cspresent_previous As Integer
Public MFtoGr As Single
Public MFtoGo As Single
Public GRtoGo As Single
Public GOtoGr As Single
Public auto_filename As String
Public MFgrGolOnly As Integer

Public CS_counter As Integer
Public NO_OF_TRIALS_GIVEN_BY_USER As Long
Public filename As String
Public OutputFilename As String
Public CFoutputFilename As String
Public OutputBaseFilename As String
Public CFoutputBaseFilename As String

Public SIMoutputFilename As String

Public NumRepeats As Integer

Public AmpModeAmp As Single
Public AmpMode As Integer
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

'******************* HVoicu - Stop after a given number of trials provided by the user -  used instead of 999
Public Sub DoWork(grancells As Integer, golgicells As Integer, StellateCells As Integer, PCcells As Integer)

Dim inAMPA As Single
Dim inNMDA As Single
Dim inI As Single
Dim inMF As Integer
Dim inGr As Single
Dim inGol As Single
Dim inGol2 As Single
Dim inPC As Integer
Dim X As Integer
Dim y As Integer
Dim y1 As Integer
Dim i As Integer
Dim i2 As Single
Dim m As Integer
Dim n As Integer
Dim syn As Integer
Dim counter As Single
Dim x_draw As Integer
Dim bkactive As Single
Dim grchange As Single
Dim Debug_weights As Single
Dim gr_elig_temp As Single
Dim gr_counter_TEMP As Integer
Dim mf_change As Single
Dim total_activity As Single
Dim raster_clear As Integer
Dim S1 As Integer
Dim S2 As Integer
Dim r As Single
Dim Z As Integer
Dim MFW As Single
Dim Rmax As Single
Dim T1 As Single
Dim T2 As Single

Dim Long1 As Long
Dim Long2 As Long

Dim count As Integer
Dim grCount As Integer

    bkactive = 0
    For X = 1 To PCNUMBER
        Pc(X).GGr = 0
        Pc(X).GStell = 0
    Next X
    For X = 1 To STELLATENUMBER
        Bk(X).GGr = 0
    Next X
    For X = 1 To NCNUMBER
        Nc(X).gPc = 0
        Nc(X).gMF = 0
    Next X
    For X = 1 To NumCF
        CF(X).GNc = 0
    Next X
  
    inI = 0

'**************************GRANULE CELLS *************************************************************
    jav_grspktotal = 0

        For X = 1 To grancells
            inAMPA = (MF2(Gr(X).MF(1)).gGrAMPA + MF2(Gr(X).MF(2)).gGrAMPA + MF2(Gr(X).MF(3)).gGrAMPA + MF2(Gr(X).MF(4)).gGrAMPA) / 4#
            inNMDA = (MF2(Gr(X).MF(1)).gGrNMDA + MF2(Gr(X).MF(2)).gGrNMDA + MF2(Gr(X).MF(3)).gGrNMDA + MF2(Gr(X).MF(4)).gGrNMDA) / 4#
           
'            Gr(X).gi = 0
'            For i = 1 To 4
'                Gr(X).gi = Gr(X).gi + Gol(Gr(X).Gol(i)).gSlow
'            Next i
'            For i = 1 To 4
'                If Gr(X).fastMask(i) = 1 Then Gr(X).gi = Gr(X).gi + Gol(Gr(X).Gol(i)).gFastFinal
'            Next i
                 
            Gr(X).v = Gr(X).v + (((gLGr + (Gr(X).g_KCa * Gr(X).g_KCa * Gr(X).g_KCa * Gr(X).g_KCa)) * (ELeakgr - Gr(X).v) - (inAMPA * Gr(X).v) - (inNMDA * Gr(X).v) + Gr(X).gi * (EGABAgr - Gr(X).v))) '  /* EE=0 */
            Gr(X).Thr = Gr(X).Thr + (ThrDecayGr * (Gr(X).ThrBase - Gr(X).Thr))
           
            If Gr(X).v > Gr(X).Thr Then
                Gr(X).act = 1
                Gr(X).Thr = ThrmaxGr
                gr_elig(X, gr_elig_counter) = 1
                jav_grspktotal = jav_grspktotal + 1
                GranAct(X) = GranAct(X) + 1
                Gr(X).g_KCa = Gr(X).g_KCa + 0.003
            Else
                Gr(X).act = 0
                Gr(X).g_KCa = Gr(X).g_KCa * 0.9999
            End If
        Next X

    If cbm_main.Scaling.Value = Checked Then
        For X = 1 To grancells
            Gr(X).g_Var = Gr(X).g_Var + 0.000001 - (Gr(X).act * 0.00005)
        Next X
    End If
    
    T2 = Timer
    TIME_GR = TIME_GR + T2 - T1
    
'    If CSisPresent Then
'        For X = 1 To grancells
'            GranActCS(X) = GranActCS(X) + Gr(X).act
'        Next X
'    End If
    
  '*******************************Golgi Cells**********************************************************
  For X = 1 To golgicells
      Gol(X).gMF = (MF2(Gol(X).MF(1)).gGol + MF2(Gol(X).MF(2)).gGol + MF2(Gol(X).MF(3)).gGol + MF2(Gol(X).MF(4)).gGol) / 4#
      
      'Debug.Print MF2(Gol(X).MF(1)).gGol, MF2(Gol(X).MF(2)).gGol, MF2(Gol(X).MF(3)).gGol, MF2(Gol(X).MF(4)).gGol
      
      inGr = Gr(Gol(X).preGr(1)).act
      For i = 2 To numGoGrDend
        inGr = inGr + Gr(Gol(X).preGr(i)).act
      Next i
      
      If DoGG = 1 Then
        inGol = 0
        inGol2 = 0
        For i = 1 To 8
            inGol = inGol + (Gol(GG(X, i)).gGG * wGG(GG(X, i), i))
            inGol2 = inGol2 + wGG(GG(X, i), i)
        Next i
        inGol = inGol / 8#
        gGGAverage = gGGAverage + inGol
      Else
        
            If GGconst < 0.05 Then
                inGol = 0.005
            ElseIf GGconst < 0.1 Then
                inGol = 0.011
            ElseIf GGconst < 0.15 Then
                inGol = 0.015
            ElseIf GGconst < 0.2 Then
                inGol = 0.018
            ElseIf GGconst < 0.25 Then
                inGol = 0.021
            ElseIf GGconst < 0.3 Then
                inGol = 0.024
            Else
                inGol = 0.0265
            End If
        
      End If
      
      Gol(X).GGr = inGr * Gol(X).g_varGr + Gol(X).GGr * gEDecayGoGr
      Gol(X).gSlow = inGr * GolMGluR + Gol(X).gSlow * gSlowDecayGo
      
      Gol(X).v = Gol(X).v + ((gLGo + Gol(X).gSlow) * (ELGo - Gol(X).v) - (Gol(X).gMF + Gol(X).GGr) * Gol(X).v) + (inGol * (ELGo - Gol(X).v)) '    /*EE for both gE1 and gE2 =0 */
      Gol(X).Thr = Gol(X).Thr + (ThrDecayGo * (Gol(X).ThrBase - Gol(X).Thr))
      
      If Gol(X).v > Gol(X).Thr Then
        Gol(X).act = 1
        Gol(X).Thr = ThrmaxGo
        Gol(X).gGG = Gol(X).gGG + ((1 - Gol(X).gGG) * GGconst)
      Else
        Gol(X).act = 0
        Gol(X).gGG = Gol(X).gGG * gIDecayGr
      End If
    Next X
    
    If GolGrHeterogenous = 1 Then
        For X = 1 To golgicells
            Gol(X).gFast = Gol(X).act * gIconstGr + Gol(X).gFast * gIFastDecayGr
            Gol(X).gSlow = Gol(X).gSlow + ((0.02 * Gol(X).gFast) * (5 * (0.2 - Gol(X).gSlow)))
            Gol(X).gSlow = Gol(X).gSlow * GABAdecay
            Gol(X).gSlow = 0
            Gol(X).gFastFinal = Gol(X).gFast * 50 * (0.1 - Gol(X).gSlow) * (0.15 - Gol(X).gSlow) + (Gol(X).gFast * 4 * (0.07 - Gol(X).gSlow))
        Next X
    End If
    
    If cbm_main.Scaling_MF_Gol.Value = Checked Then
        For X = 1 To golgicells
            Gol(X).g_varMF = Gol(X).g_varMF + 0.000001 - (Gol(X).act * 0.00005)
        Next X
    End If
    If cbm_main.Scaling_gr_Gol.Value = Checked Then
        For X = 1 To golgicells
            Gol(X).g_varGr = Gol(X).g_varGr + 0.000001 - (Gol(X).act * 0.00005)
        Next X
    End If
    T3 = Timer
    TIME_GOL = TIME_GOL + T3 - T2
    
If MFgrGolOnly = 0 Then

' ******************************** STELLATE cells ******************************
    For X = 1 To StellateCells   '240
        y = 1 + (X - 1) * PFStellsynUMBER
        y1 = (X * PFStellsynUMBER)
        inGr = 0
        For syn = y To y1 '50
            inGr = inGr + Gr(syn).act  '(Gr(syn).act * (1# - SB_STP(syn)))  used in STP paper
        Next syn
        Bk(X).GGr = Bk(X).GGr + (inGr * WEIGHTSTELL)
        Bk(X).GGr = Bk(X).GGr * GDecayGrStell
        Bk(X).v = Bk(X).v + (GLeakStell * (ELeakStell - Bk(X).v) - Bk(X).GGr * Bk(X).v)
        Bk(X).Thr = Bk(X).Thr + (ThrDecayStell * (THRBASEStell - Bk(X).Thr))
        If Bk(X).v > Bk(X).Thr Then
            Bk(X).act = 1
            Bk(X).Thr = THRMAXStell
            StellPcG(X) = StellPcG(X) + (GCONSTStellPC * (1 - StellPcG(X)))
        Else
            Bk(X).act = 0
            StellPcG(X) = StellPcG(X) * GDecayStellPC
        End If
    Next X
    
' ******************************** BASKET cells ******************************
    grCount = 0  ' which granule cell, this works because they parse by Basket cell
    For X = 1 To BasketNUMBER   '96
        
        BCells(X).GGr = 0
        For syn = 1 To PFBasketsynUMBER
            grCount = grCount + 1
            
            If Gr(grCount).act Then
                BCells(X).gE(syn) = BCells(X).gE(syn) + (BCells(X).grW(syn) * GrBCWeights)
            Else
                BCells(X).gE(syn) = BCells(X).gE(syn) * GDecayGrBC
            End If
            
            BCells(X).GGr = BCells(X).GGr + BCells(X).gE(syn)
        Next syn
        BCells(X).GGr = (BCells(X).GGr * GrBCWeights)
        If BCells(X).GGr > 0.99 Then BCells(X).GGr = 0.99
        'If X = 22 Then Debug.Print BCells(X).GGr
        
        inPC = Pc(BCells(X).PCsyn(1)).act
        For syn = 2 To 4
            inPC = inPC + Pc(BCells(X).PCsyn(syn)).act
        Next syn
        BCells(X).gPc = BCells(X).gPc + (inPC * PCBCWeights)
        BCells(X).gPc = BCells(X).gPc * GDecayPCBC
        
        If UsePCtoBasketSynapses = 0 Then
            BCells(X).gPc = TonicgI_Basket
        End If
        BCells(X).gPc = ((UsePCtoBasketSynapses) * BCells(X).gPc) + ((1 - UsePCtoBasketSynapses) * TonicgI_Basket) ' use fixed (tonic) gI if PC to basket disabled
        
        BCells(X).v = BCells(X).v + ((GLeakBC * (ELeakBC - BCells(X).v)) - (BCells(X).GGr * BCells(X).v) + (BCells(X).gPc * (-70 - BCells(X).v)))
        
        If BCells(X).v > BCells(X).Thr Then
            BCells(X).act = 1
            BCells(X).Thr = ThrMAXBC
        Else
            BCells(X).act = 0
            BCells(X).Thr = BCells(X).Thr + (ThrDecayBC * (ThrBaseBC - BCells(X).Thr))
        End If
    Next X
    
    T4 = Timer
    TIME_BK = TIME_BK + T4 - T3

' ************* PURKINJE CELLS ****************************
    For X = 1 To PCcells  '24
        y = 1 + (X - 1) * PFPCSYNUMBER
        y1 = y + PFPCSYNUMBER - 1
        For syn = y To y1
            Pc(X).GGr = Pc(X).GGr + ((grWeight(syn) + grSTP(syn)) * Gr(syn).act)
        Next syn
        Pc(X).GGr = Pc(X).GGr * GCONSTGRPC
    Next X
    
    For X = 1 To PCcells  '20
        y = 1 + (X - 1) * PCBasketsynUMBER
        y1 = y + PCBasketsynUMBER - 1
        For syn = y To y1
            Pc(X).GBC = Pc(X).GBC + (GCONSTBCPC * BCells(syn).act)
        Next syn
        Pc(X).GBC = Pc(X).GBC * GCONSTBCPC
    Next X
    
    syn = 0
    
    For X = 1 To PCcells
        For y = 1 To StellPCSYNUMBER
            syn = syn + 1
            Pc(X).GStell = Pc(X).GStell + (StellPcG(syn))
        Next y
        
        Pc(X).v = Pc(X).v + ((GLeakPC * (ELEAKPC - Pc(X).v) - (Pc(X).GGr * Pc(X).v) + (Pc(X).GStell * (VStellPC - Pc(X).v)) + (Pc(X).GBC * (VStellPC - Pc(X).v))))
        Pc(X).Thr = Pc(X).Thr + (ThrDecayPC * (Pc(X).ThrBase - Pc(X).Thr))
        
        If Pc(X).v > Pc(X).Thr Then
            Pc(X).act = 1
            Pc(X).Thr = THRMAXPC
            pcspknow = pcspknow + 1
            Purk_histo(X, bincounter5) = Purk_histo(X, bincounter5) + 1
            PurkActivity(X) = PurkActivity(X) + 1
        Else
            Pc(X).act = 0
        End If
    Next X
    T5 = Timer
    TIME_PK = TIME_PK + T5 - T4
    
' ************* NUCLEUS CELLS ********************************************
'    If response_counter < 50 Then
'        response_counter = response_counter + 1
'    Else
'        response_counter = 1
'    End If
    syn = 0
    If CFAsynchronouse = 1 Then Erase NcCfG
    For X = 1 To NCNUMBER
        Nc(X).gMF = 0
        Nc(X).gMF2 = 0
       
        For y = 1 To MFNCSYNUMBER
            
            syn = syn + 1
'            If MFS(1, syn).CStype = 1 Or MFS(1, syn).CStype = 5 Then   'For the STP paper.......
'                MFW = 0
''                If Trials_this_time > 9 Then
''                    MFW = 1# * ((118# - Trials_this_time) / 108#)
''                    If MFW < 0.2 Then MFW = 0.2
''                    MFW = mfweight(syn) * MFW
''                End If
'            Else
'                MFW = mfweight(syn)
'            End If
            
            MFW = mfweight(syn)
'            MFW = 0
            Nc(X).NMDABind(y) = Nc(X).NMDABind(y) * gNMDADecayMFNC + (MF(syn) * MFW * (1 - Nc(X).NMDABind(y)))
            Nc(X).gNMDA(y) = Nc(X).gNMDA(y) + (grNMDActivate * (Nc(X).NMDABind(y) - Nc(X).gNMDA(y)))
            Nc(X).AMPABind(y) = Nc(X).AMPABind(y) * gAMPADecayMFNC + (MF(syn) * MFW * (1 - Nc(X).AMPABind(y)))
            Nc(X).gAMPA(y) = Nc(X).gAMPA(y) + (grNMDActivate * (Nc(X).AMPABind(y) - Nc(X).gAMPA(y)))
            Nc(X).gMF = Nc(X).gMF + Nc(X).gNMDA(y)
            Nc(X).gMF2 = Nc(X).gMF2 + Nc(X).gAMPA(y)
        Next y
        Nc(X).gMF = (Nc(X).gMF * Time_step_size) / MFNCSYNUMBER
        Nc(X).gMF2 = (Nc(X).gMF2 * Time_step_size) / MFNCSYNUMBER
        I_NMDA(X) = (Nc(X).gMF * (Nc(X).v / -80))
        
        Nc(X).gPc = 0
        For y = 1 To PCNCSYNUMBER
            Nc(X).gGABA(y) = Nc(X).gGABA(y) * GDecayPCNC + (Pc(Nc(X).PCsyn(y)).act * gPURKtoNUCLEUS(Nc(X).PCsyn(y), X) * (1 - Nc(X).gGABA(y)))
            Nc(X).gPc = Nc(X).gPc + Nc(X).gGABA(y)
        Next y
        Nc(X).gPc = (Nc(X).gPc * Time_step_size) / PCNCSYNUMBER
        
        
        
        Nc(X).v = Nc(X).v + ((GLeakNC * (ELEAKNC - Nc(X).v)) - ((I_NMDA(X) + Nc(X).gMF2) * Nc(X).v) + (Nc(X).gPc * (VPCNC - Nc(X).v)))
        
        Nc(X).Thr = Nc(X).Thr + (THRdecayNC * (THRBASENC - Nc(X).Thr))
        If Nc(X).v > Nc(X).Thr Then
            Nc(X).act = 1
            Nc(X).Thr = THRMAXNC
            Nuc_histo(X, bincounter5) = Nuc_histo(X, bincounter5) + 1
        Else
            Nc(X).act = 0
        End If
        
        If CFAsynchronouse = 0 Then
            For y = 1 To NumCF
                If Nc(X).act = 1 Then
                    NcCfG(X, y) = (NcCfG(X, y) + (Nc(X).act * Nc(X).gNUCtoCF(y) * (1 - NcCfG(X, y))))
                Else
                    NcCfG(X, y) = ((1 - Nc(X).act) * NcCfG(X, y) * GDecayNCCF)
                End If
            Next y
        Else
            For y = 1 To NumCF
                If CF(y).NucInput(X) = 1 Then
                    CF(y).CurrentP(X) = CF(y).CurrentP(X) * Exp(-1 / (CFProbDecayTN * Exp(-CF(y).CurrentP(X) / CFProbDecayTTau) + CFProbDecayTO))
                    CF(y).CurrentP(X) = ((1 - Nc(X).act) * CF(y).CurrentP(X)) + (Nc(X).act * (CF(y).CurrentP(X) + 0.3 * CFProbIncN * Exp(-CF(y).CurrentP(X)) / CFProbIncTau))
                    CF(y).CurrentG(X) = CF(y).CurrentG(X) * (Exp(-1 / (-gDecayTN * Exp(-CF(y).CurrentG(X) / gDecayTTau) + gDecayTO)))
                    r = Rnd()
                    If r < CF(y).CurrentP(X) Then
                        CF(y).CurrentG(X) = CF(y).CurrentG(X) + gIncN * 2 * Exp(-CF(y).CurrentG(X) / gIncTau)
                    End If
                    NcCfG(X, y) = NcCfG(X, y) + ((CF(y).CurrentG(X) / 100) * Nc(X).gNUCtoCF(y))
                    'Debug.Print NcCfG(x, y),
                End If
            Next y
        End If
    Next X
    'Debug.Print
    T6 = Timer
    TIME_NUC = TIME_NUC + T6 - T5

'******************* CLIMBING FIBER ***************************************************************
    If NumCF > 1 And CFCoupled = 1 Then
        For X = 1 To NumCF
            CF(X).vE = 0
        Next X
        If NumCF = 12 Then
            T2 = 0.016
            
            T1 = CF(NumCF).v - CF(1).v
            CF(1).vE = CF(1).vE + (T2 * T1)    ' 1 = 12-1
            
            CF(NumCF).vE = CF(NumCF).vE + (-1 * T2 * T1)    '12 = 1-12
            
            For X = 1 To NumCF - 1
                T1 = CF(X + 1).v - CF(X).v
                CF(X).vE = CF(X).vE + (T2 * T1)
                CF(X + 1).vE = CF(X + 1).vE + (-1 * T2 * T1)
            Next X
                
            For X = 1 To 6
                T1 = CF(X + 6).v - CF(X).v
                CF(X).vE = CF(X).vE + (T2 * T1)
                CF(X + 6).vE = CF(X + 6).vE + (-1 * T2 * T1)
            Next X
            
            
        ElseIf NumCF = 4 Then
            CF(1).vE = CF(1).vE + (0.04 * (CF(2).v - CF(1).v))
            CF(2).vE = CF(2).vE + (0.04 * (CF(3).v - CF(2).v))
            CF(3).vE = CF(3).vE + (0.04 * (CF(4).v - CF(3).v))
            CF(4).vE = CF(4).vE + (0.04 * (CF(1).v - CF(4).v))
        End If
    End If
    
    For X = 1 To NumCF
        CF(X).GNc = 0
        For y = 1 To NCNUMBER
            CF(X).GNc = CF(X).GNc + (NcCfG(y, X) * CF(X).NucInput(y) * 1.5)
        Next y
        
        CF(X).v = CF(X).v + ((CF(X).GLeak * (ELEAKCF - CF(X).v)) + (CF(X).GNc * (VNCCF - CF(X).v))) + CF(X).vE
        
        If Bincounter = US_onset(1) * US1_ON Then
            If AmpMode = 0 Then
                CF(X).v = CF(X).v + MAXDRUS
            Else
                Rmax = 0
                For i = 1000 To Bincounter - 1
                    If RN_Histo(i) > Rmax Then Rmax = RN_Histo(i)
                Next i
                'cbm_main.AmpLabel = Format(Str(Rmax), "###.##")
                Rmax = Rmax / 1200
                If Rmax < AmpModeAmp Then
                    CF(X).v = CF(X).v + MAXDRUS
                End If
            End If
        End If
        If CF(X).v > CF(X).Thr Then
            CF(X).act = 1
            CF(X).Thr = THRMAXCF
            ltdtime(X) = 0
            CF_histo(0, bincounter5) = CF_histo(0, bincounter5) + 1
            CF_spike_counter(X) = CF_spike_counter(X) + 1
            DebugCF = DebugCF + 1
        Else
            CF(X).act = 0
            CF(X).Thr = CF(X).Thr + ((ThrDecayCF + ((0.02 - CF(X).GLeak) * 0.3)) * (THRBASECF - CF(X).Thr))
        End If
        
    Next X
    
' ******************* Red Nucleus Pseudocell *************************************************

    DoWorkRN

' ********************************************* MF TOTAL INPUT *********************************
    For X = 1 To 10
        jav_Mfinputaverage(bincounter5) = jav_Mfinputaverage(bincounter5) + jav_Mfinputtotal(X)
    Next X
  
' **************************************** gr-PURK PLASTICITY ****************************************
    T6b = Timer
    
    If (Time_step_size = 5 Or bincounter5_temp = 0) Then
        Debug_weights = 0#
        
        If gr_elig_counter > ((LTD_OFFSET / 5) + (LTDURATION / 5)) Then         ' *** look 20 time bins or 100 msec ago....
            gr_counter_TEMP = gr_elig_counter - (LTD_OFFSET / 5) - (LTDURATION / 5)
        Else
            gr_counter_TEMP = gr_elig_counter + 200 - (LTD_OFFSET / 5) - (LTDURATION / 5)
        End If
        
        For X = 1 To NumCF
            If ltdtime(X) <= LTDURATION Then ltdtime(X) = ltdtime(X) + 5
        Next X
           
        If CHANGEgrWEIGHTS = 1 Then
            If grPCPlasticityType = 0 Then   'graded plasticity
            
                For X = 1 To NumCF
                    S1 = 1 + ((X - 1) * (SYNUMBER / NumCF))
                    S2 = S1 + (SYNUMBER / NumCF) - 1
                    If ltdtime(X) <= LTDURATION Then
                        For m = S1 To S2
                           grchange = gr_elig(m, gr_counter_TEMP) * DELMINUS
                           grWeight(m) = grWeight(m) + grchange
                           Debug_weights = Debug_weights + grchange
                           Raster_plasticity(m) = grchange
                           If (grWeight(m) < 0#) Then grWeight(m) = 0#
                        Next m
                    Else
                        For m = S1 To S2
                           grchange = gr_elig(m, gr_counter_TEMP) * DELPLUS
                           grWeight(m) = grWeight(m) + grchange
                           Debug_weights = Debug_weights + grchange
                           Raster_plasticity(m) = grchange
                           If (grWeight(m) > 1#) Then grWeight(m) = 1#
                        Next m
                    End If
                Next X
                Debug_weights_total = Debug_weights_total + Debug_weights
                
            ElseIf grPCPlasticityType = 1 Then
                For X = 1 To NumCF
                    If ltdtime(X) <= LTDURATION Then
                        PC_form.Line (Bincounter, PC_form.ScaleTop)-(Bincounter, PC_form.ScaleTop - 1), vbWhite
                        S1 = 1 + ((X - 1) * (SYNUMBER / NumCF))
                        S2 = S1 + (SYNUMBER / NumCF) - 1
                        For m = S1 To S2
                            If Rnd() < (DELMINUS * BinaryWeightRate) And gr_elig(m, gr_counter_TEMP) = 1 Then grWeight(m) = grWeightMin
                            grchange = gr_elig(m, gr_counter_TEMP) * DELMINUS
                            grweightChangeThisSession(m, 2) = grweightChangeThisSession(m, 2) + grchange
                            grWeight(m) = grWeight(m) + grchange
                            Debug_weights = Debug_weights + grchange
                            Raster_plasticity(m) = grchange
                            If (grWeight(m) < 0#) Then grWeight(m) = 0#
                        Next m
                    Else
                        S1 = 1 + ((X - 1) * (SYNUMBER / NumCF))
                        S2 = S1 + (SYNUMBER / NumCF) - 1
                        For m = S1 To S2
                           If Rnd() < (DELPLUS * BinaryWeightRate) And gr_elig(m, gr_counter_TEMP) = 1 Then grWeight(m) = grWeightMax
                           grchange = gr_elig(m, gr_counter_TEMP) * DELPLUS
                           grweightChangeThisSession(m, 1) = grweightChangeThisSession(m, 1) + grchange
                           grWeight(m) = grWeight(m) + grchange
                           Debug_weights = Debug_weights + grchange
                           Raster_plasticity(m) = grchange
                           If (grWeight(m) > 1#) Then grWeight(m) = 1#
                        Next m
                    End If
                Next X
            Else
                For X = 1 To NumCF
                    If ltdtime(X) <= LTDURATION Then  'LTD
                        PC_form.Line (Bincounter, PC_form.ScaleTop)-(Bincounter, PC_form.ScaleTop - 1), vbWhite
                        S1 = 1 + ((X - 1) * (SYNUMBER / NumCF))
                        S2 = S1 + (SYNUMBER / NumCF) - 1
                        For m = S1 To S2
                            If gr_elig(m, gr_counter_TEMP) = 1 Then
                                Select Case grWeight(m)
                                    Case grWeightMax
                                        If Rnd() < (BinaryProbMinus) Then grWeight(m) = grWeightMin
                                    Case grWeightMin
                                        If Rnd() < (BinaryProbMinus / 2) Then grWeight(m) = grWeightMin - 0.01
                                    Case grWeightMin - 0.01
                                        If Rnd() < (BinaryProbMinus / 4) Then grWeight(m) = grWeightMin - 0.02
                                    Case grWeightMin - 0.02
                                        If Rnd() < (BinaryProbMinus / 8) Then grWeight(m) = grWeightMin - 0.03
                                    Case grWeightMax + 0.01
                                        If Rnd() < (BinaryProbMinus / 2) Then grWeight(m) = grWeightMin
                                    Case grWeightMax + 0.02
                                        If Rnd() < (BinaryProbMinus / 4) Then grWeight(m) = grWeightMin
                                    Case grWeightMax + 0.03
                                        If Rnd() < (BinaryProbMinus / 8) Then grWeight(m) = grWeightMin
                                End Select
                            End If
                        Next m
                    Else
                        S1 = 1 + ((X - 1) * (SYNUMBER / NumCF))
                        S2 = S1 + (SYNUMBER / NumCF) - 1
                        For m = S1 To S2
                            If gr_elig(m, gr_counter_TEMP) = 1 Then
                                Select Case grWeight(m)
                                    Case grWeightMin
                                        If Rnd() < (BinaryProbPlus) Then grWeight(m) = grWeightMax
                                    Case grWeightMax
                                        If Rnd() < (BinaryProbPlus / 2) Then grWeight(m) = grWeightMax + 0.01
                                    Case grWeightMax + 0.01
                                        If Rnd() < (BinaryProbPlus / 4) Then grWeight(m) = grWeightMax + 0.02
                                    Case grWeightMax + 0.02
                                        If Rnd() < (BinaryProbPlus / 8) Then grWeight(m) = grWeightMax + 0.03
                                    Case grWeightMin - 0.01
                                        If Rnd() < (BinaryProbPlus / 2) Then grWeight(m) = grWeightMax
                                    Case grWeightMin - 0.02
                                        If Rnd() < (BinaryProbPlus / 4) Then grWeight(m) = grWeightMax
                                    Case grWeightMin - 0.03
                                        If Rnd() < (BinaryProbPlus / 8) Then grWeight(m) = grWeightMax
                                End Select
                            End If
                        Next m
                    End If
                Next X
            End If
        End If
        
        If MLIPlasticity Then
            'Debug.Print Bincounter, bincounter5, gr_counter_TEMP, ltdtime(1), LTDURATION
            counter = 0
            S1 = 0
            For X = 1 To NumCF
                If ltdtime(X) <= LTDURATION Then
                    For i = 1 To 8   ' each Cf talks to 8 BCs
                        S1 = S1 + 1
                        For m = 1 To PFBasketsynUMBER
                            counter = counter + 1
                            grchange = gr_elig(counter, gr_counter_TEMP) * DELMINUS * -1
                            BCells(S1).grW(m) = BCells(S1).grW(m) + grchange
                            If BCells(S1).grW(m) > 1 Then BCells(S1).grW(m) = 1
                        Next m
                    Next i
                Else
                    For i = 1 To 8
                        S1 = S1 + 1
                        For m = 1 To PFBasketsynUMBER
                            counter = counter + 1
                            grchange = gr_elig(counter, gr_counter_TEMP) * DELPLUS * -1
                            BCells(S1).grW(m) = BCells(S1).grW(m) + grchange
                            If BCells(S1).grW(m) < 0 Then BCells(S1).grW(m) = 0
                        Next m
                    Next i
                End If
            Next X
        End If
        
    End If
   
'************************** Presynaptic short-term plasticity *************************************************
    
    For m = 1 To SYNUMBER  'always calculate eligibility in case it's being used for STP at gr to stellate or gr to basket synapses.
        grPreElig(m) = (Gr(m).act * 1 * grPreElig(m)) + ((1 - Gr(m).act) * (grPreElig(m) + ((grEligBase - grPreElig(m)) * grPreEligDecay))) '.92  0.90   2.2
    Next m
    
    If DoLTPShortTermPlasticity = 1 Then
        If STPblockedbyCF Or NumCF = 1 Then
            For X = 1 To NumCF
                S1 = 1 + ((X - 1) * (SYNUMBER / NumCF))
                S2 = S1 + (SYNUMBER / NumCF) - 1
                For m = S1 To S2
                    If grPreElig(m) > 1 Then
                        If ((ltdtime(X) >= LTDURATION)) Then 'And (grPreEligTimeout(m) = 0)) Then
                            grPreElig(m) = grEligBase
                            grSTP(m) = grSTP(m) + STPAmp
                            If grSTP(m) > grSTPMax Then grSTP(m) = grSTPMax
                            'grPreEligTimeout(m) = 1
                        End If
                    ElseIf grPreElig(m) < grEligBase Then
                        grPreElig(m) = grEligBase
                    End If
                    'If grPreEligTimeout(m) > 0 Then grPreEligTimeout(m) = grPreEligTimeout(m) - 1
                Next m
            Next X
            
        Else  'STP not blocked by CF
            For m = 1 To SYNUMBER
                If grPreElig(m) > 1 Then
                    grPreElig(m) = grEligBase
                    'If grPreEligTimeout(m) = 0 Then
                        grSTP(m) = grSTP(m) + STPAmp
                        If grSTP(m) > grSTPMax Then grSTP(m) = grSTPMax
                        'grPreEligTimeout(m) = 1
                    'End If
                ElseIf grPreElig(m) < grEligBase Then
                    grPreElig(m) = grEligBase
                End If
                'If grPreEligTimeout(m) > 0 Then grPreEligTimeout(m) = grPreEligTimeout(m) - 1
            Next m
        End If
    End If
    
'    If DoStellateBasketSTP Then
'        For m = S1 To S2
'            If grPreElig(m) > 1 Then
'                grPreElig(m) = grEligBase
'                'If grPreEligTimeout(m) = 0 Then
'                    SB_STP(m) = SB_STP(m) + STPAmp
'                    If SB_STP(m) > sb_STPMax Then SB_STP(m) = sb_STPMax
'                    'grPreEligTimeout(m) = 1
'                'End If
'            ElseIf grPreElig(m) < grEligBase Then
'                grPreElig(m) = grEligBase
'            End If
'            'If grPreEligTimeout(m) > 0 Then grPreEligTimeout(m) = grPreEligTimeout(m) - 1
'        Next m
'
'    End If
' ****************************************** MF-NUC plasticity **************************************************
    If Time_step_size = 5 Or bincounter5_temp = 0 Then
        If pcrunavcounter = PCRUNAVWINDOW Then
            pcrunavcounter = 1
        Else
            pcrunavcounter = pcrunavcounter + 1
        End If
        pcrunav = pcrunav - runavpc(pcrunavcounter) + pcspknow
        runavpc(pcrunavcounter) = pcspknow
        pcspknow = 0
    End If
    
    If (CHANGEmfWEIGHTS And (Time_step_size = 5 Or bincounter5_temp = 0)) Then
        'If bincounter5_temp = 0 Then   '  step to new bin (every bin for 5 ms, every 5 for 1 ms
            If ((pcrunav / (PCRUNAVWINDOW) > MF_LTD_threshold)) Then
                mfltdperiod = 1
                hyper_switch = 1
            ElseIf (pcrunav / PCRUNAVWINDOW > ((MF_LTD_threshold + MF_LTP_threshold) / 2)) Then
                hyper_switch = 1
            Else
                mfltdperiod = 0
            End If
            If (((pcrunav / (PCRUNAVWINDOW)) < MF_LTP_threshold) And (hyper_switch = 1)) Then
                mfltpperiod = 1
                hyper_switch = 0       '/* Can't do LTP again until cell becomes hyperpolarized */
            Else
                mfltpperiod = 0
            End If
            
            If (mfltpperiod + mfltdperiod) <> 0 Then
                 For m = 1 To MFNUMBER
                    rasters_MF_plasticity(m) = MfElig(m) * ((mfltdperiod * MFDELMINUS) + (mfltpperiod * MFDELPLUS))
                    mfweight(m) = mfweight(m) + rasters_MF_plasticity(m)
                    If (mfweight(m) > 1#) Then
                        mfweight(m) = 1#
                    ElseIf (mfweight(m) < 0#) Then
                        mfweight(m) = 0#
                    End If
                 Next m
                If mfltdperiod = 1 Then
                    'PC_form.Line (bincounter, PC_form.ScaleTop)-(bincounter, PC_form.ScaleTop + (0.005 * PC_form.ScaleHeight)), vbRed
                ElseIf mfltpperiod = 1 Then
                    PC_form.Line (Bincounter, PC_form.ScaleTop)-(Bincounter, PC_form.ScaleTop + (0.005 * PC_form.ScaleHeight)), vbGreen
                End If
            Else
                For m = 1 To MFNUMBER
                    rasters_MF_plasticity(m) = 0
                Next m
            End If
        'End If
        If Bincounter >= 5000 Then
            For m = 1 To SYNUMBER
                gr_PC_weights_rasters(m, Trials_this_time) = grWeight(m)
            Next m
            For m = 1 To MFNUMBER
                mf_nuc_weights_rasters(m, Trials_this_time) = mfweight(m)
            Next m
        End If
    End If

' ******************** HOMEOSTATIC PLASTICITY **********************

' PURKINJE CELL Synaptic Scaling

    If Bincounter >= 5000 And DoPurkSynapticScaling Then PurkSynapticScaling

' PURKINJE CELL PRE
    If Bincounter >= 5000 And DoPurkPre Then PurkPresynapticPlasticity

' PURKINJE CELL INTRINSIC EXCITABILITY

    If Bincounter >= 5000 And DoPurkIntrinsic Then PurkIntrinsicPlasticity

'NUCLEUS CELL PRE
    If Bincounter >= 5000 And DoNucPre Then NucleusPresynapticPlasticity
    
    If Bincounter >= 5000 Then
        For m = 1 To PCNUMBER
            PurkActivity(m) = 0
        Next m
    End If

T7 = Timer
TIME_Plasticity = TIME_Plasticity + T7 - T6
'****************** Write activity each trial *****************************

If cbm_main.Check2.Value = Checked Then
    For X = 1 To PCNUMBER
        PurkinjeActivity(X, Trials_this_time) = PurkinjeActivity(X, Trials_this_time) + Pc(X).act
    Next X
    For X = 1 To NCNUMBER
        NucleusActivity(X, Trials_this_time) = NucleusActivity(X, Trials_this_time) + Nc(X).act
    Next X
    For X = 1 To NumCF
        ClimbingFiberActivity(Trials_this_time, X) = ClimbingFiberActivity(Trials_this_time, X) + CF(X).act
        If CF(X).act = 1 Then
            CFCounter(X) = CFCounter(X) + 1
            CFactivityTimes(Trials_this_time, X, CFCounter(X)) = Bincounter
        End If
    Next X
    For X = 1 To SYNUMBER
        GranuleActivity(Trials_this_time) = GranuleActivity(Trials_this_time) + Gr(X).act
    Next X
    For X = 1 To 900
        GolgiActivity(Trials_this_time) = GolgiActivity(Trials_this_time) + Gol(X).act
    Next X
    For X = 1 To MFNUMBER
        MFActivity(Trials_this_time) = MFActivity(Trials_this_time) + MF(X)
    Next X
End If

' ***************AutoSave PC NUC CF each trial **********************
If cbm_main.PCNucCFCheck.Value = Checked Then
    For X = 1 To PCNUMBER
        PurkinjeActivity(X, 1) = PurkinjeActivity(X, 1) + Pc(X).act
    Next X
    For X = 1 To NCNUMBER
        NucleusActivity(X, 1) = NucleusActivity(X, 1) + Nc(X).act
    Next X
    For X = 1 To NumCF
        ClimbingFiberActivity(1, X) = ClimbingFiberActivity(1, X) + CF(X).act
    Next X
End If
'****************** Main window display *****************************

If PC_form.WindowState <> 1 And PC_form.Visible = True Then
'    If PCFormColors = 1 Then
'        PCcolor = &HFF8080
'        NCcolor = vbWhite
'        CFcolor = RGB(255, 0, 0)
'    Else
'        PCcolor = vbBlack
'        NCcolor = vbBlack
'        CFcolor = vbBlack
'    End If
    For X = 1 To PCNUMBER
       PC_form.PSet (Bincounter, (X * 17) - ((Pc(X).v * 0.7) + ELEAKPC) - 108), PCcolor ' was 25
       If Pc(X).act = 1 Then PC_form.Line -(Bincounter, (X * 17) - ((Pc(X).v * 0.7) + ELEAKPC) - 120), PCcolor ' was 25  and -123
    Next X
    For X = 1 To NCNUMBER
        PC_form.PSet (Bincounter, ((X + 21.5) * 17) - ((Nc(X).v * 0.7) + ELEAKNC) - 80), NCcolor 'was *25
        If Nc(X).act = 1 Then PC_form.Line -(Bincounter, ((X + 21.5) * 17) - ((Nc(X).v * 0.7) + ELEAKNC) - 90), NCcolor  ' - 95
    Next X
    
    If raster_form.cell_menu2(6).Checked = False Then
        For X = 1 To NumCF
            y = NumCF - X + 1
            PC_form.PSet (Bincounter, ((27) * 19) - ((CF(y).v * 0.7) + ELEAKNC) - 70 + ((NumCF - X) * 14)), CFcolor
            If CF(y).act = 1 Then PC_form.Line -(Bincounter, ((27) * 19) - ((CF(y).v * 0.7) + ELEAKNC) - 80 + ((NumCF - X) * 14)), CFcolor
        Next X
    End If
    
    PC_form.PSet (Bincounter, PC_form.ScaleHeight - (RN.Vm + 1)), vbWhite
End If
End If   '  if MFgrGolOnly = 0  toggles abbreviated simulation mode

'***************************** Conductance Form

If ConductanceForm.Visible = True Then
    Select Case ConductanceFormMode

        Case 1
            If Bincounter = 1 Then ConductanceForm.Cls
            For X = 1 To NumCF
                y = NumCF - X + 1
                ConductanceForm.PSet (Bincounter, (CF(y).GNc * CFgDisplayMultiplier) + X - 1), RGB(0, 255, 0)
            Next X
'            ConductanceForm.PSet (bincounter, CF(1).CurrentP(1)), RGB(255, 255, 0)
'            ConductanceForm.PSet (bincounter, CF(1).CurrentP(2)), RGB(255, 255, 0)
'            ConductanceForm.PSet (bincounter, CF(2).CurrentP(3) + 1), RGB(255, 255, 0)
'            ConductanceForm.PSet (bincounter, CF(2).CurrentP(4) + 1), RGB(255, 255, 0)
'            ConductanceForm.PSet (bincounter, CF(3).CurrentP(5) + 2), RGB(255, 255, 0)
'            ConductanceForm.PSet (bincounter, CF(3).CurrentP(6) + 2), RGB(255, 255, 0)
'            ConductanceForm.PSet (bincounter, CF(4).CurrentP(7) + 3), RGB(255, 255, 0)
'            ConductanceForm.PSet (bincounter, CF(4).CurrentP(8) + 3), RGB(255, 255, 0)
            ConductanceForm.PSet (Bincounter, CF(1).CurrentG(1)), RGB(255, 255, 255)
            ConductanceForm.PSet (Bincounter, CF(1).CurrentG(2)), RGB(255, 255, 255)
            ConductanceForm.PSet (Bincounter, CF(2).CurrentG(3) + 1), RGB(255, 255, 255)
            ConductanceForm.PSet (Bincounter, CF(2).CurrentG(4) + 1), RGB(255, 255, 255)
            ConductanceForm.PSet (Bincounter, CF(3).CurrentG(5) + 2), RGB(255, 255, 255)
            ConductanceForm.PSet (Bincounter, CF(3).CurrentG(6) + 2), RGB(255, 255, 255)
            ConductanceForm.PSet (Bincounter, CF(4).CurrentG(7) + 3), RGB(255, 255, 255)
            ConductanceForm.PSet (Bincounter, CF(4).CurrentG(8) + 3), RGB(255, 255, 255)
        Case 2
            TotalSTP = 0
            For i = 1 To SYNUMBER
                TotalSTP = TotalSTP + (grSTP(i) * Gr(i).act)
            Next i
            If Bincounter = 1 Then ConductanceForm.Cls
            ConductanceForm.PSet (Bincounter, (TotalSTP * CFgDisplayMultiplier) / 200#), vbWhite
        Case 3
            TotalSTP = 0
            STPcount = 1
            For i = 1 To SYNUMBER
                TotalSTP = TotalSTP + (grSTP(i) * Gr(i).act)
                STPcount = STPcount + Gr(i).act
            Next i
            
            If Bincounter = 1 Then ConductanceForm.Cls
            ConductanceForm.PSet (Bincounter, (TotalSTP * CFgDisplayMultiplier) / STPcount), vbWhite
        Case 0
    End Select
End If


'************** activity window displays .... this way, extra calculations are spared when this window is off  *****
total_activity = 0
If activity_window.WindowState <> 1 Then
    If activity_window.Label1(0).Visible = True Then  'granule cells
        activity_window.PSet (Bincounter, 1 - (jav_grspktotal / (100 * Time_step_size))), &HFFFF00
    End If
    If activity_window.Label1(1).Visible = True Then  'Purkinje cells
        activity_window.PSet (Bincounter, 2 - (pcrunav / (1# * PCNUMBER * 5))), RGB(255, 255, 255)
        activity_window.PSet (Bincounter, 2 - ((MF_LTP_threshold * PCRUNAVWINDOW) / (1# * PCNUMBER * 5))), RGB(0, 255, 0)
        activity_window.PSet (Bincounter, 2 - ((MF_LTD_threshold * PCRUNAVWINDOW) / (1# * PCNUMBER * 5))), RGB(255, 0, 0)
    End If
    If activity_window.Label1(2).Visible = True Then  'Nucleus cells
        'activity_window.PSet (bincounter, 3 - (responses(trials_this_time, bincounter5) / 60#)), RGB(255, 255, 0)
    End If
    If activity_window.Label1(3).Visible = True Then  'Golgi cells
        total_activity = 0
        For i = 1 To 900
            total_activity = total_activity + Gol(i).act
        Next i
        activity_window.PSet (Bincounter, 4 - (total_activity / (100 * Time_step_size))), &H8080FF
    End If
    If activity_window.Label1(4).Visible = True Then  'Mossy fibers
        total_activity = 0
        For i = 1 To 600
            total_activity = total_activity + MF(i)
        Next i
        activity_window.PSet (Bincounter, 5 - (total_activity / (60 * Time_step_size))), &HC0FFC0
    End If
    If activity_window.Label1(5).Visible = True Then  'basket cells
        total_activity = 0
        For i = 1 To STELLATENUMBER
            total_activity = total_activity + Bk(i).act
        Next i
        activity_window.PSet (Bincounter, 6 - (total_activity / (40 * Time_step_size))), &HFFFF&
    End If

    If activity_window.Label1(6).Visible = True Then  'Purkinje conductances
            total_activity = 0
            For i = 1 To 20
                total_activity = total_activity + Pc(i).GStell
            Next i
        activity_window.PSet (Bincounter, 7 - (total_activity / (2 * Time_step_size))), vbRed
        
        total_activity = 0
        For i = 1 To 20
            total_activity = total_activity + Pc(i).GGr
        Next i
        
        activity_window.PSet (Bincounter, 7 - (total_activity / (2 * Time_step_size))), vbGreen
    End If
    If activity_window.Label1(7).Visible = True Then  'Nucleus conductances
        activity_window.PSet (Bincounter, 8 - (pcrunav / (60))), vbRed
        activity_window.PSet (Bincounter, 8 - (jav_Mfinputaverage(bincounter5) / Time_step_size / 1.5)), vbGreen
    End If
End If

' *********** Rasters, Histograms, etc. ***********************************
If Raster_Histos_ON Then '* Is_This_a_Trial = 1 Then
    For X = 1 To grancells
        GR_histo(X, bincounter5) = GR_histo(X, bincounter5) + Gr(X).act
    Next X
    For X = 1 To golgicells
        Go_histo(X, bincounter5) = Go_histo(X, bincounter5) + Gol(X).act
    Next X
    For X = 1 To MFNUMBER
        MF_histo(X, bincounter5) = MF_histo(X, bincounter5) + MF(X)
    Next X
    For X = 1 To STELLATENUMBER
        Stellate_histo(X, bincounter5) = Stellate_histo(X, bincounter5) + Bk(X).act
    Next X
    For X = 1 To BasketNUMBER
        Basket_histo(X, bincounter5) = Basket_histo(X, bincounter5) + BCells(X).act
    Next X
End If
If raster_form.WindowState <> 1 And raster_form.Visible = True Then
    If Bincounter / Time_step_size = 1 Then raster_clear = 1 Else raster_clear = 0
    If raster_Cell_type = 0 Or raster_Cell_type = 6 Or raster_Cell_type = 8 Then
        raster_end = raster_start + 599
    ElseIf raster_Cell_type = 1 Then
        raster_end = raster_start + 999
    ElseIf raster_Cell_type = 2 Then
        raster_end = raster_start + 899
    ElseIf raster_Cell_type = 3 Then
        raster_end = raster_start + 239
    ElseIf raster_Cell_type = 4 Then
        raster_end = raster_start + 999
    ElseIf raster_Cell_type = 5 Then
        raster_end = raster_start + 999
    End If
    If raster_mode = 5 Then
        If CSisPresent = 1 Then
            Do_rasters raster_Cell_type, raster_start, raster_end, Bincounter, raster_clear, 1
        Else
            Do_rasters raster_Cell_type, raster_start, raster_end, Bincounter, raster_clear, 0
        End If
    Else
        Do_rasters raster_Cell_type, raster_start, raster_end, Bincounter, raster_clear, raster_mode
    End If
End If


If DoPCPackedRecord And Trials_this_time < 109 Then  'keep all PC spikes in bit packed longs for export to matlab
    Long1 = 0
    For X = 1 To PCNUMBER
        Long1 = Long1 + (Pc(X).act * 2 ^ (X - 1))
    Next X
    PCPacked(Bincounter, Trials_this_time) = Long1
End If


If DoSpecialRecord Then
'    Debug.Print Trials_this_time, bincounter5
    
'    If bincounter5 = 1 Then
'        If Trials_this_time <> 1 Then
'            ' write the last trial's spikes to disk
'            Open "PCspikes" For Append As #11
'            For X = 1 To PCSpikes
'        End If
''        Erase PCSpikes
''        Erase BCSpikes
''        Erase SCSpikes
'    End If
    Long1 = Bincounter
    Long2 = Trials_this_time
    
    For X = 1 To PCNUMBER
        If Pc(X).act = 1 Then
            PCSpikes(X, 0) = PCSpikes(X, 0) + 1
            PCSpikes(X, PCSpikes(X, 0)) = Long1 + ((Long2 - 1) * 5000)
        End If
    Next X
    
    For X = 1 To BasketNUMBER
        If BCells(X).act = 1 Then
            BCSpikes(X, 0) = BCSpikes(X, 0) + 1
            BCSpikes(X, BCSpikes(X, 0)) = Long1 + ((Long2 - 1) * 5000)
        End If
    Next X
    For X = 1 To STELLATENUMBER
        If Bk(X).act = 1 Then
            SCSpikes(X, 0) = SCSpikes(X, 0) + 1
            SCSpikes(X, SCSpikes(X, 0)) = Long1 + ((Long2 - 1) * 5000)
        End If
    Next X
    
End If

If Do_Big_Rasters = 1 Then
    Debug.Print Trials_this_time

'    For X = 1 To PCNUMBER
'        If Pc(X).act = 1 Then
'            Rasters(X + 12, bincounter5, Trials_this_time) = True
'        Else
'            'Rasters(x + 12, bincounter5, Trials_this_time) = False
'        End If
'    Next X
    For X = 1 To NCNUMBER
        If Nc(X).act = 1 Then
            Rasters(X + 4, bincounter5, Trials_this_time) = True
        Else
            'Rasters(x + 4, bincounter5, Trials_this_time) = False
        End If
    Next X
    For X = 1 To NumCF
        If CF(X).act = 1 Then Rasters(X, bincounter5, Trials_this_time) = True 'Else Rasters(x, bincounter5, Trials_this_time) = False
    Next X

ElseIf DoRasters = 1 Then
    For X = 1 To 36
        Select Case RRCellType(X)
            Case 1  'MF
                If MF(RRCellNum(X)) = 1 Then Rasters(X, bincounter5, Trials_this_time) = True
            Case 2  'PC
                If Pc(RRCellNum(X)).act = 1 Then Rasters(X, bincounter5, Trials_this_time) = True
            Case 3  'Nuc
                If Nc(RRCellNum(X)).act = 1 Then Rasters(X, bincounter5, Trials_this_time) = True
            Case 4  'cf
                If CF(RRCellNum(X)).act = 1 Then Rasters(X, bincounter5, Trials_this_time) = True
            Case 5  'gr
                If Gr(RRCellNum(X)).act = 1 Then Rasters(X, bincounter5, Trials_this_time) = True
            Case 6  'Go
                If Gol(RRCellNum(X)).act = 1 Then Rasters(X, bincounter5, Trials_this_time) = True
            Case 7  'Stellate
                If Bk(RRCellNum(X)).act = 1 Then Rasters(X, bincounter5, Trials_this_time) = True
            Case 8  'Basket
                If BCells(RRCellNum(X)).act = 1 Then Rasters(X, bincounter5, Trials_this_time) = True
            Case 0
        End Select
    Next X
    
End If

' ********************** plasticity monitor real time analysis ********************************************
If plasticity.Visible = True And DoRealTime = 1 Then
    plasticity.Cls
    For X = 1 To SYNUMBER
        plasticity.PSet (X, grWeight(X)), RGB(255, 255 * Gr(X).act, 255 * Gr(X).act)
    Next X
    DoEvents
'    For x = 1 To 600
'        plasticity.PSet (x * 10, mfweight(x) * 10), RGB(255, 255 * (1 - Gr(x).act), 255 * (1 - Gr(x).act))
'    Next x
End If

jav_Mfinputaverage(bincounter5) = 0#

' ************** Real time granule and Golgi readout

If ShowRealTime = 1 Then
    GWin.Cls
     GWin.DrawWidth = 2
    y = 1
    For i = 1 To 12000
        If X = 120 Then
            X = 1
            y = y + 1
        Else
            X = X + 1
        End If
        If Gr(i).act = 1 Then GWin.PSet (X, y), vbGreen
    Next i
     GWin.DrawWidth = 4
    y = 1
    X = 0
    For i = 1 To 900
        If X = 30 Then
            X = 1
            y = y + 1
        Else
            X = X + 1
        End If
       
        If Gol(i).act = 1 Then GWin.PSet (X * 4, y * 3.3), vbRed
        'GWin.PSet (x * 40, y * 33), vbRed
    Next i
    DoEvents

End If

'******************************** AudioMonitor **********************************************************
    If DoPCAudio > 0 Then
        If Pc(DoPCAudio).act = 1 Then Beep 300, 2
    ElseIf DoNCAudio > 0 Then
        If Nc(DoNCAudio).act = 1 Then Beep 300, 2
    ElseIf DoGoAudio > 0 Then
        If Gol(DoGoAudio).act = 1 Then Beep 300, 2
    ElseIf DoGrAudio > 0 Then
        If Gr(DoGrAudio).act = 1 Then Beep 300, 2
    ElseIf DoMFAudio > 0 Then
        If MF(DoMFAudio) = 1 Then Beep 300, 2
    ElseIf DoBSAudio > 0 Then
        If DoBSAudio < 1000 Then
            If Bk(DoBSAudio).act = 1 Then Beep 300, 2
        Else
            If BCells(DoBSAudio - 1000).act = 1 Then Beep 300, 2
        End If
    End If
    

'  ************************* Neuryalynx Recording *****************************************

If NeuralynxRecordingCells(0) = 1 Then
        If NeuralynxRecordingCells(1) > 0 Then
            'If Gr(27).act = 1 Then Debug.Print (bincounter + (TrialCounter * 5000) * 1000)
            If Gr(NeuralynxRecordingCells(1)).act = 1 Then
                NeuralynxSpikes(1, 0) = NeuralynxSpikes(1, 0) + 1
                NeuralynxSpikes(1, NeuralynxSpikes(1, 0)) = (Bincounter + (TrialCounter * 5000)) * 100
                'Debug.Print NeuralynxRecordingCells(1), NeuralynxSpikes(1, NeuralynxSpikes(1, 0))
            End If
        End If
End If

' ******************************* Simpson Analysis ******************************************
If DoSimpson = 1 Then SimpsonCalc

End Sub


Public Sub MFinput(Arg As Integer)
Dim mfspikes As Integer
Dim i As Integer
Dim j As Integer
Dim X As Integer
Dim a As Integer
Dim b As Integer
Dim Spike As Single
Dim PhasicStop As Single
Dim PhasicStop2 As Single
Dim r As Single
Dim avg As Single
Dim counter As Integer
Dim p As Integer
Dim Data(2499) As Integer
Dim CSisON As Integer
Dim f As String
Dim FakeTS As Double
Dim NttOut As tetdata
Dim CSFreq As Single
Dim temp As Single
Dim TempCounter As Integer
Dim Candidate As Integer
Dim CandidateSwapOut As Integer
Dim MFtemp As Integer
Dim k As Integer

csnumber = 1
If Arg = 0 Then
    TotalBins = TotalBins + 1
    mfspikes = 0
    TIME_START = Timer
    
    Bincounter = Bincounter + Time_step_size
    bincounter5_temp = bincounter5_temp + Time_step_size
    If bincounter5_temp >= 5 Then
        bincounter5_temp = 0
        bincounter5 = bincounter5 + 1
    End If
    If Time_step_size = 5 Or bincounter5_temp = 0 Then  'keep track of bincounter5 for plasticity eligibility
        If gr_elig_counter = 200 Then gr_elig_counter = 1 Else gr_elig_counter = gr_elig_counter + 1  '  200 is 5 msec times 200 to get to one second
        For i = 1 To SYNUMBER
            gr_elig(i, gr_elig_counter) = 0
        Next i
        Erase MfElig
    End If  'bincounter stuff
    
    If Bincounter > 5000 Then    'Start a new trial
        cbm_main.Caption = CStr(DebugCF) + "  " + CStr(Debug_weights_total)
        
        gGGAverage = gGGAverage / (900# * 5000#)
        ConductanceForm.GotoGoLabel.Caption = gGGAverage
        gGGAverage = 0
        
        DebugCF = 0
        Debug_weights_total = 0
        Erase CFCounter
        If cbm_main.Check2.Value = Checked Then SaveWeights
        If PM_Form.WindowState = 0 Then ShowWeights
        If Raster_Histos_ON = 1 Then HistoDivisor = HistoDivisor + 1
        WeightHistoryDraw
        If Trials_this_time = 1 Then OpenDatFile
        Bincounter = Time_step_size
        bincounter5 = 1
        bincounter5_temp = 0
        TrialCounter = TrialCounter + 1
        Trials_this_time = Trials_this_time + 1
        TrialsThisSession = TrialsThisSession + 1
        If cbm_main.PCNucCFCheck.Value = Checked Then
            Close #2
            Open "PCNucCFActivity" For Append As #2
            Write #2, TrialCounter,
            For X = 1 To PCNUMBER
                Write #2, PurkinjeActivity(X, 1),
                PurkinjeActivity(X, 1) = 0
            Next X
            For X = 1 To NCNUMBER
                Write #2, NucleusActivity(X, 1),
                NucleusActivity(X, 1) = 0
            Next X
            For X = 1 To NumCF
                Write #2, ClimbingFiberActivity(1, X),
                ClimbingFiberActivity(1, X) = 0
            Next X
            Write #2,
            Close #2
        End If
        
        If cbm_main.CompressedDatFile_Menu.Checked = False Then    ' normal data file mode
            For i = 1 To 2500
                Data(i - 1) = Int(((RN_Histo(i + 800)) - 2000) / 2.5)
            Next i
        Else                                                            ' compressed data file mode originally for subtraction experiments
            j = 0
            For i = 599 To 5000 Step 2
                Data(j) = Int((RN_Histo(i) + RN_Histo(i + 1)) / 5#)
                j = j + 1
            Next i
            X = Data(j)
            For i = j + 1 To 2500
                Data(i) = X
            Next i
        End If
        
        For i = 4 To 2499
            Data(i) = Int((Data(i - 4) + Data(i - 3) + Data(i - 2) + Data(i - 1) + Data(i)) / 5)
        Next i
        
        Put #24, , Data
        
        Erase RN_Histo
        If cbm_main.Weights_CHECK.Value = Checked And (TrialCounter Mod 10 = 0) Then SaveGrWeights
        
        If keep_the_time <> 0 Then
            caption_temp = "Total activity window" + Str$(Timer - keep_the_time)
            activity_window.Caption = caption_temp
        End If
        keep_the_time = Timer
        TIME_TOTAL = Timer - start_time
        If MFgrGolOnly = 0 Then
            If TIME_TOTAL > 0 And TIME_TOTAL < 1000 And TIME_MF > 0 And TIME_MF < 100 And TIME_GR > 0 And TIME_MF < 100 And TIME_GOL > 0 And TIME_GOL < 100 And TIME_Plasticity > 0 And TIME_Plasticity < 100 Then
                cbm_main.StatusBar1.Panels(1).Text = Str$(Format(TIME_TOTAL, "###.##")) + "  " + Str$(Format(TIME_MF, "##.##")) + "  " + Str$(Format(TIME_GR, "##.##")) + "  " + Str$(Format(TIME_GOL, "##.##")) + "  " + Str$(Format(TIME_Plasticity + TIME_PK, "##.##")) ' + "   " + Str$(Format(TIME_BK, "##.##")) + "   " + Str$(Format(TIME_PK, "##.##")) + "   " + Str$(Format(TIME_NUC, "##.##")) + "   " + Str$(Format(TIME_Plasticity, "##.##"))
            End If
        Else
            If TIME_TOTAL > 0 And TIME_TOTAL < 1000 And TIME_MF > 0 And TIME_MF < 100 And TIME_GR > 0 And TIME_MF < 100 And TIME_GOL > 0 And TIME_GOL < 100 Then
                cbm_main.StatusBar1.Panels(1).Text = Str$(Format(TIME_TOTAL, "###.##")) + "   " + Str$(Format(TIME_MF, "##.##")) + "   " + Str$(Format(TIME_GR, "##.##")) + "   " + Str$(Format(TIME_GOL, "##.##"))
            End If
        End If
        
        If STPFadeSession = 1 Then
            For i = 1 To SYNUMBER
                grSTP(i) = grSTP(i) * 0.96 '0.96
                SB_STP(i) = SB_STP(i) * 0.96
            Next i
        End If
        start_time = Timer
        ResetTimers
        
        If Trials_this_time > TrialThatEndsSession Then          'End of this session
            TrialsThisSession = 1
            MakeBVIFile
            TrialsPerSession(0, 1) = TrialsPerSession(0, 1) + 1
            SessionCounter = SessionCounter + 1
            SessionsThisExp = SessionsThisExp + 1
            If cbm_main.Check2.Value = Checked Then cbm_main.Command9.Value = True
            If SessionsThisExp > NumberofSessions Then  ' End of Experiment
                If cbm_main.ChimeCheck.Value = Checked Then Beep 800, 700
                If RepeatMode = 1 Then
                
                    Select Case UseSecondExperiment
                        Case 0  ' no second experiment, this is simple repeat mode
                            If RepeatExperimentsCounter < RepeatExperimentsGoal Then  'with no number specified, default is 32000 or forever
                                ReadyToRepeat
                            Else
                                keepgoing = 0
                                Close
                            End If
                        Case 1  ' alternate first and second experiment, repeat mode counter applies
                            If RepeatExperimentsCounter < RepeatExperimentsGoal Then  'with no number specified, default is 32000 or forever
                                SwapExperiments
                                ReadyToRepeat
                            Else
                                keepgoing = 0
                                Close
                            End If
                        Case 2  ' repeat second experiment
                            If RepeatSecondExperimentCounter = 0 Then
                                RepeatSecondExperimentCounter = 1
                                SwapExperiments
                                ReadyToRepeat
                            ElseIf RepeatSecondExperimentCounter = RepeatExperimentsGoal Then
                                RepeatSecondExperimentCounter = 0
                                SwapExperiments
                                ReadyToRepeat
                            Else
                                RepeatSecondExperimentCounter = RepeatSecondExperimentCounter + 1
                                ReadyToRepeat
                            End If
                    End Select
                Else
                    keepgoing = 0
                    Close
                End If
                
                If CommandLine = 1 Then
                    AutoSaveSimulation
                    cbm_main.QuitUserInterface
                End If
            Else
                TrialThatEndsSession = TrialThatEndsSession + TrialsPerSession(TrialsPerSession(0, 1), 1)
                OpenDatFile
                CurrentContext = SessionContexts(TrialsPerSession(0, 1))
                STPFadeSession = STPFades(TrialsPerSession(0, 1))
                If CurrentContext = 1 Then
                    cbm_main.ContextOption(0).Value = True
                Else
                    cbm_main.ContextOption(1).Value = True
                End If
                If raster_form.ResetHistoMenu(2).Checked = True Then
                    HistoDivisor = 1
                    Erase GR_histo
                    Erase Go_histo
                    Erase Stellate_histo
                    Erase Basket_histo
                End If
            End If
           
        End If  'Session?
         
        If keepgoing = 1 Then
            LoadTrial (Trials_this_time)
            cbm_main.StatusBar1.Panels(2).Text = CStr(TrialCounter) + " " + CStr(TrialsThisSession) + " of " + CStr(TrialsPerSession(TrialsPerSession(0, 1), 1))
            PC_form.Cls
            If AmpMode = 1 Then
                PC_form.Line (0, PC_form.ScaleHeight - (12 * AmpModeAmp))-(5000, PC_form.ScaleHeight - (12 * AmpModeAmp)), RGB(255, 150, 150)
            End If
            activity_window.Cls
             f = ""
            i = Len(SessionNames(TrialsPerSession(0, 1)))
            While f <> "\"
                i = i - 1
                f = Mid$(SessionNames(TrialsPerSession(0, 1)), i, 1)
            Wend
                 
            cbm_main.StatusBar1.Panels(3).Text = MasterTrialNames(Trials_this_time)
            cbm_main.StatusBar1.Panels(4).Text = CStr(TrialsPerSession(0, 1)) + "  " + Mid$(SessionNames(TrialsPerSession(0, 1)), i + 1, Len(SessionNames(TrialsPerSession(0, 1))) - i - 4)
        End If
        '********* Neuralynx stuff
        ' write data to disk
        For i = 1 To NeuralynxSpikes(1, 0)
            'Debug.Print NeuralynxSpikes(1, i)
            NttOut.TS = DoubletoTS(NeuralynxSpikes(1, i))
            Put 31, , NttOut
            NeuralynxSpikes(1, i) = 0
        Next i
        
        ' erase counters
        For i = 1 To 6
            NeuralynxSpikes(i, 0) = 0
        Next i
        
        Close 31
    End If   'end of trial?
    If Bincounter = 1 Then   ' Import mossy fiber spikes
        If AddMossyFiberForm.Check1.Value = Checked Then
            DoMossyFibersFromSpikes Trials_this_time
        End If
    End If
End If  'arg = 0
    
'*************************** Calculate MF activity ***************************************************************
CSisPresent = 0

If keepgoing = 1 Then
    
    
    For i = 1 To MFNUMBER
        CSisON = 0
        CSFreq = 0
'************Determine whether a CS is on **********************
        If CS_ON(1) = 1 And Bincounter > cs_onset(1) And Bincounter < cs_onset(1) + cs_duration(1) Then  'is the CS on and is it within its duration?
            If MFS(1, i).CStype = 1 Then                                                                   'if it's a tonic then yes
                CSisON = 1
                CSFreq = MFS(1, i).CSFreq
            End If
            If MFS(1, i).CStype = 5 And (Bincounter < cs_onset(1) + PHASICDUR) Then                      'if it's phasic then yes
                CSisON = 1
                CSFreq = MFS(1, i).CSFreq
            End If
            CSisPresent = 1
        End If
       
        If CS_ON(2) = 1 And Bincounter > cs_onset(2) And Bincounter < cs_onset(2) + cs_duration(2) Then  'is the CS on and is it within its duration?
            If MFS(1, i).CStype = 2 Then                                                        'if it's a tonic then yes
                CSisON = 1
                CSFreq = MFS(1, i).CSFreq
            End If
            If MFS(1, i).CStype = 6 And (Bincounter < cs_onset(2) + PHASICDUR) Then             'if it's phasic then yes
                CSisON = 1
                CSFreq = MFS(1, i).CSFreq
            End If
            CSisPresent = 1
        End If
    
        If CS_ON(3) = 1 And Bincounter > cs_onset(3) And Bincounter < cs_onset(3) + cs_duration(3) Then  'is the CS on and is it within its duration?
            If MFS(1, i).CStype = 3 Then                                                                'if it's a tonic then yes
                CSisON = 1
                CSFreq = MFS(1, i).CSFreq
            End If
            If MFS(1, i).CStype = 7 And (Bincounter < cs_onset(3) + PHASICDUR) Then                       'if it's phasic then yes
                CSisON = 1
                CSFreq = MFS(1, i).CSFreq
            End If
            CSisPresent = 1
        End If
    
        If CS_ON(4) = 1 And Bincounter > cs_onset(4) And Bincounter < cs_onset(4) + cs_duration(4) Then  'is the CS on and is it within its duration?
            If MFS(1, i).CStype = 4 Then                                                      'if it's a tonic then yes
                CSisON = 1
                CSFreq = MFS(1, i).CSFreq
            End If
            If MFS(1, i).CStype = 8 And (Bincounter < cs_onset(4) + PHASICDUR) Then                       'if it's phasic then yes
                CSisON = 1
                CSFreq = MFS(1, i).CSFreq
            End If
            CSisPresent = 1
            
        End If
       
        If UseUBCs = 1 And CS_ON(1) = 1 And Bincounter > cs_onset(1) + 200 And MFS(1, i).CStype = 4 Then
            If Bincounter < (cs_onset(1) + MFS(1, i).UBCduration) Then
                CSisON = 1
                If Bincounter > cs_onset(1) + 100 Then '250 Then  MBL2011
                    CSFreq = MFS(1, i).CSFreq
                Else
                    temp = cs_onset(1) + 250 - Bincounter
                    temp = 50 - temp
                    temp = temp / 50
                    If temp = 0 Then temp = 0.001
                    CSFreq = MFS(1, i).CSFreq * temp
                End If
            End If
        End If
        
        MFS(1, i).Thr = MFS(1, i).Thr + ((1 - MFS(1, i).Thr) * THRDECAYMF)
        
'        This determines whether each mossy fiber is active this trial

        If Rnd < ((CSisON * CSFreq) + MFS(CurrentContext, i).bfreq) * MFS(1, i).Thr Then
            MF(i) = 1
            MfElig(i) = 1
            MFS(1, i).Thr = 0
'            Debug.Print i
        Else
            MF(i) = 0
        End If
        If Arg > 0 Then
            If Arg = 2 Then
                MFS(1, i).Thr = 0
            End If
        End If
    Next i
    
    If ResponseDrivenMFs = 1 Then
        For j = 0 To 74 Step DrivenMFsStepper
            For i = 1 To 8
                If Nc(i).act = 1 Then
                    If Rnd() < 0.3 Then MF(i + (j * 8)) = 1
                End If
            Next i
        Next j
    End If
    

    '*******************************************
    
    If DoMFCollaterals = 1 Then
        TempCounter = 0
        For j = 1 To CollateralStepper
            For i = 1 To 8
                TempCounter = TempCounter + 1
                MF(i + (j * 120)) = Nc(i).act
                mfweight(i + (j * 120)) = 0   '  these are just collaterals, don't want them to have mf to nuc synapses.
            Next i
        Next j
    End If
    
    
   
    '  if CS is on and if number of swaps is greater than zero,
    '  randomly pick that number from CS1 and UBC
    '  randomly pick non-CS1/UBC
    '  make sure it's not a mfcollateral
    '  swap their spikes/non-spikes for this trial
    
    
    If CompetingStimulusNumber > 0 And CS_ON(1) = 1 And Bincounter > cs_onset(1) And Bincounter < cs_onset(1) + cs_duration(1) Then  ' the CS is on, otherwise no swapping
    
        If Bincounter = cs_onset(1) + 1 Then  ' first bin of this CS, so set up the swap table
            CompetingStimulusTotal = 0
            Erase CompetingStimulusSwap
            TempCounter = 0
            
            For i = 1 To MFNUMBER  ' figure out which MFs are CS1 phase, tonic or UBC  the count is stored in CompetingStimulusTotal
                If MFS(1, i).CStype = 1 Or MFS(1, i).CStype = 5 Or (UseUBCs * MFS(1, i).CStype = 4) Then
                    CompetingStimulusTotal = CompetingStimulusTotal + 1
                    CompetingStimulusSwap(1, CompetingStimulusTotal) = i
                End If
            Next i
            
            If DoMFCollaterals = 1 Then
                For j = 1 To CollateralStepper   ' figure out which MFs are collaterals   the count is stored in TempCounter
                    For i = 1 To 8
                        TempCounter = TempCounter + 1
                        CompetingStimulusSwap(2, TempCounter) = i + (j * 120)
                    Next i
                Next j
            End If
            
        ' now CompetingStimuluSwap(1,1 to CompetingStimulusTotal) lists the CS phasic, tonic and UBC MFs
        ' and CompetingStimuluSwap(2,1 to TempCounter) lists the collateral MFs (these are excluded from swapping
            
            i = 0
            Randomize Timer
            While i < CompetingStimulusNumber  'populate the swap list for this CS
                j = 1
                While j = 1   ' this picks a candidate MF to swap out
                    CandidateSwapOut = (Rnd() * CompetingStimulusTotal - 1) + 1
                    For k = 1 To 100                ' make sure it hasn't already been picked
                        If CompetingStimulusSwap(3, k) = CompetingStimulusSwap(1, CandidateSwapOut) Then CandidateSwapOut = 0
                    Next k
                    If CandidateSwapOut <> 0 Then j = 0  'if not already picked we'll keep it and eject from this loop, otherwise try again
                Wend
                
                Candidate = Rnd() * 599 + 1
                For j = 1 To CompetingStimulusTotal
                    If Candidate = CompetingStimulusSwap(1, j) Then Candidate = 0  ' candidate is already a CS MF
                Next j
                For j = 1 To TempCounter
                    If Candidate = CompetingStimulusSwap(2, j) Then Candidate = 0  ' candidate is a DCN collateral and not available to swap
                Next j
                For j = 1 To CompetingStimulusTotal
                    If Candidate = CompetingStimulusSwap(3, j) Then Candidate = 0  ' candidate is already on the swap list
                Next j
                
                If Candidate > 0 Then   ' if it's an acceptable swapper, put it on the swap list and increment i
                    CompetingStimulusSwap(3, CandidateSwapOut) = Candidate
                    'Debug.Print i, Candidate
                    i = i + 1
                End If
            Wend
            DoEvents
            
        End If
        '' do the actual swapping right here
            For i = 1 To CompetingStimulusTotal
                'Debug.Print CompetingStimulusSwap(1, i), CompetingStimulusSwap(2, i), CompetingStimulusSwap(3, i)
                If CompetingStimulusSwap(3, i) <> 0 Then
                    MFtemp = MF(CompetingStimulusSwap(3, i))
                    MF(CompetingStimulusSwap(3, i)) = MF(CompetingStimulusSwap(1, i))
                    MF(CompetingStimulusSwap(1, i)) = MFtemp
                End If
            Next i
    End If
    
    
    
    ' ****** Update "presynaptically stored" conductances of MF to gr and MF to Gol
    
    For i = 1 To MFNUMBER
    
        MF2(i).gGrAMPA = MF2(i).gGrAMPA * grAMPADecayMF
        MF2(i).gGrNMDA = MF2(i).gGrNMDA * grNMDADecayMF
        MF2(i).gGol = MF2(i).gGol * gEDecayGoMF
        
        MF2(i).gGrAMPA = MF2(i).gGrAMPA + (1 - MF2(i).gGrAMPA) * (MF(i) * 0#)
        MF2(i).gGrNMDA = MF2(i).gGrNMDA + (1 - MF2(i).gGrNMDA) * (MF(i) * gEconstGr)
        MF2(i).gGol = MF2(i).gGol + (1 - MF2(i).gGol) * (MF(i) * gEconstGoMF)
        
    Next i
   
    '*******************************************************************************
    
    
    If Arg = 0 Then
        If Bincounter = cs_onset(1) And CS_ON(1) = 1 Then
            PC_form.Line (Bincounter, PC_form.Height)-(Bincounter + cs_duration(1), 0), RGB(0, 0, 80), BF
            raster_form.Line (Bincounter, raster_form.Height)-(Bincounter + cs_duration(1), 0), RGB(0, 0, 80), BF
        ElseIf Bincounter = cs_onset(2) And CS_ON(2) = 1 Then
            PC_form.Line (Bincounter, PC_form.Height)-(Bincounter + cs_duration(2), 0), RGB(0, 0, 90), BF
            raster_form.Line (Bincounter, raster_form.Height)-(Bincounter + cs_duration(2), 0), RGB(0, 0, 90), BF
        ElseIf Bincounter = cs_onset(3) And CS_ON(3) = 1 Then
            PC_form.Line (Bincounter, PC_form.Height)-(Bincounter + cs_duration(3), 0), RGB(0, 0, 100), BF
            raster_form.Line (Bincounter, raster_form.Height)-(Bincounter + cs_duration(3), 0), RGB(0, 0, 100), BF
        ElseIf Bincounter = cs_onset(4) And CS_ON(4) = 1 Then
            PC_form.Line (Bincounter, PC_form.Height)-(Bincounter + cs_duration(4), 0), RGB(0, 0, 110), BF
            raster_form.Line (Bincounter, raster_form.Height)-(Bincounter + cs_duration(4), 0), RGB(0, 0, 110), BF
        End If
        If Bincounter = US_onset(1) And raster_form.cell_menu2(6).Checked = True Then
            PC_form.PaintPicture PuffImage, US_onset(1) - 190, 820, 380, 90
            raster_form.PaintPicture PuffImage, US_onset(1) - 190, 725, 380, 55
        End If
        T1 = Timer
        TIME_MF = TIME_MF + T1 - TIME_START
    End If
    
    If AddMossyFiberForm.Check1.Value = Checked Then
        For i = 1 To MFsAddedTotal
            MF(MFsAddedIdentity(i)) = MFsAddedSpikes(i, Bincounter)
        Next i
    End If
End If 'keepgoing
End Sub




Public Sub GetInputFile(filename As String)
Dim Indat(200)
Dim indat2(200)
Dim i As Integer

    Close #1
    Open filename For Input As #1
    i = 1
    While Not EOF(1)
        Input #1, Indat(i)
        
        If Left(Indat(i), 1) = "*" Or Indat(i) = "" Then
        
        Else
            Input #1, indat2(i)
            Debug.Print Indat(i), indat2(i)
            i = i + 1
        End If
       
    Wend

End Sub



























