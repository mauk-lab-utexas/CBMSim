Attribute VB_Name = "granule_cells"
Option Explicit
'***************** Granule cell parameters *******
Public Const MaxGrDend = 4    '7 for 5E  Don't forget to change do work
Public Const MinGrDend = 4
Public Const GrX = 120
Public Const GrY = 100
Public Const grTau = 6


Public gEconstGr As Single
Public Const grNMDADecayTau = 55
Public grNMDAdecay As Single


Public gEDecayGr As Single
Public Const gETauGr = 4.5
Public gIconstGr As Single
Public gIDecayGr As Single
Public Const gITauGr = 25 '50

Public Const gITauGrFast = 3
Public gIFastDecayGr As Single
Public Const GABAtau = 300
Public GABAdecay As Single

Public gLGr  As Single
Public Const EGABAgr = -80 '*** 2014 -70

Public Const ThrmaxGr = -20 '-20 Feb2008
Public Const ThrBasegr = -40 '-40 '-50 Feb2008
Public Const ThrTauGr = 3
Public ThrDecayGr As Single

Public Const ELeakgr = -70 '*** 2014 -85
Public Const ThrBaseVarGr = 0#
Public prex(MaxGrDend) As Integer
Public GolGrHeterogenous As Integer
Public SaveRastersON As Integer
Public SaveWeightsON As Integer

Public Type gran
  act As Integer
  numdend As Integer
  Thr As Single
  v As Single
  gi As Single
  MF(MaxGrDend) As Integer
  Gol(MaxGrDend) As Integer
  'gE As Single
'  gE(4) As Single
  g_KCa As Single
'  gE_NMDA As Single
  gIfast(MaxGrDend) As Single
  gIslow(MaxGrDend) As Single
  fastMask(MaxGrDend) As Byte
  g_Var As Single   'Mike in May added
  ThrBase As Single
End Type

Public Gr(GrX * GrY) As gran


Public Sub gr_TDV()
Dim X As Integer
    gLGr = 0.1 / (6 - Time_step_size)        '$$0.07
    gIDecayGr = Exp(-Time_step_size / gITauGr)
    gEDecayGr = Exp(-Time_step_size / gETauGr)
    ThrDecayGr = 1 - Exp(-Time_step_size / ThrTauGr)
    gIFastDecayGr = Exp(-Time_step_size / gITauGrFast)
    GABAdecay = Exp(-Time_step_size / GABAtau)
    'grAMPAdecay = Exp(-Time_step_size / granuleAMPATau)
    grNMDAdecay = Exp(-Time_step_size / grNMDADecayTau)
    'grgKDecay = Exp(-Time_step_size / grgKTau)
    
        gEconstGr = 0.0145 * MFtoGr '0.0125 0.02 '  MAY2014  0.02 '.025  0.035  'was .015 before 1-gE test Mike in May  0.025
          
        gIconstGr = 0.01 * GOtoGr '0.009  .03  0.022 '0.035 for slow '0.03 Feb2008
       
       
        For X = 1 To SYNUMBER
            Gr(X).g_Var = gEconstGr
        Next X
    
End Sub


