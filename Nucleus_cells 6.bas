Attribute VB_Name = "Nucleus_cells"
Option Explicit
Public Const NCNUMBER = 8
Public Const MFNCSYNUMBER = 75  '60
Public Const PCNCMaxSYNUMBER = 12

Public PCNCSYNUMBER As Integer  '9 or 6

Public NCTimeConstant As Single
Public Const NCTau = 6

Public GLeakNC  As Single

Public Const ELEAKNC = -65      '$$50  /*62*/
Public Const THRMAXNC = -40   ' /*-56*/
Public Const THRBASENC = -70 '-69#        '/*-53*/

Public grNMDActivate As Single
Public gNMDADecayMFNC As Single
Public gAMPADecayMFNC As Single

Public Const gtauAMPA = 6
Public Const GtauNMDA = 50
Public I_NMDA(NCNUMBER) As Single

Public GDecayPCNC As Single
Public Const GtauPCNC = 4.15
Public Const VPCNC = -80

Public Const gPurktoNucBeginAverage = 0.17 '  0.2 when there was convergence 9, now convergence 12



Public Const ThrtauNC = 5
Public Const GCONSTMFNC = 0.008 '0.022 '.025
Public THRdecayNC As Single

Public Const ELDECAYMF = 0.2
Public Type Nucleus
    act As Integer
    Thr As Single
    v As Single
    PCsyn(PCNCMaxSYNUMBER) As Integer
    gAMPA(MFNCSYNUMBER) As Single
    AMPABind(MFNCSYNUMBER) As Single
    gNMDA(MFNCSYNUMBER) As Single
    NMDABind(MFNCSYNUMBER) As Single
    
    gGABA(PCNCMaxSYNUMBER) As Single
    gK As Single
    gCa As Single
    
    gPc As Single
    gMF As Single
    gMF2 As Single
    gNUCtoCF(12) As Single
End Type

Public Nc(NCNUMBER) As Nucleus

Public Sub NUC_TDV()
Dim i As Integer
    GDecayPCNC = Exp(-Time_step_size / GtauPCNC)
    gNMDADecayMFNC = Exp(-Time_step_size / GtauNMDA)
    gAMPADecayMFNC = Exp(-Time_step_size / gtauAMPA)
    
    THRdecayNC = 1 - Exp(-Time_step_size / ThrtauNC)
    NCTimeConstant = 1 - Exp(-Time_step_size / NCTau)
    grNMDActivate = 1 - Exp(-Time_step_size / 3)
    GLeakNC = 0.15 / (6 - Time_step_size)       '$$ 0.06  'Javs was .04
    
End Sub
