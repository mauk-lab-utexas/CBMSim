Attribute VB_Name = "Purkinje_cells"
Option Explicit
Public Const PCNUMBER = 24
Public Const PFPCSYNUMBER = 500   '$$2000
Public Const StellPCSYNUMBER = 10
Public Const BasketPCSYNUMBER = 16  'New Basket


Public GLeakPC  As Single
Public Const VStellPC = -80
Public Const ELEAKPC = -60  '-60 mike
Public Const THRMAXPC = -48
Public Const THRBASEPC = -60  ' /*-56*/
Public ThrDecayPC As Single

Public GDecayGRPC As Single
Public GDecayStellPC As Single
Public GDecayBCPC As Single
Public Const ThrtauPC = 5 '3.11 ' /*0.5*/
Public Const GtauGRPC = 4.15
Public Const GtauStellPC = 4.15
Public Const GtauBCPC = 5

Public GCONSTGRPC As Single
Public GCONSTStellPC As Single
Public GCONSTBCPC As Single

Public gPURKtoNUCLEUS(PCNUMBER, NCNUMBER) As Single

Public PCFormColors As Integer
Public PCcolor As ColorConstants
Public NCcolor As ColorConstants
Public CFcolor As ColorConstants

Public Type Purkinje
    act As Integer
    GGr As Single
    GStell As Single
    GBC As Single               'New Basket
'    LTDmeter(10) As Single
'    LTPmeter(10) As Single
    v As Single
    Thr As Single
    Stellsyn(StellPCSYNUMBER) As Integer
    ThrBase As Single
End Type

Public StellPcG(STELLATENUMBER) As Single

Public Pc(PCNUMBER) As Purkinje
Public PurkActivity(PCNUMBER) As Long

Public Sub PC_TDV()
Dim i As Integer
Dim j As Integer
    GDecayGRPC = Exp(-Time_step_size / GtauGRPC)  '/*0.1, 0.35*/
    GDecayStellPC = Exp(-Time_step_size / GtauStellPC)
    GDecayBCPC = Exp(-Time_step_size / GtauBCPC)
    ThrDecayPC = 1 - Exp(-Time_step_size / ThrtauPC)  ' /*0.5*/
    GLeakPC = 0.2 / (6 - Time_step_size)    '/*0.6, 0.12*/
    
    GCONSTGRPC = 0.0096
    GCONSTStellPC = 0.009 '.009 for basket cell paper prior to 2018C, .011 is 2018d
    GCONSTBCPC = 0.1   '.14 for bcreviewers, .1 normal
End Sub

    
