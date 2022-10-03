Attribute VB_Name = "StellateCells"
Option Explicit
Public Const STELLATENUMBER = 240
Public Const PFStellsynUMBER = 50

Public GLeakStell As Single
Public WEIGHTSTELL As Single
Public Const ELeakStell = -60
Public Const THRMAXStell = 0
Public Const THRBASEStell = -50      '/*-58*/
Public GDecayGrStell As Single
Public ThrDecayStell As Single
Public Const ThrTauStell = 22  '0.2
Public Const GTauGrStell = 4.15 ' 0.3

Public Type Stellate
    act As Integer
    GGr As Single
    v As Single
    Thr As Single
End Type

Public Bk(STELLATENUMBER) As Stellate

Public Sub Stellate_TDV()
    GDecayGrStell = Exp(-Time_step_size / GTauGrStell)
    ThrDecayStell = 1 - Exp(-Time_step_size / ThrTauStell)
    GLeakStell = 0.2 / (6 - Time_step_size)      '/*0.03-> Basket=50Hz*/
    WEIGHTSTELL = 0.2
    
End Sub
