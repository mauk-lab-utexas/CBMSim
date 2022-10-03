Attribute VB_Name = "BasketCells"
Option Explicit

Public Const BasketNUMBER = 96
Public Const PFBasketsynUMBER = 125
Public Const PCBasketsynUMBER = 4

Public GLeakBC As Single
Public GrBCWeights As Single
Public PCBCWeights As Single
Public Const ELeakBC = -70
Public Const ThrMAXBC = 0
Public Const ThrBaseBC = -50 '-45 BCreviewers ' -50 normal
Public GDecayGrBC As Single
Public GDecayPCBC As Single
Public ThrDecayBC As Single
Public Const ThrTauBC = 8 ' 10   2015 for basket cell paper correlations
Public Const GTauGrBC = 10 '4.15
Public Const GTauPCBC = 5

Public UsePCtoBasketSynapses As Integer

Public TonicgI_Basket As Single

Public Type Basket
    act As Integer
    GGr As Single
    v As Single
    Thr As Single
    gPc As Single
    PCsyn(PCBasketsynUMBER) As Integer
    grW(PFBasketsynUMBER) As Single
    gE(PFBasketsynUMBER) As Single
End Type

Public BCells(BasketNUMBER) As Basket

Public Sub BC_TDV()
    GDecayGrBC = Exp(-Time_step_size / GTauGrBC)
    GDecayPCBC = Exp(-Time_step_size / GTauPCBC)
    
    ThrDecayBC = 1 - Exp(-Time_step_size / ThrTauBC)
    
    GLeakBC = 0.38 / (6 - Time_step_size)      '/*0.03-> Basket=50Hz*/
    
    GrBCWeights = 0.35 '0.35=normal '0.4= BCreviewers
    PCBCWeights = 0.25 '
    TonicgI_Basket = 0.3
End Sub
