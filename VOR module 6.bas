Attribute VB_Name = "VOR"
Option Base 0
Option Explicit

'***************************************************
' ********* VOR related variables ******************
'***************************************************

Public Const numVest_Afferents = 500
Public Const ThrBaseVest = -60
Public Const ThrmaxVest = 0
Public Const ELeakVest = -58
Public Const gEConstVest = 0.03
Public vest_in As Single
Public Vest_L(numVest_Afferents) As Vestibular_afferents
Public Vest_R(numVest_Afferents) As Vestibular_afferents
Public Const numPVP = 200
Public Const ThrBasePVP = -55
Public Const ThrMaxPVP = 0
Public Const ELeakPVP = -60
Public Const gEConstPVP = 0.18
Public Const gLeakPVP = 0.2
Public PVP_L(numPVP) As PVP_neurons
Public PVP_R(numPVP) As PVP_neurons
Public MN_L As motor_neurons
Public MN_R As motor_neurons

'*** VESTIBULAR CELLULAR PROPERTIES *************
Public Const ThrDecayVest = 0.27
Public Const gEDecayVest = 0.2
Public Const ThrDecayPVP = 0.2
Public Const gEDecayPVP = 0.3
Public Const gIDecayPVP = 0.2
Public GLeakVest As Single

Public Type Vestibular_afferents
    act As Integer
    V As Single
    Thr As Single
    gVel As Single
    gAcc As Single
    gVelNoise As Single
    gAccNoise As Single
    gVelDecayConst As Single
    gAccDecayConst As Single
    gVelConst As Single
    gAccConst As Single
    last_spike As Integer
End Type

Public Type PVP_neurons
    act As Integer
    V As Single
    Thr As Single
    gE As Single
    gi As Single
    last_spike As Integer
End Type

Public Type motor_neurons
    act As Integer
    V As Single
    Thr As Single
    gE As Single
    gi As Single
End Type

Public Sub Vest_TDV()
    GLeakVest = 0.02
End Sub

Public Sub init_VOR()
    Dim i As Integer
    Dim X As Single

    For i = 1 To numVest_Afferents
        Vest_L(i).act = 0
        Vest_L(i).gAcc = 0
        Vest_L(i).gVel = 0
        Vest_L(i).Thr = ThrBaseVest + Rnd() * 20
        Vest_L(i).V = ELeakVest + Rnd() * 0
        Vest_L(i).gVelConst = gEConstVest
        Vest_L(i).gVelDecayConst = 0.2
        Vest_L(i).gAccDecayConst = 0.2
        Vest_L(i).gVelNoise = 0.003

        X = (i * 1#) * (1 / numVest_Afferents)
        Vest_L(i).gAccConst = (gEConstVest * 5) * X
        Vest_L(i).gAccNoise = 0.012 * X

        Vest_R(i).act = 0
        Vest_R(i).gAcc = 0
        Vest_R(i).gVel = 0
        Vest_R(i).Thr = ThrBaseVest + Rnd() * 20
        Vest_R(i).V = ELeakVest + Rnd() * 0
        Vest_R(i).gVelConst = gEConstVest
        Vest_R(i).gVelDecayConst = 0.2
        Vest_R(i).gAccDecayConst = 0.2
        Vest_R(i).gVelNoise = 0.003

        X = (i * 1#) * (1 / numVest_Afferents)
        Vest_R(i).gAccConst = (gEConstVest * 5) * X
        Vest_R(i).gAccNoise = 0.012 * X
    Next i

    For i = 1 To numPVP
        PVP_L(i).act = 0
        PVP_L(i).V = ELeakPVP
        PVP_L(i).gE = 0
        PVP_L(i).Thr = ThrBasePVP
        PVP_L(i).gi = 0
        PVP_L(i).last_spike = 0

        PVP_R(i).act = 0
        PVP_R(i).V = ELeakPVP
        PVP_R(i).gE = 0
        PVP_R(i).Thr = ThrBasePVP
        PVP_R(i).gi = 0
        PVP_R(i).last_spike = 0
    Next i
    MN_L.act = 0
    MN_L.V = -70
    MN_L.gE = 0
    MN_L.gi = 0
    MN_L.Thr = 0
    MN_R.act = 0
    MN_R.V = -70
    MN_R.gE = 0
    MN_R.gi = 0
    MN_R.Thr = 0

End Sub


