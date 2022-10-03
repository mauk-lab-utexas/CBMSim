Attribute VB_Name = "Climbing_fibers"
Option Explicit
Public Const GTauNCCF = 4.15
Public Const ThrTauCF = 122

Public GDecayNCCF As Single
Public ThrDecayCF As Single

Public GCONSTNCCF As Single

Public Const VNCCF = -80
Public Const ELEAKCF = -60
Public Const THRMAXCF = 10
Public Const THRBASECF = -61

Public Const gNuctoCFBeginAverage = 0.13

Public CFCoupled As Integer

Public CF_spike_counter(12) As Integer

'Asynchronous release parameters

Public Const CFProbIncN = 0.2
Public Const CFProbIncTau = 0.8
Public Const CFProbDecayTTau = 1
Public Const CFProbDecayTN = 40
Public Const CFProbDecayTO = 78

Public Const gIncN = 0.3
Public Const gIncTau = 300
Public Const gDecayTTau = 70
Public Const gDecayTN = 50
Public Const gDecayTO = 56

' ******************************

Public NcCfG(NCNUMBER, 12) As Single
Public Type Climbing_fiber
    act As Integer
    GNc As Single
    v As Single
    Thr As Single
    NucInput(NCNUMBER) As Integer
    vE As Single
    GLeak As Single
    CurrentP(NCNUMBER) As Single
    CurrentG(NCNUMBER) As Single
End Type

Public NCCFSYNUMBER As Integer
Public NumCF As Integer

Public CFAsynchronouse As Integer

Public CF(12) As Climbing_fiber

Public Sub CF_TDV()
Dim i As Integer
Dim Changer As Single

    GDecayNCCF = Exp(-Time_step_size / GTauNCCF)
    ThrDecayCF = 1 - Exp(-Time_step_size / ThrTauCF)
    'Debug.Print "Climbing fiber", ThrDecayCF
    Randomize Timer
    For i = 1 To 12
        CF(i).GLeak = 0.1 / (6 - Time_step_size)
        'Debug.Print i, CF(i).GLeak
        
        'If SimOptionsForm.UniformityOption(1).Value = True Then
            Changer = (Rnd() - 0.5) * 0.25
            CF(i).GLeak = CF(i).GLeak + (CF(i).GLeak * Changer)
        'End If
        'Debug.Print i, CF(i).GLeak
    Next i
'    CF(1).GLeak = 0.019
'    CF(2).GLeak = 0.022
'    CF(3).GLeak = 0.018
'    CF(4).GLeak = 0.021
    
    GCONSTNCCF = 0.033 '0.027 '0.018  this is outdated and  is not used anymore
    If Time_step_size = 1 Then GCONSTNCCF = GCONSTNCCF * 0.3
End Sub
