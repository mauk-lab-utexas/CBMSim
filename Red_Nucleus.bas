Attribute VB_Name = "Red_Nucleus"
Option Explicit

Public Const gETauRN = 15
Public Const ELeakRN = 0
Public gLeakRN As Single
Public gEDecayRN As Single

Public Type RedNucleusCell
  Vm As Single
  gE(NCNUMBER) As Single
End Type

Public RN As RedNucleusCell

Public Sub RN_TDV()
Dim X As Integer
    gLeakRN = 0.025 / (6 - Time_step_size)
    gEDecayRN = Exp(-Time_step_size / gETauRN)
End Sub

Public Sub DoWorkRN()
Dim i As Integer
Dim gE As Single
    gE = 0
    For i = 1 To NCNUMBER
        RN.gE(i) = RN.gE(i) * gEDecayRN
        RN.gE(i) = RN.gE(i) + (Nc(i).act * 0.012)
        gE = gE + RN.gE(i)
    Next i
    gE = gE - 0.05
    If gE < 0 Then gE = 0
    gE = gE * gE * gE * 5#
    If gE < 0.02 Then gE = 0
    RN.Vm = RN.Vm + (gLeakRN * (-1 * RN.Vm)) + (gE * (80 - RN.Vm))
    RN_Histo(Bincounter) = RN.Vm * 100
End Sub
