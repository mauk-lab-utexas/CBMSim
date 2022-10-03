Attribute VB_Name = "Golgi_cells"
Option Explicit
Public Const GOLnumber = 900
Public Const GoY = 30
Public Const GoX = 30
Public Const numGoGrDend = 256 '64
Public Const numGoGlDend = 4

Public gEconstGoMF As Single
Public gEconstGoGr As Single
Public gEDecayGoMF As Single
Public gEDecayGoGr As Single
Public Const ELGo = -70
Public ThrDecayGo As Single
Public Const ThrmaxGo = -2#  '-10#
Public Const ThrBaseGo = -32
Public Const ThrBaseVarGo = 0#
Public Const gETauGoMF = 4.5
Public Const gETauGoGr = 55 '4.5  '40 for mGluR and 4.5 for not?
Public Const ThrTauGo = 20
Public gLGo   As Single

Public gSlowDecayGo As Single
Public Const SlowTauGo = 100 '100
Public Const GolMGluR = 0 ' 0.0001  .0001 for mGluR and 0 for not?

Public wGG(900, 8) As Single
Public GG(900, 8) As Integer

Public GGconst As Single
Public DoGG As Integer
Public Const GGPercent = 70
Public Const GGdefault = 0.05

Public gGGAverage As Double

Public Type golgi
  act As Integer
  ThrBase As Single
  preGr(numGoGrDend) As Integer
  preGl(numGoGlDend) As Integer
  MF(numGoGlDend) As Integer
  
  gGG As Single
  
  gMF As Single
  g_varMF As Single
  g_varGr As Single
  
  GGr As Single
  
  Thr As Single
  v As Single
  gFast As Single
  gFastFinal As Single
  gSlow As Single
  
End Type

Public Gol(GoX * GoY) As golgi

Public Sub Gol_TDV()
Dim X As Integer

    gLGo = 0.025 / (6 - Time_step_size)
    gEDecayGoMF = Exp(-Time_step_size / gETauGoMF)
    gEDecayGoGr = Exp(-Time_step_size / gETauGoGr)
    gSlowDecayGo = Exp(-Time_step_size / SlowTauGo)
    ThrDecayGo = 1 - Exp(-Time_step_size / ThrTauGo)

   
    gEconstGoMF = 0.04 '0.05
    gEconstGoGr = 0.0015 '0.014  this is the value for 64 dendrites
   
    For X = 1 To GoX * GoY
        Gol(X).g_varMF = gEconstGoMF * MFtoGo
        Gol(X).g_varGr = gEconstGoGr * GRtoGo
    Next X
    GGconst = 0.035 ' 0.02 '0.006
End Sub
