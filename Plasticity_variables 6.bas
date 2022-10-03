Attribute VB_Name = "Plasticity_variables"
Option Explicit
Public WEIGHTS_filename As String
Public grWEIGHTS_filename As String
Public mfWEIGHTS_filename As String

Public CHANGEgrWEIGHTS As Integer
Public CHANGEmfWEIGHTS As Integer

Public Const SYNUMBER = 12000   ' /*Set it to GrX*GrY*/
Public gr_elig(SYNUMBER, 200) As Single
Public gr_elig_counter As Integer

'Public Const LTDURATION = 100
Public LTDURATION As Integer
Public LTD_OFFSET As Integer


Public DELMINUS As Single ' -0.0008 ' Mike in June -0.00015     '$$-0.002         '/*0.0015*/
Public Const DELPLUS = 0.0004 ' 0.00008 'Mike in June 0.000015   '$$ 0.00007        '/*(0.00021)*/

Public Const INTPCRUNAVWINDOW = 8    '/*For array declaration we need integer*/
Public Const INTPCRUNAVWINDOW2 = 250
Public Const PCRUNAVWINDOW = 8#
Public Const PCRUNAVWINDOW2 = 250#
Public Const MFDELMINUS = -0.0000025 '-0.000005    '/*(-0.00012)*/
Public Const MFDELPLUS = 0.0002  '0.001  '$$0.005   '/*(0.00012)*/

Public Const MF_LTP_threshold = 3#
Public Const MF_LTD_threshold = 9

Public Const LTP_Decrement = 0.4   ' The percent of LTP that recovers between sessions

Public runavpc(INTPCRUNAVWINDOW) As Single
Public runavpc2(INTPCRUNAVWINDOW2) As Single
Public ltdtime(12) As Integer
Public ltdperiod As Integer
Public ltpperiod As Integer
Public mfltptime As Integer
Public mfltpperiod As Integer
Public mfltdperiod As Integer
Public pcrunavcounter As Integer
Public pcrunav As Single

Public grWeight(SYNUMBER) As Single
Public PreviousgrWeight(SYNUMBER) As Single
Public mfweight(MFNUMBER) As Single

Public grweightChangeThisSession(SYNUMBER, 2) As Single 'used to keep track of recent changes in order to implement STP

Public MfElig(MFNUMBER) As Single
Public MAXDRUS As Single

'********* STP at gr to PC synapses
Public Const grEligTau = 20# '25#
Public Const grEligBase = 0.002   '.002  .04
Public Const STPAmp = 0.00025 '.0005 for big runs
Public Const grSTPMax = 0.075 '.1 was too much

Public Const sb_STPMax = 0.4

Public grPreElig(SYNUMBER) As Single
Public grPreEligStatus(SYNUMBER) As Integer
Public grPreEligDecay As Single
Public grPreEligTimeout(SYNUMBER) As Integer

Public DoLTPShortTermPlasticity As Integer
Public Const DoStellateBasketSTP = 0
Public grSTP(SYNUMBER) As Single
Public SB_STP(SYNUMBER) As Single
Public grSTPAmp As Single
Public STPFadeSession As Integer
Public STPFades(25) As Integer
Public STPFades2(25) As Integer
Public STPFadesTEMP(25) As Integer
Public STPblockedbyCF As Integer

Public grPCPlasticityType As Integer
Public Const grWeightMax = 0.7
Public Const grWeightMin = 0.1
Public Const BinaryWeightRate = 200
Public Const BinaryProbPlus = 0.0025
Public Const BinaryProbMinus = 0.0225

'**************** Homeostatic plasticity
Public PCHomeoValue(PCNUMBER) As Single
Public NCHomeoValue(NCNUMBER) As Single

Public MLIPlasticity As Integer


Sub STP_TDV()
Dim m As Integer
    grPreEligDecay = 1 - Exp(-Time_step_size / grEligTau)
    For m = 1 To SYNUMBER
        grPreElig(m) = grEligBase
        grPreEligTimeout(m) = 0
    Next m
End Sub
