Attribute VB_Name = "MossyFibers"
Option Explicit

Public Const MFNUMBER = 600      '/*Set it to MFNUMBER*/

Public Const NUMCONTEXT = 0.1 '.1 0.03  MAY2014
Public Const NUMCONTEXT2 = 0.1

Public Const NUMPHASIC = 0.03
Public Const NUMTONIC = 0.03  '0.03 MAY2014

Public Const NUMPHASIC2 = 0
Public Const NUMTONIC2 = 0.03

Public Const NUMPHASIC3 = 0
Public Const NUMTONIC3 = 0.03

Public Const NUMPHASIC4 = 0
Public Const NUMTONIC4 = 0.03

Public Const PHASICDUR = 40 '30

Public CurrentContext As Integer

'these are used for the backgound rates for CS cells
Public MFBGROUNDFREQMIN_CS As Single
Public MFBGROUNDFREQMAX_CS As Single

Public MFCONTEXTFREQMIN As Single
Public MFCONTEXTFREQMAX As Single

Public MFCONTEXTFREQMIN2 As Single
Public MFCONTEXTFREQMAX2 As Single

'these are used as the background rates for all other cells
Public MFBGROUNDFREQMIN As Single
Public MFBGROUNDFREQMAX As Single

Public MFTONICFREQ_INCREMENT As Single
Public MFTONICFREQ_INCREMENT2 As Single
Public MFTONICFREQ_INCREMENT3 As Single
Public MFTONICFREQ_INCREMENT4 As Single

Public MFPHASICFREQ_INCREMENT As Single
Public MFPHASICFREQ_INCREMENT2 As Single
Public MFPHASICFREQ_INCREMENT3 As Single
Public MFPHASICFREQ_INCREMENT4 As Single

Public Const ThrtauMF = 4 '0.1
Public THRDECAYMF As Single

Public MFsAdded As Integer
Public MFsAddedTotal As Integer
Public MFsAddedFilenames(32) As String
Public MFFilenames(200) As String
Public MFsAddedSequence(128, 200) As Integer
Public MFsAddedSpikes(128, 5000) As Byte
Public MFsAddedNumber(32) As Integer
Public MFsAddedIdentity(200) As Integer
Public MFsAddedS(32) As Integer
Public ResponseDrivenMFs As Integer
Public DoMFCollaterals As Integer
Public DrivenMFsStepper As Integer
Public CollateralStepper As Integer

Public grAMPADecayMF As Single
Public grNMDADecayMF As Single
Public GolDecayMF As Single

Public Const grAMPATau = 4
Public Const grNMDATau = 25
Public Const grAMPAConst = 0.05
Public Const grNMDAConst = 0.05

Public Type mossy
  Thr As Single
  bfreq As Single
  CSFreq As Single
  CStype As Integer
  UBCduration As Integer
End Type

Public Type Mossy2
  gGrAMPA As Single
  gGrNMDA As Single
  gGol As Single
End Type

Public MFS(2, MFNUMBER) As mossy  ' 2 contexts
Public mfsBackup(2, MFNUMBER) As mossy

Public MF(MFNUMBER) As Integer
Public MF2(MFNUMBER) As Mossy2

Public MFTypetoChange As Integer
Public MFNumbertoDegrade As Integer
Public CompetingStimulusSwap(3, 100) As Integer
Public CompetingStimulusNumber As Integer
Public CompetingStimulusTotal As Integer





Public Sub MF_TDV()
    THRDECAYMF = 1 - Exp(-Time_step_size / ThrtauMF)
    grAMPADecayMF = Exp(-Time_step_size / grAMPATau)
    grNMDADecayMF = Exp(-Time_step_size / grNMDATau)
End Sub


Public Sub DoMossyFibersFromSpikes(Trials As Integer)
Dim i As Integer
Dim BC As Integer
Dim Wbyte As Integer
Dim bins As Integer
Dim FromSpikes(624, 200) As Byte
Dim NumberofTrials As Integer

    For i = 1 To MFsAddedTotal
        Close #20
        Open MFFilenames(i) For Binary As #20
        Get #20, , NumberofTrials
        Get #20, , FromSpikes
        Close #20
        BC = 0
        
        For bins = 1 To 5000
            Wbyte = bins Mod 8
            Select Case Wbyte
                Case 1
                    If FromSpikes(BC, MFsAddedSequence(Trials, i)) > 127 Then
                        MFsAddedSpikes(i, bins) = 1
                        FromSpikes(BC, MFsAddedSequence(Trials, i)) = FromSpikes(BC, MFsAddedSequence(Trials, i)) - 128
                    Else
                        MFsAddedSpikes(i, bins) = 0
                    End If
                Case 2
                    If FromSpikes(BC, MFsAddedSequence(Trials, i)) > 63 Then
                        MFsAddedSpikes(i, bins) = 1
                        FromSpikes(BC, MFsAddedSequence(Trials, i)) = FromSpikes(BC, MFsAddedSequence(Trials, i)) - 64
                    Else
                        MFsAddedSpikes(i, bins) = 0
                    End If
                Case 3
                    If FromSpikes(BC, MFsAddedSequence(Trials, i)) > 31 Then
                        MFsAddedSpikes(i, bins) = 1
                        FromSpikes(BC, MFsAddedSequence(Trials, i)) = FromSpikes(BC, MFsAddedSequence(Trials, i)) - 32
                    Else
                        MFsAddedSpikes(i, bins) = 0
                    End If
                Case 4
                    If FromSpikes(BC, MFsAddedSequence(Trials, i)) > 15 Then
                        MFsAddedSpikes(i, bins) = 1
                        FromSpikes(BC, MFsAddedSequence(Trials, i)) = FromSpikes(BC, MFsAddedSequence(Trials, i)) - 16
                    Else
                        MFsAddedSpikes(i, bins) = 0
                    End If
                Case 5
                    If FromSpikes(BC, MFsAddedSequence(Trials, i)) > 7 Then
                       MFsAddedSpikes(i, bins) = 1
                        FromSpikes(BC, MFsAddedSequence(Trials, i)) = FromSpikes(BC, MFsAddedSequence(Trials, i)) - 8
                    Else
                        MFsAddedSpikes(i, bins) = 0
                    End If
                Case 6
                    If FromSpikes(BC, MFsAddedSequence(Trials, i)) > 3 Then
                        MFsAddedSpikes(i, bins) = 1
                        FromSpikes(BC, MFsAddedSequence(Trials, i)) = FromSpikes(BC, MFsAddedSequence(Trials, i)) - 4
                    Else
                        MFsAddedSpikes(i, bins) = 0
                    End If
                Case 7
                    If FromSpikes(BC, MFsAddedSequence(Trials, i)) > 1 Then
                        MFsAddedSpikes(i, bins) = 1
                        FromSpikes(BC, MFsAddedSequence(Trials, i)) = FromSpikes(BC, MFsAddedSequence(Trials, i)) - 2
                    Else
                        MFsAddedSpikes(i, bins) = 0
                    End If
                Case 0
                    If FromSpikes(BC, MFsAddedSequence(Trials, i)) > 0 Then
                        MFsAddedSpikes(i, bins) = 1
                        FromSpikes(BC, MFsAddedSequence(Trials, i)) = FromSpikes(BC, MFsAddedSequence(Trials, i)) - 1
                    Else
                        MFsAddedSpikes(i, bins) = 0
                    End If
                    BC = BC + 1
            End Select
            
        Next bins
        'Debug.Print i, MFsAdded, MFsAddedTotal, MFFilenames(i), MFsAddedSequence(Trials, i)_this_time
    Next i


End Sub
