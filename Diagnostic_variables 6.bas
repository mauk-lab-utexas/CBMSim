Attribute VB_Name = "Diagnostic_variables"
Option Explicit

Public Debug_gran As Single
Public Debug_Gol As Single
Public Debug_BS As Single
Public Debug_MF As Single
Public TIME_TOTAL As Double
Public TIME_START As Double
Public TIME_MF As Double
Public TIME_GR As Double
Public TIME_GOL As Double
Public TIME_BK As Double
Public TIME_PK As Double
Public TIME_NUC As Double
Public TIME_Plasticity As Double
Public NeuralynxRecordingCells(6) As Long
Public NeuralynxSpikes(6, 2000) As Double
Public ShowRealTime As Integer

Public CFgDisplayMultiplier As Single

Public T1 As Double
Public T2 As Double
Public T3 As Double
Public T4 As Double
Public T5 As Double
Public T6 As Double
Public T6b As Double
Public T7 As Double

Public DoSpecialRecord As Integer
Public DoPCPackedRecord As Integer

Public PCSpikes(PCNUMBER, 120000) As Long
Public BCSpikes(BasketNUMBER, 120000) As Long
Public SCSpikes(STELLATENUMBER, 120000) As Long


Public GranAct(SYNUMBER) As Long
Public GranActCS(SYNUMBER) As Long
Public PurkinjeActivity(PCNUMBER, 1001) As Integer
Public grPurkinjeWeights(PCNUMBER, 1001) As Single
Public grPurkinjeActivity(PCNUMBER, 1001) As Single

Public PurkinjeNucleusWeights(NCNUMBER) As Single

Public NucleusActivity(NCNUMBER, 1001) As Integer
Public ClimbingFiberActivity(1001, 12) As Integer
Public CFactivityTimes(1000, 12, 16) As Integer
Public CFCounter(12) As Integer
Public GranuleActivity(1001) As Double
Public GolgiActivity(1001) As Single
Public MFActivity(1001) As Single
Public PNWeightsBYPURK(PCNUMBER, 1001) As Single
Public PNWeightsBYNUC(NCNUMBER, 1001) As Single
Public grPCWeights(1001) As Single


'Public MF_ISI(20000) As Long
'Public gr_ISI(20000) As Long
'Public Gol_ISI(20000) As Long


Public GeneologyCell As Long
Public GeneologyCellType As Integer

Public ConductanceFormMode As Integer
Public TotalSTP As Single
Public STPcount As Single




Public Type Timestamp       'Record to read in the 8-byte timestamp
    B0 As Byte              'Can't use integer type because the timestamp
    B1 As Byte              'is an unsigned value
    B2 As Byte
    B3 As Byte
    B4 As Byte
    B5 As Byte
    B6 As Byte
    B7 As Byte
End Type

Public Type tetdata            'Record for reading in Cheetah .Ntt file
    TS As Timestamp
    ScNumber As Long
    CellNumber As Long
    Param(7) As Long
    Chans(127) As Integer
End Type


Public Sub OpenDatFile()
Dim filename As String
Dim filename2 As String
Dim l As Integer

    If cbm_main.OutputCD.filename <> "" Then
        filename = cbm_main.OutputCD.filename
        l = Len(filename)
        filename = Mid$(filename, 1, l - 4)
        If SessionCounter < 10 Then
            filename = filename + "0" + Mid$(Str$(SessionCounter), 2, 1)
        ElseIf SessionCounter < 100 Then
            filename = filename + Mid$(Str$(SessionCounter), 2, 2)
        Else
            filename = filename + Mid$(Str$(SessionCounter), 2, 3)
        End If
        filename2 = filename + ".bvi"
        filename = filename + ".dat"
    Else
        filename = "CBM" + Str$(Year(Date)) + Str$(Month(Date)) + Str$(Day(Date)) + Str$(Hour(Time)) + Str$(Minute(Time))
        filename2 = filename + ".bvi"
        filename = filename + ".dat"
        cbm_main.OutputCD.filename = filename
    End If
    Close #24
    Open filename For Binary As 24
    Close #25
    Open filename2 For Output As #25
    
End Sub
Public Sub AutoSaveSimulation()
Dim filename As String
Dim filename2 As String
Dim l As Integer
    If cbm_main.OutputCD.filename <> "" Then
        
        filename = cbm_main.OutputCD.filename
        l = Len(filename)
        filename = Mid$(filename, 1, l - 4)
        If SessionsThisExp < 10 Then
            filename = filename + "0" + Mid$(Str$(SessionsThisExp), 2, 1)
        ElseIf SessionCounter < 100 Then
            filename = filename + Mid$(Str$(SessionsThisExp), 2, 2)
        Else
            filename = filename + Mid$(Str$(SessionsThisExp), 2, 3)
        End If
        
        
            filename = filename + ".cbm"
    Else
        filename = "CBM" + Str$(Year(Date)) + Str$(Month(Date)) + Str$(Day(Date)) + Str$(Hour(Time)) + Str$(Minute(Time))
        filename = filename + ".cbm"
    End If
    
    If CommandLine = 1 Then
        filename = Mid$(OutFileName, 1, Len(OutFileName) - 4) + ".cbm"
    End If
    
        Close #1
        Open filename For Binary As #1
        SaveSim
    If CommandLine = 1 And SaveRastersON = 1 Then
        filename = Mid$(OutFileName, 1, Len(OutFileName) - 4) + ".gr"
        Close #1
        Open filename For Binary As #1
        GR_histo(0, 0) = (HistoDivisor)
        Put #1, , GR_histo
        
        filename = Mid$(OutFileName, 1, Len(OutFileName) - 4) + ".gol"
        Close #1
        Open filename For Binary As #1
        Go_histo(0, 0) = (HistoDivisor)
        Put #1, , Go_histo
                    
        filename = Mid$(OutFileName, 1, Len(OutFileName) - 4) + ".mf"
        Close #1
        Open filename For Binary As #1
        MF_histo(0, 0) = (HistoDivisor)
        Put #1, , MF_histo
        Close #1
    End If
    
    If CommandLine = 1 And SaveWeightsON = 1 Then
        filename = Mid$(OutFileName, 1, Len(OutFileName) - 4) + ".wts"
        Close #1
        Open filename For Binary As #1
        
        Put #1, , grWeight
        Close #1
    End If
    End Sub
    
    
    
'  Neuralynx data formats to output .ntt mimicking files.

Public Function DoubletoTS(Number As Double) As Timestamp
Dim remainder As Double

    DoubletoTS.B7 = CByte(Number \ (2 ^ 56))
    Number = Number Mod (2 ^ 56)
    DoubletoTS.B6 = CByte(Number \ (2 ^ 48))
    Number = Number Mod (2 ^ 48)
    DoubletoTS.B5 = CByte(Number \ (2 ^ 40))
    Number = Number Mod (2 ^ 40)
    DoubletoTS.B4 = CByte(Number \ (2 ^ 32))
    Number = Number Mod (2 ^ 32)
    DoubletoTS.B3 = CByte(Number \ (2 ^ 24))
    Number = Number Mod (2 ^ 24)
    DoubletoTS.B2 = CByte(Number \ (2 ^ 16))
    Number = Number Mod (2 ^ 16)
    DoubletoTS.B1 = CByte(Number \ (2 ^ 8))
    Number = Number Mod (2 ^ 8)
    DoubletoTS.B0 = CByte(Number)
End Function

'Convert timestamp to a double type
Public Function TStoDouble(TS As Timestamp) As Double
    TStoDouble = (2 ^ 56 * TS.B7 + 2 ^ 48 * TS.B6 + 2 ^ 40 * TS.B5 + 2 ^ 32 * TS.B4 + _
    2 ^ 24 * TS.B3 + 2 ^ 16 * TS.B2 + 2 ^ 8 * TS.B1 + TS.B0)
End Function

