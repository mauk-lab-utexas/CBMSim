Attribute VB_Name = "Simpson"
Option Explicit

Public SimpsonData(12000, 1000) As Single
Public SimpsonCounter(12000) As Integer
Public DoSimpson As Integer  ' remember to set DoSimpson to 0 at load time
Public SimpsonBinCount As Long
Public SimpsonCellType As Integer
Public SortedISI(1000) As Single

Public Sub StartSimpson(CellType As Integer)
    Erase SimpsonData
    Erase SimpsonCounter
    DoSimpson = 1
    SimpsonBinCount = 0
    cbm_main.SimpsonMainMenu.Enabled = False
    SimpsonCellType = CellType
End Sub

Public Sub SimpsonCalc()
Dim i As Integer
Dim j As Integer
Dim ISI(1000) As Single

Dim LogISI(1000) As Double
Dim CV2 As Double
Dim CVLog As Double
Dim X As Double
Dim X2 As Double
Dim LogISI_SD As Double
Dim LoopEnd As Long
Dim Spike As Integer

Dim Freq As Double
Dim MidValue As Integer
Dim MedianISI As Single
Dim ISI5P As Single

    SimpsonBinCount = SimpsonBinCount + 1
    Select Case SimpsonCellType
        Case 0
            LoopEnd = 12000
        Case 1
            LoopEnd = 900
        Case 2
            LoopEnd = 240
        Case 3
            LoopEnd = 96
    End Select
    
    If SimpsonBinCount < 60000 Then  ' stop after a minute
        For i = 1 To LoopEnd
            Spike = 0
            Select Case SimpsonCellType
                Case 0
                    If Gr(i).act = 1 Then Spike = 1
                Case 1
                    If Gol(i).act = 1 Then Spike = 1
                Case 2
                    If Bk(i).act = 1 Then Spike = 1
                Case 3
                     If BCells(i).act = 1 Then Spike = 1
            End Select
            If Spike = 1 Then
                If SimpsonCounter(i) < 1000 Then
                    SimpsonCounter(i) = SimpsonCounter(i) + 1
                    SimpsonData(i, SimpsonCounter(i)) = TotalBins
                End If
            End If
        Next i
    Else
        DoSimpson = 0
    End If
    
    If DoSimpson = 0 Then
        cbm_main.Text2.LinkTopic = "Excel|Sheet1"
        cbm_main.Text2.LinkMode = vbLinkManual
        
        For i = 1 To LoopEnd
            If SimpsonCounter(i) > 10 Then
                Erase ISI
                Freq = 0
                CV2 = 0
                Erase LogISI
                CVLog = 0
                X = 0
                X2 = 0
                
                For j = 2 To SimpsonCounter(i)
                    ISI(j - 1) = SimpsonData(i, j) - SimpsonData(i, j - 1)
                    SortedISI(j - 1) = ISI(j - 1)
                    Freq = Freq + ISI(j - 1)
                    LogISI(j - 1) = Log(ISI(j - 1))
                    CVLog = CVLog + LogISI(j - 1)
                    X2 = X2 + (LogISI(j - 1) * LogISI(j - 1))
                    X = X + LogISI(j - 1)
                    If j > 2 Then
                        CV2 = CV2 + (2 * Abs(ISI(j - 1) - ISI(j - 2))) / (ISI(j - 1) + ISI(j - 2))
                    End If
                Next j
                '  sort ISIs to find median and 5th percentile
                QSort 1, SimpsonCounter(i) - 1
                MidValue = Int((SimpsonCounter(i) - 1) / 2)
                MedianISI = SortedISI(MidValue) / 1000#
                
                MidValue = Int((SimpsonCounter(i) - 1) * 0.05)
                ISI5P = SortedISI(MidValue) / 1000#
                
                CV2 = CV2 / (1# * (SimpsonCounter(i) - 2))
                CVLog = CVLog / (1# * (SimpsonCounter(i) - 2))
               
                Freq = Freq / (1 * SimpsonCounter(i))
                Freq = 1000# / Freq
                
                ' calculate standard deviation of log ISIs
                LogISI_SD = (X * X) / (SimpsonCounter(i) * 1#)
                LogISI_SD = X2 - LogISI_SD
                LogISI_SD = LogISI_SD / (SimpsonCounter(i) - 1)
                
                CVLog = LogISI_SD / CVLog
                
                cbm_main.Text2.Text = Freq
                cbm_main.Text2.LinkItem = "R" & i & "C1"
                cbm_main.Text2.LinkPoke
                cbm_main.Text2.Text = MedianISI
                cbm_main.Text2.LinkItem = "R" & i & "C2"
                cbm_main.Text2.LinkPoke

                cbm_main.Text2.Text = CVLog
                cbm_main.Text2.LinkItem = "R" & i & "C3"
                cbm_main.Text2.LinkPoke

                cbm_main.Text2.Text = CV2
                cbm_main.Text2.LinkItem = "R" & i & "C4"
                cbm_main.Text2.LinkPoke
                cbm_main.Text2.Text = ISI5P
                cbm_main.Text2.LinkItem = "R" & i & "C5"
                cbm_main.Text2.LinkPoke
            End If
            cbm_main.Text2.Text = SimpsonCounter(i)
            cbm_main.Text2.LinkItem = "R" & i & "C6"
            cbm_main.Text2.LinkPoke
            DoEvents
        Next i
        cbm_main.SimpsonMainMenu.Enabled = True
    End If
End Sub
Public Sub QSort(ByVal First As Long, ByVal Last As Long)
    Dim Low As Long, High As Long
    Dim MidValue As Single
    Dim v As Single
    
    Low = First
    High = Last
    MidValue = SortedISI((First + Last - 1) \ 2)
    
    Do
        While SortedISI(Low) < MidValue
            Low = Low + 1
        Wend
        
        While SortedISI(High) > MidValue
            High = High - 1
        Wend
        
        If Low <= High Then
            v = SortedISI(Low)           'swap values
            SortedISI(Low) = SortedISI(High)        '
            SortedISI(High) = v          '
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    
    If First < High Then QSort First, High
    If Low < Last Then QSort Low, Last
End Sub
