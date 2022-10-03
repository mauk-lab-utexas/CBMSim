VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form plasticity 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Plasticity monitor"
   ClientHeight    =   6900
   ClientLeft      =   4440
   ClientTop       =   9315
   ClientWidth     =   23925
   LinkTopic       =   "Form1"
   ScaleHeight     =   -34.21
   ScaleLeft       =   -0.01
   ScaleMode       =   0  'User
   ScaleTop        =   20.5
   ScaleWidth      =   1.149
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.wts|*.wts"
   End
   Begin VB.Menu WeightsFileMainMenu 
      Caption         =   "File"
      Begin VB.Menu WeightsMenu 
         Caption         =   "Open Previously saved weights to buffer"
         Index           =   1
      End
      Begin VB.Menu WeightsMenu 
         Caption         =   "Save Weights"
         Index           =   2
      End
   End
   Begin VB.Menu RealTimeMenu 
      Caption         =   "Real Time"
   End
   Begin VB.Menu RefreshMainMenu 
      Caption         =   "Refresh"
      Begin VB.Menu RefreshMenu 
         Caption         =   "In sequence"
         Index           =   1
      End
      Begin VB.Menu RefreshMenu 
         Caption         =   "By activity"
         Index           =   2
      End
      Begin VB.Menu RefreshMenu 
         Caption         =   "By CS activity"
         Index           =   3
      End
      Begin VB.Menu RefreshMenu 
         Caption         =   "Difference in sequence"
         Index           =   4
      End
      Begin VB.Menu RefreshMenu 
         Caption         =   "Difference by activity"
         Index           =   5
      End
      Begin VB.Menu RefreshMenu 
         Caption         =   "Difference by CS activity"
         Index           =   6
      End
      Begin VB.Menu RefreshMenu 
         Caption         =   "PC to Basket in sequence"
         Index           =   7
      End
      Begin VB.Menu RefreshMenu 
         Caption         =   "PC to basket by activity"
         Index           =   8
      End
   End
   Begin VB.Menu SendtoExcelMenu 
      Caption         =   "SendWeightstoExcel"
   End
End
Attribute VB_Name = "plasticity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

End Sub

Private Sub Check1_Click()

End Sub

Private Sub Form_Load()
    plasticity.ScaleTop = 1.1
    plasticity.ScaleHeight = -1.2
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'    If Check1.Visible = True Then Check1.Visible = False Else Check1.Visible = True
'    If Command1(0).Visible = True Then Command1(0).Visible = False Else Command1(0).Visible = True
'    If Command1(1).Visible = True Then Command1(1).Visible = False Else Command1(1).Visible = True
End Sub

Private Sub RealTimeMenu_Click()
    If DoRealTime = 1 Then DoRealTime = 0 Else DoRealTime = 1
End Sub

Private Sub RefreshMenu_Click(Index As Integer)
Dim X As Integer
Dim i As Integer
Dim min As Long
Dim minX As Integer
Dim temp(SYNUMBER) As Long
Dim Temp2(SYNUMBER) As Single

Dim c As ColorConstants
Dim count As Integer
Dim BCount As Integer

    plasticity.Visible = True
    plasticity.ScaleWidth = SYNUMBER
  
    plasticity.Cls
    plasticity.DrawWidth = 3
    
    If Index < 4 Or Index > 6 Then
        plasticity.ScaleTop = 1.1
        plasticity.ScaleHeight = -1.2
    Else
        plasticity.ScaleTop = 1.1
        plasticity.ScaleHeight = -2.2
    End If
    
    Select Case Index
        Case 1
            For X = 1 To SYNUMBER
                If X < 501 Then
                    c = vbRed
                ElseIf X < 1001 Then
                    c = vbBlue
                ElseIf X < 1501 Then
                    c = vbGreen
                ElseIf X < 2001 Then
                    c = vbMagenta
                ElseIf X < 2501 Then
                    c = vbCyan
                ElseIf X < 3001 Then
                    c = vbRed
                ElseIf X > 3501 Then
                    c = vbBlack
                ElseIf X < 4001 Then
                    c = vbBlue
                ElseIf X < 4501 Then
                    c = vbGreen
                ElseIf X < 5001 Then
                    c = vbMagenta
                ElseIf X < 5501 Then
                    c = vbCyan
                Else
                    c = vbBlack
                End If
                
                plasticity.PSet (X, grWeight(X)), c
            Next X
        Case 2
            For X = 1 To SYNUMBER
                temp(X) = GranAct(X)
            Next X
            For X = 1 To SYNUMBER
                min = 6000000
                For i = 1 To SYNUMBER
                    If temp(i) < min Then
                        min = temp(i)
                        minX = i
                    End If
                Next i
                plasticity.PSet (X, grWeight(minX)), vbBlack
                temp(minX) = 6000000
            Next X
        Case 3
            For X = 1 To SYNUMBER
                temp(X) = GranActCS(X)
            Next X
            For X = 1 To SYNUMBER
                min = 6000000
                For i = 1 To SYNUMBER
                    If temp(i) < min Then
                        min = temp(i)
                        minX = i
                    End If
                Next i
                plasticity.PSet (X, grWeight(minX)), vbBlack
                temp(minX) = 6000000
            Next X
        
        Case 4
            For X = 1 To SYNUMBER
                plasticity.PSet (X, PreviousgrWeight(X) - grWeight(X)), vbBlack
            Next X
        Case 5
            For X = 1 To SYNUMBER
                temp(X) = GranAct(X)
            Next X
            For X = 1 To SYNUMBER
                min = 6000000
                For i = 1 To SYNUMBER
                    If temp(i) < min Then
                        min = temp(i)
                        minX = i
                    End If
                Next i
                plasticity.PSet (X, PreviousgrWeight(minX) - grWeight(minX)), vbBlack
                temp(minX) = 6000000
            Next X
        Case 6
            For X = 1 To SYNUMBER
                temp(X) = GranActCS(X)
            Next X
            For X = 1 To SYNUMBER
                min = 6000000
                For i = 1 To SYNUMBER
                    If temp(i) < min Then
                        min = temp(i)
                        minX = i
                    End If
                Next i
                plasticity.PSet (X, PreviousgrWeight(minX) - grWeight(minX)), vbBlack
                temp(minX) = 6000000
            Next X
        Case 7  'pf to basket in sequence
            count = 0
            BCount = 1
            For X = 1 To SYNUMBER
                count = count + 1
                If count > PFBasketsynUMBER Then
                    count = 1
                    BCount = BCount + 1
                End If
                plasticity.PSet (X, BCells(BCount).grW(count)), vbBlack
            Next X
        Case 8   'pf to basket by activity
            BCount = 1
            count = 0
            For X = 1 To SYNUMBER
                temp(X) = GranAct(X)
                count = count + 1
                If count > PFBasketsynUMBER Then
                    count = 1
                    BCount = BCount + 1
                End If
                Temp2(X) = BCells(BCount).grW(count)
            Next X
            For X = 1 To SYNUMBER
                min = 6000000
                For i = 1 To SYNUMBER
                    If temp(i) < min Then
                        min = temp(i)
                        minX = i
                    End If
                Next i
                plasticity.PSet (X, Temp2(minX)), vbBlack
                temp(minX) = 6000000
            Next X
        
    End Select
    
'    For X = 1 To 600
'        plasticity.PSet (X * 10, mfweight(X) * 10), RGB(255, 255, 255)
'    Next X
    
End Sub

Private Sub SendtoExcelMenu_Click()
Dim X As Integer
Dim i As Integer
Dim min As Long
Dim minX As Integer
Dim n As Integer
Dim p As Integer

Dim temp(SYNUMBER) As Long
    CommonDialog1.filename = ""
    CommonDialog1.ShowSave
    If CommonDialog1.filename <> "" Then
        Close #13
        Open CommonDialog1.filename For Output As #13
        For X = 1 To SYNUMBER
            temp(X) = GranAct(X)
        Next X
        For X = 1 To SYNUMBER
            min = 6000000
            For i = 1 To SYNUMBER
                If temp(i) < min Then
                    min = temp(i)
                    minX = i
                End If
            Next i
            Write #13, temp(minX), grWeight(minX), PreviousgrWeight(minX), grWeight(X), PreviousgrWeight(X)
            temp(minX) = 6000000
        Next X
        For X = 1 To SYNUMBER
                temp(X) = GranActCS(X)
            Next X
            For X = 1 To SYNUMBER
                min = 6000000
                For i = 1 To SYNUMBER
                    If temp(i) < min Then
                        min = temp(i)
                        minX = i
                    End If
                Next i
                Write #13, PreviousgrWeight(minX), grWeight(minX)
                temp(minX) = 6000000
            Next X
            
        For n = 1 To 6
            For p = 1 To 10
                Write #13, gPURKtoNUCLEUS(Nc(n).PCsyn(p), n)
            Next p
        Next n
        For n = 1 To NCNUMBER
            For p = 1 To NumCF
                Write #13, Nc(n).gNUCtoCF(p)
            Next p
        Next n
        Close #13
    End If
End Sub

Private Sub WeightsandDiffstoExcelmenu_Click()
    
End Sub

Private Sub WeightsMenu_Click(Index As Integer)
    CommonDialog1.filename = ""
    If Index = 1 Then
        CommonDialog1.ShowOpen
        If Not CommonDialog1.filename = "" Then
            Close #20
            Open CommonDialog1.filename For Binary As #20
            Get #20, , PreviousgrWeight
            Close #20
        End If
    Else
        CommonDialog1.ShowSave
        If Not CommonDialog1.filename = "" Then
            Close #20
            Open CommonDialog1.filename For Binary As #20
            Put #20, , grWeight
            Close #20
        End If
    End If
End Sub
