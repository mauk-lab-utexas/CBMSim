VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form raster_form 
   BackColor       =   &H00000000&
   Caption         =   "Rasters: granule 1:1000"
   ClientHeight    =   12585
   ClientLeft      =   1005
   ClientTop       =   1395
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   12585
   ScaleWidth      =   12600
   Begin MSComDlg.CommonDialog Raster_CDB 
      Left            =   240
      Top             =   9600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.rl | *.rl"
   End
   Begin VB.Menu raster_mode_menu 
      Caption         =   "Mode"
      Begin VB.Menu rm_menu 
         Caption         =   "Activity this time bin"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu rm_menu 
         Caption         =   "Cumulative activity grey scale coded"
         Index           =   1
      End
      Begin VB.Menu rm_menu 
         Caption         =   "Current plasticity color coded"
         Index           =   2
      End
      Begin VB.Menu rm_menu 
         Caption         =   "Current activity coded by synapse strength"
         Index           =   3
      End
      Begin VB.Menu rm_menu 
         Caption         =   "Cumulative activity coded by synapse strength"
         Index           =   4
      End
      Begin VB.Menu rm_menu 
         Caption         =   "Cumulative during CS only"
         Index           =   5
      End
   End
   Begin VB.Menu cell_type_menu 
      Caption         =   "Cell type"
      Begin VB.Menu cell_menu2 
         Caption         =   "Mossy fibers"
         Index           =   0
      End
      Begin VB.Menu cell_menu2 
         Caption         =   "granule cells"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu cell_menu2 
         Caption         =   "Golgi cells"
         Index           =   2
      End
      Begin VB.Menu cell_menu2 
         Caption         =   "Stellate and Basket cells"
         Index           =   3
      End
      Begin VB.Menu cell_menu2 
         Caption         =   "STP Eligibility"
         Index           =   4
      End
      Begin VB.Menu cell_menu2 
         Caption         =   "STP color coding"
         Index           =   5
      End
      Begin VB.Menu cell_menu2 
         Caption         =   "Inputs Mode"
         Index           =   6
      End
      Begin VB.Menu cell_menu2 
         Caption         =   "Imported Mossy Fibers Red"
         Index           =   7
      End
      Begin VB.Menu cell_menu2 
         Caption         =   "Mossy fibers color coded"
         Index           =   8
      End
   End
   Begin VB.Menu build_cell_menu 
      Caption         =   "Build"
      Begin VB.Menu Sort_by_CS_Menu 
         Caption         =   "Sort by CS input"
      End
      Begin VB.Menu build_menu 
         Caption         =   "Consecquitive with different starting cell"
         Index           =   0
         Begin VB.Menu starting_cell_menu 
            Caption         =   "Reset to start at 1"
            Index           =   0
         End
         Begin VB.Menu starting_cell_menu 
            Caption         =   "Increase 1000"
            Index           =   1
         End
         Begin VB.Menu starting_cell_menu 
            Caption         =   "Decrease 1000"
            Index           =   2
         End
      End
      Begin VB.Menu build_menu 
         Caption         =   "Sort by latency to peak"
         Index           =   1
      End
      Begin VB.Menu build_menu 
         Caption         =   "Enrich for CS active cells"
         Index           =   2
      End
      Begin VB.Menu build_menu 
         Caption         =   "Save raster list to file"
         Index           =   3
      End
      Begin VB.Menu build_menu 
         Caption         =   "Open raster list file"
         Index           =   4
      End
   End
   Begin VB.Menu adjust_menu 
      Caption         =   "Adjust"
   End
   Begin VB.Menu ResetHistogramMenu 
      Caption         =   "Reset Histograms"
      Begin VB.Menu ResetHistoMenu 
         Caption         =   "Reset Histograms Now"
         Index           =   1
      End
      Begin VB.Menu ResetHistoMenu 
         Caption         =   "Reset at End of Sessions"
         Index           =   2
      End
   End
   Begin VB.Menu NM_menu 
      Caption         =   "Neuralynx Mode"
      Enabled         =   0   'False
   End
   Begin VB.Menu ShowATCMenu 
      Caption         =   "Show ATC"
      Enabled         =   0   'False
   End
   Begin VB.Menu ShowActivityMenu 
      Caption         =   "Show Total Activity"
   End
End
Attribute VB_Name = "raster_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub adjust_menu_Click()
    Raster_adjust.Visible = True
End Sub

Private Sub build_menu_Click(Index As Integer)
Dim i As Integer
Dim j As Integer
Dim sc As Single
Dim sp As Single
Dim target_cell As Integer
    If Index = 1 Then
        
    ElseIf Index = 2 Then
        i = 1
        target_cell = 1
        While i < 1001
            sc = 0
            sp = 0
            For j = cs_onset(1) / 5 To (cs_onset(1) + US_onset(1)) / 5
                sc = sc + GR_histo(target_cell, j) / Trials_this_time
                If GR_histo(target_cell, j) / Trials_this_time > sp Then
                    sp = GR_histo(target_cell, j) / Trials_this_time
                End If
            Next j
            sc = sc / (US_onset(1) / 5)
            If sp > 0.2 Then
                Raster_list(i) = target_cell
                i = i + 1
                target_cell = target_cell + 1
            Else
                target_cell = target_cell + 1
            End If
            If target_cell = 10001 Then i = 1001
        Wend
    ElseIf Index = 3 Then
        Raster_CDB.filename = ""
        Raster_CDB.ShowSave
        If Raster_CDB.filename <> "" Then
            Close #9
            Open Raster_CDB.filename For Binary As #9
            Put #9, , Raster_list
            Close #9
        End If
    ElseIf Index = 4 Then
        Raster_CDB.filename = ""
        Raster_CDB.ShowOpen
        If Raster_CDB.filename <> "" Then
            Close #9
            Open Raster_CDB.filename For Binary As #9
            Get #9, , Raster_list
            Close #9
        End If
    End If
End Sub

Private Sub cell_menu_Click()
End Sub

Private Sub cell_menu2_Click(Index As Integer)
Dim i As Integer
Dim X As String

    For i = 0 To 8
        cell_menu2(i).Checked = False
    Next i
    cell_menu2(Index).Checked = True
    raster_Cell_type = Index
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        DoGoAudio = 0
        DoGrAudio = 0
        DoMFAudio = 0
        DoBSAudio = 0
        
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim c As String
Dim i As Integer
Dim cell_num As Integer
Dim j As Integer
Dim y_count As Integer

    If Button = 1 Then
        Debug.Print Int(y)
        If cell_menu2(1).Checked = True Then  'granule cell raster
            DoGrAudio = Int(y)
            If DoGrAudio > 1000 Then DoGrAudio = 1000
            If DoGrAudio < 1 Then DoGrAudio = 1
        ElseIf cell_menu2(2).Checked = True Then 'Golgi cell raster
            DoGoAudio = Int(y)
            If DoGoAudio > 900 Then DoGoAudio = 900
            If DoGoAudio < 1 Then DoGoAudio = 1
        ElseIf cell_menu2(0).Checked = True Or cell_menu2(6).Checked = True Then
            DoMFAudio = Int(y)
            If DoMFAudio > 599 Then DoMFAudio = 599
            If DoMFAudio < 1 Then DoMFAudio = 1
        ElseIf cell_menu2(3).Checked = True Then   '  Stellate or Basket raster
            DoBSAudio = Int(y)
            If DoBSAudio < 1 Then DoBSAudio = 1
            If DoBSAudio < 248 Then
                If DoBSAudio > 240 Then DoBSAudio = 240
            Else
                DoBSAudio = DoBSAudio - 249
                If DoBSAudio < 1 Then DoBSAudio = 1
                If DoBSAudio > 96 Then DoBSAudio = 96
                DoBSAudio = DoBSAudio + 1000
            End If
        End If
        
    ElseIf Button = 2 Then
        OScope.CellNumber.Text = Int(y)
        GWin.Left = 0
        GWin.Top = 0
        GWin.Width = SYNUMBER
        GWin.Height = 18000
        
        GWin.ScaleLeft = -50
        GWin.ScaleWidth = 1100
        GWin.BackColor = vbWhite
        GWin.Cls
        GWin.Line (200, GWin.Height)-(200 + cs_duration(1) / 5, 0), RGB(240, 240, 240), BF
        For i = 0 To 2
            GWin.Command1(i).Visible = False
        Next i
        cell_num = Raster_list(Int(y))
        GeneologyCell = cell_num
        
        GWin.Command1(1).Visible = False
        GeneologyCellType = 0
        If cell_menu2(1).Checked = True Then  'granule cell raster
            For i = 0 To 2
                GWin.Command1(i).Visible = True
            Next i
            c = "granule cell number" + CStr(cell_num)
            GWin.Caption = c
            GWin.Visible = True
            y_count = 12
            GWin.ScaleTop = 15
            GWin.ScaleHeight = -15
            For i = 1 To 1000
                GWin.Line (i, y_count)-(i, y_count + GR_histo(cell_num, i) / HistoDivisor), vbBlue
            Next i
            GWin.Label1.Caption = grWeight(cell_num)
            If cell_num < GrX * GrY Then
                For j = 1 To Gr(cell_num).numdend
                    y_count = y_count - 1
                    For i = 1 To 1000
                        GWin.Line (i, y_count)-(i, y_count + MF_histo(Gr(cell_num).MF(j), i) / HistoDivisor), vbBlack
                        GWin.Line (i, y_count - 7)-(i, (y_count - 7) + Go_histo(Gr(cell_num).Gol(j), i) / HistoDivisor), vbRed
                    Next i
                Next j
                
                GWin.Command1(1).Visible = True
                GeneologyCellType = 1
            End If
        ElseIf cell_menu2(0).Checked = True Then 'Mossy Fiber raster
            For i = 0 To 2
                GWin.Command1(i).Visible = True
            Next i
            c = "Mossy Fiber number" + CStr(cell_num)
            GWin.Caption = c
            GWin.Visible = True
            y_count = 5
            GWin.ScaleTop = 10
            GWin.ScaleHeight = -10
            For i = 1 To 1000
                GWin.Line (i, y_count)-(i, y_count + MF_histo(cell_num, i) / HistoDivisor), vbBlue
            Next i
            
        ElseIf cell_menu2(2).Checked = True Then 'Golgi cell raster
            c = "Golgi cell number" + CStr(cell_num)
            GWin.Caption = c
            GWin.Visible = True
            y_count = 105
            GWin.ScaleTop = 107
            GWin.ScaleHeight = -107
            cell_num = Int(y)
            If cell_num < GoX * GoY Then
                For i = 1 To 1000
                    GWin.Line (i, y_count)-(i, y_count + Go_histo(cell_num, i) / HistoDivisor), vbBlue
                Next i
                
                For j = 1 To numGoGlDend
                    y_count = y_count - 1
                    For i = 1 To 1000
                        GWin.Line (i, y_count)-(i, y_count + (2 * MF_histo(Gol(cell_num).MF(j), i)) / HistoDivisor), vbRed
                    Next i
                Next j
                For j = 1 To numGoGrDend
                    y_count = y_count - 1
                    For i = 1 To 1000
                        GWin.Line (i, y_count)-(i, y_count + GR_histo(Gol(cell_num).preGr(j), i) / HistoDivisor), vbBlack
                    Next i
                Next j
                
                GWin.Command1(1).Visible = True
                GeneologyCellType = 2
            End If
            
        ElseIf cell_menu2(3).Checked = True Then 'Stellate or Basket cell raster
            cell_num = Int(y)
            
            If cell_num < 1 Then
                cell_num = 1
            End If
            If cell_num < 248 Then
                If cell_num > 240 Then
                    cell_num = 240
                End If
                GeneologyCellType = 3  'stellate
                c = "Stellate cell number" + CStr(cell_num)
                GWin.Caption = c
                GWin.Visible = True
            Else
                cell_num = cell_num - 249
                If cell_num < 1 Then cell_num = 1
                If cell_num > 96 Then cell_num = 96
                GeneologyCellType = 4 'basket
                c = "Basket cell number" + CStr(cell_num)
                GWin.Caption = c
                GWin.Visible = True
            End If
            GWin.ScaleTop = 2
            GWin.ScaleHeight = -8
            If GeneologyCellType = 3 Then
                Debug.Print cell_num
                For i = 1 To 1000
                    GWin.Line (i, 0)-(i, Stellate_histo(cell_num, i) / HistoDivisor), vbBlue
                Next i
            ElseIf GeneologyCellType = 4 Then
                Debug.Print cell_num, cell_num
                For i = 1 To 1000
                    GWin.Line (i, 0)-(i, Basket_histo(cell_num, i) / HistoDivisor), vbBlue
                Next i
            End If
        End If
        
    End If
End Sub

Private Sub Form_Resize()
    If raster_form.WindowState <> 1 Then
        raster_form.ScaleWidth = 5000
        raster_form.ScaleHeight = 1000
    End If
End Sub

Private Sub ResetHistoMenu_Click(Index As Integer)
Dim i As Integer
Dim j As Integer

    Select Case Index
        Case 1
            HistoDivisor = 1
            Erase GR_histo
            Erase Go_histo
            Erase Stellate_histo
            Erase Basket_histo
        Case 2
            If ResetHistoMenu(2).Checked = True Then ResetHistoMenu(2).Checked = False Else ResetHistoMenu(2).Checked = True
    End Select
End Sub

Private Sub rm_menu_Click(Index As Integer)
Dim i As Integer
    For i = 0 To 5
        rm_menu(i).Checked = False
    Next i
    rm_menu(Index).Checked = True
    raster_mode = Index
End Sub

Private Sub Sort_by_CS_Menu_Click()
Dim i As Integer
Dim j As Integer
Dim t As Integer
Dim p As Integer
Dim d(SYNUMBER, 2) As Integer
Dim temp(1000, 2) As Integer
    If (cell_menu2(1).Checked = True) Or (cell_menu2(5).Checked = True) Or (cell_menu2(4).Checked = True) Then    'granule cells
        t = 0
        p = 0
        For i = 1 To SYNUMBER
            t = 0
            For j = 1 To Gr(i).numdend
                If MFS(1, Gr(i).MF(j)).CStype = 1 Or MFS(1, Gr(i).MF(j)).CStype = 2 Then 'phasic
                    p = p + 1
                    If p <= 1000 Then Raster_list(p) = i
                    j = 1000
                End If
            Next j
        Next i
    End If
    
End Sub

Private Sub starting_cell_menu_Click(Index As Integer)
Dim i As Integer
Dim s As String
    If cell_menu2(1).Checked = True Then
        If Index = 0 Then
            Raster_list(1) = 1
        ElseIf Index = 1 Then
            If Raster_list(1) < 10001 Then Raster_list(1) = Raster_list(1) + 1000 Else Raster_list(1) = 11001
            
        Else
            If Raster_list(1) > 1000 Then Raster_list(1) = Raster_list(1) - 1000 Else Raster_list(1) = 1
        End If
        
        For i = 2 To 1000
            Raster_list(i) = Raster_list(i - 1) + 1
        Next i
        s = "Rasters: granule " + Str$(Raster_list(1)) + ":" + Str$(Raster_list(1000))
        raster_form.Caption = s
    End If
End Sub
