Attribute VB_Name = "Histograms"
Option Explicit
Public GR_histo(12001, 1001) As Integer
Public Go_histo(GOLnumber, 1002) As Integer
Public Purk_histo(PCNUMBER, 1002) As Integer
Public Nuc_histo(NCNUMBER, 1002) As Integer
Public CF_histo(1, 1002) As Integer
Public MF_histo(MFNUMBER, 1002) As Integer
Public Stellate_histo(STELLATENUMBER, 1002) As Integer
Public Basket_histo(BasketNUMBER, 1002) As Integer
Public response_histo(1002) As Integer
Public RN_Histo(5000) As Single

Public histo_switch As Integer

Public response_counter As Integer

Public jav_Mfinputaverage(1000) As Single

Public Rasters(36, 1000, 1000) As Boolean
Public PCPacked(5000, 108) As Long

Public RecordFormCSDuration As Integer
Public RecordFormCSDurationFromMenu As Integer
Public RecordFormBackColor As ColorConstants
Public RecordFormDotColor As ColorConstants
Public RecordFormHistoColor As ColorConstants
Public RecordFormCSColor As ColorConstants
Public RecordFormDotSize As Integer
Public RecordFormScaleRows As Integer
Public RecordFormHisto(1000) As Integer

'Public RRaster(1, 1000, 1000) As Boolean
Public RRCellType(36) As Integer
Public RRCellNum(36) As Integer
Public RRScale As Single

Public gr_PC_weights_rasters(SYNUMBER, 1000) As Single
Public mf_nuc_weights_rasters(600, 1000) As Single

Public Raster_list(1000) As Integer
Public Raster_plasticity(SYNUMBER) As Single
Public rasters_MF_plasticity(600) As Single

Public DoRealTime As Integer

Public Sub Do_rasters(Cell_type As Integer, start_cell As Integer, stop_cell As Integer, time_bin As Integer, clear As Integer, mode As Integer)
Dim i As Integer
Dim raster_color As Integer
Dim rc As Single
Dim rc2 As ColorConstants
Dim r As Integer
Dim b As Integer
Dim g As Integer
Dim w As Single
Dim r1 As Single
Dim g1 As Single
Dim B1 As Single
Dim X As Integer
Dim y As Integer

'mode 0 = draw current activity
'mode 1 = draw total histogram activity grey scale coded
'mode 2 = draw current activity color coded for plasticity at gr to Purk synapses
'mode 3 = draw current activity color coded strength of gr_Purk synapses
'mode 4 = draw total histogram activity coded by strength of gr_Purk synapses
'mode 5 = draw current activity for non-CS, for CS draw total activity gray scale coded
'mode 7 = draw only mossy fibers imported from spikes

    If clear = 1 Then raster_form.Cls
    If CF(1).act = 1 Then raster_form.Line (time_bin, 0)-(time_bin, 5), RGB(255, 255, 255)
    If time_bin = cs_onset(1) Then
        'raster_form.Line (time_bin, 0)-(time_bin + US_onset(1), raster_form.ScaleHeight), RGB(0, 0, 255), BF
    End If
    If Cell_type = 0 Or Cell_type = 6 Or Cell_type = 7 Then    'MOSSY FIBERS
        If mode = 0 Then
            raster_form.ForeColor = vbWhite
            For i = start_cell To stop_cell
                If MF(i) = 1 Then raster_form.PSet (time_bin, i)
            Next i
            raster_form.ForeColor = vbRed
            For i = 1 To MFsAddedTotal
                If MF(MFsAddedIdentity(i)) = 1 Then raster_form.PSet (time_bin, MFsAddedIdentity(i))
            Next i
        ElseIf mode = 1 Then
            For i = start_cell To stop_cell
                rc = (MF_histo(i, bincounter5) / HistoDivisor) * 255
                raster_color = Int(rc)
                raster_form.PSet (time_bin, i), RGB(raster_color, raster_color, raster_color)
            Next i
        ElseIf mode = 2 Then
            For i = start_cell To stop_cell
                If MF(i) = 1 Then
                    If rasters_MF_plasticity(i) > 0 Then
                        rc2 = vbGreen
                    ElseIf rasters_MF_plasticity(i) < 0 Then
                        rc2 = vbRed
                    Else
                        rc2 = vbWhite
                    End If
                    raster_form.PSet (time_bin, i), rc2
                End If
            Next i
        ElseIf mode = 3 Then
            
            For i = start_cell To stop_cell
                If MF(i) = 1 Then
                    w = (mfweight(i) / MF_weights_denom) * 511
                   
                    If w > 511 Then w = 511
                    If w > 255 Then
                        g = 255
                        r = 512 - w
                        b = r
                    Else
                        g = w
                        r = 255
                        b = g
                    End If
                    raster_form.PSet (time_bin, i), RGB(r, g, b)
                End If
            Next i
            
        ElseIf mode = 4 Then
            For i = start_cell To stop_cell
                w = mfweight(i) / MF_weights_denom
                w = w * 511
                If w > 511 Then w = 511
                If w > 255 Then
                    g = 255
                    r = 512 - ((w))
                    b = r
                Else
                    g = w
                    r = 255
                    b = g
                End If
                rc = (MF_histo(i, bincounter5) / HistoDivisor)
                r = r * rc
                g = g * rc
                b = b * rc
                raster_form.PSet (time_bin, i), RGB(r, g, b)
            Next i
            
        End If
        
        If Cell_type = 6 Then
            For X = 1 To NumCF
                y = NumCF - X + 1
                raster_form.PSet (time_bin, -200 + ((32) * 25) - ((CF(y).v * 0.7) + ELEAKNC) - 70 + ((4 - X) * 25)), RGB(255, 0, 0)
                If CF(y).act = 1 Then raster_form.Line -(time_bin, -200 + ((32) * 25) - ((CF(y).v * 0.7) + ELEAKNC) - 85 + ((4 - X) * 25)), RGB(255, 0, 0)
            Next X
        End If
    ElseIf Cell_type = 8 Then    ' Color code mossy fiber by CS identity
        
        
        For i = start_cell To stop_cell
            If MF(i) = 1 And MFS(1, i).CStype <> 0 Then
'                If MFS(1, i).CStype = 1 Or MFS(1, i).CStype = 5 Then
'                    raster_form.ForeColor = vbRed
'                ElseIf MFS(1, i).CStype = 2 Or MFS(1, i).CStype = 6 Then
'                    raster_form.ForeColor = vbCyan
'                ElseIf MFS(1, i).CStype = 3 Or MFS(1, i).CStype = 7 Then
'                    raster_form.ForeColor = vbYellow
                If MFS(1, i).CStype = 4 Then
                    raster_form.ForeColor = vbWhite
                    raster_form.PSet (time_bin, i)
                End If
                
            End If
        Next i
        
    ElseIf Cell_type = 1 Then     'granule cells
        If mode = 0 Then
            raster_form.ForeColor = RGB(235, 235, 255)
            For i = start_cell To stop_cell
                If Gr(Raster_list(i)).act = 1 Then raster_form.PSet (time_bin, i)
            Next i
        ElseIf mode = 1 Then
            For i = start_cell To stop_cell
                rc = (GR_histo(Raster_list(i), bincounter5) / HistoDivisor) * 255
                raster_color = Int(rc)
                raster_form.PSet (time_bin, i), RGB(raster_color, raster_color, raster_color)
            Next i
        ElseIf mode = 2 Then
            For i = start_cell To stop_cell
                If Raster_plasticity(Raster_list(i)) <> 0 Then
                    If Raster_plasticity(Raster_list(i)) > 0 Then
                        rc2 = vbGreen
                    Else
                        rc2 = vbWhite
                    End If
                    raster_form.PSet (time_bin - 100, i), rc2
                End If
            Next i
        ElseIf mode = 3 Then
            For i = start_cell To stop_cell
                If Gr(Raster_list(Raster_list(i))).act = 1 Then
                    w = grWeight(Raster_list(i)) / Gr_weights_denom
                    w = w * 511
                    If w > 511 Then w = 511
                    If w > 255 Then
                        g = 255
                        r = 512 - ((w))
                        b = r
                    Else
                        g = w
                        r = 255
                        b = g
                    End If
                    raster_form.PSet (time_bin, i), RGB(r, g, b)
                End If
            Next i
        ElseIf mode = 4 Then
            For i = start_cell To stop_cell
                w = grWeight(Raster_list(i)) / Gr_weights_denom
                w = w * 511
                If w > 511 Then w = 511
                If w > 255 Then
                    g = 255
                    r = 512 - ((w))
                    b = r
                Else
                    g = w
                    r = 255
                    b = g
                End If
                rc = (GR_histo(Raster_list(i), bincounter5) / HistoDivisor)
                r = r * rc
                g = g * rc
                b = b * rc
                raster_form.PSet (time_bin, i), RGB(r, g, b)
            Next i
        End If
    ElseIf Cell_type = 2 Then
        If mode = 0 Then
            raster_form.ForeColor = vbWhite
            For i = start_cell To stop_cell
                If Gol(i).act = 1 Then raster_form.PSet (time_bin, i)
            Next i
        ElseIf mode = 1 Then
            For i = start_cell To stop_cell
                rc = (Go_histo(i, bincounter5) / HistoDivisor) * 255
                raster_color = Int(rc)
                raster_form.PSet (time_bin, i), RGB(raster_color, raster_color, raster_color)
            Next i
        End If
        
    ElseIf Cell_type = 3 Then
        For i = start_cell To stop_cell
            If Bk(i).act = 1 Then raster_form.PSet (time_bin, i)
            If i < 97 Then
                If BCells(i).act = 1 Then raster_form.PSet (time_bin, i + 250)
            End If
        Next i
    
    
    ElseIf Cell_type = 4 Then   '  STP eligibility
    
        For i = start_cell To stop_cell
            r = grPreElig(Raster_list(i)) * 255
            g = r
            b = r
            raster_form.PSet (time_bin, i), RGB(r, g, b)
        Next i
        
        
    ElseIf Cell_type = 5 Then   ' for drawing color coded status of STP at each synapse
        If mode = 0 Then
            r = 255
            For i = start_cell To stop_cell
                If Gr(Raster_list(i)).act = 1 Then
                    g = Int(255 * (1 - (2 * grSTP(Raster_list(i)))))
                    If g < 0 Then g = 0
                    b = g
                    raster_form.PSet (time_bin, i), RGB(r, g, b)
                End If
            Next i
        Else
            For i = start_cell To stop_cell
                r = Int((GR_histo(Raster_list(i), bincounter5) / HistoDivisor) * 255)
                g = Int(r * (1 - (2 * grSTP(Raster_list(i)))))
                If g < 0 Then g = 0
                b = g
                raster_form.PSet (time_bin, i), RGB(r, g, b)
            Next i
        End If
    End If
End Sub


Public Function ColorMap(X As Single, c As Integer)
Dim r, g, b As Single
Dim Ans As Integer
    If X <= 1# / 7# Then  ' blue and red rising to violet
        r = 2 * X
        g = 0
        b = 3.5 * X
    ElseIf X <= 2# / 7# Then  'blue rising, red falling to blue
        r = (2# / 7#) - ((X - 1# / 7#) * (2#))
        g = 0
        b = 3.5 * X
    ElseIf X <= 3# / 7# Then   'blue falling
        r = 0
        g = (X - 2# / 7#) * 7#
        b = 1
    ElseIf X <= 4# / 7# Then   'red rising
        r = 0
        g = 1
        b = 1# - (X - 3# / 7#) * 7#
    ElseIf X <= 5# / 7# Then 'green falling
        r = (X - 4# / 7#) * 8#
        g = 1
        b = 0
    ElseIf X <= 6# / 7# Then                   ' green and blue rising
        r = 1
        g = 1# - (X - 5# / 7#) * 6#
        b = 0
    Else
        r = 1
        g = (X - 6# / 7#) * 8#
        b = g
    End If
    'Debug.Print x, r, g, b
    Select Case c
        Case 1
            Ans = Int(255 * r)
        Case 2
            Ans = Int(255 * g)
        Case 3
            Ans = Int(255 * b)
    End Select
    
    If Ans > 255 Then Ans = 255
    
    If Ans < 0 Then Ans = 0
    ColorMap = Ans
End Function
