Attribute VB_Name = "simulation_graphics"
Option Explicit
Public max As Single
Public onsets(1000, 4) As Integer
Public EyeImages(70) As Picture
Public PuffImage As Picture
Public PCImage As Picture
Public PCImageX As Integer


Public Sub draw_histos(data_type As Integer, start_cell As Integer, stop_cell As Integer, max_cells As Integer, rows As Integer, columns As Integer, starter As Integer, stopper As Integer, threshold As Integer, cs_onset As Integer, cs_duration As Integer, max As Single)

Dim X As Integer
Dim Y As Integer
Dim times As Integer
Dim cap As String
Dim x_spacer As Single
Dim is_zero As Single
Dim current_sweeps(50) As Integer
Dim i As Integer
Dim j As Integer
Dim smooth(1000) As Single
Dim smooth2(1000) As Single
Dim smoother As Single
Dim baseline As Single
Dim base_counter As Single
Dim is_a_max As Integer
Dim baseline2 As Single
Dim integrate As Single
Dim integrate2 As Single
Dim counter As Integer

Histo_form.Cls
If start_cell < 1 Then
    start_cell = 1
    histo_info_form.cell_num_text.Text = 1
End If
x_spacer = (stopper - starter) + 50
Histo_form.ScaleWidth = columns * x_spacer
Histo_form.ScaleHeight = (rows * -1) - 0.05
Histo_form.ScaleTop = rows
counter = start_cell - 1
For X = 0 To columns - 1
    For Y = 0 To rows - 1
        counter = counter + 1
        If counter > max_cells Then counter = max_cells
        is_zero = 0
        
        If data_type = 0 Then
            While is_zero = 0
                For times = 1 To 1000
                   is_zero = is_zero + gr_histo(counter, times)
                Next times
                If is_zero < threshold Then
                    counter = counter + 1
                    If counter > max_cells Then counter = max_cells
                    is_zero = 0
                End If
            Wend
        End If
        
        Histo_form.Label1((Y * 5) + X).Top = Y + 0.8
        Histo_form.Label1((Y * 5) + X).Left = X * x_spacer
        Histo_form.Label1((Y * 5) + X).Caption = counter
        
        current_sweeps((Y * 5) + X + 1) = counter
        Histo_form.Label1((Y * 5) + X).Visible = True
        
        If data_type = 0 Then
            For times = starter To cs_onset - 1
                Histo_form.Line ((X * x_spacer) + times - starter, Y)-((X * x_spacer) + times - starter, Y + (gr_histo(counter, times) / max)), RGB(0, 0, 0)
            Next times
            
            For times = cs_onset To cs_onset + cs_duration - 1
                Histo_form.Line ((X * x_spacer) + times - starter, Y)-((X * x_spacer) + times - starter, Y + (gr_histo(counter, times) / max)), RGB(0, 0, 255)
            Next times
                    
            For times = cs_onset + cs_duration To stopper
                Histo_form.Line ((X * x_spacer) + times - starter, Y)-((X * x_spacer) + times - starter, Y + (gr_histo(counter, times) / max)), RGB(0, 0, 0)
            Next times
        End If
        
      If histo_info_form.Check1.Value = Checked Then
                   
        For times = starter + 2 To stopper - 2
            smoother = 0
            For i = times - 2 To times + 2
                smoother = smoother + gr_histo(counter, i)
            Next i
            smoother = smoother / 5#
            smooth(times) = smoother
        Next times
        
        For j = 1 To 10
            
            
            For times = starter + 2 To stopper - 2
                smoother = 0
                For i = times - 2 To times + 2
                    smoother = smoother + smooth(i)
                Next i
                smoother = smoother / 5#
                smooth(times) = smoother
                
            Next times
            DoEvents
        Next j
        
        baseline = 0
        base_counter = 0
        
        For times = starter To cs_onset - 1
            baseline = baseline + smooth(times)
            base_counter = base_counter + 1
        Next times
        baseline = baseline / base_counter
        
        
        For times = starter To stopper
            Histo_form.PSet ((X * x_spacer) + times - starter, Y + 0.2 + ((smooth(times) - baseline) / max)), RGB(255, 0, 0)
        Next times
        
        For times = cs_onset To stopper - 20
                is_a_max = 1
                integrate = 0
                
            For i = times - 20 To times + 20
                If (smooth(i) > smooth(times)) Then
                    is_a_max = 0
                    i = times + 20
                End If
                integrate = integrate + smooth(i)
            Next i
            
            If is_a_max = 1 Then
                integrate = integrate / 41
               If (smooth(times) < integrate + 4) Then
                    is_a_max = 0
               End If
            End If
            
            If is_a_max = 1 Then 'And smooth(times) > baseline + 20 Then
                Histo_form.Line ((X * x_spacer) + times - starter, Y + 0.2 + ((smooth(times) - baseline) / max))-((X * x_spacer) + times - starter, Y + 0.3 + ((smooth(times) - baseline) / max)), RGB(255, 0, 0)
            End If
        Next times
        
        If baseline > 30 Then
        For times = starter To stopper
            smooth2(times) = -1 * smooth(times)
        Next times
        baseline = -1 * baseline
        
         For times = cs_onset To stopper - 40
                is_a_max = 1
                integrate = 0
                integrate2 = 0
            For i = times - 40 To times + 40
                If i > times - 20 And i < times + 40 Then
                    If (smooth2(i) > smooth2(times)) Then
                        is_a_max = 0
                        i = times + 40
                    End If
                End If
                
                If i < times Then
                    integrate = integrate + smooth2(i)
                ElseIf i > times Then
                    integrate2 = integrate2 + smooth2(i)
                End If
            Next i
            
            If is_a_max = 1 Then
                
                integrate = integrate / 39
                integrate2 = integrate2 / 39
               If smooth2(times) < integrate + 5 Then is_a_max = 0
               
               If smooth2(times) < integrate2 + 5 Then is_a_max = 0
               
            End If
            If is_a_max = 1 And smooth2(times) > (baseline * -1) + 100 Then
                
                Histo_form.Line ((X * x_spacer) + times - starter, Y + 0.2 + ((smooth(times) - baseline) / max))-((X * x_spacer) + times - starter, Y + 0.3 + ((smooth(times) - baseline) / max)), RGB(0, 0, 255)
                
            End If
            
        Next times
        
        End If
      End If
        
        Histo_form.Line ((X * x_spacer), Y)-((X * x_spacer) + cs_onset - starter, Y - 0.03), RGB(0, 0, 0), BF
        Histo_form.Line ((X * x_spacer) + (cs_onset - starter), Y)-((X * x_spacer) + cs_onset + cs_duration - starter, Y - 0.03), RGB(0, 0, 255), BF
        Histo_form.Line ((X * x_spacer) + cs_onset + cs_duration - starter, Y)-((X * x_spacer) + stopper - starter, Y - 0.03), RGB(0, 0, 0), BF
        
        DoEvents
    Next Y
Next X

cap = "Granule cell histograms" + Str$(starter) + " to " + Str$(counter)
Histo_form.Caption = cap
End Sub

