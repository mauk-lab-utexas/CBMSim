Attribute VB_Name = "SynGenesis"
Public Sub SynaptoGenesis(Connectivity As Integer, UseUBCinSime As Integer)   ' 0 = self contained loops, 1 = intermixed
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim X As Integer
  Dim m As Integer
  Dim n As Integer
  Dim posxx As Integer
  Dim posyy As Integer
  Dim posx As Integer
  Dim posy As Integer
  Dim done As Integer
  Dim temp(MFNUMBER) As Integer
  Dim PCcount(PCNUMBER) As Integer
  Dim Temp3(6) As Integer
  Dim Temp2(24) As Integer
  Dim Temp4(8) As Integer
  Dim s As Double
  
  Dim BC As Integer
  
  Randomize Timer

'  Progress.Caption = "Synaptogenesis"
'  Progress.Visible = True
  
  ' Glomerulus assignment
Erase Gl
m = GlX * GlY
For i = 1 To MFNUMBER    'make sure each MF has at least one glom
    done = 0
    Do Until done = 1
        j = Int(Rnd * m)
        If j = 0 Then j = 1
        If Gl(j).MF = 0 Then
            Gl(j).MF = i
            done = 1
        End If
    Loop
Next i
Erase temp
For i = 1 To m   'assign the extra gloms a mf and make sure each mf isn't used more than twice
    If Gl(i).MF = 0 Then
        done = 0
        Do Until done = 1
            j = Int(Rnd * MFNUMBER)
            If j > 0 And temp(j) = 0 Then
                Gl(i).MF = j
                temp(j) = 1
                done = 1
            End If
        Loop
    End If
Next i

'  For i = 1 To GlX * GlY  '  This is the old version used until December 2007
'    Gl(i).MF = Int(Rnd * MFNUMBER)
'    If Gl(i).MF = 0 Then Gl(i).MF = 1
'  Next i

  'MF to Granule synaptogenesis
  
'  Progress.Label1.Caption = "MF to granule connections"
  'Progress.ProgressBar1.Max = GrX * GrY
  For X = 1 To GrX * GrY
    'Progress.ProgressBar1.Value = X
    posxx = X Mod GrX
    If posxx = 0 Then posxx = GrX
    posyy = Int((X - 1) / GrX) + 1
    'Debug.Print x, posxx, posyy
    
    Gr(X).numdend = Int((MaxGrDend - MinGrDend + 1) * Rnd + MinGrDend)
    For i = 1 To Gr(X).numdend
      done = 0
      Do Until done = 1
        DoEvents
        posy = Int(posyy + ((GrDendSpanY * Rnd) - (GrDendSpanY / 2#)))
        posx = Int(posxx + ((GrDendSpanX * Rnd) - (GrDendSpanX / 2#)))
        posy = posy / GrGlScaleY
        posx = posx / GrGlScaleX
        
        If posx > GlX Then posx = posx - GlX
        If posx < 1 Then posx = posx + GlX
        If posy > GlY Then posy = posy - GlY
        If posy < 1 Then posy = posy + GlY
        'Debug.Print x, i, posxx, posyy, posx, posy,
        posx = ((posy - 1) * GlX) + posx
        done = 1
        For j = 1 To i
          If posx = prex(j) Then done = 0
        Next j
      Loop
      
      prex(i) = posx      'for visual graphics only
      Gr(X).MF(i) = Gl(posx).MF
      'Debug.Print posx, Gr(x).MF(i)
    Next i
    For i = 1 To MaxGrDend
        'If Rnd() < 0.6 Then fastMask(i) = 1 Else fastMask(i) = 0
        If Rnd() < 0.6 Then Gr(X).fastMask(i) = 1 Else Gr(X).fastMask(i) = 0
    Next i
    
  Next X
  
  ' Golgi to granule connections
'  Progress.Label1.Caption = "Golgi to granule connections"
  
  For X = 1 To GrX * GrY
    posxx = X Mod GrX
    If posxx = 0 Then posxx = GrX
    posyy = Int((X - 1) / GrX) + 1
    For i = 1 To Gr(X).numdend
      done = 0
      Do Until done = 1
        DoEvents
        posy = Int(posyy + ((GrGoDendSpanY * Rnd) - (GrGoDendSpanY / 2#)))
        posx = Int(posxx + ((GrGoDendSpanX * Rnd) - (GrGoDendSpanX / 2#)))
        posy = posy / GrGoScaleY
        posx = posx / GrGoScaleX
        
        If posx > GoX Then posx = posx - GoX
        If posx < 1 Then posx = posx + GoX
        If posy > GoY Then posy = posy - GoY
        If posy < 1 Then posy = posy + GoY
        'Debug.Print x, i, posxx, posyy, posx, posy,
        posx = ((posy - 1) * GoX) + posx
        done = 1
        For j = 1 To i
          If posx = Gr(X).Gol(j) Then done = 0
        Next j
      Loop
      Gr(X).Gol(i) = posx
    Next i
  Next X
  
  
  ' Glomerulus to Golgi connections
'  Progress.Label1.Caption = "Glomerulus to Golgi connections"
  
  For X = 1 To GoX * GoY
    posxx = X Mod GoX
    If posxx = 0 Then posxx = GoX
    posyy = (Int((X - 1) / GoX) + 1)
    For i = 1 To numGoGlDend
      done = 0
      Do Until done = 1
        DoEvents
        posy = Int(posyy + ((GoGlDendSpanY * Rnd) - (GoGlDendSpanY / 2#)))
        posx = Int(posxx + ((GoGlDendSpanX * Rnd) - (GoGlDendSpanX / 2#)))
        'posy = posy / GrGlScaleY
        'posx = posx / GrGlScaleX
        
        If posx > GlX Then posx = posx - GlX
        If posx < 1 Then posx = posx + GlX
        If posy > GlY Then posy = posy - GlY
        If posy < 1 Then posy = posy + GlY
        'Debug.Print X, i, posxx, posyy, posx, posy
        posx = ((posy - 1) * GlX) + posx
        done = 1
        For j = 1 To i
          If posx = Gol(X).preGl(j) Then done = 0
        Next j
      Loop
      Gol(X).preGl(i) = posx
      'Debug.Print X, i, gol(X).preGl(i)
      Gol(X).MF(i) = Gl(posx).MF
      'Debug.Print X, gol(X).preGl(i), posx
    Next i
  Next X
  
  
  ' granule to Golgi connections
'  Progress.Label1.Caption = "Granule to Golgi connections"
  
  For X = 1 To GoX * GoY
    posxx = (X Mod GoX) * GrGoScaleX
    If posxx = 0 Then posxx = GoX
    posyy = (Int((X - 1) / GoX) + 1) * GrGoScaleY
    
    For i = 1 To numGoGrDend
      done = 0
      Do Until done = 1
        DoEvents
        posy = Int(posyy + ((GoDendSpanY * Rnd) - (GoDendSpanY / 2#)))
        posx = Int(posxx + ((GoDendSpanX * Rnd) - (GoDendSpanX / 2#)))
        'posy = posy / GrGlScaleY
        'posx = posx / GrGlScaleX
        
        If posx > GrX Then posx = posx - GrX
        If posx < 1 Then posx = posx + GrX
        If posy > GrY Then posy = posy - GrY
        If posy < 1 Then posy = posy + GrY
        'Debug.Print x, i, posxx, posyy, posx, posy,
        posx = ((posy - 1) * GrX) + posx
        done = 1
        For j = 1 To i
          If posx = Gol(X).preGr(j) Then done = 0
        Next j
      Loop
      Gol(X).preGr(i) = posx
      
    Next i
  Next X
  
  
'  Progress.Label1.Caption = "STELLATE"
  DoEvents
    For i = 1 To PCNUMBER   '24
        n = 1
      While (n <= StellPCSYNUMBER)   '10
        
        Stellsynapse = Int((Rnd() * (STELLATENUMBER)) + 1)  '240
        
        already_connected = 0
        m = 1
        While ((already_connected = 0) And (m <= n))
          If (Pc(i).Stellsyn(m) = Stellsynapse) Then
            already_connected = 1
          Else
            m = m + 1
          End If
        Wend
        If (Not already_connected) Then
          Pc(i).Stellsyn(n) = Stellsynapse
          n = n + 1
        End If
      Wend
    Next i
    
'  Progress.Label1.Caption = "BASKET"
  DoEvents
   
    For i = 1 To PCNUMBER  '' create PC to basket synapses      PCNUMBER is 24
        BC = (PCBasketsynUMBER * (i - 1)) + 1    ' PCBasketsynUMBER is 4
        BCells(BC).PCsyn(1) = i + 1              ' each BC gets input from 4 PCs
        BCells(BC).PCsyn(2) = i + 2
        BCells(BC).PCsyn(3) = i - 1
        BCells(BC).PCsyn(4) = i - 2
            
        BC = (PCBasketsynUMBER * (i - 1)) + 2
        BCells(BC).PCsyn(1) = i + 1
        BCells(BC).PCsyn(2) = i + 3
        BCells(BC).PCsyn(3) = i - 1
        BCells(BC).PCsyn(4) = i - 3
    
        BC = (PCBasketsynUMBER * (i - 1)) + 3
        BCells(BC).PCsyn(1) = i + 3
        BCells(BC).PCsyn(2) = i + 6
        BCells(BC).PCsyn(3) = i - 3
        BCells(BC).PCsyn(4) = i - 6
        
        BC = (PCBasketsynUMBER * (i - 1)) + 4
        BCells(BC).PCsyn(1) = i + 4
        BCells(BC).PCsyn(2) = i + 9
        BCells(BC).PCsyn(3) = i - 4
        BCells(BC).PCsyn(4) = i - 9
    Next i
    
    For i = 1 To BasketNUMBER
        For j = 1 To PCBasketsynUMBER
            If BCells(i).PCsyn(j) > 24 Then BCells(i).PCsyn(j) = BCells(i).PCsyn(j) - 24
            If BCells(i).PCsyn(j) < 1 Then BCells(i).PCsyn(j) = BCells(i).PCsyn(j) + 24
        Next j
    Next i
    
    
    
    
'  Progress.Label1.Caption = "PURKINJE-NUCLEUS"
  DoEvents
    If Connectivity < 2 Then   'loops
        PCNCSYNUMBER = 12
        Nc(1).PCsyn(1) = 1
        Nc(1).PCsyn(2) = 2
        Nc(1).PCsyn(3) = 3
        Nc(1).PCsyn(4) = 4
        Nc(1).PCsyn(5) = 5
        Nc(1).PCsyn(6) = 6
       
        Nc(2).PCsyn(1) = 1
        Nc(2).PCsyn(2) = 2
        Nc(2).PCsyn(3) = 3
        Nc(2).PCsyn(4) = 4
        Nc(2).PCsyn(5) = 5
        Nc(2).PCsyn(6) = 6
       
        Nc(3).PCsyn(1) = 7
        Nc(3).PCsyn(2) = 8
        Nc(3).PCsyn(3) = 9
        Nc(3).PCsyn(4) = 10
        Nc(3).PCsyn(5) = 11
        Nc(3).PCsyn(6) = 12
        
        Nc(4).PCsyn(1) = 7
        Nc(4).PCsyn(2) = 8
        Nc(4).PCsyn(3) = 9
        Nc(4).PCsyn(4) = 10
        Nc(4).PCsyn(5) = 11
        Nc(4).PCsyn(6) = 12
       
        Nc(5).PCsyn(1) = 13
        Nc(5).PCsyn(2) = 14
        Nc(5).PCsyn(3) = 15
        Nc(5).PCsyn(4) = 16
        Nc(5).PCsyn(5) = 17
        Nc(5).PCsyn(6) = 18
     
        Nc(6).PCsyn(1) = 13
        Nc(6).PCsyn(2) = 14
        Nc(6).PCsyn(3) = 15
        Nc(6).PCsyn(4) = 16
        Nc(6).PCsyn(5) = 17
        Nc(6).PCsyn(6) = 18
    
        Nc(7).PCsyn(1) = 19
        Nc(7).PCsyn(2) = 20
        Nc(7).PCsyn(3) = 21
        Nc(7).PCsyn(4) = 22
        Nc(7).PCsyn(5) = 23
        Nc(7).PCsyn(6) = 24

        Nc(8).PCsyn(1) = 19
        Nc(8).PCsyn(2) = 20
        Nc(8).PCsyn(3) = 21
        Nc(8).PCsyn(4) = 22
        Nc(8).PCsyn(5) = 23
        Nc(8).PCsyn(6) = 24
    ElseIf Connectivity = 2 Then                   ' 4 CFs random
        PCNCSYNUMBER = 12
        Nc(1).PCsyn(1) = 1
        Nc(1).PCsyn(2) = 2
        Nc(1).PCsyn(3) = 3
'        Nc(1).Pcsyn(4) = 4
'        Nc(1).Pcsyn(5) = 5
'        Nc(1).Pcsyn(6) = 6
       
'        Nc(2).Pcsyn(1) = 1
'        Nc(2).Pcsyn(2) = 2
'        Nc(2).Pcsyn(3) = 3
        Nc(2).PCsyn(4) = 4
        Nc(2).PCsyn(5) = 5
        Nc(2).PCsyn(6) = 6
       
        Nc(3).PCsyn(1) = 7
        Nc(3).PCsyn(2) = 8
        Nc(3).PCsyn(3) = 9
'        Nc(3).Pcsyn(4) = 10
'        Nc(3).Pcsyn(5) = 11
'        Nc(3).Pcsyn(6) = 12
        
'        Nc(4).Pcsyn(1) = 7
'        Nc(4).Pcsyn(2) = 8
'        Nc(4).Pcsyn(3) = 9
        Nc(4).PCsyn(4) = 10
        Nc(4).PCsyn(5) = 11
        Nc(4).PCsyn(6) = 12
       
        Nc(5).PCsyn(1) = 13
        Nc(5).PCsyn(2) = 14
        Nc(5).PCsyn(3) = 15
'        Nc(5).Pcsyn(4) = 16
'        Nc(5).Pcsyn(5) = 17
'        Nc(5).Pcsyn(6) = 18
     
'        Nc(6).Pcsyn(1) = 13
'        Nc(6).Pcsyn(2) = 14
'        Nc(6).Pcsyn(3) = 15
        Nc(6).PCsyn(4) = 16
        Nc(6).PCsyn(5) = 17
        Nc(6).PCsyn(6) = 18
    
        Nc(7).PCsyn(1) = 19
        Nc(7).PCsyn(2) = 20
        Nc(7).PCsyn(3) = 21
'        Nc(7).Pcsyn(4) = 22
'        Nc(7).Pcsyn(5) = 23
'        Nc(7).Pcsyn(6) = 24

'        Nc(8).Pcsyn(1) = 19
'        Nc(8).Pcsyn(2) = 20
'        Nc(8).Pcsyn(3) = 21
        Nc(8).PCsyn(4) = 22
        Nc(8).PCsyn(5) = 23
        Nc(8).PCsyn(6) = 24
        
        
        
        
        For i = 1 To NCNUMBER
            For j = 1 To PCNCSYNUMBER
                If Nc(i).PCsyn(j) > 0 Then
                    PCcount(Nc(i).PCsyn(j)) = PCcount(Nc(i).PCsyn(j)) + 1
                End If
            Next j
        Next i

'        For i = 1 To 24
'            Debug.Print i, PCcount(i)
'        Next i
        
        For i = 1 To NCNUMBER
            For j = 1 To PCNCSYNUMBER
                If Nc(i).PCsyn(j) = 0 Then
                    done = 0
                    While done = 0
                        k = Int(Rnd() * 24) + 1
                        If PCcount(k) < 3 Then
                            Nc(i).PCsyn(j) = k
                            PCcount(k) = PCcount(k) + 1
                            done = 1
                        End If
                    Wend
                End If
            Next j
        Next i
'        For i = 1 To 24
'            Debug.Print i, PCcount(i)
'        Next i
    ElseIf Connectivity = 3 Then  ' this is the 12 CF model
        PCNCSYNUMBER = 12
        Nc(1).PCsyn(1) = 1   ' half are fixed
        Nc(1).PCsyn(2) = 2
        Nc(1).PCsyn(3) = 3
'        Nc(1).Pcsyn(4) = 4
'        Nc(1).Pcsyn(5) = 5
'        Nc(1).Pcsyn(6) = 6
       
'        Nc(2).Pcsyn(1) = 1
'        Nc(2).Pcsyn(2) = 2
'        Nc(2).Pcsyn(3) = 3
        Nc(2).PCsyn(4) = 4
        Nc(2).PCsyn(5) = 5
        Nc(2).PCsyn(6) = 6
       
        Nc(3).PCsyn(1) = 7
        Nc(3).PCsyn(2) = 8
        Nc(3).PCsyn(3) = 9
'        Nc(3).Pcsyn(4) = 10
'        Nc(3).Pcsyn(5) = 11
'        Nc(3).Pcsyn(6) = 12
        
'        Nc(4).PCsyn(1) = 10
'        Nc(4).PCsyn(2) = 11
'        Nc(4).PCsyn(3) = 12
'        Nc(4).PCsyn(4) = 4
'        Nc(4).PCsyn(5) = 5
'        Nc(4).PCsyn(6) = 6
       
'        Nc(5).PCsyn(1) = 13
'        Nc(5).PCsyn(2) = 14
'        Nc(5).PCsyn(3) = 15
'        Nc(5).PCsyn(4) = 16
'        Nc(5).PCsyn(5) = 17
'        Nc(5).PCsyn(6) = 18
     
'        Nc(6).Pcsyn(1) = 13
'        Nc(6).Pcsyn(2) = 14
'        Nc(6).Pcsyn(3) = 15
        Nc(6).PCsyn(4) = 16
        Nc(6).PCsyn(5) = 17
        Nc(6).PCsyn(6) = 18
    
        Nc(7).PCsyn(1) = 19
        Nc(7).PCsyn(2) = 20
        Nc(7).PCsyn(3) = 21
'        Nc(7).Pcsyn(4) = 22
'        Nc(7).Pcsyn(5) = 23
'        Nc(7).Pcsyn(6) = 24

'        Nc(8).Pcsyn(1) = 19
'        Nc(8).Pcsyn(2) = 20
'        Nc(8).Pcsyn(3) = 21
        Nc(8).PCsyn(4) = 22
        Nc(8).PCsyn(5) = 23
        Nc(8).PCsyn(6) = 24
        
        
        For i = 1 To PCNUMBER
            PCcount(i) = i
        Next i
        
        For i = 1 To 12
            done = 0
            While done = 0
                k = Int(Rnd() * 24) + 1
                If PCcount(k) > 0 Then
                    Nc(4).PCsyn(i) = k
                    PCcount(k) = 0
                    done = 1
                End If
            Wend
        Next i
        
        For i = 1 To 12
            done = 0
            While done = 0
                k = Int(Rnd() * 24) + 1
                If PCcount(k) > 0 Then
                    Nc(5).PCsyn(i) = k
                    PCcount(k) = 0
                    done = 1
                End If
            Wend
        Next i
'
    
        For i = 1 To PCNUMBER
                PCcount(i) = i
        Next i

        For i = 1 To 12
            done = 0
            While done = 0
                k = Int(Rnd() * 24) + 1
                If PCcount(k) > 0 Then
                    Nc(1).PCsyn(i) = k
                    PCcount(k) = 0
                    done = 1
                End If
            Wend
        Next i
        
        For i = 1 To 12
            done = 0
            While done = 0
                k = Int(Rnd() * 24) + 1
                If PCcount(k) > 0 Then
                    Nc(2).PCsyn(i) = k
                    PCcount(k) = 0
                    done = 1
                End If
            Wend
        Next i

        For i = 1 To PCNUMBER
                PCcount(i) = i
        Next i

        For i = 1 To 12
            done = 0
            While done = 0
                k = Int(Rnd() * 24) + 1
                If PCcount(k) > 0 Then
                    Nc(3).PCsyn(i) = k
                    PCcount(k) = 0
                    done = 1
                End If
            Wend
        Next i
        
        For i = 1 To 12
            done = 0
            While done = 0
                k = Int(Rnd() * 24) + 1
                If PCcount(k) > 0 Then
                    Nc(6).PCsyn(i) = k
                    PCcount(k) = 0
                    done = 1
                End If
            Wend
        Next i

        For i = 1 To PCNUMBER
                PCcount(i) = i
        Next i

        For i = 1 To 12
            done = 0
            While done = 0
                k = Int(Rnd() * 24) + 1
                If PCcount(k) > 0 Then
                    Nc(7).PCsyn(i) = k
                    PCcount(k) = 0
                    done = 1
                End If
            Wend
        Next i
        
        For i = 1 To 12
            done = 0
            While done = 0
                k = Int(Rnd() * 24) + 1
                If PCcount(k) > 0 Then
                    Nc(8).PCsyn(i) = k
                    PCcount(k) = 0
                    done = 1
                End If
            Wend
        Next i
        For i = 1 To NCNUMBER
            For j = 1 To PCNCSYNUMBER
                Debug.Print i, j, Nc(i).PCsyn(j)
            Next j
            Debug.Print
        Next i
        
        
        
'        For i = 1 To NCNUMBER
'            For j = 1 To PCNCSYNUMBER
'                If Nc(i).PCsyn(j) > 0 Then
'                    PCcount(Nc(i).PCsyn(j)) = PCcount(Nc(i).PCsyn(j)) + 1
'                End If
'            Next j
'        Next i
    
    End If
    
    

        
'    For i = 1 To NCNUMBER
'        For j = 1 To PCNCSYNUMBER
'            If Nc(i).PCsyn(j) = 0 Then
'                done = 0
'                While done = 0
'                    k = Int(Rnd() * 24) + 1
'                    If PCcount(k) < 8 Then
'                        Nc(i).PCsyn(j) = k
'                        PCcount(k) = PCcount(k) + 1
'                        done = 1
'                    End If
'                Wend
'            End If
'        Next j
'    Next i
'  For i = 1 To NCNUMBER
'        For j = 1 To PCNCSYNUMBER
'            Debug.Print i, j, Nc(i).PCsyn(j)
'        Next j
'        Debug.Print
'   Next i
   
   
   
    If NumCF = 4 Then
        NCCFSYNUMBER = 2
        CF(1).NucInput(1) = 1
        CF(1).NucInput(2) = 1
        
        CF(2).NucInput(3) = 1
        CF(2).NucInput(4) = 1
        
        CF(3).NucInput(5) = 1
        CF(3).NucInput(6) = 1
       
        CF(4).NucInput(7) = 1
        CF(4).NucInput(8) = 1
        
    ElseIf NumCF = 1 Then   'NEED TO ADJUST STRENGTH OF NUC TO CF DEPENDING ON 1 VERSUS 4 CF SO THAT EACH CAN START CLOSE TO EQUIL
    ' NEED TO ALTER PC WINDOW TO MAKE ROOM FOR EXTRA NUC CELLS....
        NCCFSYNUMBER = 8
        CF(1).NucInput(1) = 1
        CF(1).NucInput(2) = 1
        CF(1).NucInput(3) = 1
        CF(1).NucInput(4) = 1
        CF(1).NucInput(5) = 1
        CF(1).NucInput(6) = 1
        CF(1).NucInput(7) = 1
        CF(1).NucInput(8) = 1
        
    ElseIf NumCF = 12 Then
        NCCFSYNUMBER = 1
        CF(1).NucInput(4) = 1
        CF(2).NucInput(4) = 1
        CF(3).NucInput(4) = 1
        CF(4).NucInput(4) = 1
        CF(5).NucInput(4) = 1
        CF(6).NucInput(4) = 1
       
        CF(7).NucInput(5) = 1
        CF(8).NucInput(5) = 1
        CF(9).NucInput(5) = 1
        CF(10).NucInput(5) = 1
        CF(11).NucInput(5) = 1
        CF(12).NucInput(5) = 1
    End If
' ************************** Golgi to Golgi inhibition connectivity (nearest neighbor)
   
   For X = 1 To 900
'    Debug.Print x, ToLeft(x), ToRight(x), ToTop(x), ToBottom(x),
'    i = ToLeft(x)
'    Debug.Print ToTop(i), ToBottom(i),
'    i = ToRight(x)
'    Debug.Print ToTop(i), ToBottom(i)

        GG(X, 1) = ToLeft(X)
        GG(X, 2) = ToTop(GG(X, 1))
        GG(X, 3) = ToBottom(GG(X, 1))
        GG(X, 4) = ToRight(X)
        GG(X, 5) = ToTop(GG(X, 4))
        GG(X, 6) = ToBottom(GG(X, 4))
        GG(X, 7) = ToTop(X)
        GG(X, 8) = ToBottom(X)
        
   Next X
   
   
   Erase wGG
    Gol(0).act = 0    '  this makes 0 input when GG does not exist
 ' This randomly pick GGPercent worth of the Golgi to Golgi connections to include

    For X = 1 To 900
        For i = 1 To 8
            s = Rnd
            If s > (GGPercent * 0.01) Then wGG(X, i) = 0 Else wGG(X, i) = 1
        Next i
    Next X

    
' This picks a fixed number of Golgi to Golgi connections for each Golgi cell to exclude
'    Erase Temp2
'    Erase Temp4
'    For X = 1 To 900
'        i = 1
'        Temp2(i) = Int((Rnd() * 8) + 1)
'        'Debug.Print Temp2(1);
'        While i < 4  '  this is how many of the connection to make out of the 8
'            j = Int((Rnd() * 8) + 1)
'
'            For k = 1 To 3  '  this is connected to the 4 above
'                If j = Temp2(k) Then j = 100
'            Next k
'            If j <> 100 Then
'                i = i + 1
'                Temp2(i) = j
'         '       Debug.Print j;
'                Temp4(j) = Temp4(j) + 1
'                wGG(X, j) = 1
'            End If
'        Wend
'        'Debug.Print
'
'    Next X
'    For i = 1 To 8
'        Debug.Print Temp4(i)
'    Next i
    
DoEvents



End Sub
Public Sub Init_stuff()
Dim X As Integer
Dim m As Integer
Dim i As Integer
Dim j As Integer
Dim i2, change As Single

hyper_switch = 1
gr_elig_counter = 0

For X = 1 To GoX * GoY
 Gol(X).ThrBase = ThrBaseGo '     /*+GAUSS(ThrBaseVarGo)*/
Next X
                                 'zero background firing rate of all mossy fibers
    For X = 1 To MFNUMBER
        MFS(1, X).CStype = 0
    Next X

'*************** TONIC MOSSY FIBERS FOR CS 1 through 4

    X = 0                                       'identify tonic mossy fibers for CS1
    While (X < MFNUMBER * NUMTONIC)
        i = Rnd() / NUMTONIC + X / NUMTONIC
        'Debug.Print x, i
        If (MFS(1, i).CStype = 0) Then
          MFS(1, i).CStype = 1
          X = X + 1
        End If
    Wend

    X = 0
    While (X < MFNUMBER * NUMTONIC2)                    ' set CS2 mossy fibers to tonic
        i = Rnd() / NUMTONIC2 + X / NUMTONIC2
        If (MFS(1, i).CStype = 0) Then
          MFS(1, i).CStype = 2
          X = X + 1
          'Debug.Print "Tonic "; X, i
        End If
    Wend

    X = 0
    While (X < MFNUMBER * NUMTONIC3)                    ' set CS3 mossy fibers to tonic
        i = Rnd() / NUMTONIC3 + X / NUMTONIC3
        If (MFS(1, i).CStype = 0) Then
          MFS(1, i).CStype = 3
          X = X + 1
          'Debug.Print "Tonic "; X, i
        End If
    Wend
    
    If UseUBCs = 0 Then
        X = 0
        While (X < MFNUMBER * NUMTONIC4)                    ' set CS4 mossy fibers to tonic
            i = Rnd() / NUMTONIC4 + X / NUMTONIC4
            If (MFS(1, i).CStype = 0) Then
              MFS(1, i).CStype = 4
              X = X + 1
              Debug.Print "Tonic 4"; X, i
            End If
        Wend
    Else
        X = 0
        While (X < MFNUMBER * NUMUBC)                    ' set CS4 mossy fibers to UBCs if they are turned on
            i = Rnd() / NUMTONIC4 + X / NUMTONIC4
            If (MFS(1, i).CStype = 0) Then
              MFS(1, i).CStype = 4
              X = X + 1
              Debug.Print "UBC "; X, i
              MFS(1, i).UBCduration = Int(MinUBCDuration + (Rnd() * (MaxUBCDuration - MinUBCDuration)))
            End If
        Wend
    End If
'****************** PHASIC MOSSY FIBERS FOR CS 1 THROUGH 4
                                           
    X = 0                                                   'phasic mossy fibers  CS1
    While (X < MFNUMBER * NUMPHASIC)
        i = Rnd() / NUMPHASIC + X / NUMPHASIC
        If (MFS(1, i).CStype = 0) Then
          MFS(1, i).CStype = 5
          X = X + 1
          'Debug.Print "Phasic "; X, i
        End If
    Wend


    X = 0                                                   'phasic mossy fibers  CS2
    While (X < MFNUMBER * NUMPHASIC2)
        i = Rnd() / NUMPHASIC2 + X / NUMPHASIC2
        If (MFS(1, i).CStype = 0) Then
          MFS(1, i).CStype = 6
          X = X + 1
          'Debug.Print "Phasic "; X, i
        End If
    Wend

    X = 0                                                   'phasic mossy fibers  CS3
    While (X < MFNUMBER * NUMPHASIC3)
        i = Rnd() / NUMPHASIC3 + X / NUMPHASIC3
        If (MFS(1, i).CStype = 0) Then
          MFS(1, i).CStype = 7
          X = X + 1
          'Debug.Print "Phasic "; X, i
        End If
    Wend
    
    X = 0                                                   'phasic mossy fibers  CS4
    While (X < MFNUMBER * NUMPHASIC4)
        i = Rnd() / NUMPHASIC4 + X / NUMPHASIC4
        If (MFS(1, i).CStype = 0) Then
          MFS(1, i).CStype = 8
          X = X + 1
          'Debug.Print "Phasic "; X, i
        End If
    Wend
    
'*****************CONTEXT MOSSY FIBERS FOR CONTEXT 1 and 2
    X = 0
    
    While (X < MFNUMBER * NUMCONTEXT)
        i = Rnd() / NUMCONTEXT + X / NUMCONTEXT
        If (MFS(1, i).CStype = 0) Then
          MFS(1, i).CStype = 9
          X = X + 1
          'Debug.Print "Context "; X, i
        End If
    Wend
    
    X = 0
    While (X < MFNUMBER * NUMCONTEXT2)
        i = Rnd() / NUMCONTEXT2 + X / NUMCONTEXT2
        If (MFS(1, i).CStype = 0) Then
          MFS(1, i).CStype = 10
          X = X + 1
          'Debug.Print "Context "; X, i
        End If
    Wend
    
'    For i = 1 To 600
'        Debug.Print i, MFS(1, i).CStype
'    Next i
                                                    
AssignMFs

    For m = 1 To MFNUMBER
        MfElig(m) = 0#
        'If MFS(1, m).CStype > 10 Then Debug.Print m, MFS(1, m).CStype
    Next m

    For m = 1 To PCNUMBER
      Pc(m).Thr = THRBASEPC
      Pc(m).GGr = 0#
      Pc(m).GStell = 0#
      Pc(m).act = 0
      Pc(m).v = ELEAKPC + (Rnd() * 6)
      Pc(m).ThrBase = THRBASEPC
    Next m
    
    For m = 1 To STELLATENUMBER
      StellPcG(m) = 0#
      Bk(m).Thr = THRBASEStell
      Bk(m).GGr = 0#
      Bk(m).act = 0
      Bk(m).v = ELeakStell
    Next m
    
    For m = 1 To BasketNUMBER
        BCells(m).GGr = 0
        BCells(m).gPc = 0
        BCells(m).v = ELeakBC
        BCells(m).Thr = ThrBaseBC
        For i = 1 To PFBasketsynUMBER
            BCells(m).grW(i) = 0.4  'BCReviewers  .4 normal
        Next i
    Next m
    
    For m = 1 To NCNUMBER
      Nc(m).Thr = THRBASENC
      Nc(m).gPc = 0#
      Nc(m).gMF = 0#
      Nc(m).act = 0
      Nc(m).v = ELEAKNC
    Next m
    For m = 1 To GrX * GrY
'        Gr(m).gi = 0#
'        Gr(m).gE = 0
        Gr(m).v = ELeakgr
        Gr(m).act = 0
        Gr(m).Thr = ThrBasegr
        Gr(m).ThrBase = ThrBasegr
        'Gr(m).g_Var = gEconstGr  'Mike in May added
    Next m
    For m = 1 To GoX * GoY
     
    Gol(m).Thr = ThrBaseGo
    Gol(m).gMF = 0#
    Gol(m).GGr = 0
    Gol(m).v = ELGo
    Gol(m).act = 0
     
    Next m
    For m = 1 To NumCF
        CF(m).Thr = THRMAXCF
        CF(m).GNc = 0#
        CF(m).act = 0
        CF(m).v = ELEAKCF
    Next m
    For X = 1 To SYNUMBER
        grWeight(X) = 0.54
    Next X

    For X = 1 To MFNUMBER
        mfweight(X) = GCONSTMFNC
        If MFS(1, X).CStype = 1 Or MFS(1, X).CStype = 5 Then
            mfweight(X) = 0
        End If
    Next X

    For i = 1 To PCNUMBER
        For j = 1 To NCNUMBER
            gPURKtoNUCLEUS(i, j) = gPurktoNucBeginAverage '0.27 ' gPURKtoNUCLEUS = 0.27
            If SimOptionsForm.UniformityOption(1).Value = True Then
                change = ((Rnd() - 0.5) * 0.2) * gPURKtoNUCLEUS(i, j)
                gPURKtoNUCLEUS(i, j) = gPURKtoNUCLEUS(i, j) + change
            End If
        Next j
    Next i
    For i = 1 To NCNUMBER
        For j = 1 To NumCF
            Nc(i).gNUCtoCF(j) = gNuctoCFBeginAverage '0.1 '0.045 '0.037
            If SimOptionsForm.UniformityOption(1).Value = True Then
                change = ((Rnd() - 0.5) * 0.2) * Nc(i).gNUCtoCF(j)
                Nc(i).gNUCtoCF(j) = Nc(i).gNUCtoCF(j) + change
            End If
        Next j
    Next i
End Sub

Public Sub AssignMFs()
Dim X As Integer
Dim i As Integer

'set CS firing rates for mossy fibers
    For X = 1 To MFNUMBER
        For i = 1 To 2
            Select Case MFS(1, X).CStype
                Case 1
                    MFS(i, X).bfreq = MFBGROUNDFREQMIN_CS + Rnd * (MFBGROUNDFREQMAX_CS - MFBGROUNDFREQMIN_CS)
                    MFS(i, X).CSFreq = MFTONICFREQ_INCREMENT
                Case 2
                    MFS(i, X).bfreq = MFBGROUNDFREQMIN_CS + Rnd * (MFBGROUNDFREQMAX_CS - MFBGROUNDFREQMIN_CS)
                    MFS(i, X).CSFreq = MFTONICFREQ_INCREMENT2
                Case 3
                    MFS(i, X).bfreq = MFBGROUNDFREQMIN_CS + Rnd * (MFBGROUNDFREQMAX_CS - MFBGROUNDFREQMIN_CS)
                    MFS(i, X).CSFreq = MFTONICFREQ_INCREMENT3
                Case 4
                    MFS(i, X).bfreq = MFBGROUNDFREQMIN_CS + Rnd * (MFBGROUNDFREQMAX_CS - MFBGROUNDFREQMIN_CS)
                    MFS(i, X).CSFreq = MFTONICFREQ_INCREMENT4
                Case 5
                    MFS(i, X).bfreq = MFBGROUNDFREQMIN_CS + Rnd * (MFBGROUNDFREQMAX_CS - MFBGROUNDFREQMIN_CS)
                    MFS(i, X).CSFreq = MFPHASICFREQ_INCREMENT
                Case 6
                    MFS(i, X).bfreq = MFBGROUNDFREQMIN_CS + Rnd * (MFBGROUNDFREQMAX_CS - MFBGROUNDFREQMIN_CS)
                    MFS(i, X).CSFreq = MFPHASICFREQ_INCREMENT2
                Case 7
                    MFS(i, X).bfreq = MFBGROUNDFREQMIN_CS + Rnd * (MFBGROUNDFREQMAX_CS - MFBGROUNDFREQMIN_CS)
                    MFS(i, X).CSFreq = MFPHASICFREQ_INCREMENT3
                Case 8
                    MFS(i, X).bfreq = MFBGROUNDFREQMIN_CS + Rnd * (MFBGROUNDFREQMAX_CS - MFBGROUNDFREQMIN_CS)
                    MFS(i, X).CSFreq = MFPHASICFREQ_INCREMENT4
                Case 9  'context 1
                    MFS(1, X).bfreq = MFCONTEXTFREQMIN + Rnd * (MFCONTEXTFREQMAX - MFCONTEXTFREQMIN)
                    MFS(1, X).CSFreq = 0
                    MFS(2, X).bfreq = MFBGROUNDFREQMIN_CS + Rnd * (MFBGROUNDFREQMAX_CS - MFBGROUNDFREQMIN_CS)
                Case 10 'context 2
                    MFS(2, X).bfreq = MFCONTEXTFREQMIN2 + Rnd * (MFCONTEXTFREQMAX2 - MFCONTEXTFREQMIN2)
                    MFS(2, X).CSFreq = 0
                    MFS(1, X).bfreq = MFBGROUNDFREQMIN_CS + Rnd * (MFBGROUNDFREQMAX_CS - MFBGROUNDFREQMIN_CS)
                Case 0
                    MFS(i, X).bfreq = MFBGROUNDFREQMIN + Rnd * (MFBGROUNDFREQMAX - MFBGROUNDFREQMIN)
                    MFS(i, X).CSFreq = 0
            End Select
        Next i
        
        For i = 1 To 2
            'If MFS(i, x).CStype > 8 Then MFS(i, x).CStype = 0
            MFS(i, X).bfreq = MFS(i, X).bfreq * (Time_step_size / 1000)
            MFS(i, X).CSFreq = MFS(i, X).CSFreq * (Time_step_size / 1000)
            MFS(i, X).Thr = 1
        Next i
        
     Next X
End Sub


Public Function ToLeft(X As Integer)
    If (X + 29) Mod 30 = 0 Then
        ToLeft = X + 29
    Else
        ToLeft = X - 1
    End If
End Function

Public Function ToRight(X As Integer)
    If X Mod 30 = 0 Then
        ToRight = X - 29
    Else
        ToRight = X + 1
    End If
End Function
Public Function ToTop(X As Integer)
    If X < 31 Then
        ToTop = X + 870
    Else
        ToTop = X - 30
    End If
End Function
Public Function ToBottom(X As Integer)
    If X > 870 Then
        ToBottom = X - 870
    Else
        ToBottom = X + 30
    End If
End Function


