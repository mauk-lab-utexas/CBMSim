Attribute VB_Name = "HomeoStatic"
Public DoPurkClassic As Integer
Public DoPurkPre As Integer
Public DoNucPre As Integer
Public DoPurkSynapticScaling As Integer
Public DoPurkIntrinsic As Integer

Public Sub PurkSynapticScaling()
Dim X As Integer
Dim y As Integer
Dim y1 As Integer
Dim syn As Integer

    'Average = 0
    For X = 1 To PCNUMBER
        y = 1 + (X - 1) * PFPCSYNUMBER
        y1 = y + PFPCSYNUMBER - 1
        
        If PurkActivity(X) / 5# < PCHomeoValue(X) Then
            For syn = y To y1
                grWeight(syn) = grWeight(syn) * 1.003
                If grWeight(syn) > 1 Then
                    grWeight(syn) = 1
                End If
            Next syn
        ElseIf PurkActivity(X) / 5# > PCHomeoValue(X) Then
            For syn = y To y1
                grWeight(syn) = grWeight(syn) / 1.003
            Next syn
        End If
        
    Next X
    
End Sub

Public Sub PurkPresynapticPlasticity()
Dim X As Integer
Dim y As Integer
Dim HomeoValue As Single

'    For X = 1 To NCNUMBER
'        HomeoValue = PCHomeoValue(X)
'        HomeoValue = ((1000 - NCHomeoValue(X)) * 0.000001) / PCHomeoValue(X)
'        For y = 1 To PCNCSYNUMBER
'            gPURKtoNUCLEUS(Nc(X).PCsyn(y), X) = gPURKtoNUCLEUS(Nc(X).PCsyn(y), X) + (Pc(Nc(X).PCsyn(y)).act * HomeoValue) - ((1 - Pc(Nc(X).PCsyn(y)).act) * 0.000001)           '*****************  PURKINJE PRE
'            If gPURKtoNUCLEUS(Nc(X).PCsyn(y), X) > 0.6 Then
'                gPURKtoNUCLEUS(Nc(X).PCsyn(y), X) = 0.6
'            ElseIf gPURKtoNUCLEUS(Nc(X).PCsyn(y), X) < 0.1 Then
'                gPURKtoNUCLEUS(Nc(X).PCsyn(y), X) = 0.1
'            End If
'        Next y
'    Next X

     For X = 1 To PCNUMBER
     
        If PurkActivity(X) / 5# < PCHomeoValue(X) Then
            For y = 1 To NCNUMBER
                gPURKtoNUCLEUS(X, y) = gPURKtoNUCLEUS(X, y) / 1.03
            Next y
            
        ElseIf PurkActivity(X) / 5# > PCHomeoValue(X) Then
            For y = 1 To NCNUMBER
                gPURKtoNUCLEUS(X, y) = gPURKtoNUCLEUS(X, y) * 1.03
            Next y
        End If
        
        For y = 1 To NCNUMBER
            If gPURKtoNUCLEUS(X, y) > 0.6 Then
                gPURKtoNUCLEUS(X, y) = 0.6
            ElseIf gPURKtoNUCLEUS(X, y) < 0.05 Then
                gPURKtoNUCLEUS(X, y) = 0.05
            End If
        Next y
        
    Next X

End Sub

Public Sub NucleusPresynapticPlasticity()
Dim X, y As Integer
    For X = 1 To NCNUMBER
        For y = 1 To NumCF
            Nc(X).gNUCtoCF(y) = Nc(X).gNUCtoCF(y) + (0.000003 * Nc(X).act) - ((1 - Nc(X).act) * 0.00000003)                            'NUC PRE ***********************************
            If Nc(X).gNUCtoCF(y) > 1 Then
                Nc(X).gNUCtoCF(y) = 1
            ElseIf Nc(X).gNUCtoCF(y) < 0 Then
                Nc(X).gNUCtoCF(y) = 0
            End If
        Next y
    Next X
End Sub

Public Sub PurkIntrinsicPlasticity()
Dim X As Integer
Dim HV As Single
    
    For X = 1 To PCNUMBER
        HV = PCHomeoValue(X)
        HV = (1000 - HV) / HV
        If Pc(X).act = 1 Then
            Pc(X).ThrBase = Pc(X).ThrBase + (HV * 0.00002)
        Else
            Pc(X).ThrBase = Pc(X).ThrBase - (0.00002)
        End If
        'Pc(x).ThrBase = Pc(x).ThrBase - 0.00001
    Next X
End Sub
