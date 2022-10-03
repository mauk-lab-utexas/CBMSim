VERSION 5.00
Begin VB.Form GWin 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   17490
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   17490
   ScaleWidth      =   11850
   Begin VB.CommandButton Command1 
      Caption         =   "Save this histogram to file"
      Height          =   615
      Index           =   3
      Left            =   8760
      TabIndex        =   4
      Top             =   16800
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Granule Latencies to onset"
      Height          =   615
      Index           =   2
      Left            =   2760
      TabIndex        =   3
      Top             =   16800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Granule weights "
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   16800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Two Layer Search to File"
      Height          =   615
      Index           =   1
      Left            =   5400
      TabIndex        =   1
      Top             =   16800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "GWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
Dim filename As String
Dim i, j, k As Integer
Dim max As Single
Dim MaxLat As Integer

Dim MF1(2, 4) As Integer
Dim GR1(2, 256) As Integer
Dim GO1(2, 256) As Integer

Dim GrData(1000, 4, 64) As Integer
Dim MFData(1000, 4, 4) As Integer

    Select Case Index
    Case 0
        filename = "GranuletoPurkWeights at" & Str(TrialCounter)
         Close 22
        Open filename For Output As #22
        
        For i = 1 To SYNUMBER
            Write #22, i, grWeight(i)
        Next i
        Close 22
    Case 1
    
        filename = "Histogram for gr " & Str(GeneologyCell)
        Close 22
        Open filename For Output As #22
        
        For i = 1 To 1000
            Write #22, GR_histo(GeneologyCell, i)
        Next i
        Close 22
        
        filename = "MF and Golgi Histograms for gr " & Str(GeneologyCell)
        Close 22
        Open filename For Output As #22
        
        For i = 1 To 1000
            For j = 1 To Gr(GeneologyCell).numdend
                Write #22, MF_histo(Gl(Gr(GeneologyCell).Gol(j)).MF, i),
            Next j
            
            For j = 1 To Gr(GeneologyCell).numdend
                Write #22, Go_histo(Gr(GeneologyCell).Gol(j), i),
            Next j
            Write #22,
        Next i
        
        Close #22
        
        For i = 1 To Gr(GeneologyCell).numdend  ' for each of the four Golgi cell synapses onto the granule cell
            For j = 1 To numGoGrDend  'for each of the 64 granule cells that synapse onto that Golgi cell
                For k = 1 To 1000
                    'gr_histo(Gol(Gr(GeneologyCell).preGr)) = gr_histo(Gol(Gr(GeneologyCell).preGr))
                    'Debug.Print "For gr #:"; GeneologyCell; "Dendrite "; i; "of Golgi "; Gr(GeneologyCell).Gol(i); " granule input "; Gol(Gr(GeneologyCell).Gol(i)).preGr(j)
                    GrData(k, i, j) = GR_histo(Gol(Gr(GeneologyCell).Gol(i)).preGr(j), k)
                Next k
            Next j
            
            For j = 1 To numGoGlDend   '4
                    For k = 1 To 1000
                        MFData(k, i, j) = MF_histo(Gl(Gol(Gr(GeneologyCell).Gol(i)).preGl(j)).MF, k)
                    Next k
            Next j
        Next i
        filename = "Gr GeneologyFor gr" & Str(GeneologyCell)
        Close 22
        Open filename For Output As #22
    
        For i = 1 To 1000
            For j = 1 To Gr(GeneologyCell).numdend
                For k = 1 To numGoGrDend
                    Write #22, GrData(i, j, k),
                Next k
            Next j
            Write #22,
        Next i
        Close 22
        filename = "MF GeneologyFor gr" & Str(GeneologyCell)
        Open filename For Output As #22
    
        For i = 1 To 1000
            For j = 1 To Gr(GeneologyCell).numdend
                For k = 1 To numGoGlDend
                    Write #22, MFData(i, j, k),
                Next k
            Next j
            Write #22,
        Next i
        Close 22
    Case 2
        filename = "GranuleLatenciestoPeak at" & Str(TrialCounter)
        Close 22
        Open filename For Output As #22
        For i = 1 To SYNUMBER
            max = 0
            For j = 1 To 1000
                If GR_histo(i, j) / HistoDivisor > max Then
                    max = GR_histo(i, j) / HistoDivisor
                    MaxLat = j
                End If
            Next j
            Write #22, i, max, MaxLat
        Next i
        Close 22
    Case 3
    
        filename = "SimHisto" & Str(GeneologyCell)
        Close 22
        Open filename For Output As #22
        
        For i = 1 To 1000
            Write #22, GR_histo(GeneologyCell, i)
        Next i
        Close 22
    
    End Select
   
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If PCImageX = 0 Then PCImageX = -6547 Else PCImageX = 0
'    GWin.Cls
'    GWin.PaintPicture PCImage, PCImageX, 0, 12000, 16000
End Sub

