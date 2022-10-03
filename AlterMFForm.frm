VERSION 5.00
Begin VB.Form AlterMFForm 
   BackColor       =   &H00000000&
   Caption         =   "Alter Mossy Fibers"
   ClientHeight    =   9330
   ClientLeft      =   750
   ClientTop       =   1365
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   7695
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   52
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   51
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   50
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   49
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   4
      Left            =   4320
      TabIndex        =   48
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   5
      Left            =   4320
      TabIndex        =   47
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   6
      Left            =   4320
      TabIndex        =   46
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   7
      Left            =   4320
      TabIndex        =   45
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   8
      Left            =   1560
      TabIndex        =   44
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   9
      Left            =   1560
      TabIndex        =   43
      Top             =   7680
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   10
      Left            =   1560
      TabIndex        =   42
      Top             =   8160
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   11
      Left            =   1560
      TabIndex        =   41
      Top             =   8640
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   12
      Left            =   2760
      TabIndex        =   40
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   13
      Left            =   2760
      TabIndex        =   39
      Top             =   7680
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   14
      Left            =   2760
      TabIndex        =   38
      Top             =   8160
      Width           =   615
   End
   Begin VB.TextBox MFText 
      Height          =   375
      Index           =   15
      Left            =   2760
      TabIndex        =   37
      Top             =   8640
      Width           =   615
   End
   Begin VB.CommandButton UpdateMFCommand 
      Caption         =   "UPDATE MFs"
      Height          =   495
      Left            =   4200
      TabIndex        =   36
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton CancelMFCommand 
      Caption         =   "Cancel MF changes"
      Height          =   495
      Left            =   4200
      TabIndex        =   35
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Change"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   34
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox CSDegradeText 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   3960
      TabIndex        =   33
      Text            =   "0"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox CSDegradeText 
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   32
      Text            =   "0"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox CSDegradeText 
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   31
      Text            =   "0"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox CSDegradeText 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   30
      Text            =   "0"
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "View"
      Height          =   495
      Left            =   6360
      TabIndex        =   28
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Restore MFs"
      Height          =   495
      Left            =   6360
      TabIndex        =   27
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton DoneCommand 
      Caption         =   "Done"
      Height          =   495
      Left            =   6360
      TabIndex        =   26
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   25
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   24
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox CSActivityText 
      Height          =   285
      Index           =   7
      Left            =   3960
      TabIndex        =   23
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox CSActivityText 
      Height          =   285
      Index           =   6
      Left            =   3360
      TabIndex        =   22
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox CSActivityText 
      Height          =   285
      Index           =   5
      Left            =   2760
      TabIndex        =   21
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox CSActivityText 
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   20
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox CSActivityText 
      Height          =   285
      Index           =   3
      Left            =   3960
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox CSActivityText 
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox CSActivityText 
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox CSActivityText 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MF background frequencies (Hz)"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   72
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   71
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   70
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Mossy fiber context frequencies (Hz)"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   3120
      TabIndex        =   69
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Context 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   3240
      TabIndex        =   68
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   3480
      TabIndex        =   67
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   3480
      TabIndex        =   66
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Non-CS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   7
      Left            =   480
      TabIndex        =   65
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   8
      Left            =   720
      TabIndex        =   64
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   9
      Left            =   720
      TabIndex        =   63
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   10
      Left            =   480
      TabIndex        =   62
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Context 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   11
      Left            =   3240
      TabIndex        =   61
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minumum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   3480
      TabIndex        =   60
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   3480
      TabIndex        =   59
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CS 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   14
      Left            =   480
      TabIndex        =   58
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CS 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   15
      Left            =   480
      TabIndex        =   57
      Top             =   7800
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CS 3:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   16
      Left            =   480
      TabIndex        =   56
      Top             =   8280
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CS 4:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   17
      Left            =   480
      TabIndex        =   55
      Top             =   8760
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Phasic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   18
      Left            =   1560
      TabIndex        =   54
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tonic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Index           =   19
      Left            =   2760
      TabIndex        =   53
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Degrade CS to this # cells:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   0
      TabIndex        =   29
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of MFs:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   19
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   3960
      TabIndex        =   18
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   17
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   16
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   15
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   14
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   13
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   12
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   11
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Average activity:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CS4"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   4080
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CS3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CS2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CS1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make activity exactly:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make Average activity:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "AlterMFForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelMFCommand_Click()
    MFText(0).Text = MFBGROUNDFREQMIN
    MFText(1).Text = MFBGROUNDFREQMAX
    
    MFText(2).Text = MFBGROUNDFREQMIN_CS
    MFText(3).Text = MFBGROUNDFREQMAX_CS
    
    MFText(4).Text = MFCONTEXTFREQMIN
    MFText(5).Text = MFCONTEXTFREQMAX
    
    MFText(6).Text = MFCONTEXTFREQMIN2
    MFText(7).Text = MFCONTEXTFREQMAX2
    
    MFText(8).Text = MFPHASICFREQ_INCREMENT
    MFText(9).Text = MFPHASICFREQ_INCREMENT2
    MFText(10).Text = MFPHASICFREQ_INCREMENT3
    MFText(11).Text = MFPHASICFREQ_INCREMENT4
    
    MFText(12).Text = MFTONICFREQ_INCREMENT
    MFText(13).Text = MFTONICFREQ_INCREMENT2
    MFText(14).Text = MFTONICFREQ_INCREMENT3
    MFText(15).Text = MFTONICFREQ_INCREMENT4
    UpdateMFCommand.Enabled = False
    CancelMFCommand.Enabled = False
End Sub

Private Sub Command1_Click(Index As Integer)
Dim i As Integer

    Debug.Print MFTypetoChange
    If Index = 1 Then
        For i = 1 To MFNUMBER
            If MFS(1, i).CStype = MFTypetoChange Then
                
                MFS(1, i).CSFreq = Val(CSActivityText(MFTypetoChange + 3)) / 1000#
            End If
        Next i
    End If
    
    
    For i = 4 To 7
        CSActivityText(i).Text = ""
    Next i
    Command1(Index).Enabled = False
End Sub

Private Sub Command2_Click()
    AlterMFForm.Visible = False
End Sub

Private Sub Command3_Click()
Dim i As Integer
Dim j As Integer

    For i = 1 To 2
        For j = 1 To MFNUMBER
            MFS(i, j) = mfsBackup(i, j)
        Next j
    Next i
End Sub

Private Sub Command4_Click()
Dim i As Integer
Dim j As Integer
    For j = 1 To 4
        For i = 1 To MFNUMBER
                If MFS(1, i).CStype = j Then
                    Debug.Print j, i, MFS(1, i).bfreq, MFS(1, i).CSFreq
                End If
        Next i
        Debug.Print
    Next j
    Debug.Print
End Sub

Private Sub Command5_Click()
Dim i As Integer
Dim X As Integer
Dim done As Integer
Dim MFtoChange As Integer
Dim temp As Single
Dim TempX As Integer

    'Debug.Print MFTypetoChange, MFNumbertoDegrade
    For i = 1 To MFNumbertoDegrade
        done = 0
        
        While done = 0
            X = Rnd() * 600
            
            If X > 0 And X < MFNUMBER + 1 Then
                
                If MFS(1, X).CStype = MFTypetoChange Then
                    Debug.Print i, X
                    temp = MFS(1, X).CSFreq
                    MFS(1, X).CStype = 2000
                
                    TempX = X
                    
                    While temp <> 0
                        TempX = TempX + 1
                        Debug.Print "Tempx", TempX
                        If TempX > MFNUMBER + 1 Then TempX = 1
                        If MFS(1, TempX).CStype = 0 Then
                            MFS(1, TempX).CStype = 3000
                            MFS(1, TempX).CSFreq = temp
                            temp = 0
                        End If
                    Wend
                    
                    done = 1
                End If
            End If
        Wend
    Next i
    For i = 1 To MFNUMBER
        If MFS(1, i).CStype = 2000 Then MFS(1, i).CStype = 0
        If MFS(1, i).CStype = 3000 Then MFS(1, i).CStype = MFTypetoChange
    Next i
    Command5.Enabled = False
End Sub

Private Sub CSActivityText_Change(Index As Integer)

    If CSActivityText(Index).DataChanged = True Then
    
        If Index < 4 Then
            Command1(0).Enabled = True
        Else
            Command1(1).Enabled = True
        End If
    End If
    
    If Index < 4 Then
        MFTypetoChange = Index + 1
    Else
        MFTypetoChange = Index - 3
    End If
    
End Sub

Private Sub CSDegradeText_Change(Index As Integer)
    MFTypetoChange = Index + 1
    Command5.Enabled = True
    MFNumbertoDegrade = Val(Label2(Index + 4)) - Val(CSDegradeText(Index))
    Debug.Print MFTypetoChange, MFNumbertoDegrade
    
End Sub

Private Sub DoneCommand_Click()
    AlterMFForm.Visible = False
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim j As Integer
Dim CSFreq(4) As Single

Dim NumCS(4) As Single

    For i = 1 To 2
        For j = 1 To MFNUMBER
            mfsBackup(i, j) = MFS(i, j)
        Next j
    Next i
    
    For i = 1 To 4
        CSFreq(i) = 0
        NumCS(i) = 0
    Next i
    
    
    For j = 1 To 4
        For i = 1 To MFNUMBER
            If MFS(1, i).CStype = j Then
                CSFreq(j) = CSFreq(j) + (MFS(1, i).CSFreq * 1000)
                NumCS(j) = NumCS(j) + 1
            End If
        Next i
        
        If NumCS(j) <> 0 Then
            CSFreq(j) = CSFreq(j) / NumCS(j)
            i = Int(CSFreq(j))
            Label2(j - 1).Caption = i
        End If
        Label2(j + 3).Caption = NumCS(j)
        CSDegradeText(j - 1).Text = NumCS(j)
        Command5.Enabled = False
    Next j
    
     MFText(0).Text = MFBGROUNDFREQMIN
    MFText(1).Text = MFBGROUNDFREQMAX
    
    MFText(2).Text = MFBGROUNDFREQMIN_CS
    MFText(3).Text = MFBGROUNDFREQMAX_CS
    
    MFText(4).Text = MFCONTEXTFREQMIN
    MFText(5).Text = MFCONTEXTFREQMAX
    
    MFText(6).Text = MFCONTEXTFREQMIN2
    MFText(7).Text = MFCONTEXTFREQMAX2
    
    MFText(8).Text = MFPHASICFREQ_INCREMENT
    MFText(9).Text = MFPHASICFREQ_INCREMENT2
    MFText(10).Text = MFPHASICFREQ_INCREMENT3
    MFText(11).Text = MFPHASICFREQ_INCREMENT4
    
    MFText(12).Text = MFTONICFREQ_INCREMENT
    MFText(13).Text = MFTONICFREQ_INCREMENT2
    MFText(14).Text = MFTONICFREQ_INCREMENT3
    MFText(15).Text = MFTONICFREQ_INCREMENT4
    
    UpdateMFCommand.Enabled = False
    CancelMFCommand.Enabled = False
    
End Sub

Private Sub MFText_Change(Index As Integer)
    UpdateMFCommand.Enabled = True
    CancelMFCommand.Enabled = True
End Sub

Private Sub UpdateMFCommand_Click()
    MFBGROUNDFREQMIN = Val(MFText(0).Text)
    MFBGROUNDFREQMAX = Val(MFText(1).Text)
    
    MFBGROUNDFREQMIN_CS = Val(MFText(2).Text)
    MFBGROUNDFREQMAX_CS = Val(MFText(3).Text)
    
    MFCONTEXTFREQMIN = Val(MFText(4).Text)
    MFCONTEXTFREQMAX = Val(MFText(5).Text)
    
    MFCONTEXTFREQMIN2 = Val(MFText(6).Text)
    MFCONTEXTFREQMAX2 = Val(MFText(7).Text)
    
    MFPHASICFREQ_INCREMENT = Val(MFText(8).Text)
    MFPHASICFREQ_INCREMENT2 = Val(MFText(9).Text)
    MFPHASICFREQ_INCREMENT3 = Val(MFText(10).Text)
    MFPHASICFREQ_INCREMENT4 = Val(MFText(11).Text)
    
    MFTONICFREQ_INCREMENT = Val(MFText(12).Text)
    MFTONICFREQ_INCREMENT2 = Val(MFText(13).Text)
    MFTONICFREQ_INCREMENT3 = Val(MFText(14).Text)
    MFTONICFREQ_INCREMENT4 = Val(MFText(15).Text)
    
    AssignMFs
    
    UpdateMFCommand.Enabled = False
    CancelMFCommand.Enabled = False
End Sub
