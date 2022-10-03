VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form TrialsForm 
   BackColor       =   &H00400000&
   Caption         =   "Trials Builder"
   ClientHeight    =   5400
   ClientLeft      =   7305
   ClientTop       =   2115
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   4110
   Begin VB.CommandButton Command1 
      Caption         =   "Save As"
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   61
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   0
      Left            =   1455
      TabIndex        =   36
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   35
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   34
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   33
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   29
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   27
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   26
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   24
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   20
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   6
      Left            =   2160
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   7
      Left            =   1440
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   7
      Left            =   2160
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   9
      Left            =   1440
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cancel_buttons 
      Caption         =   "Cancel"
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Change_button 
      Caption         =   "Change"
      Height          =   255
      Index           =   8
      Left            =   1440
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   59
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   58
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Index           =   4
      Left            =   1560
      TabIndex        =   57
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   56
      Top             =   4440
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.str | *.str"
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS3 duration"
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   45
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS3 onset"
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   44
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS1 duration"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   43
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS2 duration"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   42
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CheckBox US1 
      BackColor       =   &H00000000&
      Caption         =   "US 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   3600
      Width           =   735
   End
   Begin VB.CheckBox CS1 
      BackColor       =   &H00000000&
      Caption         =   "CS 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   240
      Width           =   735
   End
   Begin VB.CheckBox CS2 
      BackColor       =   &H00000000&
      Caption         =   "CS 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS1 onset"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   38
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   37
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   32
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS2 onset"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   31
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   2
      Left            =   840
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   3
      Left            =   840
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   4
      Left            =   840
      TabIndex        =   22
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   5
      Left            =   840
      TabIndex        =   21
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   6
      Left            =   840
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   7
      Left            =   840
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS4 onset"
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   16
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change CS4 duration"
      Height          =   255
      Index           =   7
      Left            =   1440
      TabIndex        =   15
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CheckBox CS3 
      BackColor       =   &H00000000&
      Caption         =   "CS 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   735
   End
   Begin VB.CheckBox CS4 
      BackColor       =   &H00000000&
      Caption         =   "CS 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change US2 onset"
      Height          =   255
      Index           =   9
      Left            =   1440
      TabIndex        =   8
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   9
      Left            =   840
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox US2 
      BackColor       =   &H00000000&
      Caption         =   "US 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Change_stimuli 
      Caption         =   "Change US1 onset"
      Height          =   255
      Index           =   8
      Left            =   1440
      TabIndex        =   5
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox change_text 
      Height          =   285
      Index           =   8
      Left            =   840
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   60
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "1000"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   55
      Top             =   240
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "550"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   54
      Top             =   600
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "1000"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   53
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "500"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   52
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "1000"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   51
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "500"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   50
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "500"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   840
      TabIndex        =   49
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "1000"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   840
      TabIndex        =   48
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "1500"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   8
      Left            =   840
      TabIndex        =   47
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label OnsetLabels 
      BackColor       =   &H00000000&
      Caption         =   "1500"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   9
      Left            =   840
      TabIndex        =   46
      Top             =   3960
      Width           =   495
   End
End
Attribute VB_Name = "TrialsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_buttons_Click(Index As Integer)
    change_text(Index).Visible = False
    Change_button(Index).Visible = False
    Cancel_buttons(Index).Visible = False
End Sub

Private Sub Change_button_Click(Index As Integer)
If Int(change_text(Index).Text) > 0 And Int(change_text(Index).Text) < 5000 Then
        OnsetLabels(Index).Caption = change_text(Index).Text
        Select Case Index
        Case 0
            cs_onset(1) = Int(OnsetLabels(0))
        Case 1
            cs_duration(1) = Int(OnsetLabels(1))
        Case 2
            cs_onset(2) = Int(OnsetLabels(2))
        Case 3
            cs_duration(2) = Int(OnsetLabels(3))
        Case 4
            cs_onset(3) = Int(OnsetLabels(4))
        Case 5
            cs_duration(3) = Int(OnsetLabels(5))
        Case 6
            cs_onset(4) = Int(OnsetLabels(6))
        Case 7
            cs_duration(4) = Int(OnsetLabels(7))
        Case 8
            US_onset(1) = Int(OnsetLabels(8))
        Case 9
            US_onset(2) = Int(OnsetLabels(9))
        End Select
    End If
    change_text(Index).Visible = False
    Change_button(Index).Visible = False
    Cancel_buttons(Index).Visible = False
End Sub

Private Sub Change_stimuli_Click(Index As Integer)
    change_text(Index).Text = OnsetLabels(Index).Caption
    change_text(Index).Visible = True
    Change_button(Index).Visible = True
    Cancel_buttons(Index).Visible = True
End Sub

Private Sub Command1_Click(Index As Integer)
Dim TrialInfo As TrialStruct
Dim i As Integer

    Select Case Index
        Case 0
            CommonDialog1.filename = ""
            CommonDialog1.ShowOpen
            If CommonDialog1.filename <> "" Then
                Label1.Caption = CommonDialog1.filename
                Close #19
                Open CommonDialog1.filename For Binary As #19
                Get 19, , TrialInfo
                CS1.Value = TrialInfo.CSon(1)
                CS2.Value = TrialInfo.CSon(2)
                CS3.Value = TrialInfo.CSon(3)
                CS4.Value = TrialInfo.CSon(4)
                
                US1.Value = TrialInfo.USon(1)
                US2.Value = TrialInfo.USon(2)
                
                OnsetLabels(0) = Str(TrialInfo.CSonset(1))
                OnsetLabels(2) = Str(TrialInfo.CSonset(2))
                OnsetLabels(4) = Str(TrialInfo.CSonset(3))
                OnsetLabels(6) = Str(TrialInfo.CSonset(4))
                
                OnsetLabels(1) = Str(TrialInfo.CSduration(1))
                OnsetLabels(3) = Str(TrialInfo.CSduration(2))
                OnsetLabels(5) = Str(TrialInfo.CSduration(3))
                OnsetLabels(7) = Str(TrialInfo.CSduration(4))
                
                OnsetLabels(8) = Str(TrialInfo.USonset(1))
                OnsetLabels(9) = Str(TrialInfo.USonset(2))
                
                Close 19
            End If
        Case 1
            CommonDialog1.filename = ""
            CommonDialog1.ShowSave
            If CommonDialog1.filename <> "" Then
                Label1.Caption = CommonDialog1.filename
                Close #19
                Open CommonDialog1.filename For Binary As #19
                TrialInfo.CSon(1) = CS1.Value
                TrialInfo.CSon(2) = CS2.Value
                TrialInfo.CSon(3) = CS3.Value
                TrialInfo.CSon(4) = CS4.Value
                
                TrialInfo.USon(1) = US1.Value
                TrialInfo.USon(2) = US2.Value
                
                TrialInfo.CSonset(1) = Val(OnsetLabels(0))
                TrialInfo.CSonset(2) = Val(OnsetLabels(2))
                TrialInfo.CSonset(3) = Val(OnsetLabels(4))
                TrialInfo.CSonset(4) = Val(OnsetLabels(6))
                
                TrialInfo.CSduration(1) = Val(OnsetLabels(1))
                TrialInfo.CSduration(2) = Val(OnsetLabels(3))
                TrialInfo.CSduration(3) = Val(OnsetLabels(5))
                TrialInfo.CSduration(4) = Val(OnsetLabels(7))
                
                TrialInfo.USonset(1) = Val(OnsetLabels(8))
                TrialInfo.USonset(2) = Val(OnsetLabels(9))
                
                Put 19, , TrialInfo
                Close 19
            End If
        Case 2
            TrialsForm.Visible = False
        Case 3
            For i = 0 To 9
                change_text(i).Text = ""
                OnsetLabels(i).Caption = ""
            Next i
            CS1.Value = 0
            CS2.Value = 0
            CS3.Value = 0
            CS4.Value = 0
            US1.Value = 0
            US2.Value = 0
            Label1.Caption = ""
        Case 4
            If Label1.Caption <> "" Then
                Close #19
                Open Label1.Caption For Binary As #19
                TrialInfo.CSon(1) = CS1.Value
                TrialInfo.CSon(2) = CS2.Value
                TrialInfo.CSon(3) = CS3.Value
                TrialInfo.CSon(4) = CS4.Value
                
                TrialInfo.USon(1) = US1.Value
                TrialInfo.USon(2) = US2.Value
                
                TrialInfo.CSonset(1) = Val(OnsetLabels(0))
                TrialInfo.CSonset(2) = Val(OnsetLabels(2))
                TrialInfo.CSonset(3) = Val(OnsetLabels(4))
                TrialInfo.CSonset(4) = Val(OnsetLabels(6))
                
                TrialInfo.CSduration(1) = Val(OnsetLabels(1))
                TrialInfo.CSduration(2) = Val(OnsetLabels(3))
                TrialInfo.CSduration(3) = Val(OnsetLabels(5))
                TrialInfo.CSduration(4) = Val(OnsetLabels(7))
                
                TrialInfo.USonset(1) = Val(OnsetLabels(8))
                TrialInfo.USonset(2) = Val(OnsetLabels(9))
                
                Put 19, , TrialInfo
                Close 19
            End If
    End Select
End Sub

