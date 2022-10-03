VERSION 5.00
Begin VB.Form ActivityHistoryForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Activity History"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   10470
   ClientWidth     =   16575
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   0  'User
   ScaleWidth      =   1105
   Begin VB.Menu PurkinjeCellsMenu 
      Caption         =   "Purkinje Cells"
      Begin VB.Menu PCMenu 
         Caption         =   "ShowAll"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu PCMenu 
         Caption         =   "ShowAverage"
         Index           =   2
      End
      Begin VB.Menu SelectMenu 
         Caption         =   "SelectToView"
         Begin VB.Menu SMenu 
            Caption         =   "1"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu SMenu 
            Caption         =   "2"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu SMenu 
            Caption         =   "3"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu SMenu 
            Caption         =   "4"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu SMenu 
            Caption         =   "5"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu SMenu 
            Caption         =   "6"
            Checked         =   -1  'True
            Index           =   6
         End
         Begin VB.Menu SMenu 
            Caption         =   "7"
            Checked         =   -1  'True
            Index           =   7
         End
         Begin VB.Menu SMenu 
            Caption         =   "8"
            Checked         =   -1  'True
            Index           =   8
         End
         Begin VB.Menu SMenu 
            Caption         =   "9"
            Checked         =   -1  'True
            Index           =   9
         End
         Begin VB.Menu SMenu 
            Caption         =   "10"
            Checked         =   -1  'True
            Index           =   10
         End
         Begin VB.Menu SMenu 
            Caption         =   "11"
            Checked         =   -1  'True
            Index           =   11
         End
         Begin VB.Menu SMenu 
            Caption         =   "12"
            Checked         =   -1  'True
            Index           =   12
         End
         Begin VB.Menu SMenu 
            Caption         =   "13"
            Checked         =   -1  'True
            Index           =   13
         End
         Begin VB.Menu SMenu 
            Caption         =   "14"
            Checked         =   -1  'True
            Index           =   14
         End
         Begin VB.Menu SMenu 
            Caption         =   "15"
            Checked         =   -1  'True
            Index           =   15
         End
         Begin VB.Menu SMenu 
            Caption         =   "16"
            Checked         =   -1  'True
            Index           =   16
         End
         Begin VB.Menu SMenu 
            Caption         =   "17"
            Checked         =   -1  'True
            Index           =   17
         End
         Begin VB.Menu SMenu 
            Caption         =   "18"
            Checked         =   -1  'True
            Index           =   18
         End
         Begin VB.Menu SMenu 
            Caption         =   "19"
            Checked         =   -1  'True
            Index           =   19
         End
         Begin VB.Menu SMenu 
            Caption         =   "20"
            Checked         =   -1  'True
            Index           =   20
         End
      End
   End
End
Attribute VB_Name = "ActivityHistoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PCMenu_Click(Index As Integer)
Dim i As Integer

    If Index = 1 Then
        If PCMenu(1).Checked = True Then
            PCMenu(1).Checked = False
        Else
            PCMenu(1).Checked = True
            For i = 1 To 20
                ActivityHistoryForm.SMenu(i).Checked = True
            Next i
        End If
    ElseIf Index = 2 Then
        If PCMenu(2).Checked = True Then
            PCMenu(2).Checked = False
        Else
            PCMenu(2).Checked = True
        End If
    End If
End Sub

Private Sub SMenu_Click(Index As Integer)
If SMenu(Index).Checked = True Then SMenu(Index).Checked = False Else SMenu(Index).Checked = True
End Sub
