VERSION 5.00
Begin VB.Form PC_form 
   BackColor       =   &H00000000&
   Caption         =   "PC activity"
   ClientHeight    =   11640
   ClientLeft      =   7980
   ClientTop       =   375
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   ScaleHeight     =   776
   ScaleMode       =   0  'User
   ScaleWidth      =   5002
End
Attribute VB_Name = "PC_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        DoPCAudio = 0
        DoNCAudio = 0
        DoCFAudio = 0
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim WhichCell As Integer
    If Button = 1 Then
        DoPCAudio = 0
        DoNCAudio = 0
        If y < 456 Then
            WhichCell = Int((y - 1) / 19) + 1
            DoPCAudio = WhichCell
           
        ElseIf y < 610 Then
            WhichCell = Int((y) / 19) + 1
            DoNCAudio = WhichCell - 24
            
        End If
    End If
End Sub
