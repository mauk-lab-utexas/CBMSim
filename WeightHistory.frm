VERSION 5.00
Begin VB.Form WeightHistory 
   BackColor       =   &H00000000&
   Caption         =   "Synaptic Weights over time"
   ClientHeight    =   4920
   ClientLeft      =   4290
   ClientTop       =   16980
   ClientWidth     =   21240
   LinkTopic       =   "Form1"
   ScaleHeight     =   328
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1416
End
Attribute VB_Name = "WeightHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    WeightHistory.ScaleHeight = -1
    WeightHistory.ScaleTop = 1
    WeightHistory.ScaleWidth = 1000
    
End Sub
