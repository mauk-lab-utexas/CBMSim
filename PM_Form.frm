VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PM_Form 
   BackColor       =   &H00000000&
   Caption         =   "Plasticty Monitor II"
   ClientHeight    =   7905
   ClientLeft      =   4785
   ClientTop       =   555
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1272
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line2 
      X1              =   1016
      X2              =   1096
      Y1              =   224
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   1128
      X2              =   1184
      Y1              =   216
      Y2              =   120
   End
   Begin VB.Menu SendtoExcelMenu 
      Caption         =   "SendtoExcel"
   End
End
Attribute VB_Name = "PM_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
Dim i As Integer
Dim p As Integer
Dim n As Integer
    If KeyAscii > 48 And KeyAscii < 57 Then
        PM_Form.ScaleTop = 25
        PM_Form.ScaleHeight = -25
        PM_Form.ScaleLeft = -500
        PM_Form.ScaleWidth = 1500
        PM_Form.Cls
        PM_Form.FillColor = RGB(100, 100, 255)
        PM_Form.FillStyle = 0
        
        For i = 1 To PCNUMBER
            PM_Form.Circle (-450, i), 10, RGB(100, 100, 255)
        Next i
        
        PM_Form.FillColor = vbWhite
        For i = 1 To NCNUMBER
            PM_Form.Circle (-350, -1 + i * 3), 12, vbWhite
        Next i
        
        PM_Form.FillColor = vbRed
        For i = 1 To NumCF
            PM_Form.Circle (-250, -2.5 + i * 6), 12, vbRed
        Next i
        
        n = KeyAscii - 48
            For p = 1 To PCNCSYNUMBER
                PM_Form.Line (-360, -1 + n * 1)-(-440, Nc(n).PCsyn(p)), RGB(100, 100, 255)
            Next p
            
        p = KeyAscii - 48
        For n = 1 To NumCF
       
            If CF(n).NucInput(p) = 1 Then
                 PM_Form.Line (-340, -1 + p * 3)-(-260, -2.5 + n * 6), vbWhite
            End If
       
        Next n
    
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
PM_Form.ScaleTop = 25
PM_Form.ScaleHeight = -25
PM_Form.ScaleLeft = -500
PM_Form.ScaleWidth = 1500
DoEvents



End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim i As Integer
Dim p As Integer
Dim n As Integer
    PM_Form.ScaleTop = 25
    PM_Form.ScaleHeight = -25
    PM_Form.ScaleLeft = -500
    PM_Form.ScaleWidth = 1500
    PM_Form.Cls
    PM_Form.FillColor = RGB(100, 100, 255)
    PM_Form.FillStyle = 0
    
    For i = 1 To PCNUMBER
        PM_Form.Circle (-450, i), 10, RGB(100, 100, 255)
    Next i
    
    PM_Form.FillColor = vbWhite
    For i = 1 To NCNUMBER
        PM_Form.Circle (-350, -1 + i * 3), 12, vbWhite
    Next i
    
    PM_Form.FillColor = vbRed
    For i = 1 To NumCF
        PM_Form.Circle (-250, -2.5 + i * 6), 12, vbRed
    Next i
    
    For n = 1 To NCNUMBER
        For p = 1 To PCNCSYNUMBER
            PM_Form.Line (-360, -1 + n * 3)-(-440, Nc(n).PCsyn(p)), RGB(100, 100, 255)
        Next p
    Next n
    
    For n = 1 To NumCF
        For p = 1 To 8
            If CF(n).NucInput(p) = 1 Then
                 PM_Form.Line (-340, -1 + p * 3)-(-260, -2.5 + n * 6), vbWhite
            End If
        Next p
    Next n

End Sub

Private Sub SendtoExcelMenu_Click()
    ShowWeights
End Sub
