Attribute VB_Name = "Connectivity"
Option Explicit
Public Const GlY = 30
Public Const GlX = 30
Public Const GrDendSpanY = 50
Public Const GrDendSpanX = 50
Public Const GoDendSpanY = 10       '20
Public Const GoDendSpanX = 80       '60
Public Const GrGoDendSpanY = 10     '20
Public Const GrGoDendSpanX = 10     '20
Public Const GoGlDendSpanY = 20     '50
Public Const GoGlDendSpanX = 20     '50

Public Const NUMBERCS = 4

Public GrGlScaleX As Single
Public GrGoScaleX As Single
Public GrGlScaleY As Single
Public GrGoScaleY As Single

Public Type glom
  MF As Integer
End Type
Public Gl(GlX * GlY) As glom
