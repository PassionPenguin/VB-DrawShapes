VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19110
   FillColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11697.18
   ScaleMode       =   0  'User
   ScaleWidth      =   28609.12
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


‘ 项目主题 MAIN CODE
Dim aShape(10#) As Shape
Dim curShape As Integer

Public Function CreateShape(ByVal x, ByVal y, ByVal a, ByVal b, ByVal t As Integer, Optional ByVal bc, Optional ByVal fc As String = vbRed)
With aShape(curShape)
.Width = a
.Height = b
.Left = x
.Top = y
.BorderColor = bc
.FillColor = fc
.Shape = t
.FillStyle = 0
.Visible = True
End With
End Function

Private Sub form_load()
curShape = 0
For i = 0 To 10# - 1
Set aShape(i) = Controls.Add("VB.Shape", Replace("Shape" + Str(i), " ", ""))
Next
End Sub

Private Sub Form_DblClick()
Dim x, y, a, shapetype As Integer
a = 2500
b = 2500
x = 3000
y = 3000
shapetype = 3

CreateShape x, y, a, b, shapetype, vbBlue, vbGreen


'program ends with code 0
End Sub


Rem &copy Penguin, Open source under Mozilla Public License Version 2.0
Rem Github Project REF: https://github.com/PassionPenguin/VB-DrawCircles
Rem POWERD BY GLACIER ELEMENT, a Penguin's Proj
