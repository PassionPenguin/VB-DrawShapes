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
Public Function CreateRings(ByVal x As Integer, ByVal y As Integer, ByVal r As Integer, ByVal num As Integer, ByVal colors As String)
For k = 1 To num
'for different cols
x = x + 3000
Circle (x, y), r, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Next
End Function

Private Sub Form_DblClick()
Dim colors As Variant, x, y, r, num As Integer
r = 2500
DrawWidth = 10


'first row
num = 5
x = 0
y = 3000
CreateRings x, y, r, num, colors
'second row
num = 4
x = 1500
y = 4500
CreateRings x, y, r, num, colors
'third row
num = 3
x = 3000
y = 6000
CreateRings x, y, r, num, colors
'program ends with code 0
End Sub


Rem &copy Penguin, Open source under Mozilla Public License Version 2.0
Rem Github Project REF: https://github.com/PassionPenguin/VB-DrawCircles
Rem POWERD BY GLACIER ELEMENT, a Penguin's Proj
