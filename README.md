![GLACIER](/GlacierElement.png)

# VB-DrawShape

## Intro

BASE ON Visual Basic 6.0(VB 1998)

This func can help you draw shapes easily and freely (inifinity!)

## Features
a. Random colors in RGB color space. (via changing some codes)

-copy the code below to Function CreateShape(Line 32 to 33)

    .BorderColor = RGB(Rnd*255,Rnd*255,Rnd*255)
	.FillColor = RGB(Rnd*255,Rnd*255,Rnd*255)

b. auto draw shapes (in loop) via only changing some code and you can decide how many Shapes you need.

c. We provided a array contains 10# Shape!!!

d. changeable code you know...

## Usage
    Public Function CreateShape(ByVal x, ByVal y, ByVal a, ByVal b, ByVal t As Integer, Optional ByVal bc, Optional ByVal fc As String = vbRed)

eg. 
    CreateShape x, y, a, b, shapetype, vbBlue, vbGreen

©Penguin, Open Source Under Moz Via GitHub.
