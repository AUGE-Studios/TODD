Attribute VB_Name = "modCore"
Option Explicit

'colors
Public Enum tdColorConstant
        tdLightGreen = 12648384
        tdSand = 12648447
        tdLightRed = 12632319
        tdDarkGreen = 32768
        tdDarkRed = 192
        tdDarkBlue = 12936533
        tdStandard = -2147483633
        tdBlack = 0
End Enum

Sub Main()
Load frmMixer
frmMixer.Show
End Sub
