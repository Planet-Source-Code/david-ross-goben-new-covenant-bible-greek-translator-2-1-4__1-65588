Attribute VB_Name = "modToRGB"
Option Explicit
'~modToRGB.bas;
'Convert a RGB color value back to individual colors
'*******************************************************************************
' modToRGB - Convert a RGB color value back to individual colors. The intrinsic
'            RGB function provides a convenient tool for converting separate Red,
'            Green, and Blue colors into a Long value, but there is no way to
'            take a long value and break these elements back out.
'
'            The ToRGB() function returns these values as separate integer
'            value in a RGB_Type variable.
'
'            The GetSystemColor() function returns a long value that holds
'            actual color values as defined for the user's preferences for intrisic
'            desktop colors, such as vbMenuText, vbScrollBars, vbActiveTitleBar, etc.
'            This function is also used by ToRGB() if you supply ToRGB() with one
'            of these intrinsic colors
'EXAMPLE:
'  Dim MyRGB As RGB_Type, MyColor As Long
'  MyColor = RGB(64, 128, 255) 'build the test color
'  MyRGB = ToRGB(MyColor)      'get color breakdown from the test color
'  MsgBox "Color breakdown for the test color is:" & vbCrLf & _
'              "Red  : " & MyRGB.rgbRed & vbCrLf & _
'              "Green: " & MyRGB.rgbGreen & vbCrLf & _
'              "Blue : " & MyRGB.rgbBlue
'  MyRGB = ToRGB(vbActiveTitleBar) 'get color breakdown from vbActiveTitleBar
'  MsgBox "Color breakdown for user's vbActiveTitleBar is:" & vbCrLf & _
'              "Red  : " & MyRGB.rgbRed & vbCrLf & _
'              "Green: " & MyRGB.rgbGreen & vbCrLf & _
'              "Blue : " & MyRGB.rgbBlue
'
'NOTE: Be aware that intrinsic desktop color values, such as vbActiveTitleBar
'      and vbDesktop, are special color values in the range of 0 through 31, with
'      &H10000000 added to them. ToRGB() will detect this and return the colors
'      assigned for these values as defined by the user.
'
'Predefined VB intrinsic color values are:
'      vbMenuText             vbScrollBars            vbWindowBackground
'      vbWindowFrame          vbActiveBorder          vbActiveTitleBar
'      vbTitleBarText         vbApplicationWorkspace  vbButtonFace
'      vb3DHightlight         vb3DDKShadow            vbButtonText
'      vbDesktop              vbGrayText              vbHighlight
'      vbHightlightText       vbInacntiveBorder       vbInactiveTitleBar
'      vbInanctiveCaptionText vbMenuBar
'*******************************************************************************
Public Type RGB_Type
  rgbBlue As Integer
  rgbGreen As Integer
  rgbRed As Integer
End Type

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'*******************************************************************************
' Function Name     : ToRGB
' Purpose           : Break down a color value to its individual color constants
'*******************************************************************************
Public Function ToRGB(ClrVal As Long) As RGB_Type
  Dim Clr As String
  
  If ClrVal And &H10000000 Then               'intrinsic color
    Clr = Right$("000000" & Hex$(GetSystemColor(ClrVal)), 6)
  Else                                        'normal color value
    Clr = Right$("000000" & Hex$(ClrVal), 6)  'get 6 character string
  End If
  With ToRGB
    .rgbRed = Val("&h" & Right$(Clr, 2))      'extract red
    .rgbGreen = Val("&h" & Mid$(Clr, 3, 2))   'extract green
    .rgbBlue = Val("&h" & Left$(Clr, 2))      'extract blue
  End With
End Function

'*******************************************************************************
' Function Name     : GetSystemColor
' Purpose           : Get the color value for an intrinsic color
'*******************************************************************************
Public Function GetSystemColor(SysColor As Long) As Long
  GetSystemColor = GetSysColor(SysColor & &HFFFFFF) 'strip &h10000000 flag
End Function
