Attribute VB_Name = "modGetScreenColorCount"
Option Explicit
'~modGetScreenColorCount.bas;
'Retrieves screen color size flag
'*************************************************
' modGetScreenColorCount:
'
' The GetScreenColorCount() function retrieves screen
'     color size flag (and optional count).
'
' Intger Function Result Returns:
'  -1: Unknown
'   0: 1 bit color
'   1: 2 bit color
'   2: 4 bit color
'   3: 8 bit color
'   4: 16 bit high color
'   5: 24 bit true color
'   6: 32 bit true color
'
' Optional Double variable Count will contain the actual color
' count (0 - 4,294,967,296)
'
' The GetScreenColorBits() function retrieves the number
'     of bits needed to support the current color resolution.
' Results:
'   1 bit color        (2^1  = 2 colors)
'   2 bit color        (2^2  = 4 colors)
'   4 bit color        (2^4  = 16 colors)
'   8 bit color        (2^8  = 256 colors)
'   16 bit high color  (2^16 = 65536 colors)
'   24 bit true color  (2^24 = 16,777,216 colors)
'   32 bit true color  (2^32 = 4,294,967,296 colors)
'
'*************************************************

'*************************************************
' API calls
'*************************************************
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Const COLORRES = 108
Private Const BITSPIXEL = 12

'*************************************************
' GetScreenColorCount(): Retrieve screen color size flag (and optional count)
'*************************************************
Public Function GetScreenColorCount(FormhDC As Long, Optional Count As Double) As Integer
  Dim flg As Integer, caps As Long
  caps = GetDeviceCaps(FormhDC, COLORRES)       'get the bits per pixel count
  Select Case caps
    Case 1, 2, 4, 8, 15, 16, 24, 32
    Case Else
      caps = GetDeviceCaps(FormhDC, BITSPIXEL)  'get the bits per pixel count
  End Select
  
  Select Case caps
    Case 1
      flg = 0                  '2
    Case 2
      flg = 1                  '4
    Case 4
      flg = 2                  '16
    Case 8
      flg = 3                  '256
    Case 15
      flg = 4                  '32768          (high color)
    Case 16
      flg = 5                  '65536          (high color)
    Case 24
      flg = 6                  '16,777,216     (true color)
    Case 32
      flg = 7                  '4,294,967,296  (true color)
    Case Else
      GetScreenColorCount = -1 'unknown
  End Select
'
' if Count variable present, then stuff with color count
'
  If Not IsMissing(Count) Then
    Select Case caps
        Case 1, 2, 4, 8, 15, 16, 24, 32
          Count = 2# ^ CDbl(caps) 'set value to power of 2
        Case Else
          Count = 0#              'set null count if unknown
    End Select
  End If
  GetScreenColorCount = flg       'set flag
  
End Function

'*************************************************
' GetScreenColorBits(): Retrieve the number of bits needed to support
'                       The current color resolution (4=16, 8=256, etc)
'*************************************************
Public Function GetScreenColorBits(FormhDC As Long) As Integer
  GetScreenColorBits = GetDeviceCaps(FormhDC, BITSPIXEL)  'get the bits per pixel count
End Function

