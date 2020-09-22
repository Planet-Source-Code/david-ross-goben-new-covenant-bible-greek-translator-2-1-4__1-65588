Attribute VB_Name = "modGetStockObject"
Option Explicit
'~modGetStockObject.bas;
'Retrieve system stock object bitmaps and icons
'********************************************************************************
'modGetStockObject:
'  Retrieve system stock object bitmaps and icons
'  (these are graphical items stored internal to
'   windows). Load them to a provided PictureBox.
'   The PictureBox height and width will be adjusted
'   to accomodate the bitmap or icon.
'********************************************************************************

'********************************************************************************
'Constants, types, and API calls
'********************************************************************************
Private Const SRCCOPY = &HCC0020

Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Public Enum StockObjects
  OIC_SAMPLE = 32512
  OIC_HAND = 32513
  OIC_QUES = 32514
  OIC_BANG = 32515
  OIC_NOTE = 32516
  OIC_WINLOGO = 32517
  IDI_APPLICATION = 32512&
  IDI_HAND = 32513&
  IDI_QUESTION = 32514&
  IDI_EXCLAMATION = 32515&
  IDI_ASTERISK = 32516&
  IDI_WINLOGO = 32517&
  OBM_LFARROWI = 32734
  OBM_RGARROWI = 32735
  OBM_DNARROWI = 32736
  OBM_UPARROWI = 32737
  OBM_COMBO = 32738
  OBM_MNARROW = 32739
  OBM_LFARROWD = 32740
  OBM_RGARROWD = 32741
  OBM_DNARROWD = 32742
  OBM_UPARROWD = 32743
  OBM_RESTORED = 32744
  OBM_ZOOMD = 32745
  OBM_REDUCED = 32746
  OBM_RESTORE = 32747
  OBM_ZOOM = 32748
  OBM_REDUCE = 32749
  OBM_LFARROW = 32750
  OBM_RGARROW = 32751
  OBM_DNARROW = 32752
  OBM_UPARROW = 32753
  OBM_CLOSE = 32754
  OBM_BTNCORNERS = 32758
  OBM_CHECKBOXES = 32759
  OBM_CHECK = 32760
  OBM_BTSIZE = 32761
  OBM_SIZE = 32766
End Enum

Private Declare Function CreateCompatibleDC& Lib "gdi32" (ByVal hDC As Long)
Private Declare Function LoadBitmapBynum& Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Long)
Private Declare Function GetObjectAPI& Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any)
Private Declare Function SelectObject& Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long)
Private Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Private Declare Function DeleteDC& Lib "gdi32" (ByVal hDC As Long)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Private Declare Function LoadIconBynum& Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long)
Private Declare Function DrawIcon& Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long)

'********************************************************************************
'GetStockObject():
'  Retrieve system stock object bitmaps and icons
'  (these are graphical items stored internal to
'   windows). Load them to a provided PictureBox.
'   The PictureBox height and width will be adjusted
'   to accomodate the bitmap or icon.
'
'Available stock objects are:
'(MessageBox)ICONS:
'  IDI_HAND        or OIC_HAND
'  IDI_APPLICATION or OIC_SAMPLE
'  IDI_QUESTION    or OIC_QUES
'  IDI_EXCLAMATION or OIC_BANG
'  IDI_ASTERISK    or OIC_NOTE
'  IDI_WINLOGO     or OIC_WINLOGO
'BITMAPS:
'  OBM_LFARROW    'left arrow, up
'  OBM_RGARROW    'right arrow, up
'  OBM_DNARROW    'down arrow, up
'  OBM_UPARROW    'up arrow, up
'  OBM_LFARROWI   'left arrow, disabled
'  OBM_RGARROWI   'right arrow, disabled
'  OBM_DNARROWI   'down arrow, disabled
'  OBM_UPARROWI   'up arrow, disabled
'  OBM_LFARROWD   'left arrow, pressed
'  OBM_RGARROWD   'right arrow, pressed
'  OBM_DNARROWD   'down arrow, pressed
'  OBM_UPARROWD   'up arrow, pressed
'  OBM_RESTORE    'window restore button, up
'  OBM_RESTORED   'window restore button, pressed
'  OBM_ZOOM       'window Zoom button, up
'  OBM_ZOOMD      'window Zoom button, down
'  OBM_REDUCE     'window reduce button, up
'  OBM_REDUCED    'window reduce button, down
'  OBM_CLOSE      'windows logo (used in system menu on window)
'  OBM_BTNCORNERS 'small black filled circle. Used as a tag
'  OBM_CHECKBOXES 'mutliple check/readio box images
'  OBM_CHECK      'checkmark
'  OBM_SIZE       'status bar sizing image
'  OBM_BTSIZE     'status bar sizing image (same as above)
'  OBM_MNARROW    'menu right triangle arrow (for subs)
'  OBM_COMBO      'combobox down arrow
'
'********************************************************************************
Public Sub GetStockObject(StockObject As StockObjects, Picture As PictureBox)
  Dim ShadowDC As Long
  Dim isbm As Integer
  Dim param As String
  Dim idlong As Long
  Dim objhandle As Long, oldobject As Long
  Dim di As Long
  Dim bm As BITMAP
'
' Clear the picture control
'
  Picture.Cls
  ' Find out if it's a bitmap or an icon
  If StockObject > 32700 Then isbm = -1
  
  ' Extract the id value to use
  idlong = StockObject

  If isbm Then   ' It's a stock bitmap
'
' Create a memory device context compatible with the picture control
'
    ShadowDC = CreateCompatibleDC(Picture.hDC)
'
' Load the bitmap
'
    objhandle& = LoadBitmapBynum(0, idlong&)
'
' Retrieve the height and width of the bitmap
'
    di = GetObjectAPI(objhandle, Len(bm), bm)
'
' Select the bitmap into the memory DC, keeping a handle to the prior bitmap.
'
    oldobject = SelectObject(ShadowDC, objhandle)
'
' BitBlt the bitmap into the picture control
'
    Picture.Width = (bm.bmWidth + 2) * Screen.TwipsPerPixelX
    Picture.Height = (bm.bmHeight + 2) * Screen.TwipsPerPixelY
    di = BitBlt(Picture.hDC, 0, 0, bm.bmWidth, bm.bmHeight, ShadowDC&, 0, 0, SRCCOPY)
'
' Select the bitmap OUT of the memory DC...
'
    di = SelectObject(ShadowDC, oldobject)
'
' and delete it (yes - even though they are system
' bitmaps - this doesn't destroy them, just releases
' your private copy of the bitmap.
    di = DeleteObject(objhandle)
'
' Finally, delete the memory DC
'
    di = DeleteDC(ShadowDC)
  Else    ' It's an icon - a much easier process
'
' Get the stock icon
'
    objhandle& = LoadIconBynum(0, idlong&)
'
' Draw it directly onto the picture control
'
    Picture.Width = 38 * Screen.TwipsPerPixelX
    Picture.Height = 38 * Screen.TwipsPerPixelY
    di = DrawIcon(Picture.hDC, 2, 2, objhandle&)
  End If
End Sub

