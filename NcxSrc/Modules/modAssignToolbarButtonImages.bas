Attribute VB_Name = "modAssignToolbarButtonImages"
Option Explicit
'~modAssignToolbarButtonImages.bas;
'Assign images to a toolbar at run-time
'***************************************************************************
' modAssignToolbarButtonImages -- The AssignToolbarButtonImages()
'          function will assign images to a toolbar at run-time.
'          This simplifies the process of experimenting with different
'          images, because this allows one to not have to assign and
'          unassign image lists as the testing goes on.
'
'          For this to work, The KEY value in the Image list assigned
'          to images should correspond to KEY values assigned to the
'          corresponding button in the toolbar.
'
' NOTE: This module requires the Project Component
'       "Microsoft Windows Common Controls 6.0" (MSCOMCTL.OCX) be loaded.
'       (Obviously, because it contains the ImageList and ToolBar.)
'***************************************************************************

Public Sub AssignToolbarButtonImages(TlBar As MSComctlLib.Toolbar, ImgList As MSComctlLib.ImageList)
  Dim MyButton As MSComctlLib.Button
  
  With TlBar
    .ImageList = ImgList            'assign image list
    On Error Resume Next
    For Each MyButton In .Buttons   'scan through each button in the toolbar
      If Len(MyButton.Key) <> 0 Then
        MyButton.Image = MyButton.Key 'assign an image
    End If
    Next MyButton                   'do all buttons
  End With 'TlBar

End Sub
