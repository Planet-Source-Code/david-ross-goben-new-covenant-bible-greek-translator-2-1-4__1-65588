Attribute VB_Name = "modMessageBox"
Option Explicit
'~modMessageBox.bas;frmMessageBox.frm;modGetStockObject.bas;
'Display a custom messagebox
'******************************************************************************
' This module, using frmMessageBox.frm, displays a message box that is customizable.
' By setting IconImage(4).Picture to an image of your choice, and providing the
' vbSystemModal flag, your custom image will be displayed as the message icon.
'
'EXAMPLE
'  frmMessageBox.IconImage(4).Picture = MyImage.Picture
'  If MessageBox(Me, "Are you nuts?", vbYesNo Or vbSystemModal, "Sanity Check") = vbYes Then
'    MessageboxBox Me, "Obviously"
'  Else
'    MessageBox Me, "I doubt that"
'  End If
'
' NOTE: This form uses modMessageBox.bas.
' NOTE: This form uses modGetStockObject.bas.
'******************************************************************************

Public Function MessageBox(iForm As Form, Prompt As String, Optional Flags As VbMsgBoxStyle, Optional PromptCaption As String = vbNullString) As VbMsgBoxResult
  MessageBox = frmMessageBox.dwMessageBox(iForm, Prompt, Flags, PromptCaption)
End Function
