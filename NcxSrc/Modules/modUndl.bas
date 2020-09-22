Attribute VB_Name = "modUndl"
Option Explicit

'*******************************************************************************
' Subroutine Name   : UnderlineBbl
' Purpose           : Underline obvious bible passage reference
'*******************************************************************************
Public Sub UnderlineBbl(Rtb As RichTextBox)
  Dim I As Long, J As Long, K As Long, L As Long, IL As Long
  Dim S As String, T As String, Ary() As String
  
  With Rtb
    S = UCase$(.Text)                             'get text to process
    K = 0                                         'init to start of text -1
    K = InStr(K + 1, S, ":")                      'found a possible entry?
    Do While K <> 0                               'process while this is so
      If IsNumeric(Mid$(S, K - 1, 1)) And IsNumeric(Mid$(S, K + 1, 1)) Then
        I = K + 1                                 'point past ":"
        Do While IsNumeric(Mid$(S, I, 1))
          I = I + 1                               'find end of verse data
        Loop
        J = K                                     'no search over chapter data
        Do
          J = J - 1
        Loop While IsNumeric(Mid$(S, J, 1))
        If Mid$(S, J, 1) = " " Then               'space before chapter?
          J = J - 1
          If Mid$(S, J, 1) = "." Then J = J - 1   'possible book abbreviation?
          Do
            J = J - 1                             'point to start of book name
            Select Case Mid$(S, J, 1)
              Case "A" To "Z"
              Case Else
                Exit Do
            End Select
          Loop
          If Mid$(S, J, 1) = " " Then             'space before book?
            If IsNumeric(Mid$(S, J - 1, 1)) Then J = J - 2  'allow things like "2 John"
            T = UCase$(Mid$(S, J + 1, I - J - 1)) 'get bible data
            L = InStr(3, T, " ") - 1              'find book
            If Mid$(T, L, 1) = "." Then L = L - 1
            T = Left$(T, L)
            For IL = 1 To 27
              Ary = Split(Books(IL), ",")
              If StrComp(Left(Ary(3), Len(T)), T, vbTextCompare) = 0 Then Exit For
            Next IL
            If IL <= 27 Then                        'valid New Cov book?
              .SelStart = J                         'set start of selection, if yes
              .SelLength = I - J - 1                'length of selection
              .SelUnderline = True                  'underline it
            End If
          End If
        End If
      End If
      K = InStr(K + 1, S, ":")                    'check for another possible entry
    Loop
    .SelStart = 0                                 'reset pointers if all done
    .SelLength = 0
  End With
End Sub

'*******************************************************************************
' Function Name     : CheckUndl
' Purpose           : upon a double-click, if the user selected underlined text,
'                     point to that entry
'*******************************************************************************
Public Function CheckUndl(Rtb As RichTextBox) As Boolean
  Dim I As Long, K As Long, ibk As Long, ich As Long, ivs As Long
  Dim S As String, T As String, Ary() As String
  With Rtb
      S = .Text                                   ' Get text
      .SelLength = 1                              'check for underlined selection
      If .SelUnderline = True Then
        K = .SelStart + 1                         'is, so find limits of underline
        I = K
        Do
          .SelStart = I
          .SelLength = 1
          If .SelUnderline = False Then Exit Do
          I = I + 1
        Loop
        Do
          K = K - 1
          .SelStart = K
          .SelLength = 1
          If .SelUnderline = False Then Exit Do
        Loop
        S = Trim$(Mid$(S, K + 1, I - K))          'get underlined text
        K = InStrRev(S, ":")                      'find chp:vrs split
        If K <> 0 Then
          ivs = CLng(Mid$(S, K + 1))              'grab verse
          S = Left$(S, K - 1)
          K = InStrRev(S, " ")
          ich = CLng(Mid$(S, K + 1))              'grab chapter
          S = UCase$(Left$(S, K - 1))             'get book text
          If Right$(S, 1) = "." Then S = Left$(S, Len(S) - 1)
          ibk = 0
          For I = 1 To 27                         'find book
            Ary = Split(Books(I), ",")
            T = UCase$(Ary(3))
            If Left$(T, Len(S)) = S Then          'mstch?
              ibk = I                             'yes, so save index
              Exit For
            End If
          Next I
          If ibk <> 0 Then                          'if book was found
            S = .TextRTF                            'save text
            I = .BackColor                          'save background color
            K = .SelStart                           'save start position
            Bk = ibk                                'get book, chapter and verse
            Chp = ich
            Vrs = ivs
            ChpCnt = CLng(Ary(4))                   'get the chapter count
            Call frmGrkXlate.GetVerseCount          'get the verse count
            Call frmGrkXlate.UpdateVerse            'display the verse
            CheckUndl = True
            .TextRTF = S
            .BackColor = I
            .SelStart = K
            frmGrkXlate.cmdBack.Visible = True      'ensure RESET button available
          End If
        End If
      End If
    .SelLength = 0
  End With
End Function
