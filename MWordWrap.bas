Attribute VB_Name = "MWordWrap"
'Module bas code
Option Explicit

 Public Function MyJustify(theWholeText As String, theMaxLenOfLine As Integer, Optional blnForceSplit As Boolean = True) As String
      Dim astrForcedLines() As String
      Dim lngCounter As Long
      Dim theTempString As String
      'fragment for newlines
      astrForcedLines = Split(theWholeText, vbCrLf)
      'for each line (element of array)
      'decide if line  needs to be divided
      'in smallest pieces
      For lngCounter = LBound(astrForcedLines) To UBound(astrForcedLines)
         If Len(astrForcedLines(lngCounter)) > theMaxLenOfLine Then
            astrForcedLines(lngCounter) = fnctSplitIt(astrForcedLines(lngCounter), theMaxLenOfLine, blnForceSplit)
         End If
      Next lngCounter
      'rebuild final string
      MyJustify = Join(astrForcedLines, vbCrLf)
 End Function

Private Function fnctSplitIt(strInput As String, Maxlen As Integer, blnForceSplit As Boolean) As String
   Dim intFoundPos As Integer
   Dim theSplitChar As String
   Dim theSplitLen As Integer
  
   'strip unnecessary right spaces
   strInput = RTrim(strInput)
   'the blnForceSplit is to determine if you want
   'fixed len lines where division is made with "-" char
   '(when breaking a word) and vbcrlf
   If blnForceSplit Then
      Do While Len(strInput) > Maxlen
         If Mid(strInput, Maxlen, 1) = " " Then
            theSplitChar = ""
         Else
            theSplitChar = "-"
         End If
         fnctSplitIt = fnctSplitIt & Left(strInput, Maxlen) & theSplitChar & vbCrLf
         If Len(strInput) - Maxlen > 0 Then
            strInput = Right(strInput, Len(strInput) - Maxlen)
            'trim non significant left spaces
            strInput = LTrim(strInput)
         End If
      Loop
      'add last piece of string
      
      fnctSplitIt = fnctSplitIt & strInput
   'If the blnForceSplit is false, then you want this code
   'do divide lines where a break is found (a space), and
   'only if a space is not found, you want it to divide lines
   'like before, with a "-" char when breaking a word
   Else
      Do While Len(strInput) > Maxlen
         intFoundPos = InStrRev(strInput, " ", Maxlen)
         If intFoundPos > 0 Then
            fnctSplitIt = fnctSplitIt & Left(strInput, intFoundPos - 1) & vbCrLf
            theSplitLen = intFoundPos
         Else
            fnctSplitIt = fnctSplitIt & Left(strInput, Maxlen - 1) & "-" & vbCrLf
            theSplitLen = Maxlen
         End If
         If Len(strInput) - theSplitLen > 0 Then
               strInput = Right(strInput, Len(strInput) + 1 - theSplitLen)
               'trim non significant left spaces
               strInput = LTrim(strInput)
         End If
      Loop
      'add last piece of string
      fnctSplitIt = fnctSplitIt & strInput
   End If
End Function
