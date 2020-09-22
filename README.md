<div align="center">

## Better Line Wrap


</div>

### Description

Simple and fast Linewrap written

with a bit of Vb standards. This one

does same job as Sam Kopetzky one plus

one feature: it can split lines in same

length and if no space is found to split

them, it will divide the word with a "-"

char (lines will result 1 char longer than

max you said!)

The Sam Kopetzky code is compatible with all Vb versions, while this one is only for Vb6 - as it uses functions like Split and Join - but it may be easier to understand. If you want to take a look to Sam job, go here: [url]http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=35460&lngWId=1[/url]
 
### More Info
 
1)theWholeText is for the string you want to be

splitted (first parameter),

2)theMaxLenOfLine is for the maximun line length

you want it to be

3)blnForceSplit (optional) is to decide which kind

of splitting you want:

the kind of split Sams did(blnForceSplit= false),

or a more fixed length one where words can be splitted (blnForceSplit= true)

No copyright on this code. You can use it as you like.

The whole mess is returned as value of the function


<span>             |<span>
---                |---
**Submitted On**   |2002-06-06 15:07:38
**By**             |[Cesare Imperiali](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cesare-imperiali.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Better\_Lin91342672002\.zip](https://github.com/Planet-Source-Code/cesare-imperiali-better-line-wrap__1-35534/archive/master.zip)

### API Declarations

```
'No copyright on this code. You can use it as you
'like.
'Module bas code
Option Explicit
 Public Function MyJustify(theWholeText As String, theMaxLenOfLine As Integer, Optional blnForceSplit As Boolean = True) As String
  Dim astrForcedLines() As String
  Dim lngCounter As Long
  Dim theTempString As String
  'fragment for newlines
  astrForcedLines = Split(theWholeText, vbCrLf)
  'for each line (element of array)
  'decide if line needs to be divided
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
```





