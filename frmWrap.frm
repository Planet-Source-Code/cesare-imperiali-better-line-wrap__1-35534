VERSION 5.00
Begin VB.Form frmWrap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quite a ""justify"""
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSplit 
      Caption         =   "Split the string"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtTheText 
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmWrap.frx":0000
      Top             =   0
      Width           =   8055
   End
   Begin VB.CheckBox chkForce 
      Caption         =   "ForceSplit"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "frmWrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Better Line Wrap
'Simple and fast Linewrap Code that works, is simple and FREE
 
'**************************************
' Name: Better Line Wrap
' Description:Simple and fast Linewrap written
'             with a bit of Vb standards. This one
'             do same job as Sam Kopetzky one plus
'             one feature: it can split lines in same
'             length and if no space is found to split
'             them, it will divide the word with a "-"
'             char
' Inputs: 1)theWholeText is for the string you want to be
'         splitted (first parameter),
'         2)theMaxLenOfLine is for the maximun line length
'         you want it to be
'         3)blnForceSplit (optional) is to decide which kind
'         of splitting you want:
'         the kind of split Sams did(blnForceSplit= false),
'         or a more fixed length one where words can be split-
'         ted (blnForceSplit= true)
' Returns:The whole mess is returned as value of the function
'
' Assumes:if you copy-paste the code example for the form,
'         you are supposed to have a textbox on form (named
'         "txtTheText") a command button (named "cmdSplit")
'         and a checkBox (named "chkForce")
'
' Side Effects:none
'
'This code is NOT copyrighted and you can play with it as
'you like.
'
'**************************************
 
  


'--------------------------
'On a form put a text box (multiline= true, scrollbars = both)
'To let user decide which kind of "justifcation" he wants,
'add a check box, and pass to the function  the check value
'as last parameter
'call the function this way:
'--------------------------
'Form Code
Option Explicit
Private Sub cmdSplit_Click()
   'first parameter is the text you want to "justify",
   'second is the line length
   'third is a boolean to say (if true) you want the lines to be splitted
   'even if no space is found (thus breaking words and in this case adding
   'a "-" char) or (if false) to serach for spaces first, and thus having
   'lines of different length - and only if no space is found you have the
   'line splitted in a word divided by the "-" char.
   txtTheText.Text = MyJustify(txtTheText.Text, 60, chkForce.Value)
   'Note:
   '"chkForce" is a CheckBox control
End Sub

Private Sub Form_Load()
txtTheText.Text = "This is a quite long string just to demonstrate how this stuff works. Decide which kind of split you want,"
txtTheText.Text = txtTheText.Text & " by checking or uncheching the checkbox and clicking on command button."
txtTheText.Text = txtTheText.Text & vbCrLf & "It is only a demo on how to use the bas code. You can improve this as you like, and use"
txtTheText.Text = txtTheText.Text & " this code wherever you want without asking for any permission. Have a nice day and happy coding."
txtTheText.FontName = "Courier New"
End Sub
