<div align="center">

## PrintJustify Update


</div>

### Description

This code will take a paragraph of words and print them out left and right justified, just like newsprint. You have total print control as to where on the page the print occurs as well as fontsize, attributes, line spacing, etc. As far as I can tell, there is nothing like this on PSC. Please vote and give comments.
 
### More Info
 
This is an update of my previous submission. Fixed a potential bug and neatened code with the addition of variable yPrint.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Charles Roland](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/charles-roland.md)
**Level**          |Beginner
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/charles-roland-printjustify-update__1-42036/archive/master.zip)





### Source Code

```
Option Explicit
Dim strPara As String  'String containing the Paragraph to be printed
Private Sub Form_Load()
strPara = "This is an example of how to print a paragraph that is left and right justified, just like a column of newsprint. xLeft contains the x position of the column start and xRight contains the x position of the column end. yStart contains the y position of the first line and ySpacing the amount of space you want between lines and is optional (250 is the default). Set all Printer.Font attributes before calling the PrintJustify subroutine. Play with the settings in Form_Load to fully test and check this code. As far as I could tell, there is nothing like this on PSC. Have Fun and please vote if you found this code informative and useful. Any comments would be appreciated."
Printer.FontSize = 8
PrintJustify strPara, 700, 6000, 1000, 250
Printer.EndDoc
End
End Sub
Private Sub PrintJustify(ByVal strText As String, xLeft As Integer, xRight As Integer, yStart As Integer, Optional ySpacing As Integer)
Dim aWords() As String 'Array containing all the words in strText
Dim iWords As Integer  'The Number of words that will be printed in the current iLine
Dim iWidth As Integer  'The available space for a line to print in
Dim iLine As Integer  'The current line to be printed
Dim xSpace As Integer  'The amount of Space between each word to be printed
Dim xPrint As Integer  'The x position that the next word will be printed in
Dim yPrint As Integer  'The y position that the next word will be printed in
Dim strWords As String 'A string containing the words of the line to be printed
Dim i As Integer    'For/Next counter Variable
Dim j As Integer    'For/Next counter Variable
'Replace any CRLF with space
strText = Replace(strText, vbCrLf, " ")
'Replace any " " with " "
strText = Replace(strText, " ", " ")
'Remove any before or after spaces
strText = Trim(strText)
'Initialize Line Counter
iLine = 0
'Set ySpacing Default
If ySpacing = 0 Then
  ySpacing = 250
End If
'Calculate Width of Print Column
iWidth = xRight - xLeft
'Keep Processing until all lines have been printed
Do Until strText = ""
  'Increment Line Counter
  iLine = iLine + 1
  'Calculate yPrint
  yPrint = yStart + ((iLine - 1) * ySpacing)
  'Break strText into pieces
  Erase aWords
  aWords = Split(strText, " ")
  'Determine How many words can fit in line
  strWords = ""
  For i = 0 To UBound(aWords)
   If i = 0 Then
     strWords = aWords(0)
     Else
     strWords = strWords & " " & aWords(i)
   End If
   If Printer.TextWidth(strWords) > iWidth Then
     iWords = i - 1 'last word becomes first word of next line
     Exit For
   End If
  Next i
  'Rewrite StrText if Words are still left
  strText = ""
  If iWords < UBound(aWords) Then
   For i = iWords + 1 To UBound(aWords)
     If i = iWords + 1 Then
      strText = aWords(i)
      Else
      strText = strText & " " & aWords(i)
     End If
   Next i
  Else 'Print Last Line with No Justification
   Printer.CurrentX = xLeft
   Printer.CurrentY = yPrint
   Printer.Print strWords
   Exit Sub
  End If
  'Now we can Print Justified Line
  'Get Width of all Words to be Printed (No Spaces)
  strWords = ""
  For i = 0 To iWords
   strWords = strWords & aWords(i)
  Next i
  'Get Width of Blank Space
  xSpace = iWidth - Printer.TextWidth(strWords)
  'Calculate Blank Space Between Words
  xSpace = Int(xSpace / iWords)
  'Print Last Word Right Justified
  Printer.CurrentX = xRight - Printer.TextWidth(aWords(iWords))
  Printer.CurrentY = yPrint
  Printer.Print aWords(iWords)
  'Print Words xSpace apart
  For i = 0 To iWords - 1 'We just printed the last word
   'Calculate xPrint
   xPrint = xLeft
   For j = 1 To i
     xPrint = xPrint + Printer.TextWidth(aWords(j - 1)) + xSpace
   Next j
   Printer.CurrentX = xPrint
   Printer.CurrentY = yPrint
   Printer.Print aWords(i)
  Next i
Loop
End Sub
```

