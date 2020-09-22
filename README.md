<div align="center">

## A 'strReplace' function\.


</div>

### Description

When called with a string, will search through the string and replace a character of your choice with another character of your choice. For example, if you sent the string:

"Hello to the world"

And sent "o" as the character to be replaced,

and sent "a" as the replacement

It will return you with:

"Hella ta the warld".
 
### More Info
 
OldString, OldLetter and NewLetter

AND THIS IS FOR VB5.....

The modified string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Danny Young](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/danny-young.md)
**Level**          |Intermediate
**User Rating**    |4.3 (64 globes from 15 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/danny-young-a-strreplace-function__1-5250/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
  Dim oldstring As String, newletter As String, oldletter As String, newstring As String
  oldstring = "hello To the world"
  newletter = "YEAH"
  oldletter = "hello"
  newstring = Replace(oldstring, newletter, oldletter)
  MsgBox newstring
End Sub
Public Function Replace(oldstring, newletter, oldletter) As String
  Dim i As Integer
  i = 1
  Do While InStr(i, oldstring, oldletter, vbTextCompare) <> 0
    Replace = Replace & Mid(oldstring, i, InStr(i, oldstring, oldletter, vbTextCompare) - i) & newletter
    i = InStr(i, oldstring, oldletter, vbTextCompare) + Len(oldletter)
  Loop
  Replace = Replace & Right(oldstring, Len(oldstring) - i + 1)
End Function
```

