<div align="center">

## Switch Case


</div>

### Description

Did you know that the capital letters are different from the small letters only by one bit?

A = 01000001 while a = 01100000 the 3rd bit is the capitalization
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lefteris Eleftheriades](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lefteris-eleftheriades.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lefteris-eleftheriades-switch-case__1-62834/archive/master.zip)





### Source Code

```
Private Sub Form_Load()
 MsgBox SwitchCase("Sex") 'returns sEX
End Sub
Function SwitchCase(Text As String) As String
 Dim i&, out$
 For i = 1 To Len(Text)
   out = out & Chr(Asc(Mid(Text, i, 1)) Xor 32)
 Next
 SwitchCase = out
End Function
```

