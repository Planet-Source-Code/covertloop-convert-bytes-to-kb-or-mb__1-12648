<div align="center">

## Convert Bytes to KB or MB


</div>

### Description

This code will enable you to convert bytes to kilobytes or megabytes, whichever you choose. Vote if ya wanna.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[CovertLoop](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/covertloop.md)
**Level**          |Beginner
**User Rating**    |3.0 (18 globes from 6 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/covertloop-convert-bytes-to-kb-or-mb__1-12648/archive/master.zip)





### Source Code

```
Private Declare Function StrFormatByteSize Lib _
"shlwapi" Alias "StrFormatByteSizeA" (ByVal _
dw As Long, ByVal pszBuf As String, ByRef _
cchBuf As Long) As String
Public Function FormatKB(ByVal Amount As Long) _
As String
Dim Buffer As String
Dim Result As String
Buffer = Space$(255)
Result = StrFormatByteSize(Amount, Buffer, _
Len(Buffer))
If InStr(Result, vbNullChar) > 1 Then
FormatKB = Left$(Result, InStr(Result, _
vbNullChar) - 1)
End If
End Function
```

