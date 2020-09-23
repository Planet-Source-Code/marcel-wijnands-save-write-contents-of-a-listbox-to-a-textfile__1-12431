<div align="center">

## Save / Write contents of a listbox to a textfile


</div>

### Description

Save /Write contents of a listbox to a textfile
 
### More Info
 
output filename and listbox name

Usage:

ListBox2File "c:\yourfile.txt", ListBox1


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Marcel Wijnands](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marcel-wijnands.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marcel-wijnands-save-write-contents-of-a-listbox-to-a-textfile__1-12431/archive/master.zip)





### Source Code

```
Public Sub ListBox2File(sFile As String, oList As ListBox)
Dim fnum, x As Integer
Dim sTemp As String
 fnum = FreeFile()
 x = 0
 Open sFile For Output As fnum
  While x <> oList.ListCount
   Print #fnum, oList.List(x)
   x = x + 1
  Wend
 Close fnum
End Sub
'Check out http://www.vb2delphi.com for more code!
```

