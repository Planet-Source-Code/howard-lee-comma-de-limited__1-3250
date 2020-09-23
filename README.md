<div align="center">

## Comma De\-Limited


</div>

### Description

This code will read a file line by line and create a comma delimited text file which can then be imported into Excel.
 
### More Info
 
FileRead = File To be read as input

FileOutPut = file that will be created as output


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Howard Lee](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/howard-lee.md)
**Level**          |Unknown
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/howard-lee-comma-de-limited__1-3250/archive/master.zip)





### Source Code

```
Private Sub Form_Load ()
Dim instring As String
Dim outstring As String
On Error GoTo Clnup
Open "FileRead.txt" For Input As #1 ' file opened for reading
Open "FileOutPut.txt" For Output As #2 ' file created
Line Input #1, instring
While Not EOF(1)
  Line Input #1, instring
  If Len(outstring) = 0 Then
    outstring = instring
  Else
    outstring = outstring & "," & instring
  End If
Wend
Print #2, outstring
Close #1
Close #2
Clnup:
Close
End
End Sub
```

