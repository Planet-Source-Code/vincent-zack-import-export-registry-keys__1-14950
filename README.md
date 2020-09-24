<div align="center">

## Import Export Registry Keys


</div>

### Description

Import And Export Registry Keys
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Vincent Zack](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vincent-zack.md)
**Level**          |Intermediate
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Registry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/registry__1-36.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vincent-zack-import-export-registry-keys__1-14950/archive/master.zip)





### Source Code

```
Function ExportNode(sKeyPath As String, sOutFile As String)
'
'Example:
'ExportNode "HKEY_LOCAL_MACHINE\software\microsoft","c:\windows\desktop\out.reg"
'
'/E (Export) switch
Shell "regedit /E " & sOutFile & " " & sKeyPath
End Function
Function ImportNode(sInFile As String)
'
'Example:
'ImportNode "c:\windows\desktop\reg.reg"
'
'/I (Import) /S (Silent) switchs
Shell "regedit /I /S " & sInFile
End Function
```

