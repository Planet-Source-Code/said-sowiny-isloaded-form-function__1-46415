<div align="center">

## IsLoaded Form Function


</div>

### Description

This Function checks if a specified Form is loaded

by looping through the forms collection

it returns TRUE if it is and FALSE if it is not

you can decide what you want according to the

return value e.g accessing its properties or methods or controls ,...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Said Sowiny](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/said-sowiny.md)
**Level**          |Advanced
**User Rating**    |2.5 (15 globes from 6 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/said-sowiny-isloaded-form-function__1-46415/archive/master.zip)





### Source Code

<B><FONT SIZE=1><P ALIGN="LEFT"> </P>
<P ALIGN="LEFT">Public Function IsLoaded(ByVal strForm As String)As Boolean</P><DIR>
<DIR>
<P ALIGN="LEFT">Dim frmloaded As Form</P>
<P ALIGN="LEFT">IsLoaded = False</P>
<P ALIGN="LEFT">If strForm = "" Then Exit Function</P>
<P ALIGN="LEFT">For Each frmloaded In Forms</P><DIR>
<DIR>
<P ALIGN="LEFT">If frmloaded.Name = strForm Then</P><DIR>
<DIR>
<P ALIGN="LEFT">IsLoaded = True</P>
<P ALIGN="LEFT">Exit Function</P></DIR>
</DIR>
<P ALIGN="LEFT">End If</P></DIR>
</DIR>
<P ALIGN="LEFT">Next</P></DIR>
</DIR>
<P ALIGN="LEFT">End Function</P></B></FONT>

