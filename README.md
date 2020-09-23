<div align="center">

## Check  if dynamic array dimensioned already


</div>

### Description

Tells if a dynamic array has been dimensioned or not.

Lu <learly@ix.netcom.com>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Newsgroup Posting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/newsgroup-posting.md)
**Level**          |Unknown
**User Rating**    |4.2 (159 globes from 38 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/newsgroup-posting-check-if-dynamic-array-dimensioned-already__1-534/archive/master.zip)





### Source Code

```
Function Member(ary$(), text$)
  On Local Error GoTo MemberExit
  For i = 1 To UBound(ary$)
    If text$ = ary$(i) Then
      subscript = i
      Exit For
    End If
  Next
MemberExit:
  Member = subscript
End Function
;========================================
another possibility;
Function ArrayElements(ary$())
  elements = 0
  On Local Error GoTo MemberExit
  elements = UBound(ary$)
MemberExit:
  ArrayElements = elements
End Function
```

