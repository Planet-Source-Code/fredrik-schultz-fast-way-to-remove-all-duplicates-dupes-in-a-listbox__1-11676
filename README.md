<div align="center">

## Fast way to remove all duplicates \(dupes\) in a ListBox


</div>

### Description

This method removes all duplicates in a listbox, regardless if sorting is turned on or not.

AND it's fast, short and simple

(no double loops like in some other submissions).

It's also case-insensitive.
 
### More Info
 
Call by using: RemoveDupes MyListBox

(where MyListBox is the listbox that will

be cleaned of all duplicates.)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Fredrik Schultz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/fredrik-schultz.md)
**Level**          |Intermediate
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/fredrik-schultz-fast-way-to-remove-all-duplicates-dupes-in-a-listbox__1-11676/archive/master.zip)





### Source Code

```
Private Sub RemoveDupes(lst As ListBox)
 Dim iPos As Integer
 iPos=0
 '-- if listbox empty then exit..
 If lst.ListCount < 1 Then Exit Sub
 Do While iPos < lst.ListCount
  lst.Text = lst.List(iPos)
  '-- check if text already exists..
  If lst.ListIndex <> iPos Then
   '-- if so, remove it and keep iPos..
   lst.RemoveItem iPos
  Else
   '-- if not, increase iPos..
   iPos = iPos + 1
  End If
 Loop
 '-- used to unselect the last selected line..
 lst.Text = "~~~^^~~~"
End Sub
```

