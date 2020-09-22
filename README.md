<div align="center">

## \[AA\+\] Remove Duplicates From Listview


</div>

### Description

Fast way to Remove duplicates from ListView. I havn't seen anthing that is the same on PSC, so I hope this will help in the quest for a faster way to remove duplicate items from a listview.. though I really don't think there is a "Fast" way to do it.. Who knows.. ENJOY!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[AngryLoop](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/angryloop.md)
**Level**          |Beginner
**User Rating**    |4.7 (28 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/angryloop-aa-remove-duplicates-from-listview__1-58490/archive/master.zip)





### Source Code

```
Private Sub RemoveDuplicates(lst As ListView)
Dim lRet As ListItem
Dim strTemp As String
Dim intCnt As Integer
intCnt = 0
Do While intCnt <= lst.ListItems.Count - 1
 intCnt = intCnt + 1
 'Save the text that was in the listvew index
 strTemp = lst.ListItems.Item(intCnt).Text
 Do
 lst.ListItems.Item(intCnt).Text = "" 'Remove the text inside the specific index
 'Use the FindItem() call to search for the specific item
 Set lRet = lst.FindItem(strTemp, lvwText, lvwPartial)
 'If the item is found, then it is a duplicate and is removed
 If Not lRet Is Nothing Then
 lst.ListItems.Remove (lRet.Index)
 End If
 Loop While Not lRet Is Nothing 'If no item is found the loop is exited
 lst.ListItems.Item(intCnt).Text = strTemp 'reset the listitem index text back to what it was, and then continue
 Debug.Print intCnt
 DoEvents 'Added to ensure that the application does not lock up when doing large amounts of data.
Loop
End Sub
```

