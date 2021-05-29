Sub CopyAttachments(objSourceItem, objTargetItem)

Set fso = CreateObject("Scripting.FileSystemObject")
Set fldTemp = fso.GetSpecialFolder(2) ' TemporaryFolder
      
strPath = fldTemp.Path & "\"
      
For Each objatt In objSourceItem.Attachments
 If LCase(Right(objatt.FileName, 3)) = "jpg" Or LCase(Right(objatt.FileName, 3)) = "png" Or LCase(Right(objatt.FileName, 3)) = "gif" Then 'Added this line to not attach jpg's, png's, or gif's
 'Do Nothing
 Else
 strFile = strPath & objatt.FileName
 objatt.SaveAsFile strFile
 objTargetItem.Attachments.Add strFile, , , objatt.DisplayName
 fso.DeleteFile strFile
 End If
Next

Set fldTemp = Nothing
Set fso = Nothing

End Sub
