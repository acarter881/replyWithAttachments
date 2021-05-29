Sub ReplyWithAttachments()

Dim oReply As Outlook.MailItem
Dim oItem As Object
Dim oInspector As Inspector
 
Set oItem = GetCurrentItem()

If Not oItem Is Nothing Then
    Set oReply = oItem.Reply
    CopyAttachments oItem, oReply
    oReply.Display
    oItem.UnRead = False
    
    Set oInspector = Application.ActiveInspector
    
    If oInspector.IsWordMail Then
        Dim oDoc As Object, oWrdApp As Object, oSelection As Object
        Set oDoc = oInspector.WordEditor
        Set oWrdApp = oDoc.Application
        Set oSelection = oWrdApp.Selection
        Set recips = oReply.Recipients
        
        fromUser = recips(1)
        intspace = InStr(1, fromUser, " ")
        
        If intspace <> 0 Then
            theName = Left(fromUser, intspace - 1)
        End If
        
        If InStr(theName, ",") Then
            theName = Replace(theName, ",", "")
        ElseIf InStr(fromUser, ".") Then
            theName = Replace(fromUser, ".", " ")
        End If
        
        fMsg = "Hello, " & theName & ":" & vbNewLine & vbTab
        
        oSelection.InsertAfter fMsg
        oSelection.Collapse 0
        
        Set oSelection = Nothing
        Set oWrdApp = Nothing
        Set oDoc = Nothing
    End If
End If
 
Set oReply = Nothing
Set oItem = Nothing
Set recips = Nothing

End Sub
