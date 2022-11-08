Attribute VB_Name = "Module未讀OUTLOOK"
Sub ListUnReadMail()
        Dim objOutLook As Object
        Dim objNmspc As Object
        Dim objFolder As Object
        On Error GoTo errHandle
        Set objOutLook = CreateObject("Outlook.Application")
        Set objNmspc = objOutLook.GetNamespace("MAPI")
        Set objFolder = objNmspc.GetDefaultFolder(6)
        ListAllFolders objFolder
errHandle:
        objOutLook.Quit
        Set objOutLook = Nothing
        Set objNmspc = Nothing
        Set objFolder = Nothing
End Sub
Sub ListAllFolders(ByVal objFolder As Object)
        Dim objItem As Object
        Dim objSubFolder As Object
        Dim strSub$, strSender$, str$, strCC$, strBody$
        Dim k As Boolean
        k = True
        For Each objItem In objFolder.Items
                If objItem.UnRead Then
                
                        strSub = objItem.Subject
                        strSender = objItem.SenderEmailAddress
                        strCC = objItem.CC
                        strBody = objItem.Body
                        
                        MsgBox "發送者：" & strSender & vbCrLf _
                        & "CC：" & strCC & vbCrLf _
                        & "主題：" & strSub & vbCrLf _
                        & "內文：" & strBody
                        k = False
                End If
        Next objItem
        
        For Each objSubFolder In objFolder.Folders
                ListAllFolders objSubFolder
        Next objSubFolder
        
        If k Then MsgBox "Outlook 裡沒有未讀信件", vbInformation
        
        Set objItem = Nothing
        Set objSubFolder = Nothing

End Sub
