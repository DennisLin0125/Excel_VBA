Attribute VB_Name = "Module��ŪOUTLOOK"
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
                        
                        MsgBox "�o�e�̡G" & strSender & vbCrLf _
                        & "CC�G" & strCC & vbCrLf _
                        & "�D�D�G" & strSub & vbCrLf _
                        & "����G" & strBody
                        k = False
                End If
        Next objItem
        
        For Each objSubFolder In objFolder.Folders
                ListAllFolders objSubFolder
        Next objSubFolder
        
        If k Then MsgBox "Outlook �̨S����Ū�H��", vbInformation
        
        Set objItem = Nothing
        Set objSubFolder = Nothing

End Sub
