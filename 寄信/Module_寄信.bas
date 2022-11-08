Attribute VB_Name = "Module寄信"
Sub 寄信()
      strBody = ""
      Dim outApp As Object
      Set outApp = CreateObject("Outlook.Application")
      
      Dim outMail As Object
      Set outMail = outApp.CreateItem(0)
      
      With outMail
            .To = ""
            .CC = ""
            .BCC = ""
            .Subject = "2廠領料單"
            .HtmlBody = strBody
            '.Body = "請幫忙領料 謝謝"
            .Attachments.Add ""
            
'            Dim fd As FileDialog
'            ans = MsgBox("是否要附加W3M分析報告?", vbYesNo)
'            If ans = vbYes Then
'                  Set fd = Application.FileDialog(msoFileDialogFilePicker)
'                  fd.Title = "請選擇要附加的W3M檔案"
'                  MsgBox "請選擇要附加的W3M檔案"
'                  If fd.Show = -1 Then
'                        .Attachments.Add fd.SelectedItems(1)
'
'                        Dim temp$, myTitle$
'                        temp = Mid(fd.SelectedItems(1), 1, InStr(fd.SelectedItems(1), ".") - 1)
'                        myTitle = Mid(temp, InStrRev(temp, "\") + 1)
'
'                        .Subject = myTitle
'                        .Body = "附件為    " & myTitle & "   保固分析報告"
'                  End If
'            Else
'                  Set fd = Application.FileDialog(msoFileDialogFilePicker)
'                  fd.Title = "請選擇附加檔案"
'                  MsgBox "請選擇附加檔案"
'                  If fd.Show = -1 Then
'                        .Attachments.Add fd.SelectedItems(1)
'                  End If
                  
'                  Dim main1$, main2$
'                  main1 = InputBox("請輸入主旨")
'                  main2 = InputBox("請輸入內文")
'                  .Subject = main1
'                  .Body = main2
                  
'            End If
            .Display
            '.Send
      End With
            
'      Set outApp = Nothing
'      Set outMail = Nothing
'
'      Dim wbtest As Workbook
'      fath = ""
'      Set wbtest = Workbooks.Open(fath, UpdateLinks:=0)
'      wbtest.Activate
'
'      temp = 1
'
'      For i = 2 To 100
'            If Range("A" & i) <> "" Then
'                   temp = temp + 1
'            End If
'            If Range("A" & i) = "" Then Exit For
'      Next
'
'      Range("A2:N" & temp) = ""
'      wbtest.Close True
'      MsgBox "完成,  Tom 的電話 206"
End Sub

