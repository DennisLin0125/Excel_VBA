Attribute VB_Name = "Module�H�H"
Sub �H�H()
        strBody = "<B>Hi Jack</B><br><br>" & _
              "<B>�w����</B>" & _
              "<H3><B>All the best</B></H3>" & _
              "<B>Dennis Lin</B><br>" & _
              "<B>**********************************</B><br>" & _
              "<B>Kaitek Coproration</B><br>" & _
              "<B>14F., No. 659, Bannan Rd, Zhonghe Dist.,</B><br>" & _
              "<B>New Taipei City 23557, Taiwan R.O.C.</B><br>" & _
              "<B>TEL: 02-3234-8222 #251</B><br>" & _
              "<B>**********************************</B>"
      Dim outApp As Object
      Set outApp = CreateObject("Outlook.Application")
      
      Dim outMail As Object
      Set outMail = outApp.CreateItem(0)
      
      With outMail
            '.To = "jack.chen@kaitek.com.tw"
            .To = "tom.lin@kaitek.com.tw"
            .CC = ""
            .BCC = ""
            .Subject = "2�t��Ƴ�"
            .HtmlBody = strBody
            '.Body = "��������� ����"
            .Attachments.Add "P:\Service\Parts data\��ƲM��_2021_2�t.xls"
            
'            Dim fd As FileDialog
'            ans = MsgBox("�O�_�n���[W3M���R���i?", vbYesNo)
'            If ans = vbYes Then
'                  Set fd = Application.FileDialog(msoFileDialogFilePicker)
'                  fd.Title = "�п�ܭn���[��W3M�ɮ�"
'                  MsgBox "�п�ܭn���[��W3M�ɮ�"
'                  If fd.Show = -1 Then
'                        .Attachments.Add fd.SelectedItems(1)
'
'                        Dim temp$, myTitle$
'                        temp = Mid(fd.SelectedItems(1), 1, InStr(fd.SelectedItems(1), ".") - 1)
'                        myTitle = Mid(temp, InStrRev(temp, "\") + 1)
'
'                        .Subject = myTitle
'                        .Body = "����    " & myTitle & "   �O�T���R���i"
'                  End If
'            Else
'                  Set fd = Application.FileDialog(msoFileDialogFilePicker)
'                  fd.Title = "�п�ܪ��[�ɮ�"
'                  MsgBox "�п�ܪ��[�ɮ�"
'                  If fd.Show = -1 Then
'                        .Attachments.Add fd.SelectedItems(1)
'                  End If
                  
'                  Dim main1$, main2$
'                  main1 = InputBox("�п�J�D��")
'                  main2 = InputBox("�п�J����")
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
'      fath = "P:\Service\Parts data\��ƲM��_2021_2�t.xls"
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
'      MsgBox "����,  Tom ���q�� 206"
End Sub

