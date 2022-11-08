Attribute VB_Name = "Module寄信Excel"
Option Explicit
Sub SendMailEnvelope()
        
        Dim MailTO$, MailCC
        
        MailTO = ""
        
        MailCC = Array("", "")
        
        Application.ScreenUpdating = False
        
        Dim strText$, k%
        
        k = Range("A" & Rows.Count).End(xlUp).Row + 1
        
        Range("A1:C" & k).Select
        
        ActiveWorkbook.EnvelopeVisible = True
        
        With ActiveSheet.MailEnvelope

                strText = ""
        
                .Introduction = strText
                
                With .Item
                        
                        .To = MailTO
                        
                        .CC = Join(MailCC, ";")
                
                        .Subject = "請幫忙購買零件"
                        
                        Dim fd As FileDialog

                        Set fd = Application.FileDialog(msoFileDialogFilePicker)
                
                        fd.AllowMultiSelect = True
                        fd.Title = "請選擇要附加的檔案"
                        
                        MsgBox "請選擇要附加的檔案"
                        
                        If fd.Show = -1 Then
                                For Each sPath In fd.SelectedItems
                                        .Attachments.Add sPath
                                Next
                        End If
                        
                        Set fd = Nothing
                            
                        '.send
                                        
                End With
        
        End With

        'ActiveWorkbook.EnvelopeVisible = False
        
        Application.ScreenUpdating = True
End Sub

