Attribute VB_Name = "Module�H�HExcel"
Option Explicit
Sub SendMailEnvelope()
        
        Dim MailTO$, MailCC
        
        MailTO = "bella.chang@kaitek.com.tw"
        
        MailCC = Array("fortran.wu@kaitek.com.tw", "ivan.feng@kaitek.com.tw")
        
        Application.ScreenUpdating = False
        
        Dim strText$, k%
        
        k = Range("A" & Rows.Count).End(xlUp).Row + 1
        
        Range("A1:C" & k).Select
        
        ActiveWorkbook.EnvelopeVisible = True
        
        With ActiveSheet.MailEnvelope

                strText = "Bella" & vbCrLf & vbCrLf & "�������T�{�PIssue �Ƹ�" & vbCrLf & _
                "�ڻݭn�q�ʷs�ơA�w��gParts order" & vbCrLf & _
                "�O�γ������ڱ�RD" & vbCrLf & "����"
        
                .Introduction = strText
                
                With .Item
                        
                        .To = MailTO
                        
                        .CC = Join(MailCC, ";")
                
                        .Subject = "�������ʶR�s��"
                        
                        Dim fd As FileDialog

                        Set fd = Application.FileDialog(msoFileDialogFilePicker)
                
                        fd.AllowMultiSelect = True
                        fd.Title = "�п�ܭn���[���ɮ�"
                        
                        MsgBox "�п�ܭn���[���ɮ�"
                        
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

