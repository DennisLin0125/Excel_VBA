Attribute VB_Name = "Module�ɮפj�p"
Sub �ɮ׮e�q()
Attribute �ɮ׮e�q.VB_ProcData.VB_Invoke_Func = "o\n14"
        ActiveWorkbook.Save
        Dim fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        Myname = [F7] & ".xls"
        Dim i As Integer, ans As Boolean, dataSize As Double, x As Double
        For i = 2021 To 2006 Step -1
                ans = fs.FileExists("P:\Service\RMA\WR\" & i & "\" & Myname)
                If ans Then
                        dataSize = (VBA.FileLen("P:\Service\RMA\WR\" & i & "\" & Myname))
                        dataSize = dataSize / 1024 / 1024
                        x = Math.Round(dataSize, 2)
                        
                        If x > 2 Then
                                MsgBox "�ɮפj�p :  " & x & " MB" & Chr(10) & Chr(10) & "�ɮ׶W�L���ۢСA�����Y�δ�ַӤ��I�I", vbCritical
                                Exit Sub
                        Else
                                MsgBox "OK     �ɮפj�p :  " & x & " MB", vbInformation
                        End If
                        
                        Exit For
                End If
                If i = 2006 Then
                        MsgBox Range("F7") & " �ɮפ��s�b"
                        Exit Sub
                End If
        Next
End Sub
