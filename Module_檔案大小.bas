Attribute VB_Name = "Module檔案大小"
Sub 檔案容量()
Attribute 檔案容量.VB_ProcData.VB_Invoke_Func = "o\n14"
        ActiveWorkbook.Save
        Dim fs As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        Myname = [F7] & ".xls"
        Dim i As Integer, ans As Boolean, dataSize As Double, x As Double
        For i = 2021 To 2006 Step -1
                ans = fs.FileExists("")
                If ans Then
                        dataSize = (VBA.FileLen(""))
                        dataSize = dataSize / 1024 / 1024
                        x = Math.Round(dataSize, 2)
                        
                        If x > 2 Then
                                MsgBox "檔案大小 :  " & x & " MB" & Chr(10) & Chr(10) & "檔案超過２ＭＢ，請壓縮或減少照片！！", vbCritical
                                Exit Sub
                        Else
                                MsgBox "OK     檔案大小 :  " & x & " MB", vbInformation
                        End If
                        
                        Exit For
                End If
                If i = 2006 Then
                        MsgBox Range("F7") & " 檔案不存在"
                        Exit Sub
                End If
        Next
End Sub
