Attribute VB_Name = "Module移除重複"
Sub repeat()
        Dim shMain As Worksheet
        Set shMain = ActiveSheet

        Dim d As Object, arr, brr, i&, k&, s$
        Set d = CreateObject("scripting.dictionary")
    
        arr = [A1].CurrentRegion
    
        ReDim brr(1 To UBound(arr), 1 To 11)
        
        For i = UBound(arr) To 1 Step -1
    
                s = arr(i, 1)
       
                If Not d.exists(s) Then
                        d(s) = ""
                        k = k + 1
                        brr(k, 1) = arr(i, 1)
                        brr(k, 2) = arr(i, 2)
                        brr(k, 3) = arr(i, 3)
                        brr(k, 4) = arr(i, 4)
                        brr(k, 5) = arr(i, 5)
                        brr(k, 6) = arr(i, 6)
                        brr(k, 7) = arr(i, 7)
                        brr(k, 8) = arr(i, 8)
                        brr(k, 9) = arr(i, 9)
                        brr(k, 10) = arr(i, 10)
                        brr(k, 11) = arr(i, 11)
                End If
        Next
        Cells.ClearContents
        With shMain.[A1].Resize(k, 11)
                .Value = brr
        End With
        
        MsgBox "全部共有 " & UBound(arr) & " 個" & Chr(10) & Chr(10) & "共有：" & k & " 個不重複值" & Chr(10) & Chr(10) & "共刪除了 " & UBound(arr) - k & " 個重複"
        Set d = Nothing
End Sub

