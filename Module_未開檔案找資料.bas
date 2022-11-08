Attribute VB_Name = "Module���}�ɮק���"
Private Function GetValue(Path, file, sheet, ref)

        If Right(Path, 1) <> "\" Then Path = Path & "\"
        
        If Dir(Path & file) = "" Then
                GetValue = "File Not Found"
                Exit Function
        End If
        
        Dim arg$
        arg = "'" & Path & "[" & file & "]" & sheet & "'!" & Range(ref).Range("A1").Address(, , xlR1C1)
         
        GetValue = ExecuteExcel4Macro(arg)
        
End Function
Sub TestGetValue()
        P = "P:\Service\RMA\WR\2020"
        f = "P20109.xls"
        s = "RMA"
        a = "F11"
        
        Dim strLen
        strLen = GetValue(P, f, s, a)
        Debug.Print strLen
End Sub

