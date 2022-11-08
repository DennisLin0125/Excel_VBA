Attribute VB_Name = "Module複製主席"
Sub 複製主席()
    Dim ChairmanRng As Range
    Set ChairmanRng = Workbooks("RMA by Dennis.xls").Worksheets("主席").Range("A2:A14")
    
    Dim searchRng As Range
    Set searchRng = Workbooks("RMA by Dennis.xls").Worksheets("搜尋").Range("A7")
    
    Dim R
    For Each R In ChairmanRng
        If R = searchRng Then
            Dim sh As Worksheet
            Set sh = Workbooks("RMA by Dennis.xls").Worksheets("主席")
                Dim ChairmanRngROW%
                ChairmanRngROW = sh.Range(R.Address).Row
                sh.Range("H" & ChairmanRngROW) = Workbooks("RMA by Dennis.xls").Worksheets("搜尋").Range("G8") '工程師
                sh.Range("J" & ChairmanRngROW) = Workbooks("RMA by Dennis.xls").Worksheets("搜尋").Range("F8") '送回日期
                Exit For
        End If
    Next
    Set ChairmanRng = Nothing
    Set searchRng = Nothing
    Set sh = Nothing
    
    MsgBox "複製完成"
End Sub
