Attribute VB_Name = "Module�ƻs�D�u"
Sub �ƻs�D�u()
    Dim ChairmanRng As Range
    Set ChairmanRng = Workbooks("RMA by Dennis.xls").Worksheets("�D�u").Range("A2:A14")
    
    Dim searchRng As Range
    Set searchRng = Workbooks("RMA by Dennis.xls").Worksheets("�j�M").Range("A7")
    
    Dim R
    For Each R In ChairmanRng
        If R = searchRng Then
            Dim sh As Worksheet
            Set sh = Workbooks("RMA by Dennis.xls").Worksheets("�D�u")
                Dim ChairmanRngROW%
                ChairmanRngROW = sh.Range(R.Address).Row
                sh.Range("H" & ChairmanRngROW) = Workbooks("RMA by Dennis.xls").Worksheets("�j�M").Range("G8") '�u�{�v
                sh.Range("J" & ChairmanRngROW) = Workbooks("RMA by Dennis.xls").Worksheets("�j�M").Range("F8") '�e�^���
                Exit For
        End If
    Next
    Set ChairmanRng = Nothing
    Set searchRng = Nothing
    Set sh = Nothing
    
    MsgBox "�ƻs����"
End Sub
