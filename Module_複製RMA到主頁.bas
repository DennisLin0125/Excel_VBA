Attribute VB_Name = "Module�ƻsRMA��D��"
Sub �ƻsRMA()
Attribute �ƻsRMA.VB_ProcData.VB_Invoke_Func = " \n14"
        Application.ScreenUpdating = False
        Dim myRow%, DRow%
        
        myRow = Range("A" & Rows.Count).End(xlUp).Row
        DRow = Sheets("Dennis").Range("A" & Rows.Count).End(xlUp).Row
        
        Range("A2:O" & myRow).Copy
        Sheets("Dennis").Range("A" & DRow + 1).Insert Shift:=xlDown
        
        Application.CutCopyMode = False
        Sheets("Dennis").Select
        Application.ScreenUpdating = True
End Sub

