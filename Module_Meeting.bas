Attribute VB_Name = "Modulemeeting"
Sub meeting()
Attribute meeting.VB_ProcData.VB_Invoke_Func = " \n14"
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        Dim rng As Range
        Set rng = Sheets("Meeting").Columns("A:P")
        rng.Delete
        
        Dim rng1 As Range
        Set rng1 = Sheets("Dennis").Columns("A:K")
        
        rng1.Copy Sheets("Meeting").Columns("A:K")
        
        Set rng = Nothing
        Set rng1 = Nothing
        
        Sheets("Meeting").Select
        
        LF = Range("A" & Rows.Count).End(xlUp).Row + 1
        
        Range("H" & LF & ":" & "I" & LF + 10).Delete Shift:=xlUp
        
        
        If Range("A2") = "" Then
                Rows(4).Select
                ActiveWindow.FreezePanes = False
                ActiveWindow.FreezePanes = True
        Else
                LF = Range("A1").End(xlDown).Row
                LF = LF + 3
                Rows(LF).Select
                ActiveWindow.FreezePanes = False
                ActiveWindow.FreezePanes = True
        End If
        
        Range("A1").Select
        
        Sheets("Dennis").Select
        Range("A1").Select
        
        MsgBox "meeting  OK"
        Application.ScreenUpdating = True
End Sub
