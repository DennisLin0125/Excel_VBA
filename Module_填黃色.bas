Attribute VB_Name = "Module填黃色"
Sub 填黃色()
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        Dim myRng As Range
        Set myRng = Workbooks("RMA by Dennis.xls").Worksheets("Dennis").Range("A1:A" & Range("G" & Rows.Count).End(xlUp).Row)
        
        Dim wb As Workbook
        Set wb = Workbooks.Open("", UpdateLinks:=0)
        
        wb.Activate
        Dim rng As Range
        Set rng = wb.Worksheets("RMA List").Range("I1:I" & Range("I" & Rows.Count).End(xlUp).Row)
        
        For Each c In rng
                If c.Interior.Color = vbYellow And Range("F" & c.Row) = "Dennis" Then
                        For Each k In myRng
                                If k = Range("I" & c.Row) Then
                                        k.Interior.Color = vbYellow
                                        Exit For
                                End If
                        Next k
                End If
        Next c
        wb.Close False
        Workbooks("RMA by Dennis.xls").Worksheets("Dennis").Activate
        
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        
        MsgBox "完成"
End Sub
