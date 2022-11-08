Attribute VB_Name = "Module_字典"
Sub test()
        Application.ScreenUpdating = False 
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        
        Dim d As Object, arr, i&
        Set d = CreateObject("scripting.dictionary")
        
        Dim wb As Workbook
        Set wb = Workbooks.Open("")
        
        With wb
                arr = [A1].CurrentRegion
        End With
        
        wb.Close False
        
        For i = 1 To UBound(arr)
                d(arr(i, 1)) = arr(i, 2)
        Next
        
        brr = [A1:B8]
        
        For i = 1 To UBound(brr)
                If d.exists(brr(i, 1)) Then
                        brr(i, 2) = d(brr(i, 1))
                Else
                        brr(i, 2) = "查無此資料"
                End If
        Next
        
        [A1].Resize(UBound(brr), 2) = brr
       
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
End Sub
