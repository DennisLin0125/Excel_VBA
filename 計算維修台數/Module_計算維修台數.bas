Attribute VB_Name = "Module計算維修台數"
Private Function countW3M(ByRef arr As Variant, ByVal engineer As String) As Integer
        Dim i%, times%
        times = 0
        For i = LBound(arr) To UBound(arr)
                If arr(i, 20) = engineer And arr(i, 17) = 3 Then
                        times = times + 1
                End If
        Next i
        countW3M = times
End Function
Private Function countManchine(ByRef arr As Variant, ByVal engineer As String) As Integer
        Dim i%, times%
        times = 0
        For i = LBound(arr) To UBound(arr)
                If InStr(arr(i, 20), engineer) > 0 Then
                        times = times + 1
                End If
        Next i
        countManchine = times
End Function

Sub 全部()
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        myTime = Time
        
        Range("C10:E22").ClearContents
        
        Dim wb As Workbook
        
        Dim Year%
        
        Year = Worksheets("工程師保固").Cells(4, "E")
        Set wb = Workbooks.Open(Filename:="")
        
        With wb.Worksheets("Master")
                Dim arr
                arr = .[A1].CurrentRegion
        End With
        
        Dim man
        
        man = Array("Jacky", "Ken", "Roy", "Mark", "Dennis", "Tim", "Lantis", "Bill", "Jeff", "Roma", "Eric", "Steven", "Frank")
        
        Dim mat()
        
        ReDim mat(UBound(man, 1), LBound(man, 1) + 2)
        
        Dim i%, j%
        
        For i = LBound(man, 1) To UBound(man, 1)
                mat(i, 0) = man(i)
                mat(i, 1) = countManchine(arr, man(i))
                mat(i, 2) = countW3M(arr, man(i))
        Next i
        
        wb.Close False
        
        With Workbooks("待修分析.xlsm").Worksheets("工程師保固")
                .[C10].Resize(UBound(mat, 1) + 1, UBound(mat, 2) + 1) = mat
                .Activate
        End With
        
        Set wb = Nothing
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        
        myTime = Time - myTime
        myMin = Minute(myTime)
        mySec = Second(myTime)
        MsgBox "搜尋完成" & Chr(10) & Chr(10) & "搜尋時間" & myMin & "分" & mySec & "秒"
    
End Sub

