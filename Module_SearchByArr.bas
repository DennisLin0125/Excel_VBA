Attribute VB_Name = "ModuleSearchByArr"
Sub SearchByArr()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    T1 = FormatDateTime(Time, vbGeneralDate)
    
    Dim rng As Range
    Set rng = Workbooks("RMA by Dennis.xls").Worksheets("搜尋").Range("A6").CurrentRegion
    rng.Offset(1).ClearContents
    Set rng = Nothing
    
    Dim snRng As Range
    Set snRng = Workbooks("RMA by Dennis.xls").Worksheets("搜尋").Range("B1")
    
    Dim snDennis As Worksheet
    Set snDennis = Workbooks("RMA by Dennis.xls").Worksheets("搜尋")
    
    Const snColm = 11
    
    Const ManchineColm = 7
    
    Dim Row%
    Row = 7
    
    Dim RmaStartYear As Range
    Dim RmaStopYear As Range
    
    Set RmaStartYear = Range("B3")
    Set RmaStopYear = Range("B4")
    
    For j = RmaStartYear To RmaStopYear Step -1
    
        
        Dim fname$
        fname = "P:\Service\RMA\Main\Kaitek RMA " & j & " main.xls"
        
        Dim wb As Workbook
        Set wb = Workbooks.Open(fname)
        
        With wb.Worksheets("Master")
            Dim arr
            arr = .Cells(1, 1).CurrentRegion
        End With
        
        wb.Close False
        
        Set wb = Nothing
        
        
        Dim i%
        
        For i = UBound(arr) To LBound(arr) Step -1
        
            If InStr(arr(i, snColm), snRng) * InStr(arr(i, ManchineColm), "Rapid Source") Then
                
                snDennis.Range("A" & Row) = arr(i, 1) 'RMA
                snDennis.Range("B" & Row) = arr(i, 4) '客戶
                snDennis.Range("C" & Row) = arr(i, 7) '機種
                snDennis.Range("D" & Row) = arr(i, 9) 'MN
                snDennis.Range("E" & Row) = arr(i, 11) 'SN
                snDennis.Range("F" & Row) = arr(i, 16) '送回日期
                snDennis.Range("G" & Row) = arr(i, 20) '工程師
                snDennis.Range("H" & Row) = arr(i, 17) 'W3M
                snDennis.Range("I" & Row) = arr(i, 21) 'NPO
                snDennis.Range("J" & Row) = arr(i, 25) '故障描述
                Row = Row + 1
                
            End If
        Next i
    Next j
    
    snDennis.Activate
    
    Set snRng = Nothing
    Set snDennis = Nothing
    Set RmaStartYear = Nothing
    Set RmaStopYear = Nothing
    
    Erase arr
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    t2 = FormatDateTime(Time, vbGeneralDate)
    
    MsgBox "處理完成" & Chr(10) & Chr(10) & "查詢時間" & DateDiff("s", T1, t2) & "秒。"
    
End Sub


