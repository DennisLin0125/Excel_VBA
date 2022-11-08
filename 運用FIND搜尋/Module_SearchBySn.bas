Attribute VB_Name = "ModuleSearchBySn"
Sub SearchBySn()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim main As Worksheet
    On Error Resume Next
    
    myTime = Time
    
    Dim LF, Row, RmaStartYear, RmaStopYear As Integer
    Dim temp As Double
    
    Range("A7:L30") = ""
    Range("A1").Select
    
    Set Dennis = Workbooks("RMA by Dennis.xls").Worksheets("搜尋")
    
    Row = 7
    
    RmaStartYear = Dennis.Range("B3")
    RmaStopYear = Dennis.Range("B4")
    
    For i = RmaStartYear To RmaStopYear Step -1
    
        Workbooks.Open Filename:="P:\Service\RMA\Main\Kaitek RMA " & i & " main.xls"
        
        Set main = Workbooks("Kaitek RMA " & i & " main.xls").Worksheets("Master")
        main.Activate
        LF = Range("A1").End(xlDown).Row
        
        Do While LF > 1
            
            
            'main.Range("K" & LF).Select
            
            If main.Range("K" & LF) <> "" Then   '非空白才列印
            
                '數字
                temp = Val(Trim(Dennis.Range("B1")))
                
                '文字
                If main.Range("K" & LF) = Dennis.Range("B1") Then
                    
                    Dennis.Range("B" & Row) = main.Range("A" & LF)  'RMA
                    Dennis.Range("C" & Row) = main.Range("C" & LF)  'call date
                    Dennis.Range("D" & Row) = main.Range("D" & LF)  '客戶
                    Dennis.Range("E" & Row) = main.Range("G" & LF)  '機種
                    Dennis.Range("F" & Row) = main.Range("I" & LF)  'MN
                    Dennis.Range("G" & Row) = main.Range("K" & LF)  'SN
                    Dennis.Range("H" & Row) = main.Range("P" & LF)  'Ship date
                    Dennis.Range("I" & Row) = main.Range("T" & LF)  'Engineer
                    Dennis.Range("J" & Row) = main.Range("Q" & LF)  'Warranty Type
                    Dennis.Range("K" & Row) = main.Range("U" & LF)  'NPO
                    Dennis.Range("L" & Row) = main.Range("Y" & LF)  '故障內容
                    Row = Row + 1
                    
                ElseIf main.Range("K" & LF) = temp Then
        
                    Dennis.Range("B" & Row) = main.Range("A" & LF)  'RMA
                    Dennis.Range("C" & Row) = main.Range("C" & LF)  'call date
                    Dennis.Range("D" & Row) = main.Range("D" & LF)  '客戶
                    Dennis.Range("E" & Row) = main.Range("G" & LF)  '機種
                    Dennis.Range("F" & Row) = main.Range("I" & LF)  'MN
                    Dennis.Range("G" & Row) = main.Range("K" & LF)  'SN
                    Dennis.Range("H" & Row) = main.Range("P" & LF)  'Ship date
                    Dennis.Range("I" & Row) = main.Range("T" & LF)  'Engineer
                    Dennis.Range("J" & Row) = main.Range("Q" & LF)  'Warranty Type
                    Dennis.Range("K" & Row) = main.Range("U" & LF)  'NPO
                    Dennis.Range("L" & Row) = main.Range("Y" & LF)  '故障內容
                    Row = Row + 1
                End If
            End If
            LF = LF - 1
        Loop
        Workbooks("Kaitek RMA " & i & " main.xls").Close False
    Next i
    
    Row = Row - 1
    For i = 7 To Row
        If Dennis.Range("H" & i + 1) = "" Then
                Dennis.Range("A" & i) = ""
        Else
                Dennis.Range("A" & i) = (Dennis.Range("C" & i) - Dennis.Range("H" & i + 1)) & " 天"
        End If
    Next
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Set Dennis = Nothing
    Set main = Nothing
    
    myTime = Time - myTime
    myMin = Minute(myTime)
    mySec = Second(myTime)
    
    MsgBox ("處理完成" & Chr(10) & Chr(10) & "使用時間" & myMin & "分" & mySec & "秒。")
End Sub

