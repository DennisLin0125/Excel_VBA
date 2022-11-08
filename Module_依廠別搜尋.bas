Attribute VB_Name = "Module依廠別搜尋"
Sub 廠別()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim LF, Row, RmaStartYear, RmaStopYear As Integer
    Dim RmaCus As String
    
    On Error Resume Next
    
    Range("A7:L200").Select
    Selection.ClearContents
    Range("A1").Select
    
    Set Dennis = Workbooks("RMA by Dennis.xls").Worksheets("搜尋")
    
    Row = 7
    
    RmaCus = Dennis.Range("B2").Value
    
    RmaStartYear = Dennis.Range("B3").Value
    RmaStopYear = Dennis.Range("B4").Value
    
    For i = RmaStartYear To RmaStopYear Step -1
    
        Workbooks.Open Filename:="P:\Service\RMA\Main\Kaitek RMA " & i & " main.xls"
        
        Set main = Workbooks("Kaitek RMA " & i & " main.xls").Worksheets("Master")
        
        LF = Range("A1").End(xlDown).Row
        
        Do While LF > 1
            
            
            main.Range("D" & LF).Select
            
            '文字
            If main.Range("D" & LF).Value = RmaCus And main.Range("G" & LF).Value = "Rapid Source" Then
                
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
    Dennis.Activate
    MsgBox ("處理完成")
End Sub


