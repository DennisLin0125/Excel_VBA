Attribute VB_Name = "ModuleFIND搜尋"
Sub FIND搜尋()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    myTime = Time
    
    Dim rng As Range
    Set rng = Workbooks("RMA by Dennis.xls").Worksheets("搜尋").Range("A6").CurrentRegion
    rng.Offset(1).ClearContents
    Set rng = Nothing
    
    Dim snRng As Range
    Set snRng = Workbooks("RMA by Dennis.xls").Worksheets("搜尋").Range("B1")
    
    Dim snDennis As Worksheet
    Set snDennis = Workbooks("RMA by Dennis.xls").Worksheets("搜尋")
    
    Dim RmaStartYear As Range
    Dim RmaStopYear As Range
    
    Set RmaStartYear = Range("B3")
    Set RmaStopYear = Range("B4")
    
    Dim RowIndex%, Row%, LF%
    
    Row = 7
    
    For i = RmaStartYear To RmaStopYear Step -1
    
        Dim fname$
        fname = ""
        
        Dim wb As Workbook
        Set wb = Workbooks.Open(fname)
        
        wb.Activate
        
        LF = Range("A1").End(xlDown).Row
        
        If wb.Worksheets("Master").FilterMode Then
                wb.Worksheets("Master").ShowAllData
        End If
        
         With wb.Worksheets("Master")
                .AutoFilter.Sort.SortFields.Clear
                .AutoFilter.Sort.SortFields.Add Key:=Range("A1:A" & LF), Order:=xlDescending
                .AutoFilter.Sort.Apply
        End With
        
        Dim machine As Range
        Set machine = Range("K1:K" & LF).Find(What:=snRng, LookAt:=xlWhole)
        
        If Not machine Is Nothing Then
            Dim firstRng As Range
            Set firstRng = machine
            Do
                RowIndex = machine.Row
                snDennis.Range("B" & Row) = Range("A" & RowIndex)  'RMA
                snDennis.Range("C" & Row) = Range("C" & RowIndex) 'call date
                snDennis.Range("D" & Row) = Range("D" & RowIndex)  '客戶
                snDennis.Range("E" & Row) = Range("G" & RowIndex)  '機種
                snDennis.Range("F" & Row) = Range("I" & RowIndex)  'MN
                snDennis.Range("G" & Row) = Range("K" & RowIndex) 'SN
                snDennis.Range("H" & Row) = Range("P" & RowIndex)  'Ship date
                snDennis.Range("I" & Row) = Range("T" & RowIndex) 'Engineer
                snDennis.Range("J" & Row) = Range("Q" & RowIndex)  'Warranty Type
                snDennis.Range("K" & Row) = Range("U" & RowIndex)  'NPO
                snDennis.Range("L" & Row) = Range("Y" & RowIndex)  '故障內容
                Row = Row + 1
                
                Set machine = Range("K1:K" & LF).FindNext(machine)
                
            Loop Until machine.Address = firstRng.Address
            
        End If
        
        wb.Close False
        
    Next i
    
    snDennis.Activate
    
    Row = Row - 1
    For i = 7 To Row
        If snDennis.Range("H" & i + 1) = "" Then
                snDennis.Range("A" & i) = ""
        Else
                snDennis.Range("A" & i) = (snDennis.Range("C" & i) - snDennis.Range("H" & i + 1)) & " 天"
        End If
    Next
    
    Set snRng = Nothing
    Set snDennis = Nothing
    Set RmaStartYear = Nothing
    Set RmaStopYear = Nothing
    Set wb = Nothing
    Set machine = Nothing
    Set firstRng = Nothing
    
    myTime = Time - myTime
    myMin = Minute(myTime)
    mySec = Second(myTime)
    
    MsgBox "搜尋完畢" & vbLf & vbLf & "使用時間" & myMin & "分" & mySec & "秒。"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

