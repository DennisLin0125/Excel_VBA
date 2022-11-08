Attribute VB_Name = "Module陣列分析"
Private Function KTcounter(ByVal keyword As String, ByVal colum As Integer, ByVal index As Integer, ByVal colum2 As Integer) As Integer

        If [a2] = "" Then
                LF = Range("G4").End(xlDown).Row
                For i = 4 To LF
                        If InStr(1, Cells(i, colum), keyword, vbBinaryCompare) + (index * (InStr(1, Cells(i, colum2), keyword, vbBinaryCompare))) Then
                                times = times + 1
                        End If
                Next i
        Else
                bgnRow = Range("G1").End(xlDown).Row + 3
                endRow = Range("A" & Rows.Count).End(xlUp).Row
                For i = bgnRow To endRow
                        If InStr(1, Cells(i, colum), keyword, vbBinaryCompare) + (index * (InStr(1, Cells(i, colum2), keyword, vbBinaryCompare))) Then
                                times = times + 1
                        End If
                Next i
        End If
        
        KTcounter = times
        
End Function
Private Function RPcounter() As Integer
        
        If [a2] = "" Then
                RPTime = 0
        Else
                LF = [A1].End(xlDown).Row
                RPTime = LF - 1
        End If
        RPcounter = RPTime

End Function
Private Function find(ByVal keyword As String, ByVal col As Integer, temp As Integer) As Integer
        Dim machine As Range, sRow%, nRow%, num%
        Dim firstRng As Range
        num = 0
        
        If [a2] = "" Then
                sRow = 4
        Else
                sRow = [A1].End(xlDown).Row + 3
        End If
        
        nRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'temp計算是否有RPS 報價或待料    1正常   0報價或待料
        If temp = 1 Then
                With Range(Cells(sRow, col), Cells(nRow, col))
                        Set machine = .find(What:=keyword, After:=Cells(sRow, col), Lookat:=xlPart)
                        If Not machine Is Nothing Then
                                Set firstRng = machine
                                Do
                                        Set machine = .FindNext(machine)
                                        num = num + 1
                                Loop Until machine.Address = firstRng.Address
                        End If
                End With
        Else
                With Range(Cells(sRow, col), Cells(nRow, col))
                        Set machine = .find(What:=keyword, After:=Cells(sRow, col), Lookat:=xlWhole)
                        If Not machine Is Nothing Then
                                Set firstRng = machine
                                Do
                                        Dim myStr$
                                        myStr = Left(Range("G" & machine.Row), 3)
                                        If myStr <> "WFC" And myStr <> "WFP" Then
                                                num = num + 1
                                        End If
                                        Set machine = .FindNext(machine)
                                Loop Until machine.Address = firstRng.Address
                        End If
                End With
                
        End If
        
        find = num
        
End Function

Sub 陣列()
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        
        t1 = FormatDateTime(Time, vbGeneralDate)
        
        Dim wb As Workbook, ans%
        For Each wb In Workbooks
                If wb.Name = "RMA by Jack.xlsx" Then
                        wb.Activate
                        ans = 1
                End If
        Next
        
        Set wb = Nothing
        
        If ans = 0 Then
                Dim fpath$
                fpath = ""
                Set wb = Workbooks.Open(fpath, UpdateLink = 0)
        End If
        
        Dim arrJackA
        arrJackA = Array("Jacky(214)", "Ken(229)", "Roy(231)", "Mark(217)")
        
        Dim engArrA As Variant
        engArrA = Array("Jacky", "Ken Lin", "Roy Yeh", "Mark Lai")
        
        Dim matA()
        ReDim matA(UBound(engArrA, 1), 5)
        
        Dim i%, j%
        For i = LBound(arrJackA) To UBound(arrJackA)
                Worksheets(arrJackA(i)).Activate
                matA(i, 0) = engArrA(i)
                matA(i, 1) = RPcounter
                matA(i, 2) = find("WR", 7, 1)   '1為正常WR
                matA(i, 3) = find("WFC", 7, 1)  '1為正常WR
                matA(i, 4) = find("WFP", 7, 1)  '1為正常WR
                matA(i, 5) = KTcounter("KAITEK", 2, 1, 7)
        Next
'***************************************************************************************************
        Dim arrJackB
        arrJackB = Array("Roma(223)", "Bill(216)", "Lantis(220)", "Tim(221)")
        
        Dim engArrB As Variant
        engArrB = Array("Roma", "Bill", "Lantis Sun", "Tim Chang")
        
        Dim matB()
        ReDim matB(UBound(engArrB, 1), 5)
        
        For i = LBound(arrJackB) To UBound(arrJackB)
                Worksheets(arrJackB(i)).Activate
                matB(i, 0) = engArrB(i)
                matB(i, 1) = RPcounter
                matB(i, 2) = find("WR", 7, 1) '1為正常WR
                matB(i, 3) = find("WFC", 7, 1) '1為正常WR
                matB(i, 4) = find("WFP", 7, 1) '1為正常WR
                matB(i, 5) = KTcounter("KAITEK", 2, 1, 7)
        Next
'************************************************************************************************************
        Dim arrJack As Variant
        arrJack = Array("Jacky(214)", "Ken(229)", "Roy(231)", "Mark(217)", "Bill(216)", "Lantis(220)", "Tim(221)", "Roma(223)")
        
        Dim EngArr As Variant
        EngArr = Array("Jacky", "Ken Lin", "Roy Yeh", "Mark Lai", "Bill", "Lantis Sun", "Tim Chang", "Roma")
        
        Dim matWR()
        ReDim matWR(UBound(EngArr), 1)
        
        For i = LBound(arrJack) To UBound(arrJack)
                Worksheets(arrJack(i)).Activate
                matWR(i, 0) = EngArr(i)
                matWR(i, 1) = find("Rapid Source", 8, 0) + find("Xstream Sources", 8, 0)
        Next
        
'************************************************************************************************************
        With Workbooks("待修分析.xlsm").Worksheets("待修")
                .[E3].Resize(UBound(matA, 1) + 1, 6) = matA
                .[L3].Resize(UBound(matB, 1) + 1, 6) = matB
                .[S4].Resize(UBound(matWR) + 1, 2) = matWR
                .Activate
        End With
        
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    
        t2 = FormatDateTime(Time, vbGeneralDate)
        
        MsgBox "處理完成" & Chr(10) & Chr(10) & "查詢時間" & DateDiff("s", t1, t2) & "秒。"
End Sub

