Attribute VB_Name = "ModuleAWeekW3M"
Option Explicit
Private Enum myDate
        index0
        index1
        index2
        index3
        index4
        index5
        index6
End Enum
Sub WeekW3M()
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
    
        Dim rng5 As Range
        Set rng5 = Workbooks("待修分析.xlsm").Worksheets("本周保固").Cells(1, 1).CurrentRegion
        rng5.Offset(1).ClearContents
    
        
        Dim wb As Workbook, fpathThisYear$, fpathBeforeYear$, UpdateLinks%, Year%, myDate$
        
        myDate = Format(Date, "yyyymmdd")
        
        Year = Mid(myDate, 1, 4) '取年分
        
        fpathThisYear = ""
        fpathBeforeYear = ""
        Set wb = Workbooks.Open(fpathThisYear, UpdateLinks:=0)
    
        Dim arr
        arr = Array(Date - index6, Date - index5, Date - index4, Date - index3, Date - index2, Date - index1, Date - index0)
        
        Dim LF%, Total%, RPS%, W3M%, i%, j%
        wb.Activate
        LF = Cells(1, 1).End(xlDown).Row
    
        Total = 0
        RPS = 0
        W3M = 0
    
        For i = LBound(arr) To UBound(arr)
                For j = 1 To LF
                        If InStr(Cells(j, "C"), arr(i)) Then
                            Total = Total + 1
                        End If
                        If InStr(Cells(j, "C"), arr(i)) * InStr(Cells(j, "Q"), "3") Then
                            W3M = W3M + 1
                        End If
                Next j
        Next i
        
        Dim TempMyRow%
        TempMyRow = 0
        Dim firstRng   As Range
    
        For j = LBound(arr) To UBound(arr)
    
                    For i = 1 To LF
                
                    If InStr(Cells(i, "C"), arr(j)) * InStr(Cells(i, "Q"), "3") Then
                                Dim tempSn$, tempCust$
                                tempSn = Cells(i, "K")    '  暫存SN
                                tempCust = Cells(i, "D") '  暫存客戶名稱
                                Dim rngSn As Range
                                Set rngSn = Range("K1:K" & i - 1).Find(What:=tempSn, After:=Cells(i - 1, "K"), LookAt:=xlWhole, SearchDirection:=xlPrevious) '搜尋是否有保固
                                
                                If Not rngSn Is Nothing Then
                                        Dim myRow%, engineer$, RMA$, ShipDate As Date
                                        myRow = rngSn.Row
                                        If Range("D" & myRow) = tempCust Then
                                                engineer = Cells(myRow, "T")
                                                RMA = Cells(myRow, "A")
                                                ShipDate = Cells(myRow, "P")
                                        ElseIf Left(Range("D" & myRow), 3) = "UMC" And Left(tempCust, 3) = "UMC" Then  '判斷是否為UMC
                                                engineer = Cells(myRow, "T")
                                                RMA = Cells(myRow, "A")
                                                ShipDate = Cells(myRow, "P")
                                        Else
                                                Set firstRng = rngSn
                                                Do                                     '找下一個保固
                                                        Set rngSn = Range("K1:K" & i - 1).FindPrevious(rngSn)
                                                        myRow = rngSn.Row
                                                        If Range("D" & myRow) = tempCust Then
                                                                engineer = Cells(myRow, "T")
                                                                RMA = Cells(myRow, "A")
                                                                ShipDate = Cells(myRow, "P")
                                                                Exit Do
                                                        ElseIf Left(Range("D" & myRow), 3) = "UMC" And Left(tempCust, 3) = "UMC" Then  '判斷是否為UMC
                                                                engineer = Cells(myRow, "T")
                                                                RMA = Cells(myRow, "A")
                                                                ShipDate = Cells(myRow, "P")
                                                                Exit Do
                                                        End If
                                                Loop Until rngSn.Address = firstRng.Address
                                        End If
                                Else
                                        Dim wb2 As Workbook
                                        Set wb2 = Workbooks.Open(fpathBeforeYear, UpdateLinks:=0)
                                        wb2.Activate
                                        LF = Range("A1").End(xlDown).Row
                                        Dim rng2 As Range
                                        Set rng2 = Range("K1:K" & LF).Find(What:=tempSn, After:=Cells(LF, "K"), LookAt:=xlWhole, SearchDirection:=xlPrevious)
                                
                                        If Not rng2 Is Nothing Then
                                                myRow = rng2.Row
                                                If Range("D" & myRow) = tempCust Then
                                                        engineer = Cells(myRow, "T")
                                                        RMA = Cells(myRow, "A")
                                                        ShipDate = Cells(myRow, "P")
                                                ElseIf Left(Range("D" & myRow), 3) = "UMC" And Left(tempCust, 3) = "UMC" Then  '判斷是否為UMC
                                                        engineer = Cells(myRow, "T")
                                                        RMA = Cells(myRow, "A")
                                                        ShipDate = Cells(myRow, "P")
                                                Else
                                                        Set firstRng = rng2
                                                        Do                                     '找下一個保固
                                                                Set rng2 = Range("K1:K" & LF).FindPrevious(rng2)
                                                                myRow = rng2.Row
                                                                If Range("D" & myRow) = tempCust Then
                                                                        engineer = Cells(myRow, "T")
                                                                        RMA = Cells(myRow, "A")
                                                                        ShipDate = Cells(myRow, "P")
                                                                        Exit Do
                                                                 ElseIf Left(Range("D" & myRow), 3) = "UMC" And Left(tempCust, 3) = "UMC" Then  '判斷是否為UMC
                                                                        engineer = Cells(myRow, "T")
                                                                        RMA = Cells(myRow, "A")
                                                                        ShipDate = Cells(myRow, "P")
                                                                        Exit Do
                                                                End If
                                                        Loop Until rng2.Address = firstRng.Address
                                                End If
                                        End If
                                        wb2.Close False
                                        Set wb2 = Nothing
                                End If
                                
                                Dim MyArr(30, 10)
                                
                                MyArr(TempMyRow, 0) = Cells(i, "A") 'RMA
                                MyArr(TempMyRow, 1) = Cells(i, "D") 'Customer
                                MyArr(TempMyRow, 2) = Cells(i, "C") 'DateIN
                                MyArr(TempMyRow, 3) = Cells(i, "I") 'MN
                                MyArr(TempMyRow, 4) = Cells(i, "K") 'SN
                                MyArr(TempMyRow, 5) = Cells(i, "Y") 'Customer Complaint
                                MyArr(TempMyRow, 6) = Cells(i, "G") 'Model Type
                                MyArr(TempMyRow, 7) = engineer
                                MyArr(TempMyRow, 8) = Cells(i, "C") - ShipDate
                                MyArr(TempMyRow, 9) = ShipDate
                                
                                engineer = ""
                                'ShipDate = ""
                                
                                Dim RPtimes%
                                RPtimes = Application.WorksheetFunction.CountIf(wb.Worksheets("Master").Range("K1:K" & LF), tempSn)
                                If RPtimes = 1 Then
                                        RPtimes = 1
                                ElseIf RPtimes = 2 Then
                                        RPtimes = 1
                                Else
                                        RPtimes = RPtimes - 1
                                End If
                                MyArr(TempMyRow, 10) = RPtimes
                                TempMyRow = TempMyRow + 1
                        End If
                Next i
        Next j
        wb.Close False
        Set wb = Nothing
        With Workbooks("待修分析.xlsm").Worksheets("本周保固")
                .[A2].Resize(TempMyRow, 11) = MyArr
                .Activate
        End With
        
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        
        MsgBox "處理完成" & Chr(10) & Chr(10) & _
               "這禮拜收到 " & Total & " 台機器" & Chr(10) & Chr(10) & _
               "共有 " & W3M & " 台保固"
End Sub
