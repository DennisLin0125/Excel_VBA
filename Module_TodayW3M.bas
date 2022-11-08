Attribute VB_Name = "ModuleTodayW3M"
Option Explicit
Sub 搜尋今日保固()
Attribute 搜尋今日保固.VB_ProcData.VB_Invoke_Func = "w\n14"
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        Dim wb As Workbook, fpathThisYear$, fpathBeforeYear$, myDate$, Year%
        
        myDate = Format(Date, "yyyymmdd")
        
        Year = Mid(myDate, 1, 4) '取年分
        
        fpathThisYear = ""
        fpathBeforeYear = ""
        
        Dim UpdateLinks%
        Set wb = Workbooks.Open(fpathThisYear, UpdateLinks:=0)
        
        Dim times%, Total%, RPS$
        times = 1
        Total = 0
        RPS = 0
        
        Dim myStr$, LF%
        Dim firstRng As Range
        
        Dim myDay As Date
        myDay = Date
        
        wb.Activate
        LF = Range("A" & Rows.Count).End(xlUp).Row  '最後一列
        
        Dim i%
        For i = 1 To LF  '尋遍Main檔
                If InStr(Cells(i, "C"), myDay) Then   '計算當天送修台數
                        Total = Total + 1
                End If
        Next i
        
        For i = 1 To LF   '尋遍Main檔
                If InStr(Cells(i, "C"), myDay) * InStr(Cells(i, "Q"), "3") Then '當天是否有保固
                        Dim tempSn$, tempCust$
                        tempSn = Cells(i, "K")    '  暫存SN
                        tempCust = Cells(i, "D") '  暫存客戶名稱
                        Dim rngSn As Range
                        Set rngSn = Range("K1:K" & i - 1).Find(What:=tempSn, After:=Cells(i - 1, "K"), LookAt:=xlWhole, SearchDirection:=xlPrevious) '搜尋是否有保固
                        
                        If Not rngSn Is Nothing Then
                                Dim myRow%
                                myRow = rngSn.Row
                                If Range("D" & myRow) = tempCust Then      '確認保固的SN是同個廠商
                                        Dim engineer$, LifeTime As Double, RMA$, Reson$
                                        engineer = Cells(myRow, "T")
                                        LifeTime = Date - Cells(myRow, "P")
                                        RMA = Cells(myRow, "A")
                                        Reson = Cells(myRow, "Y")
                                ElseIf Left(Range("D" & myRow), 3) = "UMC" And Left(tempCust, 3) = "UMC" Then  '判斷是否為UMC
                                        engineer = Cells(myRow, "T")
                                        LifeTime = Date - Cells(myRow, "P")
                                        RMA = Cells(myRow, "A")
                                        Reson = Cells(myRow, "Y")
                                Else
                                        Set firstRng = rngSn
                                        Do                                     '找下一個保固
                                                Set rngSn = Range("K1:K" & i - 1).FindPrevious(rngSn)
                                                myRow = rngSn.Row
                                                If Range("D" & myRow) = tempCust Then
                                                        engineer = Cells(myRow, "T")
                                                        LifeTime = Date - Cells(myRow, "P")
                                                        RMA = Cells(myRow, "A")
                                                        Reson = Cells(myRow, "Y")
                                                        Exit Do
                                                ElseIf Left(Range("D" & myRow), 3) = "UMC" And Left(tempCust, 3) = "UMC" Then  '判斷是否為UMC
                                                        engineer = Cells(myRow, "T")
                                                        LifeTime = Date - Cells(myRow, "P")
                                                        RMA = Cells(myRow, "A")
                                                        Reson = Cells(myRow, "Y")
                                                        Exit Do
                                                End If
                                        Loop Until rngSn.Address = firstRng.Address
                                End If
                        Else
                                Dim wb2 As Workbook
                                Set wb2 = Workbooks.Open(fpathBeforeYear, UpdateLinks:=0)
                        
                                Dim LF2%
                                wb2.Activate
                                LF2 = Range("A1").End(xlDown).Row
                                Dim rng As Range
                                Set rng = Range("K1:K" & LF2).Find(What:=tempSn, After:=Cells(LF2, "K"), LookAt:=xlWhole, SearchDirection:=xlPrevious)
                                
                                If Not rng Is Nothing Then
                                        myRow = rng.Row
                                        If Range("D" & myRow) = tempCust Then   '確認保固的SN是同個廠商
                                                engineer = Cells(myRow, "T")
                                                LifeTime = Date - Cells(myRow, "P")
                                                RMA = Cells(myRow, "A")
                                                Reson = Cells(myRow, "Y")
                                        ElseIf Left(Range("D" & myRow), 3) = "UMC" And Left(tempCust, 3) = "UMC" Then  '判斷是否為UMC
                                                engineer = Cells(myRow, "T")
                                                LifeTime = Date - Cells(myRow, "P")
                                                RMA = Cells(myRow, "A")
                                                Reson = Cells(myRow, "Y")
                                        Else
                                                Set firstRng = rng
                                                Do                                     '找下一個保固
                                                        Set rng = Range("K1:K" & LF2).FindPrevious(rng)
                                                        myRow = rng.Row
                                                        If Range("D" & myRow) = tempCust Then
                                                                engineer = Cells(myRow, "T")
                                                                LifeTime = Date - Cells(myRow, "P")
                                                                RMA = Cells(myRow, "A")
                                                                Reson = Cells(myRow, "Y")
                                                                Exit Do
                                                        ElseIf Left(Range("D" & myRow), 3) = "UMC" And Left(tempCust, 3) = "UMC" Then  '判斷是否為UMC
                                                                engineer = Cells(myRow, "T")
                                                                LifeTime = Date - Cells(myRow, "P")
                                                                RMA = Cells(myRow, "A")
                                                                Reson = Cells(myRow, "Y")
                                                                Exit Do
                                                        End If
                                                Loop Until rng.Address = firstRng.Address
                                        End If
                                End If
                                wb2.Close False
                                Set wb2 = Nothing
                        End If
                        
                        If times = 1 Then
                                MsgBox "日期 : " & myDay & Chr(10) & Chr(10) & "共收到 " & Total & " 台機器" & Chr(10) & Chr(10) & "今日有保固"
                        End If
                        myStr = "今日第 " & times & " 台保固" & Chr(10) & Chr(10)
                        myStr = myStr & "保固工程師 : " & engineer & Chr(10) & Chr(10)
                        myStr = myStr & "客戶 : " & Cells(i, "D") & Chr(10) & Chr(10)
                        myStr = myStr & "MN : " & Cells(i, "I") & Chr(10) & Chr(10)
                        myStr = myStr & "SN : " & Cells(i, "K") & Chr(10) & Chr(10)
                        myStr = myStr & "機種 : " & Cells(i, "H") & Chr(10) & Chr(10)
                        
                        myStr = myStr & "前次RMA : " & RMA & Chr(10) & Chr(10)
                        myStr = myStr & "這次RMA : " & Cells(i, "A") & Chr(10) & Chr(10)
                        
                        myStr = myStr & "前次下機原因 : " & Reson & Chr(10) & Chr(10)
                        myStr = myStr & "此次下機原因 : " & Cells(i, "Y") & Chr(10) & Chr(10)
                        
                        myStr = myStr & "LifeTime :  " & LifeTime & " 天"
                        times = times + 1
                        MsgBox myStr
                End If
        Next i
        
        If times = 1 Then
                MsgBox "日期 : " & myDay & Chr(10) & Chr(10) & "共收到 " & Total & " 台機器" & Chr(10) & Chr(10) & "目前沒有保固"
        End If
        
        wb.Close False
        Set wb = Nothing
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
End Sub


