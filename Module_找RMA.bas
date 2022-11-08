Attribute VB_Name = "Module找RMA"
Option Explicit
Function searchMN(ByVal MN As String) As String
        If MN = "3155031-014" Or MN = "3155031-036" Or MN = "3155031-020" Or MN = "3155031-038" Or MN = "3155031-039" Then
                searchMN = "AZX 63"
        ElseIf MN = "3155031-043" Or MN = "3155031-033" Or MN = "3155031-037" Then
                searchMN = "AZX 72"
                
        ElseIf MN = "3155053-003" Or MN = "3155053-005" Or MN = "3155053-007" Then
                searchMN = "RFG2K2V"
        
        ElseIf MN = "3155051-010" Or MN = "3155051-015" Or MN = "3155051-115" Then
                searchMN = "RFG5500"
        
        ElseIf MN = "3155027-000" Or MN = "3155027-000" Or MN = "3155027-003" Or MN = "3155027-003J" Or MN = "3155027-005" Or MN = "3155027-008" Or MN = "3155027-028" Then
                searchMN = "RFG1250"
        
        ElseIf MN = "3155059-026" Or MN = "3155059-001" Then
                searchMN = "RFDS1250"
                
        ElseIf MN = "3155094-003" Or MN = "3155077-003" Or MN = "3155094-007" Or MN = "3155094-006" Then
                searchMN = "MFA"
                
        ElseIf MN = "BG578830-T" Or MN = "102074526" Then
                searchMN = "FMB"
                
        ElseIf MN = "102026212" Then
                searchMN = "FM800"
                
        ElseIf MN = "61300017" Then
                searchMN = "ASPECT Platform"
                
        ElseIf MN = "3152420-120" Then
                searchMN = "Pinnacle II"
        End If
End Function
Sub 從Main檔找RMA()
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        Dim snDennis As Worksheet
        Set snDennis = Workbooks("RMA by Dennis.xls").Worksheets("RMA")
        
        Dim rg As Range
        Set rg = snDennis.Range("B1:N" & Range("B" & Rows.Count).End(xlUp).Row)
        rg.Offset(1).ClearContents
        Set rg = Nothing
        
        If [A2] = "" Then Exit Sub
        
        Dim rng As Range, DennisRow%
        DennisRow = Range("A" & Rows.Count).End(xlUp).Row
        Set rng = snDennis.Range("A1:A" & DennisRow)
        
        Dim fname$
        fname = ""
        
        Dim wb As Workbook
        Set wb = Workbooks.Open(fname, UpdateLinks:=0)
        
        wb.Activate
        
        If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
        
        Dim c As Range, oArr(30, 13)
        For Each c In rng
                Dim machine As Range, LF%
                LF = Range("A1").End(xlDown).Row
                Set machine = Range("A1:A" & LF).Find(What:=c, LookAt:=xlWhole)
                
                If Not machine Is Nothing Then
                        Dim myRow%, machineRow%
                        
                        machineRow = machine.Row
                        oArr(myRow, 0) = wb.Worksheets("Master").Range("D" & machineRow)
                        oArr(myRow, 1) = wb.Worksheets("Master").Range("C" & machineRow)
                        oArr(myRow, 2) = wb.Worksheets("Master").Range("I" & machineRow)
                        oArr(myRow, 3) = wb.Worksheets("Master").Range("K" & machineRow)
                        oArr(myRow, 4) = wb.Worksheets("Master").Range("Y" & machineRow)
                        oArr(myRow, 5) = "WR"
                        oArr(myRow, 6) = wb.Worksheets("Master").Range("G" & machineRow)
                        If wb.Worksheets("Master").Range("Q" & machineRow) = 3 Then
                               oArr(myRow, 7) = "*"
                        End If
                        oArr(myRow, 9) = "=IF(A" & myRow + 2 & "<>"""",TODAY()-C" & myRow + 2 & ","""")"
                        oArr(myRow, 11) = searchMN(oArr(myRow, 2))
                        oArr(myRow, 13) = wb.Worksheets("Master").Range("AB" & machineRow)
                        myRow = myRow + 1
                End If
        Next
        wb.Close False
        
        With snDennis
                .[B2].Resize(myRow, 14) = oArr
        End With
'=========================================================================
        fname = ""

        Set wb = Workbooks.Open(fname, UpdateLinks:=0)
        wb.Activate
        
        Dim rng2 As Range
        Set rng2 = snDennis.Range("M2:M" & DennisRow)
        
        Dim MystrAZX$, Mystr2$
        
        Dim TotalCus, CusIndex%, requestRow%
        TotalCus = Split("CMO3,AUO-L6A,AUO-L5C-T1,CMO5,TSMC6,TSMC3,TSMC5,TSMC8,TSMC15,VISC1,VISC2,VISC3,UMC,PRA,美光,凌巨-BD1,凌巨-BD2", ",")
        
        Dim tempRow%
        tempRow = 1
        Dim sh As Worksheet, i%, MyCus$
        For Each sh In wb.Worksheets
                sh.Activate
                Select Case sh.name
                Case Is = "AZX 維修需求"
                        Dim AZXRow%
                        AZXRow = Range("A" & Rows.Count).End(xlUp).Row
                        MystrAZX = "注意事項 : " & Chr(10)
                        For i = 1 To AZXRow
                                If UCase(Range("A" & i)) = "ALL" And Range("A" & i).Font.Strikethrough <> True Then
                                        MystrAZX = MystrAZX & Range("B" & i) & Chr(10)
                                End If
                        Next i
                        For Each c In rng2
                                If c = "AZX 63" Or c = "AZX 72" Then
                                        MyCus = snDennis.Range("B" & c.Row)
                                        For CusIndex = 0 To UBound(TotalCus)
                                                If MyCus = TotalCus(CusIndex) Then
                                                        For i = 1 To AZXRow
                                                                If Range("A" & i) = TotalCus(CusIndex) Then
                                                                        Mystr2 = Mystr2 & Range("B" & i) & Chr(10)
                                                                End If
                                                        Next
                                                        snDennis.Range("L" & c.Row) = MystrAZX & Chr(10) & Mystr2
                                                        Mystr2 = ""
                                                ElseIf MyCus = "TSMC14" Then
                                                        For i = 1 To AZXRow
                                                                If Left(Range("A" & i), 6) = "TSMC14" Then
                                                                        Mystr2 = Mystr2 & Range("B" & i) & Chr(10)
                                                                End If
                                                        Next
                                                        snDennis.Range("L" & c.Row) = MystrAZX & Chr(10) & Mystr2
                                                        Mystr2 = ""
                                                ElseIf Left(MyCus, 3) = "UMC" Then
                                                        For i = 1 To AZXRow
                                                                If Left(Range("A" & i), 3) = "UMC" Then
                                                                        Mystr2 = Mystr2 & Range("B" & i) & Chr(10)
                                                                End If
                                                        Next
                                                        snDennis.Range("L" & c.Row) = MystrAZX & Chr(10) & Mystr2
                                                        Mystr2 = ""
                                                End If
                                        Next CusIndex
                                End If
                        Next c
'=====================================================================================
                Case Is = "RFDS1250"
                        Dim RFDS1250Row%, MystrRFDS1250$, Mystr6$
                        RFDS1250Row = Range("A" & Rows.Count).End(xlUp).Row
                        MystrRFDS1250 = "注意事項 : " & Chr(10)
                        For i = 1 To RFDS1250Row
                                If UCase(Range("A" & i)) = "ALL" And Range("A" & i).Font.Strikethrough <> True Then
                                        MystrRFDS1250 = MystrRFDS1250 & Range("B" & i) & Chr(10)
                                End If
                        Next i
                        For Each c In rng2
                                If c = "RFDS1250" Then
                                        MyCus = snDennis.Range("B" & c.Row)
                                        For CusIndex = 0 To UBound(TotalCus)
                                                If MyCus = TotalCus(CusIndex) Then
                                                        For i = 1 To RFDS1250Row
                                                                If Range("A" & i) = TotalCus(CusIndex) Then
                                                                        Mystr6 = Mystr6 & Range("B" & i) & Chr(10)
                                                                End If
                                                        Next i
                                                        snDennis.Range("L" & c.Row) = MystrRFDS1250 & Chr(10) & Mystr6
                                                        Mystr6 = ""
                                                End If
                                        Next CusIndex
                                End If
                        Next c
'=====================================================================================
                Case Is = "MFA"
                        Dim MFARow%, MystrMFA$, Mystr7$
                        MFARow = Range("A" & Rows.Count).End(xlUp).Row
                        MystrMFA = "注意事項 : " & Chr(10)
                        For i = 1 To MFARow
                                If UCase(Range("A" & i)) = "ALL" And Range("A" & i).Font.Strikethrough <> True Then
                                        MystrMFA = MystrMFA & Range("B" & i) & Chr(10)
                                End If
                        Next i
                        For Each c In rng2
                                If c = "MFA" Then
                                        MyCus = snDennis.Range("B" & c.Row)
                                        For CusIndex = 0 To UBound(TotalCus)
                                                If MyCus = TotalCus(CusIndex) Then
                                                        For i = 1 To MFARow
                                                                If Left(Range("A" & i), 4) = "VISC" And Right(Range("A" & i), 3) = "All" Then   'VISC-All
                                                                        Mystr7 = Mystr7 & Range("B" & i) & Chr(10)
                                                                End If
                                                                If Range("A" & i) = TotalCus(CusIndex) Then
                                                                        Mystr7 = Mystr7 & Range("B" & i) & Chr(10)
                                                                End If
                                                        Next i
                                                        snDennis.Range("L" & c.Row) = MystrMFA & Chr(10) & Mystr7
                                                        Mystr7 = ""
                                                End If
                                        Next CusIndex
                                End If
                        Next c
'=====================================================================================
                Case Is = "RFG1250 "
                        Dim RFG1250Row%, Mystr1250$, Mystr5$
                        RFG1250Row = Range("A" & Rows.Count).End(xlUp).Row
                        Mystr1250 = "注意事項 : " & Chr(10)
                        For i = 1 To RFG1250Row
                                If UCase(Range("A" & i)) = "ALL" And Range("A" & i).Font.Strikethrough <> True Then
                                        Mystr1250 = Mystr1250 & Range("B" & i) & Chr(10)
                                End If
                        Next i
                        For Each c In rng2
                                If c = "RFG1250" Then
                                        MyCus = snDennis.Range("B" & c.Row)
                                        For CusIndex = 0 To UBound(TotalCus)
                                                If MyCus = TotalCus(CusIndex) Then
                                                        For i = 1 To RFG1250Row
                                                                If Range("A" & i) = TotalCus(CusIndex) Then
                                                                        Mystr5 = Mystr5 & Range("B" & i) & Chr(10)
                                                                End If
                                                        Next i
                                                        snDennis.Range("L" & c.Row) = Mystr1250 & Chr(10) & Mystr5
                                                        Mystr5 = ""
                                                End If
                                        Next CusIndex
                                End If
                        Next c
'===========================================================================================
                Case Is = "RFG2K2V"
                        Dim RFG2K2VRow%, Mystr2K2V$, Mystr3
                        RFG2K2VRow = Range("A" & Rows.Count).End(xlUp).Row
                        Mystr2K2V = "注意事項 : " & Chr(10)
                        For i = 1 To RFG2K2VRow
                                If UCase(Range("A" & i)) = "ALL" And Range("A" & i).Font.Strikethrough <> True Then
                                        Mystr2K2V = Mystr2K2V & Range("B" & i) & Chr(10)
                                End If
                        Next i
                        For Each c In rng2
                                If c = "RFG2K2V" Then
                                        MyCus = snDennis.Range("B" & c.Row)
                                        For CusIndex = 0 To UBound(TotalCus)
                                                If MyCus = TotalCus(CusIndex) Then
                                                        For i = 1 To RFG2K2VRow
                                                                If Range("A" & i) = TotalCus(CusIndex) Then
                                                                        Mystr3 = Mystr3 & Range("B" & i) & Chr(10)
                                                                End If
                                                        Next
                                                        snDennis.Range("L" & c.Row) = Mystr2K2V & Chr(10) & Mystr3
                                                        Mystr3 = ""
                                                ElseIf Left(MyCus, 3) = "UMC" Then
                                                        For i = 1 To AZXRow
                                                                If Left(Range("A" & i), 3) = "UMC" Then
                                                                        Mystr3 = Mystr3 & Range("B" & i) & Chr(10)
                                                                End If
                                                        Next
                                                        snDennis.Range("L" & c.Row) = Mystr2K2V & Chr(10) & Mystr3
                                                        Mystr3 = ""
                                                End If
                                        Next CusIndex
                                End If
                        Next c
                
'============================================================================================
                Case Is = "RFG5500"
                Dim RFG5500Row%, Mystr5500$, Mystr4$
                RFG5500Row = Range("A" & Rows.Count).End(xlUp).Row
                Mystr5500 = "注意事項 : " & Chr(10)
                For i = 1 To RFG5500Row
                        If UCase(Range("A" & i)) = "ALL" And Range("A" & i).Font.Strikethrough <> True Then
                                Mystr5500 = Mystr5500 & Range("B" & i) & Chr(10)
                        End If
                Next i
                For Each c In rng2
                        If c = "RFG5500" Then
                                MyCus = snDennis.Range("B" & c.Row)
                                For CusIndex = 0 To UBound(TotalCus)
                                        If MyCus = TotalCus(CusIndex) Then
                                                For i = 1 To RFG5500Row
                                                        If Left(Range("A" & i), 4) = "VISC" And Right(Range("A" & i), 3) = "All" Then   'VISC-All
                                                                Mystr4 = Mystr4 & Range("B" & i) & Chr(10)
                                                        End If
                                                        If Range("A" & i) = TotalCus(CusIndex) Then
                                                                Mystr4 = Mystr4 & Range("B" & i) & Chr(10)
                                                        End If
                                                Next i
                                                snDennis.Range("L" & c.Row) = Mystr5500 & Chr(10) & Mystr4
                                                Mystr4 = ""
                                        ElseIf Left(MyCus, 3) = "UMC" Then
                                                For i = 1 To RFG5500Row
                                                        If Left(Range("A" & i), 3) = "UMC" Then
                                                                Mystr4 = Mystr4 & Range("B" & i) & Chr(10)
                                                        End If
                                                Next i
                                                snDennis.Range("L" & c.Row) = Mystr5500 & Chr(10) & Mystr4
                                                Mystr4 = ""
                                        End If
                                Next CusIndex
                        End If
                Next c
'===========================================================================================
                Case Is = "PNCL"
                        Dim PNCLRow%, MystrPNCL$, Mystr8$
                        PNCLRow = Range("A" & Rows.Count).End(xlUp).Row
                        MystrPNCL = "注意事項 : " & Chr(10)
                        For i = 1 To PNCLRow
                                If UCase(Range("A" & i)) = "ALL" And Range("A" & i).Font.Strikethrough <> True Then
                                        MystrPNCL = MystrPNCL & Range("C" & i) & Chr(10)
                                End If
                        Next i
                        For Each c In rng2
                                If c = "Pinnacle II" Then
                                        MyCus = snDennis.Range("B" & c.Row)
                                        For CusIndex = 0 To UBound(TotalCus)
                                                If MyCus = TotalCus(CusIndex) Then
                                                        For i = 1 To PNCLRow
                                                                If Range("A" & i) = TotalCus(CusIndex) Then
                                                                        Mystr8 = Mystr8 & Range("C" & i) & Chr(10)
                                                                End If
                                                        Next i
                                                        snDennis.Range("L" & c.Row) = MystrPNCL & Chr(10) & Mystr8
                                                        Mystr8 = ""
                                                End If
                                        Next CusIndex
                                End If
                        Next c
'=================================================================================
                Case Is = "AZX"                      '增加阻抗
                        Dim oMn$, oStr$, AZXarr
                        AZXarr = Split("TSMC6,TSMC3,TSMC5,TSMC8,TSMC14,TSMC15,VISC1,VISC2,UMC-所有廠區,VISC3", ",")
                        For Each c In rng2
                                If c = "AZX 63" Or c = "AZX 72" Then
                                        MyCus = snDennis.Range("B" & c.Row)
                                        oMn = Right(snDennis.Range("D" & c.Row), 3)
                                        For CusIndex = 0 To UBound(AZXarr)
                                                Dim CuzRng As Range
                                                Set CuzRng = Range("B1:P1").Find(What:=AZXarr(CusIndex), LookAt:=xlWhole)
                                                
                                                If Not CuzRng Is Nothing Then
                                                        Dim oROW%, oColumn%
                                                        oROW = CuzRng.Row
                                                        oColumn = CuzRng.Column
                                
                                                        If CuzRng.Value = MyCus And MyCus = "TSMC8" Then
                                                                If oMn = "014" Or oMn = "020" Or oMn = "036" Or oMn = "038" Or oMn = "039" Then
                                                                        snDennis.Range("N" & c.Row) = 阻抗內容(oROW, oColumn)
                                                                End If
                                                        ElseIf CuzRng.Value = MyCus And MyCus = "TSMC5" Then
                                                                If oMn = "014" Or oMn = "020" Or oMn = "036" Then
                                                                        snDennis.Range("N" & c.Row) = 阻抗內容(oROW, oColumn)
                                                                End If
                                                        ElseIf CuzRng.Value = MyCus And MyCus = "TSMC14" Then
                                                                If oMn = "037" Or oMn = "043" Then
                                                                        snDennis.Range("N" & c.Row) = 阻抗內容(oROW, oColumn)
                                                                End If
                                                        ElseIf CuzRng.Value = MyCus And MyCus = "TSMC15" Then
                                                                If oMn = "037" Or oMn = "043" Then
                                                                        snDennis.Range("N" & c.Row) = 阻抗內容(oROW, oColumn)
                                                                End If
                                                        ElseIf CuzRng.Value = MyCus And MyCus = "VISC1" Then
                                                                If oMn = "014" Or oMn = "020" Or oMn = "036" Or oMn = "038" Or oMn = "039" Then
                                                                        snDennis.Range("N" & c.Row) = 阻抗內容(oROW, oColumn)
                                                                End If
                                                        ElseIf CuzRng.Value = MyCus And MyCus = "VISC2" Then
                                                                If oMn = "014" Or oMn = "020" Or oMn = "036" Or oMn = "038" Or oMn = "039" Then
                                                                        snDennis.Range("N" & c.Row) = 阻抗內容(oROW, oColumn)
                                                                End If
                                                        ElseIf CuzRng.Value = MyCus And MyCus = "VISC3" Then
                                                                If oMn = "014" Or oMn = "020" Or oMn = "036" Or oMn = "038" Or oMn = "039" Then
                                                                        snDennis.Range("N" & c.Row) = 阻抗內容(oROW, oColumn)
                                                                End If
'======================================================================================
                                                        ElseIf Left(CuzRng.Value, 3) = Left(MyCus, 3) And Left(MyCus, 3) = "UMC" Then
                                                                If oMn = "014" Or oMn = "020" Or oMn = "036" Or oMn = "038" Or oMn = "039" Then
                                                                        snDennis.Range("N" & c.Row) = 阻抗內容(oROW, oColumn)
                                                                End If
'======================================================================================
                                                        ElseIf CuzRng.Value = MyCus And MyCus = "TSMC3" Then
                                                                If oMn = "014" Then
                                                                        snDennis.Range("N" & c.Row) = 阻抗內容(oROW, oColumn)
                                                                ElseIf oMn = "023" Or oMn = "020" Or oMn = "039" Then
                                                                        snDennis.Range("N" & c.Row) = 阻抗內容2(oROW, oColumn)
                                                                End If
                                                        ElseIf CuzRng.Value = MyCus And MyCus = "TSMC6" Then
                                                                If oMn = "014" Then
                                                                        snDennis.Range("N" & c.Row) = 阻抗內容(oROW, oColumn)
                                                                ElseIf oMn = "023" Or oMn = "020" Or oMn = "036" Or oMn = "038" Or oMn = "039" Then
                                                                        snDennis.Range("N" & c.Row) = 阻抗內容2(oROW, oColumn)
                                                                End If
                                                        End If
                                                End If
                                        Next CusIndex
                                End If
                        Next c
                End Select
        Next
        wb.Close False
        snDennis.Activate
        MsgBox "完成"
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
End Sub
Private Function 阻抗內容(ByVal oROW As Integer, ByVal oColumn As Integer) As String
        Dim oStr$
        oStr = "Shunt 電壓 : " & Cells(oROW + 2, oColumn) & "V" & Chr(10)
        oStr = oStr & "Series 電壓 : " & Cells(oROW + 3, oColumn) & "V" & Chr(10) & Chr(10)
        oStr = oStr & "阻抗點 : " & Cells(oROW + 4, oColumn) & Chr(10) & Chr(10)
        oStr = oStr & "E-chuck 正電壓 : " & Cells(oROW + 5, oColumn) & Chr(10)
        oStr = oStr & "E-chuck 負電壓 : " & Cells(oROW + 6, oColumn) & Chr(10) & Chr(10)
        oStr = oStr & "OFF SET電壓 : " & Cells(oROW + 7, oColumn) & Chr(10)
        oStr = oStr & "OFF SET電流 : " & Cells(oROW + 8, oColumn) & Chr(10)
        阻抗內容 = oStr
End Function
Private Function 阻抗內容2(ByVal oROW As Integer, ByVal oColumn As Integer) As String
        Dim oStr$
        oStr = "Shunt 電壓 : " & Cells(oROW + 2, oColumn + 1) & "V" & Chr(10)
        oStr = oStr & "Series 電壓 : " & Cells(oROW + 3, oColumn + 1) & "V" & Chr(10) & Chr(10)
        oStr = oStr & "阻抗點 : " & Cells(oROW + 4, oColumn + 1) & Chr(10) & Chr(10)
        oStr = oStr & "E-chuck 正電壓 : " & Cells(oROW + 5, oColumn) & Chr(10)
        oStr = oStr & "E-chuck 負電壓 : " & Cells(oROW + 6, oColumn) & Chr(10) & Chr(10)
        oStr = oStr & "OFF SET電壓 : " & Cells(oROW + 7, oColumn) & Chr(10)
        oStr = oStr & "OFF SET電流 : " & Cells(oROW + 8, oColumn) & Chr(10)
        阻抗內容2 = oStr
End Function








