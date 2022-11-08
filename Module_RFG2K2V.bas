Attribute VB_Name = "ModuleRFG2K2V"
Option Explicit
Sub RFG2K2V報告()
        
        If ActiveSheet.name <> "RMA" Then
                MsgBox "請到RMA頁面執行", vbCritical
                Exit Sub
        End If
'************************************************************************
        Dim DEinput, DEoutput, DEtemp
        
        DEinput = InputBox("請輸入D/E  輸入的 VPP")
        DEinput = DEinput * 10
        
        DEoutput = InputBox("請輸入D/E  輸出的 VPP")
        DEoutput = DEoutput * 10
        
        DEtemp = Math.Round(DEoutput / DEinput, 2)
        
        
        Dim MystrKTRepair(4) As String
        MystrKTRepair(0) = "1. The PA were failed."
        MystrKTRepair(1) = "2. The dinner plate was failed."
        MystrKTRepair(2) = "3. The aux board was failed."
        MystrKTRepair(3) = "4. The bypass board was failed."
        MystrKTRepair(4) = "5. The control cable was failed in inverter section."
        
        [A19] = Join(MystrKTRepair, vbCrLf)
 '****************************************************************************
        Dim MystrKT(38) As String
        MystrKT(0) = "故障確認 :" & vbCrLf
        
        MystrKT(1) = "維修內容 :"
        MystrKT(2) = "1. 檢查KT PA(K8705109-01) -- NG"
        MystrKT(3) = Space(4) & "＞PA短路,更換KT PA新品. (x2 K8705109-01)"
        
        MystrKT(4) = "2. 檢查Drv/extr(8705006) -- OK"
        
        MystrKT(5) = "3. 檢查Dinner plate(8705012R) -- NG "
        MystrKT(6) = Space(4) & "＞依標準補焊. (8705012R)"
        
        MystrKT(7) = "4. 檢查Clamp(8705014)(8705015) -- OK"
        
        MystrKT(8) = "5. 檢查Interconnect board(1315207R) -- OK"
        MystrKT(9) = "6. 檢查Measurement board(1315032R) -- OK"
        
        MystrKT(10) = "7. 檢查Bypass board(R1305164)電容使用超過4年 -- NG"
        MystrKT(11) = Space(4) & "＞依標準更換. (R1305164)"
        
        MystrKT(12) = "8. 檢查Inverter section排線, Overhaul更換新品. "
        MystrKT(13) = Space(4) & "＞ (K1345701-00)(1345324)"
        MystrKT(14) = ""
        
        MystrKT(15) = "1. 檢查AC input section -- OK"
        MystrKT(16) = "2. 檢查Aux board(1310009-06R) -- NG"
        MystrKT(17) = Space(4) & "＞ J3,J4碳化,更換新品. (x2 3501217)"
        MystrKT(18) = Space(4) & "＞ RV1 thermistor 25Ω焦黑,更換新品. (1191003) "
        MystrKT(19) = Space(4) & "＞ R15 0.1Ω焦黑,更換新品. (1141028)"
        MystrKT(20) = Space(4) & "＞ C17 2.2u損壞,更換新品. (1251044)"
        MystrKT(21) = Space(4) & "＞ C23 1u損壞,更換新品. (1261036)"
        
        MystrKT(22) = "3. 檢查Control board(1305251R) -- OK"
        MystrKT(23) = "4. 檢查Phase control board(1305362R) -- OK"
        MystrKT(24) = "5. 檢查Inverter board(1305787R) -- OK"
        MystrKT(25) = "6. 檢查Isotop board(1305340R) -- OK"
        MystrKT(26) = ""
        
        MystrKT(27) = "檢測項目 :"
        MystrKT(28) = "1. 校正PA(KT)工作電壓  0.3V、電流及相位 -- ok "
        MystrKT(29) = "2. 量測Drv/extr, input " & DEinput & " Vpp; Output " & DEoutput & " Vpp,"
        MystrKT(30) = Space(4) & "放大 " & DEtemp & " 倍."
        
        MystrKT(31) = "3. 量測Aux board輸出電壓 : "
        MystrKT(32) = Space(4) & "＞ +30V -> +30.06V"
        MystrKT(33) = Space(4) & "＞ +24V -> +23.98V"
        MystrKT(34) = Space(4) & "＞ +5V   -> +5.002V "
        MystrKT(35) = Space(4) & "＞ -15V  ->  -14.96V"
        MystrKT(36) = Space(4) & "＞ +15V -> +14.98V"
        
        MystrKT(37) = "4. 確認Full power 2000W -- 在+/- 0.5%範圍內 "
        MystrKT(38) = "5. 最後檢查: Jack "
        
        [J19] = Join(MystrKT, vbCrLf)
'**************************************************************************************************
        [F11] = "Dennis"
        [H12] = "Yes"
        [F42] = "10"
        [B41] = "2"
        [D41] = Date
        
        [B46] = "K8705109-01"
        [B47] = "R1305164"
        [B48] = "K1345701-00"
        [B49] = "1345324"
        [B50] = "1310009-06R"
        
        [G46] = 2
        [G47:G50] = 1
        
        If [H9] = "" Then
                [H9] = "=H8"
                [H10] = "=H8"
        Else
                [H10] = "=H9"
        End If
        
        Dim myStr(3) As String
        myStr(0) = "1. Machine cleaning."
        myStr(1) = "2. Replace fail parts."
        myStr(2) = "3. According the test proccedure tested --- pass."
        myStr(3) = "4. Burn-in."
        
        [A33] = Join(myStr, vbCrLf)
        
        Dim sh As Worksheet
        Dim TestEquipmentSN As Double, LoadResistance As Double
        TestEquipmentSN = "74000348"
        LoadResistance = "49.1"
        For Each sh In ActiveWorkbook.Worksheets
                sh.Select
                Select Case sh.name
                Case Is = "Test Table RF"
                        Dim Power(9) As Integer, i%, oROW%
                        oROW = 0
                        For i = 200 To 2000 Step 200
                                Power(oROW) = i
                                oROW = oROW + 1
                        Next i
                        [C22].Resize(oROW, 1) = Application.WorksheetFunction.Transpose(Power)
                        [E33] = TestEquipmentSN
                        [E34] = LoadResistance
                        
                        MsgBox "請選擇2張波形圖"
                        Call 波形圖(37, 1)
                Case Is = "Failure Photo"
                        MsgBox "請選 " & ActiveSheet.name & " (可複選)"
                        Call 貼上損壞照片
                        RFGError.Show
                        
                Case Is = "2   1"
                        [E33] = TestEquipmentSN
                        [E34] = LoadResistance
                        
                 Case Is = " 3   1"
                        [E33] = TestEquipmentSN
                        [E34] = LoadResistance
                        
                        [D30:F31,H30:H31] = "*"
                        
                Case Is = "進出廠照片"
                        MsgBox "請選 " & ActiveSheet.name & " (可複選)"
                        Call 貼上進出廠圖片
                        
                Case Is = "Use parts"
                        [B1] = 8
                        
                        Dim temp(9) As String
                        Dim temp2(9) As String
                        
                        temp(0) = "9214036"
                        temp(1) = "9215032"
                        temp(2) = "3501217"
                        temp(3) = "1191003"
                        temp(4) = "1251044"
                        temp(5) = "1261036"
                        temp(6) = "K8705109-01"
                        
                        temp(7) = "K1345701-00"
                        temp(8) = "1345324"
                        temp(9) = "R1305164"
                        
                        [A4].Resize(UBound(temp) + 1) = Application.WorksheetFunction.Transpose(temp)
                        
                        temp2(0) = "1"
                        temp2(1) = "2"
                        temp2(2) = "2"
                        temp2(3) = "1"
                        temp2(4) = "1"
                        temp2(5) = "1"
                        temp2(6) = "2"
                        
                        temp2(7) = "1"
                        temp2(8) = "1"
                        temp2(9) = "1"
                        
                        [B4].Resize(UBound(temp2) + 1) = Application.WorksheetFunction.Transpose(temp2)
                        
                        [C4:C9] = "Aux board"
                        [C10] = "PA"
                        [C11:C12] = " Inverter section"
                        [C13] = "Bypass board"
                        
                End Select
        Next sh
        Worksheets("RMA").Select
        MsgBox "完成"
End Sub

