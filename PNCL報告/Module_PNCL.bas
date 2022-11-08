Attribute VB_Name = "ModulePNCL"
Option Explicit
Sub PNCL報告()
        
        If ActiveSheet.name <> "RMA" Then
                MsgBox "請到RMA頁面執行", vbCritical
                Exit Sub
        End If
'************************************************************************
        Dim MystrKTRepair(3) As String
        MystrKTRepair(0) = "1. The input section was failed."
        MystrKTRepair(1) = "2. The power amplification section was failed."
        MystrKTRepair(2) = "3. The control section was failed."
        MystrKTRepair(3) = "4. The fan was failed."
        
        [A19] = Join(MystrKTRepair, vbCrLf)
 '****************************************************************************
        Dim MystrKT(30) As String
        MystrKT(0) = "故障確認 :" & vbCrLf
        
        MystrKT(1) = "維修內容 :"
        MystrKT(2) = "1. 檢查Aux board F1~F8… OK"
        MystrKT(3) = "2. 檢查Output bridge… OK"
        MystrKT(4) = "3. 檢查Inverter board…OK "
        MystrKT(5) = "4. 檢查Inter connect board… NG"
        MystrKT(6) = Space(4) & "A,B side的Inter connect board (1303318R) x6"
        MystrKT(7) = Space(4) & "C3、C5 (9215027)損壞，量測為0uF,標準為0.18uF"
        MystrKT(8) = Space(4) & "所以更換。"
        
        MystrKT(9) = "5. 檢查cap board… OK"
        MystrKT(10) = "6. 檢查AC input section … NG"
        MystrKT(11) = Space(4) & "Bridge炸燬，所以更換，"
        MystrKT(12) = Space(4) & "Bridge (1501225) x1"
        MystrKT(13) = Space(4) & "Contactor (3301189-R) x1"
        MystrKT(14) = Space(4) & "Breaker (3341029) x1"
        MystrKT(15) = Space(4) & "由於風扇損壞可能造成Bridge 損壞,"
        MystrKT(16) = Space(4) & "故預防性更換Bridge (1501225) x1"
        
        MystrKT(17) = "7. 檢查風扇(3311020)… OK"
        MystrKT(18) = "8. 量測Aux board電壓: OK"
        MystrKT(19) = "9. 送電檢查logic board (1303357)… NG"
        MystrKT(20) = Space(4) & "發現軟體參數run-time、idle-time時間均為-1"
        MystrKT(21) = Space(4) & "無法計數計時，更換logic board的Nov-ram"
        MystrKT(22) = "10. 使用User port 連線… OK "
        MystrKT(23) = "11. 測試Aux、Water、Vac三種Interlock… OK "
        MystrKT(24) = "12. ARC test (Open)… OK "
        MystrKT(25) = "13. 測試Master/Slave連線，並輸出電壓、電流、功率… OK "
        MystrKT(26) = "14. 依標準更換溫感線(1341338-01) x 2 "
        MystrKT(27) = "15. Logic版本7421419K.00 "
        MystrKT(28) = Space(4) & " Config版本7202177D.00"
        MystrKT(29) = "16. 17.5KW熱機時，漏電流值(地線): 6mA"
        MystrKT(30) = "17. 最後檢查: Jack "
        
        [J19] = Join(MystrKT, vbCrLf)


'**************************************************************************************************
        [F11] = "Dennis"
        [H12] = "Yes"
        [F42] = "7"
        [B41] = "1"
        [D41] = Date
        
        [B46] = "1303318R"
        [B47] = "1303357R"
        [B48] = "3311020"

        
        [G46] = 6
        [G47] = 1
        [G48] = 2
        
        If [H9] = "" Then
                [H9] = "=H8"
                [H10] = "=H8"
        Else
                [H10] = "=H9"
        End If
        
        Dim myStr(3) As String
        myStr(0) = "1. Check and replace all failed parts."
        myStr(1) = "2. According the test procedure tested."
        myStr(2) = "3. Test Aebus card and user port."
        myStr(3) = "4. Burn-in one hour."
        
        [A33] = Join(myStr, vbCrLf)
        
        Dim sh As Worksheet
        
        For Each sh In ActiveWorkbook.Worksheets
                sh.Select
                Select Case sh.name
                Case Is = "Test Table DC"
                        PNCL20K
                        [G21] = "S"
                        [H21] = "M"
                        
                        [G22:H23] = "N"
                        [G24:H24] = "20K"
                        [G25:H25] = 1
                        
                        [G34:H34] = 150
                        
                        [G41:H41] = 50
                        [G42:H42] = 0
                        
                Case Is = "Failure Photo"
                        MsgBox "請選 " & ActiveSheet.name & " (可複選)"
                        Call 貼上損壞照片
                        PNCLError.Show
                         
                Case Is = "進出廠照片"
                        MsgBox "請選 " & ActiveSheet.name & " (可複選)"
                        Call 貼上進出廠圖片
                        
                Case Is = "Use parts"
                        
                        [B1] = 7
                        
                        Dim temp(3) As String
                        Dim temp2(9) As String
                        
                        temp(0) = "9215027"
                        temp(1) = "3311020"
                        temp(2) = "1341338-01"
                        [A4].Resize(UBound(temp) + 1) = Application.WorksheetFunction.Transpose(temp)
                        
                        temp2(0) = "12"
                        temp2(1) = "2"
                        temp2(2) = "2"
                        [B4].Resize(UBound(temp2) + 1) = Application.WorksheetFunction.Transpose(temp2)
                        
                        [C4] = "Inter connect board"
                        [C5] = "FAN"
                        [C6] = "Thermo sensor"
                        
                End Select
        Next sh
        Worksheets("RMA").Select
        MsgBox "完成"
End Sub

