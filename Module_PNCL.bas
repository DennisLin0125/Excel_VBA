Attribute VB_Name = "ModulePNCL"
Option Explicit
Sub PNCL���i()
        
        If ActiveSheet.name <> "RMA" Then
                MsgBox "�Ш�RMA��������", vbCritical
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
        MystrKT(0) = "�G�ٽT�{ :" & vbCrLf
        
        MystrKT(1) = "���פ��e :"
        MystrKT(2) = "1. �ˬdAux board F1~F8�K OK"
        MystrKT(3) = "2. �ˬdOutput bridge�K OK"
        MystrKT(4) = "3. �ˬdInverter board�KOK "
        MystrKT(5) = "4. �ˬdInter connect board�K NG"
        MystrKT(6) = Space(4) & "A,B side��Inter connect board (1303318R) x6"
        MystrKT(7) = Space(4) & "C3�BC5 (9215027)�l�a�A�q����0uF,�зǬ�0.18uF"
        MystrKT(8) = Space(4) & "�ҥH�󴫡C"
        
        MystrKT(9) = "5. �ˬdcap board�K OK"
        MystrKT(10) = "6. �ˬdAC input section �K NG"
        MystrKT(11) = Space(4) & "Bridge����A�ҥH�󴫡A"
        MystrKT(12) = Space(4) & "Bridge (1501225) x1"
        MystrKT(13) = Space(4) & "Contactor (3301189-R) x1"
        MystrKT(14) = Space(4) & "Breaker (3341029) x1"
        MystrKT(15) = Space(4) & "�ѩ󭷮��l�a�i��y��Bridge �l�a,"
        MystrKT(16) = Space(4) & "�G�w���ʧ�Bridge (1501225) x1"
        
        MystrKT(17) = "7. �ˬd����(3311020)�K OK"
        MystrKT(18) = "8. �q��Aux board�q��: OK"
        MystrKT(19) = "9. �e�q�ˬdlogic board (1303357)�K NG"
        MystrKT(20) = Space(4) & "�o�{�n��Ѽ�run-time�Bidle-time�ɶ�����-1"
        MystrKT(21) = Space(4) & "�L�k�p�ƭp�ɡA��logic board��Nov-ram"
        MystrKT(22) = "10. �ϥ�User port �s�u�K OK "
        MystrKT(23) = "11. ����Aux�BWater�BVac�T��Interlock�K OK "
        MystrKT(24) = "12. ARC test (Open)�K OK "
        MystrKT(25) = "13. ����Master/Slave�s�u�A�ÿ�X�q���B�q�y�B�\�v�K OK "
        MystrKT(26) = "14. �̼зǧ󴫷ŷP�u(1341338-01) x 2 "
        MystrKT(27) = "15. Logic����7421419K.00 "
        MystrKT(28) = Space(4) & " Config����7202177D.00"
        MystrKT(29) = "16. 17.5KW�����ɡA�|�q�y��(�a�u): 6mA"
        MystrKT(30) = "17. �̫��ˬd: Jack "
        
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
                        MsgBox "�п� " & ActiveSheet.name & " (�i�ƿ�)"
                        Call �K�W�l�a�Ӥ�
                        PNCLError.Show
                         
                Case Is = "�i�X�t�Ӥ�"
                        MsgBox "�п� " & ActiveSheet.name & " (�i�ƿ�)"
                        Call �K�W�i�X�t�Ϥ�
                        
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
        MsgBox "����"
End Sub

