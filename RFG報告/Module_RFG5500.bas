Attribute VB_Name = "ModuleRFG5500"
Option Explicit
Sub RFG5500���i()
        
        If ActiveSheet.name <> "RMA" Then
                MsgBox "�Ш�RMA��������", vbCritical
                Exit Sub
        End If
        
        [F11] = "Dennis"
        [H12] = "Yes"
        [F42] = "12"
        [B41] = "2"
        [D41] = Date
        
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
        
        Worksheets("�i�X�t�Ӥ�").Copy Before:=Worksheets("�i�X�t�Ӥ�")
        Worksheets("�i�X�t�Ӥ�").Copy Before:=Worksheets("�i�X�t�Ӥ�")
        Worksheets("�i�X�t�Ӥ� (2)").name = "Failure Photo (Master)"
        Worksheets("�i�X�t�Ӥ� (3)").name = "Failure Photo (Slave)"
        
        Dim sh As Worksheet
        For Each sh In ActiveWorkbook.Worksheets
                sh.Select
                Select Case sh.name
                Case Is = "Test Table RF"
                        Dim Power(9) As Integer, i%, oROW%
                        oROW = 0
                        For i = 500 To 5000 Step 500
                                Power(oROW) = i
                                oROW = oROW + 1
                        Next i
                        [C22].Resize(oROW, 1) = Application.WorksheetFunction.Transpose(Power)
                        [E33] = "74000348"
                        [E34] = "49.1"
                        
                        MsgBox "�п��2�i�i�ι�"
                        Call �i�ι�(37, 1)
                Case Is = "Failure Photo"
                        ActiveSheet.name = "Failure Photo (5500)"
                        MsgBox "�пﱱ��O�MSN�Ӥ�"
                        Call �K�W�l�a�Ӥ�
                        
                        With [A38:D38].Borders
                                .LineStyle = xlContinuous
                        End With
                        
                        With [E38:H38].Borders
                                .LineStyle = xlContinuous
                        End With
                        
                        With [A38:D38]
                                .Merge
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                        End With
                        
                        With [E38:H38]
                                .Merge
                                .HorizontalAlignment = xlCenter
                                .VerticalAlignment = xlCenter
                        End With
                        
                        [A38] = "The control board was failed."
                        [E38] = "Replaced the failed parts."
                        
                        With [A38:E38].Font
                            .name = "Tahoma"
                            .Size = 12
                        End With
                Case Is = "Failure Photo (Master)"
                        MsgBox "�п� " & ActiveSheet.name & " (�i�ƿ�)"
                        Call �K�W�l�a�Ӥ�
                        RFGError.Show
                        
                 Case Is = "Failure Photo (Slave)"
                        MsgBox "�п� " & ActiveSheet.name & " (�i�ƿ�)"
                        Call �K�W�l�a�Ӥ�
                        RFGError.Show

                Case Is = "�i�X�t�Ӥ�"
                        MsgBox "�п� " & ActiveSheet.name & " (�i�ƿ�)"
                        Call �K�W�i�X�t�Ϥ�
                        
                End Select
        Next sh
        Worksheets("RMA").Select
        MsgBox "����"
End Sub
