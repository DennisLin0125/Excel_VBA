Attribute VB_Name = "Module�p�q����"
Sub �p�q����()
        
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        
        Dim snDennis As Worksheet
        Set snDennis = Workbooks("�ݭפ��R.xlsm").Worksheets("�p�q����")
     
        If [B5] <> "" Then
                Dim rg As Range
                Set rg = snDennis.Range("B5:Q" & Range("B" & Rows.Count).End(xlUp).Row)
                rg.ClearContents
                Set rg = Nothing
        End If
        
        
        Dim rng As Range, DennisRow%
        DennisRow = Range("A" & Rows.Count).End(xlUp).Row
        Set rng = snDennis.Range("A5:A" & DennisRow)
        

'=======================================================================
        Dim RMAnum As String
        Dim Year As String, fs As Object
        Dim WRAns, CompAns As Boolean

        Set fs = CreateObject("Scripting.FileSystemObject")


        For Each c In rng

                RMAnum = c.Value
                DRow = c.Row

                Year = "2021"

                For i = Year To 2006 Step -1

                        WRAns = fs.FileExists("P:\Service\RMA\WR\" & i & "\" & RMAnum & ".xls")
                        CompAns = fs.FileExists("P:\Service\RMA\Complete\" & i & "\" & RMAnum & ".xls")

                        If WRAns = True Then
                                
                                Set wb = Workbooks.Open("P:\Service\RMA\WR\" & i & "\" & RMAnum & ".xls")
                                If wb.Sheets(1).[B12] = "�s�ˬ�Ǥu�~��ϤO�椻��8��" Then
                                        snDennis.Range("B" & DRow) = "TSMC12-P1"
                                ElseIf wb.Sheets(1).[B12] = "741 �x�n��Ƕ�ϫn��_��1��1��" Then
                                        snDennis.Range("B" & DRow) = "TSMC14"
                                ElseIf wb.Sheets(1).[B12] = "�x��������Ƕ�Ϭ춮����1��" Then
                                        snDennis.Range("B" & DRow) = "TSMC15"
                                ElseIf wb.Sheets(1).[B12] = "�s�˥���Ƕ�Ϭ�s�@��9�� (3�t)" Then
                                        snDennis.Range("B" & DRow) = "TSMC3"
                                ElseIf wb.Sheets(1).[B12] = "�s�ˬ�Ƕ�϶�ϤT��121��" Then
                                        snDennis.Range("B" & DRow) = "TSMC5"
                                ElseIf wb.Sheets(1).[B12] = "741 �x�n��Ƕ�ϫn��_��1�� (6�t)" Then
                                        snDennis.Range("B" & DRow) = "TSMC6"
                                ElseIf wb.Sheets(1).[B12] = "�s�˥���Ƕ�ϤO���25�� (8�t)" Then
                                        snDennis.Range("B" & DRow) = "TSMC8"
                                        
                                ElseIf Left(wb.Sheets(1).[B11], 4) = "�p�عq�l" Then
                                        snDennis.Range("B" & DRow) = Right(wb.Sheets(1).[B11], 8)
                                        
                                ElseIf wb.Sheets(1).[B12] = "�s�ˬ�Ǥu�~��϶�ϤT��123��" Then
                                        snDennis.Range("B" & DRow) = "VISC1"
                                ElseIf wb.Sheets(1).[B12] = "�s�ˬ�Ǥu�~��ϤO���9��" Then
                                        snDennis.Range("B" & DRow) = "VISC2"
                                ElseIf wb.Sheets(1).[B12] = "��鿤Ī�˶m�n�r���@�q336�� " Then
                                        snDennis.Range("B" & DRow) = "VISC3"
                                        
                                ElseIf wb.Sheets(1).[B12] = "�s�˥����s�Ϥ��H�n��17��20�� " Then
                                        snDennis.Range("B" & DRow) = "�ӽ����"
                                ElseIf wb.Sheets(1).[B12] = "�s�˥����s�Ϥ��H�n��17��20�� " Then
                                        snDennis.Range("B" & DRow) = "�ӽ����"
                                ElseIf wb.Sheets(1).[B12] = "�s�˥����A��327��4��32�� " Then
                                        snDennis.Range("B" & DRow) = "�_�i"
                                End If
                                
                                
                                snDennis.Range("C" & DRow) = wb.Sheets(1).[F8]
                                snDennis.Range("D" & DRow) = wb.Sheets(1).[F9]
                                snDennis.Range("E" & DRow) = wb.Sheets(1).[F11]
                                
                                If wb.Sheets(1).[F10] = 3 Then
                                        snDennis.Range("F" & DRow) = "*"
                                Else
                                        snDennis.Range("F" & DRow) = ""
                                End If
                                
                                snDennis.Range("G" & DRow) = wb.Sheets(1).[E13]
                                
                                snDennis.Range("H" & DRow) = wb.Sheets(2).[K17]
                                snDennis.Range("I" & DRow) = wb.Sheets(2).[J17]
                                snDennis.Range("J" & DRow) = wb.Sheets(2).[L17]
                                
                                snDennis.Range("K" & DRow) = wb.Sheets(2).[N17]
                                snDennis.Range("L" & DRow) = wb.Sheets(2).[M17]
                                snDennis.Range("M" & DRow) = wb.Sheets(2).[O17]
                                
                                snDennis.Range("N" & DRow) = wb.Sheets(2).[M32]
                                snDennis.Range("O" & DRow) = wb.Sheets(2).[N32]
                                
                                snDennis.Range("P" & DRow) = wb.Sheets(2).[M36]
                                snDennis.Range("Q" & DRow) = wb.Sheets(2).[N36]
                                wb.Close
                                Exit For
                        ElseIf CompAns = True Then
                                
                                Set wb = Workbooks.Open("P:\Service\RMA\Complete\" & i & "\" & RMAnum & ".xls")
                                If wb.Sheets(1).[B12] = "�s�ˬ�Ǥu�~��ϤO�椻��8��" Then
                                        snDennis.Range("B" & DRow) = "TSMC12-P1"
                                ElseIf wb.Sheets(1).[B12] = "741 �x�n��Ƕ�ϫn��_��1��1��" Then
                                        snDennis.Range("B" & DRow) = "TSMC14"
                                ElseIf wb.Sheets(1).[B12] = "�x��������Ƕ�Ϭ춮����1��" Then
                                        snDennis.Range("B" & DRow) = "TSMC15"
                                ElseIf wb.Sheets(1).[B12] = "�s�˥���Ƕ�Ϭ�s�@��9�� (3�t)" Then
                                        snDennis.Range("B" & DRow) = "TSMC3"
                                ElseIf wb.Sheets(1).[B12] = "�s�ˬ�Ƕ�϶�ϤT��121��" Then
                                        snDennis.Range("B" & DRow) = "TSMC5"
                                ElseIf wb.Sheets(1).[B12] = "741 �x�n��Ƕ�ϫn��_��1�� (6�t)" Then
                                        snDennis.Range("B" & DRow) = "TSMC6"
                                ElseIf wb.Sheets(1).[B12] = "�s�˥���Ƕ�ϤO���25�� (8�t)" Then
                                        snDennis.Range("B" & DRow) = "TSMC8"
                                        
                                ElseIf Left(wb.Sheets(1).[B11], 4) = "�p�عq�l" Then
                                        snDennis.Range("B" & DRow) = Right(wb.Sheets(1).[B11], 8)
                                        
                                ElseIf wb.Sheets(1).[B12] = "�s�ˬ�Ǥu�~��϶�ϤT��123��" Then
                                        snDennis.Range("B" & DRow) = "VISC1"
                                ElseIf wb.Sheets(1).[B12] = "�s�ˬ�Ǥu�~��ϤO���9��" Then
                                        snDennis.Range("B" & DRow) = "VISC2"
                                ElseIf wb.Sheets(1).[B12] = "��鿤Ī�˶m�n�r���@�q336�� " Then
                                        snDennis.Range("B" & DRow) = "VISC3"
                                        
                                ElseIf wb.Sheets(1).[B12] = "�s�˥����s�Ϥ��H�n��17��20�� " Then
                                        snDennis.Range("B" & DRow) = "�ӽ����"
                                ElseIf wb.Sheets(1).[B12] = "�s�˥����s�Ϥ��H�n��17��20�� " Then
                                        snDennis.Range("B" & DRow) = "�ӽ����"
                                ElseIf wb.Sheets(1).[B12] = "�s�˥����A��327��4��32�� " Then
                                        snDennis.Range("B" & DRow) = "�_�i"
                                End If
                                snDennis.Range("C" & DRow) = wb.Sheets(1).[F8]
                                snDennis.Range("D" & DRow) = wb.Sheets(1).[F9]
                                snDennis.Range("E" & DRow) = wb.Sheets(1).[F11]
                                
                                If wb.Sheets(1).[F10] = 3 Then
                                        snDennis.Range("F" & DRow) = "*"
                                Else
                                        snDennis.Range("F" & DRow) = ""
                                End If
                                
                                snDennis.Range("G" & DRow) = wb.Sheets(1).[E13]
                                
                                snDennis.Range("H" & DRow) = wb.Sheets(2).[K17]
                                snDennis.Range("I" & DRow) = wb.Sheets(2).[J17]
                                snDennis.Range("J" & DRow) = wb.Sheets(2).[L17]
                                
                                snDennis.Range("K" & DRow) = wb.Sheets(2).[N17]
                                snDennis.Range("L" & DRow) = wb.Sheets(2).[M17]
                                snDennis.Range("M" & DRow) = wb.Sheets(2).[O17]
                                
                                snDennis.Range("N" & DRow) = wb.Sheets(2).[M32]
                                snDennis.Range("O" & DRow) = wb.Sheets(2).[N32]
                                
                                snDennis.Range("P" & DRow) = wb.Sheets(2).[M36]
                                snDennis.Range("Q" & DRow) = wb.Sheets(2).[N36]
                                wb.Close
                                Exit For
                        End If
                Next i
        Next c
        MsgBox "�d�ߧ���"
        Application.DisplayAlerts = True
End Sub


