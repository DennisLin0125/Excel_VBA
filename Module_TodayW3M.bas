Attribute VB_Name = "ModuleTodayW3M"
Option Explicit
Sub �j�M����O�T()
Attribute �j�M����O�T.VB_ProcData.VB_Invoke_Func = "w\n14"
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        Dim wb As Workbook, fpathThisYear$, fpathBeforeYear$, myDate$, Year%
        
        myDate = Format(Date, "yyyymmdd")
        
        Year = Mid(myDate, 1, 4) '���~��
        
        fpathThisYear = "P:\Service\RMA\Main\Kaitek RMA " & Year & " main.xls"
        fpathBeforeYear = "P:\Service\RMA\Main\Kaitek RMA " & Year - 1 & " main.xls"
        
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
        LF = Range("A" & Rows.Count).End(xlUp).Row  '�̫�@�C
        
        Dim i%
        For i = 1 To LF  '�M�MMain��
                If InStr(Cells(i, "C"), myDay) Then   '�p���Ѱe�ץx��
                        Total = Total + 1
                End If
        Next i
        
        For i = 1 To LF   '�M�MMain��
                If InStr(Cells(i, "C"), myDay) * InStr(Cells(i, "Q"), "3") Then '��ѬO�_���O�T
                        Dim tempSn$, tempCust$
                        tempSn = Cells(i, "K")    '  �ȦsSN
                        tempCust = Cells(i, "D") '  �Ȧs�Ȥ�W��
                        Dim rngSn As Range
                        Set rngSn = Range("K1:K" & i - 1).Find(What:=tempSn, After:=Cells(i - 1, "K"), LookAt:=xlWhole, SearchDirection:=xlPrevious) '�j�M�O�_���O�T
                        
                        If Not rngSn Is Nothing Then
                                Dim myRow%
                                myRow = rngSn.Row
                                If Range("D" & myRow) = tempCust Then      '�T�{�O�T��SN�O�P�Ӽt��
                                        Dim engineer$, LifeTime As Double, RMA$, Reson$
                                        engineer = Cells(myRow, "T")
                                        LifeTime = Date - Cells(myRow, "P")
                                        RMA = Cells(myRow, "A")
                                        Reson = Cells(myRow, "Y")
                                ElseIf Left(Range("D" & myRow), 3) = "UMC" And Left(tempCust, 3) = "UMC" Then  '�P�_�O�_��UMC
                                        engineer = Cells(myRow, "T")
                                        LifeTime = Date - Cells(myRow, "P")
                                        RMA = Cells(myRow, "A")
                                        Reson = Cells(myRow, "Y")
                                Else
                                        Set firstRng = rngSn
                                        Do                                     '��U�@�ӫO�T
                                                Set rngSn = Range("K1:K" & i - 1).FindPrevious(rngSn)
                                                myRow = rngSn.Row
                                                If Range("D" & myRow) = tempCust Then
                                                        engineer = Cells(myRow, "T")
                                                        LifeTime = Date - Cells(myRow, "P")
                                                        RMA = Cells(myRow, "A")
                                                        Reson = Cells(myRow, "Y")
                                                        Exit Do
                                                ElseIf Left(Range("D" & myRow), 3) = "UMC" And Left(tempCust, 3) = "UMC" Then  '�P�_�O�_��UMC
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
                                        If Range("D" & myRow) = tempCust Then   '�T�{�O�T��SN�O�P�Ӽt��
                                                engineer = Cells(myRow, "T")
                                                LifeTime = Date - Cells(myRow, "P")
                                                RMA = Cells(myRow, "A")
                                                Reson = Cells(myRow, "Y")
                                        ElseIf Left(Range("D" & myRow), 3) = "UMC" And Left(tempCust, 3) = "UMC" Then  '�P�_�O�_��UMC
                                                engineer = Cells(myRow, "T")
                                                LifeTime = Date - Cells(myRow, "P")
                                                RMA = Cells(myRow, "A")
                                                Reson = Cells(myRow, "Y")
                                        Else
                                                Set firstRng = rng
                                                Do                                     '��U�@�ӫO�T
                                                        Set rng = Range("K1:K" & LF2).FindPrevious(rng)
                                                        myRow = rng.Row
                                                        If Range("D" & myRow) = tempCust Then
                                                                engineer = Cells(myRow, "T")
                                                                LifeTime = Date - Cells(myRow, "P")
                                                                RMA = Cells(myRow, "A")
                                                                Reson = Cells(myRow, "Y")
                                                                Exit Do
                                                        ElseIf Left(Range("D" & myRow), 3) = "UMC" And Left(tempCust, 3) = "UMC" Then  '�P�_�O�_��UMC
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
                                MsgBox "��� : " & myDay & Chr(10) & Chr(10) & "�@���� " & Total & " �x����" & Chr(10) & Chr(10) & "���馳�O�T"
                        End If
                        myStr = "����� " & times & " �x�O�T" & Chr(10) & Chr(10)
                        myStr = myStr & "�O�T�u�{�v : " & engineer & Chr(10) & Chr(10)
                        myStr = myStr & "�Ȥ� : " & Cells(i, "D") & Chr(10) & Chr(10)
                        myStr = myStr & "MN : " & Cells(i, "I") & Chr(10) & Chr(10)
                        myStr = myStr & "SN : " & Cells(i, "K") & Chr(10) & Chr(10)
                        myStr = myStr & "���� : " & Cells(i, "H") & Chr(10) & Chr(10)
                        
                        myStr = myStr & "�e��RMA : " & RMA & Chr(10) & Chr(10)
                        myStr = myStr & "�o��RMA : " & Cells(i, "A") & Chr(10) & Chr(10)
                        
                        myStr = myStr & "�e���U����] : " & Reson & Chr(10) & Chr(10)
                        myStr = myStr & "�����U����] : " & Cells(i, "Y") & Chr(10) & Chr(10)
                        
                        myStr = myStr & "LifeTime :  " & LifeTime & " ��"
                        times = times + 1
                        MsgBox myStr
                End If
        Next i
        
        If times = 1 Then
                MsgBox "��� : " & myDay & Chr(10) & Chr(10) & "�@���� " & Total & " �x����" & Chr(10) & Chr(10) & "�ثe�S���O�T"
        End If
        
        wb.Close False
        Set wb = Nothing
        
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
End Sub


