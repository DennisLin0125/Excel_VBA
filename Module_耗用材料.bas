Attribute VB_Name = "Module�ӥΧ���"
Sub �ӥΧ���()
Attribute �ӥΧ���.VB_ProcData.VB_Invoke_Func = "m\n14"

        Application.ScreenUpdating = False '������s�e��
        Application.DisplayAlerts = False
        
        If ActiveSheet.name <> "RMA" Then
                MsgBox "�Ш�RMA��������", vbCritical
                Exit Sub
        End If
        
        Worksheets("RMA").Select
        Dim sh As Worksheet
        For Each sh In ActiveWorkbook.Worksheets
                sh.Select
                If sh.name = "User parts" Then
                        Worksheets("User parts").name = "Use parts"
                End If
        Next
        
        ActiveWorkbook.Save
        
        ActiveWindow.TabRatio = 0.9
        
        Dim Mainwb As Workbook
        Set Mainwb = ActiveWorkbook
        
        Dim mainSh As Worksheet
        Set mainSh = Mainwb.Worksheets("RMA")
        
        Dim MainUser As Worksheet
        Set MainUser = Mainwb.Worksheets("Use parts")
        
        Dim wb As Workbook
        Set wb = Workbooks.Open("P:\Service\Service Bulletin\Parts Description.xlsx")
        With wb
                Dim myTempParts
                myTempParts = .Worksheets("part description").Cells(1, 1).CurrentRegion
        End With

        mainSh.Activate
        
        Dim W3M As Integer, Customer As String, Myname As String, oW3M As Integer
        W3M = Range("F10")
        oW3M = Range("F41")
        Customer = Range("B11")
        Myname = ActiveWorkbook.name
        
        If Range("B42").Value = "0" Then
                MsgBox ("�п�J�u��"), vbCritical
                Exit Sub
        ElseIf Range("F11").Value = "" Then
                MsgBox ("�п�J�u�{�v�W�r"), vbCritical
                Exit Sub
        ElseIf Range("H12").Value = "" Then
                MsgBox ("�п�J�G�ٽT�{ Yes �� No"), vbCritical
                Exit Sub
        ElseIf Range("B41").Value = "" Then
                MsgBox ("�п�J�����ɶ�"), vbCritical
                Exit Sub
        ElseIf MainUser.[B1] = "" Then
                MsgBox ("�п�JUse parts ���פu��"), vbCritical
                MainUser.Select
                [B1].Select
                Exit Sub
        ElseIf Range("F41").Value = 3 And Range("F42").Value <> "" Then
                MsgBox ("���x���O�T,�u�ɼg����m"), vbCritical
                Exit Sub
        End If
        
        If W3M = 3 Then
                a = MsgBox("���x���O�T,�нT�{�O�_����W3M���R���i", vbYesNo)
                If a = vbYes Then
                        a = MsgBox("�O�_��NPO??", vbYesNo)
                                If a = vbYes Then
                                        Range("D42").Value = 2
                                        Range("A19").Value = "1.No problem observed."
                                        Worksheets("Use parts").[B1] = 2
                                        Range("H12").Value = "No"
                                        Range("H41").Value = 4
                                End If
                Else
                        MsgBox ("�Ч������R���i")
                        Exit Sub
                End If
        End If
        
        Dim fs
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        Dim i As Integer, ans As Boolean, dataSize As Double, x As Double
        For i = 2022 To 2006 Step -1
                ans = fs.FileExists("P:\Service\RMA\WR\" & i & "\" & Myname)
                If ans Then
                        dataSize = (VBA.FileLen("P:\Service\RMA\WR\" & i & "\" & Myname))
                        dataSize = dataSize / 1024 / 1024
                        x = Math.Round(dataSize, 2)
                        If x > 2.5 Then
                                MsgBox "�ɮפj�p :  " & x & " MB" & Chr(10) & Chr(10) & "�ɮ׶W�L2.5�ۢСA�����Y�δ�ַӤ��I�I", vbCritical
                                Exit Sub
                        End If
                        Exit For
                End If
                If i = 2006 Then
                        MsgBox Range("F7") & "�ɮפ��s�b"
                        Exit Sub
                End If
        Next
        
        If Range("B46") <> "" Then
        
                Dim rg As Range, c As Range
                Set rg = Mainwb.ActiveSheet.Range("B46:B100")
                
                For Each c In rg
                        If c = "" Then
                                Dim oROW%, rgTemp As Range
                                oROW = c.Row - 1
                                Set rgTemp = ActiveSheet.Range("B46:B" & oROW)
                                Exit For
                        End If
                Next
                
                Dim myError%, tempRow%
                myError = 1
                
                For Each c In rgTemp
                        Dim Parts As Range, endRow As Long, tempPart As String
                        endRow = UBound(myTempParts, 1)
                        tempRow = c.Row
                        Set Parts = wb.Worksheets("part description").Range("A1:A" & endRow).Find(What:=c, LookAt:=xlWhole)
                        If Not Parts Is Nothing Then
'                                Dim myRow As Long
'                                myRow = Parts.Row
'                                mainSh.Range("C" & tempRow) = wb.Worksheets("part description").Range("B" & myRow)
                        Else
                                MsgBox "RMA��  �Ƹ� :  " & c & "  ���`", vbCritical
                                Exit Sub
                        End If
                Next
                
                Set rg = Nothing
                Set Parts = Nothing
        End If
'************************************************************************************************
        Sheets("Use parts").Select
        If Range("A4") <> "" Then
                Set rg = MainUser.Range("A4:A100")
                For Each c In rg
                        If c = "" Then
                                oROW = c.Row - 1
                                Set rgTemp = MainUser.Range("A4:A" & oROW)
                                Exit For
                        End If
                Next
                
                For Each c In rgTemp
                        endRow = UBound(myTempParts, 1)
                        tempRow = c.Row
                        Set Parts = wb.Worksheets("part description").Range("A1:A" & endRow).Find(What:=c, LookAt:=xlWhole)
                        If Not Parts Is Nothing Then
'                                MainUser.Range("A" & tempRow).Interior.Color = xlNone
'                                MainUser.Range("D" & tempRow) = ""
                        Else
                                MsgBox "Use parts�� �Ƹ� :  " & c & "  ���`", vbCritical
                                Exit Sub
                        End If
                Next
        End If
        
        wb.Close False
        
        Set Parts = Nothing
        Set rgTemp = Nothing
        Set rg = Nothing
        Set Mainwb = Nothing
        Set mainSh = Nothing
        Set MainUser = Nothing
        Set wb = Nothing

        ActiveWorkbook.Save
        Worksheets("RMA").Select
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        
        MsgBox "�ɮפj�p :  " & x & " MB" & Chr(10) & Chr(10) & "�Ƹ� ���`", vbInformation
        
End Sub
