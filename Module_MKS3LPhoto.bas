Attribute VB_Name = "ModuleMKS3LPhoto"
Sub �ˬdMKS3L�Ӥ�()
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        Dim wb As Workbook, sPath$
        sPath = "P:\Service\�޳N�Q�ױM��\Engineer\4. Jacky\MKS 3L.xlsx"
        Set wb = Workbooks.Open(sPath, UpdateLink = 0)
        
        wb.Worksheets("�`��").Activate
        
        oROW = Range("B2").End(xlDown).Row
        
        Dim rg As Range
        Set rg = wb.Worksheets("�`��").Range("B3:B" & oROW)
        
       
        Dim sh As Worksheet
        For Each sh In wb.Worksheets
        
                sh.Select
                
                Select Case sh.name
                
                Case Is = "VISC1", "VISC2", "VISC3", "TSMC3", "TSMC5", "TSMC6", "TSMC8", "TSMC14", "AUO-L3D", "UMC-8F"
                
                        Dim i%
                        For i = 2 To 200 Step 13
                                If Cells(3, i) <> "" Then
                                        RMA = Cells(3, i)
                                        
                                        Dim num As Range
                                        Set num = rg.Find(What:=RMA, LookAt:=xlWhole)
                                        
                                        If num Is Nothing Then
                                                MsgBox "�`��̨S�� " & sh.name & " �� " & RMA, vbCritical
                                                Application.ScreenUpdating = True
                                                Cells(3, i).Select
                                                Set num = Nothing
                                                Exit Sub
                                        End If
                                        
                                End If
                        Next
                End Select
        Next
        
        wb.Close False
        Set rg = Nothing
        Set wb = Nothing
        
        MsgBox "�ɮק��s�b", vbInformation
        
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True

End Sub
