Attribute VB_Name = "ModuleAZX���i"
Option Explicit
Sub �i�ι�(ByVal oROW As Integer, ByVal colum As Integer)
        Dim fd As FileDialog, myWidth%, myHeight%, sPath, iTop
        Set fd = Application.FileDialog(msoFileDialogFilePicker)

        With fd
                .AllowMultiSelect = True
                .Title = "�п�ܷӤ�"
                .ButtonName = "�N�O�A�F!!!!"
        
                myWidth = 395
                myHeight = 295
                
                Dim rng As Range
                Dim sShape As Shape
                
                If .Show Then
                        For Each sPath In .SelectedItems
                                Set rng = Cells(oROW, colum)
                                Set sShape = ActiveSheet.Shapes.AddPicture(sPath, msoFalse, msoCTrue, rng.Left, iTop, myWidth, myHeight)
                                sShape.Cut
                                rng.Select
                                ActiveSheet.Paste
                                colum = colum + 4
                                Set rng = Nothing
                                Set sShape = Nothing
                        Next
                End If
        End With
        Set fd = Nothing
End Sub

Sub AZX���i()
        
        Dim myTime As Date
        myTime = Time

        If ActiveSheet.name <> "RMA" Then
                MsgBox "�Ш�RMA��������", vbCritical
                Exit Sub
        End If
        
        If [F10] = 2 Then
                AzxNormal.Show
                [H12] = "Yes"
                [F42] = "6"
                [B41] = "0.5"
                [D41] = Date
        Else
                AZXW3M.Show
        End If
        
        AZXRMA.Show
        
        Dim cus$
        cus = [B12]
        
        If [H9] = "" Then
                [H9] = "=H8"
                [H10] = "=H8"
        Else
                [H10] = "=H9"
        End If
        
        Dim sh As Worksheet
        
        For Each sh In ActiveWorkbook.Worksheets
                sh.Select
                Select Case sh.name
                Case Is = "Test Table Tuner (-020,-023)", "Test Table Tuner (-036,-039)", "Test Table Tuner (-014)", "Test Table Tuner-014"
                
                        If cus = "�s�˥���Ƕ�ϤO���25�� (8�t)" Then
                                AZX���.Show
                                Dim IdleV$, IdleI$, AfterSelect2$
                                IdleV = [K36]
                                IdleI = [L36]
                        
                                AfterSelect2 = [P36]
                                
                                MsgBox "T8���2�i�v�K����"
                                Call �i�ι�(37, 1)
                                
                                Worksheets("�i�X�t�Ӥ�").Copy Before:=Sheets("Failure Photo")
                                Worksheets("�i�X�t�Ӥ� (2)").name = "Failure Photo(�Ȥ�)"
                                Worksheets("Failure Photo(�Ȥ�)").Copy Before:=Sheets("Failure Photo")
                                Worksheets("Failure Photo(�Ȥ�) (2)").name = "Failure Photo(�Ȥ�-2)"
                                
                                Dim MystrT8(4) As String
                                MystrT8(0) = "Customer request"
                                MystrT8(1) = "1. The input impedance of phase mag board: 0.1 ohms"
                                MystrT8(2) = "2. Idle V/I = " & IdleV & "mV/" & IdleI & "mV"
                                MystrT8(3) = "3. Chuck On V/I = 2.45V/" & AfterSelect2 & "V "
                                MystrT8(4) = "4. Chuck On V/I(Max) = 2.45V/" & AfterSelect2 & "V "
                                
                                Worksheets("RMA").[E33] = Join(MystrT8, vbCrLf)
        
                                With Worksheets("RMA").[E33]
                                        .HorizontalAlignment = xlGeneral
                                        .VerticalAlignment = xlTop
                                End With
                                
                                Worksheets("Failure Photo(�Ȥ�)").Activate
                                [A17:E17] = ""
                                MsgBox "��ܵ��Ȥ�Ϥ�(�U�@�i�N�n)"
                                �K�W�l�a�Ӥ�
                                
                                
                                Worksheets("Failure Photo(�Ȥ�-2)").Activate
                                [A17:E17] = ""
                                With Range("A36:H36").Borders
                                        .LineStyle = xlContinuous
                                End With
                                
                                With Range("A58:D58").Borders
                                        .LineStyle = xlContinuous
                                End With
                                
                                With Range("A36:H36")
                                        .Merge
                                        .HorizontalAlignment = xlCenter
                                        .VerticalAlignment = xlCenter
                                End With
                                
                                With Range("A58:D58")
                                        .Merge
                                        .HorizontalAlignment = xlCenter
                                        .VerticalAlignment = xlCenter
                                End With
                                
                                [A36] = "Monitor ESC voltage out"
                                [A58] = "MN"
                                
                                With [A36].Font
                                        .name = "Tahoma"
                                        .Size = 12
                                End With
                                
                                With [A58].Font
                                        .name = "Tahoma"
                                        .Size = 12
                                End With
                        Else
                                AZX���.Show
                                MsgBox "���1�i�v�K����"
                                Call �i�ι�(36, 2)
                        End If
                        
                Case Is = "Test Table Tuner-020-023"
                
                        If cus = "741 �x�n��Ƕ�ϫn��_��1�� (6�t)" Then
                                AZX���.Show
                                MsgBox "T6���1�i�v�K����"
                                Call �i�ι�(41, 5)
                        End If
                
                Case Is = "Test Table Tuner (-037)", "Test Table Tuner (-043)", "Test Table Tuner", "Test Table Tuner (-039)"
                        AZX���.Show
                        MsgBox "���1�i�v�K����"
                        Call �i�ι�(36, 2)
                
                Case Is = "Failure Photo"
                        MsgBox "�п� " & ActiveSheet.name & " (�i�ƿ�)"
                        �K�W�l�a�Ӥ�
                        AZXError.Show
                        
                Case Is = "�i�X�t�Ӥ�"
                        MsgBox "�п� " & ActiveSheet.name & " (�i�ƿ�)"
                        �K�W�i�X�t�Ϥ�
         
                End Select
        Next sh
        
        Dim myMin%, mySec%
        myTime = Time - myTime
        myMin = Minute(myTime)
        mySec = Second(myTime)
        
        Worksheets("RMA").Select
        
        MsgBox "�B�z����" & Chr(10) & Chr(10) & "����ɶ�" & myMin & "��" & mySec & "��C", vbInformation
        
End Sub
