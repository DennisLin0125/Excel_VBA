Attribute VB_Name = "Module1�ƻs�Ӥ���P��"
Sub �إ߷Ӥ���P��()

        SN = InputBox("�п�JSN")
        If SN = "" Then Exit Sub
        
        Dim myPath$
        myPath = "D:\Users\Dlin\Desktop\�Ӥ�\" & SN
        If Dir(myPath & "\") = "" Then
                MsgBox "�Ӥ���Ƨ��̡A�S��SN : " & SN & " ���Ӥ��I", vbCritical
                Exit Sub
        End If
        
        RMA = InputBox("�п�JRMA")
        If RMA = "" Then Exit Sub
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        If fs.FolderExists("P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN) Then
                fs.DeleteFolder "P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN
        End If
        
        With fs
                .CreateFolder ("P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN)
                .CreateFolder ("P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA)
                .CreateFolder ("P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\LOG\")
                .CreateFolder ("P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\�i�X�t�Ӥ�")
                .CreateFolder ("P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\�G�ٹϤ�")
                .CreateFolder ("P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\�G�ٹϤ�\���׫e")
                .CreateFolder ("P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\�G�ٹϤ�\���׫�")
                .CreateFolder ("P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\�i�X�t�Ӥ�\���׫e")
                .CreateFolder ("P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\�i�X�t�Ӥ�\���׫�")
        End With
        
        Dim RpBeforPath$, RpAfterPath$, MachAfterPath$, MachBeforPath$, SourcePath$, LogPath$
        
        RpBeforPath = "P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\�G�ٹϤ�\���׫e\"
        RpAfterPath = "P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\�G�ٹϤ�\���׫�\"
        
        MachAfterPath = "P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\�i�X�t�Ӥ�\���׫�\"
        MachBeforPath = "P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\�i�X�t�Ӥ�\���׫e\"
        
        SourcePath = "P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA
        LogPath = "P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN & "\" & RMA & "\LOG\"
        
        '���׫e�~�[�Ӥ�
        Dim fdMachBefor As FileDialog
        Set fdMachBefor = Application.FileDialog(msoFileDialogFilePicker)
        
        With fdMachBefor
                .AllowMultiSelect = True
                .InitialFileName = myPath
                .ButtonName = "�N�O�A�F!!"
                .Title = "�п�� �i�X�t�Ӥ���(���׫e)�Ӥ�"
                MsgBox "���U�ӡA�N��Ӥ��M��ư�����" & Chr(10) & Chr(10) & "�п�� �i�X�t�Ӥ��� (���׫e) �Ӥ�"
        
                If .Show Then
                        For Each sPath In .SelectedItems
                                fs.CopyFile sPath, MachBeforPath
                        Next
                End If
        End With
        
        '���׫�~�[�Ӥ�
        Dim fdMachAfter As FileDialog
        Set fdMachAfter = Application.FileDialog(msoFileDialogFilePicker)
        
        With fdMachAfter
                .AllowMultiSelect = True
                .Title = "�п�� �i�X�t�Ӥ��� (���׫�) �Ӥ�"
                .ButtonName = "�N�O�A�F!!"
                MsgBox "�п�� �i�X�t�Ӥ���(���׫�)�Ӥ�"
               
                If .Show Then
                        For Each sPath In .SelectedItems
                                fs.CopyFile sPath, MachAfterPath
                        Next
                End If
        End With
        
        '���׫e�G�ٷӤ�
        Dim fdRpBefor As FileDialog
        Set fdRpBefor = Application.FileDialog(msoFileDialogFilePicker)
        
        With fdRpBefor
                .AllowMultiSelect = True
                .Title = "�п�� �G�ٷӤ��� (���׫e) �Ӥ�"
                .ButtonName = "�N�O�A�F!!"
                MsgBox "�п�� �G�ٷӤ���(���׫e)�Ӥ�"
               
                If .Show Then
                        For Each sPath In .SelectedItems
                                fs.CopyFile sPath, RpBeforPath
                        Next
                End If
        End With
        
        '���׫e�G�ٷӤ�
        Dim fdRpAfter As FileDialog
        Set fdRpAfter = Application.FileDialog(msoFileDialogFilePicker)
        
        With fdRpAfter
                .AllowMultiSelect = True
                .Title = "�п�� �G�ٷӤ��� (���׫�) �Ӥ�"
                .ButtonName = "�N�O�A�F!!"
                MsgBox "�п�� �G�ٷӤ���(���׫�)�Ӥ�"
                
                If .Show Then
                        For Each sPath In .SelectedItems
                                fs.CopyFile sPath, RpAfterPath
                        Next
                End If
        End With
        
        'LOG �I���q��
        Dim fdLog As FileDialog
        Set fdLog = Application.FileDialog(msoFileDialogFilePicker)
        
        With fdLog
                .AllowMultiSelect = True
                .Title = "�п��Log���"
                .ButtonName = "�N�O�A�F!!"
                MsgBox "�п�� LOG ���"
                
                If .Show Then
                        For Each sPath In .SelectedItems
                                fs.CopyFile sPath, LogPath
                        Next
                End If
        End With
        
        '����ո˷Ӥ�
'        Dim fdSource As FileDialog
'        Set fdSource = Application.FileDialog(msoFileDialogFolderPicker)
'
'        With fdSource
'                .InitialFileName = "P:\Service\Repair Picture\RF Product\��ˤH���ȩ�Source"
'                .Title = "�п�ܲոˤH���Ӥ��HRMA�R�W����Ƨ�"
'                .ButtonName = "�N�O�A�F!!"
'                MsgBox "�п�� �ոˤH���Ӥ��H RMA �R�W����Ƨ�" & Chr(10) & Chr(10) & "�p�G���饼��Ы� ����"
'
'                If .Show Then
'                        fs.CopyFolder .SelectedItems(1), SourcePath
'                End If
'        End With
        
        Set fs = Nothing
        Set fdLog = Nothing
        Set fdSource = Nothing
        Set fdRpAfter = Nothing
        Set fdRpBefor = Nothing
        Set fdMachAfter = Nothing
        Set fdMachBefor = Nothing
        
        
        '����
        Source = "P:\Service\�޳N�Q�ױM��\Engineer\Dennis\MKS\" & SN
        
        Dim fd2 As FileDialog, a%
        Set fd2 = Application.FileDialog(msoFileDialogFolderPicker)
        
        With fd2
                .Title = "�п�ܭn���ʨ���Ӹ�Ƨ�"
                .InitialFileName = "P:\Service\Product\Repair Picture"
                .ButtonName = "�N�O�A�F!!"
                MsgBox "�п�ܭn���ʨ���Ӹ�Ƨ�"
                    
                If .Show Then
                        tpath = .SelectedItems(1) & "\" & SN
                        Set fs = CreateObject("scripting.FileSystemObject")
                        fs.CopyFolder Source, tpath
                        fs.DeleteFolder Source
                        a = MsgBox("�O�_�n�R�� " & SN & " ��Ƨ� ?", vbYesNo)
                        If a = vbYes Then
                                fs.DeleteFolder myPath
                        End If
                        MsgBox "�B�z����"
                End If
        End With
        
        Set fd2 = Nothing
        Set fs = Nothing
End Sub

