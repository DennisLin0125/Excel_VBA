Attribute VB_Name = "Module1�����ɮ�"
Sub �����ɮ�()
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    SN = InputBox("�п�J�n���ʪ��ɮ�SN:")
    
    fd.Title = "�п�ܭn���ʪ��ɮ�"
    
    fd.InitialFileName = "D:\Users\Dlin\Desktop\�Ӥ�"
    
    If fd.Show = -1 Then
        
        Source = fd.SelectedItems(1)
    
        Set fd2 = Application.FileDialog(msoFileDialogFolderPicker)
        
        fd2.Title = "�п�ܭn���ʨ���Ӹ�Ƨ�"
        
        fd2.InitialFileName = "P:\Service\Repair Picture"
            
        If fd2.Show = -1 Then
        
            tpath = fd2.SelectedItems(1) & "\" & SN
                
            Set fs = CreateObject("scripting.FileSystemObject")
            
            fs.CopyFolder Source, tpath
            
            fs.DeleteFolder Source
                
            MsgBox "�B�z����"
            
        End If
        
    End If
    
    Set fs = Nothing
    Set fd = Nothing
    Set fd2 = Nothing
    
End Sub

