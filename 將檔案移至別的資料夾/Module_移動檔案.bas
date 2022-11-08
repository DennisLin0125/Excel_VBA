Attribute VB_Name = "Module1移動檔案"
Sub 移動檔案()
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    SN = InputBox("請輸入要移動的檔案SN:")
    
    fd.Title = "請選擇要移動的檔案"
    
    fd.InitialFileName = ""
    
    If fd.Show = -1 Then
        
        Source = fd.SelectedItems(1)
    
        Set fd2 = Application.FileDialog(msoFileDialogFolderPicker)
        
        fd2.Title = "請選擇要移動到哪個資料夾"
        
        fd2.InitialFileName = ""
            
        If fd2.Show = -1 Then
        
            tpath = fd2.SelectedItems(1) & "\" & SN
                
            Set fs = CreateObject("scripting.FileSystemObject")
            
            fs.CopyFolder Source, tpath
            
            fs.DeleteFolder Source
                
            MsgBox "處理完成"
            
        End If
        
    End If
    
    Set fs = Nothing
    Set fd = Nothing
    Set fd2 = Nothing
    
End Sub

