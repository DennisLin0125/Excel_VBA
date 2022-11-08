Attribute VB_Name = "Module複製腔體照片"
Sub 複製腔體照片()
    SN = InputBox("請輸入SN")
    RMA = InputBox("請輸入RMA")
    
    If SN = "" Then Exit Sub
    If RMA = "" Then Exit Sub
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If Not fs.FolderExists("") Then
        MsgBox "RMA : " & RMA & " 沒有拆腔體"
        Exit Sub
    End If
    
    Source = "" & RMA
    path = "D:\Users\Dlin\Desktop\照片\" & SN
            
    fs.CopyFolder Source, path
    
    MsgBox "SN : " & SN & " 照片複製完成"
    
End Sub
