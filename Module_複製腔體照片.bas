Attribute VB_Name = "Module�ƻs����Ӥ�"
Sub �ƻs����Ӥ�()
    SN = InputBox("�п�JSN")
    RMA = InputBox("�п�JRMA")
    
    If SN = "" Then Exit Sub
    If RMA = "" Then Exit Sub
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If Not fs.FolderExists("P:\Service\Repair Picture\RF Product\��ˤH���ȩ�Source\" & RMA) Then
        MsgBox "RMA : " & RMA & " �S�������"
        Exit Sub
    End If
    
    Source = "P:\Service\Repair Picture\RF Product\��ˤH���ȩ�Source\" & RMA
    path = "D:\Users\Dlin\Desktop\�Ӥ�\" & SN
            
    fs.CopyFolder Source, path
    
    MsgBox "SN : " & SN & " �Ӥ��ƻs����"
    
End Sub
