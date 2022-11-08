Attribute VB_Name = "Module選擇開檔案"
Option Explicit
Sub 選擇開檔()
Attribute 選擇開檔.VB_ProcData.VB_Invoke_Func = "r\n14"
        Dim RMAnum$
        Dim Year$, fs As Object
        Dim WRAns As Boolean, CompAns As Boolean
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        RMAnum = InputBox("選擇開啟RMA")
        
        If RMAnum = "" Then Exit Sub
        
        Dim wb As Workbook
        For Each wb In Workbooks
                If wb.name = RMAnum & ".xls" Then
                        MsgBox "檔案 " & RMAnum & " 已開啟", vbCritical
                        Exit Sub
                End If
        Next wb
        
        Year = "2022"
        Dim i%
        For i = Year To 2006 Step -1
            
                WRAns = fs.FileExists("P:\Service\RMA\WR\" & i & "\" & RMAnum & ".xls")
                CompAns = fs.FileExists("P:\Service\RMA\Complete\" & i & "\" & RMAnum & ".xls")
                
                If WRAns = True Then
                        Application.DisplayAlerts = False
                        Workbooks.Open Filename:="P:\Service\RMA\WR\" & i & "\" & RMAnum & ".xls"
                        If ActiveWorkbook.ReadOnly = True Then
                            MsgBox "請注意, 目前為唯讀"
                        End If
                        Application.DisplayAlerts = True
                        Exit For
                ElseIf CompAns = True Then
                        Application.DisplayAlerts = False
                        Workbooks.Open Filename:="P:\Service\RMA\Complete\" & i & "\" & RMAnum & ".xls"
                        If ActiveWorkbook.ReadOnly = True Then
                            MsgBox "請注意, 目前為唯讀"
                        End If
                        Application.DisplayAlerts = True
                        Exit For
                        If i = 2006 Then
                                MsgBox "檔案不存在!"
                        End If
                End If
        Next i
End Sub

