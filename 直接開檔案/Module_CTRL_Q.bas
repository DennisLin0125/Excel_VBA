Attribute VB_Name = "ModuleCTRL_Q"
Option Explicit
Sub 直接開檔CTRL_Q()
Attribute 直接開檔CTRL_Q.VB_ProcData.VB_Invoke_Func = "q\n14"
        Dim RMAnum As String
        Dim Year As String, fs As Object
        Dim WRAns, CompAns As Boolean
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        RMAnum = Trim(Range(ActiveCell.Address))
        
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
        
                WRAns = fs.FileExists("")
                CompAns = fs.FileExists("")
                
                If WRAns = True Then
                        Application.DisplayAlerts = False
                        Workbooks.Open Filename:=""
                        If ActiveWorkbook.ReadOnly = True Then
                                MsgBox "請注意, 目前為唯讀"
                        End If
                        Application.DisplayAlerts = True
                        Exit For
                ElseIf CompAns = True Then
                        Application.DisplayAlerts = False
                        Workbooks.Open Filename:=""
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

