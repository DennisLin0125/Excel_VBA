Attribute VB_Name = "ModuleCTRL_Q"
Option Explicit
Sub �����}��CTRL_Q()
Attribute �����}��CTRL_Q.VB_ProcData.VB_Invoke_Func = "q\n14"
        Dim RMAnum As String
        Dim Year As String, fs As Object
        Dim WRAns, CompAns As Boolean
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        RMAnum = Trim(Range(ActiveCell.Address))
        
        If RMAnum = "" Then Exit Sub
        
        Dim wb As Workbook
        For Each wb In Workbooks
                If wb.name = RMAnum & ".xls" Then
                        MsgBox "�ɮ� " & RMAnum & " �w�}��", vbCritical
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
                                MsgBox "�Ъ`�N, �ثe����Ū"
                        End If
                        Application.DisplayAlerts = True
                        Exit For
                ElseIf CompAns = True Then
                        Application.DisplayAlerts = False
                        Workbooks.Open Filename:="P:\Service\RMA\Complete\" & i & "\" & RMAnum & ".xls"
                        If ActiveWorkbook.ReadOnly = True Then
                                MsgBox "�Ъ`�N, �ثe����Ū"
                        End If
                        Application.DisplayAlerts = True
                        Exit For
                        If i = 2006 Then
                                MsgBox "�ɮפ��s�b!"
                        End If
                End If
        Next i
End Sub

