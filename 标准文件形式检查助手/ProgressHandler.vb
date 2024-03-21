' ProgressHandler.vb
Imports System.Windows.Controls
Imports System.Windows.Forms
Imports Application = Microsoft.Office.Interop.Word.Application
Imports System.IO
Imports Microsoft.Office.Interop.Word

Public Class ProgressHandler
    Public Shared progressForm As ProgressForm = New ProgressForm()

    Public Sub ProgressStart()
        ' 初始化进度表单
        progressForm.Show()
        Globals.ThisAddIn.Application.ScreenUpdating = False
    End Sub
    Public Sub ProgressStartWaiting()
        ProgressStart()
        UpdateProgress(100, "处理中...")
    End Sub
    Public Sub UpdateProgress(currentProgress As Integer, message As String)
        progressForm.UpdateProgress(currentProgress)
        progressForm.UpdateMessage(message)
        If currentProgress Mod 10 = 0 Then
            System.Windows.Forms.Application.DoEvents()
        End If
    End Sub
    Public Sub ProgressEnd()
        Globals.ThisAddIn.Application.ScreenUpdating = True
        ' 如果有进度表单
        progressForm.Hide()
    End Sub
End Class
