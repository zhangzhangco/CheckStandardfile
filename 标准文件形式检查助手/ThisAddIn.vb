Imports System.Threading

Public Class ThisAddIn

    Public Shared UISynchronizationContext As SynchronizationContext

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        UISynchronizationContext = SynchronizationContext.Current
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
End Class
