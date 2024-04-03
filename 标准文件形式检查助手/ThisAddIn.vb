Imports System.Threading
Imports Microsoft.Office.Core
Public Class ThisAddIn

    Public Shared UISynchronizationContext As SynchronizationContext

    Public Sub ThisAddIn_Startup() Handles Me.Startup
        UISynchronizationContext = SynchronizationContext.Current
    End Sub

    Public Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub
    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon()
    End Function
End Class
