Imports System.Runtime.InteropServices

Public Class NativeMethods
    <DllImport("user32.dll")>
    Public Shared Function GetDpiForWindow(hwnd As IntPtr) As Integer
    End Function
End Class