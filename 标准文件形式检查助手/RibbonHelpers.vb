Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class RibbonHelpers
    <DllImport("oleaut32.dll", ExactSpelling:=True, PreserveSig:=False)>
    Private Shared Function OleCreatePictureIndirect(
        <MarshalAs(UnmanagedType.AsAny)> ByVal picdesc As Object,
        ByRef refiid As Guid,
        <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
    End Function

    Public Shared Function ImageToPictureDisp(image As Image) As stdole.IPictureDisp
        Dim bm As New Bitmap(image)
        Dim g As Graphics = Graphics.FromImage(bm)
        Dim hdc As IntPtr = g.GetHdc()
        g.ReleaseHdc(hdc)
        g.Dispose()
        Dim pictDesc As New PICTDESCbitmap(bm)
        Return OleCreatePictureIndirect(pictDesc, GetType(stdole.IPictureDisp).GUID, True)
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Private Structure PICTDESCbitmap
        Private cbSizeOfStruct As Integer
        Public picType As Integer
        Public hbitmap As IntPtr
        Public hpal As IntPtr
        Public unused As Integer

        Public Sub New(bitmap As Bitmap)
            Me.cbSizeOfStruct = Marshal.SizeOf(GetType(PICTDESCbitmap))
            Me.picType = 1 ' PICTYPE_BITMAP
            Me.hbitmap = bitmap.GetHbitmap()
            Me.hpal = IntPtr.Zero
            Me.unused = 0
        End Sub
    End Structure
End Class
