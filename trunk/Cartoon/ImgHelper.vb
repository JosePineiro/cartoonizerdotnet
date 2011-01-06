Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Drawing.Imaging

Public Class BitmapBytes
  ' Provide public access to the picture's byte data.
  Public ImageBytes() As Byte
  Public RowSizeBytes As Integer
  Private mPixelDataSize As Integer = 24

  ' Save a reference to the bitmap.
  Public Sub New(ByVal bm As Bitmap)
    m_Bitmap = bm
    mPixelDataSize = Image.GetPixelFormatSize(bm.PixelFormat)
  End Sub

  ' A reference to the Bitmap.
  Private m_Bitmap As Bitmap

  ' Bitmap data.
  Private m_BitmapData As BitmapData

  ' Lock the bitmap's data.
  Public Sub LockBitmap()
    ' Lock the bitmap data.
    Dim bounds As Rectangle = New Rectangle( _
        0, 0, m_Bitmap.Width, m_Bitmap.Height)
    m_BitmapData = m_Bitmap.LockBits(bounds, _
        Imaging.ImageLockMode.ReadWrite, _
        m_Bitmap.PixelFormat)
    RowSizeBytes = m_BitmapData.Stride

    ' Allocate room for the data.
    Dim total_size As Integer = m_BitmapData.Stride * _
        m_BitmapData.Height
    ReDim ImageBytes(total_size)

    ' Copy the data into the ImageBytes array.
    Marshal.Copy(m_BitmapData.Scan0, ImageBytes, _
        0, total_size)
  End Sub

  ' Copy the data back into the Bitmap
  ' and release resources.
  Public Sub UnlockBitmap()
    ' Copy the data back into the bitmap.
    Dim total_size As Integer = m_BitmapData.Stride * _
        m_BitmapData.Height
    Marshal.Copy(ImageBytes, 0, _
        m_BitmapData.Scan0, total_size)

    ' Unlock the bitmap.
    m_Bitmap.UnlockBits(m_BitmapData)

    ' Release resources.
    ImageBytes = Nothing
    m_BitmapData = Nothing
  End Sub
End Class

Public Class ImageHelper

  Public Shared Function GetBitmapBits(ByRef bm As Bitmap, ByRef array As Byte(,,)) As Integer

    Dim byteDepth As Integer = Image.GetPixelFormatSize(bm.PixelFormat) \ 8

    'Resize to hold image data
    ReDim array(byteDepth - 1, bm.Width - 1, bm.Height - 1)

    'Get the image data and store into Sbyte array)
    Dim bm_bytes As New BitmapBytes(bm)
    bm_bytes.LockBitmap()

    Dim i As Integer = 0
    For y As Integer = 0 To bm.Height - 1
      For x As Integer = 0 To bm.Width - 1
        For b As Integer = 0 To byteDepth - 1
          i += 1
          array(b, x, y) = bm_bytes.ImageBytes(i)
        Next
      Next
    Next

    bm_bytes.UnlockBitmap()
    Return byteDepth
  End Function

  Public Shared Function SetBitmapBits(ByRef bm As Bitmap, ByRef array As Byte(,,)) As Integer

    Dim byteDepth As Integer = Image.GetPixelFormatSize(bm.PixelFormat) \ 8

    Dim bm_bytes As New BitmapBytes(bm)

    bm_bytes.LockBitmap()

    Dim i As Integer = 0
    For y As Integer = 0 To bm.Height - 1
      For x As Integer = 0 To bm.Width - 1
        For b As Integer = 0 To byteDepth - 1
          i += 1
          bm_bytes.ImageBytes(i) = array(b, x, y)
        Next
      Next
    Next

    bm_bytes.UnlockBitmap()

    Return byteDepth
  End Function

End Class
