Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Drawing.Imaging


''' <summary>
''' Provides fast access to a bitmap
''' </summary>
''' <remarks></remarks>
Public Class FastBitmapAccess


#Region "Support routines"


  ''' <summary>
  ''' Copys bitmap bytes to local BitmapBytesMatrix 
  ''' </summary>
  ''' <remarks></remarks>
  Public Shared Sub GetBytesFromBitmap(ByRef bm As Bitmap, ByRef byteArray As Byte(,,))

    Dim l_StartIndexForLine As Integer

    Dim l_StartIndexNextPixel As Integer

    Dim l_BytesVector() As Byte

    Dim l_BitmapData As BitmapData

    Dim l_BytesPerLine As Integer

    Dim l_BytesTotal As Integer

    Dim l_ByteDepth As Integer


    l_ByteDepth = Image.GetPixelFormatSize(bm.PixelFormat) \ 8

    If l_ByteDepth = 3 Then

      l_BitmapData = bm.LockBits(New Rectangle(0, 0, bm.Width, bm.Height), Imaging.ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb)

      l_BytesPerLine = l_BitmapData.Stride
      l_BytesTotal = l_BytesPerLine * l_BitmapData.Height

      ReDim l_BytesVector(l_BytesTotal - 1)
      ReDim byteArray(l_ByteDepth - 1, bm.Width - 1, bm.Height - 1)

      Marshal.Copy(l_BitmapData.Scan0, l_BytesVector, 0, l_BytesTotal)

      For y As Integer = 0 To bm.Height - 1

        l_StartIndexNextPixel = l_StartIndexForLine

        For x As Integer = 0 To bm.Width - 1
          byteArray(0, x, y) = l_BytesVector(l_StartIndexNextPixel)
          l_StartIndexNextPixel = l_StartIndexNextPixel + 1
          byteArray(1, x, y) = l_BytesVector(l_StartIndexNextPixel)
          l_StartIndexNextPixel = l_StartIndexNextPixel + 1
          byteArray(2, x, y) = l_BytesVector(l_StartIndexNextPixel)
          l_StartIndexNextPixel = l_StartIndexNextPixel + 1
        Next

        l_StartIndexForLine = l_StartIndexForLine + l_BytesPerLine

      Next

      bm.UnlockBits(l_BitmapData)

    Else
      Throw New Exception("ByteDepth not supported")
    End If

  End Sub



  ''' <summary>
  ''' Copies local BitmapBytesMatrix back to original bitmap 
  ''' </summary>
  ''' <remarks></remarks>
  Public Shared Sub PutBytesToBitmap(ByRef bm As Bitmap, ByRef byteArray As Byte(,,))

    Dim l_StartIndexForLine As Integer

    Dim l_StartIndexNextPixel As Integer

    Dim l_BytesVector() As Byte

    Dim l_BitmapData As BitmapData

    Dim l_BytesPerLine As Integer

    Dim l_BytesTotal As Integer

    Dim l_byteDepth As Integer


    l_byteDepth = Image.GetPixelFormatSize(bm.PixelFormat) \ 8

    If l_byteDepth = 3 Then

      l_BitmapData = bm.LockBits(New Rectangle(0, 0, bm.Width, bm.Height), Imaging.ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb)

      l_BytesPerLine = l_BitmapData.Stride
      l_BytesTotal = l_BytesPerLine * l_BitmapData.Height

      ReDim l_BytesVector(l_BytesTotal - 1)

      For y As Integer = 0 To bm.Height - 1

        l_StartIndexNextPixel = l_StartIndexForLine

        For x As Integer = 0 To bm.Width - 1

          l_BytesVector(l_StartIndexNextPixel) = byteArray(0, x, y)
          l_StartIndexNextPixel = l_StartIndexNextPixel + 1
          l_BytesVector(l_StartIndexNextPixel) = byteArray(1, x, y)
          l_StartIndexNextPixel = l_StartIndexNextPixel + 1
          l_BytesVector(l_StartIndexNextPixel) = byteArray(2, x, y)
          l_StartIndexNextPixel = l_StartIndexNextPixel + 1
        Next

        l_StartIndexForLine = l_StartIndexForLine + l_BytesPerLine

      Next

      Marshal.Copy(l_BytesVector, 0, l_BitmapData.Scan0, l_BytesTotal)

      bm.UnlockBits(l_BitmapData)

      ReDim byteArray(-1, -1, -1)

    Else
      Throw New Exception("ByteDepth not supported")
    End If

  End Sub

#End Region

End Class
