Imports System.Drawing
Imports System.Math

Public Class Filter

  ''' <summary>
  ''' Bilateral Image filter
  ''' </summary>
  ''' <param name="inImg">input image </param>
  ''' <param name="sigD">spatial sigma</param>
  ''' <param name="sigR">edge sigma</param>
  ''' <param name="hH">half-height of window to use</param>
  ''' <param name="hW">half-width of window to use</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function Bilateral(ByRef inImg As Bitmap, _
                                   ByVal sigD As Double, _
                                   ByVal sigR As Double, _
                                   ByVal hH As Integer, _
                                   ByVal hW As Integer) As Bitmap

    Dim bm_bytes As New BitmapBytes(inImg)
    bm_bytes.LockBitmap()
    Dim inImgBytes As Byte() = bm_bytes.ImageBytes
    bm_bytes.UnlockBitmap()

    Dim nCols As Integer = inImg.Width
    Dim mRows As Integer = inImg.Height

    Dim tmpImgBytes(inImgBytes.Length - 1) As Byte
    Dim outImgBytes(inImgBytes.Length - 1) As Byte
    Dim row, col, erow, ecol As Integer
    Dim num, den, c, s As Double

    'ignore pixels too close to the edge, since we can't fit a whole window
    'around them. (could be smarter and do symmetry or something)
    For row = hH To mRows - hH - 1
      For col = hW To nCols - hW - 1
        'integrate over the window *
        num = 0
        den = 0
        erow = row
        For ecol = col - hW To col + (hW - 1)
          'compute the spatial weight for this pixel 
          s = Math.Exp(-0.5 * Math.Pow(Math.Abs(ecol - col) / sigD, 2))
          'compute the edge weight for this pixel 
          c = Math.Exp(-0.5 * Math.Pow(Math.Abs(CInt(inImgBytes(erow * nCols + ecol)) - CInt(inImgBytes(row * nCols + col))) / sigR, 2))
          'update the numerator and denominator
          num += CInt(inImgBytes(erow * nCols + ecol)) * c * s
          den += c * s
        Next
        tmpImgBytes(row * nCols + col) = CByte(If(den = 0, 0, num / den))
      Next
    Next

    'now do our second pass, this time for vertical filtering
    For row = hH To mRows - hH - 1
      For col = hW To nCols - hW - 1
        'integrate over the window 
        num = 0
        den = 0
        ecol = col
        For ecol = col - hW To col + (hW - 1)
          'compute the spatial weight for this pixel 
          s = Math.Exp(-0.5 * Math.Pow(Math.Abs(ecol - col) / sigD, 2))
          'compute the edge weight for this pixel 
          c = Math.Exp(-0.5 * Math.Pow(Math.Abs(CInt(inImgBytes(erow * nCols + ecol)) - CInt(inImgBytes(row * nCols + col))) / sigR, 2))
          'update the numerator and denominator 
          num += CInt(inImgBytes(erow * nCols + ecol)) * c * s
          den += c * s
        Next
        'fill in the filtered value for this pixel 
        outImgBytes(row * nCols + col) = CByte(If(den = 0, 0, num / den))
      Next
    Next

    Dim outBM As New Bitmap(inImg.Width, inImg.Height, inImg.PixelFormat)
    Dim outBm_bytes As New BitmapBytes(outBM)
    outBm_bytes.LockBitmap()
    outBm_bytes.ImageBytes = outImgBytes
    outBm_bytes.UnlockBitmap()

    tmpImgBytes = Nothing
    outImgBytes = Nothing

    Return outBM

  End Function

  Public Enum FilterType
    Gaussian
    Average
    Bilateral
  End Enum

  Public Shared Function Bilateral2(ByRef inputImage As Bitmap) As Bitmap

    Dim bm_bytes As New BitmapBytes(inputImage)
    bm_bytes.LockBitmap()
    Dim inputImageDouble(bm_bytes.ImageBytes.Length - 1) As Double
    For i As Integer = 0 To bm_bytes.ImageBytes.Length - 1
      inputImageDouble(i) = CDbl(bm_bytes.ImageBytes(i))
    Next
    bm_bytes.UnlockBitmap()

    Dim outputImageBytes(inputImageDouble.Length - 1) As Byte

    Dim windowSize As Integer = 4
    Dim center As Integer = windowSize / 2
    Dim filter As Integer = FilterType.Bilateral

    Dim sigma1 As Integer = 1
    Dim sigma2 As Integer = 50

    Dim WIDTH As Integer = inputImage.Width
    Dim HEIGHT As Integer = inputImage.Height

    Dim div, divR, divG, divB As Double

    Dim kernel(windowSize, windowSize) As Double
    Dim kernelR(windowSize, windowSize) As Double
    Dim kernelG(windowSize, windowSize) As Double
    Dim kernelB(windowSize, windowSize) As Double

    'init Kernel   
    For row As Integer = -center To center
      For col As Integer = -center To center
        kernel(row + center, col + center) = 0.0
        If filter = FilterType.Average Then
          kernel(row + center, col + center) = 1.0
        ElseIf filter = FilterType.Gaussian Then
          kernel(row + center, col + center) = 1.0 / (Pow(sigma1, 2.0) * 2 * Pi) * Exp(-0.5 * Pow(Sqrt(Pow(col, 2.0) + Pow(row, 2.0)) / sigma1, 2))
        End If
        div += kernel(row + center, col + center)
      Next
    Next

    'convolution   
    For r As Integer = 0 To HEIGHT - 1
      For c As Integer = 0 To WIDTH - 1

        If filter = FilterType.Bilateral Then
          'init the kernel for current position   
          divR = 0
          divG = 0
          divB = 0
          For row As Integer = -center To center
            For col As Integer = -center To center
              Dim tempValueG As Double = 0
              Dim tempValueR As Double = 0
              Dim tempValueB As Double = 0

              If (r + row >= 0 AndAlso r + row < HEIGHT) AndAlso (c + col >= 0 AndAlso c + col < WIDTH) Then
                tempValueR = inputImageDouble((((r + row) * (c + col)) * 3) + 2)
                tempValueG = inputImageDouble((((r + row) * (c + col)) * 3) + 1)
                tempValueB = inputImageDouble((((r + row) * (c + col)) * 3) + 0)
              End If

              Dim domain As Double = 1.0 / (Pow(sigma1, 2.0) * 2 * Pi) * Exp(-0.5 * Pow(Sqrt(Pow(col, 2.0) + Pow(row, 2.0)) / sigma1, 2))
              Dim rangeR As Double = 1.0 / (sigma2 * Sqrt(2 * Pi)) * Exp(-0.5 * Pow((tempValueR - inputImageDouble(((r * c) * 3) + 2)) / sigma2, 2))
              Dim rangeG As Double = 1.0 / (sigma2 * Sqrt(2 * Pi)) * Exp(-0.5 * Pow((tempValueG - inputImageDouble(((r * c) * 3) + 1)) / sigma2, 2))
              Dim rangeB As Double = 1.0 / (sigma2 * Sqrt(2 * Pi)) * Exp(-0.5 * Pow((tempValueB - inputImageDouble(((r * c) * 3) + 0)) / sigma2, 2))

              kernelR(row + center, col + center) = domain * rangeR
              kernelG(row + center, col + center) = domain * rangeG
              kernelB(row + center, col + center) = domain * rangeB
              divR += kernelR(row + center, col + center)
              divG += kernelG(row + center, col + center)
              divB += kernelB(row + center, col + center)
            Next
          Next
        End If

        Dim tempG As Double = 0
        Dim tempR As Double = 0
        Dim tempB As Double = 0
        If filter = FilterType.Bilateral Then
          For row As Integer = -center To center
            For col As Integer = -center To center
              Dim tempValueR As Double = 0
              Dim tempValueG As Double = 0
              Dim tempValueB As Double = 0

              If (r + row >= 0 AndAlso r + row < HEIGHT) AndAlso (c + col >= 0 AndAlso c + col < WIDTH) Then
                tempValueR = inputImageDouble((((r + row) * (c + col)) * 3) + 2)
                tempValueG = inputImageDouble((((r + row) * (c + col)) * 3) + 1)
                tempValueB = inputImageDouble((((r + row) * (c + col)) * 3) + 0)
              End If

              tempR += tempValueR * kernelR(row + center, col + center) * 1.0
              tempG += tempValueG * kernelG(row + center, col + center) * 1.0
              tempB += tempValueB * kernelB(row + center, col + center) * 1.0
            Next
          Next
          tempR = tempR / divR
          tempG = tempG / divG
          tempB = tempB / divB
        Else
          For row As Integer = -center To center
            For col As Integer = -center To center
              Dim tempValueR As Double = 0
              Dim tempValueG As Double = 0
              Dim tempValueB As Double = 0

              If (r + row >= 0 AndAlso r + row < HEIGHT) AndAlso (c + col >= 0 AndAlso c + col < WIDTH) Then
                tempValueR = inputImageDouble((((r + row) * (c + col)) * 3) + 2)
                tempValueG = inputImageDouble((((r + row) * (c + col)) * 3) + 1)
                tempValueB = inputImageDouble((((r + row) * (c + col)) * 3) + 0)
              End If

              tempR += tempValueR * kernel(row + center, col + center) * 1.0
              tempG += tempValueG * kernel(row + center, col + center) * 1.0
              tempB += tempValueB * kernel(row + center, col + center) * 1.0
            Next
          Next

          tempR = tempR / div
          tempG = tempG / div
          tempB = tempB / div
        End If

        outputImageBytes(((r * c) * 3) + 2) = CByte(tempR)
        outputImageBytes(((r * c) * 3) + 1) = CByte(tempG)
        outputImageBytes(((r * c) * 3) + 0) = CByte(tempB)
      Next
    Next

    Dim outImg As New Bitmap(inputImage.Width, inputImage.Height, inputImage.PixelFormat)
    Dim outBMBytes As New BitmapBytes(outImg)
    bm_bytes.LockBitmap()
    outBMBytes.ImageBytes = outputImageBytes
    bm_bytes.UnlockBitmap()
    Return outImg

  End Function


End Class
