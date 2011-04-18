Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Math

Public Class Effects

  'Author :Roberto Mior
  '     reexre@gmail.com
  '
  'If you use source code or part of it please cite the author
  'You can use this code however you like providing the above credits remain intact
  '--------------------------------------------------------------------------------

#Region "Declarations"

  Enum eSource
    source
  End Enum

  Enum IntensityMode
    Gaussian
    NotGaussian
  End Enum

  Private Structure tHSP
    Public H As Single
    Public S As Single
    Public P As Single
  End Structure

  Private Structure tVector
    Public X As Single
    Public Y As Single
    Public L As Single
    Public A As Single
    Public TA As Single
    Public CosA As Single
    Public SinA As Single
  End Structure

  Private Structure tLAB
    Public L As Single
    Public A As Single
    Public B As Single
  End Structure

  Private Structure Bitmap
    Public bmType As Long
    Public bmWidth As Long
    Public bmHeight As Long
    Public bmWidthBytes As Long
    Public bmPlanes As Integer
    Public bmBitsPixel As Integer
    Public bmBits As Long
  End Structure

  Private Structure tREGION
    Public cx As Long
    Public cy As Long
    Public NP As Long
    Public X() As Long
    Public Y() As Long
    Public mR As Single
    Public mG As Single
    Public mB As Single
  End Structure

  Private Sbyte1(0, 0, 0) As Byte
  Private Sbyte2(0, 0, 0) As Byte

  Private BlurByte() As Byte

  Private SepaByte() As Byte

  Private BILAByte(0, 0, 0) As Byte
  Private ContByte(0, 0, 0) As Byte
  Private ContByte2(0, 0, 0) As Byte
  Private COMPUTEsingle(0, 0, 0) As Single
  Private COMPUTEsingle2(0, 0, 0) As Single
  Private BILASingle(0, 0, 0) As Single

  Private HSP(0, 0) As tHSP
  Private Vec(0, 0) As tVector
  Private LAB(0, 0) As tLAB

  Private hBmp As Bitmap

  Private pW As Long
  Private pH As Long
  Private pB As Long

  Private Fast_ExpIntensity(51000) As Single 'Fast_ExpIntensity(x+25500)

  Private Fast_ExpSpatial() As Single
  Private Fast_SpatialDomain(0, 0) As Single 'Fast_SpatialDomain(X + 30, Y + 30)
  Private Fast_SpatialDomain2(0, 0) As Single 'Fast_SpatialDomain2(X + 30, Y + 30)

  Public Event PercDONE(ByVal Filter As String, ByVal PercValue As Single, ByVal CurrIteration As Long)

  'BrightNess Contrast Saturation
  Private Class THistSingle
    Public A(0 To 255) As Single
  End Class

  Private histR(0 To 255) As Long
  Private histG(0 To 255) As Long
  Private histB(0 To 255) As Long

  'Luminance segmentation video frames NEW MODE
  Private Structure tSEG
    Public Start As Single
    Public [End] As Single
    Public Value As Single
  End Structure

  Private Structure tHistoCache
    Public Seg() As tSEG
  End Structure

  Dim HistoCache(0 To 4) As tHistoCache
  Private hcIDX As Long
  Private Const hcSIZE As Long = 5

  Private Radius As Long

#End Region

#Region "Initializations"

  Public Sub SetUpHistoChache(ByVal Nsegm)
    Dim I As Long
    For I = 0 To hcSIZE - 1
      ReDim HistoCache(I).Seg(0 To Nsegm - 1) 'Seg(x-1)
    Next
  End Sub

  Private Function Fast_IntensityDomain(ByVal AbsDeltaIntensity As Single) As Single    '
    Fast_IntensityDomain = Fast_ExpIntensity(AbsDeltaIntensity + 25500)
  End Function

  Public Sub zInit_IntensityDomain(ByVal SigmaI As Single, ByVal Mode As Integer)

    Dim V As Single
    Dim V2 As Single
    Dim Cos0 As Boolean

    If SigmaI = 0 Then SigmaI = 0.00001

    Select Case Mode
      Case 0
        SigmaI = 2 * SigmaI * SigmaI
      Case 1
        SigmaI = 2 * SigmaI
      Case 2
        SigmaI = SigmaI
      Case 3
        SigmaI = 3 * SigmaI
      Case 4
        SigmaI = Atan(1) * 2 * SigmaI
    End Select

    ReDim Fast_ExpIntensity(0 To 51000)

    Cos0 = False
    For V = -25500 To 25500

      V2 = Abs(V / 25500)
      Select Case Mode
        Case 0                'Gaussian
          V2 = V2 * V2
          Fast_ExpIntensity(V + 25500) = Exp(-(V2 / SigmaI))
        Case 1                'Gaussian2
          Fast_ExpIntensity(V + 25500) = (Exp(-(V2 ^ 3 / SigmaI ^ 3)))
        Case 2                'NotGaussian
          Fast_ExpIntensity(V + 25500) = Exp(-(V2 / SigmaI))
        Case 3
          Fast_ExpIntensity(V + 25500) = 1 - V2 / SigmaI
          If Fast_ExpIntensity(V + 25500) < 0 Then Fast_ExpIntensity(V + 25500) = 0
        Case 4
          If Not (Cos0) Then
            Fast_ExpIntensity(V + 25500) = Cos(Atan(1) * V2 / SigmaI)
            If Fast_ExpIntensity(V + 25500) < 0 Then
              Fast_ExpIntensity(V + 25500) = 0
              Cos0 = True
            End If
          Else
            Fast_ExpIntensity(V + 25500) = 0
          End If
      End Select

    Next
  End Sub

  Public Sub zInit_SpatialDomain(ByVal SigmaS As Single)
    zInit_SpatialDomain(SigmaS, Fast_SpatialDomain)
  End Sub

  Public Sub zInit_SpatialDomain(ByVal SigmaS As Single, ByRef arr As Single(,))
    Dim V As Single
    Dim X As Long
    Dim Y As Long
    Dim D As Long

    If SigmaS = 0 Then Exit Sub
    SigmaS = SigmaS * 2

    ReDim Fast_ExpSpatial(1800)
    For V = 0 To 1800
      Fast_ExpSpatial(V) = Exp(-(V / (SigmaS)))
    Next

    ReDim Fast_SpatialDomain(60, 60)
    For X = 0 To 60
      For Y = 0 To 60
        D = (X - 30) * (X - 30) + (Y - 30) * (Y - 30)
        arr(X, Y) = Fast_ExpSpatial(D)
      Next
    Next

  End Sub

  Public Sub zInit_SpatialDomain2LoG(ByVal SigmaS)
    Dim V As Single
    Dim V2 As Single
    Dim X As Long
    Dim Y As Long
    Dim D As Long

    If SigmaS = 0 Then Exit Sub
    SigmaS = SigmaS * SigmaS


    ReDim Fast_SpatialDomain2(0 To 60, 0 To 60)
    For X = -30 To 30
      For Y = -30 To 30
        D = X * X + Y * Y
        Fast_SpatialDomain2(X + 30, Y + 30) = (D - 2 * SigmaS) / (SigmaS * SigmaS) * Exp(-(D / (2 * SigmaS)))
      Next
    Next

  End Sub

#End Region

#Region "Limiters"

  Public Function zLimitMin0(ByVal V) As Byte
    If V < 0 Then
      Return CByte(0)
    ElseIf V > 255 Then
      Return CByte(255)
    Else
      Return CByte(V)
    End If
  End Function

  Public Function zLimitMax255(ByVal V As Single) As Byte
    If V > 255 Then
      Return CByte(255)
    Else
      Return CByte(V)
    End If
  End Function

#End Region

#Region "Spatial Preview"

  Sub zPreview_Intensity(ByRef bm As Drawing.Bitmap, ByVal cSigma As Single, ByVal Mode As Integer)

    Dim X As Single

    Dim V As Single
    Dim ky As Single

    Dim x1 As Single
    Dim y1 As Single
    Dim X2 As Single
    Dim Y2 As Single
    Dim KX As Single

    zInit_IntensityDomain(cSigma, Mode)
    ky = 255 / bm.Height
    KX = 32

    Dim g As Graphics = Graphics.FromImage(bm)
    For X = 0 To KX
      V = Fast_IntensityDomain(Abs(X * 100)) * 255

      x1 = ((bm.Width / KX) * X)
      y1 = bm.Height - V / ky
      X2 = ((bm.Width / KX) * (X + 1))
      Y2 = bm.Height

      Dim myBrush As New Drawing.SolidBrush(Color.FromArgb(V, 0, 0))
      Dim myRectF As New RectangleF(x1, y1, X2, Y2)
      g.FillRectangle(myBrush, myRectF)
    Next
    g.Dispose()

  End Sub

  Sub zPreview_Spatial(ByRef bm As Drawing.Bitmap, ByVal nn As Integer, ByVal cSigma As Single)
    Dim X As Long
    Dim Y As Long
    Dim C As Integer
    Dim X2
    Dim Y2
    Dim K As Single

    K = bm.Width / ((nn * 2) + 1)

    zInit_SpatialDomain(cSigma, Fast_SpatialDomain)
    Dim g As Graphics = Graphics.FromImage(bm)
    For X = -nn To nn
      For Y = -nn To nn
        C = Fast_SpatialDomain(X + 30, Y + 30) * 255
        X2 = (nn + X) * K
        Y2 = (nn + Y) * K
        Dim myBrush As New Drawing.SolidBrush(Color.FromArgb(C, 0, 0))
        Dim myRectF As New RectangleF(X2, Y2, X2 + K, Y2 + K)
        g.FillRectangle(myBrush, myRectF)
      Next
    Next

    g.Dispose()

  End Sub

#End Region

#Region "Effect Source Input/Output"

  Public Sub zSet_Source(ByVal bm As Drawing.Bitmap)

    hBmp.bmBitsPixel = Image.GetPixelFormatSize(bm.PixelFormat)
    hBmp.bmWidth = bm.Width
    hBmp.bmHeight = bm.Height
    hBmp.bmWidthBytes = hBmp.bmWidth * hBmp.bmBitsPixel

    pW = hBmp.bmWidth - 1
    pH = hBmp.bmHeight - 1
    pB = (hBmp.bmBitsPixel \ 8) - 1

    If pB = 3 Then pB = 2

    'Resize to hold image data
    ReDim Sbyte1(pB, pW, pH)

    'Get the image data and store into Sbyte array)
    'ImageHelper.GetBitmapBits(bm, Sbyte1)
    FastBitmapAccess.GetBytesFromBitmap(bm, Sbyte1)

  End Sub

  Public Sub zGet_Effect(ByRef bm As Drawing.Bitmap)

    'ImageHelper.SetBitmapBits(bm, BILAByte)
    FastBitmapAccess.PutBytesToBitmap(bm, BILAByte)

    Erase BILAByte
    Erase BILASingle
    Erase COMPUTEsingle

  End Sub

#End Region

#Region "Effect - Contour"

  Public Sub zEFF_ContourBySobel(ByVal Contour_0_100 As Single, ByVal LumHue01 As Single)
    Dim X As Long
    Dim Y As Long

    Dim ContAmount As Single
    Dim PercLUM As Single
    Dim PercHUE As Single

    Dim Lx As Single
    Dim Ly As Single
    Dim Ax As Single
    Dim Ay As Single
    Dim Bx As Single
    Dim By As Single
    Dim II As Single

    Dim dL As Single
    Dim dA As Single
    Dim dB As Single

    Dim ProgXStep As Long
    Dim ProgX As Long

    PercHUE = LumHue01
    PercLUM = 1 - PercHUE

    PercHUE = PercHUE * 2

    ContAmount = 0.000075 * Contour_0_100

    ReDim ContByte(0 To pB, 0 To pW, 0 To pH)
    ReDim ContByte2(0 To pB, 0 To pW, 0 To pH)
    ReDim Vec(0 To pW, 0 To pH)

    ProgXStep = Round(2 * pW / 100)
    ProgX = 0

    'sobel
    For X = 1 To pW - 1
      For Y = 1 To pH - 1

        Ly = (PercLUM * (-(-LAB(X - 1, Y - 1).L - 2 * LAB(X - 1, Y).L - LAB(X - 1, Y + 1).L + LAB(X + 1, Y - 1).L + 2 * LAB(X + 1, Y).L + LAB(X + 1, Y + 1).L)))
        Lx = (PercLUM * ((-LAB(X - 1, Y - 1).L - 2 * LAB(X, Y - 1).L - LAB(X + 1, Y - 1).L + LAB(X - 1, Y + 1).L + 2 * LAB(X, Y + 1).L + LAB(X + 1, Y + 1).L)))

        If PercHUE > 0 Then
          Ay = (PercHUE * (-(-LAB(X - 1, Y - 1).A - 2 * LAB(X - 1, Y).A - LAB(X - 1, Y + 1).A + LAB(X + 1, Y - 1).A + 2 * LAB(X + 1, Y).A + LAB(X + 1, Y + 1).A)))
          Ax = (PercHUE * ((-LAB(X - 1, Y - 1).A - 2 * LAB(X, Y - 1).A - LAB(X + 1, Y - 1).A + LAB(X - 1, Y + 1).A + 2 * LAB(X, Y + 1).A + LAB(X + 1, Y + 1).A)))

          By = (PercHUE * (-(-LAB(X - 1, Y - 1).B - 2 * LAB(X - 1, Y).B - LAB(X - 1, Y + 1).B + LAB(X + 1, Y - 1).B + 2 * LAB(X + 1, Y).B + LAB(X + 1, Y + 1).B)))
          Bx = (PercHUE * ((-LAB(X - 1, Y - 1).B - 2 * LAB(X, Y - 1).B - LAB(X + 1, Y - 1).B + LAB(X - 1, Y + 1).B + 2 * LAB(X, Y + 1).B + LAB(X + 1, Y + 1).B)))
        End If

        dL = Sqrt(Lx * Lx + Ly * Ly)
        dA = Sqrt(Ax * Ax + Ay * Ay)
        dB = Sqrt(Bx * Bx + By * By)

        II = (dL * dL + dA * dA + dB * dB) ^ 0.95
        II = II * ContAmount ^ (1 / 0.95)

        ContByte(0, X, Y) = zLimitMax255(II)

      Next

      If X > ProgX Then
        RaiseEvent PercDONE("Contour", X / pW, 0)
        ProgX = ProgX + ProgXStep
      End If

    Next

    RaiseEvent PercDONE("Contour", 1, 0)

  End Sub

  Public Sub zEFF_ContourBySobel2(ByVal Contour_0_100 As Single, ByVal LumHue01 As Single)
    Dim X As Long
    Dim Y As Long

    Dim ContAmount As Single
    Dim PercLUM As Single
    Dim PercHUE As Single

    Dim Lx As Single
    Dim Ly As Single
    Dim Ax As Single
    Dim Ay As Single
    Dim Bx As Single
    Dim By As Single
    Dim II As Single

    Dim dL As Single
    Dim dA As Single
    Dim dB As Single

    Dim ProgXStep As Long
    Dim ProgX As Long

    PercHUE = LumHue01
    PercLUM = 1 - PercHUE

    PercHUE = PercHUE * 2

    ContAmount = 0.000075 * Contour_0_100
    ContAmount = ContAmount * 4 / 12

    ReDim ContByte(0 To pB, 0 To pW, 0 To pH)
    ReDim ContByte2(0 To pB, 0 To pW, 0 To pH)
    ReDim Vec(0 To pW, 0 To pH)

    ProgXStep = Round(2 * pW / 100)
    ProgX = 0

    'sobel
    For X = 2 To pW - 2
      For Y = 2 To pH - 2

        Ly = (PercLUM * (-(-2 * LAB(X - 1, Y - 1).L - 4 * LAB(X - 1, Y).L - 2 * LAB(X - 1, Y + 1).L + 2 * LAB(X + 1, Y - 1).L + 4 * LAB(X + 1, Y).L + 2 * LAB(X + 1, Y + 1).L)))
        Ly = Ly + (PercLUM * (-(-LAB(X - 2, Y - 1).L - 2 * LAB(X - 2, Y).L - LAB(X - 2, Y + 1).L + LAB(X + 2, Y - 1).L + 2 * LAB(X + 2, Y).L + LAB(X + 2, Y + 1).L)))

        Lx = (PercLUM * ((-2 * LAB(X - 1, Y - 1).L - 4 * LAB(X, Y - 1).L - 2 * LAB(X + 1, Y - 1).L + 2 * LAB(X - 1, Y + 1).L + 4 * LAB(X, Y + 1).L + 2 * LAB(X + 1, Y + 1).L)))
        Lx = Lx + (PercLUM * ((-LAB(X - 1, Y - 2).L - 2 * LAB(X, Y - 2).L - LAB(X + 1, Y - 2).L + LAB(X - 1, Y + 2).L + 2 * LAB(X, Y + 2).L + LAB(X + 1, Y + 2).L)))

        If PercHUE > 0 Then
          Ay = (PercLUM * (-(-2 * LAB(X - 1, Y - 1).A - 4 * LAB(X - 1, Y).A - 2 * LAB(X - 1, Y + 1).A + 2 * LAB(X + 1, Y - 1).A + 4 * LAB(X + 1, Y).A + 2 * LAB(X + 1, Y + 1).A)))
          Ay = Ay + (PercLUM * (-(-LAB(X - 2, Y - 1).A - 2 * LAB(X - 2, Y).A - LAB(X - 2, Y + 1).A + LAB(X + 2, Y - 1).A + 2 * LAB(X + 2, Y).A + LAB(X + 2, Y + 1).A)))
          Ax = (PercLUM * ((-2 * LAB(X - 1, Y - 1).A - 4 * LAB(X, Y - 1).A - 2 * LAB(X + 1, Y - 1).A + 2 * LAB(X - 1, Y + 1).A + 4 * LAB(X, Y + 1).A + 2 * LAB(X + 1, Y + 1).A)))
          Ax = Ax + (PercLUM * ((-LAB(X - 1, Y - 2).A - 2 * LAB(X, Y - 2).A - LAB(X + 1, Y - 2).A + LAB(X - 1, Y + 2).A + 2 * LAB(X, Y + 2).A + LAB(X + 1, Y + 2).A)))

          By = (PercLUM * (-(-2 * LAB(X - 1, Y - 1).B - 4 * LAB(X - 1, Y).B - 2 * LAB(X - 1, Y + 1).B + 2 * LAB(X + 1, Y - 1).B + 4 * LAB(X + 1, Y).B + 2 * LAB(X + 1, Y + 1).B)))
          By = By + (PercLUM * (-(-LAB(X - 2, Y - 1).B - 2 * LAB(X - 2, Y).B - LAB(X - 2, Y + 1).B + LAB(X + 2, Y - 1).B + 2 * LAB(X + 2, Y).B + LAB(X + 2, Y + 1).B)))
          Bx = (PercLUM * ((-2 * LAB(X - 1, Y - 1).B - 4 * LAB(X, Y - 1).B - 2 * LAB(X + 1, Y - 1).B + 2 * LAB(X - 1, Y + 1).B + 4 * LAB(X, Y + 1).B + 2 * LAB(X + 1, Y + 1).B)))
          Bx = Bx + (PercLUM * ((-LAB(X - 1, Y - 2).B - 2 * LAB(X, Y - 2).B - LAB(X + 1, Y - 2).B + LAB(X - 1, Y + 2).B + 2 * LAB(X, Y + 2).B + LAB(X + 1, Y + 2).B)))
        End If

        dL = Sqrt(Lx * Lx + Ly * Ly)
        dA = Sqrt(Ax * Ax + Ay * Ay)
        dB = Sqrt(Bx * Bx + By * By)

        II = (dL * dL + dA * dA + dB * dB) ^ 0.95
        II = II * ContAmount ^ (1 / 0.95)

        ContByte(0, X, Y) = zLimitMax255(II)

      Next

      If X > ProgX Then
        RaiseEvent PercDONE("Contour", X / pW, 0)
        ProgX = ProgX + ProgXStep
      End If

    Next

    RaiseEvent PercDONE("Contour", 1, 0)

  End Sub

  Public Sub zEFF_ContourbyDOG(ByVal Contour_0_100 As Single, ByVal Thresh As Single)
    Dim X As Long
    Dim Y As Long
    Dim XP As Long
    Dim YP As Long

    Dim Xfrom As Long
    Dim Xto As Long
    Dim Yfrom As Long
    Dim Yto As Long

    Dim D As Single
    Dim Dsum As Single

    Dim R As Single
    Dim G As Single
    Dim B As Single

    ReDim ContByte(0 To pB, 0 To pW, 0 To pH)
    ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)
    ReDim COMPUTEsingle2(0 To pB, 0 To pW, 0 To pH)

    Dim XYstep

    XYstep = 3

    zInit_SpatialDomain(6, Fast_SpatialDomain2)

    Dsum = 0
    For X = -XYstep To XYstep
      For Y = -XYstep To XYstep
        Dsum = Dsum + Fast_SpatialDomain2(X, Y)
      Next
    Next

    For X = XYstep To pW - XYstep
      Xfrom = X - XYstep : Xto = X + XYstep
      For Y = XYstep To pH - XYstep
        Yfrom = Y - XYstep : Yto = Y + XYstep
        R = 0
        G = 0
        B = 0

        For XP = Xfrom To Xto
          For YP = Yfrom To Yto

            D = Fast_SpatialDomain2(X - XP, Y - YP)

            R = R + BILAByte(2, XP, YP) * D
            G = G + BILAByte(1, XP, YP) * D
            B = B + BILAByte(0, XP, YP) * D
          Next
        Next

        R = R * 0.299 + G * 0.587 + B * 0.114
        R = R / Dsum
        COMPUTEsingle(2, X, Y) = R
      Next
    Next

    zInit_SpatialDomain(2, Fast_SpatialDomain2)

    Dsum = 0

    For X = -XYstep To XYstep
      For Y = -XYstep To XYstep
        Dsum = Dsum + Fast_SpatialDomain2(X, Y)
      Next
    Next

    For X = XYstep To pW - XYstep
      Xfrom = X - XYstep : Xto = X + XYstep
      For Y = XYstep To pH - XYstep
        Yfrom = Y - XYstep : Yto = Y + XYstep
        R = 0
        G = 0
        B = 0

        For XP = Xfrom To Xto
          For YP = Yfrom To Yto

            D = Fast_SpatialDomain2(X - XP, Y - YP)

            R = R + BILAByte(2, XP, YP) * D
            G = G + BILAByte(1, XP, YP) * D
            B = B + BILAByte(0, XP, YP) * D
          Next
        Next
        R = R * 0.299 + G * 0.587 + B * 0.114
        R = R / Dsum
        COMPUTEsingle2(2, X, Y) = R
      Next
    Next

    For X = 0 To pW
      For Y = 0 To pH
        R = Contour_0_100 * (COMPUTEsingle(2, X, Y) - COMPUTEsingle2(2, X, Y) - Thresh) * 0.5
        If R < 0 Then R = 0

        R = R * 0.4
        If R > 255 Then R = 255

        ContByte(2, X, Y) = R
        ContByte(1, X, Y) = R
        ContByte(0, X, Y) = R
      Next
    Next

  End Sub

  Public Sub zEFF_ContourByLoG(ByVal Contour_0_100 As Single, ByVal Thresh As Single)
    'by Laplacian of Gaussian
    Dim X As Long
    Dim Y As Long
    Dim XP As Long
    Dim YP As Long

    Dim Xfrom As Long
    Dim Xto As Long
    Dim Yfrom As Long
    Dim Yto As Long

    Dim D As Single
    Dim Dsum As Single

    Dim R As Single
    Dim G As Single
    Dim B As Single

    ReDim ContByte(0 To pB, 0 To pW, 0 To pH)
    ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)

    Dim XYstep

    XYstep = 2

    zInit_SpatialDomain2LoG(0.75)

    Dsum = 0
    For X = -XYstep To XYstep
      For Y = -XYstep To XYstep
        Dsum = Dsum + Fast_SpatialDomain2(X, Y)
      Next
    Next

    Dsum = Abs(Dsum)

    For X = XYstep To pW - XYstep
      Xfrom = X - XYstep : Xto = X + XYstep
      For Y = XYstep To pH - XYstep
        Yfrom = Y - XYstep : Yto = Y + XYstep
        R = 0
        G = 0
        B = 0

        For XP = Xfrom To Xto
          For YP = Yfrom To Yto

            D = Fast_SpatialDomain2(X - XP, Y - YP)

            R = R + BILAByte(2, XP, YP) * D
            G = G + BILAByte(1, XP, YP) * D
            B = B + BILAByte(0, XP, YP) * D
          Next
        Next

        R = R * 0.299 + G * 0.587 + B * 0.114

        COMPUTEsingle(2, X, Y) = R

      Next
    Next


    For X = 0 To pW
      For Y = 0 To pH
        R = Contour_0_100 * (COMPUTEsingle(2, X, Y))
        If R < 0 Then R = 0
        R = R * 0.0025

        If R > 255 Then R = 255
        If R < Thresh * 255 Then R = 0

        ContByte(2, X, Y) = R
        ContByte(1, X, Y) = R
        ContByte(0, X, Y) = R
      Next

    Next

  End Sub

  Public Sub zEFF_Contour_Apply()
    Dim X As Long
    Dim Y As Long

    For X = 0 + 1 To pW - 1
      For Y = 0 + 1 To pH - 1
        BILAByte(0, X, Y) = zLimitMin0(BILAByte(0, X, Y) \ 1 - ContByte(0, X, Y) \ 1)
        BILAByte(1, X, Y) = zLimitMin0(BILAByte(1, X, Y) \ 1 - ContByte(0, X, Y) \ 1)
        BILAByte(2, X, Y) = zLimitMin0(BILAByte(2, X, Y) \ 1 - ContByte(0, X, Y) \ 1)
      Next
    Next

  End Sub

#End Region

#Region "Effect - Bilateral"

  Public Sub zEFF_BilateralFilter(ByVal N As Long, ByVal Iterations As Long, ByVal RGBmode As Boolean, Optional ByVal Directional As Boolean = False)
    'Author :Roberto Mior
    '     reexre@gmail.com
    '
    'If you use source code or part of it please cite the author
    'You can use this code however you like providing the above credits remain intact

    Radius = N

    Const d100 As Single = 1 / 100

    Dim I As Long

    Dim X As Long
    Dim Y As Long
    Dim ProgX As Long
    Dim ProgXStep As Long

    Dim XP As Long
    Dim YP As Long
    Dim XmN As Long
    Dim XpN As Long
    Dim YmN As Long
    Dim YpN As Long

    Dim dR As Single
    Dim dG As Single
    Dim dB As Single
    Dim TR As Long
    Dim TG As Long
    Dim TB As Long

    Dim RDiv As Single
    Dim GDiv As Single
    Dim BDiv As Single

    Dim SpatialD As Single

    Dim LL As Single
    Dim AA As Single
    Dim BB As Single
    Dim RRR As Single
    Dim GGG As Single
    Dim BBB As Single

    Dim Xfrom As Long
    Dim Xto As Long
    Dim Yfrom As Long
    Dim Yto As Long

    Dim XPmX As Long

    Dim NDX As Long
    Dim NDY As Long

    ReDim LAB(0 To pW, 0 To pH)

    ProgXStep = Round(3 * pW / 100)
    ProgX = 0

    If RGBmode Then
      ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
      ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
      ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)
      For X = 0 To pW
        For Y = 0 To pH
          COMPUTEsingle(2, X, Y) = CSng(Sbyte1(2, X, Y)) * 100
          COMPUTEsingle(1, X, Y) = CSng(Sbyte1(1, X, Y)) * 100
          COMPUTEsingle(0, X, Y) = CSng(Sbyte1(0, X, Y)) * 100
        Next
        If X > ProgX Then
          RaiseEvent PercDONE("Bilateral INIT", X / pW, 0)
          ProgX = ProgX + ProgXStep
        End If
      Next
    Else
      ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
      ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
      ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)

      For X = 0 To pW
        For Y = 0 To pH
          RGB_CieLAB(Sbyte1(2, X, Y), Sbyte1(1, X, Y), Sbyte1(0, X, Y), LL, AA, BB)
          COMPUTEsingle(2, X, Y) = LL * 100
          COMPUTEsingle(1, X, Y) = AA * 100
          COMPUTEsingle(0, X, Y) = BB * 100

          BILASingle(2, X, Y) = 12750
          BILASingle(1, X, Y) = AA * 100
          BILASingle(0, X, Y) = BB * 100

          BILAByte(1, X, Y) = AA
          BILAByte(0, X, Y) = BB
        Next
        If X > ProgX Then
          RaiseEvent PercDONE("Bilateral INIT", X / pW, 0)
          ProgX = ProgX + ProgXStep
        End If
      Next

    End If

    ProgXStep = Round(2 * pW / (100 / Iterations))
    Xfrom = 0 + N
    Xto = pW - N
    Yfrom = 0 + N
    Yto = pH - N

    'Classic MODE
    If RGBmode Then
      For I = 1 To Iterations
        ProgX = 0
        For X = Xfrom To Xto
          XmN = X - N
          XpN = X + N
          For Y = Yfrom To Yto
            TR = 0
            TG = 0
            TB = 0
            RDiv = 0
            GDiv = 0
            BDiv = 0
            YmN = Y - N
            YpN = Y + N
            For XP = XmN To XpN
              XPmX = XP - X
              For YP = YmN To YpN

                dR = Fast_ExpIntensity((COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y)) + 25500)
                dG = Fast_ExpIntensity((COMPUTEsingle(1, XP, YP) - COMPUTEsingle(1, X, Y)) + 25500)
                dB = Fast_ExpIntensity((COMPUTEsingle(0, XP, YP) - COMPUTEsingle(0, X, Y)) + 25500)

                SpatialD = Fast_SpatialDomain(XPmX + 30, (YP - Y) + 30)

                dR = dR * SpatialD
                dG = dG * SpatialD
                dB = dB * SpatialD

                TR = TR + (COMPUTEsingle(2, XP, YP)) * dR
                TG = TG + (COMPUTEsingle(1, XP, YP)) * dG
                TB = TB + (COMPUTEsingle(0, XP, YP)) * dB

                RDiv = RDiv + dR
                GDiv = GDiv + dG
                BDiv = BDiv + dB

              Next YP
            Next XP

            TR = TR / RDiv
            TG = TG / GDiv
            TB = TB / BDiv

            BILASingle(2, X, Y) = TR
            BILASingle(1, X, Y) = TG
            BILASingle(0, X, Y) = TB

          Next Y

          If X > ProgX Then
            RaiseEvent PercDONE("Bilateral", (I - 1) / Iterations + (X / pW) / Iterations, I)
            ProgX = ProgX + ProgXStep
          End If

        Next X

        COMPUTEsingle = BILASingle

      Next I

    Else
      'CIE LAB faster mode
      ' Compute only byte "2" that represents the LL Luminance of LL AA BB

      For I = 1 To Iterations
        ProgX = 0
        For X = Xfrom To Xto
          XmN = X - N
          XpN = X + N
          For Y = Yfrom To Yto
            TR = 0
            RDiv = 0
            YmN = Y - N
            YpN = Y + N
            For XP = XmN To XpN
              XPmX = XP - X
              For YP = YmN To YpN

                dR = Fast_ExpIntensity((COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y)) + 25500)
                SpatialD = Fast_SpatialDomain(XPmX + 30, (YP - Y) + 30)
                dR = dR * SpatialD
                TR = TR + (COMPUTEsingle(2, XP, YP)) * dR
                RDiv = RDiv + dR
              Next YP
            Next XP
            TR = TR / RDiv
            BILASingle(2, X, Y) = TR
          Next Y

          If X > ProgX Then
            RaiseEvent PercDONE("Bilateral", (I - 1) / Iterations + (X / pW) / Iterations, I)
            ProgX = ProgX + ProgXStep
          End If

        Next X

        COMPUTEsingle = BILASingle

      Next I
    End If


    If RGBmode Then
      For X = 0 To pW
        For Y = 0 To pH
          dR = COMPUTEsingle(2, X, Y) * d100
          dG = COMPUTEsingle(1, X, Y) * d100
          dB = COMPUTEsingle(0, X, Y) * d100
          RGB_CieLAB(dR, dG, dB, LAB(X, Y).L, LAB(X, Y).A, LAB(X, Y).B)

          BILAByte(2, X, Y) = dR
          BILAByte(1, X, Y) = dG
          BILAByte(0, X, Y) = dB

        Next
      Next
    Else

      For X = 0 To pW
        For Y = 0 To pH

          LL = COMPUTEsingle(2, X, Y) * d100
          AA = COMPUTEsingle(1, X, Y) * d100
          BB = COMPUTEsingle(0, X, Y) * d100
          LAB(X, Y).L = LL
          LAB(X, Y).A = AA
          LAB(X, Y).B = BB
          CieLAB_RGB(LL, AA, BB, RRR, GGG, BBB)

          BILAByte(2, X, Y) = RRR
          BILAByte(1, X, Y) = GGG
          BILAByte(0, X, Y) = BBB
        Next
      Next

    End If

    RaiseEvent PercDONE("Bilateral", 1, Iterations)

  End Sub

  'Public Sub zEFF_BilateralFilterDirectional(ByVal N As Long, ByVal Iterations As Long, ByVal RGBmode As Boolean)
  '  'Author :Roberto Mior
  '  '     reexre@gmail.com
  '  '
  '  'If you use source code or part of it please cite the author
  '  'You can use this code however you like providing the above credits remain intact
  '  '
  '  '
  '  '
  '  '
  '  Radius = N

  '  Const d100 As Single = 1 / 100

  '  Dim I As Long

  '  Dim X As Long
  '  Dim Y As Long
  '  Dim ProgX As Long        'For Progress Bar
  '  Dim ProgXStep As Long        'For Progress Bar

  '  Dim XP As Long
  '  Dim YP As Long
  '  Dim XmN As Long
  '  Dim XpN As Long
  '  Dim YmN As Long
  '  Dim YpN As Long

  '  Dim dR As Single
  '  Dim dG As Single
  '  Dim dB As Single
  '  Dim TR As Long
  '  Dim TG As Long
  '  Dim TB As Long

  '  Dim RDiv As Single
  '  Dim GDiv As Single
  '  Dim BDiv As Single

  '  Dim SpatialD As Single

  '  Dim LL As Single
  '  Dim AA As Single
  '  Dim BB As Single

  '  Dim Xfrom As Long
  '  Dim Xto As Long
  '  Dim Yfrom As Long
  '  Dim Yto As Long

  '  Dim XPmX As Long

  '  Dim NDX As Long
  '  Dim NDY As Long

  '  Dim CosA As Single
  '  Dim SinA As Single

  '  ReDim LAB(0 To pW, 0 To pH)

  '  If RGBmode Then
  '    ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
  '    ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
  '    ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)
  '    For X = 0 To pW
  '      For Y = 0 To pH
  '        COMPUTEsingle(2, X, Y) = CSng(Sbyte1(2, X, Y)) * 100
  '        COMPUTEsingle(1, X, Y) = CSng(Sbyte1(1, X, Y)) * 100
  '        COMPUTEsingle(0, X, Y) = CSng(Sbyte1(0, X, Y)) * 100
  '      Next
  '    Next
  '  Else
  '    ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
  '    ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
  '    ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)

  '    For X = 0 To pW
  '      For Y = 0 To pH
  '        RGB_CieLAB(Sbyte1(2, X, Y), Sbyte1(1, X, Y), Sbyte1(0, X, Y), LL, AA, BB)
  '        COMPUTEsingle(2, X, Y) = LL * 100
  '        COMPUTEsingle(1, X, Y) = AA * 100
  '        COMPUTEsingle(0, X, Y) = BB * 100

  '        'for black border (LAB)
  '        BILASingle(2, X, Y) = 12750    '0

  '        BILASingle(1, X, Y) = 12750
  '        BILASingle(0, X, Y) = 12750

  '        BILAByte(1, X, Y) = AA
  '        BILAByte(0, X, Y) = BB
  '      Next
  '    Next

  '  End If






  '  ProgXStep = Round(pW / (100 / Iterations))
  '  Xfrom = 0 + N * 2
  '  Xto = pW - N * 2
  '  Yfrom = 0 + N * 2
  '  Yto = pH - N * 2

  '  '-------------------------------------------------------------------------------------
  '  'Classic MODE
  '  If RGBmode Then
  '    For I = 1 To Iterations
  '      TEST2()
  '      ProgX = 0
  '      'If Directional Then ComputeSlopes
  '      For X = Xfrom To Xto
  '        'For X = 0 To pW
  '        XmN = X - N
  '        XpN = X + N
  '        For Y = Yfrom To Yto
  '          'For Y = 0 To PH
  '          TR = 0
  '          TG = 0
  '          TB = 0
  '          RDiv = 0
  '          GDiv = 0
  '          BDiv = 0
  '          YmN = Y - N
  '          YpN = Y + N
  '          For XP = XmN To XpN
  '            XPmX = XP - X
  '            For YP = YmN To YpN

  '              dR = Fast_ExpIntensity(COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y))
  '              dG = Fast_ExpIntensity(COMPUTEsingle(1, XP, YP) - COMPUTEsingle(1, X, Y))
  '              dB = Fast_ExpIntensity(COMPUTEsingle(0, XP, YP) - COMPUTEsingle(0, X, Y))


  '              SpatialD = Fast_SpatialDomain(XPmX, YP - Y)

  '              ' If Directional Then SpatialD = SpatialD * Vec(xP, yP).L
  '              '  SpatialD = Fast_SpatialDomain((xP - X) * (N - Vec(xP, yP).X) / N, (yP - Y) * (N - Vec(xP, yP).Y) / N)

  '              dR = dR * SpatialD
  '              dG = dG * SpatialD
  '              dB = dB * SpatialD

  '              TR = TR + (COMPUTEsingle(2, XP, YP)) * dR
  '              TG = TG + (COMPUTEsingle(1, XP, YP)) * dG
  '              TB = TB + (COMPUTEsingle(0, XP, YP)) * dB

  '              RDiv = RDiv + dR
  '              GDiv = GDiv + dG
  '              BDiv = BDiv + dB

  '            Next YP
  '          Next XP

  '          TR = TR / RDiv
  '          TG = TG / GDiv
  '          TB = TB / BDiv

  '          ''   TR = IIf(TR < 255000, TR, 255000)
  '          ''   TG = IIf(TG < 255000, TG, 255000)
  '          ''   TB = IIf(TB < 255000, TB, 255000)

  '          BILASingle(2, X, Y) = TR
  '          BILASingle(1, X, Y) = TG
  '          BILASingle(0, X, Y) = TB

  '        Next Y

  '        ' for the progress bar
  '        If X > ProgX Then
  '          RaiseEvent PercDONE("Bilateral", (I - 1) / Iterations + (X / pW) / Iterations, I)
  '          ProgX = ProgX + ProgXStep
  '        End If

  '      Next X

  '      'CopyMemory ByVal VarPtr(COMPUTEsingle(0, 0, 0)), ByVal VarPtr(BILASingle(0, 0, 0)), (CLng(pB + 1) * CLng(pW + 1) * CLng(pH + 1)) * 4    'Single Lenght=4 Bytes
  '      COMPUTEsingle = BILASingle

  '    Next I

  '  Else
  '    '------------------------------------------------------------------------------------
  '    'CIE LAB faster mode
  '    ' Compute only byte "2" that rappresent the LL Luminance of LL AA BB

  '    For I = 1 To Iterations
  '      TEST2()

  '      ProgX = 0
  '      'If Directional Then ComputeSlopes
  '      For X = Xfrom To Xto
  '        XmN = X - N
  '        XpN = X + N
  '        For Y = Yfrom To Yto
  '          TR = 0
  '          RDiv = 0
  '          YmN = Y - N
  '          YpN = Y + N
  '          CosA = Vec(X, Y).CosA
  '          SinA = Vec(X, Y).SinA
  '          For XP = XmN To XpN
  '            XPmX = XP - X
  '            For YP = YmN To YpN



  '              NDX = X + 2 * (XP - X) * CosA + 0.5 * (YP - Y) * SinA
  '              NDY = Y - 2 * (XP - X) * SinA + 0.5 * (YP - Y) * CosA


  '              dR = Fast_ExpIntensity(COMPUTEsingle(2, NDX, NDY) - COMPUTEsingle(2, X, Y))
  '              'SpatialD = Fast_SpatialDomain(NDX - X, NDY - Y)
  '              'dR = dR * SpatialD
  '              TR = TR + (COMPUTEsingle(2, NDX, NDY)) * dR


  '              RDiv = RDiv + dR
  '            Next YP
  '          Next XP

  '          TR = TR / RDiv

  '          BILASingle(2, X, Y) = TR
  '        Next Y

  '        If X > ProgX Then
  '          RaiseEvent PercDONE("Bilateral", (I - 1) / Iterations + (X / pW) / Iterations, I)
  '          ProgX = ProgX + ProgXStep
  '        End If

  '      Next X

  '      COMPUTEsingle = BILASingle

  '    Next I
  '  End If


  '  If RGBmode Then
  '    For X = 0 To pW
  '      For Y = 0 To pH
  '        dR = COMPUTEsingle(2, X, Y) * d100
  '        dG = COMPUTEsingle(1, X, Y) * d100
  '        dB = COMPUTEsingle(0, X, Y) * d100
  '        RGB_CieLAB(dR, dG, dB, LAB(X, Y).L, LAB(X, Y).A, LAB(X, Y).B)

  '        BILAByte(2, X, Y) = dR
  '        BILAByte(1, X, Y) = dG
  '        BILAByte(0, X, Y) = dB

  '      Next
  '    Next
  '  Else

  '    For X = 0 To pW
  '      For Y = 0 To pH
  '        'Notice that source AA,BB can be invariant
  '        'This changes only LL . so this method is faster
  '        LAB(X, Y).L = COMPUTEsingle(2, X, Y) * d100
  '        LAB(X, Y).A = BILAByte(1, X, Y)
  '        LAB(X, Y).B = BILAByte(0, X, Y)

  '        CieLAB_RGB(COMPUTEsingle(2, X, Y) * d100, _
  '                   BILAByte(1, X, Y), _
  '                   BILAByte(0, X, Y), LL, AA, BB)


  '        BILAByte(2, X, Y) = LL    '(this is R)
  '        BILAByte(1, X, Y) = AA    '(this is G)
  '        BILAByte(0, X, Y) = BB    '(this is B)
  '      Next
  '    Next

  '  End If

  '  RaiseEvent PercDONE("Bilateral", 1, Iterations)

  'End Sub

  Public Sub zEFF_BilateralFilterNOSPATIAL(ByVal N As Long, ByVal Sigma As Single, ByVal SigmaSpatial As Single, ByVal Iterations As Long, ByVal IntensityMode As Integer, ByVal RGBmode As Boolean, Optional ByVal Directional As Boolean = False)
    'Author :Roberto Mior
    '     reexre@gmail.com
    '
    'If you use source code or part of it please cite the author
    'You can use this code however you like providing the above credits remain intact

    Const d100 As Single = 1 / 100

    Dim I As Long

    Dim X As Long
    Dim Y As Long
    Dim ProgX As Long        'For Progress Bar
    Dim ProgXStep As Long        'For Progress Bar

    Dim XP As Long
    Dim YP As Long
    Dim XmN As Long
    Dim XpN As Long
    Dim YmN As Long
    Dim YpN As Long

    Dim dR As Single
    Dim dG As Single
    Dim dB As Single
    Dim TR As Long
    Dim TG As Long
    Dim TB As Long

    Dim RDiv As Single
    Dim GDiv As Single
    Dim BDiv As Single

    Dim SpatialD As Single

    Dim LL As Single
    Dim AA As Single
    Dim BB As Single

    Dim Xfrom As Long
    Dim Xto As Long
    Dim Yfrom As Long
    Dim Yto As Long

    Dim XPmX As Long

    Dim tUPr As Single
    Dim tUPg As Single
    Dim tUPb As Single
    Dim tUPrDiv As Single
    Dim tUPgDiv As Single
    Dim tUPbDiv As Single

    Dim tDownR As Single
    Dim tDownG As Single
    Dim tDownB As Single
    Dim tDOWNrDiv As Single
    Dim tDOWNgDiv As Single
    Dim tDOWNbDiv As Single

    Dim tTotR As Single
    Dim tTotG As Single
    Dim tTotB As Single
    Dim tTotrDiv As Single
    Dim tTotgDiv As Single
    Dim tTotbDiv As Single


    If RGBmode Then
      ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
      ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
      ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)
      For X = 0 To pW
        For Y = 0 To pH
          COMPUTEsingle(2, X, Y) = CSng(Sbyte1(2, X, Y)) * 100
          COMPUTEsingle(1, X, Y) = CSng(Sbyte1(1, X, Y)) * 100
          COMPUTEsingle(0, X, Y) = CSng(Sbyte1(0, X, Y)) * 100
        Next
      Next
    Else
      ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
      ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
      ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)

      For X = 0 To pW
        For Y = 0 To pH
          RGB_CieLAB(Sbyte1(2, X, Y), Sbyte1(1, X, Y), Sbyte1(0, X, Y), LL, AA, BB)
          COMPUTEsingle(2, X, Y) = LL * 100
          COMPUTEsingle(1, X, Y) = AA * 100
          COMPUTEsingle(0, X, Y) = BB * 100

          'for black border (LAB)
          BILASingle(2, X, Y) = 12750
          BILASingle(1, X, Y) = 12750
          BILASingle(0, X, Y) = 12750

          BILAByte(1, X, Y) = AA
          BILAByte(0, X, Y) = BB
        Next
      Next

    End If

    ProgXStep = Round(pW / (100 / Iterations))
    Xfrom = 0 + N
    Xto = pW - N
    Yfrom = 0 + N
    Yto = pH - N

    'Classic MODE
    If RGBmode Then
      For I = 1 To Iterations
        ProgX = 0
        For X = Xfrom To Xto
          XmN = X - N
          XpN = X + N
          For Y = Yfrom To Yto
            TR = 0
            TG = 0
            TB = 0
            RDiv = 0
            GDiv = 0
            BDiv = 0
            YmN = Y - N
            YpN = Y + N
            For XP = XmN To XpN
              XPmX = XP - X
              For YP = YmN To YpN

                dR = Fast_ExpIntensity((COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y)) + 25500)
                dG = Fast_ExpIntensity((COMPUTEsingle(1, XP, YP) - COMPUTEsingle(1, X, Y)) + 25500)
                dB = Fast_ExpIntensity((COMPUTEsingle(0, XP, YP) - COMPUTEsingle(0, X, Y)) + 25500)

                TR = TR + (COMPUTEsingle(2, XP, YP)) * dR
                TG = TG + (COMPUTEsingle(1, XP, YP)) * dG
                TB = TB + (COMPUTEsingle(0, XP, YP)) * dB

                RDiv = RDiv + dR
                GDiv = GDiv + dG
                BDiv = BDiv + dB

              Next YP
            Next XP

            TR = TR / RDiv
            TG = TG / GDiv
            TB = TB / BDiv

            BILASingle(2, X, Y) = TR
            BILASingle(1, X, Y) = TG
            BILASingle(0, X, Y) = TB

          Next Y

          ' for the progress bar
          If X > ProgX Then
            RaiseEvent PercDONE("Bilateral", (I - 1) / Iterations + (X / pW) / Iterations, I)
            ProgX = ProgX + ProgXStep
          End If

        Next X

        COMPUTEsingle = BILASingle

      Next I

    Else
      'CIE LAB faster mode
      ' Compute only byte "2" that rappresent the LL Luminance of LL AA BB

      For I = 1 To Iterations
        ProgX = 0


        For X = Xfrom To Xto
          XmN = X - N
          XpN = X + N

          tTotR = 0
          tTotrDiv = 0
          Y = N
          For XP = XmN To XpN
            For YP = Y - N + 1 To Y + N - 1
              dR = Fast_ExpIntensity((COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y)) + 25500)
              tTotR = tTotR + (COMPUTEsingle(2, XP, YP)) * dR
              tTotrDiv = tTotrDiv + dR

            Next
          Next

          tDownR = 0
          tDOWNrDiv = 0
          YP = Y + N
          For XP = XmN To XpN
            dR = Fast_ExpIntensity((COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y)) + 25500)
            tDownR = tDownR + (COMPUTEsingle(2, XP, YP)) * dR
            tDOWNrDiv = tDOWNrDiv + dR
          Next

          tUPr = 0
          tUPrDiv = 0
          YP = Y - N
          For XP = XmN To XpN
            dR = Fast_ExpIntensity((COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y)) + 25500)
            tUPr = tUPr + (COMPUTEsingle(2, XP, YP)) * dR
            tUPrDiv = tUPrDiv + dR
          Next

          tTotR = (tTotR + tUPr + tDownR) / (tTotrDiv + tUPrDiv + tDOWNrDiv)

          BILASingle(2, X, Y) = tTotR

          For Y = Yfrom + 1 To Yto
            tDownR = 0
            tDOWNrDiv = 0
            YP = Y + N
            For XP = XmN To XpN
              dR = Fast_ExpIntensity((COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y)) + 25500)
              tDownR = tDownR + (COMPUTEsingle(2, XP, YP)) * dR
              tDOWNrDiv = tDOWNrDiv + dR
            Next XP

            tTotR = (tTotR - tUPr + tDownR) / (tTotrDiv - tUPrDiv + tDOWNrDiv)

            BILASingle(2, X, Y) = tTotR

            tUPr = 0
            tUPrDiv = 0
            YP = Y - N
            For XP = XmN To XpN
              dR = Fast_ExpIntensity((COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y)) + 25500)
              tUPr = tUPr + (COMPUTEsingle(2, XP, YP)) * dR
              tUPrDiv = tUPrDiv + dR
            Next XP
          Next Y

          ' for the progress bar
          If X > ProgX Then
            RaiseEvent PercDONE("Bilateral", (I - 1) / Iterations + (X / pW) / Iterations, I)
            ProgX = ProgX + ProgXStep
          End If

        Next X

        COMPUTEsingle = BILASingle

      Next I
    End If


    If RGBmode Then
      For X = 0 To pW
        For Y = 0 To pH
          BILAByte(2, X, Y) = COMPUTEsingle(2, X, Y) * d100
          BILAByte(1, X, Y) = COMPUTEsingle(1, X, Y) * d100
          BILAByte(0, X, Y) = COMPUTEsingle(0, X, Y) * d100
        Next
      Next
    Else

      For X = 0 To pW
        For Y = 0 To pH
          'But Notice that source AA,BB can be invariant
          'it changes only LL . so this method is fater
          CieLAB_RGB(COMPUTEsingle(2, X, Y) * d100, _
                     BILAByte(1, X, Y), _
                     BILAByte(0, X, Y), LL, AA, BB)

          BILAByte(2, X, Y) = LL    '(this is R)
          BILAByte(1, X, Y) = AA    '(this is G)
          BILAByte(0, X, Y) = BB    '(this is B)
        Next
      Next

    End If

    RaiseEvent PercDONE("Bialteral", 1, Iterations)

  End Sub

#End Region

#Region "Effect - Median"

  Public Sub zEFF_MedianFilter(ByVal N As Long, ByVal Iterations As Long)

    Dim I As Long

    Dim X As Long
    Dim Y As Long

    Dim XP As Long
    Dim YP As Long
    Dim XmN As Long
    Dim XpN As Long
    Dim YmN As Long
    Dim YpN As Long

    Dim TR As Long
    Dim TG As Long
    Dim TB As Long

    Dim RR() As Byte
    Dim GG() As Byte
    Dim BB() As Byte
    Dim T As Byte

    Dim Area As Long
    Dim MidP As Long

    Dim C As Long
    Dim CC As Long

    Dim Xfrom As Long
    Dim Xto As Long
    Dim Yfrom As Long
    Dim Yto As Long

    Area = (N * 2 + 1) ^ 2

    ReDim RR(Area)
    ReDim GG(Area)
    ReDim BB(Area)
    MidP = Area \ 2 + 1

    ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)

    For I = 1 To Iterations
      Xfrom = 0 + N
      Xto = pW - N
      For X = Xfrom To Xto

        XmN = X - N
        XpN = X + N

        Yfrom = 0 + N
        Yto = pH - N

        For Y = Yfrom To Yto

          TR = 0
          TG = 0
          TB = 0

          YmN = Y - N
          YpN = Y + N
          C = 0
          For XP = XmN To XpN
            For YP = YmN To YpN

              C = C + 1

              RR(C) = Sbyte1(2, XP, YP)
              GG(C) = Sbyte1(1, XP, YP)
              BB(C) = Sbyte1(0, XP, YP)

              CC = C
              While (CC > 0) And (RR(CC) < RR(CC - 1))
                T = RR(CC)
                RR(CC) = RR(CC - 1)
                RR(CC - 1) = T
                CC = CC - 1
              End While

              CC = C
              While (CC > 0) And (GG(CC) < GG(CC - 1))
                T = GG(CC)
                GG(CC) = GG(CC - 1)
                GG(CC - 1) = T
                CC = CC - 1
              End While

              CC = C
              While (CC > 0) And (BB(CC) < BB(CC - 1))
                T = BB(CC)
                BB(CC) = BB(CC - 1)
                BB(CC - 1) = T
                CC = CC - 1
              End While

            Next
          Next

          BILAByte(2, X, Y) = RR(MidP)
          BILAByte(1, X, Y) = GG(MidP)
          BILAByte(0, X, Y) = BB(MidP)

        Next
      Next

      Sbyte1 = BILAByte

    Next

  End Sub

#End Region

#Region "Effect - Blur"

  Public Sub zEFF_BLUR(ByVal N As Long, ByVal Iterations As Long)
    Dim I As Long

    Dim X As Long
    Dim Y As Long

    Dim XP As Long
    Dim YP As Long
    Dim XmN As Long
    Dim XpN As Long
    Dim YmN As Long
    Dim YpN As Long

    Dim TR As Long
    Dim TG As Long
    Dim TB As Long

    Dim RR As Long
    Dim GG As Long
    Dim BB As Long

    Dim Area As Long
    Dim C As Long

    Dim Xfrom As Long
    Dim Xto As Long
    Dim Yfrom As Long
    Dim Yto As Long

    Area = (N * 2 + 1) ^ 2

    ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)

    For I = 1 To Iterations
      Xfrom = 0 + N
      Xto = pW - N
      For X = Xfrom To Xto

        XmN = X - N
        XpN = X + N

        Yfrom = 0 + N
        Yto = pH - N

        For Y = Yfrom To Yto

          TR = 0
          TG = 0
          TB = 0

          YmN = Y - N
          YpN = Y + N
          C = 0
          RR = 0
          GG = 0
          BB = 0

          For XP = XmN To XpN
            For YP = YmN To YpN
              RR = RR + Sbyte1(2, XP, YP) \ 1
              GG = GG + Sbyte1(1, XP, YP) \ 1
              BB = BB + Sbyte1(0, XP, YP) \ 1
            Next
          Next

          BILAByte(2, X, Y) = RR \ Area
          BILAByte(1, X, Y) = GG \ Area
          BILAByte(0, X, Y) = BB \ Area

        Next
      Next

      Sbyte1 = BILAByte

    Next

  End Sub

#End Region

#Region "Effect - Quantize Luminance"

  Public Sub zEFF_QuantizeLuminance(ByVal Segments As Long, ByVal Presence As Single, ByVal Radius As Integer, ByVal IsThisVideo As Boolean, Optional ByVal Display As Boolean = False)

    Dim X As Long
    Dim Y As Long
    Dim R As Single
    Dim G As Single
    Dim B As Single

    Dim Histo(0 To 255) As Double
    Dim blurHisto(0 To 255) As Double

    Dim Cache As New tHistoCache

    Dim HistoMax As Double

    Dim Area As Double
    Dim SegmentLimit As Double

    Dim I As Long
    Dim prI As Long
    Dim J As Long
    Dim K As Long
    Dim ws As Double
    Dim S2 As Double

    Dim ACC As Double

    Dim WeiSUM As Long
    Dim StartI As Long

    Dim NotPres As Single

    Dim Xfrom As Long
    Dim Xto As Long
    Dim Yfrom As Long
    Dim Yto As Long

    Dim Xm1 As Long
    Dim Xp1 As Long
    Dim Ym1 As Long
    Dim Yp1 As Long

    Dim Curseg As Long

    ReDim Cache.Seg(0 To Segments - 1)
    Dim S As Single
    Dim E As Single
    Dim V As Single

    If Presence <= 0 Then Exit Sub
    If Presence > 1 Then Presence = 1
    NotPres = 1 - Presence

    If Segments < 2 Then Exit Sub

    RaiseEvent PercDONE("Lum Segment", 0.2, 0)

    'Fine quantization---------------------------------------------------------

    Xfrom = Radius
    Yfrom = Radius
    Xto = pW - Radius
    Yto = pH - Radius

    For X = Xfrom To Xto
      For Y = Yfrom To Yto
        Histo(LAB(X, Y).L \ 1) = Histo(LAB(X, Y).L \ 1) + 1
      Next
    Next

    RaiseEvent PercDONE("Lum Segment", 0.1, 0)

    Area = (pW + 1) * (pH + 1)
    SegmentLimit = Area / Segments

    ACC = 0
    StartI = 0
    prI = StartI
    For I = StartI To 255
      ACC = ACC + Histo(I)
      If ACC >= SegmentLimit Then
        S2 = 0
        ws = 0
        For K = prI To I
          ws = ws + Histo(K) * K
          S2 = S2 + Histo(K)
        Next
        ws = ws / (S2 + 1)
        WeiSUM = ws

        If IsThisVideo Then
          Curseg = Curseg + 1
          Cache.Seg(Curseg - 1).Start = prI
          Cache.Seg(Curseg - 1).End = I
          Cache.Seg(Curseg - 1).Value = WeiSUM
          If Curseg = 1 Then
            Cache.Seg(Curseg - 1).Start = 0
            Cache.Seg(Curseg - 1).End = I
          ElseIf Curseg = Segments Then
            Cache.Seg(Curseg - 1).Start = prI
            Cache.Seg(Curseg - 1).End = 255
          End If
        End If

        For J = prI To I
          Histo(J) = WeiSUM
        Next
        prI = I + 1
        ACC = ACC - SegmentLimit
      End If
    Next

    I = 255

    S2 = 0
    ws = 0
    For K = prI To I
      ws = ws + Histo(K) * K
      S2 = S2 + Histo(K)
    Next
    ws = ws / (S2 + 1)
    WeiSUM = ws

    If IsThisVideo Then
      Curseg = Curseg + 1
      Cache.Seg(Curseg - 1).Start = prI
      Cache.Seg(Curseg - 1).End = I
      Cache.Seg(Curseg - 1).Value = WeiSUM
      If Curseg = 1 Then
        Cache.Seg(Curseg - 1).Start = 0
        Cache.Seg(Curseg - 1).End = I
      ElseIf Curseg = Segments Then
        Cache.Seg(Curseg - 1).Start = prI
        Cache.Seg(Curseg - 1).End = 255
      End If
    End If

    For J = prI To I
      Histo(J) = WeiSUM
    Next

    RaiseEvent PercDONE("Lum Segment", 0.25, 0)

    If IsThisVideo Then
      HistoCache(hcIDX Mod hcSIZE).Seg = Cache.Seg

      For I = 1 To Segments
        S = 0
        E = 0
        V = 0
        For K = 0 To hcSIZE - 1
          S = S + HistoCache(K).Seg(I).Start
          E = E + HistoCache(K).Seg(I).End
          V = V + HistoCache(K).Seg(I).Value
        Next
        S = S / hcSIZE
        E = E / hcSIZE
        V = V / hcSIZE
        'Stop

        For J = S To E
          Histo(J) = V
        Next
      Next
      hcIDX = hcIDX + 1
    End If

    'Blur Histo
    For I = 0 To 255
      blurHisto(I) = Histo(I)
    Next
    For I = 1 To 255 - 1
      Histo(I) = 0.25 * (blurHisto(I - 1) + 2 * blurHisto(I) + blurHisto(I + 1))
    Next

    'Merge
    For X = 0 To pW
      For Y = 0 To pH
        LAB(X, Y).L = Presence * Histo(LAB(X, Y).L \ 1) + NotPres * LAB(X, Y).L
      Next
    Next

    'reConvert to RGB
    RaiseEvent PercDONE("Lum Segment", 0.5, 0)
    For X = 0 To pW
      For Y = 0 To pH
        With LAB(X, Y)
          CieLAB_RGB(.L, .A, .B, R, G, B)
          BILAByte(2, X, Y) = R
          BILAByte(1, X, Y) = G
          BILAByte(0, X, Y) = B
        End With
      Next
    Next

    RaiseEvent PercDONE("Lum Segment", 1, 0)

  End Sub

#End Region

#Region "Effect - Brightness/Contrast"

  'Author :Roberto Mior

  Public Sub PreBrightNessAndContrast(ByVal what As eSource, ByVal BrightnessM1P1 As Single, ByVal ContrastM1P1 As Single)
    Dim R As Single
    Dim G As Single
    Dim B As Single
    Dim X As Long
    Dim Y As Long

    Dim Cmul As Single

    Dim output(0 To 2, 0 To pW, 0 To pH) As Byte
    Dim inputi(0 To 2, 0 To pW, 0 To pH) As Byte

    inputi = Sbyte1

    Cmul = (Tan((ContrastM1P1 + 1) * Atan(1)))

    For X = 0 To pW
      For Y = 0 To pH

        R = inputi(2, X, Y)
        G = inputi(1, X, Y)
        B = inputi(0, X, Y)

        If BrightnessM1P1 < 0 Then

          R = R * (1 + BrightnessM1P1)
          G = G * (1 + BrightnessM1P1)
          B = B * (1 + BrightnessM1P1)

        Else
          R = R + (255 - R) * BrightnessM1P1
          G = G + (255 - G) * BrightnessM1P1
          B = B + (255 - B) * BrightnessM1P1

        End If

        R = (R - 127) * Cmul + 127
        G = (G - 127) * Cmul + 127
        B = (B - 127) * Cmul + 127

        If R < 0 Then R = 0 Else If R > 255 Then R = 255
        If G < 0 Then G = 0 Else If G > 255 Then G = 255
        If B < 0 Then B = 0 Else If B > 255 Then B = 255

        output(2, X, Y) = R \ 1
        output(1, X, Y) = G \ 1
        output(0, X, Y) = B \ 1

      Next
    Next

    Sbyte1 = output

    Erase inputi
    Erase output

  End Sub

#End Region

#Region "Effect - Brightness/Contrast/Saturation"

  Public Sub MagneKleverBCS(ByVal BRIGHT As Long, ByVal CONTRAST As Long, ByVal SATURATION As Long)
    '(c)2009 by Roy Magne Klever    www.rmklever.com
    ' Delphi to VB6 by reexre

    Dim CR As Long
    Dim Cg As Long
    Dim cB As Long
    Dim X As Long
    Dim Y As Long

    Dim I As Long
    Dim K As Long
    Dim V As Long
    Dim ci1 As Long
    Dim ci2 As Long
    Dim ci3 As Long

    Dim Alpha As Long
    Dim A As Single

    Dim ContrastLut(0 To 255) As Long
    Dim BCLut(0 To 255) As Long
    Dim SATGrays(0 To 767) As Long
    Dim SATAlpha(0 To 255) As Long

    Dim output(0 To 2, 0 To pW, 0 To pH) As Byte


    If CONTRAST = 100 Then CONTRAST = 99
    If CONTRAST > 0 Then
      A = 1 / Cos(CONTRAST * (Pi / 200))
    Else
      A = 1 * Cos(CONTRAST * (Pi / 200))
    End If

    For I = 0 To 255
      V = Round(A * (I - 170) + 170)
      If V > 255 Then V = 255 Else If V < 0 Then V = 0
      ContrastLut(I) = V
    Next

    Alpha = BRIGHT
    For I = 0 To 255
      K = 256 - Alpha
      V = (K + Alpha * I) \ 256
      If V < 0 Then V = 0 Else If V > 255 Then V = 255
      BCLut(I) = ContrastLut(V)
    Next

    For I = 0 To 255
      SATAlpha(I) = (((I + 1) * SATURATION) \ 256)
    Next I

    X = 0
    For I = 0 To 255
      Y = I - SATAlpha(I)
      SATGrays(X) = Y
      X = X + 1
      SATGrays(X) = Y
      X = X + 1
      SATGrays(X) = Y
      X = X + 1
    Next

    For Y = 0 To pH
      For X = 0 To pW

        CR = Sbyte1(2, X, Y)
        Cg = Sbyte1(1, X, Y)
        cB = Sbyte1(0, X, Y)

        V = CR + Cg + cB

        ci1 = SATGrays(V) + SATAlpha(cB)
        ci2 = SATGrays(V) + SATAlpha(Cg)
        ci3 = SATGrays(V) + SATAlpha(CR)
        If ci1 < 0 Then ci1 = 0 Else If ci1 > 255 Then ci1 = 255
        If ci2 < 0 Then ci2 = 0 Else If ci2 > 255 Then ci2 = 255
        If ci3 < 0 Then ci3 = 0 Else If ci3 > 255 Then ci3 = 255
        output(0, X, Y) = BCLut(ci1)
        output(1, X, Y) = BCLut(ci2)
        output(2, X, Y) = BCLut(ci3)

      Next
    Next

    Sbyte1 = output

    Erase output

  End Sub

#End Region

#Region "Effect - Exposure"

  Public Sub MagneKleverExposure(ByVal K As Single)
    '(c)2009 by Roy Magne Klever    www.rmklever.com
    ' Delphi to VB6 by reexre

    Dim I As Long
    Dim X As Long
    Dim Y As Long
    Dim LUT(0 To 255) As Long

    Dim output(0 To 2, 0 To pW, 0 To pH) As Byte

    For I = 0 To 255
      If K < 0 Then
        LUT(I) = I - ((-Round((1 - Exp((I / -128) * (K / 128))) * 256) * (I Xor 255)) \ 256)
      Else
        LUT(I) = I + ((Round((1 - Exp((I / -128) * (K / 128))) * 256) * (I Xor 255)) \ 256)
      End If

      If LUT(I) < 0 Then LUT(I) = 0 Else If LUT(I) > 255 Then LUT(I) = 255
    Next

    For Y = 0 To pH
      For X = 0 To pW
        output(2, X, Y) = LUT(Sbyte1(2, X, Y))
        output(1, X, Y) = LUT(Sbyte1(1, X, Y))
        output(0, X, Y) = LUT(Sbyte1(0, X, Y))
      Next
    Next

    Sbyte1 = output
    Erase output

  End Sub

#End Region

#Region "Effect - Histogram"

  Public Sub MagneKleverfxHistCalc()
    '(c)2009 by Roy Magne Klever    www.rmklever.com
    ' Delphi to VB6 by reexre

    Dim X As Long
    Dim Y As Long
    For X = 0 To 255
      histR(X) = 0
      histG(X) = 0
      histB(X) = 0
    Next
    For Y = 0 To pH
      For X = 0 To pW
        histR(Sbyte1(2, X, Y)) = histR(Sbyte1(2, X, Y)) + 1
        histG(Sbyte1(1, X, Y)) = histG(Sbyte1(1, X, Y)) + 1
        histB(Sbyte1(0, X, Y)) = histB(Sbyte1(0, X, Y)) + 1
      Next

    Next
  End Sub

  Private Function MagneKleverCumSum(ByVal Hist As THistSingle) As THistSingle
    '(c)2009 by Roy Magne Klever    www.rmklever.com
    ' Delphi to VB6 by reexre
    Dim X As Long
    Dim Temp As New THistSingle

    Temp.A(0) = Hist.A(0)
    For X = 1 To 255
      Temp.A(X) = Temp.A(X - 1) + Hist.A(X)
    Next

    MagneKleverCumSum = Temp
  End Function

  Public Sub MagneKleverHistogramEQU(ByVal Z As Single)
    '(c)2009 by Roy Magne Klever    www.rmklever.com
    ' Delphi to VB6 by reexre

    Dim X As Long
    Dim Y As Long

    Dim Q1 As Single
    Dim Q2 As Single
    Dim Q3 As Single

    Dim Hist As New THistSingle
    Dim VCumSumR As New THistSingle
    Dim VCumSumG As New THistSingle
    Dim VCumSumB As New THistSingle

    Dim output(0 To 2, 0 To pW, 0 To pH) As Byte

    MagneKleverfxHistCalc()

    Q1 = 0                        '// RED Channel
    For X = 0 To 255
      Hist.A(X) = histR(X) ^ Z
      Q1 = Q1 + Hist.A(X)
    Next
    VCumSumR = MagneKleverCumSum(Hist)

    Q2 = 0
    For X = 0 To 255
      Hist.A(X) = histG(X) ^ Z
      Q2 = Q2 + Hist.A(X)
    Next
    VCumSumG = MagneKleverCumSum(Hist)

    Q3 = 0
    For X = 0 To 255
      Hist.A(X) = histB(X) ^ Z
      Q3 = Q3 + Hist.A(X)
    Next
    VCumSumB = MagneKleverCumSum(Hist)

    For Y = 0 To pH
      For X = 0 To pW
        output(2, X, Y) = Fix((255 / Q1) * VCumSumR.A(Sbyte1(2, X, Y)))
        output(1, X, Y) = Fix((255 / Q2) * VCumSumG.A(Sbyte1(1, X, Y)))
        output(0, X, Y) = Fix((255 / Q3) * VCumSumB.A(Sbyte1(0, X, Y)))
      Next
    Next
    Sbyte1 = output
    Erase output

  End Sub

#End Region

#Region "Compute Slopes (unused)"

  Public Sub ComputeSlopes(ByVal scrRAD As Integer)
    Dim X As Long
    Dim Y As Long

    ReDim HSP(0 To pW, 0 To pH)
    ReDim Vec(0 To pW, 0 To pH)

    For X = 0 To pW
      For Y = 0 To pH
        With HSP(X, Y)
          RGBtoHSP(Sbyte1(2, X, Y), Sbyte1(1, X, Y), Sbyte1(0, X, Y), .H, .S, .P)
        End With
      Next
    Next


    For X = 1 To pW - 1
      For Y = 1 To pH - 1

        With Vec(X, Y)
          .Y = -(-HSP(X - 1, Y - 1).P - 2 * HSP(X - 1, Y).P - HSP(X - 1, Y + 1).P + HSP(X + 1, Y - 1).P + 2 * HSP(X + 1, Y).P + HSP(X + 1, Y + 1).P)
          .X = (-HSP(X - 1, Y - 1).P - 2 * HSP(X, Y - 1).P - HSP(X + 1, Y - 1).P + HSP(X - 1, Y + 1).P + 2 * HSP(X, Y + 1).P + HSP(X + 1, Y + 1).P)

          .X = Abs(.Y) / 255
          .Y = Abs(.X) / 255

          If .X > Val(scrRAD) Then .X = Val(scrRAD)
          If .Y > Val(scrRAD) Then .Y = Val(scrRAD)
          If .X > 1 Then .X = 1
          If .Y > 1 Then .Y = 1
          .X = 1 - .X
          .Y = 1 - .Y

          .L = (.X * .X + .Y * .Y)

        End With
      Next
    Next

  End Sub

#End Region

#Region "Deconstructor"

  Protected Overrides Sub Finalize()

    Erase Sbyte1
    Erase Sbyte2
    Erase ContByte
    Erase BlurByte
    Erase SepaByte

  End Sub

#End Region

End Class