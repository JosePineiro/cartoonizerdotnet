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

  Public Sub zInit_IntensityDomain(ByVal SigmaI, ByVal Mode)

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
            If Fast_ExpIntensity(V + 25500) < 0 Then Fast_ExpIntensity(V + 25500) = 0 : Cos0 = True

          Else
            Fast_ExpIntensity(V + 25500) = 0
          End If

      End Select

    Next
  End Sub

  Public Sub zInit_SpatialDomain(ByVal SigmaS)
    Dim V As Single
    Dim V2 As Single
    Dim X As Long
    Dim Y As Long
    Dim D As Long

    If SigmaS = 0 Then Exit Sub
    SigmaS = SigmaS * 2

    ReDim Fast_ExpSpatial(30 * 30 * 2)
    For V = 0 To 30 * 30 * 2
      V2 = V
      Fast_ExpSpatial(V) = Exp(-(V2 / (SigmaS)))
    Next

    ReDim Fast_SpatialDomain(0 To 60, 0 To 60)
    For X = -30 To 30
      For Y = -30 To 30
        D = X * X + Y * Y
        Fast_SpatialDomain(X + 30, Y + 30) = Fast_ExpSpatial(D)
      Next
    Next

  End Sub

  Public Sub zInit_SpatialDomain2(ByVal SigmaS)
    Dim V As Single
    Dim V2 As Single
    Dim X As Long
    Dim Y As Long
    Dim D As Long

    If SigmaS = 0 Then Exit Sub
    SigmaS = SigmaS * 2

    ReDim Fast_ExpSpatial(30 * 30 * 2)
    For V = 0 To 30 * 30 * 2
      V2 = V
      Fast_ExpSpatial(V) = Exp(-(V2 / (SigmaS)))
    Next

    ReDim Fast_SpatialDomain2(0 To 60, 0 To 60)
    For X = -30 To 30
      For Y = -30 To 30
        D = X * X + Y * Y
        Fast_SpatialDomain2(X + 30, Y + 30) = Fast_ExpSpatial(D)
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
    KX = 32                   '25

    Dim g As Graphics = Graphics.FromImage(bm)
    For X = 0 To KX
      V = Fast_IntensityDomain(Abs(X * 100)) * 255    '2.55
      x1 = ((bm.Width / KX) * X)    '* 0.1
      y1 = bm.Height - V / ky
      X2 = ((bm.Width / KX) * (X + 1))    '* 0.1
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

    zInit_SpatialDomain(cSigma)
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

    'Resize to hold image data
    ReDim Sbyte1(pB, pW, pH)

    'Get the image data and store into Sbyte array)
    ImageHelper.GetBitmapBits(bm, Sbyte1)

  End Sub

  Public Sub zGet_Effect(ByRef bm As Drawing.Bitmap)

    ImageHelper.SetBitmapBits(bm, BILAByte)

    Erase BILAByte
    Erase BILASingle
    Erase COMPUTEsingle

  End Sub

#End Region

#Region "Effect - Contour"

  Public Sub zEFF_Contour(ByVal Contour_0_100 As Single, ByVal LumHue01 As Single)
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

    ContAmount = 0.00004 * Contour_0_100
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

  Public Sub zEFF_ContourbyDOG(ByVal Contour_0_100 As Single, ByVal LumHue01 As Single)
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

    zInit_SpatialDomain2(5)

    Dsum = 0
    For X = -2 To 2
      For Y = -2 To 2
        Dsum = Dsum + Fast_SpatialDomain2(X + 30, Y + 30)
      Next
    Next

    For X = 2 To pW - 2
      Xfrom = X - 2 : Xto = X + 2
      For Y = 2 To pH - 2
        Yfrom = Y - 2 : Yto = Y + 2
        R = 0
        G = 0
        B = 0

        For XP = Xfrom To Xto
          For YP = Yfrom To Yto

            D = Fast_SpatialDomain2((X - XP) + 30, (Y - YP) + 30)

            R = R + BILAByte(2, XP, YP) * D
            G = G + BILAByte(1, XP, YP) * D
            B = B + BILAByte(0, XP, YP) * D
          Next
        Next
        R = R / Dsum
        G = G / Dsum
        B = B / Dsum
        COMPUTEsingle(2, X, Y) = R
        COMPUTEsingle(1, X, Y) = G
        COMPUTEsingle(0, X, Y) = B

      Next
    Next

    zInit_SpatialDomain2(2)

    Dsum = 0
    For X = -2 To 2
      For Y = -2 To 2
        Dsum = Dsum + Fast_SpatialDomain2(X + 30, Y + 30)
      Next
    Next

    For X = 2 To pW - 2
      Xfrom = X - 2 : Xto = X + 2
      For Y = 2 To pH - 2
        Yfrom = Y - 2 : Yto = Y + 2
        R = 0
        G = 0
        B = 0

        For XP = Xfrom To Xto
          For YP = Yfrom To Yto

            D = Fast_SpatialDomain2((X - XP) + 30, (Y - YP) + 30)

            R = R + BILAByte(2, XP, YP) * D
            G = G + BILAByte(1, XP, YP) * D
            B = B + BILAByte(0, XP, YP) * D
          Next
        Next
        R = R / Dsum
        G = G / Dsum
        B = B / Dsum
        COMPUTEsingle2(2, X, Y) = R
        COMPUTEsingle2(1, X, Y) = G
        COMPUTEsingle2(0, X, Y) = B
      Next
    Next


    For X = 0 To pW
      For Y = 0 To pH
        R = 10 * Abs(COMPUTEsingle(2, X, Y) - COMPUTEsingle2(2, X, Y)) ^ 2
        G = 10 * Abs(COMPUTEsingle(1, X, Y) - COMPUTEsingle2(1, X, Y)) ^ 2
        B = 10 * Abs(COMPUTEsingle(0, X, Y) - COMPUTEsingle2(0, X, Y)) ^ 2
        If R > 255 Then R = 255
        If G > 255 Then G = 255
        If B > 255 Then B = 255
        ContByte(2, X, Y) = R
        ContByte(1, X, Y) = G
        ContByte(0, X, Y) = B
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

  Public Sub zEFF_BilateralFilter(ByVal N As Long, ByVal Sigma As Single, ByVal SigmaSpatial As Single, ByVal Iterations As Long, ByVal IntensityMode As Integer, ByVal RGBmode As Boolean, Optional ByVal Directional As Boolean = False)
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

    Dim LL As Byte
    Dim AA As Byte
    Dim BB As Byte

    Dim Xfrom As Long
    Dim Xto As Long
    Dim Yfrom As Long
    Dim Yto As Long

    Dim XPmX As Long

    ReDim LAB(0 To pW, 0 To pH)

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
          BILASingle(2, X, Y) = 12750    '0

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

          ' for the progress bar
          If X > ProgX Then
            RaiseEvent PercDONE("Bilateral", (I - 1) / Iterations + (X / pW) / Iterations, I)
            ProgX = ProgX + ProgXStep
          End If

        Next X

        COMPUTEsingle = BILASingle

      Next I
      '------------------------------------------------------------------------------------
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
          'Notice that source AA,BB can be invariant
          'it changes only LL . so this method is faster
          LAB(X, Y).L = COMPUTEsingle(2, X, Y) * d100
          LAB(X, Y).A = BILAByte(1, X, Y)
          LAB(X, Y).B = BILAByte(0, X, Y)


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
      HistoCache(hcIDX Mod hcSIZE).Seg = (Cache.Seg) 'Cache.Seg -1 (porting issue, need to find VB.NET equivalent)

      For I = 1 To Segments
        S = 0
        E = 0
        V = 0
        For K = 0 To hcSIZE - 1
          S = S + HistoCache(K).Seg(I - 1).Start
          E = E + HistoCache(K).Seg(I - 1).End
          V = V + HistoCache(K).Seg(I - 1).Value
        Next
        S = S / hcSIZE
        E = E / hcSIZE
        V = V / hcSIZE

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

#Region "Experimental"

  Public Sub TEST(ByRef bm1 As Drawing.Bitmap, ByRef bm2 As Drawing.Bitmap)
    Dim X As Long
    Dim Y As Long
    Dim xX As Long
    Dim yY As Long

    Dim cx As Long
    Dim cy As Long

    Dim cXI As Long
    Dim cYI As Long

    Dim P00 As Single
    Dim P01 As Single
    Dim P10 As Single
    Dim P11 As Single


    Dim x00 As Single
    Dim x01 As Single
    Dim x10 As Single
    Dim x11 As Single

    Dim y00 As Single
    Dim y01 As Single
    Dim y10 As Single
    Dim y11 As Single

    Dim pDX As Single
    Dim pDY As Single

    Dim C20 As Single
    Dim C02 As Single

    Dim C30 As Single
    Dim C03 As Single

    Dim C21 As Single
    Dim C31 As Single
    Dim C12 As Single
    Dim C13 As Single
    Dim C11 As Single

    Dim Csum1 As Single
    Dim Csum2 As Single

    Dim U As Single
    Dim U2 As Single
    Dim U3 As Single

    Dim V As Single
    Dim V2 As Single
    Dim V3 As Single

    Dim Pressure
    Dim Density

    Dim myByteDepth As Integer = ImageHelper.GetBitmapBits(bm1, Sbyte1)

    pW = bm1.Width - 1
    pH = bm1.Height - 1
    pB = myByteDepth - 1

    Dim mul(0 To pB, (pW + 1) * 4, (pH + 1) * 4) As Byte
    Dim out(0 To pB, (pW + 1) * 4, (pH + 1) * 4) As Byte


    For X = 0 To pW - 1
      For Y = 0 To pH - 1

        For xX = X * 4 To (X + 1) * 4
          For yY = Y * 4 To (Y + 1) * 4

            mul(2, xX, yY) = Sbyte1(2, X, Y)
            mul(1, xX, yY) = Sbyte1(1, X, Y)
            mul(0, xX, yY) = Sbyte1(0, X, Y)


          Next
        Next
      Next
    Next

    Stop

    For X = 2 To pW * 4 - 4
      For Y = 2 To pH * 4 - 4


        cx = (X \ 4) * 4
        cy = (Y \ 4) * 4


        cXI = ((X + 4) \ 4) * 4
        cYI = ((Y + 4) \ 4) * 4



        '          Stop

        P00 = mul(2, cx, cy) / 255 / 4
        x00 = cx
        y00 = cy

        P01 = mul(2, cx, cYI) / 255 / 4
        x01 = cx
        y01 = cYI

        P10 = mul(2, cXI, cy) / 255 / 4
        x10 = cXI
        y10 = cy

        P11 = mul(2, cXI, cYI) / 255 / 4
        x11 = cXI
        y11 = cYI

        pDX = P10 - P00
        pDY = P01 - P00

        C20 = 3 * pDX - x10 - 2 * x00
        C02 = 3 * pDY - y01 - 2 * y00
        C30 = -2 * pDX + x10 + x00
        C03 = -2 * pDY + y01 + y00

        Csum1 = P00 + y00 + C02 + C03
        Csum2 = P00 + x00 + C20 + C30

        C21 = 3 * P11 - 2 * x01 - x11 - 3 * Csum1 - C20
        C31 = (-2 * P11 + x01 + x11 + 2 * Csum1) - C30
        C12 = 3 * P11 - 2 * y10 - y11 - 3 * Csum2 - C02
        C13 = (-2 * P11 + y10 + y11 + 2 * Csum2) - C03
        C11 = x01 - C13 - C12 - x00

        U = (X - cx)
        U2 = U * U
        U3 = U * U2
        V = (Y - cy)
        V2 = V * V
        V3 = V * V2

        Density = P00 + x00 * U + y00 * V + _
                  C20 * U2 + C02 * V2 + _
                  C30 * U3 + C03 * V3 + _
                  C21 * U2 * V + C31 * U3 * V + _
                  C12 * U * V2 + C13 * U * V3 + C11 * U * V


        'Pressure = (Density - 1)

        Pressure = Density

        If Pressure < 0 Then Pressure = 0
        If Pressure > 255 Then Pressure = 255
        out(2, X, Y) = Pressure \ 1

      Next
    Next

    ImageHelper.SetBitmapBits(bm2, out)

  End Sub

  Public Sub IWH(ByRef bm1 As Drawing.Bitmap, ByRef bm2 As Drawing.Bitmap)

    Dim I As Long
    Dim K As Long


    Dim X As Long
    Dim Y As Long
    Dim dX As Long
    Dim dY As Long

    Dim X2 As Long
    Dim Y2 As Long

    Dim pX As Long
    Dim pY As Long

    Dim R As Single
    Dim G As Single
    Dim B As Single

    Dim MinI As Single
    Dim MaxI As Single
    Dim Inte As Single

    Dim CR(0 To 8) As Long 'CR(x+4)
    Dim Cg(0 To 8) As Long 'cG(x+4)
    Dim cB(0 To 8) As Long 'cB(x+4)

    Dim VIx As Long
    Dim VIy As Long
    Dim HIx As Long
    Dim HIy As Long

    Dim FOUND As Boolean

    Dim DC(0 To 8) As Long 'DC(x+4)
    Dim S As String

    Dim IT As Long

    Dim NR As Long
    Dim CurR As Long
    Dim Reg() As tREGION

    Dim myByteDepth As Integer = ImageHelper.GetBitmapBits(bm1, Sbyte1)

    pW = bm1.Width - 1
    pH = bm1.Height - 1
    pB = myByteDepth - 1

    Dim out(0 To pB, pW, pH) As Byte
    Dim rnds(0 To pB, pW, pH) As Byte

    Dim blur(0 To pB, 0 To pW, 0 To pH) As Single

    Dim XdirV(0 To 1, 0 To pW, 0 To pH) As Long
    Dim YdirV(0 To 1, 0 To pW, 0 To pH) As Long
    Dim XdirH(0 To 1, 0 To pW, 0 To pH) As Long
    Dim YdirH(0 To 1, 0 To pW, 0 To pH) As Long

    Dim LUM(0 To pW, 0 To pH) As Single

    Dim regC(0 To pB, 0 To pW, 0 To pH) As Single

    Dim vallx(0 To pW, 0 To pH) As Long
    Dim vally(0 To pW, 0 To pH) As Long

    Dim hillx(0 To pW, 0 To pH) As Long
    Dim hilly(0 To pW, 0 To pH) As Long

    ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
    ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
    ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)

    For X = 0 To pW
      For Y = 0 To pH
        COMPUTEsingle(2, X, Y) = CSng(Sbyte1(2, X, Y)) * 100
        COMPUTEsingle(1, X, Y) = CSng(Sbyte1(1, X, Y)) * 100
        COMPUTEsingle(0, X, Y) = CSng(Sbyte1(0, X, Y)) * 100
        rnds(2, X, Y) = Rnd() * 255
        rnds(1, X, Y) = Rnd() * 255
        rnds(0, X, Y) = Rnd() * 255

      Next
    Next

    'PreBlur
    For IT = 1 To 30
      NR = 0
      ReDim Reg(0)


      For X = 1 To pW - 1
        For Y = 1 To pH - 1
          R = 0
          G = 0
          B = 0

          For pX = X - 1 To X + 1
            For pY = Y - 1 To Y + 1

              R = R + COMPUTEsingle(2, pX, pY)
              G = G + COMPUTEsingle(1, pX, pY)
              B = B + COMPUTEsingle(0, pX, pY)
              If pX - X = 0 And pY - Y = 0 Then
                R = R + COMPUTEsingle(2, pX, pY) * 5
                G = G + COMPUTEsingle(1, pX, pY) * 5
                B = B + COMPUTEsingle(0, pX, pY) * 5
              End If

            Next
          Next
          R = R / (9 + 5)
          G = G / (9 + 5)
          B = B / (9 + 5)

          blur(2, X, Y) = R
          blur(1, X, Y) = G
          blur(0, X, Y) = B

          LUM(X, Y) = (R * 0.241) ^ 2 + (G * 0.691) ^ 2 + (B * 0.068) ^ 2

        Next
      Next
      COMPUTEsingle = blur

      'directions
      For X = 1 To pW - 1
        For Y = 1 To pH - 1

          Inte = LUM(X, Y)
          MinI = Inte
          MaxI = Inte

          XdirV(0, X, Y) = 0
          YdirV(0, X, Y) = 0

          XdirH(0, X, Y) = 0
          YdirH(0, X, Y) = 0

          For pX = X - 1 To X + 1
            For pY = Y - 1 To Y + 1

              Inte = LUM(pX, pY)

              If Inte < MinI Then
                MinI = Inte
                XdirV(0, X, Y) = pX - X
                YdirV(0, X, Y) = pY - Y
              End If
              If Inte > MaxI Then
                MaxI = Inte
                XdirH(0, X, Y) = pX - X
                YdirH(0, X, Y) = pY - Y
              End If

            Next
          Next
        Next


      Next

      For X = 0 To pW

        For Y = 0 To pH


          dX = XdirV(0, X, Y)
          dY = YdirV(0, X, Y)

          X2 = X
          Y2 = Y

          Do

            X2 = X2 + dX
            Y2 = Y2 + dY
            dX = XdirV(0, X2, Y2)
            dY = YdirV(0, X2, Y2)

          Loop While Not (dX = 0 And dY = 0)

          vallx(X, Y) = X2
          vally(X, Y) = Y2


          'FIND REGIONS

          FOUND = False
          For I = 1 To NR
            If (Reg(I).cx = X2) Then If (Reg(I).cy = Y2) Then FOUND = True : CurR = I : Exit For
          Next

          If Not (FOUND) Then
            NR = NR + 1
            ReDim Preserve Reg(NR)
            CurR = NR
          End If

          With Reg(CurR)
            .cx = X2
            .cy = Y2
            .NP = .NP + 1
            ReDim Preserve .X(.NP)
            ReDim Preserve .Y(.NP)
            .X(.NP) = X
            .Y(.NP) = Y
          End With

          dX = XdirH(0, X, Y)
          dY = YdirH(0, X, Y)

          X2 = X
          Y2 = Y
          Do
            X2 = X2 + dX
            Y2 = Y2 + dY
            dX = XdirH(0, X2, Y2)
            dY = YdirH(0, X2, Y2)

          Loop While Not (dX = 0 And dY = 0)

          hillx(X, Y) = X2
          hilly(X, Y) = Y2
        Next
      Next

      'Average regions
      For I = 1 To NR
        With Reg(I)
          R = 0
          G = 0
          B = 0
          For K = 1 To .NP
            R = R + blur(2, .X(K), .Y(K))
            G = G + blur(1, .X(K), .Y(K))
            B = B + blur(0, .X(K), .Y(K))
          Next
          R = R / .NP
          G = G / .NP
          B = B / .NP
          R = R * 0.01
          G = G * 0.01
          B = B * 0.01

          For K = 1 To .NP
            regC(2, .X(K), .Y(K)) = R
            regC(1, .X(K), .Y(K)) = G
            regC(0, .X(K), .Y(K)) = B
          Next


        End With
      Next

      For I = -4 To 4
        CR(I + 4) = 100 + Rnd() * 155
        Cg(I + 4) = 100 + Rnd() * 155
        cB(I + 4) = 100 + Rnd() * 155

        If I = 0 Then
          CR(I + 4) = 0
          Cg(I + 4) = 0
          cB(I + 4) = 0
        End If
      Next


      For X = 0 To pW
        For Y = 0 To pH

          I = XdirV(0, X, Y) * 3 + YdirV(0, X, Y) * 1
          DC(I + 4) = DC(I + 4) + 1

          out(2, X, Y) = CR(I + 4)
          out(1, X, Y) = Cg(I + 4)
          out(0, X, Y) = cB(I + 4)

          VIx = vallx(X, Y)
          VIy = vally(X, Y)

          HIx = hillx(X, Y)
          HIy = hilly(X, Y)

          GoTo nn
          If I <> 0 Then
            out(2, X, Y) = (rnds(2, VIx, VIy))
            out(1, X, Y) = (rnds(1, VIx, VIy))
            out(0, X, Y) = (rnds(0, VIx, VIy))
          Else
            out(2, X, Y) = 0
            out(1, X, Y) = 0
            out(0, X, Y) = 0
          End If

nn:

          out(2, X, Y) = regC(2, X, Y)
          out(1, X, Y) = regC(1, X, Y)
          out(0, X, Y) = regC(0, X, Y)

        Next

      Next

skip:

      ImageHelper.SetBitmapBits(bm2, out)

      S = "Iter " & IT & vbCrLf & vbCrLf
      S = S + "Regions " & NR & vbCrLf & vbCrLf

      For I = -4 To 4
        S = S + "Dir " & I & "   " & DC(I + 4) & vbCrLf
      Next
      MsgBox(S)

    Next IT
  End Sub

#End Region

End Class