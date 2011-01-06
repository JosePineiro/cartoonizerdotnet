Imports System.Math
Imports System.Drawing

Public Class Sketch

  Private Structure hBitmap
    Public bmType As Long
    Public bmWidth As Long
    Public bmHeight As Long
    Public bmWidthBytes As Long
    Public bmPlanes As Integer
    Public bmBitsPixel As Integer
    Public bmBits As Long
  End Structure



  Private hBmp As hBitmap

  Private pW As Integer
  Private pH As Integer
  Private pB As Integer


  Private Const PI As Single = 3.14159265358979


  Private Sbyte1(0, 0, 0) As Byte
  Private ResSingle(0, 0, 0) As Byte
  Private ANGULAR(0, 0, 0, 0) As Single

  Private GaborFilter(0 To 6, 0 To 6, 0 To 15) As Single 'x+3 , y+3


  Public Event PercDONE(ByVal PercValue As Single, ByVal CurrIteration As Long)

  Private Function max(ByVal A, ByVal B)
    max = IIf(A > B, A, B)

  End Function

  Private Function ceil(ByVal A) As Integer
    Dim B As Single
    B = Int(A)

    If B <> A Then
      ceil = B + 1
    Else

      ceil = B
    End If
  End Function
  Public Sub SetSource(ByVal bm As Bitmap)
 
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
    Dim bm_bytes As New BitmapBytes(bm)
    bm_bytes.LockBitmap()

    Dim i As Integer = 0
    For y As Integer = 0 To pH
      For x As Integer = 0 To pW
        For b As Integer = 0 To pB
          i += 1
          Sbyte1(b, x, y) = bm_bytes.ImageBytes(i)
        Next
      Next
    Next

    bm_bytes.UnlockBitmap()

    ReDim ResSingle(0 To pB, 0 To pW, 0 To pH)
    ReDim ANGULAR(0 To pB, 0 To pW, 0 To pH, 0 To 15)

  End Sub

  Public Sub SetSourceToMIX(ByVal bm As Bitmap)

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
    Dim bm_bytes As New BitmapBytes(bm)
    bm_bytes.LockBitmap()

    Dim i As Integer = 0
    For y As Integer = 0 To pH
      For x As Integer = 0 To pW
        For b As Integer = 0 To pB
          i += 1
          Sbyte1(b, x, y) = bm_bytes.ImageBytes(i)
        Next
      Next
    Next

    bm_bytes.UnlockBitmap()
  End Sub

  Public Sub MIX(ByVal Amount)
    Dim R As Single
    Dim G As Single
    Dim B As Single
    Dim X As Long
    Dim Y As Long

    Amount = Amount / 100

    For X = 0 To pW
      For Y = 0 To pH
        R = Sbyte1(2, X, Y) - (200 - (ResSingle(2, X, Y))) * Amount
        G = Sbyte1(1, X, Y) - (200 - (ResSingle(1, X, Y))) * Amount
        B = Sbyte1(0, X, Y) - (200 - (ResSingle(0, X, Y))) * Amount
        If R < 0 Then R = 0 Else If R > 255 Then R = 255
        If G < 0 Then G = 0 Else If G > 255 Then G = 255
        If B < 0 Then B = 0 Else If B > 255 Then B = 255


        ResSingle(2, X, Y) = R
        ResSingle(1, X, Y) = G
        ResSingle(0, X, Y) = B
      Next
    Next

  End Sub

  Public Sub GetEffect(ByVal bm As Bitmap)

    Dim bm_bytes As New BitmapBytes(bm)

    bm_bytes.LockBitmap()

    Dim i As Integer = 0
    For y As Integer = 0 To pH
      For x As Integer = 0 To pW
        For b As Integer = 0 To pB
          i += 1
          bm_bytes.ImageBytes(i) = ResSingle(b, x, y)
        Next
      Next
    Next

    bm_bytes.UnlockBitmap()

  End Sub

  'FROM WIKI:
  'function gb=gabor_fn(sigma,theta,lambda,psi,gamma)
  '
  'sigma_x = sigma;
  'sigma_y = sigma/gamma;
  '
  '% Bounding box
  'nstds = 3;
  'xmax = max(abs(nstds*sigma_x*cos(theta)),abs(nstds*sigma_y*sin(theta)));
  'xmax = ceil(max(1,xmax));
  'ymax = max(abs(nstds*sigma_x*sin(theta)),abs(nstds*sigma_y*cos(theta)));
  'ymax = ceil(max(1,ymax));
  'xmin = -xmax; ymin = -ymax;
  '[x,y] = meshgrid(xmin:xmax,ymin:ymax);
  '
  '% Rotation
  'x_theta=x*cos(theta)+y*sin(theta);
  'y_theta=-x*sin(theta)+y*cos(theta);
  '
  'gb=exp(-.5*(x_theta.^2/sigma_x^2+y_theta.^2/sigma_y^2)).*cos(2*pi/lambda*x_theta+psi);
  Private Sub InitGaborFilter(ByVal Sigma, ByVal Lambda, ByVal Psi, ByVal Gamma)
    Dim SigmaX As Single
    Dim SigmaY As Single
    Dim Xmax As Single
    Dim Ymax As Single
    Dim Xmin As Single
    Dim Ymin As Single
    Dim Xtheta As Single
    Dim Ytheta As Single
    Dim nstds As Long
    Dim X As Long
    Dim Y As Long
    Dim GB As Single
    Dim Theta As Single
    Dim CC As Single
    Dim A As Long


    SigmaX = Sigma
    SigmaY = Sigma / Gamma
    'Bounding box
    nstds = 3

    A = 0

    For Theta = 0 To PI * 2 Step (2 * PI / 16)
      Xmax = max(Abs(nstds * SigmaX * Cos(Theta)), Abs(nstds * SigmaY * Sin(Theta)))
      Xmax = ceil(max(1, Xmax))
      Ymax = max(Abs(nstds * SigmaX * Sin(Theta)), Abs(nstds * SigmaY * Cos(Theta)))
      Ymax = ceil(max(1, Ymax))
      Xmin = -Xmax
      Ymin = -Ymax

      ' Rotation
      For X = -nstds To nstds
        For Y = -nstds To nstds
          Xtheta = X * Cos(Theta) + Y * Sin(Theta)
          Ytheta = -X * Sin(Theta) + Y * Cos(Theta)
          '
          GB = Exp(-0.5 * (Xtheta ^ 2 / SigmaX ^ 2 + Ytheta ^ 2 / SigmaY ^ 2)) * Cos(2 * PI / Lambda * Xtheta + Psi)

          GaborFilter(X + 3, Y + 3, A) = GB

          CC = 0
          If GB < 0 Then CC = -GB : GB = 0
        Next
      Next
      A = A + 1
    Next

    For A = 0 To 15
      CC = 0
      For X = -nstds To nstds
        For Y = -nstds To nstds
          CC = CC + GaborFilter(X + 3, Y + 3, A)
        Next
      Next
      CC = CC / 49
      For X = -nstds To nstds
        For Y = -nstds To nstds
          GaborFilter(X + 3, Y + 3, A) = GaborFilter(X + 3, Y + 3, A) - CC
        Next
      Next
    Next

  End Sub

  Public Sub Sketch()
    InitGaborFilter(1, 5, 2, 0.3)
    ApplyGaborFilter()
  End Sub

  Private Sub ApplyGaborFilter()
    Dim A As Long
    Dim XP As Long
    Dim YP As Long
    Dim MaxR As Single
    Dim MaxG As Single
    Dim MaxB As Single

    Dim X As Long
    Dim Y As Long
    Dim ProgX As Long        'For Progress Bar
    Dim ProgXStep As Long        'For Progress Bar

    Dim Xfrom As Long
    Dim Xto As Long
    Dim Yfrom As Long
    Dim Yto As Long

    ProgXStep = Round(2 * pW / 100)
    ProgX = 0

    For X = 3 To pW - 4

      Xfrom = X - 3
      Xto = X + 3

      For Y = 3 To pH - 4

        ResSingle(0, X, Y) = 0
        ResSingle(0, X, Y) = 0
        ResSingle(0, X, Y) = 0

        For A = 0 To 15
          ANGULAR(0, X, Y, A) = 0
          ANGULAR(1, X, Y, A) = 0
          ANGULAR(2, X, Y, A) = 0


          Yfrom = Y - 3
          Yto = Y + 3

          For XP = Xfrom To Xto
            For YP = Yfrom To Yto


              ANGULAR(0, X, Y, A) = ANGULAR(0, X, Y, A) + GaborFilter((X - XP) + 3, (Y - YP) + 3, A) * Sbyte1(0, XP, YP)
              ANGULAR(1, X, Y, A) = ANGULAR(1, X, Y, A) + GaborFilter((X - XP) + 3, (Y - YP) + 3, A) * Sbyte1(1, XP, YP)
              ANGULAR(2, X, Y, A) = ANGULAR(2, X, Y, A) + GaborFilter((X - XP) + 3, (Y - YP) + 3, A) * Sbyte1(2, XP, YP)

            Next YP
          Next XP
        Next A


        MaxR = -9999999
        MaxG = -9999999
        MaxB = -9999999

        For A = 0 To 15
          If ANGULAR(2, X, Y, A) > MaxR Then MaxR = ANGULAR(2, X, Y, A)
          If ANGULAR(1, X, Y, A) > MaxG Then MaxG = ANGULAR(1, X, Y, A)
          If ANGULAR(0, X, Y, A) > MaxB Then MaxB = ANGULAR(0, X, Y, A)
        Next

        If MaxR < 0 Then MaxR = 0 Else If MaxR > 255 Then MaxR = 255
        If MaxG < 0 Then MaxG = 0 Else If MaxG > 255 Then MaxG = 255
        If MaxB < 0 Then MaxB = 0 Else If MaxB > 255 Then MaxB = 255

        MaxR = (MaxR + MaxG + MaxB) / 3
        MaxG = MaxR
        MaxB = MaxR

        ResSingle(2, X, Y) = Not (CByte(MaxR))
        ResSingle(1, X, Y) = Not (CByte(MaxG))
        ResSingle(0, X, Y) = Not (CByte(MaxB))

      Next Y

      If X > ProgX Then
        RaiseEvent PercDONE(X / pW, 0)
        ProgX = ProgX + ProgXStep
      End If
    Next X
    RaiseEvent PercDONE(1, 1)

  End Sub


End Class
