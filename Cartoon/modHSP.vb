Imports System.Math
Module modHSP

  '©2006 Darel Rex Finley
  '
  'http://www.alienryderflex.com/hsp.html

  Private Const Pr As Double = 0.299    ' 0.241

  Private Const Pg As Double = 0.587    '0.691
  Private Const pB As Double = 0.114    '0.068


  Private Const D12 As Double = 0.5
  Private Const D13 As Double = 0.333333333333333
  Private Const D14 As Double = 0.25
  Private Const D15 As Double = 0.2
  Private Const D16 As Double = 0.166666666666667
  Private Const D23 As Double = 0.666666666666667

  Public Sub RGBtoHSP(ByVal R, ByVal G, ByVal B, ByRef rH, ByRef rS, ByRef rP)
    Dim P As Double
    Dim H As Double
    Dim S As Double


    P = Sqrt(R * R * Pr + G * G * Pg + B * B * pB)


    '  //  Calculate the Hue and Saturation.  (This part works
    '  //  the same way as in the HSV/B and HSL systems???.)
    If (R = G And R = B) Then H = 0 : S = 0 : GoTo RET

    If (R >= G And R >= B) Then
      '   //  R is largest
      If (B >= G) Then
        H = 1 - D16 * (B - G) / (R - G) : S = 1 - G / R
      Else
        H = D16 * (G - B) / (R - B) : S = 1 - B / R
      End If
    Else
      If (G >= R And G >= B) Then
        '//  G is largest
        If (R >= B) Then
          H = D13 - D16 * (R - B) / (G - B) : S = 1 - B / G
        Else
          H = D13 + D16 * (B - R) / (G - R) : S = 1 - R / G
        End If

      Else
        '//  B is largest
        If (G >= R) Then
          H = D23 - D16 * (G - R) / (B - R) : S = 1 - R / B
        Else
          H = D23 + D16 * (R - G) / (B - G) : S = 1 - G / B
        End If
      End If
    End If

RET:

    'Stop

    rP = P
    rH = H * 255
    rS = S * 255

  End Sub


  Public Sub HSPtoRGB(ByVal H, ByVal S, ByVal P, ByRef R, ByRef G, ByRef B)

    Dim Part As Double
    Dim minOverMax As Double


    H = H / 255
    S = S / 255

    minOverMax = 1 - S
    'Stop

    If (minOverMax > 0) Then

      If (H < 1 / 6) Then       '   //  R>G>B
        H = 6 * (H - 0 / 6) : Part = 1 + H * (1 / minOverMax - 1)
        B = P / Sqrt(Pr / minOverMax / minOverMax + Pg * Part * Part + pB)
        R = (B) / minOverMax : G = (B) + H * (R - B)
      Else

        If (H < 2 / 6) Then   '   //  G>R>B
          H = 6 * (-H + 2 / 6) : Part = 1 + H * (1 / minOverMax - 1)
          B = P / Sqrt(Pg / minOverMax / minOverMax + Pr * Part * Part + pB)
          G = (B) / minOverMax : R = (B) + H * (G - B)
        Else
          If (H < 3 / 6) Then    '   //  G>B>R
            H = 6 * (H - 2 / 6) : Part = 1 + H * (1 / minOverMax - 1)
            R = P / Sqrt(Pg / minOverMax / minOverMax + pB * Part * Part + Pr)
            G = (R) / minOverMax : B = (R) + H * (G - R)
          Else
            If (H < 4 / 6) Then    '   //  B>G>R
              H = 6 * (-H + 4 / 6) : Part = 1 + H * (1 / minOverMax - 1)
              R = P / Sqrt(pB / minOverMax / minOverMax + Pg * Part * Part + Pr)
              B = (R) / minOverMax : G = (R) + H * (B - R)
            Else


              If (H < 5 / 6) Then    '  //  B>R>G
                H = 6 * (H - 4 / 6) : Part = 1 + H * (1 / minOverMax - 1)
                G = P / Sqrt(pB / minOverMax / minOverMax + Pr * Part * Part + Pg)
                B = (G) / minOverMax : R = (G) + H * (B - G)
              Else      '   //  R>B>G
                H = 6 * (-H + 6 / 6) : Part = 1 + H * (1 / minOverMax - 1)
                G = P / Sqrt(Pr / minOverMax / minOverMax + pB * Part * Part + Pg)
                R = (G) / minOverMax : B = (G) + H * (R - G)
              End If
            End If
          End If
        End If
      End If

    Else

      If (H < 1 / 6) Then       '   //  R>G>B
        H = 6.0# * (H - 0.0# / 6.0#) : R = Sqrt(P * P / (Pr + Pg * H * H)) : G = (R) * H : B = 0
      Else
        If (H < 2 / 6) Then   '   //  G>R>B
          H = 6.0# * (-H + 2.0# / 6.0#) : G = Sqrt(P * P / (Pg + Pr * H * H)) : R = (G) * H : B = 0
        Else
          If (H < 3 / 6) Then    '   //  G>B>R
            H = 6.0# * (H - 2.0# / 6.0#) : G = Sqrt(P * P / (Pg + pB * H * H)) : B = (G) * H : R = 0
          Else
            If (H < 4 / 6) Then    '   //  B>G>R
              H = 6.0# * (-H + 4.0# / 6.0#) : B = Sqrt(P * P / (pB + Pg * H * H)) : G = (B) * H : R = 0
            Else
              If (H < 5 / 6) Then    '  //  B>R>G
                H = 6.0# * (H - 4.0# / 6.0#) : B = Sqrt(P * P / (pB + Pr * H * H)) : R = (B) * H : G = 0
              Else      '   //  R>B>G
                H = 6.0# * (-H + 6.0# / 6.0#) : R = Sqrt(P * P / (Pr + pB * H * H)) : B = (R) * H : G = 0
              End If
            End If
          End If
        End If
      End If
    End If

    If R > 255 Then R = 255
    If G > 255 Then G = 255
    If B > 255 Then B = 255

  End Sub


  Public Function HUEDifference(ByVal H1, ByVal H2) As Single
    Dim D As Single
    D = H2 - (-H1)
    If D < -127.5 Then D = D + 255
    If D > 127.5 Then D = D - 255
    HUEDifference = Abs(D)

  End Function

End Module
