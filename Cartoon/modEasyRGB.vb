Module modEasyRGB

  Private Const K1d3 As Single = 0.333333333333333    '1/3
  Private Const K16d116 As Single = 0.137931034482759    '16/116
  Private Const K1d2p4 As Single = 0.416666666666667    '1/2.4


  Private Const LumMUL = 2.55
  Private Const abMUL = 1.0625     '255/240
  Private Const LumDIV = 1 / 2.55
  Private Const abDIV = 240 / 255


  Private Const d7787 As Single = 1 / 7.787
  Private Const d1292 As Single = 1 / 12.92
  Private Const d1055 As Single = 1 / 1.055
  Private Const d255 As Single = 1 / 255

  Private Const OneD500 As Single = 1 / 500
  Private Const OneD116 As Single = 1 / 116
  Private Const OneD200 As Single = 1 / 200



  Public Declare Function GetPixel Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
  Public Declare Function SetPixel Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

  Public Function CieLAB_RGB(ByVal cL As Single, ByVal cA As Single, ByVal cB As Single, _
                             ByRef R As Single, ByRef G As Single, ByRef B As Single) As Boolean

    Dim X As Single
    Dim Y As Single
    Dim Z As Single
    Dim X3 As Single
    Dim Y3 As Single
    Dim Z3 As Single

    CieLAB_RGB = True

    cL = cL * LumDIV
    cA = cA * abDIV
    cB = cB * abDIV

    cA = cA - 120
    cB = cB - 120

    Y = (cL + 16) * OneD116
    X = cA * OneD500 + Y
    Z = Y - cB * OneD200

    X3 = X * X * X
    Y3 = Y * Y * Y
    Z3 = Z * Z * Z

    If (X3 > 0.008856) Then
      X = X3
    Else
      X = (X - K16d116) * d7787
    End If

    If (Y3 > 0.008856) Then
      Y = Y3
    Else
      Y = (Y - K16d116) * d7787
    End If

    If (Z3 > 0.008856) Then
      Z = Z3
    Else
      Z = (Z - K16d116) * d7787
    End If

    X = X * 0.95047               'Observer= 2°, Illuminant= D65
    'Y = Y * 1
    Z = Z * 1.08883


    R = (X * 3.2406 + Y * -1.5372 + Z * -0.4986)
    G = (X * -0.9689 + Y * 1.8758 + Z * 0.0415)
    B = (X * 0.0557 + Y * -0.204 + Z * 1.057)

    If (R > 0.0031308) Then
      R = (1.055 * (R ^ K1d2p4) - 0.055)
    Else
      R = 12.92 * R
    End If

    If (G > 0.0031308) Then
      G = (1.055 * (G ^ K1d2p4) - 0.055)
    Else
      G = 12.92 * G
    End If

    If (B > 0.0031308) Then
      B = (1.055 * (B ^ K1d2p4) - 0.055)
    Else
      B = 12.92 * B
    End If

    R = R * 255
    G = G * 255
    B = B * 255

    If R < 0 Then
      R = 0
      CieLAB_RGB = False
    ElseIf R > 255 Then
      R = 255
      CieLAB_RGB = False
    End If

    If G < 0 Then
      G = 0
      CieLAB_RGB = False
    ElseIf G > 255 Then
      G = 255
      CieLAB_RGB = False
    End If

    If B < 0 Then
      B = 0
      CieLAB_RGB = False
    ElseIf B > 255 Then
      B = 255
      CieLAB_RGB = False
    End If

  End Function

  Public Sub RGB_CieLAB(ByVal R As Single, ByVal G As Single, ByVal B As Single, _
                             ByRef cL As Single, ByRef cA As Single, ByRef cB As Single)

    Dim X As Single
    Dim Y As Single
    Dim Z As Single

    R = (R * d255)
    G = (G * d255)
    B = (B * d255)

    If (R > 0.04045) Then
      R = ((R + 0.055) * d1055) ^ 2.4
    Else
      R = R * d1292
    End If

    If (G > 0.04045) Then
      G = ((G + 0.055) * d1055) ^ 2.4
    Else
      G = G * d1292
    End If

    If (B > 0.04045) Then
      B = ((B + 0.055) * d1055) ^ 2.4
    Else
      B = B * d1292
    End If

    'Observer. = 2°, Illuminant = D65
    X = R * 0.4124 + G * 0.3576 + B * 0.1805
    Y = R * 0.2126 + G * 0.7152 + B * 0.0722
    Z = R * 0.0193 + G * 0.1192 + B * 0.9505

    X = X / 0.95047               'Observer= 2°, Illuminant= D65
    'Y = Y / 1
    Z = Z / 1.08883

    If (X > 0.008856) Then
      X = X ^ K1d3
    Else
      X = (7.787 * X) + (K16d116)
    End If

    If (Y > 0.008856) Then
      Y = Y ^ K1d3
    Else
      Y = (7.787 * Y) + (K16d116)
    End If

    If (Z > 0.008856) Then
      Z = Z ^ K1d3
    Else
      Z = (7.787 * Z) + (K16d116)
    End If

    cL = 116 * Y - 16
    cA = 500 * (X - Y)
    cB = 200 * (Y - Z)

    cA = cA + 120
    cB = cB + 120

    cL = cL * LumMUL
    cA = cA * abMUL
    cB = cB * abMUL

  End Sub

End Module
