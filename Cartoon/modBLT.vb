Imports System.Drawing

Module modBLT

  'Declarations
  Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal Hdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
  Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal Hdc As Long) As Long
  Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
  Private Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32" (ByVal Hdc As Long) As Long

  Public Declare Function StretchBlt Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
  Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal Hdc As Long, ByVal hStretchMode As Long) As Long

  Public Const STRETCHMODE As Integer = 4 'vbPaletteModeNone    'You can find other modes in the "PaletteModeConstants" section of your Object Browser

  Public Const Pi As Double = 3.14159265358979
  Public Const Pi2 As Double = 6.28318530717959
  Public Const KDegToRad As Double = 0.0174532925199433

End Module
