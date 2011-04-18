Imports System.Drawing
Imports System.Runtime.InteropServices

Module modBLT

  'Declarations
  Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal Hdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
  Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal Hdc As Long) As Long
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

  ''' <summary>
  '''    Performs a bit-block transfer of the color data corresponding to a
  '''    rectangle of pixels from the specified source device context into
  '''    a destination device context.
  ''' </summary>
  ''' <param name="hdc">Handle to the destination device context.</param>
  ''' <param name="nXDest">The leftmost x-coordinate of the destination rectangle (in pixels).</param>
  ''' <param name="nYDest">The topmost y-coordinate of the destination rectangle (in pixels).</param>
  ''' <param name="nWidth">The width of the source and destination rectangles (in pixels).</param>
  ''' <param name="nHeight">The height of the source and the destination rectangles (in pixels).</param>
  ''' <param name="hdcSrc">Handle to the source device context.</param>
  ''' <param name="nXSrc">The leftmost x-coordinate of the source rectangle (in pixels).</param>
  ''' <param name="nYSrc">The topmost y-coordinate of the source rectangle (in pixels).</param>
  ''' <param name="dwRop">A raster-operation code.</param>
  ''' <returns>
  '''    <c>true</c> if the operation succeeded, <c>false</c> otherwise.
  ''' </returns>
  <DllImport("gdi32.dll")> _
  Public Function BitBlt(ByVal hdc As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As TernaryRasterOperations) As Boolean
  End Function


  ''' <summary>
  '''     Specifies a raster-operation code. These codes define how the color data for the
  '''     source rectangle is to be combined with the color data for the destination
  '''     rectangle to achieve the final color.
  ''' </summary>
  Enum TernaryRasterOperations As UInteger
    ''' <summary>dest = source</summary>
    SRCCOPY = &HCC0020
    ''' <summary>dest = source OR dest</summary>
    SRCPAINT = &HEE0086
    ''' <summary>dest = source AND dest</summary>
    SRCAND = &H8800C6
    ''' <summary>dest = source XOR dest</summary>
    SRCINVERT = &H660046
    ''' <summary>dest = source AND (NOT dest)</summary>
    SRCERASE = &H440328
    ''' <summary>dest = (NOT source)</summary>
    NOTSRCCOPY = &H330008
    ''' <summary>dest = (NOT src) AND (NOT dest)</summary>
    NOTSRCERASE = &H1100A6
    ''' <summary>dest = (source AND pattern)</summary>
    MERGECOPY = &HC000CA
    ''' <summary>dest = (NOT source) OR dest</summary>
    MERGEPAINT = &HBB0226
    ''' <summary>dest = pattern</summary>
    PATCOPY = &HF00021
    ''' <summary>dest = DPSnoo</summary>
    PATPAINT = &HFB0A09
    ''' <summary>dest = pattern XOR dest</summary>
    PATINVERT = &H5A0049
    ''' <summary>dest = (NOT dest)</summary>
    DSTINVERT = &H550009
    ''' <summary>dest = BLACK</summary>
    BLACKNESS = &H42
    ''' <summary>dest = WHITE</summary>
    WHITENESS = &HFF0062
    ''' <summary>
    ''' Capture window as seen on screen.  This includes layered windows
    ''' such as WPF windows with AllowsTransparency="true"
    ''' </summary>
    CAPTUREBLT = &H40000000
  End Enum



End Module
