Imports System.Drawing
Imports System.Windows.Forms

Public Class CartoonizerForm

  Private WithEvents FX As New Effects
  Private WithEvents Sketcher As New Sketch
  Private mBitmap As Bitmap
  Private mBitmapFileName As String

  Private SigmaIntensity As Single
  Private SigmaSpatial As Single

  Private m_PanStartPoint As New Point

#Region "Pre Effect Panel"

  Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPreEffect.SelectedIndexChanged
    UpdatePreEffectGroupBoxSize()
    pnExposure.Visible = (cbPreEffect.SelectedIndex = 2)
    pnBCS.Visible = (cbPreEffect.SelectedIndex = 3)
    UpdatePreEffectLabels()
  End Sub

  Private Sub tbExposure_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbExposure.Scroll
    UpdatePreEffectLabels()
  End Sub

  Private Sub tbBrightness_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbBrightness.Scroll
    UpdatePreEffectLabels()
  End Sub

  Private Sub tbContrast_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbContrast.Scroll
    UpdatePreEffectLabels()
  End Sub

  Private Sub tbSaturation_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSaturation.Scroll
    UpdatePreEffectLabels()
  End Sub

  Private Sub UpdatePreEffectLabels()
    lbExposure.Text = String.Format("Exposure: {0}", tbExposure.Value)
    lbExposure.Refresh()
    lbBrightness.Text = String.Format("Brightness: {0} %", CInt(200 * (tbBrightness.Value / tbBrightness.Maximum)))
    lbBrightness.Refresh()
    lbContrast.Text = String.Format("Contrast: {0} %", tbContrast.Value)
    lbContrast.Refresh()
    lbSaturation.Text = String.Format("Saturation: {0} %", CInt(200 * (tbSaturation.Value / tbSaturation.Maximum)))
    lbSaturation.Refresh()
  End Sub

  Private Sub UpdatePreEffectGroupBoxSize()
    Dim width As Integer = 209
    Select Case cbPreEffect.SelectedIndex
      Case 0, 1
        gbPreEffect.Size = New Size(width, 80)
      Case 2
        gbPreEffect.Size = New Size(width, 135)
      Case 3
        gbPreEffect.Size = New Size(width, 218)
    End Select
  End Sub


#End Region

#Region "Default Values"

  Private Sub SetDefaults()
    cbPreEffect.SelectedIndex = 0 'None
    cbColorMode.SelectedIndex = 1 'LAB
    cbIntensityMode.SelectedIndex = 1 'Gaussian2
    cbContourMethod.SelectedIndex = 0 'Sobel
    SigmaIntensityChange()
    SigmaSpatialChange()
    UpdatePreEffectLabels()
    UpdateBilateralLabels()
    UpdateContourLabels()
    UpdateLumSegLabels()
  End Sub

#End Region

#Region "Form Events"

  Private Sub CartoonizerForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    SetDefaults()
  End Sub

#End Region

#Region "Bilateral Panel"

  Private Sub tbIterations_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbIterations.Scroll
    UpdateBilateralLabels()
  End Sub

  Private Sub tbSpatialSigma_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSpatialSigma.Scroll
    UpdateBilateralLabels()
    SigmaSpatialChange()
  End Sub

  Private Sub tbSpatialIntensity_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSpatialIntensity.Scroll
    UpdateBilateralLabels()
    SigmaIntensityChange()
  End Sub

  Private Sub tbRadius_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbRadius.Scroll
    UpdateBilateralLabels()
    SigmaSpatialChange()
  End Sub

  Private Sub cbIntensityMode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbIntensityMode.SelectedIndexChanged
    UpdateBilateralLabels()
    SigmaIntensityChange()
  End Sub

  Private Sub SigmaIntensityChange()
    Dim myBM As New Bitmap(pbIntensity.Width, pbIntensity.Height)
    FX.zPreview_Intensity(myBM, tbSpatialIntensity.Value / 10000, cbIntensityMode.SelectedIndex)
    pbIntensity.Image = myBM
    pbIntensity.Refresh()
  End Sub

  Private Sub SigmaSpatialChange()
    Dim myBM As New Bitmap(pbSpatial.Width, pbSpatial.Height)
    FX.zPreview_Spatial(myBM, tbRadius.Value, tbSpatialSigma.Value / 10)
    pbSpatial.Image = myBM
    pbSpatial.Refresh()
  End Sub

  Private Sub UpdateBilateralLabels()
    lbRadius.Text = String.Format("Radius (HalfLate): {0}", tbRadius.Value)
    lbRadius.Refresh()
    lbSpatialIntensity.Text = String.Format("Spatial Intensity: {0}", tbSpatialIntensity.Value / 1000)
    lbSpatialIntensity.Refresh()
    lbSpatialSigma.Text = String.Format("Spatial Sigma: {0}", tbSpatialSigma.Value / 10)
    lbSpatialSigma.Refresh()
    lbIterations.Text = String.Format("Iterations: {0}", tbIterations.Value)
    lbIterations.Refresh()
  End Sub

#End Region

#Region "Contour Panel"

  Private Sub cbContourMethod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbContourMethod.SelectedIndexChanged
    pnThreshold.Visible = (cbContourMethod.SelectedIndex = 1)
    pnLumHue.Visible = (cbContourMethod.SelectedIndex = 0)
    UpdateContourLabels()
  End Sub

  Private Sub tbContourAmount_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbContourAmount.Scroll
    UpdateContourLabels()
  End Sub

  Private Sub tbLumHue_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbLumHue.Scroll
    UpdateContourLabels()
  End Sub

  Private Sub tbContourThreshold_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbContourThreshold.Scroll
    UpdateContourLabels()
  End Sub

  Private Sub UpdateContourLabels()
    lbContourAmount.Text = String.Format("Contour Amount: {0}", tbContourAmount.Value)
    lbContourAmount.Refresh()
    lbLumHue.Text = String.Format("Based On {0}% L {1}% AB :", tbLumHue.Value, 100 - tbLumHue.Value)
    lbLumHue.Refresh()
    lbContourThreshold.Text = String.Format("Threshold: {0}", tbContourThreshold.Value / 1000)
    lbContourThreshold.Refresh()
  End Sub

#End Region

#Region "Luminance Segmentation Panel"

  Private Sub tbSegments_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbSegments.Scroll
    UpdateLumSegLabels()
  End Sub

  Private Sub tbPresence_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPresence.Scroll
    UpdateLumSegLabels()
  End Sub

  Private Sub UpdateLumSegLabels()
    lbSegments.Text = String.Format("Segments: {0}", tbSegments.Value)
    lbSegments.Refresh()
    lbPresence.Text = String.Format("Presence: {0}%", tbPresence.Value)
    lbPresence.Refresh()
  End Sub

#End Region

#Region "Progress Bar Events"

  Private Sub FX_PercDONE(ByVal Filter As String, ByVal PercValue As Single, ByVal CurrIteration As Long) Handles FX.PercDONE
    pbCartoonize.Value = CInt(PercValue * 100)
    pbCartoonize.Refresh()
    Application.DoEvents()
    If PercValue > 0 Then
      pbCartoonize.CreateGraphics().DrawString(String.Format("{0} {1} {2}%", Filter, CurrIteration, pbCartoonize.Value.ToString()), New Font("Arial", CSng(8.25), FontStyle.Regular), Brushes.Black, New PointF(10, pbCartoonize.Height / 2 - 7))
    End If
  End Sub

  Private Sub Sketcher_PercDONE(ByVal PercValue As Single, ByVal CurrIteration As Long) Handles Sketcher.PercDONE
    pbCartoonize.Value = CInt(PercValue * 100)
    pbCartoonize.Refresh()
    Application.DoEvents()
    If PercValue > 0 Then
      pbCartoonize.CreateGraphics().DrawString(String.Format("Sketching {0} {1}%", CurrIteration, pbCartoonize.Value.ToString()), New Font("Arial", CSng(8.25), FontStyle.Regular), Brushes.Black, New PointF(10, pbCartoonize.Height / 2 - 7))
    End If
  End Sub

#End Region

#Region "Load/Save Bitmap"

  Private Sub LoadBitmap()
    Dim bm As New Bitmap(mBitmapFileName)
    mBitmap = New Bitmap(bm.Width, bm.Height, Imaging.PixelFormat.Format24bppRgb)
    mBitmap.SetResolution(bm.HorizontalResolution, bm.VerticalResolution)
    Dim g As Graphics = Graphics.FromImage(mBitmap)
    g.DrawImage(bm, 0, 0)
    g.Dispose()
    bm.Dispose()
  End Sub

  Private Sub btBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btBrowse.Click
    If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
      mBitmapFileName = OpenFileDialog1.FileName
      tbImage.Text = mBitmapFileName
      tbImage.Refresh()
      LoadBitmap()
      PictureBox1.Image = mBitmap
      PictureBox1.Refresh()
    End If
  End Sub

  Private Sub btSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btSave.Click
    If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
      PictureBox1.Image.Save(SaveFileDialog1.FileName)
    End If
  End Sub

#End Region

#Region "Cartoonize"

  Private Sub btCartoonize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCartoonize.Click
    PictureBox1.Image = mBitmap
    PictureBox1.Refresh()
    Cartoonize()
  End Sub

  Private Sub DoPreEffect()
    Select Case cbPreEffect.SelectedIndex
      Case 0
      Case 1
        FX.MagneKleverHistogramEQU(0.3)
      Case 2
        FX.MagneKleverExposure(tbExposure.Value)
      Case 3
        FX.MagneKleverBCS(tbBrightness.Value, tbContrast.Value, tbSaturation.Value)
    End Select
  End Sub

  Private Sub DoContour()
    If tbContourAmount.Value > 0 Then
      Select Case cbContourMethod.SelectedIndex
        Case 0
          FX.zEFF_Contour(tbContourAmount.Value, tbLumHue.Value / 100)
        Case 1
          FX.zEFF_ContourbyDOG(tbContourAmount.Value, tbContourThreshold.Value / 1000)
      End Select
    End If
  End Sub

  Private Sub ApplyContour()
    If tbContourAmount.Value > 0 Then
      FX.zEFF_Contour_Apply()
    End If
  End Sub

  Private Sub DoBilateral()
    FX.zEFF_BilateralFilter(tbRadius.Value, tbSpatialIntensity.Value / 10000, tbSpatialSigma.Value / 100, tbIterations.Value, cbIntensityMode.SelectedIndex, cbColorMode.SelectedIndex = 0)
  End Sub

  Private Sub DoQuantizeLuminance()
    FX.zEFF_QuantizeLuminance(tbSegments.Value, tbPresence.Value / 100, tbRadius.Value, False)
  End Sub

  'Private Sub InitDomains()
  '  FX.zInit_IntensityDomain(tbSpatialIntensity.Value / 10000, cbIntensityMode.SelectedIndex)
  '  FX.zInit_SpatialDomain(tbSpatialSigma.Value / 100)
  'End Sub

  Public Sub Cartoonize()
    Dim bm As Bitmap = mBitmap.Clone()
    FX.zSet_Source(bm)
    DoPreEffect()
    'Uncomment this if you want to call this without the UI
    'InitDomains
    DoBilateral()
    DoContour()
    DoQuantizeLuminance()
    ApplyContour()

    FX.zGet_Effect(bm)
    PictureBox1.Image = bm
    PictureBox1.Refresh()
    FX_PercDONE("", 0, 0)
  End Sub

#End Region

#Region "Sketch"

  Public Sub Sketch()
    Dim bm As Bitmap = mBitmap.Clone()
    Sketcher.SetSource(bm)
    Sketcher.Sketch()
    Sketcher.GetEffect(bm)
    PictureBox1.Image = bm
    PictureBox1.Refresh()
    FX_PercDONE("", 0, 0)
  End Sub

#End Region

#Region "Image Panning"

  Private Sub CartoonizerForm_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDown
    Cursor = Cursors.Hand
    m_PanStartPoint = New Point(e.X, e.Y)
  End Sub

  Private Sub PictureBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseUp
    Cursor = Cursors.Default
  End Sub

  Private Sub CartoonizerForm_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
    pbCartoonize.Focus()
    If e.Button = Windows.Forms.MouseButtons.Left Then
      Dim DeltaX As Integer = (m_PanStartPoint.X - e.X)
      Dim DeltaY As Integer = (m_PanStartPoint.Y - e.Y)
      pnPic.AutoScrollPosition = New Drawing.Point((DeltaX - pnPic.AutoScrollPosition.X), (DeltaY - pnPic.AutoScrollPosition.Y))
    End If
  End Sub

  Private Sub CartoonizerForm_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel
    Dim newWidth As Integer
    Dim newHeight As Integer

    If e.Delta > 0 Then
      newWidth = PictureBox1.Width * 1.25
      newHeight = PictureBox1.Height * 1.25
    Else
      newWidth = PictureBox1.Width / 1.25
      newHeight = PictureBox1.Height / 1.25
    End If

    If newWidth > pnPic.ClientSize.Width Or newHeight > pnPic.ClientSize.Height Then
      PictureBox1.Dock = DockStyle.None
      PictureBox1.Location = New Point(0, 0)
      PictureBox1.Width = newWidth
      PictureBox1.Height = newHeight
    Else
      PictureBox1.Dock = DockStyle.Fill
    End If

    PictureBox1.Refresh()
  End Sub

#End Region

End Class