<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CartoonizerForm
  Inherits System.Windows.Forms.Form

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()> _
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    Try
      If disposing AndAlso components IsNot Nothing Then
        components.Dispose()
      End If
    Finally
      MyBase.Dispose(disposing)
    End Try
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.gbPreEffect = New System.Windows.Forms.GroupBox
    Me.lbPreEffect = New System.Windows.Forms.Label
    Me.pnExposure = New System.Windows.Forms.Panel
    Me.lbExposure = New System.Windows.Forms.Label
    Me.tbExposure = New System.Windows.Forms.TrackBar
    Me.pnBCS = New System.Windows.Forms.Panel
    Me.lbBrightness = New System.Windows.Forms.Label
    Me.lbContrast = New System.Windows.Forms.Label
    Me.lbSaturation = New System.Windows.Forms.Label
    Me.tbSaturation = New System.Windows.Forms.TrackBar
    Me.tbBrightness = New System.Windows.Forms.TrackBar
    Me.tbContrast = New System.Windows.Forms.TrackBar
    Me.cbPreEffect = New System.Windows.Forms.ComboBox
    Me.gbBilateral = New System.Windows.Forms.GroupBox
    Me.lbIntensityMode = New System.Windows.Forms.Label
    Me.cbIntensityMode = New System.Windows.Forms.ComboBox
    Me.lbSpatialIntensity = New System.Windows.Forms.Label
    Me.lbPicSpatial = New System.Windows.Forms.Label
    Me.lbPicIntensity = New System.Windows.Forms.Label
    Me.pbSpatial = New System.Windows.Forms.PictureBox
    Me.pbIntensity = New System.Windows.Forms.PictureBox
    Me.lbIterations = New System.Windows.Forms.Label
    Me.tbIterations = New System.Windows.Forms.TrackBar
    Me.lbSpatialSigma = New System.Windows.Forms.Label
    Me.tbSpatialSigma = New System.Windows.Forms.TrackBar
    Me.lbRadius = New System.Windows.Forms.Label
    Me.tbRadius = New System.Windows.Forms.TrackBar
    Me.lbColorMode = New System.Windows.Forms.Label
    Me.cbColorMode = New System.Windows.Forms.ComboBox
    Me.tbSpatialIntensity = New System.Windows.Forms.TrackBar
    Me.gbContour = New System.Windows.Forms.GroupBox
    Me.pnThreshold = New System.Windows.Forms.Panel
    Me.lbContourThreshold = New System.Windows.Forms.Label
    Me.tbContourThreshold = New System.Windows.Forms.TrackBar
    Me.pnLumHue = New System.Windows.Forms.Panel
    Me.lbLumHue = New System.Windows.Forms.Label
    Me.tbLumHue = New System.Windows.Forms.TrackBar
    Me.pnContourAmount = New System.Windows.Forms.Panel
    Me.lbContourAmount = New System.Windows.Forms.Label
    Me.tbContourAmount = New System.Windows.Forms.TrackBar
    Me.lbContourMethod = New System.Windows.Forms.Label
    Me.cbContourMethod = New System.Windows.Forms.ComboBox
    Me.flpEffect = New System.Windows.Forms.FlowLayoutPanel
    Me.gbLumSeg = New System.Windows.Forms.GroupBox
    Me.lbPresence = New System.Windows.Forms.Label
    Me.tbPresence = New System.Windows.Forms.TrackBar
    Me.lbSegments = New System.Windows.Forms.Label
    Me.tbSegments = New System.Windows.Forms.TrackBar
    Me.pnPicCtrl = New System.Windows.Forms.Panel
    Me.pnPic = New System.Windows.Forms.Panel
    Me.PictureBox1 = New System.Windows.Forms.PictureBox
    Me.pnCtrl = New System.Windows.Forms.Panel
    Me.tbImage = New System.Windows.Forms.TextBox
    Me.btBrowse = New System.Windows.Forms.Button
    Me.btCartoonize = New System.Windows.Forms.Button
    Me.pbCartoonize = New System.Windows.Forms.ProgressBar
    Me.btSave = New System.Windows.Forms.Button
    Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
    Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
    Me.gbPreEffect.SuspendLayout()
    Me.pnExposure.SuspendLayout()
    CType(Me.tbExposure, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.pnBCS.SuspendLayout()
    CType(Me.tbSaturation, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.tbBrightness, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.tbContrast, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.gbBilateral.SuspendLayout()
    CType(Me.pbSpatial, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.pbIntensity, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.tbIterations, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.tbSpatialSigma, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.tbRadius, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.tbSpatialIntensity, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.gbContour.SuspendLayout()
    Me.pnThreshold.SuspendLayout()
    CType(Me.tbContourThreshold, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.pnLumHue.SuspendLayout()
    CType(Me.tbLumHue, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.pnContourAmount.SuspendLayout()
    CType(Me.tbContourAmount, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.flpEffect.SuspendLayout()
    Me.gbLumSeg.SuspendLayout()
    CType(Me.tbPresence, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.tbSegments, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.pnPicCtrl.SuspendLayout()
    Me.pnPic.SuspendLayout()
    CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.pnCtrl.SuspendLayout()
    Me.SuspendLayout()
    '
    'gbPreEffect
    '
    Me.gbPreEffect.Controls.Add(Me.pnExposure)
    Me.gbPreEffect.Controls.Add(Me.pnBCS)
    Me.gbPreEffect.Controls.Add(Me.cbPreEffect)
    Me.gbPreEffect.Controls.Add(Me.lbPreEffect)
    Me.gbPreEffect.Location = New System.Drawing.Point(3, 3)
    Me.gbPreEffect.Name = "gbPreEffect"
    Me.gbPreEffect.Size = New System.Drawing.Size(209, 218)
    Me.gbPreEffect.TabIndex = 0
    Me.gbPreEffect.TabStop = False
    Me.gbPreEffect.Text = "Pre Effect"
    '
    'lbPreEffect
    '
    Me.lbPreEffect.AutoSize = True
    Me.lbPreEffect.Location = New System.Drawing.Point(23, 28)
    Me.lbPreEffect.Name = "lbPreEffect"
    Me.lbPreEffect.Size = New System.Drawing.Size(35, 13)
    Me.lbPreEffect.TabIndex = 23
    Me.lbPreEffect.Text = "Effect"
    '
    'pnExposure
    '
    Me.pnExposure.Controls.Add(Me.lbExposure)
    Me.pnExposure.Controls.Add(Me.tbExposure)
    Me.pnExposure.Location = New System.Drawing.Point(20, 78)
    Me.pnExposure.Name = "pnExposure"
    Me.pnExposure.Size = New System.Drawing.Size(174, 54)
    Me.pnExposure.TabIndex = 1
    Me.pnExposure.Visible = False
    '
    'lbExposure
    '
    Me.lbExposure.AutoSize = True
    Me.lbExposure.Location = New System.Drawing.Point(3, 0)
    Me.lbExposure.Name = "lbExposure"
    Me.lbExposure.Size = New System.Drawing.Size(54, 13)
    Me.lbExposure.TabIndex = 2
    Me.lbExposure.Text = "Exposure:"
    '
    'tbExposure
    '
    Me.tbExposure.Location = New System.Drawing.Point(-3, 16)
    Me.tbExposure.Maximum = 256
    Me.tbExposure.Minimum = -127
    Me.tbExposure.Name = "tbExposure"
    Me.tbExposure.Size = New System.Drawing.Size(174, 45)
    Me.tbExposure.TabIndex = 1
    Me.tbExposure.TickStyle = System.Windows.Forms.TickStyle.None
    '
    'pnBCS
    '
    Me.pnBCS.Controls.Add(Me.lbBrightness)
    Me.pnBCS.Controls.Add(Me.lbContrast)
    Me.pnBCS.Controls.Add(Me.lbSaturation)
    Me.pnBCS.Controls.Add(Me.tbSaturation)
    Me.pnBCS.Controls.Add(Me.tbBrightness)
    Me.pnBCS.Controls.Add(Me.tbContrast)
    Me.pnBCS.Location = New System.Drawing.Point(20, 78)
    Me.pnBCS.Name = "pnBCS"
    Me.pnBCS.Size = New System.Drawing.Size(174, 131)
    Me.pnBCS.TabIndex = 2
    Me.pnBCS.Visible = False
    '
    'lbBrightness
    '
    Me.lbBrightness.AutoSize = True
    Me.lbBrightness.Location = New System.Drawing.Point(3, 0)
    Me.lbBrightness.Name = "lbBrightness"
    Me.lbBrightness.Size = New System.Drawing.Size(59, 13)
    Me.lbBrightness.TabIndex = 8
    Me.lbBrightness.Text = "Brightness:"
    '
    'lbContrast
    '
    Me.lbContrast.AutoSize = True
    Me.lbContrast.Location = New System.Drawing.Point(3, 42)
    Me.lbContrast.Name = "lbContrast"
    Me.lbContrast.Size = New System.Drawing.Size(49, 13)
    Me.lbContrast.TabIndex = 6
    Me.lbContrast.Text = "Contrast:"
    '
    'lbSaturation
    '
    Me.lbSaturation.AutoSize = True
    Me.lbSaturation.Location = New System.Drawing.Point(3, 88)
    Me.lbSaturation.Name = "lbSaturation"
    Me.lbSaturation.Size = New System.Drawing.Size(58, 13)
    Me.lbSaturation.TabIndex = 4
    Me.lbSaturation.Text = "Saturation:"
    '
    'tbSaturation
    '
    Me.tbSaturation.Location = New System.Drawing.Point(-3, 104)
    Me.tbSaturation.Maximum = 512
    Me.tbSaturation.Name = "tbSaturation"
    Me.tbSaturation.Size = New System.Drawing.Size(174, 45)
    Me.tbSaturation.TabIndex = 3
    Me.tbSaturation.TickStyle = System.Windows.Forms.TickStyle.None
    Me.tbSaturation.Value = 256
    '
    'tbBrightness
    '
    Me.tbBrightness.Location = New System.Drawing.Point(-3, 16)
    Me.tbBrightness.Maximum = 512
    Me.tbBrightness.Name = "tbBrightness"
    Me.tbBrightness.Size = New System.Drawing.Size(174, 45)
    Me.tbBrightness.TabIndex = 7
    Me.tbBrightness.TickStyle = System.Windows.Forms.TickStyle.None
    Me.tbBrightness.Value = 256
    '
    'tbContrast
    '
    Me.tbContrast.Location = New System.Drawing.Point(-3, 59)
    Me.tbContrast.Maximum = 100
    Me.tbContrast.Minimum = -100
    Me.tbContrast.Name = "tbContrast"
    Me.tbContrast.Size = New System.Drawing.Size(174, 45)
    Me.tbContrast.TabIndex = 5
    Me.tbContrast.TickStyle = System.Windows.Forms.TickStyle.None
    '
    'cbPreEffect
    '
    Me.cbPreEffect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cbPreEffect.FormattingEnabled = True
    Me.cbPreEffect.Items.AddRange(New Object() {"None", "Auto Equalize", "Exposure", "BCS Manual"})
    Me.cbPreEffect.Location = New System.Drawing.Point(20, 47)
    Me.cbPreEffect.Name = "cbPreEffect"
    Me.cbPreEffect.Size = New System.Drawing.Size(174, 21)
    Me.cbPreEffect.TabIndex = 0
    '
    'gbBilateral
    '
    Me.gbBilateral.Controls.Add(Me.lbPicIntensity)
    Me.gbBilateral.Controls.Add(Me.lbSpatialIntensity)
    Me.gbBilateral.Controls.Add(Me.lbPicSpatial)
    Me.gbBilateral.Controls.Add(Me.pbSpatial)
    Me.gbBilateral.Controls.Add(Me.pbIntensity)
    Me.gbBilateral.Controls.Add(Me.lbIterations)
    Me.gbBilateral.Controls.Add(Me.tbIterations)
    Me.gbBilateral.Controls.Add(Me.lbSpatialSigma)
    Me.gbBilateral.Controls.Add(Me.tbSpatialSigma)
    Me.gbBilateral.Controls.Add(Me.lbRadius)
    Me.gbBilateral.Controls.Add(Me.tbRadius)
    Me.gbBilateral.Controls.Add(Me.tbSpatialIntensity)
    Me.gbBilateral.Controls.Add(Me.cbColorMode)
    Me.gbBilateral.Controls.Add(Me.cbIntensityMode)
    Me.gbBilateral.Controls.Add(Me.lbColorMode)
    Me.gbBilateral.Controls.Add(Me.lbIntensityMode)
    Me.gbBilateral.Location = New System.Drawing.Point(3, 227)
    Me.gbBilateral.Name = "gbBilateral"
    Me.gbBilateral.Size = New System.Drawing.Size(209, 414)
    Me.gbBilateral.TabIndex = 1
    Me.gbBilateral.TabStop = False
    Me.gbBilateral.Text = "Bilateral Filter"
    '
    'lbIntensityMode
    '
    Me.lbIntensityMode.AutoSize = True
    Me.lbIntensityMode.Location = New System.Drawing.Point(23, 78)
    Me.lbIntensityMode.Name = "lbIntensityMode"
    Me.lbIntensityMode.Size = New System.Drawing.Size(76, 13)
    Me.lbIntensityMode.TabIndex = 22
    Me.lbIntensityMode.Text = "Intensity Mode"
    '
    'cbIntensityMode
    '
    Me.cbIntensityMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cbIntensityMode.FormattingEnabled = True
    Me.cbIntensityMode.Items.AddRange(New Object() {"Gaussian", "Gaussian 2", "InvProp", "Linear"})
    Me.cbIntensityMode.Location = New System.Drawing.Point(20, 94)
    Me.cbIntensityMode.Name = "cbIntensityMode"
    Me.cbIntensityMode.Size = New System.Drawing.Size(174, 21)
    Me.cbIntensityMode.TabIndex = 21
    '
    'lbSpatialIntensity
    '
    Me.lbSpatialIntensity.AutoSize = True
    Me.lbSpatialIntensity.Location = New System.Drawing.Point(23, 127)
    Me.lbSpatialIntensity.Name = "lbSpatialIntensity"
    Me.lbSpatialIntensity.Size = New System.Drawing.Size(84, 13)
    Me.lbSpatialIntensity.TabIndex = 12
    Me.lbSpatialIntensity.Text = "Spatial Intensity:"
    '
    'lbPicSpatial
    '
    Me.lbPicSpatial.AutoSize = True
    Me.lbPicSpatial.Location = New System.Drawing.Point(130, 303)
    Me.lbPicSpatial.Name = "lbPicSpatial"
    Me.lbPicSpatial.Size = New System.Drawing.Size(39, 13)
    Me.lbPicSpatial.TabIndex = 20
    Me.lbPicSpatial.Text = "Spatial"
    '
    'lbPicIntensity
    '
    Me.lbPicIntensity.AutoSize = True
    Me.lbPicIntensity.Location = New System.Drawing.Point(38, 303)
    Me.lbPicIntensity.Name = "lbPicIntensity"
    Me.lbPicIntensity.Size = New System.Drawing.Size(46, 13)
    Me.lbPicIntensity.TabIndex = 19
    Me.lbPicIntensity.Text = "Intensity"
    '
    'pbSpatial
    '
    Me.pbSpatial.BackColor = System.Drawing.Color.White
    Me.pbSpatial.Location = New System.Drawing.Point(110, 320)
    Me.pbSpatial.Name = "pbSpatial"
    Me.pbSpatial.Size = New System.Drawing.Size(80, 80)
    Me.pbSpatial.TabIndex = 18
    Me.pbSpatial.TabStop = False
    '
    'pbIntensity
    '
    Me.pbIntensity.BackColor = System.Drawing.Color.White
    Me.pbIntensity.Location = New System.Drawing.Point(20, 320)
    Me.pbIntensity.Name = "pbIntensity"
    Me.pbIntensity.Size = New System.Drawing.Size(80, 80)
    Me.pbIntensity.TabIndex = 17
    Me.pbIntensity.TabStop = False
    '
    'lbIterations
    '
    Me.lbIterations.AutoSize = True
    Me.lbIterations.Location = New System.Drawing.Point(23, 261)
    Me.lbIterations.Name = "lbIterations"
    Me.lbIterations.Size = New System.Drawing.Size(53, 13)
    Me.lbIterations.TabIndex = 16
    Me.lbIterations.Text = "Iterations:"
    '
    'tbIterations
    '
    Me.tbIterations.Location = New System.Drawing.Point(17, 277)
    Me.tbIterations.Maximum = 25
    Me.tbIterations.Minimum = 1
    Me.tbIterations.Name = "tbIterations"
    Me.tbIterations.Size = New System.Drawing.Size(184, 45)
    Me.tbIterations.TabIndex = 15
    Me.tbIterations.TickStyle = System.Windows.Forms.TickStyle.None
    Me.tbIterations.Value = 3
    '
    'lbSpatialSigma
    '
    Me.lbSpatialSigma.AutoSize = True
    Me.lbSpatialSigma.Location = New System.Drawing.Point(23, 214)
    Me.lbSpatialSigma.Name = "lbSpatialSigma"
    Me.lbSpatialSigma.Size = New System.Drawing.Size(74, 13)
    Me.lbSpatialSigma.TabIndex = 14
    Me.lbSpatialSigma.Text = "Spatial Sigma:"
    '
    'tbSpatialSigma
    '
    Me.tbSpatialSigma.Location = New System.Drawing.Point(17, 230)
    Me.tbSpatialSigma.Maximum = 10000
    Me.tbSpatialSigma.Name = "tbSpatialSigma"
    Me.tbSpatialSigma.Size = New System.Drawing.Size(184, 45)
    Me.tbSpatialSigma.TabIndex = 13
    Me.tbSpatialSigma.TickStyle = System.Windows.Forms.TickStyle.None
    Me.tbSpatialSigma.Value = 150
    '
    'lbRadius
    '
    Me.lbRadius.AutoSize = True
    Me.lbRadius.Location = New System.Drawing.Point(23, 171)
    Me.lbRadius.Name = "lbRadius"
    Me.lbRadius.Size = New System.Drawing.Size(92, 13)
    Me.lbRadius.TabIndex = 10
    Me.lbRadius.Text = "Radius (HalfLate):"
    '
    'tbRadius
    '
    Me.tbRadius.Location = New System.Drawing.Point(17, 187)
    Me.tbRadius.Maximum = 30
    Me.tbRadius.Minimum = 1
    Me.tbRadius.Name = "tbRadius"
    Me.tbRadius.Size = New System.Drawing.Size(184, 45)
    Me.tbRadius.TabIndex = 9
    Me.tbRadius.TickStyle = System.Windows.Forms.TickStyle.None
    Me.tbRadius.Value = 3
    '
    'lbColorMode
    '
    Me.lbColorMode.AutoSize = True
    Me.lbColorMode.Location = New System.Drawing.Point(23, 28)
    Me.lbColorMode.Name = "lbColorMode"
    Me.lbColorMode.Size = New System.Drawing.Size(61, 13)
    Me.lbColorMode.TabIndex = 4
    Me.lbColorMode.Text = "Color Mode"
    '
    'cbColorMode
    '
    Me.cbColorMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cbColorMode.FormattingEnabled = True
    Me.cbColorMode.Items.AddRange(New Object() {"RGB", "Cie LAB"})
    Me.cbColorMode.Location = New System.Drawing.Point(20, 44)
    Me.cbColorMode.Name = "cbColorMode"
    Me.cbColorMode.Size = New System.Drawing.Size(174, 21)
    Me.cbColorMode.TabIndex = 3
    '
    'tbSpatialIntensity
    '
    Me.tbSpatialIntensity.Location = New System.Drawing.Point(17, 143)
    Me.tbSpatialIntensity.Maximum = 1000
    Me.tbSpatialIntensity.Name = "tbSpatialIntensity"
    Me.tbSpatialIntensity.Size = New System.Drawing.Size(184, 45)
    Me.tbSpatialIntensity.TabIndex = 11
    Me.tbSpatialIntensity.TickFrequency = 10
    Me.tbSpatialIntensity.TickStyle = System.Windows.Forms.TickStyle.None
    Me.tbSpatialIntensity.Value = 330
    '
    'gbContour
    '
    Me.gbContour.Controls.Add(Me.pnThreshold)
    Me.gbContour.Controls.Add(Me.pnLumHue)
    Me.gbContour.Controls.Add(Me.pnContourAmount)
    Me.gbContour.Controls.Add(Me.cbContourMethod)
    Me.gbContour.Controls.Add(Me.lbContourMethod)
    Me.gbContour.Location = New System.Drawing.Point(3, 647)
    Me.gbContour.Name = "gbContour"
    Me.gbContour.Size = New System.Drawing.Size(209, 169)
    Me.gbContour.TabIndex = 2
    Me.gbContour.TabStop = False
    Me.gbContour.Text = "Contour"
    '
    'pnThreshold
    '
    Me.pnThreshold.Controls.Add(Me.lbContourThreshold)
    Me.pnThreshold.Controls.Add(Me.tbContourThreshold)
    Me.pnThreshold.Location = New System.Drawing.Point(20, 117)
    Me.pnThreshold.Name = "pnThreshold"
    Me.pnThreshold.Size = New System.Drawing.Size(174, 46)
    Me.pnThreshold.TabIndex = 28
    Me.pnThreshold.Visible = False
    '
    'lbContourThreshold
    '
    Me.lbContourThreshold.AutoSize = True
    Me.lbContourThreshold.Location = New System.Drawing.Point(6, 0)
    Me.lbContourThreshold.Name = "lbContourThreshold"
    Me.lbContourThreshold.Size = New System.Drawing.Size(57, 13)
    Me.lbContourThreshold.TabIndex = 10
    Me.lbContourThreshold.Text = "Threshold:"
    '
    'tbContourThreshold
    '
    Me.tbContourThreshold.Location = New System.Drawing.Point(0, 17)
    Me.tbContourThreshold.Maximum = 1000
    Me.tbContourThreshold.Name = "tbContourThreshold"
    Me.tbContourThreshold.Size = New System.Drawing.Size(174, 45)
    Me.tbContourThreshold.TabIndex = 9
    Me.tbContourThreshold.TickStyle = System.Windows.Forms.TickStyle.None
    '
    'pnLumHue
    '
    Me.pnLumHue.Controls.Add(Me.lbLumHue)
    Me.pnLumHue.Controls.Add(Me.tbLumHue)
    Me.pnLumHue.Location = New System.Drawing.Point(20, 117)
    Me.pnLumHue.Name = "pnLumHue"
    Me.pnLumHue.Size = New System.Drawing.Size(174, 46)
    Me.pnLumHue.TabIndex = 27
    Me.pnLumHue.Visible = False
    '
    'lbLumHue
    '
    Me.lbLumHue.AutoSize = True
    Me.lbLumHue.Location = New System.Drawing.Point(6, 0)
    Me.lbLumHue.Name = "lbLumHue"
    Me.lbLumHue.Size = New System.Drawing.Size(144, 13)
    Me.lbLumHue.TabIndex = 10
    Me.lbLumHue.Text = "Based On 100% L 100% AB :"
    '
    'tbLumHue
    '
    Me.tbLumHue.Location = New System.Drawing.Point(0, 17)
    Me.tbLumHue.Maximum = 100
    Me.tbLumHue.Name = "tbLumHue"
    Me.tbLumHue.Size = New System.Drawing.Size(174, 45)
    Me.tbLumHue.TabIndex = 9
    Me.tbLumHue.TickStyle = System.Windows.Forms.TickStyle.None
    '
    'pnContourAmount
    '
    Me.pnContourAmount.Controls.Add(Me.lbContourAmount)
    Me.pnContourAmount.Controls.Add(Me.tbContourAmount)
    Me.pnContourAmount.Location = New System.Drawing.Point(20, 77)
    Me.pnContourAmount.Name = "pnContourAmount"
    Me.pnContourAmount.Size = New System.Drawing.Size(174, 40)
    Me.pnContourAmount.TabIndex = 26
    '
    'lbContourAmount
    '
    Me.lbContourAmount.AutoSize = True
    Me.lbContourAmount.Location = New System.Drawing.Point(6, -2)
    Me.lbContourAmount.Name = "lbContourAmount"
    Me.lbContourAmount.Size = New System.Drawing.Size(86, 13)
    Me.lbContourAmount.TabIndex = 12
    Me.lbContourAmount.Text = "Contour Amount:"
    '
    'tbContourAmount
    '
    Me.tbContourAmount.Location = New System.Drawing.Point(0, 14)
    Me.tbContourAmount.Maximum = 200
    Me.tbContourAmount.Name = "tbContourAmount"
    Me.tbContourAmount.Size = New System.Drawing.Size(174, 45)
    Me.tbContourAmount.TabIndex = 11
    Me.tbContourAmount.TickStyle = System.Windows.Forms.TickStyle.None
    '
    'lbContourMethod
    '
    Me.lbContourMethod.AutoSize = True
    Me.lbContourMethod.Location = New System.Drawing.Point(23, 28)
    Me.lbContourMethod.Name = "lbContourMethod"
    Me.lbContourMethod.Size = New System.Drawing.Size(120, 13)
    Me.lbContourMethod.TabIndex = 25
    Me.lbContourMethod.Text = "Edge Detection Method"
    '
    'cbContourMethod
    '
    Me.cbContourMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cbContourMethod.FormattingEnabled = True
    Me.cbContourMethod.Items.AddRange(New Object() {"Sobel", "Difference of Gaussians"})
    Me.cbContourMethod.Location = New System.Drawing.Point(20, 44)
    Me.cbContourMethod.Name = "cbContourMethod"
    Me.cbContourMethod.Size = New System.Drawing.Size(174, 21)
    Me.cbContourMethod.TabIndex = 24
    '
    'flpEffect
    '
    Me.flpEffect.AutoScroll = True
    Me.flpEffect.BackColor = System.Drawing.SystemColors.Control
    Me.flpEffect.Controls.Add(Me.gbPreEffect)
    Me.flpEffect.Controls.Add(Me.gbBilateral)
    Me.flpEffect.Controls.Add(Me.gbContour)
    Me.flpEffect.Controls.Add(Me.gbLumSeg)
    Me.flpEffect.Dock = System.Windows.Forms.DockStyle.Left
    Me.flpEffect.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
    Me.flpEffect.Location = New System.Drawing.Point(0, 0)
    Me.flpEffect.Name = "flpEffect"
    Me.flpEffect.Size = New System.Drawing.Size(232, 437)
    Me.flpEffect.TabIndex = 3
    Me.flpEffect.WrapContents = False
    '
    'gbLumSeg
    '
    Me.gbLumSeg.Controls.Add(Me.lbPresence)
    Me.gbLumSeg.Controls.Add(Me.tbPresence)
    Me.gbLumSeg.Controls.Add(Me.lbSegments)
    Me.gbLumSeg.Controls.Add(Me.tbSegments)
    Me.gbLumSeg.Location = New System.Drawing.Point(3, 822)
    Me.gbLumSeg.Name = "gbLumSeg"
    Me.gbLumSeg.Size = New System.Drawing.Size(209, 141)
    Me.gbLumSeg.TabIndex = 4
    Me.gbLumSeg.TabStop = False
    Me.gbLumSeg.Text = "Luminance Segmentation"
    '
    'lbPresence
    '
    Me.lbPresence.AutoSize = True
    Me.lbPresence.Location = New System.Drawing.Point(26, 74)
    Me.lbPresence.Name = "lbPresence"
    Me.lbPresence.Size = New System.Drawing.Size(55, 13)
    Me.lbPresence.TabIndex = 10
    Me.lbPresence.Text = "Presence:"
    '
    'tbPresence
    '
    Me.tbPresence.Location = New System.Drawing.Point(20, 91)
    Me.tbPresence.Maximum = 100
    Me.tbPresence.Name = "tbPresence"
    Me.tbPresence.Size = New System.Drawing.Size(174, 45)
    Me.tbPresence.TabIndex = 9
    Me.tbPresence.TickStyle = System.Windows.Forms.TickStyle.None
    '
    'lbSegments
    '
    Me.lbSegments.AutoSize = True
    Me.lbSegments.Location = New System.Drawing.Point(26, 26)
    Me.lbSegments.Name = "lbSegments"
    Me.lbSegments.Size = New System.Drawing.Size(57, 13)
    Me.lbSegments.TabIndex = 12
    Me.lbSegments.Text = "Segments:"
    '
    'tbSegments
    '
    Me.tbSegments.Location = New System.Drawing.Point(20, 42)
    Me.tbSegments.Minimum = 2
    Me.tbSegments.Name = "tbSegments"
    Me.tbSegments.Size = New System.Drawing.Size(174, 45)
    Me.tbSegments.TabIndex = 11
    Me.tbSegments.TickStyle = System.Windows.Forms.TickStyle.None
    Me.tbSegments.Value = 2
    '
    'pnPicCtrl
    '
    Me.pnPicCtrl.Controls.Add(Me.pnPic)
    Me.pnPicCtrl.Controls.Add(Me.pnCtrl)
    Me.pnPicCtrl.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnPicCtrl.Location = New System.Drawing.Point(232, 0)
    Me.pnPicCtrl.Name = "pnPicCtrl"
    Me.pnPicCtrl.Size = New System.Drawing.Size(636, 437)
    Me.pnPicCtrl.TabIndex = 4
    '
    'pnPic
    '
    Me.pnPic.AutoScroll = True
    Me.pnPic.BackColor = System.Drawing.Color.Black
    Me.pnPic.Controls.Add(Me.PictureBox1)
    Me.pnPic.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnPic.Location = New System.Drawing.Point(0, 0)
    Me.pnPic.Name = "pnPic"
    Me.pnPic.Size = New System.Drawing.Size(636, 398)
    Me.pnPic.TabIndex = 1
    '
    'PictureBox1
    '
    Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
    Me.PictureBox1.Name = "PictureBox1"
    Me.PictureBox1.Size = New System.Drawing.Size(636, 398)
    Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
    Me.PictureBox1.TabIndex = 0
    Me.PictureBox1.TabStop = False
    '
    'pnCtrl
    '
    Me.pnCtrl.Controls.Add(Me.tbImage)
    Me.pnCtrl.Controls.Add(Me.btBrowse)
    Me.pnCtrl.Controls.Add(Me.btCartoonize)
    Me.pnCtrl.Controls.Add(Me.pbCartoonize)
    Me.pnCtrl.Controls.Add(Me.btSave)
    Me.pnCtrl.Dock = System.Windows.Forms.DockStyle.Bottom
    Me.pnCtrl.Location = New System.Drawing.Point(0, 398)
    Me.pnCtrl.Name = "pnCtrl"
    Me.pnCtrl.Size = New System.Drawing.Size(636, 39)
    Me.pnCtrl.TabIndex = 0
    '
    'tbImage
    '
    Me.tbImage.Location = New System.Drawing.Point(87, 8)
    Me.tbImage.Name = "tbImage"
    Me.tbImage.ReadOnly = True
    Me.tbImage.Size = New System.Drawing.Size(238, 20)
    Me.tbImage.TabIndex = 4
    '
    'btBrowse
    '
    Me.btBrowse.Location = New System.Drawing.Point(6, 6)
    Me.btBrowse.Name = "btBrowse"
    Me.btBrowse.Size = New System.Drawing.Size(75, 23)
    Me.btBrowse.TabIndex = 3
    Me.btBrowse.Text = "Load Image"
    Me.btBrowse.UseVisualStyleBackColor = True
    '
    'btCartoonize
    '
    Me.btCartoonize.Location = New System.Drawing.Point(490, 6)
    Me.btCartoonize.Name = "btCartoonize"
    Me.btCartoonize.Size = New System.Drawing.Size(79, 23)
    Me.btCartoonize.TabIndex = 2
    Me.btCartoonize.Text = "Cartoonize"
    Me.btCartoonize.UseVisualStyleBackColor = True
    '
    'pbCartoonize
    '
    Me.pbCartoonize.ForeColor = System.Drawing.Color.GreenYellow
    Me.pbCartoonize.Location = New System.Drawing.Point(331, 6)
    Me.pbCartoonize.Name = "pbCartoonize"
    Me.pbCartoonize.Size = New System.Drawing.Size(153, 23)
    Me.pbCartoonize.Step = 1
    Me.pbCartoonize.Style = System.Windows.Forms.ProgressBarStyle.Continuous
    Me.pbCartoonize.TabIndex = 1
    '
    'btSave
    '
    Me.btSave.Location = New System.Drawing.Point(575, 6)
    Me.btSave.Name = "btSave"
    Me.btSave.Size = New System.Drawing.Size(58, 23)
    Me.btSave.TabIndex = 0
    Me.btSave.Text = "Save"
    Me.btSave.UseVisualStyleBackColor = True
    '
    'OpenFileDialog1
    '
    Me.OpenFileDialog1.FileName = "OpenFileDialog1"
    Me.OpenFileDialog1.Filter = "Image Files|*.jpg;*.png;*.tif;*.tiff;*.bmp"
    '
    'SaveFileDialog1
    '
    Me.SaveFileDialog1.Filter = "Image Files|*.jpg;*.png;*.tif;*.tiff;*.bmp"
    '
    'CartoonizerForm
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(868, 437)
    Me.Controls.Add(Me.pnPicCtrl)
    Me.Controls.Add(Me.flpEffect)
    Me.KeyPreview = True
    Me.Name = "CartoonizerForm"
    Me.Text = "Cartoonizer"
    Me.gbPreEffect.ResumeLayout(False)
    Me.gbPreEffect.PerformLayout()
    Me.pnExposure.ResumeLayout(False)
    Me.pnExposure.PerformLayout()
    CType(Me.tbExposure, System.ComponentModel.ISupportInitialize).EndInit()
    Me.pnBCS.ResumeLayout(False)
    Me.pnBCS.PerformLayout()
    CType(Me.tbSaturation, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.tbBrightness, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.tbContrast, System.ComponentModel.ISupportInitialize).EndInit()
    Me.gbBilateral.ResumeLayout(False)
    Me.gbBilateral.PerformLayout()
    CType(Me.pbSpatial, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.pbIntensity, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.tbIterations, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.tbSpatialSigma, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.tbRadius, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.tbSpatialIntensity, System.ComponentModel.ISupportInitialize).EndInit()
    Me.gbContour.ResumeLayout(False)
    Me.gbContour.PerformLayout()
    Me.pnThreshold.ResumeLayout(False)
    Me.pnThreshold.PerformLayout()
    CType(Me.tbContourThreshold, System.ComponentModel.ISupportInitialize).EndInit()
    Me.pnLumHue.ResumeLayout(False)
    Me.pnLumHue.PerformLayout()
    CType(Me.tbLumHue, System.ComponentModel.ISupportInitialize).EndInit()
    Me.pnContourAmount.ResumeLayout(False)
    Me.pnContourAmount.PerformLayout()
    CType(Me.tbContourAmount, System.ComponentModel.ISupportInitialize).EndInit()
    Me.flpEffect.ResumeLayout(False)
    Me.gbLumSeg.ResumeLayout(False)
    Me.gbLumSeg.PerformLayout()
    CType(Me.tbPresence, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.tbSegments, System.ComponentModel.ISupportInitialize).EndInit()
    Me.pnPicCtrl.ResumeLayout(False)
    Me.pnPic.ResumeLayout(False)
    CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.pnCtrl.ResumeLayout(False)
    Me.pnCtrl.PerformLayout()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents gbPreEffect As System.Windows.Forms.GroupBox
  Friend WithEvents pnExposure As System.Windows.Forms.Panel
  Friend WithEvents lbExposure As System.Windows.Forms.Label
  Friend WithEvents tbExposure As System.Windows.Forms.TrackBar
  Friend WithEvents cbPreEffect As System.Windows.Forms.ComboBox
  Friend WithEvents pnBCS As System.Windows.Forms.Panel
  Friend WithEvents lbBrightness As System.Windows.Forms.Label
  Friend WithEvents lbContrast As System.Windows.Forms.Label
  Friend WithEvents lbSaturation As System.Windows.Forms.Label
  Friend WithEvents tbSaturation As System.Windows.Forms.TrackBar
  Friend WithEvents tbBrightness As System.Windows.Forms.TrackBar
  Friend WithEvents tbContrast As System.Windows.Forms.TrackBar
  Friend WithEvents gbBilateral As System.Windows.Forms.GroupBox
  Friend WithEvents cbColorMode As System.Windows.Forms.ComboBox
  Friend WithEvents lbColorMode As System.Windows.Forms.Label
  Friend WithEvents lbRadius As System.Windows.Forms.Label
  Friend WithEvents tbRadius As System.Windows.Forms.TrackBar
  Friend WithEvents lbSpatialSigma As System.Windows.Forms.Label
  Friend WithEvents tbSpatialSigma As System.Windows.Forms.TrackBar
  Friend WithEvents lbSpatialIntensity As System.Windows.Forms.Label
  Friend WithEvents tbSpatialIntensity As System.Windows.Forms.TrackBar
  Friend WithEvents pbIntensity As System.Windows.Forms.PictureBox
  Friend WithEvents lbIterations As System.Windows.Forms.Label
  Friend WithEvents tbIterations As System.Windows.Forms.TrackBar
  Friend WithEvents pbSpatial As System.Windows.Forms.PictureBox
  Friend WithEvents lbIntensityMode As System.Windows.Forms.Label
  Friend WithEvents cbIntensityMode As System.Windows.Forms.ComboBox
  Friend WithEvents lbPicSpatial As System.Windows.Forms.Label
  Friend WithEvents lbPicIntensity As System.Windows.Forms.Label
  Friend WithEvents lbPreEffect As System.Windows.Forms.Label
  Friend WithEvents gbContour As System.Windows.Forms.GroupBox
  Friend WithEvents lbContourMethod As System.Windows.Forms.Label
  Friend WithEvents cbContourMethod As System.Windows.Forms.ComboBox
  Friend WithEvents pnContourAmount As System.Windows.Forms.Panel
  Friend WithEvents lbContourAmount As System.Windows.Forms.Label
  Friend WithEvents lbLumHue As System.Windows.Forms.Label
  Friend WithEvents tbContourAmount As System.Windows.Forms.TrackBar
  Friend WithEvents tbLumHue As System.Windows.Forms.TrackBar
  Friend WithEvents pnThreshold As System.Windows.Forms.Panel
  Friend WithEvents lbContourThreshold As System.Windows.Forms.Label
  Friend WithEvents tbContourThreshold As System.Windows.Forms.TrackBar
  Friend WithEvents pnLumHue As System.Windows.Forms.Panel
  Friend WithEvents flpEffect As System.Windows.Forms.FlowLayoutPanel
  Friend WithEvents gbLumSeg As System.Windows.Forms.GroupBox
  Friend WithEvents lbPresence As System.Windows.Forms.Label
  Friend WithEvents tbPresence As System.Windows.Forms.TrackBar
  Friend WithEvents lbSegments As System.Windows.Forms.Label
  Friend WithEvents tbSegments As System.Windows.Forms.TrackBar
  Friend WithEvents pnPicCtrl As System.Windows.Forms.Panel
  Friend WithEvents pnPic As System.Windows.Forms.Panel
  Friend WithEvents pnCtrl As System.Windows.Forms.Panel
  Friend WithEvents btSave As System.Windows.Forms.Button
  Friend WithEvents tbImage As System.Windows.Forms.TextBox
  Friend WithEvents btBrowse As System.Windows.Forms.Button
  Friend WithEvents btCartoonize As System.Windows.Forms.Button
  Friend WithEvents pbCartoonize As System.Windows.Forms.ProgressBar
  Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
  Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
  Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
End Class
