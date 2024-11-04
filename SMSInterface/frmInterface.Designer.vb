<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInterface
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInterface))
        Me.grpMessage = New System.Windows.Forms.GroupBox()
        Me.txtMsg = New System.Windows.Forms.TextBox()
        Me.ofd = New System.Windows.Forms.OpenFileDialog()
        Me.fbd = New System.Windows.Forms.FolderBrowserDialog()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Rtb1 = New System.Windows.Forms.RichTextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.dtpNext = New System.Windows.Forms.DateTimePicker()
        Me.dtpLast = New System.Windows.Forms.DateTimePicker()
        Me.dtpNow = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.BtnBack = New System.Windows.Forms.Button()
        Me.ChkLenght = New System.Windows.Forms.CheckBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.ChkSize = New System.Windows.Forms.CheckBox()
        Me.ChkGrade = New System.Windows.Forms.CheckBox()
        Me.ChkStandard = New System.Windows.Forms.CheckBox()
        Me.ChkProductType = New System.Windows.Forms.CheckBox()
        Me.ChkItem = New System.Windows.Forms.CheckBox()
        Me.btnProcess = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.pnlHead = New System.Windows.Forms.Panel()
        Me.picLogo = New System.Windows.Forms.PictureBox()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.DtNowDisplay = New System.Windows.Forms.DateTimePicker()
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.btnSchedule = New System.Windows.Forms.Button()
        Me.btnSetting = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtInterval = New System.Windows.Forms.TextBox()
        Me.grpMessage.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.pnlHead.SuspendLayout()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpMessage
        '
        Me.grpMessage.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpMessage.Controls.Add(Me.txtMsg)
        Me.grpMessage.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpMessage.Location = New System.Drawing.Point(24, 494)
        Me.grpMessage.Name = "grpMessage"
        Me.grpMessage.Size = New System.Drawing.Size(652, 39)
        Me.grpMessage.TabIndex = 47
        Me.grpMessage.TabStop = False
        '
        'txtMsg
        '
        Me.txtMsg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMsg.BackColor = System.Drawing.Color.White
        Me.txtMsg.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtMsg.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMsg.ForeColor = System.Drawing.Color.DarkBlue
        Me.txtMsg.Location = New System.Drawing.Point(6, 12)
        Me.txtMsg.Multiline = True
        Me.txtMsg.Name = "txtMsg"
        Me.txtMsg.ReadOnly = True
        Me.txtMsg.Size = New System.Drawing.Size(640, 20)
        Me.txtMsg.TabIndex = 1
        Me.txtMsg.TabStop = False
        Me.txtMsg.Text = "Message"
        Me.txtMsg.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'Rtb1
        '
        Me.Rtb1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Rtb1.Location = New System.Drawing.Point(23, 132)
        Me.Rtb1.Name = "Rtb1"
        Me.Rtb1.ReadOnly = True
        Me.Rtb1.Size = New System.Drawing.Size(654, 218)
        Me.Rtb1.TabIndex = 53
        Me.Rtb1.Text = ""
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.DarkBlue
        Me.Label13.Location = New System.Drawing.Point(20, 98)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(90, 13)
        Me.Label13.TabIndex = 55
        Me.Label13.Text = "NEXT PROCESS"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.DarkBlue
        Me.Label14.Location = New System.Drawing.Point(20, 66)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(88, 13)
        Me.Label14.TabIndex = 54
        Me.Label14.Text = "LAST PROCESS"
        '
        'dtpNext
        '
        Me.dtpNext.CustomFormat = "dd MMM yyyy  hh:mm:ss tt"
        Me.dtpNext.Enabled = False
        Me.dtpNext.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpNext.Location = New System.Drawing.Point(145, 94)
        Me.dtpNext.Name = "dtpNext"
        Me.dtpNext.Size = New System.Drawing.Size(213, 20)
        Me.dtpNext.TabIndex = 59
        '
        'dtpLast
        '
        Me.dtpLast.CustomFormat = "dd MMM yyyy  hh:mm:ss tt"
        Me.dtpLast.Enabled = False
        Me.dtpLast.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpLast.Location = New System.Drawing.Point(145, 64)
        Me.dtpLast.Name = "dtpLast"
        Me.dtpLast.Size = New System.Drawing.Size(213, 20)
        Me.dtpLast.TabIndex = 58
        '
        'dtpNow
        '
        Me.dtpNow.CustomFormat = "dd MMM yyyy  hh:mm:ss tt"
        Me.dtpNow.Enabled = False
        Me.dtpNow.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpNow.Location = New System.Drawing.Point(490, 64)
        Me.dtpNow.Name = "dtpNow"
        Me.dtpNow.Size = New System.Drawing.Size(213, 20)
        Me.dtpNow.TabIndex = 60
        Me.dtpNow.Visible = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.BtnBack)
        Me.GroupBox3.Controls.Add(Me.ChkLenght)
        Me.GroupBox3.Controls.Add(Me.CheckBox1)
        Me.GroupBox3.Controls.Add(Me.CheckBox2)
        Me.GroupBox3.Controls.Add(Me.CheckBox3)
        Me.GroupBox3.Controls.Add(Me.ChkSize)
        Me.GroupBox3.Controls.Add(Me.ChkGrade)
        Me.GroupBox3.Controls.Add(Me.ChkStandard)
        Me.GroupBox3.Controls.Add(Me.ChkProductType)
        Me.GroupBox3.Controls.Add(Me.ChkItem)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.DarkBlue
        Me.GroupBox3.Location = New System.Drawing.Point(621, 710)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(62, 99)
        Me.GroupBox3.TabIndex = 64
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "SELECT MASTER"
        Me.GroupBox3.Visible = False
        '
        'BtnBack
        '
        Me.BtnBack.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnBack.BackColor = System.Drawing.SystemColors.Control
        Me.BtnBack.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnBack.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnBack.Location = New System.Drawing.Point(-32, 61)
        Me.BtnBack.Name = "BtnBack"
        Me.BtnBack.Size = New System.Drawing.Size(88, 28)
        Me.BtnBack.TabIndex = 51
        Me.BtnBack.Text = "       &BACK"
        Me.BtnBack.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.BtnBack.UseVisualStyleBackColor = False
        '
        'ChkLenght
        '
        Me.ChkLenght.AutoSize = True
        Me.ChkLenght.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkLenght.Location = New System.Drawing.Point(146, 72)
        Me.ChkLenght.Name = "ChkLenght"
        Me.ChkLenght.Size = New System.Drawing.Size(70, 17)
        Me.ChkLenght.TabIndex = 10
        Me.ChkLenght.Text = "LENGTH"
        Me.ChkLenght.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(402, 72)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(94, 17)
        Me.CheckBox1.TabIndex = 9
        Me.CheckBox1.Text = "JOURNAL CR"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox2.Location = New System.Drawing.Point(402, 25)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(94, 17)
        Me.CheckBox2.TabIndex = 8
        Me.CheckBox2.Text = "JOURNAL DP"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox3.Location = New System.Drawing.Point(402, 49)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(122, 17)
        Me.CheckBox3.TabIndex = 7
        Me.CheckBox3.Text = "JOURNAL INVOICE"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'ChkSize
        '
        Me.ChkSize.AutoSize = True
        Me.ChkSize.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkSize.Location = New System.Drawing.Point(146, 49)
        Me.ChkSize.Name = "ChkSize"
        Me.ChkSize.Size = New System.Drawing.Size(87, 17)
        Me.ChkSize.TabIndex = 6
        Me.ChkSize.Text = "SIZE CLASS"
        Me.ChkSize.UseVisualStyleBackColor = True
        '
        'ChkGrade
        '
        Me.ChkGrade.AutoSize = True
        Me.ChkGrade.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkGrade.Location = New System.Drawing.Point(146, 26)
        Me.ChkGrade.Name = "ChkGrade"
        Me.ChkGrade.Size = New System.Drawing.Size(64, 17)
        Me.ChkGrade.TabIndex = 3
        Me.ChkGrade.Text = "GRADE"
        Me.ChkGrade.UseVisualStyleBackColor = True
        '
        'ChkStandard
        '
        Me.ChkStandard.AutoSize = True
        Me.ChkStandard.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkStandard.Location = New System.Drawing.Point(21, 72)
        Me.ChkStandard.Name = "ChkStandard"
        Me.ChkStandard.Size = New System.Drawing.Size(86, 17)
        Me.ChkStandard.TabIndex = 2
        Me.ChkStandard.Text = "STANDARD"
        Me.ChkStandard.UseVisualStyleBackColor = True
        '
        'ChkProductType
        '
        Me.ChkProductType.AutoSize = True
        Me.ChkProductType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkProductType.Location = New System.Drawing.Point(21, 49)
        Me.ChkProductType.Name = "ChkProductType"
        Me.ChkProductType.Size = New System.Drawing.Size(110, 17)
        Me.ChkProductType.TabIndex = 1
        Me.ChkProductType.Text = "PRODUCT TYPE"
        Me.ChkProductType.UseVisualStyleBackColor = True
        '
        'ChkItem
        '
        Me.ChkItem.AutoSize = True
        Me.ChkItem.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkItem.Location = New System.Drawing.Point(21, 26)
        Me.ChkItem.Name = "ChkItem"
        Me.ChkItem.Size = New System.Drawing.Size(100, 17)
        Me.ChkItem.TabIndex = 0
        Me.ChkItem.Text = "MASTER ITEM"
        Me.ChkItem.UseVisualStyleBackColor = True
        '
        'btnProcess
        '
        Me.btnProcess.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnProcess.BackColor = System.Drawing.SystemColors.Control
        Me.btnProcess.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnProcess.Image = CType(resources.GetObject("btnProcess.Image"), System.Drawing.Image)
        Me.btnProcess.Location = New System.Drawing.Point(543, 371)
        Me.btnProcess.Name = "btnProcess"
        Me.btnProcess.Size = New System.Drawing.Size(134, 28)
        Me.btnProcess.TabIndex = 51
        Me.btnProcess.Text = "  &MANUAL PROCESS"
        Me.btnProcess.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnProcess.UseVisualStyleBackColor = False
        '
        'btnStart
        '
        Me.btnStart.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStart.BackColor = System.Drawing.SystemColors.Control
        Me.btnStart.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStart.Image = CType(resources.GetObject("btnStart.Image"), System.Drawing.Image)
        Me.btnStart.Location = New System.Drawing.Point(448, 371)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(89, 28)
        Me.btnStart.TabIndex = 52
        Me.btnStart.Text = "  S&TART"
        Me.btnStart.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnStart.UseVisualStyleBackColor = False
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnClose.BackColor = System.Drawing.SystemColors.Control
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Image = CType(resources.GetObject("btnClose.Image"), System.Drawing.Image)
        Me.btnClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnClose.Location = New System.Drawing.Point(24, 374)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(88, 28)
        Me.btnClose.TabIndex = 49
        Me.btnClose.Text = "       &CLOSE"
        Me.btnClose.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'pnlHead
        '
        Me.pnlHead.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlHead.BackColor = System.Drawing.Color.White
        Me.pnlHead.BackgroundImage = CType(resources.GetObject("pnlHead.BackgroundImage"), System.Drawing.Image)
        Me.pnlHead.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlHead.Controls.Add(Me.picLogo)
        Me.pnlHead.Controls.Add(Me.lblTitle)
        Me.pnlHead.Location = New System.Drawing.Point(4, 1)
        Me.pnlHead.Name = "pnlHead"
        Me.pnlHead.Size = New System.Drawing.Size(704, 43)
        Me.pnlHead.TabIndex = 22
        '
        'picLogo
        '
        Me.picLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.picLogo.Image = CType(resources.GetObject("picLogo.Image"), System.Drawing.Image)
        Me.picLogo.Location = New System.Drawing.Point(0, 0)
        Me.picLogo.Name = "picLogo"
        Me.picLogo.Size = New System.Drawing.Size(82, 38)
        Me.picLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picLogo.TabIndex = 6
        Me.picLogo.TabStop = False
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.BackColor = System.Drawing.Color.Transparent
        Me.lblTitle.Font = New System.Drawing.Font("Tahoma", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.DarkBlue
        Me.lblTitle.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblTitle.Location = New System.Drawing.Point(87, 13)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(253, 25)
        Me.lblTitle.TabIndex = 1
        Me.lblTitle.Text = "Interface Data Process"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.DarkBlue
        Me.Label15.Location = New System.Drawing.Point(407, 66)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(33, 13)
        Me.Label15.TabIndex = 65
        Me.Label15.Text = "TIME"
        '
        'DtNowDisplay
        '
        Me.DtNowDisplay.CustomFormat = "dd MMM yyyy  hh:mm:ss tt"
        Me.DtNowDisplay.Enabled = False
        Me.DtNowDisplay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtNowDisplay.Location = New System.Drawing.Point(490, 64)
        Me.DtNowDisplay.Name = "DtNowDisplay"
        Me.DtNowDisplay.Size = New System.Drawing.Size(213, 20)
        Me.DtNowDisplay.TabIndex = 70
        '
        'Timer2
        '
        Me.Timer2.Enabled = True
        '
        'btnSchedule
        '
        Me.btnSchedule.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSchedule.BackColor = System.Drawing.SystemColors.Control
        Me.btnSchedule.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnSchedule.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSchedule.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSchedule.Location = New System.Drawing.Point(321, 398)
        Me.btnSchedule.Name = "btnSchedule"
        Me.btnSchedule.Size = New System.Drawing.Size(88, 28)
        Me.btnSchedule.TabIndex = 71
        Me.btnSchedule.Text = "    &SCHEDULE"
        Me.btnSchedule.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSchedule.UseVisualStyleBackColor = False
        Me.btnSchedule.Visible = False
        '
        'btnSetting
        '
        Me.btnSetting.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSetting.BackColor = System.Drawing.SystemColors.Control
        Me.btnSetting.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnSetting.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSetting.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSetting.Location = New System.Drawing.Point(118, 374)
        Me.btnSetting.Name = "btnSetting"
        Me.btnSetting.Size = New System.Drawing.Size(88, 28)
        Me.btnSetting.TabIndex = 72
        Me.btnSetting.Text = "      &SETTING"
        Me.btnSetting.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSetting.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.DarkBlue
        Me.Label1.Location = New System.Drawing.Point(407, 98)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 73
        Me.Label1.Text = "INTERVAL"
        '
        'txtInterval
        '
        Me.txtInterval.Location = New System.Drawing.Point(490, 94)
        Me.txtInterval.Name = "txtInterval"
        Me.txtInterval.Size = New System.Drawing.Size(33, 20)
        Me.txtInterval.TabIndex = 74
        Me.txtInterval.Text = "3"
        '
        'frmInterface
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(713, 438)
        Me.Controls.Add(Me.txtInterval)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnSetting)
        Me.Controls.Add(Me.btnSchedule)
        Me.Controls.Add(Me.DtNowDisplay)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.dtpNow)
        Me.Controls.Add(Me.dtpLast)
        Me.Controls.Add(Me.dtpNext)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.btnProcess)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.Rtb1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.grpMessage)
        Me.Controls.Add(Me.pnlHead)
        Me.Controls.Add(Me.GroupBox3)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmInterface"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Batch Proses - Interface Data"
        Me.grpMessage.ResumeLayout(False)
        Me.grpMessage.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.pnlHead.ResumeLayout(False)
        Me.pnlHead.PerformLayout()
        CType(Me.picLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents pnlHead As System.Windows.Forms.Panel
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents grpMessage As System.Windows.Forms.GroupBox
    Public WithEvents txtMsg As System.Windows.Forms.TextBox
    Friend WithEvents ofd As System.Windows.Forms.OpenFileDialog
    Friend WithEvents fbd As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Public WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnProcess As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents Rtb1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents picLogo As System.Windows.Forms.PictureBox
    Friend WithEvents dtpNext As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpLast As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpNow As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents ChkSize As System.Windows.Forms.CheckBox
    Friend WithEvents ChkGrade As System.Windows.Forms.CheckBox
    Friend WithEvents ChkStandard As System.Windows.Forms.CheckBox
    Friend WithEvents ChkProductType As System.Windows.Forms.CheckBox
    Friend WithEvents ChkItem As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLenght As System.Windows.Forms.CheckBox
    Public WithEvents BtnBack As System.Windows.Forms.Button
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents DtNowDisplay As System.Windows.Forms.DateTimePicker
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Public WithEvents btnSchedule As System.Windows.Forms.Button
    Public WithEvents btnSetting As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtInterval As System.Windows.Forms.TextBox

End Class
