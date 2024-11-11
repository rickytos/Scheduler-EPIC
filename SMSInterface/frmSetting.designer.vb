<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSetting
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
        Me.btnSave = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TabEmailSettings = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.txtSubject_DMMS_PR = New System.Windows.Forms.TextBox()
        Me.txtEmailTo_DMMS_PR = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.txtEmailFrom_DMMS_PR = New System.Windows.Forms.TextBox()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.txtSMTPPort_DMMS_PR = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtSMTPAddress_DMMS_PR = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtSMTPPassword_DMMS_PR = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtSMTPUser_DMMS_PR = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabPathFolders = New System.Windows.Forms.TabControl()
        Me.tpEPIC1SM = New System.Windows.Forms.TabPage()
        Me.txtJSONPathBackup = New System.Windows.Forms.TextBox()
        Me.txtJSONPath = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tpESSBudget = New System.Windows.Forms.TabPage()
        Me.txtESSBudgetPathBackup = New System.Windows.Forms.TextBox()
        Me.txtESSBudgetPath = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.tpSAP = New System.Windows.Forms.TabPage()
        Me.txtSAPFilePathBackup = New System.Windows.Forms.TextBox()
        Me.txtSAPFilePath = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.tpIAPrice = New System.Windows.Forms.TabPage()
        Me.txtIAPriceDestinationPath = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.tpSAPCostMonitoring = New System.Windows.Forms.TabPage()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.txtIAMIPressPartLocalComponentPathBackup = New System.Windows.Forms.TextBox()
        Me.txtIAMIPressPartLocalComponentPath = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtIAMIPressPartLocalComponentXMLToSAPPath = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtCostHistoryXMLToSAPPath = New System.Windows.Forms.TextBox()
        Me.tpPurchaseRequest = New System.Windows.Forms.TabPage()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.txtPurchaseRequestPathOut = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.txtPurchaseRequestPathFailed = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.txtGoodsReceiveXMLToDMMSPath = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.txtPurchaseRequestPathBackup = New System.Windows.Forms.TextBox()
        Me.txtPurchaseRequestPath = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.tpMaterialMaster = New System.Windows.Forms.TabPage()
        Me.txtMaterialMasterXMLtoSAPBackupPath_Res = New System.Windows.Forms.TextBox()
        Me.txtMaterialMasterXMLtoSAPPath_Res = New System.Windows.Forms.TextBox()
        Me.txtMaterialMasterXMLtoSAPBackupPath_Req = New System.Windows.Forms.TextBox()
        Me.txtMaterialMasterXMLtoSAPPath_Req = New System.Windows.Forms.TextBox()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.tpInfoRecord = New System.Windows.Forms.TabPage()
        Me.txtIR_XML_ResponseBackupPath = New System.Windows.Forms.TextBox()
        Me.txtIR_XML_ResponsePath = New System.Windows.Forms.TextBox()
        Me.txtIR_XML_RequestBackupPath = New System.Windows.Forms.TextBox()
        Me.txtIR_XML_RequestPath = New System.Windows.Forms.TextBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.tpQuotationMaspis = New System.Windows.Forms.TabPage()
        Me.txtQuotationResponseBackupPath = New System.Windows.Forms.TextBox()
        Me.txtQuotationResponsePath = New System.Windows.Forms.TextBox()
        Me.txtQuotationRequestPath = New System.Windows.Forms.TextBox()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.Label45 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtInfoRecordMaspis = New System.Windows.Forms.TextBox()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.txtInfoRecordMaspisResp = New System.Windows.Forms.TextBox()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.txtInfoRecordMaspisRespBackup = New System.Windows.Forms.TextBox()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.TabEmailSettings.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPathFolders.SuspendLayout()
        Me.tpEPIC1SM.SuspendLayout()
        Me.tpESSBudget.SuspendLayout()
        Me.tpSAP.SuspendLayout()
        Me.tpIAPrice.SuspendLayout()
        Me.tpSAPCostMonitoring.SuspendLayout()
        Me.tpPurchaseRequest.SuspendLayout()
        Me.tpMaterialMaster.SuspendLayout()
        Me.tpInfoRecord.SuspendLayout()
        Me.tpQuotationMaspis.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(622, 511)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(82, 27)
        Me.btnSave.TabIndex = 21
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TabEmailSettings)
        Me.GroupBox1.Location = New System.Drawing.Point(28, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(676, 196)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "SMTP Settings"
        '
        'TabEmailSettings
        '
        Me.TabEmailSettings.Controls.Add(Me.TabPage1)
        Me.TabEmailSettings.Location = New System.Drawing.Point(6, 20)
        Me.TabEmailSettings.Name = "TabEmailSettings"
        Me.TabEmailSettings.SelectedIndex = 0
        Me.TabEmailSettings.Size = New System.Drawing.Size(664, 164)
        Me.TabEmailSettings.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.Label27)
        Me.TabPage1.Controls.Add(Me.txtSubject_DMMS_PR)
        Me.TabPage1.Controls.Add(Me.txtEmailTo_DMMS_PR)
        Me.TabPage1.Controls.Add(Me.Label26)
        Me.TabPage1.Controls.Add(Me.txtEmailFrom_DMMS_PR)
        Me.TabPage1.Controls.Add(Me.Label25)
        Me.TabPage1.Controls.Add(Me.txtSMTPPort_DMMS_PR)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.txtSMTPAddress_DMMS_PR)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.txtSMTPPassword_DMMS_PR)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.txtSMTPUser_DMMS_PR)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(656, 138)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "DMMS - PR"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(288, 101)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(43, 13)
        Me.Label27.TabIndex = 30
        Me.Label27.Text = "Subject"
        '
        'txtSubject_DMMS_PR
        '
        Me.txtSubject_DMMS_PR.Location = New System.Drawing.Point(352, 98)
        Me.txtSubject_DMMS_PR.MaxLength = 4000
        Me.txtSubject_DMMS_PR.Name = "txtSubject_DMMS_PR"
        Me.txtSubject_DMMS_PR.Size = New System.Drawing.Size(289, 21)
        Me.txtSubject_DMMS_PR.TabIndex = 7
        '
        'txtEmailTo_DMMS_PR
        '
        Me.txtEmailTo_DMMS_PR.Location = New System.Drawing.Point(352, 44)
        Me.txtEmailTo_DMMS_PR.MaxLength = 4000
        Me.txtEmailTo_DMMS_PR.Multiline = True
        Me.txtEmailTo_DMMS_PR.Name = "txtEmailTo_DMMS_PR"
        Me.txtEmailTo_DMMS_PR.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtEmailTo_DMMS_PR.Size = New System.Drawing.Size(289, 48)
        Me.txtEmailTo_DMMS_PR.TabIndex = 6
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(288, 47)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(46, 13)
        Me.Label26.TabIndex = 28
        Me.Label26.Text = "Email To"
        '
        'txtEmailFrom_DMMS_PR
        '
        Me.txtEmailFrom_DMMS_PR.Location = New System.Drawing.Point(352, 17)
        Me.txtEmailFrom_DMMS_PR.MaxLength = 4000
        Me.txtEmailFrom_DMMS_PR.Name = "txtEmailFrom_DMMS_PR"
        Me.txtEmailFrom_DMMS_PR.Size = New System.Drawing.Size(173, 21)
        Me.txtEmailFrom_DMMS_PR.TabIndex = 5
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(288, 20)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(58, 13)
        Me.Label25.TabIndex = 26
        Me.Label25.Text = "Email From"
        '
        'txtSMTPPort_DMMS_PR
        '
        Me.txtSMTPPort_DMMS_PR.Location = New System.Drawing.Point(99, 44)
        Me.txtSMTPPort_DMMS_PR.MaxLength = 10
        Me.txtSMTPPort_DMMS_PR.Name = "txtSMTPPort_DMMS_PR"
        Me.txtSMTPPort_DMMS_PR.Size = New System.Drawing.Size(77, 21)
        Me.txtSMTPPort_DMMS_PR.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(11, 47)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 13)
        Me.Label5.TabIndex = 24
        Me.Label5.Text = "SMTP Port"
        '
        'txtSMTPAddress_DMMS_PR
        '
        Me.txtSMTPAddress_DMMS_PR.Location = New System.Drawing.Point(99, 17)
        Me.txtSMTPAddress_DMMS_PR.MaxLength = 100
        Me.txtSMTPAddress_DMMS_PR.Name = "txtSMTPAddress_DMMS_PR"
        Me.txtSMTPAddress_DMMS_PR.Size = New System.Drawing.Size(173, 21)
        Me.txtSMTPAddress_DMMS_PR.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(75, 13)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "SMTP Address"
        '
        'txtSMTPPassword_DMMS_PR
        '
        Me.txtSMTPPassword_DMMS_PR.Location = New System.Drawing.Point(99, 98)
        Me.txtSMTPPassword_DMMS_PR.MaxLength = 100
        Me.txtSMTPPassword_DMMS_PR.Name = "txtSMTPPassword_DMMS_PR"
        Me.txtSMTPPassword_DMMS_PR.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSMTPPassword_DMMS_PR.Size = New System.Drawing.Size(173, 21)
        Me.txtSMTPPassword_DMMS_PR.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 101)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 13)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "SMTP Password"
        '
        'txtSMTPUser_DMMS_PR
        '
        Me.txtSMTPUser_DMMS_PR.Location = New System.Drawing.Point(99, 71)
        Me.txtSMTPUser_DMMS_PR.MaxLength = 100
        Me.txtSMTPUser_DMMS_PR.Name = "txtSMTPUser_DMMS_PR"
        Me.txtSMTPUser_DMMS_PR.Size = New System.Drawing.Size(173, 21)
        Me.txtSMTPUser_DMMS_PR.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 74)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 13)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "SMTP User"
        '
        'TabPathFolders
        '
        Me.TabPathFolders.Controls.Add(Me.tpEPIC1SM)
        Me.TabPathFolders.Controls.Add(Me.tpESSBudget)
        Me.TabPathFolders.Controls.Add(Me.tpSAP)
        Me.TabPathFolders.Controls.Add(Me.tpIAPrice)
        Me.TabPathFolders.Controls.Add(Me.tpSAPCostMonitoring)
        Me.TabPathFolders.Controls.Add(Me.tpPurchaseRequest)
        Me.TabPathFolders.Controls.Add(Me.tpMaterialMaster)
        Me.TabPathFolders.Controls.Add(Me.tpInfoRecord)
        Me.TabPathFolders.Controls.Add(Me.tpQuotationMaspis)
        Me.TabPathFolders.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.TabPathFolders.Location = New System.Drawing.Point(6, 20)
        Me.TabPathFolders.Multiline = True
        Me.TabPathFolders.Name = "TabPathFolders"
        Me.TabPathFolders.SelectedIndex = 0
        Me.TabPathFolders.Size = New System.Drawing.Size(645, 238)
        Me.TabPathFolders.TabIndex = 8
        '
        'tpEPIC1SM
        '
        Me.tpEPIC1SM.BackColor = System.Drawing.SystemColors.Control
        Me.tpEPIC1SM.Controls.Add(Me.txtJSONPathBackup)
        Me.tpEPIC1SM.Controls.Add(Me.txtJSONPath)
        Me.tpEPIC1SM.Controls.Add(Me.Label6)
        Me.tpEPIC1SM.Controls.Add(Me.Label3)
        Me.tpEPIC1SM.Location = New System.Drawing.Point(4, 40)
        Me.tpEPIC1SM.Name = "tpEPIC1SM"
        Me.tpEPIC1SM.Padding = New System.Windows.Forms.Padding(3)
        Me.tpEPIC1SM.Size = New System.Drawing.Size(637, 194)
        Me.tpEPIC1SM.TabIndex = 0
        Me.tpEPIC1SM.Text = "EPIC 1 Supplier Management"
        '
        'txtJSONPathBackup
        '
        Me.txtJSONPathBackup.Location = New System.Drawing.Point(215, 91)
        Me.txtJSONPathBackup.MaxLength = 200
        Me.txtJSONPathBackup.Name = "txtJSONPathBackup"
        Me.txtJSONPathBackup.Size = New System.Drawing.Size(322, 21)
        Me.txtJSONPathBackup.TabIndex = 10
        '
        'txtJSONPath
        '
        Me.txtJSONPath.Location = New System.Drawing.Point(215, 64)
        Me.txtJSONPath.MaxLength = 200
        Me.txtJSONPath.Name = "txtJSONPath"
        Me.txtJSONPath.Size = New System.Drawing.Size(322, 21)
        Me.txtJSONPath.TabIndex = 9
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(92, 95)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(99, 13)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Backup Path Folder"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(92, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 13)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "Path Folder"
        '
        'tpESSBudget
        '
        Me.tpESSBudget.BackColor = System.Drawing.SystemColors.Control
        Me.tpESSBudget.Controls.Add(Me.txtESSBudgetPathBackup)
        Me.tpESSBudget.Controls.Add(Me.txtESSBudgetPath)
        Me.tpESSBudget.Controls.Add(Me.Label7)
        Me.tpESSBudget.Controls.Add(Me.Label8)
        Me.tpESSBudget.Location = New System.Drawing.Point(4, 40)
        Me.tpESSBudget.Name = "tpESSBudget"
        Me.tpESSBudget.Padding = New System.Windows.Forms.Padding(3)
        Me.tpESSBudget.Size = New System.Drawing.Size(637, 194)
        Me.tpESSBudget.TabIndex = 1
        Me.tpESSBudget.Text = "ESS Budget"
        '
        'txtESSBudgetPathBackup
        '
        Me.txtESSBudgetPathBackup.Location = New System.Drawing.Point(215, 91)
        Me.txtESSBudgetPathBackup.MaxLength = 200
        Me.txtESSBudgetPathBackup.Name = "txtESSBudgetPathBackup"
        Me.txtESSBudgetPathBackup.Size = New System.Drawing.Size(322, 21)
        Me.txtESSBudgetPathBackup.TabIndex = 12
        '
        'txtESSBudgetPath
        '
        Me.txtESSBudgetPath.Location = New System.Drawing.Point(215, 64)
        Me.txtESSBudgetPath.MaxLength = 200
        Me.txtESSBudgetPath.Name = "txtESSBudgetPath"
        Me.txtESSBudgetPath.Size = New System.Drawing.Size(322, 21)
        Me.txtESSBudgetPath.TabIndex = 11
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(92, 95)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(99, 13)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Backup Path Folder"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(92, 68)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(62, 13)
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "Path Folder"
        '
        'tpSAP
        '
        Me.tpSAP.BackColor = System.Drawing.SystemColors.Control
        Me.tpSAP.Controls.Add(Me.txtSAPFilePathBackup)
        Me.tpSAP.Controls.Add(Me.txtSAPFilePath)
        Me.tpSAP.Controls.Add(Me.Label9)
        Me.tpSAP.Controls.Add(Me.Label10)
        Me.tpSAP.Location = New System.Drawing.Point(4, 40)
        Me.tpSAP.Name = "tpSAP"
        Me.tpSAP.Padding = New System.Windows.Forms.Padding(3)
        Me.tpSAP.Size = New System.Drawing.Size(637, 194)
        Me.tpSAP.TabIndex = 2
        Me.tpSAP.Text = "SAP - PO & GR"
        '
        'txtSAPFilePathBackup
        '
        Me.txtSAPFilePathBackup.Location = New System.Drawing.Point(215, 91)
        Me.txtSAPFilePathBackup.MaxLength = 200
        Me.txtSAPFilePathBackup.Name = "txtSAPFilePathBackup"
        Me.txtSAPFilePathBackup.Size = New System.Drawing.Size(322, 21)
        Me.txtSAPFilePathBackup.TabIndex = 14
        '
        'txtSAPFilePath
        '
        Me.txtSAPFilePath.Location = New System.Drawing.Point(215, 64)
        Me.txtSAPFilePath.MaxLength = 200
        Me.txtSAPFilePath.Name = "txtSAPFilePath"
        Me.txtSAPFilePath.Size = New System.Drawing.Size(322, 21)
        Me.txtSAPFilePath.TabIndex = 13
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(92, 95)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(99, 13)
        Me.Label9.TabIndex = 27
        Me.Label9.Text = "Backup Path Folder"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(92, 68)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(62, 13)
        Me.Label10.TabIndex = 26
        Me.Label10.Text = "Path Folder"
        '
        'tpIAPrice
        '
        Me.tpIAPrice.BackColor = System.Drawing.SystemColors.Control
        Me.tpIAPrice.Controls.Add(Me.txtIAPriceDestinationPath)
        Me.tpIAPrice.Controls.Add(Me.Label11)
        Me.tpIAPrice.Location = New System.Drawing.Point(4, 40)
        Me.tpIAPrice.Name = "tpIAPrice"
        Me.tpIAPrice.Padding = New System.Windows.Forms.Padding(3)
        Me.tpIAPrice.Size = New System.Drawing.Size(637, 194)
        Me.tpIAPrice.TabIndex = 3
        Me.tpIAPrice.Text = "IA Price Destination"
        '
        'txtIAPriceDestinationPath
        '
        Me.txtIAPriceDestinationPath.Location = New System.Drawing.Point(215, 64)
        Me.txtIAPriceDestinationPath.MaxLength = 200
        Me.txtIAPriceDestinationPath.Name = "txtIAPriceDestinationPath"
        Me.txtIAPriceDestinationPath.Size = New System.Drawing.Size(322, 21)
        Me.txtIAPriceDestinationPath.TabIndex = 15
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(92, 68)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(62, 13)
        Me.Label11.TabIndex = 27
        Me.Label11.Text = "Path Folder"
        '
        'tpSAPCostMonitoring
        '
        Me.tpSAPCostMonitoring.BackColor = System.Drawing.SystemColors.Control
        Me.tpSAPCostMonitoring.Controls.Add(Me.Label20)
        Me.tpSAPCostMonitoring.Controls.Add(Me.Label19)
        Me.tpSAPCostMonitoring.Controls.Add(Me.txtIAMIPressPartLocalComponentPathBackup)
        Me.tpSAPCostMonitoring.Controls.Add(Me.txtIAMIPressPartLocalComponentPath)
        Me.tpSAPCostMonitoring.Controls.Add(Me.Label13)
        Me.tpSAPCostMonitoring.Controls.Add(Me.Label12)
        Me.tpSAPCostMonitoring.Controls.Add(Me.txtIAMIPressPartLocalComponentXMLToSAPPath)
        Me.tpSAPCostMonitoring.Controls.Add(Me.Label14)
        Me.tpSAPCostMonitoring.Controls.Add(Me.Label15)
        Me.tpSAPCostMonitoring.Controls.Add(Me.txtCostHistoryXMLToSAPPath)
        Me.tpSAPCostMonitoring.Location = New System.Drawing.Point(4, 40)
        Me.tpSAPCostMonitoring.Name = "tpSAPCostMonitoring"
        Me.tpSAPCostMonitoring.Padding = New System.Windows.Forms.Padding(3)
        Me.tpSAPCostMonitoring.Size = New System.Drawing.Size(637, 194)
        Me.tpSAPCostMonitoring.TabIndex = 6
        Me.tpSAPCostMonitoring.Text = "SAP - Cost Monitoring"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.Location = New System.Drawing.Point(23, 121)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(76, 13)
        Me.Label20.TabIndex = 45
        Me.Label20.Text = "Cost History"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(23, 13)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(208, 13)
        Me.Label19.TabIndex = 44
        Me.Label19.Text = "IAMI Press Part && Local Component"
        '
        'txtIAMIPressPartLocalComponentPathBackup
        '
        Me.txtIAMIPressPartLocalComponentPathBackup.Location = New System.Drawing.Point(232, 91)
        Me.txtIAMIPressPartLocalComponentPathBackup.MaxLength = 200
        Me.txtIAMIPressPartLocalComponentPathBackup.Name = "txtIAMIPressPartLocalComponentPathBackup"
        Me.txtIAMIPressPartLocalComponentPathBackup.Size = New System.Drawing.Size(322, 21)
        Me.txtIAMIPressPartLocalComponentPathBackup.TabIndex = 18
        '
        'txtIAMIPressPartLocalComponentPath
        '
        Me.txtIAMIPressPartLocalComponentPath.Location = New System.Drawing.Point(232, 64)
        Me.txtIAMIPressPartLocalComponentPath.MaxLength = 200
        Me.txtIAMIPressPartLocalComponentPath.Name = "txtIAMIPressPartLocalComponentPath"
        Me.txtIAMIPressPartLocalComponentPath.Size = New System.Drawing.Size(322, 21)
        Me.txtIAMIPressPartLocalComponentPath.TabIndex = 17
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(23, 94)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(119, 13)
        Me.Label13.TabIndex = 43
        Me.Label13.Text = "Path Folder Backup - In"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(23, 67)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(82, 13)
        Me.Label12.TabIndex = 42
        Me.Label12.Text = "Path Folder - In"
        '
        'txtIAMIPressPartLocalComponentXMLToSAPPath
        '
        Me.txtIAMIPressPartLocalComponentXMLToSAPPath.Location = New System.Drawing.Point(232, 37)
        Me.txtIAMIPressPartLocalComponentXMLToSAPPath.MaxLength = 200
        Me.txtIAMIPressPartLocalComponentXMLToSAPPath.Name = "txtIAMIPressPartLocalComponentXMLToSAPPath"
        Me.txtIAMIPressPartLocalComponentXMLToSAPPath.Size = New System.Drawing.Size(322, 21)
        Me.txtIAMIPressPartLocalComponentXMLToSAPPath.TabIndex = 16
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(23, 40)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(90, 13)
        Me.Label14.TabIndex = 39
        Me.Label14.Text = "Path Folder - Out"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(23, 148)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(90, 13)
        Me.Label15.TabIndex = 37
        Me.Label15.Text = "Path Folder - Out"
        '
        'txtCostHistoryXMLToSAPPath
        '
        Me.txtCostHistoryXMLToSAPPath.Location = New System.Drawing.Point(232, 145)
        Me.txtCostHistoryXMLToSAPPath.MaxLength = 200
        Me.txtCostHistoryXMLToSAPPath.Name = "txtCostHistoryXMLToSAPPath"
        Me.txtCostHistoryXMLToSAPPath.Size = New System.Drawing.Size(322, 21)
        Me.txtCostHistoryXMLToSAPPath.TabIndex = 19
        '
        'tpPurchaseRequest
        '
        Me.tpPurchaseRequest.BackColor = System.Drawing.SystemColors.Control
        Me.tpPurchaseRequest.Controls.Add(Me.Label24)
        Me.tpPurchaseRequest.Controls.Add(Me.txtPurchaseRequestPathOut)
        Me.tpPurchaseRequest.Controls.Add(Me.Label18)
        Me.tpPurchaseRequest.Controls.Add(Me.txtPurchaseRequestPathFailed)
        Me.tpPurchaseRequest.Controls.Add(Me.Label23)
        Me.tpPurchaseRequest.Controls.Add(Me.txtGoodsReceiveXMLToDMMSPath)
        Me.tpPurchaseRequest.Controls.Add(Me.Label22)
        Me.tpPurchaseRequest.Controls.Add(Me.Label21)
        Me.tpPurchaseRequest.Controls.Add(Me.txtPurchaseRequestPathBackup)
        Me.tpPurchaseRequest.Controls.Add(Me.txtPurchaseRequestPath)
        Me.tpPurchaseRequest.Controls.Add(Me.Label16)
        Me.tpPurchaseRequest.Controls.Add(Me.Label17)
        Me.tpPurchaseRequest.Location = New System.Drawing.Point(4, 40)
        Me.tpPurchaseRequest.Name = "tpPurchaseRequest"
        Me.tpPurchaseRequest.Padding = New System.Windows.Forms.Padding(3)
        Me.tpPurchaseRequest.Size = New System.Drawing.Size(637, 194)
        Me.tpPurchaseRequest.TabIndex = 7
        Me.tpPurchaseRequest.Text = "DMMS - PR & GR"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(23, 121)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(140, 13)
        Me.Label24.TabIndex = 52
        Me.Label24.Text = "Path Folder Response - Out"
        '
        'txtPurchaseRequestPathOut
        '
        Me.txtPurchaseRequestPathOut.Location = New System.Drawing.Point(232, 118)
        Me.txtPurchaseRequestPathOut.MaxLength = 200
        Me.txtPurchaseRequestPathOut.Name = "txtPurchaseRequestPathOut"
        Me.txtPurchaseRequestPathOut.Size = New System.Drawing.Size(322, 21)
        Me.txtPurchaseRequestPathOut.TabIndex = 23
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(23, 94)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(113, 13)
        Me.Label18.TabIndex = 50
        Me.Label18.Text = "Path Folder Failed - In"
        '
        'txtPurchaseRequestPathFailed
        '
        Me.txtPurchaseRequestPathFailed.Location = New System.Drawing.Point(232, 91)
        Me.txtPurchaseRequestPathFailed.MaxLength = 200
        Me.txtPurchaseRequestPathFailed.Name = "txtPurchaseRequestPathFailed"
        Me.txtPurchaseRequestPathFailed.Size = New System.Drawing.Size(322, 21)
        Me.txtPurchaseRequestPathFailed.TabIndex = 22
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(23, 177)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(90, 13)
        Me.Label23.TabIndex = 48
        Me.Label23.Text = "Path Folder - Out"
        '
        'txtGoodsReceiveXMLToDMMSPath
        '
        Me.txtGoodsReceiveXMLToDMMSPath.Location = New System.Drawing.Point(232, 174)
        Me.txtGoodsReceiveXMLToDMMSPath.MaxLength = 200
        Me.txtGoodsReceiveXMLToDMMSPath.Name = "txtGoodsReceiveXMLToDMMSPath"
        Me.txtGoodsReceiveXMLToDMMSPath.Size = New System.Drawing.Size(322, 21)
        Me.txtGoodsReceiveXMLToDMMSPath.TabIndex = 24
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(23, 150)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(90, 13)
        Me.Label22.TabIndex = 46
        Me.Label22.Text = "Goods Receive"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.Location = New System.Drawing.Point(23, 13)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(109, 13)
        Me.Label21.TabIndex = 45
        Me.Label21.Text = "Purchase Request"
        '
        'txtPurchaseRequestPathBackup
        '
        Me.txtPurchaseRequestPathBackup.Location = New System.Drawing.Point(232, 64)
        Me.txtPurchaseRequestPathBackup.MaxLength = 200
        Me.txtPurchaseRequestPathBackup.Name = "txtPurchaseRequestPathBackup"
        Me.txtPurchaseRequestPathBackup.Size = New System.Drawing.Size(322, 21)
        Me.txtPurchaseRequestPathBackup.TabIndex = 21
        '
        'txtPurchaseRequestPath
        '
        Me.txtPurchaseRequestPath.Location = New System.Drawing.Point(232, 37)
        Me.txtPurchaseRequestPath.MaxLength = 200
        Me.txtPurchaseRequestPath.Name = "txtPurchaseRequestPath"
        Me.txtPurchaseRequestPath.Size = New System.Drawing.Size(322, 21)
        Me.txtPurchaseRequestPath.TabIndex = 20
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(23, 67)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(119, 13)
        Me.Label16.TabIndex = 39
        Me.Label16.Text = "Path Folder Backup - In"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(23, 40)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(82, 13)
        Me.Label17.TabIndex = 38
        Me.Label17.Text = "Path Folder - In"
        '
        'tpMaterialMaster
        '
        Me.tpMaterialMaster.BackColor = System.Drawing.SystemColors.Control
        Me.tpMaterialMaster.Controls.Add(Me.txtMaterialMasterXMLtoSAPBackupPath_Res)
        Me.tpMaterialMaster.Controls.Add(Me.txtMaterialMasterXMLtoSAPPath_Res)
        Me.tpMaterialMaster.Controls.Add(Me.txtMaterialMasterXMLtoSAPBackupPath_Req)
        Me.tpMaterialMaster.Controls.Add(Me.txtMaterialMasterXMLtoSAPPath_Req)
        Me.tpMaterialMaster.Controls.Add(Me.Label33)
        Me.tpMaterialMaster.Controls.Add(Me.Label32)
        Me.tpMaterialMaster.Controls.Add(Me.Label31)
        Me.tpMaterialMaster.Controls.Add(Me.Label30)
        Me.tpMaterialMaster.Controls.Add(Me.Label29)
        Me.tpMaterialMaster.Controls.Add(Me.Label28)
        Me.tpMaterialMaster.Location = New System.Drawing.Point(4, 40)
        Me.tpMaterialMaster.Name = "tpMaterialMaster"
        Me.tpMaterialMaster.Padding = New System.Windows.Forms.Padding(3)
        Me.tpMaterialMaster.Size = New System.Drawing.Size(637, 194)
        Me.tpMaterialMaster.TabIndex = 8
        Me.tpMaterialMaster.Text = "Material Master"
        '
        'txtMaterialMasterXMLtoSAPBackupPath_Res
        '
        Me.txtMaterialMasterXMLtoSAPBackupPath_Res.Location = New System.Drawing.Point(328, 156)
        Me.txtMaterialMasterXMLtoSAPBackupPath_Res.Name = "txtMaterialMasterXMLtoSAPBackupPath_Res"
        Me.txtMaterialMasterXMLtoSAPBackupPath_Res.Size = New System.Drawing.Size(286, 21)
        Me.txtMaterialMasterXMLtoSAPBackupPath_Res.TabIndex = 47
        '
        'txtMaterialMasterXMLtoSAPPath_Res
        '
        Me.txtMaterialMasterXMLtoSAPPath_Res.Location = New System.Drawing.Point(328, 125)
        Me.txtMaterialMasterXMLtoSAPPath_Res.Name = "txtMaterialMasterXMLtoSAPPath_Res"
        Me.txtMaterialMasterXMLtoSAPPath_Res.Size = New System.Drawing.Size(286, 21)
        Me.txtMaterialMasterXMLtoSAPPath_Res.TabIndex = 46
        '
        'txtMaterialMasterXMLtoSAPBackupPath_Req
        '
        Me.txtMaterialMasterXMLtoSAPBackupPath_Req.Location = New System.Drawing.Point(328, 67)
        Me.txtMaterialMasterXMLtoSAPBackupPath_Req.Name = "txtMaterialMasterXMLtoSAPBackupPath_Req"
        Me.txtMaterialMasterXMLtoSAPBackupPath_Req.Size = New System.Drawing.Size(286, 21)
        Me.txtMaterialMasterXMLtoSAPBackupPath_Req.TabIndex = 45
        '
        'txtMaterialMasterXMLtoSAPPath_Req
        '
        Me.txtMaterialMasterXMLtoSAPPath_Req.Location = New System.Drawing.Point(328, 38)
        Me.txtMaterialMasterXMLtoSAPPath_Req.Name = "txtMaterialMasterXMLtoSAPPath_Req"
        Me.txtMaterialMasterXMLtoSAPPath_Req.Size = New System.Drawing.Size(286, 21)
        Me.txtMaterialMasterXMLtoSAPPath_Req.TabIndex = 44
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(24, 155)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(204, 13)
        Me.Label33.TabIndex = 43
        Me.Label33.Text = "Material XML to SAP - Backup Path Folder"
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(24, 128)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(167, 13)
        Me.Label32.TabIndex = 41
        Me.Label32.Text = "Material XML to SAP - Path Folder"
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(23, 67)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(204, 13)
        Me.Label31.TabIndex = 38
        Me.Label31.Text = "Material XML to SAP - Backup Path Folder"
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(23, 40)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(167, 13)
        Me.Label30.TabIndex = 37
        Me.Label30.Text = "Material XML to SAP - Path Folder"
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.Location = New System.Drawing.Point(24, 101)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(62, 13)
        Me.Label29.TabIndex = 1
        Me.Label29.Text = "Response"
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.Location = New System.Drawing.Point(23, 13)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(54, 13)
        Me.Label28.TabIndex = 0
        Me.Label28.Text = "Request"
        '
        'tpInfoRecord
        '
        Me.tpInfoRecord.BackColor = System.Drawing.SystemColors.Control
        Me.tpInfoRecord.Controls.Add(Me.txtIR_XML_ResponseBackupPath)
        Me.tpInfoRecord.Controls.Add(Me.txtIR_XML_ResponsePath)
        Me.tpInfoRecord.Controls.Add(Me.txtIR_XML_RequestBackupPath)
        Me.tpInfoRecord.Controls.Add(Me.txtIR_XML_RequestPath)
        Me.tpInfoRecord.Controls.Add(Me.Label34)
        Me.tpInfoRecord.Controls.Add(Me.Label35)
        Me.tpInfoRecord.Controls.Add(Me.Label36)
        Me.tpInfoRecord.Controls.Add(Me.Label37)
        Me.tpInfoRecord.Controls.Add(Me.Label38)
        Me.tpInfoRecord.Controls.Add(Me.Label39)
        Me.tpInfoRecord.Location = New System.Drawing.Point(4, 40)
        Me.tpInfoRecord.Name = "tpInfoRecord"
        Me.tpInfoRecord.Padding = New System.Windows.Forms.Padding(3)
        Me.tpInfoRecord.Size = New System.Drawing.Size(637, 194)
        Me.tpInfoRecord.TabIndex = 9
        Me.tpInfoRecord.Text = "Info Record"
        '
        'txtIR_XML_ResponseBackupPath
        '
        Me.txtIR_XML_ResponseBackupPath.Location = New System.Drawing.Point(328, 156)
        Me.txtIR_XML_ResponseBackupPath.Name = "txtIR_XML_ResponseBackupPath"
        Me.txtIR_XML_ResponseBackupPath.Size = New System.Drawing.Size(286, 21)
        Me.txtIR_XML_ResponseBackupPath.TabIndex = 57
        '
        'txtIR_XML_ResponsePath
        '
        Me.txtIR_XML_ResponsePath.Location = New System.Drawing.Point(328, 125)
        Me.txtIR_XML_ResponsePath.Name = "txtIR_XML_ResponsePath"
        Me.txtIR_XML_ResponsePath.Size = New System.Drawing.Size(286, 21)
        Me.txtIR_XML_ResponsePath.TabIndex = 56
        '
        'txtIR_XML_RequestBackupPath
        '
        Me.txtIR_XML_RequestBackupPath.Location = New System.Drawing.Point(328, 67)
        Me.txtIR_XML_RequestBackupPath.Name = "txtIR_XML_RequestBackupPath"
        Me.txtIR_XML_RequestBackupPath.Size = New System.Drawing.Size(286, 21)
        Me.txtIR_XML_RequestBackupPath.TabIndex = 55
        '
        'txtIR_XML_RequestPath
        '
        Me.txtIR_XML_RequestPath.Location = New System.Drawing.Point(328, 38)
        Me.txtIR_XML_RequestPath.Name = "txtIR_XML_RequestPath"
        Me.txtIR_XML_RequestPath.Size = New System.Drawing.Size(286, 21)
        Me.txtIR_XML_RequestPath.TabIndex = 54
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Location = New System.Drawing.Point(24, 155)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(223, 13)
        Me.Label34.TabIndex = 53
        Me.Label34.Text = "Info Record XML to SAP - Backup Path Folder"
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Location = New System.Drawing.Point(24, 128)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(186, 13)
        Me.Label35.TabIndex = 52
        Me.Label35.Text = "Info Record XML to SAP - Path Folder"
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Location = New System.Drawing.Point(23, 67)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(223, 13)
        Me.Label36.TabIndex = 51
        Me.Label36.Text = "Info Record XML to SAP - Backup Path Folder"
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Location = New System.Drawing.Point(23, 40)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(186, 13)
        Me.Label37.TabIndex = 50
        Me.Label37.Text = "Info Record XML to SAP - Path Folder"
        '
        'Label38
        '
        Me.Label38.AutoSize = True
        Me.Label38.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.Location = New System.Drawing.Point(24, 101)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(62, 13)
        Me.Label38.TabIndex = 49
        Me.Label38.Text = "Response"
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label39.Location = New System.Drawing.Point(23, 13)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(54, 13)
        Me.Label39.TabIndex = 48
        Me.Label39.Text = "Request"
        '
        'tpQuotationMaspis
        '
        Me.tpQuotationMaspis.AutoScroll = True
        Me.tpQuotationMaspis.BackColor = System.Drawing.SystemColors.Control
        Me.tpQuotationMaspis.Controls.Add(Me.txtInfoRecordMaspisRespBackup)
        Me.tpQuotationMaspis.Controls.Add(Me.Label47)
        Me.tpQuotationMaspis.Controls.Add(Me.txtInfoRecordMaspisResp)
        Me.tpQuotationMaspis.Controls.Add(Me.Label46)
        Me.tpQuotationMaspis.Controls.Add(Me.txtInfoRecordMaspis)
        Me.tpQuotationMaspis.Controls.Add(Me.Label42)
        Me.tpQuotationMaspis.Controls.Add(Me.txtQuotationResponseBackupPath)
        Me.tpQuotationMaspis.Controls.Add(Me.txtQuotationResponsePath)
        Me.tpQuotationMaspis.Controls.Add(Me.txtQuotationRequestPath)
        Me.tpQuotationMaspis.Controls.Add(Me.Label40)
        Me.tpQuotationMaspis.Controls.Add(Me.Label41)
        Me.tpQuotationMaspis.Controls.Add(Me.Label43)
        Me.tpQuotationMaspis.Controls.Add(Me.Label44)
        Me.tpQuotationMaspis.Controls.Add(Me.Label45)
        Me.tpQuotationMaspis.Location = New System.Drawing.Point(4, 40)
        Me.tpQuotationMaspis.Name = "tpQuotationMaspis"
        Me.tpQuotationMaspis.Padding = New System.Windows.Forms.Padding(3)
        Me.tpQuotationMaspis.Size = New System.Drawing.Size(637, 194)
        Me.tpQuotationMaspis.TabIndex = 10
        Me.tpQuotationMaspis.Text = "MASPIS"
        '
        'txtQuotationResponseBackupPath
        '
        Me.txtQuotationResponseBackupPath.Location = New System.Drawing.Point(328, 148)
        Me.txtQuotationResponseBackupPath.Name = "txtQuotationResponseBackupPath"
        Me.txtQuotationResponseBackupPath.Size = New System.Drawing.Size(286, 21)
        Me.txtQuotationResponseBackupPath.TabIndex = 67
        '
        'txtQuotationResponsePath
        '
        Me.txtQuotationResponsePath.Location = New System.Drawing.Point(328, 119)
        Me.txtQuotationResponsePath.Name = "txtQuotationResponsePath"
        Me.txtQuotationResponsePath.Size = New System.Drawing.Size(286, 21)
        Me.txtQuotationResponsePath.TabIndex = 66
        '
        'txtQuotationRequestPath
        '
        Me.txtQuotationRequestPath.Location = New System.Drawing.Point(328, 40)
        Me.txtQuotationRequestPath.Name = "txtQuotationRequestPath"
        Me.txtQuotationRequestPath.Size = New System.Drawing.Size(286, 21)
        Me.txtQuotationRequestPath.TabIndex = 64
        '
        'Label40
        '
        Me.Label40.AutoSize = True
        Me.Label40.Location = New System.Drawing.Point(23, 148)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(210, 13)
        Me.Label40.TabIndex = 63
        Me.Label40.Text = "Quotation Response -  Backup Path Folder"
        '
        'Label41
        '
        Me.Label41.AutoSize = True
        Me.Label41.Location = New System.Drawing.Point(23, 122)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(170, 13)
        Me.Label41.TabIndex = 62
        Me.Label41.Text = "Quotation Response - Path Folder"
        '
        'Label43
        '
        Me.Label43.AutoSize = True
        Me.Label43.Location = New System.Drawing.Point(23, 42)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(156, 13)
        Me.Label43.TabIndex = 60
        Me.Label43.Text = "Quotation Request Path Folder"
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label44.Location = New System.Drawing.Point(23, 94)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(62, 13)
        Me.Label44.TabIndex = 59
        Me.Label44.Text = "Response"
        '
        'Label45
        '
        Me.Label45.AutoSize = True
        Me.Label45.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label45.Location = New System.Drawing.Point(23, 15)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(54, 13)
        Me.Label45.TabIndex = 58
        Me.Label45.Text = "Request"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.TabPathFolders)
        Me.GroupBox2.Location = New System.Drawing.Point(28, 214)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(676, 291)
        Me.GroupBox2.TabIndex = 11
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Interface Settings"
        '
        'txtInfoRecordMaspis
        '
        Me.txtInfoRecordMaspis.Location = New System.Drawing.Point(328, 67)
        Me.txtInfoRecordMaspis.Name = "txtInfoRecordMaspis"
        Me.txtInfoRecordMaspis.Size = New System.Drawing.Size(286, 21)
        Me.txtInfoRecordMaspis.TabIndex = 69
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(23, 69)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(122, 13)
        Me.Label42.TabIndex = 68
        Me.Label42.Text = "Info Record Path Folder"
        '
        'txtInfoRecordMaspisResp
        '
        Me.txtInfoRecordMaspisResp.Location = New System.Drawing.Point(328, 178)
        Me.txtInfoRecordMaspisResp.Name = "txtInfoRecordMaspisResp"
        Me.txtInfoRecordMaspisResp.Size = New System.Drawing.Size(286, 21)
        Me.txtInfoRecordMaspisResp.TabIndex = 71
        '
        'Label46
        '
        Me.Label46.AutoSize = True
        Me.Label46.Location = New System.Drawing.Point(23, 178)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(132, 13)
        Me.Label46.TabIndex = 70
        Me.Label46.Text = "Info Record -  Path Folder"
        '
        'txtInfoRecordMaspisRespBackup
        '
        Me.txtInfoRecordMaspisRespBackup.Location = New System.Drawing.Point(328, 208)
        Me.txtInfoRecordMaspisRespBackup.Name = "txtInfoRecordMaspisRespBackup"
        Me.txtInfoRecordMaspisRespBackup.Size = New System.Drawing.Size(286, 21)
        Me.txtInfoRecordMaspisRespBackup.TabIndex = 73
        '
        'Label47
        '
        Me.Label47.AutoSize = True
        Me.Label47.Location = New System.Drawing.Point(23, 208)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(169, 13)
        Me.Label47.TabIndex = 72
        Me.Label47.Text = "Info Record -  Backup Path Folder"
        '
        'frmSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(739, 556)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnSave)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSetting"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Setting"
        Me.GroupBox1.ResumeLayout(False)
        Me.TabEmailSettings.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPathFolders.ResumeLayout(False)
        Me.tpEPIC1SM.ResumeLayout(False)
        Me.tpEPIC1SM.PerformLayout()
        Me.tpESSBudget.ResumeLayout(False)
        Me.tpESSBudget.PerformLayout()
        Me.tpSAP.ResumeLayout(False)
        Me.tpSAP.PerformLayout()
        Me.tpIAPrice.ResumeLayout(False)
        Me.tpIAPrice.PerformLayout()
        Me.tpSAPCostMonitoring.ResumeLayout(False)
        Me.tpSAPCostMonitoring.PerformLayout()
        Me.tpPurchaseRequest.ResumeLayout(False)
        Me.tpPurchaseRequest.PerformLayout()
        Me.tpMaterialMaster.ResumeLayout(False)
        Me.tpMaterialMaster.PerformLayout()
        Me.tpInfoRecord.ResumeLayout(False)
        Me.tpInfoRecord.PerformLayout()
        Me.tpQuotationMaspis.ResumeLayout(False)
        Me.tpQuotationMaspis.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TabPathFolders As System.Windows.Forms.TabControl
    Friend WithEvents tpEPIC1SM As System.Windows.Forms.TabPage
    Friend WithEvents txtJSONPathBackup As System.Windows.Forms.TextBox
    Friend WithEvents txtJSONPath As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tpESSBudget As System.Windows.Forms.TabPage
    Friend WithEvents txtESSBudgetPathBackup As System.Windows.Forms.TextBox
    Friend WithEvents txtESSBudgetPath As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents tpSAP As System.Windows.Forms.TabPage
    Friend WithEvents txtSAPFilePathBackup As System.Windows.Forms.TextBox
    Friend WithEvents txtSAPFilePath As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents tpIAPrice As System.Windows.Forms.TabPage
    Friend WithEvents txtIAPriceDestinationPath As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents tpSAPCostMonitoring As System.Windows.Forms.TabPage
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtCostHistoryXMLToSAPPath As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents tpPurchaseRequest As System.Windows.Forms.TabPage
    Friend WithEvents txtPurchaseRequestPathBackup As System.Windows.Forms.TextBox
    Friend WithEvents txtPurchaseRequestPath As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents txtIAMIPressPartLocalComponentPathBackup As System.Windows.Forms.TextBox
    Friend WithEvents txtIAMIPressPartLocalComponentPath As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtIAMIPressPartLocalComponentXMLToSAPPath As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents txtGoodsReceiveXMLToDMMSPath As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtPurchaseRequestPathFailed As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents txtPurchaseRequestPathOut As System.Windows.Forms.TextBox
    Friend WithEvents TabEmailSettings As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents txtSubject_DMMS_PR As System.Windows.Forms.TextBox
    Friend WithEvents txtEmailTo_DMMS_PR As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents txtEmailFrom_DMMS_PR As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents txtSMTPPort_DMMS_PR As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtSMTPAddress_DMMS_PR As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSMTPPassword_DMMS_PR As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSMTPUser_DMMS_PR As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tpMaterialMaster As System.Windows.Forms.TabPage
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents tpInfoRecord As System.Windows.Forms.TabPage
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents txtMaterialMasterXMLtoSAPBackupPath_Res As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterialMasterXMLtoSAPPath_Res As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterialMasterXMLtoSAPBackupPath_Req As System.Windows.Forms.TextBox
    Friend WithEvents txtMaterialMasterXMLtoSAPPath_Req As System.Windows.Forms.TextBox
    Friend WithEvents txtIR_XML_ResponseBackupPath As System.Windows.Forms.TextBox
    Friend WithEvents txtIR_XML_ResponsePath As System.Windows.Forms.TextBox
    Friend WithEvents txtIR_XML_RequestBackupPath As System.Windows.Forms.TextBox
    Friend WithEvents txtIR_XML_RequestPath As System.Windows.Forms.TextBox
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents tpQuotationMaspis As System.Windows.Forms.TabPage
    Friend WithEvents txtQuotationResponseBackupPath As System.Windows.Forms.TextBox
    Friend WithEvents txtQuotationResponsePath As System.Windows.Forms.TextBox
    Friend WithEvents txtQuotationRequestPath As System.Windows.Forms.TextBox
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents txtInfoRecordMaspisRespBackup As System.Windows.Forms.TextBox
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents txtInfoRecordMaspisResp As System.Windows.Forms.TextBox
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents txtInfoRecordMaspis As System.Windows.Forms.TextBox
    Friend WithEvents Label42 As System.Windows.Forms.Label
End Class
