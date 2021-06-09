<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmProspDataReduction
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProspDataReduction))
        Me.fraDataReduction = New System.Windows.Forms.GroupBox()
        Me.cmdExitForm = New System.Windows.Forms.Button()
        Me.cmdPrtScr = New System.Windows.Forms.Button()
        Me.lblProspectDatasetStatus = New System.Windows.Forms.Label()
        Me.tabMain = New System.Windows.Forms.TabControl()
        Me.tabAreaDef = New System.Windows.Forms.TabPage()
        Me.tabProductSizes = New System.Windows.Forms.TabPage()
        Me.tabRecoveryAndMineability = New System.Windows.Forms.TabPage()
        Me.TabPage9 = New System.Windows.Forms.TabPage()
        Me.fraOffSpecPb = New System.Windows.Forms.GroupBox()
        Me.fraOrigMgoPlant = New System.Windows.Forms.GroupBox()
        Me.ssOffSpecPb = New AxFPSpread.AxvaSpread()
        Me.lblGen31 = New System.Windows.Forms.Label()
        Me.chkUseOrigMgoPlant = New System.Windows.Forms.CheckBox()
        Me.fraDoloflotPlant = New System.Windows.Forms.GroupBox()
        Me.ssDoloflotPlant = New AxFPSpread.AxvaSpread()
        Me.lblGen44 = New System.Windows.Forms.Label()
        Me.lblGen45 = New System.Windows.Forms.Label()
        Me.cmdSetDefaults = New System.Windows.Forms.Button()
        Me.chkUseDoloflotPlant = New System.Windows.Forms.CheckBox()
        Me.fraFcoDoloflot = New System.Windows.Forms.GroupBox()
        Me.ssDoloflotPlantFco2 = New AxFPSpread.AxvaSpread()
        Me.ssDoloflotPlantFco = New AxFPSpread.AxvaSpread()
        Me.lblGen51 = New System.Windows.Forms.Label()
        Me.chkUseDoloflotPlantFco = New System.Windows.Forms.CheckBox()
        Me.cmdSetDefaults2 = New System.Windows.Forms.Button()
        Me.TabPage10 = New System.Windows.Forms.TabPage()
        Me.fraRptAllToTextFile = New System.Windows.Forms.GroupBox()
        Me.lblRptAllCnt2 = New System.Windows.Forms.Label()
        Me.cmdRptAllToTxtFile = New System.Windows.Forms.Button()
        Me.txtRptAllToTxtFile = New System.Windows.Forms.TextBox()
        Me.cmdPrintRptAll = New System.Windows.Forms.Button()
        Me.fraDataSaveComm = New System.Windows.Forms.GroupBox()
        Me.lblGen53 = New System.Windows.Forms.Label()
        Me.fraOutputLocation2 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblTextfileComment = New System.Windows.Forms.Label()
        Me.lblGen10 = New System.Windows.Forms.Label()
        Me.lblGen37 = New System.Windows.Forms.Label()
        Me.txtProspDatasetTextfileName = New System.Windows.Forms.TextBox()
        Me.chkSurvCaddTextfile = New System.Windows.Forms.CheckBox()
        Me.chkSpecMoisTransferFile = New System.Windows.Forms.CheckBox()
        Me.chkInclMgPlt = New System.Windows.Forms.CheckBox()
        Me.chkPbAnalysisFillInSpecial = New System.Windows.Forms.CheckBox()
        Me.chkBdFormatTextfile = New System.Windows.Forms.CheckBox()
        Me.fraOutputLocation1 = New System.Windows.Forms.GroupBox()
        Me.lblGen9 = New System.Windows.Forms.Label()
        Me.lblGen11 = New System.Windows.Forms.Label()
        Me.chkSaveToDatabase = New System.Windows.Forms.CheckBox()
        Me.txtProspectDatasetName = New System.Windows.Forms.TextBox()
        Me.txtProspectDatasetDesc = New System.Windows.Forms.TextBox()
        Me.fraProspDatasetType = New System.Windows.Forms.GroupBox()
        Me.optInclSplits = New System.Windows.Forms.RadioButton()
        Me.optInclComposites = New System.Windows.Forms.RadioButton()
        Me.chkProductionCoefficient = New System.Windows.Forms.CheckBox()
        Me.optInclBoth = New System.Windows.Forms.RadioButton()
        Me.chk100Pct = New System.Windows.Forms.CheckBox()
        Me.fraExtra = New System.Windows.Forms.GroupBox()
        Me.TabPage11 = New System.Windows.Forms.TabPage()
        Me.fraOverride = New System.Windows.Forms.GroupBox()
        Me.fraSplitOverride = New System.Windows.Forms.GroupBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.ssSplitOverrides = New AxFPSpread.AxvaSpread()
        Me.cmdGetSplitOverrides = New System.Windows.Forms.Button()
        Me.chkOnlyMySplitOverride = New System.Windows.Forms.CheckBox()
        Me.cmdSaveSplitOverride = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdCancelSplitOverride = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboSplitOverrideMineName = New System.Windows.Forms.ComboBox()
        Me.cmdDeleteSplitOverride = New System.Windows.Forms.Button()
        Me.lblSplitOverrideTxtFile = New System.Windows.Forms.Label()
        Me.txtSplOverrideTxtFile = New System.Windows.Forms.TextBox()
        Me.cmdLoadOverrideTxtFile = New System.Windows.Forms.Button()
        Me.txtSplitOverrideName = New System.Windows.Forms.TextBox()
        Me.fraOverrideList = New System.Windows.Forms.GroupBox()
        Me.ssSplitOverride = New AxFPSpread.AxvaSpread()
        Me.ssRawProspMin = New AxFPSpread.AxvaSpread()
        Me.lblGen54 = New System.Windows.Forms.Label()
        Me.lblGen55 = New System.Windows.Forms.Label()
        Me.lblGen33 = New System.Windows.Forms.Label()
        Me.lblGen34 = New System.Windows.Forms.Label()
        Me.lblGen36 = New System.Windows.Forms.Label()
        Me.cmdApplySplOverrides = New System.Windows.Forms.Button()
        Me.cmdClrOverride = New System.Windows.Forms.Button()
        Me.chkUseRawProspAsOverride = New System.Windows.Forms.CheckBox()
        Me.tabReview = New System.Windows.Forms.TabPage()
        Me.fraReview = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.fraGeneratedData = New System.Windows.Forms.GroupBox()
        Me.tbcSplitResults = New System.Windows.Forms.TabControl()
        Me.tbSplit = New System.Windows.Forms.TabPage()
        Me.ssSplitReview = New AxFPSpread.AxvaSpread()
        Me.tbcHoleResults = New System.Windows.Forms.TabControl()
        Me.tbHole = New System.Windows.Forms.TabPage()
        Me.ssCompReview = New AxFPSpread.AxvaSpread()
        Me.ssCompErrors = New AxFPSpread.AxvaSpread()
        Me.lblBarrenSplComm = New System.Windows.Forms.Label()
        Me.lblGen23 = New System.Windows.Forms.Label()
        Me.lblGen24 = New System.Windows.Forms.Label()
        Me.lblGen25 = New System.Windows.Forms.Label()
        Me.lblGen26 = New System.Windows.Forms.Label()
        Me.lblNoReview = New System.Windows.Forms.Label()
        Me.cmdCopyToOverrides = New System.Windows.Forms.Button()
        Me.fraDetlDispl = New System.Windows.Forms.GroupBox()
        Me.ssDetlDisp = New AxFPSpread.AxvaSpread()
        Me.lblGen41 = New System.Windows.Forms.Label()
        Me.lblGen64 = New System.Windows.Forms.Label()
        Me.cmdPrtGrdDetlDisp = New System.Windows.Forms.Button()
        Me.fraResultCnt = New System.Windows.Forms.GroupBox()
        Me.ssResultCnt = New AxFPSpread.AxvaSpread()
        Me.lblGen65 = New System.Windows.Forms.Label()
        Me.fraDataTypeOption = New System.Windows.Forms.GroupBox()
        Me.lblRptAllCnt = New System.Windows.Forms.Label()
        Me.optProdCoeff = New System.Windows.Forms.RadioButton()
        Me.opt100Pct = New System.Windows.Forms.RadioButton()
        Me.cmdHoleSplitRpt = New System.Windows.Forms.Button()
        Me.cmdReportAll = New System.Windows.Forms.Button()
        Me.cmdAreaReport = New System.Windows.Forms.Button()
        Me.fraReptDisp = New System.Windows.Forms.GroupBox()
        Me.rtbRept1 = New System.Windows.Forms.RichTextBox()
        Me.cmdExitRept = New System.Windows.Forms.Button()
        Me.cmdPrintRept = New System.Windows.Forms.Button()
        Me.cmdReport = New System.Windows.Forms.Button()
        Me.cmdGenerateProspectDataset = New System.Windows.Forms.Button()
        Me.cmdCancelProspectDataset = New System.Windows.Forms.Button()
        Me.chkCreateOutputOnly = New System.Windows.Forms.CheckBox()
        Me.cmdSaveProspectDataset = New System.Windows.Forms.Button()
        Me.pnlContent = New System.Windows.Forms.Panel()
        Me.sbrMain = New System.Windows.Forms.StatusStrip()
        Me.lblStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblProcComm0 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblProcComm1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblProcComm2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.fraDataReduction.SuspendLayout()
        Me.tabMain.SuspendLayout()
        Me.TabPage9.SuspendLayout()
        Me.fraOffSpecPb.SuspendLayout()
        Me.fraOrigMgoPlant.SuspendLayout()
        CType(Me.ssOffSpecPb, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraDoloflotPlant.SuspendLayout()
        CType(Me.ssDoloflotPlant, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraFcoDoloflot.SuspendLayout()
        CType(Me.ssDoloflotPlantFco2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ssDoloflotPlantFco, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage10.SuspendLayout()
        Me.fraRptAllToTextFile.SuspendLayout()
        Me.fraDataSaveComm.SuspendLayout()
        Me.fraOutputLocation2.SuspendLayout()
        Me.fraOutputLocation1.SuspendLayout()
        Me.fraProspDatasetType.SuspendLayout()
        Me.TabPage11.SuspendLayout()
        Me.fraOverride.SuspendLayout()
        Me.fraSplitOverride.SuspendLayout()
        Me.Frame2.SuspendLayout()
        CType(Me.ssSplitOverrides, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraOverrideList.SuspendLayout()
        CType(Me.ssSplitOverride, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ssRawProspMin, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabReview.SuspendLayout()
        Me.fraReview.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.fraGeneratedData.SuspendLayout()
        Me.tbcSplitResults.SuspendLayout()
        Me.tbSplit.SuspendLayout()
        CType(Me.ssSplitReview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbcHoleResults.SuspendLayout()
        Me.tbHole.SuspendLayout()
        CType(Me.ssCompReview, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ssCompErrors, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraDetlDispl.SuspendLayout()
        CType(Me.ssDetlDisp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraResultCnt.SuspendLayout()
        CType(Me.ssResultCnt, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraDataTypeOption.SuspendLayout()
        Me.fraReptDisp.SuspendLayout()
        Me.pnlContent.SuspendLayout()
        Me.sbrMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraDataReduction
        '
        Me.fraDataReduction.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraDataReduction.Controls.Add(Me.cmdExitForm)
        Me.fraDataReduction.Controls.Add(Me.cmdPrtScr)
        Me.fraDataReduction.Controls.Add(Me.lblProspectDatasetStatus)
        Me.fraDataReduction.Controls.Add(Me.tabMain)
        Me.fraDataReduction.Controls.Add(Me.cmdReport)
        Me.fraDataReduction.Controls.Add(Me.cmdGenerateProspectDataset)
        Me.fraDataReduction.Controls.Add(Me.cmdCancelProspectDataset)
        Me.fraDataReduction.Controls.Add(Me.chkCreateOutputOnly)
        Me.fraDataReduction.Controls.Add(Me.cmdSaveProspectDataset)
        Me.fraDataReduction.Location = New System.Drawing.Point(0, 0)
        Me.fraDataReduction.Margin = New System.Windows.Forms.Padding(0)
        Me.fraDataReduction.Name = "fraDataReduction"
        Me.fraDataReduction.Size = New System.Drawing.Size(1341, 645)
        Me.fraDataReduction.TabIndex = 1
        Me.fraDataReduction.TabStop = False
        '
        'cmdExitForm
        '
        Me.cmdExitForm.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdExitForm.Location = New System.Drawing.Point(1229, 608)
        Me.cmdExitForm.Name = "cmdExitForm"
        Me.cmdExitForm.Size = New System.Drawing.Size(67, 23)
        Me.cmdExitForm.TabIndex = 4
        Me.cmdExitForm.Text = "Exit"
        Me.cmdExitForm.UseVisualStyleBackColor = True
        '
        'cmdPrtScr
        '
        Me.cmdPrtScr.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPrtScr.Location = New System.Drawing.Point(1139, 608)
        Me.cmdPrtScr.Name = "cmdPrtScr"
        Me.cmdPrtScr.Size = New System.Drawing.Size(81, 23)
        Me.cmdPrtScr.TabIndex = 2
        Me.cmdPrtScr.Text = "Print Screen"
        Me.cmdPrtScr.UseVisualStyleBackColor = True
        '
        'lblProspectDatasetStatus
        '
        Me.lblProspectDatasetStatus.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblProspectDatasetStatus.AutoSize = True
        Me.lblProspectDatasetStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProspectDatasetStatus.Location = New System.Drawing.Point(25, 608)
        Me.lblProspectDatasetStatus.Name = "lblProspectDatasetStatus"
        Me.lblProspectDatasetStatus.Size = New System.Drawing.Size(152, 16)
        Me.lblProspectDatasetStatus.TabIndex = 15
        Me.lblProspectDatasetStatus.Text = "Multi-Hole Reduction"
        '
        'tabMain
        '
        Me.tabMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabMain.Controls.Add(Me.tabAreaDef)
        Me.tabMain.Controls.Add(Me.tabProductSizes)
        Me.tabMain.Controls.Add(Me.tabRecoveryAndMineability)
        Me.tabMain.Controls.Add(Me.TabPage9)
        Me.tabMain.Controls.Add(Me.TabPage10)
        Me.tabMain.Controls.Add(Me.TabPage11)
        Me.tabMain.Controls.Add(Me.tabReview)
        Me.tabMain.Location = New System.Drawing.Point(4, 9)
        Me.tabMain.Name = "tabMain"
        Me.tabMain.SelectedIndex = 0
        Me.tabMain.Size = New System.Drawing.Size(1331, 591)
        Me.tabMain.TabIndex = 6
        '
        'tabAreaDef
        '
        Me.tabAreaDef.BackColor = System.Drawing.SystemColors.Control
        Me.tabAreaDef.Location = New System.Drawing.Point(4, 22)
        Me.tabAreaDef.Name = "tabAreaDef"
        Me.tabAreaDef.Padding = New System.Windows.Forms.Padding(3)
        Me.tabAreaDef.Size = New System.Drawing.Size(1323, 565)
        Me.tabAreaDef.TabIndex = 0
        Me.tabAreaDef.Text = "Area Definition"
        '
        'tabProductSizes
        '
        Me.tabProductSizes.BackColor = System.Drawing.SystemColors.Control
        Me.tabProductSizes.Location = New System.Drawing.Point(4, 22)
        Me.tabProductSizes.Name = "tabProductSizes"
        Me.tabProductSizes.Padding = New System.Windows.Forms.Padding(3)
        Me.tabProductSizes.Size = New System.Drawing.Size(1323, 565)
        Me.tabProductSizes.TabIndex = 1
        Me.tabProductSizes.Text = "Product Sizes"
        '
        'tabRecoveryAndMineability
        '
        Me.tabRecoveryAndMineability.BackColor = System.Drawing.SystemColors.Control
        Me.tabRecoveryAndMineability.Location = New System.Drawing.Point(4, 22)
        Me.tabRecoveryAndMineability.Name = "tabRecoveryAndMineability"
        Me.tabRecoveryAndMineability.Size = New System.Drawing.Size(1323, 565)
        Me.tabRecoveryAndMineability.TabIndex = 2
        Me.tabRecoveryAndMineability.Text = "Product Rcvry/Mineability"
        '
        'TabPage9
        '
        Me.TabPage9.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage9.Controls.Add(Me.fraOffSpecPb)
        Me.TabPage9.Location = New System.Drawing.Point(4, 22)
        Me.TabPage9.Name = "TabPage9"
        Me.TabPage9.Size = New System.Drawing.Size(1323, 565)
        Me.TabPage9.TabIndex = 8
        Me.TabPage9.Text = "Off-Spec Pb (3)"
        '
        'fraOffSpecPb
        '
        Me.fraOffSpecPb.Controls.Add(Me.fraOrigMgoPlant)
        Me.fraOffSpecPb.Controls.Add(Me.fraDoloflotPlant)
        Me.fraOffSpecPb.Controls.Add(Me.fraFcoDoloflot)
        Me.fraOffSpecPb.Location = New System.Drawing.Point(13, 19)
        Me.fraOffSpecPb.Name = "fraOffSpecPb"
        Me.fraOffSpecPb.Size = New System.Drawing.Size(1008, 527)
        Me.fraOffSpecPb.TabIndex = 2
        Me.fraOffSpecPb.TabStop = False
        '
        'fraOrigMgoPlant
        '
        Me.fraOrigMgoPlant.Controls.Add(Me.ssOffSpecPb)
        Me.fraOrigMgoPlant.Controls.Add(Me.lblGen31)
        Me.fraOrigMgoPlant.Controls.Add(Me.chkUseOrigMgoPlant)
        Me.fraOrigMgoPlant.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraOrigMgoPlant.Location = New System.Drawing.Point(719, 28)
        Me.fraOrigMgoPlant.Name = "fraOrigMgoPlant"
        Me.fraOrigMgoPlant.Size = New System.Drawing.Size(265, 474)
        Me.fraOrigMgoPlant.TabIndex = 2
        Me.fraOrigMgoPlant.TabStop = False
        Me.fraOrigMgoPlant.Text = "Original MgO Plant"
        '
        'ssOffSpecPb
        '
        Me.ssOffSpecPb.Location = New System.Drawing.Point(9, 42)
        Me.ssOffSpecPb.Name = "ssOffSpecPb"
        Me.ssOffSpecPb.OcxState = CType(resources.GetObject("ssOffSpecPb.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssOffSpecPb.Size = New System.Drawing.Size(210, 130)
        Me.ssOffSpecPb.TabIndex = 3
        '
        'lblGen31
        '
        Me.lblGen31.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen31.ForeColor = System.Drawing.Color.Navy
        Me.lblGen31.Location = New System.Drawing.Point(6, 178)
        Me.lblGen31.Name = "lblGen31"
        Me.lblGen31.Size = New System.Drawing.Size(249, 252)
        Me.lblGen31.TabIndex = 2
        Me.lblGen31.Text = "Label1"
        '
        'chkUseOrigMgoPlant
        '
        Me.chkUseOrigMgoPlant.AutoSize = True
        Me.chkUseOrigMgoPlant.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUseOrigMgoPlant.Location = New System.Drawing.Point(9, 19)
        Me.chkUseOrigMgoPlant.Name = "chkUseOrigMgoPlant"
        Me.chkUseOrigMgoPlant.Size = New System.Drawing.Size(136, 17)
        Me.chkUseOrigMgoPlant.TabIndex = 0
        Me.chkUseOrigMgoPlant.Text = "Use Original MgO Plant"
        Me.chkUseOrigMgoPlant.UseVisualStyleBackColor = True
        '
        'fraDoloflotPlant
        '
        Me.fraDoloflotPlant.Controls.Add(Me.ssDoloflotPlant)
        Me.fraDoloflotPlant.Controls.Add(Me.lblGen44)
        Me.fraDoloflotPlant.Controls.Add(Me.lblGen45)
        Me.fraDoloflotPlant.Controls.Add(Me.cmdSetDefaults)
        Me.fraDoloflotPlant.Controls.Add(Me.chkUseDoloflotPlant)
        Me.fraDoloflotPlant.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraDoloflotPlant.Location = New System.Drawing.Point(21, 28)
        Me.fraDoloflotPlant.Name = "fraDoloflotPlant"
        Me.fraDoloflotPlant.Size = New System.Drawing.Size(383, 474)
        Me.fraDoloflotPlant.TabIndex = 1
        Me.fraDoloflotPlant.TabStop = False
        Me.fraDoloflotPlant.Text = "Doloflot Plant (Ona) -- 2010"
        '
        'ssDoloflotPlant
        '
        Me.ssDoloflotPlant.Location = New System.Drawing.Point(10, 48)
        Me.ssDoloflotPlant.Name = "ssDoloflotPlant"
        Me.ssDoloflotPlant.OcxState = CType(resources.GetObject("ssDoloflotPlant.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssDoloflotPlant.Size = New System.Drawing.Size(163, 129)
        Me.ssDoloflotPlant.TabIndex = 5
        '
        'lblGen44
        '
        Me.lblGen44.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen44.ForeColor = System.Drawing.Color.Navy
        Me.lblGen44.Location = New System.Drawing.Point(23, 224)
        Me.lblGen44.Name = "lblGen44"
        Me.lblGen44.Size = New System.Drawing.Size(224, 234)
        Me.lblGen44.TabIndex = 4
        Me.lblGen44.Text = "Label1"
        '
        'lblGen45
        '
        Me.lblGen45.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen45.ForeColor = System.Drawing.Color.Navy
        Me.lblGen45.Location = New System.Drawing.Point(190, 20)
        Me.lblGen45.Name = "lblGen45"
        Me.lblGen45.Size = New System.Drawing.Size(181, 168)
        Me.lblGen45.TabIndex = 3
        Me.lblGen45.Text = "Label1"
        '
        'cmdSetDefaults
        '
        Me.cmdSetDefaults.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSetDefaults.Location = New System.Drawing.Point(186, 195)
        Me.cmdSetDefaults.Name = "cmdSetDefaults"
        Me.cmdSetDefaults.Size = New System.Drawing.Size(75, 23)
        Me.cmdSetDefaults.TabIndex = 1
        Me.cmdSetDefaults.Text = "Set Defaults"
        Me.cmdSetDefaults.UseVisualStyleBackColor = True
        '
        'chkUseDoloflotPlant
        '
        Me.chkUseDoloflotPlant.AutoSize = True
        Me.chkUseDoloflotPlant.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUseDoloflotPlant.Location = New System.Drawing.Point(7, 20)
        Me.chkUseDoloflotPlant.Name = "chkUseDoloflotPlant"
        Me.chkUseDoloflotPlant.Size = New System.Drawing.Size(176, 17)
        Me.chkUseDoloflotPlant.TabIndex = 0
        Me.chkUseDoloflotPlant.Text = "Use Doloflot Plant (Ona) -- 2010"
        Me.chkUseDoloflotPlant.UseVisualStyleBackColor = True
        '
        'fraFcoDoloflot
        '
        Me.fraFcoDoloflot.Controls.Add(Me.ssDoloflotPlantFco2)
        Me.fraFcoDoloflot.Controls.Add(Me.ssDoloflotPlantFco)
        Me.fraFcoDoloflot.Controls.Add(Me.lblGen51)
        Me.fraFcoDoloflot.Controls.Add(Me.chkUseDoloflotPlantFco)
        Me.fraFcoDoloflot.Controls.Add(Me.cmdSetDefaults2)
        Me.fraFcoDoloflot.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraFcoDoloflot.Location = New System.Drawing.Point(419, 28)
        Me.fraFcoDoloflot.Name = "fraFcoDoloflot"
        Me.fraFcoDoloflot.Size = New System.Drawing.Size(279, 474)
        Me.fraFcoDoloflot.TabIndex = 0
        Me.fraFcoDoloflot.TabStop = False
        Me.fraFcoDoloflot.Text = "Doloflot Plant (FCO) -- 2011"
        '
        'ssDoloflotPlantFco2
        '
        Me.ssDoloflotPlantFco2.Location = New System.Drawing.Point(15, 146)
        Me.ssDoloflotPlantFco2.Name = "ssDoloflotPlantFco2"
        Me.ssDoloflotPlantFco2.OcxState = CType(resources.GetObject("ssDoloflotPlantFco2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssDoloflotPlantFco2.Size = New System.Drawing.Size(189, 113)
        Me.ssDoloflotPlantFco2.TabIndex = 6
        '
        'ssDoloflotPlantFco
        '
        Me.ssDoloflotPlantFco.Location = New System.Drawing.Point(15, 78)
        Me.ssDoloflotPlantFco.Name = "ssDoloflotPlantFco"
        Me.ssDoloflotPlantFco.OcxState = CType(resources.GetObject("ssDoloflotPlantFco.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssDoloflotPlantFco.Size = New System.Drawing.Size(166, 33)
        Me.ssDoloflotPlantFco.TabIndex = 5
        '
        'lblGen51
        '
        Me.lblGen51.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen51.ForeColor = System.Drawing.Color.Navy
        Me.lblGen51.Location = New System.Drawing.Point(16, 44)
        Me.lblGen51.Name = "lblGen51"
        Me.lblGen51.Size = New System.Drawing.Size(257, 31)
        Me.lblGen51.TabIndex = 4
        Me.lblGen51.Text = "Label1"
        '
        'chkUseDoloflotPlantFco
        '
        Me.chkUseDoloflotPlantFco.AutoSize = True
        Me.chkUseDoloflotPlantFco.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUseDoloflotPlantFco.Location = New System.Drawing.Point(15, 22)
        Me.chkUseDoloflotPlantFco.Name = "chkUseDoloflotPlantFco"
        Me.chkUseDoloflotPlantFco.Size = New System.Drawing.Size(177, 17)
        Me.chkUseDoloflotPlantFco.TabIndex = 1
        Me.chkUseDoloflotPlantFco.Text = "Use Doloflot Plant (FCO) -- 2011"
        Me.chkUseDoloflotPlantFco.UseVisualStyleBackColor = True
        '
        'cmdSetDefaults2
        '
        Me.cmdSetDefaults2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSetDefaults2.Location = New System.Drawing.Point(15, 277)
        Me.cmdSetDefaults2.Name = "cmdSetDefaults2"
        Me.cmdSetDefaults2.Size = New System.Drawing.Size(75, 23)
        Me.cmdSetDefaults2.TabIndex = 0
        Me.cmdSetDefaults2.Text = "Set Defaults"
        Me.cmdSetDefaults2.UseVisualStyleBackColor = True
        '
        'TabPage10
        '
        Me.TabPage10.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage10.Controls.Add(Me.fraRptAllToTextFile)
        Me.TabPage10.Controls.Add(Me.fraDataSaveComm)
        Me.TabPage10.Controls.Add(Me.fraOutputLocation2)
        Me.TabPage10.Controls.Add(Me.fraOutputLocation1)
        Me.TabPage10.Controls.Add(Me.fraProspDatasetType)
        Me.TabPage10.Controls.Add(Me.fraExtra)
        Me.TabPage10.Location = New System.Drawing.Point(4, 22)
        Me.TabPage10.Name = "TabPage10"
        Me.TabPage10.Size = New System.Drawing.Size(1323, 565)
        Me.TabPage10.TabIndex = 9
        Me.TabPage10.Text = "Output"
        '
        'fraRptAllToTextFile
        '
        Me.fraRptAllToTextFile.Controls.Add(Me.lblRptAllCnt2)
        Me.fraRptAllToTextFile.Controls.Add(Me.cmdRptAllToTxtFile)
        Me.fraRptAllToTextFile.Controls.Add(Me.txtRptAllToTxtFile)
        Me.fraRptAllToTextFile.Controls.Add(Me.cmdPrintRptAll)
        Me.fraRptAllToTextFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraRptAllToTextFile.Location = New System.Drawing.Point(477, 300)
        Me.fraRptAllToTextFile.Name = "fraRptAllToTextFile"
        Me.fraRptAllToTextFile.Size = New System.Drawing.Size(425, 123)
        Me.fraRptAllToTextFile.TabIndex = 0
        Me.fraRptAllToTextFile.TabStop = False
        Me.fraRptAllToTextFile.Text = """Report All"" to Text File"
        '
        'lblRptAllCnt2
        '
        Me.lblRptAllCnt2.AutoSize = True
        Me.lblRptAllCnt2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRptAllCnt2.Location = New System.Drawing.Point(6, 58)
        Me.lblRptAllCnt2.Name = "lblRptAllCnt2"
        Me.lblRptAllCnt2.Size = New System.Drawing.Size(39, 13)
        Me.lblRptAllCnt2.TabIndex = 3
        Me.lblRptAllCnt2.Text = "Label1"
        '
        'cmdRptAllToTxtFile
        '
        Me.cmdRptAllToTxtFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRptAllToTxtFile.Location = New System.Drawing.Point(8, 92)
        Me.cmdRptAllToTxtFile.Name = "cmdRptAllToTxtFile"
        Me.cmdRptAllToTxtFile.Size = New System.Drawing.Size(145, 23)
        Me.cmdRptAllToTxtFile.TabIndex = 2
        Me.cmdRptAllToTxtFile.Text = "Create RptAll Text"
        Me.cmdRptAllToTxtFile.UseVisualStyleBackColor = True
        '
        'txtRptAllToTxtFile
        '
        Me.txtRptAllToTxtFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRptAllToTxtFile.Location = New System.Drawing.Point(8, 20)
        Me.txtRptAllToTxtFile.Name = "txtRptAllToTxtFile"
        Me.txtRptAllToTxtFile.Size = New System.Drawing.Size(402, 20)
        Me.txtRptAllToTxtFile.TabIndex = 1
        '
        'cmdPrintRptAll
        '
        Me.cmdPrintRptAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrintRptAll.Location = New System.Drawing.Point(319, 93)
        Me.cmdPrintRptAll.Name = "cmdPrintRptAll"
        Me.cmdPrintRptAll.Size = New System.Drawing.Size(91, 23)
        Me.cmdPrintRptAll.TabIndex = 0
        Me.cmdPrintRptAll.Text = "Print RptAll Text"
        Me.cmdPrintRptAll.UseVisualStyleBackColor = True
        '
        'fraDataSaveComm
        '
        Me.fraDataSaveComm.Controls.Add(Me.lblGen53)
        Me.fraDataSaveComm.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraDataSaveComm.Location = New System.Drawing.Point(19, 442)
        Me.fraDataSaveComm.Name = "fraDataSaveComm"
        Me.fraDataSaveComm.Size = New System.Drawing.Size(452, 150)
        Me.fraDataSaveComm.TabIndex = 1
        Me.fraDataSaveComm.TabStop = False
        Me.fraDataSaveComm.Text = "Miscellaneous save comments"
        '
        'lblGen53
        '
        Me.lblGen53.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen53.ForeColor = System.Drawing.Color.Navy
        Me.lblGen53.Location = New System.Drawing.Point(15, 20)
        Me.lblGen53.Name = "lblGen53"
        Me.lblGen53.Size = New System.Drawing.Size(420, 116)
        Me.lblGen53.TabIndex = 0
        Me.lblGen53.Text = "Label1"
        '
        'fraOutputLocation2
        '
        Me.fraOutputLocation2.Controls.Add(Me.Label7)
        Me.fraOutputLocation2.Controls.Add(Me.Label6)
        Me.fraOutputLocation2.Controls.Add(Me.lblTextfileComment)
        Me.fraOutputLocation2.Controls.Add(Me.lblGen10)
        Me.fraOutputLocation2.Controls.Add(Me.lblGen37)
        Me.fraOutputLocation2.Controls.Add(Me.txtProspDatasetTextfileName)
        Me.fraOutputLocation2.Controls.Add(Me.chkSurvCaddTextfile)
        Me.fraOutputLocation2.Controls.Add(Me.chkSpecMoisTransferFile)
        Me.fraOutputLocation2.Controls.Add(Me.chkInclMgPlt)
        Me.fraOutputLocation2.Controls.Add(Me.chkPbAnalysisFillInSpecial)
        Me.fraOutputLocation2.Controls.Add(Me.chkBdFormatTextfile)
        Me.fraOutputLocation2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraOutputLocation2.Location = New System.Drawing.Point(19, 128)
        Me.fraOutputLocation2.Name = "fraOutputLocation2"
        Me.fraOutputLocation2.Size = New System.Drawing.Size(883, 166)
        Me.fraOutputLocation2.TabIndex = 2
        Me.fraOutputLocation2.TabStop = False
        Me.fraOutputLocation2.Text = "Save to SurvCADD or MOIS special transfer textfile"
        '
        'Label7
        '
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label7.Location = New System.Drawing.Point(10, 120)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(165, 2)
        Me.Label7.TabIndex = 10
        '
        'Label6
        '
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label6.Location = New System.Drawing.Point(12, 52)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(165, 2)
        Me.Label6.TabIndex = 9
        '
        'lblTextfileComment
        '
        Me.lblTextfileComment.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTextfileComment.Location = New System.Drawing.Point(200, 57)
        Me.lblTextfileComment.Name = "lblTextfileComment"
        Me.lblTextfileComment.Size = New System.Drawing.Size(216, 92)
        Me.lblTextfileComment.TabIndex = 8
        Me.lblTextfileComment.Text = "lblTextfileComment"
        '
        'lblGen10
        '
        Me.lblGen10.AutoSize = True
        Me.lblGen10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen10.Location = New System.Drawing.Point(251, 22)
        Me.lblGen10.Name = "lblGen10"
        Me.lblGen10.Size = New System.Drawing.Size(70, 13)
        Me.lblGen10.TabIndex = 7
        Me.lblGen10.Text = "Textfile name"
        '
        'lblGen37
        '
        Me.lblGen37.AutoSize = True
        Me.lblGen37.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen37.Location = New System.Drawing.Point(440, 45)
        Me.lblGen37.Name = "lblGen37"
        Me.lblGen37.Size = New System.Drawing.Size(208, 16)
        Me.lblGen37.TabIndex = 6
        Me.lblGen37.Text = "For SurvCADD Hole Textfiles"
        '
        'txtProspDatasetTextfileName
        '
        Me.txtProspDatasetTextfileName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProspDatasetTextfileName.Location = New System.Drawing.Point(320, 19)
        Me.txtProspDatasetTextfileName.Name = "txtProspDatasetTextfileName"
        Me.txtProspDatasetTextfileName.Size = New System.Drawing.Size(548, 20)
        Me.txtProspDatasetTextfileName.TabIndex = 5
        '
        'chkSurvCaddTextfile
        '
        Me.chkSurvCaddTextfile.AutoSize = True
        Me.chkSurvCaddTextfile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSurvCaddTextfile.Location = New System.Drawing.Point(12, 22)
        Me.chkSurvCaddTextfile.Name = "chkSurvCaddTextfile"
        Me.chkSurvCaddTextfile.Size = New System.Drawing.Size(149, 17)
        Me.chkSurvCaddTextfile.TabIndex = 4
        Me.chkSurvCaddTextfile.Text = "SurvCADD transfer textfile"
        Me.chkSurvCaddTextfile.UseVisualStyleBackColor = True
        '
        'chkSpecMoisTransferFile
        '
        Me.chkSpecMoisTransferFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSpecMoisTransferFile.Location = New System.Drawing.Point(12, 57)
        Me.chkSpecMoisTransferFile.Name = "chkSpecMoisTransferFile"
        Me.chkSpecMoisTransferFile.Size = New System.Drawing.Size(187, 36)
        Me.chkSpecMoisTransferFile.TabIndex = 3
        Me.chkSpecMoisTransferFile.Text = "Special MOIS transfer textfile (IMC-RAR format)"
        Me.chkSpecMoisTransferFile.UseVisualStyleBackColor = True
        '
        'chkInclMgPlt
        '
        Me.chkInclMgPlt.AutoSize = True
        Me.chkInclMgPlt.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkInclMgPlt.Location = New System.Drawing.Point(12, 94)
        Me.chkInclMgPlt.Name = "chkInclMgPlt"
        Me.chkInclMgPlt.Size = New System.Drawing.Size(137, 17)
        Me.chkInclMgPlt.TabIndex = 2
        Me.chkInclMgPlt.Text = "Include MgO plant data"
        Me.chkInclMgPlt.UseVisualStyleBackColor = True
        '
        'chkPbAnalysisFillInSpecial
        '
        Me.chkPbAnalysisFillInSpecial.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkPbAnalysisFillInSpecial.Location = New System.Drawing.Point(443, 66)
        Me.chkPbAnalysisFillInSpecial.Name = "chkPbAnalysisFillInSpecial"
        Me.chkPbAnalysisFillInSpecial.Size = New System.Drawing.Size(413, 88)
        Me.chkPbAnalysisFillInSpecial.TabIndex = 1
        Me.chkPbAnalysisFillInSpecial.Text = "chkPbAnalysisFillInSpecial"
        Me.chkPbAnalysisFillInSpecial.UseVisualStyleBackColor = True
        '
        'chkBdFormatTextfile
        '
        Me.chkBdFormatTextfile.AutoSize = True
        Me.chkBdFormatTextfile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBdFormatTextfile.Location = New System.Drawing.Point(12, 132)
        Me.chkBdFormatTextfile.Name = "chkBdFormatTextfile"
        Me.chkBdFormatTextfile.Size = New System.Drawing.Size(144, 17)
        Me.chkBdFormatTextfile.TabIndex = 0
        Me.chkBdFormatTextfile.Text = "BD format transfer textfile"
        Me.chkBdFormatTextfile.UseVisualStyleBackColor = True
        '
        'fraOutputLocation1
        '
        Me.fraOutputLocation1.Controls.Add(Me.lblGen9)
        Me.fraOutputLocation1.Controls.Add(Me.lblGen11)
        Me.fraOutputLocation1.Controls.Add(Me.chkSaveToDatabase)
        Me.fraOutputLocation1.Controls.Add(Me.txtProspectDatasetName)
        Me.fraOutputLocation1.Controls.Add(Me.txtProspectDatasetDesc)
        Me.fraOutputLocation1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraOutputLocation1.Location = New System.Drawing.Point(19, 28)
        Me.fraOutputLocation1.Name = "fraOutputLocation1"
        Me.fraOutputLocation1.Size = New System.Drawing.Size(883, 93)
        Me.fraOutputLocation1.TabIndex = 3
        Me.fraOutputLocation1.TabStop = False
        Me.fraOutputLocation1.Text = "Save to MOIS prospect database"
        '
        'lblGen9
        '
        Me.lblGen9.AutoSize = True
        Me.lblGen9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen9.Location = New System.Drawing.Point(257, 51)
        Me.lblGen9.Name = "lblGen9"
        Me.lblGen9.Size = New System.Drawing.Size(60, 13)
        Me.lblGen9.TabIndex = 4
        Me.lblGen9.Text = "Description"
        '
        'lblGen11
        '
        Me.lblGen11.AutoSize = True
        Me.lblGen11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen11.Location = New System.Drawing.Point(201, 23)
        Me.lblGen11.Name = "lblGen11"
        Me.lblGen11.Size = New System.Drawing.Size(116, 13)
        Me.lblGen11.TabIndex = 3
        Me.lblGen11.Text = "Prospect dataset name"
        '
        'chkSaveToDatabase
        '
        Me.chkSaveToDatabase.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSaveToDatabase.Location = New System.Drawing.Point(10, 23)
        Me.chkSaveToDatabase.Name = "chkSaveToDatabase"
        Me.chkSaveToDatabase.Size = New System.Drawing.Size(176, 38)
        Me.chkSaveToDatabase.TabIndex = 2
        Me.chkSaveToDatabase.Text = "Save prospect dataset to database"
        Me.chkSaveToDatabase.UseVisualStyleBackColor = True
        '
        'txtProspectDatasetName
        '
        Me.txtProspectDatasetName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProspectDatasetName.Location = New System.Drawing.Point(320, 19)
        Me.txtProspectDatasetName.MaxLength = 30
        Me.txtProspectDatasetName.Name = "txtProspectDatasetName"
        Me.txtProspectDatasetName.Size = New System.Drawing.Size(147, 20)
        Me.txtProspectDatasetName.TabIndex = 1
        '
        'txtProspectDatasetDesc
        '
        Me.txtProspectDatasetDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProspectDatasetDesc.Location = New System.Drawing.Point(320, 47)
        Me.txtProspectDatasetDesc.MaxLength = 200
        Me.txtProspectDatasetDesc.Name = "txtProspectDatasetDesc"
        Me.txtProspectDatasetDesc.Size = New System.Drawing.Size(304, 20)
        Me.txtProspectDatasetDesc.TabIndex = 0
        '
        'fraProspDatasetType
        '
        Me.fraProspDatasetType.Controls.Add(Me.optInclSplits)
        Me.fraProspDatasetType.Controls.Add(Me.optInclComposites)
        Me.fraProspDatasetType.Controls.Add(Me.chkProductionCoefficient)
        Me.fraProspDatasetType.Controls.Add(Me.optInclBoth)
        Me.fraProspDatasetType.Controls.Add(Me.chk100Pct)
        Me.fraProspDatasetType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraProspDatasetType.Location = New System.Drawing.Point(19, 300)
        Me.fraProspDatasetType.Name = "fraProspDatasetType"
        Me.fraProspDatasetType.Size = New System.Drawing.Size(452, 123)
        Me.fraProspDatasetType.TabIndex = 4
        Me.fraProspDatasetType.TabStop = False
        Me.fraProspDatasetType.Text = "Select type of prospect data"
        '
        'optInclSplits
        '
        Me.optInclSplits.AutoSize = True
        Me.optInclSplits.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optInclSplits.Location = New System.Drawing.Point(240, 35)
        Me.optInclSplits.Name = "optInclSplits"
        Me.optInclSplits.Size = New System.Drawing.Size(50, 17)
        Me.optInclSplits.TabIndex = 2
        Me.optInclSplits.TabStop = True
        Me.optInclSplits.Text = "Splits"
        Me.optInclSplits.UseVisualStyleBackColor = True
        '
        'optInclComposites
        '
        Me.optInclComposites.AutoSize = True
        Me.optInclComposites.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optInclComposites.Location = New System.Drawing.Point(240, 57)
        Me.optInclComposites.Name = "optInclComposites"
        Me.optInclComposites.Size = New System.Drawing.Size(52, 17)
        Me.optInclComposites.TabIndex = 1
        Me.optInclComposites.TabStop = True
        Me.optInclComposites.Text = "Holes"
        Me.optInclComposites.UseVisualStyleBackColor = True
        '
        'chkProductionCoefficient
        '
        Me.chkProductionCoefficient.AutoSize = True
        Me.chkProductionCoefficient.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkProductionCoefficient.Location = New System.Drawing.Point(13, 79)
        Me.chkProductionCoefficient.Name = "chkProductionCoefficient"
        Me.chkProductionCoefficient.Size = New System.Drawing.Size(175, 17)
        Me.chkProductionCoefficient.TabIndex = 3
        Me.chkProductionCoefficient.Text = "Production Coefficient (Catalog)"
        Me.chkProductionCoefficient.UseVisualStyleBackColor = True
        '
        'optInclBoth
        '
        Me.optInclBoth.AutoSize = True
        Me.optInclBoth.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optInclBoth.Location = New System.Drawing.Point(240, 79)
        Me.optInclBoth.Name = "optInclBoth"
        Me.optInclBoth.Size = New System.Drawing.Size(89, 17)
        Me.optInclBoth.TabIndex = 0
        Me.optInclBoth.TabStop = True
        Me.optInclBoth.Text = "Splits && Holes"
        Me.optInclBoth.UseVisualStyleBackColor = True
        '
        'chk100Pct
        '
        Me.chk100Pct.AutoSize = True
        Me.chk100Pct.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk100Pct.Location = New System.Drawing.Point(13, 45)
        Me.chk100Pct.Name = "chk100Pct"
        Me.chk100Pct.Size = New System.Drawing.Size(97, 17)
        Me.chk100Pct.TabIndex = 4
        Me.chk100Pct.Text = "100% Prospect"
        Me.chk100Pct.UseVisualStyleBackColor = True
        '
        'fraExtra
        '
        Me.fraExtra.Location = New System.Drawing.Point(477, 442)
        Me.fraExtra.Name = "fraExtra"
        Me.fraExtra.Size = New System.Drawing.Size(425, 150)
        Me.fraExtra.TabIndex = 5
        Me.fraExtra.TabStop = False
        '
        'TabPage11
        '
        Me.TabPage11.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage11.Controls.Add(Me.fraOverride)
        Me.TabPage11.Location = New System.Drawing.Point(4, 22)
        Me.TabPage11.Name = "TabPage11"
        Me.TabPage11.Size = New System.Drawing.Size(1323, 565)
        Me.TabPage11.TabIndex = 10
        Me.TabPage11.Text = "O-ride"
        '
        'fraOverride
        '
        Me.fraOverride.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraOverride.Controls.Add(Me.fraSplitOverride)
        Me.fraOverride.Controls.Add(Me.fraOverrideList)
        Me.fraOverride.Location = New System.Drawing.Point(4, 18)
        Me.fraOverride.Name = "fraOverride"
        Me.fraOverride.Size = New System.Drawing.Size(1068, 612)
        Me.fraOverride.TabIndex = 0
        Me.fraOverride.TabStop = False
        '
        'fraSplitOverride
        '
        Me.fraSplitOverride.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.fraSplitOverride.Controls.Add(Me.Frame2)
        Me.fraSplitOverride.Controls.Add(Me.cmdSaveSplitOverride)
        Me.fraSplitOverride.Controls.Add(Me.Label4)
        Me.fraSplitOverride.Controls.Add(Me.cmdCancelSplitOverride)
        Me.fraSplitOverride.Controls.Add(Me.Label3)
        Me.fraSplitOverride.Controls.Add(Me.cboSplitOverrideMineName)
        Me.fraSplitOverride.Controls.Add(Me.cmdDeleteSplitOverride)
        Me.fraSplitOverride.Controls.Add(Me.lblSplitOverrideTxtFile)
        Me.fraSplitOverride.Controls.Add(Me.txtSplOverrideTxtFile)
        Me.fraSplitOverride.Controls.Add(Me.cmdLoadOverrideTxtFile)
        Me.fraSplitOverride.Controls.Add(Me.txtSplitOverrideName)
        Me.fraSplitOverride.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraSplitOverride.Location = New System.Drawing.Point(23, 290)
        Me.fraSplitOverride.Name = "fraSplitOverride"
        Me.fraSplitOverride.Size = New System.Drawing.Size(939, 315)
        Me.fraSplitOverride.TabIndex = 1
        Me.fraSplitOverride.TabStop = False
        Me.fraSplitOverride.Text = "Split Override Saved Sets"
        '
        'Frame2
        '
        Me.Frame2.Controls.Add(Me.ssSplitOverrides)
        Me.Frame2.Controls.Add(Me.cmdGetSplitOverrides)
        Me.Frame2.Controls.Add(Me.chkOnlyMySplitOverride)
        Me.Frame2.Location = New System.Drawing.Point(314, 19)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Size = New System.Drawing.Size(619, 257)
        Me.Frame2.TabIndex = 4
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Select an existing set"
        '
        'ssSplitOverrides
        '
        Me.ssSplitOverrides.Location = New System.Drawing.Point(16, 42)
        Me.ssSplitOverrides.Name = "ssSplitOverrides"
        Me.ssSplitOverrides.OcxState = CType(resources.GetObject("ssSplitOverrides.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssSplitOverrides.Size = New System.Drawing.Size(528, 209)
        Me.ssSplitOverrides.TabIndex = 10
        '
        'cmdGetSplitOverrides
        '
        Me.cmdGetSplitOverrides.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGetSplitOverrides.Location = New System.Drawing.Point(201, 13)
        Me.cmdGetSplitOverrides.Name = "cmdGetSplitOverrides"
        Me.cmdGetSplitOverrides.Size = New System.Drawing.Size(121, 23)
        Me.cmdGetSplitOverrides.TabIndex = 1
        Me.cmdGetSplitOverrides.Text = "Get Overrides"
        Me.cmdGetSplitOverrides.UseVisualStyleBackColor = True
        '
        'chkOnlyMySplitOverride
        '
        Me.chkOnlyMySplitOverride.AutoSize = True
        Me.chkOnlyMySplitOverride.Checked = True
        Me.chkOnlyMySplitOverride.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOnlyMySplitOverride.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOnlyMySplitOverride.Location = New System.Drawing.Point(16, 19)
        Me.chkOnlyMySplitOverride.Name = "chkOnlyMySplitOverride"
        Me.chkOnlyMySplitOverride.Size = New System.Drawing.Size(161, 17)
        Me.chkOnlyMySplitOverride.TabIndex = 0
        Me.chkOnlyMySplitOverride.Text = "Select only my split overrides"
        Me.chkOnlyMySplitOverride.UseVisualStyleBackColor = True
        '
        'cmdSaveSplitOverride
        '
        Me.cmdSaveSplitOverride.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSaveSplitOverride.Location = New System.Drawing.Point(47, 112)
        Me.cmdSaveSplitOverride.Name = "cmdSaveSplitOverride"
        Me.cmdSaveSplitOverride.Size = New System.Drawing.Size(247, 23)
        Me.cmdSaveSplitOverride.TabIndex = 4
        Me.cmdSaveSplitOverride.Text = "Save Split Override Set"
        Me.cmdSaveSplitOverride.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(45, 73)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Mine name"
        '
        'cmdCancelSplitOverride
        '
        Me.cmdCancelSplitOverride.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancelSplitOverride.Location = New System.Drawing.Point(128, 141)
        Me.cmdCancelSplitOverride.Name = "cmdCancelSplitOverride"
        Me.cmdCancelSplitOverride.Size = New System.Drawing.Size(164, 23)
        Me.cmdCancelSplitOverride.TabIndex = 3
        Me.cmdCancelSplitOverride.Text = "Cancel Split Override Set"
        Me.cmdCancelSplitOverride.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(31, 27)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 30)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Split override name"
        '
        'cboSplitOverrideMineName
        '
        Me.cboSplitOverrideMineName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSplitOverrideMineName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSplitOverrideMineName.FormattingEnabled = True
        Me.cboSplitOverrideMineName.Location = New System.Drawing.Point(107, 69)
        Me.cboSplitOverrideMineName.Name = "cboSplitOverrideMineName"
        Me.cboSplitOverrideMineName.Size = New System.Drawing.Size(187, 21)
        Me.cboSplitOverrideMineName.TabIndex = 2
        '
        'cmdDeleteSplitOverride
        '
        Me.cmdDeleteSplitOverride.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDeleteSplitOverride.Location = New System.Drawing.Point(47, 141)
        Me.cmdDeleteSplitOverride.Name = "cmdDeleteSplitOverride"
        Me.cmdDeleteSplitOverride.Size = New System.Drawing.Size(75, 23)
        Me.cmdDeleteSplitOverride.TabIndex = 5
        Me.cmdDeleteSplitOverride.Text = "Delete Split Override Set"
        Me.cmdDeleteSplitOverride.UseVisualStyleBackColor = True
        '
        'lblSplitOverrideTxtFile
        '
        Me.lblSplitOverrideTxtFile.AutoSize = True
        Me.lblSplitOverrideTxtFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSplitOverrideTxtFile.Location = New System.Drawing.Point(31, 286)
        Me.lblSplitOverrideTxtFile.Name = "lblSplitOverrideTxtFile"
        Me.lblSplitOverrideTxtFile.Size = New System.Drawing.Size(104, 13)
        Me.lblSplitOverrideTxtFile.TabIndex = 6
        Me.lblSplitOverrideTxtFile.Text = "Split override text file"
        '
        'txtSplOverrideTxtFile
        '
        Me.txtSplOverrideTxtFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSplOverrideTxtFile.Location = New System.Drawing.Point(141, 283)
        Me.txtSplOverrideTxtFile.MaxLength = 200
        Me.txtSplOverrideTxtFile.Name = "txtSplOverrideTxtFile"
        Me.txtSplOverrideTxtFile.Size = New System.Drawing.Size(626, 20)
        Me.txtSplOverrideTxtFile.TabIndex = 1
        '
        'cmdLoadOverrideTxtFile
        '
        Me.cmdLoadOverrideTxtFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLoadOverrideTxtFile.Location = New System.Drawing.Point(818, 281)
        Me.cmdLoadOverrideTxtFile.Name = "cmdLoadOverrideTxtFile"
        Me.cmdLoadOverrideTxtFile.Size = New System.Drawing.Size(115, 23)
        Me.cmdLoadOverrideTxtFile.TabIndex = 0
        Me.cmdLoadOverrideTxtFile.Text = "Load Override Text"
        Me.cmdLoadOverrideTxtFile.UseVisualStyleBackColor = True
        '
        'txtSplitOverrideName
        '
        Me.txtSplitOverrideName.Location = New System.Drawing.Point(107, 32)
        Me.txtSplitOverrideName.MaxLength = 30
        Me.txtSplitOverrideName.Name = "txtSplitOverrideName"
        Me.txtSplitOverrideName.Size = New System.Drawing.Size(190, 22)
        Me.txtSplitOverrideName.TabIndex = 3
        '
        'fraOverrideList
        '
        Me.fraOverrideList.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraOverrideList.Controls.Add(Me.ssSplitOverride)
        Me.fraOverrideList.Controls.Add(Me.ssRawProspMin)
        Me.fraOverrideList.Controls.Add(Me.lblGen54)
        Me.fraOverrideList.Controls.Add(Me.lblGen55)
        Me.fraOverrideList.Controls.Add(Me.lblGen33)
        Me.fraOverrideList.Controls.Add(Me.lblGen34)
        Me.fraOverrideList.Controls.Add(Me.lblGen36)
        Me.fraOverrideList.Controls.Add(Me.cmdApplySplOverrides)
        Me.fraOverrideList.Controls.Add(Me.cmdClrOverride)
        Me.fraOverrideList.Controls.Add(Me.chkUseRawProspAsOverride)
        Me.fraOverrideList.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraOverrideList.Location = New System.Drawing.Point(23, 26)
        Me.fraOverrideList.Name = "fraOverrideList"
        Me.fraOverrideList.Size = New System.Drawing.Size(1023, 258)
        Me.fraOverrideList.TabIndex = 0
        Me.fraOverrideList.TabStop = False
        Me.fraOverrideList.Text = "Override Splits"
        '
        'ssSplitOverride
        '
        Me.ssSplitOverride.Location = New System.Drawing.Point(14, 30)
        Me.ssSplitOverride.Name = "ssSplitOverride"
        Me.ssSplitOverride.OcxState = CType(resources.GetObject("ssSplitOverride.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssSplitOverride.Size = New System.Drawing.Size(230, 214)
        Me.ssSplitOverride.TabIndex = 11
        '
        'ssRawProspMin
        '
        Me.ssRawProspMin.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ssRawProspMin.Location = New System.Drawing.Point(376, 28)
        Me.ssRawProspMin.Name = "ssRawProspMin"
        Me.ssRawProspMin.OcxState = CType(resources.GetObject("ssRawProspMin.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssRawProspMin.Size = New System.Drawing.Size(641, 163)
        Me.ssRawProspMin.TabIndex = 10
        '
        'lblGen54
        '
        Me.lblGen54.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen54.ForeColor = System.Drawing.Color.Navy
        Me.lblGen54.Location = New System.Drawing.Point(255, 30)
        Me.lblGen54.Name = "lblGen54"
        Me.lblGen54.Size = New System.Drawing.Size(106, 48)
        Me.lblGen54.TabIndex = 9
        Me.lblGen54.Text = "Label5"
        '
        'lblGen55
        '
        Me.lblGen55.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen55.ForeColor = System.Drawing.Color.Navy
        Me.lblGen55.Location = New System.Drawing.Point(255, 90)
        Me.lblGen55.Name = "lblGen55"
        Me.lblGen55.Size = New System.Drawing.Size(106, 48)
        Me.lblGen55.TabIndex = 8
        Me.lblGen55.Text = "Label4"
        '
        'lblGen33
        '
        Me.lblGen33.AutoSize = True
        Me.lblGen33.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen33.Location = New System.Drawing.Point(407, 14)
        Me.lblGen33.Name = "lblGen33"
        Me.lblGen33.Size = New System.Drawing.Size(283, 13)
        Me.lblGen33.TabIndex = 7
        Me.lblGen33.Text = "Minabilities set in raw prospect data for this area"
        '
        'lblGen34
        '
        Me.lblGen34.AutoSize = True
        Me.lblGen34.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen34.ForeColor = System.Drawing.Color.Navy
        Me.lblGen34.Location = New System.Drawing.Point(374, 205)
        Me.lblGen34.Name = "lblGen34"
        Me.lblGen34.Size = New System.Drawing.Size(39, 13)
        Me.lblGen34.TabIndex = 6
        Me.lblGen34.Text = "Label2"
        '
        'lblGen36
        '
        Me.lblGen36.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen36.ForeColor = System.Drawing.Color.Navy
        Me.lblGen36.Location = New System.Drawing.Point(520, 221)
        Me.lblGen36.Name = "lblGen36"
        Me.lblGen36.Size = New System.Drawing.Size(413, 31)
        Me.lblGen36.TabIndex = 5
        Me.lblGen36.Text = "Label1"
        '
        'cmdApplySplOverrides
        '
        Me.cmdApplySplOverrides.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdApplySplOverrides.Location = New System.Drawing.Point(250, 171)
        Me.cmdApplySplOverrides.Name = "cmdApplySplOverrides"
        Me.cmdApplySplOverrides.Size = New System.Drawing.Size(72, 36)
        Me.cmdApplySplOverrides.TabIndex = 4
        Me.cmdApplySplOverrides.Text = "Apply Split Overrides"
        Me.cmdApplySplOverrides.UseVisualStyleBackColor = True
        Me.cmdApplySplOverrides.Visible = False
        '
        'cmdClrOverride
        '
        Me.cmdClrOverride.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClrOverride.Location = New System.Drawing.Point(250, 210)
        Me.cmdClrOverride.Name = "cmdClrOverride"
        Me.cmdClrOverride.Size = New System.Drawing.Size(72, 34)
        Me.cmdClrOverride.TabIndex = 3
        Me.cmdClrOverride.Text = "Clear Override"
        Me.cmdClrOverride.UseVisualStyleBackColor = True
        '
        'chkUseRawProspAsOverride
        '
        Me.chkUseRawProspAsOverride.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUseRawProspAsOverride.Location = New System.Drawing.Point(377, 221)
        Me.chkUseRawProspAsOverride.Name = "chkUseRawProspAsOverride"
        Me.chkUseRawProspAsOverride.Size = New System.Drawing.Size(140, 31)
        Me.chkUseRawProspAsOverride.TabIndex = 0
        Me.chkUseRawProspAsOverride.Text = "Use raw prospect minabilities as override"
        Me.chkUseRawProspAsOverride.UseVisualStyleBackColor = True
        '
        'tabReview
        '
        Me.tabReview.BackColor = System.Drawing.SystemColors.Control
        Me.tabReview.Controls.Add(Me.fraReview)
        Me.tabReview.Controls.Add(Me.fraReptDisp)
        Me.tabReview.Location = New System.Drawing.Point(4, 22)
        Me.tabReview.Name = "tabReview"
        Me.tabReview.Size = New System.Drawing.Size(1323, 565)
        Me.tabReview.TabIndex = 11
        Me.tabReview.Text = "Review"
        '
        'fraReview
        '
        Me.fraReview.Controls.Add(Me.TableLayoutPanel1)
        Me.fraReview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.fraReview.Location = New System.Drawing.Point(0, 0)
        Me.fraReview.Name = "fraReview"
        Me.fraReview.Size = New System.Drawing.Size(1323, 565)
        Me.fraReview.TabIndex = 5
        Me.fraReview.TabStop = False
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 350.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.fraGeneratedData, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.fraDetlDispl, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.fraDataTypeOption, 1, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(3, 16)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 105.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1317, 546)
        Me.TableLayoutPanel1.TabIndex = 13
        '
        'fraGeneratedData
        '
        Me.fraGeneratedData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraGeneratedData.Controls.Add(Me.tbcSplitResults)
        Me.fraGeneratedData.Controls.Add(Me.tbcHoleResults)
        Me.fraGeneratedData.Controls.Add(Me.ssCompErrors)
        Me.fraGeneratedData.Controls.Add(Me.lblBarrenSplComm)
        Me.fraGeneratedData.Controls.Add(Me.lblGen23)
        Me.fraGeneratedData.Controls.Add(Me.lblGen24)
        Me.fraGeneratedData.Controls.Add(Me.lblGen25)
        Me.fraGeneratedData.Controls.Add(Me.lblGen26)
        Me.fraGeneratedData.Controls.Add(Me.lblNoReview)
        Me.fraGeneratedData.Controls.Add(Me.cmdCopyToOverrides)
        Me.fraGeneratedData.Location = New System.Drawing.Point(3, 3)
        Me.fraGeneratedData.Name = "fraGeneratedData"
        Me.TableLayoutPanel1.SetRowSpan(Me.fraGeneratedData, 2)
        Me.fraGeneratedData.Size = New System.Drawing.Size(961, 540)
        Me.fraGeneratedData.TabIndex = 1
        Me.fraGeneratedData.TabStop = False
        Me.fraGeneratedData.Text = "Generated Data For Review"
        '
        'tbcSplitResults
        '
        Me.tbcSplitResults.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbcSplitResults.Controls.Add(Me.tbSplit)
        Me.tbcSplitResults.Location = New System.Drawing.Point(87, 22)
        Me.tbcSplitResults.Name = "tbcSplitResults"
        Me.tbcSplitResults.SelectedIndex = 0
        Me.tbcSplitResults.Size = New System.Drawing.Size(868, 235)
        Me.tbcSplitResults.TabIndex = 13
        '
        'tbSplit
        '
        Me.tbSplit.Controls.Add(Me.ssSplitReview)
        Me.tbSplit.Location = New System.Drawing.Point(4, 22)
        Me.tbSplit.Name = "tbSplit"
        Me.tbSplit.Padding = New System.Windows.Forms.Padding(3)
        Me.tbSplit.Size = New System.Drawing.Size(860, 209)
        Me.tbSplit.TabIndex = 0
        Me.tbSplit.Text = "Splits"
        Me.tbSplit.UseVisualStyleBackColor = True
        '
        'ssSplitReview
        '
        Me.ssSplitReview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ssSplitReview.Location = New System.Drawing.Point(3, 3)
        Me.ssSplitReview.Name = "ssSplitReview"
        Me.ssSplitReview.OcxState = CType(resources.GetObject("ssSplitReview.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssSplitReview.Size = New System.Drawing.Size(854, 203)
        Me.ssSplitReview.TabIndex = 10
        '
        'tbcHoleResults
        '
        Me.tbcHoleResults.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbcHoleResults.Controls.Add(Me.tbHole)
        Me.tbcHoleResults.Location = New System.Drawing.Point(87, 263)
        Me.tbcHoleResults.Name = "tbcHoleResults"
        Me.tbcHoleResults.SelectedIndex = 0
        Me.tbcHoleResults.Size = New System.Drawing.Size(868, 235)
        Me.tbcHoleResults.TabIndex = 13
        '
        'tbHole
        '
        Me.tbHole.Controls.Add(Me.ssCompReview)
        Me.tbHole.Location = New System.Drawing.Point(4, 22)
        Me.tbHole.Name = "tbHole"
        Me.tbHole.Padding = New System.Windows.Forms.Padding(3)
        Me.tbHole.Size = New System.Drawing.Size(860, 209)
        Me.tbHole.TabIndex = 0
        Me.tbHole.Text = "Holes"
        Me.tbHole.UseVisualStyleBackColor = True
        '
        'ssCompReview
        '
        Me.ssCompReview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ssCompReview.Location = New System.Drawing.Point(3, 3)
        Me.ssCompReview.Name = "ssCompReview"
        Me.ssCompReview.OcxState = CType(resources.GetObject("ssCompReview.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssCompReview.Size = New System.Drawing.Size(854, 203)
        Me.ssCompReview.TabIndex = 12
        '
        'ssCompErrors
        '
        Me.ssCompErrors.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ssCompErrors.Location = New System.Drawing.Point(87, 500)
        Me.ssCompErrors.Name = "ssCompErrors"
        Me.ssCompErrors.OcxState = CType(resources.GetObject("ssCompErrors.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssCompErrors.Size = New System.Drawing.Size(868, 82)
        Me.ssCompErrors.TabIndex = 11
        '
        'lblBarrenSplComm
        '
        Me.lblBarrenSplComm.AutoSize = True
        Me.lblBarrenSplComm.Location = New System.Drawing.Point(87, 234)
        Me.lblBarrenSplComm.Name = "lblBarrenSplComm"
        Me.lblBarrenSplComm.Size = New System.Drawing.Size(92, 13)
        Me.lblBarrenSplComm.TabIndex = 9
        Me.lblBarrenSplComm.Text = "lblBarrenSplComm"
        '
        'lblGen23
        '
        Me.lblGen23.AutoSize = True
        Me.lblGen23.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen23.Location = New System.Drawing.Point(24, 22)
        Me.lblGen23.Name = "lblGen23"
        Me.lblGen23.Size = New System.Drawing.Size(39, 16)
        Me.lblGen23.TabIndex = 8
        Me.lblGen23.Text = "Split"
        '
        'lblGen24
        '
        Me.lblGen24.AutoSize = True
        Me.lblGen24.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen24.Location = New System.Drawing.Point(24, 263)
        Me.lblGen24.Name = "lblGen24"
        Me.lblGen24.Size = New System.Drawing.Size(48, 16)
        Me.lblGen24.TabIndex = 7
        Me.lblGen24.Text = "Comp"
        '
        'lblGen25
        '
        Me.lblGen25.AutoSize = True
        Me.lblGen25.Location = New System.Drawing.Point(19, 310)
        Me.lblGen25.Name = "lblGen25"
        Me.lblGen25.Size = New System.Drawing.Size(30, 13)
        Me.lblGen25.TabIndex = 6
        Me.lblGen25.Text = "Data"
        '
        'lblGen26
        '
        Me.lblGen26.AutoSize = True
        Me.lblGen26.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen26.Location = New System.Drawing.Point(21, 500)
        Me.lblGen26.Name = "lblGen26"
        Me.lblGen26.Size = New System.Drawing.Size(53, 16)
        Me.lblGen26.TabIndex = 5
        Me.lblGen26.Text = "Issues"
        '
        'lblNoReview
        '
        Me.lblNoReview.AutoSize = True
        Me.lblNoReview.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNoReview.Location = New System.Drawing.Point(72, 54)
        Me.lblNoReview.Name = "lblNoReview"
        Me.lblNoReview.Size = New System.Drawing.Size(93, 20)
        Me.lblNoReview.TabIndex = 4
        Me.lblNoReview.Text = "No Review"
        Me.lblNoReview.Visible = False
        '
        'cmdCopyToOverrides
        '
        Me.cmdCopyToOverrides.Location = New System.Drawing.Point(15, 81)
        Me.cmdCopyToOverrides.Name = "cmdCopyToOverrides"
        Me.cmdCopyToOverrides.Size = New System.Drawing.Size(56, 61)
        Me.cmdCopyToOverrides.TabIndex = 0
        Me.cmdCopyToOverrides.Text = "Copy to Orides"
        Me.cmdCopyToOverrides.UseVisualStyleBackColor = True
        '
        'fraDetlDispl
        '
        Me.fraDetlDispl.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraDetlDispl.Controls.Add(Me.ssDetlDisp)
        Me.fraDetlDispl.Controls.Add(Me.lblGen41)
        Me.fraDetlDispl.Controls.Add(Me.lblGen64)
        Me.fraDetlDispl.Controls.Add(Me.cmdPrtGrdDetlDisp)
        Me.fraDetlDispl.Controls.Add(Me.fraResultCnt)
        Me.fraDetlDispl.Location = New System.Drawing.Point(970, 3)
        Me.fraDetlDispl.Name = "fraDetlDispl"
        Me.fraDetlDispl.Size = New System.Drawing.Size(344, 435)
        Me.fraDetlDispl.TabIndex = 0
        Me.fraDetlDispl.TabStop = False
        '
        'ssDetlDisp
        '
        Me.ssDetlDisp.Location = New System.Drawing.Point(14, 35)
        Me.ssDetlDisp.Name = "ssDetlDisp"
        Me.ssDetlDisp.OcxState = CType(resources.GetObject("ssDetlDisp.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssDetlDisp.Size = New System.Drawing.Size(248, 177)
        Me.ssDetlDisp.TabIndex = 6
        '
        'lblGen41
        '
        Me.lblGen41.AutoSize = True
        Me.lblGen41.Location = New System.Drawing.Point(14, 19)
        Me.lblGen41.Name = "lblGen41"
        Me.lblGen41.Size = New System.Drawing.Size(82, 13)
        Me.lblGen41.TabIndex = 5
        Me.lblGen41.Text = "lblReviewComm"
        Me.lblGen41.Visible = False
        '
        'lblGen64
        '
        Me.lblGen64.AutoSize = True
        Me.lblGen64.Location = New System.Drawing.Point(14, 465)
        Me.lblGen64.Name = "lblGen64"
        Me.lblGen64.Size = New System.Drawing.Size(39, 13)
        Me.lblGen64.TabIndex = 4
        Me.lblGen64.Text = "Label1"
        Me.lblGen64.Visible = False
        '
        'cmdPrtGrdDetlDisp
        '
        Me.cmdPrtGrdDetlDisp.Location = New System.Drawing.Point(187, 251)
        Me.cmdPrtGrdDetlDisp.Name = "cmdPrtGrdDetlDisp"
        Me.cmdPrtGrdDetlDisp.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrtGrdDetlDisp.TabIndex = 1
        Me.cmdPrtGrdDetlDisp.Text = "PrtGrd"
        Me.cmdPrtGrdDetlDisp.UseVisualStyleBackColor = True
        '
        'fraResultCnt
        '
        Me.fraResultCnt.Controls.Add(Me.ssResultCnt)
        Me.fraResultCnt.Controls.Add(Me.lblGen65)
        Me.fraResultCnt.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraResultCnt.Location = New System.Drawing.Point(17, 280)
        Me.fraResultCnt.Name = "fraResultCnt"
        Me.fraResultCnt.Size = New System.Drawing.Size(355, 182)
        Me.fraResultCnt.TabIndex = 0
        Me.fraResultCnt.TabStop = False
        Me.fraResultCnt.Text = "Results Count"
        '
        'ssResultCnt
        '
        Me.ssResultCnt.Location = New System.Drawing.Point(7, 30)
        Me.ssResultCnt.Name = "ssResultCnt"
        Me.ssResultCnt.OcxState = CType(resources.GetObject("ssResultCnt.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssResultCnt.Size = New System.Drawing.Size(232, 112)
        Me.ssResultCnt.TabIndex = 2
        '
        'lblGen65
        '
        Me.lblGen65.AutoSize = True
        Me.lblGen65.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen65.Location = New System.Drawing.Point(85, 12)
        Me.lblGen65.Name = "lblGen65"
        Me.lblGen65.Size = New System.Drawing.Size(39, 13)
        Me.lblGen65.TabIndex = 1
        Me.lblGen65.Text = "Label1"
        '
        'fraDataTypeOption
        '
        Me.fraDataTypeOption.Controls.Add(Me.lblRptAllCnt)
        Me.fraDataTypeOption.Controls.Add(Me.optProdCoeff)
        Me.fraDataTypeOption.Controls.Add(Me.opt100Pct)
        Me.fraDataTypeOption.Controls.Add(Me.cmdHoleSplitRpt)
        Me.fraDataTypeOption.Controls.Add(Me.cmdReportAll)
        Me.fraDataTypeOption.Controls.Add(Me.cmdAreaReport)
        Me.fraDataTypeOption.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraDataTypeOption.Location = New System.Drawing.Point(970, 444)
        Me.fraDataTypeOption.Name = "fraDataTypeOption"
        Me.fraDataTypeOption.Size = New System.Drawing.Size(312, 99)
        Me.fraDataTypeOption.TabIndex = 2
        Me.fraDataTypeOption.TabStop = False
        Me.fraDataTypeOption.Text = "Reports"
        '
        'lblRptAllCnt
        '
        Me.lblRptAllCnt.AutoSize = True
        Me.lblRptAllCnt.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRptAllCnt.Location = New System.Drawing.Point(212, 47)
        Me.lblRptAllCnt.Name = "lblRptAllCnt"
        Me.lblRptAllCnt.Size = New System.Drawing.Size(61, 13)
        Me.lblRptAllCnt.TabIndex = 5
        Me.lblRptAllCnt.Text = "lblRptAllCnt"
        Me.lblRptAllCnt.Visible = False
        '
        'optProdCoeff
        '
        Me.optProdCoeff.AutoSize = True
        Me.optProdCoeff.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.optProdCoeff.Location = New System.Drawing.Point(17, 19)
        Me.optProdCoeff.Name = "optProdCoeff"
        Me.optProdCoeff.Size = New System.Drawing.Size(114, 17)
        Me.optProdCoeff.TabIndex = 4
        Me.optProdCoeff.TabStop = True
        Me.optProdCoeff.Text = "Product coefficient"
        Me.optProdCoeff.UseVisualStyleBackColor = True
        '
        'opt100Pct
        '
        Me.opt100Pct.AutoSize = True
        Me.opt100Pct.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.opt100Pct.Location = New System.Drawing.Point(194, 19)
        Me.opt100Pct.Name = "opt100Pct"
        Me.opt100Pct.Size = New System.Drawing.Size(51, 17)
        Me.opt100Pct.TabIndex = 3
        Me.opt100Pct.TabStop = True
        Me.opt100Pct.Text = "100%"
        Me.opt100Pct.UseVisualStyleBackColor = True
        '
        'cmdHoleSplitRpt
        '
        Me.cmdHoleSplitRpt.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdHoleSplitRpt.Location = New System.Drawing.Point(17, 42)
        Me.cmdHoleSplitRpt.Name = "cmdHoleSplitRpt"
        Me.cmdHoleSplitRpt.Size = New System.Drawing.Size(75, 23)
        Me.cmdHoleSplitRpt.TabIndex = 2
        Me.cmdHoleSplitRpt.Text = "Report"
        Me.cmdHoleSplitRpt.UseVisualStyleBackColor = True
        '
        'cmdReportAll
        '
        Me.cmdReportAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReportAll.Location = New System.Drawing.Point(111, 71)
        Me.cmdReportAll.Name = "cmdReportAll"
        Me.cmdReportAll.Size = New System.Drawing.Size(75, 23)
        Me.cmdReportAll.TabIndex = 1
        Me.cmdReportAll.Text = "Report All"
        Me.cmdReportAll.UseVisualStyleBackColor = True
        '
        'cmdAreaReport
        '
        Me.cmdAreaReport.Enabled = False
        Me.cmdAreaReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAreaReport.Location = New System.Drawing.Point(111, 42)
        Me.cmdAreaReport.Name = "cmdAreaReport"
        Me.cmdAreaReport.Size = New System.Drawing.Size(75, 23)
        Me.cmdAreaReport.TabIndex = 0
        Me.cmdAreaReport.Text = "Area Report"
        Me.cmdAreaReport.UseVisualStyleBackColor = True
        '
        'fraReptDisp
        '
        Me.fraReptDisp.Controls.Add(Me.rtbRept1)
        Me.fraReptDisp.Controls.Add(Me.cmdExitRept)
        Me.fraReptDisp.Controls.Add(Me.cmdPrintRept)
        Me.fraReptDisp.Location = New System.Drawing.Point(913, 37)
        Me.fraReptDisp.Name = "fraReptDisp"
        Me.fraReptDisp.Size = New System.Drawing.Size(654, 496)
        Me.fraReptDisp.TabIndex = 6
        Me.fraReptDisp.TabStop = False
        Me.fraReptDisp.Visible = False
        '
        'rtbRept1
        '
        Me.rtbRept1.Location = New System.Drawing.Point(23, 29)
        Me.rtbRept1.Name = "rtbRept1"
        Me.rtbRept1.Size = New System.Drawing.Size(614, 414)
        Me.rtbRept1.TabIndex = 2
        Me.rtbRept1.Text = ""
        '
        'cmdExitRept
        '
        Me.cmdExitRept.Location = New System.Drawing.Point(510, 460)
        Me.cmdExitRept.Name = "cmdExitRept"
        Me.cmdExitRept.Size = New System.Drawing.Size(127, 23)
        Me.cmdExitRept.TabIndex = 1
        Me.cmdExitRept.Text = "Exit Report"
        Me.cmdExitRept.UseVisualStyleBackColor = True
        '
        'cmdPrintRept
        '
        Me.cmdPrintRept.Location = New System.Drawing.Point(23, 460)
        Me.cmdPrintRept.Name = "cmdPrintRept"
        Me.cmdPrintRept.Size = New System.Drawing.Size(138, 23)
        Me.cmdPrintRept.TabIndex = 0
        Me.cmdPrintRept.Text = "Print Report"
        Me.cmdPrintRept.UseVisualStyleBackColor = True
        '
        'cmdReport
        '
        Me.cmdReport.Location = New System.Drawing.Point(871, 712)
        Me.cmdReport.Name = "cmdReport"
        Me.cmdReport.Size = New System.Drawing.Size(56, 23)
        Me.cmdReport.TabIndex = 0
        Me.cmdReport.Text = "Reports"
        Me.cmdReport.UseVisualStyleBackColor = True
        Me.cmdReport.Visible = False
        '
        'cmdGenerateProspectDataset
        '
        Me.cmdGenerateProspectDataset.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdGenerateProspectDataset.Location = New System.Drawing.Point(436, 607)
        Me.cmdGenerateProspectDataset.Name = "cmdGenerateProspectDataset"
        Me.cmdGenerateProspectDataset.Size = New System.Drawing.Size(144, 35)
        Me.cmdGenerateProspectDataset.TabIndex = 11
        Me.cmdGenerateProspectDataset.Text = "Generate Prospect Dataset (Split Data for Review)"
        Me.cmdGenerateProspectDataset.UseVisualStyleBackColor = True
        '
        'cmdCancelProspectDataset
        '
        Me.cmdCancelProspectDataset.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdCancelProspectDataset.Location = New System.Drawing.Point(742, 607)
        Me.cmdCancelProspectDataset.Name = "cmdCancelProspectDataset"
        Me.cmdCancelProspectDataset.Size = New System.Drawing.Size(150, 35)
        Me.cmdCancelProspectDataset.TabIndex = 9
        Me.cmdCancelProspectDataset.Text = "Cancel Reduction"
        Me.cmdCancelProspectDataset.UseVisualStyleBackColor = True
        '
        'chkCreateOutputOnly
        '
        Me.chkCreateOutputOnly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkCreateOutputOnly.AutoSize = True
        Me.chkCreateOutputOnly.Location = New System.Drawing.Point(263, 609)
        Me.chkCreateOutputOnly.Name = "chkCreateOutputOnly"
        Me.chkCreateOutputOnly.Size = New System.Drawing.Size(167, 17)
        Me.chkCreateOutputOnly.TabIndex = 8
        Me.chkCreateOutputOnly.Text = "Create output only (no review)"
        Me.chkCreateOutputOnly.UseVisualStyleBackColor = True
        '
        'cmdSaveProspectDataset
        '
        Me.cmdSaveProspectDataset.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdSaveProspectDataset.Location = New System.Drawing.Point(586, 607)
        Me.cmdSaveProspectDataset.Name = "cmdSaveProspectDataset"
        Me.cmdSaveProspectDataset.Size = New System.Drawing.Size(150, 35)
        Me.cmdSaveProspectDataset.TabIndex = 7
        Me.cmdSaveProspectDataset.Text = "Save Prospect Dataset  (Split &&/or Composite Data)"
        Me.cmdSaveProspectDataset.UseVisualStyleBackColor = True
        '
        'pnlContent
        '
        Me.pnlContent.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlContent.Controls.Add(Me.sbrMain)
        Me.pnlContent.Controls.Add(Me.fraDataReduction)
        Me.pnlContent.Location = New System.Drawing.Point(0, 0)
        Me.pnlContent.Name = "pnlContent"
        Me.pnlContent.Size = New System.Drawing.Size(1340, 667)
        Me.pnlContent.TabIndex = 19
        '
        'sbrMain
        '
        Me.sbrMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblStatus, Me.lblProcComm0, Me.lblProcComm1, Me.lblProcComm2})
        Me.sbrMain.Location = New System.Drawing.Point(0, 645)
        Me.sbrMain.Name = "sbrMain"
        Me.sbrMain.Size = New System.Drawing.Size(1340, 22)
        Me.sbrMain.TabIndex = 6
        Me.sbrMain.Text = "StatusStrip1"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = False
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(400, 17)
        Me.lblStatus.Text = "Generating prospect data set (split data for review)..."
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblProcComm0
        '
        Me.lblProcComm0.AutoSize = False
        Me.lblProcComm0.ForeColor = System.Drawing.Color.DarkBlue
        Me.lblProcComm0.Name = "lblProcComm0"
        Me.lblProcComm0.Size = New System.Drawing.Size(300, 17)
        Me.lblProcComm0.Text = "Getting Raw Prosp from Database"
        Me.lblProcComm0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblProcComm1
        '
        Me.lblProcComm1.AutoSize = False
        Me.lblProcComm1.ForeColor = System.Drawing.Color.DarkBlue
        Me.lblProcComm1.Name = "lblProcComm1"
        Me.lblProcComm1.Size = New System.Drawing.Size(250, 17)
        Me.lblProcComm1.Text = "1000 SFC items to process..."
        Me.lblProcComm1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblProcComm2
        '
        Me.lblProcComm2.AutoSize = False
        Me.lblProcComm2.ForeColor = System.Drawing.Color.DarkBlue
        Me.lblProcComm2.Name = "lblProcComm2"
        Me.lblProcComm2.Size = New System.Drawing.Size(350, 17)
        Me.lblProcComm2.Text = "Splits processed = 1000"
        Me.lblProcComm2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'frmProspDataReduction
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(1350, 668)
        Me.Controls.Add(Me.pnlContent)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MinimumSize = New System.Drawing.Size(1350, 660)
        Me.Name = "frmProspDataReduction"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Mining OIS -- Prospect Data Reduction  (Create Splits and Composites)"
        Me.fraDataReduction.ResumeLayout(False)
        Me.fraDataReduction.PerformLayout()
        Me.tabMain.ResumeLayout(False)
        Me.TabPage9.ResumeLayout(False)
        Me.fraOffSpecPb.ResumeLayout(False)
        Me.fraOrigMgoPlant.ResumeLayout(False)
        Me.fraOrigMgoPlant.PerformLayout()
        CType(Me.ssOffSpecPb, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraDoloflotPlant.ResumeLayout(False)
        Me.fraDoloflotPlant.PerformLayout()
        CType(Me.ssDoloflotPlant, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraFcoDoloflot.ResumeLayout(False)
        Me.fraFcoDoloflot.PerformLayout()
        CType(Me.ssDoloflotPlantFco2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ssDoloflotPlantFco, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage10.ResumeLayout(False)
        Me.fraRptAllToTextFile.ResumeLayout(False)
        Me.fraRptAllToTextFile.PerformLayout()
        Me.fraDataSaveComm.ResumeLayout(False)
        Me.fraOutputLocation2.ResumeLayout(False)
        Me.fraOutputLocation2.PerformLayout()
        Me.fraOutputLocation1.ResumeLayout(False)
        Me.fraOutputLocation1.PerformLayout()
        Me.fraProspDatasetType.ResumeLayout(False)
        Me.fraProspDatasetType.PerformLayout()
        Me.TabPage11.ResumeLayout(False)
        Me.fraOverride.ResumeLayout(False)
        Me.fraSplitOverride.ResumeLayout(False)
        Me.fraSplitOverride.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        CType(Me.ssSplitOverrides, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraOverrideList.ResumeLayout(False)
        Me.fraOverrideList.PerformLayout()
        CType(Me.ssSplitOverride, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ssRawProspMin, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabReview.ResumeLayout(False)
        Me.fraReview.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.fraGeneratedData.ResumeLayout(False)
        Me.fraGeneratedData.PerformLayout()
        Me.tbcSplitResults.ResumeLayout(False)
        Me.tbSplit.ResumeLayout(False)
        CType(Me.ssSplitReview, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbcHoleResults.ResumeLayout(False)
        Me.tbHole.ResumeLayout(False)
        CType(Me.ssCompReview, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ssCompErrors, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraDetlDispl.ResumeLayout(False)
        Me.fraDetlDispl.PerformLayout()
        CType(Me.ssDetlDisp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraResultCnt.ResumeLayout(False)
        Me.fraResultCnt.PerformLayout()
        CType(Me.ssResultCnt, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraDataTypeOption.ResumeLayout(False)
        Me.fraDataTypeOption.PerformLayout()
        Me.fraReptDisp.ResumeLayout(False)
        Me.pnlContent.ResumeLayout(False)
        Me.pnlContent.PerformLayout()
        Me.sbrMain.ResumeLayout(False)
        Me.sbrMain.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents fraDataReduction As System.Windows.Forms.GroupBox
    Friend WithEvents fraRptAllToTextFile As System.Windows.Forms.GroupBox
    Friend WithEvents cmdPrintRptAll As System.Windows.Forms.Button
    Friend WithEvents txtRptAllToTxtFile As System.Windows.Forms.TextBox
    Friend WithEvents cmdRptAllToTxtFile As System.Windows.Forms.Button
    Friend WithEvents lblRptAllCnt2 As System.Windows.Forms.Label
    Friend WithEvents fraDataSaveComm As System.Windows.Forms.GroupBox
    Friend WithEvents lblGen53 As System.Windows.Forms.Label
    Friend WithEvents fraOutputLocation2 As System.Windows.Forms.GroupBox
    Friend WithEvents chkSurvCaddTextfile As System.Windows.Forms.CheckBox
    Friend WithEvents chkSpecMoisTransferFile As System.Windows.Forms.CheckBox
    Friend WithEvents chkInclMgPlt As System.Windows.Forms.CheckBox
    Friend WithEvents chkPbAnalysisFillInSpecial As System.Windows.Forms.CheckBox
    Friend WithEvents chkBdFormatTextfile As System.Windows.Forms.CheckBox
    Friend WithEvents txtProspDatasetTextfileName As System.Windows.Forms.TextBox
    Friend WithEvents lblTextfileComment As System.Windows.Forms.Label
    Friend WithEvents lblGen10 As System.Windows.Forms.Label
    Friend WithEvents lblGen37 As System.Windows.Forms.Label
    Friend WithEvents fraOutputLocation1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtProspectDatasetDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtProspectDatasetName As System.Windows.Forms.TextBox
    Friend WithEvents chkSaveToDatabase As System.Windows.Forms.CheckBox
    Friend WithEvents lblGen11 As System.Windows.Forms.Label
    Friend WithEvents lblGen9 As System.Windows.Forms.Label
    Friend WithEvents fraProspDatasetType As System.Windows.Forms.GroupBox
    Friend WithEvents chkProductionCoefficient As System.Windows.Forms.CheckBox
    Friend WithEvents optInclSplits As System.Windows.Forms.RadioButton
    Friend WithEvents optInclComposites As System.Windows.Forms.RadioButton
    Friend WithEvents optInclBoth As System.Windows.Forms.RadioButton
    Friend WithEvents chk100Pct As System.Windows.Forms.CheckBox
    Friend WithEvents fraOffSpecPb As System.Windows.Forms.GroupBox
    Friend WithEvents fraFcoDoloflot As System.Windows.Forms.GroupBox
    Friend WithEvents cmdSetDefaults2 As System.Windows.Forms.Button
    Friend WithEvents chkUseDoloflotPlantFco As System.Windows.Forms.CheckBox
    Friend WithEvents lblGen51 As System.Windows.Forms.Label
    Friend WithEvents fraDoloflotPlant As System.Windows.Forms.GroupBox
    Friend WithEvents chkUseDoloflotPlant As System.Windows.Forms.CheckBox
    Friend WithEvents cmdSetDefaults As System.Windows.Forms.Button
    Friend WithEvents lblGen45 As System.Windows.Forms.Label
    Friend WithEvents lblGen44 As System.Windows.Forms.Label
    Friend WithEvents fraOrigMgoPlant As System.Windows.Forms.GroupBox
    Friend WithEvents chkUseOrigMgoPlant As System.Windows.Forms.CheckBox
    Friend WithEvents lblGen31 As System.Windows.Forms.Label
    Friend WithEvents fraReview As System.Windows.Forms.GroupBox
    Friend WithEvents fraDetlDispl As System.Windows.Forms.GroupBox
    Friend WithEvents fraResultCnt As System.Windows.Forms.GroupBox
    Friend WithEvents lblGen65 As System.Windows.Forms.Label
    Friend WithEvents cmdPrtGrdDetlDisp As System.Windows.Forms.Button
    Friend WithEvents fraDataTypeOption As System.Windows.Forms.GroupBox
    Friend WithEvents cmdHoleSplitRpt As System.Windows.Forms.Button
    Friend WithEvents cmdReportAll As System.Windows.Forms.Button
    Friend WithEvents cmdAreaReport As System.Windows.Forms.Button
    Friend WithEvents lblRptAllCnt As System.Windows.Forms.Label
    Friend WithEvents optProdCoeff As System.Windows.Forms.RadioButton
    Friend WithEvents opt100Pct As System.Windows.Forms.RadioButton
    Friend WithEvents lblGen41 As System.Windows.Forms.Label
    Friend WithEvents lblGen64 As System.Windows.Forms.Label
    Friend WithEvents fraGeneratedData As System.Windows.Forms.GroupBox
    Friend WithEvents cmdCopyToOverrides As System.Windows.Forms.Button
    Friend WithEvents lblNoReview As System.Windows.Forms.Label
    Friend WithEvents lblGen26 As System.Windows.Forms.Label
    Friend WithEvents lblGen25 As System.Windows.Forms.Label
    Friend WithEvents lblGen24 As System.Windows.Forms.Label
    Friend WithEvents lblGen23 As System.Windows.Forms.Label
    Friend WithEvents lblBarrenSplComm As System.Windows.Forms.Label
    Friend WithEvents tabMain As System.Windows.Forms.TabControl
    Friend WithEvents tabAreaDef As System.Windows.Forms.TabPage
    Friend WithEvents tabProductSizes As System.Windows.Forms.TabPage
    Friend WithEvents tabRecoveryAndMineability As System.Windows.Forms.TabPage
    Friend WithEvents TabPage9 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage10 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage11 As System.Windows.Forms.TabPage
    Friend WithEvents tabReview As System.Windows.Forms.TabPage
    Friend WithEvents fraOverride As System.Windows.Forms.GroupBox
    Friend WithEvents fraOverrideList As System.Windows.Forms.GroupBox
    Friend WithEvents chkUseRawProspAsOverride As System.Windows.Forms.CheckBox
    Friend WithEvents cmdApplySplOverrides As System.Windows.Forms.Button
    Friend WithEvents cmdClrOverride As System.Windows.Forms.Button
    Friend WithEvents lblGen54 As System.Windows.Forms.Label
    Friend WithEvents lblGen55 As System.Windows.Forms.Label
    Friend WithEvents lblGen33 As System.Windows.Forms.Label
    Friend WithEvents lblGen34 As System.Windows.Forms.Label
    Friend WithEvents lblGen36 As System.Windows.Forms.Label
    Friend WithEvents fraSplitOverride As System.Windows.Forms.GroupBox
    Friend WithEvents cmdLoadOverrideTxtFile As System.Windows.Forms.Button
    Friend WithEvents txtSplOverrideTxtFile As System.Windows.Forms.TextBox
    Friend WithEvents cboSplitOverrideMineName As System.Windows.Forms.ComboBox
    Friend WithEvents cmdCancelSplitOverride As System.Windows.Forms.Button
    Friend WithEvents Frame2 As System.Windows.Forms.GroupBox
    Friend WithEvents chkOnlyMySplitOverride As System.Windows.Forms.CheckBox
    Friend WithEvents cmdGetSplitOverrides As System.Windows.Forms.Button
    Friend WithEvents txtSplitOverrideName As System.Windows.Forms.TextBox
    Friend WithEvents cmdDeleteSplitOverride As System.Windows.Forms.Button
    Friend WithEvents cmdSaveSplitOverride As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblSplitOverrideTxtFile As System.Windows.Forms.Label
    Friend WithEvents cmdSaveProspectDataset As System.Windows.Forms.Button
    Friend WithEvents chkCreateOutputOnly As System.Windows.Forms.CheckBox
    Friend WithEvents cmdCancelProspectDataset As System.Windows.Forms.Button
    Friend WithEvents cmdGenerateProspectDataset As System.Windows.Forms.Button
    Friend WithEvents cmdReport As System.Windows.Forms.Button
    Friend WithEvents cmdExitForm As System.Windows.Forms.Button
    Friend WithEvents ssDoloflotPlantFco As AxFPSpread.AxvaSpread
    Friend WithEvents ssDoloflotPlantFco2 As AxFPSpread.AxvaSpread
    Friend WithEvents ssOffSpecPb As AxFPSpread.AxvaSpread
    Friend WithEvents ssRawProspMin As AxFPSpread.AxvaSpread
    Friend WithEvents ssSplitOverride As AxFPSpread.AxvaSpread
    Friend WithEvents ssSplitOverrides As AxFPSpread.AxvaSpread
    Friend WithEvents ssResultCnt As AxFPSpread.AxvaSpread
    Friend WithEvents ssDetlDisp As AxFPSpread.AxvaSpread
    Friend WithEvents ssSplitReview As AxFPSpread.AxvaSpread
    Friend WithEvents ssCompReview As AxFPSpread.AxvaSpread
    Public WithEvents ssCompErrors As AxFPSpread.AxvaSpread
    Friend WithEvents ssDoloflotPlant As AxFPSpread.AxvaSpread
    Friend WithEvents fraExtra As System.Windows.Forms.GroupBox
    Friend WithEvents fraReptDisp As System.Windows.Forms.GroupBox
    Friend WithEvents rtbRept1 As System.Windows.Forms.RichTextBox
    Friend WithEvents cmdExitRept As System.Windows.Forms.Button
    Friend WithEvents cmdPrintRept As System.Windows.Forms.Button
    Friend WithEvents pnlContent As Panel
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents cmdPrtScr As Button
    Friend WithEvents sbrMain As StatusStrip
    Friend WithEvents lblStatus As ToolStripStatusLabel
    Friend WithEvents lblProcComm0 As ToolStripStatusLabel
    Friend WithEvents lblProcComm1 As ToolStripStatusLabel
    Friend WithEvents lblProcComm2 As ToolStripStatusLabel
    Friend WithEvents lblProspectDatasetStatus As Label
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents tbcSplitResults As TabControl
    Friend WithEvents tbSplit As TabPage
    Friend WithEvents tbcHoleResults As TabControl
    Friend WithEvents tbHole As TabPage
End Class
