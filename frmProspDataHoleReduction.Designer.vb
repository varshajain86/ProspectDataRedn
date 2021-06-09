<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProspDataHoleReduction
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProspDataHoleReduction))
        Me.fraResults = New System.Windows.Forms.GroupBox()
        Me.cmdProspSec = New System.Windows.Forms.Button()
        Me.cmdViewCompSplit = New System.Windows.Forms.Button()
        Me.cmdViewRawProsp = New System.Windows.Forms.Button()
        Me.cmdMakeHoleUnmineable = New System.Windows.Forms.Button()
        Me.cmdPrintSplit = New System.Windows.Forms.Button()
        Me.cmdPrintHole = New System.Windows.Forms.Button()
        Me.fraMode = New System.Windows.Forms.GroupBox()
        Me.optCatalog = New System.Windows.Forms.RadioButton()
        Me.opt100Pct = New System.Windows.Forms.RadioButton()
        Me.lblGoTo = New System.Windows.Forms.Label()
        Me.lblMaxDepthComm = New System.Windows.Forms.Label()
        Me.lblMiscComm2 = New System.Windows.Forms.Label()
        Me.lblMiscComm = New System.Windows.Forms.Label()
        Me.lblCoordsElev = New System.Windows.Forms.Label()
        Me.lblOvbComm = New System.Windows.Forms.Label()
        Me.ssDrillData = New AxFPSpread.AxvaSpread()
        Me.tabDisp = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.lblHole = New System.Windows.Forms.Label()
        Me.lblUserMadeHoleUnmineable = New System.Windows.Forms.Label()
        Me.ssHoleData = New AxFPSpread.AxvaSpread()
        Me.fraMiscStuff = New System.Windows.Forms.GroupBox()
        Me.ssRawProspMin = New AxFPSpread.AxvaSpread()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.lblSplit = New System.Windows.Forms.Label()
        Me.lblCurrSplit = New System.Windows.Forms.Label()
        Me.ssSplitData = New AxFPSpread.AxvaSpread()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.tblReductionData = New System.Windows.Forms.TableLayoutPanel()
        Me.lblRdctnHole = New System.Windows.Forms.Label()
        Me.lblRdctnSplit = New System.Windows.Forms.Label()
        Me.ssSplitReview = New AxFPSpread.AxvaSpread()
        Me.ssCompReview = New AxFPSpread.AxvaSpread()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.lblHoleInMOIS = New System.Windows.Forms.Label()
        Me.lblHoleExistStatus = New System.Windows.Forms.Label()
        Me.ssHoleExistStatus = New AxFPSpread.AxvaSpread()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblInfoComm = New System.Windows.Forms.Label()
        Me.TabPage6 = New System.Windows.Forms.TabPage()
        Me.cmdPrtGrd = New System.Windows.Forms.Button()
        Me.ssCompErrors = New AxFPSpread.AxvaSpread()
        Me.TabPage7 = New System.Windows.Forms.TabPage()
        Me.lblAreaName = New System.Windows.Forms.Label()
        Me.txtAreaName = New System.Windows.Forms.TextBox()
        Me.cmdSaveAreaName = New System.Windows.Forms.Button()
        Me.TabPage8 = New System.Windows.Forms.TabPage()
        Me.lblCurrMinabilityComm = New System.Windows.Forms.Label()
        Me.lblSplitMinabilities = New System.Windows.Forms.Label()
        Me.lblHoleMinability = New System.Windows.Forms.Label()
        Me.ssSplitMinabilities = New AxFPSpread.AxvaSpread()
        Me.ssHoleMinabilities = New AxFPSpread.AxvaSpread()
        Me.TabPage9 = New System.Windows.Forms.TabPage()
        Me.lblFeAdjComm = New System.Windows.Forms.Label()
        Me.chkUseFeAdjust = New System.Windows.Forms.CheckBox()
        Me.ssFeAdjustment = New AxFPSpread.AxvaSpread()
        Me.TabPage10 = New System.Windows.Forms.TabPage()
        Me.lblSplitOverrideSet = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtSplitOverrideName = New System.Windows.Forms.TextBox()
        Me.cmdAddToOverrideSet = New System.Windows.Forms.Button()
        Me.ssSplitOverride = New AxFPSpread.AxvaSpread()
        Me.cmdRefresh = New System.Windows.Forms.Button()
        Me.cboSplitOverrideMineName = New System.Windows.Forms.ComboBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.ssSplitOverrides = New AxFPSpread.AxvaSpread()
        Me.chkOnlyMySplitOverride = New System.Windows.Forms.CheckBox()
        Me.cmdGetSplitOverrides = New System.Windows.Forms.Button()
        Me.fraSelect = New System.Windows.Forms.GroupBox()
        Me.lblAlphaNumeric = New System.Windows.Forms.Label()
        Me.lblOffSpecPbMgPlt = New System.Windows.Forms.Label()
        Me.lblHoleInMoisComm = New System.Windows.Forms.Label()
        Me.lblScenComm = New System.Windows.Forms.Label()
        Me.lblHoleLocation = New System.Windows.Forms.Label()
        Me.lblRange = New System.Windows.Forms.Label()
        Me.lblSection = New System.Windows.Forms.Label()
        Me.lblTownship = New System.Windows.Forms.Label()
        Me.cboHole = New System.Windows.Forms.ComboBox()
        Me.cboRge = New System.Windows.Forms.ComboBox()
        Me.cboSec = New System.Windows.Forms.ComboBox()
        Me.cboTwp = New System.Windows.Forms.ComboBox()
        Me.lblOtherDefn = New System.Windows.Forms.Label()
        Me.lblIpComm = New System.Windows.Forms.Label()
        Me.lblUseOrigHoleComm = New System.Windows.Forms.Label()
        Me.cmdReduceHole = New System.Windows.Forms.Button()
        Me.chkMyParams = New System.Windows.Forms.CheckBox()
        Me.cmdRefreshParams = New System.Windows.Forms.Button()
        Me.cboOtherDefn = New System.Windows.Forms.ComboBox()
        Me.cboProdSizeDefn = New System.Windows.Forms.ComboBox()
        Me.lblProdSizeDefn = New System.Windows.Forms.Label()
        Me.cmdTest3 = New System.Windows.Forms.Button()
        Me.cmdTest2 = New System.Windows.Forms.Button()
        Me.cmdTest = New System.Windows.Forms.Button()
        Me.chkOverrideMaxDepth = New System.Windows.Forms.CheckBox()
        Me.chkUseOrigHole = New System.Windows.Forms.CheckBox()
        Me.cmdPrtScr = New System.Windows.Forms.Button()
        Me.sbrMain = New System.Windows.Forms.StatusStrip()
        Me.fraSaveToMois = New System.Windows.Forms.GroupBox()
        Me.cmdSaveMinabilities = New System.Windows.Forms.Button()
        Me.fraRdctnType = New System.Windows.Forms.GroupBox()
        Me.lblSaveToMoisComm = New System.Windows.Forms.Label()
        Me.cmdSaveCompAndSplits = New System.Windows.Forms.Button()
        Me.optBothRdctn = New System.Windows.Forms.RadioButton()
        Me.optCatalogRdctn = New System.Windows.Forms.RadioButton()
        Me.opt100PctRdctn = New System.Windows.Forms.RadioButton()
        Me.chkSaveRawProspectMinabilities = New System.Windows.Forms.CheckBox()
        Me.cboMineName = New System.Windows.Forms.ComboBox()
        Me.fraSurvCadd = New System.Windows.Forms.GroupBox()
        Me.cmdCreateSurvCadd = New System.Windows.Forms.Button()
        Me.lblSurvCaddComm = New System.Windows.Forms.Label()
        Me.chkPbAnalysisFillInSpecial = New System.Windows.Forms.CheckBox()
        Me.optInclBoth = New System.Windows.Forms.RadioButton()
        Me.optInclSplits = New System.Windows.Forms.RadioButton()
        Me.optInclComposites = New System.Windows.Forms.RadioButton()
        Me.txtSurvCaddTextfile = New System.Windows.Forms.TextBox()
        Me.lblSurvCaddTxtFile = New System.Windows.Forms.Label()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.fraResults.SuspendLayout()
        Me.fraMode.SuspendLayout()
        CType(Me.ssDrillData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabDisp.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.ssHoleData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraMiscStuff.SuspendLayout()
        CType(Me.ssRawProspMin, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.ssSplitData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        Me.tblReductionData.SuspendLayout()
        CType(Me.ssSplitReview, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ssCompReview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.ssHoleExistStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        Me.TabPage6.SuspendLayout()
        CType(Me.ssCompErrors, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage7.SuspendLayout()
        Me.TabPage8.SuspendLayout()
        CType(Me.ssSplitMinabilities, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ssHoleMinabilities, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage9.SuspendLayout()
        CType(Me.ssFeAdjustment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage10.SuspendLayout()
        CType(Me.ssSplitOverride, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame2.SuspendLayout()
        CType(Me.ssSplitOverrides, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraSelect.SuspendLayout()
        Me.fraSaveToMois.SuspendLayout()
        Me.fraRdctnType.SuspendLayout()
        Me.fraSurvCadd.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraResults
        '
        Me.fraResults.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraResults.Controls.Add(Me.cmdProspSec)
        Me.fraResults.Controls.Add(Me.cmdViewCompSplit)
        Me.fraResults.Controls.Add(Me.cmdViewRawProsp)
        Me.fraResults.Controls.Add(Me.cmdMakeHoleUnmineable)
        Me.fraResults.Controls.Add(Me.cmdPrintSplit)
        Me.fraResults.Controls.Add(Me.cmdPrintHole)
        Me.fraResults.Controls.Add(Me.fraMode)
        Me.fraResults.Controls.Add(Me.lblGoTo)
        Me.fraResults.Controls.Add(Me.lblMaxDepthComm)
        Me.fraResults.Controls.Add(Me.lblMiscComm2)
        Me.fraResults.Controls.Add(Me.lblMiscComm)
        Me.fraResults.Controls.Add(Me.lblCoordsElev)
        Me.fraResults.Controls.Add(Me.lblOvbComm)
        Me.fraResults.Controls.Add(Me.ssDrillData)
        Me.fraResults.Controls.Add(Me.tabDisp)
        Me.fraResults.Location = New System.Drawing.Point(40, 135)
        Me.fraResults.Name = "fraResults"
        Me.fraResults.Size = New System.Drawing.Size(1132, 461)
        Me.fraResults.TabIndex = 0
        Me.fraResults.TabStop = False
        Me.fraResults.Text = "Select split minabilities"
        '
        'cmdProspSec
        '
        Me.cmdProspSec.Location = New System.Drawing.Point(16, 217)
        Me.cmdProspSec.Name = "cmdProspSec"
        Me.cmdProspSec.Size = New System.Drawing.Size(75, 23)
        Me.cmdProspSec.TabIndex = 15
        Me.cmdProspSec.Text = "Prosp Sec"
        Me.cmdProspSec.UseVisualStyleBackColor = True
        Me.cmdProspSec.Visible = False
        '
        'cmdViewCompSplit
        '
        Me.cmdViewCompSplit.Location = New System.Drawing.Point(16, 188)
        Me.cmdViewCompSplit.Name = "cmdViewCompSplit"
        Me.cmdViewCompSplit.Size = New System.Drawing.Size(75, 23)
        Me.cmdViewCompSplit.TabIndex = 14
        Me.cmdViewCompSplit.Text = "Comp/Splits"
        Me.cmdViewCompSplit.UseVisualStyleBackColor = True
        Me.cmdViewCompSplit.Visible = False
        '
        'cmdViewRawProsp
        '
        Me.cmdViewRawProsp.Location = New System.Drawing.Point(16, 158)
        Me.cmdViewRawProsp.Name = "cmdViewRawProsp"
        Me.cmdViewRawProsp.Size = New System.Drawing.Size(75, 23)
        Me.cmdViewRawProsp.TabIndex = 13
        Me.cmdViewRawProsp.Text = "Raw Prosp"
        Me.cmdViewRawProsp.UseVisualStyleBackColor = True
        Me.cmdViewRawProsp.Visible = False
        '
        'cmdMakeHoleUnmineable
        '
        Me.cmdMakeHoleUnmineable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdMakeHoleUnmineable.Location = New System.Drawing.Point(932, 227)
        Me.cmdMakeHoleUnmineable.Name = "cmdMakeHoleUnmineable"
        Me.cmdMakeHoleUnmineable.Size = New System.Drawing.Size(184, 23)
        Me.cmdMakeHoleUnmineable.TabIndex = 12
        Me.cmdMakeHoleUnmineable.Text = "Make Hole Unmineable"
        Me.cmdMakeHoleUnmineable.UseVisualStyleBackColor = True
        '
        'cmdPrintSplit
        '
        Me.cmdPrintSplit.Enabled = False
        Me.cmdPrintSplit.Location = New System.Drawing.Point(16, 111)
        Me.cmdPrintSplit.Name = "cmdPrintSplit"
        Me.cmdPrintSplit.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrintSplit.TabIndex = 11
        Me.cmdPrintSplit.Text = "Print Split"
        Me.cmdPrintSplit.UseVisualStyleBackColor = True
        '
        'cmdPrintHole
        '
        Me.cmdPrintHole.Enabled = False
        Me.cmdPrintHole.Location = New System.Drawing.Point(16, 84)
        Me.cmdPrintHole.Name = "cmdPrintHole"
        Me.cmdPrintHole.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrintHole.TabIndex = 10
        Me.cmdPrintHole.Text = "Print Hole"
        Me.cmdPrintHole.UseVisualStyleBackColor = True
        '
        'fraMode
        '
        Me.fraMode.Controls.Add(Me.optCatalog)
        Me.fraMode.Controls.Add(Me.opt100Pct)
        Me.fraMode.Location = New System.Drawing.Point(16, 20)
        Me.fraMode.Name = "fraMode"
        Me.fraMode.Size = New System.Drawing.Size(75, 58)
        Me.fraMode.TabIndex = 9
        Me.fraMode.TabStop = False
        '
        'optCatalog
        '
        Me.optCatalog.AutoSize = True
        Me.optCatalog.Location = New System.Drawing.Point(7, 35)
        Me.optCatalog.Name = "optCatalog"
        Me.optCatalog.Size = New System.Drawing.Size(61, 17)
        Me.optCatalog.TabIndex = 1
        Me.optCatalog.TabStop = True
        Me.optCatalog.Text = "Catalog"
        Me.optCatalog.UseVisualStyleBackColor = True
        '
        'opt100Pct
        '
        Me.opt100Pct.AutoSize = True
        Me.opt100Pct.Location = New System.Drawing.Point(7, 11)
        Me.opt100Pct.Name = "opt100Pct"
        Me.opt100Pct.Size = New System.Drawing.Size(51, 17)
        Me.opt100Pct.TabIndex = 0
        Me.opt100Pct.TabStop = True
        Me.opt100Pct.Text = "100%"
        Me.opt100Pct.UseVisualStyleBackColor = True
        '
        'lblGoTo
        '
        Me.lblGoTo.AutoSize = True
        Me.lblGoTo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGoTo.Location = New System.Drawing.Point(32, 142)
        Me.lblGoTo.Name = "lblGoTo"
        Me.lblGoTo.Size = New System.Drawing.Size(42, 13)
        Me.lblGoTo.TabIndex = 8
        Me.lblGoTo.Text = "Go To"
        Me.lblGoTo.Visible = False
        '
        'lblMaxDepthComm
        '
        Me.lblMaxDepthComm.AutoSize = True
        Me.lblMaxDepthComm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaxDepthComm.ForeColor = System.Drawing.Color.Red
        Me.lblMaxDepthComm.Location = New System.Drawing.Point(130, 53)
        Me.lblMaxDepthComm.Name = "lblMaxDepthComm"
        Me.lblMaxDepthComm.Size = New System.Drawing.Size(45, 13)
        Me.lblMaxDepthComm.TabIndex = 7
        Me.lblMaxDepthComm.Text = "Label6"
        '
        'lblMiscComm2
        '
        Me.lblMiscComm2.AutoSize = True
        Me.lblMiscComm2.ForeColor = System.Drawing.Color.Navy
        Me.lblMiscComm2.Location = New System.Drawing.Point(456, 34)
        Me.lblMiscComm2.Name = "lblMiscComm2"
        Me.lblMiscComm2.Size = New System.Drawing.Size(39, 13)
        Me.lblMiscComm2.TabIndex = 6
        Me.lblMiscComm2.Text = "Label5"
        '
        'lblMiscComm
        '
        Me.lblMiscComm.AutoSize = True
        Me.lblMiscComm.Location = New System.Drawing.Point(647, 20)
        Me.lblMiscComm.Name = "lblMiscComm"
        Me.lblMiscComm.Size = New System.Drawing.Size(39, 13)
        Me.lblMiscComm.TabIndex = 5
        Me.lblMiscComm.Text = "Label2"
        '
        'lblCoordsElev
        '
        Me.lblCoordsElev.AutoSize = True
        Me.lblCoordsElev.Location = New System.Drawing.Point(130, 18)
        Me.lblCoordsElev.Name = "lblCoordsElev"
        Me.lblCoordsElev.Size = New System.Drawing.Size(39, 13)
        Me.lblCoordsElev.TabIndex = 4
        Me.lblCoordsElev.Text = "Label3"
        '
        'lblOvbComm
        '
        Me.lblOvbComm.AutoSize = True
        Me.lblOvbComm.Location = New System.Drawing.Point(130, 36)
        Me.lblOvbComm.Name = "lblOvbComm"
        Me.lblOvbComm.Size = New System.Drawing.Size(39, 13)
        Me.lblOvbComm.TabIndex = 3
        Me.lblOvbComm.Text = "Label2"
        '
        'ssDrillData
        '
        Me.ssDrillData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ssDrillData.Location = New System.Drawing.Point(108, 71)
        Me.ssDrillData.Name = "ssDrillData"
        Me.ssDrillData.OcxState = CType(resources.GetObject("ssDrillData.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssDrillData.Size = New System.Drawing.Size(1008, 150)
        Me.ssDrillData.TabIndex = 2
        '
        'tabDisp
        '
        Me.tabDisp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabDisp.Controls.Add(Me.TabPage1)
        Me.tabDisp.Controls.Add(Me.TabPage2)
        Me.tabDisp.Controls.Add(Me.TabPage3)
        Me.tabDisp.Controls.Add(Me.TabPage4)
        Me.tabDisp.Controls.Add(Me.TabPage5)
        Me.tabDisp.Controls.Add(Me.TabPage6)
        Me.tabDisp.Controls.Add(Me.TabPage7)
        Me.tabDisp.Controls.Add(Me.TabPage8)
        Me.tabDisp.Controls.Add(Me.TabPage9)
        Me.tabDisp.Controls.Add(Me.TabPage10)
        Me.tabDisp.Location = New System.Drawing.Point(16, 244)
        Me.tabDisp.Name = "tabDisp"
        Me.tabDisp.SelectedIndex = 0
        Me.tabDisp.Size = New System.Drawing.Size(1100, 211)
        Me.tabDisp.TabIndex = 1
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.lblHole)
        Me.TabPage1.Controls.Add(Me.lblUserMadeHoleUnmineable)
        Me.TabPage1.Controls.Add(Me.ssHoleData)
        Me.TabPage1.Controls.Add(Me.fraMiscStuff)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1092, 185)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Hole Composite"
        '
        'lblHole
        '
        Me.lblHole.AutoSize = True
        Me.lblHole.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHole.Location = New System.Drawing.Point(6, 12)
        Me.lblHole.Name = "lblHole"
        Me.lblHole.Size = New System.Drawing.Size(45, 13)
        Me.lblHole.TabIndex = 1
        Me.lblHole.Text = "Label1"
        '
        'lblUserMadeHoleUnmineable
        '
        Me.lblUserMadeHoleUnmineable.AutoSize = True
        Me.lblUserMadeHoleUnmineable.Location = New System.Drawing.Point(517, 14)
        Me.lblUserMadeHoleUnmineable.Name = "lblUserMadeHoleUnmineable"
        Me.lblUserMadeHoleUnmineable.Size = New System.Drawing.Size(39, 13)
        Me.lblUserMadeHoleUnmineable.TabIndex = 2
        Me.lblUserMadeHoleUnmineable.Text = "Label2"
        '
        'ssHoleData
        '
        Me.ssHoleData.Location = New System.Drawing.Point(7, 31)
        Me.ssHoleData.Name = "ssHoleData"
        Me.ssHoleData.OcxState = CType(resources.GetObject("ssHoleData.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssHoleData.Size = New System.Drawing.Size(952, 130)
        Me.ssHoleData.TabIndex = 0
        '
        'fraMiscStuff
        '
        Me.fraMiscStuff.Controls.Add(Me.ssRawProspMin)
        Me.fraMiscStuff.Location = New System.Drawing.Point(769, 14)
        Me.fraMiscStuff.Name = "fraMiscStuff"
        Me.fraMiscStuff.Size = New System.Drawing.Size(727, 157)
        Me.fraMiscStuff.TabIndex = 1
        Me.fraMiscStuff.TabStop = False
        Me.fraMiscStuff.Visible = False
        '
        'ssRawProspMin
        '
        Me.ssRawProspMin.Location = New System.Drawing.Point(473, 11)
        Me.ssRawProspMin.Name = "ssRawProspMin"
        Me.ssRawProspMin.OcxState = CType(resources.GetObject("ssRawProspMin.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssRawProspMin.Size = New System.Drawing.Size(237, 137)
        Me.ssRawProspMin.TabIndex = 4
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage2.Controls.Add(Me.lblSplit)
        Me.TabPage2.Controls.Add(Me.lblCurrSplit)
        Me.TabPage2.Controls.Add(Me.ssSplitData)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1092, 185)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Split Reduced"
        '
        'lblSplit
        '
        Me.lblSplit.AutoSize = True
        Me.lblSplit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSplit.Location = New System.Drawing.Point(6, 12)
        Me.lblSplit.Name = "lblSplit"
        Me.lblSplit.Size = New System.Drawing.Size(45, 13)
        Me.lblSplit.TabIndex = 1
        Me.lblSplit.Text = "Label1"
        '
        'lblCurrSplit
        '
        Me.lblCurrSplit.AutoSize = True
        Me.lblCurrSplit.Location = New System.Drawing.Point(518, 13)
        Me.lblCurrSplit.Name = "lblCurrSplit"
        Me.lblCurrSplit.Size = New System.Drawing.Size(39, 13)
        Me.lblCurrSplit.TabIndex = 2
        Me.lblCurrSplit.Text = "Label2"
        '
        'ssSplitData
        '
        Me.ssSplitData.Location = New System.Drawing.Point(7, 31)
        Me.ssSplitData.Name = "ssSplitData"
        Me.ssSplitData.OcxState = CType(resources.GetObject("ssSplitData.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssSplitData.Size = New System.Drawing.Size(952, 130)
        Me.ssSplitData.TabIndex = 0
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage3.Controls.Add(Me.tblReductionData)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(1092, 185)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Reduction Data"
        '
        'tblReductionData
        '
        Me.tblReductionData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tblReductionData.ColumnCount = 2
        Me.tblReductionData.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.tblReductionData.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.tblReductionData.Controls.Add(Me.lblRdctnHole, 1, 0)
        Me.tblReductionData.Controls.Add(Me.lblRdctnSplit, 0, 0)
        Me.tblReductionData.Controls.Add(Me.ssSplitReview, 0, 1)
        Me.tblReductionData.Controls.Add(Me.ssCompReview, 1, 1)
        Me.tblReductionData.Location = New System.Drawing.Point(6, 3)
        Me.tblReductionData.Name = "tblReductionData"
        Me.tblReductionData.RowCount = 2
        Me.tblReductionData.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 16.0!))
        Me.tblReductionData.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tblReductionData.Size = New System.Drawing.Size(1094, 182)
        Me.tblReductionData.TabIndex = 4
        '
        'lblRdctnHole
        '
        Me.lblRdctnHole.AutoSize = True
        Me.lblRdctnHole.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRdctnHole.Location = New System.Drawing.Point(550, 0)
        Me.lblRdctnHole.Name = "lblRdctnHole"
        Me.lblRdctnHole.Size = New System.Drawing.Size(33, 13)
        Me.lblRdctnHole.TabIndex = 3
        Me.lblRdctnHole.Text = "Hole"
        '
        'lblRdctnSplit
        '
        Me.lblRdctnSplit.AutoSize = True
        Me.lblRdctnSplit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRdctnSplit.Location = New System.Drawing.Point(3, 0)
        Me.lblRdctnSplit.Name = "lblRdctnSplit"
        Me.lblRdctnSplit.Size = New System.Drawing.Size(32, 13)
        Me.lblRdctnSplit.TabIndex = 2
        Me.lblRdctnSplit.Text = "Split"
        '
        'ssSplitReview
        '
        Me.ssSplitReview.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ssSplitReview.Location = New System.Drawing.Point(3, 19)
        Me.ssSplitReview.Name = "ssSplitReview"
        Me.ssSplitReview.OcxState = CType(resources.GetObject("ssSplitReview.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssSplitReview.Size = New System.Drawing.Size(541, 160)
        Me.ssSplitReview.TabIndex = 0
        '
        'ssCompReview
        '
        Me.ssCompReview.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ssCompReview.Location = New System.Drawing.Point(550, 19)
        Me.ssCompReview.Name = "ssCompReview"
        Me.ssCompReview.OcxState = CType(resources.GetObject("ssCompReview.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssCompReview.Size = New System.Drawing.Size(541, 160)
        Me.ssCompReview.TabIndex = 1
        '
        'TabPage4
        '
        Me.TabPage4.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage4.Controls.Add(Me.lblHoleInMOIS)
        Me.TabPage4.Controls.Add(Me.lblHoleExistStatus)
        Me.TabPage4.Controls.Add(Me.ssHoleExistStatus)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(1092, 185)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "MOIS"
        '
        'lblHoleInMOIS
        '
        Me.lblHoleInMOIS.AutoSize = True
        Me.lblHoleInMOIS.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHoleInMOIS.Location = New System.Drawing.Point(3, 17)
        Me.lblHoleInMOIS.Name = "lblHoleInMOIS"
        Me.lblHoleInMOIS.Size = New System.Drawing.Size(176, 13)
        Me.lblHoleInMOIS.TabIndex = 2
        Me.lblHoleInMOIS.Text = "Hole/Split Existence in  MOIS"
        '
        'lblHoleExistStatus
        '
        Me.lblHoleExistStatus.AutoSize = True
        Me.lblHoleExistStatus.Location = New System.Drawing.Point(3, 51)
        Me.lblHoleExistStatus.Name = "lblHoleExistStatus"
        Me.lblHoleExistStatus.Size = New System.Drawing.Size(39, 13)
        Me.lblHoleExistStatus.TabIndex = 1
        Me.lblHoleExistStatus.Text = "Label2"
        '
        'ssHoleExistStatus
        '
        Me.ssHoleExistStatus.Location = New System.Drawing.Point(6, 85)
        Me.ssHoleExistStatus.Name = "ssHoleExistStatus"
        Me.ssHoleExistStatus.OcxState = CType(resources.GetObject("ssHoleExistStatus.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssHoleExistStatus.Size = New System.Drawing.Size(752, 81)
        Me.ssHoleExistStatus.TabIndex = 0
        '
        'TabPage5
        '
        Me.TabPage5.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage5.Controls.Add(Me.Label1)
        Me.TabPage5.Controls.Add(Me.lblInfoComm)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Size = New System.Drawing.Size(1092, 185)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Info"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(22, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(212, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "When saving comp/splits to MOIS..."
        '
        'lblInfoComm
        '
        Me.lblInfoComm.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblInfoComm.ForeColor = System.Drawing.Color.Navy
        Me.lblInfoComm.Location = New System.Drawing.Point(22, 53)
        Me.lblInfoComm.Name = "lblInfoComm"
        Me.lblInfoComm.Size = New System.Drawing.Size(1061, 119)
        Me.lblInfoComm.TabIndex = 1
        Me.lblInfoComm.Text = "Label2"
        '
        'TabPage6
        '
        Me.TabPage6.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage6.Controls.Add(Me.cmdPrtGrd)
        Me.TabPage6.Controls.Add(Me.ssCompErrors)
        Me.TabPage6.Location = New System.Drawing.Point(4, 22)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Size = New System.Drawing.Size(1092, 185)
        Me.TabPage6.TabIndex = 5
        Me.TabPage6.Text = "Hole Issues"
        '
        'cmdPrtGrd
        '
        Me.cmdPrtGrd.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.cmdPrtGrd.Location = New System.Drawing.Point(508, 153)
        Me.cmdPrtGrd.Name = "cmdPrtGrd"
        Me.cmdPrtGrd.Size = New System.Drawing.Size(75, 23)
        Me.cmdPrtGrd.TabIndex = 0
        Me.cmdPrtGrd.Text = "PrtGrd"
        Me.cmdPrtGrd.UseVisualStyleBackColor = True
        '
        'ssCompErrors
        '
        Me.ssCompErrors.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ssCompErrors.Location = New System.Drawing.Point(6, 19)
        Me.ssCompErrors.Name = "ssCompErrors"
        Me.ssCompErrors.OcxState = CType(resources.GetObject("ssCompErrors.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssCompErrors.Size = New System.Drawing.Size(1074, 125)
        Me.ssCompErrors.TabIndex = 1
        '
        'TabPage7
        '
        Me.TabPage7.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage7.Controls.Add(Me.lblAreaName)
        Me.TabPage7.Controls.Add(Me.txtAreaName)
        Me.TabPage7.Controls.Add(Me.cmdSaveAreaName)
        Me.TabPage7.Location = New System.Drawing.Point(4, 22)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Size = New System.Drawing.Size(1092, 185)
        Me.TabPage7.TabIndex = 6
        Me.TabPage7.Text = "Add Area"
        '
        'lblAreaName
        '
        Me.lblAreaName.AutoSize = True
        Me.lblAreaName.Location = New System.Drawing.Point(21, 21)
        Me.lblAreaName.Name = "lblAreaName"
        Me.lblAreaName.Size = New System.Drawing.Size(58, 13)
        Me.lblAreaName.TabIndex = 1
        Me.lblAreaName.Text = "Area name"
        Me.lblAreaName.Visible = False
        '
        'txtAreaName
        '
        Me.txtAreaName.Location = New System.Drawing.Point(81, 17)
        Me.txtAreaName.MaxLength = 30
        Me.txtAreaName.Name = "txtAreaName"
        Me.txtAreaName.Size = New System.Drawing.Size(175, 20)
        Me.txtAreaName.TabIndex = 2
        Me.txtAreaName.Visible = False
        '
        'cmdSaveAreaName
        '
        Me.cmdSaveAreaName.Location = New System.Drawing.Point(93, 71)
        Me.cmdSaveAreaName.Name = "cmdSaveAreaName"
        Me.cmdSaveAreaName.Size = New System.Drawing.Size(142, 23)
        Me.cmdSaveAreaName.TabIndex = 0
        Me.cmdSaveAreaName.Text = "Save Area Name"
        Me.cmdSaveAreaName.UseVisualStyleBackColor = True
        Me.cmdSaveAreaName.Visible = False
        '
        'TabPage8
        '
        Me.TabPage8.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage8.Controls.Add(Me.lblCurrMinabilityComm)
        Me.TabPage8.Controls.Add(Me.lblSplitMinabilities)
        Me.TabPage8.Controls.Add(Me.lblHoleMinability)
        Me.TabPage8.Controls.Add(Me.ssSplitMinabilities)
        Me.TabPage8.Controls.Add(Me.ssHoleMinabilities)
        Me.TabPage8.Location = New System.Drawing.Point(4, 22)
        Me.TabPage8.Name = "TabPage8"
        Me.TabPage8.Size = New System.Drawing.Size(1092, 185)
        Me.TabPage8.TabIndex = 7
        Me.TabPage8.Text = "Current Minabilities"
        '
        'lblCurrMinabilityComm
        '
        Me.lblCurrMinabilityComm.ForeColor = System.Drawing.Color.Navy
        Me.lblCurrMinabilityComm.Location = New System.Drawing.Point(31, 119)
        Me.lblCurrMinabilityComm.Name = "lblCurrMinabilityComm"
        Me.lblCurrMinabilityComm.Size = New System.Drawing.Size(239, 55)
        Me.lblCurrMinabilityComm.TabIndex = 4
        Me.lblCurrMinabilityComm.Text = "Label1"
        '
        'lblSplitMinabilities
        '
        Me.lblSplitMinabilities.AutoSize = True
        Me.lblSplitMinabilities.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSplitMinabilities.Location = New System.Drawing.Point(349, 36)
        Me.lblSplitMinabilities.Name = "lblSplitMinabilities"
        Me.lblSplitMinabilities.Size = New System.Drawing.Size(249, 13)
        Me.lblSplitMinabilities.TabIndex = 2
        Me.lblSplitMinabilities.Text = "Split Minabilities in Raw Prospect (Current)"
        '
        'lblHoleMinability
        '
        Me.lblHoleMinability.AutoSize = True
        Me.lblHoleMinability.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHoleMinability.Location = New System.Drawing.Point(31, 36)
        Me.lblHoleMinability.Name = "lblHoleMinability"
        Me.lblHoleMinability.Size = New System.Drawing.Size(250, 13)
        Me.lblHoleMinability.TabIndex = 1
        Me.lblHoleMinability.Text = "Hole Minabilities in Raw Prospect (Current)"
        '
        'ssSplitMinabilities
        '
        Me.ssSplitMinabilities.Location = New System.Drawing.Point(352, 55)
        Me.ssSplitMinabilities.Name = "ssSplitMinabilities"
        Me.ssSplitMinabilities.OcxState = CType(resources.GetObject("ssSplitMinabilities.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssSplitMinabilities.Size = New System.Drawing.Size(463, 112)
        Me.ssSplitMinabilities.TabIndex = 3
        '
        'ssHoleMinabilities
        '
        Me.ssHoleMinabilities.Location = New System.Drawing.Point(31, 58)
        Me.ssHoleMinabilities.Name = "ssHoleMinabilities"
        Me.ssHoleMinabilities.OcxState = CType(resources.GetObject("ssHoleMinabilities.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssHoleMinabilities.Size = New System.Drawing.Size(228, 49)
        Me.ssHoleMinabilities.TabIndex = 0
        '
        'TabPage9
        '
        Me.TabPage9.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage9.Controls.Add(Me.lblFeAdjComm)
        Me.TabPage9.Controls.Add(Me.chkUseFeAdjust)
        Me.TabPage9.Controls.Add(Me.ssFeAdjustment)
        Me.TabPage9.Location = New System.Drawing.Point(4, 22)
        Me.TabPage9.Name = "TabPage9"
        Me.TabPage9.Size = New System.Drawing.Size(1092, 185)
        Me.TabPage9.TabIndex = 8
        Me.TabPage9.Text = "Miscellaneous"
        '
        'lblFeAdjComm
        '
        Me.lblFeAdjComm.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFeAdjComm.ForeColor = System.Drawing.Color.Navy
        Me.lblFeAdjComm.Location = New System.Drawing.Point(397, 29)
        Me.lblFeAdjComm.Name = "lblFeAdjComm"
        Me.lblFeAdjComm.Size = New System.Drawing.Size(691, 92)
        Me.lblFeAdjComm.TabIndex = 2
        Me.lblFeAdjComm.Text = "Label1"
        '
        'chkUseFeAdjust
        '
        Me.chkUseFeAdjust.AutoSize = True
        Me.chkUseFeAdjust.Location = New System.Drawing.Point(20, 37)
        Me.chkUseFeAdjust.Name = "chkUseFeAdjust"
        Me.chkUseFeAdjust.Size = New System.Drawing.Size(269, 17)
        Me.chkUseFeAdjust.TabIndex = 0
        Me.chkUseFeAdjust.Text = "Use """"Adjusted"""" Fe to determine product minability"
        Me.chkUseFeAdjust.UseVisualStyleBackColor = True
        '
        'ssFeAdjustment
        '
        Me.ssFeAdjustment.Location = New System.Drawing.Point(43, 74)
        Me.ssFeAdjustment.Name = "ssFeAdjustment"
        Me.ssFeAdjustment.OcxState = CType(resources.GetObject("ssFeAdjustment.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssFeAdjustment.Size = New System.Drawing.Size(178, 32)
        Me.ssFeAdjustment.TabIndex = 1
        '
        'TabPage10
        '
        Me.TabPage10.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage10.Controls.Add(Me.lblSplitOverrideSet)
        Me.TabPage10.Controls.Add(Me.Label4)
        Me.TabPage10.Controls.Add(Me.txtSplitOverrideName)
        Me.TabPage10.Controls.Add(Me.cmdAddToOverrideSet)
        Me.TabPage10.Controls.Add(Me.ssSplitOverride)
        Me.TabPage10.Controls.Add(Me.cmdRefresh)
        Me.TabPage10.Controls.Add(Me.cboSplitOverrideMineName)
        Me.TabPage10.Controls.Add(Me.Frame2)
        Me.TabPage10.Location = New System.Drawing.Point(4, 22)
        Me.TabPage10.Name = "TabPage10"
        Me.TabPage10.Size = New System.Drawing.Size(1092, 185)
        Me.TabPage10.TabIndex = 9
        Me.TabPage10.Text = "Add to Override"
        '
        'lblSplitOverrideSet
        '
        Me.lblSplitOverrideSet.Location = New System.Drawing.Point(19, 37)
        Me.lblSplitOverrideSet.Name = "lblSplitOverrideSet"
        Me.lblSplitOverrideSet.Size = New System.Drawing.Size(76, 34)
        Me.lblSplitOverrideSet.TabIndex = 7
        Me.lblSplitOverrideSet.Text = "Split Override Set"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(36, 79)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Mine name"
        '
        'txtSplitOverrideName
        '
        Me.txtSplitOverrideName.Location = New System.Drawing.Point(98, 37)
        Me.txtSplitOverrideName.MaxLength = 30
        Me.txtSplitOverrideName.Name = "txtSplitOverrideName"
        Me.txtSplitOverrideName.Size = New System.Drawing.Size(209, 20)
        Me.txtSplitOverrideName.TabIndex = 3
        '
        'cmdAddToOverrideSet
        '
        Me.cmdAddToOverrideSet.Location = New System.Drawing.Point(39, 133)
        Me.cmdAddToOverrideSet.Name = "cmdAddToOverrideSet"
        Me.cmdAddToOverrideSet.Size = New System.Drawing.Size(179, 23)
        Me.cmdAddToOverrideSet.TabIndex = 4
        Me.cmdAddToOverrideSet.Text = "Add Splits to Override Set"
        Me.cmdAddToOverrideSet.UseVisualStyleBackColor = True
        '
        'ssSplitOverride
        '
        Me.ssSplitOverride.Location = New System.Drawing.Point(860, 28)
        Me.ssSplitOverride.Name = "ssSplitOverride"
        Me.ssSplitOverride.OcxState = CType(resources.GetObject("ssSplitOverride.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssSplitOverride.Size = New System.Drawing.Size(209, 115)
        Me.ssSplitOverride.TabIndex = 5
        '
        'cmdRefresh
        '
        Me.cmdRefresh.Location = New System.Drawing.Point(862, 153)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(58, 23)
        Me.cmdRefresh.TabIndex = 0
        Me.cmdRefresh.Text = "Refresh"
        Me.cmdRefresh.UseVisualStyleBackColor = True
        '
        'cboSplitOverrideMineName
        '
        Me.cboSplitOverrideMineName.FormattingEnabled = True
        Me.cboSplitOverrideMineName.Location = New System.Drawing.Point(98, 75)
        Me.cboSplitOverrideMineName.Name = "cboSplitOverrideMineName"
        Me.cboSplitOverrideMineName.Size = New System.Drawing.Size(133, 21)
        Me.cboSplitOverrideMineName.TabIndex = 1
        '
        'Frame2
        '
        Me.Frame2.Controls.Add(Me.ssSplitOverrides)
        Me.Frame2.Controls.Add(Me.chkOnlyMySplitOverride)
        Me.Frame2.Controls.Add(Me.cmdGetSplitOverrides)
        Me.Frame2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame2.Location = New System.Drawing.Point(317, 13)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Size = New System.Drawing.Size(510, 154)
        Me.Frame2.TabIndex = 2
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Select an existing set"
        '
        'ssSplitOverrides
        '
        Me.ssSplitOverrides.Location = New System.Drawing.Point(6, 48)
        Me.ssSplitOverrides.Name = "ssSplitOverrides"
        Me.ssSplitOverrides.OcxState = CType(resources.GetObject("ssSplitOverrides.OcxState"), System.Windows.Forms.AxHost.State)
        Me.ssSplitOverrides.Size = New System.Drawing.Size(498, 100)
        Me.ssSplitOverrides.TabIndex = 2
        '
        'chkOnlyMySplitOverride
        '
        Me.chkOnlyMySplitOverride.AutoSize = True
        Me.chkOnlyMySplitOverride.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOnlyMySplitOverride.Location = New System.Drawing.Point(6, 21)
        Me.chkOnlyMySplitOverride.Name = "chkOnlyMySplitOverride"
        Me.chkOnlyMySplitOverride.Size = New System.Drawing.Size(161, 17)
        Me.chkOnlyMySplitOverride.TabIndex = 1
        Me.chkOnlyMySplitOverride.Text = "Select only my split overrides"
        Me.chkOnlyMySplitOverride.UseVisualStyleBackColor = True
        '
        'cmdGetSplitOverrides
        '
        Me.cmdGetSplitOverrides.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGetSplitOverrides.Location = New System.Drawing.Point(230, 19)
        Me.cmdGetSplitOverrides.Name = "cmdGetSplitOverrides"
        Me.cmdGetSplitOverrides.Size = New System.Drawing.Size(132, 23)
        Me.cmdGetSplitOverrides.TabIndex = 0
        Me.cmdGetSplitOverrides.Text = "Get Overrides"
        Me.cmdGetSplitOverrides.UseVisualStyleBackColor = True
        '
        'fraSelect
        '
        Me.fraSelect.Controls.Add(Me.lblAlphaNumeric)
        Me.fraSelect.Controls.Add(Me.lblOffSpecPbMgPlt)
        Me.fraSelect.Controls.Add(Me.lblHoleInMoisComm)
        Me.fraSelect.Controls.Add(Me.lblScenComm)
        Me.fraSelect.Controls.Add(Me.lblHoleLocation)
        Me.fraSelect.Controls.Add(Me.lblRange)
        Me.fraSelect.Controls.Add(Me.lblSection)
        Me.fraSelect.Controls.Add(Me.lblTownship)
        Me.fraSelect.Controls.Add(Me.cboHole)
        Me.fraSelect.Controls.Add(Me.cboRge)
        Me.fraSelect.Controls.Add(Me.cboSec)
        Me.fraSelect.Controls.Add(Me.cboTwp)
        Me.fraSelect.Controls.Add(Me.lblOtherDefn)
        Me.fraSelect.Controls.Add(Me.lblIpComm)
        Me.fraSelect.Controls.Add(Me.lblUseOrigHoleComm)
        Me.fraSelect.Controls.Add(Me.cmdReduceHole)
        Me.fraSelect.Controls.Add(Me.chkMyParams)
        Me.fraSelect.Controls.Add(Me.cmdRefreshParams)
        Me.fraSelect.Controls.Add(Me.cboOtherDefn)
        Me.fraSelect.Controls.Add(Me.cboProdSizeDefn)
        Me.fraSelect.Controls.Add(Me.lblProdSizeDefn)
        Me.fraSelect.Controls.Add(Me.cmdTest3)
        Me.fraSelect.Controls.Add(Me.cmdTest2)
        Me.fraSelect.Controls.Add(Me.cmdTest)
        Me.fraSelect.Controls.Add(Me.chkOverrideMaxDepth)
        Me.fraSelect.Controls.Add(Me.chkUseOrigHole)
        Me.fraSelect.Location = New System.Drawing.Point(40, 19)
        Me.fraSelect.Name = "fraSelect"
        Me.fraSelect.Size = New System.Drawing.Size(1132, 107)
        Me.fraSelect.TabIndex = 3
        Me.fraSelect.TabStop = False
        Me.fraSelect.Text = "Select hole to reduce"
        '
        'lblAlphaNumeric
        '
        Me.lblAlphaNumeric.AutoSize = True
        Me.lblAlphaNumeric.Location = New System.Drawing.Point(284, 64)
        Me.lblAlphaNumeric.Name = "lblAlphaNumeric"
        Me.lblAlphaNumeric.Size = New System.Drawing.Size(83, 13)
        Me.lblAlphaNumeric.TabIndex = 25
        Me.lblAlphaNumeric.Text = "lblAlphaNumeric"
        Me.lblAlphaNumeric.Visible = False
        '
        'lblOffSpecPbMgPlt
        '
        Me.lblOffSpecPbMgPlt.AutoSize = True
        Me.lblOffSpecPbMgPlt.Location = New System.Drawing.Point(715, 89)
        Me.lblOffSpecPbMgPlt.Name = "lblOffSpecPbMgPlt"
        Me.lblOffSpecPbMgPlt.Size = New System.Drawing.Size(39, 13)
        Me.lblOffSpecPbMgPlt.TabIndex = 24
        Me.lblOffSpecPbMgPlt.Text = "Label5"
        Me.lblOffSpecPbMgPlt.Visible = False
        '
        'lblHoleInMoisComm
        '
        Me.lblHoleInMoisComm.AutoSize = True
        Me.lblHoleInMoisComm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHoleInMoisComm.ForeColor = System.Drawing.Color.Red
        Me.lblHoleInMoisComm.Location = New System.Drawing.Point(23, 82)
        Me.lblHoleInMoisComm.Name = "lblHoleInMoisComm"
        Me.lblHoleInMoisComm.Size = New System.Drawing.Size(20, 13)
        Me.lblHoleInMoisComm.TabIndex = 23
        Me.lblHoleInMoisComm.Text = "lbl"
        '
        'lblScenComm
        '
        Me.lblScenComm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblScenComm.ForeColor = System.Drawing.Color.Red
        Me.lblScenComm.Location = New System.Drawing.Point(914, 13)
        Me.lblScenComm.Name = "lblScenComm"
        Me.lblScenComm.Size = New System.Drawing.Size(200, 30)
        Me.lblScenComm.TabIndex = 22
        Me.lblScenComm.Text = "lbl"
        Me.lblScenComm.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblScenComm.Visible = False
        '
        'lblHoleLocation
        '
        Me.lblHoleLocation.AutoSize = True
        Me.lblHoleLocation.Location = New System.Drawing.Point(150, 55)
        Me.lblHoleLocation.Name = "lblHoleLocation"
        Me.lblHoleLocation.Size = New System.Drawing.Size(29, 13)
        Me.lblHoleLocation.TabIndex = 21
        Me.lblHoleLocation.Text = "Hole"
        '
        'lblRange
        '
        Me.lblRange.AutoSize = True
        Me.lblRange.Location = New System.Drawing.Point(20, 53)
        Me.lblRange.Name = "lblRange"
        Me.lblRange.Size = New System.Drawing.Size(27, 13)
        Me.lblRange.TabIndex = 20
        Me.lblRange.Text = "Rge"
        '
        'lblSection
        '
        Me.lblSection.AutoSize = True
        Me.lblSection.Location = New System.Drawing.Point(153, 20)
        Me.lblSection.Name = "lblSection"
        Me.lblSection.Size = New System.Drawing.Size(26, 13)
        Me.lblSection.TabIndex = 19
        Me.lblSection.Text = "Sec"
        '
        'lblTownship
        '
        Me.lblTownship.AutoSize = True
        Me.lblTownship.Location = New System.Drawing.Point(19, 20)
        Me.lblTownship.Name = "lblTownship"
        Me.lblTownship.Size = New System.Drawing.Size(28, 13)
        Me.lblTownship.TabIndex = 18
        Me.lblTownship.Text = "Twp"
        '
        'cboHole
        '
        Me.cboHole.FormattingEnabled = True
        Me.cboHole.Location = New System.Drawing.Point(182, 50)
        Me.cboHole.Name = "cboHole"
        Me.cboHole.Size = New System.Drawing.Size(74, 21)
        Me.cboHole.TabIndex = 17
        '
        'cboRge
        '
        Me.cboRge.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRge.FormattingEnabled = True
        Me.cboRge.Location = New System.Drawing.Point(50, 50)
        Me.cboRge.Name = "cboRge"
        Me.cboRge.Size = New System.Drawing.Size(74, 21)
        Me.cboRge.TabIndex = 16
        '
        'cboSec
        '
        Me.cboSec.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSec.FormattingEnabled = True
        Me.cboSec.Location = New System.Drawing.Point(182, 16)
        Me.cboSec.Name = "cboSec"
        Me.cboSec.Size = New System.Drawing.Size(74, 21)
        Me.cboSec.TabIndex = 15
        '
        'cboTwp
        '
        Me.cboTwp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTwp.FormattingEnabled = True
        Me.cboTwp.Location = New System.Drawing.Point(50, 16)
        Me.cboTwp.Name = "cboTwp"
        Me.cboTwp.Size = New System.Drawing.Size(74, 21)
        Me.cboTwp.TabIndex = 14
        '
        'lblOtherDefn
        '
        Me.lblOtherDefn.Location = New System.Drawing.Point(377, 59)
        Me.lblOtherDefn.Name = "lblOtherDefn"
        Me.lblOtherDefn.Size = New System.Drawing.Size(123, 30)
        Me.lblOtherDefn.TabIndex = 13
        Me.lblOtherDefn.Text = "Rcvry, Prod Adj, Prod Qual, Minability Scenario"
        '
        'lblIpComm
        '
        Me.lblIpComm.AutoSize = True
        Me.lblIpComm.ForeColor = System.Drawing.Color.Navy
        Me.lblIpComm.Location = New System.Drawing.Point(508, 44)
        Me.lblIpComm.Name = "lblIpComm"
        Me.lblIpComm.Size = New System.Drawing.Size(155, 13)
        Me.lblIpComm.TabIndex = 12
        Me.lblIpComm.Text = "IP && OS won't transfer to MOIS!"
        '
        'lblUseOrigHoleComm
        '
        Me.lblUseOrigHoleComm.AutoSize = True
        Me.lblUseOrigHoleComm.ForeColor = System.Drawing.Color.Navy
        Me.lblUseOrigHoleComm.Location = New System.Drawing.Point(786, 88)
        Me.lblUseOrigHoleComm.Name = "lblUseOrigHoleComm"
        Me.lblUseOrigHoleComm.Size = New System.Drawing.Size(39, 13)
        Me.lblUseOrigHoleComm.TabIndex = 11
        Me.lblUseOrigHoleComm.Text = "Label3"
        '
        'cmdReduceHole
        '
        Me.cmdReduceHole.Location = New System.Drawing.Point(287, 17)
        Me.cmdReduceHole.Name = "cmdReduceHole"
        Me.cmdReduceHole.Size = New System.Drawing.Size(75, 43)
        Me.cmdReduceHole.TabIndex = 10
        Me.cmdReduceHole.Text = "Reduce Hole"
        Me.cmdReduceHole.UseVisualStyleBackColor = True
        '
        'chkMyParams
        '
        Me.chkMyParams.Checked = True
        Me.chkMyParams.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMyParams.Location = New System.Drawing.Point(735, 14)
        Me.chkMyParams.Name = "chkMyParams"
        Me.chkMyParams.Size = New System.Drawing.Size(129, 35)
        Me.chkMyParams.TabIndex = 9
        Me.chkMyParams.Text = "Select only my parameter definitions"
        Me.chkMyParams.UseVisualStyleBackColor = True
        '
        'cmdRefreshParams
        '
        Me.cmdRefreshParams.Location = New System.Drawing.Point(735, 55)
        Me.cmdRefreshParams.Name = "cmdRefreshParams"
        Me.cmdRefreshParams.Size = New System.Drawing.Size(148, 23)
        Me.cmdRefreshParams.TabIndex = 8
        Me.cmdRefreshParams.Text = "Refresh Parameters"
        Me.cmdRefreshParams.UseVisualStyleBackColor = True
        '
        'cboOtherDefn
        '
        Me.cboOtherDefn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboOtherDefn.FormattingEnabled = True
        Me.cboOtherDefn.Location = New System.Drawing.Point(506, 61)
        Me.cboOtherDefn.Name = "cboOtherDefn"
        Me.cboOtherDefn.Size = New System.Drawing.Size(209, 21)
        Me.cboOtherDefn.TabIndex = 7
        '
        'cboProdSizeDefn
        '
        Me.cboProdSizeDefn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProdSizeDefn.FormattingEnabled = True
        Me.cboProdSizeDefn.Location = New System.Drawing.Point(506, 13)
        Me.cboProdSizeDefn.Name = "cboProdSizeDefn"
        Me.cboProdSizeDefn.Size = New System.Drawing.Size(209, 21)
        Me.cboProdSizeDefn.TabIndex = 6
        '
        'lblProdSizeDefn
        '
        Me.lblProdSizeDefn.AutoSize = True
        Me.lblProdSizeDefn.Location = New System.Drawing.Point(378, 17)
        Me.lblProdSizeDefn.Name = "lblProdSizeDefn"
        Me.lblProdSizeDefn.Size = New System.Drawing.Size(122, 13)
        Me.lblProdSizeDefn.TabIndex = 5
        Me.lblProdSizeDefn.Text = "Product size designation"
        '
        'cmdTest3
        '
        Me.cmdTest3.Location = New System.Drawing.Point(422, 32)
        Me.cmdTest3.Name = "cmdTest3"
        Me.cmdTest3.Size = New System.Drawing.Size(15, 23)
        Me.cmdTest3.TabIndex = 4
        Me.cmdTest3.Text = "T"
        Me.cmdTest3.UseVisualStyleBackColor = True
        '
        'cmdTest2
        '
        Me.cmdTest2.Location = New System.Drawing.Point(401, 32)
        Me.cmdTest2.Name = "cmdTest2"
        Me.cmdTest2.Size = New System.Drawing.Size(15, 23)
        Me.cmdTest2.TabIndex = 3
        Me.cmdTest2.Text = "T"
        Me.cmdTest2.UseVisualStyleBackColor = True
        '
        'cmdTest
        '
        Me.cmdTest.Location = New System.Drawing.Point(380, 32)
        Me.cmdTest.Name = "cmdTest"
        Me.cmdTest.Size = New System.Drawing.Size(15, 23)
        Me.cmdTest.TabIndex = 2
        Me.cmdTest.Text = "T"
        Me.cmdTest.UseVisualStyleBackColor = True
        '
        'chkOverrideMaxDepth
        '
        Me.chkOverrideMaxDepth.AutoSize = True
        Me.chkOverrideMaxDepth.Location = New System.Drawing.Point(917, 46)
        Me.chkOverrideMaxDepth.Name = "chkOverrideMaxDepth"
        Me.chkOverrideMaxDepth.Size = New System.Drawing.Size(118, 17)
        Me.chkOverrideMaxDepth.TabIndex = 1
        Me.chkOverrideMaxDepth.Text = "Override max depth"
        Me.chkOverrideMaxDepth.UseVisualStyleBackColor = True
        '
        'chkUseOrigHole
        '
        Me.chkUseOrigHole.AutoSize = True
        Me.chkUseOrigHole.Location = New System.Drawing.Point(917, 69)
        Me.chkUseOrigHole.Name = "chkUseOrigHole"
        Me.chkUseOrigHole.Size = New System.Drawing.Size(190, 17)
        Me.chkUseOrigHole.TabIndex = 0
        Me.chkUseOrigHole.Text = "Use original hole that was redrilled*"
        Me.chkUseOrigHole.UseVisualStyleBackColor = True
        '
        'cmdPrtScr
        '
        Me.cmdPrtScr.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPrtScr.Location = New System.Drawing.Point(1037, 710)
        Me.cmdPrtScr.Name = "cmdPrtScr"
        Me.cmdPrtScr.Size = New System.Drawing.Size(66, 23)
        Me.cmdPrtScr.TabIndex = 4
        Me.cmdPrtScr.Text = "PrtScr"
        Me.cmdPrtScr.UseVisualStyleBackColor = True
        '
        'sbrMain
        '
        Me.sbrMain.Location = New System.Drawing.Point(0, 710)
        Me.sbrMain.Name = "sbrMain"
        Me.sbrMain.Size = New System.Drawing.Size(1202, 22)
        Me.sbrMain.TabIndex = 5
        Me.sbrMain.Text = "StatusStrip1"
        '
        'fraSaveToMois
        '
        Me.fraSaveToMois.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.fraSaveToMois.Controls.Add(Me.cmdSaveMinabilities)
        Me.fraSaveToMois.Controls.Add(Me.fraRdctnType)
        Me.fraSaveToMois.Location = New System.Drawing.Point(40, 602)
        Me.fraSaveToMois.Name = "fraSaveToMois"
        Me.fraSaveToMois.Size = New System.Drawing.Size(612, 100)
        Me.fraSaveToMois.TabIndex = 6
        Me.fraSaveToMois.TabStop = False
        Me.fraSaveToMois.Text = "Save to MOIS"
        '
        'cmdSaveMinabilities
        '
        Me.cmdSaveMinabilities.Location = New System.Drawing.Point(7, 19)
        Me.cmdSaveMinabilities.Name = "cmdSaveMinabilities"
        Me.cmdSaveMinabilities.Size = New System.Drawing.Size(152, 75)
        Me.cmdSaveMinabilities.TabIndex = 1
        Me.cmdSaveMinabilities.Text = "Save Minabilities Only (to Raw Prospect)"
        Me.cmdSaveMinabilities.UseVisualStyleBackColor = True
        '
        'fraRdctnType
        '
        Me.fraRdctnType.Controls.Add(Me.lblSaveToMoisComm)
        Me.fraRdctnType.Controls.Add(Me.cmdSaveCompAndSplits)
        Me.fraRdctnType.Controls.Add(Me.optBothRdctn)
        Me.fraRdctnType.Controls.Add(Me.optCatalogRdctn)
        Me.fraRdctnType.Controls.Add(Me.opt100PctRdctn)
        Me.fraRdctnType.Controls.Add(Me.chkSaveRawProspectMinabilities)
        Me.fraRdctnType.Controls.Add(Me.cboMineName)
        Me.fraRdctnType.Location = New System.Drawing.Point(165, 0)
        Me.fraRdctnType.Name = "fraRdctnType"
        Me.fraRdctnType.Size = New System.Drawing.Size(441, 100)
        Me.fraRdctnType.TabIndex = 0
        Me.fraRdctnType.TabStop = False
        '
        'lblSaveToMoisComm
        '
        Me.lblSaveToMoisComm.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSaveToMoisComm.ForeColor = System.Drawing.Color.Navy
        Me.lblSaveToMoisComm.Location = New System.Drawing.Point(278, 50)
        Me.lblSaveToMoisComm.Name = "lblSaveToMoisComm"
        Me.lblSaveToMoisComm.Size = New System.Drawing.Size(141, 35)
        Me.lblSaveToMoisComm.TabIndex = 6
        Me.lblSaveToMoisComm.Text = "(100% Prospect && Catalog)"
        '
        'cmdSaveCompAndSplits
        '
        Me.cmdSaveCompAndSplits.Location = New System.Drawing.Point(6, 17)
        Me.cmdSaveCompAndSplits.Name = "cmdSaveCompAndSplits"
        Me.cmdSaveCompAndSplits.Size = New System.Drawing.Size(75, 77)
        Me.cmdSaveCompAndSplits.TabIndex = 5
        Me.cmdSaveCompAndSplits.Text = "Save Reduced Composites && Splits"
        Me.cmdSaveCompAndSplits.UseVisualStyleBackColor = True
        '
        'optBothRdctn
        '
        Me.optBothRdctn.AutoSize = True
        Me.optBothRdctn.Location = New System.Drawing.Point(97, 68)
        Me.optBothRdctn.Name = "optBothRdctn"
        Me.optBothRdctn.Size = New System.Drawing.Size(47, 17)
        Me.optBothRdctn.TabIndex = 4
        Me.optBothRdctn.TabStop = True
        Me.optBothRdctn.Text = "Both"
        Me.optBothRdctn.UseVisualStyleBackColor = True
        '
        'optCatalogRdctn
        '
        Me.optCatalogRdctn.AutoSize = True
        Me.optCatalogRdctn.Location = New System.Drawing.Point(97, 44)
        Me.optCatalogRdctn.Name = "optCatalogRdctn"
        Me.optCatalogRdctn.Size = New System.Drawing.Size(61, 17)
        Me.optCatalogRdctn.TabIndex = 3
        Me.optCatalogRdctn.TabStop = True
        Me.optCatalogRdctn.Text = "Catalog"
        Me.optCatalogRdctn.UseVisualStyleBackColor = True
        '
        'opt100PctRdctn
        '
        Me.opt100PctRdctn.AutoSize = True
        Me.opt100PctRdctn.Location = New System.Drawing.Point(97, 19)
        Me.opt100PctRdctn.Name = "opt100PctRdctn"
        Me.opt100PctRdctn.Size = New System.Drawing.Size(96, 17)
        Me.opt100PctRdctn.TabIndex = 2
        Me.opt100PctRdctn.TabStop = True
        Me.opt100PctRdctn.Text = "100% Prospect"
        Me.opt100PctRdctn.UseVisualStyleBackColor = True
        '
        'chkSaveRawProspectMinabilities
        '
        Me.chkSaveRawProspectMinabilities.Location = New System.Drawing.Point(193, 22)
        Me.chkSaveRawProspectMinabilities.Name = "chkSaveRawProspectMinabilities"
        Me.chkSaveRawProspectMinabilities.Size = New System.Drawing.Size(79, 62)
        Me.chkSaveRawProspectMinabilities.TabIndex = 1
        Me.chkSaveRawProspectMinabilities.Text = "Save raw prospect minabilities"
        Me.chkSaveRawProspectMinabilities.UseVisualStyleBackColor = True
        '
        'cboMineName
        '
        Me.cboMineName.FormattingEnabled = True
        Me.cboMineName.Location = New System.Drawing.Point(278, 19)
        Me.cboMineName.Name = "cboMineName"
        Me.cboMineName.Size = New System.Drawing.Size(141, 21)
        Me.cboMineName.TabIndex = 0
        '
        'fraSurvCadd
        '
        Me.fraSurvCadd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.fraSurvCadd.Controls.Add(Me.cmdCreateSurvCadd)
        Me.fraSurvCadd.Controls.Add(Me.lblSurvCaddComm)
        Me.fraSurvCadd.Controls.Add(Me.chkPbAnalysisFillInSpecial)
        Me.fraSurvCadd.Controls.Add(Me.optInclBoth)
        Me.fraSurvCadd.Controls.Add(Me.optInclSplits)
        Me.fraSurvCadd.Controls.Add(Me.optInclComposites)
        Me.fraSurvCadd.Controls.Add(Me.txtSurvCaddTextfile)
        Me.fraSurvCadd.Controls.Add(Me.lblSurvCaddTxtFile)
        Me.fraSurvCadd.Location = New System.Drawing.Point(658, 602)
        Me.fraSurvCadd.Name = "fraSurvCadd"
        Me.fraSurvCadd.Size = New System.Drawing.Size(514, 100)
        Me.fraSurvCadd.TabIndex = 7
        Me.fraSurvCadd.TabStop = False
        Me.fraSurvCadd.Text = "SurvCADD"
        '
        'cmdCreateSurvCadd
        '
        Me.cmdCreateSurvCadd.Location = New System.Drawing.Point(342, 61)
        Me.cmdCreateSurvCadd.Name = "cmdCreateSurvCadd"
        Me.cmdCreateSurvCadd.Size = New System.Drawing.Size(154, 23)
        Me.cmdCreateSurvCadd.TabIndex = 7
        Me.cmdCreateSurvCadd.Text = "Create Textfile"
        Me.cmdCreateSurvCadd.UseVisualStyleBackColor = True
        '
        'lblSurvCaddComm
        '
        Me.lblSurvCaddComm.AutoSize = True
        Me.lblSurvCaddComm.ForeColor = System.Drawing.Color.Navy
        Me.lblSurvCaddComm.Location = New System.Drawing.Point(252, 66)
        Me.lblSurvCaddComm.Name = "lblSurvCaddComm"
        Me.lblSurvCaddComm.Size = New System.Drawing.Size(84, 13)
        Me.lblSurvCaddComm.TabIndex = 6
        Me.lblSurvCaddComm.Text = "(100% Prospect)"
        '
        'chkPbAnalysisFillInSpecial
        '
        Me.chkPbAnalysisFillInSpecial.AutoSize = True
        Me.chkPbAnalysisFillInSpecial.Location = New System.Drawing.Point(20, 71)
        Me.chkPbAnalysisFillInSpecial.Name = "chkPbAnalysisFillInSpecial"
        Me.chkPbAnalysisFillInSpecial.Size = New System.Drawing.Size(205, 17)
        Me.chkPbAnalysisFillInSpecial.TabIndex = 5
        Me.chkPbAnalysisFillInSpecial.Text = "Set Pebble analysis to MgPlt if missing"
        Me.chkPbAnalysisFillInSpecial.UseVisualStyleBackColor = True
        '
        'optInclBoth
        '
        Me.optInclBoth.AutoSize = True
        Me.optInclBoth.Location = New System.Drawing.Point(145, 48)
        Me.optInclBoth.Name = "optInclBoth"
        Me.optInclBoth.Size = New System.Drawing.Size(89, 17)
        Me.optInclBoth.TabIndex = 4
        Me.optInclBoth.TabStop = True
        Me.optInclBoth.Text = "Splits && Holes"
        Me.optInclBoth.UseVisualStyleBackColor = True
        '
        'optInclSplits
        '
        Me.optInclSplits.AutoSize = True
        Me.optInclSplits.Location = New System.Drawing.Point(86, 48)
        Me.optInclSplits.Name = "optInclSplits"
        Me.optInclSplits.Size = New System.Drawing.Size(50, 17)
        Me.optInclSplits.TabIndex = 3
        Me.optInclSplits.TabStop = True
        Me.optInclSplits.Text = "Splits"
        Me.optInclSplits.UseVisualStyleBackColor = True
        '
        'optInclComposites
        '
        Me.optInclComposites.AutoSize = True
        Me.optInclComposites.Location = New System.Drawing.Point(20, 47)
        Me.optInclComposites.Name = "optInclComposites"
        Me.optInclComposites.Size = New System.Drawing.Size(52, 17)
        Me.optInclComposites.TabIndex = 2
        Me.optInclComposites.TabStop = True
        Me.optInclComposites.Text = "Holes"
        Me.optInclComposites.UseVisualStyleBackColor = True
        '
        'txtSurvCaddTextfile
        '
        Me.txtSurvCaddTextfile.Location = New System.Drawing.Point(100, 17)
        Me.txtSurvCaddTextfile.Name = "txtSurvCaddTextfile"
        Me.txtSurvCaddTextfile.Size = New System.Drawing.Size(396, 20)
        Me.txtSurvCaddTextfile.TabIndex = 1
        '
        'lblSurvCaddTxtFile
        '
        Me.lblSurvCaddTxtFile.AutoSize = True
        Me.lblSurvCaddTxtFile.Location = New System.Drawing.Point(27, 21)
        Me.lblSurvCaddTxtFile.Name = "lblSurvCaddTxtFile"
        Me.lblSurvCaddTxtFile.Size = New System.Drawing.Size(70, 13)
        Me.lblSurvCaddTxtFile.TabIndex = 0
        Me.lblSurvCaddTxtFile.Text = "Textfile name"
        '
        'cmdExit
        '
        Me.cmdExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdExit.Location = New System.Drawing.Point(1109, 710)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(66, 23)
        Me.cmdExit.TabIndex = 8
        Me.cmdExit.Text = "Exit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'frmProspDataHoleReduction
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1202, 732)
        Me.Controls.Add(Me.cmdPrtScr)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.fraSurvCadd)
        Me.Controls.Add(Me.fraSaveToMois)
        Me.Controls.Add(Me.sbrMain)
        Me.Controls.Add(Me.fraSelect)
        Me.Controls.Add(Me.fraResults)
        Me.Name = "frmProspDataHoleReduction"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Mining OIS -- Prospect Data Single Hole Reduction  (Split Minability Determinatio" &
    "n Utility)"
        Me.fraResults.ResumeLayout(False)
        Me.fraResults.PerformLayout()
        Me.fraMode.ResumeLayout(False)
        Me.fraMode.PerformLayout()
        CType(Me.ssDrillData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabDisp.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.ssHoleData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraMiscStuff.ResumeLayout(False)
        CType(Me.ssRawProspMin, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout
        CType(Me.ssSplitData,System.ComponentModel.ISupportInitialize).EndInit
        Me.TabPage3.ResumeLayout(false)
        Me.tblReductionData.ResumeLayout(false)
        Me.tblReductionData.PerformLayout
        CType(Me.ssSplitReview,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.ssCompReview,System.ComponentModel.ISupportInitialize).EndInit
        Me.TabPage4.ResumeLayout(false)
        Me.TabPage4.PerformLayout
        CType(Me.ssHoleExistStatus,System.ComponentModel.ISupportInitialize).EndInit
        Me.TabPage5.ResumeLayout(false)
        Me.TabPage5.PerformLayout
        Me.TabPage6.ResumeLayout(false)
        CType(Me.ssCompErrors,System.ComponentModel.ISupportInitialize).EndInit
        Me.TabPage7.ResumeLayout(false)
        Me.TabPage7.PerformLayout
        Me.TabPage8.ResumeLayout(false)
        Me.TabPage8.PerformLayout
        CType(Me.ssSplitMinabilities,System.ComponentModel.ISupportInitialize).EndInit
        CType(Me.ssHoleMinabilities,System.ComponentModel.ISupportInitialize).EndInit
        Me.TabPage9.ResumeLayout(false)
        Me.TabPage9.PerformLayout
        CType(Me.ssFeAdjustment,System.ComponentModel.ISupportInitialize).EndInit
        Me.TabPage10.ResumeLayout(false)
        Me.TabPage10.PerformLayout
        CType(Me.ssSplitOverride,System.ComponentModel.ISupportInitialize).EndInit
        Me.Frame2.ResumeLayout(false)
        Me.Frame2.PerformLayout
        CType(Me.ssSplitOverrides,System.ComponentModel.ISupportInitialize).EndInit
        Me.fraSelect.ResumeLayout(false)
        Me.fraSelect.PerformLayout
        Me.fraSaveToMois.ResumeLayout(false)
        Me.fraRdctnType.ResumeLayout(false)
        Me.fraRdctnType.PerformLayout
        Me.fraSurvCadd.ResumeLayout(false)
        Me.fraSurvCadd.PerformLayout
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents fraResults As System.Windows.Forms.GroupBox
    Friend WithEvents tabDisp As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents ssCompReview As AxFPSpread.AxvaSpread
    Friend WithEvents ssSplitReview As AxFPSpread.AxvaSpread
    Friend WithEvents lblRdctnHole As System.Windows.Forms.Label
    Friend WithEvents lblRdctnSplit As System.Windows.Forms.Label
    Friend WithEvents lblCurrSplit As System.Windows.Forms.Label
    Friend WithEvents lblSplit As System.Windows.Forms.Label
    Friend WithEvents ssSplitData As AxFPSpread.AxvaSpread
    Friend WithEvents ssHoleData As AxFPSpread.AxvaSpread
    Friend WithEvents lblUserMadeHoleUnmineable As System.Windows.Forms.Label
    Friend WithEvents lblHole As System.Windows.Forms.Label
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage8 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage9 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage10 As System.Windows.Forms.TabPage
    Friend WithEvents cmdRefresh As System.Windows.Forms.Button
    Friend WithEvents cboSplitOverrideMineName As System.Windows.Forms.ComboBox
    Friend WithEvents Frame2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdGetSplitOverrides As System.Windows.Forms.Button
    Friend WithEvents chkOnlyMySplitOverride As System.Windows.Forms.CheckBox
    Friend WithEvents ssSplitOverrides As AxFPSpread.AxvaSpread
    Friend WithEvents txtSplitOverrideName As System.Windows.Forms.TextBox
    Friend WithEvents cmdAddToOverrideSet As System.Windows.Forms.Button
    Friend WithEvents ssSplitOverride As AxFPSpread.AxvaSpread
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblSplitOverrideSet As System.Windows.Forms.Label
    Friend WithEvents ssHoleMinabilities As AxFPSpread.AxvaSpread
    Friend WithEvents lblHoleMinability As System.Windows.Forms.Label
    Friend WithEvents lblSplitMinabilities As System.Windows.Forms.Label
    Friend WithEvents ssSplitMinabilities As AxFPSpread.AxvaSpread
    Friend WithEvents lblCurrMinabilityComm As System.Windows.Forms.Label
    Friend WithEvents chkUseFeAdjust As System.Windows.Forms.CheckBox
    Friend WithEvents ssFeAdjustment As AxFPSpread.AxvaSpread
    Friend WithEvents lblFeAdjComm As System.Windows.Forms.Label
    Friend WithEvents fraMiscStuff As System.Windows.Forms.GroupBox
    Friend WithEvents cmdPrtGrd As System.Windows.Forms.Button
    Friend WithEvents ssCompErrors As AxFPSpread.AxvaSpread
    Friend WithEvents cmdSaveAreaName As System.Windows.Forms.Button
    Friend WithEvents txtAreaName As System.Windows.Forms.TextBox
    Friend WithEvents lblAreaName As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblInfoComm As System.Windows.Forms.Label
    Friend WithEvents ssHoleExistStatus As AxFPSpread.AxvaSpread
    Friend WithEvents lblHoleExistStatus As System.Windows.Forms.Label
    Friend WithEvents lblHoleInMOIS As System.Windows.Forms.Label
    Friend WithEvents fraSelect As System.Windows.Forms.GroupBox
    Friend WithEvents lblAlphaNumeric As System.Windows.Forms.Label
    Friend WithEvents lblOffSpecPbMgPlt As System.Windows.Forms.Label
    Friend WithEvents lblHoleInMoisComm As System.Windows.Forms.Label
    Friend WithEvents lblScenComm As System.Windows.Forms.Label
    Friend WithEvents lblHoleLocation As System.Windows.Forms.Label
    Friend WithEvents lblRange As System.Windows.Forms.Label
    Friend WithEvents lblSection As System.Windows.Forms.Label
    Friend WithEvents lblTownship As System.Windows.Forms.Label
    Friend WithEvents cboHole As System.Windows.Forms.ComboBox
    Friend WithEvents cboRge As System.Windows.Forms.ComboBox
    Friend WithEvents cboSec As System.Windows.Forms.ComboBox
    Friend WithEvents cboTwp As System.Windows.Forms.ComboBox
    Friend WithEvents lblOtherDefn As System.Windows.Forms.Label
    Friend WithEvents lblIpComm As System.Windows.Forms.Label
    Friend WithEvents lblUseOrigHoleComm As System.Windows.Forms.Label
    Friend WithEvents cmdReduceHole As System.Windows.Forms.Button
    Friend WithEvents chkMyParams As System.Windows.Forms.CheckBox
    Friend WithEvents cmdRefreshParams As System.Windows.Forms.Button
    Friend WithEvents cboOtherDefn As System.Windows.Forms.ComboBox
    Friend WithEvents cboProdSizeDefn As System.Windows.Forms.ComboBox
    Friend WithEvents lblProdSizeDefn As System.Windows.Forms.Label
    Friend WithEvents cmdTest3 As System.Windows.Forms.Button
    Friend WithEvents cmdTest2 As System.Windows.Forms.Button
    Friend WithEvents cmdTest As System.Windows.Forms.Button
    Friend WithEvents chkOverrideMaxDepth As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseOrigHole As System.Windows.Forms.CheckBox
    Friend WithEvents cmdPrtScr As System.Windows.Forms.Button
    Friend WithEvents ssDrillData As AxFPSpread.AxvaSpread
    Friend WithEvents lblCoordsElev As System.Windows.Forms.Label
    Friend WithEvents lblOvbComm As System.Windows.Forms.Label
    Friend WithEvents lblMiscComm2 As System.Windows.Forms.Label
    Friend WithEvents lblMiscComm As System.Windows.Forms.Label
    Friend WithEvents lblMaxDepthComm As System.Windows.Forms.Label
    Friend WithEvents lblGoTo As System.Windows.Forms.Label
    Friend WithEvents fraMode As System.Windows.Forms.GroupBox
    Friend WithEvents optCatalog As System.Windows.Forms.RadioButton
    Friend WithEvents opt100Pct As System.Windows.Forms.RadioButton
    Friend WithEvents cmdPrintSplit As System.Windows.Forms.Button
    Friend WithEvents cmdPrintHole As System.Windows.Forms.Button
    Friend WithEvents cmdMakeHoleUnmineable As System.Windows.Forms.Button
    Friend WithEvents cmdProspSec As System.Windows.Forms.Button
    Friend WithEvents cmdViewCompSplit As System.Windows.Forms.Button
    Friend WithEvents cmdViewRawProsp As System.Windows.Forms.Button
    Friend WithEvents sbrMain As System.Windows.Forms.StatusStrip
    Friend WithEvents fraSaveToMois As System.Windows.Forms.GroupBox
    Friend WithEvents fraRdctnType As System.Windows.Forms.GroupBox
    Friend WithEvents cboMineName As System.Windows.Forms.ComboBox
    Friend WithEvents chkSaveRawProspectMinabilities As System.Windows.Forms.CheckBox
    Friend WithEvents cmdSaveCompAndSplits As System.Windows.Forms.Button
    Friend WithEvents optBothRdctn As System.Windows.Forms.RadioButton
    Friend WithEvents optCatalogRdctn As System.Windows.Forms.RadioButton
    Friend WithEvents opt100PctRdctn As System.Windows.Forms.RadioButton
    Friend WithEvents lblSaveToMoisComm As System.Windows.Forms.Label
    Friend WithEvents cmdSaveMinabilities As System.Windows.Forms.Button
    Friend WithEvents fraSurvCadd As System.Windows.Forms.GroupBox
    Friend WithEvents txtSurvCaddTextfile As System.Windows.Forms.TextBox
    Friend WithEvents lblSurvCaddTxtFile As System.Windows.Forms.Label
    Friend WithEvents optInclBoth As System.Windows.Forms.RadioButton
    Friend WithEvents optInclSplits As System.Windows.Forms.RadioButton
    Friend WithEvents optInclComposites As System.Windows.Forms.RadioButton
    Friend WithEvents chkPbAnalysisFillInSpecial As System.Windows.Forms.CheckBox
    Friend WithEvents cmdCreateSurvCadd As System.Windows.Forms.Button
    Friend WithEvents lblSurvCaddComm As System.Windows.Forms.Label
    Friend WithEvents ssRawProspMin As AxFPSpread.AxvaSpread
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents tblReductionData As TableLayoutPanel
End Class
