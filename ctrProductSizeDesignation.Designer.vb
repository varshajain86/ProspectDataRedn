<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ctrProductSizeDesignation
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ctrProductSizeDesignation))
        Me.ErrorProvider = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.ProductSizeBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DetailsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.grpProductDesignationDetail = New DevExpress.XtraEditors.GroupControl()
        Me.pnlDetails = New System.Windows.Forms.Panel()
        Me.txtDefinedOn = New System.Windows.Forms.TextBox()
        Me.lblDefinedOn = New System.Windows.Forms.Label()
        Me.txtDefinedBy = New System.Windows.Forms.TextBox()
        Me.lblDefinedBy = New System.Windows.Forms.Label()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.lblName = New System.Windows.Forms.Label()
        Me.lblMineName = New System.Windows.Forms.Label()
        Me.cboMineName = New System.Windows.Forms.ComboBox()
        Me.grpSFCDistribution = New DevExpress.XtraEditors.GroupControl()
        Me.btnPrintGrid = New System.Windows.Forms.Button()
        Me.grdSFCDistribution = New DevExpress.XtraGrid.GridControl()
        Me.grdSFCDistributionView = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.colSFCCode = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colSFCDescription = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colMaterial = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colOversize = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.checkEdit = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.colCoarsePb = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colFinePb = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colIp = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colCoarseFd_Cn = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colFineFd_Cn = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colClay = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.cboProductGroup = New System.Windows.Forms.ComboBox()
        Me.lblProductGroup = New System.Windows.Forms.Label()
        Me.tspPSize = New System.Windows.Forms.ToolStrip()
        Me.btnAddNew = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnSave = New System.Windows.Forms.ToolStripButton()
        Me.btnSaveAs = New System.Windows.Forms.ToolStripButton()
        Me.separator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnDelete = New System.Windows.Forms.ToolStripButton()
        Me.separator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnCancel = New System.Windows.Forms.ToolStripButton()
        Me.grpProdSizes = New DevExpress.XtraEditors.GroupControl()
        Me.btnGetProductDesignations = New System.Windows.Forms.Button()
        Me.chkMyDesignationsOnly = New System.Windows.Forms.CheckBox()
        Me.grdProductDesignation = New DevExpress.XtraGrid.GridControl()
        Me.grdProductDesignationView = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.colName = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colWho = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colWhen = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colMine = New DevExpress.XtraGrid.Columns.GridColumn()
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ProductSizeBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DetailsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpProductDesignationDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpProductDesignationDetail.SuspendLayout()
        Me.pnlDetails.SuspendLayout()
        CType(Me.grpSFCDistribution, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpSFCDistribution.SuspendLayout()
        CType(Me.grdSFCDistribution, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdSFCDistributionView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.checkEdit, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tspPSize.SuspendLayout()
        CType(Me.grpProdSizes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpProdSizes.SuspendLayout()
        CType(Me.grdProductDesignation, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdProductDesignationView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ErrorProvider
        '
        Me.ErrorProvider.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink
        Me.ErrorProvider.ContainerControl = Me
        Me.ErrorProvider.DataSource = Me.ProductSizeBindingSource
        '
        'ProductSizeBindingSource
        '
        Me.ProductSizeBindingSource.DataSource = GetType(ProspectDataReduction.ViewModels.ProductSizeDesignation)
        '
        'DetailsBindingSource
        '
        Me.DetailsBindingSource.DataSource = GetType(ProspectDataReduction.ViewModels.ProductSFCDistribution)
        '
        'grpProductDesignationDetail
        '
        Me.grpProductDesignationDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpProductDesignationDetail.AppearanceCaption.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.grpProductDesignationDetail.AppearanceCaption.Options.UseFont = True
        Me.grpProductDesignationDetail.Controls.Add(Me.pnlDetails)
        Me.grpProductDesignationDetail.Controls.Add(Me.tspPSize)
        Me.grpProductDesignationDetail.Location = New System.Drawing.Point(594, 12)
        Me.grpProductDesignationDetail.Name = "grpProductDesignationDetail"
        Me.grpProductDesignationDetail.Size = New System.Drawing.Size(751, 540)
        Me.grpProductDesignationDetail.TabIndex = 6
        Me.grpProductDesignationDetail.Text = "Product Size Designation Detail"
        '
        'pnlDetails
        '
        Me.pnlDetails.Controls.Add(Me.txtDefinedOn)
        Me.pnlDetails.Controls.Add(Me.lblDefinedOn)
        Me.pnlDetails.Controls.Add(Me.txtDefinedBy)
        Me.pnlDetails.Controls.Add(Me.lblDefinedBy)
        Me.pnlDetails.Controls.Add(Me.txtName)
        Me.pnlDetails.Controls.Add(Me.lblName)
        Me.pnlDetails.Controls.Add(Me.lblMineName)
        Me.pnlDetails.Controls.Add(Me.cboMineName)
        Me.pnlDetails.Controls.Add(Me.grpSFCDistribution)
        Me.pnlDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlDetails.Enabled = False
        Me.pnlDetails.Location = New System.Drawing.Point(2, 45)
        Me.pnlDetails.Name = "pnlDetails"
        Me.pnlDetails.Size = New System.Drawing.Size(747, 493)
        Me.pnlDetails.TabIndex = 30
        '
        'txtDefinedOn
        '
        Me.txtDefinedOn.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDefinedOn.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.ProductSizeBindingSource, "WhenDefined", True))
        Me.txtDefinedOn.Location = New System.Drawing.Point(599, 37)
        Me.txtDefinedOn.Name = "txtDefinedOn"
        Me.txtDefinedOn.ReadOnly = True
        Me.txtDefinedOn.Size = New System.Drawing.Size(135, 21)
        Me.txtDefinedOn.TabIndex = 29
        Me.txtDefinedOn.TabStop = False
        '
        'lblDefinedOn
        '
        Me.lblDefinedOn.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDefinedOn.AutoSize = True
        Me.lblDefinedOn.Location = New System.Drawing.Point(526, 40)
        Me.lblDefinedOn.Name = "lblDefinedOn"
        Me.lblDefinedOn.Size = New System.Drawing.Size(59, 13)
        Me.lblDefinedOn.TabIndex = 28
        Me.lblDefinedOn.Text = "Defined on"
        '
        'txtDefinedBy
        '
        Me.txtDefinedBy.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDefinedBy.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.ProductSizeBindingSource, "WhoDefined", True))
        Me.txtDefinedBy.Location = New System.Drawing.Point(599, 10)
        Me.txtDefinedBy.Name = "txtDefinedBy"
        Me.txtDefinedBy.ReadOnly = True
        Me.txtDefinedBy.Size = New System.Drawing.Size(135, 21)
        Me.txtDefinedBy.TabIndex = 25
        Me.txtDefinedBy.TabStop = False
        '
        'lblDefinedBy
        '
        Me.lblDefinedBy.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDefinedBy.AutoSize = True
        Me.lblDefinedBy.Location = New System.Drawing.Point(526, 13)
        Me.lblDefinedBy.Name = "lblDefinedBy"
        Me.lblDefinedBy.Size = New System.Drawing.Size(59, 13)
        Me.lblDefinedBy.TabIndex = 26
        Me.lblDefinedBy.Text = "Defined by"
        '
        'txtName
        '
        Me.txtName.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.ProductSizeBindingSource, "ProductSizeDesignationName", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.txtName.Location = New System.Drawing.Point(74, 10)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(298, 21)
        Me.txtName.TabIndex = 23
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Location = New System.Drawing.Point(23, 13)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(34, 13)
        Me.lblName.TabIndex = 22
        Me.lblName.Text = "Name"
        '
        'lblMineName
        '
        Me.lblMineName.AutoSize = True
        Me.lblMineName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMineName.Location = New System.Drawing.Point(27, 40)
        Me.lblMineName.Name = "lblMineName"
        Me.lblMineName.Size = New System.Drawing.Size(30, 13)
        Me.lblMineName.TabIndex = 19
        Me.lblMineName.Text = "Mine"
        '
        'cboMineName
        '
        Me.cboMineName.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.ProductSizeBindingSource, "MineName", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.cboMineName.DisplayMember = "Name"
        Me.cboMineName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMineName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboMineName.FormattingEnabled = True
        Me.cboMineName.Location = New System.Drawing.Point(74, 37)
        Me.cboMineName.Name = "cboMineName"
        Me.cboMineName.Size = New System.Drawing.Size(298, 21)
        Me.cboMineName.TabIndex = 18
        Me.cboMineName.ValueMember = "Name"
        '
        'grpSFCDistribution
        '
        Me.grpSFCDistribution.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpSFCDistribution.AppearanceCaption.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.grpSFCDistribution.AppearanceCaption.Options.UseFont = True
        Me.grpSFCDistribution.Controls.Add(Me.btnPrintGrid)
        Me.grpSFCDistribution.Controls.Add(Me.grdSFCDistribution)
        Me.grpSFCDistribution.Controls.Add(Me.cboProductGroup)
        Me.grpSFCDistribution.Controls.Add(Me.lblProductGroup)
        Me.grpSFCDistribution.Location = New System.Drawing.Point(12, 75)
        Me.grpSFCDistribution.Name = "grpSFCDistribution"
        Me.grpSFCDistribution.Size = New System.Drawing.Size(723, 402)
        Me.grpSFCDistribution.TabIndex = 1
        Me.grpSFCDistribution.Text = "Size Fraction Distribution"
        '
        'btnPrintGrid
        '
        Me.btnPrintGrid.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrintGrid.Location = New System.Drawing.Point(631, 26)
        Me.btnPrintGrid.Name = "btnPrintGrid"
        Me.btnPrintGrid.Size = New System.Drawing.Size(75, 23)
        Me.btnPrintGrid.TabIndex = 7
        Me.btnPrintGrid.Text = "Print Grid"
        Me.btnPrintGrid.UseVisualStyleBackColor = True
        '
        'grdSFCDistribution
        '
        Me.grdSFCDistribution.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grdSFCDistribution.DataSource = Me.DetailsBindingSource
        Me.grdSFCDistribution.Location = New System.Drawing.Point(14, 58)
        Me.grdSFCDistribution.MainView = Me.grdSFCDistributionView
        Me.grdSFCDistribution.Name = "grdSFCDistribution"
        Me.grdSFCDistribution.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.checkEdit})
        Me.grdSFCDistribution.Size = New System.Drawing.Size(692, 330)
        Me.grdSFCDistribution.TabIndex = 2
        Me.grdSFCDistribution.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.grdSFCDistributionView})
        '
        'grdSFCDistributionView
        '
        Me.grdSFCDistributionView.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.colSFCCode, Me.colSFCDescription, Me.colMaterial, Me.colOversize, Me.colCoarsePb, Me.colFinePb, Me.colIp, Me.colCoarseFd_Cn, Me.colFineFd_Cn, Me.colClay})
        Me.grdSFCDistributionView.GridControl = Me.grdSFCDistribution
        Me.grdSFCDistributionView.IndicatorWidth = 33
        Me.grdSFCDistributionView.Name = "grdSFCDistributionView"
        Me.grdSFCDistributionView.OptionsCustomization.AllowFilter = False
        Me.grdSFCDistributionView.OptionsCustomization.AllowSort = False
        Me.grdSFCDistributionView.OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never
        Me.grdSFCDistributionView.OptionsView.ShowGroupPanel = False
        '
        'colSFCCode
        '
        Me.colSFCCode.AppearanceHeader.Options.UseTextOptions = True
        Me.colSFCCode.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSFCCode.Caption = "SFC"
        Me.colSFCCode.FieldName = "SFCCode"
        Me.colSFCCode.Name = "colSFCCode"
        Me.colSFCCode.OptionsColumn.AllowEdit = False
        Me.colSFCCode.Visible = True
        Me.colSFCCode.VisibleIndex = 0
        Me.colSFCCode.Width = 56
        '
        'colSFCDescription
        '
        Me.colSFCDescription.AppearanceHeader.Options.UseTextOptions = True
        Me.colSFCDescription.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSFCDescription.Caption = "Description"
        Me.colSFCDescription.FieldName = "SFCDescription"
        Me.colSFCDescription.Name = "colSFCDescription"
        Me.colSFCDescription.OptionsColumn.AllowEdit = False
        Me.colSFCDescription.Visible = True
        Me.colSFCDescription.VisibleIndex = 1
        Me.colSFCDescription.Width = 131
        '
        'colMaterial
        '
        Me.colMaterial.AppearanceHeader.Options.UseTextOptions = True
        Me.colMaterial.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colMaterial.Caption = "Material"
        Me.colMaterial.FieldName = "Material"
        Me.colMaterial.Name = "colMaterial"
        Me.colMaterial.OptionsColumn.AllowEdit = False
        Me.colMaterial.Visible = True
        Me.colMaterial.VisibleIndex = 2
        Me.colMaterial.Width = 55
        '
        'colOversize
        '
        Me.colOversize.AppearanceCell.Options.UseTextOptions = True
        Me.colOversize.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colOversize.AppearanceHeader.Options.UseTextOptions = True
        Me.colOversize.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colOversize.Caption = "Oversize"
        Me.colOversize.ColumnEdit = Me.checkEdit
        Me.colOversize.FieldName = "IsOversize"
        Me.colOversize.Name = "colOversize"
        Me.colOversize.Visible = True
        Me.colOversize.VisibleIndex = 3
        Me.colOversize.Width = 62
        '
        'checkEdit
        '
        Me.checkEdit.AutoHeight = False
        Me.checkEdit.Name = "checkEdit"
        '
        'colCoarsePb
        '
        Me.colCoarsePb.AppearanceCell.Options.UseTextOptions = True
        Me.colCoarsePb.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCoarsePb.AppearanceHeader.Options.UseTextOptions = True
        Me.colCoarsePb.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCoarsePb.Caption = "Coarse Pb"
        Me.colCoarsePb.ColumnEdit = Me.checkEdit
        Me.colCoarsePb.FieldName = "IsCoarsePb"
        Me.colCoarsePb.Name = "colCoarsePb"
        Me.colCoarsePb.Visible = True
        Me.colCoarsePb.VisibleIndex = 4
        Me.colCoarsePb.Width = 60
        '
        'colFinePb
        '
        Me.colFinePb.AppearanceCell.Options.UseTextOptions = True
        Me.colFinePb.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colFinePb.AppearanceHeader.Options.UseTextOptions = True
        Me.colFinePb.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colFinePb.Caption = "Fine Pb"
        Me.colFinePb.ColumnEdit = Me.checkEdit
        Me.colFinePb.FieldName = "IsFinePb"
        Me.colFinePb.Name = "colFinePb"
        Me.colFinePb.Visible = True
        Me.colFinePb.VisibleIndex = 5
        Me.colFinePb.Width = 52
        '
        'colIp
        '
        Me.colIp.AppearanceCell.Options.UseTextOptions = True
        Me.colIp.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colIp.AppearanceHeader.Options.UseTextOptions = True
        Me.colIp.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colIp.Caption = "IP"
        Me.colIp.ColumnEdit = Me.checkEdit
        Me.colIp.FieldName = "IsIp"
        Me.colIp.Name = "colIp"
        Me.colIp.Visible = True
        Me.colIp.VisibleIndex = 6
        Me.colIp.Width = 47
        '
        'colCoarseFd_Cn
        '
        Me.colCoarseFd_Cn.AppearanceCell.Options.UseTextOptions = True
        Me.colCoarseFd_Cn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCoarseFd_Cn.AppearanceHeader.Options.UseTextOptions = True
        Me.colCoarseFd_Cn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colCoarseFd_Cn.Caption = "Coarse Fd-Cn"
        Me.colCoarseFd_Cn.ColumnEdit = Me.checkEdit
        Me.colCoarseFd_Cn.FieldName = "IsCoarseFd"
        Me.colCoarseFd_Cn.Name = "colCoarseFd_Cn"
        Me.colCoarseFd_Cn.Visible = True
        Me.colCoarseFd_Cn.VisibleIndex = 7
        Me.colCoarseFd_Cn.Width = 76
        '
        'colFineFd_Cn
        '
        Me.colFineFd_Cn.AppearanceCell.Options.UseTextOptions = True
        Me.colFineFd_Cn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colFineFd_Cn.AppearanceHeader.Options.UseTextOptions = True
        Me.colFineFd_Cn.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colFineFd_Cn.Caption = "Fine Fd-Cn"
        Me.colFineFd_Cn.ColumnEdit = Me.checkEdit
        Me.colFineFd_Cn.FieldName = "IsFineFd"
        Me.colFineFd_Cn.Name = "colFineFd_Cn"
        Me.colFineFd_Cn.Visible = True
        Me.colFineFd_Cn.VisibleIndex = 8
        Me.colFineFd_Cn.Width = 63
        '
        'colClay
        '
        Me.colClay.AppearanceCell.Options.UseTextOptions = True
        Me.colClay.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colClay.AppearanceHeader.Options.UseTextOptions = True
        Me.colClay.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colClay.Caption = "Clay"
        Me.colClay.ColumnEdit = Me.checkEdit
        Me.colClay.FieldName = "IsClay"
        Me.colClay.Name = "colClay"
        Me.colClay.Visible = True
        Me.colClay.VisibleIndex = 9
        Me.colClay.Width = 55
        '
        'cboProductGroup
        '
        Me.cboProductGroup.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.ProductSizeBindingSource, "SizeFractionDistribution", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.cboProductGroup.DataBindings.Add(New System.Windows.Forms.Binding("SelectedItem", Me.ProductSizeBindingSource, "SizeFractionDistribution", True))
        Me.cboProductGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProductGroup.FormattingEnabled = True
        Me.cboProductGroup.Location = New System.Drawing.Point(259, 28)
        Me.cboProductGroup.Name = "cboProductGroup"
        Me.cboProductGroup.Size = New System.Drawing.Size(209, 21)
        Me.cboProductGroup.TabIndex = 1
        '
        'lblProductGroup
        '
        Me.lblProductGroup.AutoSize = True
        Me.lblProductGroup.Location = New System.Drawing.Point(15, 34)
        Me.lblProductGroup.Name = "lblProductGroup"
        Me.lblProductGroup.Size = New System.Drawing.Size(239, 13)
        Me.lblProductGroup.TabIndex = 0
        Me.lblProductGroup.Text = "Based on product group size fraction distribution"
        '
        'tspPSize
        '
        Me.tspPSize.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnAddNew, Me.ToolStripSeparator1, Me.btnSave, Me.btnSaveAs, Me.separator3, Me.btnDelete, Me.separator2, Me.btnCancel})
        Me.tspPSize.Location = New System.Drawing.Point(2, 20)
        Me.tspPSize.Name = "tspPSize"
        Me.tspPSize.Size = New System.Drawing.Size(747, 25)
        Me.tspPSize.TabIndex = 24
        Me.tspPSize.Text = "ToolStrip1"
        '
        'btnAddNew
        '
        Me.btnAddNew.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnAddNew.Image = CType(resources.GetObject("btnAddNew.Image"), System.Drawing.Image)
        Me.btnAddNew.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(60, 22)
        Me.btnAddNew.Text = "Add New"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'btnSave
        '
        Me.btnSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnSave.Image = CType(resources.GetObject("btnSave.Image"), System.Drawing.Image)
        Me.btnSave.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(35, 22)
        Me.btnSave.Text = "Save"
        '
        'btnSaveAs
        '
        Me.btnSaveAs.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnSaveAs.Image = CType(resources.GetObject("btnSaveAs.Image"), System.Drawing.Image)
        Me.btnSaveAs.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnSaveAs.Name = "btnSaveAs"
        Me.btnSaveAs.Size = New System.Drawing.Size(51, 22)
        Me.btnSaveAs.Text = "Save As"
        Me.btnSaveAs.Visible = False
        '
        'separator3
        '
        Me.separator3.Name = "separator3"
        Me.separator3.Size = New System.Drawing.Size(6, 25)
        '
        'btnDelete
        '
        Me.btnDelete.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnDelete.Image = CType(resources.GetObject("btnDelete.Image"), System.Drawing.Image)
        Me.btnDelete.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(44, 22)
        Me.btnDelete.Text = "Delete"
        '
        'separator2
        '
        Me.separator2.Name = "separator2"
        Me.separator2.Size = New System.Drawing.Size(6, 25)
        '
        'btnCancel
        '
        Me.btnCancel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.btnCancel.Image = CType(resources.GetObject("btnCancel.Image"), System.Drawing.Image)
        Me.btnCancel.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(47, 22)
        Me.btnCancel.Text = "Cancel"
        '
        'grpProdSizes
        '
        Me.grpProdSizes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpProdSizes.AppearanceCaption.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.grpProdSizes.AppearanceCaption.Options.UseFont = True
        Me.grpProdSizes.Controls.Add(Me.btnGetProductDesignations)
        Me.grpProdSizes.Controls.Add(Me.chkMyDesignationsOnly)
        Me.grpProdSizes.Controls.Add(Me.grdProductDesignation)
        Me.grpProdSizes.Location = New System.Drawing.Point(15, 12)
        Me.grpProdSizes.Name = "grpProdSizes"
        Me.grpProdSizes.Size = New System.Drawing.Size(563, 540)
        Me.grpProdSizes.TabIndex = 5
        Me.grpProdSizes.Text = "Product Size Designations"
        '
        'btnGetProductDesignations
        '
        Me.btnGetProductDesignations.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGetProductDesignations.Location = New System.Drawing.Point(341, 23)
        Me.btnGetProductDesignations.Name = "btnGetProductDesignations"
        Me.btnGetProductDesignations.Size = New System.Drawing.Size(204, 23)
        Me.btnGetProductDesignations.TabIndex = 2
        Me.btnGetProductDesignations.Text = "Get Product Size Designations"
        Me.btnGetProductDesignations.UseVisualStyleBackColor = True
        '
        'chkMyDesignationsOnly
        '
        Me.chkMyDesignationsOnly.AutoSize = True
        Me.chkMyDesignationsOnly.Checked = True
        Me.chkMyDesignationsOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMyDesignationsOnly.Location = New System.Drawing.Point(17, 28)
        Me.chkMyDesignationsOnly.Name = "chkMyDesignationsOnly"
        Me.chkMyDesignationsOnly.Size = New System.Drawing.Size(169, 17)
        Me.chkMyDesignationsOnly.TabIndex = 1
        Me.chkMyDesignationsOnly.Text = "My Product Designations Only"
        Me.chkMyDesignationsOnly.UseVisualStyleBackColor = True
        '
        'grdProductDesignation
        '
        Me.grdProductDesignation.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grdProductDesignation.Location = New System.Drawing.Point(17, 51)
        Me.grdProductDesignation.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
        Me.grdProductDesignation.LookAndFeel.UseDefaultLookAndFeel = False
        Me.grdProductDesignation.MainView = Me.grdProductDesignationView
        Me.grdProductDesignation.Name = "grdProductDesignation"
        Me.grdProductDesignation.Size = New System.Drawing.Size(528, 471)
        Me.grdProductDesignation.TabIndex = 0
        Me.grdProductDesignation.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.grdProductDesignationView})
        '
        'grdProductDesignationView
        '
        Me.grdProductDesignationView.Appearance.EvenRow.BackColor = System.Drawing.Color.AliceBlue
        Me.grdProductDesignationView.Appearance.EvenRow.Options.UseBackColor = True
        Me.grdProductDesignationView.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.colName, Me.colWho, Me.colWhen, Me.colMine})
        Me.grdProductDesignationView.GridControl = Me.grdProductDesignation
        Me.grdProductDesignationView.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always
        Me.grdProductDesignationView.IndicatorWidth = 30
        Me.grdProductDesignationView.Name = "grdProductDesignationView"
        Me.grdProductDesignationView.OptionsBehavior.Editable = False
        Me.grdProductDesignationView.OptionsCustomization.AllowFilter = False
        Me.grdProductDesignationView.OptionsCustomization.AllowGroup = False
        Me.grdProductDesignationView.OptionsView.EnableAppearanceEvenRow = True
        Me.grdProductDesignationView.OptionsView.ShowGroupPanel = False
        '
        'colName
        '
        Me.colName.Caption = "Name"
        Me.colName.FieldName = "ProductSizeDefinitionName"
        Me.colName.Name = "colName"
        Me.colName.Visible = True
        Me.colName.VisibleIndex = 0
        Me.colName.Width = 199
        '
        'colWho
        '
        Me.colWho.Caption = "Who"
        Me.colWho.FieldName = "WhoDefined"
        Me.colWho.Name = "colWho"
        Me.colWho.Visible = True
        Me.colWho.VisibleIndex = 1
        Me.colWho.Width = 60
        '
        'colWhen
        '
        Me.colWhen.Caption = "When"
        Me.colWhen.DisplayFormat.FormatString = "d"
        Me.colWhen.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime
        Me.colWhen.FieldName = "WhenDefined"
        Me.colWhen.Name = "colWhen"
        Me.colWhen.Visible = True
        Me.colWhen.VisibleIndex = 2
        Me.colWhen.Width = 64
        '
        'colMine
        '
        Me.colMine.Caption = "Mine/Area"
        Me.colMine.FieldName = "MineName"
        Me.colMine.Name = "colMine"
        Me.colMine.Visible = True
        Me.colMine.VisibleIndex = 3
        Me.colMine.Width = 99
        '
        'ctrProductSizeDesignation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.grpProductDesignationDetail)
        Me.Controls.Add(Me.grpProdSizes)
        Me.Name = "ctrProductSizeDesignation"
        Me.Size = New System.Drawing.Size(1360, 564)
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ProductSizeBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DetailsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpProductDesignationDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpProductDesignationDetail.ResumeLayout(False)
        Me.grpProductDesignationDetail.PerformLayout()
        Me.pnlDetails.ResumeLayout(False)
        Me.pnlDetails.PerformLayout()
        CType(Me.grpSFCDistribution, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpSFCDistribution.ResumeLayout(False)
        Me.grpSFCDistribution.PerformLayout()
        CType(Me.grdSFCDistribution, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdSFCDistributionView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.checkEdit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tspPSize.ResumeLayout(False)
        Me.tspPSize.PerformLayout()
        CType(Me.grpProdSizes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpProdSizes.ResumeLayout(False)
        Me.grpProdSizes.PerformLayout()
        CType(Me.grdProductDesignation, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdProductDesignationView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ErrorProvider As ErrorProvider
    Friend WithEvents ProductSizeBindingSource As BindingSource
    Friend WithEvents DetailsBindingSource As BindingSource
    Friend WithEvents grpProductDesignationDetail As DevExpress.XtraEditors.GroupControl
    Friend WithEvents pnlDetails As Panel
    Friend WithEvents txtDefinedOn As TextBox
    Friend WithEvents lblDefinedOn As Label
    Friend WithEvents txtDefinedBy As TextBox
    Friend WithEvents lblDefinedBy As Label
    Friend WithEvents txtName As TextBox
    Friend WithEvents lblName As Label
    Friend WithEvents lblMineName As Label
    Friend WithEvents cboMineName As ComboBox
    Friend WithEvents grpSFCDistribution As DevExpress.XtraEditors.GroupControl
    Friend WithEvents btnPrintGrid As Button
    Friend WithEvents grdSFCDistribution As DevExpress.XtraGrid.GridControl
    Friend WithEvents grdSFCDistributionView As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents colSFCCode As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colSFCDescription As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colMaterial As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colOversize As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents checkEdit As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents colCoarsePb As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colFinePb As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colIp As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colCoarseFd_Cn As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colFineFd_Cn As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colClay As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents cboProductGroup As ComboBox
    Friend WithEvents lblProductGroup As Label
    Friend WithEvents tspPSize As ToolStrip
    Friend WithEvents btnAddNew As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents btnSave As ToolStripButton
    Friend WithEvents btnSaveAs As ToolStripButton
    Friend WithEvents separator3 As ToolStripSeparator
    Friend WithEvents btnDelete As ToolStripButton
    Friend WithEvents separator2 As ToolStripSeparator
    Friend WithEvents btnCancel As ToolStripButton
    Friend WithEvents grpProdSizes As DevExpress.XtraEditors.GroupControl
    Friend WithEvents btnGetProductDesignations As Button
    Friend WithEvents chkMyDesignationsOnly As CheckBox
    Friend WithEvents grdProductDesignation As DevExpress.XtraGrid.GridControl
    Friend WithEvents grdProductDesignationView As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents colName As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colWho As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colWhen As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colMine As DevExpress.XtraGrid.Columns.GridColumn
End Class
