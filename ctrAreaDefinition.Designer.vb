<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ctrAreaDefinition
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ctrAreaDefinition))
        Me.grpAreaDefinitionDetail = New DevExpress.XtraEditors.GroupControl()
        Me.pnlDetails = New System.Windows.Forms.Panel()
        Me.txtDefinedOn = New System.Windows.Forms.TextBox()
        Me.AreaDefinitionBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.lblDefinedOn = New System.Windows.Forms.Label()
        Me.txtDefinedBy = New System.Windows.Forms.TextBox()
        Me.lblDefinedBy = New System.Windows.Forms.Label()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.lblName = New System.Windows.Forms.Label()
        Me.cboAreaName = New System.Windows.Forms.ComboBox()
        Me.lblAreaName = New System.Windows.Forms.Label()
        Me.lblMineName = New System.Windows.Forms.Label()
        Me.cboMineName = New System.Windows.Forms.ComboBox()
        Me.grpHoleFilters = New DevExpress.XtraEditors.GroupControl()
        Me.cboOriginalForRedrill = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboOwnership = New DevExpress.XtraEditors.GridLookUpEdit()
        Me.GridLookUpEdit1View = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.colOwnershipCode = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colOwnershipDesc = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.dtpAreaEndDrillDate2 = New DevExpress.XtraEditors.DateEdit()
        Me.lblGen2 = New System.Windows.Forms.Label()
        Me.dtpAreaBeginDrillDate2 = New DevExpress.XtraEditors.DateEdit()
        Me.lblOwnership = New System.Windows.Forms.Label()
        Me.lblGen4 = New System.Windows.Forms.Label()
        Me.lblGen32 = New System.Windows.Forms.Label()
        Me.lblGen1 = New System.Windows.Forms.Label()
        Me.lblGen0 = New System.Windows.Forms.Label()
        Me.cboAreaProspHoleType = New System.Windows.Forms.ComboBox()
        Me.cboAreaMinedOutStatus = New System.Windows.Forms.ComboBox()
        Me.grpArea = New DevExpress.XtraEditors.GroupControl()
        Me.grpSelectionMode = New System.Windows.Forms.GroupBox()
        Me.rdoFromFile = New System.Windows.Forms.RadioButton()
        Me.rdoHoles = New System.Windows.Forms.RadioButton()
        Me.rdoMineArea = New System.Windows.Forms.RadioButton()
        Me.rdoXYCorner = New System.Windows.Forms.RadioButton()
        Me.rdoTRSCorner = New System.Windows.Forms.RadioButton()
        Me.grpSelectHoles = New System.Windows.Forms.GroupBox()
        Me.pnlSelectHoles = New System.Windows.Forms.Panel()
        Me.pnlTRSCorner = New System.Windows.Forms.Panel()
        Me.btnClearTRS = New System.Windows.Forms.Button()
        Me.grdTRSCorner = New DevExpress.XtraGrid.GridControl()
        Me.TRSCornerBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.grdTRSCornerView = New DevExpress.XtraGrid.Views.BandedGrid.BandedGridView()
        Me.colTRSCornerSW = New DevExpress.XtraGrid.Views.BandedGrid.GridBand()
        Me.colTRSCornerTownship_SW = New DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn()
        Me.lupTRSTownship = New DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit()
        Me.TownshipBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.colTRSCornerRange_NE = New DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn()
        Me.lupTRSRange = New DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit()
        Me.RangeBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.colTRSCornerSection_NE = New DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn()
        Me.lupTRSSection = New DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit()
        Me.SectionBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.colTRSCornerNE = New DevExpress.XtraGrid.Views.BandedGrid.GridBand()
        Me.colTRSCornerTownship_NE = New DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn()
        Me.colTRSCornerRange_SW = New DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn()
        Me.colTRSCornerSection_SW = New DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn()
        Me.pnlMineArea = New System.Windows.Forms.Panel()
        Me.cboArea = New System.Windows.Forms.ComboBox()
        Me.lblArea = New System.Windows.Forms.Label()
        Me.lblMine = New System.Windows.Forms.Label()
        Me.cboMine = New System.Windows.Forms.ComboBox()
        Me.pnlXYCorner = New System.Windows.Forms.Panel()
        Me.btnClearXY = New System.Windows.Forms.Button()
        Me.grdXYCorner = New DevExpress.XtraGrid.GridControl()
        Me.XYCornerBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.grdXYCornerView = New DevExpress.XtraGrid.Views.BandedGrid.BandedGridView()
        Me.colSW = New DevExpress.XtraGrid.Views.BandedGrid.GridBand()
        Me.colSW_X = New DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn()
        Me.colSW_Y = New DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn()
        Me.colNE = New DevExpress.XtraGrid.Views.BandedGrid.GridBand()
        Me.colNE_X = New DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn()
        Me.colNE_Y = New DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn()
        Me.pnlHoles = New System.Windows.Forms.Panel()
        Me.btnClearHoles = New System.Windows.Forms.Button()
        Me.btnAddHolesFromFile = New System.Windows.Forms.Button()
        Me.btnAddHolesFromProspectGrid = New System.Windows.Forms.Button()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.grdHolesFromFile = New DevExpress.XtraGrid.GridControl()
        Me.HoleBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.grdHolesFromFileView = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.colTownship = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.lupHoleTownship = New DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit()
        Me.colRange = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.lupHoleRange = New DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit()
        Me.colSection = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.lupHoleSection = New DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit()
        Me.colHole = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.lupHoleHole = New DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit()
        Me.HolesListBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.tspArea = New System.Windows.Forms.ToolStrip()
        Me.btnAddNew = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnSave = New System.Windows.Forms.ToolStripButton()
        Me.btnSaveAs = New System.Windows.Forms.ToolStripButton()
        Me.separator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnDelete = New System.Windows.Forms.ToolStripButton()
        Me.separator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnCancel = New System.Windows.Forms.ToolStripButton()
        Me.grpAreaDefinitions = New DevExpress.XtraEditors.GroupControl()
        Me.btnGetAreaDefinitions = New System.Windows.Forms.Button()
        Me.chkMyDefinitionsOnly = New System.Windows.Forms.CheckBox()
        Me.grdAreaDefinition = New DevExpress.XtraGrid.GridControl()
        Me.grdAreaDefinitionView = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.colName = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colWho = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colWhen = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colType = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colMine = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colSubArea = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.mnuAddFrom = New DevExpress.XtraBars.PopupMenu(Me.components)
        Me.btnFromProspectGrid = New DevExpress.XtraBars.BarButtonItem()
        Me.btnFromFile = New DevExpress.XtraBars.BarButtonItem()
        Me.barAddFrom = New DevExpress.XtraBars.BarManager(Me.components)
        Me.barDockControlTop = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlBottom = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlLeft = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlRight = New DevExpress.XtraBars.BarDockControl()
        Me.ProspectCodeBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.ErrorProvider = New System.Windows.Forms.ErrorProvider(Me.components)
        CType(Me.grpAreaDefinitionDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAreaDefinitionDetail.SuspendLayout()
        Me.pnlDetails.SuspendLayout()
        CType(Me.AreaDefinitionBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpHoleFilters, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpHoleFilters.SuspendLayout()
        CType(Me.cboOwnership.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridLookUpEdit1View, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtpAreaEndDrillDate2.Properties.CalendarTimeProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtpAreaEndDrillDate2.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtpAreaBeginDrillDate2.Properties.CalendarTimeProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtpAreaBeginDrillDate2.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpArea, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpArea.SuspendLayout()
        Me.grpSelectionMode.SuspendLayout()
        Me.grpSelectHoles.SuspendLayout()
        Me.pnlSelectHoles.SuspendLayout()
        Me.pnlTRSCorner.SuspendLayout()
        CType(Me.grdTRSCorner, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TRSCornerBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdTRSCornerView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lupTRSTownship, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TownshipBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lupTRSRange, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RangeBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lupTRSSection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SectionBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlMineArea.SuspendLayout()
        Me.pnlXYCorner.SuspendLayout()
        CType(Me.grdXYCorner, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.XYCornerBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdXYCornerView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlHoles.SuspendLayout()
        CType(Me.grdHolesFromFile, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.HoleBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdHolesFromFileView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lupHoleTownship, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lupHoleRange, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lupHoleSection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lupHoleHole, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.HolesListBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tspArea.SuspendLayout()
        CType(Me.grpAreaDefinitions, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpAreaDefinitions.SuspendLayout()
        CType(Me.grdAreaDefinition, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdAreaDefinitionView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mnuAddFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.barAddFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ProspectCodeBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpAreaDefinitionDetail
        '
        Me.grpAreaDefinitionDetail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpAreaDefinitionDetail.AppearanceCaption.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.grpAreaDefinitionDetail.AppearanceCaption.Options.UseFont = True
        Me.grpAreaDefinitionDetail.Controls.Add(Me.pnlDetails)
        Me.grpAreaDefinitionDetail.Controls.Add(Me.tspArea)
        Me.grpAreaDefinitionDetail.Location = New System.Drawing.Point(691, 12)
        Me.grpAreaDefinitionDetail.Name = "grpAreaDefinitionDetail"
        Me.grpAreaDefinitionDetail.Size = New System.Drawing.Size(651, 540)
        Me.grpAreaDefinitionDetail.TabIndex = 4
        Me.grpAreaDefinitionDetail.Text = "Area Definition Detail"
        '
        'pnlDetails
        '
        Me.pnlDetails.Controls.Add(Me.txtDefinedOn)
        Me.pnlDetails.Controls.Add(Me.lblDefinedOn)
        Me.pnlDetails.Controls.Add(Me.txtDefinedBy)
        Me.pnlDetails.Controls.Add(Me.lblDefinedBy)
        Me.pnlDetails.Controls.Add(Me.txtName)
        Me.pnlDetails.Controls.Add(Me.lblName)
        Me.pnlDetails.Controls.Add(Me.cboAreaName)
        Me.pnlDetails.Controls.Add(Me.lblAreaName)
        Me.pnlDetails.Controls.Add(Me.lblMineName)
        Me.pnlDetails.Controls.Add(Me.cboMineName)
        Me.pnlDetails.Controls.Add(Me.grpHoleFilters)
        Me.pnlDetails.Controls.Add(Me.grpArea)
        Me.pnlDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlDetails.Enabled = False
        Me.pnlDetails.Location = New System.Drawing.Point(2, 45)
        Me.pnlDetails.Name = "pnlDetails"
        Me.pnlDetails.Size = New System.Drawing.Size(647, 493)
        Me.pnlDetails.TabIndex = 30
        '
        'txtDefinedOn
        '
        Me.txtDefinedOn.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.AreaDefinitionBindingSource, "WhenDefined", True))
        Me.txtDefinedOn.Location = New System.Drawing.Point(490, 37)
        Me.txtDefinedOn.Name = "txtDefinedOn"
        Me.txtDefinedOn.ReadOnly = True
        Me.txtDefinedOn.Size = New System.Drawing.Size(135, 21)
        Me.txtDefinedOn.TabIndex = 29
        Me.txtDefinedOn.TabStop = False
        '
        'AreaDefinitionBindingSource
        '
        Me.AreaDefinitionBindingSource.DataSource = GetType(ProspectDataReduction.ViewModels.ProspectAreaDefinition)
        '
        'lblDefinedOn
        '
        Me.lblDefinedOn.AutoSize = True
        Me.lblDefinedOn.Location = New System.Drawing.Point(416, 40)
        Me.lblDefinedOn.Name = "lblDefinedOn"
        Me.lblDefinedOn.Size = New System.Drawing.Size(59, 13)
        Me.lblDefinedOn.TabIndex = 28
        Me.lblDefinedOn.Text = "Defined on"
        '
        'txtDefinedBy
        '
        Me.txtDefinedBy.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.AreaDefinitionBindingSource, "WhoDefined", True))
        Me.txtDefinedBy.Location = New System.Drawing.Point(490, 10)
        Me.txtDefinedBy.Name = "txtDefinedBy"
        Me.txtDefinedBy.ReadOnly = True
        Me.txtDefinedBy.Size = New System.Drawing.Size(135, 21)
        Me.txtDefinedBy.TabIndex = 25
        Me.txtDefinedBy.TabStop = False
        '
        'lblDefinedBy
        '
        Me.lblDefinedBy.AutoSize = True
        Me.lblDefinedBy.Location = New System.Drawing.Point(416, 13)
        Me.lblDefinedBy.Name = "lblDefinedBy"
        Me.lblDefinedBy.Size = New System.Drawing.Size(59, 13)
        Me.lblDefinedBy.TabIndex = 26
        Me.lblDefinedBy.Text = "Defined by"
        '
        'txtName
        '
        Me.txtName.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.AreaDefinitionBindingSource, "AreaDefinitionName", True))
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
        'cboAreaName
        '
        Me.cboAreaName.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.AreaDefinitionBindingSource, "SpecAreaName", True))
        Me.cboAreaName.DisplayMember = "Name"
        Me.cboAreaName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAreaName.FormattingEnabled = True
        Me.cboAreaName.Location = New System.Drawing.Point(74, 64)
        Me.cboAreaName.Name = "cboAreaName"
        Me.cboAreaName.Size = New System.Drawing.Size(298, 21)
        Me.cboAreaName.TabIndex = 21
        Me.cboAreaName.ValueMember = "Name"
        '
        'lblAreaName
        '
        Me.lblAreaName.AutoSize = True
        Me.lblAreaName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAreaName.Location = New System.Drawing.Point(28, 67)
        Me.lblAreaName.Name = "lblAreaName"
        Me.lblAreaName.Size = New System.Drawing.Size(29, 13)
        Me.lblAreaName.TabIndex = 20
        Me.lblAreaName.Text = "Area"
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
        Me.cboMineName.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.AreaDefinitionBindingSource, "MineName", True))
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
        'grpHoleFilters
        '
        Me.grpHoleFilters.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.grpHoleFilters.AppearanceCaption.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.grpHoleFilters.AppearanceCaption.Options.UseFont = True
        Me.grpHoleFilters.Controls.Add(Me.cboOriginalForRedrill)
        Me.grpHoleFilters.Controls.Add(Me.Label2)
        Me.grpHoleFilters.Controls.Add(Me.cboOwnership)
        Me.grpHoleFilters.Controls.Add(Me.dtpAreaEndDrillDate2)
        Me.grpHoleFilters.Controls.Add(Me.lblGen2)
        Me.grpHoleFilters.Controls.Add(Me.dtpAreaBeginDrillDate2)
        Me.grpHoleFilters.Controls.Add(Me.lblOwnership)
        Me.grpHoleFilters.Controls.Add(Me.lblGen4)
        Me.grpHoleFilters.Controls.Add(Me.lblGen32)
        Me.grpHoleFilters.Controls.Add(Me.lblGen1)
        Me.grpHoleFilters.Controls.Add(Me.lblGen0)
        Me.grpHoleFilters.Controls.Add(Me.cboAreaProspHoleType)
        Me.grpHoleFilters.Controls.Add(Me.cboAreaMinedOutStatus)
        Me.grpHoleFilters.Location = New System.Drawing.Point(12, 327)
        Me.grpHoleFilters.Name = "grpHoleFilters"
        Me.grpHoleFilters.Size = New System.Drawing.Size(623, 159)
        Me.grpHoleFilters.TabIndex = 2
        Me.grpHoleFilters.Text = "Area Prospect Hole Filters"
        '
        'cboOriginalForRedrill
        '
        Me.cboOriginalForRedrill.DataBindings.Add(New System.Windows.Forms.Binding("SelectedItem", Me.AreaDefinitionBindingSource, "UseOriginalHoleForRedrills", True))
        Me.cboOriginalForRedrill.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboOriginalForRedrill.FormattingEnabled = True
        Me.cboOriginalForRedrill.Items.AddRange(New Object() {"", "Yes", "No"})
        Me.cboOriginalForRedrill.Location = New System.Drawing.Point(150, 126)
        Me.cboOriginalForRedrill.Name = "cboOriginalForRedrill"
        Me.cboOriginalForRedrill.Size = New System.Drawing.Size(121, 21)
        Me.cboOriginalForRedrill.TabIndex = 26
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 131)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 13)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "Use original hole for redrills"
        '
        'cboOwnership
        '
        Me.cboOwnership.DataBindings.Add(New System.Windows.Forms.Binding("EditValue", Me.AreaDefinitionBindingSource, "Ownership", True))
        Me.cboOwnership.EditValue = ""
        Me.cboOwnership.Location = New System.Drawing.Point(345, 28)
        Me.cboOwnership.Name = "cboOwnership"
        Me.cboOwnership.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.cboOwnership.Properties.DisplayMember = "ProspCodeDesc"
        Me.cboOwnership.Properties.NullText = ""
        Me.cboOwnership.Properties.ValueMember = "ProspCode"
        Me.cboOwnership.Properties.View = Me.GridLookUpEdit1View
        Me.cboOwnership.Size = New System.Drawing.Size(260, 20)
        Me.cboOwnership.TabIndex = 24
        '
        'GridLookUpEdit1View
        '
        Me.GridLookUpEdit1View.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.colOwnershipCode, Me.colOwnershipDesc})
        Me.GridLookUpEdit1View.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus
        Me.GridLookUpEdit1View.Name = "GridLookUpEdit1View"
        Me.GridLookUpEdit1View.OptionsSelection.EnableAppearanceFocusedCell = False
        Me.GridLookUpEdit1View.OptionsView.ShowColumnHeaders = False
        Me.GridLookUpEdit1View.OptionsView.ShowGroupPanel = False
        '
        'colOwnershipCode
        '
        Me.colOwnershipCode.Caption = "Code"
        Me.colOwnershipCode.FieldName = "ProspCode"
        Me.colOwnershipCode.Name = "colOwnershipCode"
        Me.colOwnershipCode.Visible = True
        Me.colOwnershipCode.VisibleIndex = 0
        Me.colOwnershipCode.Width = 25
        '
        'colOwnershipDesc
        '
        Me.colOwnershipDesc.Caption = "Description"
        Me.colOwnershipDesc.FieldName = "ProspCodeDesc"
        Me.colOwnershipDesc.Name = "colOwnershipDesc"
        Me.colOwnershipDesc.Visible = True
        Me.colOwnershipDesc.VisibleIndex = 1
        Me.colOwnershipDesc.Width = 359
        '
        'dtpAreaEndDrillDate2
        '
        Me.dtpAreaEndDrillDate2.DataBindings.Add(New System.Windows.Forms.Binding("EditValue", Me.AreaDefinitionBindingSource, "EndDrillDate", True))
        Me.dtpAreaEndDrillDate2.EditValue = Nothing
        Me.dtpAreaEndDrillDate2.Location = New System.Drawing.Point(150, 52)
        Me.dtpAreaEndDrillDate2.Name = "dtpAreaEndDrillDate2"
        Me.dtpAreaEndDrillDate2.Properties.AllowNullInput = DevExpress.Utils.DefaultBoolean.[True]
        Me.dtpAreaEndDrillDate2.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.dtpAreaEndDrillDate2.Properties.CalendarTimeProperties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.dtpAreaEndDrillDate2.Size = New System.Drawing.Size(86, 20)
        Me.dtpAreaEndDrillDate2.TabIndex = 22
        '
        'lblGen2
        '
        Me.lblGen2.AutoSize = True
        Me.lblGen2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen2.Location = New System.Drawing.Point(49, 81)
        Me.lblGen2.Name = "lblGen2"
        Me.lblGen2.Size = New System.Drawing.Size(95, 13)
        Me.lblGen2.TabIndex = 20
        Me.lblGen2.Text = "Prospect hole type"
        '
        'dtpAreaBeginDrillDate2
        '
        Me.dtpAreaBeginDrillDate2.DataBindings.Add(New System.Windows.Forms.Binding("EditValue", Me.AreaDefinitionBindingSource, "BeginningDrillDate", True))
        Me.dtpAreaBeginDrillDate2.EditValue = Nothing
        Me.dtpAreaBeginDrillDate2.Location = New System.Drawing.Point(150, 28)
        Me.dtpAreaBeginDrillDate2.Name = "dtpAreaBeginDrillDate2"
        Me.dtpAreaBeginDrillDate2.Properties.AllowNullInput = DevExpress.Utils.DefaultBoolean.[True]
        Me.dtpAreaBeginDrillDate2.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.dtpAreaBeginDrillDate2.Properties.CalendarTimeProperties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.dtpAreaBeginDrillDate2.Size = New System.Drawing.Size(86, 20)
        Me.dtpAreaBeginDrillDate2.TabIndex = 21
        '
        'lblOwnership
        '
        Me.lblOwnership.AutoSize = True
        Me.lblOwnership.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnership.Location = New System.Drawing.Point(282, 31)
        Me.lblOwnership.Name = "lblOwnership"
        Me.lblOwnership.Size = New System.Drawing.Size(57, 13)
        Me.lblOwnership.TabIndex = 19
        Me.lblOwnership.Text = "Ownership"
        '
        'lblGen4
        '
        Me.lblGen4.AutoSize = True
        Me.lblGen4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen4.Location = New System.Drawing.Point(33, 106)
        Me.lblGen4.Name = "lblGen4"
        Me.lblGen4.Size = New System.Drawing.Size(111, 13)
        Me.lblGen4.TabIndex = 18
        Me.lblGen4.Text = "Skip mined-out holes?"
        '
        'lblGen32
        '
        Me.lblGen32.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblGen32.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen32.ForeColor = System.Drawing.Color.Navy
        Me.lblGen32.Location = New System.Drawing.Point(279, 77)
        Me.lblGen32.Name = "lblGen32"
        Me.lblGen32.Size = New System.Drawing.Size(329, 63)
        Me.lblGen32.TabIndex = 17
        Me.lblGen32.Text = resources.GetString("lblGen32.Text")
        '
        'lblGen1
        '
        Me.lblGen1.AutoSize = True
        Me.lblGen1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen1.Location = New System.Drawing.Point(76, 56)
        Me.lblGen1.Name = "lblGen1"
        Me.lblGen1.Size = New System.Drawing.Size(68, 13)
        Me.lblGen1.TabIndex = 16
        Me.lblGen1.Text = "End drill date"
        '
        'lblGen0
        '
        Me.lblGen0.AutoSize = True
        Me.lblGen0.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGen0.Location = New System.Drawing.Point(48, 31)
        Me.lblGen0.Name = "lblGen0"
        Me.lblGen0.Size = New System.Drawing.Size(96, 13)
        Me.lblGen0.TabIndex = 15
        Me.lblGen0.Text = "Beginning drill date"
        '
        'cboAreaProspHoleType
        '
        Me.cboAreaProspHoleType.DataBindings.Add(New System.Windows.Forms.Binding("SelectedItem", Me.AreaDefinitionBindingSource, "HoleMetLabProcessType", True))
        Me.cboAreaProspHoleType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAreaProspHoleType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboAreaProspHoleType.FormattingEnabled = True
        Me.cboAreaProspHoleType.Items.AddRange(New Object() {"", "Regular", "Expanded"})
        Me.cboAreaProspHoleType.Location = New System.Drawing.Point(150, 76)
        Me.cboAreaProspHoleType.Name = "cboAreaProspHoleType"
        Me.cboAreaProspHoleType.Size = New System.Drawing.Size(121, 21)
        Me.cboAreaProspHoleType.TabIndex = 14
        '
        'cboAreaMinedOutStatus
        '
        Me.cboAreaMinedOutStatus.DataBindings.Add(New System.Windows.Forms.Binding("SelectedItem", Me.AreaDefinitionBindingSource, "MinedStatus", True))
        Me.cboAreaMinedOutStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAreaMinedOutStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboAreaMinedOutStatus.FormattingEnabled = True
        Me.cboAreaMinedOutStatus.Items.AddRange(New Object() {"", "Yes", "No"})
        Me.cboAreaMinedOutStatus.Location = New System.Drawing.Point(150, 101)
        Me.cboAreaMinedOutStatus.Name = "cboAreaMinedOutStatus"
        Me.cboAreaMinedOutStatus.Size = New System.Drawing.Size(121, 21)
        Me.cboAreaMinedOutStatus.TabIndex = 13
        '
        'grpArea
        '
        Me.grpArea.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.grpArea.AppearanceCaption.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.grpArea.AppearanceCaption.Options.UseFont = True
        Me.grpArea.Controls.Add(Me.grpSelectionMode)
        Me.grpArea.Controls.Add(Me.grpSelectHoles)
        Me.grpArea.Location = New System.Drawing.Point(12, 99)
        Me.grpArea.Name = "grpArea"
        Me.grpArea.Size = New System.Drawing.Size(623, 222)
        Me.grpArea.TabIndex = 1
        Me.grpArea.Text = "Select Area"
        '
        'grpSelectionMode
        '
        Me.grpSelectionMode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.grpSelectionMode.Controls.Add(Me.rdoFromFile)
        Me.grpSelectionMode.Controls.Add(Me.rdoHoles)
        Me.grpSelectionMode.Controls.Add(Me.rdoMineArea)
        Me.grpSelectionMode.Controls.Add(Me.rdoXYCorner)
        Me.grpSelectionMode.Controls.Add(Me.rdoTRSCorner)
        Me.grpSelectionMode.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.grpSelectionMode.Location = New System.Drawing.Point(19, 31)
        Me.grpSelectionMode.Name = "grpSelectionMode"
        Me.grpSelectionMode.Size = New System.Drawing.Size(176, 175)
        Me.grpSelectionMode.TabIndex = 4
        Me.grpSelectionMode.TabStop = False
        Me.grpSelectionMode.Text = "Selection Mode"
        '
        'rdoFromFile
        '
        Me.rdoFromFile.AutoSize = True
        Me.rdoFromFile.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.rdoFromFile.Location = New System.Drawing.Point(12, 110)
        Me.rdoFromFile.Name = "rdoFromFile"
        Me.rdoFromFile.Size = New System.Drawing.Size(104, 17)
        Me.rdoFromFile.TabIndex = 4
        Me.rdoFromFile.TabStop = True
        Me.rdoFromFile.Text = "Holes - From File"
        Me.rdoFromFile.UseVisualStyleBackColor = True
        Me.rdoFromFile.Visible = False
        '
        'rdoHoles
        '
        Me.rdoHoles.AutoSize = True
        Me.rdoHoles.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.rdoHoles.Location = New System.Drawing.Point(12, 87)
        Me.rdoHoles.Name = "rdoHoles"
        Me.rdoHoles.Size = New System.Drawing.Size(51, 17)
        Me.rdoHoles.TabIndex = 3
        Me.rdoHoles.TabStop = True
        Me.rdoHoles.Text = "Holes"
        Me.rdoHoles.UseVisualStyleBackColor = True
        '
        'rdoMineArea
        '
        Me.rdoMineArea.AutoSize = True
        Me.rdoMineArea.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.rdoMineArea.Location = New System.Drawing.Point(12, 64)
        Me.rdoMineArea.Name = "rdoMineArea"
        Me.rdoMineArea.Size = New System.Drawing.Size(74, 17)
        Me.rdoMineArea.TabIndex = 2
        Me.rdoMineArea.TabStop = True
        Me.rdoMineArea.Text = "Mine/Area"
        Me.rdoMineArea.UseVisualStyleBackColor = True
        '
        'rdoXYCorner
        '
        Me.rdoXYCorner.AutoSize = True
        Me.rdoXYCorner.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.rdoXYCorner.Location = New System.Drawing.Point(12, 41)
        Me.rdoXYCorner.Name = "rdoXYCorner"
        Me.rdoXYCorner.Size = New System.Drawing.Size(111, 17)
        Me.rdoXYCorner.TabIndex = 1
        Me.rdoXYCorner.TabStop = True
        Me.rdoXYCorner.Text = "X && Y Coordinates"
        Me.rdoXYCorner.UseVisualStyleBackColor = True
        '
        'rdoTRSCorner
        '
        Me.rdoTRSCorner.AutoSize = True
        Me.rdoTRSCorner.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.rdoTRSCorner.Location = New System.Drawing.Point(12, 18)
        Me.rdoTRSCorner.Name = "rdoTRSCorner"
        Me.rdoTRSCorner.Size = New System.Drawing.Size(88, 17)
        Me.rdoTRSCorner.TabIndex = 0
        Me.rdoTRSCorner.Text = "T-R-S Corner"
        Me.rdoTRSCorner.UseVisualStyleBackColor = True
        '
        'grpSelectHoles
        '
        Me.grpSelectHoles.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.grpSelectHoles.Controls.Add(Me.pnlSelectHoles)
        Me.grpSelectHoles.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.grpSelectHoles.Location = New System.Drawing.Point(202, 31)
        Me.grpSelectHoles.Name = "grpSelectHoles"
        Me.grpSelectHoles.Size = New System.Drawing.Size(403, 175)
        Me.grpSelectHoles.TabIndex = 3
        Me.grpSelectHoles.TabStop = False
        Me.grpSelectHoles.Text = "Select T-R-S Corners"
        '
        'pnlSelectHoles
        '
        Me.pnlSelectHoles.AutoScroll = True
        Me.pnlSelectHoles.Controls.Add(Me.pnlTRSCorner)
        Me.pnlSelectHoles.Controls.Add(Me.pnlMineArea)
        Me.pnlSelectHoles.Controls.Add(Me.pnlXYCorner)
        Me.pnlSelectHoles.Controls.Add(Me.pnlHoles)
        Me.pnlSelectHoles.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlSelectHoles.Location = New System.Drawing.Point(3, 17)
        Me.pnlSelectHoles.Name = "pnlSelectHoles"
        Me.pnlSelectHoles.Size = New System.Drawing.Size(397, 155)
        Me.pnlSelectHoles.TabIndex = 0
        '
        'pnlTRSCorner
        '
        Me.pnlTRSCorner.Controls.Add(Me.btnClearTRS)
        Me.pnlTRSCorner.Controls.Add(Me.grdTRSCorner)
        Me.pnlTRSCorner.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTRSCorner.Location = New System.Drawing.Point(0, 465)
        Me.pnlTRSCorner.Name = "pnlTRSCorner"
        Me.pnlTRSCorner.Size = New System.Drawing.Size(380, 155)
        Me.pnlTRSCorner.TabIndex = 18
        '
        'btnClearTRS
        '
        Me.btnClearTRS.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClearTRS.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.btnClearTRS.Location = New System.Drawing.Point(287, 129)
        Me.btnClearTRS.Name = "btnClearTRS"
        Me.btnClearTRS.Size = New System.Drawing.Size(75, 23)
        Me.btnClearTRS.TabIndex = 5
        Me.btnClearTRS.Text = "Clear Grid"
        Me.btnClearTRS.UseVisualStyleBackColor = True
        '
        'grdTRSCorner
        '
        Me.grdTRSCorner.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grdTRSCorner.DataSource = Me.TRSCornerBindingSource
        Me.grdTRSCorner.Location = New System.Drawing.Point(17, 3)
        Me.grdTRSCorner.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
        Me.grdTRSCorner.LookAndFeel.UseDefaultLookAndFeel = False
        Me.grdTRSCorner.MainView = Me.grdTRSCornerView
        Me.grdTRSCorner.Name = "grdTRSCorner"
        Me.grdTRSCorner.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.lupTRSTownship, Me.lupTRSRange, Me.lupTRSSection})
        Me.grdTRSCorner.Size = New System.Drawing.Size(345, 125)
        Me.grdTRSCorner.TabIndex = 0
        Me.grdTRSCorner.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.grdTRSCornerView})
        '
        'TRSCornerBindingSource
        '
        Me.TRSCornerBindingSource.DataSource = GetType(ProspectDataReduction.ViewModels.ProspectAreaTRSCorner)
        '
        'grdTRSCornerView
        '
        Me.grdTRSCornerView.Bands.AddRange(New DevExpress.XtraGrid.Views.BandedGrid.GridBand() {Me.colTRSCornerSW, Me.colTRSCornerNE})
        Me.grdTRSCornerView.Columns.AddRange(New DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn() {Me.colTRSCornerTownship_SW, Me.colTRSCornerRange_SW, Me.colTRSCornerSection_SW, Me.colTRSCornerTownship_NE, Me.colTRSCornerRange_NE, Me.colTRSCornerSection_NE})
        Me.grdTRSCornerView.GridControl = Me.grdTRSCorner
        Me.grdTRSCornerView.IndicatorWidth = 25
        Me.grdTRSCornerView.Name = "grdTRSCornerView"
        Me.grdTRSCornerView.OptionsCustomization.AllowFilter = False
        Me.grdTRSCornerView.OptionsCustomization.AllowGroup = False
        Me.grdTRSCornerView.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom
        Me.grdTRSCornerView.OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never
        Me.grdTRSCornerView.OptionsView.ShowGroupPanel = False
        '
        'colTRSCornerSW
        '
        Me.colTRSCornerSW.AppearanceHeader.Options.UseTextOptions = True
        Me.colTRSCornerSW.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colTRSCornerSW.Caption = "SW"
        Me.colTRSCornerSW.Columns.Add(Me.colTRSCornerTownship_SW)
        Me.colTRSCornerSW.Columns.Add(Me.colTRSCornerRange_SW)
        Me.colTRSCornerSW.Columns.Add(Me.colTRSCornerSection_SW)
        Me.colTRSCornerSW.Name = "colTRSCornerSW"
        Me.colTRSCornerSW.VisibleIndex = 0
        Me.colTRSCornerSW.Width = 160
        '
        'colTRSCornerTownship_SW
        '
        Me.colTRSCornerTownship_SW.AppearanceHeader.Options.UseTextOptions = True
        Me.colTRSCornerTownship_SW.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colTRSCornerTownship_SW.Caption = "Township SW"
        Me.colTRSCornerTownship_SW.ColumnEdit = Me.lupTRSTownship
        Me.colTRSCornerTownship_SW.FieldName = "SW_Township"
        Me.colTRSCornerTownship_SW.Name = "colTRSCornerTownship_SW"
        Me.colTRSCornerTownship_SW.Visible = True
        Me.colTRSCornerTownship_SW.Width = 55
        '
        'lupTRSTownship
        '
        Me.lupTRSTownship.AutoHeight = False
        Me.lupTRSTownship.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.lupTRSTownship.Columns.AddRange(New DevExpress.XtraEditors.Controls.LookUpColumnInfo() {New DevExpress.XtraEditors.Controls.LookUpColumnInfo("Value", "Township")})
        Me.lupTRSTownship.DataSource = Me.TownshipBindingSource
        Me.lupTRSTownship.DisplayMember = "Value"
        Me.lupTRSTownship.Name = "lupTRSTownship"
        Me.lupTRSTownship.NullText = ""
        Me.lupTRSTownship.ShowFooter = False
        Me.lupTRSTownship.ShowHeader = False
        Me.lupTRSTownship.ShowLines = False
        Me.lupTRSTownship.ValueMember = "Value"
        '
        'TownshipBindingSource
        '
        Me.TownshipBindingSource.DataSource = GetType(ProspectDataReduction.ValueListWeb.ValueListDetailDS)
        Me.TownshipBindingSource.Position = 0
        '
        'colTRSCornerRange_NE
        '
        Me.colTRSCornerRange_NE.AppearanceHeader.Options.UseTextOptions = True
        Me.colTRSCornerRange_NE.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colTRSCornerRange_NE.Caption = "Range"
        Me.colTRSCornerRange_NE.ColumnEdit = Me.lupTRSRange
        Me.colTRSCornerRange_NE.FieldName = "NE_Range"
        Me.colTRSCornerRange_NE.Name = "colTRSCornerRange_NE"
        Me.colTRSCornerRange_NE.Visible = True
        Me.colTRSCornerRange_NE.Width = 48
        '
        'lupTRSRange
        '
        Me.lupTRSRange.AutoHeight = False
        Me.lupTRSRange.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.lupTRSRange.Columns.AddRange(New DevExpress.XtraEditors.Controls.LookUpColumnInfo() {New DevExpress.XtraEditors.Controls.LookUpColumnInfo("Value", "Range")})
        Me.lupTRSRange.DataSource = Me.RangeBindingSource
        Me.lupTRSRange.DisplayMember = "Value"
        Me.lupTRSRange.Name = "lupTRSRange"
        Me.lupTRSRange.NullText = ""
        Me.lupTRSRange.ShowFooter = False
        Me.lupTRSRange.ShowHeader = False
        Me.lupTRSRange.ShowLines = False
        Me.lupTRSRange.ValueMember = "Value"
        '
        'RangeBindingSource
        '
        Me.RangeBindingSource.DataSource = GetType(ProspectDataReduction.ValueListWeb.ValueListDetailDS)
        Me.RangeBindingSource.Position = 0
        '
        'colTRSCornerSection_NE
        '
        Me.colTRSCornerSection_NE.AppearanceHeader.Options.UseTextOptions = True
        Me.colTRSCornerSection_NE.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colTRSCornerSection_NE.Caption = "Section"
        Me.colTRSCornerSection_NE.ColumnEdit = Me.lupTRSSection
        Me.colTRSCornerSection_NE.FieldName = "NE_Section"
        Me.colTRSCornerSection_NE.Name = "colTRSCornerSection_NE"
        Me.colTRSCornerSection_NE.Visible = True
        Me.colTRSCornerSection_NE.Width = 54
        '
        'lupTRSSection
        '
        Me.lupTRSSection.AutoHeight = False
        Me.lupTRSSection.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.lupTRSSection.Columns.AddRange(New DevExpress.XtraEditors.Controls.LookUpColumnInfo() {New DevExpress.XtraEditors.Controls.LookUpColumnInfo("Value", "Section")})
        Me.lupTRSSection.DataSource = Me.SectionBindingSource
        Me.lupTRSSection.DisplayMember = "Value"
        Me.lupTRSSection.Name = "lupTRSSection"
        Me.lupTRSSection.NullText = ""
        Me.lupTRSSection.ShowFooter = False
        Me.lupTRSSection.ShowHeader = False
        Me.lupTRSSection.ShowLines = False
        Me.lupTRSSection.ValueMember = "Value"
        '
        'SectionBindingSource
        '
        Me.SectionBindingSource.DataSource = GetType(ProspectDataReduction.ValueListWeb.ValueListDetailDS)
        Me.SectionBindingSource.Position = 0
        '
        'colTRSCornerNE
        '
        Me.colTRSCornerNE.AppearanceHeader.Options.UseTextOptions = True
        Me.colTRSCornerNE.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colTRSCornerNE.Caption = "NE"
        Me.colTRSCornerNE.Columns.Add(Me.colTRSCornerTownship_NE)
        Me.colTRSCornerNE.Columns.Add(Me.colTRSCornerRange_NE)
        Me.colTRSCornerNE.Columns.Add(Me.colTRSCornerSection_NE)
        Me.colTRSCornerNE.Name = "colTRSCornerNE"
        Me.colTRSCornerNE.VisibleIndex = 1
        Me.colTRSCornerNE.Width = 167
        '
        'colTRSCornerTownship_NE
        '
        Me.colTRSCornerTownship_NE.AppearanceHeader.Options.UseTextOptions = True
        Me.colTRSCornerTownship_NE.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colTRSCornerTownship_NE.Caption = "Township"
        Me.colTRSCornerTownship_NE.ColumnEdit = Me.lupTRSTownship
        Me.colTRSCornerTownship_NE.FieldName = "NE_Township"
        Me.colTRSCornerTownship_NE.Name = "colTRSCornerTownship_NE"
        Me.colTRSCornerTownship_NE.Visible = True
        Me.colTRSCornerTownship_NE.Width = 65
        '
        'colTRSCornerRange_SW
        '
        Me.colTRSCornerRange_SW.AppearanceHeader.Options.UseTextOptions = True
        Me.colTRSCornerRange_SW.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colTRSCornerRange_SW.Caption = "Range SW"
        Me.colTRSCornerRange_SW.ColumnEdit = Me.lupTRSRange
        Me.colTRSCornerRange_SW.FieldName = "SW_Range"
        Me.colTRSCornerRange_SW.Name = "colTRSCornerRange_SW"
        Me.colTRSCornerRange_SW.Visible = True
        Me.colTRSCornerRange_SW.Width = 46
        '
        'colTRSCornerSection_SW
        '
        Me.colTRSCornerSection_SW.AppearanceHeader.Options.UseTextOptions = True
        Me.colTRSCornerSection_SW.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colTRSCornerSection_SW.Caption = "Section SW"
        Me.colTRSCornerSection_SW.ColumnEdit = Me.lupTRSSection
        Me.colTRSCornerSection_SW.FieldName = "SW_Section"
        Me.colTRSCornerSection_SW.Name = "colTRSCornerSection_SW"
        Me.colTRSCornerSection_SW.Visible = True
        Me.colTRSCornerSection_SW.Width = 59
        '
        'pnlMineArea
        '
        Me.pnlMineArea.Controls.Add(Me.cboArea)
        Me.pnlMineArea.Controls.Add(Me.lblArea)
        Me.pnlMineArea.Controls.Add(Me.lblMine)
        Me.pnlMineArea.Controls.Add(Me.cboMine)
        Me.pnlMineArea.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMineArea.Location = New System.Drawing.Point(0, 310)
        Me.pnlMineArea.Name = "pnlMineArea"
        Me.pnlMineArea.Size = New System.Drawing.Size(380, 155)
        Me.pnlMineArea.TabIndex = 3
        '
        'cboArea
        '
        Me.cboArea.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.AreaDefinitionBindingSource, "SpecAreaName", True))
        Me.cboArea.DisplayMember = "Name"
        Me.cboArea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboArea.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.cboArea.FormattingEnabled = True
        Me.cboArea.Location = New System.Drawing.Point(70, 43)
        Me.cboArea.Name = "cboArea"
        Me.cboArea.Size = New System.Drawing.Size(228, 21)
        Me.cboArea.TabIndex = 17
        Me.cboArea.ValueMember = "Name"
        '
        'lblArea
        '
        Me.lblArea.AutoSize = True
        Me.lblArea.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblArea.Location = New System.Drawing.Point(22, 42)
        Me.lblArea.Name = "lblArea"
        Me.lblArea.Size = New System.Drawing.Size(29, 13)
        Me.lblArea.TabIndex = 16
        Me.lblArea.Text = "Area"
        '
        'lblMine
        '
        Me.lblMine.AutoSize = True
        Me.lblMine.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMine.Location = New System.Drawing.Point(22, 17)
        Me.lblMine.Name = "lblMine"
        Me.lblMine.Size = New System.Drawing.Size(30, 13)
        Me.lblMine.TabIndex = 15
        Me.lblMine.Text = "Mine"
        '
        'cboMine
        '
        Me.cboMine.DataBindings.Add(New System.Windows.Forms.Binding("SelectedValue", Me.AreaDefinitionBindingSource, "MineName", True))
        Me.cboMine.DisplayMember = "Name"
        Me.cboMine.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMine.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboMine.FormattingEnabled = True
        Me.cboMine.Location = New System.Drawing.Point(70, 15)
        Me.cboMine.Name = "cboMine"
        Me.cboMine.Size = New System.Drawing.Size(228, 21)
        Me.cboMine.TabIndex = 14
        Me.cboMine.ValueMember = "Name"
        '
        'pnlXYCorner
        '
        Me.pnlXYCorner.Controls.Add(Me.btnClearXY)
        Me.pnlXYCorner.Controls.Add(Me.grdXYCorner)
        Me.pnlXYCorner.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlXYCorner.Location = New System.Drawing.Point(0, 155)
        Me.pnlXYCorner.Name = "pnlXYCorner"
        Me.pnlXYCorner.Size = New System.Drawing.Size(380, 155)
        Me.pnlXYCorner.TabIndex = 6
        '
        'btnClearXY
        '
        Me.btnClearXY.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClearXY.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.btnClearXY.Location = New System.Drawing.Point(287, 130)
        Me.btnClearXY.Name = "btnClearXY"
        Me.btnClearXY.Size = New System.Drawing.Size(75, 23)
        Me.btnClearXY.TabIndex = 6
        Me.btnClearXY.Text = "Clear Grid"
        Me.btnClearXY.UseVisualStyleBackColor = True
        '
        'grdXYCorner
        '
        Me.grdXYCorner.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grdXYCorner.DataSource = Me.XYCornerBindingSource
        Me.grdXYCorner.Location = New System.Drawing.Point(17, 3)
        Me.grdXYCorner.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
        Me.grdXYCorner.LookAndFeel.UseDefaultLookAndFeel = False
        Me.grdXYCorner.MainView = Me.grdXYCornerView
        Me.grdXYCorner.Name = "grdXYCorner"
        Me.grdXYCorner.Size = New System.Drawing.Size(345, 125)
        Me.grdXYCorner.TabIndex = 0
        Me.grdXYCorner.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.grdXYCornerView})
        '
        'XYCornerBindingSource
        '
        Me.XYCornerBindingSource.AllowNew = True
        Me.XYCornerBindingSource.DataSource = GetType(ProspectDataReduction.ViewModels.ProspectAreaXYCorner)
        '
        'grdXYCornerView
        '
        Me.grdXYCornerView.Bands.AddRange(New DevExpress.XtraGrid.Views.BandedGrid.GridBand() {Me.colSW, Me.colNE})
        Me.grdXYCornerView.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Flat
        Me.grdXYCornerView.Columns.AddRange(New DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn() {Me.colSW_X, Me.colSW_Y, Me.colNE_X, Me.colNE_Y})
        Me.grdXYCornerView.GridControl = Me.grdXYCorner
        Me.grdXYCornerView.IndicatorWidth = 25
        Me.grdXYCornerView.Name = "grdXYCornerView"
        Me.grdXYCornerView.OptionsCustomization.AllowGroup = False
        Me.grdXYCornerView.OptionsMenu.EnableColumnMenu = False
        Me.grdXYCornerView.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom
        Me.grdXYCornerView.OptionsView.ShowGroupPanel = False
        '
        'colSW
        '
        Me.colSW.AppearanceHeader.Options.UseTextOptions = True
        Me.colSW.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSW.Caption = "SW"
        Me.colSW.Columns.Add(Me.colSW_X)
        Me.colSW.Columns.Add(Me.colSW_Y)
        Me.colSW.Name = "colSW"
        Me.colSW.VisibleIndex = 0
        Me.colSW.Width = 150
        '
        'colSW_X
        '
        Me.colSW_X.AppearanceHeader.Options.UseTextOptions = True
        Me.colSW_X.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSW_X.Caption = "X Coordinate"
        Me.colSW_X.DisplayFormat.FormatString = "#0.00"
        Me.colSW_X.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.colSW_X.FieldName = "SW_XCoordinate"
        Me.colSW_X.Name = "colSW_X"
        Me.colSW_X.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.[False]
        Me.colSW_X.OptionsColumn.AllowMove = False
        Me.colSW_X.OptionsColumn.AllowShowHide = False
        Me.colSW_X.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.[False]
        Me.colSW_X.OptionsFilter.AllowAutoFilter = False
        Me.colSW_X.OptionsFilter.AllowFilter = False
        Me.colSW_X.OptionsFilter.AllowFilterModeChanging = DevExpress.Utils.DefaultBoolean.[False]
        Me.colSW_X.OptionsFilter.FilterBySortField = DevExpress.Utils.DefaultBoolean.[False]
        Me.colSW_X.Visible = True
        '
        'colSW_Y
        '
        Me.colSW_Y.AppearanceHeader.Options.UseTextOptions = True
        Me.colSW_Y.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSW_Y.Caption = "Y Coordinate"
        Me.colSW_Y.DisplayFormat.FormatString = "#0.00"
        Me.colSW_Y.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.colSW_Y.FieldName = "SW_YCoordinate"
        Me.colSW_Y.Name = "colSW_Y"
        Me.colSW_Y.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.[False]
        Me.colSW_Y.OptionsColumn.AllowMove = False
        Me.colSW_Y.OptionsColumn.AllowShowHide = False
        Me.colSW_Y.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.[False]
        Me.colSW_Y.OptionsFilter.AllowAutoFilter = False
        Me.colSW_Y.OptionsFilter.AllowFilter = False
        Me.colSW_Y.OptionsFilter.AllowFilterModeChanging = DevExpress.Utils.DefaultBoolean.[False]
        Me.colSW_Y.OptionsFilter.FilterBySortField = DevExpress.Utils.DefaultBoolean.[False]
        Me.colSW_Y.Visible = True
        '
        'colNE
        '
        Me.colNE.AppearanceHeader.Options.UseTextOptions = True
        Me.colNE.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colNE.Caption = "NE"
        Me.colNE.Columns.Add(Me.colNE_X)
        Me.colNE.Columns.Add(Me.colNE_Y)
        Me.colNE.Name = "colNE"
        Me.colNE.OptionsBand.AllowMove = False
        Me.colNE.VisibleIndex = 1
        Me.colNE.Width = 150
        '
        'colNE_X
        '
        Me.colNE_X.AppearanceHeader.Options.UseTextOptions = True
        Me.colNE_X.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colNE_X.Caption = "X Coordinate"
        Me.colNE_X.DisplayFormat.FormatString = "#0.00"
        Me.colNE_X.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.colNE_X.FieldName = "NE_XCoordinate"
        Me.colNE_X.Name = "colNE_X"
        Me.colNE_X.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.[False]
        Me.colNE_X.OptionsColumn.AllowMove = False
        Me.colNE_X.OptionsColumn.AllowShowHide = False
        Me.colNE_X.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.[False]
        Me.colNE_X.OptionsFilter.AllowAutoFilter = False
        Me.colNE_X.OptionsFilter.AllowFilter = False
        Me.colNE_X.OptionsFilter.AllowFilterModeChanging = DevExpress.Utils.DefaultBoolean.[False]
        Me.colNE_X.OptionsFilter.FilterBySortField = DevExpress.Utils.DefaultBoolean.[False]
        Me.colNE_X.Visible = True
        '
        'colNE_Y
        '
        Me.colNE_Y.AppearanceHeader.Options.UseTextOptions = True
        Me.colNE_Y.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colNE_Y.Caption = "Y Coordinate"
        Me.colNE_Y.DisplayFormat.FormatString = "#0.00"
        Me.colNE_Y.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
        Me.colNE_Y.FieldName = "NE_YCoordinate"
        Me.colNE_Y.Name = "colNE_Y"
        Me.colNE_Y.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.[False]
        Me.colNE_Y.OptionsColumn.AllowMove = False
        Me.colNE_Y.OptionsColumn.AllowShowHide = False
        Me.colNE_Y.OptionsColumn.AllowSort = DevExpress.Utils.DefaultBoolean.[False]
        Me.colNE_Y.OptionsFilter.AllowAutoFilter = False
        Me.colNE_Y.OptionsFilter.AllowFilter = False
        Me.colNE_Y.OptionsFilter.AllowFilterModeChanging = DevExpress.Utils.DefaultBoolean.[False]
        Me.colNE_Y.OptionsFilter.FilterBySortField = DevExpress.Utils.DefaultBoolean.[False]
        Me.colNE_Y.Visible = True
        '
        'pnlHoles
        '
        Me.pnlHoles.Controls.Add(Me.btnClearHoles)
        Me.pnlHoles.Controls.Add(Me.btnAddHolesFromFile)
        Me.pnlHoles.Controls.Add(Me.btnAddHolesFromProspectGrid)
        Me.pnlHoles.Controls.Add(Me.btnBrowse)
        Me.pnlHoles.Controls.Add(Me.Label1)
        Me.pnlHoles.Controls.Add(Me.txtFileName)
        Me.pnlHoles.Controls.Add(Me.grdHolesFromFile)
        Me.pnlHoles.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlHoles.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.pnlHoles.Location = New System.Drawing.Point(0, 0)
        Me.pnlHoles.Name = "pnlHoles"
        Me.pnlHoles.Size = New System.Drawing.Size(380, 155)
        Me.pnlHoles.TabIndex = 7
        '
        'btnClearHoles
        '
        Me.btnClearHoles.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClearHoles.Font = New System.Drawing.Font("Tahoma", 8.25!)
        Me.btnClearHoles.Location = New System.Drawing.Point(287, 128)
        Me.btnClearHoles.Name = "btnClearHoles"
        Me.btnClearHoles.Size = New System.Drawing.Size(75, 23)
        Me.btnClearHoles.TabIndex = 9
        Me.btnClearHoles.Text = "Clear Grid"
        Me.btnClearHoles.UseVisualStyleBackColor = True
        '
        'btnAddHolesFromFile
        '
        Me.btnAddHolesFromFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAddHolesFromFile.Location = New System.Drawing.Point(173, 32)
        Me.btnAddHolesFromFile.Name = "btnAddHolesFromFile"
        Me.btnAddHolesFromFile.Size = New System.Drawing.Size(193, 23)
        Me.btnAddHolesFromFile.TabIndex = 8
        Me.btnAddHolesFromFile.Text = "Add Holes from File (specified above)"
        Me.btnAddHolesFromFile.UseVisualStyleBackColor = True
        '
        'btnAddHolesFromProspectGrid
        '
        Me.btnAddHolesFromProspectGrid.Location = New System.Drawing.Point(15, 32)
        Me.btnAddHolesFromProspectGrid.Name = "btnAddHolesFromProspectGrid"
        Me.btnAddHolesFromProspectGrid.Size = New System.Drawing.Size(155, 23)
        Me.btnAddHolesFromProspectGrid.TabIndex = 7
        Me.btnAddHolesFromProspectGrid.Text = "Add Holes from Prospect Grid"
        Me.btnAddHolesFromProspectGrid.UseVisualStyleBackColor = True
        '
        'btnBrowse
        '
        Me.btnBrowse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowse.Location = New System.Drawing.Point(312, 7)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(54, 23)
        Me.btnBrowse.TabIndex = 6
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(18, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "File name"
        '
        'txtFileName
        '
        Me.txtFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFileName.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.AreaDefinitionBindingSource, "FileName", True, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged))
        Me.txtFileName.Location = New System.Drawing.Point(70, 8)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(236, 21)
        Me.txtFileName.TabIndex = 4
        '
        'grdHolesFromFile
        '
        Me.grdHolesFromFile.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grdHolesFromFile.DataSource = Me.HoleBindingSource
        Me.grdHolesFromFile.Location = New System.Drawing.Point(18, 58)
        Me.grdHolesFromFile.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
        Me.grdHolesFromFile.LookAndFeel.UseDefaultLookAndFeel = False
        Me.grdHolesFromFile.MainView = Me.grdHolesFromFileView
        Me.grdHolesFromFile.Name = "grdHolesFromFile"
        Me.grdHolesFromFile.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.lupHoleTownship, Me.lupHoleRange, Me.lupHoleSection, Me.lupHoleHole})
        Me.grdHolesFromFile.Size = New System.Drawing.Size(345, 68)
        Me.grdHolesFromFile.TabIndex = 2
        Me.grdHolesFromFile.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.grdHolesFromFileView})
        '
        'HoleBindingSource
        '
        Me.HoleBindingSource.AllowNew = True
        Me.HoleBindingSource.DataSource = GetType(ProspectDataReduction.ViewModels.ProspectAreaHole)
        '
        'grdHolesFromFileView
        '
        Me.grdHolesFromFileView.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.colTownship, Me.colRange, Me.colSection, Me.colHole})
        Me.grdHolesFromFileView.GridControl = Me.grdHolesFromFile
        Me.grdHolesFromFileView.IndicatorWidth = 30
        Me.grdHolesFromFileView.Name = "grdHolesFromFileView"
        Me.grdHolesFromFileView.OptionsCustomization.AllowFilter = False
        Me.grdHolesFromFileView.OptionsMenu.EnableColumnMenu = False
        Me.grdHolesFromFileView.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom
        Me.grdHolesFromFileView.OptionsView.ShowGroupPanel = False
        '
        'colTownship
        '
        Me.colTownship.Caption = "Township"
        Me.colTownship.ColumnEdit = Me.lupHoleTownship
        Me.colTownship.FieldName = "Hole_Township"
        Me.colTownship.Name = "colTownship"
        Me.colTownship.Visible = True
        Me.colTownship.VisibleIndex = 0
        '
        'lupHoleTownship
        '
        Me.lupHoleTownship.AutoHeight = False
        Me.lupHoleTownship.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.lupHoleTownship.Columns.AddRange(New DevExpress.XtraEditors.Controls.LookUpColumnInfo() {New DevExpress.XtraEditors.Controls.LookUpColumnInfo("Value", "Township")})
        Me.lupHoleTownship.DataSource = Me.TownshipBindingSource
        Me.lupHoleTownship.DisplayMember = "Value"
        Me.lupHoleTownship.Name = "lupHoleTownship"
        Me.lupHoleTownship.NullText = ""
        Me.lupHoleTownship.ShowFooter = False
        Me.lupHoleTownship.ShowHeader = False
        Me.lupHoleTownship.ShowLines = False
        Me.lupHoleTownship.ValueMember = "Value"
        '
        'colRange
        '
        Me.colRange.Caption = "Range"
        Me.colRange.ColumnEdit = Me.lupHoleRange
        Me.colRange.FieldName = "Hole_Range"
        Me.colRange.Name = "colRange"
        Me.colRange.Visible = True
        Me.colRange.VisibleIndex = 1
        '
        'lupHoleRange
        '
        Me.lupHoleRange.AutoHeight = False
        Me.lupHoleRange.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.lupHoleRange.Columns.AddRange(New DevExpress.XtraEditors.Controls.LookUpColumnInfo() {New DevExpress.XtraEditors.Controls.LookUpColumnInfo("Value", "Range")})
        Me.lupHoleRange.DataSource = Me.RangeBindingSource
        Me.lupHoleRange.DisplayMember = "Value"
        Me.lupHoleRange.Name = "lupHoleRange"
        Me.lupHoleRange.NullText = ""
        Me.lupHoleRange.ShowFooter = False
        Me.lupHoleRange.ShowHeader = False
        Me.lupHoleRange.ShowLines = False
        Me.lupHoleRange.ValueMember = "Value"
        '
        'colSection
        '
        Me.colSection.Caption = "Section"
        Me.colSection.ColumnEdit = Me.lupHoleSection
        Me.colSection.FieldName = "Hole_Section"
        Me.colSection.Name = "colSection"
        Me.colSection.Visible = True
        Me.colSection.VisibleIndex = 2
        '
        'lupHoleSection
        '
        Me.lupHoleSection.AutoHeight = False
        Me.lupHoleSection.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.lupHoleSection.Columns.AddRange(New DevExpress.XtraEditors.Controls.LookUpColumnInfo() {New DevExpress.XtraEditors.Controls.LookUpColumnInfo("Value", "Section")})
        Me.lupHoleSection.DataSource = Me.SectionBindingSource
        Me.lupHoleSection.DisplayMember = "Value"
        Me.lupHoleSection.Name = "lupHoleSection"
        Me.lupHoleSection.NullText = ""
        Me.lupHoleSection.ShowFooter = False
        Me.lupHoleSection.ShowHeader = False
        Me.lupHoleSection.ShowLines = False
        Me.lupHoleSection.ValueMember = "Value"
        '
        'colHole
        '
        Me.colHole.Caption = "Hole"
        Me.colHole.ColumnEdit = Me.lupHoleHole
        Me.colHole.FieldName = "Hole_Location"
        Me.colHole.Name = "colHole"
        Me.colHole.Visible = True
        Me.colHole.VisibleIndex = 3
        '
        'lupHoleHole
        '
        Me.lupHoleHole.AutoHeight = False
        Me.lupHoleHole.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.lupHoleHole.Columns.AddRange(New DevExpress.XtraEditors.Controls.LookUpColumnInfo() {New DevExpress.XtraEditors.Controls.LookUpColumnInfo("Value", "Hole")})
        Me.lupHoleHole.DataSource = Me.HolesListBindingSource
        Me.lupHoleHole.DisplayMember = "Value"
        Me.lupHoleHole.Name = "lupHoleHole"
        Me.lupHoleHole.NullText = ""
        Me.lupHoleHole.ShowFooter = False
        Me.lupHoleHole.ShowHeader = False
        Me.lupHoleHole.ShowLines = False
        Me.lupHoleHole.ValueMember = "Value"
        '
        'HolesListBindingSource
        '
        Me.HolesListBindingSource.DataSource = GetType(ProspectDataReduction.ValueListWeb.ValueListDetailDS)
        Me.HolesListBindingSource.Position = 0
        '
        'tspArea
        '
        Me.tspArea.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnAddNew, Me.ToolStripSeparator1, Me.btnSave, Me.btnSaveAs, Me.separator3, Me.btnDelete, Me.separator2, Me.btnCancel})
        Me.tspArea.Location = New System.Drawing.Point(2, 20)
        Me.tspArea.Name = "tspArea"
        Me.tspArea.Size = New System.Drawing.Size(647, 25)
        Me.tspArea.TabIndex = 24
        Me.tspArea.Text = "ToolStrip1"
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
        'grpAreaDefinitions
        '
        Me.grpAreaDefinitions.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpAreaDefinitions.AppearanceCaption.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
        Me.grpAreaDefinitions.AppearanceCaption.Options.UseFont = True
        Me.grpAreaDefinitions.Controls.Add(Me.btnGetAreaDefinitions)
        Me.grpAreaDefinitions.Controls.Add(Me.chkMyDefinitionsOnly)
        Me.grpAreaDefinitions.Controls.Add(Me.grdAreaDefinition)
        Me.grpAreaDefinitions.Location = New System.Drawing.Point(12, 12)
        Me.grpAreaDefinitions.Name = "grpAreaDefinitions"
        Me.grpAreaDefinitions.Size = New System.Drawing.Size(662, 540)
        Me.grpAreaDefinitions.TabIndex = 5
        Me.grpAreaDefinitions.Text = "Area Definitions"
        '
        'btnGetAreaDefinitions
        '
        Me.btnGetAreaDefinitions.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGetAreaDefinitions.Location = New System.Drawing.Point(508, 23)
        Me.btnGetAreaDefinitions.Name = "btnGetAreaDefinitions"
        Me.btnGetAreaDefinitions.Size = New System.Drawing.Size(136, 23)
        Me.btnGetAreaDefinitions.TabIndex = 2
        Me.btnGetAreaDefinitions.Text = "Get Area Definitions"
        Me.btnGetAreaDefinitions.UseVisualStyleBackColor = True
        '
        'chkMyDefinitionsOnly
        '
        Me.chkMyDefinitionsOnly.AutoSize = True
        Me.chkMyDefinitionsOnly.Checked = True
        Me.chkMyDefinitionsOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMyDefinitionsOnly.Location = New System.Drawing.Point(17, 28)
        Me.chkMyDefinitionsOnly.Name = "chkMyDefinitionsOnly"
        Me.chkMyDefinitionsOnly.Size = New System.Drawing.Size(118, 17)
        Me.chkMyDefinitionsOnly.TabIndex = 1
        Me.chkMyDefinitionsOnly.Text = "My Definitions Only"
        Me.chkMyDefinitionsOnly.UseVisualStyleBackColor = True
        '
        'grdAreaDefinition
        '
        Me.grdAreaDefinition.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grdAreaDefinition.Location = New System.Drawing.Point(17, 51)
        Me.grdAreaDefinition.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat
        Me.grdAreaDefinition.LookAndFeel.UseDefaultLookAndFeel = False
        Me.grdAreaDefinition.MainView = Me.grdAreaDefinitionView
        Me.grdAreaDefinition.Name = "grdAreaDefinition"
        Me.grdAreaDefinition.Size = New System.Drawing.Size(627, 471)
        Me.grdAreaDefinition.TabIndex = 0
        Me.grdAreaDefinition.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.grdAreaDefinitionView})
        '
        'grdAreaDefinitionView
        '
        Me.grdAreaDefinitionView.Appearance.EvenRow.BackColor = System.Drawing.Color.AliceBlue
        Me.grdAreaDefinitionView.Appearance.EvenRow.Options.UseBackColor = True
        Me.grdAreaDefinitionView.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.colName, Me.colWho, Me.colWhen, Me.colType, Me.colMine, Me.colSubArea})
        Me.grdAreaDefinitionView.GridControl = Me.grdAreaDefinition
        Me.grdAreaDefinitionView.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always
        Me.grdAreaDefinitionView.IndicatorWidth = 30
        Me.grdAreaDefinitionView.Name = "grdAreaDefinitionView"
        Me.grdAreaDefinitionView.OptionsBehavior.Editable = False
        Me.grdAreaDefinitionView.OptionsCustomization.AllowFilter = False
        Me.grdAreaDefinitionView.OptionsCustomization.AllowGroup = False
        Me.grdAreaDefinitionView.OptionsView.EnableAppearanceEvenRow = True
        Me.grdAreaDefinitionView.OptionsView.ShowGroupPanel = False
        '
        'colName
        '
        Me.colName.Caption = "Name"
        Me.colName.FieldName = "AreaDefinitionName"
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
        'colType
        '
        Me.colType.Caption = "Mode"
        Me.colType.FieldName = "AreaMethod"
        Me.colType.Name = "colType"
        Me.colType.Visible = True
        Me.colType.VisibleIndex = 3
        Me.colType.Width = 88
        '
        'colMine
        '
        Me.colMine.Caption = "Mine"
        Me.colMine.FieldName = "MineName"
        Me.colMine.Name = "colMine"
        Me.colMine.Visible = True
        Me.colMine.VisibleIndex = 4
        Me.colMine.Width = 99
        '
        'colSubArea
        '
        Me.colSubArea.Caption = "Sub Area"
        Me.colSubArea.FieldName = "SpecAreaName"
        Me.colSubArea.Name = "colSubArea"
        Me.colSubArea.Visible = True
        Me.colSubArea.VisibleIndex = 5
        Me.colSubArea.Width = 74
        '
        'mnuAddFrom
        '
        Me.mnuAddFrom.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.btnFromProspectGrid), New DevExpress.XtraBars.LinkPersistInfo(Me.btnFromFile)})
        Me.mnuAddFrom.Manager = Me.barAddFrom
        Me.mnuAddFrom.Name = "mnuAddFrom"
        '
        'btnFromProspectGrid
        '
        Me.btnFromProspectGrid.Caption = "Prospect Grid"
        Me.btnFromProspectGrid.Id = 7
        Me.btnFromProspectGrid.Name = "btnFromProspectGrid"
        '
        'btnFromFile
        '
        Me.btnFromFile.Caption = "File"
        Me.btnFromFile.Id = 8
        Me.btnFromFile.Name = "btnFromFile"
        '
        'barAddFrom
        '
        Me.barAddFrom.DockControls.Add(Me.barDockControlTop)
        Me.barAddFrom.DockControls.Add(Me.barDockControlBottom)
        Me.barAddFrom.DockControls.Add(Me.barDockControlLeft)
        Me.barAddFrom.DockControls.Add(Me.barDockControlRight)
        Me.barAddFrom.Form = Me
        Me.barAddFrom.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.btnFromProspectGrid, Me.btnFromFile})
        Me.barAddFrom.MaxItemId = 9
        '
        'barDockControlTop
        '
        Me.barDockControlTop.CausesValidation = False
        Me.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.barDockControlTop.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlTop.Size = New System.Drawing.Size(1360, 0)
        '
        'barDockControlBottom
        '
        Me.barDockControlBottom.CausesValidation = False
        Me.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.barDockControlBottom.Location = New System.Drawing.Point(0, 564)
        Me.barDockControlBottom.Size = New System.Drawing.Size(1360, 0)
        '
        'barDockControlLeft
        '
        Me.barDockControlLeft.CausesValidation = False
        Me.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.barDockControlLeft.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlLeft.Size = New System.Drawing.Size(0, 564)
        '
        'barDockControlRight
        '
        Me.barDockControlRight.CausesValidation = False
        Me.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.barDockControlRight.Location = New System.Drawing.Point(1360, 0)
        Me.barDockControlRight.Size = New System.Drawing.Size(0, 564)
        '
        'ProspectCodeBindingSource
        '
        Me.ProspectCodeBindingSource.DataSource = GetType(ProspectDataReduction.RawService.ProspectCode)
        '
        'ErrorProvider
        '
        Me.ErrorProvider.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink
        Me.ErrorProvider.ContainerControl = Me
        Me.ErrorProvider.DataSource = Me.AreaDefinitionBindingSource
        '
        'ctrAreaDefinition
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.grpAreaDefinitionDetail)
        Me.Controls.Add(Me.grpAreaDefinitions)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.Name = "ctrAreaDefinition"
        Me.Size = New System.Drawing.Size(1360, 564)
        CType(Me.grpAreaDefinitionDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpAreaDefinitionDetail.ResumeLayout(False)
        Me.grpAreaDefinitionDetail.PerformLayout()
        Me.pnlDetails.ResumeLayout(False)
        Me.pnlDetails.PerformLayout()
        CType(Me.AreaDefinitionBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpHoleFilters, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpHoleFilters.ResumeLayout(False)
        Me.grpHoleFilters.PerformLayout()
        CType(Me.cboOwnership.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridLookUpEdit1View, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtpAreaEndDrillDate2.Properties.CalendarTimeProperties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtpAreaEndDrillDate2.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtpAreaBeginDrillDate2.Properties.CalendarTimeProperties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtpAreaBeginDrillDate2.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpArea, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpArea.ResumeLayout(False)
        Me.grpSelectionMode.ResumeLayout(False)
        Me.grpSelectionMode.PerformLayout()
        Me.grpSelectHoles.ResumeLayout(False)
        Me.pnlSelectHoles.ResumeLayout(False)
        Me.pnlTRSCorner.ResumeLayout(False)
        CType(Me.grdTRSCorner, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TRSCornerBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdTRSCornerView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lupTRSTownship, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TownshipBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lupTRSRange, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RangeBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lupTRSSection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SectionBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlMineArea.ResumeLayout(False)
        Me.pnlMineArea.PerformLayout()
        Me.pnlXYCorner.ResumeLayout(False)
        CType(Me.grdXYCorner, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.XYCornerBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdXYCornerView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlHoles.ResumeLayout(False)
        Me.pnlHoles.PerformLayout()
        CType(Me.grdHolesFromFile, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.HoleBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdHolesFromFileView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lupHoleTownship, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lupHoleRange, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lupHoleSection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lupHoleHole, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.HolesListBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tspArea.ResumeLayout(False)
        Me.tspArea.PerformLayout()
        CType(Me.grpAreaDefinitions, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpAreaDefinitions.ResumeLayout(False)
        Me.grpAreaDefinitions.PerformLayout()
        CType(Me.grdAreaDefinition, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdAreaDefinitionView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mnuAddFrom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.barAddFrom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ProspectCodeBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ErrorProvider, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents grpAreaDefinitionDetail As DevExpress.XtraEditors.GroupControl
    Friend WithEvents pnlDetails As Panel
    Friend WithEvents txtDefinedOn As TextBox
    Friend WithEvents lblDefinedOn As Label
    Friend WithEvents txtDefinedBy As TextBox
    Friend WithEvents lblDefinedBy As Label
    Friend WithEvents txtName As TextBox
    Friend WithEvents lblName As Label
    Friend WithEvents cboAreaName As ComboBox
    Friend WithEvents lblAreaName As Label
    Friend WithEvents lblMineName As Label
    Friend WithEvents cboMineName As ComboBox
    Friend WithEvents grpHoleFilters As DevExpress.XtraEditors.GroupControl
    Friend WithEvents cboOriginalForRedrill As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents cboOwnership As DevExpress.XtraEditors.GridLookUpEdit
    Friend WithEvents GridLookUpEdit1View As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents colOwnershipCode As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colOwnershipDesc As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents dtpAreaEndDrillDate2 As DevExpress.XtraEditors.DateEdit
    Friend WithEvents lblGen2 As Label
    Friend WithEvents dtpAreaBeginDrillDate2 As DevExpress.XtraEditors.DateEdit
    Friend WithEvents lblOwnership As Label
    Friend WithEvents lblGen4 As Label
    Friend WithEvents lblGen32 As Label
    Friend WithEvents lblGen1 As Label
    Friend WithEvents lblGen0 As Label
    Friend WithEvents cboAreaProspHoleType As ComboBox
    Friend WithEvents cboAreaMinedOutStatus As ComboBox
    Friend WithEvents grpArea As DevExpress.XtraEditors.GroupControl
    Friend WithEvents tspArea As ToolStrip
    Friend WithEvents btnAddNew As ToolStripButton
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents btnSave As ToolStripButton
    Friend WithEvents btnSaveAs As ToolStripButton
    Friend WithEvents separator3 As ToolStripSeparator
    Friend WithEvents btnDelete As ToolStripButton
    Friend WithEvents separator2 As ToolStripSeparator
    Friend WithEvents btnCancel As ToolStripButton
    Friend WithEvents grpAreaDefinitions As DevExpress.XtraEditors.GroupControl
    Friend WithEvents btnGetAreaDefinitions As Button
    Friend WithEvents chkMyDefinitionsOnly As CheckBox
    Friend WithEvents grdAreaDefinition As DevExpress.XtraGrid.GridControl
    Friend WithEvents grdAreaDefinitionView As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents colName As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colWho As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colWhen As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colType As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colMine As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colSubArea As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents TRSCornerBindingSource As BindingSource
    Friend WithEvents TownshipBindingSource As BindingSource
    Friend WithEvents RangeBindingSource As BindingSource
    Friend WithEvents SectionBindingSource As BindingSource
    Friend WithEvents AreaDefinitionBindingSource As BindingSource
    Friend WithEvents HoleBindingSource As BindingSource
    Friend WithEvents HolesListBindingSource As BindingSource
    Friend WithEvents XYCornerBindingSource As BindingSource
    Friend WithEvents mnuAddFrom As DevExpress.XtraBars.PopupMenu
    Friend WithEvents barAddFrom As DevExpress.XtraBars.BarManager
    Friend WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Friend WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Friend WithEvents ProspectCodeBindingSource As BindingSource
    Friend WithEvents ErrorProvider As ErrorProvider
    Friend WithEvents btnFromProspectGrid As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents btnFromFile As DevExpress.XtraBars.BarButtonItem
    Friend WithEvents grpSelectHoles As GroupBox
    Friend WithEvents pnlSelectHoles As Panel
    Friend WithEvents pnlTRSCorner As Panel
    Friend WithEvents btnClearTRS As Button
    Friend WithEvents grdTRSCorner As DevExpress.XtraGrid.GridControl
    Friend WithEvents grdTRSCornerView As DevExpress.XtraGrid.Views.BandedGrid.BandedGridView
    Friend WithEvents colTRSCornerSW As DevExpress.XtraGrid.Views.BandedGrid.GridBand
    Friend WithEvents colTRSCornerTownship_SW As DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn
    Friend WithEvents lupTRSTownship As DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit
    Friend WithEvents colTRSCornerRange_NE As DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn
    Friend WithEvents lupTRSRange As DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit
    Friend WithEvents colTRSCornerSection_NE As DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn
    Friend WithEvents lupTRSSection As DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit
    Friend WithEvents colTRSCornerNE As DevExpress.XtraGrid.Views.BandedGrid.GridBand
    Friend WithEvents colTRSCornerTownship_NE As DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn
    Friend WithEvents colTRSCornerRange_SW As DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn
    Friend WithEvents colTRSCornerSection_SW As DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn
    Friend WithEvents pnlMineArea As Panel
    Friend WithEvents cboArea As ComboBox
    Friend WithEvents lblArea As Label
    Friend WithEvents lblMine As Label
    Friend WithEvents cboMine As ComboBox
    Friend WithEvents pnlXYCorner As Panel
    Friend WithEvents btnClearXY As Button
    Friend WithEvents grdXYCorner As DevExpress.XtraGrid.GridControl
    Friend WithEvents grdXYCornerView As DevExpress.XtraGrid.Views.BandedGrid.BandedGridView
    Friend WithEvents colSW As DevExpress.XtraGrid.Views.BandedGrid.GridBand
    Friend WithEvents colSW_X As DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn
    Friend WithEvents colSW_Y As DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn
    Friend WithEvents colNE As DevExpress.XtraGrid.Views.BandedGrid.GridBand
    Friend WithEvents colNE_X As DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn
    Friend WithEvents colNE_Y As DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn
    Friend WithEvents pnlHoles As Panel
    Friend WithEvents btnClearHoles As Button
    Friend WithEvents btnAddHolesFromFile As Button
    Friend WithEvents btnAddHolesFromProspectGrid As Button
    Friend WithEvents btnBrowse As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents txtFileName As TextBox
    Friend WithEvents grdHolesFromFile As DevExpress.XtraGrid.GridControl
    Friend WithEvents grdHolesFromFileView As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents colTownship As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents lupHoleTownship As DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit
    Friend WithEvents colRange As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents lupHoleRange As DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit
    Friend WithEvents colSection As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents lupHoleSection As DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit
    Friend WithEvents colHole As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents lupHoleHole As DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit
    Friend WithEvents grpSelectionMode As GroupBox
    Friend WithEvents rdoFromFile As RadioButton
    Friend WithEvents rdoHoles As RadioButton
    Friend WithEvents rdoMineArea As RadioButton
    Friend WithEvents rdoXYCorner As RadioButton
    Friend WithEvents rdoTRSCorner As RadioButton
End Class
