<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProspectGridHoleSelection
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
        Me.col02 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.grdProspect = New DevExpress.XtraGrid.GridControl()
        Me.grdProspectView = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.col04 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col06 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col08 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col10 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col12 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col14 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col16 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col18 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col20 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col22 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col24 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col26 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col28 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col30 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.col32 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colHoleSuffix = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.lblCurrLoc = New System.Windows.Forms.Label()
        Me.btnEast = New System.Windows.Forms.Button()
        Me.btnSouth = New System.Windows.Forms.Button()
        Me.btnNorth = New System.Windows.Forms.Button()
        Me.btnWest = New System.Windows.Forms.Button()
        Me.btnGo = New System.Windows.Forms.Button()
        Me.cboSec = New System.Windows.Forms.ComboBox()
        Me.lblSec = New System.Windows.Forms.Label()
        Me.cboRge = New System.Windows.Forms.ComboBox()
        Me.cboTwp = New System.Windows.Forms.ComboBox()
        Me.lblRge = New System.Windows.Forms.Label()
        Me.lblTwp = New System.Windows.Forms.Label()
        Me.grdSelectedHoles = New DevExpress.XtraGrid.GridControl()
        Me.grdSelectedHolesView = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.colSelTwp = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colSelRng = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colSelSec = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colSelHole = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.colRemove = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.RepositoryItemCheckEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit()
        Me.grdNonBasicHole = New DevExpress.XtraGrid.GridControl()
        Me.grdNonBasicHoleView = New DevExpress.XtraGrid.Views.Grid.GridView()
        Me.GridColumn1 = New DevExpress.XtraGrid.Columns.GridColumn()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        CType(Me.grdProspect, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdProspectView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdSelectedHoles, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdSelectedHolesView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RepositoryItemCheckEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdNonBasicHole, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdNonBasicHoleView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'col02
        '
        Me.col02.AppearanceHeader.Options.UseTextOptions = True
        Me.col02.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col02.Caption = "02"
        Me.col02.FieldName = "Hole02"
        Me.col02.MaxWidth = 25
        Me.col02.MinWidth = 25
        Me.col02.Name = "col02"
        Me.col02.OptionsColumn.AllowSize = False
        Me.col02.Visible = True
        Me.col02.VisibleIndex = 0
        Me.col02.Width = 25
        '
        'grdProspect
        '
        Me.grdProspect.Location = New System.Drawing.Point(12, 66)
        Me.grdProspect.MainView = Me.grdProspectView
        Me.grdProspect.Name = "grdProspect"
        Me.grdProspect.Size = New System.Drawing.Size(452, 379)
        Me.grdProspect.TabIndex = 0
        Me.grdProspect.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.grdProspectView})
        '
        'grdProspectView
        '
        Me.grdProspectView.Appearance.FixedLine.BackColor = System.Drawing.SystemColors.ControlDark
        Me.grdProspectView.Appearance.FixedLine.Options.UseBackColor = True
        Me.grdProspectView.Appearance.HeaderPanel.BackColor = System.Drawing.SystemColors.Control
        Me.grdProspectView.Appearance.HeaderPanel.Options.UseBackColor = True
        Me.grdProspectView.Appearance.HorzLine.BackColor = System.Drawing.SystemColors.ControlDark
        Me.grdProspectView.Appearance.HorzLine.Options.UseBackColor = True
        Me.grdProspectView.Appearance.VertLine.BackColor = System.Drawing.SystemColors.ControlDark
        Me.grdProspectView.Appearance.VertLine.Options.UseBackColor = True
        Me.grdProspectView.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.col02, Me.col04, Me.col06, Me.col08, Me.col10, Me.col12, Me.col14, Me.col16, Me.col18, Me.col20, Me.col22, Me.col24, Me.col26, Me.col28, Me.col30, Me.col32, Me.colHoleSuffix})
        Me.grdProspectView.GridControl = Me.grdProspect
        Me.grdProspectView.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Never
        Me.grdProspectView.IndicatorWidth = 25
        Me.grdProspectView.Name = "grdProspectView"
        Me.grdProspectView.OptionsBehavior.Editable = False
        Me.grdProspectView.OptionsCustomization.AllowFilter = False
        Me.grdProspectView.OptionsCustomization.AllowSort = False
        Me.grdProspectView.OptionsSelection.UseIndicatorForSelection = False
        Me.grdProspectView.OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never
        Me.grdProspectView.OptionsView.ShowGroupPanel = False
        Me.grdProspectView.VertScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Never
        '
        'col04
        '
        Me.col04.AppearanceHeader.Options.UseTextOptions = True
        Me.col04.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col04.Caption = "04"
        Me.col04.FieldName = "Hole04"
        Me.col04.MaxWidth = 25
        Me.col04.MinWidth = 25
        Me.col04.Name = "col04"
        Me.col04.Visible = True
        Me.col04.VisibleIndex = 1
        Me.col04.Width = 25
        '
        'col06
        '
        Me.col06.AppearanceHeader.Options.UseTextOptions = True
        Me.col06.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col06.Caption = "06"
        Me.col06.FieldName = "Hole06"
        Me.col06.MaxWidth = 25
        Me.col06.MinWidth = 25
        Me.col06.Name = "col06"
        Me.col06.Visible = True
        Me.col06.VisibleIndex = 2
        Me.col06.Width = 25
        '
        'col08
        '
        Me.col08.AppearanceHeader.Options.UseTextOptions = True
        Me.col08.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col08.Caption = "08"
        Me.col08.FieldName = "Hole08"
        Me.col08.MaxWidth = 25
        Me.col08.MinWidth = 25
        Me.col08.Name = "col08"
        Me.col08.Visible = True
        Me.col08.VisibleIndex = 3
        Me.col08.Width = 25
        '
        'col10
        '
        Me.col10.AppearanceHeader.Options.UseTextOptions = True
        Me.col10.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col10.Caption = "10"
        Me.col10.FieldName = "Hole10"
        Me.col10.MaxWidth = 25
        Me.col10.MinWidth = 25
        Me.col10.Name = "col10"
        Me.col10.Visible = True
        Me.col10.VisibleIndex = 4
        Me.col10.Width = 25
        '
        'col12
        '
        Me.col12.AppearanceHeader.Options.UseTextOptions = True
        Me.col12.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col12.Caption = "12"
        Me.col12.FieldName = "Hole12"
        Me.col12.MaxWidth = 25
        Me.col12.MinWidth = 25
        Me.col12.Name = "col12"
        Me.col12.Visible = True
        Me.col12.VisibleIndex = 5
        Me.col12.Width = 25
        '
        'col14
        '
        Me.col14.AppearanceHeader.Options.UseTextOptions = True
        Me.col14.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col14.Caption = "14"
        Me.col14.FieldName = "Hole14"
        Me.col14.MaxWidth = 25
        Me.col14.MinWidth = 25
        Me.col14.Name = "col14"
        Me.col14.Visible = True
        Me.col14.VisibleIndex = 6
        Me.col14.Width = 25
        '
        'col16
        '
        Me.col16.AppearanceHeader.Options.UseTextOptions = True
        Me.col16.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col16.Caption = "16"
        Me.col16.FieldName = "Hole16"
        Me.col16.MaxWidth = 25
        Me.col16.MinWidth = 25
        Me.col16.Name = "col16"
        Me.col16.Visible = True
        Me.col16.VisibleIndex = 7
        Me.col16.Width = 25
        '
        'col18
        '
        Me.col18.AppearanceHeader.Options.UseTextOptions = True
        Me.col18.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col18.Caption = "18"
        Me.col18.FieldName = "Hole18"
        Me.col18.MaxWidth = 25
        Me.col18.MinWidth = 25
        Me.col18.Name = "col18"
        Me.col18.Visible = True
        Me.col18.VisibleIndex = 8
        Me.col18.Width = 25
        '
        'col20
        '
        Me.col20.AppearanceHeader.Options.UseTextOptions = True
        Me.col20.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col20.Caption = "20"
        Me.col20.FieldName = "Hole20"
        Me.col20.MaxWidth = 25
        Me.col20.MinWidth = 25
        Me.col20.Name = "col20"
        Me.col20.Visible = True
        Me.col20.VisibleIndex = 9
        Me.col20.Width = 25
        '
        'col22
        '
        Me.col22.AppearanceHeader.Options.UseTextOptions = True
        Me.col22.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col22.Caption = "22"
        Me.col22.FieldName = "Hole22"
        Me.col22.MaxWidth = 25
        Me.col22.MinWidth = 25
        Me.col22.Name = "col22"
        Me.col22.Visible = True
        Me.col22.VisibleIndex = 10
        Me.col22.Width = 25
        '
        'col24
        '
        Me.col24.AppearanceHeader.Options.UseTextOptions = True
        Me.col24.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col24.Caption = "24"
        Me.col24.FieldName = "Hole24"
        Me.col24.MaxWidth = 25
        Me.col24.MinWidth = 25
        Me.col24.Name = "col24"
        Me.col24.Visible = True
        Me.col24.VisibleIndex = 11
        Me.col24.Width = 25
        '
        'col26
        '
        Me.col26.AppearanceHeader.Options.UseTextOptions = True
        Me.col26.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col26.Caption = "26"
        Me.col26.FieldName = "Hole26"
        Me.col26.MaxWidth = 25
        Me.col26.MinWidth = 25
        Me.col26.Name = "col26"
        Me.col26.Visible = True
        Me.col26.VisibleIndex = 12
        Me.col26.Width = 25
        '
        'col28
        '
        Me.col28.AppearanceHeader.Options.UseTextOptions = True
        Me.col28.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col28.Caption = "28"
        Me.col28.FieldName = "Hole28"
        Me.col28.MaxWidth = 25
        Me.col28.MinWidth = 25
        Me.col28.Name = "col28"
        Me.col28.Visible = True
        Me.col28.VisibleIndex = 13
        Me.col28.Width = 25
        '
        'col30
        '
        Me.col30.AppearanceHeader.Options.UseTextOptions = True
        Me.col30.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col30.Caption = "30"
        Me.col30.FieldName = "Hole30"
        Me.col30.MaxWidth = 25
        Me.col30.MinWidth = 25
        Me.col30.Name = "col30"
        Me.col30.Visible = True
        Me.col30.VisibleIndex = 14
        Me.col30.Width = 25
        '
        'col32
        '
        Me.col32.AppearanceHeader.Options.UseTextOptions = True
        Me.col32.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.col32.Caption = "32"
        Me.col32.FieldName = "Hole32"
        Me.col32.MaxWidth = 25
        Me.col32.MinWidth = 25
        Me.col32.Name = "col32"
        Me.col32.Visible = True
        Me.col32.VisibleIndex = 15
        Me.col32.Width = 25
        '
        'colHoleSuffix
        '
        Me.colHoleSuffix.AppearanceCell.BackColor = System.Drawing.SystemColors.Control
        Me.colHoleSuffix.AppearanceCell.Options.UseBackColor = True
        Me.colHoleSuffix.AppearanceCell.Options.UseTextOptions = True
        Me.colHoleSuffix.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colHoleSuffix.Caption = " "
        Me.colHoleSuffix.FieldName = "HoleSuffix"
        Me.colHoleSuffix.MaxWidth = 25
        Me.colHoleSuffix.MinWidth = 25
        Me.colHoleSuffix.Name = "colHoleSuffix"
        Me.colHoleSuffix.Visible = True
        Me.colHoleSuffix.VisibleIndex = 16
        Me.colHoleSuffix.Width = 25
        '
        'lblCurrLoc
        '
        Me.lblCurrLoc.AutoSize = True
        Me.lblCurrLoc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrLoc.ForeColor = System.Drawing.SystemColors.Highlight
        Me.lblCurrLoc.Location = New System.Drawing.Point(339, 27)
        Me.lblCurrLoc.Name = "lblCurrLoc"
        Me.lblCurrLoc.Size = New System.Drawing.Size(49, 13)
        Me.lblCurrLoc.TabIndex = 48
        Me.lblCurrLoc.Text = "34-24-19"
        Me.lblCurrLoc.Visible = False
        '
        'btnEast
        '
        Me.btnEast.Enabled = False
        Me.btnEast.Location = New System.Drawing.Point(392, 23)
        Me.btnEast.Name = "btnEast"
        Me.btnEast.Size = New System.Drawing.Size(29, 20)
        Me.btnEast.TabIndex = 47
        Me.btnEast.Text = "E"
        Me.btnEast.UseVisualStyleBackColor = True
        '
        'btnSouth
        '
        Me.btnSouth.Enabled = False
        Me.btnSouth.Location = New System.Drawing.Point(348, 42)
        Me.btnSouth.Name = "btnSouth"
        Me.btnSouth.Size = New System.Drawing.Size(29, 20)
        Me.btnSouth.TabIndex = 46
        Me.btnSouth.Text = "S"
        Me.btnSouth.UseVisualStyleBackColor = True
        '
        'btnNorth
        '
        Me.btnNorth.Enabled = False
        Me.btnNorth.Location = New System.Drawing.Point(348, 2)
        Me.btnNorth.Name = "btnNorth"
        Me.btnNorth.Size = New System.Drawing.Size(29, 20)
        Me.btnNorth.TabIndex = 45
        Me.btnNorth.Text = "N"
        Me.btnNorth.UseVisualStyleBackColor = True
        '
        'btnWest
        '
        Me.btnWest.Enabled = False
        Me.btnWest.Location = New System.Drawing.Point(305, 23)
        Me.btnWest.Name = "btnWest"
        Me.btnWest.Size = New System.Drawing.Size(29, 20)
        Me.btnWest.TabIndex = 44
        Me.btnWest.Text = "W"
        Me.btnWest.UseVisualStyleBackColor = True
        '
        'btnGo
        '
        Me.btnGo.Enabled = False
        Me.btnGo.Location = New System.Drawing.Point(248, 33)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(36, 21)
        Me.btnGo.TabIndex = 43
        Me.btnGo.Text = "Go"
        Me.btnGo.UseVisualStyleBackColor = True
        '
        'cboSec
        '
        Me.cboSec.DisplayMember = "Value"
        Me.cboSec.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSec.FormattingEnabled = True
        Me.cboSec.Location = New System.Drawing.Point(162, 33)
        Me.cboSec.Name = "cboSec"
        Me.cboSec.Size = New System.Drawing.Size(64, 21)
        Me.cboSec.TabIndex = 42
        Me.cboSec.ValueMember = "Value"
        '
        'lblSec
        '
        Me.lblSec.AutoSize = True
        Me.lblSec.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSec.Location = New System.Drawing.Point(128, 36)
        Me.lblSec.Name = "lblSec"
        Me.lblSec.Size = New System.Drawing.Size(29, 13)
        Me.lblSec.TabIndex = 41
        Me.lblSec.Text = "Sec"
        '
        'cboRge
        '
        Me.cboRge.DisplayMember = "Value"
        Me.cboRge.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRge.FormattingEnabled = True
        Me.cboRge.Location = New System.Drawing.Point(64, 33)
        Me.cboRge.Name = "cboRge"
        Me.cboRge.Size = New System.Drawing.Size(64, 21)
        Me.cboRge.TabIndex = 40
        Me.cboRge.ValueMember = "Value"
        '
        'cboTwp
        '
        Me.cboTwp.DisplayMember = "Value"
        Me.cboTwp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTwp.FormattingEnabled = True
        Me.cboTwp.Location = New System.Drawing.Point(64, 7)
        Me.cboTwp.Name = "cboTwp"
        Me.cboTwp.Size = New System.Drawing.Size(64, 21)
        Me.cboTwp.TabIndex = 39
        Me.cboTwp.ValueMember = "Value"
        '
        'lblRge
        '
        Me.lblRge.AutoSize = True
        Me.lblRge.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRge.Location = New System.Drawing.Point(30, 36)
        Me.lblRge.Name = "lblRge"
        Me.lblRge.Size = New System.Drawing.Size(30, 13)
        Me.lblRge.TabIndex = 38
        Me.lblRge.Text = "Rge"
        '
        'lblTwp
        '
        Me.lblTwp.AutoSize = True
        Me.lblTwp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTwp.Location = New System.Drawing.Point(29, 11)
        Me.lblTwp.Name = "lblTwp"
        Me.lblTwp.Size = New System.Drawing.Size(31, 13)
        Me.lblTwp.TabIndex = 37
        Me.lblTwp.Text = "Twp"
        '
        'grdSelectedHoles
        '
        Me.grdSelectedHoles.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.grdSelectedHoles.Location = New System.Drawing.Point(596, 66)
        Me.grdSelectedHoles.MainView = Me.grdSelectedHolesView
        Me.grdSelectedHoles.Name = "grdSelectedHoles"
        Me.grdSelectedHoles.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.RepositoryItemCheckEdit1})
        Me.grdSelectedHoles.Size = New System.Drawing.Size(285, 379)
        Me.grdSelectedHoles.TabIndex = 49
        Me.grdSelectedHoles.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.grdSelectedHolesView})
        '
        'grdSelectedHolesView
        '
        Me.grdSelectedHolesView.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.colSelTwp, Me.colSelRng, Me.colSelSec, Me.colSelHole, Me.colRemove})
        Me.grdSelectedHolesView.GridControl = Me.grdSelectedHoles
        Me.grdSelectedHolesView.IndicatorWidth = 30
        Me.grdSelectedHolesView.Name = "grdSelectedHolesView"
        Me.grdSelectedHolesView.OptionsCustomization.AllowFilter = False
        Me.grdSelectedHolesView.OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never
        Me.grdSelectedHolesView.OptionsView.ShowGroupPanel = False
        '
        'colSelTwp
        '
        Me.colSelTwp.AppearanceHeader.Options.UseTextOptions = True
        Me.colSelTwp.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSelTwp.Caption = "Township"
        Me.colSelTwp.FieldName = "Township"
        Me.colSelTwp.Name = "colSelTwp"
        Me.colSelTwp.OptionsColumn.AllowEdit = False
        Me.colSelTwp.Visible = True
        Me.colSelTwp.VisibleIndex = 0
        Me.colSelTwp.Width = 53
        '
        'colSelRng
        '
        Me.colSelRng.AppearanceHeader.Options.UseTextOptions = True
        Me.colSelRng.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSelRng.Caption = "Range"
        Me.colSelRng.FieldName = "Range"
        Me.colSelRng.Name = "colSelRng"
        Me.colSelRng.OptionsColumn.AllowEdit = False
        Me.colSelRng.Visible = True
        Me.colSelRng.VisibleIndex = 1
        Me.colSelRng.Width = 53
        '
        'colSelSec
        '
        Me.colSelSec.AppearanceHeader.Options.UseTextOptions = True
        Me.colSelSec.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSelSec.Caption = "Section"
        Me.colSelSec.FieldName = "Section"
        Me.colSelSec.Name = "colSelSec"
        Me.colSelSec.OptionsColumn.AllowEdit = False
        Me.colSelSec.Visible = True
        Me.colSelSec.VisibleIndex = 2
        Me.colSelSec.Width = 53
        '
        'colSelHole
        '
        Me.colSelHole.AppearanceHeader.Options.UseTextOptions = True
        Me.colSelHole.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.colSelHole.Caption = "Hole"
        Me.colSelHole.FieldName = "HoleLocation"
        Me.colSelHole.Name = "colSelHole"
        Me.colSelHole.OptionsColumn.AllowEdit = False
        Me.colSelHole.Visible = True
        Me.colSelHole.VisibleIndex = 3
        Me.colSelHole.Width = 58
        '
        'colRemove
        '
        Me.colRemove.Caption = "Remove"
        Me.colRemove.ColumnEdit = Me.RepositoryItemCheckEdit1
        Me.colRemove.FieldName = "UnSelect"
        Me.colRemove.Name = "colRemove"
        Me.colRemove.UnboundType = DevExpress.Data.UnboundColumnType.[Boolean]
        Me.colRemove.Visible = True
        Me.colRemove.VisibleIndex = 4
        Me.colRemove.Width = 50
        '
        'RepositoryItemCheckEdit1
        '
        Me.RepositoryItemCheckEdit1.AutoHeight = False
        Me.RepositoryItemCheckEdit1.Name = "RepositoryItemCheckEdit1"
        '
        'grdNonBasicHole
        '
        Me.grdNonBasicHole.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.grdNonBasicHole.Location = New System.Drawing.Point(471, 66)
        Me.grdNonBasicHole.MainView = Me.grdNonBasicHoleView
        Me.grdNonBasicHole.Name = "grdNonBasicHole"
        Me.grdNonBasicHole.Size = New System.Drawing.Size(118, 379)
        Me.grdNonBasicHole.TabIndex = 50
        Me.grdNonBasicHole.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.grdNonBasicHoleView})
        '
        'grdNonBasicHoleView
        '
        Me.grdNonBasicHoleView.Columns.AddRange(New DevExpress.XtraGrid.Columns.GridColumn() {Me.GridColumn1})
        Me.grdNonBasicHoleView.GridControl = Me.grdNonBasicHole
        Me.grdNonBasicHoleView.IndicatorWidth = 30
        Me.grdNonBasicHoleView.Name = "grdNonBasicHoleView"
        Me.grdNonBasicHoleView.OptionsBehavior.Editable = False
        Me.grdNonBasicHoleView.OptionsCustomization.AllowFilter = False
        Me.grdNonBasicHoleView.OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never
        Me.grdNonBasicHoleView.OptionsView.ShowGroupPanel = False
        '
        'GridColumn1
        '
        Me.GridColumn1.Caption = "Non-Basic Hole"
        Me.GridColumn1.FieldName = "HoleLocation"
        Me.GridColumn1.Name = "GridColumn1"
        Me.GridColumn1.Visible = True
        Me.GridColumn1.VisibleIndex = 0
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(724, 451)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 51
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(806, 451)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 52
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'frmProspectGridHoleSelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(892, 482)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.grdNonBasicHole)
        Me.Controls.Add(Me.grdSelectedHoles)
        Me.Controls.Add(Me.lblCurrLoc)
        Me.Controls.Add(Me.btnEast)
        Me.Controls.Add(Me.btnSouth)
        Me.Controls.Add(Me.btnNorth)
        Me.Controls.Add(Me.btnWest)
        Me.Controls.Add(Me.btnGo)
        Me.Controls.Add(Me.cboSec)
        Me.Controls.Add(Me.lblSec)
        Me.Controls.Add(Me.cboRge)
        Me.Controls.Add(Me.cboTwp)
        Me.Controls.Add(Me.lblRge)
        Me.Controls.Add(Me.lblTwp)
        Me.Controls.Add(Me.grdProspect)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmProspectGridHoleSelection"
        Me.Text = "Prospect Grid Hole Selection"
        CType(Me.grdProspect, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdProspectView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdSelectedHoles, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdSelectedHolesView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RepositoryItemCheckEdit1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdNonBasicHole, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdNonBasicHoleView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grdProspect As DevExpress.XtraGrid.GridControl
    Friend WithEvents grdProspectView As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents col02 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col04 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col06 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col08 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col10 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col12 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col14 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col16 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col18 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col20 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col22 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col24 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col26 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col28 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col30 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents col32 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents lblCurrLoc As System.Windows.Forms.Label
    Friend WithEvents btnEast As System.Windows.Forms.Button
    Friend WithEvents btnSouth As System.Windows.Forms.Button
    Friend WithEvents btnNorth As System.Windows.Forms.Button
    Friend WithEvents btnWest As System.Windows.Forms.Button
    Friend WithEvents btnGo As System.Windows.Forms.Button
    Friend WithEvents cboSec As System.Windows.Forms.ComboBox
    Friend WithEvents lblSec As System.Windows.Forms.Label
    Friend WithEvents cboRge As System.Windows.Forms.ComboBox
    Friend WithEvents cboTwp As System.Windows.Forms.ComboBox
    Friend WithEvents lblRge As System.Windows.Forms.Label
    Friend WithEvents lblTwp As System.Windows.Forms.Label
    Friend WithEvents colHoleSuffix As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents grdSelectedHoles As DevExpress.XtraGrid.GridControl
    Friend WithEvents grdSelectedHolesView As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents colSelTwp As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colSelRng As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colSelSec As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colSelHole As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents colRemove As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents RepositoryItemCheckEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemCheckEdit
    Friend WithEvents grdNonBasicHole As DevExpress.XtraGrid.GridControl
    Friend WithEvents grdNonBasicHoleView As DevExpress.XtraGrid.Views.Grid.GridView
    Friend WithEvents GridColumn1 As DevExpress.XtraGrid.Columns.GridColumn
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
