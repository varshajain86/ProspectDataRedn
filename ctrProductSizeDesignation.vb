Imports ProspectDataReduction.ReductionService
Imports ProspectDataReduction.CommonMiningWeb
Imports ProspectDataReduction.ViewModels
Imports DevExpress.XtraGrid.Views.Grid
Imports ProspectDataReduction.Print

Public Class ctrProductSizeDesignation
    Private WithEvents _psizeDefinitionBinder As ProductSizeDesignation

    Private Sub btnGetProductDesignations_Click(sender As System.Object, e As System.EventArgs) Handles btnGetProductDesignations.Click
        GetProductSizeDesignations()
    End Sub

    Private Sub grdProductDesignationView_BeforeLeaveRow(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowAllowEventArgs) Handles grdProductDesignationView.BeforeLeaveRow
        If Not CheckForChanges() Then
            e.Allow = False
        End If
    End Sub

    Private Sub GetProductSizeDesignations()
        If Not CheckForChanges() Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        Dim psizes() As ProductSizeDefinitionData = {}
        Using svc As New ReductionClient
            psizes = svc.GetProspectUserProductSizeDefinitionsBase(If(chkMyDesignationsOnly.Checked, gUserName.ToLower, ""))
        End Using
        grdProductDesignation.DataSource = psizes.OrderBy(Function(a) a.ProductSizeDefinitionName)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub frmProductSizeDesignation_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.UseWaitCursor = True
        My.WebServices.CommonMining.UseDefaultCredentials = True

        Dim mines = My.WebServices.CommonMining.GetMines(False).ToList().OrderBy(Function(m) m.Name).ToList()
        mines.Insert(0, New BusinessEntityAttributeValue())
        cboMineName.DataSource = mines

        Dim groups As New List(Of String) From {""}
        Using svc As New RawService.RawClient
            groups.AddRange(svc.GetWeightTableVersionsByMine(gActiveMineNameLong).ToList())
        End Using

        cboProductGroup.DataBindings(0).DataSourceUpdateMode = DataSourceUpdateMode.Never
        cboProductGroup.DataSource = groups
        cboProductGroup.DataBindings(0).DataSourceUpdateMode = DataSourceUpdateMode.OnPropertyChanged

        GetProductSizeDesignations()
        AddNewProductSizeDesignation()
        If pnlDetails.Enabled Then
            btnSave.Visible = Not _psizeDefinitionBinder.IsReadOnly
            btnSaveAs.Visible = Not btnSave.Visible
            btnDelete.Enabled = Not _psizeDefinitionBinder.IsReadOnly
        End If

        Me.UseWaitCursor = False
    End Sub

    Private Sub grdView_CustomDrawRowIndicator(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs) Handles grdProductDesignationView.CustomDrawRowIndicator,
                                                                                                                                                        grdSFCDistributionView.CustomDrawRowIndicator
        If e.RowHandle >= 0 AndAlso Not DirectCast(sender, GridView).IsNewItemRow(e.RowHandle) Then
            e.Info.DisplayText = (e.RowHandle + 1).ToString()
            e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        End If
    End Sub

    Private Sub grdProductDesignationView_FocusedRowChanged(sender As Object, e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles grdProductDesignationView.FocusedRowChanged
        DisplaySelectedProductSizeDesignation()
    End Sub

    Private Sub DisplaySelectedProductSizeDesignation()
        Me.Cursor = Cursors.WaitCursor
        Dim psizeDefinition As ProductSizeDefinitionData = DirectCast(grdProductDesignationView.GetRow(grdProductDesignationView.FocusedRowHandle), ReductionService.ProductSizeDefinitionData)
        If Not psizeDefinition Is Nothing Then
            Using svc As New ReductionClient
                psizeDefinition = svc.GetProspectUserProductSizeDefinition(psizeDefinition.ProductSizeDefinitionName)
            End Using
            DisplayProductDesignation(psizeDefinition)
        Else
            AddNewProductSizeDesignation()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub DisplayProductDesignation(psizeDefinition As ReductionService.ProductSizeDefinitionData)
        If Not psizeDefinition Is Nothing Then
            _psizeDefinitionBinder = New ProductSizeDesignation(psizeDefinition)
        Else
            _psizeDefinitionBinder = New ProductSizeDesignation()
        End If
        ProductSizeBindingSource.DataSource = _psizeDefinitionBinder
        DetailsBindingSource.DataSource = _psizeDefinitionBinder.Details
    End Sub

    Public Sub AddNewProductSizeDesignation()
        DisplayProductDesignation(Nothing)
    End Sub

    Private Function CheckForChanges() As Boolean
        If Not _psizeDefinitionBinder Is Nothing AndAlso _psizeDefinitionBinder.IsDirty Then
            Dim result = MessageBox.Show("Changes have been made. Would you like to save before you continue?", "Display Product Size Desgination", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            Select Case result
                Case Windows.Forms.DialogResult.Cancel
                    Return False
                Case Windows.Forms.DialogResult.Yes
                    If Not Save() Then
                        Return False
                    End If
                Case Windows.Forms.DialogResult.No
                    _psizeDefinitionBinder.Cancel()
            End Select
        End If
        Return True
    End Function

    Private Sub btnAddNew_Click(sender As System.Object, e As System.EventArgs) Handles btnAddNew.Click
        AddNewProductSizeDesignation()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As System.EventArgs) Handles btnSave.Click, btnSaveAs.Click
        Save()
    End Sub

    Private Function Save()
        Dim success As Boolean = False
        If Not _psizeDefinitionBinder.IsValid Then
            MessageBox.Show(_psizeDefinitionBinder.Error, "Save Product Size Designation", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return success
        End If
        Me.Cursor = Cursors.WaitCursor
        success = _psizeDefinitionBinder.Save
        If success Then
            Dim holdDefinitionName As String = _psizeDefinitionBinder.ProductSizeDesignationName
            MessageBox.Show(String.Format("Product Size Designation {0} saved.", holdDefinitionName), "Save Product Size Designation", MessageBoxButtons.OK, MessageBoxIcon.Information)
            GetProductSizeDesignations()
            Dim rowIndex As Integer = grdProductDesignationView.LocateByValue("ProductSizeDefinitionName", holdDefinitionName)
            grdProductDesignationView.FocusedRowHandle = grdProductDesignationView.GetRowHandle(rowIndex)
        End If
        Me.Cursor = Cursors.Default
        Return success
    End Function

    Private Sub ProductSizeBindingSource_CurrentItemChanged(sender As Object, e As System.EventArgs) Handles ProductSizeBindingSource.CurrentItemChanged
        If ProductSizeBindingSource.List.Count = 0 Then Exit Sub
        btnSaveAs.Enabled = pnlDetails.Enabled AndAlso _psizeDefinitionBinder.IsReadOnly AndAlso Not _psizeDefinitionBinder.OriginalProductSizeDesignationName.ToLower.Equals(_psizeDefinitionBinder.ProductSizeDesignationName.ToLower)
    End Sub

    Private Sub ProductSizeBindingSource_DataSourceChanged(sender As Object, e As System.EventArgs) Handles ProductSizeBindingSource.DataSourceChanged
        pnlDetails.Enabled = Not ProductSizeBindingSource.DataSource Is Nothing AndAlso ProductSizeBindingSource.List.Count > 0
        btnSave.Enabled = pnlDetails.Enabled
        btnDelete.Enabled = False
        btnCancel.Enabled = pnlDetails.Enabled
        btnSaveAs.Enabled = False
        If pnlDetails.Enabled Then
            btnSave.Visible = Not _psizeDefinitionBinder.IsReadOnly
            btnSaveAs.Visible = Not btnSave.Visible
            btnDelete.Enabled = Not _psizeDefinitionBinder.IsReadOnly
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As System.EventArgs) Handles btnDelete.Click
        If MessageBox.Show(String.Format("Are you sure you want to delete {0}?", _psizeDefinitionBinder.ProductSizeDesignationName), "Delete Product Size Designation", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
            Me.Cursor = Cursors.WaitCursor
            If _psizeDefinitionBinder.IsNew Then
                DisplaySelectedProductSizeDesignation()
                Exit Sub
            End If
            If _psizeDefinitionBinder.Delete() Then GetProductSizeDesignations()
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As System.EventArgs) Handles btnCancel.Click
        If _psizeDefinitionBinder.IsDirty Then
            If MessageBox.Show("Are you sure you want to cancel and lose any changes you made?", "Cancel Product Size Designation", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
            _psizeDefinitionBinder.Cancel()
        End If
        DisplaySelectedProductSizeDesignation()
    End Sub

    Public ReadOnly Property ProductSizeDesignation As ProductSizeDesignation
        Get
            Return _psizeDefinitionBinder
        End Get
    End Property

    Private Sub checkEdit_EditValueChanged(sender As Object, e As System.EventArgs) Handles checkEdit.EditValueChanged
        grdSFCDistributionView.PostEditor()
    End Sub

    Private Sub _psizeDefinitionBinder_RebindDetails() Handles _psizeDefinitionBinder.RebindDetails
        DetailsBindingSource.DataSource = _psizeDefinitionBinder.Details
    End Sub

    Private Sub grdProductDesignationView_RowClick(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles grdProductDesignationView.RowClick
        If Not _psizeDefinitionBinder Is Nothing Then
            If Not CheckForChanges() Then
                Exit Sub
            End If
            DisplaySelectedProductSizeDesignation()
        End If
    End Sub

    Private Sub btnPrintGrid_Click(sender As System.Object, e As System.EventArgs) Handles btnPrintGrid.Click
        Dim printGridSetting As New PrintGridSetting
        With printGridSetting
            .GridObject = grdSFCDistribution
            .GridType = GridType.DevExpress
            .PrintGridHeader = "Prospect Data Reduction"
            .PrintGridSubHeader1 = "Product size designation name = " &
                                   _psizeDefinitionBinder.ProductSizeDesignationName
            .PrintGridSubHeader2 = "Based on " & _psizeDefinitionBinder.SizeFractionDistribution
            .OrientHeader = "Center"
            .OrientSubHeader1 = "Center"
            .OrientSubHeader2 = "Center"
            .PrintGridFooter = ""
            .OrientFooter = ""
            .SubHead2IsHeader = False
            .PrintGridDefaultTxtFname = ""
            .PrintMarginLeft = 1440
            .PrintMarginRight = 1440
            .PrintMarginTop = 770
            .PrintMarginBottom = 770
        End With
        Me.Cursor = Cursors.WaitCursor
        Using frm As New frmGridToText(New DevExpressGridPrinter(printGridSetting))
            frm.ShowDialog()
        End Using
        Me.Cursor = Cursors.Arrow
    End Sub

End Class
