Imports ProspectDataReduction.ReductionService
Imports ProspectDataReduction.CommonMiningWeb
Imports System.ComponentModel
Imports ProspectDataReduction.ViewModels
Imports System.IO
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraGrid

Public Class ctrAreaDefinition

    Private WithEvents _areaDefinitionBinder As ProspectAreaDefinition

    Private Sub btnGetAreaDefinitions_Click(sender As System.Object, e As System.EventArgs) Handles btnGetAreaDefinitions.Click
        GetAreaDefinitions()
    End Sub

    Private Sub GetAreaDefinitions()
        If Not CheckForChanges() Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        Dim areas() As ProspectAreaDefinitionData = {}
        Using svc As New ReductionClient
            areas = svc.GetProspectUserAreaDefinitionsBase(If(chkMyDefinitionsOnly.Checked, gUserName.ToLower, ""))
        End Using
        grdAreaDefinition.DataSource = areas.OrderBy(Function(a) a.AreaDefinitionName)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub grdAreaDefinitionView_BeforeLeaveRow(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowAllowEventArgs) Handles grdAreaDefinitionView.BeforeLeaveRow
        If Not CheckForChanges() Then
            e.Allow = False
        End If
    End Sub

    Private Sub grdView_CustomDrawRowIndicator(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs) Handles grdAreaDefinitionView.CustomDrawRowIndicator
        If e.RowHandle >= 0 AndAlso Not DirectCast(sender, GridView).IsNewItemRow(e.RowHandle) Then
            e.Info.DisplayText = (e.RowHandle + 1).ToString()
            e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        End If
    End Sub

    Private Sub grdAreaDefinitionView_FocusedRowChanged(sender As Object, e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles grdAreaDefinitionView.FocusedRowChanged
        DisplaySelectedAreaDefinition()
    End Sub

    Private Function CheckForChanges() As Boolean
        If Not _areaDefinitionBinder Is Nothing AndAlso _areaDefinitionBinder.IsDirty Then
            Dim result = MessageBox.Show("Changes have been made. Would you like to save before you continue?", "Display Area Definition", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            Select Case result
                Case Windows.Forms.DialogResult.Cancel
                    Return False
                Case Windows.Forms.DialogResult.Yes
                    If Not Save() Then
                        Return False
                    End If
                Case Windows.Forms.DialogResult.No
                    _areaDefinitionBinder.Cancel()
            End Select
        End If
        Return True
    End Function

    Private Sub DisplaySelectedAreaDefinition()
        Me.Cursor = Cursors.WaitCursor
        Dim areaDefinition As ProspectAreaDefinitionData = DirectCast(grdAreaDefinitionView.GetRow(grdAreaDefinitionView.FocusedRowHandle), ReductionService.ProspectAreaDefinitionData)
        If Not areaDefinition Is Nothing Then
            Using svc As New ReductionClient
                areaDefinition = svc.GetProspectUserAreaDefinition(areaDefinition.AreaDefinitionName)
            End Using
            DisplayAreaDefinition(areaDefinition)
        Else
            AddNewAreaDefinition()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub DisplayAreaDefinition(areaDefinition As ReductionService.ProspectAreaDefinitionData)
        If Not areaDefinition Is Nothing Then
            _areaDefinitionBinder = New ProspectAreaDefinition(areaDefinition)
        Else
            _areaDefinitionBinder = New ProspectAreaDefinition()
        End If
        AreaDefinitionBindingSource.DataSource = _areaDefinitionBinder
        HoleBindingSource.DataSource = _areaDefinitionBinder.Holes
        TRSCornerBindingSource.DataSource = _areaDefinitionBinder.TRSCorners
        XYCornerBindingSource.DataSource = _areaDefinitionBinder.XYCorners
        Select Case True
            Case _areaDefinitionBinder.ByHolesAreaMethod
                rdoHoles.Checked = True
            Case _areaDefinitionBinder.ByMineAreaAreaMethod
                rdoMineArea.Checked = True
            Case _areaDefinitionBinder.ByTRSCornersAreaMethod
                rdoTRSCorner.Checked = True
            Case _areaDefinitionBinder.ByXYCoordinatesAreaMethod
                rdoXYCorner.Checked = True
            Case Else
                rdoHoles.Checked = False
                rdoMineArea.Checked = False
                rdoTRSCorner.Checked = False
                rdoXYCorner.Checked = False
        End Select
        gHaveRawProspData = False
    End Sub

    Private Sub ctrAreaDefinition_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        My.WebServices.CommonMining.UseDefaultCredentials = True
        Dim mines = My.WebServices.CommonMining.GetMines(False).ToList().OrderBy(Function(m) m.Name).ToList()
        mines.Insert(0, New BusinessEntityAttributeValue())
        cboMineName.DataSource = mines
        Dim areas = My.WebServices.CommonMining.GetAreas(False).ToList().OrderBy(Function(a) a.Name).ToList()
        areas.Insert(0, New BusinessEntityAttributeValue())
        cboAreaName.DataSource = areas
        cboMine.DataSource = mines.ToList()
        cboArea.DataSource = areas.ToList()

        My.WebServices.ValueListService.UseDefaultCredentials = True

        Dim townshipsList = My.WebServices.ValueListService.GetDetailDSByGroupName("FL Township")
        townshipsList.ValueListDetailDS.Rows.RemoveAt(0)
        TownshipBindingSource.DataSource = townshipsList.ValueListDetailDS

        Dim sectionsList = My.WebServices.ValueListService.GetDetailDSByGroupName("Section")
        sectionsList.ValueListDetailDS.Rows.RemoveAt(0)
        SectionBindingSource.DataSource = sectionsList.ValueListDetailDS

        Dim rangesList = My.WebServices.ValueListService.GetDetailDSByGroupName("FL Range")
        rangesList.ValueListDetailDS.Rows.RemoveAt(0)
        RangeBindingSource.DataSource = rangesList.ValueListDetailDS

        Dim holesList = My.WebServices.ValueListService.GetDetailDSByGroupName("HOP Hole 1")
        holesList.ValueListDetailDS.Rows.RemoveAt(0)
        HolesListBindingSource.DataSource = holesList.ValueListDetailDS
        Dim ownserhipCodes() As RawService.ProspectCode
        Try
            Using svc As New RawService.RawClient
                ownserhipCodes = svc.GetProspCodesList("Ownership")
            End Using
        Catch ex As Exception
            MessageBox.Show("Error")
        End Try

        If Not ownserhipCodes Is Nothing Then
            cboOwnership.Properties.DataSource = ownserhipCodes
        End If

        GetAreaDefinitions()

        If pnlDetails.Enabled Then
            btnSave.Visible = Not _areaDefinitionBinder.IsReadOnly
            btnSaveAs.Visible = Not btnSave.Visible
            btnDelete.Enabled = Not _areaDefinitionBinder.IsReadOnly
        End If

    End Sub

    Private Sub rdoTRSCorner_CheckedChanged(sender As Object, e As System.EventArgs) Handles rdoTRSCorner.CheckedChanged
        pnlTRSCorner.Visible = rdoTRSCorner.Checked
        pnlHoles.Visible = Not rdoTRSCorner.Checked
        pnlXYCorner.Visible = Not rdoTRSCorner.Checked
        pnlMineArea.Visible = Not rdoTRSCorner.Checked
        _areaDefinitionBinder.ByTRSCornersAreaMethod = rdoTRSCorner.Checked
        If rdoTRSCorner.Checked Then grpSelectHoles.Text = "Select T-R-S Corners"
    End Sub

    Private Sub rdoXYCorner_CheckedChanged(sender As Object, e As System.EventArgs) Handles rdoXYCorner.CheckedChanged
        pnlXYCorner.Visible = rdoXYCorner.Checked
        pnlTRSCorner.Visible = Not rdoXYCorner.Checked
        pnlHoles.Visible = Not rdoXYCorner.Checked
        pnlMineArea.Visible = Not rdoXYCorner.Checked
        _areaDefinitionBinder.ByXYCoordinatesAreaMethod = rdoXYCorner.Checked
        If rdoXYCorner.Checked Then grpSelectHoles.Text = "Select X && Y Coordinates"
    End Sub

    Private Sub rdoMineArea_CheckedChanged(sender As Object, e As System.EventArgs) Handles rdoMineArea.CheckedChanged
        pnlMineArea.Visible = rdoMineArea.Checked
        pnlXYCorner.Visible = Not rdoMineArea.Checked
        pnlTRSCorner.Visible = Not rdoMineArea.Checked
        pnlHoles.Visible = Not rdoMineArea.Checked
        _areaDefinitionBinder.ByMineAreaAreaMethod = rdoMineArea.Checked
        If rdoMineArea.Checked Then grpSelectHoles.Text = "Select Mine/Area"
    End Sub

    Private Sub rdoHoles_CheckedChanged(sender As Object, e As System.EventArgs) Handles rdoHoles.CheckedChanged
        pnlHoles.Visible = rdoHoles.Checked
        pnlMineArea.Visible = Not rdoHoles.Checked
        pnlXYCorner.Visible = Not rdoHoles.Checked
        pnlTRSCorner.Visible = Not rdoHoles.Checked
        _areaDefinitionBinder.ByHolesAreaMethod = rdoHoles.Checked
        If rdoHoles.Checked Then grpSelectHoles.Text = "Select Holes"
    End Sub

    Private Sub pnlAreaMethod_VisibleChanged(sender As Object, e As System.EventArgs) Handles pnlHoles.VisibleChanged,
                                                                                              pnlMineArea.VisibleChanged,
                                                                                              pnlTRSCorner.VisibleChanged,
                                                                                              pnlXYCorner.VisibleChanged
        With DirectCast(sender, Panel)
            .Dock = If(.Visible, DockStyle.Fill, DockStyle.Top)
        End With
    End Sub

    Private Sub grdHolesFromFileView_CellValueChanged(sender As Object, e As DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs) Handles grdHolesFromFileView.CellValueChanged
        Dim col As GridColumn = Nothing
        If grdHolesFromFileView.IsNewItemRow(e.RowHandle) OrElse grdHolesFromFileView.DataRowCount = (e.RowHandle + 1) Then
            Select Case True
                Case e.Column Is colTownship
                    lupHoleTownship.NullText = e.Value
                    col = colTownship
                Case e.Column Is colRange
                    lupHoleRange.NullText = e.Value
                    col = colRange
                Case e.Column Is colSection
                    lupHoleSection.NullText = e.Value
                    col = colSection
                Case e.Column Is colHole
                    lupHoleHole.NullText = e.Value
                    col = colHole
            End Select
        End If
        'US 1487 : Prospect -Area Definition input screen adjustments
        gHaveRawProspData = False
    End Sub

    Private Sub grdView_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles grdHolesFromFileView.KeyUp,
                                                                                                grdXYCornerView.KeyUp,
                                                                                                grdTRSCornerView.KeyUp
        With DirectCast(sender, GridView)
            If e.KeyCode = Keys.Delete Then
                Dim rowHandle As Integer = .FocusedRowHandle
                .DeleteRow(rowHandle)
            End If
        End With
    End Sub

    Private Sub HoleBindingSource_AddingNew(sender As Object, e As System.ComponentModel.AddingNewEventArgs) Handles HoleBindingSource.AddingNew
        Dim township As Integer = 0
        Dim range As Integer = 0
        Dim section As Integer = 0
        Integer.TryParse(lupHoleTownship.NullText, township)
        Integer.TryParse(lupHoleRange.NullText, range)
        Integer.TryParse(lupHoleSection.NullText, section)
        e.NewObject = New ProspectAreaHole With {.Hole_Township = township,
                                                 .Hole_Range = range,
                                                 .Hole_Section = section}
    End Sub


    Private Sub HoleBindingSource_DataSourceChanged(sender As Object, e As System.EventArgs) Handles HoleBindingSource.DataSourceChanged
        If HoleBindingSource.List.Count > 0 Then
            Dim sourceList As BindingList(Of ProspectAreaHole) = DirectCast(HoleBindingSource.List, BindingList(Of ProspectAreaHole))
            lupHoleTownship.NullText = sourceList.Last.Hole_Township
            lupHoleRange.NullText = sourceList.Last.Hole_Range
            lupHoleSection.NullText = sourceList.Last.Hole_Section
            grdHolesFromFileView.RefreshData()
        Else
            lupHoleTownship.NullText = ""
            lupHoleRange.NullText = ""
            lupHoleSection.NullText = ""
        End If
    End Sub

    Private Sub btnAddNew_Click(sender As System.Object, e As System.EventArgs) Handles btnAddNew.Click
        AddNewAreaDefinition()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As System.EventArgs) Handles btnSave.Click, btnSaveAs.Click
        'US 1487 : Prospect -Area Definition input screen adjustments
        txtFileName.Select()
        Save()
    End Sub

    Private Function Save()
        Dim success As Boolean = False
        If Not _areaDefinitionBinder.IsValid Then
            MessageBox.Show(_areaDefinitionBinder.Error, "Save Area Defintion", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return success
        End If
        Me.Cursor = Cursors.WaitCursor
        success = _areaDefinitionBinder.Save
        If success Then
            Dim holdAreaDefinitionName As String = _areaDefinitionBinder.AreaDefinitionName
            MessageBox.Show(String.Format("Area Definition {0} saved.", holdAreaDefinitionName), "Save Area Definition", MessageBoxButtons.OK, MessageBoxIcon.Information)
            GetAreaDefinitions()
            grdAreaDefinitionView.FocusedRowHandle = grdAreaDefinitionView.LocateByValue("AreaDefinitionName", holdAreaDefinitionName)
        End If
        Me.Cursor = Cursors.Default
        Return success
    End Function

    Private Sub btnBrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowse.Click
        Dim dialog As New OpenFileDialog
        dialog.InitialDirectory = _areaDefinitionBinder.FileName
        If dialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtFileName.Text = dialog.FileName
        End If
    End Sub

    Private Sub btnFromFile_Click(sender As Object, e As EventArgs) Handles btnAddHolesFromFile.Click
        If String.IsNullOrEmpty(txtFileName.Text) Then
            MessageBox.Show("Please specify file name.", "Area Definition", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If System.IO.File.Exists(txtFileName.Text) Then
            Dim sourceList As BindingList(Of ProspectAreaHole) = DirectCast(HoleBindingSource.List, BindingList(Of ProspectAreaHole))
            Dim tempList As New List(Of ProspectAreaHole)
            Dim reader = New IO.StreamReader(txtFileName.Text)
            While Not reader.EndOfStream
                tempList.Add(New ProspectAreaHole(reader.ReadLine()))
            End While
            tempList.ForEach(Sub(h) If h.IsValidHole Then sourceList.Add(h))
        Else
            MessageBox.Show(String.Format("File {0} does not exist.", txtFileName.Text), "Area Definition", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub AreaDefinitionBindingSource_CurrentItemChanged(sender As Object, e As System.EventArgs) Handles AreaDefinitionBindingSource.CurrentItemChanged
        If AreaDefinitionBindingSource.List.Count = 0 Then Exit Sub
        btnSaveAs.Enabled = pnlDetails.Enabled AndAlso _areaDefinitionBinder.IsReadOnly AndAlso Not _areaDefinitionBinder.OriginalAreaDefintionName.ToLower.Equals(_areaDefinitionBinder.AreaDefinitionName.ToLower)
    End Sub

    Private Sub AreaDefinitionBindingSource_DataSourceChanged(sender As Object, e As System.EventArgs) Handles AreaDefinitionBindingSource.DataSourceChanged
        pnlDetails.Enabled = Not AreaDefinitionBindingSource.DataSource Is Nothing AndAlso AreaDefinitionBindingSource.List.Count > 0
        btnSave.Enabled = pnlDetails.Enabled
        btnDelete.Enabled = False
        btnCancel.Enabled = pnlDetails.Enabled
        btnSaveAs.Enabled = False
        If pnlDetails.Enabled Then
            btnSave.Visible = Not _areaDefinitionBinder.IsReadOnly
            btnSaveAs.Visible = Not btnSave.Visible
            btnDelete.Enabled = Not _areaDefinitionBinder.IsReadOnly
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As System.EventArgs) Handles btnDelete.Click
        If MessageBox.Show(String.Format("Are you sure you want to delete {0}?", _areaDefinitionBinder.AreaDefinitionName), "Delete Area Defintion", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
            Me.Cursor = Cursors.WaitCursor
            If _areaDefinitionBinder.IsNew Then
                DisplaySelectedAreaDefinition()
                Exit Sub
            End If
            If _areaDefinitionBinder.Delete() Then GetAreaDefinitions()
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As System.EventArgs) Handles btnCancel.Click
        If _areaDefinitionBinder.IsDirty Then
            If MessageBox.Show("Are you sure you want to cancel and lose any changes you made?", "Cancel Area Definition", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
            _areaDefinitionBinder.Cancel()
        End If
        DisplaySelectedAreaDefinition()
    End Sub

    Private Sub btnFromProspectGrid_Click(sender As System.Object, e As EventArgs) Handles btnAddHolesFromProspectGrid.Click
        Using frm As New frmProspectGridHoleSelection
            If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim sourceList As BindingList(Of ProspectAreaHole) = DirectCast(HoleBindingSource.List, BindingList(Of ProspectAreaHole))
                For Each item As ProspectGridHole In frm.SelectedHoles
                    sourceList.Add(New ProspectAreaHole(item))
                Next
            End If
        End Using
    End Sub

    Public ReadOnly Property AreaDefinition As ProspectAreaDefinition
        Get
            Return _areaDefinitionBinder
        End Get
    End Property

    Public Sub AddNewAreaDefinition()
        DisplayAreaDefinition(Nothing)
    End Sub

    Private Sub _areaDefinitionBinder_PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Handles _areaDefinitionBinder.PropertyChanged
        gHaveRawProspData = False
    End Sub

    Private Sub grdAreaDefinitionView_RowClick(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles grdAreaDefinitionView.RowClick
        If Not _areaDefinitionBinder Is Nothing Then
            If Not CheckForChanges() Then
                Exit Sub
            End If
            DisplaySelectedAreaDefinition()
        End If
    End Sub


    Private Sub lupHoleFromFile_EditValueChanged(sender As Object, e As System.EventArgs) Handles lupHoleHole.EditValueChanged, lupHoleTownship.EditValueChanged, lupHoleRange.EditValueChanged, lupHoleSection.EditValueChanged
        grdHolesFromFileView.PostEditor()
        grdHolesFromFileView.UpdateCurrentRow()
    End Sub

    Private Sub btnClearTRS_Click(sender As System.Object, e As System.EventArgs) Handles btnClearTRS.Click
        AreaDefinition.TRSCorners.Clear()
    End Sub

    Private Sub btnClearHoles_Click(sender As System.Object, e As System.EventArgs) Handles btnClearHoles.Click
        AreaDefinition.Holes.Clear()
    End Sub

    Private Sub btnClearXY_Click(sender As System.Object, e As System.EventArgs) Handles btnClearXY.Click
        AreaDefinition.XYCorners.Clear()
    End Sub

    Private Sub grdAreaDefinition_TextChanged(sender As Object, e As EventArgs) Handles grdAreaDefinition.TextChanged

    End Sub

    Private Sub grdAreaDefinition_FocusedViewChanged(sender As Object, e As ViewFocusEventArgs) Handles grdAreaDefinition.FocusedViewChanged

    End Sub
End Class
