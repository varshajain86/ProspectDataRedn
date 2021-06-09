Imports ProspectDataReduction.RawService
Imports ProspectDataReduction.ViewModels
Imports ProspectDataReduction.ValueListWeb
Imports System.ComponentModel
Imports Mosaic.Mining.Common.ViewModels
Imports Mosaic.Mining.Common
Imports DevExpress.XtraGrid.Views.Grid

Public Class frmProspectGridHoleSelection

    Private WithEvents _selectedHoles As New BindingList(Of ProspectGridHole)
    Private _trsHoles As New Dictionary(Of String, ProspectGridHole())
    Private _townships As ValueListDetailDS
    Private _ranges As ValueListDetailDS

    Public ReadOnly Property SelectedHoles
        Get
            Return _selectedHoles
        End Get
    End Property

    Public ReadOnly Property Township As Integer
        Get
            Dim number As Integer = 0
            If Integer.TryParse(cboTwp.SelectedValue, number) Then Return number
            Return 0
        End Get
    End Property

    Public ReadOnly Property Range As Integer
        Get
            Dim number As Integer = 0
            If Integer.TryParse(cboRge.SelectedValue, number) Then Return number
            Return 0
        End Get
    End Property

    Public ReadOnly Property Section As Integer
        Get
            Dim number As Integer = 0
            If Integer.TryParse(cboSec.SelectedValue, number) Then Return number
            Return 0
        End Get
    End Property

    Private Sub grdProspectView_CustomDrawCell(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs) Handles grdProspectView.CustomDrawCell
        Dim row As ProspectGridHoles = DirectCast(grdProspectView.GetRow(e.RowHandle), ProspectGridHoles)
        If row.IsDummy Then
            If e.Column.VisibleIndex < 17 Then
                e.DisplayText = ((e.Column.VisibleIndex + 1) * 2).ToString.PadLeft(2, "0")
                e.Appearance.BackColor = SystemColors.Control
                e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            End If
        Else
            e.Appearance.BackColor = SystemColors.Control
            If Not e.CellValue Is Nothing Then
                If TypeOf e.CellValue Is ViewModels.ProspectGridHole Then
                    With DirectCast(e.CellValue, ViewModels.ProspectGridHole)
                        If Not .MinableStatus Is Nothing Then
                            Select Case .MinableStatus
                                Case "M"
                                    e.Appearance.BackColor = If(.IsSelected, Color.Green, Color.PaleGreen)
                                Case "U"
                                    e.Appearance.BackColor = If(.IsSelected, Color.Red, Color.LightPink)
                                Case ""
                                    e.Appearance.BackColor = If(.IsSelected, Color.Yellow, Color.LightYellow)
                            End Select
                        End If
                    End With
                    e.DisplayText = ""
                End If
            End If
        End If
    End Sub

    Private Sub btnGo_Click(sender As System.Object, e As System.EventArgs) Handles btnGo.Click
        LoadData()
    End Sub

    Private Sub ResetProspectGrid()
        btnGo.Enabled = False
        lblCurrLoc.Visible = False
        btnNorth.Enabled = False
        btnSouth.Enabled = False
        btnEast.Enabled = False
        btnWest.Enabled = False
        grdProspect.DataSource = ProspectGridHoles.GetEmptyProspectGrid
    End Sub

    Private Sub frmProspectGridHoleSelection_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        _townships = My.WebServices.ValueListService.GetDetailDSByGroupName("FL Township")
        Dim vt1 As ValueListDetailDS.ValueListDetailDSRow = _townships.ValueListDetailDS.NewValueListDetailDSRow
        vt1.ID = -1
        vt1.Value = "(Twp...)"
        vt1.SortOrder = -1
        _townships.ValueListDetailDS.Rows.RemoveAt(0)
        _townships.ValueListDetailDS.Rows.InsertAt(vt1, 0)

        Dim sections = My.WebServices.ValueListService.GetDetailDSByGroupName("Section")
        Dim vs1 As ValueListDetailDS.ValueListDetailDSRow = sections.ValueListDetailDS.NewValueListDetailDSRow
        vs1.ID = -1
        vs1.Value = "(Sec...)"
        vs1.SortOrder = -1
        sections.ValueListDetailDS.Rows.RemoveAt(0)
        sections.ValueListDetailDS.Rows.InsertAt(vs1, 0)

        _ranges = My.WebServices.ValueListService.GetDetailDSByGroupName("FL Range")
        Dim vr1 As ValueListDetailDS.ValueListDetailDSRow = _ranges.ValueListDetailDS.NewValueListDetailDSRow
        vr1.ID = -1
        vr1.Value = "(Rge...)"
        vr1.SortOrder = -1
        _ranges.ValueListDetailDS.Rows.RemoveAt(0)
        _ranges.ValueListDetailDS.Rows.InsertAt(vr1, 0)

        cboTwp.DataSource = _townships.ValueListDetailDS
        cboRge.DataSource = _ranges.ValueListDetailDS
        cboSec.DataSource = sections.ValueListDetailDS

        ResetProspectGrid()
        grdSelectedHoles.DataSource = _selectedHoles
    End Sub

    Private Sub grdProspectView_CustomDrawRowIndicator(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs) Handles grdProspectView.CustomDrawRowIndicator
        If e.RowHandle >= 0 AndAlso e.RowHandle < 16 Then
            e.Info.DisplayText = (64 - ((e.RowHandle) * 2)).ToString()
        End If
        e.Info.ImageIndex = -1
    End Sub

    Private Sub grdView_CustomDrawRowIndicator(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs) Handles grdNonBasicHoleView.CustomDrawRowIndicator,
                                                                                                                                                        grdSelectedHolesView.CustomDrawRowIndicator
        If e.RowHandle >= 0 AndAlso Not DirectCast(sender, GridView).IsNewItemRow(e.RowHandle) Then
            e.Info.DisplayText = (e.RowHandle + 1).ToString()
            e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        End If
    End Sub

    Private Sub cboTRS_SelectedValueChanged(sender As Object, e As System.EventArgs) Handles cboTwp.SelectedValueChanged,
                                                                                             cboRge.SelectedValueChanged,
                                                                                             cboSec.SelectedValueChanged
        ResetProspectGrid()
        If Township = 0 Then Exit Sub
        If Range = 0 Then Exit Sub
        If Section = 0 Then Exit Sub
        lblCurrLoc.Text = String.Format("{0}-{1}-{2}", Township.ToString, Range.ToString, Section.ToString)
        lblCurrLoc.Visible = True
        btnGo.Enabled = True
    End Sub

    Private Sub grdProspectView_RowCellClick(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs) Handles grdProspectView.RowCellClick
        If Not e.CellValue Is Nothing AndAlso TypeOf e.CellValue Is ProspectGridHole Then
            With DirectCast(e.CellValue, ProspectGridHole)
                AddHandler .IsSelectedChanged, AddressOf OnIsSelectedChanged
                .IsSelected = Not .IsSelected
                RemoveHandler .IsSelectedChanged, AddressOf OnIsSelectedChanged
                grdProspectView.InvalidateRowCell(e.RowHandle, e.Column)
            End With
        End If
    End Sub

    Private Sub grdNonBasicHoleView_CustomDrawCell(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs) Handles grdNonBasicHoleView.CustomDrawCell
        With DirectCast(grdNonBasicHoleView.GetRow(e.RowHandle), ProspectGridHole)
            If Not .MinableStatus Is Nothing Then
                Select Case .MinableStatus
                    Case "M"
                        e.Appearance.BackColor = If(.IsSelected, Color.Green, Color.PaleGreen)
                        e.Appearance.ForeColor = If(.IsSelected, Color.White, Color.Black)
                    Case "U"
                        e.Appearance.BackColor = If(.IsSelected, Color.Red, Color.LightPink)
                        e.Appearance.ForeColor = If(.IsSelected, Color.White, Color.Black)
                    Case ""
                        e.Appearance.BackColor = If(.IsSelected, Color.Yellow, Color.LightYellow)
                End Select
            End If
        End With
    End Sub

    Private Sub grdNonBasicHoleView_RowCellClick(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs) Handles grdNonBasicHoleView.RowCellClick
        With DirectCast(grdNonBasicHoleView.GetRow(e.RowHandle), ProspectGridHole)
            AddHandler .IsSelectedChanged, AddressOf OnIsSelectedChanged
            .IsSelected = Not .IsSelected
            RemoveHandler .IsSelectedChanged, AddressOf OnIsSelectedChanged
            grdNonBasicHoleView.InvalidateRowCell(e.RowHandle, e.Column)
        End With
    End Sub

    Private Sub OnIsSelectedChanged(sender As Object, e As EventArgs)
        Dim clickedHole As ProspectGridHole = DirectCast(sender, ProspectGridHole)
        If clickedHole.IsSelected Then
            If Not _selectedHoles.Contains(clickedHole) Then _selectedHoles.Add(clickedHole)
        Else
            If _selectedHoles.Contains(clickedHole) Then _selectedHoles.Remove(clickedHole)
        End If
    End Sub

    Private Sub _selectedHoles_ListChanged(sender As Object, e As System.ComponentModel.ListChangedEventArgs) Handles _selectedHoles.ListChanged
        Select Case e.ListChangedType
            Case ListChangedType.ItemChanged
                OnIsSelectedChanged(_selectedHoles(e.NewIndex), New EventArgs)
            Case ListChangedType.ItemDeleted
                grdNonBasicHoleView.Invalidate()
                grdProspectView.Invalidate()
        End Select
    End Sub

    Private Sub RepositoryItemCheckEdit1_EditValueChanged(sender As System.Object, e As System.EventArgs) Handles RepositoryItemCheckEdit1.EditValueChanged
        grdSelectedHolesView.PostEditor()
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnSouth_Click(sender As System.Object, e As System.EventArgs) Handles btnSouth.Click
        GetNewSection("South")
    End Sub

    Private Sub btnWest_Click(sender As Object, e As System.EventArgs) Handles btnWest.Click
        GetNewSection("West")
    End Sub

    Private Sub btnNorth_Click(sender As Object, e As System.EventArgs) Handles btnNorth.Click
        GetNewSection("North")
    End Sub

    Private Sub btnEast_Click(sender As Object, e As System.EventArgs) Handles btnEast.Click
        GetNewSection("East")
    End Sub

    Private Function SectionMove(aDirection As String) As Boolean
        Dim currentHole As New HoleInfo With {.Section = Section, .Range = Range, .Township = Township}
        Dim nextSection As HoleInfo = MineHoleHelper.GetNextSection(currentHole, aDirection, _townships.ValueListDetailDS, _ranges.ValueListDetailDS)
        If currentHole.Township <> nextSection.Township Then
            cboTwp.SelectedValue = nextSection.Township.ToString
        End If
        If currentHole.Range <> nextSection.Range Then
            cboRge.SelectedValue = nextSection.Range.ToString
        End If
        If currentHole.Section <> nextSection.Section Then
            cboSec.SelectedValue = nextSection.Section
        End If
        Return currentHole.Section <> nextSection.Section
    End Function

    Private Sub GetNewSection(direction As String)
        Dim canMove As Boolean = SectionMove(direction)
        If canMove Then
            If Township > 0 Or Range > 0 Or Section > 0 Then
                btnGo.Enabled = True
                LoadData()
            Else
                btnGo.Enabled = False
            End If
        End If
    End Sub

    Private Sub LoadData()
        Dim prospectGridHoles() As ProspectGridHole = {}
        Dim key As String = String.Format("{0}-{1}-{2}", Township, Range, Section)
        If _trsHoles.ContainsKey(key) Then
            prospectGridHoles = _trsHoles(key)
        Else
            Dim holes As RawProspectHoles() = Nothing
            Using client = New RawClient
                holes = client.GetProspHolesList(Township, Range, Section)
            End Using
            If Not holes Is Nothing Then
                prospectGridHoles = holes.Select(Function(h) New ProspectGridHole(h)).ToArray()
            End If
            _trsHoles.Add(key, prospectGridHoles)
        End If

        Dim nonBasicHoles As New List(Of ProspectGridHole)
        nonBasicHoles.AddRange(prospectGridHoles.Where(Function(h) h.IsNonBasicHole))
        grdNonBasicHole.DataSource = nonBasicHoles
        Dim dataSourceList As New List(Of ProspectGridHoles)
        For i = 64 To 34 Step -2
            dataSourceList.Add(New ProspectGridHoles(prospectGridHoles.Where(Function(h) h.HoleSuffix.Equals(i.ToString)).ToArray(), i.ToString))
        Next
        dataSourceList.Add(ViewModels.ProspectGridHoles.GetDummyRow)
        grdProspect.DataSource = dataSourceList
        btnNorth.Enabled = True
        btnSouth.Enabled = True
        btnEast.Enabled = True
        btnWest.Enabled = True
    End Sub
End Class