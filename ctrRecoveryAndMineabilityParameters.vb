Imports ProspectDataReduction.ReductionService
Imports ProspectDataReduction.CommonMiningWeb
Imports ProspectDataReduction.ViewModels
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraEditors
Imports DevExpress.Utils

Public Class ctrRecoveryAndMineabilityParameters
    Private WithEvents _recoveryDefinitionBinder As ProductRecoveryDefinition
    Private _parentForm As frmProspDataReduction
    Private _hidingFlyoutPanel As Boolean


    Private _loadingData As Boolean
    Private Sub ctrRecoveryAndMineabilityParameters_Load(sender As Object, e As EventArgs) Handles Me.Load
        'If Not Me.IsMdiChild Then pnlScenarios.OwnerControl = Me.Parent ' this will be the tab control
        My.WebServices.CommonMining.UseDefaultCredentials = True
        Dim mines = My.WebServices.CommonMining.GetMines(False).ToList().OrderBy(Function(m) m.Name).ToList()
        mines.Insert(0, New BusinessEntityAttributeValue())
        cboMineName.DataSource = mines

        pnlScenarios.ShowPopup()
        GetProductRecoveryScenarios()
        AddNewProductRecoveryScenario()
        If pnlDetails.Enabled Then
            btnSave.Visible = Not _recoveryDefinitionBinder.IsReadOnly
            btnSaveAs.Visible = Not btnSave.Visible
            btnDelete.Enabled = Not _recoveryDefinitionBinder.IsReadOnly
        End If
    End Sub
    Public Sub New(parentForm As frmProspDataReduction)
        InitializeComponent()
        _parentForm = parentForm
    End Sub

    Private Sub btnShowScenarios_Click(sender As System.Object, e As System.EventArgs) Handles btnShowScenarios.Click
        pnlScenarios.ShowPopup()
    End Sub

    Private Sub pnlScenarios_ButtonClick(sender As Object, e As DevExpress.Utils.FlyoutPanelButtonClickEventArgs) Handles pnlScenarios.ButtonClick
        pnlScenarios.HidePopup()
    End Sub

    Private Sub pnlScenarios_Hidden(sender As Object, e As DevExpress.Utils.FlyoutPanelEventArgs) Handles pnlScenarios.Hidden
        tspScenarios.Visible = True
        _hidingFlyoutPanel = False
    End Sub

    Private Sub pnlScenarios_Hiding(sender As Object, e As DevExpress.Utils.FlyoutPanelEventArgs) Handles pnlScenarios.Hiding
        _hidingFlyoutPanel = True
    End Sub

    Private Sub pnlScenarios_Shown(sender As Object, e As DevExpress.Utils.FlyoutPanelEventArgs) Handles pnlScenarios.Shown
        tspScenarios.Visible = False
    End Sub

    Private Sub GetProductRecoveryScenarios()
        If Not CheckForChanges() Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        Dim recoveries() As ProductRecoveryDefinitionData = {}
        Using svc As New ReductionService.ReductionClient
            recoveries = svc.GetProspectUserRecoveryDefinitionsBase(If(chkMyScenariosOnly.Checked, gUserName.ToLower, ""))
        End Using
        grdScenarios.DataSource = recoveries.OrderBy(Function(a) a.RCVRY_SCENARIO_NAME)
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub btnGetScenarios_Click(sender As Object, e As System.EventArgs) Handles btnGetScenarios.Click
        GetProductRecoveryScenarios()
    End Sub

    Private Sub grdView_CustomDrawRowIndicator(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs) Handles grdScenariosView.CustomDrawRowIndicator
        If e.RowHandle >= 0 AndAlso Not DirectCast(sender, GridView).IsNewItemRow(e.RowHandle) Then
            e.Info.DisplayText = (e.RowHandle + 1).ToString()
            e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far
        End If
    End Sub

    Private Sub grdScenariosView_BeforeLeaveRow(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowAllowEventArgs) Handles grdScenariosView.BeforeLeaveRow
        pnlScenarios.Options.CloseOnOuterClick = False
        If Not CheckForChanges() Then
            e.Allow = False
        End If
        pnlScenarios.Options.CloseOnOuterClick = True
    End Sub

    Private Sub grdScenariosView_FocusedRowChanged(sender As Object, e As DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs) Handles grdScenariosView.FocusedRowChanged
        If _hidingFlyoutPanel Then Exit Sub
        DisplaySelectedProductRecoveryScenario()
    End Sub

    Private Sub DisplaySelectedProductRecoveryScenario()
        Me.Cursor = Cursors.WaitCursor
        Dim recoveryDefinition As ProductRecoveryDefinitionData = DirectCast(grdScenariosView.GetRow(grdScenariosView.FocusedRowHandle), ReductionService.ProductRecoveryDefinitionData)
        If Not recoveryDefinition Is Nothing Then
            Using svc As New ReductionClient
                recoveryDefinition = svc.GetProspectUserProductRecoveryDefinition(recoveryDefinition.RCVRY_SCENARIO_NAME)
            End Using
            DisplayProductRecovery(recoveryDefinition)
        Else
            AddNewProductRecoveryScenario()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub DisplayProductRecovery(recoveryDefinition As ReductionService.ProductRecoveryDefinitionData)
        _loadingData = True
        _parentForm.ClearRcvryEtc()

        If recoveryDefinition IsNot Nothing Then
            _recoveryDefinitionBinder = New ProductRecoveryDefinition(recoveryDefinition)
        Else
            _recoveryDefinitionBinder = New ProductRecoveryDefinition()
        End If

        ProductRecoveryBindingSource.DataSource = _recoveryDefinitionBinder
        RecoveryDilutionParamBindingSource.DataSource = _recoveryDefinitionBinder.RecoveryDilutionParamaters
        MiniabilityParamBindingSource.DataSource = _recoveryDefinitionBinder.MiniabilityParamaters
        HoleQualitySpecificationBindingSource.DataSource = _recoveryDefinitionBinder.HoleQualitySpecifications
        SplitQualitySpecificationBindingSource.DataSource = _recoveryDefinitionBinder.SplitQualitySpecifications
        PebbleRejectBindingSource.DataSource = _recoveryDefinitionBinder.PebbleRejectCriteria
        ProductQualityCoefficientsBindingSource.DataSource = _recoveryDefinitionBinder.ProductQualityCoefficients

        If _recoveryDefinitionBinder.IsNew Then
            rdoCalculated.Checked = False
            rdoLimitRules.Checked = False
            rdoMeasuredLab.Checked = False
        Else
            _parentForm.DisplayRcvryEtc(_recoveryDefinitionBinder)
            Select Case True
                Case _recoveryDefinitionBinder.DensityCalculatedCalculationMode
                    rdoCalculated.Checked = True
                Case _recoveryDefinitionBinder.DensityLimitRuleCalculationMode
                    rdoLimitRules.Checked = True
                Case _recoveryDefinitionBinder.DensityMeasuredLabCalculationMode
                    rdoMeasuredLab.Checked = True
            End Select
        End If
        _recoveryDefinitionBinder.DisplayRecoveryScenarioComplete()
        tabScenario.SelectedTabPage = pgDensity
        _loadingData = False
    End Sub

    Public Sub AddNewProductRecoveryScenario()
        DisplayProductRecovery(Nothing)
    End Sub

    Private Function CheckForChanges() As Boolean
        If Not _recoveryDefinitionBinder Is Nothing AndAlso _recoveryDefinitionBinder.IsDirty Then
            Dim result = MessageBox.Show("Changes have been made. Would you like to save before you continue?", "Display Product Recovery Scenario", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
            Select Case result
                Case Windows.Forms.DialogResult.Cancel
                    Return False
                Case Windows.Forms.DialogResult.Yes
                    If Not Save() Then
                        Return False
                    End If
                Case Windows.Forms.DialogResult.No
                    _recoveryDefinitionBinder.Cancel()
            End Select
        End If
        Return True
    End Function

    Private Sub btnAddNew_Click(sender As System.Object, e As System.EventArgs) Handles btnAddNew.Click
        AddNewProductRecoveryScenario()
    End Sub

    Private Sub EndEditOnAllBindingSources()

        If Me.components IsNot Nothing Then
            Dim Comps As System.ComponentModel.ComponentCollection = Me.components.Components
            If Comps IsNot Nothing Then
                Dim BindingSourcesQuery = From bindingsources In Me.components.Components
                                          Where (TypeOf bindingsources Is Windows.Forms.BindingSource)
                                          Select bindingsources

                For Each bindingSource As Windows.Forms.BindingSource In BindingSourcesQuery
                    bindingSource.EndEdit()
                Next
            End If
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As System.EventArgs) Handles btnSave.Click, btnSaveAs.Click
        EndEditOnAllBindingSources()
        Save()
    End Sub

    Private Function Save() As Boolean
        Dim success As Boolean = False
        If Not _recoveryDefinitionBinder.IsValid Then
            MessageBox.Show(_recoveryDefinitionBinder.ErrorInfoError, "Save Product Recovery Scenario", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return success
        End If
        Try
            Me.Cursor = Cursors.WaitCursor
            _parentForm.SaveRcvryScenario()
            success = _recoveryDefinitionBinder.Save()
            Dim holdDefinitionName As String = _recoveryDefinitionBinder.ScenarioName
            MessageBox.Show(String.Format("Product Recovery Scenario {0} saved.", holdDefinitionName), "Save Product Recovery Scenario", MessageBoxButtons.OK, MessageBoxIcon.Information)
            GetProductRecoveryScenarios()
            Dim rowIndex As Integer = grdScenariosView.LocateByValue("RCVRY_SCENARIO_NAME", holdDefinitionName)
            grdScenariosView.FocusedRowHandle = grdScenariosView.GetRowHandle(rowIndex)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Save Product Recovery Scenario", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
        Return success
    End Function

    Private Sub ProductRecoveryBindingSource_CurrentItemChanged(sender As Object, e As System.EventArgs) Handles ProductRecoveryBindingSource.CurrentItemChanged ', InsolAdjustmentBindingSource.CurrentItemChanged
        If ProductRecoveryBindingSource.List.Count = 0 Then Exit Sub
        btnSaveAs.Enabled = pnlDetails.Enabled AndAlso _recoveryDefinitionBinder.IsReadOnly AndAlso Not _recoveryDefinitionBinder.OriginalScenarioName.ToLower.Equals(_recoveryDefinitionBinder.ScenarioName.ToLower)
    End Sub

    Private Sub ProductRecoveryBindingSource_DataSourceChanged(sender As Object, e As System.EventArgs) Handles ProductRecoveryBindingSource.DataSourceChanged
        pnlDetails.Enabled = Not ProductRecoveryBindingSource.DataSource Is Nothing AndAlso ProductRecoveryBindingSource.List.Count > 0
        btnSave.Enabled = pnlDetails.Enabled
        btnDelete.Enabled = False
        btnCancel.Enabled = pnlDetails.Enabled
        btnSaveAs.Enabled = False
        If pnlDetails.Enabled Then
            btnSave.Visible = Not _recoveryDefinitionBinder.IsReadOnly
            btnSaveAs.Visible = Not btnSave.Visible
            btnDelete.Enabled = Not _recoveryDefinitionBinder.IsReadOnly
        End If
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As System.EventArgs) Handles btnDelete.Click
        If MessageBox.Show(String.Format("Are you sure you want to delete {0}?", _recoveryDefinitionBinder.ScenarioName), "Delete Product Recovery Scenario", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
            If _recoveryDefinitionBinder.IsNew Then
                DisplaySelectedProductRecoveryScenario()
                Exit Sub
            End If
            Try
                Me.Cursor = Cursors.WaitCursor
                _parentForm.DeleteRcvryScenario()
                _recoveryDefinitionBinder.Delete()
                GetProductRecoveryScenarios()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Delete Product Recovery Scenario", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Me.Cursor = Cursors.Default
            End Try
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As System.EventArgs) Handles btnCancel.Click
        If _recoveryDefinitionBinder.IsDirty Then
            If MessageBox.Show("Are you sure you want to cancel and lose any changes you made?", "Cancel Product Recovery Scenario", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
            _recoveryDefinitionBinder.Cancel()
            _parentForm.ClearRcvryEtc()
        End If
        DisplaySelectedProductRecoveryScenario()
    End Sub

    Public ReadOnly Property ProductRecoveryDefinition As ProductRecoveryDefinition
        Get
            Return _recoveryDefinitionBinder
        End Get
    End Property

    Private Sub grdScenariosView_RowClick(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs) Handles grdScenariosView.RowClick
        If Not _recoveryDefinitionBinder Is Nothing Then
            If Not CheckForChanges() Then
                Exit Sub
            End If
            DisplaySelectedProductRecoveryScenario()
        End If
        pnlScenarios.HidePopup()
    End Sub

    Private Sub rdoCalcMode_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rdoMeasuredLab.CheckedChanged, rdoCalculated.CheckedChanged, rdoLimitRules.CheckedChanged
        'manually setting the property values because binding to radio button is a little tricky when it is within a panel
        Select Case True
            Case rdoCalculated.Checked
                _recoveryDefinitionBinder.DensityCalculatedCalculationMode = True
                _recoveryDefinitionBinder.DensityLimitRuleCalculationMode = False
                _recoveryDefinitionBinder.DensityMeasuredLabCalculationMode = False
                pnlUseLimitRules.Enabled = False
                pnlUseMeasuredLabValue.Enabled = False
            Case rdoLimitRules.Checked
                _recoveryDefinitionBinder.DensityLimitRuleCalculationMode = True
                _recoveryDefinitionBinder.DensityCalculatedCalculationMode = False
                _recoveryDefinitionBinder.DensityMeasuredLabCalculationMode = False
                pnlUseMeasuredLabValue.Enabled = False
                pnlUseLimitRules.Enabled = True
            Case rdoMeasuredLab.Checked
                _recoveryDefinitionBinder.DensityMeasuredLabCalculationMode = True
                _recoveryDefinitionBinder.DensityCalculatedCalculationMode = False
                _recoveryDefinitionBinder.DensityLimitRuleCalculationMode = False
                pnlUseMeasuredLabValue.Enabled = True
                pnlUseLimitRules.Enabled = False
            Case Else
                pnlUseLimitRules.Enabled = False
                pnlUseMeasuredLabValue.Enabled = False
        End Select
    End Sub

#Region "Full Focus on Text Boxes"
    Private _enteringText As Boolean
    Private _textNeedSelect As Boolean
    Private Sub ResetEnterFlag()
        _enteringText = False
    End Sub
    Private Sub txtEdit_Enter(sender As Object, e As System.EventArgs) Handles _
                                                                                txtUpperLimit.Enter,
                                                                               txtLowerLimit.Enter, txtUpperZoneCorrection.Enter, txtLowerZoneCorrection.Enter,
                                                                               txtMaxClPctSpl.Enter, txtMaxClPctHole.Enter,
                                                                               txtMaxCaP2O5Spl.Enter, txtMaxCaP2O5Hole.Enter, txtMaxOvbThk.Enter,
                                                                               txtMinMtxThkHole.Enter, txtMinItbThkHole.Enter, txtMaxTotDepthHole.Enter,
                                                                               txtMaxMtxXSpl.Enter, txtMaxMtxXHole.Enter, txtMaxTotXHole.Enter,
                                                                               txtMinTotPrTpaHole.Enter, txtMaxClayTotPrSpl.Enter, txtMaxClayTotPrHole.Enter,
                                                                               txtMaxInbLowerZoneHole.Enter

        _enteringText = True
        BeginInvoke(New MethodInvoker(AddressOf ResetEnterFlag))
    End Sub

    Private Sub txtEdit_GotFocus(sender As Object, e As System.EventArgs) Handles _
                                                                                   txtUpperLimit.GotFocus,
                                                                                  txtLowerLimit.GotFocus, txtUpperZoneCorrection.GotFocus, txtLowerZoneCorrection.GotFocus,
                                                                                  txtMaxClPctSpl.GotFocus, txtMaxClPctHole.GotFocus,
                                                                                  txtMaxCaP2O5Spl.GotFocus, txtMaxCaP2O5Hole.GotFocus, txtMaxOvbThk.GotFocus,
                                                                                  txtMinMtxThkHole.GotFocus, txtMinItbThkHole.GotFocus, txtMaxTotDepthHole.GotFocus,
                                                                                  txtMaxMtxXSpl.GotFocus, txtMaxMtxXHole.GotFocus, txtMaxTotXHole.GotFocus,
                                                                                  txtMinTotPrTpaHole.GotFocus, txtMaxClayTotPrSpl.GotFocus, txtMaxClayTotPrHole.GotFocus,
                                                                                  txtMaxInbLowerZoneHole.GotFocus
        With DirectCast(sender, TextEdit)
            .SelectAll()
        End With
        EndEditOnAllBindingSources()
    End Sub

    Private Sub txtEdit_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles _
                                                                                                     txtUpperLimit.MouseUp,
                                                                                                    txtLowerLimit.MouseUp, txtUpperZoneCorrection.MouseUp, txtLowerZoneCorrection.MouseUp,
                                                                                                    txtMaxClPctSpl.MouseUp, txtMaxClPctHole.MouseUp,
                                                                                                    txtMaxCaP2O5Spl.MouseUp, txtMaxCaP2O5Hole.MouseUp, txtMaxOvbThk.MouseUp,
                                                                                                    txtMinMtxThkHole.MouseUp, txtMinItbThkHole.MouseUp, txtMaxTotDepthHole.MouseUp,
                                                                                                    txtMaxMtxXSpl.MouseUp, txtMaxMtxXHole.MouseUp, txtMaxTotXHole.MouseUp,
                                                                                                    txtMinTotPrTpaHole.MouseUp, txtMaxClayTotPrSpl.MouseUp, txtMaxClayTotPrHole.MouseUp,
                                                                                                    txtMaxInbLowerZoneHole.MouseUp
        If _textNeedSelect Then DirectCast(sender, TextEdit).SelectAll()
    End Sub

    Private Sub txtEdit_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles txtUpperLimit.MouseDown,
                                                                                                      txtLowerLimit.MouseDown, txtUpperZoneCorrection.MouseDown,
                                                                                                      txtLowerZoneCorrection.MouseDown,
                                                                                                      txtMaxClPctSpl.MouseDown, txtMaxClPctHole.MouseDown,
                                                                                                      txtMaxCaP2O5Spl.MouseDown, txtMaxCaP2O5Hole.MouseDown, txtMaxOvbThk.MouseDown,
                                                                                                      txtMinMtxThkHole.MouseDown, txtMinItbThkHole.MouseDown, txtMaxTotDepthHole.MouseDown,
                                                                                                      txtMaxMtxXSpl.MouseDown, txtMaxMtxXHole.MouseDown, txtMaxTotXHole.MouseDown,
                                                                                                      txtMinTotPrTpaHole.MouseDown, txtMaxClayTotPrSpl.MouseDown, txtMaxClayTotPrHole.MouseDown,
                                                                                                      txtMaxInbLowerZoneHole.MouseDown
        _textNeedSelect = _enteringText
    End Sub

    Private Sub txtEdit_Enter2(sender As Object, e As System.EventArgs) Handles txtIpInsAdj.Enter, txtIPInsAdj100.Enter, txtOvbVolRcvryCf.Enter, txtMineVolRcvryCf.Enter,
                                                                                txtTotPbInsAdj.Enter, txtTotPbInsAdj100.Enter, txtTotCnInsAdj.Enter, txtTotCnInsAdj100.Enter,
                                                                                txtTotFdBplRcvry.Enter, txtTotPbTonRcvry.Enter, txtIpTonRcvryTot.Enter, txtTotFdTonRcvry.Enter,
                                                                                txtClTonRcvryTot.Enter, txtFlotRcvryHw.Enter, txtFlotRcvryFdBplExp.Enter, txtFlotRcvryTlBplHw.Enter,
                                                                                txtFlotRcvryFdBplExp100.Enter



        _enteringText = True
        BeginInvoke(New MethodInvoker(AddressOf ResetEnterFlag))
    End Sub

    Private Sub txtEdit_GotFocus2(sender As Object, e As System.EventArgs) Handles txtIpInsAdj.GotFocus, txtIPInsAdj100.GotFocus, txtOvbVolRcvryCf.GotFocus, txtMineVolRcvryCf.GotFocus,
                                                                                    txtTotPbInsAdj.GotFocus, txtTotPbInsAdj100.GotFocus, txtTotCnInsAdj.GotFocus, txtTotCnInsAdj100.GotFocus,
                                                                                    txtTotFdBplRcvry.GotFocus, txtTotPbTonRcvry.GotFocus, txtIpTonRcvryTot.GotFocus, txtTotFdTonRcvry.GotFocus,
                                                                                    txtClTonRcvryTot.GotFocus, txtFlotRcvryHw.GotFocus, txtFlotRcvryFdBplExp.GotFocus, txtFlotRcvryTlBplHw.GotFocus,
                                                                                    txtFlotRcvryFdBplExp100.GotFocus
        With DirectCast(sender, TextEdit)
            .SelectAll()
        End With
        EndEditOnAllBindingSources()
    End Sub

    Private Sub txtEdit_MouseUp2(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles txtIpInsAdj.MouseUp, txtIPInsAdj100.MouseUp, txtOvbVolRcvryCf.MouseUp, txtMineVolRcvryCf.MouseUp,
                                                                                                    txtTotPbInsAdj.MouseUp, txtTotPbInsAdj100.MouseUp, txtTotCnInsAdj.MouseUp, txtTotCnInsAdj100.MouseUp,
                                                                                                    txtTotFdBplRcvry.MouseUp, txtTotPbTonRcvry.MouseUp, txtIpTonRcvryTot.MouseUp, txtTotFdTonRcvry.MouseUp,
                                                                                                    txtClTonRcvryTot.MouseUp, txtFlotRcvryHw.MouseUp, txtFlotRcvryFdBplExp.MouseUp, txtFlotRcvryTlBplHw.MouseUp,
                                                                                                    txtFlotRcvryFdBplExp100.MouseUp

        If _textNeedSelect Then DirectCast(sender, TextEdit).SelectAll()
    End Sub

    Private Sub txtEdit_MouseDown2(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles txtIpInsAdj.MouseDown, txtIPInsAdj100.MouseDown, txtOvbVolRcvryCf.MouseDown, txtMineVolRcvryCf.MouseDown,
                                                                                                        txtTotPbInsAdj.MouseDown, txtTotPbInsAdj100.MouseDown, txtTotCnInsAdj.MouseDown, txtTotCnInsAdj100.MouseDown,
                                                                                                        txtTotFdBplRcvry.MouseDown, txtTotPbTonRcvry.MouseDown, txtIpTonRcvryTot.MouseDown, txtTotFdTonRcvry.MouseDown,
                                                                                                        txtClTonRcvryTot.MouseDown, txtFlotRcvryHw.MouseDown, txtFlotRcvryFdBplExp.MouseDown, txtFlotRcvryTlBplHw.MouseDown,
                                                                                                        txtFlotRcvryFdBplExp100.MouseDown
        _textNeedSelect = _enteringText
    End Sub

#End Region 'Full Focus on Text Boxes

#Region "Null value on Text Boxes"
    Sub TextEditNulableDecimal_KeyDown(sender As Object, e As KeyEventArgs) Handles txtMaxClPctSpl.KeyDown, txtMaxClPctHole.KeyDown,
                                                                                    txtMaxCaP2O5Spl.KeyDown, txtMaxCaP2O5Hole.KeyDown,
                                                                                    txtMaxOvbThk.KeyDown, txtMinMtxThkHole.KeyDown, txtMinItbThkHole.KeyDown,
                                                                                    txtMaxTotDepthHole.KeyDown,
                                                                                    txtMaxMtxXSpl.KeyDown, txtMaxMtxXHole.KeyDown, txtMaxTotXHole.KeyDown, txtMinTotPrTpaHole.KeyDown,
                                                                                    txtMaxClayTotPrSpl.KeyDown, txtMaxClayTotPrHole.KeyDown,
                                                                                    txtMaxInbLowerZoneHole.KeyDown

        Dim ThisTextEdit As TextEdit = CType(sender, TextEdit)
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            If ThisTextEdit.Text.Length = 1 Or ThisTextEdit.SelectionLength = ThisTextEdit.Text.Length Then
                ThisTextEdit.Text = Nothing
            End If
        ElseIf e.KeyCode = Keys.Up OrElse e.KeyCode = Keys.Down Then
            e.Handled = True
        End If
    End Sub


    Sub TextEditNulableDecimal_KeyDown2(sender As Object, e As KeyEventArgs) Handles txtIpInsAdj.KeyDown, txtIPInsAdj100.KeyDown, txtOvbVolRcvryCf.KeyDown, txtMineVolRcvryCf.KeyDown,
                                                                                    txtTotPbInsAdj.KeyDown, txtTotPbInsAdj100.KeyDown, txtTotCnInsAdj.KeyDown, txtTotCnInsAdj100.KeyDown,
                                                                                    txtTotFdBplRcvry.KeyDown, txtTotPbTonRcvry.KeyDown, txtIpTonRcvryTot.KeyDown, txtTotFdTonRcvry.KeyDown,
                                                                                    txtClTonRcvryTot.KeyDown, txtFlotRcvryHw.KeyDown, txtFlotRcvryFdBplExp.KeyDown, txtFlotRcvryTlBplHw.KeyDown,
                                                                                    txtFlotRcvryFdBplExp100.KeyDown
        Dim ThisTextEdit As TextEdit = CType(sender, TextEdit)
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = Keys.Back Then
            If ThisTextEdit.Text.Length = 1 Or ThisTextEdit.SelectionLength = ThisTextEdit.Text.Length Then
                ThisTextEdit.Text = Nothing
            End If
        ElseIf e.KeyCode = Keys.Up OrElse e.KeyCode = Keys.Down Then
            e.Handled = True
        End If
    End Sub

#End Region 'Null value on Text Boxes

    Private Sub txtSplitPebbleRejectNumericEditor_EditValueChanged(sender As Object, e As System.EventArgs) Handles txtSplitPebbleRejectNumericEditor.EditValueChanged
        grdPebbleRejectView.PostEditor()
    End Sub

    Private Sub _recoveryDefinitionBinder_PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Handles _recoveryDefinitionBinder.PropertyChanged
        pnlCpbFpbCheckedMessage.Visible = _recoveryDefinitionBinder.PebbleIsCheckedStatusVisible
        If Not _loadingData Then
            grdPebbleRejectView.RefreshData()
            grdSplitQualityView.RefreshData()
            grdHoleQualityView.RefreshData()
        End If
    End Sub

    Private Sub grdPebbleRejectView_CustomDrawCell(sender As Object, e As DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs) Handles grdPebbleRejectView.CustomDrawCell
        Select Case e.Column.VisibleIndex
            Case 2
                e.Appearance.BackColor = Color.LightYellow
                If pnlCoarsePebbleMessage.ForeColor = Color.Red Then e.Appearance.BackColor = Color.LightPink
                If pnlCoarsePebbleMessage.ForeColor = Color.Green Then e.Appearance.BackColor = Color.LightGreen
            Case 3
                e.Appearance.BackColor = Color.LightYellow
                If pnlFinePebbleMessage.ForeColor = Color.Red Then e.Appearance.BackColor = Color.LightPink
                If pnlFinePebbleMessage.ForeColor = Color.Green Then e.Appearance.BackColor = Color.LightGreen
            Case 4
                e.Appearance.BackColor = Color.LightYellow
                If pnlIPMessage.ForeColor = Color.Red Then e.Appearance.BackColor = Color.LightPink
                If pnlIPMessage.ForeColor = Color.Green Then e.Appearance.BackColor = Color.LightGreen
            Case 5
                e.Appearance.BackColor = Color.LightYellow
                If pnlTotalPebbleMessage.ForeColor = Color.Red Then e.Appearance.BackColor = Color.LightPink
                If pnlTotalPebbleMessage.ForeColor = Color.Green Then e.Appearance.BackColor = Color.LightGreen
        End Select
    End Sub
    '118
    Private Sub grdPebbleRejectView_ShowingEditor(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles grdPebbleRejectView.ShowingEditor
        Select Case grdPebbleRejectView.FocusedColumn.VisibleIndex
            Case 2
                If Not _recoveryDefinitionBinder.EnterCoarsePebbleRejectValue Then e.Cancel = True
            Case 3
                If Not _recoveryDefinitionBinder.EnterFinePebbleRejectValue Then e.Cancel = True
            Case 4
                If Not _recoveryDefinitionBinder.EnterIPRejectValue Then e.Cancel = True
            Case 5
                If Not _recoveryDefinitionBinder.EnterTotalPebbleRejectValue Then e.Cancel = True
        End Select
    End Sub

    Private Sub chkCoarsePebbleReject_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkCoarsePebbleReject.CheckedChanged
        pnlCoarsePebbleMessage.Visible = chkCoarsePebbleReject.Checked
    End Sub

    Private Sub chkFinePebbleReject_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkFinePebbleReject.CheckedChanged
        pnlFinePebbleMessage.Visible = chkFinePebbleReject.Checked
    End Sub

    Private Sub chkIPReject_CheckedChanged(sender As Object, e As EventArgs) Handles chkIPReject.CheckedChanged
        pnlIPMessage.Visible = chkIPReject.Checked
        'EndEditOnAllBindingSources()
    End Sub

    Private Sub chkTotalPebbleReject_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkTotalPebbleReject.CheckedChanged
        pnlTotalPebbleMessage.Visible = chkTotalPebbleReject.Checked
    End Sub

    Private Sub grdQualityView_ShowingEditor(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles grdSplitQualityView.ShowingEditor, grdHoleQualityView.ShowingEditor
        Dim grdView As GridView = DirectCast(sender, GridView)
        Dim data As ProductQualitySpecification = grdView.GetRow(grdSplitQualityView.FocusedRowHandle)
        If grdView.FocusedColumn.FieldName = "CoarsePebbleValue" Then
            If _recoveryDefinitionBinder.IsProdQualReadOnly(data.Element, "Coarse pebble") Then e.Cancel = True
        End If
        If grdView.FocusedColumn.FieldName = "FinePebbleValue" Then
            If _recoveryDefinitionBinder.IsProdQualReadOnly(data.Element, "Fine pebble") Then e.Cancel = True
        End If
        If grdView.FocusedColumn.FieldName = "IpValue" Then
            If _recoveryDefinitionBinder.IsProdQualReadOnly(data.Element, "IP") Then e.Cancel = True
        End If
    End Sub

    Private Sub rdb_CheckedChanged(sender As Object, e As EventArgs) Handles rdbIsFlotRcvryMode100ExponentFeedBPL.CheckedChanged, rdbIsFlotRcvryMode100ZeroTailBPL.CheckedChanged,
                                                                             rdbIsFlotRcvryModeExponentFeedBPL.CheckedChanged, rdbIsFlotRcvryModeHardwire.CheckedChanged,
                                                                             rdbIsFlotRcvryModeLabFlotation.CheckedChanged, rdbIsFlotRcvryModeLinearModel.CheckedChanged, rdbIsTotPbInsAdjMinimum.CheckedChanged, rdbIsTotPbInsAdjMetLab.CheckedChanged, rdbIsTotPbInsAdjHardwire.CheckedChanged, rdbIsTotPbInsAdj100Minimum.CheckedChanged, rdbIsTotPbInsAdj100MetLab.CheckedChanged, rdbIsTotPbInsAdj100Hardwire.CheckedChanged, rdbIsTotCnInsAdjMinimum.CheckedChanged, rdbIsTotCnInsAdjMetLab.CheckedChanged, rdbIsTotCnInsAdjHardwire.CheckedChanged, rdbIsTotCnInsAdj100Minimum.CheckedChanged, rdbIsTotCnInsAdj100MetLab.CheckedChanged, rdbIsTotCnInsAdj100Hardwire.CheckedChanged, rdbIsIPInsAdjMinimum.CheckedChanged, rdbIsIPInsAdjMetLab.CheckedChanged, rdbIsIPInsAdjHardwire.CheckedChanged, rdbIsIPInsAdj100Minimum.CheckedChanged, rdbIsIPInsAdj100MetLab.CheckedChanged, rdbIsIPInsAdj100Hardwire.CheckedChanged
        If CType(sender, RadioButton).Checked Then
            For Each b As Binding In CType(sender, RadioButton).DataBindings
                b.WriteValue()
            Next
        End If
    End Sub

    Private Sub chk_CheckedChanged(sender As Object, e As EventArgs) Handles chkAbsoluteStop.CheckedChanged, chkFinishSplit.CheckedChanged
        If CType(sender, CheckBox).Checked Then
            For Each b As Binding In CType(sender, CheckBox).DataBindings
                b.WriteValue()
            Next
        End If
    End Sub

    Private Sub btnSetDefaultQualCoef_Click(sender As Object, e As EventArgs) Handles btnSetDefaultQualCoef.Click
        For Each ProductQualCoff As ProductQualitySpecification In _recoveryDefinitionBinder.ProductQualityCoefficients
            ProductQualCoff.PebbleValue = 100
            ProductQualCoff.IpValue = 100
            ProductQualCoff.ConcentrateValue = 100
        Next
    End Sub

    Private Sub txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtUpperLimit.KeyDown, txtLowerLimit.KeyDown, txtUpperZoneCorrection.KeyDown, txtLowerZoneCorrection.KeyDown
        If e.KeyCode = Keys.Up OrElse e.KeyCode = Keys.Down Then
            e.Handled = True
        End If

    End Sub

    Private Sub pnlScenarios_ButtonChecked(sender As Object, e As FlyoutPanelButtonCheckedEventArgs) Handles pnlScenarios.ButtonChecked

    End Sub
End Class
