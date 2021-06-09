Imports OracleInProcServer
Imports System.Math
Imports AxFPSpread
Imports Microsoft.VisualBasic
Imports System.Drawing.Printing
Imports System.Windows.Forms.AxHost
Imports System.IO

Public Class frmProspDataReduction

    Private _areaDefinitionForm As ctrAreaDefinition
    Private _productSizeDesginationForm As ctrProductSizeDesignation
    Private _recoveryScenariosForm As ctrRecoveryAndMineabilityParameters
    Private _loadingRecoveryScenarioData As Boolean = False
    Private _ATprProductResults As ctrProductResult

    Dim fProspCodeDynaset As OraDynaset
    Dim fProspDatasetTfileLoc As String
    'Dim fRawProspDynaset As OraDynaset
    Dim fProcessing As Boolean
    Private stringToPrint As String = String.Empty

    Private Sub Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.Cursor = Cursors.WaitCursor

            _areaDefinitionForm = New ctrAreaDefinition
            _areaDefinitionForm.Dock = DockStyle.Fill
            tabAreaDef.Controls.Add(_areaDefinitionForm)
            _areaDefinitionForm.Show()

            _productSizeDesginationForm = New ctrProductSizeDesignation
            _productSizeDesginationForm.Dock = DockStyle.Fill
            tabProductSizes.Controls.Add(_productSizeDesginationForm)
            _productSizeDesginationForm.Show()

            _recoveryScenariosForm = New ctrRecoveryAndMineabilityParameters(Me)
            _recoveryScenariosForm.Dock = DockStyle.Fill
            tabRecoveryAndMineability.Controls.Add(_recoveryScenariosForm)
            _recoveryScenariosForm.Show()


            Dim ItemIdx As Integer = 0
            gFormLoad = True
            gShowMainExit("Off")
            gWriteOk = WriteOK()
            fProcessing = False
            ssCompErrors.MaxRows = 0
            gHaveRawProspData = False
            GetAllProspCodes()
            fraDataReduction.Visible = False

            fraReview.Width = 974 '13500

            'Report display frame (multi-split, hole)
            fraReptDisp.Top = fraReview.Top '440
            fraReptDisp.Left = fraReview.Left '180
            fraReptDisp.Width = fraReview.Width
            fraReptDisp.Height = fraReview.Height  '5820
            'fraReptDisp.BorderStyle = 0
            fraReptDisp.Visible = False
            '-----

            lblProspectDatasetStatus.Text = "Adding new prospect dataset"

            cmdGenerateProspectDataset.Enabled = False
            cmdCancelProspectDataset.Enabled = False
            GetAllMineNames()
            fProspDatasetTfileLoc = gGetProspDatasetTfileLoc(gUserName.ToLower)
            FixSsRowHdrs()
            AddMiscVertDividers()


            cmdApplySplOverrides.Visible = False


            MarkProdsInDetlDispSprd()
            lblBarrenSplComm.Text = "Note: Barren splits shown in yellow."


            lblGen64.Text = ""     'Was lblRowNum
            lblGen54.Text = ""
            lblGen55.Text = ""
            cmdHoleSplitRpt.Enabled = False
            'cmdReportAll.Enabled = False
            optProdCoeff.Enabled = True
            opt100Pct.Enabled = True
            chkSaveToDatabase.Enabled = False

            lblTextfileComment.Text = ""

            lblGen53.Text = "The save options are currently limited to:" &
                       vbCrLf & vbCrLf &
                       "1) SurvCADD transfer text files for holes." & vbCrLf &
                       "2) MOIS transfer text files for holes+splits (combined in same file)."

            'raDataTypeOption.BorderStyle = 1
            optProdCoeff.Checked = True
            lblRptAllCnt.Text = ""
            lblRptAllCnt2.Text = ""
            'f raOutputOptions.BorderStyle = 0



            'lblGen65) was lblOffSpecPbMgPlt
            lblGen65.Text = ""



            chkInclMgPlt.Checked = False

            lblNoReview.Visible = False

            'Original MgO plant comment
            lblGen31.Text = "The 'Good' pebble quality (under the PrdQual(3) Tab) and the MgO plant input quality " &
                       "defined here should ideally match up correctly.  The 'Good' pebble 'Min BPL' should " &
                       "be the same as the 'MgO plant input BPL <' value defined here and the 'Good' pebble 'Max MgO' should " &
                       "be the same as the 'MgO plant input MgO >' value defined here." &
                       vbCrLf & vbCrLf &
                       "The oversize will always be excluded regardless of quality."

            'Doloflot comment (Ona)
            lblGen44.Text = "The 'Good' pebble quality (under the PrdQual(3) Tab) and the Doloflot FnePb MgO cutoff quality " &
                       "and the Doloflot IP MgO cutoff quality defined here should ideally match up correctly. " &
                       vbCrLf & vbCrLf &
                       "The OVERSIZE is always thrown out regardless of quality. " & vbCrLf &
                       "Any off-spec COARSE PEBBLE is always thrown out (it is " &
                       "never processed by the Doloflot plant).  " &
                       "Any off-spec FINE PEBBLE is processed by the Doloflot plant if the " &
                       "MgO > Doloflot FnePb MgO cutoff.  " &
                       "Any off-spec IP is processed by the Doloflot plant if the " &
                       "MgO > Doloflot IP MgO cutoff."

            'Doloflot comment (FCO)
            'lblGen51).Text = "Enter the Grind, Acid, P2O5, PA64, Flot minutes, and Target MgO " & _
            '                     "in the Doloflot Plant (Ona) grid to the left."

            'Changes 11/16/2011
            lblGen51.Text = "Uses the 'new' 2011 FCO Doloflot Plant Model (November, 2011)"



            lblGen36.Text = "If a hole is unminable in raw prospect then all splits for that hole will be " &
                       "overridden as unminable."


            chkPbAnalysisFillInSpecial.Text = "Assign weighted average of MgO Plant input + MgO Plant reject analysis to the Coarse, Fine and Total pebble " &
                                       "analyses for holes where there is no Coarse, Fine or Total pebble TPA that was 'OK as is' " &
                                       "without sending the pebble to the MgO Plant."



            lblGen45.Text = "BLUESTAR MgO Model" & vbCrLf & vbCrLf &
                       "Pb BPL = 36.4 + (0.339 * FdBPL) + (0.435 * pa64) + " &
                       "(30.0 * acid) - (14.5 * P2O5) - (0.350 * flotmin)" & vbCrLf & vbCrLf &
                       "BPL %Recovery = 100 * (1 - EXP(-4.56 + (0.0474 * FdMgO) + (0.0226 * grind) - " &
                       "(0.0964 * flotmin) - (5.90 * P2O5) + (12.2 * acid) + (0.121 * pa64)))"




            gFormLoad = False


            gFormProspDataReduction = Me
            AddNewProspectDataset()


            SetActionStatus("Preparing for new prospect dataset...")
            Me.Cursor = Cursors.WaitCursor


            ClearRcvryEtc()
            'ClearSplitOverride()


            lblProspectDatasetStatus.Text = "Adding new prospect dataset"

            'cmdGenerateProspectDataset.Enabled = True
            'cmdCancelProspectDataset.Enabled = True



            '05/14/2007, lss
            'The default at this time will be to create a 100% SurvCADD transfer
            'textfile for holes (composites).

            txtProspectDatasetName.Text = ""
            txtProspectDatasetDesc.Text = ""
            'txtProspDatasetTextfileName.Text = fProspDatasetTfileLoc
            'txtRptAllToTxtFile.Text = fProspDatasetTfileLoc
            chkSaveToDatabase.Checked = False
            chkSurvCaddTextfile.Checked = True
            chkSpecMoisTransferFile.Checked = False
            chkBdFormatTextfile.Checked = False
            chk100Pct.Checked = True
            chkProductionCoefficient.Checked = False
            'optInclComposites.Checked = True

            'Get scenario names that have been saved -- user may then select from
            'them where appropriate.


            'O-ride Tab
            lblGen54.Text = "M = mineable" & vbCrLf &
                             "U = unmineable" & vbCrLf &
                             "C = use calculated value"
            'Review Tab
            lblGen41.Text = ""     'Was lblReviewComm
            lblGen64.Text = ""     'Was lblRowNum


            Me.Cursor = Cursors.Default

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            gFormLoad = False
            Me.Cursor = Cursors.Arrow
        End Try
    End Sub

    Private Function WriteOK()


        'In order to add/modify data on this form gUserPermission.RawProspectReduction must
        'be "Write" or "Setup".

        WriteOK = False

        If gUserPermission.RawProspectReduction = "Write" Or
            gUserPermission.RawProspectReduction = "Setup" Or
            gUserPermission.RawProspectReduction = "Admin" Then
            WriteOK = True
        End If
    End Function

    Private Sub cmdExitForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExitForm.Click


        If MsgBox("Exit raw prospect data reduction program?", vbYesNo +
            vbDefaultButton1, "Exiting Program") = vbYes Then
            Me.Close() ' End 'Unload(Me)
        End If
    End Sub

    Private Sub Form_Unload(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed


        gShowMainExit("On")
        If gRawProspDynaset IsNot Nothing Then
            gRawProspDynaset.Close()
        End If
        On Error Resume Next
        'fRawProspDynaset.Close()
    End Sub

    Private Sub AddNewProspectDataset()
        SetActionStatus("Preparing for new reduction...")
        Me.Cursor = Cursors.WaitCursor

        _areaDefinitionForm.AddNewAreaDefinition()
        _productSizeDesginationForm.AddNewProductSizeDesignation()
        _recoveryScenariosForm.AddNewProductRecoveryScenario()
        ClearSplitOverride()
        ClearDetlDisp()

        cmdGenerateProspectDataset.Enabled = True
        cmdCancelProspectDataset.Enabled = True

        lblProcComm0.Text = ""
        lblProcComm1.Text = ""
        lblProcComm2.Text = ""
        lblGen41.Text = ""
        lblGen64.Text = ""
        lblGen65.Text = ""

        txtProspDatasetTextfileName.Text = fProspDatasetTfileLoc
        txtRptAllToTxtFile.Text = fProspDatasetTfileLoc

        optInclComposites.Checked = True
        cmdHoleSplitRpt.Enabled = False
        optProdCoeff.Enabled = True
        opt100Pct.Enabled = True
        GetSplitOverrideSets()
        fraDataReduction.Visible = True
        'Will display some calculations just in case they are needed at this point.
        DispFdRcvryCalcs()
        chkCreateOutputOnly.Checked = True
        tabMain.SelectedTab = tabAreaDef
        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Sub



    Private Sub ClearSplitOverride()

        txtSplitOverrideName.Text = ""
        txtSplitOverrideName.Enabled = True
        cboSplitOverrideMineName.Text = "None"
        cboSplitOverrideMineName.Enabled = True

        ssSplitReview.MaxRows = 0
        ssCompReview.MaxRows = 0
        ssCompErrors.MaxRows = 0
        ssSplitOverride.MaxRows = 0
        ssSplitOverrides.MaxRows = 0
        ssRawProspMin.MaxRows = 0

        chkOnlyMySplitOverride.Checked = True
        cmdSaveSplitOverride.Enabled = True
        cmdSaveSplitOverride.Enabled = True
        cmdSaveSplitOverride.Enabled = True

        txtSplOverrideTxtFile.Text = fProspDatasetTfileLoc
        lblGen34.Text = ""
        chkUseRawProspAsOverride.Checked = False
    End Sub

    Private Sub cmdPrtScr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrtScr.Click


        Dim UserResponse As Object

        UserResponse = MsgBox("Print the screen?", vbOKCancel, "Printing")

        If UserResponse = vbOK Then
            'On Error GoTo PrintError

            SetActionStatus("Printing the screen...")
            Me.Cursor = Cursors.WaitCursor
            Try
                gPrintScreen(Me.Handle)
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try


            SetActionStatus("")
            Me.Cursor = Cursors.Arrow
        End If

        Exit Sub


    End Sub


    Private Sub ssSplitOverride_Click(ByVal sender As System.Object, ByVal e As AxFPSpread._DSpreadEvents_ClickEvent) Handles ssSplitOverride.ClickEvent  '(ByVal Col As Long, ByVal Row As Long)


        If e.col = 3 Then
            gHaveRawProspData = False
        End If
    End Sub


    Private Sub SetActionStatus(ByVal aStatus As String)

        On Error Resume Next
        sbrMain.Text = aStatus
    End Sub

    Private Sub cmdCancelProspectDataset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancelProspectDataset.Click
        If MsgBox("Are you really sure you want to cancel out of this reduction?", vbYesNo +
                  vbDefaultButton1, "Cancel Reduction") = vbYes Then
            AddNewProspectDataset()
        End If
    End Sub


    Private Sub GetAllMineNames()

        On Error GoTo GetAllMineNamesError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim MineNameDynaset As OraDynaset

        SetActionStatus("Loading mine names...")
        Me.Cursor = Cursors.WaitCursor

        'Get all existing mine names
        'Set 
        params = gDBParams

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_mine_info(:pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MineNameDynaset = params("pResult").Value
        ClearParams(params)

        cboSplitOverrideMineName.Items.Add("None")

        MineNameDynaset.MoveFirst()

        Do While Not MineNameDynaset.EOF
            'Only want mines with prospect data!
            If MineNameDynaset.Fields("mine_prospect").Value = 1 Then
                cboSplitOverrideMineName.Items.Add(MineNameDynaset.Fields("mine_name").Value)
            End If
            MineNameDynaset.MoveNext()
        Loop
        cboSplitOverrideMineName.Text = "None"

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow

        Exit Sub

GetAllMineNamesError:
        MsgBox("Error getting all mine names." & vbCrLf &
        Err.Description,
        vbOKOnly + vbExclamation,
        "Mine Names Access Error")

        On Error Resume Next
        ClearParams(params)
        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub GetAllProspCodes()

        On Error GoTo GetProspCodesError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        SetActionStatus("Loading prospect codes...")
        Me.Cursor = Cursors.WaitCursor

        'Get all existing prospect codes
        ' Set 
        params = gDBParams

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_all_prosp_codes
        'pResult              IN OUT c_prospcodes)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospectnew.get_all_prosp_codes(:pResult);end;", ORASQL_FAILEXEC)
        ' Set 
        fProspCodeDynaset = params("pResult").Value
        ClearParams(params)

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow

        Exit Sub

GetProspCodesError:
        MsgBox("Error getting all prospect codes." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Prospect Code Get Error")

        On Error Resume Next
        ClearParams(params)
        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub FixSsRowHdrs()

        'With ssVolRcvry
        '    .BlockMode = True
        '    .Row = 1
        '    .Row2 = .MaxRows
        '    .Col = 0
        '    .Col2 = 0
        '    .TypeTextWordWrap = False
        '    .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft ' SS_CELL_H_ALIGN_LEFT
        '    .BlockMode = False
        'End With

        'With ssProdRcvryFctrs
        '    .BlockMode = True
        '    .Row = 1
        '    .Row2 = .MaxRows
        '    .Col = 0
        '    .Col2 = 0
        '    .TypeTextWordWrap = False
        '    .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft '  SS_CELL_H_ALIGN_LEFT
        '    .BlockMode = False
        'End With

        'With ssFlotRcvryLinear
        '    .BlockMode = True
        '    .Row = 1
        '    .Row2 = .MaxRows
        '    .Col = 0
        '    .Col2 = 0
        '    .TypeTextWordWrap = False
        '    .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft '  SS_CELL_H_ALIGN_LEFT
        '    .BlockMode = False
        'End With

        'With ssInsAdj100Pct
        '    .BlockMode = True
        '    .Row = 1
        '    .Row2 = .MaxRows
        '    .Col = 0
        '    .Col2 = 0
        '    .TypeTextWordWrap = False
        '    .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft '  SS_CELL_H_ALIGN_LEFT
        '    .BlockMode = False
        'End With

        'With ssSplitPhysMineability
        '    .BlockMode = True
        '    .Row = 1
        '    .Row2 = .MaxRows
        '    .Col = 0
        '    .Col2 = 0
        '    .TypeTextWordWrap = False
        '    .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft '  SS_CELL_H_ALIGN_LEFT
        '    .BlockMode = False
        'End With

        'With ssHolePhysMineability
        '    .BlockMode = True
        '    .Row = 1
        '    .Row2 = .MaxRows
        '    .Col = 0
        '    .Col2 = 0
        '    .TypeTextWordWrap = False
        '    .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft '  SS_CELL_H_ALIGN_LEFT
        '    .BlockMode = False
        'End With

        'With ssSplitEconMineability
        '    .BlockMode = True
        '    .Row = 1
        '    .Row2 = .MaxRows
        '    .Col = 0
        '    .Col2 = 0
        '    .TypeTextWordWrap = False
        '    .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft '  SS_CELL_H_ALIGN_LEFT
        '    .BlockMode = False
        'End With

        'With ssHoleEconMineability
        '    .BlockMode = True
        '    .Row = 1
        '    .Row2 = .MaxRows
        '    .Col = 0
        '    .Col2 = 0
        '    .TypeTextWordWrap = False
        '    .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft '  SS_CELL_H_ALIGN_LEFT
        '    .BlockMode = False
        'End With

        With ssDetlDisp
            .BlockMode = True
            .Row = 1
            .Row2 = .MaxRows
            .Col = 0
            .Col2 = 0
            .TypeTextWordWrap = False
            .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft '  SS_CELL_H_ALIGN_LEFT
            .BlockMode = False
        End With

        With ssResultCnt
            .BlockMode = True
            .Row = 1
            .Row2 = .MaxRows
            .Col = 0
            .Col2 = 0
            .TypeTextWordWrap = False
            .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft ' SS_CELL_H_ALIGN_LEFT
            .BlockMode = False
        End With

        With ssOffSpecPb
            .BlockMode = True
            .Row = 1
            .Row2 = .MaxRows
            .Col = 0
            .Col2 = 0
            .TypeTextWordWrap = False
            .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft '  SS_CELL_H_ALIGN_LEFT
            .BlockMode = False
        End With

        With ssDoloflotPlant
            .BlockMode = True
            .Row = 1
            .Row2 = .MaxRows
            .Col = 0
            .Col2 = 0
            .TypeTextWordWrap = False
            .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft '  SS_CELL_H_ALIGN_LEFT
            .BlockMode = False
        End With

        With ssDoloflotPlantFco
            .BlockMode = True
            .Row = 1
            .Row2 = .MaxRows
            .Col = 0
            .Col2 = 0
            .TypeTextWordWrap = False
            .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft ' SS_CELL_H_ALIGN_LEFT
            .BlockMode = False
        End With

        With ssDoloflotPlantFco2
            .BlockMode = True
            .Row = 1
            .Row2 = .MaxRows
            .Col = 0
            .Col2 = 0
            .TypeTextWordWrap = False
            .TypeHAlign = FPSpread.TypeHAlignConstants.TypeHAlignLeft ' SS_CELL_H_ALIGN_LEFT
            .BlockMode = False
        End With

    End Sub

    Public Sub DeleteRcvryScenario()
        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Try
            SetActionStatus("Deleting recovery scenario...")
            params = gDBParams
            params.Add("pRcvryScenarioName", _recoveryScenariosForm.ProductRecoveryDefinition.ScenarioName, ORAPARM_INPUT)
            params("pRcvryScenarioName").serverType = ORATYPE_VARCHAR2
            params.Add("pProspSetName", _recoveryScenariosForm.ProductRecoveryDefinition.ProspectSetName, ORAPARM_INPUT)
            params("pProspSetName").serverType = ORATYPE_VARCHAR2
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prosp_data_rdctn.delete_rcvry_scen(" +
                      ":pRcvryScenarioName, :pProspSetName);end;", ORASQL_FAILEXEC)
        Catch ex As Exception
            Throw ex
        Finally
            ClearParams(params)
            SetActionStatus("")
        End Try
    End Sub

    Public Sub ClearRcvryEtc()
        'With ssVolRcvry
        '    .Col = 1    'Check boxes
        '    .Row = 1
        '    .Value = 0
        '    .Col = 1
        '    .Row = 2
        '    .Value = 1  '1

        '    .Col = 2    'Constant Factor
        '    .Row = 1
        '    .Value = 0
        '    .Col = 2
        '    .Row = 2
        '    .Value = 0  '110

        '    .Col = 3    'Variable Factor
        '    .Row = 1
        '    .Value = 0
        '    .Col = 3
        '    .Row = 2
        '    .Value = 0

        '    .Col = 5    'Check boxes
        '    .Row = 1
        '    .Value = 1  '1
        '    .Col = 5
        '    .Row = 2
        '    .Value = 0

        '    .Col = 6    'Footage adjustment
        '    .Row = 1
        '    .Value = 0  '-3
        '    .Col = 6
        '    .Row = 2
        '    .Value = 0
        'End With

        'With ssAdjAssocTons
        '    .Col = 1    'OS
        '    .Row = 1
        '    .Value = 0
        '    .Col = 2    'Pb
        '    .Value = 0
        '    .Col = 3    'IP
        '    .Value = 0
        '    .Col = 4    'Fd
        '    .Value = 0
        '    .Col = 5    'Cl
        '    .Value = 0
        'End With

        'With ssProdRcvryFctrs
        '    .Col = 3
        '    .Row = 1    'CrsPb ton rcvry
        '    .Value = 0  '90
        '    .Row = 2    'IP ton rcvry
        '    .Value = 0  '0
        '    .Row = 3    'CrsFd ton rcvry
        '    .Value = 0  '112
        '    .Row = 4    'CrsFd BPL rcvry
        '    .Value = 0  '72
        '    .Row = 5    'Clay ton rcvry
        '    .Value = 0

        '    .Col = 5
        '    .Row = 1    'FnePb ton rcvry
        '    .Value = 0  '90
        '    .Row = 3    'FneFd ton rcvry
        '    .Value = 0  '118
        '    .Row = 4    'FneFd BPL rcvry
        '    .Value = 0  '72
        'End With

        'With ssCalcdResults
        '    .Row = 1
        '    .Col = 1    'Crs Conc PL*Ton Recovery
        '    .Value = 0
        '    .Col = 2    'Fne Conc PL*Ton Recovery
        '    .Value = 0
        '    .Col = 4    'Total Fd BPL dilution
        '    .Value = 0
        'End With

        'With ssFlotRcvryLinear
        '    .Row = 1
        '    .Col = 1    'Check box (Linear model)
        '    .Value = 1  '1
        '    .Row = 1
        '    .Col = 5    'Coarse Constant factor (Linear model)
        '    .Value = 0  '82
        '    .Row = 1
        '    .Col = 6    'Coarse Variable factor (Linear model)
        '    .Value = 0
        '    .Row = 2
        '    .Col = 5    'Fine Constant factor (Linear model)
        '    .Value = 0  '82
        '    .Row = 2
        '    .Col = 6    'Fine Variable factor (Linear model)
        '    .Value = 0
        'End With
        'chkTestResultVsLabFlot1.Checked = True

        'With ssFlotRcvryHardwire
        '    .Row = 1
        '    .Col = 0    'Check box  (Hard-wire)
        '    .Value = 0
        '    .Row = 1
        '    .Col = 5    'Coarse Tailings BPL  (Hard-wire)
        '    .Value = 0
        '    .Row = 1
        '    .Col = 6    'Coarse Concentrate Insol  (Hard-wire)
        '    .Value = 0
        '    .Row = 2
        '    .Col = 5    'Fine Tailings BPL  (Hard-wire)
        '    .Value = 0
        '    .Row = 2
        '    .Col = 6    'Fine Concentrate Insol  (Hard-wire)
        '    .Value = 0
        'End With
        'chkTestResultVsLabFlot2.Checked = False

        'With ssOtherFlotMethods
        '    .Row = 1
        '    .Col = 1    'Check -- Use lab flotation recovery
        '    .Value = 0
        '    .Row = 2
        '    .Col = 1    'Check -- Use SqrRt(Feed BPL) = Tail BPL
        '    .Value = 0
        'End With

        'With ssInsAdj100Pct
        '    .Row = 1    'Coarse pebble
        '    .Col = 1    'Minimum
        '    .Value = 0  '1
        '    .Row = 1
        '    .Col = 2    'Direct
        '    .Value = 0
        '    .Row = 1
        '    .Col = 3    'Incremental
        '    .Value = 0
        '    .Row = 1
        '    .Col = 4    'In-Situ
        '    .Value = 0
        '    .Row = 1
        '    .Col = 6    'Adjustment value
        '    .Value = 0  '11

        '    .Row = 2    'Fine pebble
        '    .Col = 1    'Minimum
        '    .Value = 0  '1
        '    .Row = 2
        '    .Col = 2    'Direct
        '    .Value = 0
        '    .Row = 2
        '    .Col = 3    'Incremental
        '    .Value = 0
        '    .Row = 2
        '    .Col = 4    'In-Situ
        '    .Value = 0
        '    .Row = 2
        '    .Col = 6    'Adjustment value
        '    .Value = 0  '11

        '    .Row = 3    'IP
        '    .Col = 1    'Minimum
        '    .Value = 0
        '    .Row = 3
        '    .Col = 2    'Direct
        '    .Value = 0
        '    .Row = 3
        '    .Col = 3    'Incremental
        '    .Value = 0
        '    .Row = 3
        '    .Col = 4    'In-Situ
        '    .Value = 0  '1
        '    .Row = 3
        '    .Col = 6    'Adjustment value
        '    .Value = 0

        '    .Row = 4    'Coarse concentrate
        '    .Col = 1    'Minimum
        '    .Value = 0
        '    .Row = 4
        '    .Col = 2    'Direct
        '    .Value = 0  '1
        '    .Row = 4
        '    .Col = 3    'Incremental
        '    .Value = 0
        '    .Row = 4
        '    .Col = 4    'In-Situ
        '    .Value = 0
        '    .Row = 4
        '    .Col = 6    'Adjustment value
        '    .Value = 0  '7.5

        '    .Row = 5    'Fine Concentrate
        '    .Col = 1    'Minimum
        '    .Value = 0
        '    .Row = 5
        '    .Col = 2    'Direct
        '    .Value = 0  '1
        '    .Row = 5
        '    .Col = 3    'Incremental
        '    .Value = 0
        '    .Row = 5
        '    .Col = 4    'In-Situ
        '    .Value = 0
        '    .Row = 5
        '    .Col = 6    'Adjustment value
        '    .Value = 0  '7.5
        'End With


        'fraMineability Items
        'fraMineability Items
        'fraMineability Items
        'With ssSplitPhysMineability
        '    .Row = 1
        '    .Col = 1    'Maximum %clay  Split
        '    .Value = 0  '50
        '    .Row = 2
        '    .Col = 1    'Maximum total depth
        '    .Value = 0  '75
        '    .Col = 2    'Absolute stop
        '    .Value = 0
        '    .Col = 3    'Finish split
        '    .Value = 0  '1
        'End With

        'With ssSplitEconMineability
        '    .Row = 1
        '    .Col = 1    'Maximum Mtx-X  Split
        '    .Value = 0  '7.5
        'End With

        'With ssHolePhysMineability
        '    .Row = 1
        '    .Col = 1    'Maximum %clay  Hole
        '    .Value = 0  '50
        '    .Row = 2
        '    .Col = 1    'Minimum ore thickness  Hole
        '    .Value = 0  '3
        '    .Row = 3
        '    .Col = 1    'Minimum interburden thickness
        '    .Value = 0  '3
        'End With

        'With ssHoleEconMineability
        '    .Row = 1
        '    .Col = 1    'Maximum Mtx-X  Hole
        '    .Value = 0  '7.5
        '    .Row = 2
        '    .Col = 1    'Maximum Tot-X  Hole
        '    .Value = 0  '30
        '    .Row = 3
        '    .Col = 1    'Minimum total product TPA
        '    .Value = 0  '2000
        'End With

        'chkNoUnmineableHoles.Checked = True
        'chkInclCpbAlways.Checked = True
        'chkInclFpbAlways.Checked = True
        'chkInclOsAlways.Checked = False

        'chkInclCpbNever.Checked = False
        'chkInclFpbNever.Checked = False
        'chkInclOsNever.Checked = True

        'chkMineHasOffSpecPbPlt.Checked = False
        '06/14/2010, lss  chkMineHasOffSpecPbPlt not really used anymore!
        'Have chkUseDoloflotPlant and chkUseOrigMgOPlant and
        'chkUseDoloflotPlantFco
        chkUseDoloflotPlant.Checked = False      'Ona
        chkUseDoloflotPlantFco.Checked = False  'FCO 1st Run
        chkUseOrigMgoPlant.Checked = False

        'chkCanSelectRejectTpb.Checked = False
        'chkCanSelectRejectTcn.Checked = False

        'fra100PctFlotRcvry Items
        'fra100PctFlotRcvry Items
        'fra100PctFlotRcvry Items
        'opt100PctTlZero.Checked = True

        With ssOffSpecPb
            .Row = 1
            .Col = 1
            .Value = 0
            .Row = 2
            .Value = 0
            .Row = 3
            .Value = 0
            .Row = 4
            .Value = 0
            '-----
            .Row = 5
            .Value = 0
            .Row = 6
            .Value = 0
            .Row = 7
            .Value = 0
            .Row = 8
            .Value = 0
        End With

        'lblGen65) was lblOffSpecPbMgPlt
        lblGen65.Text = ""

        With ssDoloflotPlant
            .Col = 1
            .Row = 1
            .Value = 0
            .Row = 2
            .Value = 0
            .Row = 3
            .Value = 0
            .Row = 4
            .Value = 0
            .Row = 5
            .Value = 0
            .Row = 6
            .Value = 0
            .Row = 7
            .Value = 0
            .Row = 8
            .Value = 0
            .Row = 9
            .Value = 0
            .Row = 10
            .Value = 0
        End With

        With ssDoloflotPlantFco
            .Col = 1
            .Row = 1
            .Value = 0
            .Row = 2
            .Value = 0
        End With

        'New 11/16/2011
        With ssDoloflotPlantFco2
            .Col = 1
            .Row = 1
            .Value = 0
            .Row = 2
            .Value = 0
            .Row = 3
            .Value = 0
            .Row = 4
            .Value = 0
            .Row = 5
            .Value = 0
            .Row = 6
            .Value = 0
            .Row = 7
            .Value = 0
        End With
    End Sub

    Private Sub GetRcvryEtcParamsFromForm(ByRef aRcvryData As gDataRdctnParamsType)

        ' _recoveryScenariosForm.ProductRecoveryDefinition.MineVolRcvryVf


        'fraRecovery Items
        'fraRecovery Items
        'fraRecovery Items

        aRcvryData.OvbVolRcvryMode = "Linear model"
        aRcvryData.MineVolRcvryMode = "Linear model"

        '    '.Col = 2    'Constant Factor
        '    '.Row = 1
        'aRcvryData.OvbVolRcvryCf = _recoveryScenariosForm.ProductRecoveryDefinition.RecoveryDilutionParamaters.OvbVolRcvryCf
        '    '.Col = 2
        '    '.Row = 2
        'aRcvryData.MineVolRcvryCf = _recoveryScenariosForm.ProductRecoveryDefinition.RecoveryDilutionParamaters.MineVolRcvryCf

        '    .Col = 3    'Variable Factor
        '    .Row = 1
        '    aRcvryData.OvbVolRcvryVf = .Value
        '    .Col = 3
        '    .Row = 2
        '    aRcvryData.MineVolRcvryVf = .Value

        '    .Col = 6    'Footage adjustment
        '    .Row = 1
        '    aRcvryData.OvbVolRcvryFa = .Value
        '    .Col = 6
        '    .Row = 2
        '    aRcvryData.MineVolRcvryFa = .Value
        'End With

        'With ssAdjAssocTons
        '    .Col = 1    'OS
        '    .Row = 1
        'aRcvryData.AdjOsTonsWvol = _recoveryScenariosForm.ProductRecoveryDefinition.AdjT
        '    .Col = 2    'Pb
        '    aRcvryData.AdjPbTonsWvol = .Value
        '    .Col = 3    'IP
        '    aRcvryData.AdjIpTonsWvol = .Value
        '    .Col = 4    'Fd
        '    aRcvryData.AdjFdTonsWvol = .Value
        '    .Col = 5    'Cl
        '    aRcvryData.AdjClTonsWvol = .Value
        'End With

        'With ssProdRcvryFctrs
        '    .Col = 3
        '    .Row = 1    'CrsPb ton rcvry
        '    aRcvryData.PbTonRcvryCrs = .Value
        '    '.Row = 2    'IP ton rcvry
        '    'aRcvryData.IpTonRcvryTot = .Value
        '    .Row = 3    'CrsFd ton rcvry
        '    aRcvryData.FdTonRcvryCrs = .Value
        '    .Row = 4    'CrsFd BPL rcvry
        '    aRcvryData.FdBplRcvryCrs = .Value
        '    '.Row = 5    'Clay ton rcvry
        '    'aRcvryData.ClTonRcvryTot = .Value

        '    .Col = 5
        '    .Row = 1    'FnePb ton rcvry
        '    aRcvryData.PbTonRcvryFne = .Value
        '    .Row = 3    'FneFd ton rcvry
        '    aRcvryData.FdTonRcvryFne = .Value
        '    .Row = 4    'FneFd BPL rcvry
        '    aRcvryData.FdBplRcvryFne = .Value
        'End With

        'With ssFlotRcvryLinear
        '    .Row = 1
        '    .Col = 1    'Check box (Linear model)
        '    Chk1Val = .Value
        '    .Col = 5    'Coarse Constant factor (Linear model)
        '    aRcvryData.FlotRcvryCrsCf = .Value
        '    .Col = 6    'Coarse Variable factor (Linear model)
        '    aRcvryData.FlotRcvryCrsVf = .Value
        '    .Row = 2
        '    .Col = 5    'Fine Constant factor (Linear model)
        '    aRcvryData.FlotRcvryFneCf = .Value
        '    .Col = 6    'Fine Variable factor (Linear model)
        '    aRcvryData.FlotRcvryFneVf = .Value
        'End With

        'aRcvryData.LmTest = chkTestResultVsLabFlot1.Checked

        'With ssFlotRcvryHardwire
        '    .Row = 1
        '    .Col = 1    'Check box  (Hard-wire)
        '    Chk2Val = .Value
        '    .Col = 5    'Coarse Tailings BPL  (Hard-wire)
        '    aRcvryData.FlotRcvryCrsTlBpl = .Value
        '    .Col = 6    'Coarse Concentrate Insol  (Hard-wire)
        '    aRcvryData.FlotRcvryCrsCnIns = .Value
        '    .Col = 5    'Fine Tailings BPL  (Hard-wire)
        '    aRcvryData.FlotRcvryFneTlBpl = .Value
        '    .Col = 6    'Fine Concentrate Insol  (Hard-wire)
        '    aRcvryData.FlotRcvryFneCnIns = .Value
        'End With

        'aRcvryData.HwTest = chkTestResultVsLabFlot2.Checked

        'With ssOtherFlotMethods
        '    .Row = 1
        '    .Col = 1    'Check -- Use lab flotation recovery
        '    Chk3Val = .Value
        '    .Row = 2
        '    .Col = 1    'Check -- Use SqrRt(Feed BPL) = Tail BPL
        '    Chk4Val = .Value
        'End With

        'Assume that one check-box is checked
        'If Chk1Val = 1 Then
        '    'aRcvryData.FlotRcvryMode = "Linear model"
        'End If
        'If Chk2Val = 1 Then
        '    'aRcvryData.FlotRcvryMode = "Hard-wire"
        'End If
        'If Chk3Val = 1 Then
        '    'aRcvryData.FlotRcvryMode = "Lab flotation"
        'End If
        'If Chk4Val = 1 Then
        '    'aRcvryData.FlotRcvryMode = "SqrRt feed BPL"
        'End If

        'With _recoveryScenariosForm.ProductRecoveryDefinition.InsolAdjustment
        '    aRcvryData.CrsPbInsAdjMode = .CrsPbAdjMode
        '    aRcvryData.CrsPbInsAdj = .CrsPbAdjValue
        '    aRcvryData.FnePbInsAdjMode = .FnePbAdjMode
        '    aRcvryData.FnePbInsAdj = .FnePbAdjValue
        '    aRcvryData.IpInsAdjMode = .IntermediateProductAdjMode
        '    aRcvryData.IpInsAdj = .IntermediateProductAdjValue
        '    aRcvryData.CrsCnInsAdjMode = .CrsCnAdjMode
        '    aRcvryData.CrsCnInsAdj = .CrsCnAdjValue
        '    aRcvryData.FneCnInsAdjMode = .FneCnAdjMode
        '    aRcvryData.FneCnInsAdj = .FneCnAdjValue
        'End With
        aRcvryData.AdjInsAfterQualTest = False

        'With ssInsAdj100Pct
        '    .Row = 1    'Coarse pebble
        '    .Col = 1    'Minimum
        '    Chk1Val = .Value
        '    .Col = 2    'Direct
        '    Chk2Val = .Value
        '    .Col = 3    'Incremental
        '    Chk3Val = .Value
        '    .Col = 4    'In-Situ
        '    Chk4Val = .Value
        '    .Col = 6    'Adjustment value
        '    aRcvryData.CrsPbInsAdj100 = .Value

        '    'Assume that one check-box is checked
        '    If Chk1Val = 1 Then
        '        aRcvryData.CrsPbInsAdjMode100 = "Minimum"
        '    End If
        '    If Chk2Val = 1 Then
        '        aRcvryData.CrsPbInsAdjMode100 = "Direct"
        '    End If
        '    If Chk3Val = 1 Then
        '        aRcvryData.CrsPbInsAdjMode100 = "Incremental"
        '    End If
        '    If Chk4Val = 1 Then
        '        aRcvryData.CrsPbInsAdjMode100 = "In-Situ"
        '    End If

        '    .Row = 2    'Fine pebble
        '    .Col = 1    'Minimum
        '    Chk1Val = .Value
        '    .Col = 2    'Direct
        '    Chk2Val = .Value
        '    .Col = 3    'Incremental
        '    Chk3Val = .Value
        '    .Col = 4    'In-Situ
        '    Chk4Val = .Value
        '    .Col = 6    'Adjustment value
        '    aRcvryData.FnePbInsAdj100 = .Value

        '    'Assume that one check-box is checked
        '    If Chk1Val = 1 Then
        '        aRcvryData.FnePbInsAdjMode100 = "Minimum"
        '    End If
        '    If Chk2Val = 1 Then
        '        aRcvryData.FnePbInsAdjMode100 = "Direct"
        '    End If
        '    If Chk3Val = 1 Then
        '        aRcvryData.FnePbInsAdjMode100 = "Incremental"
        '    End If
        '    If Chk4Val = 1 Then
        '        aRcvryData.FnePbInsAdjMode100 = "In-Situ"
        '    End If

        '    .Row = 3    'IP
        '    .Col = 1    'Minimum
        '    Chk1Val = .Value
        '    .Col = 2    'Direct
        '    Chk2Val = .Value
        '    .Col = 3    'Incremental
        '    Chk3Val = .Value
        '    .Col = 4    'In-Situ
        '    Chk4Val = .Value
        '    .Col = 6    'Adjustment value
        '    aRcvryData.IpInsAdj100 = .Value

        '    'Assume that one check-box is checked
        '    If Chk1Val = 1 Then
        '        aRcvryData.IpInsAdjMode100 = "Minimum"
        '    End If
        '    If Chk2Val = 1 Then
        '        aRcvryData.IpInsAdjMode100 = "Direct"
        '    End If
        '    If Chk3Val = 1 Then
        '        aRcvryData.IpInsAdjMode100 = "Incremental"
        '    End If
        '    If Chk4Val = 1 Then
        '        aRcvryData.IpInsAdjMode100 = "In-Situ"
        '    End If

        '    .Row = 4    'Coarse concentrate
        '    .Col = 1    'Minimum
        '    Chk1Val = .Value
        '    .Col = 2    'Direct
        '    Chk2Val = .Value
        '    .Col = 3    'Incremental
        '    Chk3Val = .Value
        '    .Col = 4    'In-Situ
        '    Chk4Val = .Value
        '    .Col = 6    'Adjustment value
        '    aRcvryData.CrsCnInsAdj100 = .Value

        '    'Assume that one check-box is checked
        '    If Chk1Val = 1 Then
        '        aRcvryData.CrsCnInsAdjMode100 = "Minimum"
        '    End If
        '    If Chk2Val = 1 Then
        '        aRcvryData.CrsCnInsAdjMode100 = "Direct"
        '    End If
        '    If Chk3Val = 1 Then
        '        aRcvryData.CrsCnInsAdjMode100 = "Incremental"
        '    End If
        '    If Chk4Val = 1 Then
        '        aRcvryData.CrsCnInsAdjMode100 = "In-Situ"
        '    End If

        '    .Row = 5    'Fine Concentrate
        '    .Col = 1    'Minimum
        '    Chk1Val = .Value
        '    .Col = 2    'Direct
        '    Chk2Val = .Value
        '    .Col = 3    'Incremental
        '    Chk3Val = .Value
        '    .Col = 4    'In-Situ
        '    Chk4Val = .Value
        '    .Col = 6    'Adjustment value
        '    aRcvryData.FneCnInsAdj100 = .Value

        '    'Assume that one check-box is checked
        '    If Chk1Val = 1 Then
        '        aRcvryData.FneCnInsAdjMode100 = "Minimum"
        '    End If
        '    If Chk2Val = 1 Then
        '        aRcvryData.FneCnInsAdjMode100 = "Direct"
        '    End If
        '    If Chk3Val = 1 Then
        '        aRcvryData.FneCnInsAdjMode100 = "Incremental"
        '    End If
        '    If Chk4Val = 1 Then
        '        aRcvryData.FneCnInsAdjMode100 = "In-Situ"
        '    End If
        'End With

        'fraMineability
        'fraMineability
        'fraMineability
        'With ssSplitPhysMineability
        '    .Row = 1
        '    .Col = 1    'Maximum %clay  Split
        '    aRcvryData.ClPctMaxSpl = .Value
        '    .Row = 2
        '    .Col = 1    'Maximum total depth
        '    aRcvryData.MaxTotDepthSpl = .Value
        '    .Col = 2    'Absolute stop
        '    Chk1Val = .Value
        '    .Col = 3    'Finish split
        '    Chk2Val = .Value
        'End With

        'Assume that one check-box is checked -- set value to " " just in
        'case it is not!
        'aRcvryData.MaxTotDepthModeSpl = " "
        'If Chk1Val = 1 Then
        '    aRcvryData.MaxTotDepthModeSpl = "Absolute stop"
        'End If
        'If Chk2Val = 1 Then
        '    aRcvryData.MaxTotDepthModeSpl = "Finish split"
        'End If

        'With ssSplitEconMineability
        '    .Row = 1
        '    .Col = 1    'Maximum Mtx-X  Split
        '    aRcvryData.MtxxMaxSpl = .Value
        'End With

        'With ssHolePhysMineability
        '    .Row = 1
        '    .Col = 1    'Maximum %clay  Hole
        '    aRcvryData.ClPctMaxHole = .Value
        '    .Row = 2
        '    .Col = 1    'Minimum ore thickness  Hole
        '    aRcvryData.MinOreThk = .Value
        '    .Row = 3
        '    .Col = 1    'Minimum interburden thickness
        '    aRcvryData.MinItbThk = .Value
        'End With

        'With ssHoleEconMineability
        '    .Row = 1
        '    .Col = 1    'Maximum Mtx-X  Hole
        '    aRcvryData.MtxxMaxHole = .Value
        '    .Row = 2
        '    .Col = 1    'Maximum Tot-X  Hole
        '    aRcvryData.TotxMaxHole = .Value
        '    .Row = 3
        '    .Col = 1    'Minimum total product TPA
        '    aRcvryData.TotPrTpaMinHole = .Value
        'End With

        'Currently in Private Sub GetRcvryEtcParamsFromForm

        'aRcvryData.MineFirstSpl = IIf(chkNoUnmineableHoles.Checked, 1, 0)
        'aRcvryData.InclCpbAlways = chkInclCpbAlways.Checked
        'aRcvryData.InclFpbAlways = chkInclFpbAlways.Checked
        'aRcvryData.InclOsAlways = chkInclOsAlways.Checked

        'aRcvryData.InclCpbNever = chkInclCpbNever.Checked
        'aRcvryData.InclFpbNever = chkInclFpbNever.Checked
        'aRcvryData.InclOsNever = chkInclOsNever.Checked
        'aRcvryData.CanSelectRejectTpb = chkCanSelectRejectTpb.Checked
        'aRcvryData.CanSelectRejectTcn = chkCanSelectRejectTcn.Checked
        'aRcvryData.MineHasOffSpecPbPlt = chkMineHasOffSpecPbPlt.Checked
        '06/14/2010, lss  MineHasOffSpecPbPlt not really used anymore!
        aRcvryData.UseOrigMgoPlant = chkUseOrigMgoPlant.Checked
        aRcvryData.UseDoloflotPlant2010 = chkUseDoloflotPlant.Checked
        aRcvryData.UseDoloflotPlantFco = chkUseDoloflotPlantFco.Checked

        'aRcvryData.WgFeAdjCutoffDate = _recoveryScenariosForm.ProductRecoveryDefinition.WingateCutoffDate

        With _recoveryScenariosForm.ProductRecoveryDefinition
            aRcvryData.DensCalcMode = .DensityCalculationMode
            aRcvryData.DensMlvSpecChk = .MeasuredLabTestMeasuredVsCalculated
            aRcvryData.DensLrSpecChk = .LimitTestMeasuredVsCalculated
            aRcvryData.DensLowerLimit = .DensityLowerLimit
            aRcvryData.DensUpperLimit = .DensityUpperLimit
        End With

        'If opt100PctTlZero.Checked = True Then
        '    aRcvryData.FlotRcvryMode100 = "0 tail BPL"
        'End If
        'If opt100PctSqrRtFd.Checked = True Then
        '    aRcvryData.FlotRcvryMode100 = "SqrRt feed BPL"
        'End If

        With ssOffSpecPb
            .Row = 1
            .Col = 1
            aRcvryData.MplInpBplTarg = .Value
            .Row = 2
            aRcvryData.MplInpMgoTarg = .Value
            .Row = 3
            aRcvryData.MplRejBplTarg = .Value
            .Row = 4
            aRcvryData.MplRejMgoTarg = .Value
            '-----
            .Row = 5
            aRcvryData.MplM1BpltRcvry = .Value
            .Row = 6
            aRcvryData.MplM1BplHwire = .Value
            .Row = 7
            aRcvryData.MplM1InsHwire = .Value
            .Row = 8
            aRcvryData.MplM1MgoImprove = .Value
        End With

        If chkUseRawProspAsOverride.Checked = True Then
            aRcvryData.UseRawProspAsOverride = True
        Else
            aRcvryData.UseRawProspAsOverride = False
        End If

        If chkPbAnalysisFillInSpecial.Checked = True Then
            aRcvryData.SetPbToMgPlt = True
        Else
            aRcvryData.SetPbToMgPlt = False
        End If

        'aRcvryData.UseFeAdjust = _recoveryScenariosForm.ProductRecoveryDefinition.UseAdjustedFeToDetermineMineability
        'aRcvryData.UpperZoneFeAdjust = _recoveryScenariosForm.ProductRecoveryDefinition.UpperLimitCorrection
        'aRcvryData.LowerZoneFeAdjust = _recoveryScenariosForm.ProductRecoveryDefinition.LowerLimitCorrection


        With ssDoloflotPlant
            If Not chkUseDoloflotPlantFco.Checked Then
                .Col = 1
                .Row = 1
                aRcvryData.DpFnePbMgoCutoff = .Value   'Could be TotPb MgO Max
                .Row = 2
                aRcvryData.DpIpMgoCutoff = .Value      'Could be TotPb MgO Min
            End If

            .Row = 3
            aRcvryData.DpGrind = .Value

            If Not chkUseDoloflotPlantFco.Checked Then
                .Row = 4
                aRcvryData.DpAcid = .Value
                .Row = 5
                aRcvryData.DpP2o5 = .Value
                .Row = 6
                aRcvryData.DpPa64 = .Value
                .Row = 7
                aRcvryData.DpFlotMin = .Value
            End If

            .Row = 8
            aRcvryData.DpTargMgo = .Value
            ''       .Row = 9
            ''       aRcvryData.DpAl2O3GreaterThan = .Value
            ''       .Row = 10
            ''       aRcvryData.DpFe2O3GreaterThan = .Value
        End With

        With ssDoloflotPlantFco
            If chkUseDoloflotPlantFco.Checked Then
                .Col = 1
                .Row = 1
                aRcvryData.DpFnePbMgoCutoff = .Value   'Could be TotPb MgO Max
                .Row = 2
                aRcvryData.DpIpMgoCutoff = .Value      'Could be TotPb MgO Min
            End If
        End With

        'New 11/16/2011
        'Currently in Sub GetRcvryEtcParamsFromForm
        If chkUseDoloflotPlantFco.Checked Then
            With ssDoloflotPlantFco2
                .Row = 1
                aRcvryData.DpPctWtM200Mesh = .Value
                .Row = 2
                aRcvryData.DpCondMinutes = .Value
                .Row = 3
                aRcvryData.DpCondPctSolids = .Value
                .Row = 4
                aRcvryData.DpFlotMin = .Value
                .Row = 5
                aRcvryData.DpPa64 = .Value
                .Row = 6
                aRcvryData.DpP2o5 = .Value
                .Row = 7
                aRcvryData.DpAcid = .Value
            End With
        End If

    End Sub

    Public Sub SaveRcvryScenario()

        Dim RcvryData As gDataRdctnParamsType
        Dim RcvryProdQual(0 To 14) As gDataRdctnProdQualType

        SetActionStatus("Saving recovery scenario...")

        'Get the recovery parameters from the form so we can save them!
        GetRcvryEtcParamsFromForm(RcvryData)

        With _recoveryScenariosForm.ProductRecoveryDefinition

            '.ClPctMaxSpl = RcvryData.ClPctMaxSpl
            '.ClPctMaxHole = RcvryData.ClPctMaxHole
            '.MinOreThk = RcvryData.MinOreThk
            '.MinItbThk = RcvryData.MinItbThk


            '.MtxxMaxSpl = RcvryData.MtxxMaxSpl
            '.MtxxMaxHole = RcvryData.MtxxMaxHole
            '.TotxMaxHole = RcvryData.TotxMaxHole
            '.TotPrTpaMinHole = RcvryData.TotPrTpaMinHole

            '.InclCrsPbStatus = If(RcvryData.InclCpbAlways, "Always", "")
            'If RcvryData.InclCpbNever = True Then
            '    .InclCrsPbStatus = "Never"
            'End If
            '.InclFnePbStatus = If(RcvryData.InclFpbAlways, "Always", "")
            'If RcvryData.InclFpbNever = True Then
            '    .InclFnePbStatus = "Never"
            'End If
            'If RcvryData.InclOsAlways = True Then
            '    .InclOsStatus = "Always"
            'End If
            'If RcvryData.InclOsNever = True Then
            '    .InclOsStatus = "Never"
            'End If
            '.CanSelectRejectTpb = If(RcvryData.CanSelectRejectTpb, 1, 0)
            '.CanSelectRejectTcn = If(RcvryData.CanSelectRejectTcn, 1, 0)


            .MineHasOffSpecPbPlt = If(RcvryData.MineHasOffSpecPbPlt, 1, 0)
            .UseOrigMgoPlant = If(RcvryData.UseOrigMgoPlant, 1, 0)
            If RcvryData.UseDoloflotPlant2010 = True Then
                .UseDoloflotPlant2010 = 1
            Else
                If RcvryData.UseDoloflotPlantFco = True Then
                    .UseDoloflotPlant2010 = 2
                Else
                    .UseDoloflotPlant2010 = 0
                End If
            End If

            '.OvbVolRcvryMode = RcvryData.OvbVolRcvryMode  'Linear model
            '.OvbVolRcvryCf = RcvryData.OvbVolRcvryCf
            '.OvbVolRcvryVf = RcvryData.OvbVolRcvryVf
            '.OvbVolRcvryFa = RcvryData.OvbVolRcvryFa
            '.MineVolRcvryMode = RcvryData.MineVolRcvryMode 'Linear model
            '.MineVolRcvryCf = RcvryData.MineVolRcvryCf
            '.MineVolRcvryVf = RcvryData.MineVolRcvryVf
            '.MineVolRcvryFa = RcvryData.MineVolRcvryFa

            '.AdjOsTonsWvol = RcvryData.AdjOsTonsWvol
            '.AdjPbTonsWvol = RcvryData.AdjPbTonsWvol
            '.AdjIpTonsWvol = RcvryData.AdjIpTonsWvol
            '.AdjFdTonsWvol = RcvryData.AdjFdTonsWvol
            '.AdjClTonsWvol = RcvryData.AdjClTonsWvol

            .PbTonRcvryCrs = RcvryData.PbTonRcvryCrs
            .PbTonRcvryFne = RcvryData.PbTonRcvryFne
            '.IpTonRcvryTot = RcvryData.IpTonRcvryTot
            .FdTonRcvryCrs = RcvryData.FdTonRcvryCrs
            .FdTonRcvryFne = RcvryData.FdTonRcvryFne
            .FdBplRcvryCrs = RcvryData.FdBplRcvryCrs
            .FdBplRcvryFne = RcvryData.FdBplRcvryFne
            '.ClTonRcvryTot = RcvryData.ClTonRcvryTot
            '.FlotRcvryMode = RcvryData.FlotRcvryMode
            .FlotRcvryCrsCf = RcvryData.FlotRcvryCrsCf
            .FlotRcvryCrsVf = RcvryData.FlotRcvryCrsVf
            .FlotRcvryFneCf = RcvryData.FlotRcvryFneCf
            .FlotRcvryFneVf = RcvryData.FlotRcvryFneVf
            .FlotRcvryCrsTlBpl = RcvryData.FlotRcvryCrsTlBpl
            .FlotRcvryCrsCnIns = RcvryData.FlotRcvryCrsCnIns
            .FlotRcvryFneTlBpl = RcvryData.FlotRcvryFneTlBpl
            .FlotRcvryFneCnIns = RcvryData.FlotRcvryFneCnIns
            .LmTest = RcvryData.LmTest
            .HwTest = RcvryData.HwTest

            .MaxTotDepthSpl = RcvryData.MaxTotDepthSpl
            .MaxTotDepthModeSpl = RcvryData.MaxTotDepthModeSpl

            .MineFirstSpl = RcvryData.MineFirstSpl
            '.FlotRcvryMode100 = RcvryData.FlotRcvryMode100
            .CrsPbInsAdjMode100 = RcvryData.CrsPbInsAdjMode100   '''ToDo: Comment this Line?????????????
            .CrsPbInsAdj100 = RcvryData.CrsPbInsAdj100
            .FnePbInsAdjMode100 = RcvryData.FnePbInsAdjMode100
            .FnePbInsAdj100 = RcvryData.FnePbInsAdj100
            .IpInsAdjMode100 = RcvryData.IpInsAdjMode100
            .IpInsAdj100 = RcvryData.IpInsAdj100
            .CrsCnInsAdjMode100 = RcvryData.CrsCnInsAdjMode100
            .CrsCnInsAdj100 = RcvryData.CrsCnInsAdj100
            .FneCnInsAdjMode100 = RcvryData.FneCnInsAdjMode100
            .FneCnInsAdj100 = RcvryData.FneCnInsAdj100

            .MplInpBplTarg = RcvryData.MplInpBplTarg
            .MplInpMgoTarg = RcvryData.MplInpMgoTarg
            .MplRejBplTarg = RcvryData.MplRejBplTarg
            .MplRejMgoTarg = RcvryData.MplRejMgoTarg
            .MplM1BpltRcvry = RcvryData.MplM1BpltRcvry
            .MplM1BplHwire = RcvryData.MplM1BplHwire
            .MplM1InsHwire = RcvryData.MplM1InsHwire
            .MplM1MgoImprove = RcvryData.MplM1MgoImprove
            .DpCrsPbMgoCutoff = RcvryData.DpCrsPbMgoCutoff
            .DpFnePbMgoCutoff = RcvryData.DpFnePbMgoCutoff
            .DpIpMgoCutoff = RcvryData.DpIpMgoCutoff
            .DpGrind = RcvryData.DpGrind
            .DpAcid = RcvryData.DpAcid
            .DpP2o5 = RcvryData.DpP2o5
            .DpPa64 = RcvryData.DpPa64
            .DpFlotMin = RcvryData.DpFlotMin
            .DpTargMgo = RcvryData.DpTargMgo
            .DpPctWtM200Mesh = RcvryData.DpPctWtM200Mesh
            .DpCondMinutes = RcvryData.DpCondMinutes
            .DpCondPctSolids = RcvryData.DpCondPctSolids
        End With
    End Sub


    Public Sub DisplayRcvryEtc(recoveryDefinition As ViewModels.ProductRecoveryDefinition)

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RcvryScenDynaset As OraDynaset
        Dim RcvryData As gDataRdctnParamsType
        Dim RcvryProdQual(0 To 14) As gDataRdctnProdQualType
        Dim RecordCount As Integer
        Dim TempData As gDataRdctnProdQualType
        Dim ProdRow As Integer

        Try
            _loadingRecoveryScenarioData = True
            SetActionStatus("Getting recovery scenario...")
            params = gDBParams
            params.Add("pRcvryScenarioName", recoveryDefinition.ScenarioName, ORAPARM_INPUT)
            params("pRcvryScenarioName").serverType = ORATYPE_VARCHAR2
            params.Add("pProspSetName", recoveryDefinition.ProspectSetName, ORAPARM_INPUT)
            params("pProspSetName").serverType = ORATYPE_VARCHAR2
            params.Add("pResult", 0, ORAPARM_OUTPUT)
            params("pResult").serverType = ORATYPE_CURSOR

            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prosp_data_rdctn.get_rcvry_base(" +
                     ":pRcvryScenarioName, :pProspSetName, :pResult);end;", ORASQL_FAILEXEC)

            RcvryScenDynaset = params("pResult").Value
            ClearParams(params)
            RecordCount = RcvryScenDynaset.RecordCount
            RcvryScenDynaset.MoveFirst()
            With RcvryData
                .RcvryScenarioName = RcvryScenDynaset.Fields("rcvry_scenario_name").Value
                .ProspSetName = RcvryScenDynaset.Fields("prosp_set_name").Value
                .WhoDefined = RcvryScenDynaset.Fields("who_defined").Value
                .WhenDefined = RcvryScenDynaset.Fields("when_defined").Value

                If Not IsDBNull(RcvryScenDynaset.Fields("mine_name").Value) Then
                    .MineName = RcvryScenDynaset.Fields("mine_name").Value
                Else
                    .MineName = "None"
                End If
                'Legacy .OvbVolRcvryMode = RcvryScenDynaset.Fields("ovb_vol_rcvry_mode").Value
                '.OvbVolRcvryCf = RcvryScenDynaset.Fields("ovb_vol_rcvry_cf").Value
                'Condition added by VJ to handle null values
                If Not IsDBNull(RcvryScenDynaset.Fields("ovb_vol_rcvry_vf").Value) Then
                    .OvbVolRcvryVf = RcvryScenDynaset.Fields("ovb_vol_rcvry_vf").Value
                End If

                'Legacy .OvbVolRcvryFa = RcvryScenDynaset.Fields("ovb_vol_rcvry_fa").Value
                'Legacy .MineVolRcvryMode = RcvryScenDynaset.Fields("mine_vol_rcvry_mode").Value
                '.MineVolRcvryCf = RcvryScenDynaset.Fields("mine_vol_rcvry_cf").Value
                'Condition added by VJ to handle null values
                If Not IsDBNull(RcvryScenDynaset.Fields("mine_vol_rcvry_vf").Value) Then
                    .MineVolRcvryVf = RcvryScenDynaset.Fields("mine_vol_rcvry_vf").Value
                End If

                .MineVolRcvryFa = RcvryScenDynaset.Fields("mine_vol_rcvry_fa").Value
                .AdjOsTonsWvol = RcvryScenDynaset.Fields("adj_os_tons_wvol").Value
                .AdjPbTonsWvol = RcvryScenDynaset.Fields("adj_pb_tons_wvol").Value
                .AdjIpTonsWvol = RcvryScenDynaset.Fields("adj_ip_tons_wvol").Value
                .AdjFdTonsWvol = RcvryScenDynaset.Fields("adj_fd_tons_wvol").Value
                .AdjClTonsWvol = RcvryScenDynaset.Fields("adj_cl_tons_wvol").Value
                'Legacy .PbTonRcvryCrs = RcvryScenDynaset.Fields("pb_ton_rcvry_crs").Value
                'Legacy .PbTonRcvryFne = RcvryScenDynaset.Fields("pb_ton_rcvry_fne").Value
                '.IpTonRcvryTot = RcvryScenDynaset.Fields("ip_ton_rcvry_tot").Value
                'Legacy .FdTonRcvryCrs = RcvryScenDynaset.Fields("fd_ton_rcvry_crs").Value
                'Legacy .FdTonRcvryFne = RcvryScenDynaset.Fields("fd_ton_rcvry_fne").Value
                'Legacy .FdBplRcvryCrs = RcvryScenDynaset.Fields("fd_bpl_rcvry_crs").Value
                'Legacy .FdBplRcvryFne = RcvryScenDynaset.Fields("fd_bpl_rcvry_fne").Value
                '.ClTonRcvryTot = RcvryScenDynaset.Fields("cl_ton_rcvry_tot").Value
                '.FlotRcvryMode = RcvryScenDynaset.Fields("flot_rcvry_mode").Value
                'Legacy .FlotRcvryCrsCf = RcvryScenDynaset.Fields("flot_rcvry_crs_cf").Value
                .FlotRcvryCrsVf = RcvryScenDynaset.Fields("flot_rcvry_crs_vf").Value
                'Legacy FlotRcvryFneCf = RcvryScenDynaset.Fields("flot_rcvry_fne_cf").Value
                .FlotRcvryFneVf = RcvryScenDynaset.Fields("flot_rcvry_fne_vf").Value
                'Legacy .FlotRcvryCrsTlBpl = RcvryScenDynaset.Fields("flot_rcvry_crs_tlbpl").Value
                'Legacy .FlotRcvryCrsCnIns = RcvryScenDynaset.Fields("flot_rcvry_crs_cnins").Value
                'Legacy .FlotRcvryFneTlBpl = RcvryScenDynaset.Fields("flot_rcvry_fne_tlbpl").Value
                'Legacy .FlotRcvryFneCnIns = RcvryScenDynaset.Fields("flot_rcvry_fne_cnins").Value
                'Legacy .LmTest = RcvryScenDynaset.Fields("lm_test").Value
                .HwTest = RcvryScenDynaset.Fields("hw_test").Value

                ''Product adjustments -- fraProdAdj
                'If Not IsDBNull(RcvryScenDynaset.Fields("crspb_insadj_mode").Value) Then
                '    .CrsPbInsAdjMode = RcvryScenDynaset.Fields("crspb_insadj_mode").Value
                'Else
                '    .CrsPbInsAdjMode = ""
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("crspb_insadj").Value) Then
                '    .CrsPbInsAdj = RcvryScenDynaset.Fields("crspb_insadj").Value
                'Else
                '    .CrsPbInsAdj = 0
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("fnepb_insadj_mode").Value) Then
                '    .FnePbInsAdjMode = RcvryScenDynaset.Fields("fnepb_insadj_mode").Value
                'Else
                '    .FnePbInsAdjMode = ""
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("fnepb_insadj").Value) Then
                '    .FnePbInsAdj = RcvryScenDynaset.Fields("fnepb_insadj").Value
                'Else
                '    .FnePbInsAdj = 0
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("ip_insadj_mode").Value) Then
                '    .IpInsAdjMode= RcvryScenDynaset.Fields("ip_insadj_mode").Value
                'Else
                '    .IpInsAdjMode = ""
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("ip_insadj").Value) Then
                '    .IpInsAdj = RcvryScenDynaset.Fields("ip_insadj").Value
                'Else
                '    .IpInsAdj = 0
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("crscn_insadj_mode").Value) Then
                '    .CrsCnInsAdjMode = RcvryScenDynaset.Fields("crscn_insadj_mode").Value
                'Else
                '    .CrsCnInsAdjMode = ""
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("crscn_insadj").Value) Then
                '    .CrsCnInsAdj = RcvryScenDynaset.Fields("crscn_insadj").Value
                'Else
                '    .CrsCnInsAdj = 0
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("fnecn_insadj_mode").Value) Then
                '    .FneCnInsAdjMode = RcvryScenDynaset.Fields("fnecn_insadj_mode").Value
                'Else
                '    .FneCnInsAdjMode = ""
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("fnecn_insadj").Value) Then
                '    .FneCnInsAdj = RcvryScenDynaset.Fields("fnecn_insadj").Value
                'Else
                '    .FneCnInsAdj = 0
                'End If

                .AdjInsAfterQualTest = RcvryScenDynaset.Fields("adj_ins_after_qual_test").Value

                ''fra100PctDefn
                'If Not IsDBNull(RcvryScenDynaset.Fields("crspb_insadj_mode_100").Value) Then
                '    .CrsPbInsAdjMode100 = RcvryScenDynaset.Fields("crspb_insadj_mode_100").Value
                'Else
                '    .CrsPbInsAdjMode100 = ""
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("crspb_insadj_100").Value) Then
                '    .CrsPbInsAdj100 = RcvryScenDynaset.Fields("crspb_insadj_100").Value
                'Else
                '    .CrsPbInsAdj100 = 0
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("fnepb_insadj_mode_100").Value) Then
                '    .FnePbInsAdjMode100 = RcvryScenDynaset.Fields("fnepb_insadj_mode_100").Value
                'Else
                '    .FnePbInsAdjMode100 = ""
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("fnepb_insadj_100").Value) Then
                '    .FnePbInsAdj100 = RcvryScenDynaset.Fields("fnepb_insadj_100").Value
                'Else
                '    .FnePbInsAdj100 = 0
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("ip_insadj_mode_100").Value) Then
                '    .IpInsAdjMode100 = RcvryScenDynaset.Fields("ip_insadj_mode_100").Value
                'Else
                '    .IpInsAdjMode100 = ""
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("ip_insadj_100").Value) Then
                '    .IpInsAdj100 = RcvryScenDynaset.Fields("ip_insadj_100").Value
                'Else
                '    .IpInsAdj100 = 0
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("crscn_insadj_mode_100").Value) Then
                '    .CrsCnInsAdjMode100 = RcvryScenDynaset.Fields("crscn_insadj_mode_100").Value
                'Else
                '    .CrsCnInsAdjMode100 = ""
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("crscn_insadj_100").Value) Then
                '    .CrsCnInsAdj100 = RcvryScenDynaset.Fields("crscn_insadj_100").Value
                'Else
                '    .CrsCnInsAdj100 = 0
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("fnecn_insadj_mode_100").Value) Then
                '    .FneCnInsAdjMode100 = RcvryScenDynaset.Fields("fnecn_insadj_mode_100").Value
                'Else
                '    .FneCnInsAdjMode100 = ""
                'End If
                'If Not IsDBNull(RcvryScenDynaset.Fields("fnecn_insadj_100").Value) Then
                '    .FneCnInsAdj100 = RcvryScenDynaset.Fields("fnecn_insadj_100").Value
                'Else
                '    .FneCnInsAdj100 = 0
                'End If

                'Economic and physical mineability criteria -- fraMineability
                If Not IsDBNull(RcvryScenDynaset.Fields("cl_pct_max_spl").Value) Then
                    .ClPctMaxSpl = RcvryScenDynaset.Fields("cl_pct_max_spl").Value
                End If
                If Not IsDBNull(RcvryScenDynaset.Fields("mtxx_max_spl").Value) Then
                    .MtxxMaxSpl = RcvryScenDynaset.Fields("mtxx_max_spl").Value
                End If
                If Not IsDBNull(RcvryScenDynaset.Fields("max_tot_depth_spl").Value) Then
                    .MaxTotDepthSpl = RcvryScenDynaset.Fields("max_tot_depth_spl").Value
                End If
                If Not IsDBNull(RcvryScenDynaset.Fields("max_tot_depth_mode_spl").Value) Then
                    .MaxTotDepthModeSpl = RcvryScenDynaset.Fields("max_tot_depth_mode_spl").Value
                End If
                If Not IsDBNull(RcvryScenDynaset.Fields("min_ore_thk").Value) Then
                    .MinOreThk = RcvryScenDynaset.Fields("min_ore_thk").Value
                End If
                If Not IsDBNull(RcvryScenDynaset.Fields("min_itb_thk").Value) Then
                    .MinItbThk = RcvryScenDynaset.Fields("min_itb_thk").Value
                End If
                If Not IsDBNull(RcvryScenDynaset.Fields("cl_pct_max_hole").Value) Then
                    .ClPctMaxHole = RcvryScenDynaset.Fields("cl_pct_max_hole").Value
                End If
                If Not IsDBNull(RcvryScenDynaset.Fields("mtxx_max_hole").Value) Then
                    .MtxxMaxHole = RcvryScenDynaset.Fields("mtxx_max_hole").Value
                End If
                If Not IsDBNull(RcvryScenDynaset.Fields("totx_max_hole").Value) Then
                    .TotxMaxHole = RcvryScenDynaset.Fields("totx_max_hole").Value
                End If
                If Not IsDBNull(RcvryScenDynaset.Fields("totpr_tpa_min_hole").Value) Then
                    .TotPrTpaMinHole = RcvryScenDynaset.Fields("totpr_tpa_min_hole").Value
                End If



                If Not IsDBNull(RcvryScenDynaset.Fields("mine_first_spl").Value) Then
                    .MineFirstSpl = RcvryScenDynaset.Fields("mine_first_spl").Value
                Else
                    .MineFirstSpl = 0
                End If

                .InclCpbAlways = True
                .InclFpbAlways = True
                .InclCpbNever = False
                .InclFpbNever = False
                .InclOsAlways = False
                .InclOsNever = True

                '06/15/2010, lss  .MineHasOffSpecPbPlt not really used anymore.
                .MineHasOffSpecPbPlt = False
                .UseDoloflotPlant2010 = False
                .UseDoloflotPlantFco = False
                .UseOrigMgoPlant = False
                .CanSelectRejectTpb = False
                .CanSelectRejectTcn = False

                If Not IsDBNull(RcvryScenDynaset.Fields("incl_crspb_status").Value) Then
                    If RcvryScenDynaset.Fields("incl_crspb_status").Value = "Always" Then
                        .InclCpbAlways = True
                    End If
                    If RcvryScenDynaset.Fields("incl_crspb_status").Value = "Never" Then
                        .InclCpbNever = True
                    End If
                End If
                If Not IsDBNull(RcvryScenDynaset.Fields("incl_fnepb_status").Value) Then
                    If RcvryScenDynaset.Fields("incl_fnepb_status").Value = "Always" Then
                        .InclFpbAlways = True
                    End If
                    If RcvryScenDynaset.Fields("incl_fnepb_status").Value = "Never" Then
                        .InclFpbNever = True
                    End If
                End If
                If Not IsDBNull(RcvryScenDynaset.Fields("incl_os_status").Value) Then
                    If RcvryScenDynaset.Fields("incl_os_status").Value = "Always" Then
                        .InclOsAlways = True
                    End If
                    If RcvryScenDynaset.Fields("incl_os_status").Value = "Never" Then
                        .InclOsNever = True
                    End If
                End If

                '06/15/2010, lss  .MineHasOffSpecPbPlt not really used anymore.
                If Not IsDBNull(RcvryScenDynaset.Fields("mine_has_offspec_pb_plt").Value) Then
                    If RcvryScenDynaset.Fields("mine_has_offspec_pb_plt").Value = 1 Then
                        .MineHasOffSpecPbPlt = True
                    Else
                        .MineHasOffSpecPbPlt = False
                    End If
                End If

                .UseDoloflotPlant2010 = False
                .UseDoloflotPlantFco = False
                If Not IsDBNull(RcvryScenDynaset.Fields("use_doloflot_plant_2010").Value) Then
                    If RcvryScenDynaset.Fields("use_doloflot_plant_2010").Value = 1 Then
                        .UseDoloflotPlant2010 = True
                        .UseDoloflotPlantFco = False
                    Else
                        If RcvryScenDynaset.Fields("use_doloflot_plant_2010").Value = 2 Then
                            .UseDoloflotPlant2010 = False
                            .UseDoloflotPlantFco = True
                        End If
                    End If
                Else
                    .UseDoloflotPlant2010 = False
                    .UseDoloflotPlantFco = False
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("use_orig_mgo_plant").Value) Then
                    If RcvryScenDynaset.Fields("use_orig_mgo_plant").Value = 1 Then
                        .UseOrigMgoPlant = True
                    Else
                        .UseOrigMgoPlant = False
                    End If
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("can_select_reject_tpb").Value) Then
                    If RcvryScenDynaset.Fields("can_select_reject_tpb").Value = 1 Then
                        .CanSelectRejectTpb = True
                    Else
                        .CanSelectRejectTpb = False
                    End If
                End If

                'If Not IsDBNull(RcvryScenDynaset.Fields("flot_rcvry_mode_100").Value) Then
                '    .FlotRcvryMode100 = RcvryScenDynaset.Fields("flot_rcvry_mode_100").Value
                'Else
                '    .FlotRcvryMode100 = "0 tail BPL"
                'End If

                If Not IsDBNull(RcvryScenDynaset.Fields("dens_calc_mode").Value) Then
                    .DensCalcMode = RcvryScenDynaset.Fields("dens_calc_mode").Value
                Else
                    .DensCalcMode = "Limit routine"
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("dens_mlv_spec_chk").Value) Then
                    If RcvryScenDynaset.Fields("dens_mlv_spec_chk").Value = 1 Then
                        .DensMlvSpecChk = True
                    Else
                        .DensMlvSpecChk = False
                    End If
                Else
                    .DensMlvSpecChk = False
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("dens_lr_spec_chk").Value) Then
                    If RcvryScenDynaset.Fields("dens_lr_spec_chk").Value = 1 Then
                        .DensLrSpecChk = True
                    Else
                        .DensLrSpecChk = False
                    End If
                Else
                    .DensLrSpecChk = False
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("dens_upper_limit").Value) Then
                    .DensUpperLimit = RcvryScenDynaset.Fields("dens_upper_limit").Value
                Else
                    .DensUpperLimit = 0
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("dens_lower_limit").Value) Then
                    .DensLowerLimit = RcvryScenDynaset.Fields("dens_lower_limit").Value
                Else
                    .DensLowerLimit = 0
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("mpl_inp_bpl_targ").Value) Then
                    .MplInpBplTarg = RcvryScenDynaset.Fields("mpl_inp_bpl_targ").Value
                Else
                    .MplInpBplTarg = 0
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("mpl_inp_mgo_targ").Value) Then
                    .MplInpMgoTarg = RcvryScenDynaset.Fields("mpl_inp_mgo_targ").Value
                Else
                    .MplInpMgoTarg = 0
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("mpl_rej_bpl_targ").Value) Then
                    .MplRejBplTarg = RcvryScenDynaset.Fields("mpl_rej_bpl_targ").Value
                Else
                    .MplRejBplTarg = 0
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("mpl_rej_mgo_targ").Value) Then
                    .MplRejMgoTarg = RcvryScenDynaset.Fields("mpl_rej_mgo_targ").Value
                Else
                    .MplRejMgoTarg = 0
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("mpl_m1_bplt_rcvry").Value) Then
                    .MplM1BpltRcvry = RcvryScenDynaset.Fields("mpl_m1_bplt_rcvry").Value
                Else
                    .MplM1BpltRcvry = 0
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("mpl_m1_bpl_hwire").Value) Then
                    .MplM1BplHwire = RcvryScenDynaset.Fields("mpl_m1_bpl_hwire").Value
                Else
                    .MplM1BplHwire = 0
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("mpl_m1_ins_hwire").Value) Then
                    .MplM1InsHwire = RcvryScenDynaset.Fields("mpl_m1_ins_hwire").Value
                Else
                    .MplM1InsHwire = 0
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("mpl_m1_mgo_improve").Value) Then
                    .MplM1MgoImprove = RcvryScenDynaset.Fields("mpl_m1_mgo_improve").Value
                Else
                    .MplM1MgoImprove = 0
                End If

                If Not IsDBNull(RcvryScenDynaset.Fields("can_select_reject_tcn").Value) Then
                    If RcvryScenDynaset.Fields("can_select_reject_tcn").Value = 1 Then
                        .CanSelectRejectTcn = True
                    Else
                        .CanSelectRejectTcn = False
                    End If
                End If

                'If Not IsDBNull(RcvryScenDynaset.Fields("use_adj_fe_for_minable").Value) Then
                '    If RcvryScenDynaset.Fields("use_adj_fe_for_minable").Value = 1 Then
                '        .UseFeAdjust = True
                '    Else
                '        .UseFeAdjust = False
                '    End If
                'End If

                'If Not IsDBNull(RcvryScenDynaset.Fields("upper_zone_fe_adjust").Value) Then
                '    .UpperZoneFeAdjust = RcvryScenDynaset.Fields("upper_zone_fe_adjust").Value
                'Else
                '    .UpperZoneFeAdjust = 0
                'End If

                'If Not IsDBNull(RcvryScenDynaset.Fields("lower_zone_fe_adjust").Value) Then
                '    .LowerZoneFeAdjust = RcvryScenDynaset.Fields("lower_zone_fe_adjust").Value
                'Else
                '    .LowerZoneFeAdjust = 0
                'End If

                .DpCrsPbMgoCutoff = RcvryScenDynaset.Fields("dp_crspb_mgo_cutoff").Value
                .DpFnePbMgoCutoff = RcvryScenDynaset.Fields("dp_fnepb_mgo_cutoff").Value
                .DpIpMgoCutoff = RcvryScenDynaset.Fields("dp_ip_mgo_cutoff").Value
                .DpGrind = RcvryScenDynaset.Fields("dp_grind").Value
                .DpAcid = RcvryScenDynaset.Fields("dp_acid").Value
                .DpP2o5 = RcvryScenDynaset.Fields("dp_p2o5").Value
                .DpPa64 = RcvryScenDynaset.Fields("dp_pa64").Value
                .DpFlotMin = RcvryScenDynaset.Fields("dp_flotmin").Value
                .DpTargMgo = RcvryScenDynaset.Fields("dp_targ_mgo").Value

                'New 11/16/2011
                .DpPctWtM200Mesh = RcvryScenDynaset.Fields("dp_pctwt_m200_mesh").Value
                .DpCondMinutes = RcvryScenDynaset.Fields("dp_cond_minutes").Value
                .DpCondPctSolids = RcvryScenDynaset.Fields("dp_cond_pct_solids").Value

                'If Not IsDBNull(RcvryScenDynaset.Fields("wg_feadj_cutoff_date").Value) Then
                '    .WgFeAdjCutoffDate = RcvryScenDynaset.Fields("wg_feadj_cutoff_date").Value
                'Else
                '    .WgFeAdjCutoffDate = #1/1/2001#
                'End If
            End With

            'params = gDBParams
            'params.Add("pRcvryScenarioName", recoveryDefinition.ScenarioName, ORAPARM_INPUT)
            'params("pRcvryScenarioName").serverType = ORATYPE_VARCHAR2
            'params.Add("pProspSetName", recoveryDefinition.ProspectSetName, ORAPARM_INPUT)
            'params("pProspSetName").serverType = ORATYPE_VARCHAR2
            'params.Add("pResult", 0, ORAPARM_OUTPUT)
            'params("pResult").serverType = ORATYPE_CURSOR
            'SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prosp_data_rdctn.get_rcvry_prod_qual(" +
            '          ":pRcvryScenarioName, :pProspSetName, :pResult);end;", ORASQL_FAILEXEC)
            'RcvryScenDynaset = params("pResult").Value
            'ClearParams(params)
            'RecordCount = RcvryScenDynaset.RecordCount
            'SetupDataRdctnProdQualArray(RcvryProdQual)
            'RcvryScenDynaset.MoveFirst()
            'Do While Not RcvryScenDynaset.EOF
            '    With TempData
            '        .MatlTypeName = RcvryScenDynaset.Fields("matl_type_name").Value
            '        .MatlName = RcvryScenDynaset.Fields("matl_name").Value
            '        .SpecLevel = RcvryScenDynaset.Fields("spec_level").Value
            '        .Bpl = RcvryScenDynaset.Fields("bpl_min").Value
            '        .Fe2O3 = RcvryScenDynaset.Fields("fe2o3_max").Value
            '        .Al2O3 = RcvryScenDynaset.Fields("al2o3_max").Value
            '        .Ia = RcvryScenDynaset.Fields("ia_max").Value
            '        .MgO = RcvryScenDynaset.Fields("mgo_max").Value
            '        .CaO = RcvryScenDynaset.Fields("cao_max").Value
            '        .Mer = RcvryScenDynaset.Fields("mer_max").Value
            '        .CaOP2O5 = RcvryScenDynaset.Fields("caop2o5_max").Value
            '        ProdRow = GetProdRow(TempData)
            '        If ProdRow <> 0 Then
            '            RcvryProdQual(ProdRow).MatlTypeName = .MatlTypeName
            '            RcvryProdQual(ProdRow).MatlName = .MatlName
            '            RcvryProdQual(ProdRow).SpecLevel = .SpecLevel
            '            RcvryProdQual(ProdRow).Bpl = .Bpl
            '            RcvryProdQual(ProdRow).Fe2O3 = .Fe2O3
            '            RcvryProdQual(ProdRow).Al2O3 = .Al2O3
            '            RcvryProdQual(ProdRow).Ia = .Ia
            '            RcvryProdQual(ProdRow).MgO = .MgO
            '            RcvryProdQual(ProdRow).CaO = .CaO
            '            RcvryProdQual(ProdRow).Mer = .Mer
            '            RcvryProdQual(ProdRow).CaOP2O5 = .CaOP2O5
            '        End If
            '    End With
            '    RcvryScenDynaset.MoveNext()
            'Loop

            PutParamsOnForm(RcvryData)
            DispFdRcvryCalcs()
            RcvryScenDynaset.Close()
            SetActionStatus("")

        Catch ex As Exception
            Throw ex
        Finally
            _loadingRecoveryScenarioData = False
            ClearParams(params)
            If Not RcvryScenDynaset Is Nothing Then RcvryScenDynaset.Close()
            SetActionStatus("")
        End Try

    End Sub

    Private Function GetProdRow(ByRef aTempData As gDataRdctnProdQualType) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        GetProdRow = 0

        Select Case aTempData.MatlName
            Case Is = "Coarse pebble"
                GetProdRow = 1
            Case Is = "Fine pebble"
                GetProdRow = 2
            Case Is = "IP"
                GetProdRow = 3
            Case Is = "Coarse concentrate"
                GetProdRow = 4
            Case Is = "Fine concentrate"
                GetProdRow = 5
            Case Is = "Total pebble"
                GetProdRow = 6
            Case Is = "Total concentrate"
                GetProdRow = 7
        End Select

        If aTempData.SpecLevel = "Hole" And GetProdRow <> 0 Then
            GetProdRow = GetProdRow + 7
        End If
    End Function

    Private Sub SetupDataRdctnProdQualArray(ByRef aRcvryProdQual() As gDataRdctnProdQualType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim RowIdx As Integer

        For RowIdx = 1 To 14
            aRcvryProdQual(RowIdx).MatlTypeName = "Prospect data reduction matl"
        Next RowIdx

        aRcvryProdQual(1).MatlName = "Coarse pebble"
        aRcvryProdQual(2).MatlName = "Fine pebble"
        aRcvryProdQual(3).MatlName = "IP"
        aRcvryProdQual(4).MatlName = "Coarse concentrate"
        aRcvryProdQual(5).MatlName = "Fine concentrate"
        aRcvryProdQual(6).MatlName = "Total pebble"
        aRcvryProdQual(7).MatlName = "Total concentrate"
        '-----
        aRcvryProdQual(8).MatlName = "Coarse pebble"
        aRcvryProdQual(9).MatlName = "Fine pebble"
        aRcvryProdQual(10).MatlName = "IP"
        aRcvryProdQual(11).MatlName = "Coarse concentrate"
        aRcvryProdQual(12).MatlName = "Fine concentrate"
        aRcvryProdQual(13).MatlName = "Total pebble"
        aRcvryProdQual(14).MatlName = "Total concentrate"

        For RowIdx = 1 To 7
            aRcvryProdQual(RowIdx).SpecLevel = "Split"
        Next RowIdx
        For RowIdx = 8 To 14
            aRcvryProdQual(RowIdx).SpecLevel = "Hole"
        Next RowIdx
    End Sub

    Private Sub PutParamsOnForm(ByVal aRcvryData As gDataRdctnParamsType)



        Dim Chk1Val As Integer
        Dim Chk2Val As Integer
        Dim Chk3Val As Integer
        Dim Chk4Val As Integer


        'fraRecovery Items
        'fraRecovery Items
        'fraRecovery Items
        'With ssVolRcvry
        '    .Col = 1    'Check boxes
        '    .Row = 1
        '    If aRcvryData.OvbVolRcvryMode = "Linear model" Then
        '        .Value = 1
        '    End If
        '    .Col = 1
        '    .Row = 2
        '    If aRcvryData.MineVolRcvryMode = "Linear model" Then
        '        .Value = 1
        '    End If

        '    .Col = 2    'Constant Factor
        '    .Row = 1
        '    .Value = aRcvryData.OvbVolRcvryCf
        '    .Col = 2
        '    .Row = 2
        '    .Value = aRcvryData.MineVolRcvryCf

        '    .Col = 3    'Variable Factor
        '    .Row = 1
        '    .Value = aRcvryData.OvbVolRcvryVf
        '    .Col = 3
        '    .Row = 2
        '    .Value = aRcvryData.MineVolRcvryVf

        '    .Col = 5    'Check boxes
        '    .Row = 1
        '    If aRcvryData.OvbVolRcvryMode = "Footage adjustment" Then
        '        .Value = 1
        '    End If
        '    .Col = 5
        '    .Row = 2
        '    If aRcvryData.MineVolRcvryMode = "Footage adjustment" Then
        '        .Value = 1
        '    End If

        '    .Col = 6    'Footage adjustment
        '    .Row = 1
        '    .Value = aRcvryData.OvbVolRcvryFa
        '    .Col = 6
        '    .Row = 2
        '    .Value = aRcvryData.MineVolRcvryFa
        'End With

        'With ssAdjAssocTons
        '    .Col = 1    'OS
        '    .Row = 1
        '    .Value = aRcvryData.AdjOsTonsWvol
        '    .Col = 2    'Pb
        '    .Value = aRcvryData.AdjPbTonsWvol
        '    .Col = 3    'IP
        '    .Value = aRcvryData.AdjIpTonsWvol
        '    .Col = 4    'Fd
        '    .Value = aRcvryData.AdjFdTonsWvol
        '    .Col = 5    'Cl
        '    .Value = aRcvryData.AdjClTonsWvol
        'End With

        'With ssProdRcvryFctrs
        '    .Col = 3
        '    .Row = 1    'CrsPb ton rcvry
        '    .Value = aRcvryData.PbTonRcvryCrs
        '    .Row = 2    'IP ton rcvry
        '    .Value = aRcvryData.IpTonRcvryTot
        '    .Row = 3    'CrsFd ton rcvry
        '    .Value = aRcvryData.FdTonRcvryCrs
        '    .Row = 4    'CrsFd BPL rcvry
        '    .Value = aRcvryData.FdBplRcvryCrs
        '    .Row = 5    'Clay ton rcvry
        '    .Value = aRcvryData.ClTonRcvryTot

        '    .Col = 5
        '    .Row = 1    'FnePb ton rcvry
        '    .Value = aRcvryData.PbTonRcvryFne
        '    .Row = 3    'FneFd ton rcvry
        '    .Value = aRcvryData.FdTonRcvryFne
        '    .Row = 4    'FneFd BPL rcvry
        '    .Value = aRcvryData.FdBplRcvryFne
        'End With

        'With ssCalcdResults
        '    .Row = 1
        '    .Col = 1    'Crs Conc PL*Ton Recovery
        '    .Value = 0
        '    .Col = 2    'Fne Conc PL*Ton Recovery
        '    .Value = 0
        '    .Col = 4    'Total Fd BPL dilution
        '    .Value = 0
        'End With

        'With ssFlotRcvryLinear
        '    .Row = 1
        '    .Col = 1    'Check box (Linear model)
        '    If aRcvryData.FlotRcvryMode = "Linear model" Then
        '        .Value = 1
        '    End If

        '    .Row = 1
        '    .Col = 5    'Coarse Constant factor (Linear model)
        '    .Value = aRcvryData.FlotRcvryCrsCf
        '    .Row = 1
        '    .Col = 6    'Coarse Variable factor (Linear model)
        '    .Value = aRcvryData.FlotRcvryCrsVf
        '    .Row = 2
        '    .Col = 5    'Fine Constant factor (Linear model)
        '    .Value = aRcvryData.FlotRcvryFneCf
        '    .Row = 2
        '    .Col = 6    'Fine Variable factor (Linear model)
        '    .Value = aRcvryData.FlotRcvryFneVf
        'End With
        'chkTestResultVsLabFlot1.Checked = aRcvryData.LmTest

        'With ssFlotRcvryHardwire
        '    .Row = 1
        '    .Col = 0    'Check box  (Hard-wire)
        '    If aRcvryData.FlotRcvryMode = "Hard-wire" Then
        '        .Value = 1
        '    End If

        '    .Row = 1
        '    .Col = 5    'Coarse Tailings BPL  (Hard-wire)
        '    .Value = aRcvryData.FlotRcvryCrsTlBpl
        '    .Row = 1
        '    .Col = 6    'Coarse Concentrate Insol  (Hard-wire)
        '    .Value = aRcvryData.FlotRcvryCrsCnIns
        '    .Row = 2
        '    .Col = 5    'Fine Tailings BPL  (Hard-wire)
        '    .Value = aRcvryData.FlotRcvryFneTlBpl
        '    .Row = 2
        '    .Col = 6    'Fine Concentrate Insol  (Hard-wire)
        '    .Value = aRcvryData.FlotRcvryFneCnIns
        'End With
        'chkTestResultVsLabFlot2.Checked = aRcvryData.HwTest

        'With ssOtherFlotMethods
        '    .Row = 1
        '    .Col = 1    'Check -- Use lab flotation recovery
        '    If aRcvryData.FlotRcvryMode = "Lab flotation" Then
        '        .Value = 1
        '    End If

        '    .Row = 2
        '    .Col = 1    'Check -- Use SqrRt(Feed BPL) = Tail BPL
        '    If aRcvryData.FlotRcvryMode = "SqrRt feed BPL" Then
        '        .Value = 1
        '    End If
        'End With

        ''fraProdAdj Items
        ''fraProdAdj Items
        ''fraProdAdj Items
        'With ssInsAdj
        '    .Row = 1    'Coarse pebble
        '    .Col = 1    'Minimum
        '    If aRcvryData.CrsPbInsAdjMode = "Minimum" Then
        '        .Value = 1
        '    End If
        '    .Col = 2    'Direct
        '    If aRcvryData.CrsPbInsAdjMode = "Direct" Then
        '        .Value = 1
        '    End If
        '    .Col = 3    'Incremental
        '    If aRcvryData.CrsPbInsAdjMode = "Incremental" Then
        '        .Value = 1
        '    End If
        '    .Col = 4    'In-Situ
        '    If aRcvryData.CrsPbInsAdjMode = "In-Situ" Then
        '        .Value = 1
        '    End If
        '    .Col = 6    'Adjustment value
        '    .Value = aRcvryData.CrsPbInsAdj

        '    .Row = 2    'Fine pebble
        '    .Col = 1    'Minimum
        '    If aRcvryData.FnePbInsAdjMode = "Minimum" Then
        '        .Value = 1
        '    End If
        '    .Col = 2    'Direct
        '    If aRcvryData.FnePbInsAdjMode = "Direct" Then
        '        .Value = 1
        '    End If
        '    .Col = 3    'Incremental
        '    If aRcvryData.FnePbInsAdjMode = "Incremental" Then
        '        .Value = 1
        '    End If
        '    .Col = 4    'In-Situ
        '    If aRcvryData.FnePbInsAdjMode = "In-Situ" Then
        '        .Value = 1
        '    End If
        '    .Col = 6    'Adjustment value
        '    .Value = aRcvryData.FnePbInsAdj

        '    .Row = 3    'IP
        '    .Col = 1    'Minimum
        '    If aRcvryData.IpInsAdjMode = "Minimum" Then
        '        .Value = 1
        '    End If
        '    .Col = 2    'Direct
        '    If aRcvryData.IpInsAdjMode = "Direct" Then
        '        .Value = 1
        '    End If
        '    .Col = 3    'Incremental
        '    If aRcvryData.IpInsAdjMode = "Incremental" Then
        '        .Value = 1
        '    End If
        '    .Col = 4    'In-Situ
        '    If aRcvryData.IpInsAdjMode = "In-Situ" Then
        '        .Value = 1
        '    End If
        '    .Col = 6    'Adjustment value
        '    .Value = aRcvryData.IpInsAdj

        '    .Row = 4    'Coarse concentrate
        '    .Col = 1    'Minimum
        '    If aRcvryData.CrsCnInsAdjMode = "Minimum" Then
        '        .Value = 1
        '    End If
        '    .Col = 2    'Direct
        '    If aRcvryData.CrsCnInsAdjMode = "Direct" Then
        '        .Value = 1
        '    End If
        '    .Col = 3    'Incremental
        '    If aRcvryData.CrsCnInsAdjMode = "Incremental" Then
        '        .Value = 1
        '    End If
        '    .Col = 4    'In-Situ
        '    If aRcvryData.CrsCnInsAdjMode = "In-Situ" Then
        '        .Value = 1
        '    End If
        '    .Col = 6    'Adjustment value
        '    .Value = aRcvryData.CrsCnInsAdj

        '    .Row = 5    'Fine Concentrate
        '    .Col = 1    'Minimum
        '    If aRcvryData.FneCnInsAdjMode = "Minimum" Then
        '        .Value = 1
        '    End If
        '    .Col = 2    'Direct
        '    If aRcvryData.FneCnInsAdjMode = "Direct" Then
        '        .Value = 1
        '    End If
        '    .Col = 3    'Incremental
        '    If aRcvryData.FneCnInsAdjMode = "Incremental" Then
        '        .Value = 1
        '    End If
        '    .Col = 4    'In-Situ
        '    If aRcvryData.FneCnInsAdjMode = "In-Situ" Then
        '        .Value = 1
        '    End If
        '    .Col = 6    'Adjustment value
        '    .Value = aRcvryData.FneCnInsAdj

        '    chkMakeAdjAfterTestForQual.Checked = aRcvryData.AdjInsAfterQualTest
        'End With

        ''fra100PctDefn Items
        ''fra100PctDefn Items
        ''fra100PctDefn Items
        ''May be Nulls that we have to handle!
        'With ssInsAdj100Pct
        '    .Row = 1    'Coarse pebble
        '    .Col = 1    'Minimum
        '    If aRcvryData.CrsPbInsAdjMode100 = "Minimum" Then
        '        .Value = 1
        '    End If
        '    .Col = 2    'Direct
        '    If aRcvryData.CrsPbInsAdjMode100 = "Direct" Then
        '        .Value = 1
        '    End If
        '    .Col = 3    'Incremental
        '    If aRcvryData.CrsPbInsAdjMode100 = "Incremental" Then
        '        .Value = 1
        '    End If
        '    .Col = 4    'In-Situ
        '    If aRcvryData.CrsPbInsAdjMode100 = "In-Situ" Then
        '        .Value = 1
        '    End If
        '    .Col = 6    'Adjustment value
        '    .Value = aRcvryData.CrsPbInsAdj100

        '    .Row = 2    'Fine pebble
        '    .Col = 1    'Minimum
        '    If aRcvryData.FnePbInsAdjMode100 = "Minimum" Then
        '        .Value = 1
        '    End If
        '    .Col = 2    'Direct
        '    If aRcvryData.FnePbInsAdjMode100 = "Direct" Then
        '        .Value = 1
        '    End If
        '    .Col = 3    'Incremental
        '    If aRcvryData.FnePbInsAdjMode100 = "Incremental" Then
        '        .Value = 1
        '    End If
        '    .Col = 4    'In-Situ
        '    If aRcvryData.FnePbInsAdjMode100 = "In-Situ" Then
        '        .Value = 1
        '    End If
        '    .Col = 6    'Adjustment value
        '    .Value = aRcvryData.FnePbInsAdj100

        '    .Row = 3    'IP
        '    .Col = 1    'Minimum
        '    If aRcvryData.IpInsAdjMode100 = "Minimum" Then
        '        .Value = 1
        '    End If
        '    .Col = 2    'Direct
        '    If aRcvryData.IpInsAdjMode100 = "Direct" Then
        '        .Value = 1
        '    End If
        '    .Col = 3    'Incremental
        '    If aRcvryData.IpInsAdjMode100 = "Incremental" Then
        '        .Value = 1
        '    End If
        '    .Col = 4    'In-Situ
        '    If aRcvryData.IpInsAdjMode100 = "In-Situ" Then
        '        .Value = 1
        '    End If
        '    .Col = 6    'Adjustment value
        '    .Value = aRcvryData.IpInsAdj100

        '    .Row = 4    'Coarse concentrate
        '    .Col = 1    'Minimum
        '    If aRcvryData.CrsCnInsAdjMode100 = "Minimum" Then
        '        .Value = 1
        '    End If
        '    .Col = 2    'Direct
        '    If aRcvryData.CrsCnInsAdjMode100 = "Direct" Then
        '        .Value = 1
        '    End If
        '    .Col = 3    'Incremental
        '    If aRcvryData.CrsCnInsAdjMode100 = "Incremental" Then
        '        .Value = 1
        '    End If
        '    .Col = 4    'In-Situ
        '    If aRcvryData.CrsCnInsAdjMode100 = "In-Situ" Then
        '        .Value = 1
        '    End If
        '    .Col = 6    'Adjustment value
        '    .Value = aRcvryData.CrsCnInsAdj100

        '    .Row = 5    'Fine Concentrate
        '    .Col = 1    'Minimum
        '    If aRcvryData.FneCnInsAdjMode100 = "Minimum" Then
        '        .Value = 1
        '    End If
        '    .Col = 2    'Direct
        '    If aRcvryData.FneCnInsAdjMode100 = "Direct" Then
        '        .Value = 1
        '    End If
        '    .Col = 3    'Incremental
        '    If aRcvryData.FneCnInsAdjMode100 = "Incremental" Then
        '        .Value = 1
        '    End If
        '    .Col = 4    'In-Situ
        '    If aRcvryData.FneCnInsAdjMode100 = "In-Situ" Then
        '        .Value = 1
        '    End If
        '    .Col = 6    'Adjustment value
        '    .Value = aRcvryData.FneCnInsAdj100
        'End With

        ''fraMineability Items
        ''fraMineability Items
        ''fraMineability Items
        'With ssSplitPhysMineability
        '    .Row = 1
        '    .Col = 1    'Maximum %clay  Split
        '    .Value = aRcvryData.ClPctMaxSpl
        '    .Row = 2
        '    .Col = 1    'Maximum total depth
        '    .Value = aRcvryData.MaxTotDepthSpl
        '    .Col = 2    'Absolute stop
        '    If aRcvryData.MaxTotDepthModeSpl = "Absolute stop" Then
        '        .Value = 1
        '    End If
        '    .Col = 3    'Finish split
        '    If aRcvryData.MaxTotDepthModeSpl = "Finish split" Then
        '        .Value = 1
        '    End If
        'End With

        'With ssSplitEconMineability
        '    .Row = 1
        '    .Col = 1    'Maximum Mtx-X  Split
        '    .Value = aRcvryData.MtxxMaxSpl
        'End With

        'With ssHolePhysMineability
        '    .Row = 1
        '    .Col = 1    'Maximum %clay  Hole
        '    .Value = aRcvryData.ClPctMaxHole
        '    .Row = 2
        '    .Col = 1    'Minimum ore thickness  Hole
        '    .Value = aRcvryData.MinOreThk
        '    .Row = 3
        '    .Col = 1    'Minimum interburden thickness
        '    .Value = aRcvryData.MinItbThk
        'End With

        'With ssHoleEconMineability
        '    .Row = 1
        '    .Col = 1    'Maximum Mtx-X  Hole
        '    .Value = aRcvryData.MtxxMaxHole
        '    .Row = 2
        '    .Col = 1    'Maximum Tot-X  Hole
        '    .Value = aRcvryData.TotxMaxHole
        '    .Row = 3
        '    .Col = 1    'Minimum total product TPA
        '    .Value = aRcvryData.TotPrTpaMinHole
        'End With

        'chkNoUnmineableHoles.Checked = IIf(aRcvryData.MineFirstSpl = 1, True, False)
        'chkInclCpbAlways.Checked = aRcvryData.InclCpbAlways
        'chkInclFpbAlways.Checked = aRcvryData.InclFpbAlways
        'chkInclOsAlways.Checked = aRcvryData.InclOsAlways
        'chkInclCpbNever.Checked = aRcvryData.InclCpbNever
        'chkInclFpbNever.Checked = aRcvryData.InclFpbNever
        'chkInclOsNever.Checked = aRcvryData.InclOsNever
        'chkCanSelectRejectTpb.Checked = aRcvryData.CanSelectRejectTpb

        'chkMineHasOffSpecPbPlt.Checked = aRcvryData.MineHasOffSpecPbPlt
        '06/14/2010, lss  chkMineHasOffSpecPbPlt not really used anymore!
        'Have chkUseDoloflotPlant and chkUseOrigMgOPlant
        chkUseDoloflotPlant.Checked = aRcvryData.UseDoloflotPlant2010
        chkUseOrigMgoPlant.Checked = aRcvryData.UseOrigMgoPlant
        chkUseDoloflotPlantFco.Checked = aRcvryData.UseDoloflotPlantFco


        'fra100PctFlotRcvry Items
        'fra100PctFlotRcvry Items
        'fra100PctFlotRcvry Items
        'If aRcvryData.FlotRcvryMode100 = "SqrRt feed BPL" Then
        '    opt100PctSqrRtFd.Checked = True
        'End If
        'If aRcvryData.FlotRcvryMode100 = "0 tail BPL" Then
        '    opt100PctTlZero.Checked = True
        'End If


        With ssOffSpecPb
            .Row = 1
            .Col = 1
            .Value = aRcvryData.MplInpBplTarg
            .Row = 2
            .Value = aRcvryData.MplInpMgoTarg
            .Row = 3
            .Value = aRcvryData.MplRejBplTarg
            .Row = 4
            .Value = aRcvryData.MplRejMgoTarg
            '-----
            .Row = 5
            .Value = aRcvryData.MplM1BpltRcvry
            .Row = 6
            .Value = aRcvryData.MplM1BplHwire
            .Row = 7
            .Value = aRcvryData.MplM1InsHwire
            .Row = 8
            .Value = aRcvryData.MplM1MgoImprove
        End With

        'chkCanSelectRejectTcn.Checked = IIf(aRcvryData.CanSelectRejectTcn = True, 1, 0)


        With ssDoloflotPlant
            If chkUseDoloflotPlant.Checked = True Then
                .Col = 1
                .Row = 1
                .Value = aRcvryData.DpFnePbMgoCutoff
                .Row = 2
                .Value = aRcvryData.DpIpMgoCutoff
            Else
                .Col = 1
                .Row = 1
                .Value = 0
                .Row = 2
                .Value = 0
            End If

            .Row = 3
            .Value = aRcvryData.DpGrind

            If chkUseDoloflotPlantFco.Checked <> True Then
                .Row = 4
                .Value = aRcvryData.DpAcid
                .Row = 5
                .Value = aRcvryData.DpP2o5
                .Row = 6
                .Value = aRcvryData.DpPa64
                .Row = 7
                .Value = aRcvryData.DpFlotMin
                .Row = 8
            Else
                .Row = 4
                .Value = 0
                .Row = 5
                .Value = 0
                .Row = 6
                .Value = 0
                .Row = 7
                .Value = 0
            End If

            .Row = 8
            .Value = aRcvryData.DpTargMgo
            ''        .Row = 9
            ''        .Value = aRcvryData.DpAl2O3GreaterThan
            ''        .Row = 10
            ''        .Value = aRcvryData.DpFe2O3GreaterThan
        End With

        With ssDoloflotPlantFco
            If chkUseDoloflotPlantFco.Checked = True Then
                .Col = 1
                .Row = 1
                .Value = aRcvryData.DpFnePbMgoCutoff   'Total Pebble MgO Max
                .Row = 2
                .Value = aRcvryData.DpIpMgoCutoff      'Total Pebble MgO Min
            Else
                .Col = 1
                .Row = 1
                .Value = 0
                .Row = 2
                .Value = 0
            End If
        End With

        With ssDoloflotPlantFco2
            If chkUseDoloflotPlantFco.Checked = True Then
                .Col = 1
                .Row = 1
                .Value = aRcvryData.DpPctWtM200Mesh    '%Wt -200 mesh
                .Row = 2
                .Value = aRcvryData.DpCondMinutes      'Conditioning minutes
                .Row = 3
                .Value = aRcvryData.DpCondPctSolids    'Conditioning %solids
                .Row = 4
                .Value = aRcvryData.DpFlotMin          'Flotation minutes
                .Row = 5
                .Value = aRcvryData.DpPa64             'PA64 lbs/ton
                .Row = 6
                .Value = aRcvryData.DpP2o5             'Phos acid lbs/ton
                .Row = 7
                .Value = aRcvryData.DpAcid             'Sulfuric acid lbs/ton
            Else
                .Col = 1
                .Row = 1
                .Value = 0
                .Row = 2
                .Value = 0
                .Row = 3
                .Value = 0
                .Row = 4
                .Value = 0
                .Row = 5
                .Value = 0
                .Row = 6
                .Value = 0
                .Row = 7
                .Value = 0
            End If
        End With

    End Sub

    Private Sub cmdLoadOverrideTxtFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLoadOverrideTxtFile.Click
        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SetActionStatus("Loading split override text file...")
        Me.Cursor = Cursors.WaitCursor

        txtSplitOverrideName.Text = ""
        cboSplitOverrideMineName.Text = "None"

        gLoadOverrideTxtFile(txtSplOverrideTxtFile.Text, ssSplitOverride)

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub cmdApplySplOverrides_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdApplySplOverrides.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SetActionStatus("Applying split overrides...")
        Me.Cursor = Cursors.WaitCursor


        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub cmdCancelSplitOverride_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancelSplitOverride.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        ClearSplitOverride()
    End Sub


    Private Sub cmdSaveSplitOverride_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveSplitOverride.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        If MsgBox("Save split override set?", vbYesNo +
                  vbDefaultButton1, "Save Split Override Set Definition") = vbYes Then
            SaveSplitOverrideSet("User split override set")
            ClearSplitOverride()
            GetSplitOverrideSets()
        End If
    End Sub

    Private Sub SaveSplitOverrideSet(ByVal aProspSetName As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo SaveSplitOverrideSetError

        Dim ItemCount As Integer
        Dim RowIdx As Integer
        Dim InsertSQL As String
        Dim MineName As String

        Dim Twp() As Integer
        Dim Rge() As Single
        Dim Sec() As Single
        Dim Hole() As String
        Dim Split() As Single
        Dim Mineability() As String

        Dim ThisTwp As Integer
        Dim ThisRge As Integer
        Dim ThisSec As Integer
        Dim ThisHole As String
        Dim ThisHoleLoc As String
        Dim ThisSplit As Integer
        Dim ThisMineable As String

        'Check for some problems
        If Trim(txtSplitOverrideName.Text) = "" Then
            MsgBox("You MUST enter a split override set name!" & vbCrLf &
                   "This split override set has NOT been saved.",
                   vbOKOnly + vbExclamation,
                   "Split Override Set Problem")
            Exit Sub
        End If

        If cboSplitOverrideMineName.Text = "None" Then
            MineName = ""
        Else
            MineName = cboSplitOverrideMineName.Text
        End If

        'Determine item count
        ItemCount = ssSplitOverride.MaxRows

        If ItemCount > 0 Then
            ReDim Twp(ItemCount - 1)
            ReDim Rge(ItemCount - 1)
            ReDim Sec(ItemCount - 1)
            ReDim Hole(ItemCount - 1)
            ReDim Split(ItemCount - 1)
            ReDim Mineability(ItemCount - 1)
        Else
            'Nothing to update!
            Exit Sub
        End If

        SetActionStatus("Saving split override set name...")
        Me.Cursor = Cursors.WaitCursor

        'Now get the data into the transfer arrays.
        ItemCount = 0
        With ssSplitOverride
            For RowIdx = 1 To .MaxRows
                .Row = RowIdx
                .Col = 1
                ThisHoleLoc = .Text     'tt-rr-ss hole  tt = township
                '               rr = range
                '               ss = section

                ThisTwp = Mid(ThisHoleLoc, 1, 2)
                ThisRge = Mid(ThisHoleLoc, 4, 2)
                ThisSec = Mid(ThisHoleLoc, 7, 2)
                ThisHole = Mid(ThisHoleLoc, 10)

                .Col = 2
                ThisSplit = .Value
                .Col = 3
                ThisMineable = .Text    'M, U, C

                'Need to get data into Twp(), Rge(), Sec(), Hole(),
                '                      SplitNumber(), Mineability()
                Twp(ItemCount) = ThisTwp
                Rge(ItemCount) = ThisRge
                Sec(ItemCount) = ThisSec
                Hole(ItemCount) = ThisHole
                Split(ItemCount) = ThisSplit
                Mineability(ItemCount) = ThisMineable

                ItemCount = ItemCount + 1
            Next RowIdx
        End With

        'PROCEDURE update_prosp_split_oride_set
        'pArraySize             IN     INTEGER,
        'pSplitOrideSetName     IN     VARCHAR2,
        'pProspSetName          IN     VARCHAR2,
        '--
        'pWhoDefined            IN     VARCHAR2,
        'pWhenDefined           IN     DATE,
        'pMineName              IN     VARCHAR2,
        '--
        'pTownship              IN     NUMBERARRAY,
        'pRange                 IN     NUMBERARRAY,
        'pSection               IN     NUMBERARRAY,
        'pHoleLocation          IN     VCHAR2ARRAY4,
        'pSplitNumber           IN     NUMBERARRAY,
        'pMineability           IN     VCHAR2ARRAY1,
        'pResult                IN OUT NUMBER)
        InsertSQL = "Begin mois.mois_prosp_data_rdctn.update_prosp_split_oride_set(" &
        "   :pArraySize, " &
        "   :pSplitOrideSetName, " &
        "   :pProspSetName, " &
        "   :pWhoDefined, " &
        "   :pWhenDefined, " &
        "   :pMineName, " &
        "   :pTownship, " &
        "   :pRange, " &
        "   :pSection, " &
        "   :pHoleLocation, " &
        "   :pSplitNumber, " &
        "   :pMineability, " &
        "   :pResult); " &
        "end;"
        '13
        Dim arA1() As Object = {"pArraySize", ItemCount, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA2() As Object = {"pSplitOrideSetName", txtSplitOverrideName.Text, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA3() As Object = {"pProspSetName", aProspSetName, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA4() As Object = {"pWhoDefined", gUserName.ToLower, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA5() As Object = {"pWhenDefined", Now, ORAPARM_INPUT, ORATYPE_DATE}
        Dim arA6() As Object = {"pMineName", MineName, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA7() As Object = {"pTownship", Twp, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA8() As Object = {"pRange", Rge, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA9() As Object = {"pSection", Sec, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA10() As Object = {"pHoleLocation", Hole, ORAPARM_INPUT, ORATYPE_VARCHAR2, 4}
        Dim arA11() As Object = {"pSplitNumber", Split, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA12() As Object = {"pMineability", Mineability, ORAPARM_INPUT, ORATYPE_VARCHAR2, 1}
        Dim arA13() As Object = {"pResult", 0, ORAPARM_INPUT, ORATYPE_NUMBER}

        'RunBatchSP(InsertSQL, _
        '    Array("pArraySize", ItemCount, ORAPARM_INPUT, ORATYPE_NUMBER), _

        '    Array("pSplitOrideSetName", txtSplitOverrideName.Text, ORAPARM_INPUT, ORATYPE_VARCHAR2), _

        '    Array("pProspSetName", aProspSetName, ORAPARM_INPUT, ORATYPE_VARCHAR2), _

        '    Array("pWhoDefined", gUserName, ORAPARM_INPUT, ORATYPE_VARCHAR2), _

        '    Array("pWhenDefined", Now, ORAPARM_INPUT, ORATYPE_DATE), _

        '    Array("pMineName", MineName, ORAPARM_INPUT, ORATYPE_VARCHAR2), _

        '    Array("pTownship", Twp(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _

        '    Array("pRange", Rge(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _

        '    Array("pSection", Sec(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _

        '    Array("pHoleLocation", Hole(), ORAPARM_INPUT, ORATYPE_VARCHAR2, 4), _

        '    Array("pSplitNumber", Split(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _

        '    Array("pMineability", Mineability(), ORAPARM_INPUT, ORATYPE_VARCHAR2, 1), _
        '    Array("pResult", 0, ORAPARM_INPUT, ORATYPE_NUMBER))
        RunBatchSP(InsertSQL,
                arA1,
                arA2,
                arA3,
                arA4,
                arA5,
                arA6,
                arA7,
                arA8,
                arA9,
                arA10,
                arA11,
                arA12,
                arA13)

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow

        MsgBox("Split override set saved.", vbOKOnly, "Save Status")

        Exit Sub

SaveSplitOverrideSetError:
        MsgBox("Error while saving." & Str(Err.Number) &
               Chr(10) & Chr(10) &
               Err.Description, vbExclamation,
               "Update Error")

        On Error Resume Next
        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub cmdDeleteSplitOverride_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteSplitOverride.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim CommStr As String

        CommStr = "Are you really sure you want " &
                  "to delete this split override set?" &
                  vbCrLf & vbCrLf &
                  "(Please be careful!!)"

        If MsgBox(CommStr, vbYesNo +
                  vbDefaultButton1, "Delete Split Override Set") = vbYes Then
            DeleteSplitOverrideSet("User split override set")
            ClearSplitOverride()
            GetSplitOverrideSets()
        End If
    End Sub

    Private Sub DeleteSplitOverrideSet(ByVal aProspSetName As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo DeleteSplitOverrideSetError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        Me.Cursor = Cursors.WaitCursor
        SetActionStatus("Deleting split override set...")

        'Set 
        params = gDBParams

        params.Add("pSplitOrideSetName", txtSplitOverrideName.Text, ORAPARM_INPUT)
        params("pSplitOrideSetName").serverType = ORATYPE_VARCHAR2

        params.Add("pProspSetName", aProspSetName, ORAPARM_INPUT)
        params("pProspSetName").serverType = ORATYPE_VARCHAR2

        'Procedure delete_prosp_split_oride_set
        'pSplitOrideSetName  IN     VARCHAR2,
        'pProspSetName       IN     VARCHAR2)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prosp_data_rdctn.delete_prosp_split_oride_set(" +
                      ":pSplitOrideSetName, :pProspSetName);end;", ORASQL_FAILEXEC)

        ClearParams(params)
        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
        Exit Sub

DeleteSplitOverrideSetError:
        MsgBox("Error deleting split override set." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Split Override Set Deletion Error")

        On Error Resume Next
        ClearParams(params)
        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub cmdGetSplitOverrides_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetSplitOverrides.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        GetSplitOverrideSets()
    End Sub

    Private Sub GetSplitOverrideSets()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetSplitOverrideSetsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim OrideDynaset As OraDynaset
        Dim UserName As String
        Dim ThisSplitOrideSetName As String
        Dim ThisWhoDefnd As String
        Dim ThisWhenDefnd As Date
        Dim ThisMineName As String
        Dim RecordCount As Integer

        ssSplitOverrides.MaxRows = 0

        If chkOnlyMySplitOverride.Checked = True Then
            UserName = gUserName.ToLower
        Else
            UserName = "All"
        End If

        params = gDBParams

        params.Add("pUserName", UserName, ORAPARM_INPUT)
        params("pUserName").serverType = ORATYPE_VARCHAR2

        params.Add("pProspSetName", "User split override set", ORAPARM_INPUT)
        params("pProspSetName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_split_oride_set_all
        'pUserName           IN     VARCHAR2,
        'pProspSetName       IN     VARCHAR2,
        'pResult             IN OUT c_splitoride)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prosp_data_rdctn.get_prosp_split_oride_set_all (" &
                                             ":pUserName, :pProspSetName, :pResult);end;", ORASQL_FAILEXEC)
        OrideDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = OrideDynaset.RecordCount

        OrideDynaset.MoveFirst()
        Do While Not OrideDynaset.EOF
            ThisSplitOrideSetName = OrideDynaset.Fields("split_oride_set_name").Value
            ThisWhoDefnd = OrideDynaset.Fields("who_defined").Value
            ThisWhenDefnd = OrideDynaset.Fields("when_defined").Value

            If Not IsDBNull(OrideDynaset.Fields("mine_name").Value) Then
                ThisMineName = OrideDynaset.Fields("mine_name").Value
            Else
                ThisMineName = ""
            End If

            With ssSplitOverrides
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 1
                .Text = ThisSplitOrideSetName
                .Col = 2
                .Text = ThisWhoDefnd
                .Col = 3
                .Text = Format(ThisWhenDefnd, "MM/dd/yy")
                .Col = 4
                .Text = ThisMineName
            End With
            OrideDynaset.MoveNext()
        Loop

        OrideDynaset.Close()

        Exit Sub

GetSplitOverrideSetsError:
        On Error Resume Next

        MsgBox("Error getting split override set." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Split Override Sets Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        OrideDynaset.Close()
    End Sub

    Private Sub ssSplitOverrides_Click(ByVal sender As System.Object, ByVal e As AxFPSpread._DSpreadEvents_ClickEvent) Handles ssSplitOverrides.ClickEvent '(ByVal Col As Long, ByVal Row As Long)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ThisSplitOrideSetName As String

        If e.row <> 0 Then
            With ssSplitOverrides
                .Row = e.row
                .Col = 1
                ThisSplitOrideSetName = .Text
            End With
            lblGen55.Text = ""
            'frmProspDataReduction.Refresh()

            DisplaySplitOverrideSet(ThisSplitOrideSetName, "User split override set")
            MarkHolesGreen(ssSplitOverride, True)
            gHaveRawProspData = False
            chkUseRawProspAsOverride.Checked = False
        End If
    End Sub

    Private Sub DisplaySplitOverrideSet(ByVal aSplitOrideSetName As String,
                                        ByVal aProspSetName As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo DisplaySplitOverrideSetError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim OrideDynaset As OraDynaset
        Dim RecordCount As Integer
        Dim ItemCount As Integer

        Dim SplitOrideSetName As String
        Dim MineName As String
        Dim ThisTwp As Integer
        Dim ThisRge As Integer
        Dim ThisSec As Integer
        Dim ThisHole As String
        Dim ThisSplit As Integer
        Dim ThisMineability As String
        Dim ThisHoleLoc As String

        SetActionStatus("Getting split override set...")
        Me.Cursor = Cursors.WaitCursor

        ssSplitOverride.MaxRows = 0

        ' Set 
        params = gDBParams

        params.Add("pSplitOrideSetName", aSplitOrideSetName, ORAPARM_INPUT)
        params("pSplitOrideSetName").serverType = ORATYPE_VARCHAR2

        params.Add("pProspSetName", aProspSetName, ORAPARM_INPUT)
        params("pProspSetName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_split_oride_set
        'pSplitOrideSetName  IN     VARCHAR2,
        'pProspSetName       IN     VARCHAR2,
        'pResult             IN OUT c_splitoride)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prosp_data_rdctn.get_prosp_split_oride_set(" +
                                       ":pSplitOrideSetName, :pProspSetName, :pResult);end;", ORASQL_FAILEXEC)

        OrideDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = OrideDynaset.RecordCount
        ItemCount = 0

        OrideDynaset.MoveFirst()
        Do While Not OrideDynaset.EOF
            ItemCount = ItemCount + 1
            If ItemCount = 1 Then
                SplitOrideSetName = OrideDynaset.Fields("split_oride_set_name").Value

                If Not IsDBNull(OrideDynaset.Fields("mine_name").Value) Then
                    MineName = OrideDynaset.Fields("mine_name").Value
                Else
                    MineName = "None"
                End If

                txtSplitOverrideName.Text = SplitOrideSetName
                cboSplitOverrideMineName.Text = MineName
            End If

            ThisTwp = OrideDynaset.Fields("township").Value
            ThisRge = OrideDynaset.Fields("range").Value
            ThisSec = OrideDynaset.Fields("section").Value
            ThisHole = OrideDynaset.Fields("hole_location").Value
            ThisSplit = OrideDynaset.Fields("split_number").Value
            ThisMineability = OrideDynaset.Fields("mineability").Value

            With ssSplitOverride
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 1
                .Text = gGetHoleLocationTrs(ThisSec,
                                            ThisTwp,
                                            ThisRge,
                                            ThisHole)
                .Col = 2
                .Value = ThisSplit
                .Col = 3
                .Text = ThisMineability    'M, U, C
            End With
            OrideDynaset.MoveNext()
        Loop

        OrideDynaset.Close()
        SetActionStatus("")
        Me.Cursor = Cursors.Arrow

        Exit Sub

DisplaySplitOverrideSetError:
        MsgBox("Error getting data." & vbCrLf &
        Err.Description,
        vbOKOnly + vbExclamation,
        "Data Get Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        OrideDynaset.Close()
        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub cmdGenerateProspectDataSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGenerateProspectDataset.Click

        Dim RcvryData As gDataRdctnParamsType
        Dim CommStr As String
        'Dim DataSetStatus As Integer
        Dim ProspectResult As SplitResultSet
        Dim AreaDefnData As gAreaDefnParamsType
        Dim NoReview As Boolean
        Dim SaveType As String
        Dim MineHasOffSpecPbPlt As Boolean
        Dim MineHasDoloflotPlt As Boolean
        Dim MsgResponse As Integer
        Dim RowIdx As Long

        Dim RecoveryParameters As ViewModels.ProductRecoveryDefinition = _recoveryScenariosForm.ProductRecoveryDefinition


        'Will hard-code this for now -- there is really only one defined
        'distribution that is used that I have called Standard2006.  It has
        '85 size fraction codes that are distributed by percent among 21 size
        'fraction codes.  I will save numbers as strings.  Added 2 extra rows
        'and 1 extra column for header stuff.
        Dim SfcReproData(0 To 87, 0 To 22) As String

        If chkCreateOutputOnly.Checked = True Then
            NoReview = True
        Else
            NoReview = False
        End If

        lblProcComm0.Text = ""
        lblProcComm1.Text = ""
        lblProcComm2.Text = ""
        lblGen41.Text = ""   'Was lblReviewComm
        lblGen64.Text = ""   'Was lblRowNum
        ClearDetlDisp()
        cmdHoleSplitRpt.Enabled = False
        cmdReportAll.Enabled = True
        optProdCoeff.Enabled = True
        opt100Pct.Enabled = True
        'frmProspDataReduction.Refresh()

        If chkSaveToDatabase.Checked Then
            SaveType = "Database"
        End If
        If chkSurvCaddTextfile.Checked Then
            SaveType = "SurvCaddText"
        End If
        If chkSpecMoisTransferFile.Checked Then
            SaveType = "MoisText"
        End If
        If chkBdFormatTextfile.Checked Then
            SaveType = "BdFormatText"
        End If

        'lblGen65) was lblOffSpecPbMgPlt
        If lblGen65.Text = "*OffSpec Pb Mg Plt*" Or
        lblGen65.Text = "*Doloflot Plt FCO*" Then
            MineHasOffSpecPbPlt = True
        Else
            MineHasOffSpecPbPlt = False
        End If

        If NoReview Then
            'Check comma-delimited text file name
            If Trim(txtProspDatasetTextfileName.Text).Equals(String.Empty) OrElse
                Mid(txtProspDatasetTextfileName.Text, Len(txtProspDatasetTextfileName.Text)) = "\" Then
                MsgBox("You must enter a SurvCADD/MOIS transfer textfile name " &
                       vbCrLf & "(See the 'Output' tab at the top of this form)." +
                        Chr(10) + Chr(10) + "Reduction/Save process cannot begin!",
                        vbExclamation, "Reduction/Save Process Error")
                Exit Sub
            End If
            CommStr = "Generate a prospect dataset?" & vbCrLf & vbCrLf &
                      "(Data review will not be available -- you will " &
                      "be saving this data to a text file or a database.)"
        Else
            CommStr = "Generate a prospect dataset?" & vbCrLf & vbCrLf &
                      "(Split data will be created and you can review " &
                      "it under the ""Review"" tab before actually " &
                      "saving this data to a text file or a database.)"
        End If

        tabMain.SelectedTab = tabReview

        If MsgBox(CommStr, vbYesNo +
                  vbDefaultButton1, "Generate Prospect Dataset") = vbYes Then
            SetActionStatus("Generating prospect dataset (split data for review)...")
            Me.Cursor = Cursors.WaitCursor

            'Check the product sizes selected by the user.
            If Not ProductSizeOk Then
                CommStr = "Cannot generate a prospect dataset!" & vbCrLf & vbCrLf &
                          "Product Sizes are not correct!  Every size fraction " &
                          "code must be assigned to at least one material in " &
                          "the size fraction assignment grid under the 'Product Sizes' " &
                          "tab."
                MsgBox(CommStr, vbOKOnly, "Product Size Definition Problem")
                SetActionStatus("")
                Me.Cursor = Cursors.Arrow
                Exit Sub
            End If

            'Check the area definition.
            If Not AreaDefnOk() Then
                CommStr = "Cannot generate a prospect dataset!" & vbCrLf & vbCrLf &
                          "An area has not been selected under the 'Area Definition' " &
                          "tab."
                MsgBox(CommStr, vbOKOnly, "Area Definition Problem")
                SetActionStatus("")
                Me.Cursor = Cursors.Arrow
                Exit Sub
            End If


            ssSplitReview.MaxRows = 0

            ssCompReview.MaxRows = 0
            ssCompErrors.MaxRows = 0
            With ssResultCnt
                .Col = 1
                .Row = 1
                .Value = 0
                .Row = 2
                .Value = 0
                .Row = 3
                .Value = 0
                .Row = 4
                .Value = 0
                .Row = 5
                .Value = 0
                .Row = 6
                .Value = 0
                .Row = 7
                .Value = 0
            End With

            fProcessing = True

            'Get recovery, etc. information
            GetRcvryEtcParamsFromForm(RcvryData)

            '06/15/2010, lss  RcvryData.MineHasOffSpecPbPlt not really used anymore.
            'If RcvryData.MineHasOffSpecPbPlt = True Then

            'If RcvryData.MineHasOffSpecPbPlt = True then RcvryData.UseDoloflotPlant2010 = False
            If RcvryData.UseOrigMgoPlant Then
                'lblGen65) was lblOffSpecPbMgPlt
                lblGen65.Text = "*OffSpec Pb Mg Plt*"
                MineHasOffSpecPbPlt = True
            Else
                If RcvryData.UseDoloflotPlant2010 Then
                    'lblGen65) was lblOffSpecPbMgPlt
                    lblGen65.Text = "*Doloflot Plt Ona*"
                    MineHasDoloflotPlt = True
                Else
                    If RcvryData.UseDoloflotPlantFco = True Then
                        'lblGen65) was lblOffSpecPbMgPlt
                        lblGen65.Text = "*Doloflot Plt FCO*"
                        MineHasOffSpecPbPlt = True
                    Else
                        'lblGen65) was lblOffSpecPbMgPlt
                        lblGen65.Text = ""
                        MineHasDoloflotPlt = False
                    End If
                End If
            End If

            'Get prospect raw material size definition (% distribution of SFC's)
            GetProspRawMatlSizeDefn(_productSizeDesginationForm.ProductSizeDesignation.SizeFractionDistribution,
                                    SfcReproData)

            If RcvryData.UseOrigMgoPlant = True Then
                'Check to see if the MgO plant definition meshes correctly with the
                'product quality definition for split pebble.
                Dim SplitBPLQualityLimits As ViewModels.ProductQualitySpecification = RecoveryParameters.SplitQualitySpecifications.SingleOrDefault(Function(t) t.Element = ViewModels.ElementExtensions.DisplayName(ViewModels.Element.BPL))
                Dim SplitMgOQualityLimits As ViewModels.ProductQualitySpecification = RecoveryParameters.SplitQualitySpecifications.SingleOrDefault(Function(t) t.Element = ViewModels.ElementExtensions.DisplayName(ViewModels.Element.MGO))

                If (RcvryData.MplInpBplTarg <> SplitBPLQualityLimits.CoarsePebbleValue And RcvryData.MplInpBplTarg <> 0) Or
                   (RcvryData.MplInpMgoTarg <> SplitMgOQualityLimits.CoarsePebbleValue And RcvryData.MplInpMgoTarg <> 0) Then

                    'Things are not meshed correctly!
                    CommStr = "WARNING only -- prospect dataset can still be generated!  To continue click " &
                              "'OK', to stop click 'Cancel'." & vbCrLf & vbCrLf &
                              "The 'Good' pebble quality and the MgO plant input quality" &
                              " definition do not match up correctly ('Good' split pebble 'Min BPL'" &
                              " should be the same as the 'MgO plant input BPL <' and the 'Good'" &
                              " split pebble 'Max MgO' should be the same as the 'MgO plant input MgO >')."
                    MsgResponse = MsgBox(CommStr, vbOKCancel, "Pebble Quality Problem")

                    If MsgResponse = vbCancel Then
                        SetActionStatus("")
                        Me.Cursor = Cursors.Arrow
                        fProcessing = False
                        Exit Sub
                    End If
                End If
            End If

            If RcvryData.UseDoloflotPlant2010 = True Then
                'Check to see if the Doloflot plant definition meshes correctly with the
                'product quality definition for Fine pebble and IP.
                Dim SplitMgOQualityLimits As ViewModels.ProductQualitySpecification = RecoveryParameters.SplitQualitySpecifications.SingleOrDefault(Function(t) t.Element = ViewModels.ElementExtensions.DisplayName(ViewModels.Element.MGO))

                If (RcvryData.DpFnePbMgoCutoff <> SplitMgOQualityLimits.FinePebbleValue And
                    RcvryData.DpFnePbMgoCutoff <> 0) Or
                    (RcvryData.DpIpMgoCutoff <> SplitMgOQualityLimits.IpValue And
                    RcvryData.DpIpMgoCutoff <> 0) Then

                    'Things are not meshed correctly!
                    CommStr = "WARNING only -- prospect dataset can still be generated!  To continue click " &
                              "'OK', to stop click 'Cancel'." & vbCrLf & vbCrLf &
                              "The 'Good' Fine pebble quality and the Doloflot plant FnePb MgO cutoff &/or " &
                              "the 'Good' IP quality and the Doloflot plant IP MgO cutoff do not match up correctly " &
                              "('Good' split Fine pebble 'Max MgO'" &
                              " should be the same as the Doloflot 'FnePb MgO cutoff' and the 'Good'" &
                              " split IP 'Max MgO' should be the same as the Doloflot 'IP MgO cutoff')."
                    MsgResponse = MsgBox(CommStr, vbOKCancel, "Fine pebble &/or IP Quality Problem")

                    If MsgResponse = vbCancel Then
                        SetActionStatus("")
                        Me.Cursor = Cursors.Arrow
                        fProcessing = False
                        Exit Sub
                    End If
                End If
            End If

            If RcvryData.UseDoloflotPlantFco = True Then
                'Maybe add something here?
            End If

            'Check to see if Density setup has a problem.
            'If RcvryData.DensCalcMode = "Limit routine" And RcvryData.DensLowerLimit = RcvryData.DensUpperLimit Then
            If RecoveryParameters.DensityCalculationMode = "Limit routine" AndAlso RecoveryParameters.DensityLowerLimit = RecoveryParameters.DensityUpperLimit Then
                'Density setup problem?
                CommStr = "WARNING only -- prospect dataset can still be generated!  To continue click " &
                          "'OK', to stop click 'Cancel'." & vbCrLf & vbCrLf &
                          "Density lower limit = Density upper limit."
                MsgResponse = MsgBox(CommStr, vbOKCancel, "Density Problem")

                If MsgResponse = vbCancel Then
                    SetActionStatus("")
                    Me.Cursor = Cursors.Arrow
                    fProcessing = False
                    Exit Sub
                End If
            End If

            'Process raw prospect data -- place data in ssSplitReview and
            'ssCompReview
            If NoReview Then
                'ssSplitReview.Visible = False
                tbcSplitResults.Visible = False
                tbcHoleResults.Visible = False
                lblNoReview.Visible = True
                cmdCopyToOverrides.Visible = False
                lblBarrenSplComm.Visible = False
                lblGen25.Visible = False
            Else
                'ssSplitReview.Visible = True
                tbcSplitResults.Visible = True
                tbcHoleResults.Visible = True
                lblNoReview.Visible = False
                cmdCopyToOverrides.Visible = True
                lblBarrenSplComm.Visible = True
                lblGen25.Visible = True
            End If

            gFileNumber = -99
            gOutputLines = New List(Of String)
            'gGenerateProspectDataset = 1   OK
            'gGenerateProspectDataset = 2   User escaped
            'gGenerateProspectDataset = 3   Problems

            'Public Function gGenerateProspectDataset
            ' 1) aAreaDefnData As gAreaDefnParamsType,               AreaDefnData
            ' 2) aSsAreaTrsCorner As vaSpread, _                     ssAreaTrsCorner
            ' 3) aSsAreaXyCoord As vaSpread,                         ssAreaXyCoord
            ' 4) aSsAreaByHoleSelect As vaSpread, _                  ssAreaByHoleSelect
            ' 5) aSsProdDist As vaSpread,                            ssProdDist
            ' 6) aRcvryParamsData As gDataRdctnParamsType, _         RcvryData
            ' 7) aRcvryProdQual() As gDataRdctnProdQualType, _       RcvryProdQual()
            ' 8) aSsSplitReview As vaSpread, _                       ssSplitReview
            ' 9) aSsCompReview As vaSpread, _                        ssCompReview
            '10) aSplitOverrideName As String, _                     txtSplitOverrideName.Text
            '11) aSsRawProspMin As vaSpread, _                       ssRawProspMin
            '12) aSfcReproData() As String, _                        SfcReproData()
            '13) aRawProspDynaset As OraDynaset, _                   RawProspDynaset
            '14) aScope As String, _                                 "Batch"
            '15) aNoReview As Boolean, _                             NoReview
            '16) aSaveType As String, _                              SaveType
            '17) aMineHasOffSpecPbPlt As Boolean, _                  MineHasOffSpecPbPlt
            '18) aProspectDatasetName As String, _                   txtProspectDatasetName.Text
            '19) aProspDatasetTextFileName As String, _              txtProspDatasetTextfileName
            '20) aChk100Pct As Integer, _                            chk100Pct.Value
            '21) aChkProductionCoefficient As Integer, _             chkProductionCoefficient.Value
            '22) aOptInclSplits As Boolean, _                        optInclSplits.Value
            '23) aOptInclComposites As Boolean,                      optInclComposites.Value
            '24) aOptInclBoth As Boolean, _                          optInclBoth.Value
            '25) aChkInclMgPlt As Integer, _                         chkInclMgPlt.Value
            '26) aUseOrigHole As Boolean, _                          False
            '27) aMineHasDoloflotPlt As Boolean)                     MineHasDoloflotPlt

            ProspectResult = gGenerateProspectDataset(_areaDefinitionForm.AreaDefinition,
                                                     _productSizeDesginationForm.ProductSizeDesignation,
                                                     RcvryData, RecoveryParameters,
                                                     ssSplitReview, ssCompReview, txtSplitOverrideName.Text,
                                                     ssRawProspMin, SfcReproData,
                                                     "Batch", NoReview,
                                                     SaveType, MineHasOffSpecPbPlt,
                                                     txtProspectDatasetName.Text,
                                                     txtProspDatasetTextfileName.Text,
                                                     IIf(chk100Pct.Checked, 1, 0),
                                                     IIf(chkProductionCoefficient.Checked, 1, 0),
                                                     optInclSplits.Checked,
                                                     optInclComposites.Checked,
                                                     optInclBoth.Checked,
                                                     IIf(chkInclMgPlt.Checked, 1, 0),
                                                     False, MineHasDoloflotPlt)

            SetActionStatus("Summing results...")
            SumResults()

            If NoReview Then
                On Error Resume Next
                ' Close #gFileNumber
            End If

            MarkHolesGreen(ssSplitReview, False)
            With ssRawProspMin
                .Row = 0
                .Col = 7
                .Text = " "
                .set_ColWidth(.Col, 0.17)
                For RowIdx = 0 To .MaxRows
                    .Row = RowIdx
                    .CellType = FPSpread.CellTypeConstants.CellTypeStaticText ' SS_CELL_TYPE_STATIC_TEXT
                    .BackColor = Color.Black 'vbBlack
                Next
            End With

            'gGenerateProspectDataset = 1   OK
            'gGenerateProspectDataset = 2   User escaped
            'gGenerateProspectDataset = 3   Problems

            SetActionStatus("")
            Me.Cursor = Cursors.Arrow

            Select Case ProspectResult.IntResult
                Case Is = 1
                    'Filling Split Product grids
                    Dim Tabs As New List(Of Control)
                    For Each ControlTab As Control In Me.tbcSplitResults.Controls
                        If ControlTab.Name.Contains("tab") Then
                            Tabs.Add(ControlTab)
                        End If
                    Next
                    For Each ControlTab2 As Control In Tabs
                        Me.tbcSplitResults.Controls.Remove(ControlTab2)
                    Next

                    With tbcSplitResults.Controls
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.CPb))
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.CpbRej))
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.FPb))
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.FPbRej))
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.IP))
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.IPRej))
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.CCn))
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.FCn))
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.TCn))
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.Os))
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.TPb))
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.TPbRej))
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.Tpr))
                        .Add(RdctnProductTab(ProspectResult.SplitResults, ViewModels.ProductTypes.ATpr))
                    End With

                    'Filling Hole Product grids
                    Dim HoleTabs As New List(Of Control)
                    For Each ControlTab As Control In Me.tbcHoleResults.Controls
                        If ControlTab.Name.Contains("tab") Then
                            HoleTabs.Add(ControlTab)
                        End If
                    Next
                    For Each ControlTab2 As Control In HoleTabs
                        Me.tbcHoleResults.Controls.Remove(ControlTab2)
                    Next

                    With tbcHoleResults.Controls
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.CPb))
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.CpbRej))
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.FPb))
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.FPbRej))
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.IP))
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.IPRej))
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.CCn))
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.FCn))
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.TCn))
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.Os))
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.TPb))
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.TPbRej))
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.Tpr))
                        .Add(RdctnProductTab(ProspectResult.HoleResults, ViewModels.ProductTypes.ATpr))
                    End With


                    MsgBox("Prospect dataset generation completed -- no problems!", vbOKOnly, "Create Status")

                Case Is = 2
                    lblGen41.Text = ""   'Was lblReviewComm
                    lblGen64.Text = ""   'Was lblRowNum
                    ClearDetlDisp()
                    cmdHoleSplitRpt.Enabled = False
                    optProdCoeff.Enabled = True
                    opt100Pct.Enabled = True
                    MsgBox("User terminated process.", vbOKOnly,
                           "Process Terminated")

                Case Else
                    lblGen41.Text = ""   'Was lblReviewComm
                    lblGen64.Text = ""   'Was lblRowNum
                    ClearDetlDisp()
                    cmdHoleSplitRpt.Enabled = False
                    optProdCoeff.Enabled = True
                    opt100Pct.Enabled = True
                    MsgBox("Prospect dataset generation completed -- PROBLEMS!", vbOKOnly, "Create Status")
            End Select
            lblProcComm0.Text = ""
            lblProcComm1.Text = ""
            lblProcComm2.Text = ""
        End If

        fProcessing = False
    End Sub

    Private Function RdctnProductTab(ByVal rawProspResults As List(Of gRawProspSplRdctnType), ByVal prodType As ViewModels.ProductTypes) As System.Windows.Forms.TabPage

        Dim ProductTab = New System.Windows.Forms.TabPage()
        Select Case prodType
            Case ViewModels.ProductTypes.CPb
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.Cpb, ViewModels.ProductTypes.CPb))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabCpb"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 2
                    .Text = "Cpb"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
            Case ViewModels.ProductTypes.CpbRej
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.CpbRej, ViewModels.ProductTypes.CpbRej))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabCpbRej"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 3
                    .Text = "CpbRej"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
            Case ViewModels.ProductTypes.FPb
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.Fpb, ViewModels.ProductTypes.FPb))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabFpb"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 4
                    .Text = "Fpb"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
            Case ViewModels.ProductTypes.FPbRej
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.FpbRej, ViewModels.ProductTypes.FPbRej))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabFpbRej"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 4
                    .Text = "FpbRej"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
            Case ViewModels.ProductTypes.IP
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.Ip, ViewModels.ProductTypes.IP))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabIP"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 5
                    .Text = "IP"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
            Case ViewModels.ProductTypes.IPRej
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.IpRej, ViewModels.ProductTypes.IPRej))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabIPRej"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 6
                    .Text = "IPRej"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
            Case ViewModels.ProductTypes.CCn
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.Ccn, ViewModels.ProductTypes.CCn))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabCcn"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 7
                    .Text = "Ccn"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
            Case ViewModels.ProductTypes.FCn
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.Fcn, ViewModels.ProductTypes.FCn))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabFcn"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 8
                    .Text = "FCn"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
            Case ViewModels.ProductTypes.TCn
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.Tcn, ViewModels.ProductTypes.TCn))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabTcn"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 9
                    .Text = "Tcn"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
            Case ViewModels.ProductTypes.Os
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.Os, ViewModels.ProductTypes.Os))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabOs"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 10
                    .Text = "Os"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
            Case ViewModels.ProductTypes.TPb
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.Tpb, ViewModels.ProductTypes.TPb))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabTpb"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 11
                    .Text = "Tpb"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
            Case ViewModels.ProductTypes.TPbRej
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.TpbRej, ViewModels.ProductTypes.TPbRej))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabTpbRej"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 12
                    .Text = "TpbRej"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
            Case ViewModels.ProductTypes.ATpr
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.ATpr, ViewModels.ProductTypes.ATpr))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabATpr"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 13
                    .Text = "ATpr"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
            Case ViewModels.ProductTypes.Tpr
                Dim ProdResults As New List(Of ViewModels.SplitProductResult)
                For Each RdctnCalcData As gRawProspSplRdctnType In rawProspResults
                    ProdResults.Add(GetProductResult(RdctnCalcData.HoleDesc, RdctnCalcData.SplitNumber, RdctnCalcData.Tpr, ViewModels.ProductTypes.Tpr))
                Next
                Dim ProductResultGrid As New ctrProductResult(ProdResults)
                ProductResultGrid.Dock = DockStyle.Fill
                With ProductTab
                    .Location = New System.Drawing.Point(4, 22)
                    .Name = "tabTpr"
                    .Padding = New System.Windows.Forms.Padding(3)
                    .Size = New System.Drawing.Size(860, 209)
                    .TabIndex = 14
                    .Text = "Tpr"
                    .UseVisualStyleBackColor = True
                    .Controls.Add(ProductResultGrid)
                End With
        End Select
        Return ProductTab


    End Function

    Private Function GetProductResult(ByVal holedesc As String, ByVal splitNumb As Integer, ByVal rdctnCalcData As mProdInfoType, ByVal prodType As ViewModels.ProductTypes) As ViewModels.SplitProductResult
        Dim ProductResult As New ViewModels.SplitProductResult(prodType)
        With rdctnCalcData
            ProductResult.TRSH = holedesc
            ProductResult.SplitNumber = splitNumb
            ProductResult.IsOnSpec = .IsOnSpec
            ProductResult.Tpa = .Tpa
            ProductResult.WtPct = .WtPct
            ProductResult.Al = .Al
            ProductResult.Bpl = .Bpl
            ProductResult.Ca = .Ca
            ProductResult.Fe = .Fe
            ProductResult.FeAdj = .FeAdj
            ProductResult.Mg = .Mg
            ProductResult.Ins = .Ins
            ProductResult.Ia = .Ia
            ProductResult.IaAdj = .IaAdj
            ProductResult.Mer = .Mer
            ProductResult.MgOP2O5 = .MgOP2O5
            ProductResult.Fe2O3P2O5 = .Fe2O3P2O5
            ProductResult.CaOP2O5 = .CaOP2O5
            ProductResult.BplOffSpecFlag = .BplOffSpecFlag
            ProductResult.MgOffSpecFlag = .MgOffSpecFlag
            ProductResult.AlOffSpecFlag = .AlOffSpecFlag
            ProductResult.CaOffSpecFlag = .CaOffSpecFlag
            ProductResult.FeOffSpecFlag = .FeOffSpecFlag
            ProductResult.IaOffSpecFlag = .IaOffSpecFlag
            ProductResult.InsOffSpecFlag = .InsOffSpecFlag
            ProductResult.MerOffSpecFlag = .MerOffSpecFlag
            ProductResult.CaOP2O5OffSpecFlag = .CaOP2O5OffSpecFlag
            ProductResult.Fe2O3P2O5OffSpecFlag = .Fe2O3P2O5OffSpecFlag
            ProductResult.MgOP2O5OffSpecFlag = .MgOP2O5OffSpecFlag
        End With
        Return ProductResult
    End Function

    Private Sub cmdCopyToOverrides_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyToOverrides.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim RowIdx As Integer
        Dim ThisHoleLoc As String
        Dim ThisSplit As Integer
        Dim ThisMineable As String

        'Copy from ssSplitReview to ssSplitOverride

        SetActionStatus("Copying oride data...")
        Me.Cursor = Cursors.WaitCursor

        ssSplitOverride.MaxRows = 0

        With ssSplitReview
            ssSplitOverride.ReDraw = False
            For RowIdx = 1 To .MaxRows
                .Row = RowIdx
                .Col = 1
                ThisHoleLoc = .Text
                .Col = 2
                ThisSplit = .Value
                .Col = 5
                ThisMineable = .Text    'Mineable override -- M, U, C
                If Trim(ThisMineable) = "" Then
                    ThisMineable = "C"
                End If

                'With ssSplitOverride
                '    .MaxRows = .MaxRows + 1
                '    .Row = .MaxRows
                '    .Col = 1
                '    .Text = ThisHoleLoc
                '    .Col = 2
                '    .Value = ThisSplit
                '    .Col = 3
                '    .Text = ThisMineable    'M, U, C
                'End With
            Next RowIdx
            ssSplitOverride.ReDraw = True
        End With

        MarkHolesGreen(ssSplitOverride, True)

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub cmdSaveProspectDataset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveProspectDataset.Click

        Dim SaveStatus As Boolean
        Dim SaveType As String
        Dim MineHasOffSpecPbPlt As Boolean
        Dim MineHasDoloflotPlt As Boolean
        Dim NoReview As Boolean
        Dim SetPbToMgPlt As Boolean

        If chkCreateOutputOnly.Checked = True Then
            NoReview = True
        Else
            NoReview = False
        End If

        SetActionStatus("Saving raw prospect data...")
        Me.Cursor = Cursors.WaitCursor

        If chkSaveToDatabase.Checked = True Then
            SaveType = "Database"
        End If
        If chkSurvCaddTextfile.Checked Then
            SaveType = "SurvCaddText"
        End If
        If chkBdFormatTextfile.Checked Then
            SaveType = "BdFormatText"
        End If
        If chkSpecMoisTransferFile.Checked = True Then
            SaveType = "MoisText"
        End If

        'gSaveProspectDataset is in modRawProspDataReduction.
        'lblGen65) was lblOffSpecPbMgPlt
        If lblGen65.Text = "*OffSpec Pb Mg Plt*" Or
        lblGen65.Text = "*Doloflot Plt FCO*" Then
            MineHasOffSpecPbPlt = True
        Else
            MineHasOffSpecPbPlt = False
        End If

        'lblGen65) was lblOffSpecPbMgPlt
        If lblGen65.Text = "*Doloflot Plt Ona*" Then
            MineHasDoloflotPlt = True
        Else
            MineHasDoloflotPlt = False
        End If

        If chkPbAnalysisFillInSpecial.Checked = True Then
            SetPbToMgPlt = True
        Else
            SetPbToMgPlt = False
        End If

        gFileNumber = -99

        ' We need to translate ssSplitReview and ssCompReview to Lists of gRawProspSplRdctnType

        Dim ResultSet As SplitResultSet = GetResultSetFromSpreadSheets(ssSplitReview, ssCompReview)

        SaveStatus = gSaveProspectDataset(SaveType,
                                          txtProspectDatasetName.Text,
                                          txtProspDatasetTextfileName.Text,
                                          IIf(chk100Pct.Checked, 1, 0),
                                           IIf(chkProductionCoefficient.Checked, 1, 0),
                                          optInclSplits.Checked,
                                          optInclComposites.Checked,
                                          optInclBoth.Checked,
                                          ssCompReview,
                                          ssSplitReview,
                                          ResultSet.SplitResults,
                                          ResultSet.HoleResults,
                                          MineHasOffSpecPbPlt,
                                          IIf(chkInclMgPlt.Checked, 1, 0),
                                          False,
                                          0,
                                          NoReview,
                                          1,
                                          SetPbToMgPlt,
                                          MineHasDoloflotPlt,
                                          _recoveryScenariosForm.ProductRecoveryDefinition.UseAdjustedFeToDetermineMineability)

        SetActionStatus("")

        If SaveType = "SurvCaddText" AndAlso Not txtProspDatasetTextfileName.Equals(String.Empty) Then
            Dim sw As StreamWriter = New StreamWriter(txtProspDatasetTextfileName.Text, False)
            sw.Write(gOutputLines)
            sw.Close()
            sw = Nothing
            gOutputLines.Clear()
        End If

        Me.Cursor = Cursors.Arrow

        If SaveStatus = True Then
            MsgBox("Prospect dataset save completed -- no problems!",
                   vbOKOnly, "Save Status")
        Else
            MsgBox("Prospect dataset save completed -- PROBLEMS!",
                   vbOKOnly, "Save Status")
        End If
    End Sub

    Private Sub cmdPrtGrdDetlDisp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrtGrdDetlDisp.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo cmdPrtGrdDetlDispClickError

        gClearGridPrint()

        ' Set 
        gGridObject = ssDetlDisp

        gPrintGridHeader = lblGen41.Text   'Was lblReviewComm
        gPrintGridSubHeader1 = "Printed: " & Format(Now, "MM/dd/yyyy hh:mm tt")

        gPrintGridSubHeader2 = ""

        gOrientHeader = "Center"
        gOrientSubHeader1 = "Center"
        gOrientSubHeader2 = "Center"

        gPrintGridFooter = ""
        gOrientFooter = ""
        gSubHead2IsHeader = False

        gPrintGridDefaultTxtFname = ""

        gPrintMarginLeft = 1440     '1440 = 1"
        gPrintMarginRight = 1440
        gPrintMarginTop = 770
        gPrintMarginBottom = 770

        SetActionStatus("Printing spreadsheet...")
        Me.Cursor = Cursors.WaitCursor
        Print.frmGridToText.ShowDialog()
        Print.frmGridToText.Dispose()
        'Load frmGridToText
        'frmGridToText.Show vbModal
        'Unload frmGridToText

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow

        Exit Sub

cmdPrtGrdDetlDispClickError:
        MsgBox("Error printing grid." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Grid Error")

        On Error Resume Next
        Me.Cursor = Cursors.Arrow
        SetActionStatus("")
    End Sub


    Private Sub DispFdRcvryCalcs()

        'Need to place some calculations in ssCalcdResults

        Dim CrsFdTonRcvry As Single
        Dim CrsFdBplRcvry As Single
        Dim CrsFlotRcvry As Single
        Dim FneFdTonRcvry As Single
        Dim FneFdBplRcvry As Single
        Dim FneFlotRcvry As Single

        'With ssProdRcvryFctrs
        '    .Row = 3
        '    .Col = 3
        '    CrsFdTonRcvry = .Value
        '    .Row = 4
        '    .Col = 3
        '    CrsFdBplRcvry = .Value
        '    .Row = 3
        '    .Col = 5
        '    FneFdTonRcvry = .Value
        '    .Row = 4
        '    .Col = 5
        '    FneFdBplRcvry = .Value
        'End With

        'With ssFlotRcvryLinear
        '    .Row = 1
        '    .Col = 5
        '    CrsFlotRcvry = .Value
        '    .Row = 2
        '    .Col = 5
        '    FneFlotRcvry = .Value
        'End With

        'If CrsFdTonRcvry > 0 And CrsFdBplRcvry > 0 And
        '    CrsFlotRcvry > 0 Then
        '    'With ssCalcdResults
        '    '    .Row = 1
        '    '    .Col = 1
        '    '    .Value = Round((CrsFdTonRcvry / 100 * CrsFdBplRcvry / 100 *
        '    '             CrsFlotRcvry / 100) * 100, 0)
        '    '    .Row = 1
        '    '    .Col = 2
        '    '    .Value = Round((FneFdTonRcvry / 100 * FneFdBplRcvry / 100 *
        '    '             FneFlotRcvry / 100) * 100, 0)

        '    '    .Row = 1
        '    '    .Col = 4
        '    '    .Value = Round((1 / ((CrsFdTonRcvry * CrsFdBplRcvry + FneFdTonRcvry * FneFdBplRcvry) /
        '    '             (CrsFdTonRcvry + FneFdTonRcvry) / 100) - 1) * 100, 1)
        '    'End With
        'End If
    End Sub

    Private Sub ssHolePhysMineability_ButtonClicked(ByVal sender As System.Object, ByVal e As AxFPSpread._DSpreadEvents_ButtonClickedEvent)  '(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

        Dim ThisVal As Single

        If e.row = 1 And e.col = 3 Then
            'With ssSplitPhysMineability
            '    .Row = 1
            '    .Col = 1
            '    ThisVal = .Value
            'End With

            'With ssHolePhysMineability
            '    .Row = 1
            '    .Col = 1
            '    .Value = ThisVal
            'End With
        End If
    End Sub

    Private Sub ssHoleEconMineability_ButtonClicked(ByVal sender As System.Object, ByVal e As AxFPSpread._DSpreadEvents_ButtonClickedEvent)   '(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)


        Dim ThisVal As Single

        If e.row = 1 And e.col = 3 Then
            'With ssSplitEconMineability
            '    .Row = 1
            '    .Col = 1
            '    ThisVal = .Value
            'End With

            'With ssHoleEconMineability
            '    .Row = 1
            '    .Col = 1
            '    .Value = ThisVal
            'End With
        End If
    End Sub


    Private Sub AddMiscVertDividers()

        Dim RowIdx As Integer

        'With ssInsAdj100Pct
        '    .Row = 0
        '    .Col = 5
        '    .Text = " "
        '    .set_ColWidth(.Col, 0.17)
        '    For RowIdx = 0 To .MaxRows
        '        .Row = RowIdx
        '        .CellType = FPSpread.CellTypeConstants.CellTypeStaticText
        '        .BackColor = Color.Black
        '    Next
        'End With

        'With ssHolePhysMineability
        '    .Row = 0
        '    .Col = 2
        '    .Text = " "
        '    .set_ColWidth(.Col, 0.17)
        '    For RowIdx = 0 To .MaxRows
        '        .Row = RowIdx
        '        .CellType = FPSpread.CellTypeConstants.CellTypeStaticText
        '        .BackColor = Color.Black
        '    Next
        'End With

        'With ssHoleEconMineability
        '    .Row = 0
        '    .Col = 2
        '    .Text = " "
        '    .set_ColWidth(.Col, 0.17)
        '    For RowIdx = 0 To .MaxRows
        '        .Row = RowIdx
        '        .CellType = FPSpread.CellTypeConstants.CellTypeStaticText
        '        .BackColor = Color.Black
        '    Next
        'End With

        'With ssVolRcvry
        '    .Row = 0
        '    .Col = 4
        '    .Text = " "
        '    .set_ColWidth(.Col, 0.17)
        '    For RowIdx = 0 To .MaxRows
        '        .Row = RowIdx
        '        .CellType = FPSpread.CellTypeConstants.CellTypeStaticText
        '        .BackColor = Color.Black
        '    Next
        'End With

        'With ssFlotRcvryLinear
        '    .Row = 0
        '    .Col = 3
        '    .Text = " "
        '    .set_ColWidth(.Col, 0.17)
        '    For RowIdx = 0 To .MaxRows
        '        .Row = RowIdx
        '        .CellType = FPSpread.CellTypeConstants.CellTypeStaticText
        '        .BackColor = Color.Black
        '    Next

        '    .Row = 0
        '    .Col = 7
        '    .Text = " "
        '    .set_ColWidth(.Col, 0.17)
        '    For RowIdx = 0 To .MaxRows
        '        .Row = RowIdx
        '        .CellType = FPSpread.CellTypeConstants.CellTypeStaticText
        '        .BackColor = Color.Black
        '    Next
        'End With

        'With ssFlotRcvryHardwire
        '    .Row = 0
        '    .Col = 3
        '    .Text = " "
        '    .set_ColWidth(.Col, 0.17)
        '    For RowIdx = 0 To .MaxRows
        '        .Row = RowIdx
        '        .CellType = FPSpread.CellTypeConstants.CellTypeStaticText
        '        .BackColor = Color.Black ' Color.Black
        '    Next

        '    .Row = 0
        '    .Col = 7
        '    .Text = " "
        '    .set_ColWidth(.Col, 0.17)
        '    For RowIdx = 0 To .MaxRows
        '        .Row = RowIdx
        '        .CellType = FPSpread.CellTypeConstants.CellTypeStaticText
        '        .BackColor = Color.Black
        '    Next
        'End With

        'With ssProdRcvryFctrs
        '    .Row = 0
        '    .Col = 1
        '    .Text = " "
        '    .set_ColWidth(.Col, 0.17)
        '    For RowIdx = 0 To .MaxRows
        '        .Row = RowIdx
        '        .CellType = FPSpread.CellTypeConstants.CellTypeStaticText 'SS_CELL_TYPE_STATIC_TEXT
        '        .BackColor = Color.Black
        '    Next

        '    .Row = 0
        '    .Col = 6
        '    .Text = " "
        '    .set_ColWidth(.Col, 0.17)
        '    For RowIdx = 0 To .MaxRows
        '        .Row = RowIdx
        '        .CellType = FPSpread.CellTypeConstants.CellTypeStaticText
        '        .BackColor = Color.Black
        '    Next
        'End With

        With ssRawProspMin
            .Row = 0
            .Col = 7
            .Text = " "
            .set_ColWidth(.Col, 0.17)
            For RowIdx = 0 To .MaxRows
                .Row = RowIdx
                .CellType = FPSpread.CellTypeConstants.CellTypeStaticText
                .BackColor = Color.Black
            Next
        End With
    End Sub

    Private Sub GetProspRawMatlSizeDefn(ByVal aWeightTableVersion As String,
                                        ByVal aSfcReproData(,) As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetProspRawMatlSizeDefnError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim WtTableVersionDynaset As OraDynaset
        Dim ThisWtTOableVersion As String
        Dim aValue As Object
        Dim RecordCount As Integer

        Dim ThisWtTableVersion As String
        Dim ThisSizeFrctnCode As String
        Dim ThisProspectProdGroup As String
        Dim ThisProdGrpSizeFrctnCode As String
        Dim ThisProdGrpSizeFrctnDesc As String
        Dim ThisRowNum As Integer
        Dim ThisColNum As Integer
        Dim ThisWeightPct As Single
        Dim ItemCount As Integer

        Dim RowIdx As Integer
        Dim ColIdx As Integer

        For RowIdx = 1 To UBound(aSfcReproData, 1)
            For ColIdx = 1 To UBound(aSfcReproData, 2)
                If (RowIdx = 1 Or RowIdx = 2) Or ColIdx = 1 Then
                    aSfcReproData(RowIdx, ColIdx) = ""
                Else
                    aSfcReproData(RowIdx, ColIdx) = "0"
                End If
            Next ColIdx
        Next RowIdx

        params = gDBParams

        params.Add("pMineName", gActiveMineNameLong, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pWeightTableVersion", aWeightTableVersion, ORAPARM_INPUT)
        params("pWeightTableVersion").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_size_frctn_weightnz
        'pMineName           IN     VARCHAR2,
        'pWeightTableVersion IN     VARCHAR2,
        'pResult             IN OUT c_matlsizedefn)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_size_frctn_weightnz(:pMineName," +
                      ":pWeightTableVersion, :pResult);end;", ORASQL_FAILEXEC)
        WtTableVersionDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = WtTableVersionDynaset.RecordCount

        'Only non-zero weight percent items have been returned!

        WtTableVersionDynaset.MoveFirst()
        Do While Not WtTableVersionDynaset.EOF
            ThisWtTableVersion = WtTableVersionDynaset.Fields("weight_table_version").Value
            ThisSizeFrctnCode = WtTableVersionDynaset.Fields("size_frctn_code").Value
            ThisProspectProdGroup = WtTableVersionDynaset.Fields("prospect_prod_grp").Value
            ThisProdGrpSizeFrctnCode = WtTableVersionDynaset.Fields("prod_grp_size_frctn_code").Value
            ThisProdGrpSizeFrctnDesc = WtTableVersionDynaset.Fields("prod_grp_size_frctn_desc").Value
            ThisRowNum = WtTableVersionDynaset.Fields("row_num").Value
            ThisColNum = WtTableVersionDynaset.Fields("col_num").Value
            ThisWeightPct = WtTableVersionDynaset.Fields("weight_pct").Value

            ItemCount = ItemCount + 1

            'The Row and Col that are stored in PROSP_SIZE_FRCTN_WEIGHT are as follows:
            '        Col  Col  Col  Col  Col  Col  Col  Col
            '         2    3    4    5    6    7    8    9   etc.
            'Row  3
            'Row  4
            'Row  5
            'Row  6
            'Row  7
            'Row  etc.

            aSfcReproData(ThisRowNum, ThisColNum) = CStr(ThisWeightPct)

            'Add row heading
            If aSfcReproData(ThisRowNum, 1) = "" Then
                aSfcReproData(ThisRowNum, 1) = ThisSizeFrctnCode
            End If

            If aSfcReproData(0, ThisColNum) = "" Then
                aSfcReproData(0, ThisColNum) = ThisProdGrpSizeFrctnCode
            End If

            'Add some column headings
            If aSfcReproData(1, ThisColNum) = "" Then
                aSfcReproData(1, ThisColNum) = ThisProspectProdGroup
            End If
            If aSfcReproData(2, ThisColNum) = "" Then
                aSfcReproData(2, ThisColNum) = ThisProdGrpSizeFrctnDesc
            End If

            WtTableVersionDynaset.MoveNext()
        Loop

        WtTableVersionDynaset.Close()

        Exit Sub

GetProspRawMatlSizeDefnError:
        MsgBox("Error getting weight table version" & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Weight Table Version Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        WtTableVersionDynaset.Close()
    End Sub

    Private Sub ssSplitReview_Click(ByVal sender As System.Object, ByVal e As AxFPSpread._DSpreadEvents_ClickEvent) Handles ssSplitReview.ClickEvent '(ByVal Col As Long, ByVal Row As Long)

        'User has clicked on a row in ssSplitReview.
        'Need to move the column values in ssSplitReiview to row values in
        'ssDetlDisp.

        Dim ThisLoc As String

        If e.row = 0 Then
            Exit Sub
        End If

        MoveColsToRows(ssSplitReview, e.row, "Split")

        With ssSplitReview
            .Row = e.row
            .Col = 1
            ThisLoc = .Text
        End With
        GoToHoleInSprd(ssCompReview, ThisLoc)

        cmdHoleSplitRpt.Enabled = True
        optProdCoeff.Enabled = True
        opt100Pct.Enabled = True
    End Sub

    Private Sub ssSplitReview_Change(ByVal sender As System.Object, ByVal e As AxFPSpread._DSpreadEvents_ChangeEvent) Handles ssSplitReview.Change '(ByVal Col As Long, ByVal Row As Long)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim MineableCalcd As String
        Dim MineableOverride As String

        If fProcessing <> True Then
            'If user has messed with the Mineability override then may
            'have to change the Key value in column 19.
            If e.col = 5 Then
                With ssSplitReview
                    .Row = e.row
                    .Col = 4
                    MineableCalcd = .Text
                    .Col = 5
                    MineableOverride = .Text
                    .Col = 19
                    If MineableCalcd = "M" Or MineableOverride = "M" Then
                        .Value = 1
                    Else
                        .Value = 0
                    End If

                    .Col = e.col
                End With
            End If
        End If
    End Sub

    Private Sub ssCompReview_Click(ByVal sender As System.Object, ByVal e As AxFPSpread._DSpreadEvents_ClickEvent) Handles ssCompReview.ClickEvent  '(ByVal Col As Long, ByVal Row As Long)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        If fProcessing <> True Then
            'User has clicked on a row in ssCompReview.
            'Need to move the column values in ssCompReview to row values in
            'ssDetlDisp.
            Dim RowIdx As Integer
            Dim ThisLoc As String

            If e.row = 0 Then
                Exit Sub
            End If

            MoveColsToRows(ssCompReview, e.row, "Comp")

            With ssCompReview
                .Row = e.row
                .Col = 1
                ThisLoc = .Text
            End With
            GoToHoleInSprd(ssSplitReview, ThisLoc)
            GoToHoleInSprd(ssCompErrors, ThisLoc)

            cmdHoleSplitRpt.Enabled = True
            optProdCoeff.Enabled = True
            opt100Pct.Enabled = True
        End If
    End Sub

    Private Sub MoveColsToRows(ByRef aSpread As AxvaSpread,
                           ByVal aRow As Integer,
                           ByVal aSprdType As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ColIdx As Integer

        'Need to move the column values in ssSplitReview or ssCompReview to
        'row values in ssDetlDisp.
        'aSprdType will be "Split" or "Comp".

        lblGen41.Text = ""   'Was lblReviewComm
        lblGen64.Text = ""   'Was lblRowNum

        If aSprdType = "Split" Then
            lblGen41.Text = "Split Data"  'Was lblReviewComm
        End If
        If aSprdType = "Comp" Then
            lblGen41.Text = "Hole Data"   'Was lblReviewComm
        End If

        'Will keep the row# here instead of "wasteing" a form-level
        'variable.
        lblGen64.Text = CStr(aRow)     'Was lblRowNum

        With aSpread
            .Row = aRow
            For ColIdx = 1 To .MaxCols
                'I originally used Cols 156 to 189 for "In plant" and
                'To the plant" stuff
                If ColIdx <= 155 Or ColIdx >= 190 Then
                    .Col = ColIdx
                    ssDetlDisp.Row = .Col
                    ssDetlDisp.Col = 1

                    'Text columns
                    'Col 1   Hole location
                    'Col 3   Prospect date
                    'Col 4   Mineable Calc'd
                    'Col 5   Mineable Oride
                    'Col 10  Ownership
                    'Col 11  Mined out status
                    'Col 12  Hole type
                    'Col 13  Expanded drill
                    'Col 18  Override
                    'Col 20  Class
                    'Col 21  Bed
                    'Col 22  Level
                    'Col 23  Horizon
                    'Col 84  Mtx color
                    'Col 85  Deg consol
                    'Col 86  Dig char
                    'Col 87  Pump char
                    'Col 88  Lithology
                    'Col 89  Phosph color
                    'Col 90  SurvCADD hole ID
                    'Col 190 Mineable Hole PC
                    'Col 191 Cpb Mineable
                    'Col 192 Fpb Mineable
                    'Col 193 Tpb Mineable
                    'Col 194 Ccn Mineable
                    'Col 195 Fcn Mineable
                    'Col 196 Tcn Mineable
                    'Col 197 Os Mineable
                    'Col 198 IP Mineable
                    'Col 313 Mineable Hole 100
                    'Col 319 CpbMinHole
                    'Col 320 FpbMinHole
                    'Col 321 TpbMinHole
                    'Col 322 CcnMinHole
                    'Col 323 FcnMinHole
                    'Col 324 TcnMinHole
                    'Col 325 OsMinHole
                    'Col 326 IpMinHole
                    'Col 327 Mtx"X" OnSpec PC Hole
                    'Col 328 Tot"X" OnSpec PC Hole
                    'Col 329 Mtx"X" All PC Hole
                    'Col 330 Tot"X" All PC Hole
                    'Col 331 Mtx"X" OnSpec 100 Hole
                    'Col 332 Tot"X" OnSpec 100 Hole
                    'Col 333 Mtx"X" All 100 Hole
                    'Col 334 Tot"X" All 100 Hole
                    'Col 335 Mtx %Mois
                    'Col 336 Mtx %Sol

                    If .Col = 1 Or .Col = 3 Or .Col = 4 Or .Col = 5 Or .Col = 10 Or
                        .Col = 11 Or .Col = 12 Or .Col = 18 Or
                        .Col = 13 Or .Col = 20 Or .Col = 21 Or .Col = 22 Or
                        .Col = 23 Or .Col = 84 Or .Col = 85 Or .Col = 86 Or
                        .Col = 87 Or .Col = 88 Or .Col = 89 Or .Col = 90 Or
                        .Col = 190 Or .Col = 191 Or .Col = 192 Or .Col = 193 Or
                        .Col = 194 Or .Col = 195 Or .Col = 196 Or .Col = 197 Or
                        .Col = 198 Or .Col = 313 Or .Col <> 319 Or .Col <> 320 Or
                        .Col <> 321 Or .Col <> 322 Or .Col <> 323 Or .Col <> 324 Or
                        .Col <> 325 Or .Col <> 326 Then
                        ssDetlDisp.Text = .Text
                    Else
                        ssDetlDisp.Value = Val(.Value)
                    End If

                    If .FontBold = True Then
                        ssDetlDisp.FontBold = True
                    Else
                        ssDetlDisp.FontBold = False
                    End If
                    If .ForeColor = Color.DarkRed Then    'Dark red &HC0& 
                        ssDetlDisp.ForeColor = Color.DarkRed        'Dark red
                    Else
                        ssDetlDisp.ForeColor = SystemColors.ControlText ' &H80000012   'Normal
                    End If
                End If
            Next ColIdx
        End With
    End Sub


    Private Sub MarkProdsInDetlDispSprd()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim BlockIdx As Integer

        For BlockIdx = 1 To 22
            With ssDetlDisp
                .BlockMode = True
                Select Case BlockIdx
                    Case Is = 1     'OS PC
                        .Row = 31
                        .Row2 = 41

                    Case Is = 2     'IP PC
                        .Row = 53
                        .Row2 = 63

                    Case Is = 3     'Cfd PC
                        .Row = 75
                        .Row2 = 77

                    Case Is = 4     'Wcly PC
                        .Row = 81
                        .Row2 = 83

                    Case Is = 5     'CrsPb PC
                        .Row = 91
                        .Row2 = 101

                    Case Is = 6     'Misc
                        .Row = 113
                        .Row2 = 113

                    Case Is = 7     'Tfd PC
                        .Row = 117
                        .Row2 = 119

                    Case Is = 8     'Misc
                        .Row = 131
                        .Row2 = 133

                    Case Is = 9     'Fcn PC
                        .Row = 145
                        .Row2 = 155

                    Case Is = 10    'Not used
                        .Row = 159
                        .Row2 = 161

                    Case Is = 11    'Not used
                        .Row = 165
                        .Row2 = 167

                    Case Is = 12    'Not used
                        .Row = 171
                        .Row2 = 173

                    Case Is = 13    'Not used
                        .Row = 179
                        .Row2 = 182

                    Case Is = 14
                        .Row = 187
                        .Row2 = 189

                    Case Is = 15    'Cpb 100
                        .Row = 199
                        .Row2 = 209

                    Case Is = 16    'Ttl 100
                        .Row = 221
                        .Row2 = 223

                    Case Is = 17    'Tpr 100
                        .Row = 227
                        .Row2 = 237

                    Case Is = 18    'Fcn 100
                        .Row = 249
                        .Row2 = 259

                    Case Is = 19    'Tpr 100
                        .Row = 271
                        .Row2 = 281

                    Case Is = 20    'Tcn 100
                        .Row = 293
                        .Row2 = 303

                    Case Is = 21    'Ffd 100
                        .Row = 307
                        .Row2 = 309

                    Case Is = 22    'Misc
                        .Row = 312
                        .Row2 = 314
                End Select

                .Col = 1
                .Col2 = 1
                .BackColor = Color.LightGreen ' &HD8FFD8      'Light light green
                .BlockMode = False
            End With
        Next BlockIdx

        'Mark mineable stuff with light blue.
        For BlockIdx = 1 To 3
            With ssDetlDisp
                .BlockMode = True
                Select Case BlockIdx
                    Case Is = 1
                        .Row = 4
                        .Row2 = 5

                    Case Is = 2
                        .Row = 190
                        .Row2 = 198

                    Case Is = 3
                        .Row = 313
                        .Row2 = 314
                End Select

                .Col = 1
                .Col2 = 1
                .BackColor = Color.LightBlue ' &HFFFFC0    'Light blue
                .BlockMode = False
            End With
        Next BlockIdx
    End Sub

    Private Sub ClearDetlDisp()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        With ssDetlDisp
            .BlockMode = True
            .Row = 1
            .Row2 = .MaxRows
            .Col = 1
            .Col2 = 1
            .Action = 12    'Clear text
            .BlockMode = False
        End With
    End Sub

    Private Sub cmdHoleSplitRpt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHoleSplitRpt.Click


        On Error GoTo cmdHoleSplitRptClickError

        Dim ConnectString As String
        Dim ProspData As gRawProspSplRdctnType
        Dim RcvryData As gDataRdctnParamsType
        Dim RcvryProdQual(0 To 14) As gDataRdctnProdQualType
        Dim PctProspect100 As Boolean
        Dim InclCpbAlways As String
        Dim InclFpbAlways As String
        Dim InclOsAlways As String
        Dim InclCpbNever As String
        Dim InclFpbNever As String
        Dim InclOsNever As String
        Dim MtxDensityCalc As Single
        Dim MineHasOffSpecPbPlt As String
        Dim CanSelectRejectTpb As String
        Dim CanSelectRejectTcn As String
        Dim UseFeAdjust As String

        Dim OsOnSpec As String
        Dim CpbOnSpec As String
        Dim FpbOnSpec As String
        Dim TpbOnSpec As String
        Dim CcnOnSpec As String
        Dim FcnOnSpec As String
        Dim TcnOnSpec As String
        Dim IpOnSpec As String

        'lblReviewComm --> lblGen41)
        'lblRowNum --> lblGen64)
        If lblGen64.Text = "" Or lblGen41.Text = "" Then
            Exit Sub
        End If

        PctProspect100 = False
        If opt100Pct.Checked = True Then
            PctProspect100 = True
        Else
            PctProspect100 = False
        End If

        GetRcvryEtcParamsFromForm(RcvryData)

        InclCpbAlways = "No"
        InclFpbAlways = "No"
        InclOsAlways = "No"
        InclCpbNever = "No"
        InclFpbNever = "No"
        InclOsNever = "No"
        CanSelectRejectTpb = "No"
        CanSelectRejectTcn = "No"
        If RcvryData.InclCpbAlways Then
            InclCpbAlways = "Yes"
        End If
        If RcvryData.InclFpbAlways Then
            InclFpbAlways = "Yes"
        End If
        If RcvryData.InclOsAlways Then
            InclOsAlways = "Yes"
        End If
        If RcvryData.InclCpbNever Then
            InclCpbNever = "Yes"
        End If
        If RcvryData.InclFpbNever Then
            InclFpbNever = "Yes"
        End If
        If RcvryData.InclOsNever Then
            InclOsNever = "Yes"
        End If
        If RcvryData.CanSelectRejectTpb Then
            CanSelectRejectTpb = "Yes"
        End If
        If RcvryData.CanSelectRejectTcn Then
            CanSelectRejectTcn = "Yes"
        End If

        '06/14/2010, lss  Don't use RcvryData.MineHasOffSpecPbPlt anymore!
        'Get comment concerning off-spec MgO plant
        'May be "Original" or "Doloflot"
        MineHasOffSpecPbPlt = ""
        If RcvryData.UseOrigMgoPlant = True Then
            MineHasOffSpecPbPlt = "Mine has off-spec pebble processing plant (Original)."
        Else
            If RcvryData.UseDoloflotPlant2010 = True Then
                MineHasOffSpecPbPlt = "Mine has Ona Doloflot plant (2010)."
            Else
                If RcvryData.UseDoloflotPlantFco = True Then
                    MineHasOffSpecPbPlt = "Mine has FCO Doloflot plant (2011)."
                Else
                    MineHasOffSpecPbPlt = ""
                End If
            End If
        End If

        'lblReviewComm --> lblGen41)
        'lblRowNum --> lblGen64)
        If lblGen41.Text = "Hole Data" Then
            ProspData = gGetDataFromReviewSprd(ssCompReview, Val(lblGen64.Text))
        End If
        'lblReviewComm --> lblGen41)
        'lblRowNum --> lblGen64)
        If lblGen41.Text = "Split Data" Then
            ProspData = gGetDataFromReviewSprd(ssSplitReview, Val(lblGen64.Text))
        End If

        If RcvryData.UseFeAdjust = True Then
            UseFeAdjust = "Fe2O3 adjust has been used to determine minabilities."
        Else
            UseFeAdjust = ""
        End If

        'Need to get whether the material is on-spec as opposed to whether the
        'material will or will not be mined!
        'lblReviewComm --> lblGen41)
        'lblRowNum --> lblGen64)
        If lblGen41.Text = "Split Data" Then
            OsOnSpec = gGetMaterialOnSpec("OS", Val(lblGen64.Text), ssSplitReview)
            CpbOnSpec = gGetMaterialOnSpec("Cpb", Val(lblGen64.Text), ssSplitReview)
            FpbOnSpec = gGetMaterialOnSpec("Fpb", Val(lblGen64.Text), ssSplitReview)
            TpbOnSpec = gGetMaterialOnSpec("Tpb", Val(lblGen64.Text), ssSplitReview)
            CcnOnSpec = gGetMaterialOnSpec("Ccn", Val(lblGen64.Text), ssSplitReview)
            FcnOnSpec = gGetMaterialOnSpec("Fcn", Val(lblGen64.Text), ssSplitReview)
            TcnOnSpec = gGetMaterialOnSpec("Tcn", Val(lblGen64.Text), ssSplitReview)
            IpOnSpec = gGetMaterialOnSpec("IP", Val(lblGen64.Text), ssSplitReview)
        End If
        'lblReviewComm --> lblGen41)
        'lblRowNum --> lblGen64)
        If lblGen41.Text = "Hole Data" Then
            OsOnSpec = gGetMaterialOnSpec("OS", Val(lblGen64.Text), ssCompReview)
            CpbOnSpec = gGetMaterialOnSpec("Cpb", Val(lblGen64.Text), ssCompReview)
            FpbOnSpec = gGetMaterialOnSpec("Fpb", Val(lblGen64.Text), ssCompReview)
            TpbOnSpec = gGetMaterialOnSpec("Tpb", Val(lblGen64.Text), ssCompReview)
            CcnOnSpec = gGetMaterialOnSpec("Ccn", Val(lblGen64.Text), ssCompReview)
            FcnOnSpec = gGetMaterialOnSpec("Fcn", Val(lblGen64.Text), ssCompReview)
            TcnOnSpec = gGetMaterialOnSpec("Tcn", Val(lblGen64.Text), ssCompReview)
            IpOnSpec = gGetMaterialOnSpec("IP", Val(lblGen64.Text), ssCompReview)
        End If

        'With ProspData
        '    rptProspRdctn.Formulas(1) = "ProspDate = '" & .ProspDate & "'"
        '    rptProspRdctn.Formulas(2) = "Section = '" & Format(.Section, "##") & "'"
        '    rptProspRdctn.Formulas(3) = "Township = '" & Format(.Township, "##") & "'"
        '    rptProspRdctn.Formulas(4) = "Range = '" & Format(.Range, "##") & "'"
        '    rptProspRdctn.Formulas(5) = "HoleLocation = '" & .HoleLocation & "'"

        '    'lblReviewComm --> lblGen41)
        '    If lblGen41.Text = "Split Data" Then
        '        rptProspRdctn.Formulas(6) = "SplitNumber = '" & Format(.SplitNumber, "##") & "'"
        '        rptProspRdctn.Formulas(7) = "SplitDepthTop = '" & Format(.SplitDepthTop, "##0.0") & "'"
        '        rptProspRdctn.Formulas(8) = "SplitDepthBot = '" & Format(.SplitDepthBot, "##0.0") & "'"
        '        rptProspRdctn.Formulas(9) = "SplitThck = '" & Format(.SplitThck, "##0.0") & "'"
        '    Else    'Hole data
        '        rptProspRdctn.Formulas(6) = "SplitNumber = '" & "--" & "'"
        '        rptProspRdctn.Formulas(7) = "SplitDepthTop = '" & "--" & "'"
        '        rptProspRdctn.Formulas(8) = "SplitDepthBot = '" & "--" & "'"
        '        rptProspRdctn.Formulas(9) = "SplitThck = '" & "--" & "'"
        '    End If

        '    rptProspRdctn.Formulas(10) = "Elevation = '" & Format(.Elevation, "#,##0.00") & "'"
        '    rptProspRdctn.Formulas(11) = "Xcoord = '" & Format(.Xcoord, "#,###,##0.00") & "'"
        '    rptProspRdctn.Formulas(12) = "Ycoord = '" & Format(.Ycoord, "#,###,##0.00") & "'"

        '    'OvbThk, ItbThk, and MtxThk will be the same for both 100% and ProdCoeff!
        '    'lblReviewComm --> lblGen41)
        '    If lblGen41.Text = "Hole Data" Then
        '        rptProspRdctn.Formulas(13) = "OvbThk = '" & Format(.OvbThk, "##0.0") & "'"
        '        rptProspRdctn.Formulas(14) = "MtxThk = '" & Format(.MtxThk, "##0.0") & "'"
        '        rptProspRdctn.Formulas(15) = "ItbThk = '" & Format(.ItbThk, "##0.0") & "'"
        '    Else    'Split data
        '        'These items do not apply to splits (only to holes).
        '        rptProspRdctn.Formulas(13) = "OvbThk = '" & "--" & "'"
        '        rptProspRdctn.Formulas(14) = "MtxThk = '" & "--" & "'"
        '        rptProspRdctn.Formulas(15) = "ItbThk = '" & "--" & "'"
        '    End If

        '    'Laboratory matrix density remains the same for both 100% and ProdCoeff!
        '    'This is the laboratory based density!
        '    rptProspRdctn.Formulas(16) = "MtxDensityLab = '" & Format(.MtxDensity, "##0.0") & "'"
        'End With

        'Public Sub gAddProdCoeffOr100Pct
        ' 1) aProspData As gRawProspSplRdctnType, _
        ' 2) aRcvryData As gDataRdctnParamsType, _
        ' 3) aOsOnSpec As String, _
        ' 4) aCpbOnSpec As String, _
        ' 5) aFpbOnSpec As String, _
        ' 6) aTpbOnSpec As String, _
        ' 7) aCcnOnSpec As String, _
        ' 8) aFcnOnSpec As String, _
        ' 9) aTcnOnSpec As String, _
        '10) aIpOnSpec As String, _
        '11) aInclCpbAlways As String, _
        '12) aInclFpbAlways As String, _
        '13) aInclOsAlways As String, _
        '14) aInclCpbNever As String, _
        '15) aInclFpbNever As String, _
        '16) aInclOsNever As String, _
        '17) aMineHasOffSpecPbPlt As String, _
        '18) aCanSelectRejectTpb As String, _
        '19) aCanSelectRejectTcn As String, _
        '20) aUseFeAdjust As String, _
        '21) aRdctnMode As String, _
        '22) aRptObj As CrystalReport, _
        '23) aSplitOrHole As String, _
        '24) aAreaDefnMineName As String)

        '    gAddProdCoeffOr100Pct(ProspData, _
        '                          RcvryData, _
        '                          OsOnSpec, _
        '                          CpbOnSpec, _
        '                          FpbOnSpec, _
        '                          TpbOnSpec, _
        '                          CcnOnSpec, _
        '                          FcnOnSpec, _
        '                          TcnOnSpec, _
        '                          IpOnSpec, _
        '                          InclCpbAlways, _
        '                          InclFpbAlways, _
        '                          InclOsAlways, _
        '                          InclCpbNever, _
        '                          InclFpbNever, _
        '                          InclOsNever, _
        '                          MineHasOffSpecPbPlt, _
        '                          CanSelectRejectTpb, _
        '                          CanSelectRejectTcn, _
        '                          UseFeAdjust, _
        '                          IIf(optProdCoeff.Checked = True, "ProdCoeff", "100%Prospect"), _
        '                          rptProspRdctn, _
        '                          IIf(lblGen41.Text = "Split Data", "Split", "Hole"), _
        '                          cboAreaDefnMineName.Text)

        ''Have all the needed data -- start the report
        'rptProspRdctn.ReportFileName = gPath + "\Reports\" + _
        '                               "ProspectReduction.rpt"

        ''Connect to Oracle database
        'ConnectString = "DSN = " + gDataSource + ";UID = " + gOracleUserName + _
        '    ";PWD = " + gOracleUserPassword + ";DSQ = "

        'rptProspRdctn.Connect = ConnectString

        ''Need to pass the company name and report type into the report
        ''lblReviewComm --> lblGen41)
        'rptProspRdctn.ParameterFields(0) = "pCompanyName;" & gCompanyName & ";TRUE"
        '    rptProspRdctn.ParameterFields(1) = "pRptType;" & lblGen41).Text & ";TRUE"
        'rptProspRdctn.ParameterFields(2) = "pPctProspect100;" & PctProspect100 & ";TRUE"
        'rptProspRdctn.ParameterFields(3) = "pMineHasOffSpecPbPlt;" & RcvryData.UseOrigMgoPlant & ";TRUE"
        'rptProspRdctn.ParameterFields(4) = "pProdSizeDesig;" & txtPsizeDefnName.Text & ";TRUE"
        'rptProspRdctn.ParameterFields(5) = "pRcvryEtcScen;" & txtRcvryEtcName.Text & ";TRUE"
        'rptProspRdctn.ParameterFields(6) = "pMineHasDoloflotPlt;" & RcvryData.UseDoloflotPlant2010 & ";TRUE"
        'rptProspRdctn.ParameterFields(7) = "pMineHasDoloflotPltFco;" & RcvryData.UseDoloflotPlantFco & ";TRUE"

        ''Report window maximized
        'rptProspRdctn.WindowState = crptMaximized

        'rptProspRdctn.WindowTitle = "Raw Data Reduction Prospect Data "

        ''User not allowed to minimize report window
        'rptProspRdctn.WindowMinButton = False

        ''Start Crystal Reports
        'rptProspRdctn.action = 1

        'rptProspRdctn.Reset
        'heyhey!
        Exit Sub

cmdHoleSplitRptClickError:
        MsgBox("Error printing prospect data from raw data reduction." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Print Error")

        On Error Resume Next
        'rptProspRdctn.Reset
    End Sub

    Private Sub MarkHolesGreen(ByRef aSpread As AxvaSpread,
                           ByVal aOverride As Boolean)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim CurrColor As String
        Dim ThisHoleLoc As String
        Dim CurrHoleLoc As String
        Dim RowIdx As Long
        Dim Row1 As Long
        Dim HoleCnt As Integer

        CurrHoleLoc = ""
        Row1 = 0
        CurrColor = "White"
        HoleCnt = 0

        With aSpread
            For RowIdx = 1 To .MaxRows
                .Row = RowIdx
                .Col = 1
                ThisHoleLoc = .Text

                If ThisHoleLoc <> CurrHoleLoc And CurrHoleLoc <> "" Then
                    HoleCnt = HoleCnt + 1
                    If CurrColor = "Green" Then
                        .BlockMode = True
                        .Row = Row1
                        .Row2 = RowIdx - 1
                        .Col = 1
                        .Col2 = .MaxCols
                        .BackColor = Color.LightGreen ' &HD8FFD8      'Light light green
                        .BlockMode = False
                        CurrColor = "White"
                    Else
                        CurrColor = "Green"
                    End If
                    Row1 = RowIdx
                    CurrHoleLoc = ThisHoleLoc
                End If

                If CurrHoleLoc = "" Then
                    CurrHoleLoc = ThisHoleLoc
                End If
            Next RowIdx

            'Need to mark the last hole!
            If .MaxRows <> 0 Then
                HoleCnt = HoleCnt + 1
                If CurrColor = "Green" Then
                    .BlockMode = True
                    .Row = Row1
                    .Row2 = RowIdx - 1
                    .Col = 1
                    .Col2 = .MaxCols
                    .BackColor = Color.LightGreen ' &HD8FFD8      'Light light green
                    .BlockMode = False
                End If
            End If
        End With

        If aOverride = True Then
            lblGen55.Text = "#Splits = " &
                             Format(ssSplitOverride.MaxRows, "###,##0") &
                             vbCrLf &
                             "#Holes = " &
                             Format(HoleCnt, "###,##0")
        End If
    End Sub

    Private ReadOnly Property ProductSizeOk() As Boolean
        Get
            If Not _productSizeDesginationForm.ProductSizeDesignation Is Nothing AndAlso _productSizeDesginationForm.ProductSizeDesignation.IsValid AndAlso Not _productSizeDesginationForm.ProductSizeDesignation.IsNew Then
                Return True
            End If
            Return False
        End Get
    End Property

    Private ReadOnly Property AreaDefnOk() As Boolean
        Get
            If _areaDefinitionForm.AreaDefinition IsNot Nothing OrElse _areaDefinitionForm.AreaDefinition.Holes.Count > 0 _
             OrElse _areaDefinitionForm.AreaDefinition.TRSCorners.Count > 0 OrElse _areaDefinitionForm.AreaDefinition.XYCorners.Count > 0 Then
                Return True
            End If
            Return False
        End Get
    End Property

    Private Sub chkSaveToDatabase_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSaveToDatabase.CheckedChanged

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        OutputSelectionSet("Database")
    End Sub

    Private Sub chkSurvCaddTextfile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSurvCaddTextfile.CheckedChanged

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        OutputSelectionSet("SurvCADD")
    End Sub

    Private Sub chkSpecMoisTransferFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSpecMoisTransferFile.CheckedChanged

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        OutputSelectionSet("SpecialMois")
    End Sub

    Private Sub chkBdFormatTextFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBdFormatTextfile.CheckedChanged

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        OutputSelectionSet("BdFormat")
    End Sub

    Private Sub OutputSelectionSet(ByVal aSelection As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'User has three "main" output choices:
        '1) Save to database                        chkSaveToDatabase
        '2) Save to SurvCADD transfer text file     chkSurvCaddTextFile
        '3) Save to MOIS transfer text file         chkSpecMoisTransferFile
        '4) Save to BD format text file             chkBdFormatTextfile

        'Only one can be selected at a time!

        Select Case aSelection
            Case Is = "Database"
                If chkSaveToDatabase.Checked = True Then
                    chkSurvCaddTextfile.Checked = False
                    chkBdFormatTextfile.Checked = False
                    chkSpecMoisTransferFile.Checked = False
                End If

            Case Is = "SurvCADD"
                '08/16/2007, lss
                'Only composites available for transfer at this time!
                If chkSurvCaddTextfile.Checked Then
                    chkSaveToDatabase.Checked = False
                    chkSpecMoisTransferFile.Checked = False
                    chkBdFormatTextfile.Checked = False

                    optInclComposites.Enabled = True
                    optInclComposites.Checked = True

                    '09/14/2009, lss
                    'Added SurvCADD splits option.
                    optInclComposites.Enabled = True

                    optInclBoth.Enabled = False
                End If

            Case Is = "BdFormat"
                '08/16/2007, lss
                'Only composites available for transfer at this time!
                If chkBdFormatTextfile.Checked Then
                    chkSaveToDatabase.Checked = False
                    chkSpecMoisTransferFile.Checked = False
                    chkSurvCaddTextfile.Checked = False

                    optInclComposites.Enabled = True
                    optInclComposites.Checked = True
                    optInclComposites.Enabled = False
                    optInclBoth.Enabled = False
                End If

            Case Is = "SpecialMois"
                'Special MOIS transfer text files are based on the IMC
                'RAR text file layout (Splits and Holes are combined).
                If chkSpecMoisTransferFile.Checked Then
                    chkSaveToDatabase.Checked = False
                    chkSurvCaddTextfile.Checked = False
                    chkBdFormatTextfile.Checked = False

                    optInclComposites.Enabled = False
                    optInclComposites.Enabled = False
                    optInclBoth.Enabled = True
                    optInclBoth.Checked = True
                End If
        End Select
    End Sub

    Private Sub chk100Pct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk100Pct.CheckedChanged

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        ProspDataTypeSet("100%")
    End Sub

    Private Sub chkProductionCoefficient_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkProductionCoefficient.CheckedChanged

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        ProspDataTypeSet("ProdCoeff")
    End Sub

    Private Sub ProspDataTypeSet(ByVal aSelection As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Select Case aSelection
            Case Is = "100%"
                If chk100Pct.Checked Then
                    chkProductionCoefficient.Checked = False
                End If

            Case Is = "ProdCoeff"
                If chkProductionCoefficient.Checked Then
                    chk100Pct.Checked = False
                End If
        End Select
    End Sub

    Private Sub chkUseOrigMgOPlant_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseOrigMgoPlant.CheckedChanged

        '06/14/2010, lss  chkMineHasOffSpecPbPlt not really used anymore!
        'Have chkUseDoloflotPlant and chkUseOrigMgOPlant and
        'chkUseDoloflotPlantFco

        If chkUseOrigMgoPlant.Checked = True Then
            'chkInclCpbNever.Checked = False
            'chkInclFpbNever.Checked = False
            'chkInclOsNever.Checked = True
            'chkInclCpbAlways.Checked = False
            'chkInclFpbAlways.Checked = False
            'chkInclOsAlways.Checked = False
            'chkCanSelectRejectTpb.Checked = False
        End If

        If chkUseOrigMgoPlant.Checked = True Then
            chkUseDoloflotPlant.Checked = False
            chkUseDoloflotPlantFco.Checked = False
        End If
    End Sub

    Private Sub SumResults()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim TotSpl As Long
        Dim TotSplMin As Long
        Dim TotSplCalcMin As Long
        Dim TotSplOrideMin As Long
        Dim TotHole As Integer
        Dim TotHoleMin As Integer
        Dim TotHoleMinForced As Integer
        Dim RowIdx As Long
        Dim MinCalcd As String
        Dim MinOride As String
        Dim TotAssigned As Long
        Dim AssignedPct As Single

        TotSpl = 0
        TotSplMin = 0
        TotSplCalcMin = 0
        TotSplOrideMin = 0
        TotHole = 0
        TotHoleMin = 0
        TotHoleMinForced = 0
        TotAssigned = 0

        'Add up the stuff.
        With ssSplitReview
            TotSpl = .MaxRows
            For RowIdx = 1 To .MaxRows
                .Row = RowIdx
                .Col = 4
                MinCalcd = .Text
                .Col = 5
                MinOride = .Text
                If MinCalcd = "M" Or MinOride = "M" Then
                    TotSplMin = TotSplMin + 1
                End If
                If MinCalcd = "M" Then
                    TotSplCalcMin = TotSplCalcMin + 1
                End If
                If MinOride = "M" Then
                    TotSplOrideMin = TotSplOrideMin + 1
                End If
            Next RowIdx
        End With

        With ssCompReview
            TotHole = .MaxRows
            For RowIdx = 1 To .MaxRows
                .Row = RowIdx
                .Col = 4
                MinCalcd = .Text
                If MinCalcd = "M" Or MinCalcd = "MF" Then
                    TotHoleMin = TotHoleMin + 1
                End If
                If MinCalcd = "MF" Then
                    TotHoleMinForced = TotHoleMinForced + 1
                End If
            Next RowIdx
        End With

        With ssResultCnt
            .Col = 1
            .Row = 1
            .Value = TotSpl
            .Row = 2
            .Value = TotSplMin
            .Row = 3
            .Value = TotSplCalcMin
            .Row = 4
            .Value = TotSplOrideMin
            .Row = 5
            .Value = TotHole
            .Row = 6
            .Value = TotHoleMin
            .Row = 7
            .Value = TotHoleMinForced
        End With

        With ssRawProspMin
            TotSpl = 0
            TotAssigned = 0
            For RowIdx = 1 To .MaxRows
                TotSpl = TotSpl + 1
                .Row = RowIdx
                .Col = 4
                If .Text <> "?" Then
                    TotAssigned = TotAssigned + 1
                End If
            Next RowIdx
        End With
        If TotSpl <> 0 Then
            AssignedPct = Round(TotAssigned / TotSpl * 100, 1)
        Else
            AssignedPct = 0
        End If

        lblGen34.Text = Format(TotSpl, "###,##0") & " splits, " & Format(TotAssigned, "###,##0") &
                         " assigned  (" & Format(AssignedPct, "##0.0") & "%)"
    End Sub

    Private Sub GoToHoleInSprd(ByRef aSpread As AxvaSpread,
                               ByVal aHoleLoc As String)

        ''**********************************************************************
        ''
        ''
        ''
        ''**********************************************************************

        Dim RowIdx As Integer

        With aSpread
            For RowIdx = 1 To .MaxRows
                .Row = RowIdx
                .Col = 1
                If .Text = aHoleLoc Then
                    .Action = FPSpread.ActionConstants.ActionActiveCell ' SS_ACTION_ACTIVE_CELL
                    .Action = FPSpread.ActionConstants.ActionGotoCell ' SS_ACTION_GOTO_CELL
                    Exit For
                End If
            Next RowIdx
        End With
    End Sub

    Private Sub ssCompErrors_Click(ByVal sender As System.Object, ByVal e As AxFPSpread._DSpreadEvents_ClickEvent) Handles ssCompErrors.ClickEvent  '(ByVal Col As Long, ByVal Row As Long)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'User has clicked on a row in ssCompErrors.

        Dim RowIdx As Integer
        Dim ThisLoc As String

        If fProcessing <> True Then
            If e.row = 0 Then
                Exit Sub
            End If

            With ssCompErrors
                .Row = e.row
                .Col = 1
                ThisLoc = .Text
            End With
            GoToHoleInSprd(ssSplitReview, ThisLoc)
            GoToHoleInSprd(ssCompReview, ThisLoc)
        End If
    End Sub

    Private Sub cmdReportAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReportAll.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo cmdReportAllClickError

        Dim PctProspect100 As Boolean

        lblRptAllCnt.Text = ""

        PctProspect100 = False
        If opt100Pct.Checked = True Then
            PctProspect100 = True
        Else
            PctProspect100 = False
        End If

        'Could print report using monospaced font
        'Printer.Font.Name = "Arial monospaced for SAP"
        'Printer.FontBold = False
        'Printer.PrintQuality = -4
        'Printer.Print "AAAAAAAAAAAAAAAAAAAA"
        'Printer.Print "aaaaaaaaaaaaaaaaaaaa"
        'Printer.Print "iiiiiiiiiiiiiiiiiiii"
        'Printer.Print "llllllllllllllllllll"
        'Printer.Print "WWWWWWWWWWWWWWWWWWWW"
        'Printer.Print "11111111111111111111"
        'Printer.EndDoc

        'Print report using RichTextBox instead...
        CreateReportAll(PctProspect100, "RTB")

        fraReview.Visible = False
        fraReptDisp.Visible = True

        Exit Sub

cmdReportAllClickError:
        MsgBox("Error printing prospect data from raw data reduction." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Print Error")

        On Error Resume Next
        'rptProspRdctn.Reset
    End Sub

    Private Sub cmdRptAllToTxtFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRptAllToTxtFile.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim PctProspect100 As Boolean
        Dim RptStatus As Boolean

        lblRptAllCnt.Text = ""

        PctProspect100 = False
        If opt100Pct.Checked = True Then
            PctProspect100 = True
        Else
            PctProspect100 = False
        End If

        'Print report using text file.
        RptStatus = CreateReportAll(PctProspect100, "TextFile")

        If RptStatus = True Then
            MsgBox("Text file completed.", vbOKOnly,
                   "Report Status")
        End If
    End Sub

    Private Sub cmdExitRept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        fraReptDisp.Visible = False
        fraReview.Visible = True
    End Sub

    Private Sub cmdPrintRept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo DocPrintError

        SetActionStatus("Printing report...")
        Me.Cursor = Cursors.WaitCursor

        'Printer.Print " "
        'rtbRept1.SelPrint Printer.hDC, 0
        'Printer.EndDoc

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow

        Exit Sub

DocPrintError:
        If Err.Number = 482 Then
            'do nothing
        Else
            Err.Raise(Err.Number)
        End If

        On Error Resume Next
        SetActionStatus("")
        On Error Resume Next
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Function CreateReportAll(ByVal a100Pct As Boolean,
                                     ByVal aRptType As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim RowIdx As Long
        Dim CurrHole As String
        Dim CurrTwp As Integer
        Dim CurrRge As Integer
        Dim CurrSec As Integer
        Dim SplitData As gRawProspSplRdctnType
        Dim HoleData As gRawProspSplRdctnType
        Dim PageNum As Integer
        Dim LineCnt As Long
        Dim MaxLines As Integer
        Dim HoleCnt As Integer
        Dim LastSplit As Boolean
        Dim UseRtf As Boolean
        Dim FileNumber As Integer

        lblRptAllCnt.Text = ""
        CreateReportAll = True

        'aRptType = "RTB"
        'Create a "Report All" report to the rich text box (rtbRept1).
        'This doesn't work so well -- it is way to slow.

        'aRptType = "TextFile"
        'Create a "Report All" report to a text file = txtRptAllToTxtFile.Text

        If aRptType = "TextFile" Then
            'Check to see if the text file exists...
            If Trim(txtRptAllToTxtFile.Text) = "" Or
                Mid(txtRptAllToTxtFile.Text,
                    Len(txtRptAllToTxtFile.Text)) = "\" Then

                MsgBox("You must enter a text file name." +
                        Chr(10) + Chr(10) + "Report not created!",
                        vbExclamation, "Report Status")
                CreateReportAll = False
                Exit Function
            End If

            FileNumber = FreeFile()
            FileOpen(FileNumber, txtRptAllToTxtFile.Text, OpenMode.Output, OpenAccess.Write) ' For Output As #FileNumber
        End If

        SetActionStatus("Generating report...")
        Me.Cursor = Cursors.WaitCursor

        PageNum = 0
        LineCnt = 77
        MaxLines = 77
        LastSplit = False

        CurrHole = ""
        CurrTwp = 0
        CurrRge = 0
        CurrSec = 0
        HoleCnt = 0

        rtbRept1.Text = ""

        With ssSplitReview
            For RowIdx = 1 To .MaxRows
                .Row = RowIdx

                If aRptType = "RTB" Then
                    If RowIdx Mod 10 = 0 Then
                        lblRptAllCnt.Text = Format(RowIdx, "###,##0")
                        'frmProspDataReduction.Refresh()
                    End If
                Else
                    If RowIdx Mod 10 = 0 Then
                        lblRptAllCnt2.Text = Format(RowIdx, "###,##0")
                        'frmProspDataReduction.Refresh()
                    End If
                End If

                'Get the split data
                SplitData = gGetDataFromReviewSprd(ssSplitReview, RowIdx)

                'Get the hole data (if split = 1)
                If SplitData.SplitNumber = 1 Then
                    HoleCnt = HoleCnt + 1
                    HoleData = gGetDataFromReviewSprd(ssCompReview, HoleCnt)
                End If

                '1) AddReportAllPageHdr
                '2) AddReportAllSplitHdr
                '3) AddReportAllSplit
                '4) AddReportAllHoleHdr
                '5) AddReportAllHole

                With SplitData
                    'Start of hole -- put a page header!
                    If .SplitNumber = 1 Then
                        PageNum = PageNum + 1
                        AddReportAllPageHdr(SplitData,
                                            HoleData,
                                            PageNum,
                                            LineCnt,
                                            MaxLines,
                                            aRptType,
                                            FileNumber)
                        LineCnt = 6
                    End If

                    'Add a header for this split
                    'If there isn't room for an entire split then go to next page
                    'first.
                    If MaxLines - LineCnt < 21 Then
                        PageNum = PageNum + 1
                        AddReportAllPageHdr(SplitData,
                                            HoleData,
                                            PageNum,
                                            LineCnt,
                                            MaxLines,
                                            aRptType,
                                            FileNumber)
                        LineCnt = 6
                    End If
                    AddReportAllSplitHdr(SplitData,
                                         aRptType,
                                         FileNumber)
                    AddReportAllSplit(SplitData,
                                      aRptType,
                                      FileNumber)
                    LineCnt = LineCnt + 21

                    'Need to take a peek at the next split to see if we have just
                    'finished the last split for this hole!
                    LastSplit = False
                    If RowIdx = ssSplitReview.MaxRows Then
                        LastSplit = True
                    Else
                        ssSplitReview.Row = RowIdx + 1
                        ssSplitReview.Col = 2
                        If ssSplitReview.Value = 1 Then
                            LastSplit = True
                        End If
                    End If

                    If LastSplit = True Then
                        'We have just finished adding the splits for a hole -- need
                        'to add the hole data.
                        If MaxLines - LineCnt < 21 Then
                            PageNum = PageNum + 1
                            AddReportAllPageHdr(SplitData,
                                                HoleData,
                                                PageNum,
                                                LineCnt,
                                                MaxLines,
                                                aRptType,
                                                FileNumber)
                            LineCnt = 6
                        End If

                        AddReportAllHoleHdr(HoleData,
                                            aRptType,
                                            FileNumber)
                        AddReportAllHole(HoleData,
                                         aRptType,
                                         FileNumber)
                        LineCnt = LineCnt + 21
                    End If
                End With
            Next RowIdx
        End With

        If aRptType = "TextFile" Then
            FileClose(FileNumber)
        End If

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Function

    Private Sub AddReportAllPageHdr(ByRef aSplitData As gRawProspSplRdctnType,
                                    ByRef aHoleData As gRawProspSplRdctnType,
                                    ByVal aPageNum As Integer,
                                    ByVal aLineCnt As Long,
                                    ByVal aMaxLines As Integer,
                                    ByVal aRptType As String,
                                    ByVal aFileNumber As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim HoleDescTitled As String
        Dim LineIdx As Integer
        Dim TextStr As String

        'Adds 6 lines to text file.

        '1) txtProspectDatasetName.Text  30 characters max
        '2) txtProspectDatasetDesc.Text  200 characters max
        '3) txtAreaDefnName.Text         30 characters max
        '4) txtPsizeDefnName.Text        30 characters max
        '5) txtRcvryEtcName.Text         30 characters max
        '6) txtSplitOverrideName.text    30 characters max

        With aSplitData
            HoleDescTitled = gGetHoleLocationTitled2(.Section, .Township,
                             .Range, .HoleLocation)

            'Make sure we are at the top of a page!
            For LineIdx = 1 To aMaxLines - aLineCnt
                If aRptType = "RTB" Then
                    rtbRept1.Text = rtbRept1.Text & vbCrLf
                Else
                    gWriteLine(aFileNumber, " ")
                End If
            Next LineIdx

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    Format(Now, "MM/dd/yyyy  hh:mm AM/PM") &
                    "                 Prospect Data Reduction Splits & Holes" &
                    "                     Page " & Format(aPageNum, "#,##0") &
                    vbCrLf &
                    Trim(txtProspectDatasetName.Text) & "  " &
                    Mid(Trim(txtProspectDatasetDesc.Text), 1, 73) & vbCrLf &
                    "Area=" & Mid(Trim(_areaDefinitionForm.AreaDefinition.AreaDefinitionName), 1, 25) & " " &
                    "Prdsize=" & Mid(Trim(_productSizeDesginationForm.ProductSizeDesignation.ProductSizeDesignationName), 1, 25) & " " &
                    "PrdAdj=" & Mid(Trim(_recoveryScenariosForm.ProductRecoveryDefinition.ScenarioName), 1, 25) & " " &
                    "Oride=" & Mid(Trim(txtSplitOverrideName.Text), 1, 25) &
                    vbCrLf & HoleDescTitled & "  " & "AOI " & Format(aHoleData.Aoi, "#,##0.0") &
                    "  Prosp date " & Format(.ProspDate, "mm/dd/yy") &
                    "  Own = " & aHoleData.Ownership & "  Ovb = " & Format(aHoleData.OvbThk, "##0.0") & "'" &
                    " Tot dpth = " & Format(aHoleData.TotDepth, "#,##0.0") & "'" & vbCrLf &
                    "----------------------------------------" &
                    "----------------------------------------" &
                    "------------------------" & vbCrLf & vbCrLf
            Else
                gWriteLine(aFileNumber, " ")
                TextStr = Format(Now, "MM/dd/yyyy  hh:mm AM/PM") &
                    "                 Prospect Data Reduction Splits & Holes" &
                    "                     Page " & Format(aPageNum, "#,##0")
                gWriteLine(aFileNumber, TextStr)

                TextStr = Trim(txtProspectDatasetName.Text) & "  " &
                    Mid(Trim(txtProspectDatasetDesc.Text), 1, 73)
                gWriteLine(aFileNumber, TextStr)

                TextStr = "Area=" & Mid(Trim(_areaDefinitionForm.AreaDefinition.AreaDefinitionName), 1, 25) & " " &
                    "Prdsize=" & Mid(Trim(_productSizeDesginationForm.ProductSizeDesignation.ProductSizeDesignationName), 1, 25) & " " &
                    "PrdAdj=" & Mid(Trim(_recoveryScenariosForm.ProductRecoveryDefinition.ScenarioName), 1, 25) & " " &
                    "Oride=" & Mid(Trim(txtSplitOverrideName.Text), 1, 25)
                gWriteLine(aFileNumber, TextStr)

                TextStr = HoleDescTitled & "  " & "AOI " & Format(aHoleData.Aoi, "#,##0.0") &
                    "  Prosp date " & Format(.ProspDate, "mm/dd/yy") &
                    "  Own = " & aHoleData.Ownership & "  Ovb = " & Format(aHoleData.OvbThk, "##0.0") & "'" &
                    " Tot dpth = " & Format(aHoleData.TotDepth, "#,##0.0") & "'"
                gWriteLine(aFileNumber, TextStr)

                TextStr = "----------------------------------------" &
                    "----------------------------------------" &
                    "------------------------"
                gWriteLine(aFileNumber, TextStr)
                gWriteLine(aFileNumber, " ")
            End If
        End With
    End Sub

    Private Sub AddReportAllSplitHdr(ByRef aSplitData As gRawProspSplRdctnType,
                                     ByVal aRptType As String,
                                     ByVal aFileNumber As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim Mineability As String
        Dim TextStr As String

        'This is the "Split" header part of the "Report All" report to the rich text box
        '(rtbRept1).
        'Adds 2 lines to text file.

        With aSplitData
            Mineability = "**UNDETERMINED**"
            Select Case .MineableOride
                Case Is = "M"
                    Mineability = "**MINEABLE**"
                Case Is = "U"
                    Mineability = "**UNMINEABLE**"
                Case Is = "C"
                    If .MineableCalcd = "M" Then
                        Mineability = "**MINEABLE**"
                    End If
                    If .MineableCalcd = "U" Then
                        Mineability = "**UNMINEABLE**"
                    End If
            End Select

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "** SPLIT # " & Format(.SplitNumber, "#0") & " **" &
                    "    From " & Format(.SplitDepthTop, "##0.0") & "'" &
                    " To " & Format(.SplitDepthBot, "##0.0") & "'" &
                    "  Thick = " & Format(.SplitThck, "#0.0") & "'" &
                    "    Density (lab) = " & Format(.MtxDensity, "##0.0") &
                    "    " & Mineability &
                    vbCrLf & vbCrLf
            Else
                TextStr = "** SPLIT # " & Format(.SplitNumber, "#0") & " **" &
                    "    From " & Format(.SplitDepthTop, "##0.0") & "'" &
                    " To " & Format(.SplitDepthBot, "##0.0") & "'" &
                    "  Thick = " & Format(.SplitThck, "#0.0") & "'" &
                    "    Density(lab) = " & Format(.MtxDensity, "##0.0") &
                    "    " & Mineability
                gWriteLine(aFileNumber, TextStr)
                gWriteLine(aFileNumber, " ")
            End If
        End With
    End Sub

    Private Sub AddReportAllSplit(ByRef aSplitData As gRawProspSplRdctnType,
                                  ByVal aRptType As String,
                                  ByVal aFileNumber As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim MatlTitle As String
        Dim TextStr As String

        'This is the "Split" data part of the "Report All" report to the rich text box
        '(rtbRept1).
        'Adds 19 lines to text file

        With aSplitData
            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "*Off Spec             Product Coefficient                            100% Prospect                " & vbCrLf &
                    "       ---------------------------------------------   ---------------------------------------------" & vbCrLf &
                    "          TPA   Pct  BPL  Ins   Ca    Fe    Al    Mg      TPA   Pct  BPL  Ins   Ca    Fe    Al    Mg" & vbCrLf &
                    "       ------ ----- ---- ---- ---- ----- ----- -----   ------ ----- ---- ---- ---- ----- ----- -----" & vbCrLf
            Else
                TextStr = "*Off Spec             Product Coefficient                            100% Prospect                "
                gWriteLine(aFileNumber, TextStr)

                TextStr = "       ---------------------------------------------   ---------------------------------------------"
                gWriteLine(aFileNumber, TextStr)

                TextStr = "          TPA   Pct  BPL  Ins   Ca    Fe    Al    Mg      TPA   Pct  BPL  Ins   Ca    Fe    Al    Mg"
                gWriteLine(aFileNumber, TextStr)

                TextStr = "       ------ ----- ---- ---- ---- ----- ----- -----   ------ ----- ---- ---- ---- ----- ----- -----"
                gWriteLine(aFileNumber, TextStr)
            End If

            'Oversize
            If .OsOnSpec = "No" Then
                MatlTitle = "O-size*"
            Else
                MatlTitle = "O-size "
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    MatlTitle & FormatSpec(.Os.Tpa, "##,##0") & " " & FormatSpec(.Os.WtPct, "##0.0") & " " &
                                FormatSpec(.Os.Bpl, "#0.0") & " " & FormatSpec(.Os.Ins, "#0.0") & " " &
                                FormatSpec(.Os.Ca, "#0.0") & " " & FormatSpec(.Os.Fe, "#0.00") & " " &
                                FormatSpec(.Os.Al, "#0.00") & " " & FormatSpec(.Os.Mg, "#0.00") & "   " &
                                FormatSpec(.Os100.Tpa, "##,##0") & " " & FormatSpec(.Os100.WtPct, "##0.0") & " " &
                                FormatSpec(.Os100.Bpl, "#0.0") & " " & FormatSpec(.Os100.Ins, "#0.0") & " " &
                                FormatSpec(.Os100.Ca, "#0.0") & " " & FormatSpec(.Os100.Fe, "#0.00") & " " &
                                FormatSpec(.Os100.Al, "#0.00") & " " & FormatSpec(.Os100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = MatlTitle & FormatSpec(.Os.Tpa, "##,##0") & " " & FormatSpec(.Os.WtPct, "##0.0") & " " &
                                      FormatSpec(.Os.Bpl, "#0.0") & " " & FormatSpec(.Os.Ins, "#0.0") & " " &
                                      FormatSpec(.Os.Ca, "#0.0") & " " & FormatSpec(.Os.Fe, "#0.00") & " " &
                                      FormatSpec(.Os.Al, "#0.00") & " " & FormatSpec(.Os.Mg, "#0.00") & "   " &
                                      FormatSpec(.Os100.Tpa, "##,##0") & " " & FormatSpec(.Os100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Os100.Bpl, "#0.0") & " " & FormatSpec(.Os100.Ins, "#0.0") & " " &
                                      FormatSpec(.Os100.Ca, "#0.0") & " " & FormatSpec(.Os100.Fe, "#0.00") & " " &
                                      FormatSpec(.Os100.Al, "#0.00") & " " & FormatSpec(.Os100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            'Coarse pebble
            If .CpbOnSpec = "No" Then
                MatlTitle = "Crs pb*"
            Else
                MatlTitle = "Crs pb "
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    MatlTitle & FormatSpec(.Cpb.Tpa, "##,##0") & " " & FormatSpec(.Cpb.WtPct, "##0.0") & " " &
                                FormatSpec(.Cpb.Bpl, "#0.0") & " " & FormatSpec(.Cpb.Ins, "#0.0") & " " &
                                FormatSpec(.Cpb.Ca, "#0.0") & " " & FormatSpec(.Cpb.Fe, "#0.00") & " " &
                                FormatSpec(.Cpb.Al, "#0.00") & " " & FormatSpec(.Cpb.Mg, "#0.00") & "   " &
                                FormatSpec(.Cpb100.Tpa, "##,##0") & " " & FormatSpec(.Cpb100.WtPct, "##0.0") & " " &
                                FormatSpec(.Cpb100.Bpl, "#0.0") & " " & FormatSpec(.Cpb100.Ins, "#0.0") & " " &
                                FormatSpec(.Cpb100.Ca, "#0.0") & " " & FormatSpec(.Cpb100.Fe, "#0.00") & " " &
                                FormatSpec(.Cpb100.Al, "#0.00") & " " & FormatSpec(.Cpb100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = MatlTitle & FormatSpec(.Cpb.Tpa, "##,##0") & " " & FormatSpec(.Cpb.WtPct, "##0.0") & " " &
                                      FormatSpec(.Cpb.Bpl, "#0.0") & " " & FormatSpec(.Cpb.Ins, "#0.0") & " " &
                                      FormatSpec(.Cpb.Ca, "#0.0") & " " & FormatSpec(.Cpb.Fe, "#0.00") & " " &
                                      FormatSpec(.Cpb.Al, "#0.00") & " " & FormatSpec(.Cpb.Mg, "#0.00") & "   " &
                                      FormatSpec(.Cpb100.Tpa, "##,##0") & " " & FormatSpec(.Cpb100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Cpb100.Bpl, "#0.0") & " " & FormatSpec(.Cpb100.Ins, "#0.0") & " " &
                                      FormatSpec(.Cpb100.Ca, "#0.0") & " " & FormatSpec(.Cpb100.Fe, "#0.00") & " " &
                                      FormatSpec(.Cpb100.Al, "#0.00") & " " & FormatSpec(.Cpb100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            'Fine pebble
            If .FpbOnSpec = "No" Then
                MatlTitle = "Fne pb*"
            Else
                MatlTitle = "Fne pb "
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    MatlTitle & FormatSpec(.Fpb.Tpa, "##,##0") & " " & FormatSpec(.Fpb.WtPct, "##0.0") & " " &
                                FormatSpec(.Fpb.Bpl, "#0.0") & " " & FormatSpec(.Fpb.Ins, "#0.0") & " " &
                                FormatSpec(.Fpb.Ca, "#0.0") & " " & FormatSpec(.Fpb.Fe, "#0.00") & " " &
                                FormatSpec(.Fpb.Al, "#0.00") & " " & FormatSpec(.Fpb.Mg, "#0.00") & "   " &
                                FormatSpec(.Fpb100.Tpa, "##,##0") & " " & FormatSpec(.Fpb100.WtPct, "##0.0") & " " &
                                FormatSpec(.Fpb100.Bpl, "#0.0") & " " & FormatSpec(.Fpb100.Ins, "#0.0") & " " &
                                FormatSpec(.Fpb100.Ca, "#0.0") & " " & FormatSpec(.Fpb100.Fe, "#0.00") & " " &
                                FormatSpec(.Fpb100.Al, "#0.00") & " " & FormatSpec(.Fpb100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = MatlTitle & FormatSpec(.Fpb.Tpa, "##,##0") & " " & FormatSpec(.Fpb.WtPct, "##0.0") & " " &
                                      FormatSpec(.Fpb.Bpl, "#0.0") & " " & FormatSpec(.Fpb.Ins, "#0.0") & " " &
                                      FormatSpec(.Fpb.Ca, "#0.0") & " " & FormatSpec(.Fpb.Fe, "#0.00") & " " &
                                      FormatSpec(.Fpb.Al, "#0.00") & " " & FormatSpec(.Fpb.Mg, "#0.00") & "   " &
                                      FormatSpec(.Fpb100.Tpa, "##,##0") & " " & FormatSpec(.Fpb100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Fpb100.Bpl, "#0.0") & " " & FormatSpec(.Fpb100.Ins, "#0.0") & " " &
                                      FormatSpec(.Fpb100.Ca, "#0.0") & " " & FormatSpec(.Fpb100.Fe, "#0.00") & " " &
                                      FormatSpec(.Fpb100.Al, "#0.00") & " " & FormatSpec(.Fpb100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            'Fine pebble
            If .TpbOnSpec = "No" Then
                MatlTitle = "Tot pb*"
            Else
                MatlTitle = "Tot pb "
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    MatlTitle & FormatSpec(.Tpb.Tpa, "##,##0") & " " & FormatSpec(.Tpb.WtPct, "##0.0") & " " &
                                FormatSpec(.Tpb.Bpl, "#0.0") & " " & FormatSpec(.Tpb.Ins, "#0.0") & " " &
                                FormatSpec(.Tpb.Ca, "#0.0") & " " & FormatSpec(.Tpb.Fe, "#0.00") & " " &
                                FormatSpec(.Tpb.Al, "#0.00") & " " & FormatSpec(.Tpb.Mg, "#0.00") & "   " &
                                FormatSpec(.Tpb100.Tpa, "##,##0") & " " & FormatSpec(.Tpb100.WtPct, "##0.0") & " " &
                                FormatSpec(.Tpb100.Bpl, "#0.0") & " " & FormatSpec(.Tpb100.Ins, "#0.0") & " " &
                                FormatSpec(.Tpb100.Ca, "#0.0") & " " & FormatSpec(.Tpb100.Fe, "#0.00") & " " &
                                FormatSpec(.Tpb100.Al, "#0.00") & " " & FormatSpec(.Tpb100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = MatlTitle & FormatSpec(.Tpb.Tpa, "##,##0") & " " & FormatSpec(.Tpb.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tpb.Bpl, "#0.0") & " " & FormatSpec(.Tpb.Ins, "#0.0") & " " &
                                      FormatSpec(.Tpb.Ca, "#0.0") & " " & FormatSpec(.Tpb.Fe, "#0.00") & " " &
                                      FormatSpec(.Tpb.Al, "#0.00") & " " & FormatSpec(.Tpb.Mg, "#0.00") & "   " &
                                      FormatSpec(.Tpb100.Tpa, "##,##0") & " " & FormatSpec(.Tpb100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tpb100.Bpl, "#0.0") & " " & FormatSpec(.Tpb100.Ins, "#0.0") & " " &
                                      FormatSpec(.Tpb100.Ca, "#0.0") & " " & FormatSpec(.Tpb100.Fe, "#0.00") & " " &
                                      FormatSpec(.Tpb100.Al, "#0.00") & " " & FormatSpec(.Tpb100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            'Coarse concentrate
            If .CcnOnSpec = "No" Then
                MatlTitle = "Crs cn*"
            Else
                MatlTitle = "Crs cn "
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    MatlTitle & FormatSpec(.Ccn.Tpa, "##,##0") & " " & FormatSpec(.Ccn.WtPct, "##0.0") & " " &
                                FormatSpec(.Ccn.Bpl, "#0.0") & " " & FormatSpec(.Ccn.Ins, "#0.0") & " " &
                                FormatSpec(.Ccn.Ca, "#0.0") & " " & FormatSpec(.Ccn.Fe, "#0.00") & " " &
                                FormatSpec(.Ccn.Al, "#0.00") & " " & FormatSpec(.Ccn.Mg, "#0.00") & "   " &
                                FormatSpec(.Ccn100.Tpa, "##,##0") & " " & FormatSpec(.Ccn100.WtPct, "##0.0") & " " &
                                FormatSpec(.Ccn100.Bpl, "#0.0") & " " & FormatSpec(.Ccn100.Ins, "#0.0") & " " &
                                FormatSpec(.Ccn100.Ca, "#0.0") & " " & FormatSpec(.Ccn100.Fe, "#0.00") & " " &
                                FormatSpec(.Ccn100.Al, "#0.00") & " " & FormatSpec(.Ccn100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = MatlTitle & FormatSpec(.Ccn.Tpa, "##,##0") & " " & FormatSpec(.Ccn.WtPct, "##0.0") & " " &
                                      FormatSpec(.Ccn.Bpl, "#0.0") & " " & FormatSpec(.Ccn.Ins, "#0.0") & " " &
                                      FormatSpec(.Ccn.Ca, "#0.0") & " " & FormatSpec(.Ccn.Fe, "#0.00") & " " &
                                      FormatSpec(.Ccn.Al, "#0.00") & " " & FormatSpec(.Ccn.Mg, "#0.00") & "   " &
                                      FormatSpec(.Ccn100.Tpa, "##,##0") & " " & FormatSpec(.Ccn100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Ccn100.Bpl, "#0.0") & " " & FormatSpec(.Ccn100.Ins, "#0.0") & " " &
                                      FormatSpec(.Ccn100.Ca, "#0.0") & " " & FormatSpec(.Ccn100.Fe, "#0.00") & " " &
                                      FormatSpec(.Ccn100.Al, "#0.00") & " " & FormatSpec(.Ccn100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            'Fine concentrate
            If .FcnOnSpec = "No" Then
                MatlTitle = "Fne cn*"
            Else
                MatlTitle = "Fne cn "
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    MatlTitle & FormatSpec(.Fcn.Tpa, "##,##0") & " " & FormatSpec(.Fcn.WtPct, "##0.0") & " " &
                                FormatSpec(.Fcn.Bpl, "#0.0") & " " & FormatSpec(.Fcn.Ins, "#0.0") & " " &
                                FormatSpec(.Fcn.Ca, "#0.0") & " " & FormatSpec(.Fcn.Fe, "#0.00") & " " &
                                FormatSpec(.Fcn.Al, "#0.00") & " " & FormatSpec(.Fcn.Mg, "#0.00") & "   " &
                                FormatSpec(.Fcn100.Tpa, "##,##0") & " " & FormatSpec(.Fcn100.WtPct, "##0.0") & " " &
                                FormatSpec(.Fcn100.Bpl, "#0.0") & " " & FormatSpec(.Fcn100.Ins, "#0.0") & " " &
                                FormatSpec(.Fcn100.Ca, "#0.0") & " " & FormatSpec(.Fcn100.Fe, "#0.00") & " " &
                                FormatSpec(.Fcn100.Al, "#0.00") & " " & FormatSpec(.Fcn100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = MatlTitle & FormatSpec(.Fcn.Tpa, "##,##0") & " " & FormatSpec(.Fcn.WtPct, "##0.0") & " " &
                                      FormatSpec(.Fcn.Bpl, "#0.0") & " " & FormatSpec(.Fcn.Ins, "#0.0") & " " &
                                      FormatSpec(.Fcn.Ca, "#0.0") & " " & FormatSpec(.Fcn.Fe, "#0.00") & " " &
                                      FormatSpec(.Fcn.Al, "#0.00") & " " & FormatSpec(.Fcn.Mg, "#0.00") & "   " &
                                      FormatSpec(.Fcn100.Tpa, "##,##0") & " " & FormatSpec(.Fcn100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Fcn100.Bpl, "#0.0") & " " & FormatSpec(.Fcn100.Ins, "#0.0") & " " &
                                      FormatSpec(.Fcn100.Ca, "#0.0") & " " & FormatSpec(.Fcn100.Fe, "#0.00") & " " &
                                      FormatSpec(.Fcn100.Al, "#0.00") & " " & FormatSpec(.Fcn100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            'Total concentrate
            If .TcnOnSpec = "No" Then
                MatlTitle = "Tot cn*"
            Else
                MatlTitle = "Tot cn "
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    MatlTitle & FormatSpec(.Tcn.Tpa, "##,##0") & " " & FormatSpec(.Tcn.WtPct, "##0.0") & " " &
                                FormatSpec(.Tcn.Bpl, "#0.0") & " " & FormatSpec(.Tcn.Ins, "#0.0") & " " &
                                FormatSpec(.Tcn.Ca, "#0.0") & " " & FormatSpec(.Tcn.Fe, "#0.00") & " " &
                                FormatSpec(.Tcn.Al, "#0.00") & " " & FormatSpec(.Tcn.Mg, "#0.00") & "   " &
                                FormatSpec(.Tcn100.Tpa, "##,##0") & " " & FormatSpec(.Tcn100.WtPct, "##0.0") & " " &
                                FormatSpec(.Tcn100.Bpl, "#0.0") & " " & FormatSpec(.Tcn100.Ins, "#0.0") & " " &
                                FormatSpec(.Tcn100.Ca, "#0.0") & " " & FormatSpec(.Tcn100.Fe, "#0.00") & " " &
                                FormatSpec(.Tcn100.Al, "#0.00") & " " & FormatSpec(.Tcn100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = MatlTitle & FormatSpec(.Tcn.Tpa, "##,##0") & " " & FormatSpec(.Tcn.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tcn.Bpl, "#0.0") & " " & FormatSpec(.Tcn.Ins, "#0.0") & " " &
                                      FormatSpec(.Tcn.Ca, "#0.0") & " " & FormatSpec(.Tcn.Fe, "#0.00") & " " &
                                      FormatSpec(.Tcn.Al, "#0.00") & " " & FormatSpec(.Tcn.Mg, "#0.00") & "   " &
                                      FormatSpec(.Tcn100.Tpa, "##,##0") & " " & FormatSpec(.Tcn100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tcn100.Bpl, "#0.0") & " " & FormatSpec(.Tcn100.Ins, "#0.0") & " " &
                                      FormatSpec(.Tcn100.Ca, "#0.0") & " " & FormatSpec(.Tcn100.Fe, "#0.00") & " " &
                                      FormatSpec(.Tcn100.Al, "#0.00") & " " & FormatSpec(.Tcn100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            'Total product
            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Tot pr " & FormatSpec(.Tpr.Tpa, "##,##0") & " " & FormatSpec(.Tpr.WtPct, "##0.0") & " " &
                                FormatSpec(.Tpr.Bpl, "#0.0") & " " & FormatSpec(.Tpr.Ins, "#0.0") & " " &
                                FormatSpec(.Tpr.Ca, "#0.0") & " " & FormatSpec(.Tpr.Fe, "#0.00") & " " &
                                FormatSpec(.Tpr.Al, "#0.00") & " " & FormatSpec(.Tpr.Mg, "#0.00") & "   " &
                                FormatSpec(.Tpr100.Tpa, "##,##0") & " " & FormatSpec(.Tpr100.WtPct, "##0.0") & " " &
                                FormatSpec(.Tpr100.Bpl, "#0.0") & " " & FormatSpec(.Tpr100.Ins, "#0.0") & " " &
                                FormatSpec(.Tpr100.Ca, "#0.0") & " " & FormatSpec(.Tpr100.Fe, "#0.00") & " " &
                                FormatSpec(.Tpr100.Al, "#0.00") & " " & FormatSpec(.Tpr100.Mg, "#0.00") & vbCrLf
                rtbRept1.Text = rtbRept1.Text & vbCrLf
            Else
                TextStr = "Tot pr " & FormatSpec(.Tpr.Tpa, "##,##0") & " " & FormatSpec(.Tpr.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tpr.Bpl, "#0.0") & " " & FormatSpec(.Tpr.Ins, "#0.0") & " " &
                                      FormatSpec(.Tpr.Ca, "#0.0") & " " & FormatSpec(.Tpr.Fe, "#0.00") & " " &
                                      FormatSpec(.Tpr.Al, "#0.00") & " " & FormatSpec(.Tpr.Mg, "#0.00") & "   " &
                                      FormatSpec(.Tpr100.Tpa, "##,##0") & " " & FormatSpec(.Tpr100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tpr100.Bpl, "#0.0") & " " & FormatSpec(.Tpr100.Ins, "#0.0") & " " &
                                      FormatSpec(.Tpr100.Ca, "#0.0") & " " & FormatSpec(.Tpr100.Fe, "#0.00") & " " &
                                      FormatSpec(.Tpr100.Al, "#0.00") & " " & FormatSpec(.Tpr100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
                gWriteLine(aFileNumber, " ")
            End If

            'Tailings
            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Tails  " & FormatSpec(.Ttl.Tpa, "##,##0") & " " & FormatSpec(.Ttl.WtPct, "##0.0") & " " &
                                FormatSpec(.Ttl.Bpl, "#0.0") &
                                "    MtxX OnSpec = " & FormatSpec(.MtxxOnSpec, "##0.00") &
                                "       " &
                                FormatSpec(.Ttl100.Tpa, "##,##0") & " " &
                                FormatSpec(.Ttl100.WtPct, "##0.0") & " " & FormatSpec(.Ttl.Bpl, "#0.0") &
                                "    MtxX OnSpec = " & FormatSpec(.MtxxOnSpec100, "##0.00") &
                                vbCrLf
            Else
                TextStr = "Tails  " & FormatSpec(.Ttl.Tpa, "##,##0") & " " & FormatSpec(.Ttl.WtPct, "##0.0") & " " &
                                      FormatSpec(.Ttl.Bpl, "#0.0") &
                                      "    MtxX OnSpec = " & FormatSpec(.MtxxOnSpec, "##0.00") &
                                      "       " &
                                      FormatSpec(.Ttl100.Tpa, "##,##0") & " " &
                                      FormatSpec(.Ttl100.WtPct, "##0.0") & " " & FormatSpec(.Ttl.Bpl, "#0.0") &
                                      "    MtxX OnSpec = " & FormatSpec(.MtxxOnSpec100, "##0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            'Waste clay
            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Wst cl " & FormatSpec(.Wcl.Tpa, "##,##0") & " " & FormatSpec(.Wcl.WtPct, "##0.0") & " " &
                                FormatSpec(.Wcl.Bpl, "#0.0") &
                                "    MtxX All    = " & FormatSpec(.MtxxAll, "##0.00") &
                                "       " &
                                FormatSpec(.Wcl100.Tpa, "##,##0") & " " &
                                FormatSpec(.Wcl100.WtPct, "##0.0") & " " & FormatSpec(.Wcl.Bpl, "#0.0") &
                                "    MtxX All    = " & FormatSpec(.MtxxAll100, "##0.00") &
                                vbCrLf
            Else
                TextStr = "Wst cl " & FormatSpec(.Wcl.Tpa, "##,##0") & " " & FormatSpec(.Wcl.WtPct, "##0.0") & " " &
                                      FormatSpec(.Wcl.Bpl, "#0.0") &
                                      "    MtxX All    = " & FormatSpec(.MtxxAll, "##0.00") &
                                      "       " &
                                      FormatSpec(.Wcl100.Tpa, "##,##0") & " " &
                                      FormatSpec(.Wcl100.WtPct, "##0.0") & " " & FormatSpec(.Wcl.Bpl, "#0.0") &
                                      "    MtxX All    = " & FormatSpec(.MtxxAll100, "##0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            'Coarse feed
            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Crs fd " & FormatSpec(.Cfd.Tpa, "##,##0") & " " & FormatSpec(.Cfd.WtPct, "##0.0") & " " &
                                FormatSpec(.Cfd.Bpl, "#0.0") &
                                "                               " &
                                FormatSpec(.Cfd100.Tpa, "##,##0") & " " &
                                FormatSpec(.Cfd100.WtPct, "##0.0") & " " & FormatSpec(.Cfd.Bpl, "#0.0") & vbCrLf
            Else
                TextStr = "Crs fd " & FormatSpec(.Cfd.Tpa, "##,##0") & " " & FormatSpec(.Cfd.WtPct, "##0.0") & " " &
                                      FormatSpec(.Cfd.Bpl, "#0.0") &
                                      "                               " &
                                      FormatSpec(.Cfd100.Tpa, "##,##0") & " " &
                                      FormatSpec(.Cfd100.WtPct, "##0.0") & " " & FormatSpec(.Cfd.Bpl, "#0.0")
                gWriteLine(aFileNumber, TextStr)
            End If

            'Fine feed
            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Fne fd " & FormatSpec(.Ffd.Tpa, "##,##0") & " " & FormatSpec(.Ffd.WtPct, "##0.0") & " " &
                                FormatSpec(.Ffd.Bpl, "#0.0") &
                                "                               " &
                                FormatSpec(.Ffd100.Tpa, "##,##0") & " " &
                                FormatSpec(.Ffd100.WtPct, "##0.0") & " " & FormatSpec(.Ffd.Bpl, "#0.0") & vbCrLf
            Else
                TextStr = "Fne fd " & FormatSpec(.Ffd.Tpa, "##,##0") & " " & FormatSpec(.Ffd.WtPct, "##0.0") & " " &
                                      FormatSpec(.Ffd.Bpl, "#0.0") &
                                      "                               " &
                                      FormatSpec(.Ffd100.Tpa, "##,##0") & " " &
                                      FormatSpec(.Ffd100.WtPct, "##0.0") & " " & FormatSpec(.Ffd.Bpl, "#0.0")
                gWriteLine(aFileNumber, TextStr)
            End If

            'Total feed
            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Tot fd " & FormatSpec(.Tfd.Tpa, "##,##0") & " " & FormatSpec(.Tfd.WtPct, "##0.0") & " " &
                                FormatSpec(.Tfd.Bpl, "#0.0") &
                                "                               " &
                                FormatSpec(.Tfd100.Tpa, "##,##0") & " " &
                                FormatSpec(.Tfd100.WtPct, "##0.0") & " " & FormatSpec(.Tfd.Bpl, "#0.0") & vbCrLf

                rtbRept1.Text = rtbRept1.Text & vbCrLf
            Else
                TextStr = "Tot fd " & FormatSpec(.Tfd.Tpa, "##,##0") & " " & FormatSpec(.Tfd.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tfd.Bpl, "#0.0") &
                                      "                               " &
                                      FormatSpec(.Tfd100.Tpa, "##,##0") & " " &
                                      FormatSpec(.Tfd100.WtPct, "##0.0") & " " & FormatSpec(.Tfd.Bpl, "#0.0")
                gWriteLine(aFileNumber, TextStr)
                gWriteLine(aFileNumber, " ")
            End If
        End With
    End Sub

    Private Sub AddReportAllHoleHdr(ByRef aHoleData As gRawProspSplRdctnType,
                                    ByVal aRptType As String,
                                    ByVal aFileNumber As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim Mineability As String
        Dim TextStr As String

        'This is the "Hole" header part of the "Report All" report to the rich text box
        '(rtbRept1).
        'Adds 2 lines to text file.

        With aHoleData
            Mineability = "**UNDETERMINED**"
            Select Case .MineableHole
                Case Is = "M"
                    Mineability = "**MINEABLE**"
                Case Is = "U"
                    Mineability = "**UNMINEABLE**"
            End Select

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "** HOLE TOTAL **" &
                    "    Ovb " & Format(.OvbThk, "##0.0") & "'" &
                    "  Mtx " & Format(.MtxThk, "##0.0") & "'" &
                    "  Itb " & Format(.ItbThk, "##0.0") & "'" &
                    "    Density (lab) = " & Format(.MtxDensity, "##0.0") &
                    "    " & Mineability &
                    vbCrLf & vbCrLf
            Else
                TextStr = "** HOLE TOTAL **" &
                    "    Ovb " & Format(.OvbThk, "##0.0") & "'" &
                    "  Mtx " & Format(.MtxThk, "##0.0") & "'" &
                    "  Itb " & Format(.ItbThk, "##0.0") & "'" &
                    "    Density (lab) = " & Format(.MtxDensity, "##0.0") &
                    "    " & Mineability
                gWriteLine(aFileNumber, TextStr)
                gWriteLine(aFileNumber, " ")
            End If
        End With
    End Sub

    Private Sub AddReportAllHole(ByRef aHoleData As gRawProspSplRdctnType,
                                 ByVal aRptType As String,
                                 ByVal aFileNumber As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim TextStr As String

        'This is the "Hole" data part of the "Report All" report to the rich text box
        '(rtbRept1).
        'Adds 19 lines to text file

        With aHoleData
            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "                      Product Coefficient                            100% Prospect                " & vbCrLf &
                    "       ---------------------------------------------   ---------------------------------------------" & vbCrLf &
                    "          TPA   Pct  BPL  Ins   Ca    Fe    Al    Mg      TPA   Pct  BPL  Ins   Ca    Fe    Al    Mg" & vbCrLf &
                    "       ------ ----- ---- ---- ---- ----- ----- -----   ------ ----- ---- ---- ---- ----- ----- -----" & vbCrLf
            Else
                TextStr = "                      Product Coefficient                            100% Prospect                "
                gWriteLine(aFileNumber, TextStr)

                TextStr = "       ---------------------------------------------   ---------------------------------------------"
                gWriteLine(aFileNumber, TextStr)

                TextStr = "          TPA   Pct  BPL  Ins   Ca    Fe    Al    Mg      TPA   Pct  BPL  Ins   Ca    Fe    Al    Mg"
                gWriteLine(aFileNumber, TextStr)

                TextStr = "       ------ ----- ---- ---- ---- ----- ----- -----   ------ ----- ---- ---- ---- ----- ----- -----"
                gWriteLine(aFileNumber, TextStr)
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "O-size " & FormatSpec(.Os.Tpa, "##,##0") & " " & FormatSpec(.Os.WtPct, "##0.0") & " " &
                                FormatSpec(.Os.Bpl, "#0.0") & " " & FormatSpec(.Os.Ins, "#0.0") & " " &
                                FormatSpec(.Os.Ca, "#0.0") & " " & FormatSpec(.Os.Fe, "#0.00") & " " &
                                FormatSpec(.Os.Al, "#0.00") & " " & FormatSpec(.Os.Mg, "#0.00") & "   " &
                                FormatSpec(.Os100.Tpa, "##,##0") & " " & FormatSpec(.Os100.WtPct, "##0.0") & " " &
                                FormatSpec(.Os100.Bpl, "#0.0") & " " & FormatSpec(.Os100.Ins, "#0.0") & " " &
                                FormatSpec(.Os100.Ca, "#0.0") & " " & FormatSpec(.Os100.Fe, "#0.00") & " " &
                                FormatSpec(.Os100.Al, "#0.00") & " " & FormatSpec(.Os100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = "O-size " & FormatSpec(.Os.Tpa, "##,##0") & " " & FormatSpec(.Os.WtPct, "##0.0") & " " &
                                      FormatSpec(.Os.Bpl, "#0.0") & " " & FormatSpec(.Os.Ins, "#0.0") & " " &
                                      FormatSpec(.Os.Ca, "#0.0") & " " & FormatSpec(.Os.Fe, "#0.00") & " " &
                                      FormatSpec(.Os.Al, "#0.00") & " " & FormatSpec(.Os.Mg, "#0.00") & "   " &
                                      FormatSpec(.Os100.Tpa, "##,##0") & " " & FormatSpec(.Os100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Os100.Bpl, "#0.0") & " " & FormatSpec(.Os100.Ins, "#0.0") & " " &
                                      FormatSpec(.Os100.Ca, "#0.0") & " " & FormatSpec(.Os100.Fe, "#0.00") & " " &
                                      FormatSpec(.Os100.Al, "#0.00") & " " & FormatSpec(.Os100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Crs pb " & FormatSpec(.Cpb.Tpa, "##,##0") & " " & FormatSpec(.Cpb.WtPct, "##0.0") & " " &
                                FormatSpec(.Cpb.Bpl, "#0.0") & " " & FormatSpec(.Cpb.Ins, "#0.0") & " " &
                                FormatSpec(.Cpb.Ca, "#0.0") & " " & FormatSpec(.Cpb.Fe, "#0.00") & " " &
                                FormatSpec(.Cpb.Al, "#0.00") & " " & FormatSpec(.Cpb.Mg, "#0.00") & "   " &
                                FormatSpec(.Cpb100.Tpa, "##,##0") & " " & FormatSpec(.Cpb100.WtPct, "##0.0") & " " &
                                FormatSpec(.Cpb100.Bpl, "#0.0") & " " & FormatSpec(.Cpb100.Ins, "#0.0") & " " &
                                FormatSpec(.Cpb100.Ca, "#0.0") & " " & FormatSpec(.Cpb100.Fe, "#0.00") & " " &
                                FormatSpec(.Cpb100.Al, "#0.00") & " " & FormatSpec(.Cpb100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = "Crs pb " & FormatSpec(.Cpb.Tpa, "##,##0") & " " & FormatSpec(.Cpb.WtPct, "##0.0") & " " &
                                      FormatSpec(.Cpb.Bpl, "#0.0") & " " & FormatSpec(.Cpb.Ins, "#0.0") & " " &
                                      FormatSpec(.Cpb.Ca, "#0.0") & " " & FormatSpec(.Cpb.Fe, "#0.00") & " " &
                                      FormatSpec(.Cpb.Al, "#0.00") & " " & FormatSpec(.Cpb.Mg, "#0.00") & "   " &
                                      FormatSpec(.Cpb100.Tpa, "##,##0") & " " & FormatSpec(.Cpb100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Cpb100.Bpl, "#0.0") & " " & FormatSpec(.Cpb100.Ins, "#0.0") & " " &
                                      FormatSpec(.Cpb100.Ca, "#0.0") & " " & FormatSpec(.Cpb100.Fe, "#0.00") & " " &
                                      FormatSpec(.Cpb100.Al, "#0.00") & " " & FormatSpec(.Cpb100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Fne pb " & FormatSpec(.Fpb.Tpa, "##,##0") & " " & FormatSpec(.Fpb.WtPct, "##0.0") & " " &
                                FormatSpec(.Fpb.Bpl, "#0.0") & " " & FormatSpec(.Fpb.Ins, "#0.0") & " " &
                                FormatSpec(.Fpb.Ca, "#0.0") & " " & FormatSpec(.Fpb.Fe, "#0.00") & " " &
                                FormatSpec(.Fpb.Al, "#0.00") & " " & FormatSpec(.Fpb.Mg, "#0.00") & "   " &
                                FormatSpec(.Fpb100.Tpa, "##,##0") & " " & FormatSpec(.Fpb100.WtPct, "##0.0") & " " &
                                FormatSpec(.Fpb100.Bpl, "#0.0") & " " & FormatSpec(.Fpb100.Ins, "#0.0") & " " &
                                FormatSpec(.Fpb100.Ca, "#0.0") & " " & FormatSpec(.Fpb100.Fe, "#0.00") & " " &
                                FormatSpec(.Fpb100.Al, "#0.00") & " " & FormatSpec(.Fpb100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = "Fne pb " & FormatSpec(.Fpb.Tpa, "##,##0") & " " & FormatSpec(.Fpb.WtPct, "##0.0") & " " &
                                      FormatSpec(.Fpb.Bpl, "#0.0") & " " & FormatSpec(.Fpb.Ins, "#0.0") & " " &
                                      FormatSpec(.Fpb.Ca, "#0.0") & " " & FormatSpec(.Fpb.Fe, "#0.00") & " " &
                                      FormatSpec(.Fpb.Al, "#0.00") & " " & FormatSpec(.Fpb.Mg, "#0.00") & "   " &
                                      FormatSpec(.Fpb100.Tpa, "##,##0") & " " & FormatSpec(.Fpb100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Fpb100.Bpl, "#0.0") & " " & FormatSpec(.Fpb100.Ins, "#0.0") & " " &
                                      FormatSpec(.Fpb100.Ca, "#0.0") & " " & FormatSpec(.Fpb100.Fe, "#0.00") & " " &
                                      FormatSpec(.Fpb100.Al, "#0.00") & " " & FormatSpec(.Fpb100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Tot pb " & FormatSpec(.Tpb.Tpa, "##,##0") & " " & FormatSpec(.Tpb.WtPct, "##0.0") & " " &
                                FormatSpec(.Tpb.Bpl, "#0.0") & " " & FormatSpec(.Tpb.Ins, "#0.0") & " " &
                                FormatSpec(.Tpb.Ca, "#0.0") & " " & FormatSpec(.Tpb.Fe, "#0.00") & " " &
                                FormatSpec(.Tpb.Al, "#0.00") & " " & FormatSpec(.Tpb.Mg, "#0.00") & "   " &
                                FormatSpec(.Tpb100.Tpa, "##,##0") & " " & FormatSpec(.Tpb100.WtPct, "##0.0") & " " &
                                FormatSpec(.Tpb100.Bpl, "#0.0") & " " & FormatSpec(.Tpb100.Ins, "#0.0") & " " &
                                FormatSpec(.Tpb100.Ca, "#0.0") & " " & FormatSpec(.Tpb100.Fe, "#0.00") & " " &
                                FormatSpec(.Tpb100.Al, "#0.00") & " " & FormatSpec(.Tpb100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = "Tot pb " & FormatSpec(.Tpb.Tpa, "##,##0") & " " & FormatSpec(.Tpb.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tpb.Bpl, "#0.0") & " " & FormatSpec(.Tpb.Ins, "#0.0") & " " &
                                      FormatSpec(.Tpb.Ca, "#0.0") & " " & FormatSpec(.Tpb.Fe, "#0.00") & " " &
                                      FormatSpec(.Tpb.Al, "#0.00") & " " & FormatSpec(.Tpb.Mg, "#0.00") & "   " &
                                      FormatSpec(.Tpb100.Tpa, "##,##0") & " " & FormatSpec(.Tpb100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tpb100.Bpl, "#0.0") & " " & FormatSpec(.Tpb100.Ins, "#0.0") & " " &
                                      FormatSpec(.Tpb100.Ca, "#0.0") & " " & FormatSpec(.Tpb100.Fe, "#0.00") & " " &
                                      FormatSpec(.Tpb100.Al, "#0.00") & " " & FormatSpec(.Tpb100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Crs cn " & FormatSpec(.Ccn.Tpa, "##,##0") & " " & FormatSpec(.Ccn.WtPct, "##0.0") & " " &
                                FormatSpec(.Ccn.Bpl, "#0.0") & " " & FormatSpec(.Ccn.Ins, "#0.0") & " " &
                                FormatSpec(.Ccn.Ca, "#0.0") & " " & FormatSpec(.Ccn.Fe, "#0.00") & " " &
                                FormatSpec(.Ccn.Al, "#0.00") & " " & FormatSpec(.Ccn.Mg, "#0.00") & "   " &
                                FormatSpec(.Ccn100.Tpa, "##,##0") & " " & FormatSpec(.Ccn100.WtPct, "##0.0") & " " &
                                FormatSpec(.Ccn100.Bpl, "#0.0") & " " & FormatSpec(.Ccn100.Ins, "#0.0") & " " &
                                FormatSpec(.Ccn100.Ca, "#0.0") & " " & FormatSpec(.Ccn100.Fe, "#0.00") & " " &
                                FormatSpec(.Ccn100.Al, "#0.00") & " " & FormatSpec(.Ccn100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = "Crs cn " & FormatSpec(.Ccn.Tpa, "##,##0") & " " & FormatSpec(.Ccn.WtPct, "##0.0") & " " &
                                      FormatSpec(.Ccn.Bpl, "#0.0") & " " & FormatSpec(.Ccn.Ins, "#0.0") & " " &
                                      FormatSpec(.Ccn.Ca, "#0.0") & " " & FormatSpec(.Ccn.Fe, "#0.00") & " " &
                                      FormatSpec(.Ccn.Al, "#0.00") & " " & FormatSpec(.Ccn.Mg, "#0.00") & "   " &
                                      FormatSpec(.Ccn100.Tpa, "##,##0") & " " & FormatSpec(.Ccn100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Ccn100.Bpl, "#0.0") & " " & FormatSpec(.Ccn100.Ins, "#0.0") & " " &
                                      FormatSpec(.Ccn100.Ca, "#0.0") & " " & FormatSpec(.Ccn100.Fe, "#0.00") & " " &
                                      FormatSpec(.Ccn100.Al, "#0.00") & " " & FormatSpec(.Ccn100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Fne cn " & FormatSpec(.Fcn.Tpa, "##,##0") & " " & FormatSpec(.Fcn.WtPct, "##0.0") & " " &
                                FormatSpec(.Fcn.Bpl, "#0.0") & " " & FormatSpec(.Fcn.Ins, "#0.0") & " " &
                                FormatSpec(.Fcn.Ca, "#0.0") & " " & FormatSpec(.Fcn.Fe, "#0.00") & " " &
                                FormatSpec(.Fcn.Al, "#0.00") & " " & FormatSpec(.Fcn.Mg, "#0.00") & "   " &
                                FormatSpec(.Fcn100.Tpa, "##,##0") & " " & FormatSpec(.Fcn100.WtPct, "##0.0") & " " &
                                FormatSpec(.Fcn100.Bpl, "#0.0") & " " & FormatSpec(.Fcn100.Ins, "#0.0") & " " &
                                FormatSpec(.Fcn100.Ca, "#0.0") & " " & FormatSpec(.Fcn100.Fe, "#0.00") & " " &
                                FormatSpec(.Fcn100.Al, "#0.00") & " " & FormatSpec(.Fcn100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = "Fne cn " & FormatSpec(.Fcn.Tpa, "##,##0") & " " & FormatSpec(.Fcn.WtPct, "##0.0") & " " &
                                      FormatSpec(.Fcn.Bpl, "#0.0") & " " & FormatSpec(.Fcn.Ins, "#0.0") & " " &
                                      FormatSpec(.Fcn.Ca, "#0.0") & " " & FormatSpec(.Fcn.Fe, "#0.00") & " " &
                                      FormatSpec(.Fcn.Al, "#0.00") & " " & FormatSpec(.Fcn.Mg, "#0.00") & "   " &
                                      FormatSpec(.Fcn100.Tpa, "##,##0") & " " & FormatSpec(.Fcn100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Fcn100.Bpl, "#0.0") & " " & FormatSpec(.Fcn100.Ins, "#0.0") & " " &
                                      FormatSpec(.Fcn100.Ca, "#0.0") & " " & FormatSpec(.Fcn100.Fe, "#0.00") & " " &
                                      FormatSpec(.Fcn100.Al, "#0.00") & " " & FormatSpec(.Fcn100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Tot cn " & FormatSpec(.Tcn.Tpa, "##,##0") & " " & FormatSpec(.Tcn.WtPct, "##0.0") & " " &
                                FormatSpec(.Tcn.Bpl, "#0.0") & " " & FormatSpec(.Tcn.Ins, "#0.0") & " " &
                                FormatSpec(.Tcn.Ca, "#0.0") & " " & FormatSpec(.Tcn.Fe, "#0.00") & " " &
                                FormatSpec(.Tcn.Al, "#0.00") & " " & FormatSpec(.Tcn.Mg, "#0.00") & "   " &
                                FormatSpec(.Tcn100.Tpa, "##,##0") & " " & FormatSpec(.Tcn100.WtPct, "##0.0") & " " &
                                FormatSpec(.Tcn100.Bpl, "#0.0") & " " & FormatSpec(.Tcn100.Ins, "#0.0") & " " &
                                FormatSpec(.Tcn100.Ca, "#0.0") & " " & FormatSpec(.Tcn100.Fe, "#0.00") & " " &
                                FormatSpec(.Tcn100.Al, "#0.00") & " " & FormatSpec(.Tcn100.Mg, "#0.00") & vbCrLf
            Else
                TextStr = "Tot cn " & FormatSpec(.Tcn.Tpa, "##,##0") & " " & FormatSpec(.Tcn.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tcn.Bpl, "#0.0") & " " & FormatSpec(.Tcn.Ins, "#0.0") & " " &
                                      FormatSpec(.Tcn.Ca, "#0.0") & " " & FormatSpec(.Tcn.Fe, "#0.00") & " " &
                                      FormatSpec(.Tcn.Al, "#0.00") & " " & FormatSpec(.Tcn.Mg, "#0.00") & "   " &
                                      FormatSpec(.Tcn100.Tpa, "##,##0") & " " & FormatSpec(.Tcn100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tcn100.Bpl, "#0.0") & " " & FormatSpec(.Tcn100.Ins, "#0.0") & " " &
                                      FormatSpec(.Tcn100.Ca, "#0.0") & " " & FormatSpec(.Tcn100.Fe, "#0.00") & " " &
                                      FormatSpec(.Tcn100.Al, "#0.00") & " " & FormatSpec(.Tcn100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Tot pr " & FormatSpec(.Tpr.Tpa, "##,##0") & " " & FormatSpec(.Tpr.WtPct, "##0.0") & " " &
                                FormatSpec(.Tpr.Bpl, "#0.0") & " " & FormatSpec(.Tpr.Ins, "#0.0") & " " &
                                FormatSpec(.Tpr.Ca, "#0.0") & " " & FormatSpec(.Tpr.Fe, "#0.00") & " " &
                                FormatSpec(.Tpr.Al, "#0.00") & " " & FormatSpec(.Tpr.Mg, "#0.00") & "   " &
                                FormatSpec(.Tpr100.Tpa, "##,##0") & " " & FormatSpec(.Tpr100.WtPct, "##0.0") & " " &
                                FormatSpec(.Tpr100.Bpl, "#0.0") & " " & FormatSpec(.Tpr100.Ins, "#0.0") & " " &
                                FormatSpec(.Tpr100.Ca, "#0.0") & " " & FormatSpec(.Tpr100.Fe, "#0.00") & " " &
                                FormatSpec(.Tpr100.Al, "#0.00") & " " & FormatSpec(.Tpr100.Mg, "#0.00") & vbCrLf

                rtbRept1.Text = rtbRept1.Text & vbCrLf
            Else
                TextStr = "Tot pr " & FormatSpec(.Tpr.Tpa, "##,##0") & " " & FormatSpec(.Tpr.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tpr.Bpl, "#0.0") & " " & FormatSpec(.Tpr.Ins, "#0.0") & " " &
                                      FormatSpec(.Tpr.Ca, "#0.0") & " " & FormatSpec(.Tpr.Fe, "#0.00") & " " &
                                      FormatSpec(.Tpr.Al, "#0.00") & " " & FormatSpec(.Tpr.Mg, "#0.00") & "   " &
                                      FormatSpec(.Tpr100.Tpa, "##,##0") & " " & FormatSpec(.Tpr100.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tpr100.Bpl, "#0.0") & " " & FormatSpec(.Tpr100.Ins, "#0.0") & " " &
                                      FormatSpec(.Tpr100.Ca, "#0.0") & " " & FormatSpec(.Tpr100.Fe, "#0.00") & " " &
                                      FormatSpec(.Tpr100.Al, "#0.00") & " " & FormatSpec(.Tpr100.Mg, "#0.00")
                gWriteLine(aFileNumber, TextStr)
                gWriteLine(aFileNumber, " ")
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Tails  " & FormatSpec(.Ttl.Tpa, "##,##0") & " " & FormatSpec(.Ttl.WtPct, "##0.0") & " " &
                                FormatSpec(.Ttl.Bpl, "#0.0") &
                                "    MtxX OnSpec = " & FormatSpec(.MtxxOnSpecPcHole, "##0.00") &
                                "       " &
                                FormatSpec(.Ttl100.Tpa, "##,##0") & " " &
                                FormatSpec(.Ttl100.WtPct, "##0.0") & " " & FormatSpec(.Ttl.Bpl, "#0.0") &
                                "    MtxX OnSpec = " & FormatSpec(.MtxxOnSpec100Hole, "##0.00") &
                                vbCrLf
            Else
                TextStr = "Tails  " & FormatSpec(.Ttl.Tpa, "##,##0") & " " & FormatSpec(.Ttl.WtPct, "##0.0") & " " &
                                      FormatSpec(.Ttl.Bpl, "#0.0") &
                                      "    MtxX OnSpec = " & FormatSpec(.MtxxOnSpecPcHole, "##0.00") &
                                      "       " &
                                      FormatSpec(.Ttl100.Tpa, "##,##0") & " " &
                                      FormatSpec(.Ttl100.WtPct, "##0.0") & " " & FormatSpec(.Ttl.Bpl, "#0.0") &
                                      "    MtxX OnSpec = " & FormatSpec(.MtxxOnSpec100Hole, "##0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Wst cl " & FormatSpec(.Wcl.Tpa, "##,##0") & " " & FormatSpec(.Wcl.WtPct, "##0.0") & " " &
                                FormatSpec(.Wcl.Bpl, "#0.0") &
                                "    TotX OnSpec = " & FormatSpec(.TotxOnSpecPcHole, "##0.00") &
                                "       " &
                                FormatSpec(.Wcl100.Tpa, "##,##0") & " " &
                                FormatSpec(.Wcl100.WtPct, "##0.0") & " " & FormatSpec(.Wcl.Bpl, "#0.0") &
                                "    TotX OnSpec = " & FormatSpec(.TotxAll100Hole, "##0.00") &
                                vbCrLf
            Else
                TextStr = "Wst cl " & FormatSpec(.Wcl.Tpa, "##,##0") & " " & FormatSpec(.Wcl.WtPct, "##0.0") & " " &
                                      FormatSpec(.Wcl.Bpl, "#0.0") &
                                      "    TotX OnSpec = " & FormatSpec(.TotxOnSpecPcHole, "##0.00") &
                                      "       " &
                                      FormatSpec(.Wcl100.Tpa, "##,##0") & " " &
                                      FormatSpec(.Wcl100.WtPct, "##0.0") & " " & FormatSpec(.Wcl.Bpl, "#0.0") &
                                      "    TotX OnSpec = " & FormatSpec(.TotxAll100Hole, "##0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Crs fd " & FormatSpec(.Cfd.Tpa, "##,##0") & " " & FormatSpec(.Cfd.WtPct, "##0.0") & " " &
                                FormatSpec(.Cfd.Bpl, "#0.0") &
                                "    MtxX All    = " & FormatSpec(.MtxxAllPcHole, "##0.00") &
                                "       " &
                                FormatSpec(.Cfd100.Tpa, "##,##0") & " " &
                                FormatSpec(.Cfd100.WtPct, "##0.0") & " " & FormatSpec(.Cfd.Bpl, "#0.0") &
                                "    MtxX All    = " & FormatSpec(.MtxxAll100Hole, "##0.00") &
                                vbCrLf
            Else
                TextStr = "Crs fd " & FormatSpec(.Cfd.Tpa, "##,##0") & " " & FormatSpec(.Cfd.WtPct, "##0.0") & " " &
                                      FormatSpec(.Cfd.Bpl, "#0.0") &
                                      "    MtxX All    = " & FormatSpec(.MtxxAllPcHole, "##0.00") &
                                      "       " &
                                      FormatSpec(.Cfd100.Tpa, "##,##0") & " " &
                                      FormatSpec(.Cfd100.WtPct, "##0.0") & " " & FormatSpec(.Cfd.Bpl, "#0.0") &
                                      "    MtxX All    = " & FormatSpec(.MtxxAll100Hole, "##0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Fne fd " & FormatSpec(.Ffd.Tpa, "##,##0") & " " & FormatSpec(.Ffd.WtPct, "##0.0") & " " &
                                FormatSpec(.Ffd.Bpl, "#0.0") &
                                "    TotX All    = " & FormatSpec(.TotxAllPcHole, "##0.00") &
                                "       " &
                                FormatSpec(.Ffd100.Tpa, "##,##0") & " " &
                                FormatSpec(.Ffd100.WtPct, "##0.0") & " " & FormatSpec(.Ffd.Bpl, "#0.0") &
                                "    TotX All    = " & FormatSpec(.TotxAll100Hole, "##0.00") &
                                vbCrLf
            Else
                TextStr = "Fne fd " & FormatSpec(.Ffd.Tpa, "##,##0") & " " & FormatSpec(.Ffd.WtPct, "##0.0") & " " &
                                      FormatSpec(.Ffd.Bpl, "#0.0") &
                                      "    TotX All    = " & FormatSpec(.TotxAllPcHole, "##0.00") &
                                      "       " &
                                      FormatSpec(.Ffd100.Tpa, "##,##0") & " " &
                                      FormatSpec(.Ffd100.WtPct, "##0.0") & " " & FormatSpec(.Ffd.Bpl, "#0.0") &
                                      "    TotX All    = " & FormatSpec(.TotxAll100Hole, "##0.00")
                gWriteLine(aFileNumber, TextStr)
            End If

            If aRptType = "RTB" Then
                rtbRept1.Text = rtbRept1.Text &
                    "Tot fd " & FormatSpec(.Tfd.Tpa, "##,##0") & " " & FormatSpec(.Tfd.WtPct, "##0.0") & " " &
                                FormatSpec(.Tfd.Bpl, "#0.0") &
                                "                               " &
                                FormatSpec(.Tfd100.Tpa, "##,##0") & " " &
                                FormatSpec(.Tfd100.WtPct, "##0.0") & " " & FormatSpec(.Tfd.Bpl, "#0.0") & vbCrLf

                rtbRept1.Text = rtbRept1.Text & vbCrLf
            Else
                TextStr = "Tot fd " & FormatSpec(.Tfd.Tpa, "##,##0") & " " & FormatSpec(.Tfd.WtPct, "##0.0") & " " &
                                      FormatSpec(.Tfd.Bpl, "#0.0") &
                                      "                               " &
                                      FormatSpec(.Tfd100.Tpa, "##,##0") & " " &
                                      FormatSpec(.Tfd100.WtPct, "##0.0") & " " & FormatSpec(.Tfd.Bpl, "#0.0")
                gWriteLine(aFileNumber, TextStr)
                gWriteLine(aFileNumber, " ")
            End If
        End With
    End Sub

    Private Function FormatSpec(ByVal aNum As Single,
                                ByVal aFormStr As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        FormatSpec = gPadLeft(Format(aNum, aFormStr), Len(aFormStr))
    End Function

    Private Sub cmdClrOverride_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClrOverride.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        ClearOverride()
    End Sub

    Private Sub chkUseRawProspAsOverride_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseRawProspAsOverride.CheckedChanged

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        ClearOverride()
    End Sub

    Private Sub ClearOverride()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        ssSplitOverride.MaxRows = 0

        fProcessing = True
        txtSplitOverrideName.Text = ""
        cboSplitOverrideMineName.Text = "None"
        fProcessing = False

        lblGen55.Text = ""
    End Sub

    Private Sub cmdPrintRptAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrintRptAll.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim a As Object
        Dim fs As Object
        Dim ThisStr As String
        Dim ErrMsg As String

        'printer.Font.Name = "Monospac821 BT"

        'Set fs = CreateObject("Scripting.FileSystemObject")

        'If Not fs.FileExists(txtRptAllToTxtFile.Text) Then
        '   ErrMsg = "Text file = " & txtRptAllToTxtFile.Text & " does not exist!"
        '   MsgBox ErrMsg, vbOKOnly, "Print Status"

        '        SetActionStatus("")
        '        Me.Cursor = Cursors.Arrow
        '   Exit Sub
        'End If

        '    SetActionStatus("Printing report...")
        '    Me.Cursor = Cursors.WaitCursor

        ''Text file exists -- continue with printing
        'Set a = fs.OpenTextFile(txtRptAllToTxtFile.Text, 1)

        'Do While Not a.atendofstream
        '    ThisStr = a.ReadLine

        '    Printer.Print ThisStr
        'Loop

        'Printer.EndDoc
        'Set fs = Nothing
        'Set a = Nothing

        '    SetActionStatus("")
        '    Me.Cursor = Cursors.Arrow
    End Sub

    'Private Sub dtpAreaBeginDrillDate_Change(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpAreaBeginDrillDate.ValueChanged

    '    '**********************************************************************
    '    '
    '    '
    '    '
    '    '**********************************************************************

    '    ssSplitReview.MaxRows = 0
    'End Sub

    Private Sub dtpAreaBeginDrillDate2_EditValueChanged(sender As Object, e As EventArgs)
        Try
            ssSplitReview.MaxRows = 0
        Catch ex As Exception
            If TypeOf ex Is InvalidActiveXStateException Then
                ssSplitReview.CreateControl()
            End If
        End Try
    End Sub


    'Private Sub dtpAreaBeginDrillDate_Click()

    '    '**********************************************************************
    '    '
    '    '
    '    '
    '    '**********************************************************************

    '    ssSplitReview.MaxRows = 0
    'End Sub

    Private Sub dtpAreaBeginDrillDate2_Click(sender As Object, e As EventArgs)
        ssSplitReview.MaxRows = 0
    End Sub

    'Private Sub dtpAreaEndDrillDate_Change(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpAreaEndDrillDate.ValueChanged

    '    '**********************************************************************
    '    '
    '    '
    '    '
    '    '**********************************************************************

    '    ssSplitReview.MaxRows = 0
    'End Sub

    Private Sub dtpAreaEndDrillDate2_EditValueChanged(sender As Object, e As EventArgs)
        Try
            ssSplitReview.MaxRows = 0
        Catch ex As Exception
            If TypeOf ex Is InvalidActiveXStateException Then
                ssSplitReview.CreateControl()
            End If
        End Try
    End Sub


    'Private Sub dtpAreaEndDrillDate_Click()

    '    '**********************************************************************
    '    '
    '    '
    '    '
    '    '**********************************************************************

    '    ssSplitReview.MaxRows = 0
    'End Sub

    Private Sub dtpAreaBEndDrillDate2_Click(sender As Object, e As EventArgs)
        ssSplitReview.MaxRows = 0
    End Sub

    Private Sub cboAreaOwnership_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        ssSplitReview.MaxRows = 0
    End Sub

    Private Sub cboAreaProspHoleType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        ssSplitReview.MaxRows = 0
    End Sub

    Private Sub cboAreaMinedOutStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        ssSplitReview.MaxRows = 0
        gHaveRawProspData = False
    End Sub

    Private Sub cboAreaDefnMineName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        gHaveRawProspData = False
    End Sub

    Private Sub cboAreaDefnSpecAreaName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        gHaveRawProspData = False
    End Sub

    Private Sub chkCreateOutputOnly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCreateOutputOnly.CheckedChanged

        If chkCreateOutputOnly.Checked = True Then
            cmdSaveProspectDataset.Enabled = False
            'ssSplitReview.Visible = False
            tbcSplitResults.Visible = False
            tbcHoleResults.Visible = False
            lblNoReview.Visible = True
            cmdCopyToOverrides.Visible = False
            lblBarrenSplComm.Visible = False
            lblGen25.Visible = False
        Else
            cmdSaveProspectDataset.Enabled = True
            'ssSplitReview.Visible = True
            tbcSplitResults.Visible = True
            tbcHoleResults.Visible = True
            lblNoReview.Visible = False
            cmdCopyToOverrides.Visible = True
            lblBarrenSplComm.Visible = True
            lblGen25.Visible = True
        End If
    End Sub

    Private Sub txtSplitOverrideName_Click()
        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        gHaveRawProspData = False
        If fProcessing = False Then
            chkUseRawProspAsOverride.Checked = False
        End If
    End Sub

    Private Sub cboSplitOverrideMineName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSplitOverrideMineName.SelectedIndexChanged

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        If fProcessing = False Then
            chkUseRawProspAsOverride.Checked = False
        End If
    End Sub


    Private Sub cmdAreaReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAreaReport.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo cmdAreaReportClickError

        Dim ConnectString As String
        Dim ProspData As gRawProspSplRdctnType
        Dim RcvryData As gDataRdctnParamsType
        Dim RcvryProdQual(0 To 14) As gDataRdctnProdQualType
        Dim PctProspect100 As Boolean
        Dim InclCpbAlways As String
        Dim InclFpbAlways As String
        Dim InclOsAlways As String
        Dim InclCpbNever As String
        Dim InclFpbNever As String
        Dim InclOsNever As String
        Dim MtxDensityCalc As Single
        Dim MineHasOffSpecPbPlt As String
        Dim OffSpecPbPlt As Boolean
        Dim SumData As gRawProspSplRdctnSumType
        Dim HoleCount As Long
        Dim MinableHoleCount As Long
        Dim CanSelectRejectTpb As String

        SetActionStatus("Generating area reserve report...")
        Me.Cursor = Cursors.WaitCursor

        'Need to sum hole data from ssCompReview
        SumTheHoleData(SumData,
                       optProdCoeff.Checked,
                       ProspData,
                       HoleCount,
                       MinableHoleCount)

        PctProspect100 = False
        If opt100Pct.Checked = True Then
            PctProspect100 = True
        Else
            PctProspect100 = False
        End If

        GetRcvryEtcParamsFromForm(RcvryData)

        InclCpbAlways = "No"
        InclFpbAlways = "No"
        InclOsAlways = "No"
        InclCpbNever = "No"
        InclFpbNever = "No"
        InclOsNever = "No"
        CanSelectRejectTpb = "No"
        If RcvryData.InclCpbAlways Then
            InclCpbAlways = "Yes"
        End If
        If RcvryData.InclFpbAlways Then
            InclFpbAlways = "Yes"
        End If
        If RcvryData.InclOsAlways Then
            InclOsAlways = "Yes"
        End If
        If RcvryData.InclCpbNever Then
            InclCpbNever = "Yes"
        End If
        If RcvryData.InclFpbNever Then
            InclFpbNever = "Yes"
        End If
        If RcvryData.InclOsNever Then
            InclOsNever = "Yes"
        End If
        If RcvryData.CanSelectRejectTpb Then
            CanSelectRejectTpb = "Yes"
        End If

        If RcvryData.MineHasOffSpecPbPlt = True Then
            MineHasOffSpecPbPlt = "Mine has off-spec pebble processing plant."
        Else
            MineHasOffSpecPbPlt = ""
        End If

        'With ProspData
        '    rptProspRdctn.Formulas(13) = "OvbThk = '" & Format(.OvbThk, "##0.0") & "'"
        '    rptProspRdctn.Formulas(14) = "MtxThk = '" & Format(.MtxThk, "##0.0") & "'"
        '    rptProspRdctn.Formulas(15) = "ItbThk = '" & Format(.ItbThk, "##0.0") & "'"

        '    rptProspRdctn.Formulas(18) = "MtxXOnSpec = '" & Format(.MtxxOnSpecPcHole, "##0.00") & "'"
        '    rptProspRdctn.Formulas(135) = "MtxxAll = '" & Format(.MtxxAllPcHole, "##0.00") & "'"
        '    rptProspRdctn.Formulas(19) = "TotXOnSpec = '" & Format(.TotxOnSpecPcHole, "##0.00") & "'"
        '    rptProspRdctn.Formulas(136) = "TotxAll = '" & Format(.TotxAllPcHole, "##0.00") & "'"

        '    rptProspRdctn.Formulas(20) = "CpbWtPct = '" & Format(.Cpb.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(21) = "CpbTpa = '" & Format(.Cpb.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(22) = "CpbBpl = '" & Format(.Cpb.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(23) = "CpbIns = '" & Format(.Cpb.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(24) = "CpbIa = '" & Format(.Cpb.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(25) = "CpbFe = '" & Format(.Cpb.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(26) = "CpbAl = '" & Format(.Cpb.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(27) = "CpbMg = '" & Format(.Cpb.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(28) = "CpbCa = '" & Format(.Cpb.Ca, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(29) = "FpbWtPct = '" & Format(.Fpb.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(30) = "FpbTpa = '" & Format(.Fpb.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(31) = "FpbBpl = '" & Format(.Fpb.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(32) = "FpbIns = '" & Format(.Fpb.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(33) = "FpbIa = '" & Format(.Fpb.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(34) = "FpbFe = '" & Format(.Fpb.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(35) = "FpbAl = '" & Format(.Fpb.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(36) = "FpbMg = '" & Format(.Fpb.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(37) = "FpbCa = '" & Format(.Fpb.Ca, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(38) = "TpbWtPct = '" & Format(.Tpb.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(39) = "TpbTpa = '" & Format(.Tpb.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(40) = "TpbBpl = '" & Format(.Tpb.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(41) = "TpbIns = '" & Format(.Tpb.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(42) = "TpbIa = '" & Format(.Tpb.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(43) = "TpbFe = '" & Format(.Tpb.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(44) = "TpbAl = '" & Format(.Tpb.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(45) = "TpbMg = '" & Format(.Tpb.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(46) = "TpbCa = '" & Format(.Tpb.Ca, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(47) = "TcnWtPct = '" & Format(.Tcn.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(48) = "TcnTpa = '" & Format(.Tcn.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(49) = "TcnBpl = '" & Format(.Tcn.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(50) = "TcnIns = '" & Format(.Tcn.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(51) = "TcnIa = '" & Format(.Tcn.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(52) = "TcnFe = '" & Format(.Tcn.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(53) = "TcnAl = '" & Format(.Tcn.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(54) = "TcnMg = '" & Format(.Tcn.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(55) = "TcnCa = '" & Format(.Tcn.Ca, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(56) = "TprWtPct = '" & Format(.Tpr.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(57) = "TprTpa = '" & Format(.Tpr.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(58) = "TprBpl = '" & Format(.Tpr.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(59) = "TprIns = '" & Format(.Tpr.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(60) = "TprIa = '" & Format(.Tpr.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(61) = "TprFe = '" & Format(.Tpr.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(62) = "TprAl = '" & Format(.Tpr.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(63) = "TprMg = '" & Format(.Tpr.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(64) = "TprCa = '" & Format(.Tpr.Ca, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(65) = "TtlWtPct = '" & Format(.Ttl.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(66) = "TtlTpa = '" & Format(.Ttl.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(67) = "TtlBpl = '" & Format(.Ttl.Bpl, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(68) = "CfdWtPct = '" & Format(.Cfd.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(69) = "CfdTpa = '" & Format(.Cfd.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(70) = "CfdBpl = '" & Format(.Cfd.Bpl, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(71) = "FfdWtPct = '" & Format(.Ffd.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(72) = "FfdTpa = '" & Format(.Ffd.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(73) = "FfdBpl = '" & Format(.Ffd.Bpl, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(74) = "TfdWtPct = '" & Format(.Tfd.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(75) = "TfdTpa = '" & Format(.Tfd.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(76) = "TfdBpl = '" & Format(.Tfd.Bpl, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(77) = "WclWtPct = '" & Format(.Wcl.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(78) = "WclTpa = '" & Format(.Wcl.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(79) = "WclBpl = '" & Format(.Wcl.Bpl, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(80) = "MineName = '" & cboAreaDefnMineName.Text & "'"
        '    '-----
        '    rptProspRdctn.Formulas(81) = "FcnWtPct = '" & Format(.Fcn.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(82) = "FcnTpa = '" & Format(.Fcn.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(83) = "FcnBpl = '" & Format(.Fcn.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(84) = "FcnIns = '" & Format(.Fcn.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(85) = "FcnIa = '" & Format(.Fcn.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(86) = "FcnFe = '" & Format(.Fcn.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(87) = "FcnAl = '" & Format(.Fcn.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(88) = "FcnMg = '" & Format(.Fcn.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(89) = "FcnCa = '" & Format(.Fcn.Ca, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(90) = "CcnWtPct = '" & Format(.Ccn.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(91) = "CcnTpa = '" & Format(.Ccn.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(92) = "CcnBpl = '" & Format(.Ccn.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(93) = "CcnIns = '" & Format(.Ccn.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(94) = "CcnIa = '" & Format(.Ccn.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(95) = "CcnFe = '" & Format(.Ccn.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(96) = "CcnAl = '" & Format(.Ccn.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(97) = "CcnMg = '" & Format(.Ccn.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(98) = "CcnCa = '" & Format(.Ccn.Ca, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(99) = "IpWtPct = '" & Format(.Ip.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(100) = "IpTpa = '" & Format(.Ip.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(101) = "IpBpl = '" & Format(.Ip.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(102) = "IpIns = '" & Format(.Ip.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(103) = "IpIa = '" & Format(.Ip.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(104) = "IpFe = '" & Format(.Ip.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(105) = "IpAl = '" & Format(.Ip.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(106) = "IpMg = '" & Format(.Ip.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(107) = "IpCa = '" & Format(.Ip.Ca, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(108) = "CpbTonRcvry = '" & Format(RcvryData.PbTonRcvryCrs, "##0.0") & "'"
        '    rptProspRdctn.Formulas(109) = "FpbTonRcvry = '" & Format(RcvryData.PbTonRcvryFne, "##0.0") & "'"
        '    rptProspRdctn.Formulas(110) = "IpTonRcvry = '" & Format(RcvryData.IpTonRcvryTot, "##0.0") & "'"
        '    rptProspRdctn.Formulas(111) = "CfdTonRcvry = '" & Format(RcvryData.FdTonRcvryCrs, "##0.0") & "'"
        '    rptProspRdctn.Formulas(112) = "FfdTonRcvry = '" & Format(RcvryData.FdTonRcvryFne, "##0.0") & "'"
        '    rptProspRdctn.Formulas(113) = "CfdBplTonRcvry = '" & Format(RcvryData.FdBplRcvryCrs, "##0.0") & "'"
        '    rptProspRdctn.Formulas(114) = "FfdBplTonRcvry = '" & Format(RcvryData.FdBplRcvryFne, "##0.0") & "'"
        '    rptProspRdctn.Formulas(115) = "WclTonRcvry = '" & Format(RcvryData.ClTonRcvryTot, "##0.0") & "'"

        '    rptProspRdctn.Formulas(116) = "OsWtPct = '" & Format(.Os.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(117) = "OsTpa = '" & Format(.Os.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(118) = "OsBpl = '" & Format(.Os.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(119) = "OsIns = '" & Format(.Os.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(120) = "OsIa = '" & Format(.Os.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(121) = "OsFe = '" & Format(.Os.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(122) = "OsAl = '" & Format(.Os.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(123) = "OsMg = '" & Format(.Os.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(124) = "OsCa = '" & Format(.Os.Ca, "##0.0") & "'"

        '    'These are special items.
        '    'lblReviewComm --> lblGen41)
        '        If lblGen41.Text = "Split Data" Then
        '            rptProspRdctn.Formulas(125) = "OsMin = '" & .OsMin & "'"
        '            rptProspRdctn.Formulas(126) = "CpbMin = '" & .CpbMin & "'"
        '            rptProspRdctn.Formulas(127) = "FpbMin = '" & .FpbMin & "'"
        '            rptProspRdctn.Formulas(128) = "TpbMin = '" & .TpbMin & "'"
        '            rptProspRdctn.Formulas(129) = "IpMin = '" & .IpMin & "'"
        '            rptProspRdctn.Formulas(130) = "CcnMin = '" & .CcnMin & "'"
        '            rptProspRdctn.Formulas(131) = "FcnMin = '" & .FcnMin & "'"
        '            rptProspRdctn.Formulas(132) = "TcnMin = '" & .TcnMin & "'"
        '            rptProspRdctn.Formulas(133) = "MineableCalcd = '" & .MineableCalcd & "'"
        '            rptProspRdctn.Formulas(134) = "MineableOride = '" & .MineableOride & "'"
        '        Else    'Hole data
        '            rptProspRdctn.Formulas(125) = "OsMin = '" & .OsMinHole & "'"
        '            rptProspRdctn.Formulas(126) = "CpbMin = '" & .CpbMinHole & "'"
        '            rptProspRdctn.Formulas(127) = "FpbMin = '" & .FpbMinHole & "'"
        '            rptProspRdctn.Formulas(128) = "TpbMin = '" & .TpbMinHole & "'"
        '            rptProspRdctn.Formulas(129) = "IpMin = '" & .IpMinHole & "'"
        '            rptProspRdctn.Formulas(130) = "CcnMin = '" & .CcnMinHole & "'"
        '            rptProspRdctn.Formulas(131) = "FcnMin = '" & .FcnMinHole & "'"
        '            rptProspRdctn.Formulas(132) = "TcnMin = '" & .TcnMinHole & "'"
        '            rptProspRdctn.Formulas(133) = "MineableCalcd = '" & .MineableHole & "'"
        '            rptProspRdctn.Formulas(134) = "MineableOride = '" & " " & "'"   'Does not apply to holes
        '        End If

        '    rptProspRdctn.Formulas(137) = "InclCpbAlways = '" & InclCpbAlways & "'"
        '    rptProspRdctn.Formulas(138) = "InclFpbAlways = '" & InclFpbAlways & "'"
        '    rptProspRdctn.Formulas(139) = "InclCpbNever = '" & InclCpbNever & "'"
        '    rptProspRdctn.Formulas(140) = "InclFpbNever = '" & InclFpbNever & "'"
        '    '-----
        '    rptProspRdctn.Formulas(141) = "InclOsAlways = '" & InclOsAlways & "'"
        '    rptProspRdctn.Formulas(142) = "InclOsNever = '" & InclOsNever & "'"
        '    '-----
        '    'Need to "recalculate" a matrix density.
        '    'lblReviewComm --> lblGen41)
        '        If lblGen41.Text = "Split Data" Then
        '            If .SplitThck <> 0 Then
        '                MtxDensityCalc = Round((.MtxTpaPc * 2000) / (.SplitThck * 43560), 1)
        '            Else
        '                MtxDensityCalc = 0
        '            End If
        '        Else    'Hole data
        '            If .MtxThk <> 0 Then
        '                MtxDensityCalc = Round((.MtxTpaPc * 2000) / (.MtxThk * 43560), 1)
        '            Else
        '                MtxDensityCalc = 0
        '            End If
        '        End If
        '    rptProspRdctn.Formulas(143) = "MtxDensityCalc = '" & Format(MtxDensityCalc, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(144) = "MtxPctSol = '" & Format(.MtxPctSol, "##0.0") & "'"

        '    rptProspRdctn.Formulas(145) = "MgPltInpWtPct = '" & Format(.MgPltInp.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(146) = "MgPltInpTpa = '" & Format(.MgPltInp.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(147) = "MgPltInpBpl = '" & Format(.MgPltInp.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(148) = "MgPltInpIns = '" & Format(.MgPltInp.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(149) = "MgPltInpIa = '" & Format(.MgPltInp.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(150) = "MgPltInpFe = '" & Format(.MgPltInp.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(151) = "MgPltInpAl = '" & Format(.MgPltInp.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(152) = "MgPltInpMg = '" & Format(.MgPltInp.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(153) = "MgPltInpCa = '" & Format(.MgPltInp.Ca, "##0.0") & "'"

        '    rptProspRdctn.Formulas(154) = "MgPltRejWtPct = '" & Format(.MgPltRej.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(155) = "MgPltRejTpa = '" & Format(.MgPltRej.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(156) = "MgPltRejBpl = '" & Format(.MgPltRej.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(157) = "MgPltRejIns = '" & Format(.MgPltRej.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(158) = "MgPltRejIa = '" & Format(.MgPltRej.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(159) = "MgPltRejFe = '" & Format(.MgPltRej.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(160) = "MgPltRejAl = '" & Format(.MgPltRej.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(161) = "MgPltRejMg = '" & Format(.MgPltRej.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(162) = "MgPltRejCa = '" & Format(.MgPltRej.Ca, "##0.0") & "'"

        '    rptProspRdctn.Formulas(163) = "MgPltProdWtPct = '" & Format(.MgPltProd.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(164) = "MgPltProdTpa = '" & Format(.MgPltProd.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(165) = "MgPltProdBpl = '" & Format(.MgPltProd.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(166) = "MgPltProdIns = '" & Format(.MgPltProd.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(167) = "MgPltProdIa = '" & Format(.MgPltProd.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(168) = "MgPltProdFe = '" & Format(.MgPltProd.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(169) = "MgPltProdAl = '" & Format(.MgPltProd.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(170) = "MgPltProdMg = '" & Format(.MgPltProd.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(171) = "MgPltProdCa = '" & Format(.MgPltProd.Ca, "##0.0") & "'"

        '    rptProspRdctn.Formulas(172) = "MgPltTcnWtPct = '" & Format(.MgPltTcn.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(173) = "MgPltTcnTpa = '" & Format(.MgPltTcn.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(174) = "MgPltTcnBpl = '" & Format(.MgPltTcn.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(175) = "MgPltTcnIns = '" & Format(.MgPltTcn.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(176) = "MgPltTcnIa = '" & Format(.MgPltTcn.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(177) = "MgPltTcnFe = '" & Format(.MgPltTcn.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(178) = "MgPltTcnAl = '" & Format(.MgPltTcn.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(179) = "MgPltTcnMg = '" & Format(.MgPltTcn.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(180) = "MgPltTcnCa = '" & Format(.MgPltTcn.Ca, "##0.0") & "'"

        '    rptProspRdctn.Formulas(181) = "MgPltTprWtPct = '" & Format(.MgPltTpr.WtPct, "##0.00") & "'"
        '    rptProspRdctn.Formulas(182) = "MgPltTprTpa = '" & Format(.MgPltTpr.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(183) = "MgPltTprBpl = '" & Format(.MgPltTpr.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(184) = "MgPltTprIns = '" & Format(.MgPltTpr.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(185) = "MgPltTprIa = '" & Format(.MgPltTpr.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(186) = "MgPltTprFe = '" & Format(.MgPltTpr.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(187) = "MgPltTprAl = '" & Format(.MgPltTpr.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(188) = "MgPltTprMg = '" & Format(.MgPltTpr.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(189) = "MgPltTprCa = '" & Format(.MgPltTpr.Ca, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(190) = "MineHasOffSpecPbPlt = '" & MineHasOffSpecPbPlt & "'"
        '    '-----
        '    rptProspRdctn.Formulas(191) = "TpbRejTpa = '" & Format(.TpbRej.Tpa, "##,##0") & "'"
        '    rptProspRdctn.Formulas(192) = "TpbRejBpl = '" & Format(.TpbRej.Bpl, "##0.0") & "'"
        '    rptProspRdctn.Formulas(193) = "TpbRejIns = '" & Format(.TpbRej.Ins, "##0.0") & "'"
        '    rptProspRdctn.Formulas(194) = "TpbRejIa = '" & Format(.TpbRej.Ia, "##0.00") & "'"
        '    rptProspRdctn.Formulas(195) = "TpbRejFe = '" & Format(.TpbRej.Fe, "##0.00") & "'"
        '    rptProspRdctn.Formulas(196) = "TpbRejAl = '" & Format(.TpbRej.Al, "##0.00") & "'"
        '    rptProspRdctn.Formulas(197) = "TpbRejMg = '" & Format(.TpbRej.Mg, "##0.00") & "'"
        '    rptProspRdctn.Formulas(198) = "TpbRejCa = '" & Format(.TpbRej.Ca, "##0.0") & "'"
        '    '-----
        '    rptProspRdctn.Formulas(199) = "HoleCount = '" & Format(HoleCount, "###,##0") & "'"
        '    rptProspRdctn.Formulas(200) = "MinableHoleCount = '" & Format(MinableHoleCount, "###,##0") & "'"
        '    rptProspRdctn.Formulas(201) = "CanSelectRejectTpb = '" & CanSelectRejectTpb & "'"
        'End With

        If MineHasOffSpecPbPlt <> "" Then
            OffSpecPbPlt = True
        Else
            OffSpecPbPlt = False
        End If

        'Have all the needed data -- start the report
        'rptProspRdctn.ReportFileName = gPath + "\Reports\" + _
        '                               "ProspectReductionArea.rpt"

        ''Connect to Oracle database
        'ConnectString = "DSN = " + gDataSource + ";UID = " + gOracleUserName + _
        '    ";PWD = " + gOracleUserPassword + ";DSQ = "

        'rptProspRdctn.Connect = ConnectString

        'Need to pass the company name and report type into the report
        'lblReviewComm --> lblGen41)
        'rptProspRdctn.ParameterFields(0) = "pCompanyName;" & gCompanyName & ";TRUE"
        '    rptProspRdctn.ParameterFields(1) = "pRptType;" & lblGen41).Text & ";TRUE"
        'rptProspRdctn.ParameterFields(2) = "pPctProspect100;" & PctProspect100 & ";TRUE"
        'rptProspRdctn.ParameterFields(3) = "pMineHasOffSpecPbPlt;" & OffSpecPbPlt & ";TRUE"
        'rptProspRdctn.ParameterFields(4) = "pProdSizeDesig;" & txtPsizeDefnName.Text & ";TRUE"
        'rptProspRdctn.ParameterFields(5) = "pRcvryEtcScen;" & txtRcvryEtcName.Text & ";TRUE"
        'rptProspRdctn.ParameterFields(6) = "pAreaDefn;" & txtAreaDefnName.Text & ";TRUE"

        ''Report window maximized
        'rptProspRdctn.WindowState = crptMaximized

        'rptProspRdctn.WindowTitle = "Raw Data Reduction Prospect Area Data"

        'User not allowed to minimize report window
        'rptProspRdctn.WindowMinButton = False

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow

        ''Start Crystal Reports
        'rptProspRdctn.action = 1

        'rptProspRdctn.Reset

        Exit Sub

cmdAreaReportClickError:
        MsgBox("Error printing prospect data from raw data reduction." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Print Error")

        On Error Resume Next
        'rptProspRdctn.Reset
    End Sub

    Private Sub SumTheHoleData(ByRef aSumData As gRawProspSplRdctnSumType,
                               ByVal aProdCoeffVal As Integer,
                               ByRef aProspData As gRawProspSplRdctnType,
                               ByRef aHoleCount As Long,
                               ByRef aMinableHoleCount As Long)

        On Error GoTo SumTheHoleDataError

        Dim RowIdx As Long
        Dim ProspData As gRawProspSplRdctnType
        Dim InclCpb As Boolean
        Dim InclFpb As Boolean
        Dim InclOs As Boolean
        Dim TotWt As Double

        Dim RcvryParamsData As gDataRdctnParamsType
        Dim RcvryProdQual(0 To 14) As gDataRdctnProdQualType

        aHoleCount = 0
        aMinableHoleCount = 0

        GetRcvryEtcParamsFromForm(RcvryParamsData)
        ZeroRdctnSumData(aSumData)

        'Only want the holes that are minable!  Only want to sum the products
        'that are minable from the holes that are minable!

        With aSumData
            For RowIdx = 1 To ssCompReview.MaxRows
                ProspData = gGetDataFromReviewSprd(ssCompReview, RowIdx)

                aHoleCount = aHoleCount + 1

                If ProspData.MineableHole = "M" Or ProspData.MineableHole = "MF" Then
                    aMinableHoleCount = aMinableHoleCount + 1
                    'Coarse pebble
                    If RcvryParamsData.InclCpbAlways Then
                        'Include the Cpb no matter what.
                        InclCpb = True
                    Else
                        If RcvryParamsData.InclCpbNever Then
                            'Exclude the Cpb no matter what.
                            InclCpb = False
                        Else
                            'Include the Cpb if the quality is OK.
                            'Need to use .CpbMinHole here
                            If ProspData.CpbOnSpec = "Yes" Or ProspData.CpbOnSpec = "ND" Then
                                InclCpb = True
                            End If
                        End If
                    End If

                    'Temp Fix!!!
                    If RcvryParamsData.MineHasOffSpecPbPlt Then
                        InclCpb = True
                        InclFpb = True
                    End If

                    'Fine pebble
                    If RcvryParamsData.InclFpbAlways Then
                        'Include the Fpb no matter what.
                        InclFpb = True
                    Else
                        If RcvryParamsData.InclFpbNever Then
                            'Exclude the Fpb no matter what.
                            InclFpb = False
                        Else
                            'Include the Fpb if the quality is OK.
                            'Need to use .FpbMinHole here
                            If ProspData.FpbOnSpec = "Yes" Or ProspData.FpbOnSpec = "ND" Then
                                InclFpb = True
                            End If
                        End If
                    End If

                    'Oversize
                    If RcvryParamsData.InclOsAlways = True Then
                        'Include the Os no matter what.
                        InclOs = True
                    Else
                        If RcvryParamsData.InclOsNever = True Then
                            'Exclude the Os no matter what.
                            InclOs = False
                        Else
                            'Include the Os if the quality is OK.
                            'Need to use .OsMinHole here
                            If ProspData.OsOnSpec = "Yes" Or ProspData.OsOnSpec = "ND" Then
                                InclOs = True
                            End If
                        End If
                    End If

                    .OvbThk = .OvbThk + ProspData.OvbThk
                    .ItbThk = .ItbThk + ProspData.ItbThk
                    .MtxThk = .MtxThk + ProspData.MtxThk

                    If aProdCoeffVal = False Then
                        '100% Prospect -- Add to aSumData.
                        If InclOs = True Then
                            .Os.Tpa = .Os.Tpa + ProspData.Os100.Tpa
                            .Os.BplTons = .Os.BplTons + ProspData.Os100.Tpa * ProspData.Os100.Bpl
                            .Os.InsTons = .Os.InsTons + ProspData.Os100.Tpa * ProspData.Os100.Ins
                            .Os.IaTons = .Os.IaTons + ProspData.Os100.Tpa * ProspData.Os100.Ia
                            .Os.FeTons = .Os.FeTons + ProspData.Os100.Tpa * ProspData.Os100.Fe
                            .Os.AlTons = .Os.AlTons + ProspData.Os100.Tpa * ProspData.Os100.Al
                            .Os.MgTons = .Os.MgTons + ProspData.Os100.Tpa * ProspData.Os100.Mg
                            .Os.CaTons = .Os.CaTons + ProspData.Os100.Tpa * ProspData.Os100.Ca
                            .Os.TwBpl = .Os.TwBpl + IIf(ProspData.Os100.Bpl > 0, ProspData.Os100.Tpa, 0)
                            .Os.TwIns = .Os.TwIns + IIf(ProspData.Os100.Ins > 0, ProspData.Os100.Tpa, 0)
                            .Os.TwIa = .Os.TwIa + IIf(ProspData.Os100.Ia > 0, ProspData.Os100.Tpa, 0)
                            .Os.TwFe = .Os.TwFe + IIf(ProspData.Os100.Fe > 0, ProspData.Os100.Tpa, 0)
                            .Os.TwAl = .Os.TwAl + IIf(ProspData.Os100.Al > 0, ProspData.Os100.Tpa, 0)
                            .Os.TwMg = .Os.TwMg + IIf(ProspData.Os100.Mg > 0, ProspData.Os100.Tpa, 0)
                            .Os.TwCa = .Os.TwCa + IIf(ProspData.Os100.Ca > 0, ProspData.Os100.Tpa, 0)
                        End If

                        If InclCpb = True Then
                            .Cpb.Tpa = .Cpb.Tpa + ProspData.Cpb100.Tpa
                            .Cpb.BplTons = .Cpb.BplTons + ProspData.Cpb100.Tpa * ProspData.Cpb100.Bpl
                            .Cpb.InsTons = .Cpb.InsTons + ProspData.Cpb100.Tpa * ProspData.Cpb100.Ins
                            .Cpb.IaTons = .Cpb.IaTons + ProspData.Cpb100.Tpa * ProspData.Cpb100.Ia
                            .Cpb.FeTons = .Cpb.FeTons + ProspData.Cpb100.Tpa * ProspData.Cpb100.Fe
                            .Cpb.AlTons = .Cpb.AlTons + ProspData.Cpb100.Tpa * ProspData.Cpb100.Al
                            .Cpb.MgTons = .Cpb.MgTons + ProspData.Cpb100.Tpa * ProspData.Cpb100.Mg
                            .Cpb.CaTons = .Cpb.CaTons + ProspData.Cpb100.Tpa * ProspData.Cpb100.Ca
                            .Cpb.TwBpl = .Cpb.TwBpl + IIf(ProspData.Cpb100.Bpl > 0, ProspData.Cpb100.Tpa, 0)
                            .Cpb.TwIns = .Cpb.TwIns + IIf(ProspData.Cpb100.Ins > 0, ProspData.Cpb100.Tpa, 0)
                            .Cpb.TwIa = .Cpb.TwIa + IIf(ProspData.Cpb100.Ia > 0, ProspData.Cpb100.Tpa, 0)
                            .Cpb.TwFe = .Cpb.TwFe + IIf(ProspData.Cpb100.Fe > 0, ProspData.Cpb100.Tpa, 0)
                            .Cpb.TwAl = .Cpb.TwAl + IIf(ProspData.Cpb100.Al > 0, ProspData.Cpb100.Tpa, 0)
                            .Cpb.TwMg = .Cpb.TwMg + IIf(ProspData.Cpb100.Mg > 0, ProspData.Cpb100.Tpa, 0)
                            .Cpb.TwCa = .Cpb.TwCa + IIf(ProspData.Cpb100.Ca > 0, ProspData.Cpb100.Tpa, 0)
                        End If

                        If InclFpb = True Then
                            .Fpb.Tpa = .Fpb.Tpa + ProspData.Fpb100.Tpa
                            .Fpb.BplTons = .Fpb.BplTons + ProspData.Fpb100.Tpa * ProspData.Fpb100.Bpl
                            .Fpb.InsTons = .Fpb.InsTons + ProspData.Fpb100.Tpa * ProspData.Fpb100.Ins
                            .Fpb.IaTons = .Fpb.IaTons + ProspData.Fpb100.Tpa * ProspData.Fpb100.Ia
                            .Fpb.FeTons = .Fpb.FeTons + ProspData.Fpb100.Tpa * ProspData.Fpb100.Fe
                            .Fpb.AlTons = .Fpb.AlTons + ProspData.Fpb100.Tpa * ProspData.Fpb100.Al
                            .Fpb.MgTons = .Fpb.MgTons + ProspData.Fpb100.Tpa * ProspData.Fpb100.Mg
                            .Fpb.CaTons = .Fpb.CaTons + ProspData.Fpb100.Tpa * ProspData.Fpb100.Ca
                            .Fpb.TwBpl = .Fpb.TwBpl + IIf(ProspData.Fpb100.Bpl > 0, ProspData.Fpb100.Tpa, 0)
                            .Fpb.TwIns = .Fpb.TwIns + IIf(ProspData.Fpb100.Ins > 0, ProspData.Fpb100.Tpa, 0)
                            .Fpb.TwIa = .Fpb.TwIa + IIf(ProspData.Fpb100.Ia > 0, ProspData.Fpb100.Tpa, 0)
                            .Fpb.TwFe = .Fpb.TwFe + IIf(ProspData.Fpb100.Fe > 0, ProspData.Fpb100.Tpa, 0)
                            .Fpb.TwAl = .Fpb.TwAl + IIf(ProspData.Fpb100.Al > 0, ProspData.Fpb100.Tpa, 0)
                            .Fpb.TwMg = .Fpb.TwMg + IIf(ProspData.Fpb100.Mg > 0, ProspData.Fpb100.Tpa, 0)
                            .Fpb.TwCa = .Fpb.TwCa + IIf(ProspData.Fpb100.Ca > 0, ProspData.Fpb100.Tpa, 0)
                        End If

                        .Tpb.Tpa = .Tpb.Tpa + ProspData.Tpb100.Tpa
                        .Tpb.BplTons = .Tpb.BplTons + ProspData.Tpb100.Tpa * ProspData.Tpb100.Bpl
                        .Tpb.InsTons = .Tpb.InsTons + ProspData.Tpb100.Tpa * ProspData.Tpb100.Ins
                        .Tpb.IaTons = .Tpb.IaTons + ProspData.Tpb100.Tpa * ProspData.Tpb100.Ia
                        .Tpb.FeTons = .Tpb.FeTons + ProspData.Tpb100.Tpa * ProspData.Tpb100.Fe
                        .Tpb.AlTons = .Tpb.AlTons + ProspData.Tpb100.Tpa * ProspData.Tpb100.Al
                        .Tpb.MgTons = .Tpb.MgTons + ProspData.Tpb100.Tpa * ProspData.Tpb100.Mg
                        .Tpb.CaTons = .Tpb.CaTons + ProspData.Tpb100.Tpa * ProspData.Tpb100.Ca
                        .Tpb.TwBpl = .Tpb.TwBpl + IIf(ProspData.Tpb100.Bpl > 0, ProspData.Tpb100.Tpa, 0)
                        .Tpb.TwIns = .Tpb.TwIns + IIf(ProspData.Tpb100.Ins > 0, ProspData.Tpb100.Tpa, 0)
                        .Tpb.TwIa = .Tpb.TwIa + IIf(ProspData.Tpb100.Ia > 0, ProspData.Tpb100.Tpa, 0)
                        .Tpb.TwFe = .Tpb.TwFe + IIf(ProspData.Tpb100.Fe > 0, ProspData.Tpb100.Tpa, 0)
                        .Tpb.TwAl = .Tpb.TwAl + IIf(ProspData.Tpb100.Al > 0, ProspData.Tpb100.Tpa, 0)
                        .Tpb.TwMg = .Tpb.TwMg + IIf(ProspData.Tpb100.Mg > 0, ProspData.Tpb100.Tpa, 0)
                        .Tpb.TwCa = .Tpb.TwCa + IIf(ProspData.Tpb100.Ca > 0, ProspData.Tpb100.Tpa, 0)

                        .TpbRej.Tpa = .TpbRej.Tpa + ProspData.TpbRej100.Tpa
                        .TpbRej.BplTons = .TpbRej.BplTons + ProspData.TpbRej100.Tpa * ProspData.TpbRej100.Bpl
                        .TpbRej.InsTons = .TpbRej.InsTons + ProspData.TpbRej100.Tpa * ProspData.TpbRej100.Ins
                        .TpbRej.IaTons = .TpbRej.IaTons + ProspData.TpbRej100.Tpa * ProspData.TpbRej100.Ia
                        .TpbRej.FeTons = .TpbRej.FeTons + ProspData.TpbRej100.Tpa * ProspData.TpbRej100.Fe
                        .TpbRej.AlTons = .TpbRej.AlTons + ProspData.TpbRej100.Tpa * ProspData.TpbRej100.Al
                        .TpbRej.MgTons = .TpbRej.MgTons + ProspData.TpbRej100.Tpa * ProspData.TpbRej100.Mg
                        .TpbRej.CaTons = .TpbRej.CaTons + ProspData.TpbRej100.Tpa * ProspData.TpbRej100.Ca
                        .TpbRej.TwBpl = .TpbRej.TwBpl + IIf(ProspData.TpbRej100.Bpl > 0, ProspData.TpbRej100.Tpa, 0)
                        .TpbRej.TwIns = .TpbRej.TwIns + IIf(ProspData.TpbRej100.Ins > 0, ProspData.TpbRej100.Tpa, 0)
                        .TpbRej.TwIa = .TpbRej.TwIa + IIf(ProspData.TpbRej100.Ia > 0, ProspData.TpbRej100.Tpa, 0)
                        .TpbRej.TwFe = .TpbRej.TwFe + IIf(ProspData.TpbRej100.Fe > 0, ProspData.TpbRej100.Tpa, 0)
                        .TpbRej.TwAl = .TpbRej.TwAl + IIf(ProspData.TpbRej100.Al > 0, ProspData.TpbRej100.Tpa, 0)
                        .TpbRej.TwMg = .TpbRej.TwMg + IIf(ProspData.TpbRej100.Mg > 0, ProspData.TpbRej100.Tpa, 0)
                        .TpbRej.TwCa = .TpbRej.TwCa + IIf(ProspData.TpbRej100.Ca > 0, ProspData.TpbRej100.Tpa, 0)

                        'IP, Fcn, Ccn always included.
                        .Ip.Tpa = .Ip.Tpa + ProspData.Ip100.Tpa
                        .Ip.BplTons = .Ip.BplTons + ProspData.Ip100.Tpa * ProspData.Ip100.Bpl
                        .Ip.InsTons = .Ip.InsTons + ProspData.Ip100.Tpa * ProspData.Ip100.Ins
                        .Ip.IaTons = .Ip.IaTons + ProspData.Ip100.Tpa * ProspData.Ip100.Ia
                        .Ip.FeTons = .Ip.FeTons + ProspData.Ip100.Tpa * ProspData.Ip100.Fe
                        .Ip.AlTons = .Ip.AlTons + ProspData.Ip100.Tpa * ProspData.Ip100.Al
                        .Ip.MgTons = .Ip.MgTons + ProspData.Ip100.Tpa * ProspData.Ip100.Mg
                        .Ip.CaTons = .Ip.CaTons + ProspData.Ip100.Tpa * ProspData.Ip100.Ca
                        .Ip.TwBpl = .Ip.TwBpl + IIf(ProspData.Ip100.Bpl > 0, ProspData.Ip100.Tpa, 0)
                        .Ip.TwIns = .Ip.TwIns + IIf(ProspData.Ip100.Ins > 0, ProspData.Ip100.Tpa, 0)
                        .Ip.TwIa = .Ip.TwIa + IIf(ProspData.Ip100.Ia > 0, ProspData.Ip100.Tpa, 0)
                        .Ip.TwFe = .Ip.TwFe + IIf(ProspData.Ip100.Fe > 0, ProspData.Ip100.Tpa, 0)
                        .Ip.TwAl = .Ip.TwAl + IIf(ProspData.Ip100.Al > 0, ProspData.Ip100.Tpa, 0)
                        .Ip.TwMg = .Ip.TwMg + IIf(ProspData.Ip100.Mg > 0, ProspData.Ip100.Tpa, 0)
                        .Ip.TwCa = .Ip.TwCa + IIf(ProspData.Ip100.Ca > 0, ProspData.Ip100.Tpa, 0)

                        .Ccn.Tpa = .Ccn.Tpa + ProspData.Ccn100.Tpa
                        .Ccn.BplTons = .Ccn.BplTons + ProspData.Ccn100.Tpa * ProspData.Ccn100.Bpl
                        .Ccn.InsTons = .Ccn.InsTons + ProspData.Ccn100.Tpa * ProspData.Ccn100.Ins
                        .Ccn.IaTons = .Ccn.IaTons + ProspData.Ccn100.Tpa * ProspData.Ccn100.Ia
                        .Ccn.FeTons = .Ccn.FeTons + ProspData.Ccn100.Tpa * ProspData.Ccn100.Fe
                        .Ccn.AlTons = .Ccn.AlTons + ProspData.Ccn100.Tpa * ProspData.Ccn100.Al
                        .Ccn.MgTons = .Ccn.MgTons + ProspData.Ccn100.Tpa * ProspData.Ccn100.Mg
                        .Ccn.CaTons = .Ccn.CaTons + ProspData.Ccn100.Tpa * ProspData.Ccn100.Ca
                        .Ccn.TwBpl = .Ccn.TwBpl + IIf(ProspData.Ccn100.Bpl > 0, ProspData.Ccn100.Tpa, 0)
                        .Ccn.TwIns = .Ccn.TwIns + IIf(ProspData.Ccn100.Ins > 0, ProspData.Ccn100.Tpa, 0)
                        .Ccn.TwIa = .Ccn.TwIa + IIf(ProspData.Ccn100.Ia > 0, ProspData.Ccn100.Tpa, 0)
                        .Ccn.TwFe = .Ccn.TwFe + IIf(ProspData.Ccn100.Fe > 0, ProspData.Ccn100.Tpa, 0)
                        .Ccn.TwAl = .Ccn.TwAl + IIf(ProspData.Ccn100.Al > 0, ProspData.Ccn100.Tpa, 0)
                        .Ccn.TwMg = .Ccn.TwMg + IIf(ProspData.Ccn100.Mg > 0, ProspData.Ccn100.Tpa, 0)
                        .Ccn.TwCa = .Ccn.TwCa + IIf(ProspData.Ccn100.Ca > 0, ProspData.Ccn100.Tpa, 0)

                        .Fcn.Tpa = .Fcn.Tpa + ProspData.Fcn100.Tpa
                        .Fcn.BplTons = .Fcn.BplTons + ProspData.Fcn100.Tpa * ProspData.Fcn100.Bpl
                        .Fcn.InsTons = .Fcn.InsTons + ProspData.Fcn100.Tpa * ProspData.Fcn100.Ins
                        .Fcn.IaTons = .Fcn.IaTons + ProspData.Fcn100.Tpa * ProspData.Fcn100.Ia
                        .Fcn.FeTons = .Fcn.FeTons + ProspData.Fcn100.Tpa * ProspData.Fcn100.Fe
                        .Fcn.AlTons = .Fcn.AlTons + ProspData.Fcn100.Tpa * ProspData.Fcn100.Al
                        .Fcn.MgTons = .Fcn.MgTons + ProspData.Fcn100.Tpa * ProspData.Fcn100.Mg
                        .Fcn.CaTons = .Fcn.CaTons + ProspData.Fcn100.Tpa * ProspData.Fcn100.Ca
                        .Fcn.TwBpl = .Fcn.TwBpl + IIf(ProspData.Fcn100.Bpl > 0, ProspData.Fcn100.Tpa, 0)
                        .Fcn.TwIns = .Fcn.TwIns + IIf(ProspData.Fcn100.Ins > 0, ProspData.Fcn100.Tpa, 0)
                        .Fcn.TwIa = .Fcn.TwIa + IIf(ProspData.Fcn100.Ia > 0, ProspData.Fcn100.Tpa, 0)
                        .Fcn.TwFe = .Fcn.TwFe + IIf(ProspData.Fcn100.Fe > 0, ProspData.Fcn100.Tpa, 0)
                        .Fcn.TwAl = .Fcn.TwAl + IIf(ProspData.Fcn100.Al > 0, ProspData.Fcn100.Tpa, 0)
                        .Fcn.TwMg = .Fcn.TwMg + IIf(ProspData.Fcn100.Mg > 0, ProspData.Fcn100.Tpa, 0)
                        .Fcn.TwCa = .Fcn.TwCa + IIf(ProspData.Fcn100.Ca > 0, ProspData.Fcn100.Tpa, 0)

                        .Tcn.Tpa = .Tcn.Tpa + ProspData.Tcn100.Tpa
                        .Tcn.BplTons = .Tcn.BplTons + ProspData.Tcn100.Tpa * ProspData.Tcn100.Bpl
                        .Tcn.InsTons = .Tcn.InsTons + ProspData.Tcn100.Tpa * ProspData.Tcn100.Ins
                        .Tcn.IaTons = .Tcn.IaTons + ProspData.Tcn100.Tpa * ProspData.Tcn100.Ia
                        .Tcn.FeTons = .Tcn.FeTons + ProspData.Tcn100.Tpa * ProspData.Tcn100.Fe
                        .Tcn.AlTons = .Tcn.AlTons + ProspData.Tcn100.Tpa * ProspData.Tcn100.Al
                        .Tcn.MgTons = .Tcn.MgTons + ProspData.Tcn100.Tpa * ProspData.Tcn100.Mg
                        .Tcn.CaTons = .Tcn.CaTons + ProspData.Tcn100.Tpa * ProspData.Tcn100.Ca
                        .Tcn.TwBpl = .Tcn.TwBpl + IIf(ProspData.Tcn100.Bpl > 0, ProspData.Tcn100.Tpa, 0)
                        .Tcn.TwIns = .Tcn.TwIns + IIf(ProspData.Tcn100.Ins > 0, ProspData.Tcn100.Tpa, 0)
                        .Tcn.TwIa = .Tcn.TwIa + IIf(ProspData.Tcn100.Ia > 0, ProspData.Tcn100.Tpa, 0)
                        .Tcn.TwFe = .Tcn.TwFe + IIf(ProspData.Tcn100.Fe > 0, ProspData.Tcn100.Tpa, 0)
                        .Tcn.TwAl = .Tcn.TwAl + IIf(ProspData.Tcn100.Al > 0, ProspData.Tcn100.Tpa, 0)
                        .Tcn.TwMg = .Tcn.TwMg + IIf(ProspData.Tcn100.Mg > 0, ProspData.Tcn100.Tpa, 0)
                        .Tcn.TwCa = .Tcn.TwCa + IIf(ProspData.Tcn100.Ca > 0, ProspData.Tcn100.Tpa, 0)

                        .Wcl.Tpa = .Wcl.Tpa + ProspData.Wcl100.Tpa
                        .Wcl.BplTons = .Wcl.BplTons + ProspData.Wcl100.Tpa * ProspData.Wcl100.Bpl
                        .Wcl.TwBpl = .Wcl.TwBpl + IIf(ProspData.Wcl100.Bpl > 0, ProspData.Wcl100.Tpa, 0)

                        .Cfd.Tpa = .Cfd.Tpa + ProspData.Cfd100.Tpa
                        .Cfd.BplTons = .Cfd.BplTons + ProspData.Cfd100.Tpa * ProspData.Cfd100.Bpl
                        .Cfd.TwBpl = .Cfd.TwBpl + IIf(ProspData.Cfd100.Bpl > 0, ProspData.Cfd100.Tpa, 0)

                        .Ffd.Tpa = .Ffd.Tpa + ProspData.Ffd100.Tpa
                        .Ffd.BplTons = .Ffd.BplTons + ProspData.Ffd100.Tpa * ProspData.Ffd100.Bpl
                        .Ffd.TwBpl = .Ffd.TwBpl + IIf(ProspData.Ffd100.Bpl > 0, ProspData.Ffd100.Tpa, 0)

                        .Tfd.Tpa = .Tfd.Tpa + ProspData.Tfd100.Tpa
                        .Tfd.BplTons = .Tfd.BplTons + ProspData.Tfd100.Tpa * ProspData.Tfd100.Bpl
                        .Tfd.TwBpl = .Tfd.TwBpl + IIf(ProspData.Tfd100.Bpl > 0, ProspData.Tfd100.Tpa, 0)

                        .Ttl.Tpa = .Ttl.Tpa + ProspData.Ttl100.Tpa
                        .Ttl.BplTons = .Ttl.BplTons + ProspData.Ttl100.Tpa * ProspData.Ttl100.Bpl
                        .Ttl.TwBpl = .Ttl.TwBpl + IIf(ProspData.Ttl100.Bpl > 0, ProspData.Ttl100.Tpa, 0)

                        .Tpr.Tpa = .Tpr.Tpa + ProspData.Tpr100.Tpa
                        .Tpr.BplTons = .Tpr.BplTons + ProspData.Tpr100.Tpa * ProspData.Tpr100.Bpl
                        .Tpr.InsTons = .Tpr.InsTons + ProspData.Tpr100.Tpa * ProspData.Tpr100.Ins
                        .Tpr.IaTons = .Tpr.IaTons + ProspData.Tpr100.Tpa * ProspData.Tpr100.Ia
                        .Tpr.FeTons = .Tpr.FeTons + ProspData.Tpr100.Tpa * ProspData.Tpr100.Fe
                        .Tpr.AlTons = .Tpr.AlTons + ProspData.Tpr100.Tpa * ProspData.Tpr100.Al
                        .Tpr.MgTons = .Tpr.MgTons + ProspData.Tpr100.Tpa * ProspData.Tpr100.Mg
                        .Tpr.CaTons = .Tpr.CaTons + ProspData.Tpr100.Tpa * ProspData.Tpr100.Ca
                        .Tpr.TwBpl = .Tpr.TwBpl + IIf(ProspData.Tpr100.Bpl > 0, ProspData.Tpr100.Tpa, 0)
                        .Tpr.TwIns = .Tpr.TwIns + IIf(ProspData.Tpr100.Ins > 0, ProspData.Tpr100.Tpa, 0)
                        .Tpr.TwIa = .Tpr.TwIa + IIf(ProspData.Tpr100.Ia > 0, ProspData.Tpr100.Tpa, 0)
                        .Tpr.TwFe = .Tpr.TwFe + IIf(ProspData.Tpr100.Fe > 0, ProspData.Tpr100.Tpa, 0)
                        .Tpr.TwAl = .Tpr.TwAl + IIf(ProspData.Tpr100.Al > 0, ProspData.Tpr100.Tpa, 0)
                        .Tpr.TwMg = .Tpr.TwMg + IIf(ProspData.Tpr100.Mg > 0, ProspData.Tpr100.Tpa, 0)
                        .Tpr.TwCa = .Tpr.TwCa + IIf(ProspData.Tpr100.Ca > 0, ProspData.Tpr100.Tpa, 0)

                        .MgPltInp.Tpa = .MgPltInp.Tpa + ProspData.MgPltInp100.Tpa
                        .MgPltInp.BplTons = .MgPltInp.BplTons + ProspData.MgPltInp100.Tpa * ProspData.MgPltInp100.Bpl
                        .MgPltInp.InsTons = .MgPltInp.InsTons + ProspData.MgPltInp100.Tpa * ProspData.MgPltInp100.Ins
                        .MgPltInp.IaTons = .MgPltInp.IaTons + ProspData.MgPltInp100.Tpa * ProspData.MgPltInp100.Ia
                        .MgPltInp.FeTons = .MgPltInp.FeTons + ProspData.MgPltInp100.Tpa * ProspData.MgPltInp100.Fe
                        .MgPltInp.AlTons = .MgPltInp.AlTons + ProspData.MgPltInp100.Tpa * ProspData.MgPltInp100.Al
                        .MgPltInp.MgTons = .MgPltInp.MgTons + ProspData.MgPltInp100.Tpa * ProspData.MgPltInp100.Mg
                        .MgPltInp.CaTons = .MgPltInp.CaTons + ProspData.MgPltInp100.Tpa * ProspData.MgPltInp100.Ca
                        .MgPltInp.TwBpl = .MgPltInp.TwBpl + IIf(ProspData.MgPltInp100.Bpl > 0, ProspData.MgPltInp100.Tpa, 0)
                        .MgPltInp.TwIns = .MgPltInp.TwIns + IIf(ProspData.MgPltInp100.Ins > 0, ProspData.MgPltInp100.Tpa, 0)
                        .MgPltInp.TwIa = .MgPltInp.TwIa + IIf(ProspData.MgPltInp100.Ia > 0, ProspData.MgPltInp100.Tpa, 0)
                        .MgPltInp.TwFe = .MgPltInp.TwFe + IIf(ProspData.MgPltInp100.Fe > 0, ProspData.MgPltInp100.Tpa, 0)
                        .MgPltInp.TwAl = .MgPltInp.TwAl + IIf(ProspData.MgPltInp100.Al > 0, ProspData.MgPltInp100.Tpa, 0)
                        .MgPltInp.TwMg = .MgPltInp.TwMg + IIf(ProspData.MgPltInp100.Mg > 0, ProspData.MgPltInp100.Tpa, 0)
                        .MgPltInp.TwCa = .MgPltInp.TwCa + IIf(ProspData.MgPltInp100.Ca > 0, ProspData.MgPltInp100.Tpa, 0)

                        .MgPltRej.Tpa = .MgPltRej.Tpa + ProspData.MgPltRej100.Tpa
                        .MgPltRej.BplTons = .MgPltRej.BplTons + ProspData.MgPltRej100.Tpa * ProspData.MgPltRej100.Bpl
                        .MgPltRej.InsTons = .MgPltRej.InsTons + ProspData.MgPltRej100.Tpa * ProspData.MgPltRej100.Ins
                        .MgPltRej.IaTons = .MgPltRej.IaTons + ProspData.MgPltRej100.Tpa * ProspData.MgPltRej100.Ia
                        .MgPltRej.FeTons = .MgPltRej.FeTons + ProspData.MgPltRej100.Tpa * ProspData.MgPltRej100.Fe
                        .MgPltRej.AlTons = .MgPltRej.AlTons + ProspData.MgPltRej100.Tpa * ProspData.MgPltRej100.Al
                        .MgPltRej.MgTons = .MgPltRej.MgTons + ProspData.MgPltRej100.Tpa * ProspData.MgPltRej100.Mg
                        .MgPltRej.CaTons = .MgPltRej.CaTons + ProspData.MgPltRej100.Tpa * ProspData.MgPltRej100.Ca
                        .MgPltRej.TwBpl = .MgPltRej.TwBpl + IIf(ProspData.MgPltRej100.Bpl > 0, ProspData.MgPltRej100.Tpa, 0)
                        .MgPltRej.TwIns = .MgPltRej.TwIns + IIf(ProspData.MgPltRej100.Ins > 0, ProspData.MgPltRej100.Tpa, 0)
                        .MgPltRej.TwIa = .MgPltRej.TwIa + IIf(ProspData.MgPltRej100.Ia > 0, ProspData.MgPltRej100.Tpa, 0)
                        .MgPltRej.TwFe = .MgPltRej.TwFe + IIf(ProspData.MgPltRej100.Fe > 0, ProspData.MgPltRej100.Tpa, 0)
                        .MgPltRej.TwAl = .MgPltRej.TwAl + IIf(ProspData.MgPltRej100.Al > 0, ProspData.MgPltRej100.Tpa, 0)
                        .MgPltRej.TwMg = .MgPltRej.TwMg + IIf(ProspData.MgPltRej100.Mg > 0, ProspData.MgPltRej100.Tpa, 0)
                        .MgPltRej.TwCa = .MgPltRej.TwCa + IIf(ProspData.MgPltRej100.Ca > 0, ProspData.MgPltRej100.Tpa, 0)

                        .MgPltProd.Tpa = .MgPltProd.Tpa + ProspData.MgPltProd100.Tpa
                        .MgPltProd.BplTons = .MgPltProd.BplTons + ProspData.MgPltProd100.Tpa * ProspData.MgPltProd100.Bpl
                        .MgPltProd.InsTons = .MgPltProd.InsTons + ProspData.MgPltProd100.Tpa * ProspData.MgPltProd100.Ins
                        .MgPltProd.IaTons = .MgPltProd.IaTons + ProspData.MgPltProd100.Tpa * ProspData.MgPltProd100.Ia
                        .MgPltProd.FeTons = .MgPltProd.FeTons + ProspData.MgPltProd100.Tpa * ProspData.MgPltProd100.Fe
                        .MgPltProd.AlTons = .MgPltProd.AlTons + ProspData.MgPltProd100.Tpa * ProspData.MgPltProd100.Al
                        .MgPltProd.MgTons = .MgPltProd.MgTons + ProspData.MgPltProd100.Tpa * ProspData.MgPltProd100.Mg
                        .MgPltProd.CaTons = .MgPltProd.CaTons + ProspData.MgPltProd100.Tpa * ProspData.MgPltProd100.Ca
                        .MgPltProd.TwBpl = .MgPltProd.TwBpl + IIf(ProspData.MgPltProd100.Bpl > 0, ProspData.MgPltProd100.Tpa, 0)
                        .MgPltProd.TwIns = .MgPltProd.TwIns + IIf(ProspData.MgPltProd100.Ins > 0, ProspData.MgPltProd100.Tpa, 0)
                        .MgPltProd.TwIa = .MgPltProd.TwIa + IIf(ProspData.MgPltProd100.Ia > 0, ProspData.MgPltProd100.Tpa, 0)
                        .MgPltProd.TwFe = .MgPltProd.TwFe + IIf(ProspData.MgPltProd100.Fe > 0, ProspData.MgPltProd100.Tpa, 0)
                        .MgPltProd.TwAl = .MgPltProd.TwAl + IIf(ProspData.MgPltProd100.Al > 0, ProspData.MgPltProd100.Tpa, 0)
                        .MgPltProd.TwMg = .MgPltProd.TwMg + IIf(ProspData.MgPltProd100.Mg > 0, ProspData.MgPltProd100.Tpa, 0)
                        .MgPltProd.TwCa = .MgPltProd.TwCa + IIf(ProspData.MgPltProd100.Ca > 0, ProspData.MgPltProd100.Tpa, 0)

                        .MgPltTcn.Tpa = .MgPltTcn.Tpa + ProspData.MgPltTcn100.Tpa
                        .MgPltTcn.BplTons = .MgPltTcn.BplTons + ProspData.MgPltTcn100.Tpa * ProspData.MgPltTcn100.Bpl
                        .MgPltTcn.InsTons = .MgPltTcn.InsTons + ProspData.MgPltTcn100.Tpa * ProspData.MgPltTcn100.Ins
                        .MgPltTcn.IaTons = .MgPltTcn.IaTons + ProspData.MgPltTcn100.Tpa * ProspData.MgPltTcn100.Ia
                        .MgPltTcn.FeTons = .MgPltTcn.FeTons + ProspData.MgPltTcn100.Tpa * ProspData.MgPltTcn100.Fe
                        .MgPltTcn.AlTons = .MgPltTcn.AlTons + ProspData.MgPltTcn100.Tpa * ProspData.MgPltTcn100.Al
                        .MgPltTcn.MgTons = .MgPltTcn.MgTons + ProspData.MgPltTcn100.Tpa * ProspData.MgPltTcn100.Mg
                        .MgPltTcn.CaTons = .MgPltTcn.CaTons + ProspData.MgPltTcn100.Tpa * ProspData.MgPltTcn100.Ca
                        .MgPltTcn.TwBpl = .MgPltTcn.TwBpl + IIf(ProspData.MgPltTcn100.Bpl > 0, ProspData.MgPltTcn100.Tpa, 0)
                        .MgPltTcn.TwIns = .MgPltTcn.TwIns + IIf(ProspData.MgPltTcn100.Ins > 0, ProspData.MgPltTcn100.Tpa, 0)
                        .MgPltTcn.TwIa = .MgPltTcn.TwIa + IIf(ProspData.MgPltTcn100.Ia > 0, ProspData.MgPltTcn100.Tpa, 0)
                        .MgPltTcn.TwFe = .MgPltTcn.TwFe + IIf(ProspData.MgPltTcn100.Fe > 0, ProspData.MgPltTcn100.Tpa, 0)
                        .MgPltTcn.TwAl = .MgPltTcn.TwAl + IIf(ProspData.MgPltTcn100.Al > 0, ProspData.MgPltTcn100.Tpa, 0)
                        .MgPltTcn.TwMg = .MgPltTcn.TwMg + IIf(ProspData.MgPltTcn100.Mg > 0, ProspData.MgPltTcn100.Tpa, 0)
                        .MgPltTcn.TwCa = .MgPltTcn.TwCa + IIf(ProspData.MgPltTcn100.Ca > 0, ProspData.MgPltTcn100.Tpa, 0)

                        .MgPltTpr.Tpa = .MgPltTpr.Tpa + ProspData.MgPltTpr100.Tpa
                        .MgPltTpr.BplTons = .MgPltTpr.BplTons + ProspData.MgPltTpr100.Tpa * ProspData.MgPltTpr100.Bpl
                        .MgPltTpr.InsTons = .MgPltTpr.InsTons + ProspData.MgPltTpr100.Tpa * ProspData.MgPltTpr100.Ins
                        .MgPltTpr.IaTons = .MgPltTpr.IaTons + ProspData.MgPltTpr100.Tpa * ProspData.MgPltTpr100.Ia
                        .MgPltTpr.FeTons = .MgPltTpr.FeTons + ProspData.MgPltTpr100.Tpa * ProspData.MgPltTpr100.Fe
                        .MgPltTpr.AlTons = .MgPltTpr.AlTons + ProspData.MgPltTpr100.Tpa * ProspData.MgPltTpr100.Al
                        .MgPltTpr.MgTons = .MgPltTpr.MgTons + ProspData.MgPltTpr100.Tpa * ProspData.MgPltTpr100.Mg
                        .MgPltTpr.CaTons = .MgPltTpr.CaTons + ProspData.MgPltTpr100.Tpa * ProspData.MgPltTpr100.Ca
                        .MgPltTpr.TwBpl = .MgPltTpr.TwBpl + IIf(ProspData.MgPltTpr100.Bpl > 0, ProspData.MgPltTpr100.Tpa, 0)
                        .MgPltTpr.TwIns = .MgPltTpr.TwIns + IIf(ProspData.MgPltTpr100.Ins > 0, ProspData.MgPltTpr100.Tpa, 0)
                        .MgPltTpr.TwIa = .MgPltTpr.TwIa + IIf(ProspData.MgPltTpr100.Ia > 0, ProspData.MgPltTpr100.Tpa, 0)
                        .MgPltTpr.TwFe = .MgPltTpr.TwFe + IIf(ProspData.MgPltTpr100.Fe > 0, ProspData.MgPltTpr100.Tpa, 0)
                        .MgPltTpr.TwAl = .MgPltTpr.TwAl + IIf(ProspData.MgPltTpr100.Al > 0, ProspData.MgPltTpr100.Tpa, 0)
                        .MgPltTpr.TwMg = .MgPltTpr.TwMg + IIf(ProspData.MgPltTpr100.Mg > 0, ProspData.MgPltTpr100.Tpa, 0)
                        .MgPltTpr.TwCa = .MgPltTpr.TwCa + IIf(ProspData.MgPltTpr100.Ca > 0, ProspData.MgPltTpr100.Tpa, 0)
                    Else
                        'Product coefficient -- Add to aSumData.
                        If InclOs = True Then
                            .Os.Tpa = .Os.Tpa + ProspData.Os.Tpa
                            .Os.BplTons = .Os.BplTons + ProspData.Os.Tpa * ProspData.Os.Bpl
                            .Os.InsTons = .Os.InsTons + ProspData.Os.Tpa * ProspData.Os.Ins
                            .Os.IaTons = .Os.IaTons + ProspData.Os.Tpa * ProspData.Os.Ia
                            .Os.FeTons = .Os.FeTons + ProspData.Os.Tpa * ProspData.Os.Fe
                            .Os.AlTons = .Os.AlTons + ProspData.Os.Tpa * ProspData.Os.Al
                            .Os.MgTons = .Os.MgTons + ProspData.Os.Tpa * ProspData.Os.Mg
                            .Os.CaTons = .Os.CaTons + ProspData.Os.Tpa * ProspData.Os.Ca
                            .Os.TwBpl = .Os.TwBpl + IIf(ProspData.Os.Bpl > 0, ProspData.Os.Tpa, 0)
                            .Os.TwIns = .Os.TwIns + IIf(ProspData.Os.Ins > 0, ProspData.Os.Tpa, 0)
                            .Os.TwIa = .Os.TwIa + IIf(ProspData.Os.Ia > 0, ProspData.Os.Tpa, 0)
                            .Os.TwFe = .Os.TwFe + IIf(ProspData.Os.Fe > 0, ProspData.Os.Tpa, 0)
                            .Os.TwAl = .Os.TwAl + IIf(ProspData.Os.Al > 0, ProspData.Os.Tpa, 0)
                            .Os.TwMg = .Os.TwMg + IIf(ProspData.Os.Mg > 0, ProspData.Os.Tpa, 0)
                            .Os.TwCa = .Os.TwCa + IIf(ProspData.Os.Ca > 0, ProspData.Os.Tpa, 0)
                        End If

                        If InclCpb = True Then
                            .Cpb.Tpa = .Cpb.Tpa + ProspData.Cpb.Tpa
                            .Cpb.BplTons = .Cpb.BplTons + ProspData.Cpb.Tpa * ProspData.Cpb.Bpl
                            .Cpb.InsTons = .Cpb.InsTons + ProspData.Cpb.Tpa * ProspData.Cpb.Ins
                            .Cpb.IaTons = .Cpb.IaTons + ProspData.Cpb.Tpa * ProspData.Cpb.Ia
                            .Cpb.FeTons = .Cpb.FeTons + ProspData.Cpb.Tpa * ProspData.Cpb.Fe
                            .Cpb.AlTons = .Cpb.AlTons + ProspData.Cpb.Tpa * ProspData.Cpb.Al
                            .Cpb.MgTons = .Cpb.MgTons + ProspData.Cpb.Tpa * ProspData.Cpb.Mg
                            .Cpb.CaTons = .Cpb.CaTons + ProspData.Cpb.Tpa * ProspData.Cpb.Ca
                            .Cpb.TwBpl = .Cpb.TwBpl + IIf(ProspData.Cpb.Bpl > 0, ProspData.Cpb.Tpa, 0)
                            .Cpb.TwIns = .Cpb.TwIns + IIf(ProspData.Cpb.Ins > 0, ProspData.Cpb.Tpa, 0)
                            .Cpb.TwIa = .Cpb.TwIa + IIf(ProspData.Cpb.Ia > 0, ProspData.Cpb.Tpa, 0)
                            .Cpb.TwFe = .Cpb.TwFe + IIf(ProspData.Cpb.Fe > 0, ProspData.Cpb.Tpa, 0)
                            .Cpb.TwAl = .Cpb.TwAl + IIf(ProspData.Cpb.Al > 0, ProspData.Cpb.Tpa, 0)
                            .Cpb.TwMg = .Cpb.TwMg + IIf(ProspData.Cpb.Mg > 0, ProspData.Cpb.Tpa, 0)
                            .Cpb.TwCa = .Cpb.TwCa + IIf(ProspData.Cpb.Ca > 0, ProspData.Cpb.Tpa, 0)
                        End If

                        If InclFpb = True Then
                            .Fpb.Tpa = .Fpb.Tpa + ProspData.Fpb.Tpa
                            .Fpb.BplTons = .Fpb.BplTons + ProspData.Fpb.Tpa * ProspData.Fpb.Bpl
                            .Fpb.InsTons = .Fpb.InsTons + ProspData.Fpb.Tpa * ProspData.Fpb.Ins
                            .Fpb.IaTons = .Fpb.IaTons + ProspData.Fpb.Tpa * ProspData.Fpb.Ia
                            .Fpb.FeTons = .Fpb.FeTons + ProspData.Fpb.Tpa * ProspData.Fpb.Fe
                            .Fpb.AlTons = .Fpb.AlTons + ProspData.Fpb.Tpa * ProspData.Fpb.Al
                            .Fpb.MgTons = .Fpb.MgTons + ProspData.Fpb.Tpa * ProspData.Fpb.Mg
                            .Fpb.CaTons = .Fpb.CaTons + ProspData.Fpb.Tpa * ProspData.Fpb.Ca
                            .Fpb.TwBpl = .Fpb.TwBpl + IIf(ProspData.Fpb.Bpl > 0, ProspData.Fpb.Tpa, 0)
                            .Fpb.TwIns = .Fpb.TwIns + IIf(ProspData.Fpb.Ins > 0, ProspData.Fpb.Tpa, 0)
                            .Fpb.TwIa = .Fpb.TwIa + IIf(ProspData.Fpb.Ia > 0, ProspData.Fpb.Tpa, 0)
                            .Fpb.TwFe = .Fpb.TwFe + IIf(ProspData.Fpb.Fe > 0, ProspData.Fpb.Tpa, 0)
                            .Fpb.TwAl = .Fpb.TwAl + IIf(ProspData.Fpb.Al > 0, ProspData.Fpb.Tpa, 0)
                            .Fpb.TwMg = .Fpb.TwMg + IIf(ProspData.Fpb.Mg > 0, ProspData.Fpb.Tpa, 0)
                            .Fpb.TwCa = .Fpb.TwCa + IIf(ProspData.Fpb.Ca > 0, ProspData.Fpb.Tpa, 0)
                        End If

                        .Tpb.Tpa = .Tpb.Tpa + ProspData.Tpb.Tpa
                        .Tpb.BplTons = .Tpb.BplTons + ProspData.Tpb.Tpa * ProspData.Tpb.Bpl
                        .Tpb.InsTons = .Tpb.InsTons + ProspData.Tpb.Tpa * ProspData.Tpb.Ins
                        .Tpb.IaTons = .Tpb.IaTons + ProspData.Tpb.Tpa * ProspData.Tpb.Ia
                        .Tpb.FeTons = .Tpb.FeTons + ProspData.Tpb.Tpa * ProspData.Tpb.Fe
                        .Tpb.AlTons = .Tpb.AlTons + ProspData.Tpb.Tpa * ProspData.Tpb.Al
                        .Tpb.MgTons = .Tpb.MgTons + ProspData.Tpb.Tpa * ProspData.Tpb.Mg
                        .Tpb.CaTons = .Tpb.CaTons + ProspData.Tpb.Tpa * ProspData.Tpb.Ca
                        .Tpb.TwBpl = .Tpb.TwBpl + IIf(ProspData.Tpb.Bpl > 0, ProspData.Tpb.Tpa, 0)
                        .Tpb.TwIns = .Tpb.TwIns + IIf(ProspData.Tpb.Ins > 0, ProspData.Tpb.Tpa, 0)
                        .Tpb.TwIa = .Tpb.TwIa + IIf(ProspData.Tpb.Ia > 0, ProspData.Tpb.Tpa, 0)
                        .Tpb.TwFe = .Tpb.TwFe + IIf(ProspData.Tpb.Fe > 0, ProspData.Tpb.Tpa, 0)
                        .Tpb.TwAl = .Tpb.TwAl + IIf(ProspData.Tpb.Al > 0, ProspData.Tpb.Tpa, 0)
                        .Tpb.TwMg = .Tpb.TwMg + IIf(ProspData.Tpb.Mg > 0, ProspData.Tpb.Tpa, 0)
                        .Tpb.TwCa = .Tpb.TwCa + IIf(ProspData.Tpb.Ca > 0, ProspData.Tpb.Tpa, 0)

                        .TpbRej.Tpa = .TpbRej.Tpa + ProspData.TpbRej.Tpa
                        .TpbRej.BplTons = .TpbRej.BplTons + ProspData.TpbRej.Tpa * ProspData.TpbRej.Bpl
                        .TpbRej.InsTons = .TpbRej.InsTons + ProspData.TpbRej.Tpa * ProspData.TpbRej.Ins
                        .TpbRej.IaTons = .TpbRej.IaTons + ProspData.TpbRej.Tpa * ProspData.TpbRej.Ia
                        .TpbRej.FeTons = .TpbRej.FeTons + ProspData.TpbRej.Tpa * ProspData.TpbRej.Fe
                        .TpbRej.AlTons = .TpbRej.AlTons + ProspData.TpbRej.Tpa * ProspData.TpbRej.Al
                        .TpbRej.MgTons = .TpbRej.MgTons + ProspData.TpbRej.Tpa * ProspData.TpbRej.Mg
                        .TpbRej.CaTons = .TpbRej.CaTons + ProspData.TpbRej.Tpa * ProspData.TpbRej.Ca
                        .TpbRej.TwBpl = .TpbRej.TwBpl + IIf(ProspData.TpbRej.Bpl > 0, ProspData.TpbRej.Tpa, 0)
                        .TpbRej.TwIns = .TpbRej.TwIns + IIf(ProspData.TpbRej.Ins > 0, ProspData.TpbRej.Tpa, 0)
                        .TpbRej.TwIa = .TpbRej.TwIa + IIf(ProspData.TpbRej.Ia > 0, ProspData.TpbRej.Tpa, 0)
                        .TpbRej.TwFe = .TpbRej.TwFe + IIf(ProspData.TpbRej.Fe > 0, ProspData.TpbRej.Tpa, 0)
                        .TpbRej.TwAl = .TpbRej.TwAl + IIf(ProspData.TpbRej.Al > 0, ProspData.TpbRej.Tpa, 0)
                        .TpbRej.TwMg = .TpbRej.TwMg + IIf(ProspData.TpbRej.Mg > 0, ProspData.TpbRej.Tpa, 0)
                        .TpbRej.TwCa = .TpbRej.TwCa + IIf(ProspData.TpbRej.Ca > 0, ProspData.TpbRej.Tpa, 0)

                        .Ip.Tpa = .Ip.Tpa + ProspData.Ip.Tpa
                        .Ip.BplTons = .Ip.BplTons + ProspData.Ip.Tpa * ProspData.Ip.Bpl
                        .Ip.InsTons = .Ip.InsTons + ProspData.Ip.Tpa * ProspData.Ip.Ins
                        .Ip.IaTons = .Ip.IaTons + ProspData.Ip.Tpa * ProspData.Ip.Ia
                        .Ip.FeTons = .Ip.FeTons + ProspData.Ip.Tpa * ProspData.Ip.Fe
                        .Ip.AlTons = .Ip.AlTons + ProspData.Ip.Tpa * ProspData.Ip.Al
                        .Ip.MgTons = .Ip.MgTons + ProspData.Ip.Tpa * ProspData.Ip.Mg
                        .Ip.CaTons = .Ip.CaTons + ProspData.Ip.Tpa * ProspData.Ip.Ca
                        .Ip.TwBpl = .Ip.TwBpl + IIf(ProspData.Ip.Bpl > 0, ProspData.Ip.Tpa, 0)
                        .Ip.TwIns = .Ip.TwIns + IIf(ProspData.Ip.Ins > 0, ProspData.Ip.Tpa, 0)
                        .Ip.TwIa = .Ip.TwIa + IIf(ProspData.Ip.Ia > 0, ProspData.Ip.Tpa, 0)
                        .Ip.TwFe = .Ip.TwFe + IIf(ProspData.Ip.Fe > 0, ProspData.Ip.Tpa, 0)
                        .Ip.TwAl = .Ip.TwAl + IIf(ProspData.Ip.Al > 0, ProspData.Ip.Tpa, 0)
                        .Ip.TwMg = .Ip.TwMg + IIf(ProspData.Ip.Mg > 0, ProspData.Ip.Tpa, 0)
                        .Ip.TwCa = .Ip.TwCa + IIf(ProspData.Ip.Ca > 0, ProspData.Ip.Tpa, 0)

                        .Ccn.Tpa = .Ccn.Tpa + ProspData.Ccn.Tpa
                        .Ccn.BplTons = .Ccn.BplTons + ProspData.Ccn.Tpa * ProspData.Ccn.Bpl
                        .Ccn.InsTons = .Ccn.InsTons + ProspData.Ccn.Tpa * ProspData.Ccn.Ins
                        .Ccn.IaTons = .Ccn.IaTons + ProspData.Ccn.Tpa * ProspData.Ccn.Ia
                        .Ccn.FeTons = .Ccn.FeTons + ProspData.Ccn.Tpa * ProspData.Ccn.Fe
                        .Ccn.AlTons = .Ccn.AlTons + ProspData.Ccn.Tpa * ProspData.Ccn.Al
                        .Ccn.MgTons = .Ccn.MgTons + ProspData.Ccn.Tpa * ProspData.Ccn.Mg
                        .Ccn.CaTons = .Ccn.CaTons + ProspData.Ccn.Tpa * ProspData.Ccn.Ca
                        .Ccn.TwBpl = .Ccn.TwBpl + IIf(ProspData.Ccn.Bpl > 0, ProspData.Ccn.Tpa, 0)
                        .Ccn.TwIns = .Ccn.TwIns + IIf(ProspData.Ccn.Ins > 0, ProspData.Ccn.Tpa, 0)
                        .Ccn.TwIa = .Ccn.TwIa + IIf(ProspData.Ccn.Ia > 0, ProspData.Ccn.Tpa, 0)
                        .Ccn.TwFe = .Ccn.TwFe + IIf(ProspData.Ccn.Fe > 0, ProspData.Ccn.Tpa, 0)
                        .Ccn.TwAl = .Ccn.TwAl + IIf(ProspData.Ccn.Al > 0, ProspData.Ccn.Tpa, 0)
                        .Ccn.TwMg = .Ccn.TwMg + IIf(ProspData.Ccn.Mg > 0, ProspData.Ccn.Tpa, 0)
                        .Ccn.TwCa = .Ccn.TwCa + IIf(ProspData.Ccn.Ca > 0, ProspData.Ccn.Tpa, 0)

                        .Fcn.Tpa = .Fcn.Tpa + ProspData.Fcn.Tpa
                        .Fcn.BplTons = .Fcn.BplTons + ProspData.Fcn.Tpa * ProspData.Fcn.Bpl
                        .Fcn.InsTons = .Fcn.InsTons + ProspData.Fcn.Tpa * ProspData.Fcn.Ins
                        .Fcn.IaTons = .Fcn.IaTons + ProspData.Fcn.Tpa * ProspData.Fcn.Ia
                        .Fcn.FeTons = .Fcn.FeTons + ProspData.Fcn.Tpa * ProspData.Fcn.Fe
                        .Fcn.AlTons = .Fcn.AlTons + ProspData.Fcn.Tpa * ProspData.Fcn.Al
                        .Fcn.MgTons = .Fcn.MgTons + ProspData.Fcn.Tpa * ProspData.Fcn.Mg
                        .Fcn.CaTons = .Fcn.CaTons + ProspData.Fcn.Tpa * ProspData.Fcn.Ca
                        .Fcn.TwBpl = .Fcn.TwBpl + IIf(ProspData.Fcn.Bpl > 0, ProspData.Fcn.Tpa, 0)
                        .Fcn.TwIns = .Fcn.TwIns + IIf(ProspData.Fcn.Ins > 0, ProspData.Fcn.Tpa, 0)
                        .Fcn.TwIa = .Fcn.TwIa + IIf(ProspData.Fcn.Ia > 0, ProspData.Fcn.Tpa, 0)
                        .Fcn.TwFe = .Fcn.TwFe + IIf(ProspData.Fcn.Fe > 0, ProspData.Fcn.Tpa, 0)
                        .Fcn.TwAl = .Fcn.TwAl + IIf(ProspData.Fcn.Al > 0, ProspData.Fcn.Tpa, 0)
                        .Fcn.TwMg = .Fcn.TwMg + IIf(ProspData.Fcn.Mg > 0, ProspData.Fcn.Tpa, 0)
                        .Fcn.TwCa = .Fcn.TwCa + IIf(ProspData.Fcn.Ca > 0, ProspData.Fcn.Tpa, 0)

                        .Tcn.Tpa = .Tcn.Tpa + ProspData.Tcn.Tpa
                        .Tcn.BplTons = .Tcn.BplTons + ProspData.Tcn.Tpa * ProspData.Tcn.Bpl
                        .Tcn.InsTons = .Tcn.InsTons + ProspData.Tcn.Tpa * ProspData.Tcn.Ins
                        .Tcn.IaTons = .Tcn.IaTons + ProspData.Tcn.Tpa * ProspData.Tcn.Ia
                        .Tcn.FeTons = .Tcn.FeTons + ProspData.Tcn.Tpa * ProspData.Tcn.Fe
                        .Tcn.AlTons = .Tcn.AlTons + ProspData.Tcn.Tpa * ProspData.Tcn.Al
                        .Tcn.MgTons = .Tcn.MgTons + ProspData.Tcn.Tpa * ProspData.Tcn.Mg
                        .Tcn.CaTons = .Tcn.CaTons + ProspData.Tcn.Tpa * ProspData.Tcn.Ca
                        .Tcn.TwBpl = .Tcn.TwBpl + IIf(ProspData.Tcn.Bpl > 0, ProspData.Tcn.Tpa, 0)
                        .Tcn.TwIns = .Tcn.TwIns + IIf(ProspData.Tcn.Ins > 0, ProspData.Tcn.Tpa, 0)
                        .Tcn.TwIa = .Tcn.TwIa + IIf(ProspData.Tcn.Ia > 0, ProspData.Tcn.Tpa, 0)
                        .Tcn.TwFe = .Tcn.TwFe + IIf(ProspData.Tcn.Fe > 0, ProspData.Tcn.Tpa, 0)
                        .Tcn.TwAl = .Tcn.TwAl + IIf(ProspData.Tcn.Al > 0, ProspData.Tcn.Tpa, 0)
                        .Tcn.TwMg = .Tcn.TwMg + IIf(ProspData.Tcn.Mg > 0, ProspData.Tcn.Tpa, 0)
                        .Tcn.TwCa = .Tcn.TwCa + IIf(ProspData.Tcn.Ca > 0, ProspData.Tcn.Tpa, 0)

                        .Wcl.Tpa = .Wcl.Tpa + ProspData.Wcl.Tpa
                        .Wcl.BplTons = .Wcl.BplTons + ProspData.Wcl.Tpa * ProspData.Wcl.Bpl
                        .Wcl.TwBpl = .Wcl.TwBpl + IIf(ProspData.Wcl.Bpl > 0, ProspData.Wcl.Tpa, 0)

                        .Cfd.Tpa = .Cfd.Tpa + ProspData.Cfd.Tpa
                        .Cfd.BplTons = .Cfd.BplTons + ProspData.Cfd.Tpa * ProspData.Cfd.Bpl
                        .Cfd.TwBpl = .Cfd.TwBpl + IIf(ProspData.Cfd.Bpl > 0, ProspData.Cfd.Tpa, 0)

                        .Ffd.Tpa = .Ffd.Tpa + ProspData.Ffd.Tpa
                        .Ffd.BplTons = .Ffd.BplTons + ProspData.Ffd.Tpa * ProspData.Ffd.Bpl
                        .Ffd.TwBpl = .Ffd.TwBpl + IIf(ProspData.Ffd.Bpl > 0, ProspData.Ffd.Tpa, 0)

                        .Tfd.Tpa = .Tfd.Tpa + ProspData.Tfd.Tpa
                        .Tfd.BplTons = .Tfd.BplTons + ProspData.Tfd.Tpa * ProspData.Tfd.Bpl
                        .Tfd.TwBpl = .Tfd.TwBpl + IIf(ProspData.Tfd.Bpl > 0, ProspData.Tfd.Tpa, 0)

                        .Ttl.Tpa = .Ttl.Tpa + ProspData.Ttl.Tpa
                        .Ttl.BplTons = .Ttl.BplTons + ProspData.Ttl.Tpa * ProspData.Ttl.Bpl
                        .Ttl.TwBpl = .Ttl.TwBpl + IIf(ProspData.Ttl.Bpl > 0, ProspData.Ttl.Tpa, 0)

                        .Tpr.Tpa = .Tpr.Tpa + ProspData.Tpr.Tpa
                        .Tpr.BplTons = .Tpr.BplTons + ProspData.Tpr.Tpa * ProspData.Tpr.Bpl
                        .Tpr.InsTons = .Tpr.InsTons + ProspData.Tpr.Tpa * ProspData.Tpr.Ins
                        .Tpr.IaTons = .Tpr.IaTons + ProspData.Tpr.Tpa * ProspData.Tpr.Ia
                        .Tpr.FeTons = .Tpr.FeTons + ProspData.Tpr.Tpa * ProspData.Tpr.Fe
                        .Tpr.AlTons = .Tpr.AlTons + ProspData.Tpr.Tpa * ProspData.Tpr.Al
                        .Tpr.MgTons = .Tpr.MgTons + ProspData.Tpr.Tpa * ProspData.Tpr.Mg
                        .Tpr.CaTons = .Tpr.CaTons + ProspData.Tpr.Tpa * ProspData.Tpr.Ca
                        .Tpr.TwBpl = .Tpr.TwBpl + IIf(ProspData.Tpr.Bpl > 0, ProspData.Tpr.Tpa, 0)
                        .Tpr.TwIns = .Tpr.TwIns + IIf(ProspData.Tpr.Ins > 0, ProspData.Tpr.Tpa, 0)
                        .Tpr.TwIa = .Tpr.TwIa + IIf(ProspData.Tpr.Ia > 0, ProspData.Tpr.Tpa, 0)
                        .Tpr.TwFe = .Tpr.TwFe + IIf(ProspData.Tpr.Fe > 0, ProspData.Tpr.Tpa, 0)
                        .Tpr.TwAl = .Tpr.TwAl + IIf(ProspData.Tpr.Al > 0, ProspData.Tpr.Tpa, 0)
                        .Tpr.TwMg = .Tpr.TwMg + IIf(ProspData.Tpr.Mg > 0, ProspData.Tpr.Tpa, 0)
                        .Tpr.TwCa = .Tpr.TwCa + IIf(ProspData.Tpr.Ca > 0, ProspData.Tpr.Tpa, 0)

                        .MgPltInp.Tpa = .MgPltInp.Tpa + ProspData.MgPltInp.Tpa
                        .MgPltInp.BplTons = .MgPltInp.BplTons + ProspData.MgPltInp.Tpa * ProspData.MgPltInp.Bpl
                        .MgPltInp.InsTons = .MgPltInp.InsTons + ProspData.MgPltInp.Tpa * ProspData.MgPltInp.Ins
                        .MgPltInp.IaTons = .MgPltInp.IaTons + ProspData.MgPltInp.Tpa * ProspData.MgPltInp.Ia
                        .MgPltInp.FeTons = .MgPltInp.FeTons + ProspData.MgPltInp.Tpa * ProspData.MgPltInp.Fe
                        .MgPltInp.AlTons = .MgPltInp.AlTons + ProspData.MgPltInp.Tpa * ProspData.MgPltInp.Al
                        .MgPltInp.MgTons = .MgPltInp.MgTons + ProspData.MgPltInp.Tpa * ProspData.MgPltInp.Mg
                        .MgPltInp.CaTons = .MgPltInp.CaTons + ProspData.MgPltInp.Tpa * ProspData.MgPltInp.Ca
                        .MgPltInp.TwBpl = .MgPltInp.TwBpl + IIf(ProspData.MgPltInp.Bpl > 0, ProspData.MgPltInp.Tpa, 0)
                        .MgPltInp.TwIns = .MgPltInp.TwIns + IIf(ProspData.MgPltInp.Ins > 0, ProspData.MgPltInp.Tpa, 0)
                        .MgPltInp.TwIa = .MgPltInp.TwIa + IIf(ProspData.MgPltInp.Ia > 0, ProspData.MgPltInp.Tpa, 0)
                        .MgPltInp.TwFe = .MgPltInp.TwFe + IIf(ProspData.MgPltInp.Fe > 0, ProspData.MgPltInp.Tpa, 0)
                        .MgPltInp.TwAl = .MgPltInp.TwAl + IIf(ProspData.MgPltInp.Al > 0, ProspData.MgPltInp.Tpa, 0)
                        .MgPltInp.TwMg = .MgPltInp.TwMg + IIf(ProspData.MgPltInp.Mg > 0, ProspData.MgPltInp.Tpa, 0)
                        .MgPltInp.TwCa = .MgPltInp.TwCa + IIf(ProspData.MgPltInp.Ca > 0, ProspData.MgPltInp.Tpa, 0)

                        .MgPltRej.Tpa = .MgPltRej.Tpa + ProspData.MgPltRej.Tpa
                        .MgPltRej.BplTons = .MgPltRej.BplTons + ProspData.MgPltRej.Tpa * ProspData.MgPltRej.Bpl
                        .MgPltRej.InsTons = .MgPltRej.InsTons + ProspData.MgPltRej.Tpa * ProspData.MgPltRej.Ins
                        .MgPltRej.IaTons = .MgPltRej.IaTons + ProspData.MgPltRej.Tpa * ProspData.MgPltRej.Ia
                        .MgPltRej.FeTons = .MgPltRej.FeTons + ProspData.MgPltRej.Tpa * ProspData.MgPltRej.Fe
                        .MgPltRej.AlTons = .MgPltRej.AlTons + ProspData.MgPltRej.Tpa * ProspData.MgPltRej.Al
                        .MgPltRej.MgTons = .MgPltRej.MgTons + ProspData.MgPltRej.Tpa * ProspData.MgPltRej.Mg
                        .MgPltRej.CaTons = .MgPltRej.CaTons + ProspData.MgPltRej.Tpa * ProspData.MgPltRej.Ca
                        .MgPltRej.TwBpl = .MgPltRej.TwBpl + IIf(ProspData.MgPltRej.Bpl > 0, ProspData.MgPltRej.Tpa, 0)
                        .MgPltRej.TwIns = .MgPltRej.TwIns + IIf(ProspData.MgPltRej.Ins > 0, ProspData.MgPltRej.Tpa, 0)
                        .MgPltRej.TwIa = .MgPltRej.TwIa + IIf(ProspData.MgPltRej.Ia > 0, ProspData.MgPltRej.Tpa, 0)
                        .MgPltRej.TwFe = .MgPltRej.TwFe + IIf(ProspData.MgPltRej.Fe > 0, ProspData.MgPltRej.Tpa, 0)
                        .MgPltRej.TwAl = .MgPltRej.TwAl + IIf(ProspData.MgPltRej.Al > 0, ProspData.MgPltRej.Tpa, 0)
                        .MgPltRej.TwMg = .MgPltRej.TwMg + IIf(ProspData.MgPltRej.Mg > 0, ProspData.MgPltRej.Tpa, 0)
                        .MgPltRej.TwCa = .MgPltRej.TwCa + IIf(ProspData.MgPltRej.Ca > 0, ProspData.MgPltRej.Tpa, 0)

                        .MgPltProd.Tpa = .MgPltProd.Tpa + ProspData.MgPltProd.Tpa
                        .MgPltProd.BplTons = .MgPltProd.BplTons + ProspData.MgPltProd.Tpa * ProspData.MgPltProd.Bpl
                        .MgPltProd.InsTons = .MgPltProd.InsTons + ProspData.MgPltProd.Tpa * ProspData.MgPltProd.Ins
                        .MgPltProd.IaTons = .MgPltProd.IaTons + ProspData.MgPltProd.Tpa * ProspData.MgPltProd.Ia
                        .MgPltProd.FeTons = .MgPltProd.FeTons + ProspData.MgPltProd.Tpa * ProspData.MgPltProd.Fe
                        .MgPltProd.AlTons = .MgPltProd.AlTons + ProspData.MgPltProd.Tpa * ProspData.MgPltProd.Al
                        .MgPltProd.MgTons = .MgPltProd.MgTons + ProspData.MgPltProd.Tpa * ProspData.MgPltProd.Mg
                        .MgPltProd.CaTons = .MgPltProd.CaTons + ProspData.MgPltProd.Tpa * ProspData.MgPltProd.Ca
                        .MgPltProd.TwBpl = .MgPltProd.TwBpl + IIf(ProspData.MgPltProd.Bpl > 0, ProspData.MgPltProd.Tpa, 0)
                        .MgPltProd.TwIns = .MgPltProd.TwIns + IIf(ProspData.MgPltProd.Ins > 0, ProspData.MgPltProd.Tpa, 0)
                        .MgPltProd.TwIa = .MgPltProd.TwIa + IIf(ProspData.MgPltProd.Ia > 0, ProspData.MgPltProd.Tpa, 0)
                        .MgPltProd.TwFe = .MgPltProd.TwFe + IIf(ProspData.MgPltProd.Fe > 0, ProspData.MgPltProd.Tpa, 0)
                        .MgPltProd.TwAl = .MgPltProd.TwAl + IIf(ProspData.MgPltProd.Al > 0, ProspData.MgPltProd.Tpa, 0)
                        .MgPltProd.TwMg = .MgPltProd.TwMg + IIf(ProspData.MgPltProd.Mg > 0, ProspData.MgPltProd.Tpa, 0)
                        .MgPltProd.TwCa = .MgPltProd.TwCa + IIf(ProspData.MgPltProd.Ca > 0, ProspData.MgPltProd.Tpa, 0)

                        .MgPltTcn.Tpa = .MgPltTcn.Tpa + ProspData.MgPltTcn.Tpa
                        .MgPltTcn.BplTons = .MgPltTcn.BplTons + ProspData.MgPltTcn.Tpa * ProspData.MgPltTcn.Bpl
                        .MgPltTcn.InsTons = .MgPltTcn.InsTons + ProspData.MgPltTcn.Tpa * ProspData.MgPltTcn.Ins
                        .MgPltTcn.IaTons = .MgPltTcn.IaTons + ProspData.MgPltTcn.Tpa * ProspData.MgPltTcn.Ia
                        .MgPltTcn.FeTons = .MgPltTcn.FeTons + ProspData.MgPltTcn.Tpa * ProspData.MgPltTcn.Fe
                        .MgPltTcn.AlTons = .MgPltTcn.AlTons + ProspData.MgPltTcn.Tpa * ProspData.MgPltTcn.Al
                        .MgPltTcn.MgTons = .MgPltTcn.MgTons + ProspData.MgPltTcn.Tpa * ProspData.MgPltTcn.Mg
                        .MgPltTcn.CaTons = .MgPltTcn.CaTons + ProspData.MgPltTcn.Tpa * ProspData.MgPltTcn.Ca
                        .MgPltTcn.TwBpl = .MgPltTcn.TwBpl + IIf(ProspData.MgPltTcn.Bpl > 0, ProspData.MgPltTcn.Tpa, 0)
                        .MgPltTcn.TwIns = .MgPltTcn.TwIns + IIf(ProspData.MgPltTcn.Ins > 0, ProspData.MgPltTcn.Tpa, 0)
                        .MgPltTcn.TwIa = .MgPltTcn.TwIa + IIf(ProspData.MgPltTcn.Ia > 0, ProspData.MgPltTcn.Tpa, 0)
                        .MgPltTcn.TwFe = .MgPltTcn.TwFe + IIf(ProspData.MgPltTcn.Fe > 0, ProspData.MgPltTcn.Tpa, 0)
                        .MgPltTcn.TwAl = .MgPltTcn.TwAl + IIf(ProspData.MgPltTcn.Al > 0, ProspData.MgPltTcn.Tpa, 0)
                        .MgPltTcn.TwMg = .MgPltTcn.TwMg + IIf(ProspData.MgPltTcn.Mg > 0, ProspData.MgPltTcn.Tpa, 0)
                        .MgPltTcn.TwCa = .MgPltTcn.TwCa + IIf(ProspData.MgPltTcn.Ca > 0, ProspData.MgPltTcn.Tpa, 0)

                        .MgPltTpr.Tpa = .MgPltTpr.Tpa + ProspData.MgPltTpr.Tpa
                        .MgPltTpr.BplTons = .MgPltTpr.BplTons + ProspData.MgPltTpr.Tpa * ProspData.MgPltTpr.Bpl
                        .MgPltTpr.InsTons = .MgPltTpr.InsTons + ProspData.MgPltTpr.Tpa * ProspData.MgPltTpr.Ins
                        .MgPltTpr.IaTons = .MgPltTpr.IaTons + ProspData.MgPltTpr.Tpa * ProspData.MgPltTpr.Ia
                        .MgPltTpr.FeTons = .MgPltTpr.FeTons + ProspData.MgPltTpr.Tpa * ProspData.MgPltTpr.Fe
                        .MgPltTpr.AlTons = .MgPltTpr.AlTons + ProspData.MgPltTpr.Tpa * ProspData.MgPltTpr.Al
                        .MgPltTpr.MgTons = .MgPltTpr.MgTons + ProspData.MgPltTpr.Tpa * ProspData.MgPltTpr.Mg
                        .MgPltTpr.CaTons = .MgPltTpr.CaTons + ProspData.MgPltTpr.Tpa * ProspData.MgPltTpr.Ca
                        .MgPltTpr.TwBpl = .MgPltTpr.TwBpl + IIf(ProspData.MgPltTpr.Bpl > 0, ProspData.MgPltTpr.Tpa, 0)
                        .MgPltTpr.TwIns = .MgPltTpr.TwIns + IIf(ProspData.MgPltTpr.Ins > 0, ProspData.MgPltTpr.Tpa, 0)
                        .MgPltTpr.TwIa = .MgPltTpr.TwIa + IIf(ProspData.MgPltTpr.Ia > 0, ProspData.MgPltTpr.Tpa, 0)
                        .MgPltTpr.TwFe = .MgPltTpr.TwFe + IIf(ProspData.MgPltTpr.Fe > 0, ProspData.MgPltTpr.Tpa, 0)
                        .MgPltTpr.TwAl = .MgPltTpr.TwAl + IIf(ProspData.MgPltTpr.Al > 0, ProspData.MgPltTpr.Tpa, 0)
                        .MgPltTpr.TwMg = .MgPltTpr.TwMg + IIf(ProspData.MgPltTpr.Mg > 0, ProspData.MgPltTpr.Tpa, 0)
                        .MgPltTpr.TwCa = .MgPltTpr.TwCa + IIf(ProspData.MgPltTpr.Ca > 0, ProspData.MgPltTpr.Tpa, 0)
                    End If
                End If
            Next RowIdx

            'Have summed the data -- put some data in aProspData now.
            'Oversize
            aProspData.Os.Tpa = .Os.Tpa
            If .Os.TwBpl > 0 Then
                aProspData.Os.Bpl = Round(.Os.BplTons / .Os.TwBpl, 1)
            Else
                aProspData.Os.Bpl = 0
            End If
            If .Os.TwIns > 0 Then
                aProspData.Os.Ins = Round(.Os.InsTons / .Os.TwIns, 1)
            Else
                aProspData.Os.Ins = 0
            End If
            If .Os.TwIa > 0 Then
                aProspData.Os.Ia = Round(.Os.IaTons / .Os.TwIa, 2)
            Else
                aProspData.Os.Ia = 0
            End If
            If .Os.TwFe > 0 Then
                aProspData.Os.Fe = Round(.Os.FeTons / .Os.TwFe, 2)
            Else
                aProspData.Os.Fe = 0
            End If
            If .Os.TwAl > 0 Then
                aProspData.Os.Al = Round(.Os.AlTons / .Os.TwAl, 2)
            Else
                aProspData.Os.Al = 0
            End If
            If .Os.TwMg > 0 Then
                aProspData.Os.Mg = Round(.Os.MgTons / .Os.TwMg, 2)
            Else
                aProspData.Os.Mg = 0
            End If
            If .Os.TwCa > 0 Then
                aProspData.Os.Ca = Round(.Os.CaTons / .Os.TwCa, 2)
            Else
                aProspData.Os.Ca = 0
            End If

            'Coarse pebble
            aProspData.Cpb.Tpa = .Cpb.Tpa
            If .Cpb.TwBpl > 0 Then
                aProspData.Cpb.Bpl = Round(.Cpb.BplTons / .Cpb.TwBpl, 1)
            Else
                aProspData.Cpb.Bpl = 0
            End If
            If .Cpb.TwIns > 0 Then
                aProspData.Cpb.Ins = Round(.Cpb.InsTons / .Cpb.TwIns, 1)
            Else
                aProspData.Cpb.Ins = 0
            End If
            If .Cpb.TwIa > 0 Then
                aProspData.Cpb.Ia = Round(.Cpb.IaTons / .Cpb.TwIa, 2)
            Else
                aProspData.Cpb.Ia = 0
            End If
            If .Cpb.TwFe > 0 Then
                aProspData.Cpb.Fe = Round(.Cpb.FeTons / .Cpb.TwFe, 2)
            Else
                aProspData.Cpb.Fe = 0
            End If
            If .Cpb.TwAl > 0 Then
                aProspData.Cpb.Al = Round(.Cpb.AlTons / .Cpb.TwAl, 2)
            Else
                aProspData.Cpb.Al = 0
            End If
            If .Cpb.TwMg > 0 Then
                aProspData.Cpb.Mg = Round(.Cpb.MgTons / .Cpb.TwMg, 2)
            Else
                aProspData.Cpb.Mg = 0
            End If
            If .Cpb.TwCa > 0 Then
                aProspData.Cpb.Ca = Round(.Cpb.CaTons / .Cpb.TwCa, 2)
            Else
                aProspData.Cpb.Ca = 0
            End If

            'Fine pebble
            aProspData.Fpb.Tpa = .Fpb.Tpa
            If .Fpb.TwBpl > 0 Then
                aProspData.Fpb.Bpl = Round(.Fpb.BplTons / .Fpb.TwBpl, 1)
            Else
                aProspData.Fpb.Bpl = 0
            End If
            If .Fpb.TwIns > 0 Then
                aProspData.Fpb.Ins = Round(.Fpb.InsTons / .Fpb.TwIns, 2)
            Else
                aProspData.Fpb.Ins = 0
            End If
            If .Fpb.TwIa > 0 Then
                aProspData.Fpb.Ia = Round(.Fpb.IaTons / .Fpb.TwIa, 2)
            Else
                aProspData.Fpb.Ia = 0
            End If
            If .Fpb.TwFe > 0 Then
                aProspData.Fpb.Fe = Round(.Fpb.FeTons / .Fpb.TwFe, 2)
            Else
                aProspData.Fpb.Fe = 0
            End If
            If .Fpb.TwAl > 0 Then
                aProspData.Fpb.Al = Round(.Fpb.AlTons / .Fpb.TwAl, 2)
            Else
                aProspData.Fpb.Al = 0
            End If
            If .Fpb.TwMg > 0 Then
                aProspData.Fpb.Mg = Round(.Fpb.MgTons / .Fpb.TwMg, 2)
            Else
                aProspData.Fpb.Mg = 0
            End If
            If .Fpb.TwCa > 0 Then
                aProspData.Fpb.Ca = Round(.Fpb.CaTons / .Fpb.TwCa, 2)
            Else
                aProspData.Fpb.Ca = 0
            End If

            'Total pebble
            aProspData.Tpb.Tpa = .Tpb.Tpa
            If .Tpb.TwBpl > 0 Then
                aProspData.Tpb.Bpl = Round(.Tpb.BplTons / .Tpb.TwBpl, 1)
            Else
                aProspData.Tpb.Bpl = 0
            End If
            If .Tpb.TwIns > 0 Then
                aProspData.Tpb.Ins = Round(.Tpb.InsTons / .Tpb.TwIns, 2)
            Else
                aProspData.Tpb.Ins = 0
            End If
            If .Tpb.TwIa > 0 Then
                aProspData.Tpb.Ia = Round(.Tpb.IaTons / .Tpb.TwIa, 2)
            Else
                aProspData.Tpb.Ia = 0
            End If
            If .Tpb.TwFe > 0 Then
                aProspData.Tpb.Fe = Round(.Tpb.FeTons / .Tpb.TwFe, 2)
            Else
                aProspData.Tpb.Fe = 0
            End If
            If .Tpb.TwAl > 0 Then
                aProspData.Tpb.Al = Round(.Tpb.AlTons / .Tpb.TwAl, 2)
            Else
                aProspData.Tpb.Al = 0
            End If
            If .Tpb.TwMg > 0 Then
                aProspData.Tpb.Mg = Round(.Tpb.MgTons / .Tpb.TwMg, 2)
            Else
                aProspData.Tpb.Mg = 0
            End If
            If .Tpb.TwCa > 0 Then
                aProspData.Tpb.Ca = Round(.Tpb.CaTons / .Tpb.TwCa, 2)
            Else
                aProspData.Tpb.Ca = 0
            End If

            'IP
            aProspData.Ip.Tpa = .Ip.Tpa
            If .Ip.TwBpl > 0 Then
                aProspData.Ip.Bpl = Round(.Ip.BplTons / .Ip.TwBpl, 1)
            Else
                aProspData.Ip.Bpl = 0
            End If
            If .Ip.TwIns > 0 Then
                aProspData.Ip.Ins = Round(.Ip.InsTons / .Ip.TwIns, 1)
            Else
                aProspData.Ip.Ins = 0
            End If
            If .Ip.TwIa > 0 Then
                aProspData.Ip.Ia = Round(.Ip.IaTons / .Ip.TwIa, 2)
            Else
                aProspData.Ip.Ia = 0
            End If
            If .Ip.TwFe > 0 Then
                aProspData.Ip.Fe = Round(.Ip.FeTons / .Ip.TwFe, 2)
            Else
                aProspData.Ip.Fe = 0
            End If
            If .Ip.TwAl > 0 Then
                aProspData.Ip.Al = Round(.Ip.AlTons / .Ip.TwAl, 2)
            Else
                aProspData.Ip.Al = 0
            End If
            If .Ip.TwMg > 0 Then
                aProspData.Ip.Mg = Round(.Ip.MgTons / .Ip.TwMg, 2)
            Else
                aProspData.Ip.Mg = 0
            End If
            If .Ip.TwCa > 0 Then
                aProspData.Ip.Ca = Round(.Ip.CaTons / .Ip.TwCa, 2)
            Else
                aProspData.Ip.Ca = 0
            End If

            'Coarse concentrate
            aProspData.Ccn.Tpa = .Ccn.Tpa
            If .Ccn.TwBpl > 0 Then
                aProspData.Ccn.Bpl = Round(.Ccn.BplTons / .Ccn.TwBpl, 1)
            Else
                aProspData.Ccn.Bpl = 0
            End If
            If .Ccn.TwIns > 0 Then
                aProspData.Ccn.Ins = Round(.Ccn.InsTons / .Ccn.TwIns, 1)
            Else
                aProspData.Ccn.Ins = 0
            End If
            If .Ccn.TwIa > 0 Then
                aProspData.Ccn.Ia = Round(.Ccn.IaTons / .Ccn.TwIa, 1)
            Else
                aProspData.Ccn.Ia = 0
            End If
            If .Ccn.TwFe > 0 Then
                aProspData.Ccn.Fe = Round(.Ccn.FeTons / .Ccn.TwFe, 1)
            Else
                aProspData.Ccn.Fe = 0
            End If
            If .Ccn.TwAl > 0 Then
                aProspData.Ccn.Al = Round(.Ccn.AlTons / .Ccn.TwAl, 1)
            Else
                aProspData.Ccn.Al = 0
            End If
            If .Ccn.TwMg > 0 Then
                aProspData.Ccn.Mg = Round(.Ccn.MgTons / .Ccn.TwMg, 1)
            Else
                aProspData.Ccn.Mg = 0
            End If
            If .Ccn.TwCa > 0 Then
                aProspData.Ccn.Ca = Round(.Ccn.CaTons / .Ccn.TwCa, 1)
            Else
                aProspData.Ccn.Ca = 0
            End If

            'Fine concentrate
            aProspData.Fcn.Tpa = .Fcn.Tpa
            If .Fcn.TwBpl > 0 Then
                aProspData.Fcn.Bpl = Round(.Fcn.BplTons / .Fcn.TwBpl, 1)
            Else
                aProspData.Fcn.Bpl = 0
            End If
            If .Fcn.TwIns > 0 Then
                aProspData.Fcn.Ins = Round(.Fcn.InsTons / .Fcn.TwIns, 1)
            Else
                aProspData.Fcn.Ins = 0
            End If
            If .Fcn.TwIa > 0 Then
                aProspData.Fcn.Ia = Round(.Fcn.IaTons / .Fcn.TwIa, 2)
            Else
                aProspData.Fcn.Ia = 0
            End If
            If .Fcn.TwFe > 0 Then
                aProspData.Fcn.Fe = Round(.Fcn.FeTons / .Fcn.TwFe, 2)
            Else
                aProspData.Fcn.Fe = 0
            End If
            If .Fcn.TwAl > 0 Then
                aProspData.Fcn.Al = Round(.Fcn.AlTons / .Fcn.TwAl, 2)
            Else
                aProspData.Fcn.Al = 0
            End If
            If .Fcn.TwMg > 0 Then
                aProspData.Fcn.Mg = Round(.Fcn.MgTons / .Fcn.TwMg, 2)
            Else
                aProspData.Fcn.Mg = 0
            End If
            If .Fcn.TwCa > 0 Then
                aProspData.Fcn.Ca = Round(.Fcn.CaTons / .Fcn.TwCa, 2)
            Else
                aProspData.Fcn.Ca = 0
            End If

            'Total concentrate
            aProspData.Tcn.Tpa = .Tcn.Tpa
            If .Tcn.TwBpl > 0 Then
                aProspData.Tcn.Bpl = Round(.Tcn.BplTons / .Tcn.TwBpl, 1)
            Else
                aProspData.Tcn.Bpl = 0
            End If
            If .Tcn.TwIns > 0 Then
                aProspData.Tcn.Ins = Round(.Tcn.InsTons / .Tcn.TwIns, 1)
            Else
                aProspData.Tcn.Ins = 0
            End If
            If .Tcn.TwIa > 0 Then
                aProspData.Tcn.Ia = Round(.Tcn.IaTons / .Tcn.TwIa, 2)
            Else
                aProspData.Tcn.Ia = 0
            End If
            If .Tcn.TwFe > 0 Then
                aProspData.Tcn.Fe = Round(.Tcn.FeTons / .Tcn.TwFe, 2)
            Else
                aProspData.Tcn.Fe = 0
            End If
            If .Tcn.TwAl > 0 Then
                aProspData.Tcn.Al = Round(.Tcn.AlTons / .Tcn.TwAl, 2)
            Else
                aProspData.Tcn.Al = 0
            End If
            If .Tcn.TwMg > 0 Then
                aProspData.Tcn.Mg = Round(.Tcn.MgTons / .Tcn.TwMg, 2)
            Else
                aProspData.Tcn.Mg = 0
            End If
            If .Tcn.TwCa > 0 Then
                aProspData.Tcn.Ca = Round(.Tcn.CaTons / .Tcn.TwCa, 2)
            Else
                aProspData.Tcn.Ca = 0
            End If

            'Waste clay
            aProspData.Wcl.Tpa = .Wcl.Tpa
            If .Wcl.TwBpl > 0 Then
                aProspData.Wcl.Bpl = Round(.Wcl.BplTons / .Wcl.TwBpl, 1)
            Else
                aProspData.Wcl.Bpl = 0
            End If

            'Coarse feed
            aProspData.Cfd.Tpa = .Cfd.Tpa
            If .Cfd.TwBpl > 0 Then
                aProspData.Cfd.Bpl = Round(.Cfd.BplTons / .Cfd.TwBpl, 1)
            Else
                aProspData.Cfd.Bpl = 0
            End If

            'Fine feed
            aProspData.Ffd.Tpa = .Ffd.Tpa
            If .Ffd.TwBpl > 0 Then
                aProspData.Ffd.Bpl = Round(.Ffd.BplTons / .Ffd.TwBpl, 1)
            Else
                aProspData.Ffd.Bpl = 0
            End If

            'Total feed
            aProspData.Tfd.Tpa = .Tfd.Tpa
            If .Tfd.TwBpl > 0 Then
                aProspData.Tfd.Bpl = Round(.Tfd.BplTons / .Tfd.TwBpl, 1)
            Else
                aProspData.Tfd.Bpl = 0
            End If

            'Total tails
            aProspData.Ttl.Tpa = .Ttl.Tpa
            If .Ttl.TwBpl > 0 Then
                aProspData.Ttl.Bpl = Round(.Ttl.BplTons / .Ttl.TwBpl, 1)
            Else
                aProspData.Ttl.Bpl = 0
            End If

            'Total product
            aProspData.Tpr.Tpa = .Tpr.Tpa
            If .Tpr.TwBpl > 0 Then
                aProspData.Tpr.Bpl = Round(.Tpr.BplTons / .Tpr.TwBpl, 1)
            Else
                aProspData.Tpr.Bpl = 0
            End If
            If .Tpr.TwIns > 0 Then
                aProspData.Tpr.Ins = Round(.Tpr.InsTons / .Tpr.TwIns, 2)
            Else
                aProspData.Tpr.Ins = 0
            End If
            If .Tpr.TwIa > 0 Then
                aProspData.Tpr.Ia = Round(.Tpr.IaTons / .Tpr.TwIa, 2)
            Else
                aProspData.Tpr.Ia = 0
            End If
            If .Tpr.TwFe > 0 Then
                aProspData.Tpr.Fe = Round(.Tpr.FeTons / .Tpr.TwFe, 2)
            Else
                aProspData.Tpr.Fe = 0
            End If
            If .Tpr.TwAl > 0 Then
                aProspData.Tpr.Al = Round(.Tpr.AlTons / .Tpr.TwAl, 2)
            Else
                aProspData.Tpr.Al = 0
            End If
            If .Tpr.TwMg > 0 Then
                aProspData.Tpr.Mg = Round(.Tpr.MgTons / .Tpr.TwMg, 2)
            Else
                aProspData.Tpr.Mg = 0
            End If
            If .Tpr.TwCa > 0 Then
                aProspData.Tpr.Ca = Round(.Tpr.CaTons / .Tpr.TwCa, 2)
            Else
                aProspData.Tpr.Ca = 0
            End If

            'MgO plant input
            aProspData.MgPltInp.Tpa = .MgPltInp.Tpa
            If .MgPltInp.TwBpl > 0 Then
                aProspData.MgPltInp.Bpl = Round(.MgPltInp.BplTons / .MgPltInp.TwBpl, 1)
            Else
                aProspData.MgPltInp.Bpl = 0
            End If
            If .MgPltInp.TwIns > 0 Then
                aProspData.MgPltInp.Ins = Round(.MgPltInp.InsTons / .MgPltInp.TwIns, 2)
            Else
                aProspData.MgPltInp.Ins = 0
            End If
            If .MgPltInp.TwIa > 0 Then
                aProspData.MgPltInp.Ia = Round(.MgPltInp.IaTons / .MgPltInp.TwIa, 2)
            Else
                aProspData.MgPltInp.Ia = 0
            End If
            If .MgPltInp.TwFe > 0 Then
                aProspData.MgPltInp.Fe = Round(.MgPltInp.FeTons / .MgPltInp.TwFe, 2)
            Else
                aProspData.MgPltInp.Fe = 0
            End If
            If .MgPltInp.TwAl > 0 Then
                aProspData.MgPltInp.Al = Round(.MgPltInp.AlTons / .MgPltInp.TwAl, 2)
            Else
                aProspData.MgPltInp.Al = 0
            End If
            If .MgPltInp.TwMg > 0 Then
                aProspData.MgPltInp.Mg = Round(.MgPltInp.MgTons / .MgPltInp.TwMg, 2)
            Else
                aProspData.MgPltInp.Mg = 0
            End If
            If .MgPltInp.TwCa > 0 Then
                aProspData.MgPltInp.Ca = Round(.MgPltInp.CaTons / .MgPltInp.TwCa, 2)
            Else
                aProspData.MgPltInp.Ca = 0
            End If

            'MgO plant reject
            aProspData.MgPltRej.Tpa = .MgPltRej.Tpa
            If .MgPltRej.TwBpl > 0 Then
                aProspData.MgPltRej.Bpl = Round(.MgPltRej.BplTons / .MgPltRej.TwBpl, 1)
            Else
                aProspData.MgPltRej.Bpl = 0
            End If
            If .MgPltRej.TwIns > 0 Then
                aProspData.MgPltRej.Ins = Round(.MgPltRej.InsTons / .MgPltRej.TwIns, 2)
            Else
                aProspData.MgPltRej.Ins = 0
            End If
            If .MgPltRej.TwIa > 0 Then
                aProspData.MgPltRej.Ia = Round(.MgPltRej.IaTons / .MgPltRej.TwIa, 2)
            Else
                aProspData.MgPltRej.Ia = 0
            End If
            If .MgPltRej.TwFe > 0 Then
                aProspData.MgPltRej.Fe = Round(.MgPltRej.FeTons / .MgPltRej.TwFe, 2)
            Else
                aProspData.MgPltRej.Fe = 0
            End If
            If .MgPltRej.TwAl > 0 Then
                aProspData.MgPltRej.Al = Round(.MgPltRej.AlTons / .MgPltRej.TwAl, 2)
            Else
                aProspData.MgPltRej.Al = 0
            End If
            If .MgPltRej.TwMg > 0 Then
                aProspData.MgPltRej.Mg = Round(.MgPltRej.MgTons / .MgPltRej.TwMg, 2)
            Else
                aProspData.MgPltRej.Mg = 0
            End If
            If .MgPltRej.TwCa > 0 Then
                aProspData.MgPltRej.Ca = Round(.MgPltRej.CaTons / .MgPltRej.TwCa, 2)
            Else
                aProspData.MgPltRej.Ca = 0
            End If

            'MgO plant product
            aProspData.MgPltProd.Tpa = .MgPltProd.Tpa
            If .MgPltProd.TwBpl > 0 Then
                aProspData.MgPltProd.Bpl = Round(.MgPltProd.BplTons / .MgPltProd.TwBpl, 1)
            Else
                aProspData.MgPltProd.Bpl = 0
            End If
            If .MgPltProd.TwIns > 0 Then
                aProspData.MgPltProd.Ins = Round(.MgPltProd.InsTons / .MgPltProd.TwIns, 2)
            Else
                aProspData.MgPltProd.Ins = 0
            End If
            If .MgPltProd.TwIa > 0 Then
                aProspData.MgPltProd.Ia = Round(.MgPltProd.IaTons / .MgPltProd.TwIa, 2)
            Else
                aProspData.MgPltProd.Ia = 0
            End If
            If .MgPltProd.TwFe > 0 Then
                aProspData.MgPltProd.Fe = Round(.MgPltProd.FeTons / .MgPltProd.TwFe, 2)
            Else
                aProspData.MgPltProd.Fe = 0
            End If
            If .MgPltProd.TwAl > 0 Then
                aProspData.MgPltProd.Al = Round(.MgPltProd.AlTons / .MgPltProd.TwAl, 2)
            Else
                aProspData.MgPltProd.Al = 0
            End If
            If .MgPltProd.TwMg > 0 Then
                aProspData.MgPltProd.Mg = Round(.MgPltProd.MgTons / .MgPltProd.TwMg, 2)
            Else
                aProspData.MgPltProd.Mg = 0
            End If
            If .MgPltProd.TwCa > 0 Then
                aProspData.MgPltProd.Ca = Round(.MgPltProd.CaTons / .MgPltProd.TwCa, 2)
            Else
                aProspData.MgPltProd.Ca = 0
            End If

            'MgO plant total concentrate
            aProspData.MgPltTcn.Tpa = .MgPltTcn.Tpa
            If .MgPltTcn.TwBpl > 0 Then
                aProspData.MgPltTcn.Bpl = Round(.MgPltTcn.BplTons / .MgPltTcn.TwBpl, 1)
            Else
                aProspData.MgPltTcn.Bpl = 0
            End If
            If .MgPltTcn.TwIns > 0 Then
                aProspData.MgPltTcn.Ins = Round(.MgPltTcn.InsTons / .MgPltTcn.TwIns, 2)
            Else
                aProspData.MgPltTcn.Ins = 0
            End If
            If .MgPltTcn.TwIa > 0 Then
                aProspData.MgPltTcn.Ia = Round(.MgPltTcn.IaTons / .MgPltTcn.TwIa, 2)
            Else
                aProspData.MgPltTcn.Ia = 0
            End If
            If .MgPltTcn.TwFe > 0 Then
                aProspData.MgPltTcn.Fe = Round(.MgPltTcn.FeTons / .MgPltTcn.TwFe, 2)
            Else
                aProspData.MgPltTcn.Fe = 0
            End If
            If .MgPltTcn.TwAl > 0 Then
                aProspData.MgPltTcn.Al = Round(.MgPltTcn.AlTons / .MgPltTcn.TwAl, 2)
            Else
                aProspData.MgPltTcn.Al = 0
            End If
            If .MgPltTcn.TwMg > 0 Then
                aProspData.MgPltTcn.Mg = Round(.MgPltTcn.MgTons / .MgPltTcn.TwMg, 2)
            Else
                aProspData.MgPltTcn.Mg = 0
            End If
            If .MgPltTcn.TwCa > 0 Then
                aProspData.MgPltTcn.Ca = Round(.MgPltTcn.CaTons / .MgPltTcn.TwCa, 2)
            Else
                aProspData.MgPltTcn.Ca = 0
            End If

            'MgO plant total product
            aProspData.MgPltTpr.Tpa = .MgPltTpr.Tpa
            If .MgPltTpr.TwBpl > 0 Then
                aProspData.MgPltTpr.Bpl = Round(.MgPltTpr.BplTons / .MgPltTpr.TwBpl, 1)
            Else
                aProspData.MgPltTpr.Bpl = 0
            End If
            If .MgPltTpr.TwIns > 0 Then
                aProspData.MgPltTpr.Ins = Round(.MgPltTpr.InsTons / .MgPltTpr.TwIns, 2)
            Else
                aProspData.MgPltTpr.Ins = 0
            End If
            If .MgPltTpr.TwIa > 0 Then
                aProspData.MgPltTpr.Ia = Round(.MgPltTpr.IaTons / .MgPltTpr.TwIa, 2)
            Else
                aProspData.MgPltTpr.Ia = 0
            End If
            If .MgPltTpr.TwFe > 0 Then
                aProspData.MgPltTpr.Fe = Round(.MgPltTpr.FeTons / .MgPltTpr.TwFe, 2)
            Else
                aProspData.MgPltTpr.Fe = 0
            End If
            If .MgPltTpr.TwAl > 0 Then
                aProspData.MgPltTpr.Al = Round(.MgPltTpr.AlTons / .MgPltTpr.TwAl, 2)
            Else
                aProspData.MgPltTpr.Al = 0
            End If
            If .MgPltTpr.TwMg > 0 Then
                aProspData.MgPltTpr.Mg = Round(.MgPltTpr.MgTons / .MgPltTpr.TwMg, 2)
            Else
                aProspData.MgPltTpr.Mg = 0
            End If
            If .MgPltTpr.TwCa > 0 Then
                aProspData.MgPltTpr.Ca = Round(.MgPltTpr.CaTons / .MgPltTpr.TwCa, 2)
            Else
                aProspData.MgPltTpr.Ca = 0
            End If

            TotWt = aProspData.Tpb.Tpa + aProspData.Tfd.Tpa + aProspData.Wcl.Tpa +
                    aProspData.Os.Tpa + aProspData.Ip.Tpa

            If TotWt > 0 Then
                aProspData.Os.WtPct = Round(aProspData.Os.Tpa / TotWt * 100, 2)
                aProspData.Cpb.WtPct = Round(aProspData.Cpb.Tpa / TotWt * 100, 2)
                aProspData.Fpb.WtPct = Round(aProspData.Fpb.Tpa / TotWt * 100, 2)
                aProspData.Tpb.WtPct = Round(aProspData.Tpb.Tpa / TotWt * 100, 2)
                aProspData.Ccn.WtPct = Round(aProspData.Ccn.Tpa / TotWt * 100, 2)
                aProspData.Fcn.WtPct = Round(aProspData.Fcn.Tpa / TotWt * 100, 2)
                aProspData.Tcn.WtPct = Round(aProspData.Tcn.Tpa / TotWt * 100, 2)
                aProspData.Tpr.WtPct = Round(aProspData.Tpr.Tpa / TotWt * 100, 2)
                aProspData.Ttl.WtPct = Round(aProspData.Ttl.Tpa / TotWt * 100, 2)
                aProspData.Wcl.WtPct = Round(aProspData.Wcl.Tpa / TotWt * 100, 2)
                aProspData.Cfd.WtPct = Round(aProspData.Cfd.Tpa / TotWt * 100, 2)
                aProspData.Ffd.WtPct = Round(aProspData.Ffd.Tpa / TotWt * 100, 2)
                aProspData.Tfd.WtPct = Round(aProspData.Tfd.Tpa / TotWt * 100, 2)
                aProspData.Ip.WtPct = Round(aProspData.Ip.Tpa / TotWt * 100, 2)
            Else
                aProspData.Os.WtPct = 0
                aProspData.Cpb.WtPct = 0
                aProspData.Fpb.WtPct = 0
                aProspData.Tpb.WtPct = 0
                aProspData.Ccn.WtPct = 0
                aProspData.Fcn.WtPct = 0
                aProspData.Tcn.WtPct = 0
                aProspData.Tpr.WtPct = 0
                aProspData.Ttl.WtPct = 0
                aProspData.Wcl.WtPct = 0
                aProspData.Cfd.WtPct = 0
                aProspData.Ffd.WtPct = 0
                aProspData.Tfd.WtPct = 0
            End If

            If aMinableHoleCount > 0 Then
                aProspData.OvbThk = Round(.OvbThk / aMinableHoleCount, 1)
            Else
                aProspData.OvbThk = 0
            End If
            If aMinableHoleCount > 0 Then
                aProspData.ItbThk = Round(.ItbThk / aMinableHoleCount, 1)
            Else
                aProspData.ItbThk = 0
            End If
            If aMinableHoleCount > 0 Then
                aProspData.MtxThk = Round(.MtxThk / aMinableHoleCount, 1)
            Else
                aProspData.MtxThk = 0
            End If

            If aProspData.Tpr.Tpa > 0 Then
                aProspData.MtxxAllPcHole = Round(((43560 * aProspData.MtxThk / 27) * aMinableHoleCount) /
                                           aProspData.Tpr.Tpa, 2)
            Else
                aProspData.MtxxAllPcHole = 0
            End If
            If aProspData.Tpr.Tpa > 0 Then
                aProspData.TotxAllPcHole = Round(((43560 * (aProspData.MtxThk + aProspData.OvbThk +
                                           aProspData.ItbThk) * aMinableHoleCount) / 27) /
                                           aProspData.Tpr.Tpa, 2)
            Else
                aProspData.TotxAllPcHole = 0
            End If
        End With

        Exit Sub

SumTheHoleDataError:
        MsgBox("Error summing holes." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Process Error")
    End Sub

    Private Sub cmdSetDefaults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSetDefaults.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'BLUESTARMgOMODEL.xls --> from Glen Oswald (2010)

        ' 1) FnePb MgO cutoff
        ' 2) IP MgO cutoff
        ' 3) Grind
        ' 4) Acid
        ' 5) P2O5
        ' 6) PA64
        ' 7) Flot minutes
        ' 8) Target MgO
        ' 9) Al2O3 >
        '10) Fe2O3 >

        '1) CrsPb MgO cutoff -- 06/11/2010, lss  Had this too but don't need.
        '                       The coarse pebble can only go to product or reject.
        '                       It never goes to the Doloflot plant.

        With ssDoloflotPlant
            .Col = 1
            .Row = 1
            .Value = 2.5    'FnePb MgO cutoff
            .Row = 2
            .Value = 2.5    'IP MgO cutoff
            .Row = 3
            .Value = 70     'Grind
            .Row = 4
            .Value = 4.6    'Acid
            .Row = 5
            .Value = 9.3    'P2O5
            .Row = 6
            .Value = 10     'PA64
            .Row = 7
            .Value = 20     'Flot minutes
            .Row = 8
            .Value = 0.9    'Target MgO
            .Row = 9
            .Value = 1      'Al2O3 >
            .Row = 10
            .Value = 1      'Fe2O3 >
        End With

        ''Changes 11/16/2011
        ''With ssDoloflotPlantFco
        ''    .Col = 1
        ''    .Row = 1
        ''    .Value = 0    'TotPb MgO Max
        ''    .Row = 2
        ''    .Value = 0    'TotPb MgO Min
        ''End With

        chkUseDoloflotPlant.Checked = True
        chkUseDoloflotPlantFco.Checked = False
    End Sub

    Private Sub cmdSetDefaults2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSetDefaults2.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'New Model --> from Glen Oswald (2011)

        ' 1) %Wt -200 mesh
        ' 2) Conditioning minutes
        ' 3) Conditioning %solids
        ' 4) Flotation minutes
        ' 5) PA64 lbs/ton
        ' 6) Phos aid lbs/ton
        ' 7) Sulfuric acid lbs/ton

        With ssDoloflotPlantFco2
            .Col = 1
            .Row = 1
            .Value = 60     '%Wt -200 mesh
            .Row = 2
            .Value = 0.5    'Conditioning minutes
            .Row = 3
            .Value = 36     'Conditioning %solids
            .Row = 4
            .Value = 5      'Flotation minutes
            .Row = 5
            .Value = 13     'PA64 lbs/ton
            .Row = 6
            .Value = 9.09   'Phos aid lbs/ton
            .Row = 7
            .Value = 10     'Sulfuric acid lbs/ton
        End With

        With ssDoloflotPlantFco
            .Col = 1
            .Row = 1
            .Value = 4    'TotPb MgO Max
            .Row = 2
            .Value = 1    'TotPb MgO Min
        End With

        chkUseDoloflotPlant.Checked = False
        chkUseDoloflotPlantFco.Checked = True
    End Sub

    Private Sub chkUseDoloflotPlant_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseDoloflotPlant.CheckedChanged

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        If chkUseDoloflotPlant.Checked = True Then
            chkUseOrigMgoPlant.Checked = False
            chkUseDoloflotPlantFco.Checked = False
        End If
    End Sub

    Private Sub chkUseDoloflotPlantFco_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseDoloflotPlantFco.CheckedChanged

        'If chkUseDoloflotPlantFco.Checked = True Then
        '    chkInclCpbNever.Checked = False
        '    chkInclFpbNever.Checked = False
        '    chkInclOsNever.Checked = True
        '    chkInclCpbAlways.Checked = False
        '    chkInclFpbAlways.Checked = False
        '    chkInclOsAlways.Checked = False
        '    chkCanSelectRejectTpb.Checked = False
        'End If

        If chkUseDoloflotPlantFco.Checked = True Then
            chkUseOrigMgoPlant.Checked = False
            chkUseDoloflotPlant.Checked = False

            With ssDoloflotPlant
                .Col = 1
                .Row = 1
                .Value = 0
                .Row = 2
                .Value = 0
            End With
        End If
    End Sub

    Private Sub cmdPrintRept_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrintRept.Click
        On Error GoTo DocPrintError

        SetActionStatus("Printing report...")
        Me.Cursor = Cursors.WaitCursor
        Dim dlg As New PrintDialog
        Dim pd As New PrintDocument()
        dlg.Document = pd
        dlg.AllowSelection = True
        dlg.AllowSomePages = False


        stringToPrint = rtbRept1.Text
        AddHandler pd.PrintPage, AddressOf printDocument1_PrintPage
        'Microsoft.VisualBasic.p.Print(" ")
        'rtbRept1.SelPrint(Printer.hDC, 0)
        'Printer.EndDoc()
        If (dlg.ShowDialog = System.Windows.Forms.DialogResult.OK) Then
            pd.Print()
        End If
        SetActionStatus("")
        Me.Cursor = Cursors.Arrow

        Exit Sub

DocPrintError:
        If Err.Number = 482 Then
            'do nothing
        Else
            Err.Raise(Err.Number)
        End If

        On Error Resume Next
        SetActionStatus("")
        On Error Resume Next
        Me.Cursor = Cursors.Arrow
    End Sub
    Private Sub printDocument1_PrintPage(ByVal sender As Object,
    ByVal e As PrintPageEventArgs)

        Dim charactersOnPage As Integer = 0
        Dim linesPerPage As Integer = 0

        ' Sets the value of charactersOnPage to the number of characters 
        ' of stringToPrint that will fit within the bounds of the page.
        e.Graphics.MeasureString(stringToPrint, Me.Font, e.MarginBounds.Size,
            StringFormat.GenericTypographic, charactersOnPage, linesPerPage)

        ' Draws the string within the bounds of the page
        e.Graphics.DrawString(stringToPrint, Me.Font, Brushes.Black,
            e.MarginBounds, StringFormat.GenericTypographic)

        ' Remove the portion of the string that has been printed.
        stringToPrint = stringToPrint.Substring(charactersOnPage)

        ' Check to see if more pages are to be printed.
        e.HasMorePages = stringToPrint.Length > 0

    End Sub

    Private Sub cmdExitRept_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExitRept.Click

        fraReptDisp.Visible = False
        fraReview.Visible = True
    End Sub

End Class