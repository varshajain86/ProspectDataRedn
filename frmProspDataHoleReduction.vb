Imports OracleInProcServer
Imports System.Math
Imports AxFPSpread
Imports System.Globalization
Imports FPSpread
Imports CrystalDecisions.CrystalReports.Engine
Imports System.IO
Imports ProspectDataReduction.ReductionService
Imports ProspectDataReduction.ViewModels

Public Class frmProspDataHoleReduction
    Dim fRcvryScenDynaset As OraDynaset
    Dim fReducing As Boolean
    Dim fUserIsAdmin As Boolean
    Dim fMinabilityUser As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Me.Dock = DockStyle.Fill
        'Me.Width = Me.Parent.Width - 22
        Dim TabFrameTops As Integer

        fUserIsAdmin = False
        fMinabilityUser = ""

        gFormLoad = True
        gShowMainExit("Off")
        gWriteOk = WriteOK()

        'chkMyParams.Checked = True

        GetAllSections()
        GetAllTownships()
        GetAllRanges()
        GetAllHoles()
        GetPsizeDefns()
        GetRcvryEtc()
        GetAllMineNames()

        TabFrameTops = 4560

        opt100Pct.Checked = True

        lblHole.Text = ""
        lblSplit.Text = ""
        cmdSaveMinabilities.Enabled = False
        cmdSaveCompAndSplits.Enabled = False
        opt100PctRdctn.Enabled = False
        optCatalogRdctn.Enabled = False
        optBothRdctn.Enabled = False
        chkSaveRawProspectMinabilities.Enabled = False
        cboMineName.Enabled = False

        txtSurvCaddTextfile.Enabled = False
        optInclComposites.Checked = True
        optInclComposites.Enabled = False
        optInclSplits.Enabled = False
        optInclBoth.Enabled = False
        cmdCreateSurvCadd.Enabled = False

        cmdReduceHole.Enabled = False

        opt100Pct.Enabled = False
        optCatalog.Enabled = False
        cmdPrintHole.Enabled = False
        cmdPrintSplit.Enabled = False
        cmdMakeHoleUnmineable.Enabled = False

        'If gUserName.ToUpper <> "SSIEBER" Then
        cmdTest.Visible = False
        cmdTest2.Visible = False
        cmdTest3.Visible = False
        'End If

        ssSplitReview.MaxRows = 0
        ssCompReview.MaxRows = 0
        lblCoordsElev.Text = ""
        lblOvbComm.Text = ""
        lblMaxDepthComm.Text = ""
        lblMiscComm.Text = ""
        lblMiscComm2.Text = ""
        lblCurrSplit.Text = ""
        lblCurrSplit.Visible = False

        lblHoleExistStatus.Text = ""
        ssHoleExistStatus.MaxRows = 0
        lblAlphaNumeric.Text = ""

        lblInfoComm.Text = "1)  The hole should have X-Coordinate, Y-Coordinate and Elevation values in the raw prospect data." &
                          vbCrLf &
                          "2)  The hole should be 'released by the MetLab' in the raw prospect data." &
                          vbCrLf &
                          "3)  The hole should be 'available for use in developing composite/split " &
                          "prospect sets' in the raw prospect data." &
                          vbCrLf &
                          "4)  The 'Rcvry,Prod Adj,Prod Qual, Minability Scenario' that has been " &
                          "selected should have its '100% Flotation Recovery Method' set " &
                          "to " & vbCrLf &
                          "      'Tail BPL = SqrRt(Feed BPL)'." &
                          vbCrLf &
                          "5)  You must select a mine."


        FixSpreads()
        ClearData()

        'fraOutputOptions.BorderStyle = 0
        txtSurvCaddTextfile.Text = gGetProspDatasetTfileLoc(gUserName)

        lblOffSpecPbMgPlt.Text = ""
        ssCompErrors.MaxRows = 0
        lblUserMadeHoleUnmineable.Text = ""

        lblHoleInMoisComm.Text = ""
        lblHoleInMoisComm.Visible = True
        txtAreaName.Text = ""
        lblScenComm.Text = ""
        lblScenComm.Visible = True
        chkOverrideMaxDepth.Checked = False
        chkPbAnalysisFillInSpecial.Checked = False

        cmdSaveMinabilities.Text = "Save Minabilities Only" & vbCrLf & "(to Raw Prospect)"
        cmdSaveCompAndSplits.Text = "Save Reduced" & vbCrLf & "Composites && Splits"

        chkUseOrigHole.Checked = False
        'lblUseOrigHoleComm.Text = "(Works only if hole redrilled once!)"
        lblUseOrigHoleComm.Text = "*Gets most recent redrilled hole if more than one redrill exists!"

        chkUseFeAdjust.Checked = False

        ssFeAdjustment.Row = 1
        ssFeAdjustment.Col = 1
        ssFeAdjustment.Value = 90
        ssFeAdjustment.Row = 2
        ssFeAdjustment.Col = 1
        ssFeAdjustment.Value = 60

        '09/09/2009, lss
        'User does not need access to chkUseFeAdjust and ssFeAdjustment anymore.
        chkUseFeAdjust.Enabled = False
        ssFeAdjustment.Enabled = False

        lblFeAdjComm.Text = "Fe2O3 adjustment parameters can now be saved in the " &
                               "'Recovery, Prod Adj, Prod Qual, Minability, Density, " &
                               "100% Defn, Off-spec Pb' scenarios on the Multiple " &
                               "hole raw prospect data reduction form."

        optBothRdctn.Checked = True
        chkSaveRawProspectMinabilities.Checked = True
        lblSaveToMoisComm.Text = "(100% Prospect && Catalog)"

        '01/05/2011, lss
        'Override Stuff
        txtSplitOverrideName.Text = ""
        chkOnlyMySplitOverride.Checked = False
        ssSplitOverrides.MaxRows = 0
        ssSplitOverride.MaxRows = 0
        GetSplitOverrideSets()

        gFormLoad = False
    End Sub

    Private Sub SetMoisHoleExist()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim HoleDynaset As OraDynaset
        Dim HoleDesc As String
        Dim ThisSec As Integer
        Dim ThisTwp As Integer
        Dim ThisRge As Integer
        Dim ThisHole As String
        Dim ThisProspDate As String
        Dim ThisNumSplits As Integer
        Dim ThisMinableSplits As String
        Dim ThisMineName As String
        Dim ThisProspStandard As String
        Dim LblSet As Boolean

        ssHoleExistStatus.MaxRows = 0

        'This hole location will be numeric -- gGetHoleMoisExist will
        'check for both numeric and alpha-numeric.

        HoleDesc = gGetHoleLocationTrs(cboSec.Text,
                                       cboTwp.Text,
                                       cboRge.Text,
                                       cboHole.Text)

        gGetHoleMoisExist(Val(cboTwp.Text),
                          Val(cboRge.Text),
                          Val(cboSec.Text),
                          cboHole.Text,
                          HoleDynaset)

        LblSet = False

        If HoleDynaset.RecordCount <> 0 Then
            HoleDynaset.MoveFirst()
            Do While Not HoleDynaset.EOF
                ThisTwp = HoleDynaset.Fields("township").Value
                ThisRge = HoleDynaset.Fields("range").Value
                ThisSec = HoleDynaset.Fields("section").Value
                ThisHole = HoleDynaset.Fields("hole_location").Value
                ThisProspDate = HoleDynaset.Fields("drill_cdate").Value
                ThisNumSplits = HoleDynaset.Fields("split_total_num").Value

                'Display the data in lblHoleInMoisComm for the 1st record only.
                If LblSet = False Then
                    lblHoleInMoisComm.Text = "Already in MOIS, " &
                                                ThisProspDate & ", " & CStr(ThisNumSplits) &
                                                " split(s)"
                    LblSet = True
                End If

                If Not IsDBNull(HoleDynaset.Fields("split_sum").Value) Then
                    ThisMinableSplits = HoleDynaset.Fields("split_sum").Value
                Else
                    ThisMinableSplits = ""
                End If

                ThisMineName = HoleDynaset.Fields("mine_name").Value
                ThisProspStandard = HoleDynaset.Fields("prosp_standard").Value

                With ssHoleExistStatus
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = 1
                    .Text = ThisMineName
                    .Col = 2
                    .Text = gGetHoleLocationTrs(ThisSec, ThisTwp,
                                                ThisRge, ThisHole)
                    .Col = 3
                    .Text = ThisProspDate
                    .Col = 4
                    .Value = ThisNumSplits
                    .Col = 5
                    .Text = ThisMinableSplits
                    .Col = 6
                    .Text = ThisProspStandard
                End With

                HoleDynaset.MoveNext()
            Loop
            lblHoleExistStatus.Text = "This hole exists in MOIS (" &
                                         HoleDesc & ")."
        Else
            lblHoleExistStatus.Text = "This hole does not exist in MOIS (" &
                                         HoleDesc & ")."
        End If

        HoleDynaset.Close()
    End Sub

    Private Sub ClearData()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SprdIdx As Integer
        Dim ThisSprd As AxvaSpread = Nothing

        ssDrillData.MaxRows = 0
        lblCoordsElev.Text = ""
        lblOvbComm.Text = ""
        lblMaxDepthComm.Text = ""
        lblMiscComm.Text = ""
        lblMiscComm2.Text = ""
        lblHole.Text = ""
        lblSplit.Text = ""
        lblRdctnSplit.Text = ""
        lblOffSpecPbMgPlt.Text = ""

        For SprdIdx = 1 To 2
            If SprdIdx = 1 Then
                ThisSprd = ssHoleData
            ElseIf SprdIdx = 2 Then
                ThisSprd = ssSplitData
            End If

            With ThisSprd
                .BlockMode = True
                .Row = 1
                .Row2 = .MaxRows
                .Col = 1
                .Col2 = 1
                .Text = ""

                .Row = 1
                .Row2 = .MaxRows
                .Col = 2
                .Col2 = 10
                .Value = 0

                .BlockMode = False

                .Col = 14   'Wt%
                .Row = 2
                .Value = 0
                .Row = 3
                .Value = 0
                .Row = 5
                .Value = 0
                .Row = 6
                .Value = 0
                .Row = 7
                .Value = 0

                .Col = 15   'TPA
                .Row = 2
                .Value = 0
                .Row = 3
                .Value = 0
                .Row = 5
                .Value = 0
                .Row = 6
                .Value = 0
                .Row = 7
                .Value = 0

                .Col = 16   'BPL
                .Row = 2
                .Value = 0
                .Row = 3
                .Value = 0
                .Row = 5
                .Value = 0
                .Row = 6
                .Value = 0
                .Row = 7
                .Value = 0

                .Col = 18
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

                .Col = 20
                .Row = 4
                .Value = 0
                .Row = 5
                .Value = 0
                .Row = 6
                .Text = ""
            End With
        Next SprdIdx

        ssCompErrors.MaxRows = 0

        ssSplitReview.MaxRows = 0
        ssCompReview.MaxRows = 0

        cmdSaveMinabilities.Enabled = False
        cmdSaveCompAndSplits.Enabled = False
        opt100PctRdctn.Enabled = False
        optCatalogRdctn.Enabled = False
        optBothRdctn.Enabled = False
        chkSaveRawProspectMinabilities.Enabled = False
        cboMineName.Enabled = False
        opt100Pct.Enabled = False
        optCatalog.Enabled = False
        cmdPrintHole.Enabled = False
        cmdPrintSplit.Enabled = False
        cmdMakeHoleUnmineable.Enabled = False

        txtSurvCaddTextfile.Enabled = False
        optInclComposites.Checked = True
        optInclComposites.Enabled = False
        optInclSplits.Enabled = False
        optInclBoth.Enabled = False
        cmdCreateSurvCadd.Enabled = False

        lblHoleInMoisComm.Text = ""
        lblHoleInMoisComm.Visible = True

        lblScenComm.Text = ""
        lblScenComm.Visible = True

        ssSplitMinabilities.MaxRows = 0

        With ssHoleMinabilities
            .Row = 1
            .Col = 1
            .Text = ""
            .Row = 2
            .Text = ""
            .Row = 3
            .Text = ""
        End With

        lblCurrMinabilityComm.Text = ""
    End Sub

    Private Sub FixSpreads()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim DivIdx As Integer
        Dim ColNum As Integer
        Dim RowIdx As Integer

        With ssHoleData
            .BlockMode = True
            .Row = 1
            .Row2 = .MaxRows
            .Col = 0
            .Col2 = 0
            .TypeTextWordWrap = False
            .TypeHAlign = TypeHAlignConstants.TypeHAlignLeft ' SS_CELL_H_ALIGN_LEFT
            .BlockMode = False
        End With
        With ssSplitData
            .BlockMode = True
            .Row = 1
            .Row2 = .MaxRows
            .Col = 0
            .Col2 = 0
            .TypeTextWordWrap = False
            .TypeHAlign = TypeHAlignConstants.TypeHAlignLeft ' SS_CELL_H_ALIGN_LEFT
            .BlockMode = False
        End With
        With ssHoleMinabilities
            .BlockMode = True
            .Row = 1
            .Row2 = .MaxRows
            .Col = 0
            .Col2 = 0
            .TypeTextWordWrap = False
            .TypeHAlign = TypeHAlignConstants.TypeHAlignLeft 'SS_CELL_H_ALIGN_LEFT
            .BlockMode = False
        End With

        'Add vertical dividers
        For DivIdx = 1 To 9
            Select Case DivIdx
                Case Is = 1
                    ColNum = 4
                Case Is = 2
                    ColNum = 12
                Case Is = 3
                    ColNum = 16
                Case Is = 4
                    ColNum = 24
                Case Is = 5
                    ColNum = 32
                Case Is = 6
                    ColNum = 35
                Case Is = 7
                    ColNum = 43
                Case Is = 8
                    ColNum = 51
                Case Is = 9
                    ColNum = 60
            End Select

            With ssDrillData
                .Row = 0
                .Col = ColNum
                .Text = " "
                .set_ColWidth(.Col, 0.17)
                For RowIdx = 0 To .MaxRows
                    .Row = RowIdx
                    .CellType = CellTypeConstants.CellTypeStaticText ' SS_CELL_TYPE_STATIC_TEXT
                    .Text = " "
                    .TypeTextShadow = False
                    .BackColor = Color.Black 'vbBlack
                Next
            End With
        Next DivIdx
    End Sub

    Private Sub cmdExitForm_Click()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        If MsgBox("Exit raw prospect data reduction program?", vbYesNo +
            vbDefaultButton1, "Exiting Program") = vbYes Then
            'Unload(Me)
            Me.Close()
        End If
    End Sub

    Private Sub Form_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.FormClosed 'ByVal Cancel As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        gShowMainExit("On")

        On Error Resume Next
    End Sub

    Private Function WriteOK()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'In order to add/modify data on this form gUserPermission.RawProspectReduction must
        'be "Write" or "Setup" or "Admin".

        WriteOK = False
        fUserIsAdmin = False

        If gUserPermission.RawProspectReduction = "Write" Or
            gUserPermission.RawProspectReduction = "Setup" Or
            gUserPermission.RawProspectReduction = "Admin" Then
            WriteOK = True
        End If

        If gUserPermission.RawProspectReduction = "Admin" Then
            fUserIsAdmin = True
        Else
            fUserIsAdmin = False
        End If
    End Function

    Private Sub cmdPrtScr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrtScr.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim UserResponse As Object

        UserResponse = MsgBox("Print the screen?", vbOKCancel, "Printing")

        If UserResponse = vbOK Then
            On Error GoTo PrintError

            SetActionStatus("Printing the screen...")
            'Me.Cursor = Cursors.WaitCursor
            Me.Cursor = Cursors.WaitCursor
            '  Try
            gPrintScreen(Me.Handle)
            'Catch ex As Exception
            '    MessageBox.Show(ex.Message)
            'End Try
            'PrepareViewForPrint()
            'Me.Refresh()
            ''Picture1.Picture = CaptureClient(Me)
            ''PrintPictureToFitPage(Printer, Picture1.Picture)
            ''Printer.EndDoc()
            'ResetViewAfterPrint()

            SetActionStatus("")
            Me.Cursor = Cursors.Arrow
            'Me.Cursor = Cursors.Arrow
        End If

        Exit Sub

PrintError:
        Me.Cursor = Cursors.Arrow
        MsgBox("Error printing screen." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Screen Printing Error")

    End Sub

    Private Sub PrepareViewForPrint()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Will leave things alone on this form when printing form!
    End Sub

    Private Sub ResetViewAfterPrint()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Will leave things alone on this form when printing form!
    End Sub

    Private Sub SetActionStatus(ByVal aStatus As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error Resume Next
        sbrMain.Text = aStatus
    End Sub

    Private Sub GetAllSections()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetAllSectionsError

        Dim ChoiceDynaset As OraDynaset
        Dim ChoicesOk As Boolean
        Dim ChoiceStr As String

        'Section number choices will be the same regardless of the mine.
        ChoicesOk = gGetChoices(ChoiceDynaset, gActiveMineNameLong,
                                "Section number", True)

        ChoiceStr = " "
        cboSec.Items.Add("(Select...)")

        ChoiceDynaset.MoveFirst()
        Do While Not ChoiceDynaset.EOF
            cboSec.Items.Add(ChoiceDynaset.Fields("combo_box_choice_text").Value)
            ChoiceStr = ChoiceStr + Chr(9) +
                        ChoiceDynaset.Fields("combo_box_choice_text").Value
            ChoiceDynaset.MoveNext()
        Loop
        cboSec.Text = "(Select...)"

        Exit Sub

GetAllSectionsError:
        MsgBox("Error loading section numbers." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Section Number Loading Error")
    End Sub

    Private Sub GetAllTownships()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetAllTownshipsError

        Dim ChoiceDynaset As OraDynaset
        Dim ChoicesOk As Boolean
        Dim ChoiceStr As String

        ChoicesOk = gGetChoices(ChoiceDynaset, "Four Corners",
                                "Township all", True)

        ChoiceStr = " "
        cboTwp.Items.Add("(Select...)")

        ChoiceDynaset.MoveFirst()
        Do While Not ChoiceDynaset.EOF
            cboTwp.Items.Add(ChoiceDynaset.Fields("combo_box_choice_text").Value)
            ChoiceStr = ChoiceStr + Chr(9) +
                        ChoiceDynaset.Fields("combo_box_choice_text").Value
            ChoiceDynaset.MoveNext()
        Loop
        cboTwp.Text = "(Select...)"

        Exit Sub

GetAllTownshipsError:
        MsgBox("Error loading Township numbers." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Township Number Loading Error")
    End Sub

    Private Sub GetAllRanges()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetAllRangesError

        Dim ChoiceDynaset As OraDynaset
        Dim ChoicesOk As Boolean
        Dim ChoiceStr As String

        ChoicesOk = gGetChoices(ChoiceDynaset, "Four Corners",
                                "Range all", True)

        ChoiceStr = " "
        cboRge.Items.Add("(Select...)")

        ChoiceDynaset.MoveFirst()
        Do While Not ChoiceDynaset.EOF
            cboRge.Items.Add(ChoiceDynaset.Fields("combo_box_choice_text").Value)
            ChoiceStr = ChoiceStr + Chr(9) +
                        ChoiceDynaset.Fields("combo_box_choice_text").Value
            ChoiceDynaset.MoveNext()
        Loop
        cboRge.Text = "(Select...)"

        Exit Sub

GetAllRangesError:
        MsgBox("Error loading Range numbers." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Range Number Loading Error")
    End Sub

    Private Sub GetAllHoles()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetAllHolesError

        Dim ChoiceDynaset As OraDynaset
        Dim ChoicesOk As Boolean
        Dim ChoiceStr As String

        ChoicesOk = gGetChoices(ChoiceDynaset, gActiveMineNameLong,
                                "Hole locations", True)

        ChoiceStr = " "
        cboHole.Items.Add("(Select...)")

        ChoiceDynaset.MoveFirst()
        Do While Not ChoiceDynaset.EOF
            cboHole.Items.Add(ChoiceDynaset.Fields("combo_box_choice_text").Value)
            ChoiceStr = ChoiceStr + Chr(9) +
                        ChoiceDynaset.Fields("combo_box_choice_text").Value
            ChoiceDynaset.MoveNext()
        Loop
        cboHole.Text = "(Select...)"

        Exit Sub

GetAllHolesError:
        MsgBox("Error loading holes." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Hole Loading Error")
    End Sub

    Private Sub GetPsizeDefns()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetPsizeDefnsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim PsizeDefnDynaset As OraDynaset
        Dim UserName As String
        Dim ThisPsizeDefnName As String
        Dim ThisWhoDefnd As String
        Dim ThisWhenDefnd As Date
        Dim ThisMineName As String
        Dim RecordCount As Integer

        If chkMyParams.Checked = True Then
            UserName = gUserName.ToLower
            'UserName = "ssieber"
        Else
            UserName = "All"
        End If

        params = gDBParams

        params.Add("pUserName", UserName, ORAPARM_INPUT)
        params("pUserName").serverType = ORATYPE_VARCHAR2

        params.Add("pProspSetName", "User psize definition", ORAPARM_INPUT)
        params("pProspSetName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_psize_defn_all
        'pUserName           IN     VARCHAR2,
        'pProspSetName       IN     VARCHAR2,
        'pResult             IN OUT c_areadefn)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prosp_data_rdctn.get_prosp_psize_defn_all (" &
                                             ":pUserName, :pProspSetName, :pResult);end;", ORASQL_FAILEXEC)
        PsizeDefnDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = PsizeDefnDynaset.RecordCount

        cboProdSizeDefn.Items.Clear()
        cboProdSizeDefn.Items.Add("(Select...)")

        PsizeDefnDynaset.MoveFirst()
        Do While Not PsizeDefnDynaset.EOF
            ThisPsizeDefnName = PsizeDefnDynaset.Fields("Psize_defn_name").Value
            ThisWhoDefnd = PsizeDefnDynaset.Fields("who_defined").Value
            ThisWhenDefnd = PsizeDefnDynaset.Fields("when_defined").Value

            'The mine name can be Null in PROSP_PSIZE_DEFN_BASE.
            If Not IsDBNull(PsizeDefnDynaset.Fields("mine_name").Value) Then
                ThisMineName = PsizeDefnDynaset.Fields("mine_name").Value
            Else
                ThisMineName = ""
            End If

            cboProdSizeDefn.Items.Add(ThisPsizeDefnName)
            PsizeDefnDynaset.MoveNext()
        Loop
        cboProdSizeDefn.Text = "(Select...)"

        PsizeDefnDynaset.Close()

        Exit Sub

GetPsizeDefnsError:
        On Error Resume Next

        MsgBox("Error getting product size designations." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Product Size Designations Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        PsizeDefnDynaset.Close()
    End Sub

    Private Sub GetRcvryEtc()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetRcvryEtcError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RcvryDynaset As OraDynaset
        Dim UserName As String
        Dim ThisRcvryScenarioName As String
        Dim ThisWhoDefnd As String
        Dim ThisWhenDefnd As Date
        Dim ThisMineName As String
        Dim RecordCount As Integer

        If chkMyParams.Checked = True Then
            UserName = gUserName.ToLower
            'UserName = "ssieber"
        Else
            UserName = "All"
        End If

        params = gDBParams

        params.Add("pUserName", UserName, ORAPARM_INPUT)
        params("pUserName").serverType = ORATYPE_VARCHAR2

        params.Add("pProspSetName", "User recovery scenario", ORAPARM_INPUT)
        params("pProspSetName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_rcvry_base_all
        'pUserName           IN     VARCHAR2,
        'pProspSetName       IN     VARCHAR2,
        'pResult             IN OUT c_rcvry)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prosp_data_rdctn.get_prosp_rcvry_base_all (" &
                                             ":pUserName, :pProspSetName, :pResult);end;", ORASQL_FAILEXEC)
        RcvryDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = RcvryDynaset.RecordCount

        cboOtherDefn.Items.Clear()
        cboOtherDefn.Items.Add("(Select...)")

        RcvryDynaset.MoveFirst()
        Do While Not RcvryDynaset.EOF
            ThisRcvryScenarioName = RcvryDynaset.Fields("rcvry_scenario_name").Value
            ThisWhoDefnd = RcvryDynaset.Fields("who_defined").Value
            ThisWhenDefnd = RcvryDynaset.Fields("when_defined").Value

            If Not IsDBNull(RcvryDynaset.Fields("mine_name").Value) Then
                ThisMineName = RcvryDynaset.Fields("mine_name").Value
            Else
                ThisMineName = ""
            End If

            cboOtherDefn.Items.Add(ThisRcvryScenarioName)

            RcvryDynaset.MoveNext()
        Loop
        cboOtherDefn.Text = "(Select...)"

        RcvryDynaset.Close()

        Exit Sub

GetRcvryEtcError:
        On Error Resume Next

        MsgBox("Error getting recovery scenarios." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Recovery Scenarios Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        RcvryDynaset.Close()
    End Sub

    Private Sub cmdRefreshParams_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefreshParams.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        GetPsizeDefns()
        GetRcvryEtc()
    End Sub

    Private Sub cboTwp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTwp.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SetReduceEnable()
        ClearData()
    End Sub

    Private Sub cboRge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRge.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SetReduceEnable()
        ClearData()
    End Sub

    Private Sub cboSec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboSec.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SetReduceEnable()
        ClearData()
    End Sub

    Private Sub cboHole_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboHole.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        If cboHole.Text <> "(Select...)" Then
            If gGetHoleLoc2(cboHole.Text, "Char") <> "???" Then
                lblAlphaNumeric.Text = gGetHoleLoc2(cboHole.Text, "Char")
            Else
                lblAlphaNumeric.Text = ""
            End If
        End If

        SetReduceEnable()
        ClearData()
    End Sub

    Private Sub chkOverrideMaxDepth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOverrideMaxDepth.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SetReduceEnable()
        ClearData()
    End Sub

    Private Sub chkUseOrigHole_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseOrigHole.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SetReduceEnable()
        ClearData()
    End Sub

    Private Sub cboHole_Validate(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboHole.Validating 'ByVal Cancel As Boolean)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim InputOk As Boolean

        InputOk = True

        'Should be in the format "####" -- 0434, 1434, etc -- length should be 3 or 4.
        If IsNumeric(cboHole.Text) And Len(Trim(cboHole.Text)) = 3 Then
            cboHole.Text = "0" & cboHole.Text
        End If
        If IsNumeric(cboHole.Text) = False Then
            InputOk = False
        End If
        If Len(Trim(cboHole.Text)) <> 3 And Len(Trim(cboHole.Text)) <> 4 Then
            InputOk = False
        End If
        If Val(Mid(cboHole.Text, 1, 2)) < 1 Or Val(Mid(cboHole.Text, 1, 2)) > 33 Or
             Val(Mid(cboHole.Text, 3, 2)) < 33 Or Val(Mid(cboHole.Text, 3, 2)) > 72 Then
            InputOk = False
        End If
        If InputOk = False Then
            MsgBox("Illegal numeric hole location!", vbOKOnly, "Illegal Input")
            cboHole.Text = "(Select...)"
        End If
    End Sub

    Private Sub cboProdSizeDefn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboProdSizeDefn.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SetReduceEnable()
        ClearData()
    End Sub

    Private Sub cboOtherDefn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboOtherDefn.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SetReduceEnable()
        ClearData()
    End Sub

    Private Sub SetReduceEnable()
        Try
            If cboTwp.Text <> "(Select...)" And cboRge.Text <> "(Select...)" And
                cboSec.Text <> "(Select...)" And cboHole.Text <> "(Select...)" And
                cboProdSizeDefn.Text <> "(Select...)" And cboOtherDefn.Text <> "(Select...)" Then
                cmdReduceHole.Enabled = True
            Else
                cmdReduceHole.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub tabDisp_Click()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Select Case tabDisp.SelectedItem.Index
        '    Case Is = 1     'Hole composite
        '        fraHole.Visible = True
        '        fraSplit.Visible = False
        '        fraRdctnData.Visible = False
        '        fraMiscDisp.Visible = False
        '        fraInfo.Visible = False
        '        fraAddArea.Visible = False
        '        fraCompErrors.Visible = False
        '        fraCurrMinabilities.Visible = False
        '        fraMisc.Visible = False
        '        fraAddToOverride.Visible = False

        '    Case Is = 2     'Split
        '        fraHole.Visible = False
        '        fraSplit.Visible = True
        '        fraRdctnData.Visible = False
        '        fraMiscDisp.Visible = False
        '        fraInfo.Visible = False
        '        fraAddArea.Visible = False
        '        fraCompErrors.Visible = False
        '        fraCurrMinabilities.Visible = False
        '        fraMisc.Visible = False
        '        fraAddToOverride.Visible = False

        '    Case Is = 3     'Reduction data
        '        fraHole.Visible = False
        '        fraSplit.Visible = False
        '        fraRdctnData.Visible = True
        '        fraMiscDisp.Visible = False
        '        fraInfo.Visible = False
        '        fraAddArea.Visible = False
        '        fraCompErrors.Visible = False
        '        fraCurrMinabilities.Visible = False
        '        fraMisc.Visible = False
        '        fraAddToOverride.Visible = False

        '    Case Is = 4     'Miscellaneous display
        '        fraHole.Visible = False
        '        fraSplit.Visible = False
        '        fraRdctnData.Visible = False
        '        fraMiscDisp.Visible = True
        '        fraInfo.Visible = False
        '        fraAddArea.Visible = False
        '        fraCompErrors.Visible = False
        '        fraCurrMinabilities.Visible = False
        '        fraMisc.Visible = False
        '        fraAddToOverride.Visible = False

        '    Case Is = 5     'Info
        '        fraHole.Visible = False
        '        fraSplit.Visible = False
        '        fraRdctnData.Visible = False
        '        fraMiscDisp.Visible = False
        '        fraInfo.Visible = True
        '        fraAddArea.Visible = False
        '        fraCompErrors.Visible = False
        '        fraCurrMinabilities.Visible = False
        '        fraMisc.Visible = False
        '        fraAddToOverride.Visible = False

        '    Case Is = 6     'Comp errors
        '        fraHole.Visible = False
        '        fraSplit.Visible = False
        '        fraRdctnData.Visible = False
        '        fraMiscDisp.Visible = False
        '        fraInfo.Visible = False
        '        fraAddArea.Visible = False
        '        fraCompErrors.Visible = True
        '        fraCurrMinabilities.Visible = False
        '        fraMisc.Visible = False
        '        fraAddToOverride.Visible = False

        '    Case Is = 7     'Add areas
        '        fraHole.Visible = False
        '        fraSplit.Visible = False
        '        fraRdctnData.Visible = False
        '        fraMiscDisp.Visible = False
        '        fraInfo.Visible = False
        '        fraAddArea.Visible = True
        '        fraCompErrors.Visible = False
        '        fraCurrMinabilities.Visible = False
        '        fraMisc.Visible = False
        '        fraAddToOverride.Visible = False

        '    Case Is = 8    'Current minabilities (in raw prospect)
        '        fraHole.Visible = False
        '        fraSplit.Visible = False
        '        fraRdctnData.Visible = False
        '        fraMiscDisp.Visible = False
        '        fraInfo.Visible = False
        '        fraAddArea.Visible = False
        '        fraCompErrors.Visible = False
        '        fraCurrMinabilities.Visible = True
        '        fraMisc.Visible = False
        '        fraAddToOverride.Visible = False

        '    Case Is = 9    'Miscellaneous
        '        fraHole.Visible = False
        '        fraSplit.Visible = False
        '        fraRdctnData.Visible = False
        '        fraMiscDisp.Visible = False
        '        fraInfo.Visible = False
        '        fraAddArea.Visible = False
        '        fraCompErrors.Visible = False
        '        fraCurrMinabilities.Visible = False
        '        fraMisc.Visible = True
        '        fraAddToOverride.Visible = False

        '    Case Is = 10   'Add to Override
        '        fraHole.Visible = False
        '        fraSplit.Visible = False
        '        fraRdctnData.Visible = False
        '        fraMiscDisp.Visible = False
        '        fraInfo.Visible = False
        '        fraAddArea.Visible = False
        '        fraCompErrors.Visible = False
        '        fraCurrMinabilities.Visible = False
        '        fraMisc.Visible = False
        '        fraAddToOverride.Visible = True
        'End Select
    End Sub

    Private Sub cmdReduceHole_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReduceHole.Click

        Dim RcvryData As gDataRdctnParamsType
        Dim RcvryProdQual(0 To 14) As gDataRdctnProdQualType
        'Dim DataSetStatus As Integer
        Dim ProspectResult As SplitResultSet
        Dim AreaDefnData As New ViewModels.ProspectAreaDefinition
        Dim SfcReproData(0 To 87, 0 To 22) As String
        Dim SplitDatesCorrect As Boolean
        Dim ProspData As gRawProspSplRdctnType
        Dim UseOrigHole As Boolean

        'I am going to piggy-back this on top of the functionality in
        'frmProspDataReduction.  It's a little hodge-podged but it
        'will work for now.

        If chkUseOrigHole.Checked = True Then
            UseOrigHole = True
        Else
            UseOrigHole = False
        End If

        SetActionStatus("Reducing hole...")
        Me.Cursor = Cursors.WaitCursor

        fReducing = True

        ClearData()

        SetMoisHoleExist()
        SetMoisCurrRawMinabilities()

        'Set area definition data
        AreaDefnData.ByHolesAreaMethod = True
        AreaDefnData.BeginningDrillDate = #12/31/8888#
        AreaDefnData.EndDrillDate = #12/31/8888#
        AreaDefnData.HoleMetLabProcessType = Nothing
        AreaDefnData.MinedStatus = Nothing
        AreaDefnData.MineName = Nothing
        AreaDefnData.Ownership = Nothing
        AreaDefnData.Holes.Add(New ViewModels.ProspectAreaHole With {.Hole_Township = cboTwp.Text,
                                                                     .Hole_Range = cboRge.Text,
                                                                     .Hole_Section = cboSec.Text,
                                                                     .Hole_Location = cboHole.Text})

        Dim ProductSizeDesignation As ViewModels.ProductSizeDesignation = GetProductSizeDistribution(cboProdSizeDefn.Text)
        Dim RecoveryDef As ViewModels.ProductRecoveryDefinition = GetRecoveryDefinition(cboOtherDefn.Text)


        'Set recovery, etc. information
        SetRcvryEtc(cboOtherDefn.Text,
                    "User recovery scenario",
                    RcvryData,
                    RcvryProdQual)

        '06/15/2010, lss  RcvryData.MineHasOffSpecPbPlt not used anymore.
        'If RcvryData.MineHasOffSpecPbPlt = True Then
        If RcvryData.UseOrigMgoPlant = True Then
            lblOffSpecPbMgPlt.Text = "*OffSpec Pb Mg Plt*"
        Else
            If RcvryData.UseDoloflotPlant2010 = True Then
                lblOffSpecPbMgPlt.Text = "*Doloflot Plt Ona*"
            Else
                If RcvryData.UseDoloflotPlantFco = True Then
                    lblOffSpecPbMgPlt.Text = "*Doloflot Plt FCO*"
                Else
                    lblOffSpecPbMgPlt.Text = ""
                End If
            End If
        End If

        lblScenComm.Text = RcvryData.MaxTotDepthModeSpl &
                              ", Max tot depth = " & Format(RcvryData.MaxTotDepthSpl, "#,##0.0")

        If chkOverrideMaxDepth.Checked = True Then
            'Just make it a real big number!
            RcvryData.MaxTotDepthSpl = 9999
        End If

        If chkUseFeAdjust.Checked = True Then
            RcvryData.UseFeAdjust = True
        Else
            RcvryData.UseFeAdjust = False
        End If

        With ssFeAdjustment
            .Row = 1
            .Col = 1
            RcvryData.UpperZoneFeAdjust = .Value
            .Row = 2
            .Col = 1
            RcvryData.LowerZoneFeAdjust = .Value
        End With

        'frmProspDataHoleReduction.Refresh()

        'Get prospect raw material size definition (% distribution of SFC's)
        'Will hard-code this for now -- there is really only one defined
        'distribution that is used that I have called Standard2006.  It has
        '85 size fraction codes that are distributed by percent among 21 size
        'fraction codes.  I will save numbers as strings.  Added 2 extra rows
        'and 1 extra column for header stuff.
        Dim StandardSFCName As String = String.Empty
        StandardSFCName = ProductSizeDesignation.SizeFractionDistribution
        If StandardSFCName.Equals(String.Empty) Then StandardSFCName = "Standard2006"
        GetProspRawMatlSizeDefn(StandardSFCName,
                                SfcReproData)

        'Process raw prospect data -- place data in ssSplitReview and
        'ssCompReview.
        gHaveRawProspData = False

        'Adjusting Fe2O3 not available thru here!!!
        RcvryData.UseFeAdjust = False
        RcvryData.UpperZoneFeAdjust = 90
        RcvryData.LowerZoneFeAdjust = 60

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
        '10) aSplitOverrideName As String, _                     ""
        '11) aSsRawProspMin As vaSpread, _                       ssRawProspMin
        '12) aSfcReproData() As String, _                        SfcReproData()
        '13) aRawProspDynaset As OraDynaset, _                   RawProspDynaset
        '14) aScope As String, _                                 "Hole"
        '15) aNoReview As Boolean, _                             False
        '16) aSaveType As String, _                              ""
        '17) aMineHasOffSpecPbPlt As Boolean, _                  False
        '18) aProspectDatasetName As String, _                   ""
        '19) aProspDatasetTextFileName As String, _              ""
        '20) aChk100Pct As Integer, _                            0
        '21) aChkProductionCoefficient As Integer, _             0
        '22) aOptInclSplits As Boolean, _                        False
        '23) aOptInclComposites As Boolean,                      False
        '24) aOptInclBoth As Boolean, _                          False
        '25) aChkInclMgPlt As Integer, _                         0
        '26) aUseOrigHole As Boolean, _                          UseOrigHole
        '27) aMineHasDoloflotPlt As Boolean)                     False



        ProspectResult = gGenerateProspectDataset(AreaDefnData,
                                                 ProductSizeDesignation,
                                                 RcvryData, RecoveryDef,
                                                 ssSplitReview,
                                                 ssCompReview, "", ssRawProspMin,
                                                 SfcReproData,
                                                 "Hole", False,
                                                 "", False,
                                                 "", "", 0, 0, False, False, False, 0, UseOrigHole, False, cboProdSizeDefn.Text)


        If ProspectResult.IntResult = 1 Then
            ProcessReducedData(True)
        End If

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow

        Select Case ProspectResult.IntResult
            Case Is = 1
                lblMaxDepthComm.Text = GetMaxDepthInfo(RcvryData)

                'Check coordinates and elevation!
                ProspData = gGetDataFromReviewSprd(ssCompReview, 1)

                If lblOffSpecPbMgPlt.Text <> "*OffSpec Pb Mg Plt*" Then
                    If UserCanSaveThisToMois() = True Then
                        cmdSaveMinabilities.Enabled = True
                        cmdSaveCompAndSplits.Enabled = True
                        cmdCreateSurvCadd.Enabled = True
                        opt100PctRdctn.Enabled = True
                        optCatalogRdctn.Enabled = True
                        optBothRdctn.Enabled = True
                        chkSaveRawProspectMinabilities.Enabled = True
                        lblCurrMinabilityComm.Text = ""
                        cmdAddToOverrideSet.Enabled = True
                        cmdSaveAreaName.Enabled = True
                    Else
                        cmdSaveMinabilities.Enabled = False
                        cmdSaveCompAndSplits.Enabled = False
                        cmdCreateSurvCadd.Enabled = False
                        opt100PctRdctn.Enabled = False
                        optCatalogRdctn.Enabled = False
                        optBothRdctn.Enabled = False
                        lblCurrMinabilityComm.Text = "You do not have permission to save raw prospect minabilities " &
                                                        "for this hole or to save this hole to Pros1."
                        cmdAddToOverrideSet.Enabled = False
                        cmdSaveAreaName.Enabled = False
                    End If
                Else
                    '01/17/2008, lss
                    'MOIS will not handle the Off-spec pebble MgO plant stuff at this time!
                    cmdSaveMinabilities.Enabled = False
                    cmdSaveCompAndSplits.Enabled = False
                    cmdCreateSurvCadd.Enabled = False
                    opt100PctRdctn.Enabled = False
                    optCatalogRdctn.Enabled = False
                    optBothRdctn.Enabled = False
                    chkSaveRawProspectMinabilities.Enabled = False
                    lblCurrMinabilityComm.Text = "Saving is not allowed for '*OffSpec Pb Mg Plt*' holes " &
                                                    "at this time."
                    cmdAddToOverrideSet.Enabled = False
                    cmdSaveAreaName.Enabled = False
                End If

                If ProspData.Xcoord <> 0 And ProspData.Ycoord <> 0 And ProspData.Elevation <> 0 Then
                    MsgBox("Hole reduction completed -- no problems!", vbOKOnly, "Create Status")
                Else
                    MsgBox("Hole reduction reduced OK but this hole is missing hole coordinates " &
                           "and/or a hole elevation!!" & vbCrLf & vbCrLf &
                           "You really should not save this to the MOIS Composite/Split data! ",
                           vbOKOnly + vbExclamation, "Create Status")
                End If

                cboMineName.Enabled = True
                opt100Pct.Enabled = True
                optCatalog.Enabled = True
                cmdPrintHole.Enabled = True
                cmdPrintSplit.Enabled = True
                cmdMakeHoleUnmineable.Enabled = True
                SetMineName()

                txtSurvCaddTextfile.Enabled = True
                optInclComposites.Enabled = True

                '01/17/2008, lss -- not ready yet
                '01/12/2009, lss -- read now
                optInclSplits.Enabled = True
                optInclBoth.Enabled = False

                cmdCreateSurvCadd.Enabled = True
            Case Else
                MsgBox("Hole reduction completed -- PROBLEMS (Hole may not exist)!", vbOKOnly + vbExclamation, "Create Status")
                cmdSaveMinabilities.Enabled = False
                cmdSaveCompAndSplits.Enabled = False
                opt100PctRdctn.Enabled = False
                optCatalogRdctn.Enabled = False
                optBothRdctn.Enabled = False
                chkSaveRawProspectMinabilities.Enabled = False
                cboMineName.Enabled = False
                opt100Pct.Enabled = False
                optCatalog.Enabled = False
                cmdPrintHole.Enabled = False
                cmdPrintSplit.Enabled = False
                cmdMakeHoleUnmineable.Enabled = False

                txtSurvCaddTextfile.Enabled = False
                optInclComposites.Enabled = False
                optInclSplits.Enabled = False
                optInclBoth.Enabled = False
                cmdCreateSurvCadd.Enabled = False
        End Select

        gHaveRawProspData = False

        SplitDatesCorrect = GetSplitDatesCorrect()

        If SplitDatesCorrect = False Then
            MsgBox("Multiple split dates exist for this hole!  Old hole may not " &
                   "have been marked correctly as a redrill?" &
                   vbCrLf & vbCrLf &
                   "You may not do anything with this reduction data!", vbOKOnly,
                   "Prospect Hole Split Problem")

            cmdSaveMinabilities.Enabled = False
            cmdSaveCompAndSplits.Enabled = False
            opt100PctRdctn.Enabled = False
            optCatalogRdctn.Enabled = False
            optBothRdctn.Enabled = False
            chkSaveRawProspectMinabilities.Enabled = False
            cboMineName.Enabled = False
            opt100Pct.Enabled = False
            optCatalog.Enabled = False
            cmdPrintHole.Enabled = False
            cmdPrintSplit.Enabled = False
            cmdMakeHoleUnmineable.Enabled = False

            txtSurvCaddTextfile.Enabled = False
            optInclComposites.Checked = True
            optInclComposites.Enabled = False
            optInclSplits.Enabled = False
            optInclBoth.Enabled = False
            cmdCreateSurvCadd.Enabled = False
        End If

        lblUserMadeHoleUnmineable.Text = ""
        fReducing = False
    End Sub

    Private Sub ProcessReducedData(ByVal aDispProbs As Boolean)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ProspDate As Date
        Dim ReviewStatus As Boolean
        Dim Released As Integer
        Dim Redrilled As Integer
        Dim UseForReduction As Integer
        Dim QaQcHole As Integer

        With ssCompReview
            .Row = 1
            .Col = 3
            If .Text <> "None" Then
                Dim dteTest As DateTime
                If DateTime.TryParse(.Text, dteTest) Then
                    ProspDate = CDate(.Text)
                Else
                    ProspDate = #12/31/8888#
                End If
            Else
                ProspDate = #12/31/8888#
            End If
        End With

        lblHole.Text = GetHoleTitle()

        If opt100Pct.Checked = True Then
            lblHole.Text = lblHole.Text & "  " &
                              Format(ProspDate, "MM/dd/yyyy") &
                              "   (100% Prospect)    * = Offspec"
        Else    'Catalog
            lblHole.Text = lblHole.Text & "  " &
                              Format(ProspDate, "MM/dd/yyyy") &
                              "   (Catalog)    * = Offspec"
        End If

        'Will always display Split#1 to start.
        lblSplit.Text = GetHoleTitle() & "  Split# 1"
        lblCurrSplit.Text = "1"
        If opt100Pct.Checked = True Then
            lblSplit.Text = lblSplit.Text & " " &
                               Format(ProspDate, "MM/dd/yyyy") &
                               "   (100% Prospect)    * = Offspec"
        Else    'Catalog
            lblSplit.Text = lblSplit.Text & " " &
                               Format(ProspDate, "MM/dd/yyyy") &
                               "   (Catalog)    * = Offspec"
        End If

        lblRdctnHole.Text = "Hole  (" & GetHoleTitle() & ")"
        lblRdctnSplit.Text = "Split  (" & GetHoleTitle() & "  Split #1)"


        'Special data not available in ssCompReview or ssSplitReview
        ReviewStatus = gGetProspRawStatus(Val(cboTwp.Text),
                                          Val(cboRge.Text),
                                          Val(cboSec.Text),
                                          cboHole.Text,
                                          ProspDate,
                                          Redrilled,
                                          Released,
                                          UseForReduction,
                                          QaQcHole)
        lblMiscComm.Text = ""
        If ReviewStatus = True Then
            If Redrilled = 1 Then
                lblMiscComm.Text = lblMiscComm.Text & "Redrilled!  "
            End If
            If Released <> 1 Then
                lblMiscComm.Text = lblMiscComm.Text & "UnReleased!  "
            End If
            If UseForReduction <> 1 Then
                lblMiscComm.Text = lblMiscComm.Text & "No Reduction Use!  "
            End If
            If QaQcHole = 1 Then
                lblMiscComm.Text = lblMiscComm.Text & "QA/QC Hole!  "
            End If

            'Important that the user knows that the MetLab has not release this hole!
            If aDispProbs = True Then
                If Released <> 1 Then
                    MsgBox("WARNING:  This hole has not been released by the MetLab!!",
                           vbOKOnly, "Hole MetLab Release Status")
                End If
            End If
        Else
            lblMiscComm.Text = "Problems!!!"
        End If

        lblMiscComm.ForeColor = Color.DarkRed ' &HC0&     'Dark red
        'lblMiscComm.Font.Bol = True

        PopulateSsDrillDataEtc()
    End Sub


    Private Sub SetRcvryEtc(ByVal aRcvryScen As String,
                            ByVal aProspSetName As String,
                            ByRef aRcvryData As gDataRdctnParamsType,
                            ByRef aRcvryProdQual() As gDataRdctnProdQualType)


        On Error GoTo SetRcvryEtcError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer
        Dim TempData As gDataRdctnProdQualType
        Dim ProdQualDynaset As OraDynaset
        Dim ProdRow As Integer

        'Prospect recovery base data
        'Prospect recovery base data
        'Prospect recovery base data

        params = gDBParams

        params.Add("pRcvryScenarioName", aRcvryScen, ORAPARM_INPUT)
        params("pRcvryScenarioName").serverType = ORATYPE_VARCHAR2

        params.Add("pProspSetName", aProspSetName, ORAPARM_INPUT)
        params("pProspSetName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_rcvry_base
        'pRcvryScenarioName     IN     VARCHAR2,
        'pProspSetName          IN     VARCHAR2,
        'pResult                IN OUT c_rcvry);
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prosp_data_rdctn.get_rcvry_base(" +
                                       ":pRcvryScenarioName, :pProspSetName, :pResult);end;", ORASQL_FAILEXEC)

        fRcvryScenDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = fRcvryScenDynaset.RecordCount

        'Only one record will be returned!
        fRcvryScenDynaset.MoveFirst()

        With aRcvryData
            .RcvryScenarioName = fRcvryScenDynaset.Fields("rcvry_scenario_name").Value
            .ProspSetName = fRcvryScenDynaset.Fields("prosp_set_name").Value
            .WhoDefined = fRcvryScenDynaset.Fields("who_defined").Value
            .WhenDefined = fRcvryScenDynaset.Fields("when_defined").Value

            If Not IsDBNull(fRcvryScenDynaset.Fields("mine_name").Value) Then
                .MineName = fRcvryScenDynaset.Fields("mine_name").Value
            Else
                .MineName = "None"
            End If
            '-----
            '.OvbVolRcvryMode = fRcvryScenDynaset.Fields("ovb_vol_rcvry_mode").Value
            aRcvryData.OvbVolRcvryMode = "Linear model"
            '.OvbVolRcvryCf = fRcvryScenDynaset.Fields("ovb_vol_rcvry_cf").Value
            '.OvbVolRcvryVf = fRcvryScenDynaset.Fields("ovb_vol_rcvry_vf").Value
            '.OvbVolRcvryFa = fRcvryScenDynaset.Fields("ovb_vol_rcvry_fa").Value
            '.MineVolRcvryMode = fRcvryScenDynaset.Fields("mine_vol_rcvry_mode").Value
            aRcvryData.MineVolRcvryMode = "Linear model"
            '.MineVolRcvryCf = fRcvryScenDynaset.Fields("mine_vol_rcvry_cf").Value
            '.MineVolRcvryVf = fRcvryScenDynaset.Fields("mine_vol_rcvry_vf").Value
            '.MineVolRcvryFa = fRcvryScenDynaset.Fields("mine_vol_rcvry_fa").Value
            '.AdjOsTonsWvol = fRcvryScenDynaset.Fields("adj_os_tons_wvol").Value
            '.AdjPbTonsWvol = fRcvryScenDynaset.Fields("adj_pb_tons_wvol").Value
            '.AdjIpTonsWvol = fRcvryScenDynaset.Fields("adj_ip_tons_wvol").Value
            '.AdjFdTonsWvol = fRcvryScenDynaset.Fields("adj_fd_tons_wvol").Value
            '.AdjClTonsWvol = fRcvryScenDynaset.Fields("adj_cl_tons_wvol").Value
            '.PbTonRcvryCrs = fRcvryScenDynaset.Fields("pb_ton_rcvry_crs").Value
            '.PbTonRcvryFne = fRcvryScenDynaset.Fields("pb_ton_rcvry_fne").Value
            '.IpTonRcvryTot = fRcvryScenDynaset.Fields("ip_ton_rcvry_tot").Value
            '.FdTonRcvryCrs = fRcvryScenDynaset.Fields("fd_ton_rcvry_crs").Value
            '.FdTonRcvryFne = fRcvryScenDynaset.Fields("fd_ton_rcvry_fne").Value
            '.FdBplRcvryCrs = fRcvryScenDynaset.Fields("fd_bpl_rcvry_crs").Value
            '.FdBplRcvryFne = fRcvryScenDynaset.Fields("fd_bpl_rcvry_fne").Value
            '.ClTonRcvryTot = fRcvryScenDynaset.Fields("cl_ton_rcvry_tot").Value
            '.FlotRcvryMode = fRcvryScenDynaset.Fields("flot_rcvry_mode").Value
            '.FlotRcvryCrsCf = fRcvryScenDynaset.Fields("flot_rcvry_crs_cf").Value
            '.FlotRcvryCrsVf = fRcvryScenDynaset.Fields("flot_rcvry_crs_vf").Value
            '.FlotRcvryFneCf = fRcvryScenDynaset.Fields("flot_rcvry_fne_cf").Value
            '.FlotRcvryFneVf = fRcvryScenDynaset.Fields("flot_rcvry_fne_vf").Value
            '.FlotRcvryCrsTlBpl = fRcvryScenDynaset.Fields("flot_rcvry_crs_tlbpl").Value
            '.FlotRcvryCrsCnIns = fRcvryScenDynaset.Fields("flot_rcvry_crs_cnins").Value
            '.FlotRcvryFneTlBpl = fRcvryScenDynaset.Fields("flot_rcvry_fne_tlbpl").Value
            '.FlotRcvryFneCnIns = fRcvryScenDynaset.Fields("flot_rcvry_fne_cnins").Value
            '.LmTest = fRcvryScenDynaset.Fields("lm_test").Value
            '.HwTest = fRcvryScenDynaset.Fields("hw_test").Value

            'Insol adjustments -- ProdCoeff (Catalog)
            '.CrsPbInsAdjMode = fRcvryScenDynaset.Fields("crspb_insadj_mode").Value
            '.CrsPbInsAdj = fRcvryScenDynaset.Fields("crspb_insadj").Value
            '.FnePbInsAdjMode = fRcvryScenDynaset.Fields("fnepb_insadj_mode").Value
            '.FnePbInsAdj = fRcvryScenDynaset.Fields("fnepb_insadj").Value
            '.IpInsAdjMode = fRcvryScenDynaset.Fields("ip_insadj_mode").Value
            '.IpInsAdj = fRcvryScenDynaset.Fields("ip_insadj").Value
            '.CrsCnInsAdjMode = fRcvryScenDynaset.Fields("crscn_insadj_mode").Value
            '.CrsCnInsAdj = fRcvryScenDynaset.Fields("crscn_insadj").Value
            '.FneCnInsAdjMode = fRcvryScenDynaset.Fields("fnecn_insadj_mode").Value
            '.FneCnInsAdj = fRcvryScenDynaset.Fields("fnecn_insadj").Value
            '.AdjInsAfterQualTest = fRcvryScenDynaset.Fields("adj_ins_after_qual_test").Value
            .AdjInsAfterQualTest = False

            'Insol adjustments -- 100%
            'If Not IsDBNull(fRcvryScenDynaset.Fields("crspb_insadj_mode_100").Value) Then
            '    .CrsPbInsAdjMode100 = fRcvryScenDynaset.Fields("crspb_insadj_mode_100").Value
            'Else
            '    .CrsPbInsAdjMode100 = ""
            'End If
            'If Not IsDBNull(fRcvryScenDynaset.Fields("crspb_insadj_100").Value) Then
            '    .CrsPbInsAdj100 = fRcvryScenDynaset.Fields("crspb_insadj_100").Value
            'Else
            '    .CrsPbInsAdj100 = 0
            'End If
            'If Not IsDBNull(fRcvryScenDynaset.Fields("fnepb_insadj_mode_100").Value) Then
            '    .FnePbInsAdjMode100 = fRcvryScenDynaset.Fields("fnepb_insadj_mode_100").Value
            'Else
            '    .FnePbInsAdjMode100 = ""
            'End If
            'If Not IsDBNull(fRcvryScenDynaset.Fields("fnepb_insadj_100").Value) Then
            '    .FnePbInsAdj100 = fRcvryScenDynaset.Fields("fnepb_insadj_100").Value
            'Else
            '    .FnePbInsAdj100 = 0
            'End If
            'If Not IsDBNull(fRcvryScenDynaset.Fields("ip_insadj_mode_100").Value) Then
            '    .IpInsAdjMode100 = fRcvryScenDynaset.Fields("ip_insadj_mode_100").Value
            'Else
            '    .IpInsAdjMode100 = ""
            'End If
            'If Not IsDBNull(fRcvryScenDynaset.Fields("ip_insadj_100").Value) Then
            '    .IpInsAdj100 = fRcvryScenDynaset.Fields("ip_insadj_100").Value
            'Else
            '    .IpInsAdj100 = 0
            'End If
            'If Not IsDBNull(fRcvryScenDynaset.Fields("crscn_insadj_mode_100").Value) Then
            '    .CrsCnInsAdjMode100 = fRcvryScenDynaset.Fields("crscn_insadj_mode_100").Value
            'Else
            '    .CrsCnInsAdjMode100 = ""
            'End If
            'If Not IsDBNull(fRcvryScenDynaset.Fields("crscn_insadj_100").Value) Then
            '    .CrsCnInsAdj100 = fRcvryScenDynaset.Fields("crscn_insadj_100").Value
            'Else
            '    .CrsCnInsAdj100 = 0
            'End If
            'If Not IsDBNull(fRcvryScenDynaset.Fields("fnecn_insadj_mode_100").Value) Then
            '    .FneCnInsAdjMode100 = fRcvryScenDynaset.Fields("fnecn_insadj_mode_100").Value
            'Else
            '    .FneCnInsAdjMode100 = ""
            'End If
            'If Not IsDBNull(fRcvryScenDynaset.Fields("fnecn_insadj_100").Value) Then
            '    .FneCnInsAdj100 = fRcvryScenDynaset.Fields("fnecn_insadj_100").Value
            'Else
            '    .FneCnInsAdj100 = 0
            'End If

            'Economic and physical mineability criteria
            '.ClPctMaxSpl = fRcvryScenDynaset.Fields("cl_pct_max_spl").Value
            '.MtxxMaxSpl = fRcvryScenDynaset.Fields("mtxx_max_spl").Value
            '.MaxTotDepthSpl = fRcvryScenDynaset.Fields("max_tot_depth_spl").Value
            '.MaxTotDepthModeSpl = fRcvryScenDynaset.Fields("max_tot_depth_mode_spl").Value
            '.MinOreThk = fRcvryScenDynaset.Fields("min_ore_thk").Value
            '.MinItbThk = fRcvryScenDynaset.Fields("min_itb_thk").Value
            '.ClPctMaxHole = fRcvryScenDynaset.Fields("cl_pct_max_hole").Value
            '.MtxxMaxHole = fRcvryScenDynaset.Fields("mtxx_max_hole").Value
            '.TotxMaxHole = fRcvryScenDynaset.Fields("totx_max_hole").Value
            '.TotPrTpaMinHole = fRcvryScenDynaset.Fields("totpr_tpa_min_hole").Value

            'If Not IsDBNull(fRcvryScenDynaset.Fields("mine_first_spl").Value) Then
            '    .MineFirstSpl = fRcvryScenDynaset.Fields("mine_first_spl").Value
            'Else
            '    .MineFirstSpl = 0
            'End If

            '.InclCpbAlways = False
            '.InclFpbAlways = False
            '.InclCpbNever = False
            '.InclFpbNever = False
            '.InclOsAlways = False
            '.InclOsNever = False
            '.CanSelectRejectTpb = False
            '.CanSelectRejectTcn = False

            If Not IsDBNull(fRcvryScenDynaset.Fields("incl_crspb_status").Value) Then
                If fRcvryScenDynaset.Fields("incl_crspb_status").Value = "Always" Then
                    .InclCpbAlways = True
                End If
                If fRcvryScenDynaset.Fields("incl_crspb_status").Value = "Never" Then
                    .InclCpbNever = True
                End If
            End If
            If Not IsDBNull(fRcvryScenDynaset.Fields("incl_fnepb_status").Value) Then
                If fRcvryScenDynaset.Fields("incl_fnepb_status").Value = "Always" Then
                    .InclFpbAlways = True
                End If
                If fRcvryScenDynaset.Fields("incl_fnepb_status").Value = "Never" Then
                    .InclFpbNever = True
                End If
            End If
            If Not IsDBNull(fRcvryScenDynaset.Fields("incl_os_status").Value) Then
                If fRcvryScenDynaset.Fields("incl_os_status").Value = "Always" Then
                    .InclOsAlways = True
                End If
                If fRcvryScenDynaset.Fields("incl_os_status").Value = "Never" Then
                    .InclOsNever = True
                End If
            End If

            If Not IsDBNull(fRcvryScenDynaset.Fields("flot_rcvry_mode_100").Value) Then
                .FlotRcvryMode100 = fRcvryScenDynaset.Fields("flot_rcvry_mode_100").Value
            Else
                .FlotRcvryMode100 = "0 tail BPL"
            End If

            If Not IsDBNull(fRcvryScenDynaset.Fields("dens_calc_mode").Value) Then
                .DensCalcMode = fRcvryScenDynaset.Fields("dens_calc_mode").Value
            Else
                .DensCalcMode = "Limit routine"
            End If

            If Not IsDBNull(fRcvryScenDynaset.Fields("dens_mlv_spec_chk").Value) Then
                If fRcvryScenDynaset.Fields("dens_mlv_spec_chk").Value = 1 Then
                    .DensMlvSpecChk = True
                Else
                    .DensMlvSpecChk = False
                End If
            Else
                .DensMlvSpecChk = False
            End If

            If Not IsDBNull(fRcvryScenDynaset.Fields("dens_lr_spec_chk").Value) Then
                If fRcvryScenDynaset.Fields("dens_lr_spec_chk").Value = 1 Then
                    .DensLrSpecChk = True
                Else
                    .DensLrSpecChk = False
                End If
            Else
                .DensLrSpecChk = False
            End If

            If Not IsDBNull(fRcvryScenDynaset.Fields("dens_upper_limit").Value) Then
                .DensUpperLimit = fRcvryScenDynaset.Fields("dens_upper_limit").Value
            Else
                .DensUpperLimit = 0
            End If

            If Not IsDBNull(fRcvryScenDynaset.Fields("dens_lower_limit").Value) Then
                .DensLowerLimit = fRcvryScenDynaset.Fields("dens_lower_limit").Value
            Else
                .DensLowerLimit = 0
            End If

            '06/15/2010, lss "mine_has_offspec_pb_plt" not really used anymore.
            If Not IsDBNull(fRcvryScenDynaset.Fields("mine_has_offspec_pb_plt").Value) Then
                If fRcvryScenDynaset.Fields("mine_has_offspec_pb_plt").Value = 1 Then
                    .MineHasOffSpecPbPlt = True
                Else
                    .MineHasOffSpecPbPlt = False
                End If
            End If

            If Not IsDBNull(fRcvryScenDynaset.Fields("use_doloflot_plant_2010").Value) Then
                If fRcvryScenDynaset.Fields("use_doloflot_plant_2010").Value = 1 Then
                    .UseDoloflotPlant2010 = True
                Else
                    .UseDoloflotPlant2010 = False
                End If
            End If

            If Not IsDBNull(fRcvryScenDynaset.Fields("use_doloflot_plant_2010").Value) Then
                If fRcvryScenDynaset.Fields("use_doloflot_plant_2010").Value = 2 Then
                    .UseDoloflotPlantFco = True
                Else
                    .UseDoloflotPlantFco = False
                End If
            End If

            If Not IsDBNull(fRcvryScenDynaset.Fields("use_orig_mgo_plant").Value) Then
                If fRcvryScenDynaset.Fields("use_orig_mgo_plant").Value = 1 Then
                    .UseOrigMgoPlant = True
                Else
                    .UseOrigMgoPlant = False
                End If
            End If

            If Not IsDBNull(fRcvryScenDynaset.Fields("can_select_reject_tpb").Value) Then
                If fRcvryScenDynaset.Fields("can_select_reject_tpb").Value = 1 Then
                    .CanSelectRejectTpb = True
                Else
                    .CanSelectRejectTpb = False
                End If
            End If

            ''.MineHasOffSpecPbPlt = fRcvryScenDynaset.Fields("mine_has_offspec_pb_plt").Value

            .MplInpBplTarg = fRcvryScenDynaset.Fields("mpl_inp_bpl_targ").Value
            .MplInpMgoTarg = fRcvryScenDynaset.Fields("mpl_inp_mgo_targ").Value
            .MplRejBplTarg = fRcvryScenDynaset.Fields("mpl_rej_bpl_targ").Value
            .MplRejMgoTarg = fRcvryScenDynaset.Fields("mpl_rej_mgo_targ").Value
            .MplM1BpltRcvry = fRcvryScenDynaset.Fields("mpl_m1_bplt_rcvry").Value
            .MplM1BplHwire = fRcvryScenDynaset.Fields("mpl_m1_bpl_hwire").Value
            .MplM1InsHwire = fRcvryScenDynaset.Fields("mpl_m1_ins_hwire").Value
            .MplM1MgoImprove = fRcvryScenDynaset.Fields("mpl_m1_mgo_improve").Value

            If Not IsDBNull(fRcvryScenDynaset.Fields("can_select_reject_tcn").Value) Then
                If fRcvryScenDynaset.Fields("can_select_reject_tcn").Value = 1 Then
                    .CanSelectRejectTcn = True
                Else
                    .CanSelectRejectTcn = False
                End If
            End If

            If Not IsDBNull(fRcvryScenDynaset.Fields("use_adj_fe_for_minable").Value) Then
                If fRcvryScenDynaset.Fields("use_adj_fe_for_minable").Value = 1 Then
                    .UseFeAdjust = True
                Else
                    .UseFeAdjust = False
                End If
            End If

            If Not IsDBNull(fRcvryScenDynaset.Fields("upper_zone_fe_adjust").Value) Then
                .UpperZoneFeAdjust = fRcvryScenDynaset.Fields("upper_zone_fe_adjust").Value
            Else
                .UpperZoneFeAdjust = 0
            End If

            If Not IsDBNull(fRcvryScenDynaset.Fields("lower_zone_fe_adjust").Value) Then
                .LowerZoneFeAdjust = fRcvryScenDynaset.Fields("lower_zone_fe_adjust").Value
            Else
                .LowerZoneFeAdjust = 0
            End If

            .DpCrsPbMgoCutoff = fRcvryScenDynaset.Fields("dp_crspb_mgo_cutoff").Value
            .DpFnePbMgoCutoff = fRcvryScenDynaset.Fields("dp_fnepb_mgo_cutoff").Value
            .DpIpMgoCutoff = fRcvryScenDynaset.Fields("dp_ip_mgo_cutoff").Value
            .DpGrind = fRcvryScenDynaset.Fields("dp_grind").Value
            .DpAcid = fRcvryScenDynaset.Fields("dp_acid").Value
            .DpP2o5 = fRcvryScenDynaset.Fields("dp_p2o5").Value
            .DpPa64 = fRcvryScenDynaset.Fields("dp_pa64").Value
            .DpFlotMin = fRcvryScenDynaset.Fields("dp_flotmin").Value
            .DpTargMgo = fRcvryScenDynaset.Fields("dp_targ_mgo").Value
        End With

        'Set values in chkUseFeAdjust and ssFeAdjustments
        If aRcvryData.UseFeAdjust Then
            chkUseFeAdjust.Checked = True
        Else
            chkUseFeAdjust.Checked = False
        End If

        With ssFeAdjustment
            .Row = 1
            .Col = 1
            .Value = aRcvryData.UpperZoneFeAdjust
            .Row = 2
            .Col = 1
            .Value = aRcvryData.LowerZoneFeAdjust
        End With

        'Prospect recovery product quality data
        'Prospect recovery product quality data
        'Prospect recovery product quality data

        'params = gDBParams

        'params.Add("pRcvryScenarioName", aRcvryScen, ORAPARM_INPUT)
        'params("pRcvryScenarioName").serverType = ORATYPE_VARCHAR2

        'params.Add("pProspSetName", aProspSetName, ORAPARM_INPUT)
        'params("pProspSetName").serverType = ORATYPE_VARCHAR2

        'params.Add("pResult", 0, ORAPARM_OUTPUT)
        'params("pResult").serverType = ORATYPE_CURSOR

        ''PROCEDURE get_rcvry_prod_qual
        ''pRcvryScenarioName     IN     VARCHAR2,
        ''pProspSetName          IN     VARCHAR2,
        ''pResult                IN OUT c_rcvry);
        'SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prosp_data_rdctn.get_rcvry_prod_qual(" +
        '                               ":pRcvryScenarioName, :pProspSetName, :pResult);end;", ORASQL_FAILEXEC)

        'ProdQualDynaset = params("pResult").Value
        'ClearParams(params)

        'RecordCount = ProdQualDynaset.RecordCount
        'SetupDataRdctnProdQualArray(aRcvryProdQual)

        'ProdQualDynaset.MoveFirst()
        'Do While Not ProdQualDynaset.EOF
        '    With TempData
        '        .MatlTypeName = ProdQualDynaset.Fields("matl_type_name").Value
        '        .MatlName = ProdQualDynaset.Fields("matl_name").Value
        '        .SpecLevel = ProdQualDynaset.Fields("spec_level").Value
        '        .Bpl = ProdQualDynaset.Fields("bpl_min").Value
        '        .Fe2O3 = ProdQualDynaset.Fields("fe2o3_max").Value
        '        .Al2O3 = ProdQualDynaset.Fields("al2o3_max").Value
        '        .Ia = ProdQualDynaset.Fields("ia_max").Value
        '        .MgO = ProdQualDynaset.Fields("mgo_max").Value
        '        .CaO = ProdQualDynaset.Fields("cao_max").Value
        '        .Mer = ProdQualDynaset.Fields("mer_max").Value
        '        .CaOP2O5 = ProdQualDynaset.Fields("caop2o5_max").Value

        '        ProdRow = GetProdRow(TempData)

        '        If ProdRow <> 0 Then
        '            aRcvryProdQual(ProdRow).MatlTypeName = .MatlTypeName
        '            aRcvryProdQual(ProdRow).MatlName = .MatlName
        '            aRcvryProdQual(ProdRow).SpecLevel = .SpecLevel
        '            aRcvryProdQual(ProdRow).Bpl = .Bpl
        '            aRcvryProdQual(ProdRow).Fe2O3 = .Fe2O3
        '            aRcvryProdQual(ProdRow).Al2O3 = .Al2O3
        '            aRcvryProdQual(ProdRow).Ia = .Ia
        '            aRcvryProdQual(ProdRow).MgO = .MgO
        '            aRcvryProdQual(ProdRow).CaO = .CaO
        '            aRcvryProdQual(ProdRow).Mer = .Mer
        '            aRcvryProdQual(ProdRow).CaOP2O5 = .CaOP2O5
        '        End If
        '    End With
        '    ProdQualDynaset.MoveNext()
        'Loop

        'ProdQualDynaset.Close()

        Exit Sub

SetRcvryEtcError:
        MsgBox("Error getting data." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Data Get Error")

        'On Error Resume Next
        ClearParams(params)
        'On Error Resume Next
        'Try

        'Catch ex As Exception

        'End Try
        'ProdQualDynaset.Close()
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

    Private Sub GetAllMineNames()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetAllMineNamesError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim MineNameDynaset As OraDynaset

        SetActionStatus("Loading mine names...")
        Me.Cursor = Cursors.WaitCursor

        'Get all existing mine names
        params = gDBParams

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_mine_info(:pResult);end;", ORASQL_FAILEXEC)
        MineNameDynaset = params("pResult").Value
        ClearParams(params)

        cboMineName.Items.Add("(Select mine...)")
        cboSplitOverrideMineName.Items.Add("None")

        MineNameDynaset.MoveFirst()
        Do While Not MineNameDynaset.EOF
            'Only want mines with prospect data!
            If MineNameDynaset.Fields("mine_prospect").Value = 1 Then
                cboMineName.Items.Add(MineNameDynaset.Fields("mine_name").Value)
                cboSplitOverrideMineName.Items.Add(MineNameDynaset.Fields("mine_name").Value)
            End If
            MineNameDynaset.MoveNext()
        Loop
        cboMineName.Text = "(Select mine...)"
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

    Private Sub cmdTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTest.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo cmdTestClickError

        cboTwp.Text = "34"
        cboRge.Text = "22"
        cboSec.Text = "22"
        cboHole.Text = "2845"

        chkMyParams.Checked = False
        'cmdRefreshParams.Value = True

        cboProdSizeDefn.Text = "RS Win Size Fractions"
        cboOtherDefn.Text = "Ona Pion West Pion- BCD"

        'cmdReduceHole.Value = True

cmdTestClickError:
        'Don't do anything!
    End Sub

    Private Sub cmdTest2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTest2.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo cmdTest2ClickError

        cboTwp.Text = "31"
        cboRge.Text = "21"
        cboSec.Text = "26"
        cboHole.Text = "2636"

        chkMyParams.Checked = False
        'cmdRefreshParams.Value = True

        cboProdSizeDefn.Text = "RS FCO"
        cboOtherDefn.Text = "RS FCO 1112 REC tentative"

        'cmdReduceHole.Value = True

cmdTest2ClickError:
        'Don't do anything!
    End Sub

    Private Sub cmdTest3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTest3.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo cmdTest3ClickError

        chkMyParams.Checked = False
        'cmdRefreshParams.Value = True

        cboTwp.Text = "33"
        cboRge.Text = "26"
        cboSec.Text = "6"
        cboHole.Text = "1248"

        cboProdSizeDefn.Text = "RS SFM"
        cboOtherDefn.Text = "RS SFM 2013 RECOVERIES"

        ' cmdReduceHole.Value = True

cmdTest3ClickError:
        'Don't do anything!
    End Sub


    Private Function GetHoleTitle() As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        GetHoleTitle = gGetHoleLocationTrs(cboSec.Text,
                                           cboTwp.Text,
                                           cboRge.Text,
                                           cboHole.Text)
    End Function

    Private Sub PopulateSsDrillDataEtc()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Move the data from ssSplitReview and ssCompReview (which the reduction
        'process populated) to ssDrillData where the user can manipulate it!

        Dim ProspData As gRawProspSplRdctnType
        Dim ProspType As String
        Dim SplIdx As Integer
        Dim CumulativeThk As Single

        CumulativeThk = 0

        If opt100Pct.Checked = True Then
            ProspType = "100%"
        Else
            ProspType = "Catalog"
        End If

        ssDrillData.ReDraw = False
        ssHoleData.ReDraw = False
        ssSplitData.ReDraw = False

        ssDrillData.MaxRows = 0

        'Populate the composite data first!
        ProspData = gGetDataFromReviewSprd(ssCompReview, 1)
        PutDataInSsHoleDataOrSsSplitData(ProspType,
                                         ProspData,
                                         "Hole",
                                         ssHoleData)

        lblMiscComm2.Text = ""
        If ProspData.MineableHole = "MF" Then
            lblMiscComm2.ForeColor = Color.DarkRed ' &HC0&     'Dark red
            lblMiscComm2.Text = "Hole forced mineable!"
        End If
        If ProspData.MineableHole = "U" Then
            lblMiscComm2.ForeColor = Color.DarkRed '&HC0&     'Dark red
            lblMiscComm2.Text = "Hole not minable! If you add splits to make it minable you " &
                                   "must RECHECK ALL SPLITS that you want to be minable to make " &
                                   "the hole composite correct!"
        End If

        lblCoordsElev.Text = "Xcoord = " & Format(ProspData.Xcoord, "#,###,##0.00") &
                                ",  Ycoord = " & Format(ProspData.Ycoord, "#,###,##0.00") &
                                ",  Elev = " & Format(ProspData.Elevation, "#,##0.00") &
                                ",  Prospect date = " & DateTime.ParseExact(ProspData.ProspDate, "MM/dd/yyyy", CultureInfo.InvariantCulture).ToString("MM/dd/yyyy")

        If ProspData.Xcoord <= 0 Or ProspData.Ycoord <= 0 Or ProspData.Elevation <= 0 Then
            lblCoordsElev.ForeColor = Color.DarkRed ' &HC0&     'Dark red
            'lblCoordsElev.FontBold = True
        Else
            lblCoordsElev.ForeColor = Color.Black ' vbButtonText
            'lblCoordsElev.FontBold = False
        End If

        'Populate the split data (Split #1).
        For SplIdx = 1 To ssSplitReview.MaxRows
            ProspData = gGetDataFromReviewSprd(ssSplitReview, SplIdx)

            If SplIdx = 1 Then
                PutDataInSsHoleDataOrSsSplitData(ProspType,
                                                 ProspData,
                                                 "Split",
                                                 ssSplitData)

                With ssDrillData
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = 0
                    .Text = "Ovb*"

                    .Col = 1
                    .CellType = CellTypeConstants.CellTypeStaticText 'SS_CELL_TYPE_STATIC_TEXT
                    .Text = " "

                    .Col = 2
                    .Value = ProspData.SplitDepthTop

                    '01/08/2009, lss  Added cumulative thickness stuff.
                    .Col = 3
                    .Value = ProspData.SplitDepthTop

                    CumulativeThk = ProspData.SplitDepthTop

                End With
            End If

            'Add the split to ssDrillData -- this is where the user can manipulate
            'minabilities.
            CumulativeThk = CumulativeThk + ProspData.SplitThck
            PutSplitInSsDrillData(ProspData, CumulativeThk)
        Next SplIdx

        FixSpreads()

        ssDrillData.ReDraw = True
        ssHoleData.ReDraw = True
        ssSplitData.ReDraw = True
    End Sub

    Private Sub PutDataInSsHoleDataOrSsSplitData(ByVal aProspType As String,
                                                 ByRef aProspData As gRawProspSplRdctnType,
                                                 ByVal aMode As String,
                                                 ByRef aSprd As AxvaSpread)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'aProspType will be "100%" or "Catalog"
        'aMode will be "Hole" or "Split"

        If aProspType = "100%" Then
            With aSprd
                .Col = 1
                .Row = 1
                .Text = IIf(aProspData.OsOnSpec = "No", "*", "")
                .Row = 2
                .Text = IIf(aProspData.CpbOnSpec = "No", "*", "")
                .Row = 3
                .Text = IIf(aProspData.FpbOnSpec = "No", "*", "")
                .Row = 4
                .Text = IIf(aProspData.TpbOnSpec = "No", "*", "")
                .Row = 5
                .Text = IIf(aProspData.IpOnSpec = "No", "*", "")
                .Row = 6
                .Text = IIf(aProspData.TcnOnSpec = "No", "*", "")
                .Row = 7
                .Text = ""
                '-----
                .Col = 2   'Prod Wt%  100%
                .Row = 1
                .Value = aProspData.Os100.WtPct
                .Row = 2
                .Value = aProspData.Cpb100.WtPct
                .Row = 3
                .Value = aProspData.Fpb100.WtPct
                .Row = 4
                .Value = aProspData.Tpb100.WtPct
                .Row = 5
                .Value = aProspData.Ip100.WtPct
                .Row = 6
                .Value = aProspData.Tcn100.WtPct
                .Row = 7
                .Value = aProspData.Tpr100.WtPct
                '-----
                .Col = 3   'Prod TPA  100%
                .Row = 1
                .Value = aProspData.Os100.Tpa
                .Row = 2
                .Value = aProspData.Cpb100.Tpa
                .Row = 3
                .Value = aProspData.Fpb100.Tpa
                .Row = 4
                .Value = aProspData.Tpb100.Tpa
                .Row = 5
                .Value = aProspData.Ip100.Tpa
                .Row = 6
                .Value = aProspData.Tcn100.Tpa
                .Row = 7
                .FontBold = True
                .Value = aProspData.Tpr100.Tpa
                '-----
                .Col = 4   'Prod BPL  100%
                .Row = 1
                .Value = aProspData.Os100.Bpl
                .Row = 2
                .Value = aProspData.Cpb100.Bpl
                .Row = 3
                .Value = aProspData.Fpb100.Bpl
                .Row = 4
                .Value = aProspData.Tpb100.Bpl
                .Row = 5
                .Value = aProspData.Ip100.Bpl
                .Row = 6
                .Value = aProspData.Tcn100.Bpl
                .Row = 7
                .FontBold = True
                .Value = aProspData.Tpr100.Bpl
                '-----
                .Col = 5   'Prod Insol  100%
                .Row = 1
                .Value = aProspData.Os100.Ins
                .Row = 2
                .Value = aProspData.Cpb100.Ins
                .Row = 3
                .Value = aProspData.Fpb100.Ins
                .Row = 4
                .Value = aProspData.Tpb100.Ins
                .Row = 5
                .Value = aProspData.Ip100.Ins
                .Row = 6
                .Value = aProspData.Tcn100.Ins
                .Row = 7
                .Value = aProspData.Tpr100.Ins
                '-----
                .Col = 6   'Prod Fe2O3  100%
                .Row = 1
                .Value = aProspData.Os100.Fe
                .Row = 2
                .Value = aProspData.Cpb100.Fe
                .Row = 3
                .Value = aProspData.Fpb100.Fe
                .Row = 4
                .Value = aProspData.Tpb100.Fe
                .Row = 5
                .Value = aProspData.Ip100.Fe
                .Row = 6
                .Value = aProspData.Tcn100.Fe
                .Row = 7
                .Value = aProspData.Tpr100.Fe
                '-----
                .Col = 7   'Prod Al2O3  100%
                .Row = 1
                .Value = aProspData.Os100.Al
                .Row = 2
                .Value = aProspData.Cpb100.Al
                .Row = 3
                .Value = aProspData.Fpb100.Al
                .Row = 4
                .Value = aProspData.Tpb100.Al
                .Row = 5
                .Value = aProspData.Ip100.Al
                .Row = 6
                .Value = aProspData.Tcn100.Al
                .Row = 7
                .Value = aProspData.Tpr100.Al
                '-----
                .Col = 8   'Prod I&A  100%
                .Row = 1
                .Value = aProspData.Os100.Ia
                .Row = 2
                .Value = aProspData.Cpb100.Ia
                .Row = 3
                .Value = aProspData.Fpb100.Ia
                .Row = 4
                .Value = aProspData.Tpb100.Ia
                .Row = 5
                .Value = aProspData.Ip100.Ia
                .Row = 6
                .Value = aProspData.Tcn100.Ia
                .Row = 7
                .Value = aProspData.Tpr100.Ia
                '-----
                .Col = 9   'Prod MgO  100%
                .Row = 1
                .Value = aProspData.Os100.Mg
                .Row = 2
                .Value = aProspData.Cpb100.Mg
                .Row = 3
                .Value = aProspData.Fpb100.Mg
                .Row = 4
                .Value = aProspData.Tpb100.Mg
                .Row = 5
                .Value = aProspData.Ip100.Mg
                .Row = 6
                .Value = aProspData.Tcn100.Mg
                .Row = 7
                .FontBold = True
                .Value = aProspData.Tpr100.Mg
                '-----
                .Col = 10   'Prod MER1  100%  -- New 01/19/2011, lss
                .Row = 1
                .Value = gGetMer(aProspData.Os100.Bpl, aProspData.Os100.Fe, aProspData.Os100.Al,
                                 aProspData.Os100.Mg, 3)
                .Row = 2
                .Value = gGetMer(aProspData.Cpb100.Bpl, aProspData.Cpb100.Fe, aProspData.Cpb100.Al,
                                 aProspData.Cpb100.Mg, 3)
                .Row = 3
                .Value = gGetMer(aProspData.Fpb100.Bpl, aProspData.Fpb100.Fe, aProspData.Fpb100.Al,
                                 aProspData.Fpb100.Mg, 3)
                .Row = 4
                .Value = gGetMer(aProspData.Tpb100.Bpl, aProspData.Tpb100.Fe, aProspData.Tpb100.Al,
                                 aProspData.Tpb100.Mg, 3)
                .Row = 5
                .Value = gGetMer(aProspData.Ip100.Bpl, aProspData.Ip100.Fe, aProspData.Ip100.Al,
                                 aProspData.Ip100.Mg, 3)
                .Row = 6
                .Value = gGetMer(aProspData.Tcn100.Bpl, aProspData.Tcn100.Fe, aProspData.Tcn100.Al,
                                 aProspData.Tcn100.Mg, 3)
                .Row = 7
                .FontBold = True
                .Value = gGetMer(aProspData.Tpr100.Bpl, aProspData.Tpr100.Fe, aProspData.Tpr100.Al,
                                 aProspData.Tpr100.Mg, 3)
                '-----
                .Col = 11   'Prod MER2  100%  -- New 01/19/2011, lss
                .Row = 1
                .Value = gGetMerAt(aProspData.Os100.Bpl, aProspData.Os100.Fe, aProspData.Os100.Al,
                                   aProspData.Os100.Mg, 3)
                .Row = 2
                .Value = gGetMerAt(aProspData.Cpb100.Bpl, aProspData.Cpb100.Fe, aProspData.Cpb100.Al,
                                   aProspData.Cpb100.Mg, 3)
                .Row = 3
                .Value = gGetMerAt(aProspData.Fpb100.Bpl, aProspData.Fpb100.Fe, aProspData.Fpb100.Al,
                                   aProspData.Fpb100.Mg, 3)
                .Row = 4
                .Value = gGetMerAt(aProspData.Tpb100.Bpl, aProspData.Tpb100.Fe, aProspData.Tpb100.Al,
                                   aProspData.Tpb100.Mg, 3)
                .Row = 5
                .Value = gGetMerAt(aProspData.Ip100.Bpl, aProspData.Ip100.Fe, aProspData.Ip100.Al,
                                   aProspData.Ip100.Mg, 3)
                .Row = 6
                .Value = gGetMerAt(aProspData.Tcn100.Bpl, aProspData.Tcn100.Fe, aProspData.Tcn100.Al,
                                   aProspData.Tcn100.Mg, 3)
                .Row = 7
                .FontBold = True
                .Value = gGetMerAt(aProspData.Tpr100.Bpl, aProspData.Tpr100.Fe, aProspData.Tpr100.Al,
                                   aProspData.Tpr100.Mg, 3)
                '-----
                .Col = 12   'Prod CaO  100%
                .Row = 1
                .Value = aProspData.Os100.Ca
                .Row = 2
                .Value = aProspData.Cpb100.Ca
                .Row = 3
                .Value = aProspData.Fpb100.Ca
                .Row = 4
                .Value = aProspData.Tpb100.Ca
                .Row = 5
                .Value = aProspData.Ip100.Ca
                .Row = 6
                .Value = aProspData.Tcn100.Ca
                .Row = 7
                .Value = aProspData.Tpr100.Ca
                '-----
                .Col = 14   'Wt%  100%
                .Row = 2
                .Value = aProspData.Ttl100.WtPct
                .Row = 3
                .FontBold = True
                .Value = aProspData.Wcl100.WtPct
                .Row = 5
                .Value = aProspData.Cfd100.WtPct
                .Row = 6
                .Value = aProspData.Ffd100.WtPct
                .Row = 7
                .Value = aProspData.Tfd100.WtPct
                '-----
                .Col = 15   'TPA  100%
                .Row = 2
                .Value = aProspData.Ttl100.Tpa
                .Row = 3
                .Value = aProspData.Wcl100.Tpa
                .Row = 5
                .Value = aProspData.Cfd100.Tpa
                .Row = 6
                .Value = aProspData.Ffd100.Tpa
                .Row = 7
                .Value = aProspData.Tfd100.Tpa
                '-----
                .Col = 16   'BPL  100%
                .Row = 2
                .Value = aProspData.Ttl100.Bpl
                .Row = 3
                .Value = aProspData.Wcl100.Bpl
                .Row = 5
                .Value = aProspData.Cfd100.Bpl
                .Row = 6
                .Value = aProspData.Ffd100.Bpl
                .Row = 7
                .FontBold = True
                .Value = aProspData.Tfd100.Bpl
                '-----
                If aMode = "Hole" Then
                    .Col = 18
                    .Row = 1
                    .Value = aProspData.OvbThk
                    .Row = 2
                    .Value = aProspData.ItbThk
                    .Row = 3
                    .Value = aProspData.MtxThk
                    .Row = 4
                    .FontBold = True
                    .Value = aProspData.MtxxAll100Hole
                    .Row = 5
                    .FontBold = True
                    .Value = aProspData.TotxAll100Hole

                    'If aProspData.OvbThk = 0 Then we have a problem!
                    'The Fishtail depth and Ovb cored for this hole are both probably zero!
                    If aProspData.OvbThk <= 0 Then
                        lblOvbComm.Text = "Ovb <= 0  Probably Fishtail depth && Ovb cored problem for this hole!"
                        lblOvbComm.ForeColor = Color.DarkRed ' &HC0&     'Dark red
                        'lblOvbComm.FontBold = True
                    Else
                        lblOvbComm.Text = ""
                        lblOvbComm.ForeColor = Color.Black 'vbButtonText
                        'lblOvbComm.FontBold = False
                    End If

                Else    'Split
                    .Col = 18
                    .Row = 1
                    .Value = 0
                    .Row = 2
                    .Value = 0
                    .Row = 3
                    .Value = aProspData.SplitThck
                    .Row = 4
                    .FontBold = True
                    .Value = aProspData.MtxxAll100
                    .Row = 5
                    .FontBold = True
                    .Value = aProspData.TotxAll100
                End If

                .Row = 6
                .Value = aProspData.MtxDensity
                .Row = 7
                .Value = aProspData.MtxPctSol
                '-----
                If aMode = "Hole" Then
                    .Col = 20
                    .Row = 4
                    .Value = aProspData.MtxxOnSpec100Hole
                    .Row = 5
                    .Value = aProspData.TotxOnSpec100Hole
                    .Row = 6
                    .FontBold = True
                    .Text = aProspData.MineableHole100

                    'Will override it here!
                    If aProspData.MtxThk = 0 Then
                        .Text = "U"
                    End If
                Else    'Split
                    .Col = 20
                    .Row = 4
                    .Value = aProspData.MtxxOnSpec100
                    .Row = 5
                    .Value = aProspData.TotxOnSpec100
                    .Row = 6
                    .FontBold = True
                    .Text = aProspData.MineableCalcd
                End If
            End With
        Else    'Catalog
            With aSprd
                .Col = 1
                .Row = 1
                .Text = IIf(aProspData.OsOnSpec = "No", "*", "")
                .Row = 2
                .Text = IIf(aProspData.CpbOnSpec = "No", "*", "")
                .Row = 3
                .Text = IIf(aProspData.FpbOnSpec = "No", "*", "")
                .Row = 4
                .Text = IIf(aProspData.TpbOnSpec = "No", "*", "")
                .Row = 5
                .Text = IIf(aProspData.IpOnSpec = "No", "*", "")
                .Row = 6
                .Text = IIf(aProspData.TcnOnSpec = "No", "*", "")
                .Row = 7
                .Text = ""
                '-----
                .Col = 2   'Prod Wt%  Catalog
                .Row = 1
                .Value = aProspData.Os.WtPct
                .Row = 2
                .Value = aProspData.Cpb.WtPct
                .Row = 3
                .Value = aProspData.Fpb.WtPct
                .Row = 4
                .Value = aProspData.Tpb.WtPct
                .Row = 5
                .Value = aProspData.Ip.WtPct
                .Row = 6
                .Value = aProspData.Tcn.WtPct
                .Row = 7
                .Value = aProspData.Tpr.WtPct
                '-----
                .Col = 3   'Prod TPA  Catalog
                .Row = 1
                .Value = aProspData.Os.Tpa
                .Row = 2
                .Value = aProspData.Cpb.Tpa
                .Row = 3
                .Value = aProspData.Fpb.Tpa
                .Row = 4
                .Value = aProspData.Tpb.Tpa
                .Row = 5
                .Value = aProspData.Ip.Tpa
                .Row = 6
                .Value = aProspData.Tcn.Tpa
                .Row = 7
                .FontBold = True
                .Value = aProspData.Tpr.Tpa
                '-----
                .Col = 4   'Prod BPL  Catalog
                .Row = 1
                .Value = aProspData.Os.Bpl
                .Row = 2
                .Value = aProspData.Cpb.Bpl
                .Row = 3
                .Value = aProspData.Fpb.Bpl
                .Row = 4
                .Value = aProspData.Tpb.Bpl
                .Row = 5
                .Value = aProspData.Ip.Bpl
                .Row = 6
                .Value = aProspData.Tcn.Bpl
                .Row = 7
                .FontBold = True
                .Value = aProspData.Tpr.Bpl
                '-----
                .Col = 5   'Prod Insol  Catalog
                .Row = 1
                .Value = aProspData.Os.Bpl
                .Row = 2
                .Value = aProspData.Cpb.Ins
                .Row = 3
                .Value = aProspData.Fpb.Ins
                .Row = 4
                .Value = aProspData.Tpb.Ins
                .Row = 5
                .Value = aProspData.Ip.Ins
                .Row = 6
                .Value = aProspData.Tcn.Ins
                .Row = 7
                .Value = aProspData.Tpr.Ins
                '-----
                .Col = 6   'Prod Fe2O3  Catalog
                .Row = 1
                .Value = aProspData.Os.Fe
                .Row = 2
                .Value = aProspData.Cpb.Fe
                .Row = 3
                .Value = aProspData.Fpb.Fe
                .Row = 4
                .Value = aProspData.Tpb.Fe
                .Row = 5
                .Value = aProspData.Ip.Fe
                .Row = 6
                .Value = aProspData.Tcn.Fe
                .Row = 7
                .Value = aProspData.Tpr.Fe
                '-----
                .Col = 7   'Prod Al2O3  Catalog
                .Row = 1
                .Value = aProspData.Os.Al
                .Row = 2
                .Value = aProspData.Cpb.Al
                .Row = 3
                .Value = aProspData.Fpb.Al
                .Row = 4
                .Value = aProspData.Tpb.Al
                .Row = 5
                .Value = aProspData.Ip.Al
                .Row = 6
                .Value = aProspData.Tcn.Al
                .Row = 7
                .Value = aProspData.Tpr.Al
                '-----
                .Col = 8   'Prod I&A  Catalog
                .Row = 1
                .Value = aProspData.Os.Ia
                .Row = 2
                .Value = aProspData.Cpb.Ia
                .Row = 3
                .Value = aProspData.Fpb.Ia
                .Row = 4
                .Value = aProspData.Tpb.Ia
                .Row = 5
                .Value = aProspData.Ip.Ia
                .Row = 6
                .Value = aProspData.Tcn.Ia
                .Row = 7
                .Value = aProspData.Tpr.Ia
                '-----
                .Col = 9   'Prod MgO  Catalog
                .Row = 1
                .Value = aProspData.Os.Mg
                .Row = 2
                .Value = aProspData.Cpb.Mg
                .Row = 3
                .Value = aProspData.Fpb.Mg
                .Row = 4
                .Value = aProspData.Tpb.Mg
                .Row = 5
                .Value = aProspData.Ip.Mg
                .Row = 6
                .Value = aProspData.Tcn.Mg
                .Row = 7
                .FontBold = True
                .Value = aProspData.Tpr.Mg
                '-----
                .Col = 10   'Prod MER1  Catalog -- New 01/19/2011, lss
                .Row = 1
                .Value = gGetMer(aProspData.Os.Bpl, aProspData.Os.Fe, aProspData.Os.Al,
                                 aProspData.Os.Mg, 3)
                .Row = 2
                .Value = gGetMer(aProspData.Cpb.Bpl, aProspData.Cpb.Fe, aProspData.Cpb.Al,
                                 aProspData.Cpb.Mg, 3)
                .Row = 3
                .Value = gGetMer(aProspData.Fpb.Bpl, aProspData.Fpb.Fe, aProspData.Fpb.Al,
                                 aProspData.Fpb.Mg, 3)
                .Row = 4
                .Value = gGetMer(aProspData.Tpb.Bpl, aProspData.Tpb.Fe, aProspData.Tpb.Al,
                                 aProspData.Tpb.Mg, 3)
                .Row = 5
                .Value = gGetMer(aProspData.Ip.Bpl, aProspData.Ip.Fe, aProspData.Ip.Al,
                                 aProspData.Ip.Mg, 3)
                .Row = 6
                .Value = gGetMer(aProspData.Tcn.Bpl, aProspData.Tcn.Fe, aProspData.Tcn.Al,
                                 aProspData.Tcn.Mg, 3)
                .Row = 7
                .FontBold = True
                .Value = gGetMer(aProspData.Tpr.Bpl, aProspData.Tpr.Fe, aProspData.Tpr.Al,
                                 aProspData.Tpr.Mg, 3)
                '-----
                .Col = 11   'Prod MER2  Catalog -- New 01/19/2011, lss
                .Row = 1
                .Value = gGetMerAt(aProspData.Os.Bpl, aProspData.Os.Fe, aProspData.Os.Al,
                                   aProspData.Os.Mg, 3)
                .Row = 2
                .Value = gGetMerAt(aProspData.Cpb.Bpl, aProspData.Cpb.Fe, aProspData.Cpb.Al,
                                   aProspData.Cpb.Mg, 3)
                .Row = 3
                .Value = gGetMerAt(aProspData.Fpb.Bpl, aProspData.Fpb.Fe, aProspData.Fpb.Al,
                                   aProspData.Fpb.Mg, 3)
                .Row = 4
                .Value = gGetMerAt(aProspData.Tpb.Bpl, aProspData.Tpb.Fe, aProspData.Tpb.Al,
                                   aProspData.Tpb.Mg, 3)
                .Row = 5
                .Value = gGetMerAt(aProspData.Ip.Bpl, aProspData.Ip.Fe, aProspData.Ip.Al,
                                   aProspData.Ip.Mg, 3)
                .Row = 6
                .Value = gGetMerAt(aProspData.Tcn.Bpl, aProspData.Tcn.Fe, aProspData.Tcn.Al,
                                   aProspData.Tcn.Mg, 3)
                .Row = 7
                .FontBold = True
                .Value = gGetMerAt(aProspData.Tpr.Bpl, aProspData.Tpr.Fe, aProspData.Tpr.Al,
                                   aProspData.Tpr.Mg, 3)

                '-----
                .Col = 12   'Prod CaO  Catalog
                .Row = 1
                .Value = aProspData.Os.Ca
                .Row = 2
                .Value = aProspData.Cpb.Ca
                .Row = 3
                .Value = aProspData.Fpb.Ca
                .Row = 4
                .Value = aProspData.Tpb.Ca
                .Row = 5
                .Value = aProspData.Ip.Ca
                .Row = 6
                .Value = aProspData.Tcn.Ca
                .Row = 7
                .Value = aProspData.Tpr.Ca
                '-----
                .Col = 14
                .Row = 2
                .Value = aProspData.Ttl.WtPct
                .Row = 3
                .Value = aProspData.Wcl.WtPct
                .Row = 5
                .Value = aProspData.Cfd.WtPct
                .Row = 6
                .Value = aProspData.Ffd.WtPct
                .Row = 7
                .Value = aProspData.Tfd.WtPct
                '-----
                .Col = 15
                .Row = 2
                .Value = aProspData.Ttl.Tpa
                .Row = 3
                .Value = aProspData.Wcl.Tpa
                .Row = 5
                .Value = aProspData.Cfd.Tpa
                .Row = 6
                .Value = aProspData.Ffd.Tpa
                .Row = 7
                .Value = aProspData.Tfd.Tpa
                '-----
                .Col = 16
                .Row = 2
                .Value = aProspData.Ttl.Bpl
                .Row = 3
                .Value = aProspData.Wcl.Bpl
                .Row = 5
                .Value = aProspData.Cfd.Bpl
                .Row = 6
                .Value = aProspData.Ffd.Bpl
                .Row = 7
                .FontBold = True
                .Value = aProspData.Tfd.Bpl
                '-----
                If aMode = "Hole" Then
                    .Col = 18
                    .Row = 1
                    .Value = aProspData.OvbThk
                    .Row = 2
                    .Value = aProspData.ItbThk
                    .Row = 3
                    .Value = aProspData.MtxThk
                    .Row = 4
                    .FontBold = True
                    .Value = aProspData.MtxxAllPcHole
                    .Row = 5
                    .FontBold = True
                    .Value = aProspData.TotxAllPcHole
                Else    'Split
                    .Col = 18
                    .Row = 1
                    .Value = 0
                    .Row = 2
                    .Value = 0
                    .Row = 3
                    .Value = aProspData.SplitThck
                    .Row = 4
                    .FontBold = True
                    .Value = aProspData.MtxxAll
                    .Row = 5
                    .FontBold = True
                    .Value = aProspData.TotxAll
                End If

                .Row = 6
                .Value = aProspData.MtxDensity
                .Row = 7
                .Value = aProspData.MtxPctSol
                '-----
                If aMode = "Hole" Then
                    .Col = 20
                    .Row = 4
                    .Value = aProspData.MtxxOnSpecPcHole
                    .Row = 5
                    .Value = aProspData.TotxOnSpecPcHole
                    .Row = 6
                    .FontBold = True
                    .Text = aProspData.MineableHole

                    'Will override it here!
                    If aProspData.MtxThk = 0 Then
                        .Text = "U"
                    End If
                Else    'Split
                    .Col = 20
                    .Row = 4
                    .Value = aProspData.MtxxOnSpec
                    .Row = 5
                    .Value = aProspData.TotxOnSpec
                    .Row = 6
                    .FontBold = True
                    .Text = aProspData.MineableCalcd
                End If
            End With
        End If
    End Sub

    Private Sub PutSplitInSsDrillData(ByRef aProspData As gRawProspSplRdctnType,
                                      ByVal aCumulativeThk As Single)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Columns in ssDrillData
        ' 1) Mineability check box
        ' 2) Thickness
        ' 3) Cumulative thickness
        ' 4) Divider
        ' 5) Tpr BPL
        ' 6) Tpr Ins
        ' 7) Tpr Fe
        ' 8) Tpr Al
        ' 9) Tpr Mg
        '10) Tpr Ca
        '11) Tpr TPA
        '12) Divider
        '13) Mtx"X" All
        '14) Mtx"X" OnSpec
        '15) %Clay
        '16) Divider
        '17) Tpb BPL
        '18) Tpb Ins
        '19) Tpb Fe
        '20) Tpb Al
        '21) Tpb Mg
        '22) Tpb Ca
        '23) Tpb TPA
        '24) Divider
        '25) Tcn BPL
        '26) Tcn Ins
        '27) Tcn Fe
        '28) Tcn Al
        '29) Tcn Mg
        '30) Tcn Ca
        '31) Tcn TPA
        '32) Divider
        '33) Tfd BPL
        '34) Tfd TPA
        '35) Divider
        '36) IP BPL
        '37) IP Ins
        '38) IP Fe
        '39) IP Al
        '40) IP Mg
        '41) IP Ca
        '42) IP TPA
        '43) Divider
        '44) OS BPL
        '45) OS Ins
        '46) OS Fe
        '47) OS Al
        '48) OS Mg
        '49) OS Ca
        '50) OS TPA
        '51) Divider
        '52) OS OffSpec
        '53) Cpb OffSpec
        '54) Fpb OffSpec
        '55) Tpb OffSpec
        '56) IP OffSpec
        '57) Ccn OffSpec
        '58) Fcn OffSpec
        '59) Tcn OffSpec
        '60) Divider
        '61) Sample ID
        '62) Original mineability (calculated by reduction)

        With ssDrillData
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows

            .Col = 0
            .Text = "Spl" & CStr(aProspData.SplitNumber)


            .Col = 1
            If aProspData.MineableCalcd = "M" Then
                .Value = 1
            Else
                .Value = 0
            End If


            .Col = 2
            .Value = aProspData.SplitThck


            .Col = 3
            .Value = aCumulativeThk



            If opt100Pct.Checked = True Then
                .Col = 5
                .Value = aProspData.Tpr100.Bpl
                .Col = 6
                .Value = aProspData.Tpr100.Ins
                .Col = 7
                .Value = aProspData.Tpr100.Fe
                .Col = 8
                .Value = aProspData.Tpr100.Al
                .Col = 9
                .Value = aProspData.Tpr100.Mg
                .Col = 10
                .Value = aProspData.Tpr100.Ca
                .Col = 11
                .Value = aProspData.Tpr100.Tpa
                '-----
                .Col = 13
                .Value = aProspData.MtxxAll100
                .Col = 14
                .Value = aProspData.MtxxOnSpec100
                .Col = 15
                .Value = aProspData.Wcl100.WtPct
                '-----
                .Col = 17
                .Value = aProspData.Tpb100.Bpl
                .Col = 18
                .Value = aProspData.Tpb100.Ins
                .Col = 19
                .Value = aProspData.Tpb100.Fe
                .Col = 20
                .Value = aProspData.Tpb100.Al
                .Col = 21
                .Value = aProspData.Tpb100.Mg
                .Col = 22
                .Value = aProspData.Tpb100.Ca
                .Col = 23
                .Value = aProspData.Tpb100.Tpa
                '-----
                .Col = 25
                .Value = aProspData.Tcn100.Bpl
                .Col = 26
                .Value = aProspData.Tcn100.Ins
                .Col = 27
                .Value = aProspData.Tcn100.Fe
                .Col = 28
                .Value = aProspData.Tcn100.Al
                .Col = 29
                .Value = aProspData.Tcn100.Mg
                .Col = 30
                .Value = aProspData.Tcn100.Ca
                .Col = 31
                .Value = aProspData.Tcn100.Tpa
                '-----
                .Col = 33
                .Value = aProspData.Tfd100.Bpl
                .Col = 34
                .Value = aProspData.Tfd100.Tpa
                '-----
                .Col = 36
                .Value = aProspData.Ip100.Bpl
                .Col = 37
                .Value = aProspData.Ip100.Ins
                .Col = 38
                .Value = aProspData.Ip100.Fe
                .Col = 39
                .Value = aProspData.Ip100.Al
                .Col = 40
                .Value = aProspData.Ip100.Mg
                .Col = 41
                .Value = aProspData.Ip100.Ca
                .Col = 42
                .Value = aProspData.Ip100.Tpa
                '-----
                .Col = 44
                .Value = aProspData.Os100.Bpl
                .Col = 45
                .Value = aProspData.Os100.Ins
                .Col = 46
                .Value = aProspData.Os100.Fe
                .Col = 47
                .Value = aProspData.Os100.Al
                .Col = 48
                .Value = aProspData.Os100.Mg
                .Col = 49
                .Value = aProspData.Os100.Ca
                .Col = 50
                .Value = aProspData.Os100.Tpa
                '-----
                .Col = 52
                .Text = IIf(aProspData.OsOnSpec = "No", "*", "")
                .Row = 53
                .Text = IIf(aProspData.CpbOnSpec = "No", "*", "")
                .Row = 54
                .Text = IIf(aProspData.FpbOnSpec = "No", "*", "")
                .Row = 55
                .Text = IIf(aProspData.TpbOnSpec = "No", "*", "")
                .Row = 56
                .Text = IIf(aProspData.IpOnSpec = "No", "*", "")
                .Row = 57
                .Text = IIf(aProspData.CcnOnSpec = "No", "*", "")
                .Row = 58
                .Text = IIf(aProspData.FcnOnSpec = "No", "*", "")
                .Row = 59
                .Text = IIf(aProspData.TcnOnSpec = "No", "*", "")
                '-----
                .Col = 61
                .Text = aProspData.SampleId
                .Col = 62
                .Text = aProspData.MineableCalcd
            Else    'Catalog
                .Col = 5
                .Value = aProspData.Tpr.Bpl
                .Col = 6
                .Value = aProspData.Tpr.Ins
                .Col = 7
                .Value = aProspData.Tpr.Fe
                .Col = 8
                .Value = aProspData.Tpr.Al
                .Col = 9
                .Value = aProspData.Tpr.Mg
                .Col = 10
                .Value = aProspData.Tpr.Ca
                .Col = 11
                .Value = aProspData.Tpr.Tpa
                '-----
                .Col = 13
                .Value = aProspData.MtxxAll
                .Col = 14
                .Value = aProspData.MtxxOnSpec
                .Col = 15
                .Value = aProspData.Wcl.WtPct
                '-----
                .Col = 17
                .Value = aProspData.Tpb.Bpl
                .Col = 18
                .Value = aProspData.Tpb.Ins
                .Col = 19
                .Value = aProspData.Tpb.Fe
                .Col = 20
                .Value = aProspData.Tpb.Al
                .Col = 21
                .Value = aProspData.Tpb.Mg
                .Col = 22
                .Value = aProspData.Tpb.Ca
                .Col = 23
                .Value = aProspData.Tpb.Tpa
                '-----
                .Col = 25
                .Value = aProspData.Tcn.Bpl
                .Col = 26
                .Value = aProspData.Tcn.Ins
                .Col = 27
                .Value = aProspData.Tcn.Fe
                .Col = 28
                .Value = aProspData.Tcn.Al
                .Col = 29
                .Value = aProspData.Tcn.Mg
                .Col = 30
                .Value = aProspData.Tcn.Ca
                .Col = 31
                .Value = aProspData.Tcn.Tpa
                '-----
                .Col = 33
                .Value = aProspData.Tfd.Bpl
                .Col = 34
                .Value = aProspData.Tfd.Tpa
                '-----
                .Col = 36
                .Value = aProspData.Ip.Bpl
                .Col = 37
                .Value = aProspData.Ip.Ins
                .Col = 38
                .Value = aProspData.Ip.Fe
                .Col = 39
                .Value = aProspData.Ip.Al
                .Col = 40
                .Value = aProspData.Ip.Mg
                .Col = 41
                .Value = aProspData.Ip.Ca
                .Col = 42
                .Value = aProspData.Ip.Tpa
                '-----
                .Col = 44
                .Value = aProspData.Os.Bpl
                .Col = 45
                .Value = aProspData.Os.Ins
                .Col = 46
                .Value = aProspData.Os.Fe
                .Col = 47
                .Value = aProspData.Os.Al
                .Col = 48
                .Value = aProspData.Os.Mg
                .Col = 49
                .Value = aProspData.Os.Ca
                .Col = 50
                .Value = aProspData.Os.Tpa
                '-----
                .Col = 52
                .Text = IIf(aProspData.OsOnSpec = "No", "*", "")
                .Row = 53
                .Text = IIf(aProspData.CpbOnSpec = "No", "*", "")
                .Row = 54
                .Text = IIf(aProspData.FpbOnSpec = "No", "*", "")
                .Row = 55
                .Text = IIf(aProspData.TpbOnSpec = "No", "*", "")
                .Row = 56
                .Text = IIf(aProspData.IpOnSpec = "No", "*", "")
                .Row = 57
                .Text = IIf(aProspData.CcnOnSpec = "No", "*", "")
                .Row = 58
                .Text = IIf(aProspData.FcnOnSpec = "No", "*", "")
                .Row = 59
                .Text = IIf(aProspData.TcnOnSpec = "No", "*", "")
                '-----
                .Col = 61
                .Text = aProspData.SampleId
                .Col = 62
                .Text = aProspData.MineableCalcd
            End If

        End With
    End Sub

    Private Sub ssDrillData_ButtonClicked(ByVal sender As System.Object, ByVal e As AxFPSpread._DSpreadEvents_ButtonClickedEvent) Handles ssDrillData.ButtonClicked 'ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

        Dim ThkTable(0 To 30) As gSplThkType
        Dim ProcessStatus As Boolean
        Dim RcvryData As gDataRdctnParamsType = Nothing
        Dim RcvryProdQual(0 To 14) As gDataRdctnProdQualType
        Dim Minability As String
        Dim ThisSplDesc As String
        Dim ThisSplNum As Integer

        If fReducing = False Then
            lblUserMadeHoleUnmineable.Text = ""

            If e.col = 1 And e.row >= 2 Then
                'Need to recalculate the composite!
                'The split data that needs to be recalculated is in ssSplitReview.

                'SetRcvryOnly cboOtherDefn.Text, _
                '"User recovery scenario", _
                'RcvryData

                'Set recovery, etc. information
                SetRcvryEtc(cboOtherDefn.Text,
                    "User recovery scenario",
                    RcvryData,
                    RcvryProdQual)

                If chkOverrideMaxDepth.Checked = True Then
                    'Just make it a real big number!
                    RcvryData.MaxTotDepthSpl = 9999
                End If

                With ssDrillData
                    .Row = e.row
                    .Col = e.col
                    If .Value = 1 Then
                        Minability = "M"
                    Else
                        Minability = "U"
                    End If

                    .Row = e.row
                    .Col = 0
                    ThisSplDesc = .Text     'Spl1, Spl2, Spl3, etc.
                    ThisSplNum = Val(Mid(ThisSplDesc, 4))
                End With

                With ssSplitReview
                    .Row = e.row - 1
                    .Col = 4
                    .Text = Minability
                    .Col = 190
                    .Text = Minability
                    .Col = 313
                    .Text = Minability
                End With

                ssCompReview.MaxRows = 0

                Dim ProductSizeDesignation As ViewModels.ProductSizeDesignation = GetProductSizeDistribution(cboProdSizeDefn.Text)
                Dim RecoveryDef As ViewModels.ProductRecoveryDefinition = GetRecoveryDefinition(cboOtherDefn.Text)

                Dim IsIPDistributedTo As Boolean = False
                IsIPDistributedTo = MatlDistributedTo(ProductSizeDesignation, "IP")

                Dim Holes As List(Of gRawProspSplRdctnType)
                Holes = gCompositeSplitData(ssSplitReview,
                                            RcvryData,
                                            RecoveryDef,
                                            True,
                                            IsIPDistributedTo)
                For Each Hole As gRawProspSplRdctnType In Holes
                    AssignCompositedHoleData(ssCompReview, Hole)
                Next

                'New composite data is in ssCompReview -- need to put it into
                'ssHoleData.
                PopulateComp()

                'If the split whose minability has been changed is currently
                'being displayed in ssSplitData then we need to change the
                'minability in Col18, Row6.

                If lblCurrSplit.Text = CStr(ThisSplNum) Then
                    'The split is currently displayed!
                    With ssSplitData
                        .Row = 6
                        .Col = 20   '01/19/2011, lss  Was 18
                        .Text = Minability
                    End With
                End If
            End If
        End If
    End Sub

    Private Function GetProductSizeDistribution(ByVal ProdSizeDefnName As String) As ViewModels.ProductSizeDesignation
        Dim productSizeDesignation As ViewModels.ProductSizeDesignation = Nothing
        Using svc As New ReductionService.ReductionClient
            Dim psizeDefndata = svc.GetProspectUserProductSizeDefinition(ProdSizeDefnName)
            If Not psizeDefndata Is Nothing Then
                productSizeDesignation = New ViewModels.ProductSizeDesignation(psizeDefndata)
            End If
        End Using
        Return productSizeDesignation
    End Function

    Private Function GetRecoveryDefinition(ByVal RecoveryDefName As String) As ViewModels.ProductRecoveryDefinition
        Dim recoveryDefinition As ViewModels.ProductRecoveryDefinition = New ViewModels.ProductRecoveryDefinition
        Using svc As New ReductionService.ReductionClient
            Dim scenarioData = svc.GetProspectUserProductRecoveryDefinition(RecoveryDefName)
            If Not scenarioData Is Nothing Then
                recoveryDefinition = New ViewModels.ProductRecoveryDefinition(scenarioData)
            End If
        End Using
        Return recoveryDefinition
    End Function


    Private Sub ssDrillData_Click(ByVal sender As System.Object, ByVal e As AxFPSpread._DSpreadEvents_ClickEvent) Handles ssDrillData.ClickEvent 'ByVal Col As Long, ByVal Row As Long)

        If e.col <> 1 And e.row > 1 Then
            'Need to display another split in ssSplitData.
            PopulateSplit(e.row - 1)
        End If
    End Sub

    Private Sub PopulateSplit(ByVal aSplitNumber As Integer)

        Dim ProspData As gRawProspSplRdctnType = Nothing
        Dim ProspType As String

        'Move the data from ssSplitReview to ssSplitData.

        If opt100Pct.Checked = True Then
            ProspType = "100%"
        Else
            ProspType = "Catalog"
        End If

        'Populate the split data in ssSplitData.
        ProspData = gGetDataFromReviewSprd(ssSplitReview, aSplitNumber)

        PutDataInSsHoleDataOrSsSplitData(ProspType,
                                         ProspData,
                                         "Split",
                                         ssSplitData)

        lblSplit.Text = GetHoleTitle() & "  Split# " & CStr(aSplitNumber)
        lblCurrSplit.Text = CStr(aSplitNumber)

        If opt100Pct.Checked = True Then
            lblSplit.Text = lblSplit.Text & "   (100% Prospect)"
        Else    'Catalog
            lblSplit.Text = lblSplit.Text & "   (Catalog)"
        End If
    End Sub

    Private Sub PopulateComp()

        Dim ProspData As gRawProspSplRdctnType
        Dim ProspType As String

        'Move the data from ssCompReview to ssHoleData.

        If opt100Pct.Checked = True Then
            ProspType = "100%"
        Else
            ProspType = "Catalog"
        End If

        'Populate the composite data in ssHoleData.
        ProspData = gGetDataFromReviewSprd(ssCompReview, 1)

        PutDataInSsHoleDataOrSsSplitData(ProspType,
                                         ProspData,
                                         "Hole",
                                         ssHoleData)
    End Sub

    Private Sub opt100Pct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles opt100Pct.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        If ssDrillData.MaxRows > 1 Then
            SetActionStatus("Getting data...")
            Me.Cursor = Cursors.WaitCursor

            fReducing = True
            ProcessReducedData(False)
            fReducing = False

            SetActionStatus("")
            Me.Cursor = Cursors.Arrow
        End If
    End Sub

    Private Sub optCatalog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCatalog.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        If ssDrillData.MaxRows > 1 Then
            SetActionStatus("Getting data...")
            Me.Cursor = Cursors.WaitCursor

            fReducing = True
            ProcessReducedData(False)
            fReducing = False

            SetActionStatus("")
            Me.Cursor = Cursors.Arrow
        End If
    End Sub

    Private Sub SetMineName()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ThisSampleId As String
        Dim ThisMine As String

        'For now will get the mine name that was assigned to the first
        'split for the hole (in raw prospect data, probaly when Ernest
        'entered the hole in MOIS).
        With ssSplitReview
            If .MaxRows <> 0 Then
                .Row = 1
                .Col = 156
                ThisSampleId = .Text

                ThisMine = gGetMineForSampleId(ThisSampleId)

                If Trim(ThisMine) <> "" Then
                    cboMineName.Text = ThisMine
                Else
                    cboMineName.Text = "(Select mine...)"
                End If
            End If
        End With
    End Sub

    Private Sub cmdSaveMinabilities_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveMinabilities.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SaveMinabilities(True)
    End Sub

    Private Sub SaveMinabilities(ByVal aAskIfOk As Boolean)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo SaveMinabilitiesError

        'Need to save the split minabilities that have been determined into
        'the raw prospect data.  Need to save the hole mineability also!

        'Do not need the mine name for this -- only saving to the raw prospect
        'which doesn't really need the mine name.

        Dim HoleMinable As Integer
        Dim SampleId() As String
        Dim SplitMinable() As Integer
        Dim ItemCount As Integer
        Dim InsertSQL As String
        Dim RowIdx As Integer
        Dim WhoModified As String
        Dim WhenModified As Date
        Dim MinableSplitExists As Boolean
        Dim ProspDate As Date
        Dim MsgStr As String
        Dim TprTpa As Single

        If aAskIfOk = True Then
            MsgStr = "You are going to be saving your minability selections to " &
                     "the raw prospect data.  If you have marked all of the " &
                     "splits as unmineable then the hole will also be marked as " &
                     "unminable in the raw prospect data." & vbCrLf & vbCrLf &
                     "OK to continue? "

            If MsgBox(MsgStr, vbYesNo +
                vbDefaultButton1, "Save OK?") <> vbYes Then
                Exit Sub
            End If
        End If

        SetActionStatus("Saving minabilities...")
        Me.Cursor = Cursors.WaitCursor

        'Need to assign some dates and stuff.
        WhoModified = StrConv(gUserName, vbUpperCase)
        'WhenModified = CDate(Format(Now, "MM/dd/yyyy hh:mm AM/PM"))   *****************************  Changed 12/10/2018
        WhenModified = Now

        ItemCount = ssDrillData.MaxRows - 1

        'If data exists then redimension transfer arrays
        If ItemCount > 0 Then
            ReDim SplitMinable(ItemCount - 1)
            ReDim SampleId(ItemCount - 1)
        Else
            'Nothing to update
            Exit Sub
        End If

        With ssCompReview
            .Row = 1
            .Col = 3
            If .Text <> "None" Then
                ProspDate = CDate(.Text)
            Else
                ProspDate = #12/31/8888#
            End If
        End With

        'Now place the data into the transfer arrays.
        ItemCount = 0
        MinableSplitExists = False

        With ssDrillData
            For RowIdx = 2 To .MaxRows
                .Row = RowIdx
                .Col = 1   'Minability
                SplitMinable(ItemCount) = .Value   'Will be 0 or 1

                If .Value = 1 Then
                    MinableSplitExists = True
                End If

                .Col = 61  'Sample ID  (was Col 60)
                SampleId(ItemCount) = .Text

                ItemCount = ItemCount + 1
            Next RowIdx
        End With

        'This is not correct!  Even though a mineable split exists the hole may be unmineable!
        'If MinableSplitExists = True Then
        'If the total product TPA for the hole is zero then the hole is unmineable!
        With ssHoleData
            .Row = 7
            .Col = 3   '01/19/2011, lss  This is still OK -- Col 3 is still Col 3 (TPA)
            TprTpa = .Value
        End With

        If TprTpa > 0 Then
            HoleMinable = 1
        Else
            HoleMinable = 0
        End If

        'Need to get the prospect date for this hole.
        With ssCompReview
            .Row = 1
            .Col = 3
            If .Text <> "None" Then
                ProspDate = CDate(.Text)
            Else
                ProspDate = #12/31/8888#
            End If
        End With

        'PROCEDURE update_raw_minabilities
        'pArraySize        IN     INTEGER,
        'pTownship         IN     NUMBER,
        'pRange            IN     NUMBER,
        'pSection          IN     NUMBER,
        'pHoleLocation     IN     VARCHAR2,
        'pProspDate        IN     DATE,
        'pHoleMinable      IN     NUMBER,
        'pWhoModified      IN     VARCHAR2,
        'pWhenModified     IN     DATE,
        'pSampleId         IN     VCHAR2ARRAY8,
        'pSplitMinable     IN     NUMBERARRAY,
        'pResult           IN OUT NUMBER)
        InsertSQL = "Begin mois.mois_raw_prospectnew.update_raw_minabilities(" &
        "   :pArraySize, " &
        "   :pTownship, " &
        "   :pRange, " &
        "   :pSection, " &
        "   :pHoleLocation, " &
        "   :pProspDate, " &
        "   :pHoleMinable, " &
        "   :pWhoModified, " &
        "   :pWhenModified, " &
        "   :pSampleId, " &
        "   :pSplitMinable, " &
        "   :pResult); " &
        "end;"
        Dim arA1() As Object = {"pArraySize", ItemCount, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA2() As Object = {"pTownship", cboTwp.Text, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA3() As Object = {"pRange", cboRge.Text, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA4() As Object = {"pSection", cboSec.Text, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA5() As Object = {"pHoleLocation", cboHole.Text, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA6() As Object = {"pProspDate", ProspDate, ORAPARM_INPUT, ORATYPE_DATE}
        Dim arA7() As Object = {"pHoleMinable", HoleMinable, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA8() As Object = {"pWhoModified", WhoModified, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA9() As Object = {"pWhenModified", WhenModified, ORAPARM_INPUT, ORATYPE_DATE}
        Dim arA10() As Object = {"pSampleId", SampleId, ORAPARM_INPUT, ORATYPE_VARCHAR2, 8}
        Dim arA11() As Object = {"pSplitMinable", SplitMinable, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA12() As Object = {"pResult", ItemCount, ORAPARM_INPUT, ORATYPE_NUMBER}

        'RunBatchSP(InsertSQL, _
        '    Array("pArraySize", ItemCount, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pTownship", cboTwp.Text, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pRange", cboRge.Text, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pSection", cboSec.Text, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pHoleLocation", cboHole.Text, ORAPARM_INPUT, ORATYPE_VARCHAR2), _
        '    Array("pProspDate", ProspDate, ORAPARM_INPUT, ORATYPE_DATE), _
        '    Array("pHoleMinable", HoleMinable, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pWhoModified", WhoModified, ORAPARM_INPUT, ORATYPE_VARCHAR2), _
        '    Array("pWhenModified", WhenModified, ORAPARM_INPUT, ORATYPE_DATE), _
        '    Array("pSampleId", SampleId(), ORAPARM_INPUT, ORATYPE_VARCHAR2, 8), _
        '    Array("pSplitMinable", SplitMinable(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pResult", 0, ORAPARM_OUTPUT, ORATYPE_NUMBER))
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
        arA12)

        If aAskIfOk = True Then
            MsgBox("Minabilities saved to raw prospect.", vbOKOnly, "Save Status")
        End If

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow

        Exit Sub

SaveMinabilitiesError:
        MsgBox("Error while saving." & Str(Err.Number) &
               Chr(10) & Chr(10) &
               Err.Description, vbExclamation,
               "Update Error")

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub cmdSaveCompAndSplits_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveCompAndSplits.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ProspData As gRawProspSplRdctnType
        Dim MsgStr As String
        Dim Status As Boolean
        Dim RowIdx As Integer
        Dim ProspDate As Date
        Dim UserSetHoleUnmineable As Boolean
        Dim OvbThk As Single
        Dim ProspDateMois As Date
        Dim SplCntMois As Integer
        Dim SumMtxThk As Single
        Dim HoleCompOvbThk As Single
        Dim RdctnWhoWhenStatus As Boolean
        Dim StandardsToSave As String

        UserSetHoleUnmineable = False

        '1) opt100PctRdctn
        '2) optCatalogRdctn
        '3) optBothRdctn
        If opt100PctRdctn.Checked = True Then
            StandardsToSave = "100% Prospect Only"
        End If
        If optCatalogRdctn.Checked = True Then
            StandardsToSave = "Catalog Only"
        End If
        If optBothRdctn.Checked = True Then
            StandardsToSave = "100% Prospect and Catalog"
        End If

        'Need to check the hole we will be saving to MOIS to see if the hole
        'that it may be replacing in MOIS may be newer than the hole that
        'we are replacing it with!

        'Need to save the data in ssCompReview and ssSplitReview to the
        'MOIS prospect composite and split tables.
        'Will save the composite first and then the splits.
        'If the hole already exists in MOIS then it will be replaced.
        'Need to save the minability selections to raw prospect also!
        'Will check the 100% Prospect hole and assume that if the Catalog hole
        'exists then it will be the same.

        'Could use this procedure but already have existing MOIS data in
        'ssHoleExistStatus.
        'gGetHoleDateAndNumSplits cboMineName.Text, _
        'cboSec.Text, _
        'cboTwp.Text, _
        'cboRge.Text, _
        'cboHole.Text, _
        '"100%PROSPECT", _
        'ProspDateMois, _
        'SplCntMois
        Try

            'Need the prospect date for the hole we have reduced.
            With ssCompReview
                .Row = 1
                .Col = 3
                If .Text <> "None" Then
                    ProspDate = CDate(.Text)
                Else
                    ProspDate = #12/31/8888#
                End If
            End With

            'Have existing MOIS data in ssHoleExistStatus.
            If ssHoleExistStatus.MaxRows <> 0 Then
                With ssHoleExistStatus
                    .Row = 1
                    .Col = 3

                    'Let's try to get a real date for the character date!
                    'DRILL_CDATE in PROSPECT_COMP_BASE may be some goofy date from GEOCOMP!
                    ProspDateMois = gGetDrillDate(.Text)

                    .Col = 4
                    SplCntMois = .Value
                End With

                'Need the prospect date for the hole we have reduced.
                With ssCompReview
                    .Row = 1
                    .Col = 3
                    If .Text <> "None" Then
                        ProspDate = CDate(.Text)
                    Else
                        ProspDate = #12/31/8888#
                    End If
                End With

                'See if we are replacing a hole that is more recent than what we just
                'reduced!
                If ProspDateMois > ProspDate And ProspDateMois <> #12/31/8888# Then
                    MsgStr = "You are going to replace a hole in MOIS that has " &
                         "prospect date = " & Format(ProspDateMois, "MM/dd/yyyy") &
                         " and split count = " & CStr(SplCntMois) & vbCrLf & vbCrLf &
                         "This may create a problem -- make sure you know what you " &
                         "are doing!! OK to Continue with the Save?"

                    If MsgBox(MsgStr, vbYesNo +
                    vbDefaultButton1, "Save OK?") <> vbYes Then
                        Exit Sub
                    End If
                End If
            End If

            'Need to check a special situation:
            'Hole was unminable -- user has made the hole minable and is now saving it.
            'Max depth does not come into play on this hole.
            'In this case the sum of the minable matrix thicknesses from ssDrillData should be the
            'hole composite matrix thickness value in ssHoleData!
            'Also the sum of the interburden (waste) thicknesses from ssDrillData should be the
            'hole composite waste thickness value in ssHoleData!
            If InStr(lblMaxDepthComm.Text, "None") <> 0 And
            InStr(lblMiscComm2.Text, "Hole not minable!") <> 0 Then
                SumMtxThk = 0
                With ssDrillData
                    For RowIdx = 3 To .MaxRows
                        .Row = RowIdx
                        .Col = 1
                        If .Text = "1" Then
                            .Col = 2
                            SumMtxThk = SumMtxThk + .Value
                        End If
                    Next RowIdx
                End With
                With ssHoleData
                    .Col = 18   '01/19/2011, lss  Was Col 16
                    .Row = 3
                    HoleCompOvbThk = .Value
                End With

                If SumMtxThk <> HoleCompOvbThk Then
                    MsgStr = "Hole composite Mtx' = " & Format(HoleCompOvbThk, "##0.0") &
                         " and the sum of the minable split Mtx' = " &
                         Format(SumMtxThk, "##0.0") & "!" & vbCrLf & vbCrLf &
                         "You may NOT SAVE this data to the MOIS reduced composite and split tables!"
                    MsgBox(MsgStr, vbOKOnly, "Save Problem")
                    Exit Sub
                End If
            End If

            If chkSaveRawProspectMinabilities.Checked = False Then
                MsgStr = "Your minability selections WILL NOT be saved to " &
                     "the raw prospect data (hole or splits)." & vbCrLf & vbCrLf &
                     "You will be saving the hole composite and splits to " &
                     "MOIS.  If the hole already exists in MOIS it will be " &
                     "overwritten!"
            Else
                MsgStr = "You are going to be saving your minability selections to " &
                     "the raw prospect data.  If you have marked all of the " &
                     "splits as unmineable then the hole will be marked as " &
                     "unminable in the raw prospect data." & vbCrLf & vbCrLf &
                     "You will be saving the hole composite and splits to " &
                     "MOIS also.  If the hole already exists in MOIS it will be " &
                     "overwritten!"
            End If

            If opt100PctRdctn.Checked = True Then
                MsgStr = MsgStr & vbCrLf & vbCrLf &
                     "You will be saving the 100% Prospect data only!"
            End If
            If optCatalogRdctn.Checked = True Then
                MsgStr = MsgStr & vbCrLf & vbCrLf &
                     "You will be saving the Catalog data only!"
            End If
            If optBothRdctn.Checked = True Then
                MsgStr = MsgStr & vbCrLf & vbCrLf &
                     "You will be saving both the 100% Prospect and the Catalog data!"
            End If

            MsgStr = MsgStr & vbCrLf & vbCrLf &
                 "You will be saving this hole to " & cboMineName.Text & "." &
                 vbCrLf & vbCrLf &
                 "OK to continue? "

            If MsgBox(MsgStr, vbYesNo +
            vbDefaultButton1, "Save OK?") <> vbYes Then
                Exit Sub
            End If

            '01/06/2010, lss
            'User now has the option to save or not save the minability selections to raw prospect.
            'Save the minability selections to raw prospect first!
            If chkSaveRawProspectMinabilities.Checked Then
                SaveMinabilities(False)
            End If

            'Save the reduction who and when to raw prospect.
            gSaveRdctnWhoAndWhen(cboTwp.Text,
                             cboRge.Text,
                             cboSec.Text,
                             cboHole.Text,
                             ProspDate,
                             False,
                             "", #12/31/8888#)

            'Add the composite.
            SetActionStatus("Saving hole (composite)...")
            Me.Cursor = Cursors.WaitCursor
            ProspData = gGetDataFromReviewSprd(ssCompReview, 1)

            With ssHoleData
                .Row = 7
                .Col = 3   '01/19/2011, lss  This is still OK -- Col 3 is still Col 3  (TPA)
                If .Value = 0 And ProspData.Tpr.Tpa <> 0 Then
                    UserSetHoleUnmineable = True
                End If
            End With

            If UserSetHoleUnmineable Then
                'gGetDataFromReviewSprd above assigned ProspData from ssCompReview.  The
                'problem is that some of the splits in the composite may be mineable although
                'the user wants the hole to be unmineable -- the data in ssCompReview will
                'not reflect this!  SetOverrideZeros will zero out some of the data in
                'ssCompReview so that the hole will be unmineable in MOIS when the data
                'is save to MOIS (AddCompositeToMois below).

                With ssDrillData
                    .Row = 2
                    .Col = 2
                    OvbThk = .Value 'This is the depth to the top of the 1st split
                    'and will be the Ovb thk for an unmineable hole.
                End With

                gSetOverrideZeros(ProspData, OvbThk)
            End If

            Status = AddCompositeToMois(ProspData,
                                    UserSetHoleUnmineable,
                                    StandardsToSave)

            'Add the splits -- the splits will be saved one at a time to Oracle.
            With ssSplitReview
                For RowIdx = 1 To .MaxRows
                    SetActionStatus("Saving split #" & CStr(RowIdx) & "...")
                    ProspData = gGetDataFromReviewSprd(ssSplitReview, RowIdx)
                    Status = AddSplitToMois(ProspData, StandardsToSave)
                Next RowIdx
            End With

            SetActionStatus("")
            Me.Cursor = Cursors.Arrow

            MsgBox("Hole/Split saved to MOIS, minabilities saved to raw prospect.", vbOKOnly, "Save Status")


        Catch ex As Exception
            MessageBox.Show("Error saving: " & ex.Message)
        End Try

    End Sub

    Private Function AddCompositeToMois(ByRef aProspData As gRawProspSplRdctnType,
                                        ByVal aUserSetHoleUnmineable As Boolean,
                                        ByVal aStandardsToSave As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo AddCompositeToMoisError

        'IMPORTANT: Adding a composite hole will also delete any and all
        '           existing splits for that composite hole!!

        '01/06/2010, lss
        'aStandardsToSave
        'Now can add to "100% Prospect Only", "Catalog Only" or
        '               "100% Prospect and Catalog"

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim HoleLocAlpha As String
        Dim HoleLoc As String
        Dim ThisProspStandard As String
        Dim MtxWetDensity As Single
        Dim PitBottomElevation As Single
        Dim ProspGridType As String
        Dim HasCtlgReserves As Boolean
        Dim ProspSets As Integer
        Dim DataIdx As Integer
        Dim SplitsSummarized As String
        Dim TotNumSplits As Integer
        Dim RowIdx As Integer
        Dim SplitThkSum As Single
        Dim OkToContinue As Boolean

        If aUserSetHoleUnmineable = True Then
            SplitThkSum = 0
            With ssDrillData
                For RowIdx = 3 To .MaxRows
                    .Row = RowIdx
                    .Col = 2
                    SplitThkSum = SplitThkSum + .Value
                Next RowIdx
            End With
        End If

        'Need to have a mine name to add it to MOIS composite/split.
        If cboMineName.Text = "(Select mine...)" Then
            MsgBox("You must select a mine!", vbOKOnly, "Missing Mine Name")
            AddCompositeToMois = False
            Exit Function
        End If

        '11/15/2007, lss
        'This functionality is not available for South Fort Meade and Hookers Prairie
        'until they are merged into SurvCADD!  Currently they are transferred into
        'Geocomp and then loaded into MOIS from a text file that Geocomp creates.

        '04/23/2008, lss
        'SurvCADD is now being used for South Fort Meade and Hookers Prairie.
        'If cboMineName.Text = "South Fort Meade" Or cboMineName.Text = "Hookers Prairie" Then
        '    MsgBox "Not available for South Fort Meade or Hookers Prairie!", vbOKOnly, "Illegal Mine Name"
        '    AddCompositeToMois = False
        '    Exit Function
        'End If

        TotNumSplits = ssSplitReview.MaxRows
        SplitsSummarized = ""
        If aUserSetHoleUnmineable = False Then
            With ssDrillData
                For RowIdx = 3 To .MaxRows
                    .Row = RowIdx
                    .Col = 1
                    If .Text = "1" Then
                        'This split is indicated as minable -- add to list.
                        SplitsSummarized = SplitsSummarized & CStr(RowIdx - 2) & " "
                    End If
                Next RowIdx
            End With
            SplitsSummarized = Trim(SplitsSummarized)
        Else
            SplitsSummarized = ""
        End If

        'If we don't have an alpha-numeric hole location then don't add it
        'to the database at this time!
        'aProspData.HoleLocation is numeric -- may need an alpha-numeric hole
        'location!
        ProspGridType = gGetProspGridType(cboMineName.Text)

        'Add new composite prospect data
        If ProspGridType = "Alpha-numeric" Then
            'Need to get the alpha-numeric hole location.
            'If it won't translate will get "???".
            HoleLocAlpha = gGetHoleLoc2(aProspData.HoleLocation, "Char")

            If HoleLocAlpha = "???" Then
                'Cannot transfer this hole into MOIS right now!
                MsgBox("Cannot load this hole into MOIS (won't fit the alpha-numeric hole location scheme)!",
                       vbOKOnly, "Illegal Hole Location")
                AddCompositeToMois = False
                Exit Function
            End If
        End If

        HasCtlgReserves = gGetHasCtlgReserves(cboMineName.Text)
        If HasCtlgReserves = True Then
            ProspSets = 2
        Else
            ProspSets = 1
        End If

        'All mines have 100% prospect on the MOIS comp/split side
        'Only some mines currently have Catalog reserves as well
        '(South Fort Meade, Hookers Prairie and Wingate do not).

        For DataIdx = 1 To ProspSets
            'ProspSets will be 1 or 2
            'DataIdx = 1  Saving to 100% Prospect
            'DataIdx = 2  Saving to Catalog

            OkToContinue = False
            If DataIdx = 1 Then
                'Going to save 100% Prospect
                If aStandardsToSave = "100% Prospect Only" Or
                    aStandardsToSave = "100% Prospect and Catalog" Then
                    OkToContinue = True
                Else
                    OkToContinue = False
                End If
            End If
            If DataIdx = 2 Then
                'Going to save Catalog
                If aStandardsToSave = "Catalog Only" Or
                    aStandardsToSave = "100% Prospect and Catalog" Then
                    OkToContinue = True
                Else
                    OkToContinue = False
                End If
            End If

            If OkToContinue = True Then
                With aProspData
                    'Need to create a wet density from a dry density.
                    If .MtxPctSol <> 0 Then
                        MtxWetDensity = Round(.MtxDensity / (.MtxPctSol / 100), 2)
                    Else
                        MtxWetDensity = 0
                    End If

                    'NOTE: The interburden thickness is in both .ItbThk and .WstThk.
                    If aUserSetHoleUnmineable = False Then
                        PitBottomElevation = .Elevation - .OvbThk - .MtxThk - .WstThk
                    Else
                        PitBottomElevation = .Elevation - .OvbThk - SplitThkSum
                    End If

                    If PitBottomElevation < 0 Then
                        PitBottomElevation = 0
                    End If

                    params = gDBParams

                    params.Add("pMineName", cboMineName.Text, ORAPARM_INPUT)
                    params("pMineName").serverType = ORATYPE_VARCHAR2

                    params.Add("pTownShip", .Township, ORAPARM_INPUT)
                    params("pTownShip").serverType = ORATYPE_NUMBER

                    params.Add("pRange", .Range, ORAPARM_INPUT)
                    params("pRange").serverType = ORATYPE_NUMBER

                    params.Add("pSection", .Section, ORAPARM_INPUT)
                    params("pSection").serverType = ORATYPE_NUMBER

                    params.Add("pXSPCoordinate", .Xcoord, ORAPARM_INPUT)
                    params("pXSPCoordinate").serverType = ORATYPE_NUMBER

                    params.Add("pYSPCoordinate", .Ycoord, ORAPARM_INPUT)
                    params("pYSPCoordinate").serverType = ORATYPE_NUMBER

                    If ProspGridType = "Alpha-numeric" Then
                        HoleLoc = HoleLocAlpha
                    Else    'ProspGridType = "Numeric"
                        HoleLoc = .HoleLocation
                    End If

                    params.Add("pHoleLocation", HoleLoc, ORAPARM_INPUT)
                    params("pHoleLocation").serverType = ORATYPE_VARCHAR2

                    params.Add("pDrillDate", .ProspDate, ORAPARM_INPUT)
                    params("pDrillDate").serverType = ORATYPE_VARCHAR2

                    'Not really even sure what this date is in GEOCOMP.
                    'It is the "Date of analysis" date sometimes called the "Report" date.
                    'Will put the prospect date in here for now -- need to put
                    'the raw prospect wash date in here eventually!
                    params.Add("pWashDate", .ProspDate, ORAPARM_INPUT)
                    params("pWashDate").serverType = ORATYPE_VARCHAR2

                    params.Add("pAreaOfInfluence", .Aoi, ORAPARM_INPUT)
                    params("pAreaOfInfluence").serverType = ORATYPE_NUMBER

                    params.Add("pOvbThickness", .OvbThk, ORAPARM_INPUT)
                    params("pOvbThickness").serverType = ORATYPE_NUMBER

                    params.Add("pMtxThickness", .MtxThk, ORAPARM_INPUT)
                    params("pMtxThickness").serverType = ORATYPE_NUMBER

                    params.Add("pMtxWetDensity", MtxWetDensity, ORAPARM_INPUT)
                    params("pMtxWetDensity").serverType = ORATYPE_NUMBER

                    params.Add("pMtxPercentSolids", .MtxPctSol, ORAPARM_INPUT)
                    params("pMtxPercentSolids").serverType = ORATYPE_NUMBER

                    params.Add("pMtxBPL", .MtxBPL, ORAPARM_INPUT)
                    params("pMtxBPL").serverType = ORATYPE_NUMBER

                    '----------

                    If DataIdx = 1 Then
                        '100%  100%  100%  100%  100%  100%
                        '100%  100%  100%  100%  100%  100%
                        '100%  100%  100%  100%  100%  100%

                        params.Add("pMtxTons", .MtxTPA, ORAPARM_INPUT)
                        params("pMtxTons").serverType = ORATYPE_NUMBER

                        'Coarse pebble

                        params.Add("pCpWeightPercent", .Cpb100.WtPct, ORAPARM_INPUT)
                        params("pCpWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pCpTonsPerAcre", .Cpb100.Tpa, ORAPARM_INPUT)
                        params("pCpTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pCpBPL", .Cpb100.Bpl, ORAPARM_INPUT)
                        params("pCpBPL").serverType = ORATYPE_NUMBER

                        params.Add("pCpInsol", .Cpb100.Ins, ORAPARM_INPUT)
                        params("pCpInsol").serverType = ORATYPE_NUMBER

                        params.Add("pCpFe2O3", .Cpb100.Fe, ORAPARM_INPUT)
                        params("pCpFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pCpAl2O3", .Cpb100.Al, ORAPARM_INPUT)
                        params("pCpAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pCpMgO", .Cpb100.Mg, ORAPARM_INPUT)
                        params("pCpMgO").serverType = ORATYPE_NUMBER

                        params.Add("pCpCaO", .Cpb100.Ca, ORAPARM_INPUT)
                        params("pCpCaO").serverType = ORATYPE_NUMBER

                        '----------
                        'Fine pebble

                        params.Add("pFpWeightPercent", .Fpb100.WtPct, ORAPARM_INPUT)
                        params("pFpWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pFpTonsPerAcre", .Fpb100.Tpa, ORAPARM_INPUT)
                        params("pFpTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pFpBPL", .Fpb100.Bpl, ORAPARM_INPUT)
                        params("pFpBPL").serverType = ORATYPE_NUMBER

                        params.Add("pFpInsol", .Fpb100.Ins, ORAPARM_INPUT)
                        params("pFpInsol").serverType = ORATYPE_NUMBER

                        params.Add("pFpFe2O3", .Fpb100.Fe, ORAPARM_INPUT)
                        params("pFpFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pFpAl2O3", .Fpb100.Al, ORAPARM_INPUT)
                        params("pFpAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pFpMgO", .Fpb100.Mg, ORAPARM_INPUT)
                        params("pFpMgO").serverType = ORATYPE_NUMBER

                        params.Add("pFpCaO", .Fpb100.Ca, ORAPARM_INPUT)
                        params("pFpCaO").serverType = ORATYPE_NUMBER

                        '----------

                        params.Add("pTfWeightPercent", .Tfd100.WtPct, ORAPARM_INPUT)
                        params("pTfWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pTfTonsPerAcre", .Tfd100.Tpa, ORAPARM_INPUT)
                        params("pTfTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pTfBPL", .Tfd100.Bpl, ORAPARM_INPUT)
                        params("pTfBPL").serverType = ORATYPE_NUMBER

                        '----------

                        params.Add("pWcWeightPercent", .Wcl100.WtPct, ORAPARM_INPUT)
                        params("pWcWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pWcTonsPerAcre", .Wcl100.Tpa, ORAPARM_INPUT)
                        params("pWcTonsPerAcre").serverType = ORATYPE_NUMBER

                        '----------

                        params.Add("pCfBPL", .Cfd100.Bpl, ORAPARM_INPUT)
                        params("pCfBPL").serverType = ORATYPE_NUMBER

                        params.Add("pFfBPL", .Ffd100.Bpl, ORAPARM_INPUT)
                        params("pFfBPL").serverType = ORATYPE_NUMBER

                        params.Add("pCfTonsPerAcre", .Cfd100.Tpa, ORAPARM_INPUT)
                        params("pCfTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pFfTonsPerAcre", .Ffd100.Tpa, ORAPARM_INPUT)
                        params("pFfTonsPerAcre").serverType = ORATYPE_NUMBER

                        '----------
                        'Concentrate

                        'Have to be careful here -- If we have a hole that has
                        'used the Off-spec pebble MgO plant then we have something
                        'different to send to MOIS than the "normal" float plant
                        'concentrate!

                        If lblOffSpecPbMgPlt.Text <> "*OffSpec Pb Mg Plt*" Then
                            params.Add("pCnWeightPercent", .Tcn100.WtPct, ORAPARM_INPUT)
                            params("pCnWeightPercent").serverType = ORATYPE_NUMBER

                            params.Add("pCnTonsPerAcre", .Tcn100.Tpa, ORAPARM_INPUT)
                            params("pCnTonsPerAcre").serverType = ORATYPE_NUMBER

                            params.Add("pCnBPL", .Tcn100.Bpl, ORAPARM_INPUT)
                            params("pCnBPL").serverType = ORATYPE_NUMBER

                            params.Add("pCnInsol", .Tcn100.Ins, ORAPARM_INPUT)
                            params("pCnInsol").serverType = ORATYPE_NUMBER

                            params.Add("pCnFe2O3", .Tcn100.Fe, ORAPARM_INPUT)
                            params("pCnFe2O3").serverType = ORATYPE_NUMBER

                            params.Add("pCnAl2O3", .Tcn100.Al, ORAPARM_INPUT)
                            params("pCnAl2O3").serverType = ORATYPE_NUMBER

                            params.Add("pCnMgO", .Tcn100.Mg, ORAPARM_INPUT)
                            params("pCnMgO").serverType = ORATYPE_NUMBER

                            params.Add("pCnCaO", .Tcn100.Ca, ORAPARM_INPUT)
                            params("pCnCaO").serverType = ORATYPE_NUMBER
                        Else
                            params.Add("pCnWeightPercent", .MgPltTcn100.WtPct, ORAPARM_INPUT)
                            params("pCnWeightPercent").serverType = ORATYPE_NUMBER

                            params.Add("pCnTonsPerAcre", .MgPltTcn100.Tpa, ORAPARM_INPUT)
                            params("pCnTonsPerAcre").serverType = ORATYPE_NUMBER

                            params.Add("pCnBPL", .MgPltTcn100.Bpl, ORAPARM_INPUT)
                            params("pCnBPL").serverType = ORATYPE_NUMBER

                            params.Add("pCnInsol", .MgPltTcn100.Ins, ORAPARM_INPUT)
                            params("pCnInsol").serverType = ORATYPE_NUMBER

                            params.Add("pCnFe2O3", .MgPltTcn100.Fe, ORAPARM_INPUT)
                            params("pCnFe2O3").serverType = ORATYPE_NUMBER

                            params.Add("pCnAl2O3", .MgPltTcn100.Al, ORAPARM_INPUT)
                            params("pCnAl2O3").serverType = ORATYPE_NUMBER

                            params.Add("pCnMgO", .MgPltTcn100.Mg, ORAPARM_INPUT)
                            params("pCnMgO").serverType = ORATYPE_NUMBER

                            params.Add("pCnCaO", .MgPltTcn100.Ca, ORAPARM_INPUT)
                            params("pCnCaO").serverType = ORATYPE_NUMBER
                        End If
                        '----------

                        params.Add("pTpWeightPercent", .Tpr100.WtPct, ORAPARM_INPUT)
                        params("pTpWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pTpTonsPerAcre", .Tpr100.Tpa, ORAPARM_INPUT)
                        params("pTpTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pTpBPL", .Tpr100.Bpl, ORAPARM_INPUT)
                        params("pTpBPL").serverType = ORATYPE_NUMBER

                        params.Add("pTpInsol", .Tpr100.Ins, ORAPARM_INPUT)
                        params("pTpInsol").serverType = ORATYPE_NUMBER

                        params.Add("pTpFe2O3", .Tpr100.Fe, ORAPARM_INPUT)
                        params("pTpFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pTpAl2O3", .Tpr100.Al, ORAPARM_INPUT)
                        params("pTpAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pTpMgO", .Tpr100.Mg, ORAPARM_INPUT)
                        params("pTpMgO").serverType = ORATYPE_NUMBER

                        params.Add("pTpCaO", .Tpr100.Ca, ORAPARM_INPUT)
                        params("pTpCaO").serverType = ORATYPE_NUMBER

                        '----------

                        params.Add("pMtxX", .MtxxAll100Hole, ORAPARM_INPUT)
                        params("pMtxX").serverType = ORATYPE_NUMBER
                    Else    'DataIdx = 2
                        'Catalog  Catalog  Catalog  Catalog
                        'Catalog  Catalog  Catalog  Catalog
                        'Catalog  Catalog  Catalog  Catalog

                        params.Add("pMtxTons", .MtxTpaPc, ORAPARM_INPUT)
                        params("pMtxTons").serverType = ORATYPE_NUMBER

                        params.Add("pCpWeightPercent", .Cpb.WtPct, ORAPARM_INPUT)
                        params("pCpWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pCpTonsPerAcre", .Cpb.Tpa, ORAPARM_INPUT)
                        params("pCpTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pCpBPL", .Cpb.Bpl, ORAPARM_INPUT)
                        params("pCpBPL").serverType = ORATYPE_NUMBER

                        params.Add("pCpInsol", .Cpb.Ins, ORAPARM_INPUT)
                        params("pCpInsol").serverType = ORATYPE_NUMBER

                        params.Add("pCpFe2O3", .Cpb.Fe, ORAPARM_INPUT)
                        params("pCpFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pCpAl2O3", .Cpb.Al, ORAPARM_INPUT)
                        params("pCpAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pCpMgO", .Cpb.Mg, ORAPARM_INPUT)
                        params("pCpMgO").serverType = ORATYPE_NUMBER

                        params.Add("pCpCaO", .Cpb.Ca, ORAPARM_INPUT)
                        params("pCpCaO").serverType = ORATYPE_NUMBER

                        '----------

                        params.Add("pFpWeightPercent", .Fpb.WtPct, ORAPARM_INPUT)
                        params("pFpWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pFpTonsPerAcre", .Fpb.Tpa, ORAPARM_INPUT)
                        params("pFpTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pFpBPL", .Fpb.Bpl, ORAPARM_INPUT)
                        params("pFpBPL").serverType = ORATYPE_NUMBER

                        params.Add("pFpInsol", .Fpb.Ins, ORAPARM_INPUT)
                        params("pFpInsol").serverType = ORATYPE_NUMBER

                        params.Add("pFpFe2O3", .Fpb.Fe, ORAPARM_INPUT)
                        params("pFpFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pFpAl2O3", .Fpb.Al, ORAPARM_INPUT)
                        params("pFpAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pFpMgO", .Fpb.Mg, ORAPARM_INPUT)
                        params("pFpMgO").serverType = ORATYPE_NUMBER

                        params.Add("pFpCaO", .Fpb.Ca, ORAPARM_INPUT)
                        params("pFpCaO").serverType = ORATYPE_NUMBER

                        '----------

                        params.Add("pTfWeightPercent", .Tfd.WtPct, ORAPARM_INPUT)
                        params("pTfWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pTfTonsPerAcre", .Tfd.Tpa, ORAPARM_INPUT)
                        params("pTfTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pTfBPL", .Tfd.Bpl, ORAPARM_INPUT)
                        params("pTfBPL").serverType = ORATYPE_NUMBER

                        '----------

                        params.Add("pWcWeightPercent", .Wcl.WtPct, ORAPARM_INPUT)
                        params("pWcWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pWcTonsPerAcre", .Wcl.Tpa, ORAPARM_INPUT)
                        params("pWcTonsPerAcre").serverType = ORATYPE_NUMBER

                        '----------

                        params.Add("pCfBPL", .Cfd.Bpl, ORAPARM_INPUT)
                        params("pCfBPL").serverType = ORATYPE_NUMBER

                        params.Add("pFfBPL", .Ffd.Bpl, ORAPARM_INPUT)
                        params("pFfBPL").serverType = ORATYPE_NUMBER

                        params.Add("pCfTonsPerAcre", .Cfd.Tpa, ORAPARM_INPUT)
                        params("pCfTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pFfTonsPerAcre", .Ffd.Tpa, ORAPARM_INPUT)
                        params("pFfTonsPerAcre").serverType = ORATYPE_NUMBER

                        '----------
                        If lblOffSpecPbMgPlt.Text <> "*OffSpec Pb Mg Plt*" Then
                            params.Add("pCnWeightPercent", .Tcn.WtPct, ORAPARM_INPUT)
                            params("pCnWeightPercent").serverType = ORATYPE_NUMBER

                            params.Add("pCnTonsPerAcre", .Tcn.Tpa, ORAPARM_INPUT)
                            params("pCnTonsPerAcre").serverType = ORATYPE_NUMBER

                            params.Add("pCnBPL", .Tcn.Bpl, ORAPARM_INPUT)
                            params("pCnBPL").serverType = ORATYPE_NUMBER

                            params.Add("pCnInsol", .Tcn.Ins, ORAPARM_INPUT)
                            params("pCnInsol").serverType = ORATYPE_NUMBER

                            params.Add("pCnFe2O3", .Tcn.Fe, ORAPARM_INPUT)
                            params("pCnFe2O3").serverType = ORATYPE_NUMBER

                            params.Add("pCnAl2O3", .Tcn.Al, ORAPARM_INPUT)
                            params("pCnAl2O3").serverType = ORATYPE_NUMBER

                            params.Add("pCnMgO", .Tcn.Mg, ORAPARM_INPUT)
                            params("pCnMgO").serverType = ORATYPE_NUMBER

                            params.Add("pCnCaO", .Tcn.Ca, ORAPARM_INPUT)
                            params("pCnCaO").serverType = ORATYPE_NUMBER
                        Else
                            params.Add("pCnWeightPercent", .MgPltTcn.WtPct, ORAPARM_INPUT)
                            params("pCnWeightPercent").serverType = ORATYPE_NUMBER

                            params.Add("pCnTonsPerAcre", .MgPltTcn.Tpa, ORAPARM_INPUT)
                            params("pCnTonsPerAcre").serverType = ORATYPE_NUMBER

                            params.Add("pCnBPL", .MgPltTcn.Bpl, ORAPARM_INPUT)
                            params("pCnBPL").serverType = ORATYPE_NUMBER

                            params.Add("pCnInsol", .MgPltTcn.Ins, ORAPARM_INPUT)
                            params("pCnInsol").serverType = ORATYPE_NUMBER

                            params.Add("pCnFe2O3", .MgPltTcn.Fe, ORAPARM_INPUT)
                            params("pCnFe2O3").serverType = ORATYPE_NUMBER

                            params.Add("pCnAl2O3", .MgPltTcn.Al, ORAPARM_INPUT)
                            params("pCnAl2O3").serverType = ORATYPE_NUMBER

                            params.Add("pCnMgO", .MgPltTcn.Mg, ORAPARM_INPUT)
                            params("pCnMgO").serverType = ORATYPE_NUMBER

                            params.Add("pCnCaO", .MgPltTcn.Ca, ORAPARM_INPUT)
                            params("pCnCaO").serverType = ORATYPE_NUMBER
                        End If

                        '----------

                        params.Add("pTpWeightPercent", .Tpr.WtPct, ORAPARM_INPUT)
                        params("pTpWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pTpTonsPerAcre", .Tpr.Tpa, ORAPARM_INPUT)
                        params("pTpTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pTpBPL", .Tpr.Bpl, ORAPARM_INPUT)
                        params("pTpBPL").serverType = ORATYPE_NUMBER

                        params.Add("pTpInsol", .Tpr.Ins, ORAPARM_INPUT)
                        params("pTpInsol").serverType = ORATYPE_NUMBER

                        params.Add("pTpFe2O3", .Tpr.Fe, ORAPARM_INPUT)
                        params("pTpFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pTpAl2O3", .Tpr.Al, ORAPARM_INPUT)
                        params("pTpAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pTpMgO", .Tpr.Mg, ORAPARM_INPUT)
                        params("pTpMgO").serverType = ORATYPE_NUMBER

                        params.Add("pTpCaO", .Tpr.Ca, ORAPARM_INPUT)
                        params("pTpCaO").serverType = ORATYPE_NUMBER

                        '----------

                        params.Add("pMtxX", .MtxxAllPcHole, ORAPARM_INPUT)
                        params("pMtxX").serverType = ORATYPE_NUMBER
                    End If

                    params.Add("pHoleElevation", .Elevation, ORAPARM_INPUT)
                    params("pHoleElevation").serverType = ORATYPE_NUMBER

                    params.Add("pTotalNumberSplits", TotNumSplits, ORAPARM_INPUT)
                    params("pTotalNumberSplits").serverType = ORATYPE_NUMBER

                    params.Add("pPitBottomElevation", PitBottomElevation, ORAPARM_INPUT)
                    params("pPitBottomElevation").serverType = ORATYPE_NUMBER

                    'Don't have a triangle code.
                    params.Add("pTriangleCode", "", ORAPARM_INPUT)
                    params("pTriangleCode").serverType = ORATYPE_VARCHAR2

                    'Don't have a prospector code.
                    params.Add("pProspectorCode", 0, ORAPARM_INPUT)
                    params("pProspectorCode").serverType = ORATYPE_VARCHAR2

                    params.Add("pSplitsSummarized", SplitsSummarized, ORAPARM_INPUT)
                    params("pSplitsSummarized").serverType = ORATYPE_VARCHAR2

                    '----------

                    If DataIdx = 1 Then
                        '100%  100%  100%  100%  100%  100%
                        '100%  100%  100%  100%  100%  100%
                        '100%  100%  100%  100%  100%  100%

                        params.Add("pTpbWeightPercent", .Tpb100.WtPct, ORAPARM_INPUT)
                        params("pTpbWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pTpbTonsPerAcre", .Tpb100.Tpa, ORAPARM_INPUT)
                        params("pTpbTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pTpbBPL", .Tpb100.Bpl, ORAPARM_INPUT)
                        params("pTpbBPL").serverType = ORATYPE_NUMBER

                        params.Add("pTpbInsol", .Tpb100.Ins, ORAPARM_INPUT)
                        params("pTpbInsol").serverType = ORATYPE_NUMBER

                        params.Add("pTpbFe2O3", .Tpb100.Fe, ORAPARM_INPUT)
                        params("pTpbFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pTpbAl2O3", .Tpb100.Al, ORAPARM_INPUT)
                        params("pTpbAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pTpbMgO", .Tpb100.Mg, ORAPARM_INPUT)
                        params("pTpbMgO").serverType = ORATYPE_NUMBER

                        params.Add("pTpbCaO", .Tpb100.Ca, ORAPARM_INPUT)
                        params("pTpbCaO").serverType = ORATYPE_NUMBER
                    Else    'DataIdx = 2
                        'Catalog  Catalog  Catalog  Catalog
                        'Catalog  Catalog  Catalog  Catalog
                        'Catalog  Catalog  Catalog  Catalog

                        params.Add("pTpbWeightPercent", .Tpb.WtPct, ORAPARM_INPUT)
                        params("pTpbWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pTpbTonsPerAcre", .Tpb.Tpa, ORAPARM_INPUT)
                        params("pTpbTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pTpbBPL", .Tpb.Bpl, ORAPARM_INPUT)
                        params("pTpbBPL").serverType = ORATYPE_NUMBER

                        params.Add("pTpbInsol", .Tpb.Ins, ORAPARM_INPUT)
                        params("pTpbInsol").serverType = ORATYPE_NUMBER

                        params.Add("pTpbFe2O3", .Tpb.Fe, ORAPARM_INPUT)
                        params("pTpbFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pTpbAl2O3", .Tpb.Al, ORAPARM_INPUT)
                        params("pTpbAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pTpbMgO", .Tpb.Mg, ORAPARM_INPUT)
                        params("pTpbMgO").serverType = ORATYPE_NUMBER

                        params.Add("pTpbCaO", .Tpb.Ca, ORAPARM_INPUT)
                        params("pTpbCaO").serverType = ORATYPE_NUMBER
                    End If

                    '----------
                    'NOTE: The interburden thickness is in both .ItbThk and .WstThk.
                    params.Add("pWstThck", .WstThk, ORAPARM_INPUT)
                    params("pWstThck").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pTotX", 0, ORAPARM_INPUT)
                    params("pTotX").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pMinableSplits", "", ORAPARM_INPUT)
                    params("pMinableSplits").serverType = ORATYPE_VARCHAR2

                    'Pioneer
                    params.Add("pHoleMinable", "", ORAPARM_INPUT)
                    params("pHoleMinable").serverType = ORATYPE_VARCHAR2

                    'Pioneer
                    params.Add("pCpbMinable", "", ORAPARM_INPUT)
                    params("pCpbMinable").serverType = ORATYPE_VARCHAR2

                    'Pioneer
                    params.Add("pFpbMinable", "", ORAPARM_INPUT)
                    params("pFpbMinable").serverType = ORATYPE_VARCHAR2

                    'Pioneer
                    params.Add("pCpbFeTpaWt", 0, ORAPARM_INPUT)
                    params("pCpbFeTpaWt").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pCpbAlTpaWt", 0, ORAPARM_INPUT)
                    params("pCpbAlTpaWt").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pCpbIaTpaWt", 0, ORAPARM_INPUT)
                    params("pCpbIaTpaWt").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pCpbCaTpaWt", 0, ORAPARM_INPUT)
                    params("pCpbCaTpaWt").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pFpbFeTpaWt", 0, ORAPARM_INPUT)
                    params("pFpbFeTpaWt").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pFpbAlTpaWt", 0, ORAPARM_INPUT)
                    params("pFpbAlTpaWt").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pFpbIaTpaWt", 0, ORAPARM_INPUT)
                    params("pFpbIaTpaWt").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pFpbCaTpaWt", 0, ORAPARM_INPUT)
                    params("pFpbCaTpaWt").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pCnFeTpaWt", 0, ORAPARM_INPUT)
                    params("pCnFeTpaWt").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pCnAlTpaWt", 0, ORAPARM_INPUT)
                    params("pCnAlTpaWt").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pCnIaTpaWt", 0, ORAPARM_INPUT)
                    params("pCnIaTpaWt").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pCnCaTpaWt", 0, ORAPARM_INPUT)
                    params("pCnCaTpaWt").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pCpIa", 0, ORAPARM_INPUT)
                    params("pCpIa").serverType = ORATYPE_NUMBER

                    params.Add("pFpIa", 0, ORAPARM_INPUT)
                    params("pFpIa").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pCnIa", 0, ORAPARM_INPUT)
                    params("pCnIa").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pTpIa", 0, ORAPARM_INPUT)
                    params("pTpIa").serverType = ORATYPE_NUMBER

                    'Pioneer
                    params.Add("pTpbIa", 0, ORAPARM_INPUT)
                    params("pTpbIa").serverType = ORATYPE_NUMBER

                    '----------
                    'Don't have Cadmium as an analyte in the metlab.

                    params.Add("pCpbCd", 0, ORAPARM_INPUT)
                    params("pCpbCd").serverType = ORATYPE_NUMBER

                    params.Add("pFpbCd", 0, ORAPARM_INPUT)
                    params("pFpbCd").serverType = ORATYPE_NUMBER

                    params.Add("pTcnCd", 0, ORAPARM_INPUT)
                    params("pTcnCd").serverType = ORATYPE_NUMBER

                    params.Add("pTprCd", 0, ORAPARM_INPUT)
                    params("pTprCd").serverType = ORATYPE_NUMBER

                    params.Add("pTpbCd", 0, ORAPARM_INPUT)
                    params("pTpbCd").serverType = ORATYPE_NUMBER

                    '----------

                    'Don't have a hardpan code.
                    params.Add("pHardpanCode", 0, ORAPARM_INPUT)
                    params("pHardpanCode").serverType = ORATYPE_NUMBER

                    '----------

                    If DataIdx = 1 Then
                        ThisProspStandard = "100%PROSPECT"
                    End If
                    If DataIdx = 2 Then
                        ThisProspStandard = "CATALOG"
                    End If
                    params.Add("pProspStandard", ThisProspStandard, ORAPARM_INPUT)
                    params("pProspStandard").serverType = ORATYPE_VARCHAR2

                    params.Add("pMinedStatus", 0, ORAPARM_INPUT)
                    params("pMinedStatus").serverType = ORATYPE_NUMBER
                End With

                'WARNING  Procedure load_composite will delete this hole (including the splits)
                '         if it already exists in MOIS!

                'Procedure load_composite
                'pMineName           IN     VARCHAR2,    -- 1
                'pTownship           IN     NUMBER,      -- 2
                'pRange              IN     NUMBER,      -- 3
                'pSection            IN     NUMBER,      -- 4
                'pXSPCoordinate      IN     NUMBER,      -- 5
                'pYSPCoordinate      IN     NUMBER,      -- 6
                'pHoleLocation       IN     VARCHAR2,    -- 7
                'pDrillDate          IN     VARCHAR2,    -- 8
                'pWashDate           IN     VARCHAR2,    -- 9
                'pAreaOfInfluence    IN     NUMBER,      -- 10
                'pOvbThickness       IN     NUMBER,      -- 11
                'pMtxThickness       IN     NUMBER,      -- 12
                'pMtxWetDensity      IN     NUMBER,      -- 13
                'pMtxPercentSolids   IN     NUMBER,      -- 14
                'pMtxTons            IN     NUMBER,      -- 15
                'pMtxBPL             IN     NUMBER,      -- 16
                'pCpWeightPercent    IN     NUMBER,      -- 17
                'pCpTonsPerAcre      IN     NUMBER,      -- 18
                'pCpBPL              IN     NUMBER,      -- 19
                'pCpInsol            IN     NUMBER,      -- 20
                'pCpFe2O3            IN     NUMBER,      -- 21
                'pCpAl2O3            IN     NUMBER,      -- 22
                'pCpMgO              IN     NUMBER,      -- 23
                'pCpCaO              IN     NUMBER,      -- 24
                'pFpWeightPercent    IN     NUMBER,      -- 25
                'pFpTonsPerAcre      IN     NUMBER,      -- 26
                'pFpBPL              IN     NUMBER,      -- 27
                'pFpInsol            IN     NUMBER,      -- 28
                'pFpFe2O3            IN     NUMBER,      -- 29
                'pFpAl2O3            IN     NUMBER,      -- 30
                'pFpMgO              IN     NUMBER,      -- 31
                'pFpCaO              IN     NUMBER,      -- 32
                'pTfWeightPercent    IN     NUMBER,      -- 33
                'pTfTonsPerAcre      IN     NUMBER,      -- 34
                'pTfBPL              IN     NUMBER,      -- 35
                'pWcWeightPercent    IN     NUMBER,      -- 36
                'pWcTonsPerAcre      IN     NUMBER,      -- 37
                'pCfBPL              IN     NUMBER,      -- 38
                'pFfBPL              IN     NUMBER,      -- 39
                'pCfTonsPerAcre      IN     NUMBER,      -- 40
                'pFfTonsPerAcre      IN     NUMBER,      -- 41
                'pCnWeightPercent    IN     NUMBER,      -- 42
                'pCnTonsPerAcre      IN     NUMBER,      -- 43
                'pCnBPL              IN     NUMBER,      -- 44
                'pCnInsol            IN     NUMBER,      -- 45
                'pCnFe2O3            IN     NUMBER,      -- 46
                'pCnAl2O3            IN     NUMBER,      -- 47
                'pCnMgO              IN     NUMBER,      -- 48
                'pCnCaO              IN     NUMBER,      -- 49
                'pTpWeightPercent    IN     NUMBER,      -- 50
                'pTpTonsPerAcre      IN     NUMBER,      -- 51
                'pTpBPL              IN     NUMBER,      -- 52
                'pTpInsol            IN     NUMBER,      -- 53
                'pTpFe2O3            IN     NUMBER,      -- 54
                'pTpAl2O3            IN     NUMBER,      -- 55
                'pTpMgO              IN     NUMBER,      -- 56
                'pTpCaO              IN     NUMBER,      -- 57
                'pMtxX               IN     NUMBER,      -- 58
                'pHoleElevation      IN     NUMBER,      -- 59
                'pTotalNumberSplits  IN     NUMBER,      -- 60
                'pPitBottomElevation IN     NUMBER,      -- 61
                'pTriangleCode       IN     VARCHAR2,    -- 62
                'pProspectorCode     IN     VARCHAR2,    -- 63
                'pSplitsSummarized   IN     VARCHAR2,    -- 64
                '----------
                'pTpbWeightPercent   IN     NUMBER,      -- 65
                'pTpbTonsPerAcre     IN     NUMBER,      -- 66
                'pTpbBPL             IN     NUMBER,      -- 67
                'pTpbInsol           IN     NUMBER,      -- 68
                'pTpbFe2O3           IN     NUMBER,      -- 69
                'pTpbAl2O3           IN     NUMBER,      -- 70
                'pTpbMgO             IN     NUMBER,      -- 71
                'pTpbCaO             IN     NUMBER,      -- 72
                '----------
                'pWstThck            IN     NUMBER,      -- 73
                'pTotX               IN     NUMBER,      -- 74
                'pMinableSplits      IN     VARCHAR2,    -- 75
                'pHoleMinable        IN     VARCHAR2,    -- 76
                'pCpbMinable         IN     VARCHAR2,    -- 77
                'pFpbMinable         IN     VARCHAR2,    -- 78
                'pCpbFeTpaWt         IN     NUMBER,      -- 79
                'pCpbAlTpaWt         IN     NUMBER,      -- 80
                'pCpbIaTpaWt         IN     NUMBER,      -- 81
                'pCpbCaTpaWt         IN     NUMBER,      -- 82
                'pFpbFeTpaWt         IN     NUMBER,      -- 83
                'pFpbAlTpaWt         IN     NUMBER,      -- 84
                'pFpbIaTpaWt         IN     NUMBER,      -- 85
                'pFpbCaTpaWt         IN     NUMBER,      -- 86
                'pCnFeTpaWt          IN     NUMBER,      -- 87
                'pCnAlTpaWt          IN     NUMBER,      -- 88
                'pCnIaTpaWt          IN     NUMBER,      -- 89
                'pCnCaTpaWt          IN     NUMBER,      -- 90
                '----------
                'pCpIa               IN     NUMBER,      -- 91
                'pFpIa               IN     NUMBER,      -- 92
                'pCnIa               IN     NUMBER,      -- 93
                'pTpIa               IN     NUMBER,      -- 94
                'pTpbIa              IN     NUMBER,      -- 95
                '----------
                'pCpbCd              IN     NUMBER,      -- 96
                'pFpbCd              IN     NUMBER,      -- 97
                'pTcnCd              IN     NUMBER,      -- 98
                'pTprCd              IN     NUMBER,      -- 99
                'pTpbCd              IN     NUMBER,      -- 100
                '----------
                'pHardpanCode        IN     NUMBER,      -- 101
                'pProspStandard      IN     VARCHAR2,    -- 102
                'pMinedStatus        IN     NUMBER);     -- 103

                SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect.load_composite(:pMineName, :pTownship, :pRange, :pSection, :pXSPCoordinate," +
                              ":pYSPCoordinate, :pHoleLocation, :pDrillDate, :pWashDate, :pAreaOfInfluence," +
                              ":pOvbThickness, :pMtxThickness, :pMtxWetDensity, :pMtxPercentSolids, :pMtxTons," +
                              ":pMtxBPL, :pCpWeightPercent, :pCpTonsPerAcre, :pCpBPL, :pCpInsol," +
                              ":pCpFe2O3, :pCpAl2O3, :pCpMgO, :pCpCaO, :pFpWeightPercent," +
                              ":pFpTonsPerAcre, :pFpBPL, :pFpInsol, :pFpFe2O3, :pFpAl2O3," +
                              ":pFpMgO, :pFpCaO, :pTfWeightPercent, :pTfTonsPerAcre, :pTfBPL," +
                              ":pWcWeightPercent, :pWcTonsPerAcre, :pCfBPL, :pFfBPL, :pCfTonsPerAcre," +
                              ":pFfTonsPerAcre, :pCnWeightPercent, :pCnTonsPerAcre, :pCnBPL, :pCnInsol," +
                              ":pCnFe2O3, :pCnAl2O3, :pCnMgO, :pCnCaO, :pTpWeightPercent," +
                              ":pTpTonsPerAcre, :pTpBPL, :pTpInsol, :pTpFe2O3, :pTpAl2O3," +
                              ":pTpMgO, :pTpCaO, :pMtxX, :pHoleElevation, :pTotalNumberSplits," +
                              ":pPitBottomElevation, :pTriangleCode, :pProspectorCode," +
                              ":pSplitsSummarized, :pTpbWeightPercent, :pTpbTonsPerAcre," +
                              ":pTpbBPL, :pTpbInsol, :pTpbFe2O3, :pTpbAl2O3, :pTpbMgO, :pTpbCaO," +
                              ":pWstThck, :pTotX, :pMinableSplits, :pHoleMinable, :pCpbMinable, " +
                              ":pFpbMinable, :pCpbFeTpaWt, :pCpbAlTpaWt, :pCpbIaTpaWt, :pCpbCaTpaWt, " +
                              ":pFpbFeTpaWt, :pFpbAlTpaWt, :pFpbIaTpaWt, :pFpbCaTpaWt, " +
                              ":pCnFeTpaWt, :pCnAlTpaWt, :pCnIaTpaWt, :pCnCaTpaWt, " +
                              ":pCpIa, :pFpIa, :pCnIa, :pTpIa, :pTpbIa, " +
                              ":pCpbCd, :pFpbCd, :pTcnCd, :pTprCd, :pTpbCd, :pHardpanCode, :pProspStandard, :pMinedStatus);end;", ORASQL_FAILEXEC)

                ClearParams(params)
            End If
        Next DataIdx

        AddCompositeToMois = True

        Exit Function

AddCompositeToMoisError:
        MsgBox("Error saving composite." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Save Error")

        On Error Resume Next
        ClearParams(params)
    End Function

    Private Function AddSplitToMois(ByRef aProspData As gRawProspSplRdctnType,
                                    ByVal aStandardsToSave As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo AddSplitToMoisError

        '01/06/2010, lss
        'aStandardsToSave
        'Now can add to "100% Prospect Only", "Catalog Only" or
        '               "100% Prospect and Catalog"

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        Dim HasCtlgReserves As Boolean
        Dim ProspSets As Integer
        Dim DataIdx As Integer
        Dim HoleLoc As String
        Dim TotNumSplits As Integer
        Dim MinableStatus As String
        Dim ProspGridType As String
        Dim HoleLocAlpha As String
        Dim ThisProspStandard As String
        Dim MtxWetDensity As Single
        Dim CaP2O5Val As Single
        Dim SplitElev As Single
        Dim ThisColor As String
        Dim OkToContinue As Boolean

        TotNumSplits = ssSplitReview.MaxRows

        ProspGridType = gGetProspGridType(cboMineName.Text)

        ThisColor = gGetProspCodeDesc("Phosphate color", aProspData.PhosphColor)

        'Add new composite prospect data
        If ProspGridType = "Alpha-numeric" Then
            'Need to get the alpha-numeric hole location.
            'If it won't translate will get "???".
            HoleLocAlpha = gGetHoleLoc2(aProspData.HoleLocation, "Char")

            'Since we got this far (Hole has been saved to Oracle) we know that
            'the numeric hole location will translate into an alpha-numeric
            'hole location (so we don't have to worry about HoleLocAlpha = "???").
        End If

        HasCtlgReserves = gGetHasCtlgReserves(cboMineName.Text)
        If HasCtlgReserves = True Then
            ProspSets = 2
        Else
            ProspSets = 1
        End If

        'All mines have 100% prospect on the MOIS comp/split side
        'Only some mines currently have Catalog reserves as well
        '(South Fort Meade, Hookers Prairie and Wingate do not).

        For DataIdx = 1 To ProspSets
            'ProspSets will be 1 or 2
            'DataIdx = 1  Saving to 100% Prospect
            'DataIdx = 2  Saving to Catalog

            OkToContinue = False
            If DataIdx = 1 Then
                'Going to save 100% Prospect
                If aStandardsToSave = "100% Prospect Only" Or
                    aStandardsToSave = "100% Prospect and Catalog" Then
                    OkToContinue = True
                Else
                    OkToContinue = False
                End If
            End If
            If DataIdx = 2 Then
                'Going to save Catalog
                If aStandardsToSave = "Catalog Only" Or
                    aStandardsToSave = "100% Prospect and Catalog" Then
                    OkToContinue = True
                Else
                    OkToContinue = False
                End If
            End If

            If OkToContinue = True Then
                With aProspData
                    params = gDBParams

                    params.Add("pMineName", cboMineName.Text, ORAPARM_INPUT)     '1
                    params("pMineName").serverType = ORATYPE_VARCHAR2

                    params.Add("pTownShip", .Township, ORAPARM_INPUT)            '2
                    params("pTownShip").serverType = ORATYPE_NUMBER

                    params.Add("pRange", .Range, ORAPARM_INPUT)                  '3
                    params("pRange").serverType = ORATYPE_NUMBER

                    params.Add("pSection", .Section, ORAPARM_INPUT)              '4
                    params("pSection").serverType = ORATYPE_NUMBER

                    If ProspGridType = "Alpha-numeric" Then
                        HoleLoc = HoleLocAlpha
                    Else    'ProspGridType = "Numeric"
                        HoleLoc = .HoleLocation
                    End If
                    params.Add("pHoleLocation", HoleLoc, ORAPARM_INPUT)          '5
                    params("pHoleLocation").serverType = ORATYPE_VARCHAR2

                    params.Add("pSplit", .SplitNumber, ORAPARM_INPUT)            '6
                    params("pSplit").serverType = ORATYPE_NUMBER

                    params.Add("pDrillDate", .ProspDate, ORAPARM_INPUT)          '7
                    params("pDrillDate").serverType = ORATYPE_VARCHAR2

                    'Need the wash date -- don't have it available in ProspDate yet!
                    'Will put the prospect date in here for now!
                    params.Add("pWashDate", .ProspDate, ORAPARM_INPUT)               '8
                    params("pWashDate").serverType = ORATYPE_VARCHAR2

                    params.Add("pAreaOfInfluence", .Aoi, ORAPARM_INPUT)              '9
                    params("pAreaOfInfluence").serverType = ORATYPE_NUMBER

                    'Don't have a prospector code!
                    params.Add("pProspectorCode", " ", ORAPARM_INPUT)                '10
                    params("pProspectorCode").serverType = ORATYPE_VARCHAR2

                    params.Add("pTopOfSplitDepth", .SplitDepthTop, ORAPARM_INPUT)    '11
                    params("pTopOfSplitDepth").serverType = ORATYPE_NUMBER

                    params.Add("pBotOfSplitDepth", .SplitDepthBot, ORAPARM_INPUT)    '12
                    params("pBotOfSplitDepth").serverType = ORATYPE_NUMBER

                    params.Add("pSplitThickness", .SplitThck, ORAPARM_INPUT)         '13
                    params("pSplitThickness").serverType = ORATYPE_NUMBER

                    params.Add("pSampleNumber", .SampleId, ORAPARM_INPUT)            '14
                    params("pSampleNumber").serverType = ORATYPE_VARCHAR2

                    'Don't care about density volume right now.
                    'Don't really need it in MOIS for anything.
                    params.Add("pWetDensityVolume", 0, ORAPARM_INPUT)                '15
                    params("pWetDensityVolume").serverType = ORATYPE_NUMBER

                    'Don't care about wet density weight right now.
                    'Don't really need it in MOIS for anything.
                    params.Add("pWetDensityWeight", 0, ORAPARM_INPUT)                '16
                    params("pWetDensityWeight").serverType = ORATYPE_NUMBER

                    'Need to create a wet density from a dry density.
                    If .MtxPctSol <> 0 Then
                        MtxWetDensity = Round(.MtxDensity / (.MtxPctSol / 100), 1)
                    Else
                        MtxWetDensity = 0
                    End If
                    params.Add("pWetDensity", MtxWetDensity, ORAPARM_INPUT)          '17
                    params("pWetDensity").serverType = ORATYPE_NUMBER

                    'Don't care about wet matrix lbs right now.
                    'Don't really need it in MOIS for anything.
                    params.Add("pWetMtxLbs", 0, ORAPARM_INPUT)                       '18
                    params("pWetMtxLbs").serverType = ORATYPE_NUMBER

                    'Mtx %mois wet wt - Mtx %mois tare wt
                    'Don't really need it in MOIS for anything.
                    params.Add("pMtxGmsWet", 0, ORAPARM_INPUT)                       '19
                    params("pMtxGmsWet").serverType = ORATYPE_NUMBER

                    'Mtx %mois dry wt - Mtx %mois tare wt
                    'Don't really need it in MOIS for anything.
                    params.Add("pMtxGmsDry", 0, ORAPARM_INPUT)                       '20
                    params("pMtxGmsDry").serverType = ORATYPE_NUMBER

                    params.Add("pPercentSolidsMtx", .MtxPctSol, ORAPARM_INPUT)       '21
                    params("pPercentSolidsMtx").serverType = ORATYPE_NUMBER

                    params.Add("pDryDensity", .MtxDensity, ORAPARM_INPUT)            '22
                    params("pDryDensity").serverType = ORATYPE_NUMBER

                    'Wet feed lbs
                    'Don't really need it in MOIS for anything.
                    params.Add("pWetFeedLbs", 0, ORAPARM_INPUT)                      '23
                    params("pWetFeedLbs").serverType = ORATYPE_NUMBER

                    'Fd %mois wet wt - Fd %mois tare wt
                    'Don't really need it in MOIS for anything.
                    params.Add("pFeedMoistWetGms", 0, ORAPARM_INPUT)                 '24
                    params("pFeedMoistWetGms").serverType = ORATYPE_NUMBER

                    'Fd %mois dry wt - Fd %mois tare wt
                    'Don't really need it in MOIS for anything.
                    params.Add("pFeedMoistDryGms", 0, ORAPARM_INPUT)                 '25
                    params("pFeedMoistDryGms").serverType = ORATYPE_NUMBER

                    'Dry density volume? -- 0's in GEOCOMP.
                    params.Add("pDryDensityVolume", 0, ORAPARM_INPUT)                '26
                    params("pDryDensityVolume").serverType = ORATYPE_NUMBER

                    'Dry density weight? -- 0's in GEOCOMP.
                    params.Add("pDryDensityWeight", 0, ORAPARM_INPUT)                '27
                    params("pDryDensityWeight").serverType = ORATYPE_NUMBER

                    'Don't really have a matrix BPL.
                    params.Add("pMtxBPL", .MtxBPL, ORAPARM_INPUT)                    '28
                    params("pMtxBPL").serverType = ORATYPE_NUMBER

                    'Do not have matrix Insol.
                    params.Add("pMtxInsol", 0, ORAPARM_INPUT)                        '29
                    params("pMtxInsol").serverType = ORATYPE_NUMBER

                    If DataIdx = 1 Then
                        '100%  100%  100%  100%  100%  100%
                        '100%  100%  100%  100%  100%  100%
                        '100%  100%  100%  100%  100%  100%

                        'Coarse pebble grams.
                        'Don't really need it in MOIS for anything.
                        params.Add("pCpGrams", 0, ORAPARM_INPUT)                     '30
                        params("pCpGrams").serverType = ORATYPE_NUMBER

                        params.Add("pCpBPL", .Cpb100.Bpl, ORAPARM_INPUT)             '31
                        params("pCpBPL").serverType = ORATYPE_NUMBER

                        params.Add("pCpInsol", .Cpb100.Ins, ORAPARM_INPUT)           '32
                        params("pCpInsol").serverType = ORATYPE_NUMBER

                        params.Add("pCpFe2O3", .Cpb100.Fe, ORAPARM_INPUT)            '33
                        params("pCpFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pCpAl2O3", .Cpb100.Al, ORAPARM_INPUT)            '34
                        params("pCpAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pCpMgO", .Cpb100.Mg, ORAPARM_INPUT)              '35
                        params("pCpMgO").serverType = ORATYPE_NUMBER

                        params.Add("pCpCaO", .Cpb100.Ca, ORAPARM_INPUT)              '36
                        params("pCpCaO").serverType = ORATYPE_NUMBER

                        'Fine pebble grams.
                        'Don't really need it in MOIS for anything.
                        params.Add("pFpGrams", 0, ORAPARM_INPUT)                     '37
                        params("pFpGrams").serverType = ORATYPE_NUMBER

                        params.Add("pFpBPL", .Fpb100.Bpl, ORAPARM_INPUT)             '38
                        params("pFpBPL").serverType = ORATYPE_NUMBER

                        params.Add("pFpInsol", .Fpb100.Ins, ORAPARM_INPUT)           '39
                        params("pFpInsol").serverType = ORATYPE_NUMBER

                        params.Add("pFpFe2O3", .Fpb100.Fe, ORAPARM_INPUT)            '40
                        params("pFpFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pFpAl2O3", .Fpb100.Al, ORAPARM_INPUT)            '41
                        params("pFpAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pFpMgO", .Fpb100.Mg, ORAPARM_INPUT)              '42
                        params("pFpMgO").serverType = ORATYPE_NUMBER

                        params.Add("pFpCaO", .Fpb100.Ca, ORAPARM_INPUT)              '43
                        params("pFpCaO").serverType = ORATYPE_NUMBER

                        params.Add("pFpFeAl", .Fpb100.Ia, ORAPARM_INPUT)             '44
                        params("pFpFeAl").serverType = ORATYPE_NUMBER

                        CaP2O5Val = gGetCaP2O5(.Fpb100.Ca, .Fpb100.Bpl, 2)
                        params.Add("pFpCaOP2O5", CaP2O5Val, ORAPARM_INPUT)           '45
                        params("pFpCaOP2O5").serverType = ORATYPE_NUMBER

                        'Total feed wet weight floated (grams)
                        'Don't really need it in MOIS for anything.
                        params.Add("pTfGrams", 0, ORAPARM_INPUT)                     '46
                        params("pTfGrams").serverType = ORATYPE_NUMBER

                        'Total feed BPL -- measured in the lab
                        'Don't have this -- will set to calculated total feed BPL
                        params.Add("pTfBPL2", .Tfd100.Bpl, ORAPARM_INPUT)            '47
                        params("pTfBPL2").serverType = ORATYPE_NUMBER

                        'Don't really have a clay BPL for here.
                        params.Add("pWcBPL", .Wcl100.Bpl, ORAPARM_INPUT)             '48
                        params("pWcBPL").serverType = ORATYPE_NUMBER

                        'Total feed rougher tail dry weight (grams)
                        'In recent years this is actually RghTl and ClnTl combined.
                        'Don't really need it in MOIS for anything.
                        params.Add("pFatGrams", 0, ORAPARM_INPUT)                    '49
                        params("pFatGrams").serverType = ORATYPE_NUMBER

                        'Total feed rougher tail BPL (grams)
                        'In recent years this is actually RghTl and ClnTl combined.
                        'Don't really need it in MOIS for anything.
                        params.Add("pFatBPL", 0, ORAPARM_INPUT)                      '50
                        params("pFatBPL").serverType = ORATYPE_NUMBER

                        'Total feed cleaner tail dry weight (grams)
                        'Don't really need it in MOIS for anything.
                        params.Add("pAtGrams", 0, ORAPARM_INPUT)                     '51
                        params("pAtGrams").serverType = ORATYPE_NUMBER

                        'Total feed cleaner tail BPL (grams)
                        'Don't really need it in MOIS for anything.
                        params.Add("pAtBPL", 0, ORAPARM_INPUT)                       '52
                        params("pAtBPL").serverType = ORATYPE_NUMBER

                        'Lab concentrate grams
                        'Don't really need it in MOIS for anything.
                        params.Add("pLcnGrams", 0, ORAPARM_INPUT)                    '53
                        params("pLcnGrams").serverType = ORATYPE_NUMBER

                        'Lab concentrate BPL
                        'Will set to calculated concentrate value.
                        params.Add("pLcnBPL", .Tcn100.Bpl, ORAPARM_INPUT)            '54
                        params("pLcnBPL").serverType = ORATYPE_NUMBER

                        'Lab concentrate Insol
                        'Will set to calculated concentrate value.
                        params.Add("pLcnInsol", .Tcn100.Ins, ORAPARM_INPUT)          '55
                        params("pLcnInsol").serverType = ORATYPE_NUMBER

                        'Lab concentrate Fe2O3
                        'Will set to calculated concentrate value.
                        params.Add("pLcnFe2O3", .Tcn100.Fe, ORAPARM_INPUT)           '56
                        params("pLcnFe2O3").serverType = ORATYPE_NUMBER

                        'Lab concentrate Al2O3
                        'Will set to calculated concentrate value.
                        params.Add("pLcnAl2O3", .Tcn100.Al, ORAPARM_INPUT)           '57
                        params("pLcnAl2O3").serverType = ORATYPE_NUMBER

                        'Lab concentrate MgO
                        'Will set to calculated concentrate value.
                        params.Add("pLcnMgO", .Tcn100.Mg, ORAPARM_INPUT)             '58
                        params("pLcnMgO").serverType = ORATYPE_NUMBER

                        'Lab concentrate CaO
                        'Will set to calculated concentrate value.
                        params.Add("pLcnCaO", .Tcn100.Ca, ORAPARM_INPUT)             '59
                        params("pLcnCaO").serverType = ORATYPE_NUMBER

                        'Lab concentrate IA
                        'Will set to calculated concentrate value.
                        params.Add("pLcnFeAl", .Tcn100.Ia, ORAPARM_INPUT)            '60
                        params("pLcnFeAl").serverType = ORATYPE_NUMBER

                        params.Add("pCfWeightPercent", .Cfd100.WtPct, ORAPARM_INPUT) '61
                        params("pCfWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pCfBPL", .Cfd100.Bpl, ORAPARM_INPUT)             '62
                        params("pCfBPL").serverType = ORATYPE_NUMBER

                        params.Add("pFfWeightPercent", .Ffd100.WtPct, ORAPARM_INPUT) '63
                        params("pFfWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pFfBPL", .Ffd100.Bpl, ORAPARM_INPUT)             '64
                        params("pFfBPL").serverType = ORATYPE_NUMBER

                        params.Add("pTotalNumberSplits", TotNumSplits, ORAPARM_INPUT) '65
                        params("pTotalNumberSplits").serverType = ORATYPE_NUMBER

                        params.Add("pMtxTonsPerAcre", .MtxTPA, ORAPARM_INPUT)        '66
                        params("pMtxTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pWcWeightPercent", .Wcl100.WtPct, ORAPARM_INPUT) '67
                        params("pWcWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pWcTonsPerAcre", .Wcl100.Tpa, ORAPARM_INPUT)     '68
                        params("pWcTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pCpWeightPercent", .Cpb100.WtPct, ORAPARM_INPUT) '69
                        params("pCpWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pCpTonsPerAcre", .Cpb100.Tpa, ORAPARM_INPUT)     '70
                        params("pCpTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pFpWeightPercent", .Fpb100.WtPct, ORAPARM_INPUT) '71
                        params("pFpWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pFpTonsPerAcre", .Fpb100.Tpa, ORAPARM_INPUT)     '72
                        params("pFpTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pTfweightPercent", .Tfd100.WtPct, ORAPARM_INPUT) '73
                        params("pTfweightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pTfTonsPerAcre", .Tfd100.Tpa, ORAPARM_INPUT)     '74
                        params("pTfTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pTfBPL", .Tfd100.Bpl, ORAPARM_INPUT)             '75
                        params("pTfBPL").serverType = ORATYPE_NUMBER

                        params.Add("pFfTonsPerAcre", .Ffd100.Tpa, ORAPARM_INPUT)     '76
                        params("pFfTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pCfTonsPerAcre", .Cfd100.Tpa, ORAPARM_INPUT)     '77
                        params("pCfTonsPerAcre").serverType = ORATYPE_NUMBER

                        'Amine concentrate TPA
                        params.Add("pAcnTonsPerAcre", .Tcn100.Tpa, ORAPARM_INPUT)    '78
                        params("pAcnTonsPerAcre").serverType = ORATYPE_NUMBER

                        'Fine amine tails TPA
                        params.Add("pFatTonsPerAcre", 0, ORAPARM_INPUT)              '79
                        params("pFatTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pAtTonsPerAcre", 0, ORAPARM_INPUT)
                        params("pAtTonsPerAcre").serverType = ORATYPE_NUMBER        '80

                        params.Add("pCalcLossPercent", 0, ORAPARM_INPUT)
                        params("pCalcLossPercent").serverType = ORATYPE_NUMBER      '81

                        params.Add("pCalcLossTPA", 0, ORAPARM_INPUT)                 '82
                        params("pCalcLossTPA").serverType = ORATYPE_NUMBER

                        params.Add("pCalcLossBPL", 0, ORAPARM_INPUT)                 '83
                        params("pCalcLossBPL").serverType = ORATYPE_NUMBER

                        params.Add("pCnWeightPercent", .Tcn100.WtPct, ORAPARM_INPUT) '84
                        params("pCnWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pCnTonsPerAcre", .Tcn100.Tpa, ORAPARM_INPUT)     '85
                        params("pCnTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pCnBPL", .Tcn100.Bpl, ORAPARM_INPUT)             '86
                        params("pCnBPL").serverType = ORATYPE_NUMBER

                        params.Add("pCnInsol", .Tcn100.Ins, ORAPARM_INPUT)           '87
                        params("pCnInsol").serverType = ORATYPE_NUMBER

                        params.Add("pCnFeAl", .Tcn100.Ia, ORAPARM_INPUT)             '88
                        params("pCnFeAl").serverType = ORATYPE_NUMBER

                        CaP2O5Val = gGetCaP2O5(.Tcn100.Ca, .Tcn100.Bpl, 2)
                        params.Add("pCnCaOP2O5", CaP2O5Val, ORAPARM_INPUT)           '89
                        params("pCnCaOP2O5").serverType = ORATYPE_NUMBER

                        params.Add("pTpWeightPercent", .Tpr100.WtPct, ORAPARM_INPUT) '90
                        params("pTpWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pTpTonsPerAcre", .Tpr100.Tpa, ORAPARM_INPUT)     '91
                        params("pTpTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pTpBPL", .Tpr100.Bpl, ORAPARM_INPUT)             '92
                        params("pTpBPL").serverType = ORATYPE_NUMBER

                        params.Add("pTpInsol", .Tpr100.Ins, ORAPARM_INPUT)           '93
                        params("pTpInsol").serverType = ORATYPE_NUMBER

                        params.Add("pTpFe2O3", .Tpr100.Fe, ORAPARM_INPUT)            '94
                        params("pTpFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pTpAl2O3", .Tpr100.Al, ORAPARM_INPUT)            '95
                        params("pTpAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pTpMgO", .Tpr100.Mg, ORAPARM_INPUT)              '96
                        params("pTpMgO").serverType = ORATYPE_NUMBER

                        params.Add("pTpCaO", .Tpr100.Ca, ORAPARM_INPUT)              '97
                        params("pTpCaO").serverType = ORATYPE_NUMBER

                        params.Add("pTpFeAl", .Tpr100.Ia, ORAPARM_INPUT)             '98
                        params("pTpFeAl").serverType = ORATYPE_NUMBER

                        CaP2O5Val = gGetCaP2O5(.Tpr100.Ca, .Tpr100.Bpl, 2)
                        params.Add("pTpCaOP2O5", CaP2O5Val, ORAPARM_INPUT)           '99
                        params("pTpCaOP2O5").serverType = ORATYPE_NUMBER

                        params.Add("pMtxX", .MtxxAll100, ORAPARM_INPUT)              '100
                        params("pMtxX").serverType = ORATYPE_NUMBER

                        'Don't care about this right now!
                        params.Add("pRatioOfConc", 0, ORAPARM_INPUT)                 '101
                        params("pRatioOfConc").serverType = ORATYPE_NUMBER
                    Else
                        'Catalog  Catalog  Catalog  Catalog
                        'Catalog  Catalog  Catalog  Catalog
                        'Catalog  Catalog  Catalog  Catalog

                        'Coarse pebble grams.
                        'Don't really need it in MOIS for anything.
                        params.Add("pCpGrams", 0, ORAPARM_INPUT)
                        params("pCpGrams").serverType = ORATYPE_NUMBER

                        params.Add("pCpBPL", .Cpb.Bpl, ORAPARM_INPUT)
                        params("pCpBPL").serverType = ORATYPE_NUMBER

                        params.Add("pCpInsol", .Cpb.Ins, ORAPARM_INPUT)
                        params("pCpInsol").serverType = ORATYPE_NUMBER

                        params.Add("pCpFe2O3", .Cpb.Fe, ORAPARM_INPUT)
                        params("pCpFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pCpAl2O3", .Cpb.Al, ORAPARM_INPUT)
                        params("pCpAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pCpMgO", .Cpb.Mg, ORAPARM_INPUT)
                        params("pCpMgO").serverType = ORATYPE_NUMBER

                        params.Add("pCpCaO", .Cpb.Ca, ORAPARM_INPUT)
                        params("pCpCaO").serverType = ORATYPE_NUMBER

                        'Fine pebble grams.
                        'Don't really need it in MOIS for anything.
                        params.Add("pFpGrams", 0, ORAPARM_INPUT)
                        params("pFpGrams").serverType = ORATYPE_NUMBER

                        params.Add("pFpBPL", .Fpb.Bpl, ORAPARM_INPUT)
                        params("pFpBPL").serverType = ORATYPE_NUMBER

                        params.Add("pFpInsol", .Fpb.Ins, ORAPARM_INPUT)
                        params("pFpInsol").serverType = ORATYPE_NUMBER

                        params.Add("pFpFe2O3", .Fpb.Fe, ORAPARM_INPUT)
                        params("pFpFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pFpAl2O3", .Fpb.Al, ORAPARM_INPUT)
                        params("pFpAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pFpMgO", .Fpb.Mg, ORAPARM_INPUT)
                        params("pFpMgO").serverType = ORATYPE_NUMBER

                        params.Add("pFpCaO", .Fpb.Ca, ORAPARM_INPUT)
                        params("pFpCaO").serverType = ORATYPE_NUMBER

                        params.Add("pFpFeAl", .Fpb.Ia, ORAPARM_INPUT)
                        params("pFpFeAl").serverType = ORATYPE_NUMBER

                        CaP2O5Val = gGetCaP2O5(.Fpb.Ca, .Fpb.Bpl, 2)
                        params.Add("pFpCaOP2O5", CaP2O5Val, ORAPARM_INPUT)
                        params("pFpCaOP2O5").serverType = ORATYPE_NUMBER

                        'Total feed wet weight floated (grams)
                        'Don't really need it in MOIS for anything.
                        params.Add("pTfGrams", 0, ORAPARM_INPUT)
                        params("pTfGrams").serverType = ORATYPE_NUMBER

                        'Total feed BPL -- measured in the lab
                        'Don't have this -- will set to calculated total feed BPL
                        params.Add("pTfBPL2", .Tfd.Bpl, ORAPARM_INPUT)
                        params("pTfBPL2").serverType = ORATYPE_NUMBER

                        'Don't really have a clay BPL for here.
                        params.Add("pWcBPL", .Wcl.Bpl, ORAPARM_INPUT)
                        params("pWcBPL").serverType = ORATYPE_NUMBER

                        'Total feed rougher tail dry weight (grams)
                        'In recent years this is actually RghTl and ClnTl combined.
                        'Don't really need it in MOIS for anything.
                        params.Add("pFatGrams", 0, ORAPARM_INPUT)
                        params("pFatGrams").serverType = ORATYPE_NUMBER

                        'Total feed rougher tail BPL (grams)
                        'In recent years this is actually RghTl and ClnTl combined.
                        'Don't really need it in MOIS for anything.
                        params.Add("pFatBPL", 0, ORAPARM_INPUT)
                        params("pFatBPL").serverType = ORATYPE_NUMBER

                        'Total feed cleaner tail dry weight (grams)
                        'Don't really need it in MOIS for anything.
                        params.Add("pAtGrams", 0, ORAPARM_INPUT)
                        params("pAtGrams").serverType = ORATYPE_NUMBER

                        'Total feed cleaner tail BPL (grams)
                        'Don't really need it in MOIS for anything.
                        params.Add("pAtBPL", 0, ORAPARM_INPUT)
                        params("pAtBPL").serverType = ORATYPE_NUMBER

                        'Lab concentrate grams
                        'Don't really need it in MOIS for anything.
                        params.Add("pLcnGrams", 0, ORAPARM_INPUT)
                        params("pLcnGrams").serverType = ORATYPE_NUMBER

                        'Lab concentrate BPL
                        'Will set to calculated concentrate value.
                        params.Add("pLcnBPL", .Tcn.Bpl, ORAPARM_INPUT)
                        params("pLcnBPL").serverType = ORATYPE_NUMBER

                        'Lab concentrate Insol
                        'Will set to calculated concentrate value.
                        params.Add("pLcnInsol", .Tcn.Ins, ORAPARM_INPUT)
                        params("pLcnInsol").serverType = ORATYPE_NUMBER

                        'Lab concentrate Fe2O3
                        'Will set to calculated concentrate value.
                        params.Add("pLcnFe2O3", .Tcn.Fe, ORAPARM_INPUT)
                        params("pLcnFe2O3").serverType = ORATYPE_NUMBER

                        'Lab concentrate Al2O3
                        'Will set to calculated concentrate value.
                        params.Add("pLcnAl2O3", .Tcn.Al, ORAPARM_INPUT)
                        params("pLcnAl2O3").serverType = ORATYPE_NUMBER

                        'Lab concentrate MgO
                        'Will set to calculated concentrate value.
                        params.Add("pLcnMgO", .Tcn.Mg, ORAPARM_INPUT)
                        params("pLcnMgO").serverType = ORATYPE_NUMBER

                        'Lab concentrate CaO
                        'Will set to calculated concentrate value.
                        params.Add("pLcnCaO", .Tcn.Ca, ORAPARM_INPUT)
                        params("pLcnCaO").serverType = ORATYPE_NUMBER

                        'Lab concentrate IA
                        'Will set to calculated concentrate value.
                        params.Add("pLcnFeAl", .Tcn.Ia, ORAPARM_INPUT)
                        params("pLcnFeAl").serverType = ORATYPE_NUMBER

                        params.Add("pCfWeightPercent", .Cfd.WtPct, ORAPARM_INPUT)
                        params("pCfWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pCfBPL", .Cfd.Bpl, ORAPARM_INPUT)
                        params("pCfBPL").serverType = ORATYPE_NUMBER

                        params.Add("pFfWeightPercent", .Ffd.WtPct, ORAPARM_INPUT)
                        params("pFfWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pFfBPL", .Ffd.Bpl, ORAPARM_INPUT)
                        params("pFfBPL").serverType = ORATYPE_NUMBER

                        params.Add("pTotalNumberSplits", TotNumSplits, ORAPARM_INPUT)
                        params("pTotalNumberSplits").serverType = ORATYPE_NUMBER

                        params.Add("pMtxTonsPerAcre", .MtxTpaPc, ORAPARM_INPUT)
                        params("pMtxTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pWcWeightPercent", .Wcl.WtPct, ORAPARM_INPUT)
                        params("pWcWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pWcTonsPerAcre", .Wcl.Tpa, ORAPARM_INPUT)
                        params("pWcTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pCpWeightPercent", .Cpb.WtPct, ORAPARM_INPUT)
                        params("pCpWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pCpTonsPerAcre", .Cpb.Tpa, ORAPARM_INPUT)
                        params("pCpTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pFpWeightPercent", .Fpb.WtPct, ORAPARM_INPUT)
                        params("pFpWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pFpTonsPerAcre", .Fpb.Tpa, ORAPARM_INPUT)
                        params("pFpTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pTfWeightPercent", .Tfd.WtPct, ORAPARM_INPUT)
                        params("pTfWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pTfTonsPerAcre", .Tfd.Tpa, ORAPARM_INPUT)
                        params("pTfTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pTfBPL", .Tfd.Bpl, ORAPARM_INPUT)
                        params("pTfBPL").serverType = ORATYPE_NUMBER

                        params.Add("pCfTonsPerAcre", .Cfd.Tpa, ORAPARM_INPUT)
                        params("pCfTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pFfTonsPerAcre", .Ffd.Tpa, ORAPARM_INPUT)
                        params("pFfTonsPerAcre").serverType = ORATYPE_NUMBER

                        'Amine concentrate TPA
                        params.Add("pAcnTonsPerAcre", .Tcn.Tpa, ORAPARM_INPUT)
                        params("pAcnTonsPerAcre").serverType = ORATYPE_NUMBER

                        'Fine amine tails TPA
                        params.Add("pFatTonsPerAcre", 0, ORAPARM_INPUT)
                        params("pFatTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pAtTonsPerAcre", 0, ORAPARM_INPUT)
                        params("pAtTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pCalcLossPercent", 0, ORAPARM_INPUT)
                        params("pCalcLossPercent").serverType = ORATYPE_NUMBER

                        params.Add("pCalcLossTPA", 0, ORAPARM_INPUT)
                        params("pCalcLossTPA").serverType = ORATYPE_NUMBER

                        params.Add("pCalcLossBPL", 0, ORAPARM_INPUT)
                        params("pCalcLossBPL").serverType = ORATYPE_NUMBER

                        params.Add("pCnWeightPercent", .Tcn.WtPct, ORAPARM_INPUT)
                        params("pCnWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pCnTonsPerAcre", .Tcn.Tpa, ORAPARM_INPUT)
                        params("pCnTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pCnBPL", .Tcn.Bpl, ORAPARM_INPUT)
                        params("pCnBPL").serverType = ORATYPE_NUMBER

                        params.Add("pCnInsol", .Tcn.Ins, ORAPARM_INPUT)
                        params("pCnInsol").serverType = ORATYPE_NUMBER

                        params.Add("pCnFeAl", .Tcn.Ia, ORAPARM_INPUT)
                        params("pCnFeAl").serverType = ORATYPE_NUMBER

                        CaP2O5Val = gGetCaP2O5(.Tcn.Ca, .Tcn.Bpl, 2)
                        params.Add("pCnCaOP2O5", CaP2O5Val, ORAPARM_INPUT)
                        params("pCnCaOP2O5").serverType = ORATYPE_NUMBER

                        params.Add("pTpWeightPercent", .Tpr.WtPct, ORAPARM_INPUT)
                        params("pTpWeightPercent").serverType = ORATYPE_NUMBER

                        params.Add("pTpTonsPerAcre", .Tpr.Tpa, ORAPARM_INPUT)
                        params("pTpTonsPerAcre").serverType = ORATYPE_NUMBER

                        params.Add("pTpBPL", .Tpr.Bpl, ORAPARM_INPUT)
                        params("pTpBPL").serverType = ORATYPE_NUMBER

                        params.Add("pTpInsol", .Tpr.Ins, ORAPARM_INPUT)
                        params("pTpInsol").serverType = ORATYPE_NUMBER

                        params.Add("pTpFe2O3", .Tpr.Fe, ORAPARM_INPUT)
                        params("pTpFe2O3").serverType = ORATYPE_NUMBER

                        params.Add("pTpAl2O3", .Tpr.Al, ORAPARM_INPUT)
                        params("pTpAl2O3").serverType = ORATYPE_NUMBER

                        params.Add("pTpMgO", .Tpr.Mg, ORAPARM_INPUT)
                        params("pTpMgO").serverType = ORATYPE_NUMBER

                        params.Add("pTpCaO", .Tpr.Ca, ORAPARM_INPUT)
                        params("pTpCaO").serverType = ORATYPE_NUMBER

                        params.Add("pTpFeAl", .Tpr.Ia, ORAPARM_INPUT)
                        params("pTpFeAl").serverType = ORATYPE_NUMBER

                        CaP2O5Val = gGetCaP2O5(.Tpr.Ca, .Tpr.Bpl, 2)
                        params.Add("pTpCaOP2O5", CaP2O5Val, ORAPARM_INPUT)
                        params("pTpCaOP2O5").serverType = ORATYPE_NUMBER

                        params.Add("pMtxX", .MtxxAll, ORAPARM_INPUT)
                        params("pMtxX").serverType = ORATYPE_NUMBER

                        'Don't care about this right now!
                        params.Add("pRatioOfConc", 0, ORAPARM_INPUT)
                        params("pRatioOfConc").serverType = ORATYPE_NUMBER
                    End If

                    '.MineableCalc will be set to "M" or "U".
                    'Minability status
                    'GEOCOMP minability codes are:
                    'A = Active split
                    'M = Mined out
                    'I = Unminable split
                    'B = Bottom split
                    'O = Unminable hole
                    If .MineableCalcd = "M" Then
                        'Minable split
                        MinableStatus = "A"
                    Else
                        'Unminable split
                        'Minability status
                        'GEOCOMP minability codes are:
                        'A = Active split
                        'M = Mined out
                        'I = Unminable split
                        'B = Bottom split
                        'O = Unminable hole
                        MinableStatus = "I"
                    End If

                    params.Add("pMineableStatus", MinableStatus, ORAPARM_INPUT)      '102
                    params("pMineableStatus").serverType = ORATYPE_VARCHAR2

                    'Don't have a county code for here.
                    params.Add("pCountyCode", " ", ORAPARM_INPUT)                    '103
                    params("pCountyCode").serverType = ORATYPE_VARCHAR2

                    'Don't have a mining code for here.
                    params.Add("pMiningCode", " ", ORAPARM_INPUT)                    '104
                    params("pMiningCode").serverType = ORATYPE_VARCHAR2

                    'Don't have a pumping code for here.
                    params.Add("pPumpingCode", " ", ORAPARM_INPUT)                   '105
                    params("pPumpingCode").serverType = ORATYPE_VARCHAR2

                    'Don't have a MetLab ID for here -- some identification
                    'number for the metlab that processed the split.
                    params.Add("pMetLabID", " ", ORAPARM_INPUT)                      '106
                    params("pMetLabID").serverType = ORATYPE_VARCHAR2

                    'Don't have a ChemLab ID for here -- some identification
                    'number for the chemlab that processed the split.
                    params.Add("pChemLabID", " ", ORAPARM_INPUT)                     '107
                    params("pChemLabID").serverType = ORATYPE_VARCHAR2

                    'Will have a color code here -- get the color and will
                    'save that in MOIS.
                    'Will set phosphate color to "" for now!
                    'Need to translate to GEOCOMP color codes!
                    params.Add("pColor", " ", ORAPARM_INPUT)                         '108
                    params("pColor").serverType = ORATYPE_VARCHAR2

                    'Split elevation
                    SplitElev = .Elevation - .SplitDepthTop
                    If SplitElev < 0 Then
                        SplitElev = 0
                    End If
                    params.Add("pSplitElevation", SplitElev, ORAPARM_INPUT)          '109
                    params("pSplitElevation").serverType = ORATYPE_NUMBER

                    params.Add("pHoleNumber", Val(.HoleLocation), ORAPARM_INPUT)     '110
                    params("pHoleNumber").serverType = ORATYPE_NUMBER

                    'Don't have a triangle code for here.
                    params.Add("pTriangleCode", "", ORAPARM_INPUT)                   '111
                    params("pTriangleCode").serverType = ORATYPE_VARCHAR2

                    params.Add("pSampleNumber2", .SampleId, ORAPARM_INPUT)            '112
                    params("pSampleNumber2").serverType = ORATYPE_VARCHAR2

                    '----------

                    'Don't have any Cpb Cd values.
                    params.Add("pCpCd", 0, ORAPARM_INPUT)                            '113
                    params("pCpCd").serverType = ORATYPE_NUMBER

                    'Don't have any Fpb Cd values.
                    params.Add("pFpCd", 0, ORAPARM_INPUT)                            '114
                    params("pFpCd").serverType = ORATYPE_NUMBER

                    'Don't have any Lab concentrate Cd values.
                    params.Add("pLcnCd", 0, ORAPARM_INPUT)                           '115
                    params("pLcnCd").serverType = ORATYPE_NUMBER

                    'Don't have any Total product Cd values.
                    params.Add("pTpCd", 0, ORAPARM_INPUT)                            '116
                    params("pTpCd").serverType = ORATYPE_NUMBER

                    'Don't have a Hardpan Code for here.
                    params.Add("pHardpanCode", 0, ORAPARM_INPUT)                     '117
                    params("pHardpanCode").serverType = ORATYPE_NUMBER

                    'Prospect standard
                    If DataIdx = 1 Then
                        ThisProspStandard = "100%PROSPECT"
                    End If
                    If DataIdx = 2 Then
                        ThisProspStandard = "CATALOG"
                    End If
                    params.Add("pProspStandard", ThisProspStandard, ORAPARM_INPUT)   '118
                    params("pProspStandard").serverType = ORATYPE_VARCHAR2

                    'Procedure load_split
                    'pMineName           IN     VARCHAR2,    -- 1
                    'pTownShip           IN     NUMBER,      -- 2
                    'pRange              IN     NUMBER,      -- 3
                    'pSection            IN     NUMBER,      -- 4
                    'pHoleLocation       IN     VARCHAR2,    -- 5
                    'pSplit              IN     NUMBER,      -- 6
                    'pDrillDate          IN     VARCHAR2,    -- 7
                    'pWashDate           IN     VARCHAR2,    -- 8
                    'pAreaOfInfluence    IN     NUMBER,      -- 9
                    'pProspectorCode     IN     VARCHAR2,    -- 10
                    'pTopOfSplitDepth    IN     NUMBER,      -- 11
                    'pBotOfSplitDepth    IN     NUMBER,      -- 12
                    'pSplitThickness     IN     NUMBER,      -- 13
                    'pSampleNumber       IN     VARCHAR2,    -- 14
                    'pWetDensityVolume   IN     NUMBER,      -- 15
                    'pWetDensityWeight   IN     NUMBER,      -- 16
                    'pWetDensity         IN     NUMBER,      -- 17
                    'pWetMtxLbs          IN     NUMBER,      -- 18
                    'pMtxGmsWet          IN     NUMBER,      -- 19
                    'pMtxGmsDry          IN     NUMBER,      -- 20
                    'pPercentSolidsMtx   IN     NUMBER,      -- 21
                    'pDryDensity         IN     NUMBER,      -- 22
                    'pWetFeedLbs         IN     NUMBER,      -- 23
                    'pFeedMoistWetGms    IN     NUMBER,      -- 24
                    'pFeedMoistDryGms    IN     NUMBER,      -- 25
                    'pDryDensityVolume   IN     NUMBER,      -- 26
                    'pDryDensityWeight   IN     NUMBER,      -- 27
                    'pMtxBPL             IN     NUMBER,      -- 28
                    'pMtxInsol           IN     NUMBER,      -- 29
                    'pCpGrams            IN     NUMBER,      -- 30
                    'pCpBPL              IN     NUMBER,      -- 31
                    'pCpInsol            IN     NUMBER,      -- 32
                    'pCpFe2O3            IN     NUMBER,      -- 33
                    'pCpAl2O3            IN     NUMBER,      -- 34
                    'pCpMgO              IN     NUMBER,      -- 35
                    'pCpCaO              IN     NUMBER,      -- 36
                    'pFpGrams            IN     NUMBER,      -- 37
                    'pFpBPL              IN     NUMBER,      -- 38
                    'pFpInsol            IN     NUMBER,      -- 39
                    'pFpFe2O3            IN     NUMBER,      -- 40
                    'pFpAl2O3            IN     NUMBER,      -- 41
                    'pFpMgO              IN     NUMBER,      -- 42
                    'pFpCaO              IN     NUMBER,      -- 43
                    'pFpFeAl             IN     NUMBER,      -- 44
                    'pFpCaOP2O5          IN     NUMBER,      -- 45
                    'pTfGrams            IN     NUMBER,      -- 46
                    'pTfBPL2             IN     NUMBER,      -- 47
                    'pWcBPL              IN     NUMBER,      -- 48
                    'pFatGrams           IN     NUMBER,      -- 49
                    'pFatBPL             IN     NUMBER,      -- 50
                    'pAtGrams            IN     NUMBER,      -- 51
                    'pAtBPL              IN     NUMBER,      -- 52
                    'pLcnGrams           IN     NUMBER,      -- 53
                    'pLcnBPL             IN     NUMBER,      -- 54
                    'pLcnInsol           IN     NUMBER,      -- 55
                    'pLcnFe2O3           IN     NUMBER,      -- 56
                    'pLcnAl2O3           IN     NUMBER,      -- 57
                    'pLcnMgO             IN     NUMBER,      -- 58
                    'pLcnCaO             IN     NUMBER,      -- 59
                    'pLcnFeAl            IN     NUMBER,      -- 60
                    'pCfWeightPercent    IN     NUMBER,      -- 61
                    'pCfBPL              IN     NUMBER,      -- 62
                    'pFfWeightPercent    IN     NUMBER,      -- 63
                    'pFfBPL              IN     NUMBER,      -- 64
                    'pTotalNumberSplits  IN     NUMBER,      -- 65
                    'pMtxTonsPerAcre     IN     NUMBER,      -- 66
                    'pWcWeightPercent    IN     NUMBER,      -- 67
                    'pWcTonsPerAcre      IN     NUMBER,      -- 68
                    'pCpWeightPercent    IN     NUMBER,      -- 69
                    'pCpTonsPerAcre      IN     NUMBER,      -- 70
                    'pFpWeightPercent    IN     NUMBER,      -- 71
                    'pFpTonsPerAcre      IN     NUMBER,      -- 72
                    'pTfweightPercent    IN     NUMBER,      -- 73
                    'pTfTonsPerAcre      IN     NUMBER,      -- 74
                    'pTfBPL              IN     NUMBER,      -- 75
                    'pFfTonsPerAcre      IN     NUMBER,      -- 76
                    'pCfTonsPerAcre      IN     NUMBER,      -- 77
                    'pAcnTonsPerAcre     IN     NUMBER,      -- 78
                    'pFatTonsPerAcre     IN     NUMBER,      -- 79
                    'pAtTonsPerAcre      IN     NUMBER,      -- 80
                    'pCalcLossPercent    IN     NUMBER,      -- 81
                    'pCalcLossTPA        IN     NUMBER,      -- 82
                    'pCalcLossBPL        IN     NUMBER,      -- 83
                    'pCnWeightPercent    IN     NUMBER,      -- 84
                    'pCnTonsPerAcre      IN     NUMBER,      -- 85
                    'pCnBPL              IN     NUMBER,      -- 86
                    'pCnInsol            IN     NUMBER,      -- 87
                    'pCnFeAl             IN     NUMBER,      -- 88
                    'pCnCaOP2O5          IN     NUMBER,      -- 89
                    'pTpWeightPercent    IN     NUMBER,      -- 90
                    'pTpTonsPerAcre      IN     NUMBER,      -- 91
                    'pTpBPL              IN     NUMBER,      -- 92
                    'pTpInsol            IN     NUMBER,      -- 93
                    'pTpFe2O3            IN     NUMBER,      -- 94
                    'pTpAl2O3            IN     NUMBER,      -- 95
                    'pTpMgO              IN     NUMBER,      -- 96
                    'pTpCaO              IN     NUMBER,      -- 97
                    'pTpFeAl             IN     NUMBER,      -- 98
                    'pTpCaOP2O5          IN     NUMBER,      -- 99
                    'pMtxX               IN     NUMBER,      -- 100
                    'pRatioOfConc        IN     NUMBER,      -- 101
                    'pMineableStatus     IN     VARCHAR2,    -- 102
                    'pCountyCode         IN     VARCHAR2,    -- 103
                    'pMiningCode         IN     VARCHAR2,    -- 104
                    'pPumpingCode        IN     VARCHAR2,    -- 105
                    'pMetLabID           IN     VARCHAR2,    -- 106
                    'pChemLabID          IN     VARCHAR2,    -- 107
                    'pColor              IN     VARCHAR2,    -- 108
                    'pSplitElevation     IN     NUMBER,      -- 109
                    'pHoleNumber         IN     NUMBER,      -- 110
                    'pTriangleCode       IN     VARCHAR2,    -- 111
                    'pSampleNumber2      IN     VARCHAR2,    -- 112
                    '--
                    'pCpCd               IN     NUMBER,      -- 113
                    'pFpCd               IN     NUMBER,      -- 114
                    'pLcnCd              IN     NUMBER,      -- 115
                    'pTpCd               IN     NUMBER,      -- 116
                    'pHardpanCode        IN     NUMBER,      -- 117
                    'pProspStandard      IN     VARCHAR2)    -- 118

                    SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect.load_split(:pMineName, " +
                         ":pTownShip, :pRange, :pSection, :pHoleLocation, " +
                         ":pSplit, :pDrillDate, :pWashDate, :pAreaOfInfluence, :pProspectorCode, " +
                         ":pTopOfSplitDepth, :pBotOfSplitDepth, :pSplitThickness, :pSampleNumber, :pWetDensityVolume, " +
                         ":pWetDensityWeight, :pWetDensity, :pWetMtxLbs, :pMtxGmsWet, :pMtxGmsDry, " +
                         ":pPercentSolidsMtx, :pDryDensity, :pWetFeedLbs, :pFeedMoistWetGms, :pFeedMoistDryGms, " +
                         ":pDryDensityVolume, :pDryDensityWeight, :pMtxBPL, :pMtxInsol, :pCpGrams, " +
                         ":pCpBPL, :pCpInsol, :pCpFe2O3, :pCpAl2O3, :pCpMgO, " +
                         ":pCpCaO, :pFpGrams, :pFpBPL, :pFpInsol, :pFpFe2O3, " +
                         ":pFpAl2O3, :pFpMgO, :pFpCaO, :pFpFeAl, :pFpCaOP2O5, " +
                         ":pTfGrams, :pTfBPL2, :pWcBPL, :pFatGrams, :pFatBPL, " +
                         ":pAtGrams, :pAtBPL, :pLcnGrams, :pLcnBPL, :pLcnInsol, " +
                         ":pLcnFe2O3, :pLcnAl2O3, :pLcnMgO, :pLcnCaO, :pLcnFeAl, " +
                         ":pCfWeightPercent, :pCfBPL, :pFfWeightPercent, :pFfBPL, :pTotalNumberSplits, " +
                         ":pMtxTonsPerAcre, :pWcWeightPercent, :pWcTonsPerAcre, :pCpWeightPercent, :pCpTonsPerAcre, " +
                         ":pFpWeightPercent, :pFpTonsPerAcre, :pTfweightPercent, :pTfTonsPerAcre, :pTfBPL, " +
                         ":pFfTonsPerAcre, :pCfTonsPerAcre, :pAcnTonsPerAcre, :pFatTonsPerAcre, :pAtTonsPerAcre, " +
                         ":pCalcLossPercent, :pCalcLossTPA, :pCalcLossBPL, :pCnWeightPercent, :pCnTonsPerAcre, " +
                         ":pCnBPL, :pCnInsol, :pCnFeAl, :pCnCaOP2O5, :pTpWeightPercent, " +
                         ":pTpTonsPerAcre, :pTpBPL, :pTpInsol, :pTpFe2O3, :pTpAl2O3, " +
                         ":pTpMgO, :pTpCaO, :pTpFeAl, :pTpCaOP2O5, :pMtxX, " +
                         ":pRatioOfConc, :pMineableStatus, :pCountyCode, :pMiningCode, :pPumpingCode, " +
                         ":pMetLabID, :pChemLabID, :pColor, :pSplitElevation, :pHoleNumber, " +
                         ":pTriangleCode, :pSampleNumber2, :pCpCd, :pFpCd, :pLcnCd, :pTpCd, :pHardpanCode, :pProspStandard);end;", ORASQL_FAILEXEC)

                    ClearParams(params)
                End With
            End If
        Next DataIdx

        AddSplitToMois = True

        Exit Function

AddSplitToMoisError:
        MsgBox("Error saving split." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Save Error")

        On Error Resume Next
        ClearParams(params)
    End Function

    Private Sub cmdPrintHole_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrintHole.Click
        Me.Cursor = Cursors.WaitCursor
        PrintReport("Hole")
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub cmdPrintSplit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrintSplit.Click
        Me.Cursor = Cursors.WaitCursor
        PrintReport("Split")
        Me.Cursor = Cursors.Default
    End Sub

    'Private Sub PrintReport2(ByVal aRptMode As String)

    '    'Declaring Parameters
    '    Dim RptTitle As String
    '    Dim PctProspect100 As Boolean
    '    Dim RcvryData As gDataRdctnParamsType
    '    Dim RcvryProdQual(0 To 14) As gDataRdctnProdQualType


    '    'Declaring Report
    '    Dim OutputDir As String = System.Environment.GetEnvironmentVariable("Temp") & "\OIS\"
    '    Dim OutputFile As String
    '    Dim OutputFileExtension As String
    '    Dim ReportFormat As Integer
    '    Dim Params() As ReportWeb.ReportParameter = Nothing
    '    Dim ReportInfo As New ReportWeb.Report
    '    Dim buf As Byte() = Nothing

    '    'Parameters

    '    RptTitle = aRptMode & " Data"

    '    PctProspect100 = False
    '    If opt100Pct.Checked = True Then
    '        PctProspect100 = True
    '    Else
    '        PctProspect100 = False
    '    End If

    '    SetRcvryEtc(cboOtherDefn.Text, "User recovery scenario", RcvryData, RcvryProdQual)

    '    Params = Nothing
    '    Params = Array.CreateInstance(GetType(ReportWeb.ReportParameter), 8)

    '    Params(0) = New ReportWeb.ReportParameter
    '    With Params(0)
    '        .Name = "pCompanyName"
    '        .CurrentValues(0) = New ReportWeb.ReportParameterValue
    '        .CurrentValues(0).CurrentValue = gCompanyName
    '    End With

    '    Params(1) = New ReportWeb.ReportParameter
    '    With Params(1)
    '        .Name = "pRptType"
    '        .CurrentValues(0) = New ReportWeb.ReportParameterValue
    '        .CurrentValues(0).CurrentValue = RptTitle
    '    End With

    '    Params(2) = New ReportWeb.ReportParameter
    '    With Params(2)
    '        .Name = "pPctProspect100"
    '        .CurrentValues(0) = New ReportWeb.ReportParameterValue
    '        .CurrentValues(0).CurrentValue = PctProspect100
    '    End With

    '    Params(3) = New ReportWeb.ReportParameter
    '    With Params(3)
    '        .Name = "pMineHasOffSpecPbPlt"
    '        .CurrentValues(0) = New ReportWeb.ReportParameterValue
    '        .CurrentValues(0).CurrentValue = RcvryData.UseOrigMgoPlant
    '    End With

    '    Params(4) = New ReportWeb.ReportParameter
    '    With Params(4)
    '        .Name = "pProdSizeDesig"
    '        .CurrentValues(0) = New ReportWeb.ReportParameterValue
    '        .CurrentValues(0).CurrentValue = cboProdSizeDefn.Text
    '    End With

    '    Params(5) = New ReportWeb.ReportParameter
    '    With Params(5)
    '        .Name = "pRcvryEtcScen"
    '        .CurrentValues(0) = New ReportWeb.ReportParameterValue
    '        .CurrentValues(0).CurrentValue = cboOtherDefn.Text
    '    End With

    '    Params(6) = New ReportWeb.ReportParameter
    '    With Params(6)
    '        .Name = "pMineHasDoloflotPlt"
    '        .CurrentValues(0) = New ReportWeb.ReportParameterValue
    '        .CurrentValues(0).CurrentValue = RcvryData.UseDoloflotPlant2010
    '    End With

    '    Params(7) = New ReportWeb.ReportParameter
    '    With Params(7)
    '        .Name = "pMineHasDoloflotPltFco"
    '        .CurrentValues(0) = New ReportWeb.ReportParameterValue
    '        .CurrentValues(0).CurrentValue = RcvryData.UseDoloflotPlantFco
    '    End With


    '    'Build Report
    '    ReportInfo.ReportFile = My.Settings("ProspectReductionReportFile")
    '    ReportInfo.ReportDescription = ""
    '    ReportInfo.ReportID = 0

    '    ReportFormat = ReportWeb.FileFormats.PDF
    '    OutputFileExtension = "PDF"

    '    My.WebServices.ReportService.Timeout = 300000
    '    My.WebServices.ReportService.UseDefaultCredentials = True
    '    buf = My.WebServices.ReportService.GenerateReport(ReportInfo, Params, ReportFormat)

    '    'Write the Results out to a temporary file
    '    If Not System.IO.Directory.Exists(OutputDir) Then
    '        System.IO.Directory.CreateDirectory(OutputDir)
    '    End If

    'End Sub



    Private Sub PrintReport(ByVal aRptMode As String)


        ' MessageBox.Show("Functionality is in pending developement stage")


        '    '**********************************************************************
        '    '
        '    '
        '    '
        '    '**********************************************************************

        '    On Error GoTo PrintReportError

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
        Dim MineName As String
        Dim RptTitle As String
        Dim MineHasOffSpecPbPlt As String
        Dim OffSpecPbPlt As Boolean
        Dim MaxDepthComm As String
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
        Dim RowNum As Integer
        Dim rptProspRdctn As ReportDocument

        Try


            If cboMineName.Text <> "(Select mine...)" Then
                MineName = cboMineName.Text
            Else
                MineName = "Not assigned"
            End If

            RptTitle = aRptMode & " Data"

            PctProspect100 = False
            If opt100Pct.Checked Then
                PctProspect100 = True
            Else
                PctProspect100 = False
            End If

            'GetRcvryEtcParamsFromForm RcvryData, RcvryProdQual()
            SetRcvryEtc(cboOtherDefn.Text,
                        "User recovery scenario",
                        RcvryData,
                        RcvryProdQual)

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

            '06/15/2010, lss  RcvryData.MineHasOffSpecPbPlt not used anymore
            'If RcvryData.MineHasOffSpecPbPlt = True Then
            If RcvryData.UseOrigMgoPlant Then
                MineHasOffSpecPbPlt = "Mine has off-spec pebble processing plant (Original)."
            Else
                If RcvryData.UseDoloflotPlant2010 Then
                    MineHasOffSpecPbPlt = "Mine has Ona Doloflot plant."
                Else
                    If RcvryData.UseDoloflotPlantFco Then
                        MineHasOffSpecPbPlt = "Mine has FCO Doloflot plant."
                    Else
                        MineHasOffSpecPbPlt = ""
                    End If
                End If
            End If

            If RcvryData.CanSelectRejectTpb Then
                CanSelectRejectTpb = "Yes"
            End If
            If RcvryData.CanSelectRejectTcn Then
                CanSelectRejectTcn = "Yes"
            End If

            If chkUseFeAdjust.Checked Then
                RcvryData.UseFeAdjust = True
            Else
                RcvryData.UseFeAdjust = False
            End If

            If RcvryData.UseFeAdjust Then
                UseFeAdjust = "Fe2O3 adjust has been used to determine minabilities."
            Else
                UseFeAdjust = ""
            End If

            If aRptMode = "Hole" Then
                ProspData = gGetDataFromReviewSprd(ssCompReview, 1)
                RowNum = 1
            End If

            If aRptMode = "Split" Then
                ProspData = gGetDataFromReviewSprd(ssSplitReview, Val(lblCurrSplit.Text))
                RowNum = ProspData.SplitNumber
            End If

            'Need to get whether the material is on-spec as opposed to whether the
            'material will or will not be mined!
            If aRptMode = "Split" Then
                OsOnSpec = gGetMaterialOnSpec("OS", RowNum, ssSplitReview)
                CpbOnSpec = gGetMaterialOnSpec("Cpb", RowNum, ssSplitReview)
                FpbOnSpec = gGetMaterialOnSpec("Fpb", RowNum, ssSplitReview)
                TpbOnSpec = gGetMaterialOnSpec("Tpb", RowNum, ssSplitReview)
                CcnOnSpec = gGetMaterialOnSpec("Ccn", RowNum, ssSplitReview)
                FcnOnSpec = gGetMaterialOnSpec("Fcn", RowNum, ssSplitReview)
                TcnOnSpec = gGetMaterialOnSpec("Tcn", RowNum, ssSplitReview)
                IpOnSpec = gGetMaterialOnSpec("IP", RowNum, ssSplitReview)
            End If
            If aRptMode = "Hole" Then
                OsOnSpec = gGetMaterialOnSpec("OS", RowNum, ssCompReview)
                CpbOnSpec = gGetMaterialOnSpec("Cpb", RowNum, ssCompReview)
                FpbOnSpec = gGetMaterialOnSpec("Fpb", RowNum, ssCompReview)
                TpbOnSpec = gGetMaterialOnSpec("Tpb", RowNum, ssCompReview)
                CcnOnSpec = gGetMaterialOnSpec("Ccn", RowNum, ssCompReview)
                FcnOnSpec = gGetMaterialOnSpec("Fcn", RowNum, ssCompReview)
                TcnOnSpec = gGetMaterialOnSpec("Tcn", RowNum, ssCompReview)
                IpOnSpec = gGetMaterialOnSpec("IP", RowNum, ssCompReview)
            End If

            '06/15/2009, lss
            'New report format.
            'Dim strReportPath As String = Application.StartupPath & "\ProspectReduction.rpt"
            Dim strReportPath As String = My.Settings.ReportPath
            If Not IO.File.Exists(strReportPath) Then
                Throw (New Exception("Unable to locate report file:" & vbCrLf & strReportPath))
            End If
            rptProspRdctn = New ReportDocument
            rptProspRdctn.Load(strReportPath)
            With ProspData
                rptProspRdctn.DataDefinition.FormulaFields("ProspDate").Text = "'" & .ProspDate & "'"
                'rptProspRdctn.Formulas(2) = "Section = '" & Format(.Section, "##") & "'"
                rptProspRdctn.DataDefinition.FormulaFields("Section").Text = "'" & Format(.Section, "##") & "'"
                'rptProspRdctn.Formulas(3) = "Township = '" & Format(.Township, "##") & "'"
                rptProspRdctn.DataDefinition.FormulaFields("Township").Text = "'" & Format(.Township, "##") & "'"
                'rptProspRdctn.Formulas(4) = "Range = '" & Format(.Range, "##") & "'"
                rptProspRdctn.DataDefinition.FormulaFields("Range").Text = "'" & Format(.Range, "##") & "'"
                'rptProspRdctn.Formulas(5) = "HoleLocation = '" & .HoleLocation & "'"
                rptProspRdctn.DataDefinition.FormulaFields("HoleLocation").Text = "'" & .HoleLocation & "'"

                If aRptMode = "Split" Then
                    '    rptProspRdctn.Formulas(6) = "SplitNumber = '" & Format(.SplitNumber, "##") & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("SplitNumber").Text = "'" & Format(.SplitNumber, "##") & "'"
                    '    rptProspRdctn.Formulas(7) = "SplitDepthTop = '" & Format(.SplitDepthTop, "##0.0") & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("SplitDepthTop").Text = "'" & Format(.SplitDepthTop, "##0.0") & "'"
                    '    rptProspRdctn.Formulas(8) = "SplitDepthBot = '" & Format(.SplitDepthBot, "##0.0") & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("SplitDepthBot").Text = "'" & Format(.SplitDepthBot, "##0.0") & "'"
                    '    rptProspRdctn.Formulas(9) = "SplitThck = '" & Format(.SplitThck, "##0.0") & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("SplitThck").Text = "'" & Format(.SplitThck, "##0.0") & "'"
                Else    'Hole data
                    '    rptProspRdctn.Formulas(6) = "SplitNumber = '" & "--" & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("SplitNumber").Text = "'" & "--" & "'"
                    '    rptProspRdctn.Formulas(7) = "SplitDepthTop = '" & "--" & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("SplitDepthTop").Text = "'" & "--" & "'"
                    '    rptProspRdctn.Formulas(8) = "SplitDepthBot = '" & "--" & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("SplitDepthBot").Text = "'" & "--" & "'"
                    '    rptProspRdctn.Formulas(9) = "SplitThck = '" & "--" & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("SplitThck").Text = "'" & "--" & "'"
                End If

                'rptProspRdctn.Formulas(10) = "Elevation = '" & Format(.Elevation, "#,##0.00") & "'"
                rptProspRdctn.DataDefinition.FormulaFields("Elevation").Text = "'" & Format(.Elevation, "#,##0.00") & "'"
                'rptProspRdctn.Formulas(11) = "Xcoord = '" & Format(.Xcoord, "#,###,##0.00") & "'"
                rptProspRdctn.DataDefinition.FormulaFields("Xcoord").Text = "'" & Format(.Xcoord, "#,###,##0.00") & "'"
                'rptProspRdctn.Formulas(12) = "Ycoord = '" & Format(.Ycoord, "#,###,##0.00") & "'"
                rptProspRdctn.DataDefinition.FormulaFields("Ycoord").Text = "'" & Format(.Ycoord, "#,###,##0.00") & "'"


                'OvbThk, ItbThk, and MtxThk will be the same for both 100% and ProdCoeff!
                If aRptMode = "Hole" Then
                    'rptProspRdctn.Formulas(13) = "OvbThk = '" & Format(.OvbThk, "##0.0") & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("OvbThk").Text = "'" & Format(.OvbThk, "##0.0") & "'"
                    'rptProspRdctn.Formulas(14) = "MtxThk = '" & Format(.MtxThk, "##0.0") & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("MtxThk").Text = "'" & Format(.MtxThk, "##0.0") & "'"
                    'rptProspRdctn.Formulas(15) = "ItbThk = '" & Format(.ItbThk, "##0.0") & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("ItbThk").Text = "'" & Format(.ItbThk, "##0.0") & "'"
                Else    'Split data
                    'These items do not apply to splits (only to holes).
                    'rptProspRdctn.Formulas(13) = "OvbThk = '" & "--" & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("OvbThk").Text = "'" & "--" & "'"
                    'rptProspRdctn.Formulas(14) = "MtxThk = '" & "--" & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("MtxThk").Text = "'" & "--" & "'"
                    'rptProspRdctn.Formulas(15) = "ItbThk = '" & "--" & "'"
                    rptProspRdctn.DataDefinition.FormulaFields("ItbThk").Text = "'" & "--" & "'"
                End If

                ''Laboratory matrix density remains the same for both 100% and ProdCoeff!
                ''This is the laboratory based density!
                'rptProspRdctn.Formulas(16) = "MtxDensityLab = '" & Format(.MtxDensity, "##0.0") & "'"
                rptProspRdctn.DataDefinition.FormulaFields("MtxDensityLab").Text = "'" & Format(.MtxDensity, "##0.0") & "'"
            End With
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

            gAddProdCoeffOr100Pct(ProspData,
                                  RcvryData,
                                  OsOnSpec,
                                  CpbOnSpec,
                                  FpbOnSpec,
                                  TpbOnSpec,
                                  CcnOnSpec,
                                  FcnOnSpec,
                                  TcnOnSpec,
                                  IpOnSpec,
                                  InclCpbAlways,
                                  InclFpbAlways,
                                  InclOsAlways,
                                  InclCpbNever,
                                  InclFpbNever,
                                  InclOsNever,
                                  MineHasOffSpecPbPlt,
                                  CanSelectRejectTpb,
                                  CanSelectRejectTcn,
                                  UseFeAdjust,
                                  If(optCatalog.Checked, "ProdCoeff", "100%Prospect"),
                                  rptProspRdctn,
                                  IIf(aRptMode = "Split", "Split", "Hole"),
                                  MineName)

            '    If MineHasOffSpecPbPlt <> "" Then
            '        OffSpecPbPlt = True
            '    Else
            '        OffSpecPbPlt = False
            '    End If
            '
            'Have all the needed data -- start the report
            'rptProspRdctn.ReportFileName = gPath + "\Reports\" + "ProspectReduction.rpt"

            ''Connect to Oracle database
            'ConnectString = "DSN = " + gDataSource + ";UID = " + gOracleUserName + _
            '    ";PWD = " + gOracleUserPassword + ";DSQ = "

            'rptProspRdctn.Connect = ConnectString

            'Need to pass the company name and report type into the report
            With rptProspRdctn
                .SetParameterValue("pCompanyName", "Mosaic")
                .SetParameterValue("pRptType", RptTitle)
                .SetParameterValue("pPctProspect100", PctProspect100)
                .SetParameterValue("pMineHasOffSpecPbPlt", RcvryData.UseOrigMgoPlant)
                .SetParameterValue("pProdSizeDesig", cboProdSizeDefn.Text)
                .SetParameterValue("pRcvryEtcScen", cboOtherDefn.Text)
                .SetParameterValue("pMineHasDoloflotPlt", RcvryData.UseDoloflotPlant2010)
                .SetParameterValue("pMineHasDoloflotPltFco", RcvryData.UseDoloflotPlantFco)
            End With

            'rptProspRdctn.ParameterFields(0) = "pCompanyName;" & gCompanyName & ";TRUE"
            'rptProspRdctn.ParameterFields(1) = "pRptType;" & RptTitle & ";TRUE"
            'rptProspRdctn.ParameterFields(2) = "pPctProspect100;" & PctProspect100 & ";TRUE"
            'rptProspRdctn.ParameterFields(3) = "pMineHasOffSpecPbPlt;" & RcvryData.UseOrigMgoPlant & ";TRUE"
            'rptProspRdctn.ParameterFields(4) = "pProdSizeDesig;" & cboProdSizeDefn.Text & ";TRUE"
            'rptProspRdctn.ParameterFields(5) = "pRcvryEtcScen;" & cboOtherDefn.Text & ";TRUE"
            'rptProspRdctn.ParameterFields(6) = "pMineHasDoloflotPlt;" & RcvryData.UseDoloflotPlant2010 & ";TRUE"
            'rptProspRdctn.ParameterFields(7) = "pMineHasDoloflotPltFco;" & RcvryData.UseDoloflotPlantFco & ";TRUE"

            ''Report window maximized
            'rptProspRdctn.WindowState = crptMaximized

            'rptProspRdctn.WindowTitle = "Raw Data Reduction Prospect Data "

            ''User not allowed to minimize report window
            'rptProspRdctn.WindowMinButton = False

            ''Start Crystal Reports
            'rptProspRdctn.action = 1

            Dim ReportForm As New frmReport()
            With ReportForm.CrystalReportViewer1
                .ReportSource = rptProspRdctn
                .LogOnInfo.Item(0).ConnectionInfo.UserID = "mois"
                .LogOnInfo.Item(0).ConnectionInfo.Password = "legs2"
                .Refresh()
            End With

            ReportForm.ShowDialog()


        Catch ex As Exception
            MessageBox.Show("Error printing prospect data from raw data reduction." & vbCrLf &
                           ex.Message)
        End Try

    End Sub

    Private Sub cmdViewCompSplit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdViewCompSplit.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'SetActionStatus("Accessing Comp/Split...")
        'Me.Cursor = Cursors.WaitCursor
        'Dim frmPD As New frmProspectData
        'frmPD.ShowDialog()
        'frmPD.Dispose()
        ''Load(frmProspectData)
        ''frmProspectData.Show(vbModal)
        ''Unload(frmProspectData)

        'SetActionStatus("")
        'Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub cmdProspSec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdProspSec.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SetActionStatus("Accessing Prospect Section...")
        Me.Cursor = Cursors.WaitCursor

        'Load(frmPdProspView)
        'frmPdProspView.Show(vbModal)
        'Unload(frmPdProspView)

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub cmdViewRawProsp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdViewRawProsp.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SetActionStatus("Accessing raw prospect...")
        Me.Cursor = Cursors.WaitCursor

        'Load(frmProspRawDataGen)
        'frmProspRawDataGen.Show(vbModal)
        'Unload(frmProspRawDataGen)

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub cmdCreateSurvCadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreateSurvCadd.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SaveStatus As Boolean
        Dim MineHasOffSpecPbPlt As Boolean
        Dim MineHasDoloflotPlt As Boolean
        Dim UserSetHoleUnmineable As Boolean
        Dim ProspData As gRawProspSplRdctnType
        Dim OvbThk As Single
        Dim SetPbToMgPlt As Boolean

        'Holes only right now!
        '01/12/2010, lss -- Splits are available now.
        If optInclComposites.Checked = False And optInclSplits.Checked = False Then
            MsgBox("Only Holes or Splits are available right now for SurvCADD transfer textfiles!",
                   vbOKOnly, "Illegal Output Option")
            Exit Sub
        End If

        'Is the textfile name OK?
        If Len(Trim(txtSurvCaddTextfile.Text)) = 0 Then
            MsgBox("You must enter a textfilename!", vbOKOnly, "Missing Textfile Name")
            Exit Sub
        Else
            If Mid(txtSurvCaddTextfile.Text, Len(txtSurvCaddTextfile.Text)) = "\" Then
                MsgBox("You must enter a textfilename!", vbOKOnly, "Missing Textfile Name")
                Exit Sub
            End If
        End If

        SetActionStatus("Creating SurvCADD transfer textfile...")
        Me.Cursor = Cursors.WaitCursor

        'gSaveProspectDataset is in modRawProspDataReduction.

        'SurvCADD transfer text files will only be 100%!!

        If lblOffSpecPbMgPlt.Text = "*OffSpec Pb Mg Plt*" Or
            lblOffSpecPbMgPlt.Text = "*Doloflot Plt FCO*" Then
            MineHasOffSpecPbPlt = True
        Else
            MineHasOffSpecPbPlt = False
        End If

        If lblOffSpecPbMgPlt.Text = "*Doloflot Plt Ona*" Then
            MineHasDoloflotPlt = True
        Else
            MineHasDoloflotPlt = False
        End If

        'Make sure that the user has not set the hole to unmineable!
        ProspData = gGetDataFromReviewSprd(ssCompReview, 1)
        UserSetHoleUnmineable = False
        OvbThk = 0
        With ssHoleData
            .Row = 7
            .Col = 3   '01/19/2011, lss  This is still OK -- Col 3 is still Col 3  (TPA)
            If .Value = 0 And ProspData.Tpr.Tpa <> 0 Then
                UserSetHoleUnmineable = True

                With ssDrillData
                    .Row = 2
                    .Col = 2
                    OvbThk = .Value 'This is the depth to the top of the 1st split
                    'and will be the Ovb thk for an unmineable hole.
                End With
            End If
        End With

        If chkPbAnalysisFillInSpecial.Checked = True Then
            SetPbToMgPlt = True
        Else
            SetPbToMgPlt = False
        End If

        '01/12/2010, lss
        'Added this line!
        gFileNumber = -99

        Dim ResultSet As SplitResultSet = GetResultSetFromSpreadSheets(ssSplitReview, ssCompReview)

        SaveStatus = gSaveProspectDataset("SurvCaddText",
                                          "",
                                          txtSurvCaddTextfile.Text,
                                          1,
                                          0,
                                          optInclSplits.Checked,
                                          optInclComposites.Checked,
                                          optInclBoth.Checked,
                                          ssCompReview,
                                          ssSplitReview,
                                          ResultSet.SplitResults,
                                          ResultSet.HoleResults,
                                          MineHasOffSpecPbPlt,
                                          0,
                                          UserSetHoleUnmineable,
                                          OvbThk,
                                          False,
                                          1,
                                          SetPbToMgPlt,
                                          MineHasDoloflotPlt,
                                          False)

        If Not txtSurvCaddTextfile.Equals(String.Empty) Then
            Dim sw As StreamWriter = New StreamWriter(txtSurvCaddTextfile.Text, False)
            sw.Write(gOutputLines)
            sw.Close()
            sw = Nothing
            gOutputLines.Clear()
        End If


        SetActionStatus("")
        Me.Cursor = Cursors.Arrow

        If SaveStatus = True Then
            MsgBox("SurvCADD (100% Prospect) transfer textfile completed -- no problems!",
                   vbOKOnly, "Textfile Create Status")
        Else
            MsgBox("SurvCADD (100% Prospect) transfer textfile completed -- PROBLEMS!",
                   vbOKOnly, "Textfile Create Status")
        End If
    End Sub

    Private Sub lblSurvCaddTxtFile_Click()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        txtSurvCaddTextfile.Text = gGetProspDatasetTfileLoc(gUserName)
    End Sub

    Private Function GetSplitDatesCorrect() As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim RowIdx As Integer
        Dim ThisDate As String

        'All dates for splits should be the same!!!

        ThisDate = ""
        GetSplitDatesCorrect = True

        With ssSplitReview
            For RowIdx = 1 To .MaxRows
                .Row = RowIdx
                .Col = 3
                If .Text <> ThisDate And ThisDate <> "" Then
                    'We have a problem!
                    GetSplitDatesCorrect = False
                    Exit Function
                End If
                ThisDate = .Text
            Next RowIdx
        End With
    End Function

    Private Sub cmdPrtGrd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrtGrd.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo cmdPrtGrdClickError

        gClearGridPrint()

        gGridObject = ssCompErrors

        gPrintGridHeader = "Prospect Data Reduction -- Composite Issues"
        gPrintGridSubHeader1 = "Product size designation name = " &
                               cboProdSizeDefn.Text

        gPrintGridSubHeader2 = "Based on " & cboOtherDefn.Text

        gOrientHeader = "Center"
        gOrientSubHeader1 = "Center"
        gOrientSubHeader2 = "Center"

        gPrintGridFooter = ""
        gOrientFooter = ""
        gSubHead2IsHeader = False

        gPrintGridDefaultTxtFname = ""

        gPrintMarginLeft = 0     '1440 = 1"
        gPrintMarginRight = 0
        gPrintMarginTop = 770
        gPrintMarginBottom = 770

        SetActionStatus("Printing spreadsheet...")
        Me.Cursor = Cursors.WaitCursor
        Print.frmGridToText.ShowDialog()
        Print.frmGridToText.Dispose()
        'Load(frmGridToText)
        'frmGridToText.Show(vbModal)
        'Unload(frmGridToText)

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow

        Exit Sub

cmdPrtGrdClickError:
        MsgBox("Error printing grid." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Grid Error")

        On Error Resume Next
        Me.Cursor = Cursors.Arrow
        SetActionStatus("")
    End Sub

    Private Sub cmdMakeHoleUnmineable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMakeHoleUnmineable.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim OvbThk As Single

        lblUserMadeHoleUnmineable.Text = "User made hole unmineable!"
        lblUserMadeHoleUnmineable.ForeColor = Color.DarkRed ' &HC0&     'Dark red

        'Will set the overburden thickness to the depth to the first split!
        With ssDrillData
            .Row = 2
            .Col = 2
            OvbThk = .Value
        End With

        With ssHoleData
            .BlockMode = True
            .Row = 1
            .Row2 = .MaxRows
            .Col = 1
            .Col2 = 1
            .Text = ""

            .Row = 1
            .Row2 = .MaxRows
            .Col = 2
            .Col2 = 12   '01/19/2011, lss  Was 10
            .Value = 0

            .BlockMode = False

            .Col = 14   '01/19/2011, lss  Was 12
            .Row = 2    'Ttl Wt%
            .Value = 0
            .Row = 3    'Wcl Wt%
            .Value = 0
            .Row = 5    'Cfd Wt%
            .Value = 0
            .Row = 6    'Ffd Wt%
            .Value = 0
            .Row = 7    'Tfd Wt%
            .Value = 0

            .Col = 15   '01/19/2011, lss  Was 13
            .Row = 2    'Ttl TPA
            .Value = 0
            .Row = 3    'Wcl TPA
            .Value = 0
            .Row = 5    'Cfd TPA
            .Value = 0
            .Row = 6    'Ffd TPA
            .Value = 0
            .Row = 7    'Tfd TPA
            .Value = 0

            .Col = 16   '01/19/2011, lss  Was 14
            .Row = 2    'Ttl BPL
            .Value = 0
            .Row = 3    'Wcl BPL
            .Value = 0
            .Row = 5    'Cfd BPL
            .Value = 0
            .Row = 6    'Ffd BPL
            .Value = 0
            .Row = 7    'Tfd BPL
            .Value = 0

            .Col = 18   '01/19/2011, lss  Was 16
            .Row = 1    'Ovb thk'
            .Value = OvbThk
            .Row = 2    'Itb thk'
            .Value = 0
            .Row = 3    'Mtx thk'
            .Value = 0
            .Row = 4    'Mtx"X" All
            .Value = 0
            .Row = 5    'Tot"X" All
            .Value = 0
            .Row = 6    'Density
            .Value = 0
            .Row = 7    '%Solids
            .Value = 0

            .Col = 20   '01/19/2011, lss  Was 18
            .Row = 4    'Mtx"X" OnSpec
            .Value = 0
            .Row = 5    'Tot"X" On Spec
            .Value = 0
            .Row = 6    'Minability
            .Text = "U"
        End With
    End Sub

    Private Sub cmdSaveAreaName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveAreaName.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        If MsgBox("Save area name?", vbYesNo +
            vbDefaultButton1, "Save") = vbYes Then
            SaveAreaName()
        End If
    End Sub

    Private Sub SaveAreaName()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo SaveAreaNameError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim CheckSuccess As Integer

        If Trim(txtAreaName.Text) = "" Then
            MsgBox("Area name may not be blank." +
                    Chr(10) + Chr(10) + "Area name NOT ADDED!" _
                    , vbExclamation, "Error Adding Area Name")
            Exit Sub
        End If

        SetActionStatus("Saving area name...")
        Me.Cursor = Cursors.WaitCursor

        params = gDBParams

        params.Add("pMineName", "Four Corners", ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pCategoryName", "Prospect area", ORAPARM_INPUT)
        params("pCategoryName").serverType = ORATYPE_VARCHAR2

        params.Add("pComboBoxChoiceText", Trim(txtAreaName.Text), ORAPARM_INPUT)
        params("pComboBoxChoiceText").serverType = ORATYPE_VARCHAR2

        params.Add("pStartShiftDate", #1/1/2000#, ORAPARM_INPUT)
        params("pStartShiftDate").serverType = ORATYPE_DATE

        params.Add("pStartShift", "1ST", ORAPARM_INPUT)
        params("pStartShift").serverType = ORATYPE_VARCHAR2

        params.Add("pStopShiftDate", #12/31/8888#, ORAPARM_INPUT)
        params("pStopShiftDate").serverType = ORATYPE_DATE

        params.Add("pStopShift", "NIGHT", ORAPARM_INPUT)
        params("pStopShift").serverType = ORATYPE_VARCHAR2

        params.Add("pOrderNum", 1, ORAPARM_INPUT)
        params("pOrderNum").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        'PROCEDURE update_choices
        'pMineName               IN     VARCHAR2,
        'pCategoryName           IN     VARCHAR2,
        'pComboBoxChoiceText     IN     VARCHAR2,
        'pStartShiftDate         IN     DATE,
        'pStartShift             IN     VARCHAR2,
        'pStopShiftDate          IN     DATE,
        'pStopShift              IN     VARCHAR2,
        'pOrderNum               IN     NUMBER,
        'pResult                 IN OUT NUMBER)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.update_choices(:pMineName," +
                      ":pCategoryName, :pComboBoxChoiceText," +
                      ":pStartShiftDate, :pStartShift, :pStopShiftDate," +
                      ":pStopShift, :pOrderNum, :pResult);end;", ORASQL_FAILEXEC)
        CheckSuccess = params("pResult").Value
        ClearParams(params)

        SetActionStatus("")
        Me.Cursor = Cursors.Arrow

        Exit Sub

SaveAreaNameError:
        MsgBox("Oracle returned an error while attempting to update the data." + Str(Err.Number) + Chr(10) + Chr(10) +
               Err.Description, vbExclamation, "Error Updating Data")

        On Error Resume Next
        ClearParams(params)
        SetActionStatus("")
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Function GetMaxDepthInfo(ByRef aRcvryData As gDataRdctnParamsType) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SplIdx As Integer
        Dim OvbDepthTo1stSpl As Single
        Dim CurrDepth As Single
        Dim ThisSplTop As Single
        Dim ThisSplBot As Single
        Dim ThisSplThk As Single

        GetMaxDepthInfo = "Splits affected by max total depth = "

        With ssDrillData
            .Row = 2
            .Col = 2
            OvbDepthTo1stSpl = .Value
            CurrDepth = OvbDepthTo1stSpl

            For SplIdx = 3 To .MaxRows
                .Row = SplIdx
                .Col = 2
                ThisSplThk = .Value
                ThisSplTop = CurrDepth
                ThisSplBot = ThisSplTop + ThisSplThk

                If ThisSplTop > aRcvryData.MaxTotDepthSpl Or
                    ThisSplBot > aRcvryData.MaxTotDepthSpl Then
                    'This split is affected by the max depth cutoff
                    GetMaxDepthInfo = GetMaxDepthInfo & CStr(SplIdx - 2) & ", "
                End If

                CurrDepth = ThisSplBot
            Next SplIdx
        End With

        If GetMaxDepthInfo = "Splits affected by max total depth = " Then
            GetMaxDepthInfo = "Splits affected by max total depth = None"
        End If
        If InStr(GetMaxDepthInfo, ",") <> 0 Then
            'Strip the extra ", " off
            GetMaxDepthInfo = Mid(GetMaxDepthInfo, 1, Len(GetMaxDepthInfo) - 2)
        End If
    End Function

    Private Function UserCanSaveThisToMois() As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim MinabilitySetUserIsAdmin As Boolean

        'Write -- Save to MOIS if an "Admin" has not already saved it to MOIS.
        'Setup -- Save to MOIS if an "Admin" has not already saved it to MOIS.
        'Admin -- Can do anything (basically only Allen Truesdell).

        'Have fMinabilityUser.
        'If the user is an "Admin" in raw prospect reduction then they can save.
        'If fUserIsAdmin = True Then
        If AppShared.IsUserAdminRole Then
            UserCanSaveThisToMois = True
            Exit Function
        End If

        'If no minabilities have been set for this hole then the user
        If fMinabilityUser = "" Then
            UserCanSaveThisToMois = True
            Exit Function
        End If

        'At this point we know that the current user is not "Admin" for raw prospect reduction!
        'Need to determine if this user that set minability is an administrator!
        If gGetMineUserPermissions(fMinabilityUser,
                                   "Raw Prospect Reduction",
                                   gActiveMineNameLong) = "Admin" Then
            MinabilitySetUserIsAdmin = True
        Else
            MinabilitySetUserIsAdmin = False
        End If

        If MinabilitySetUserIsAdmin = True Then
            UserCanSaveThisToMois = False
        Else
            UserCanSaveThisToMois = True
        End If
    End Function

    Private Sub SetMoisCurrRawMinabilities()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SampleDynaset As OraDynaset
        Dim SampsOk As Boolean
        Dim ThisSplit As Integer
        Dim ThisProspDate As String
        Dim ThisHoleMinable As String
        Dim ThisHoleMinableWhen As String
        Dim ThisHoleMinableWho As String
        Dim ThisSplitMinable As String
        Dim ThisSplitMinableWhen As String
        Dim ThisSplitMinableWho As String
        Dim ItemCnt As Integer

        ssSplitMinabilities.MaxRows = 0
        ItemCnt = 0
        fMinabilityUser = ""

        SampsOk = gGetDrillHoleDateSpec(Val(cboSec.Text),
                                        Val(cboTwp.Text),
                                        Val(cboRge.Text),
                                        cboHole.Text,
                                        1,
                                        SampleDynaset)

        If SampleDynaset.RecordCount <> 0 Then
            SampleDynaset.MoveFirst()
            Do While Not SampleDynaset.EOF
                ThisProspDate = Format(SampleDynaset.Fields("prosp_date").Value, "MM/dd/yyyy")
                ThisSplit = SampleDynaset.Fields("split_number").Value

                If Not IsDBNull(SampleDynaset.Fields("hole_minable").Value) Then
                    If SampleDynaset.Fields("hole_minable").Value = 1 Then
                        ThisHoleMinable = "Yes"
                    Else
                        ThisHoleMinable = "No"
                    End If
                Else
                    ThisHoleMinable = "NA"
                End If
                If Not IsDBNull(SampleDynaset.Fields("hole_minable_when").Value) Then
                    ThisHoleMinableWhen = Format(SampleDynaset.Fields("hole_minable_when").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    ThisHoleMinableWhen = "None"
                End If
                If Not IsDBNull(SampleDynaset.Fields("hole_minable_who").Value) Then
                    ThisHoleMinableWho = SampleDynaset.Fields("hole_minable_who").Value
                    fMinabilityUser = ThisHoleMinableWho
                Else
                    ThisHoleMinableWho = "None"
                End If

                If Not IsDBNull(SampleDynaset.Fields("split_minable").Value) Then
                    If SampleDynaset.Fields("split_minable").Value = 1 Then
                        ThisSplitMinable = "Yes"
                    Else
                        ThisSplitMinable = "No"
                    End If
                Else
                    ThisSplitMinable = "NA"
                End If
                If Not IsDBNull(SampleDynaset.Fields("split_minable_when").Value) Then
                    ThisSplitMinableWhen = Format(SampleDynaset.Fields("split_minable_when").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    ThisSplitMinableWhen = "None"
                End If
                If Not IsDBNull(SampleDynaset.Fields("split_minable_who").Value) Then
                    ThisSplitMinableWho = SampleDynaset.Fields("split_minable_who").Value
                    If fMinabilityUser <> "" Then
                        fMinabilityUser = ThisSplitMinableWho
                    End If
                Else
                    ThisSplitMinableWho = "None"
                End If

                With ssSplitMinabilities
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = 0
                    .Text = "Spl" & CStr(ThisSplit)
                    .TypeHAlign = TypeHAlignConstants.TypeHAlignLeft ' SS_CELL_H_ALIGN_LEFT
                    .Col = 1
                    .Text = ThisProspDate
                    .Col = 2
                    .Text = ThisSplitMinable
                    .Col = 3
                    .Text = ThisSplitMinableWhen
                    .Col = 4
                    .Text = ThisSplitMinableWho
                End With

                If ItemCnt = 0 Then
                    With ssHoleMinabilities
                        .Row = 1
                        .Col = 1
                        .Text = ThisHoleMinable
                        .Row = 2
                        .Text = ThisHoleMinableWhen
                        .Row = 3
                        .Text = ThisHoleMinableWho
                    End With
                End If

                ItemCnt = ItemCnt + 1
                SampleDynaset.MoveNext()
            Loop
        End If

        SampleDynaset.Close()
    End Sub

    Private Sub opt100PctRdctn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles opt100PctRdctn.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        lblSaveToMoisComm.Text = "(100% Prospect Only)"
    End Sub

    Private Sub optCatalogRdctn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCatalogRdctn.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        lblSaveToMoisComm.Text = "(Catalog Only)"
    End Sub

    Private Sub optBothRdctn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optBothRdctn.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        lblSaveToMoisComm.Text = "(100% Prospect && Catalog)"
    End Sub

    Private Sub cboMineName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMineName.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        If cboMineName.Text = "South Fort Meade" Or
            cboMineName.Text = "Hookers Prairie" Or
            cboMineName.Text = "Wingate" Then
            opt100PctRdctn.Checked = True
            optCatalogRdctn.Enabled = False
            optBothRdctn.Enabled = False
        Else
            opt100PctRdctn.Enabled = True
            optCatalogRdctn.Enabled = True
            optBothRdctn.Enabled = True
            optBothRdctn.Checked = True
        End If
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

    Private Sub ssSplitOverrides_Click(ByVal sender As System.Object, ByVal e As AxFPSpread._DSpreadEvents_ClickEvent) Handles ssSplitOverrides.ClickEvent 'ByVal Col As Long, ByVal Row As Long)

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

            'frmProspDataHoleReduction.Refresh()

            DisplaySplitOverrideSet(ThisSplitOrideSetName, "User split override set")
            MarkHolesGreen(ssSplitOverride, True)
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

    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        GetSplitOverrideSets()
        ssSplitOverride.MaxRows = 0
    End Sub

    Private Sub cmdAddToOverrideSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddToOverrideSet.Click

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SaveSplitOverrideSet("User split override set")
        DisplaySetAgain()
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
        'First row is headers.
        'Second row is overburden.
        ItemCount = ssDrillData.MaxRows - 2

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
        With ssDrillData
            For RowIdx = 3 To .MaxRows
                .Row = RowIdx

                ThisTwp = Val(cboTwp.Text)
                ThisRge = Val(cboRge.Text)
                ThisSec = Val(cboSec.Text)
                ThisHole = cboHole.Text
                ThisSplit = RowIdx - 2

                .Col = 1
                If .Text = "1" Then
                    ThisMineable = "M"    'M, U, C
                Else
                    ThisMineable = "U"
                End If

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

        'PROCEDURE update_prosp_split_oride_set2
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
        InsertSQL = "Begin mois.mois_prosp_data_rdctn.update_prosp_split_oride_set2(" &
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
        Dim arA1() As Object = {"pArraySize", ItemCount, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA2() As Object = {"pSplitOrideSetName", txtSplitOverrideName.Text, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA3() As Object = {"pProspSetName", aProspSetName, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA4() As Object = {"pWhoDefined", gUserName, ORAPARM_INPUT, ORATYPE_VARCHAR2}
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

        MsgBox("Split override split(s) saved.", vbOKOnly, "Save Status")

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

    Private Sub DisplaySetAgain()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        DisplaySplitOverrideSet(txtSplitOverrideName.Text, "User split override set")
        MarkHolesGreen(ssSplitOverride, True)
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
    End Sub


    'Private Sub cmdSaveMinabilities_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveMinabilities.Click

    'End Sub

    'Private Sub cboClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboProdSizeDefn.Click

    'End Sub

    'Private Sub cboValidating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboProdSizeDefn.Validating

    'End Sub

    'Private Sub chkClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOverrideMaxDepth.Click

    'End Sub

    'Private Sub optClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles opt100PctRdctn.Click

    'End Sub

    'Private Sub ssButClick(ByVal sender As System.Object, ByVal e As AxFPSpread._DSpreadEvents_ButtonClickedEvent) Handles ssHoleData.ButtonClicked

    'End Sub

    'Private Sub ssclick(ByVal sender As System.Object, ByVal e As AxFPSpread._DSpreadEvents_ClickEvent) Handles ssHoleData.ClickEvent

    'End Sub

    'Private Sub frmProspDataHoleReduction_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    'End Sub

    Private Sub cboProdSizeDefn_IndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboProdSizeDefn.SelectedIndexChanged
        Try

            If cboProdSizeDefn.SelectedIndex > 0 Then
                '**********************************************************************
                '
                '
                '
                '**********************************************************************

                SetReduceEnable()
                ClearData()

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboOtherDefn_SelectIndex(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboOtherDefn.SelectedIndexChanged
        Try

            If cboOtherDefn.SelectedIndex > 0 Then
                '**********************************************************************
                '
                '
                '
                '**********************************************************************

                SetReduceEnable()
                ClearData()

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub cmdExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        If MsgBox("Exit raw prospect data hole reduction program?", vbYesNo +
       vbDefaultButton1, "Exiting Program") = vbYes Then
            Me.Close() ' End 'Unload(Me)
        End If
    End Sub

End Class
