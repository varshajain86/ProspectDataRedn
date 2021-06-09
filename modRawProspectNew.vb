Option Explicit On
Imports OracleInProcServer
Imports System.Math
Imports AxFPSpread

Module modRawProspectNew
    'Attribute VB_Name = "modRawProspectNew"
    '**********************************************************************
    'RAW PROSPECT MODULE NEW -- September 2006 Version
    '
    '
    '**********************************************************************
    '   Maintenance Log
    '
    '   09/19/2006, lss
    '       Added this module.
    '       Added Function gGetProspCodes.
    '   10/31/2006, lss
    '       Added Public Sub gGetProspHoleSplits.
    '   02/21/2007, lss
    '       Finished Function gGetExpandedStatus.
    '   03/05/2007, lss
    '       gGetMetLabErrorsNew will not show any errors for barren splits
    '       from now on. (per Earnest Terry)
    '   03/07/2007, lss
    '       Fixed gGetMtxDensityComp -- Replaced 11 with 181.
    '   03/13/2007, lss
    '       Added Function gGetProspRawCoordsExist.
    '   05/17/2007, lss
    '       Added Public Function gGetIntervalsNew (same as in
    '       modRawProspect).
    '   05/25/2007, lss
    '       Modified error check for low recovery.
    '       Error = Recovery < 75 and feed grade > 8.
    '       Error = Recovery < 40 and any feed grade.
    '   09/05/2007, lss
    '       Added Function gGetHoleDatesInMois.
    '   10/16/2007, lss
    '       Added Function gUpdateRawProspect.
    '       Added Sub gUpdateProspRawSizeFrctn.
    '   10/25/2007, lss
    '       Added Function gGetRawHoleSplCnt.
    '   10/30/2007, lss
    '       Added HoleMinable, HoleMinableWhen, HoleMinableWho,
    '       SpecAreaName, ManufacturedData, SplitMinable, SplitMinableWhen,
    '       SplitMinableWho, SampleIdCargill, BedCode to gRawProspBaseType.
    '       Added HoleMinable, HoleMinableWhen, HoleMinableWho,
    '       SpecAreaName, ManufacturedData to gRawProspBaseHoleType.
    '   10/30/2007, lss
    '       Fixed problem in gGetMetLabErrorsNew when checking for interval
    '       footage problems.
    '   11/14/2007, lss
    '       Added Public Function gGetMineForSampleId.
    '   11/15/2007, lss
    '       Added Public Function gDispMinabilities.
    '       Added Public Function gGetDrillHoleNew()
    '   11/16/2007, lss
    '       Added Public Function gGetProspRawStatus.
    '   11/20/2007, lss
    '       Added Public Function gGetProspCodeDesc.
    '   11/27/2007, lss
    '       Added MoisLoadWhen and MoisLoadWho to gRawProspBaseType.
    '       Added to Function gGetProspRawDataNew also.
    '   11/28/2007, lss
    '       Added Public Sub gSaveRdctnWhoAndWhen.
    '   12/03/2007, lss
    '       Modified Function gGetHoleDatesInMois to look for alpha-numeric
    '       hole locations also.
    '   12/12/2007, lss
    '       Added Public Function gGetRawProspSampleId.
    '   01/03/2007, lss
    '       Added DryDensityOverride to Public Type gRawProspBaseType.
    '       Added DryDensity to Public Function gUpdateRawProspect.
    '   01/28/2008, lss
    '       Modified gGetProspRawCoordsExist -- returns 0, 1, 2, 3, 4 or 5
    '       now.
    '   02/06/2008, lss
    '       Modified Sub gGetRawProspCalcData for Pb & Cn MgO & I&A.
    '   02/11/2008, lss
    '       Replaced get_prosp_raw_hole_lith2 with get_prosp_raw_hole_lith3
    '       in Function gGetDrillHoleSamplesNewLithSpl.
    '   03/27/2008, lss
    '       Added Function gGetSampIdSpec.
    '   04/15/2008, lss
    '       Fixed Public Sub gGetMtxDensityComp so it displays
    '       .DensityProblem correctly.
    '       Added fix to Public Function gCalcDryDensity for 0.33 instead
    '       of 0.3 problem!
    '   04/28/2008, lss
    '       Added ClaySettlingLvl and PbClayPct to gRawProspBaseType.
    '   06/16/2008, lss
    '       Added Function gGetSampIdProspDate.
    '   07/10/2008, lss
    '       Modified Function gGetHoleDatesInMois to show blank prospect
    '       dates as "??/??/????".
    '   08/20/2008, lss
    '       Added Public Sub gGetMineAreaRawProsp.
    '   09/09/2008, lss
    '       Added aWho and aWhen to Sub gSaveRdctnWhoAndWhen.
    '   09/24/2008, lss
    '       Added Function gGetProspRawHoleDataOnly2.
    '       Added Sub gGetProspCodesToGrid.
    '   09/25/2008, lss
    '       Added Function gGetSampIdSpec2.
    '   09/29/2008, lss
    '       Added Sub gFixHoleCol, gFixLithologyCol, gAppendLithology.
    '   10/02/2008, lss
    '       Added Public Function gGetMineAreaSpecAbbrv.
    '       Added aSkipBarrenSplits to gGetMetLabErrorsNew.  As of 10/2008
    '       Earnest wants errors for barren splits to be displayed.
    '   03/02/2009, lss
    '       Added Function gGetDrillHoleDateSpec.
    '   10/16/2009, lss
    '       Changed Feed %solids outside range %solids (75% to 85%) to
    '       75% to 86%.
    '   02/11/2010, lss
    '       Added Public Function gCalcMoist2.
    '       Modified Public Function gCalcMoist.
    '   02/16/2010, lss
    '       Modified Function gUpdateRawProspect -- added the second set of
    '       matrix %moisture weights (wet, dry, tare).  Added QA-QC Hole?
    '
    '       Modified Sub gGetRawProspCalcData -- calls gCalcMoist2 instead
    '       of gCalcMoist in order to handle 2 sets of matrix %moisture
    '       samples.
    '       Modified Function gGetMatlWtPct -- calls gCalcMoist2 instead
    '       of gCalcMoist in order to handle 2 sets of matrix %moisture
    '       samples.
    '   02/16/2010, lss
    '       Added gCalcSolids, gCalcSolids2.
    '       Modified Function gCalcDryDensity.
    '   02/22/2010, lss
    '       Modified Sub gGetMetLabErrorsNew -- added error checking for
    '       matrix %moisture density 1st and 2nd set data problems.
    '   03/15/2010, lss
    '       Added MtxPctMoist1 and MtxPctMoist2 to gRawProspBaseType.
    '   03/31/2010, lss
    '       Added aQaQcHole to Function gGetProspRawStatus.
    '   04/27/2010, lss
    '       Added .QaQcHole to Function gGetProspRawHoleDataOnly.
    '   04/27/2010, lss
    '       Added QaQcHole As Integer to Public Type gRawProspBaseHoleType.
    '   10/18/2011, lss
    '       Modified for Hardpan stuff.
    '   01/13/2012, lss
    '       Added error check for missing clay settling level for Ona and
    '       Wingate.
    '
    '**********************************************************************


    Public Structure gRawProspBaseType
        Public SampleId As String           '1
        Public Township As Integer          '2
        Public Range As Integer            '3
        Public Section As Integer           '4
        Public HoleLocation As String       '5
        Public Forty As Integer             '6
        Public State As String              '7
        Public Quadrant As Integer          '8
        Public MineName As String           '9
        Public ExpDrill As Integer          '10
        Public SplitTotalNum As Integer     '11
        Public Xcoord As Double             '12
        Public Ycoord As Double             '13
        Public FtlDepth As Single           '14
        Public OvbCored As Single           '15
        Public Ownership As String          '16
        Public ProspDate As String          '17
        Public MinedStatus As Integer       '18
        Public Elevation As Single          '19
        Public TotDepth As Single           '20
        Public Aoi As Single                '21
        Public CoordSurveyed As Integer     '22
        Public HoleComment As String        '23
        Public HoleLocationChar As String   '24
        Public WhoModifiedHole As String    '25
        Public WhenModifiedHole As String   '26
        Public LogDate As String            '27
        Public Released As Integer          '28
        Public Redrilled As Integer         '29
        Public RedrillDate As String        '30
        Public UseForReduction As Integer   '31
        '-----
        Public SplitNumber As Integer       '32
        Public Barren As Integer            '33
        Public SplitFtlBottom As Single    '34
        Public MtxTotWetWt As Single        '35
        Public MtxMoistWetWt As Single      '36
        Public MtxMoistDryWt As Single      '37
        Public MtxMoistTareWt As Single     '38
        Public FdTotWetWt As Single         '39
        Public FdTotWetWtMsr As Single      '40
        Public FdMoistWetWt As Single       '41
        Public FdMoistDryWt As Single       '42
        Public FdMoistTareWt As Single      '43
        Public FdScrnSampWt As Single       '44
        Public DensCylSize As Single        '45
        Public DensCylWetWt As Single       '46
        Public DensCylH2oWt As Single       '47
        Public DryDensity As Single         '48
        Public FlotFdWetWt As Single        '49
        Public MtxProcWetWt As Single       '50
        Public ExpExcessWt As Single        '51
        Public MtxColor As String           '52
        Public DegConsol As String          '53
        Public DigChar As String            '54
        Public PumpChar As String           '55
        Public Lithology As String          '56
        Public PhosphColor As String        '57
        Public PhysMineable As Integer      '58
        Public ClaySettChar As String       '59
        Public FdScrnSampWtComp As Single   '60
        Public RecordLocked As Integer      '61
        Public DateChemLab As String        '62
        Public WhoChemLab As String         '63
        Public RerunStatus As Integer       '64
        Public DateRerun As String          '65
        Public MetLabComment As String      '66
        Public ChemLabComment As String     '67
        Public DateMetLab As String         '68
        Public WhoMetLab As String          '69
        Public SplitDepthTop As Single      '70
        Public SplitDepthBot As Single      '71
        Public SplitThck As Single          '72
        Public WashDate As String           '73
        Public WhoModifiedSplit As String   '74
        Public WhenModifiedSplit As String  '75
        Public OrigData As Integer          '76
        '-----
        Public MtxPctMoist As Single        '77
        Public FdPctMoist As Single         '78
        Public FlotFdBplActual As Single    '79
        Public FlotFdWtActual As Single     '80
        Public FlotFdBplCalc As Single      '81
        Public FlotFdWtCalc As Single       '82
        Public FlotPctRcvry As Single       '83
        Public FdBplDiff As Single          '84
        Public PbWtPct As Single            '85
        Public FdWtPct As Single            '86
        Public ClWtPct As Single            '87
        Public FdWtDiff As Single           '88
        '----
        Public County As String             '89
        Public BankCode As String           '90
        '-----
        Public HoleMinable As Integer       '91
        Public HoleMinableWhen As String    '92
        Public HoleMinableWho As String     '93
        Public SpecAreaName As String       '94
        Public ManufacturedData As Integer  '95
        '-----
        Public SplitMinable As Integer      '96
        Public SplitMinableWhen As String   '97
        Public SplitMinableWho As String    '98
        Public SampleIdCargill As String    '99
        Public BedCode As String            '100
        '-----
        Public MoisLoadWhen As String       '101
        Public MoisLoadWho As String        '102
        '-----
        Public DryDensityOverride As Single '103
        Public ClaySettlingLvl As Integer   '104
        Public PbClayPct As Integer         '105
        '-----
        'Added 2/11/2010, lss
        Public MtxMoistWetWt2 As Single     '106
        Public MtxMoistDryWt2 As Single     '107
        Public MtxMoistTareWt2 As Single    '108
        '-----
        'Added 2/16/2010, lss
        Public QaQcHole As Integer          '109
        '-----
        'Added 3/15/2010, lss
        Public MtxPctMoist1 As Single       '110
        Public MtxPctMoist2 As Single       '111
        '-----
        'Added 10/18/2011, lss
        Public HardpanFrom As Single        '112
        Public HardpanTo As Single          '113
        Public HardpanCode As String        '114
        Public HardpanThck As Single        '115
    End Structure

    Public Structure gRawProspBaseHoleType
        Public SampleId As String           '1
        Public Township As Integer          '2
        Public Range As Integer            '3
        Public Section As Integer           '4
        Public HoleLocation As String       '5
        Public Forty As Integer             '6
        Public State As String              '7
        Public Quadrant As Integer          '8
        Public MineName As String           '9
        Public ExpDrill As Integer          '10
        Public SplitTotalNum As Integer     '11
        Public Xcoord As Double             '12
        Public Ycoord As Double             '13
        Public FtlDepth As Single           '14
        Public OvbCored As Single           '15
        Public Ownership As String          '16
        Public ProspDate As String          '17
        Public MinedStatus As Integer       '18
        Public Elevation As Single          '19
        Public TotDepth As Single           '20
        Public Aoi As Single                '21
        Public CoordSurveyed As Integer     '22
        Public HoleComment As String        '23
        Public HoleLocationChar As String   '24
        Public WhoModifiedHole As String    '25
        Public WhenModifiedHole As String   '26
        Public LogDate As String            '27
        Public Released As Integer          '28
        Public Redrilled As Integer         '29
        Public RedrillDate As String        '30
        Public UseForReduction As Integer   '31
        '-----
        Public County As String             '32
        Public BankCode As String           '33
        '-----
        Public HoleMinable As Integer       '34
        Public HoleMinableWhen As String    '35
        Public HoleMinableWho As String     '36
        Public SpecAreaName As String       '37
        Public ManufacturedData As Integer  '38
        Public SavedMoisWhen As String      '39
        Public SavedMoisWho As String       '40
        '-----
        Public QaQcHole As Integer          '41
        '-----
        'Added 10/18/2011, lss
        Public HardpanFrom As Single        '42
        Public HardpanTo As Single          '43
        Public HardpanCode As String        '44
        Public HardpanThck As Single        '45
    End Structure

    Public Structure gRawProspLoctnType
        Public SampleId As String
        Public Township As Integer
        Public Range As Integer
        Public Section As Integer
        Public HoleLocation As String
        Public SplitNumber As Integer
        Public ProspDate As String
    End Structure

    Public Structure gRawProspSfcType
        Public SampleId As String           '1
        Public Township As Integer          '2
        Public Range As Integer            '3
        Public Section As Integer           '4
        Public HoleLocation As String       '5
        Public ProspDate As String          '6
        Public SplitNumber As Integer       '7
        Public SizeFrctnCode As String      '8
        Public SfcDescription As String     '9
        Public SfcMatlName As String        '10
        Public SfcMatlAbbrv As String       '11
        Public SfcOrderNum As Integer       '12
        Public Bpl As Single                '13
        Public FeAl As Single               '14
        Public Insol As Single              '15
        Public CaO As Single                '16
        Public MgO As Single                '17
        Public Fe2O3 As Single              '18
        Public Al2O3 As Single              '19
        Public Cd As Single                 '20
        Public SizeFrctnWt As Single        '21    Adjusted values for "T" types
        Public SizeFrctnWtMsr As Single     '22    Measured value for "T" types
        Public SizeFrctnType As String      '23
        Public WhoModified As String        '24
        Public WhenModified As String       '25    09/15/2008, lss -- Changed from Date to String
        Public OrderNum As Integer          '26
    End Structure

    Public Structure gRawProspSfcSprdType
        Public SizeFrctnCode As String      '1
        Public SfcDescription As String     '2
        Public SfcMatlName As String        '3
        Public SfcOrderNum As Integer       '4
        Public Bpl As Single                '5
        Public FeAl As Single               '6
        Public Insol As Single              '7
        Public CaO As Single                '8
        Public MgO As Single                '9
        Public Fe2O3 As Single              '10
        Public Al2O3 As Single              '11
        Public Cd As Single                 '12
        Public SizeFrctnWt As Single        '13
        Public SizeFrctnWtMsr As Single     '14
        Public SizeFrctnWtAdj As Single     '15
        Public SizeFrctnType As String      '16
    End Structure

    'This type is sort of based on the "MET LAB PROSPECTING DATA --- (ALL TYPES VER 3.1)
    'report from the met lab at Four Corners (Glen Oswald's programs)
    Public Structure gRawProspCalcType
        Public SplitThck As Single          '1
        Public HeadCalc As Single           '2
        Public HeadMgo As Single            '3
        Public CorePctSol As Single         '4
        Public FdPctSol As Single           '5
        Public CorePctMoist As Single       '6
        Public FdPctMoist As Single         '7
        Public CoreLbsPerFt As Single       '8
        Public DryDensity As Single         '9
        Public XmitLbs As Single            '10
        Public FlotBpl As Single            '11
        Public FlotActWt As Single          '12
        Public FlotCalcWt As Single         '13
        Public Plus35FdPct As Single        '14
        Public Plus35FdWt As Single         '15
        Public Plus35FdBpl As Single        '16
        Public Minus35FdPct As Single       '17
        Public Minus35FdWt As Single        '18
        Public Minus35FdBpl As Single       '19
        Public PbBpl As Single              '20
        Public FdBpl As Single              '21
        Public ClBpl As Single              '22
        Public CnBpl As Single              '23
        Public GmtBpl As Single             '24
        Public CnIns As Single              '25
        Public FlotPctRcvry As Single       '26
        Public DryPbLbsAdj As Single        '27
        Public DryFdLbsAdj As Single        '28
        Public DryClLbsAdj As Single        '29
        Public DryCoreLbsTot As Single      '30
        Public DryCoreLbsProc As Single     '31
        Public PbPctWt As Single            '32
        Public FdPctWt As Single            '33
        Public ClPctWt As Single            '34
        Public TotPctWt As Single           '35
        '----
        Public SampleId As String           '36
        Public Township As Integer          '37
        Public Range As Integer            '38
        Public Section As Integer           '39
        Public HoleLocation As String       '40
        Public SplitNumber As Integer       '41
        Public ProspDate As String          '42
        '----
        Public PbWtDryGms As Single         '43
        Public FdWtDryGms As Single         '44
        '-----
        'Added 02/06/2008, lss
        Public PbMgO As Single              '45
        Public PbIa As Single               '46
        Public CnMgO As Single              '47
        Public CnIa As Single               '48
    End Structure

    Public Structure gZzzRatioDataType
        Public Bpl As Single
        Public Ins As Single
        Public InsCalc As Single
        Public Mg As Single
        Public Fe As Single
        Public Al As Single
        Public ZzzRatio As Single
        Public ZlBpl As Single
        Public AssayOk As Boolean
    End Structure

    Public Structure gSfcDataType
        Public Weight As Single
        Public Bpl As Single
        Public Insol As Single
        Public CaO As Single
        Public Fe2O3 As Single
        Public Al2O3 As Single
        Public FeAl As Single
        Public MgO As Single
        Public Cd As Single
        Public Type As String
    End Structure

    Public Structure gMtxPctSolModType
        Public SplitDepthBot As Single
        Public Plus35FdBpl As Single
        Public Plus35FdWt As Single
        Public Minus35FdBpl As Single
        Public Minus35FdWt As Single
        Public DryFdLbsAdj As Single
        Public PbWtDryGms As Single
        Public MtxTotWetWt As Single
        Public CorePctSol As Single
        Public PbPctWt As Single
        Public FdPctWt As Single
        Public ClPctWt As Single
        Public DryCoreLbsTot As Single
        Public DryCoreLbsProc As Single
        Public FdBplCalc As Single
        Public PctSand As Single
        Public PctMoistCalc As Single
        Public PctSolidsModel As Single
        Public PctSolidsLowerLimit As Single
        Public PctMoistProblem As Boolean
    End Structure

    Public Structure gMtxDensityModType
        Public MtxPctMoist As Single
        Public MtxPctSol As Single
        Public PbPctWt As Single
        Public FdPctWt As Single
        Public ClPctWt As Single
        Public FdBpl As Single
        Public DenFac As Single
        Public CalDen1 As Single
        Public CalDen As Single
        Public LowerLimit As Single
        Public UpperLimit As Single
        Public LabMsrdDryDensity As Single
        Public DensityProblem As Boolean
    End Structure

    Public Structure gFlotCalcDataType
        Public FdBplAct As Single
        Public FdWtAct As Single
        Public FdBplCalc As Single
        Public FdWtCalc As Single
        Public Rcvry As Single
        Public CnBplAct As Single
        Public TlBplAct As Single
    End Structure

    Public Structure gSfcPlusMinusDataType
        Public PlusWt As Single
        Public PlusBpl As Single
        Public MinusWt As Single
        Public MinusBpl As Single
    End Structure

    Public Structure gRawProspCoordType
        Public SampleId As String
        Public Township As Integer
        Public Range As Integer
        Public Section As Integer
        Public HoleLocation As String
        Public Xcoord As Double
        Public Ycoord As Double
    End Structure

    Public Function gGetProspCodes(ByVal aProspCodeTypeName As String,
                                   ByVal aDescriptions As Boolean,
                                   ByVal aAddSelect As Boolean,
                                   ByVal aAddBlankSelection As Boolean) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspCodesError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim CodeDynaset As OraDynaset
        Dim ThisCode As String
        Dim ThisCodeDesc As String
        Dim CodeList As String

        gGetProspCodes = ""

        params = gDBParams

        params.Add("pProspCodeTypeName", aProspCodeTypeName, ORAPARM_INPUT)
        params("pProspCodeTypeName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_codes
        'pProspCodeTypeName   IN     VARCHAR2,
        'pResult              IN OUT c_prospcodes)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospectnew.get_prosp_codes(" +
                      ":pProspCodeTypeName, :pResult);end;", ORASQL_FAILEXEC)
        CodeDynaset = params("pResult").Value
        ClearParams(params)

        'Create a combo box selection string for use in a spreadsheet
        If aAddSelect = True Then
            CodeList = "(Select...)"
        Else
            CodeList = ""
        End If

        If aAddBlankSelection = True Then
            If CodeList = "" Then
                CodeList = " "
            Else
                CodeList = CodeList + Chr(9) + " "
            End If
        End If

        CodeDynaset.MoveFirst()

        Do While Not CodeDynaset.EOF
            ThisCode = CodeDynaset.Fields("prosp_code").Value
            ThisCodeDesc = CodeDynaset.Fields("prosp_code_desc").Value

            If aDescriptions = True Then
                CodeList = CodeList + Chr(9) + ThisCode & "-" &
                       ThisCodeDesc
            Else
                CodeList = CodeList + Chr(9) + ThisCode
            End If

            CodeDynaset.MoveNext()
        Loop

        CodeDynaset.Close()
        gGetProspCodes = CodeList

        Exit Function

gGetProspCodesError:
        MsgBox("Error getting prospect codes." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Prospect Codes Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        CodeDynaset.Close()
        gGetProspCodes = ""
    End Function

    Public Function gGetProspRawDataNew(ByVal aSampleId As String,
                                        ByVal aTwp As Integer,
                                        ByVal aRge As Integer,
                                        ByVal aSec As Integer,
                                        ByVal aHloc As String,
                                        ByVal aProspDate As Date,
                                        ByVal aSplitNum As Integer,
                                        ByRef aProspRawData As gRawProspBaseType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspRawDataError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ProspRawDynaset As OraDynaset
        Dim RecordCount As Integer
        Dim NeedSplitOnly As Boolean

        If aTwp = 0 Then
            NeedSplitOnly = True
        Else
            NeedSplitOnly = False
        End If

        If NeedSplitOnly = False Then
            'Hole level data  Hole level data  Hole level data
            'Hole level data  Hole level data  Hole level data
            'Hole level data  Hole level data  Hole level data

            'PROCEDURE get_prosp_raw_base
            'pTownship           IN     NUMBER,
            'pRange              IN     NUMBER,
            'pSection            IN     NUMBER,
            'pHoleLocation       IN     VARCHAR2,
            'pProspDate          IN     DATE,
            'pResult             IN OUT c_prosprawbase)

            params = gDBParams

            params.Add("pTownship", aTwp, ORAPARM_INPUT)
            params("pTownship").serverType = ORATYPE_NUMBER

            params.Add("pRange", aRge, ORAPARM_INPUT)
            params("pRange").serverType = ORATYPE_NUMBER

            params.Add("pSection", aSec, ORAPARM_INPUT)
            params("pSection").serverType = ORATYPE_NUMBER

            params.Add("pHoleLocation", aHloc, ORAPARM_INPUT)
            params("pHoleLocation").serverType = ORATYPE_VARCHAR2

            params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
            params("pProspDate").serverType = ORATYPE_DATE

            params.Add("pResult", 0, ORAPARM_OUTPUT)
            params("pResult").serverType = ORATYPE_CURSOR

            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_base(" &
                      ":pTownship, :pRange, :pSection, :pHoleLocation, " &
                      ":pProspDate, :pResult);end;", ORASQL_FAILEXEC)

            ProspRawDynaset = params("pResult").Value
            ClearParams(params)

            RecordCount = ProspRawDynaset.RecordCount

            'Should be only one record returned!
            ProspRawDynaset.MoveFirst()

            With aProspRawData
                .Township = ProspRawDynaset.Fields("township").Value
                .Range = ProspRawDynaset.Fields("range").Value
                .Section = ProspRawDynaset.Fields("section").Value
                .HoleLocation = ProspRawDynaset.Fields("hole_location").Value
                .Forty = ProspRawDynaset.Fields("forty").Value
                .State = ProspRawDynaset.Fields("state").Value
                .Quadrant = ProspRawDynaset.Fields("quadrant").Value

                If Not IsDBNull(ProspRawDynaset.Fields("mine_name").Value) Then
                    .MineName = ProspRawDynaset.Fields("mine_name").Value
                Else
                    .MineName = ""
                End If

                .ExpDrill = ProspRawDynaset.Fields("exp_drill").Value
                .SplitTotalNum = ProspRawDynaset.Fields("split_total_num").Value
                .Xcoord = ProspRawDynaset.Fields("x_coord").Value
                .Ycoord = ProspRawDynaset.Fields("y_coord").Value
                .FtlDepth = ProspRawDynaset.Fields("ftl_depth").Value
                .OvbCored = ProspRawDynaset.Fields("ovb_cored").Value
                .Ownership = ProspRawDynaset.Fields("ownership").Value

                .ProspDate = Format(ProspRawDynaset.Fields("prosp_date").Value, "MM/dd/yyyy")

                .MinedStatus = ProspRawDynaset.Fields("mined_status").Value
                .Elevation = ProspRawDynaset.Fields("elevation").Value
                .TotDepth = ProspRawDynaset.Fields("tot_depth").Value
                .Aoi = ProspRawDynaset.Fields("aoi").Value
                .CoordSurveyed = ProspRawDynaset.Fields("coord_surveyed").Value

                If Not IsDBNull(ProspRawDynaset.Fields("long_comment").Value) Then
                    .HoleComment = ProspRawDynaset.Fields("long_comment").Value
                Else
                    .HoleComment = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("hole_location_char").Value) Then
                    .HoleLocationChar = ProspRawDynaset.Fields("hole_location_char").Value
                Else
                    .HoleLocationChar = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("who_modified").Value) Then
                    .WhoModifiedHole = ProspRawDynaset.Fields("who_modified").Value
                Else
                    .WhoModifiedHole = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("when_modified").Value) Then
                    'Want date and time!
                    .WhenModifiedHole = Format(ProspRawDynaset.Fields("when_modified").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .WhenModifiedHole = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("log_date").Value) Then
                    .LogDate = Format(ProspRawDynaset.Fields("log_date").Value, "MM/dd/yyyy")
                Else
                    .LogDate = ""
                End If

                .Released = ProspRawDynaset.Fields("released").Value
                .Redrilled = ProspRawDynaset.Fields("redrilled").Value

                If Not IsDBNull(ProspRawDynaset.Fields("redrill_date").Value) Then
                    .RedrillDate = Format(ProspRawDynaset.Fields("redrill_date").Value, "MM/dd/yyyy")
                Else
                    .RedrillDate = ""
                End If

                .UseForReduction = ProspRawDynaset.Fields("use_for_reduction").Value

                If Not IsDBNull(ProspRawDynaset.Fields("county").Value) Then
                    .County = ProspRawDynaset.Fields("county").Value
                Else
                    .County = ""
                End If
                If Not IsDBNull(ProspRawDynaset.Fields("bank_code").Value) Then
                    .BankCode = ProspRawDynaset.Fields("bank_code").Value
                Else
                    .BankCode = ""
                End If

                '-----
                'New columns added 10/30/2007, lss

                If Not IsDBNull(ProspRawDynaset.Fields("hole_minable").Value) Then
                    .HoleMinable = ProspRawDynaset.Fields("hole_minable").Value
                Else
                    'A null hole minable value will be represented with -1.
                    'It will be displayed as "NA" = Not assigned.
                    .HoleMinable = -1
                End If
                If IsDate(ProspRawDynaset.Fields("hole_minable_when").Value) Then
                    'Want date and time!
                    .HoleMinableWhen = Format(ProspRawDynaset.Fields("hole_minable_when").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .HoleMinableWhen = ""
                End If
                If Not IsDBNull(ProspRawDynaset.Fields("hole_minable_who").Value) Then
                    .HoleMinableWho = ProspRawDynaset.Fields("hole_minable_who").Value
                Else
                    .HoleMinableWho = ""
                End If
                If Not IsDBNull(ProspRawDynaset.Fields("spec_area_name").Value) Then
                    .SpecAreaName = ProspRawDynaset.Fields("spec_area_name").Value
                Else
                    .SpecAreaName = " "
                End If
                .ManufacturedData = ProspRawDynaset.Fields("manufactured_data").Value
                .QaQcHole = ProspRawDynaset.Fields("qaqc_hole").Value

                '-----
                'New columns added 11/27/2007, lss

                If IsDate(ProspRawDynaset.Fields("saved_mois_when").Value) Then
                    'Want date and time!
                    .MoisLoadWhen = Format(ProspRawDynaset.Fields("saved_mois_when").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .MoisLoadWhen = ""
                End If
                If Not IsDBNull(ProspRawDynaset.Fields("saved_mois_who").Value) Then
                    .MoisLoadWho = ProspRawDynaset.Fields("saved_mois_who").Value
                Else
                    .MoisLoadWho = ""
                End If

                '-----
                'New columns added 10/17/2011, lss
                If Not IsDBNull(ProspRawDynaset.Fields("hardpan_from").Value) Then
                    .HardpanFrom = ProspRawDynaset.Fields("hardpan_from").Value
                Else
                    .HardpanFrom = 0
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("hardpan_to").Value) Then
                    .HardpanTo = ProspRawDynaset.Fields("hardpan_to").Value
                Else
                    .HardpanTo = 0
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("hardpan_code").Value) Then
                    .HardpanCode = ProspRawDynaset.Fields("hardpan_code").Value
                Else
                    .HardpanCode = "0"
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("hardpan_thck").Value) Then
                    .HardpanThck = ProspRawDynaset.Fields("hardpan_thck").Value
                Else
                    .HardpanThck = 0
                End If
            End With
            ProspRawDynaset.Close()
        End If

        'Split level data  Split level data  Split level data
        'Split level data  Split level data  Split level data
        'Split level data  Split level data  Split level data

        params = gDBParams

        params.Add("pSampleId", aSampleId, ORAPARM_INPUT)
        params("pSampleId").serverType = ORATYPE_VARCHAR2

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_VARCHAR2

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHloc, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pSplitNumber", aSplitNum, ORAPARM_INPUT)
        params("pSplitNumber").serverType = ORATYPE_NUMBER

        params.Add("pSampleIdOnly", 1, ORAPARM_INPUT)
        params("pSampleIdOnly").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_raw_split
        'pSampleId           IN     VARCHAR2
        'pTownship           IN     NUMBER,
        'pRange              IN     NUMBER,
        'pSection            IN     NUMBER,
        'pHoleLocation       IN     VARCHAR2,
        'pProspDate          IN     DATE,
        'pSplitNumber        IN     NUMBER,
        'pSampleIdOnly       IN     NUMBER,
        'pResult             IN OUT c_prosprawsplit)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_split(" &
                  ":pSampleId, :pTownship, :pRange, :pSection, :pHoleLocation, " &
                  ":pProspDate, :pSplitNumber, :pSampleIdOnly, :pResult);end;", ORASQL_FAILEXEC)

        ProspRawDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = ProspRawDynaset.RecordCount

        'Should be only one record returned!
        ProspRawDynaset.MoveFirst()

        If RecordCount = 1 Then
            With aProspRawData
                If NeedSplitOnly = True Then
                    .Township = ProspRawDynaset.Fields("township").Value
                    .Range = ProspRawDynaset.Fields("range").Value
                    .Section = ProspRawDynaset.Fields("section").Value
                    .HoleLocation = ProspRawDynaset.Fields("hole_location").Value
                    .ProspDate = Format(ProspRawDynaset.Fields("prosp_date").Value, "MM/dd/yyyy")
                End If

                .SampleId = ProspRawDynaset.Fields("sample_id").Value
                .SplitNumber = ProspRawDynaset.Fields("split_number").Value
                .Barren = ProspRawDynaset.Fields("barren").Value
                .SplitFtlBottom = ProspRawDynaset.Fields("split_ftl_bottom").Value
                .MtxTotWetWt = ProspRawDynaset.Fields("mtx_tot_wet_wt").Value

                .MtxMoistWetWt = ProspRawDynaset.Fields("mtx_moist_wet_wt").Value
                .MtxMoistDryWt = ProspRawDynaset.Fields("mtx_moist_dry_wt").Value
                .MtxMoistTareWt = ProspRawDynaset.Fields("mtx_moist_tare_wt").Value

                .MtxMoistWetWt2 = ProspRawDynaset.Fields("mtx_moist_wet_wt2").Value
                .MtxMoistDryWt2 = ProspRawDynaset.Fields("mtx_moist_dry_wt2").Value
                .MtxMoistTareWt2 = ProspRawDynaset.Fields("mtx_moist_tare_wt2").Value

                .FdTotWetWt = ProspRawDynaset.Fields("fd_tot_wet_wt").Value
                .FdTotWetWtMsr = ProspRawDynaset.Fields("fd_tot_wet_wt_msr").Value
                .FdMoistWetWt = ProspRawDynaset.Fields("fd_moist_wet_wt").Value
                .FdMoistDryWt = ProspRawDynaset.Fields("fd_moist_dry_wt").Value
                .FdMoistTareWt = ProspRawDynaset.Fields("fd_moist_tare_wt").Value
                .FdScrnSampWt = ProspRawDynaset.Fields("fd_scrn_samp_wt").Value
                .DensCylSize = ProspRawDynaset.Fields("dens_cyl_size").Value
                .DensCylWetWt = ProspRawDynaset.Fields("dens_cyl_wet_wt").Value
                .DensCylH2oWt = ProspRawDynaset.Fields("dens_cyl_h2o_wt").Value
                .DryDensity = ProspRawDynaset.Fields("dry_density").Value
                .FlotFdWetWt = ProspRawDynaset.Fields("flot_wet_wt").Value
                .MtxProcWetWt = ProspRawDynaset.Fields("mtx_proc_wet_wt").Value
                .ExpExcessWt = ProspRawDynaset.Fields("exp_excess_wt").Value

                If Not IsDBNull(ProspRawDynaset.Fields("mtx_color").Value) Then
                    .MtxColor = ProspRawDynaset.Fields("mtx_color").Value
                Else
                    .MtxColor = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("deg_consol").Value) Then
                    .DegConsol = ProspRawDynaset.Fields("deg_consol").Value
                Else
                    .DegConsol = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("dig_char").Value) Then
                    .DigChar = ProspRawDynaset.Fields("dig_char").Value
                Else
                    .DigChar = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("pump_char").Value) Then
                    .PumpChar = ProspRawDynaset.Fields("pump_char").Value
                Else
                    .PumpChar = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("lithology").Value) Then
                    .Lithology = ProspRawDynaset.Fields("lithology").Value
                Else
                    .Lithology = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("phosph_color").Value) Then
                    .PhosphColor = ProspRawDynaset.Fields("phosph_color").Value
                Else
                    .PhosphColor = ""
                End If

                .PhysMineable = ProspRawDynaset.Fields("phys_mineable").Value

                If Not IsDBNull(ProspRawDynaset.Fields("clay_sett_char").Value) Then
                    .ClaySettChar = ProspRawDynaset.Fields("clay_sett_char").Value
                Else
                    .ClaySettChar = ""
                End If

                .FdScrnSampWtComp = ProspRawDynaset.Fields("fd_scrn_samp_wt_comp").Value
                .RecordLocked = ProspRawDynaset.Fields("record_locked").Value

                If Not IsDBNull(ProspRawDynaset.Fields("date_chem_lab").Value) Then
                    'Want date and time!
                    .DateChemLab = Format(ProspRawDynaset.Fields("date_chem_lab").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .DateChemLab = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("who_chem_lab").Value) Then
                    .WhoChemLab = ProspRawDynaset.Fields("who_chem_lab").Value
                Else
                    .WhoChemLab = ""
                End If

                .RerunStatus = ProspRawDynaset.Fields("rerun_status").Value

                If Not IsDBNull(ProspRawDynaset.Fields("date_rerun").Value) Then
                    .DateRerun = Format(ProspRawDynaset.Fields("date_rerun").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .DateRerun = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("metlab_comment").Value) Then
                    .MetLabComment = ProspRawDynaset.Fields("metlab_comment").Value
                Else
                    .MetLabComment = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("chemlab_comment").Value) Then
                    .ChemLabComment = ProspRawDynaset.Fields("chemlab_comment").Value
                Else
                    .ChemLabComment = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("date_met_lab").Value) Then
                    'Want date and time!
                    .DateMetLab = Format(ProspRawDynaset.Fields("date_met_lab").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .DateMetLab = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("who_met_lab").Value) Then
                    .WhoMetLab = ProspRawDynaset.Fields("who_met_lab").Value
                Else
                    .WhoMetLab = ""
                End If

                .SplitDepthTop = ProspRawDynaset.Fields("split_depth_top").Value
                .SplitDepthBot = ProspRawDynaset.Fields("split_depth_bot").Value
                .SplitThck = ProspRawDynaset.Fields("split_thck").Value

                If Not IsDBNull(ProspRawDynaset.Fields("wash_date").Value) Then
                    .WashDate = Format(ProspRawDynaset.Fields("wash_date").Value, "MM/dd/yyyy")
                Else
                    .WashDate = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("who_modified").Value) Then
                    .WhoModifiedSplit = ProspRawDynaset.Fields("who_modified").Value
                Else
                    .WhoModifiedSplit = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("when_modified").Value) Then
                    .WhenModifiedSplit = Format(ProspRawDynaset.Fields("when_modified").Value, "MM/dd/yyyy")
                Else
                    .WhenModifiedSplit = ""
                End If

                .OrigData = ProspRawDynaset.Fields("orig_data").Value

                '-----
                'New columns added 10/30/2007, lss

                If Not IsDBNull(ProspRawDynaset.Fields("split_minable").Value) Then
                    .SplitMinable = ProspRawDynaset.Fields("split_minable").Value
                Else
                    'A null hole minable value will be represented with -1.
                    'It will be displayed as "NA" = Not assigned.
                    .SplitMinable = -1
                End If
                If IsDate(ProspRawDynaset.Fields("split_minable_when").Value) Then
                    'Want date and time!
                    .SplitMinableWhen = Format(ProspRawDynaset.Fields("split_minable_when").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .SplitMinableWhen = ""
                End If
                If Not IsDBNull(ProspRawDynaset.Fields("split_minable_who").Value) Then
                    .SplitMinableWho = ProspRawDynaset.Fields("split_minable_who").Value
                Else
                    .SplitMinableWho = ""
                End If
                If Not IsDBNull(ProspRawDynaset.Fields("sample_id_cargill").Value) Then
                    .SampleIdCargill = ProspRawDynaset.Fields("sample_id_cargill").Value
                Else
                    .SampleIdCargill = ""
                End If
                If Not IsDBNull(ProspRawDynaset.Fields("bed_code").Value) Then
                    .BedCode = ProspRawDynaset.Fields("bed_code").Value
                Else
                    .BedCode = ""
                End If

                .ClaySettlingLvl = ProspRawDynaset.Fields("clay_settling_lvl").Value
                .PbClayPct = ProspRawDynaset.Fields("pb_clay_pct").Value

                '03/15/2010, lss  Added this functionality.
                'Need to calculate the matrix %moisture(s) here.
                If .MtxMoistWetWt > 0 And .MtxMoistWetWt2 > 0 Then
                    'Should have two %moistures -- lets calculate them here.
                    .MtxPctMoist1 = gCalcMoist(.MtxMoistWetWt,
                                               .MtxMoistDryWt,
                                               .MtxMoistTareWt)

                    .MtxPctMoist2 = gCalcMoist(.MtxMoistWetWt2,
                                               .MtxMoistDryWt2,
                                               .MtxMoistTareWt2)
                Else
                    .MtxPctMoist1 = .MtxPctMoist1
                    .MtxPctMoist2 = 0
                End If
            End With
        Else
            With aProspRawData
                .SampleId = ProspRawDynaset.Fields("sample_id").Value
            End With

            MsgBox("Error -- multiple sample#.",
                    vbOKOnly + vbExclamation,
                    "Too Many Sample#'s Error")
        End If

        ProspRawDynaset.Close()

        gGetProspRawDataNew = True

        Exit Function

gGetProspRawDataError:
        gGetProspRawDataNew = False

        MsgBox("Error accessing raw prospect data." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Raw Prospect Data Access Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        ProspRawDynaset.Close()
    End Function

    Public Sub gZeroProspRawDataNew(ByRef aProspRawData As gRawProspBaseType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        With aProspRawData
            .SampleId = ""
            .Township = 0
            .Range = 0
            .Section = 0
            .HoleLocation = ""
            .Forty = 0
            .State = ""
            .Quadrant = 0
            .MineName = ""
            .ExpDrill = 0
            .SplitTotalNum = 0
            .Xcoord = 0
            .Ycoord = 0
            .FtlDepth = 0
            .OvbCored = 0
            .Ownership = ""
            .ProspDate = ""
            .MinedStatus = 0
            .Elevation = 0
            .TotDepth = 0
            .Aoi = 0
            .CoordSurveyed = 0
            .HoleComment = ""
            .HoleLocationChar = ""
            .WhoModifiedHole = ""
            .WhenModifiedHole = ""
            .LogDate = ""
            .Released = 0
            .Redrilled = 0
            .RedrillDate = ""
            .UseForReduction = 0
            '-----
            .SplitNumber = 0
            .Barren = 0
            .SplitFtlBottom = 0
            .MtxTotWetWt = 0
            '-----
            .MtxMoistWetWt = 0
            .MtxMoistDryWt = 0
            .MtxMoistTareWt = 0
            .MtxMoistWetWt2 = 0
            .MtxMoistDryWt2 = 0
            .MtxMoistTareWt2 = 0
            '-----
            .FdTotWetWt = 0
            .FdTotWetWtMsr = 0
            .FdMoistWetWt = 0
            .FdMoistDryWt = 0
            .FdMoistTareWt = 0
            .FdScrnSampWt = 0
            .DensCylSize = 0
            .DensCylWetWt = 0
            .DensCylH2oWt = 0
            .DryDensity = 0
            .FlotFdWetWt = 0
            .MtxProcWetWt = 0
            .ExpExcessWt = 0
            .MtxColor = ""
            .DegConsol = ""
            .DigChar = ""
            .PumpChar = ""
            .Lithology = ""
            .PhosphColor = ""
            .PhysMineable = 0
            .ClaySettChar = ""
            .FdScrnSampWtComp = 0
            .RecordLocked = 0
            .DateChemLab = ""
            .WhoChemLab = ""
            .RerunStatus = 0
            .DateRerun = ""
            .MetLabComment = ""
            .ChemLabComment = ""
            .DateMetLab = ""
            .WhoMetLab = ""
            .SplitDepthTop = 0
            .SplitDepthBot = 0
            .SplitThck = 0
            .WashDate = ""
            .WhoModifiedSplit = ""
            .WhenModifiedSplit = ""
            .OrigData = 0
            '-----
            .MtxPctMoist = 0
            .FdPctMoist = 0
            .FlotFdBplActual = 0
            .FlotFdWtActual = 0
            .FlotFdBplCalc = 0
            .FlotFdWtCalc = 0
            .FlotPctRcvry = 0
            .FdBplDiff = 0
            .PbWtPct = 0
            .FdWtPct = 0
            .ClWtPct = 0
            .FdWtDiff = 0
            '-----
            .County = ""
            .BankCode = ""
            '-----
            .DryDensityOverride = 0
            '-----
            .QaQcHole = 0
            .MtxPctMoist1 = 0
            .MtxPctMoist2 = 0
        End With
    End Sub

    Public Function gGetProspLoctnData(ByVal aSampleId As String,
                                       ByRef aProspRawLoctnData As gRawProspLoctnType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspLoctnDataError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ProspRawDynaset As OraDynaset
        Dim RecordCount As Integer
        params = gDBParams

        params.Add("pSampleId", aSampleId, ORAPARM_INPUT)
        params("pSampleId").serverType = ORATYPE_VARCHAR2

        params.Add("pTownship", 0, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", 0, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", 0, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", "", ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", #12/31/8888#, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pSplitNumber", 0, ORAPARM_INPUT)
        params("pSplitNumber").serverType = ORATYPE_NUMBER

        params.Add("pSampleIdOnly", 1, ORAPARM_INPUT)
        params("pSampleIdOnly").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_raw_split
        'pSampleId           IN     VARCHAR2
        'pTownship           IN     NUMBER,
        'pRange              IN     NUMBER,
        'pSection            IN     NUMBER,
        'pHoleLocation       IN     VARCHAR2,
        'pProspDate          IN     DATE,
        'pSplitNumber        IN     NUMBER,
        'pSampleIdOnly       IN     NUMBER,
        'pResult             IN OUT c_prosprawsplit)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_split(" &
                  ":pSampleId, :pTownship, :pRange, :pSection, :pHoleLocation, " &
                  ":pProspDate, :pSplitNumber, :pSampleIdOnly, :pResult);end;", ORASQL_FAILEXEC)

        ProspRawDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = ProspRawDynaset.RecordCount

        'Should be only one record returned!
        ProspRawDynaset.MoveFirst()

        With aProspRawLoctnData
            .SampleId = ""
            .Section = 0
            .Township = 0
            .Range = 0
            .HoleLocation = ""
            .ProspDate = #12/31/8888#
            .SplitNumber = 0
        End With

        If RecordCount = 1 Then
            With aProspRawLoctnData
                .SampleId = ProspRawDynaset.Fields("sample_id").Value
                .Section = ProspRawDynaset.Fields("section").Value
                .Township = ProspRawDynaset.Fields("township").Value
                .Range = ProspRawDynaset.Fields("range").Value
                .HoleLocation = ProspRawDynaset.Fields("hole_location").Value
                .ProspDate = ProspRawDynaset.Fields("prosp_date").Value
                .SplitNumber = ProspRawDynaset.Fields("split_number").Value
            End With
        Else
            MsgBox("Error -- multiple sample#.",
                    vbOKOnly + vbExclamation,
                    "Too Many Sample#'s Error")
        End If

        ProspRawDynaset.Close()

        gGetProspLoctnData = True

        Exit Function

gGetProspLoctnDataError:
        gGetProspLoctnData = False

        MsgBox("Error accessing raw prospect data." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Raw Prospect Data Access Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        ProspRawDynaset.Close()
    End Function

    Public Function gGetProspRawSfcData(ByVal aSampleId As String,
                                        ByVal aTwp As Integer,
                                        ByVal aRge As Integer,
                                        ByVal aSec As Integer,
                                        ByVal aHloc As String,
                                        ByVal aProspDate As Date,
                                        ByVal aSplitNum As Integer,
                                        ByRef aProspRawSfcData() As gRawProspSfcType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspRawSfcDataError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ProspRawSfcDynaset As OraDynaset
        Dim RecordCount As Integer
        Dim SfcIdx As Integer

        params = gDBParams

        params.Add("pSampleId", aSampleId, ORAPARM_INPUT)
        params("pSampleId").serverType = ORATYPE_VARCHAR2

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHloc, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pSplitNumber", aSplitNum, ORAPARM_INPUT)
        params("pSplitNumber").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_raw_size_frctn
        'pSampleId           IN     VARCHAR2,
        'pTownship           IN     NUMBER,
        'pRange              IN     NUMBER,
        'pSection            IN     NUMBER,
        'pHoleLocation       IN     VARCHAR2,
        'pProspDate          IN     DATE,
        'pSplitNumber        IN     NUMBER,
        'pResult             IN OUT c_prosprawsplit)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_size_frctn(" &
                  ":pSampleId, :pTownship, :pRange, :pSection, :pHoleLocation, " &
                  ":pProspDate, :pSplitNumber, :pResult);end;", ORASQL_FAILEXEC)

        ProspRawSfcDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = ProspRawSfcDynaset.RecordCount
        SfcIdx = 0

        If RecordCount <> 0 Then
            ReDim aProspRawSfcData(RecordCount)

            ProspRawSfcDynaset.MoveFirst()
            Do While Not ProspRawSfcDynaset.EOF
                SfcIdx = SfcIdx + 1

                With aProspRawSfcData(SfcIdx)
                    .SampleId = ProspRawSfcDynaset.Fields("sample_id").Value
                    .Township = ProspRawSfcDynaset.Fields("township").Value
                    .Range = ProspRawSfcDynaset.Fields("range").Value
                    .Section = ProspRawSfcDynaset.Fields("section").Value
                    .HoleLocation = ProspRawSfcDynaset.Fields("hole_location").Value
                    .ProspDate = ProspRawSfcDynaset.Fields("prosp_date").Value
                    .SplitNumber = ProspRawSfcDynaset.Fields("split_number").Value
                    .SizeFrctnCode = ProspRawSfcDynaset.Fields("size_frctn_code").Value
                    .SfcDescription = ProspRawSfcDynaset.Fields("description").Value
                    .SfcOrderNum = ProspRawSfcDynaset.Fields("sfc_order_num").Value
                    .SfcMatlName = ProspRawSfcDynaset.Fields("matl_name").Value
                    .SfcMatlAbbrv = ProspRawSfcDynaset.Fields("matl_abbrv").Value
                    .Bpl = ProspRawSfcDynaset.Fields("bpl").Value
                    .FeAl = ProspRawSfcDynaset.Fields("feal").Value
                    .Insol = ProspRawSfcDynaset.Fields("insol").Value
                    .CaO = ProspRawSfcDynaset.Fields("cao").Value
                    .MgO = ProspRawSfcDynaset.Fields("mgo").Value
                    .Fe2O3 = ProspRawSfcDynaset.Fields("fe2o3").Value
                    .Al2O3 = ProspRawSfcDynaset.Fields("al2o3").Value
                    .Cd = ProspRawSfcDynaset.Fields("cd").Value
                    .SizeFrctnWt = ProspRawSfcDynaset.Fields("size_frctn_wt").Value
                    .SizeFrctnWtMsr = ProspRawSfcDynaset.Fields("size_frctn_wt_msr").Value
                    .SizeFrctnType = ProspRawSfcDynaset.Fields("size_frctn_type").Value

                    If Not IsDBNull(ProspRawSfcDynaset.Fields("who_modified").Value) Then
                        .WhoModified = ProspRawSfcDynaset.Fields("who_modified").Value
                    Else
                        .WhoModified = ""
                    End If

                    If Not IsDBNull(ProspRawSfcDynaset.Fields("when_modified").Value) Then
                        .WhenModified = Format(ProspRawSfcDynaset.Fields("when_modified").Value, "MM/dd/yyyy")
                    Else
                        .WhenModified = ""
                    End If

                    .OrderNum = ProspRawSfcDynaset.Fields("prsfc_order_num").Value
                End With
                ProspRawSfcDynaset.MoveNext()
            Loop
            gGetProspRawSfcData = True
        Else
            'There is no size fraction code data for this sample/split!
            gGetProspRawSfcData = False
        End If

        ProspRawSfcDynaset.Close()

        Exit Function

gGetProspRawSfcDataError:
        gGetProspRawSfcData = False

        MsgBox("Error accessing raw prospect sfc data." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Raw Prospect SFC Data Access Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        ProspRawSfcDynaset.Close()
    End Function

    '    Public Function gDispIntervalsNew(ByVal aSection As Integer, _
    '                                      ByVal aTownship As Integer, _
    '                                      ByVal aRange As Integer, _
    '                                      ByVal aHoleLocation As String, _
    '                                      ByVal aSplitNum As Integer, _
    '                                      ByVal aProspDate As Date, _
    '                                      ByRef aDispSpread As vaSpread) As Boolean

    '        '**********************************************************************
    '        '
    '        '
    '        '
    '        '**********************************************************************

    '        On Error GoTo DisplayIntervalsError

    '        Dim SampleDynaset As OraDynaset
    '        Dim SplitThk As Single
    '        Dim MetComment As String
    '        Dim ChemComment As String
    '        Dim DisplayedSplit As Integer
    '        Dim RecordCount As Integer
    '        Dim SampsOk As Boolean
    '        Dim ThisSplit As Integer

    '        DisplayedSplit = aSplitNum

    '        'Intervals will be displayed in aDispSpread.
    '        aDispSpread.MaxRows = 0

    '        SampsOk = gGetDrillHoleSamplesNew(aSection, _
    '                                          aTownship, _
    '                                          aRange, _
    '                                          aHoleLocation, _
    '                                          aProspDate, _
    '                                          SampleDynaset)

    '        If SampsOk = False Then
    '            gDispIntervalsNew = False
    '            Exit Function
    '        End If

    '        RecordCount = SampleDynaset.RecordCount

    '        If RecordCount = 0 Then
    '            gDispIntervalsNew = False
    '            Exit Function
    '        Else
    '            gDispIntervalsNew = True
    '        End If

    '        SampleDynaset.MoveFirst()

    '        Do While Not SampleDynaset.EOF
    '            With aDispSpread
    '                .MaxRows = .MaxRows + 1
    '                .Row = .MaxRows

    '                ThisSplit = SampleDynaset.Fields("split_number").Value
    '                .Col = 0
    '                .Text = "Spl" & CStr(ThisSplit)

    '                'Col1   From
    '                .Col = 1
    '                .Value = SampleDynaset.Fields("split_depth_top").Value

    '                'Col1   To
    '                .Col = 2
    '                .Value = SampleDynaset.Fields("split_depth_bot").Value

    '                'Col3   Thickness
    '                SplitThk = SampleDynaset.Fields("split_depth_bot").Value - _
    '                           SampleDynaset.Fields("split_depth_top").Value
    '                .Col = 3
    '                .Value = SplitThk

    '                'Col4   Sample#
    '                .Col = 4
    '                .Text = SampleDynaset.Fields("sample_id").Value

    '                'Col5   Drill date
    '                .Col = 5
    '                .Text = Format(SampleDynaset.Fields("prosp_date").Value, "mm/dd/yyyy")

    '                'Col6   Comments
    '                If Not isdbnull(SampleDynaset.Fields("metlab_comment").Value) Then
    '                    MetComment = SampleDynaset.Fields("metlab_comment").Value
    '                Else
    '                    MetComment = ""
    '                End If
    '                If Not isdbnull(SampleDynaset.Fields("chemlab_comment").Value) Then
    '                    ChemComment = SampleDynaset.Fields("chemlab_comment").Value
    '                Else
    '                    ChemComment = ""
    '                End If

    '                .Col = 6
    '                .Text = Trim(MetComment) + vbCrLf + Trim(ChemComment)

    '                If ThisSplit = DisplayedSplit Then
    '                    .BlockMode = True
    '                    .Row = .MaxRows
    '                    .Row2 = .MaxRows
    '                    .Col = 1
    '                    .Col2 = .MaxCols
    '                    .BackColor = &HC0FFC0   'Light green
    '                    .BlockMode = False
    '                End If

    '                SampleDynaset.MoveNext()
    '            End With
    '        Loop

    '        With aDispSpread
    '            .BlockMode = True
    '            .Row = 1
    '            .Row2 = .MaxRows
    '            .Col = 0
    '            .Col2 = 0
    '            .TypeTextWordWrap = False
    '            .TypeHAlign = SS_CELL_H_ALIGN_LEFT
    '            .BlockMode = False
    '        End With

    '        SampleDynaset.Close()

    '        Exit Function

    'DisplayIntervalsError:
    '        MsgBox("Error getting all sample#'s for this hole." & vbCrLf & _
    '            Err.Description, _
    '            vbOKOnly + vbExclamation, _
    '            "All Hole Sample#'s Access Error")

    '        On Error Resume Next
    '        gDispIntervalsNew = False
    '        SampleDynaset.Close()
    '    End Function

    Public Function gGetDrillHoleSamplesNew(ByVal aSection As Integer,
                                            ByVal aTownship As Integer,
                                            ByVal aRange As Integer,
                                            ByVal aHoleLocation As String,
                                            ByVal aProspDate As Date,
                                            ByRef aSampleDynaset As OraDynaset) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetDrillHoleSamplesNewError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer

        Dim ProspDateStr As String

        If aProspDate <> #12/31/8888# Then
            ProspDateStr = Format(aProspDate, "MM/dd/yyyy")
        Else
            ProspDateStr = "All"
        End If

        gGetDrillHoleSamplesNew = False

        params = gDBParams

        params.Add("pSection", aSection, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTownship, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRange, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHoleLocation, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        'Note:  In this case pProspDate is a VARCHAR2 not a DATE!
        params.Add("pProspDate", ProspDateStr, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_all_loc_prosprawbase
        'pSection             IN     NUMBER,
        'pTownship            IN     NUMBER,
        'pRange               IN     NUMBER,
        'pHoleLocation        IN     VARCHAR2,
        'pProspDate           IN     VARCHAR2,
        'pResult              IN OUT c_prosprawbase)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_all_loc_prosprawbase(:pSection," +
                  ":pTownship, :pRange, :pHoleLocation, :pProspDate, " +
                  ":pResult);end;", ORASQL_FAILEXEC)
        aSampleDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = aSampleDynaset.RecordCount
        gGetDrillHoleSamplesNew = True

        Exit Function

gGetDrillHoleSamplesNewError:
        MsgBox("Error getting all sample#'s for this prospect hole." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "All Hole Sample#'s Access Error")

        On Error Resume Next
        ClearParams(params)
        gGetDrillHoleSamplesNew = False
    End Function

    Public Function gGetDrillHoleNew(ByVal aSection As Integer,
                                     ByVal aTownship As Integer,
                                     ByVal aRange As Integer,
                                     ByVal aHoleLocation As String,
                                     ByVal aProspDate As Date,
                                     ByRef aSampleDynaset As OraDynaset) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetDrillHoleNewError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer

        gGetDrillHoleNew = False

        params = gDBParams

        params.Add("pTownship", aTownship, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRange, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSection, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHoleLocation, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_raw_base
        'pTownship           IN     NUMBER,
        'pRange              IN     NUMBER,
        'pSection            IN     NUMBER,
        'pHoleLocation       IN     VARCHAR2,
        'pProspDate          IN     DATE,
        'pResult             IN OUT c_prosprawbase)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_base(:pTownship," +
                  ":pRange, :pSection, :pHoleLocation, :pProspDate, " +
                  ":pResult);end;", ORASQL_FAILEXEC)
        aSampleDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = aSampleDynaset.RecordCount
        gGetDrillHoleNew = True

        Exit Function

gGetDrillHoleNewError:
        MsgBox("Error getting all raw prospect hole." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Process Error")

        On Error Resume Next
        ClearParams(params)
        gGetDrillHoleNew = False
    End Function

    Public Sub gGetProspHoleSplits(ByVal aSection As Integer,
                                   ByVal aTownship As Integer,
                                   ByVal aRange As Integer,
                                   ByVal aHoleLocation As String,
                                   ByVal aProspDate As Date,
                                   ByRef aHoleSplitDynaset As OraDynaset)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspHoleSplitsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer

        params = gDBParams

        params.Add("pTownship", aTownship, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRange, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSection, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHoleLocation, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_hole_sample_id
        'pTownship                  IN     NUMBER,
        'pRange                     IN     NUMBER,
        'pSection                   IN     NUMBER,
        'pHoleLocation              IN     VARCHAR2,
        'pProspDate                 IN     DATE,
        'pResult                    IN OUT c_prosprawbase)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_hole_sample_id(" &
                  ":pTownship, :pRange, :pSection, :pHoleLocation, " &
                  ":pProspDate, :pResult);end;", ORASQL_FAILEXEC)

        aHoleSplitDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = aHoleSplitDynaset.RecordCount

        Exit Sub

gGetProspHoleSplitsError:
        MsgBox("Error accessing hole splits." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Hole Splits Access Error")

        On Error Resume Next
        ClearParams(params)
    End Sub

    Public Function gSampNumExistsNew(ByVal aSampleId As String,
                                      ByVal aSec As Integer,
                                      ByVal aTwp As Integer,
                                      ByVal aRge As Integer,
                                      ByVal aHloc As String,
                                      ByVal aProspDate As Date) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gSampNumExistsNewError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer

        gSampNumExistsNew = -1

        'Does this sample number exist?
        params = gDBParams

        params.Add("pSampleId", aSampleId, ORAPARM_INPUT)
        params("pSampleId").serverType = ORATYPE_VARCHAR2

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHloc, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pResult", "", ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        'PROCEDURE sample_num_exists
        'pSampleId              IN     VARCHAR2,
        'pSection               IN     NUMBER,
        'pTownship              IN     NUMBER,
        'pRange                 IN     NUMBER,
        'pHoleLocation          IN     VARCHAR2,
        'pProspDate             IN     DATE,
        'pResult                IN OUT NUMBER);
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.sample_num_exists(:pSampleId, " +
                  ":pSection, :pTownship, :pRange, :pHoleLocation, :pProspDate, :pResult);end;", ORASQL_FAILEXEC)

        gSampNumExistsNew = params("pResult").Value

        ClearParams(params)

        Exit Function

gSampNumExistsNewError:
        MsgBox("Error checking if sample# exists." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Check if Sample# Exists Error")

        On Error Resume Next
        gSampNumExistsNew = -1
        ClearParams(params)
    End Function

    Public Function gSampNumExistsIdOnly(ByVal aSampleId As String,
                                         ByRef aSplitBaseData As gRawProspLoctnType) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gSampNumExistsIdOnlyError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer
        Dim ProspRawDynaset As OraDynaset

        gSampNumExistsIdOnly = -1

        'Does this sample number exist?
        params = gDBParams

        params.Add("pSampleId", aSampleId, ORAPARM_INPUT)
        params("pSampleId").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", "", ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE sample_num_exists_idonly2
        'pSampleId              IN     VARCHAR2,
        'pResult                IN OUT c_prosprawbase);
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.sample_num_exists_idonly2(:pSampleId, " +
                  ":pResult);end;", ORASQL_FAILEXEC)

        ProspRawDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = ProspRawDynaset.RecordCount
        gSampNumExistsIdOnly = RecordCount

        If RecordCount = 1 Then
            With aSplitBaseData
                .SampleId = ProspRawDynaset.Fields("sample_id").Value
                .Township = ProspRawDynaset.Fields("township").Value
                .Range = ProspRawDynaset.Fields("range").Value
                .Section = ProspRawDynaset.Fields("section").Value
                .HoleLocation = ProspRawDynaset.Fields("hole_location").Value
                .ProspDate = ProspRawDynaset.Fields("prosp_date").Value
                .SplitNumber = ProspRawDynaset.Fields("split_number").Value
            End With
        Else
            With aSplitBaseData
                .SampleId = ""
                .Township = 0
                .Range = 0
                .Section = 0
                .HoleLocation = ""
                .ProspDate = ""
                .SplitNumber = 0
            End With
        End If

        Exit Function

gSampNumExistsIdOnlyError:
        MsgBox("Error checking if sample# exists." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Check if Sample# Exists Error")

        On Error Resume Next
        gSampNumExistsIdOnly = -1
        ClearParams(params)
    End Function

    Public Function gGetSfcMatlAbbrv(ByVal aMatl As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        gGetSfcMatlAbbrv = ""

        Select Case aMatl
            Case Is = "Pebble"
                gGetSfcMatlAbbrv = "Pb"
            Case Is = "Concentrate"
                gGetSfcMatlAbbrv = "Cn"
            Case Is = "Feed"
                gGetSfcMatlAbbrv = "Fd"
            Case Is = "Pan"
                gGetSfcMatlAbbrv = "Pan"
            Case Is = "Tails"
                gGetSfcMatlAbbrv = "Tl"
            Case Is = "Clay"
                gGetSfcMatlAbbrv = "Cl"
            Case Is = "Head"
                gGetSfcMatlAbbrv = "Hd"
        End Select
    End Function

    Public Sub gGetMetLabErrorsNew(ByRef aRawProspCalcData As gRawProspCalcType,
                                   ByRef aRawProspBase As gRawProspBaseType,
                                   ByRef aSfcDataSprd() As gRawProspSfcSprdType,
                                   ByRef aErrComms() As String,
                                   ByVal aErrType As String,
                                   ByVal aSkipBarrenSplits As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ErrIdx As Integer
        Dim ErrCount As Integer
        Dim RowIdx As Integer
        Dim MtxPctSolMod As gMtxPctSolModType
        Dim ZzzRatioData As gZzzRatioDataType
        Dim MtxDensityMod As gMtxDensityModType
        Dim RatioOk As Boolean
        Dim ChemAnalOk As Boolean
        Dim SfcIdx As Integer
        Dim MatlIdx As Integer
        Dim TargMatl As String
        Dim FdOnlyNeeded As Boolean
        Dim FlotDryWtGms As Single

        Dim IntervalStat As Boolean
        Dim MissingSplit As Boolean
        Dim DoubledSplit As Boolean
        Dim ZeroSplit As Boolean
        Dim TotalSplit As Boolean
        Dim AllSplits() As gHoleIntervalType
        Dim DateProbs As Boolean
        Dim CurrDate As Date
        Dim PrevDate As Date
        Dim IntvProbs As Boolean
        Dim Top As Single
        Dim Bottom As Single
        Dim PrevSplit As Integer
        Dim CurrSplit As Integer
        Dim LowRcvry1 As Boolean
        Dim PctDiffWet As Single
        Dim PctDiffDry As Single
        Dim PctMoist1 As Single
        Dim PctMoist2 As Single
        Dim SpecFdCase As Boolean

        'Errors to check for -- from the IMC stuff from Glen
        ' 1)  Low flotation %recovery < 65%
        ' 2)  Concentrate Insol > 6.0%
        ' 3)  Feed grade difference > 2 BPL
        ' 4)  Chopping density outside 15% range (105.1 to 77.7)
        ' 5)  Core %solids outside range %solids (65% to 95%)
        ' 6)  Negative weight on clay.
        ' 7)  ZZZ ratio assay check
        ' 8)  Flotation feed > 1075 grams
        ' 9)  Feed %solids outside 75-85% solids  10/16/2009 -- changed to 75-86%
        '10)  Core %solids < 6% of model (value)
        '----------
        '11)  Waste clay% out of range
        '12)  High flotation %recovery > 95%
        '13)  Incomplete chemical analyses -- Fd, Cn, Pb, Tl         (from Cargill)
        '14)  I&A check for Pb & Cn (if > 4.00 then a problem)       (from Cargill)
        '15)  MgO check for Pb & Cn (if > 4.00 then a problem)       (from Cargill)
        '16)  CaO check for Pb & Cn (if > 59 Or < 10 then a problem) (from Cargill)
        '17)  Feed flotation weight problems                         (from Cargill)
        '18)  Pb wt + Fd wt > Mtx wt                                 (from Cargill)
        '----------
        '19)  For Ona or Wingate --> Clay settle lvl should not be zero.
        '     (aRawProspBase.ClaySettlingLvl)

        'aErrType will be "MetLab" or "ChemLab"

        '05/25/2007, lss
        '1)  Low flotation %recovery < 65% is returning too many errors per
        'Earnest Terry -- will change it to:
        '1)  Low flotation %recovery < 75%  (and feed grade > 8).
        '2)  Low flotation %recovery < 40%  (any feed grade).

        '----------
        '02/22/2010, lss
        'Now have 2 sets of matrix %moisture samples.  Need to have some error
        'checking on this.

        ErrCount = 0
        LowRcvry1 = False

        'Make sure that ErrComms is empty
        'For RowIdx = 1 To UBound(aErrComms, 1)
        '    aErrComms(RowIdx, 1) = ""
        '    aErrComms(RowIdx, 2) = ""
        '    aErrComms(RowIdx, 3) = ""
        '    aErrComms(RowIdx, 4) = ""   'Will be ChemLab, MetLab, Both
        'Next RowIdx

        'May not want to show any errors for "barren" splits -- depending on the
        'value of aSkipBarrenSplits that is passed into this procedure!
        If aRawProspBase.Barren = 1 And aSkipBarrenSplits = 1 Then
            Exit Sub
        End If

        'With aRawProspCalcData
        '    '1a) Low flotation recovery
        '    If .FlotPctRcvry < 75 And .FdBpl > 8 Then
        '        ErrCount = ErrCount + 1
        '        aErrComms(ErrCount, 1) = "Low flotation recovery = " & _
        '                                 Format(.FlotPctRcvry / 100, "##0.0%") & _
        '                                 "  (< 75%," & " Feed BPL = " & _
        '                                 Format(.FdBpl, "##0.0") & _
        '                                 ")"
        '        aErrComms(ErrCount, 2) = "%Rcvry Problem"
        '        aErrComms(ErrCount, 3) = Format(.FlotPctRcvry / 100, "##0.0%")
        '        aErrComms(ErrCount, 4) = "Both"
        '        LowRcvry1 = True
        '    End If

        ''1b) Low flotation recovery
        'If .FlotPctRcvry < 40 And LowRcvry1 = False Then
        '    ErrCount = ErrCount + 1
        '    aErrComms(ErrCount, 1) = "Low flotation recovery = " & _
        '                             Format(.FlotPctRcvry / 100, "##0.0%") & _
        '                             "  (< 40%" & ")"
        '    aErrComms(ErrCount, 2) = "%Rcvry Problem"
        '    aErrComms(ErrCount, 3) = Format(.FlotPctRcvry / 100, "##0.0%")
        '    aErrComms(ErrCount, 4) = "Both"
        'End If

        ''2) Concentrate Insol > 6.0%
        'If .CnIns > 6 Then
        '    ErrCount = ErrCount + 1
        '    aErrComms(ErrCount, 1) = "Concentrate Insol = " & _
        '                             Format(.CnIns, "#0.0") & "  (> 6.0)"
        '    aErrComms(ErrCount, 2) = "Ins>6 Cn"
        '    aErrComms(ErrCount, 3) = Format(.CnIns, "#0.0")
        '    aErrComms(ErrCount, 4) = "Both"
        'End If

        ''3) Feed grade difference > 2 BPL
        'If Abs(.FdBpl - .FlotBpl) > 2 Then
        '    ErrCount = ErrCount + 1
        '    aErrComms(ErrCount, 1) = "Feed BPL difference = " & _
        '                             Format(Abs(.FdBpl - .FlotBpl), "#0.0") & _
        '                             "  (> 2.0)"
        '    aErrComms(ErrCount, 2) = "Fd BPL Diff>2"
        '    aErrComms(ErrCount, 3) = Format(Abs(.FdBpl - .FlotBpl), "#0.0")
        '    aErrComms(ErrCount, 4) = "Both"
        'End If

        ''4)  Matrix (Core) %solids outside range %solids (65% to 95%)
        'If .CorePctSol < 65 Or .CorePctSol > 95 Then
        '    ErrCount = ErrCount + 1
        '    aErrComms(ErrCount, 1) = "Mtx %solids = " & _
        '                             Format(.CorePctSol, "##0.0") & _
        '                             " outside range  (65 to 95)"
        '    aErrComms(ErrCount, 2) = "Mtx %Sol Range"
        '    aErrComms(ErrCount, 3) = Format(.CorePctSol, "##0.0")
        '    aErrComms(ErrCount, 4) = "Both"
        'End If

        ''5)  Negative weight on clay
        'If .ClPctWt < 0 Then
        '    ErrCount = ErrCount + 1
        '    aErrComms(ErrCount, 1) = "Negative weight on clay  (%Clay = " & _
        '                             Format(.ClPctWt, "###0.0") & _
        '                             ")"
        '    aErrComms(ErrCount, 2) = "WClay Problem"
        '    aErrComms(ErrCount, 3) = Format(.ClPctWt, "###0.0")
        '    aErrComms(ErrCount, 4) = "Both"
        'End If

        ''6)  Feed %solids outside range %solids (75% to 85%)
        ''    10/16/2009 -- changed to 75% to 86%
        'If .FdPctSol < 75 Or .FdPctSol > 86 Then
        '    ErrCount = ErrCount + 1
        '    aErrComms(ErrCount, 1) = "Feed %solids = " & _
        '                             Format(.FdPctSol, "##0.0") & _
        '                             " outside range  (75 to 86)"
        '    aErrComms(ErrCount, 2) = "Fd %Solids"
        '    aErrComms(ErrCount, 3) = Format(.FdPctSol, "##0.0")
        '    aErrComms(ErrCount, 4) = "Both"
        'End If

        ''7)  Flotation feed weight > 1075 grams
        'If .FdWtDryGms > 1075 Then
        '    ErrCount = ErrCount + 1
        '    aErrComms(ErrCount, 1) = "Flotation feed weight > 1075 grams  (" & _
        '                             Format(.FdWtDryGms, "###0") & ")"
        '    aErrComms(ErrCount, 2) = "FlotWt >1075"
        '    aErrComms(ErrCount, 3) = Format(.FdWtDryGms, "###0")
        '    aErrComms(ErrCount, 4) = "Both"
        'End If

        ''8)  Waste clay < 10%
        'If .ClPctWt < 10 Then
        '    ErrCount = ErrCount + 1
        '    aErrComms(ErrCount, 1) = "Waste clay% < 10  (" & _
        '                             Format(.ClPctWt, "##0.0") & ")"
        '    aErrComms(ErrCount, 2) = "WClay Problem"
        '    aErrComms(ErrCount, 3) = Format(.ClPctWt, "##0.0")
        '    aErrComms(ErrCount, 4) = "Both"
        'End If

        ''9)  Waste clay > 10%
        'If .ClPctWt > 85 Then
        '    ErrCount = ErrCount + 1
        '    aErrComms(ErrCount, 1) = "Waste clay% > 85  (" & _
        '                             Format(.ClPctWt, "##0.0") & ")"
        '    aErrComms(ErrCount, 2) = "WClay Problem"
        '    aErrComms(ErrCount, 3) = Format(.ClPctWt, "##0.0")
        '    aErrComms(ErrCount, 4) = "Both"
        'End If

        ''10) High flotation recovery
        'If .FlotPctRcvry > 95 Then
        '    ErrCount = ErrCount + 1
        '    aErrComms(ErrCount, 1) = "High flotation recovery = " & _
        '                             Format(.FlotPctRcvry / 100, "##0.0%") & _
        '                             "  (> 95%)"
        '    aErrComms(ErrCount, 2) = "%Rcvry Problem"
        '    aErrComms(ErrCount, 3) = Format(.FlotPctRcvry / 100, "##0.0%")
        '    aErrComms(ErrCount, 4) = "Both"
        'End If
        'End With

        'Special error checks  Special error checks  Special error checks
        'Special error checks  Special error checks  Special error checks
        'Special error checks  Special error checks  Special error checks

        '1)  ZZZ Ratio
        'For SfcIdx = 1 To UBound(aSfcDataSprd)
        '    RatioOk = gGetZzzRatioData(aSfcDataSprd(SfcIdx).SizeFrctnCode, _
        '                               ZzzRatioData, _
        '                               aSfcDataSprd())
        '    If RatioOk = True Then
        '        If ZzzRatioData.AssayOk = False Then
        '            ErrCount = ErrCount + 1
        '            aErrComms(ErrCount, 1) = "ZZZ ratio problem -- " & _
        '                                     aSfcDataSprd(SfcIdx).SizeFrctnCode & _
        '                                     " (" & _
        '                                     Format(ZzzRatioData.ZzzRatio, "##0.00") & _
        '                                     " not between 0.9 and 1.10)"
        '            aErrComms(ErrCount, 2) = "ZZZ Ratio"
        '            aErrComms(ErrCount, 3) = aSfcDataSprd(SfcIdx).SizeFrctnCode
        '            aErrComms(ErrCount, 4) = "Both"
        '        End If
        '    End If
        'Next SfcIdx

        ''2)  Matrix (Core) %solids < 6% of model
        'gGetMtxPctSolComp(aRawProspBase, _
        '                  aRawProspCalcData, _
        '                  MtxPctSolMod)

        'With MtxPctSolMod
        '    If .PctMoistProblem = True Then
        '        ErrCount = ErrCount + 1
        '        aErrComms(ErrCount, 1) = "Mtx %solids < 6% of model  (%Sol = " & _
        '                                 Format(.CorePctSol, "##0.0") & _
        '                                 "   Model %Sol = " & _
        '                                 Format(.PctSolidsModel, "##0.0") & _
        '                                 "   Lo limit = " & _
        '                                 Format(.PctSolidsLowerLimit, "##0.0") & _
        '                                 ")"
        '        aErrComms(ErrCount, 2) = "Mtx %Sol Model"
        '        aErrComms(ErrCount, 3) = Format(.CorePctSol, "##0.0") & _
        '                                 "   Mod " & _
        '                                 Format(.PctSolidsModel, "##0.0") & _
        '                                 "   Lo " & _
        '                                 Format(.PctSolidsLowerLimit, "##0.0") & _
        '                                 ")"
        '        aErrComms(ErrCount, 4) = "Both"
        '    End If
        'End With

        ''3)  Matrix (Chopping) density outside 15% range (105.1 to 77.7)
        'gGetMtxDensityComp(aRawProspBase, _
        '                   aRawProspCalcData, _
        '                   MtxDensityMod)

        'With MtxDensityMod
        '    If .DensityProblem = True Then
        '        ErrCount = ErrCount + 1
        '        aErrComms(ErrCount, 1) = "Matrix (chopping) density = " & _
        '                                 Format(.LabMsrdDryDensity, "##0.0") & _
        '                                 " outside 15% range  (" & _
        '                                 Format(.LowerLimit, "##0.0") & _
        '                                 " to " & _
        '                                 Format(.UpperLimit, "##0.0") & ")"
        '        aErrComms(ErrCount, 2) = "Mtx Density Range"
        '        aErrComms(ErrCount, 3) = Format(.LabMsrdDryDensity, "##0.0") & _
        '                                 " (" & _
        '                                 Format(.LowerLimit, "##0.0") & _
        '                                 " to " & _
        '                                 Format(.UpperLimit, "##0.0") & ")"
        '        aErrComms(ErrCount, 4) = "Both"
        '    End If
        'End With

        ''4)  Incomplete chem analysis -- Pb, Cn, Fd, Tl
        'For MatlIdx = 1 To 4
        '    Select Case MatlIdx
        '        Case Is = 1
        '            TargMatl = "Pb"
        '            FdOnlyNeeded = False
        '        Case Is = 2
        '            TargMatl = "Cn"
        '            FdOnlyNeeded = False
        '        Case Is = 3
        '            TargMatl = "Fd"
        '            FdOnlyNeeded = True
        '        Case Is = 4
        '            TargMatl = "Tl"
        '            FdOnlyNeeded = True
        '    End Select

        '    For SfcIdx = 1 To UBound(aSfcDataSprd)
        '        'Pb or Cn -- should have BPL, Insol, CaO, MgO, Fe2O3, Al2O3
        '        '            if the weight is not 0
        '        'Fd or Tl -- should have BPL if the weight is not 0

        '        '12/06/2012, lss Was this!
        '        'If aSfcDataSprd(SfcIdx).SfcMatlName = TargMatl Then
        '        '    If (TargMatl = "Fd" And (aRawProspBase.MineName = "Ona" Or _
        '        '        aRawProspBase.MineName = "Ona-Pioneer") And _
        '        '        aSfcDataSprd(SfcIdx).SizeFrctnCode = "051") Then
        '        '        SpecFdCase = True
        '        '    Else
        '        '        SpecFdCase = False
        '        '    End If

        '        If aSfcDataSprd(SfcIdx).SfcMatlName = TargMatl Then
        '            If TargMatl = "Fd" Then
        '                If ((aRawProspBase.MineName = "Ona" Or _
        '                    aRawProspBase.MineName = "Ona-Pioneer") And _
        '                    aSfcDataSprd(SfcIdx).SizeFrctnCode = "051") Then
        '                    SpecFdCase = True    'Need complete analysis for the feed.
        '                    FdOnlyNeeded = False  '12/06/2012, lss  Added this line.
        '                Else
        '                    SpecFdCase = False
        '                    FdOnlyNeeded = True
        '                End If
        '            Else
        '                SpecFdCase = False
        '            End If

        '            ChemAnalOk = gGetChemAnalComplete(aSfcDataSprd(SfcIdx).SizeFrctnCode, _
        '                                              aSfcDataSprd(), _
        '                                              aSfcDataSprd(SfcIdx).SfcMatlName, _
        '                                              aRawProspBase.MineName, _
        '                                              SpecFdCase)
        '            If ChemAnalOk = False Then
        '                ErrCount = ErrCount + 1
        '                If FdOnlyNeeded = False Then
        '                    aErrComms(ErrCount, 1) = aSfcDataSprd(SfcIdx).SfcMatlName & _
        '                                             " " & aSfcDataSprd(SfcIdx).SizeFrctnCode & _
        '                                             " " & aSfcDataSprd(SfcIdx).SfcDescription & _
        '                                             " -- Incomplete chemical analysis"
        '                    aErrComms(ErrCount, 2) = "Incomplete Analysis"
        '                    aErrComms(ErrCount, 3) = aSfcDataSprd(SfcIdx).SizeFrctnCode
        '                    aErrComms(ErrCount, 4) = "Both"
        '                Else
        '                    If SpecFdCase = False Then
        '                        aErrComms(ErrCount, 1) = aSfcDataSprd(SfcIdx).SfcMatlName & _
        '                                                 " " & aSfcDataSprd(SfcIdx).SizeFrctnCode & _
        '                                                 " (" & aSfcDataSprd(SfcIdx).SfcDescription & _
        '                                                 " -- " & _
        '                                                 TargMatl & " BPL is missing)"
        '                        aErrComms(ErrCount, 2) = TargMatl & " Anal"
        '                        aErrComms(ErrCount, 3) = "Incomplete"
        '                        aErrComms(ErrCount, 4) = "Both"
        '                    End If
        '                End If
        '            End If
        '        End If
        '    Next SfcIdx
        'Next MatlIdx

        ''5)  I&A check for Pb & Cn (if > 4.00 then a problem)
        'For SfcIdx = 1 To UBound(aSfcDataSprd)
        '    If aSfcDataSprd(SfcIdx).SfcMatlName = "Pb" Or _
        '        aSfcDataSprd(SfcIdx).SfcMatlName = "Cn" Then
        '        ChemAnalOk = gGetChemAnalOk(aSfcDataSprd(SfcIdx).SizeFrctnCode, _
        '                                    aSfcDataSprd(), _
        '                                    "I&A")
        '        If ChemAnalOk = False Then
        '            ErrCount = ErrCount + 1
        '            aErrComms(ErrCount, 1) = aSfcDataSprd(SfcIdx).SfcMatlName & _
        '                                     " " & aSfcDataSprd(SfcIdx).SizeFrctnCode & _
        '                                     " " & aSfcDataSprd(SfcIdx).SfcDescription & _
        '                                     " -- I&A > 4.00  (" & _
        '                                     Format(aSfcDataSprd(SfcIdx).FeAl, "#0.00") & ")"
        '            aErrComms(ErrCount, 2) = "I&A>4"
        '            aErrComms(ErrCount, 3) = aSfcDataSprd(SfcIdx).SizeFrctnCode
        '            aErrComms(ErrCount, 4) = "Both"
        '        End If
        '    End If
        'Next SfcIdx

        ''6)  MgO check for Pb & Cn (if > 4.00 then a problem)
        'For SfcIdx = 1 To UBound(aSfcDataSprd)
        '    If aSfcDataSprd(SfcIdx).SfcMatlName = "Pb" Or _
        '        aSfcDataSprd(SfcIdx).SfcMatlName = "Cn" Then
        '        ChemAnalOk = gGetChemAnalOk(aSfcDataSprd(SfcIdx).SizeFrctnCode, _
        '                                    aSfcDataSprd(), _
        '                                    "MgO")
        '        If ChemAnalOk = False Then
        '            ErrCount = ErrCount + 1
        '            aErrComms(ErrCount, 1) = aSfcDataSprd(SfcIdx).SfcMatlName & _
        '                                     " " & aSfcDataSprd(SfcIdx).SizeFrctnCode & _
        '                                     " " & aSfcDataSprd(SfcIdx).SfcDescription & _
        '                                     " -- MgO > 4.00  (" & _
        '                                     Format(aSfcDataSprd(SfcIdx).MgO, "#0.00") & ")"
        '            aErrComms(ErrCount, 2) = "MgO>4"
        '            aErrComms(ErrCount, 3) = aSfcDataSprd(SfcIdx).SizeFrctnCode
        '            aErrComms(ErrCount, 4) = "Both"
        '        End If
        '    End If
        'Next SfcIdx

        ''7)  CaO check for Pb & Cn (if CaO > 59 Or CaO < 10 then a problem)
        'For SfcIdx = 1 To UBound(aSfcDataSprd)
        '    If aSfcDataSprd(SfcIdx).SfcMatlName = "Pb" Or _
        '        aSfcDataSprd(SfcIdx).SfcMatlName = "Cn" Then
        '        ChemAnalOk = gGetChemAnalOk(aSfcDataSprd(SfcIdx).SizeFrctnCode, _
        '                                    aSfcDataSprd(), _
        '                                    "CaO")
        '        If ChemAnalOk = False Then
        '            ErrCount = ErrCount + 1
        '            aErrComms(ErrCount, 1) = aSfcDataSprd(SfcIdx).SfcMatlName & _
        '                                     " " & aSfcDataSprd(SfcIdx).SizeFrctnCode & _
        '                                     " " & aSfcDataSprd(SfcIdx).SfcDescription & _
        '                                     " -- CaO > 59 Or CaO < 10  (" & _
        '                                     Format(aSfcDataSprd(SfcIdx).CaO, "#0.00") & ")"
        '            aErrComms(ErrCount, 2) = "CaO Problem"
        '            aErrComms(ErrCount, 3) = aSfcDataSprd(SfcIdx).SizeFrctnCode
        '            aErrComms(ErrCount, 4) = "Both"
        '        End If
        '    End If
        'Next SfcIdx

        ''8)  Feed flotation loss check (> 150 grams is a problem)
        ''    or if sum of Cn & Tl dry gms > original dry flotation grams then
        ''    a problem.
        'With aRawProspCalcData
        '    'Compare the flotation feed sample weight (in dry grams) to the
        '    'sum of the Cn and Tl samples in PROSP_RAW_SIZE_FRCTN
        '    If aRawProspBase.OrigData = 1 Then
        '        FlotDryWtGms = Round(1250 * ((100 - .FdPctMoist) / 100), 0)
        '    Else
        '        FlotDryWtGms = Round(aRawProspBase.FlotFdWetWt * ((100 - .FdPctMoist) / 100), 0)
        '    End If

        '    If FlotDryWtGms - .FdWtDryGms >= 0 Then
        '        If FlotDryWtGms - .FdWtDryGms > 150 Then
        '            ErrCount = ErrCount + 1
        '            aErrComms(ErrCount, 1) = "Flotation feed loss > 150 grams  (" & _
        '                                     Format(FlotDryWtGms - .FdWtDryGms, "###0") & ")"
        '            aErrComms(ErrCount, 2) = "Loss> 150gm"
        '            aErrComms(ErrCount, 3) = Format(FlotDryWtGms - .FdWtDryGms, "###0")
        '            aErrComms(ErrCount, 4) = "Both"
        '        End If
        '    Else
        '        ErrCount = ErrCount + 1
        '        aErrComms(ErrCount, 1) = "Dry flotation feed total is too high (" & _
        '                                 Format(FlotDryWtGms, "###0") & " vs " & _
        '                                 Format(.FdWtDryGms, "###0") & ")"
        '        aErrComms(ErrCount, 2) = "DryTot High"
        '        aErrComms(ErrCount, 3) = "Problem"
        '        aErrComms(ErrCount, 4) = "Both"
        '    End If
        'End With

        ''9)  Pb wt + Fd wt > Mtx wt
        ''    Will work with "adjusted" values here (adjusted to total matrix pounds)
        'With aRawProspCalcData
        '    If .DryFdLbsAdj + Round(.PbWtDryGms / 453.6, 0) > .DryCoreLbsTot Then
        '        ErrCount = ErrCount + 1
        '        aErrComms(ErrCount, 1) = "Pebble + feed weight > dry matrix weight"
        '        aErrComms(ErrCount, 2) = "Pb+Fd> DryMtx"
        '        aErrComms(ErrCount, 3) = "Problem"
        '        aErrComms(ErrCount, 4) = "Both"
        '    End If
        'End With

        ''Extra data checks (from old system, requested by Earnest Terry)
        ''Extra data checks (from old system, requested by Earnest Terry)
        ''Extra data checks (from old system, requested by Earnest Terry)
        'With aRawProspBase
        '    'Check for interval footage problems
        '    IntervalStat = gGetIntervalsNew(.Section, .Township, _
        '                                    .Range, .HoleLocation, .SplitNumber, _
        '                                    .ProspDate, AllSplits())

        '    If IntervalStat = True Then
        '        'Check for drill date problems -- should be the same for all splits for a given
        '        'prospect hole
        '        DateProbs = False
        '        PrevDate = AllSplits(1).DrillDate
        '        For RowIdx = 2 To UBound(AllSplits)
        '            If Not IsDBNull(AllSplits(RowIdx).DrillDate) Then
        '                CurrDate = AllSplits(RowIdx).DrillDate
        '            Else
        '                CurrDate = #12/31/8888#
        '            End If
        '            If CurrDate <> PrevDate Then
        '                DateProbs = True
        '            End If
        '            PrevDate = CurrDate
        '        Next RowIdx

        '        If DateProbs = True Then
        '            ErrCount = ErrCount + 1
        '            aErrComms(ErrCount, 1) = "Drill date problems for this hole"
        '            aErrComms(ErrCount, 2) = "Spl Dates<>"
        '            aErrComms(ErrCount, 3) = "Problem"
        '            aErrComms(ErrCount, 4) = "MetLab"
        '        End If
        '        'heyhey
        '        'AllSplits()   Row 1   Top of seam depth
        '        '              Row 2   Bottom of seam depth
        '        '              Row 3   Sample#
        '        '              Row 4   Prospect date  (Drill date)
        '        '              Row 5   Split#

        '        IntvProbs = False
        '        Top = AllSplits(1).TosDepth
        '        Bottom = AllSplits(1).BosDepth

        '        If Bottom <= Top Then
        '            IntvProbs = True
        '        End If

        '        If IntvProbs = False Then
        '            For RowIdx = 2 To UBound(AllSplits)
        '                Top = AllSplits(RowIdx).TosDepth

        '                If Top <> Bottom Then
        '                    IntvProbs = True
        '                    Exit For
        '                End If

        '                Bottom = AllSplits(RowIdx).BosDepth

        '                If Bottom <= Top Then
        '                    IntvProbs = True
        '                    Exit For
        '                End If
        '            Next RowIdx
        '        End If

        '        If IntvProbs = True Then
        '            ErrCount = ErrCount + 1
        '            aErrComms(ErrCount, 1) = "Interval footage problems for this hole."
        '            aErrComms(ErrCount, 2) = "FtLngth"
        '            aErrComms(ErrCount, 3) = "Problem"
        '            aErrComms(ErrCount, 4) = "MetLab"
        '        End If

        '        'Check for missing splits (& doubled splits)
        '        MissingSplit = False
        '        DoubledSplit = False
        '        ZeroSplit = False
        '        TotalSplit = False

        '        For RowIdx = 1 To UBound(AllSplits)
        '            If AllSplits(RowIdx).Split = 0 Then
        '                ZeroSplit = True
        '            End If
        '        Next RowIdx

        '        PrevSplit = AllSplits(1).Split
        '        For RowIdx = 2 To UBound(AllSplits)
        '            CurrSplit = AllSplits(RowIdx).Split

        '            If CurrSplit <> PrevSplit + 1 Then
        '                MissingSplit = True
        '                Exit For
        '            End If

        '            PrevSplit = CurrSplit
        '        Next RowIdx

        '        PrevSplit = AllSplits(1).Split
        '        For RowIdx = 2 To UBound(AllSplits)
        '            CurrSplit = AllSplits(RowIdx).Split

        '            If CurrSplit = PrevSplit Then
        '                DoubledSplit = True
        '                Exit For
        '            End If

        '            PrevSplit = CurrSplit
        '        Next RowIdx

        '        If UBound(AllSplits) <> .SplitTotalNum Then
        '            TotalSplit = True
        '        End If

        '        If MissingSplit = True Then
        '            ErrCount = ErrCount + 1
        '            aErrComms(ErrCount, 1) = "Split is missing for this hole."
        '            aErrComms(ErrCount, 2) = "MsngSpl"
        '            aErrComms(ErrCount, 3) = "Problem"
        '            aErrComms(ErrCount, 4) = "MetLab"
        '        End If

        '        If DoubledSplit = True Then
        '            ErrCount = ErrCount + 1
        '            aErrComms(ErrCount, 1) = "Split is used 2X for this hole."
        '            aErrComms(ErrCount, 2) = "DblSpl"
        '            aErrComms(ErrCount, 3) = "Problem"
        '            aErrComms(ErrCount, 4) = "MetLab"
        '        End If

        '        If ZeroSplit = True Then
        '            ErrCount = ErrCount + 1
        '            aErrComms(ErrCount, 1) = "Split with zero value for this hole."
        '            aErrComms(ErrCount, 2) = "ZeroSpl"
        '            aErrComms(ErrCount, 3) = "Problem"
        '            aErrComms(ErrCount, 4) = "MetLab"
        '        End If

        '        If TotalSplit = True Then
        '            ErrCount = ErrCount + 1
        '            aErrComms(ErrCount, 1) = "Problem checking for interval footage problems."
        '            aErrComms(ErrCount, 2) = "IntrvFtge"
        '            aErrComms(ErrCount, 3) = "Problem"
        '            aErrComms(ErrCount, 4) = "MetLab"
        '        End If
        '    Else
        '        ErrCount = ErrCount + 1
        '        aErrComms(ErrCount, 1) = "Problem checking for interval footage problems."
        '        aErrComms(ErrCount, 2) = "IntrvFtge"
        '        aErrComms(ErrCount, 3) = "Problem"
        '        aErrComms(ErrCount, 4) = "MetLab"
        '    End If
        'End With

        ''Check matrix %moisture samples
        ''Check matrix %moisture samples
        ''Check matrix %moisture samples

        ''03/15/2010, lss
        ''Don't really need to check the moisture sample weights!  I misunderstood what
        ''Earnest Terry was talking about or he didn't know what he was talking about!
        ''With aRawProspBase
        ''    If (.MtxMoistWetWt > 0 And .MtxMoistDryWt <= 0) Or _
        ''        (.MtxMoistDryWt > 0 And .MtxMoistWetWt <= 0) Or _
        ''        (.MtxMoistWetWt2 > 0 And .MtxMoistDryWt2 <= 0) Or _
        ''        (.MtxMoistDryWt2 > 0 And .MtxMoistWetWt2 <= 0) Then
        ''        ErrCount = ErrCount + 1
        ''        aErrComms(ErrCount, 1) = "Mtx %moist dry or wet weight is missing."
        ''        aErrComms(ErrCount, 2) = "Mtx%Moist Wts"
        ''        aErrComms(ErrCount, 3) = "Problem"
        ''        aErrComms(ErrCount, 4) = "MetLab"
        ''    End If

        ''    If (.MtxMoistWetWt2 > 0 Or .MtxMoistDryWt2 > 0) And _
        ''        (.MtxMoistDryWt <= 0 Or .MtxMoistWetWt <= 0) Then
        ''        ErrCount = ErrCount + 1
        ''        aErrComms(ErrCount, 1) = "Has mtx %moist 2nd set data but incomplete 1st set data."
        ''        aErrComms(ErrCount, 2) = "Mtx%Moist Wts"
        ''        aErrComms(ErrCount, 3) = "Problem"
        ''        aErrComms(ErrCount, 4) = "MetLab"
        ''    End If

        ''    If .MtxMoistWetWt2 > 0 Then
        ''        PctDiffWet = Round(Abs(.MtxMoistWetWt2 - .MtxMoistWetWt) / .MtxMoistWetWt * 100, 1)
        ''    Else
        ''        PctDiffWet = 0
        ''    End If

        ''    If .MtxMoistDryWt2 > 0 Then
        ''        PctDiffDry = Round(Abs(.MtxMoistDryWt2 - .MtxMoistDryWt) / .MtxMoistDryWt * 100, 1)
        ''    Else
        ''        PctDiffDry = 0
        ''    End If

        ''    If PctDiffWet > 10 Then
        ''        ErrCount = ErrCount + 1
        ''        aErrComms(ErrCount, 1) = "Mtx %moist wet weight set difference > 10%."
        ''        aErrComms(ErrCount, 2) = "Mtx%Moist Wts"
        ''        aErrComms(ErrCount, 3) = "Problem"
        ''        aErrComms(ErrCount, 4) = "MetLab"
        ''    End If

        ''    If PctDiffDry > 10 Then
        ''        ErrCount = ErrCount + 1
        ''        aErrComms(ErrCount, 1) = "Mtx %moist dry weight set difference > 10%."
        ''        aErrComms(ErrCount, 2) = "Mtx%Moist Wts"
        ''        aErrComms(ErrCount, 3) = "Problem"
        ''        aErrComms(ErrCount, 4) = "MetLab"
        ''    End If

        ''03/15/2010, lss
        ''Need to check the difference between the two %moisture calculations (if two of
        ''them exist).
        'With aRawProspBase
        '    If .MtxPctMoist1 > 0 And .MtxPctMoist2 > 0 Then
        '        'Have already checked that .MtxPctMoist1 > 0 so don't need to check again here.
        '        If Abs(Round((.MtxPctMoist2 - .MtxPctMoist1) / .MtxPctMoist1 * 100, 1)) > 10 Then
        '            ErrCount = ErrCount + 1
        '            aErrComms(ErrCount, 1) = "Mtx %moistures difference > 10%."
        '            aErrComms(ErrCount, 2) = "Mtx%Moists"
        '            aErrComms(ErrCount, 3) = "Problem"
        '            aErrComms(ErrCount, 4) = "MetLab"
        '        End If
        '    End If
        'End With

        ''01/13/2012, lss
        ''For Ona or Ona-Pioneer or Wingate --> Clay settle lvl should not be zero.
        'If aRawProspBase.MineName = "Ona" Or aRawProspBase.MineName = "Wingate" Or _
        '    aRawProspBase.MineName = "Ona-Pioneer" Then
        '    If aRawProspBase.ClaySettlingLvl <= 0 Then
        '        ErrCount = ErrCount + 1
        '        aErrComms(ErrCount, 1) = "Clay settling level <=0."
        '        aErrComms(ErrCount, 2) = "ClySettLvl"
        '        aErrComms(ErrCount, 3) = "Problem"
        '        aErrComms(ErrCount, 4) = "MetLab"
        '    End If
        'End If

        '01/13/2012, lss
        'For Ona or Ona-Pioneer --> Should have full analysis in the F1(+20M) feed sample.

        '12/06/2012, lss  This is checked in 4) above.  Don't need this code here.
        '                 Commented it out.
        'If aRawProspBase.MineName = "Ona" Or aRawProspBase.MineName = "Ona-Pioneer" Then
        '    TargMatl = "Fd"
        '    FdOnlyNeeded = False
        '    SpecFdCase = True
        '    For SfcIdx = 1 To UBound(aSfcDataSprd)
        '        If aSfcDataSprd(SfcIdx).SfcMatlName = TargMatl And _
        '            aSfcDataSprd(SfcIdx).SizeFrctnCode = "051" Then
        '            ChemAnalOk = gGetChemAnalComplete(aSfcDataSprd(SfcIdx).SizeFrctnCode, _
        '                                              aSfcDataSprd(), _
        '                                              aSfcDataSprd(SfcIdx).SfcMatlName, _
        '                                              aRawProspBase.MineName, _
        '                                              SpecFdCase)
        '            If ChemAnalOk = False Then
        '                ErrCount = ErrCount + 1
        '                aErrComms(ErrCount, 1) = aSfcDataSprd(SfcIdx).SfcMatlName & _
        '                                         " " & aSfcDataSprd(SfcIdx).SizeFrctnCode & _
        '                                         " " & aSfcDataSprd(SfcIdx).SfcDescription & _
        '                                         " -- Incomplete chemical analysis"
        '                aErrComms(ErrCount, 2) = "Incomplete Analysis"
        '                aErrComms(ErrCount, 3) = aSfcDataSprd(SfcIdx).SizeFrctnCode
        '                aErrComms(ErrCount, 4) = "MetLab"
        '            End If
        '        End If
        '    Next SfcIdx
        'End If
    End Sub

    Public Function gGetSfcDataSfc(ByRef aSfcData As gSfcDataType,
                                   ByVal aSfc As String,
                                   ByRef aSfcDataSprd() As gRawProspSfcSprdType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ThisSfc As String
        Dim SfcIdx As Integer

        'Get size fraction code data for a single sized fraction code.

        gZeroSfcData(aSfcData)

        aSfc = gPadLeftChar(aSfc, 3, "0")

        For SfcIdx = 1 To UBound(aSfcDataSprd)
            ThisSfc = aSfcDataSprd(SfcIdx).SizeFrctnCode
            If ThisSfc = aSfc Then
                If IsNumeric(aSfcDataSprd(SfcIdx).SizeFrctnWt) Then
                    aSfcData.Weight = aSfcDataSprd(SfcIdx).SizeFrctnWt
                End If

                If IsNumeric(aSfcDataSprd(SfcIdx).Bpl) Then
                    aSfcData.Bpl = aSfcDataSprd(SfcIdx).Bpl
                End If

                If IsNumeric(aSfcDataSprd(SfcIdx).Insol) Then
                    aSfcData.Insol = aSfcDataSprd(SfcIdx).Insol
                End If

                If IsNumeric(aSfcDataSprd(SfcIdx).CaO) Then
                    aSfcData.CaO = aSfcDataSprd(SfcIdx).CaO
                End If

                If IsNumeric(aSfcDataSprd(SfcIdx).MgO) Then
                    aSfcData.MgO = aSfcDataSprd(SfcIdx).MgO
                End If

                If IsNumeric(aSfcDataSprd(SfcIdx).Fe2O3) Then
                    aSfcData.Fe2O3 = aSfcDataSprd(SfcIdx).Fe2O3
                End If

                If IsNumeric(aSfcDataSprd(SfcIdx).Al2O3) Then
                    aSfcData.Al2O3 = aSfcDataSprd(SfcIdx).Al2O3
                End If

                If IsNumeric(aSfcDataSprd(SfcIdx).FeAl) Then
                    aSfcData.FeAl = aSfcDataSprd(SfcIdx).FeAl
                End If

                If IsNumeric(aSfcDataSprd(SfcIdx).Cd) Then
                    aSfcData.Cd = aSfcDataSprd(SfcIdx).Cd
                End If
            End If
        Next SfcIdx
    End Function

    Public Sub gZeroSfcData(ByRef aSfcData As gSfcDataType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        With aSfcData
            .Weight = 0
            .Bpl = 0
            .Insol = 0
            .CaO = 0
            .Fe2O3 = 0
            .Al2O3 = 0
            .FeAl = 0
            .MgO = 0
            .Cd = 0
            .Type = ""
        End With
    End Sub

    Public Function gGetZzzRatioData(ByVal aSfc As String,
                                     ByRef aZzzRatioData As gZzzRatioDataType,
                                     ByRef aSfcDataSprd() As gRawProspSfcSprdType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SfcData As gSfcDataType

        gGetSfcDataSfc(SfcData, aSfc, aSfcDataSprd)

        'Is this size fraction code OK to get a ZZZ ratio from?
        With SfcData
            If .Bpl <> 0 And .Insol <> 0 And .MgO <> 0 And .Fe2O3 <> 0 And
                .Al2O3 <> 0 Then
                gGetZzzRatioData = True
            Else
                gGetZzzRatioData = False
            End If
        End With

        If gGetZzzRatioData = True Then
            With aZzzRatioData
                .Bpl = SfcData.Bpl
                .Ins = SfcData.Insol
                .InsCalc = Round((1 - (SfcData.Bpl * 0.012844) -
                             (SfcData.MgO * 0.048303) -
                             (SfcData.Fe2O3 * 0.016062) +
                             (SfcData.Al2O3 * 0.015825)) / 0.012591, 4)
                .Mg = SfcData.MgO
                .Fe = SfcData.Fe2O3
                .Al = SfcData.Al2O3
                .ZzzRatio = Round((SfcData.Bpl * 0.012844) +
                            (SfcData.Insol * 0.012591) -
                            (SfcData.Al2O3 * 0.015825) +
                            (SfcData.Fe2O3 * 0.016062) +
                            (SfcData.MgO * 0.048303), 3)
                .ZlBpl = Round(100 * SfcData.Bpl / (100 - SfcData.Insol), 2)
                If .ZzzRatio >= 0.9 And .ZzzRatio <= 1.1 Then
                    .AssayOk = True
                Else
                    .AssayOk = False
                End If
            End With
        Else
            With aZzzRatioData
                .Bpl = 0
                .Ins = 0
                .InsCalc = 0
                .Mg = 0
                .Fe = 0
                .Al = 0
                .ZzzRatio = 0
                .ZlBpl = 0
                .AssayOk = False
            End With
        End If
    End Function

    Public Sub gGetMtxPctSolComp(ByRef aRawProspBase As gRawProspBaseType,
                                 ByRef aRawProspCalcData As gRawProspCalcType,
                                 ByRef aMtxPctSolMod As gMtxPctSolModType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim PctMoistCalc As Single
        Dim PctSolidsModel As Single

        Dim Part1 As Single
        Dim Part2 As Single
        Dim Part3 As Single
        Dim Part4 As Single

        With aMtxPctSolMod
            .SplitDepthBot = aRawProspBase.SplitDepthBot
            .Plus35FdBpl = aRawProspCalcData.Plus35FdBpl
            .Plus35FdWt = aRawProspCalcData.Plus35FdWt
            .Minus35FdBpl = aRawProspCalcData.Minus35FdBpl
            .Minus35FdWt = aRawProspCalcData.Minus35FdWt
            .DryFdLbsAdj = aRawProspCalcData.DryFdLbsAdj
            .PbWtDryGms = aRawProspCalcData.PbWtDryGms
            .MtxTotWetWt = aRawProspBase.MtxTotWetWt
            .CorePctSol = aRawProspCalcData.CorePctSol
            .PbPctWt = aRawProspCalcData.PbPctWt
            .FdPctWt = aRawProspCalcData.FdPctWt
            .ClPctWt = aRawProspCalcData.ClPctWt
            .DryCoreLbsTot = aRawProspCalcData.DryCoreLbsTot
            .DryCoreLbsProc = aRawProspCalcData.DryCoreLbsProc

            'Will calculate the overall feed BPL from the +35M and
            '-35M feed BPL's from above.
            If IIf(.Plus35FdBpl > 0, .Plus35FdWt, 0) +
                IIf(.Minus35FdBpl > 0, .Minus35FdWt, 0) <> 0 Then
                .FdBplCalc = Round((.Plus35FdBpl * .Plus35FdWt + .Minus35FdBpl * .Minus35FdWt) /
                             (IIf(.Plus35FdBpl > 0, .Plus35FdWt, 0) +
                             IIf(.Minus35FdBpl > 0, .Minus35FdWt, 0)), 2)
            Else
                .FdBplCalc = 0
            End If

            .PctSand = ((1 - .FdBplCalc / 70)) * aRawProspCalcData.FdPctWt

            Part1 = 0.1032 + (0.518 * (.PbPctWt / 100)) +
                    (0.603 * (.ClPctWt / 100)) +
                    (0.18 * (.FdPctWt / 100))

            Part2 = (0.298 * (.PbPctWt / 100) * (.PbPctWt / 100)) -
                    (0.392 * (.ClPctWt / 100) * (.ClPctWt / 100)) -
                    (0.17 * (.PctSand / 100) * (.PctSand / 100))

            Part3 = -(1.166 * (.PbPctWt / 100) * (.ClPctWt / 100)) -
                    (0.516 * (.ClPctWt / 100) * (.PctSand / 100))

            Part4 = -(1.069 * (.PbPctWt / 100) * (.PctSand / 100)) +
                    (0.000354 * .SplitDepthBot)


            .PctMoistCalc = Round((0.1032 + (0.518 * (.PbPctWt / 100)) +
                            (0.603 * (.ClPctWt / 100)) +
                            (0.18 * (.FdPctWt / 100)) -
                            (0.298 * (.PbPctWt / 100) * (.PbPctWt / 100)) -
                            (0.392 * (.ClPctWt / 100) * (.ClPctWt / 100)) -
                            (0.17 * (.PctSand / 100) * (.PctSand / 100)) -
                            (1.166 * (.PbPctWt / 100) * (.ClPctWt / 100)) -
                            (0.516 * (.ClPctWt / 100) * (.PctSand / 100)) -
                            (1.069 * (.PbPctWt / 100) * (.PctSand / 100)) +
                            (0.000354 * .SplitDepthBot)) * 100, 2)

            .PctSolidsModel = 100 - .PctMoistCalc

            'Lower limt = calculates model value - 6%
            'There is no upper limit
            .PctSolidsLowerLimit = Round(.PctSolidsModel - 6, 1)
            If .CorePctSol < .PctSolidsLowerLimit Then
                .PctMoistProblem = True
            Else
                .PctMoistProblem = False
            End If
        End With
    End Sub

    Public Sub gGetMtxDensityComp(ByRef aRawProspBase As gRawProspBaseType,
                                  ByRef aRawProspCalcData As gRawProspCalcType,
                                  ByRef aMtxDensityMod As gMtxDensityModType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim PctMoistCalc As Single
        Dim PctSolidsModel As Single

        'Note This subroutine also exists in modRawProspDataReduction
        '     with different parameters (called gGetMtxDensityComp2).

        With aMtxDensityMod
            .LabMsrdDryDensity = aRawProspBase.DryDensity
            .MtxPctSol = aRawProspCalcData.CorePctSol
            .MtxPctMoist = aRawProspCalcData.CorePctMoist
            .PbPctWt = aRawProspCalcData.PbPctWt
            .FdPctWt = aRawProspCalcData.FdPctWt
            .ClPctWt = aRawProspCalcData.ClPctWt
            .FdBpl = aRawProspCalcData.FdBpl
            .DenFac = (131 * (.ClPctWt / 100)) +
                      (181 * (.PbPctWt / 100)) +
                      ((.FdPctWt / 100) * ((0.183 * .FdBpl) + 165))

            If .MtxPctMoist <> 0 And .MtxPctSol <> 0 Then
                .CalDen1 = Round((0.826 * .DenFac) /
                           ((.MtxPctSol / 100) + (((.MtxPctMoist / 100) / (.MtxPctSol / 100)) _
                           * .DenFac / 62.4)), 4)
            Else
                .CalDen1 = 0
            End If

            .CalDen = Round(0.521341 + (1.063872 * .CalDen1), 4)
            .LowerLimit = Round(.CalDen * 0.85, 1)
            .UpperLimit = Round(.CalDen * 1.15, 1)

            '04/15/2008, lss
            'Was comparing .CalDen with .LowerLimit and .UpperLimit
            'Needed to compare .LabMsrdDryDensity with the two values!
            If .LabMsrdDryDensity > .LowerLimit And .LabMsrdDryDensity < .UpperLimit Then
                .DensityProblem = False
            Else
                .DensityProblem = True
            End If
        End With
    End Sub

    Public Function gGetChemAnalComplete(ByVal aSfc As String,
                                         ByRef aSfcDataSprd() As gRawProspSfcSprdType,
                                         ByVal aMatlName As String,
                                         ByVal aMineName As String,
                                         ByVal aSpecFdCase As Boolean) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SfcData As gSfcDataType

        gGetSfcDataSfc(SfcData, aSfc, aSfcDataSprd)

        If aMatlName = "Pb" Or aMatlName = "Cn" Or aSpecFdCase = True Then
            'Complete analysis = BPL, Insol, MgO, Fe2O3, Al2O3, CaO
            With SfcData
                If .Weight <> 0 Then
                    If .Bpl = 0 Or .Insol = 0 Or .MgO = 0 Or .Fe2O3 = 0 Or
                        .Al2O3 = 0 Or .CaO = 0 Then
                        gGetChemAnalComplete = False
                    Else
                        gGetChemAnalComplete = True
                    End If
                Else
                    'The weight is zero therefore analysis is not needed or
                    'available
                    gGetChemAnalComplete = True
                End If
            End With
        End If

        If aSpecFdCase = False Then
            If aMatlName = "Fd" Or aMatlName = "Tl" Then
                'Complete analysis = BPL
                With SfcData
                    If .Weight <> 0 Then
                        If .Bpl = 0 Then
                            gGetChemAnalComplete = False
                        Else
                            gGetChemAnalComplete = True
                        End If
                    Else
                        'The weight is zero therefore analysis is not needed or
                        'available
                        gGetChemAnalComplete = True
                    End If
                End With
            End If
        End If
    End Function

    Public Function gGetChemAnalOk(ByVal aSfc As String,
                                   ByRef aSfcDataSprd() As gRawProspSfcSprdType,
                                   ByVal aAnalyte As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SfcData As gSfcDataType

        gGetSfcDataSfc(SfcData, aSfc, aSfcDataSprd)

        'Only have checks for I&A and MgO
        With SfcData
            If .Weight <> 0 Then
                gGetChemAnalOk = True
                Select Case aAnalyte
                    Case Is = "I&A"
                        If .FeAl > 4 Then
                            gGetChemAnalOk = False
                        End If
                    Case Is = "MgO"
                        If .MgO > 4 Then
                            gGetChemAnalOk = False
                        End If
                    Case Is = "CaO"
                        If .CaO > 59 Or .CaO < 10 Then
                            gGetChemAnalOk = False
                        End If
                End Select
            Else
                'The weight is zero therefore analysis is not needed or
                'available
                gGetChemAnalOk = True
            End If
        End With
    End Function

    Public Sub gGetRawProspCalcData(ByRef aRawProspCalcData As gRawProspCalcType,
                                    ByRef aRawProspBase As gRawProspBaseType,
                                    ByRef aSfcDataSprd() As gRawProspSfcSprdType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim FlotCalcData As gFlotCalcDataType
        Dim SfcData As gSfcDataType
        Dim SfcPlusMinusData As gSfcPlusMinusDataType
        Dim FlotFdWetWt As Single

        If aRawProspBase.OrigData = 1 Then
            If aRawProspBase.FlotFdWetWt <> 0 Then
                FlotFdWetWt = aRawProspBase.FlotFdWetWt
            Else
                FlotFdWetWt = 1250
            End If
        Else
            FlotFdWetWt = aRawProspBase.FlotFdWetWt
        End If

        gCalcFlotationData(FlotCalcData,
                           aRawProspBase.FdMoistWetWt,
                           aRawProspBase.FdMoistDryWt,
                           aRawProspBase.FdMoistTareWt,
                           FlotFdWetWt,
                           aSfcDataSprd)

        With aRawProspCalcData
            .SplitThck = aRawProspBase.SplitThck
            '.CorePctMoist = gCalcMoist(aRawProspBase.MtxMoistWetWt, _
            '                           aRawProspBase.MtxMoistDryWt, _
            '                           aRawProspBase.MtxMoistTareWt)
            '.CorePctSol = 100 - .CorePctMoist

            '-----
            '02/16/2010, lss
            'Need to get matrix %moisture based on 2 sets of matrix moisture samples!
            'The second set of matrix moisture samples (wet, dry, tare) will all be
            'zeros for older holes -- gCalcMoist2 will handle this correctly!
            .CorePctMoist = gCalcMoist2(aRawProspBase.MtxMoistWetWt,
                                        aRawProspBase.MtxMoistDryWt,
                                        aRawProspBase.MtxMoistTareWt,
                                        aRawProspBase.MtxMoistWetWt2,
                                        aRawProspBase.MtxMoistDryWt2,
                                        aRawProspBase.MtxMoistTareWt2)
            .CorePctSol = 100 - .CorePctMoist
            '-----

            .FdPctMoist = gCalcMoist(aRawProspBase.FdMoistWetWt,
                                     aRawProspBase.FdMoistDryWt,
                                     aRawProspBase.FdMoistTareWt)
            .FdPctSol = 100 - .FdPctMoist

            If aRawProspBase.SplitThck <> 0 Then
                .CoreLbsPerFt = Round(aRawProspBase.MtxTotWetWt / aRawProspBase.SplitThck, 1)
            Else
                .CoreLbsPerFt = 0
            End If

            .DryDensity = aRawProspBase.DryDensity
            .XmitLbs = aRawProspBase.MtxTotWetWt
            .FdBpl = FlotCalcData.FdBplAct
            .FlotBpl = FlotCalcData.FdBplCalc
            .FlotActWt = FlotCalcData.FdWtAct
            .FlotCalcWt = FlotCalcData.FdWtCalc

            .Plus35FdPct = gGetSfcPctPlusData(SfcPlusMinusData,
                                              "Fd",
                                              "+35",
                                              aSfcDataSprd)
            .Plus35FdBpl = SfcPlusMinusData.PlusBpl
            .Plus35FdWt = SfcPlusMinusData.PlusWt

            .Minus35FdPct = 100 - .Plus35FdPct
            .Minus35FdBpl = SfcPlusMinusData.MinusBpl
            .Minus35FdWt = SfcPlusMinusData.MinusWt

            gGetSfcDataMatl(SfcData,
                            "Pb",
                            aSfcDataSprd)
            .PbBpl = SfcData.Bpl
            .PbWtDryGms = SfcData.Weight
            .PbMgO = SfcData.MgO                   'Added 02/06/2008, lss
            .PbIa = SfcData.Fe2O3 + SfcData.Al2O3  'Added 02/06/2008, lss

            .FdBpl = FlotCalcData.FdBplAct
            .FdWtDryGms = FlotCalcData.FdWtAct

            gGetSfcDataMatl(SfcData,
                            "Cl",
                            aSfcDataSprd)
            .ClBpl = SfcData.Bpl

            .CnBpl = FlotCalcData.CnBplAct

            .GmtBpl = FlotCalcData.TlBplAct

            gGetSfcDataMatl(SfcData,
                            "Cn",
                            aSfcDataSprd)
            .CnIns = SfcData.Insol
            .CnMgO = SfcData.MgO                   'Added 02/06/2008, lss
            .CnIa = SfcData.Fe2O3 + SfcData.Al2O3  'Added 02/06/2008, lss

            .FlotPctRcvry = FlotCalcData.Rcvry

            .PbPctWt = gGetMatlWtPct("Pb", aRawProspBase, aSfcDataSprd)
            .ClPctWt = gGetMatlWtPct("Cl", aRawProspBase, aSfcDataSprd)
            .FdPctWt = gGetMatlWtPct("Fd", aRawProspBase, aSfcDataSprd)
            .TotPctWt = .PbPctWt + .ClPctWt + .FdPctWt

            .DryCoreLbsTot = Round((.CorePctSol / 100) * aRawProspBase.MtxTotWetWt, 1)
            .DryPbLbsAdj = Round((.PbPctWt / 100) * .DryCoreLbsTot, 1)
            .DryFdLbsAdj = Round((.FdPctWt / 100) * .DryCoreLbsTot, 1)
            .DryClLbsAdj = Round((.ClPctWt / 100) * .DryCoreLbsTot, 1)

            .DryCoreLbsProc = Round((.CorePctSol / 100) * aRawProspBase.MtxProcWetWt, 1)

            'For older prospect holes clay BPL will be available and we will show this
            'value.
            If .DryClLbsAdj <> 0 And (.DryPbLbsAdj + .DryFdLbsAdj + .DryClLbsAdj) <> 0 Then
                .HeadCalc = Round((.PbBpl * .DryPbLbsAdj + .FdBpl * .DryFdLbsAdj +
                            .ClBpl * .DryClLbsAdj) / (.DryPbLbsAdj + .DryFdLbsAdj + .DryClLbsAdj), 1)
            Else
                .HeadCalc = 0
            End If

            .SampleId = aRawProspBase.SampleId
            .Section = aRawProspBase.Section
            .Township = aRawProspBase.Township
            .Range = aRawProspBase.Range
            .HoleLocation = aRawProspBase.HoleLocation
            .SplitNumber = aRawProspBase.SplitNumber
            .ProspDate = aRawProspBase.ProspDate
        End With
    End Sub

    Public Sub gCalcFlotationData(ByRef aFlotCalcData As gFlotCalcDataType,
                                  ByVal aFdWetWt As Single,
                                  ByVal aFdDryWt As Single,
                                  ByVal aFdTareWt As Single,
                                  ByVal aFlotFdWetWt As Single,
                                  ByRef aSfcDataSprd() As gRawProspSfcSprdType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim PctRcvry As Single
        Dim FdBplCalc As Single
        Dim FdWtCalc As Single
        Dim FdBplAct As Single
        Dim FdWtAct As Single
        Dim SfcData As gSfcDataType
        Dim CnBpl As Single
        Dim CnWt As Single
        Dim TlBpl As Single
        Dim TlWt As Single

        PctRcvry = 0
        FdBplCalc = 0
        FdWtCalc = 0
        FdBplAct = 0
        FdWtAct = 0

        gGetSfcDataMatl(SfcData,
                        "Cn",
                        aSfcDataSprd)
        CnBpl = SfcData.Bpl
        CnWt = SfcData.Weight

        gGetSfcDataMatl(SfcData,
                        "Tl",
                        aSfcDataSprd)
        TlBpl = SfcData.Bpl
        TlWt = SfcData.Weight

        gGetSfcDataMatl(SfcData,
                        "Fd",
                        aSfcDataSprd)
        FdBplAct = SfcData.Bpl
        FdWtAct = CnWt + TlWt

        'Calculate flotation BPL
        If CnWt + TlWt <> 0 Then
            FdBplCalc = Round(((CnBpl * CnWt) + (TlBpl * TlWt)) /
                        (CnWt + TlWt), 1)
        Else
            FdBplCalc = 0
        End If

        'Calculate flotation feed weight
        If aFdWetWt - aFdTareWt <> 0 Then
            FdWtCalc = Round(aFlotFdWetWt * ((aFdDryWt - aFdTareWt) /
                       (aFdWetWt - aFdTareWt)), 0)
        Else
            FdWtCalc = 0
        End If

        'Calculate flotation %recovery
        If (CnBpl * CnWt) + (TlBpl * TlWt) <> 0 Then
            PctRcvry = Round((CnBpl * CnWt) /
                       ((CnBpl * CnWt) + (TlBpl * TlWt)) * 100, 1)
        Else
            PctRcvry = 0
        End If

        With aFlotCalcData
            .FdBplCalc = FdBplCalc
            .FdBplAct = FdBplAct
            .FdWtCalc = FdWtCalc
            .FdWtAct = FdWtAct
            .Rcvry = PctRcvry
            .CnBplAct = CnBpl
            .TlBplAct = TlBpl
        End With
    End Sub

    Public Function gGetSfcDataMatl(ByRef aSfcData As gSfcDataType,
                                    ByVal aMatl As String,
                                    ByRef aSfcDataSprd() As gRawProspSfcSprdType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim RowIdx As Integer
        Dim ThisMatl As String
        Dim ThisWt As Single

        'Get size fraction code data based on a material (Pb, Cn, Tl, etc.)
        'All size fraction code data for the particular material will
        'be summed and "returned" through aSfcData.

        'Pebble weight returned is the adjusted gram value!

        'These are summing variables
        Dim WtSum As Single
        Dim WtNum As Integer
        Dim WtWBpl As Single
        Dim WtBpl As Single
        Dim WtWIns As Single
        Dim WtIns As Single
        Dim WtWCa As Single
        Dim WtCa As Single
        Dim WtWMg As Single
        Dim WtMg As Single
        Dim WtWFe As Single
        Dim WtFe As Single
        Dim WtWAl As Single
        Dim WtAl As Single
        Dim WtWFeAl As Single
        Dim WtFeAl As Single
        Dim WtWCd As Single
        Dim WtCd As Single

        WtWBpl = 0
        WtBpl = 0
        WtWIns = 0
        WtIns = 0
        WtWCa = 0
        WtCa = 0
        WtWMg = 0
        WtMg = 0
        WtWFe = 0
        WtFe = 0
        WtWAl = 0
        WtAl = 0
        WtWFeAl = 0
        WtFeAl = 0
        WtWCd = 0
        WtCd = 0

        'aMatl will be Pb, Cn, Tl, Fd, Cl

        gZeroSfcData(aSfcData)

        For RowIdx = 1 To UBound(aSfcDataSprd)
            ThisMatl = aSfcDataSprd(RowIdx).SfcMatlName
            ThisWt = 0
            If ThisMatl = aMatl Then
                WtSum = WtSum + aSfcDataSprd(RowIdx).SizeFrctnWt
                WtNum = WtNum + 1
                ThisWt = aSfcDataSprd(RowIdx).SizeFrctnWt

                If aSfcDataSprd(RowIdx).Bpl <> 0 Then
                    WtWBpl = WtWBpl + ThisWt
                End If
                WtBpl = WtBpl + aSfcDataSprd(RowIdx).Bpl * ThisWt

                If aSfcDataSprd(RowIdx).Insol <> 0 Then
                    WtWIns = WtWIns + ThisWt
                End If
                WtIns = WtIns + aSfcDataSprd(RowIdx).Insol * ThisWt

                If aSfcDataSprd(RowIdx).CaO <> 0 Then
                    WtWCa = WtWCa + ThisWt
                End If
                WtCa = WtCa + aSfcDataSprd(RowIdx).CaO * ThisWt

                If aSfcDataSprd(RowIdx).MgO <> 0 Then
                    WtWMg = WtWMg + ThisWt
                End If
                WtMg = WtMg + aSfcDataSprd(RowIdx).MgO * ThisWt

                If aSfcDataSprd(RowIdx).Fe2O3 <> 0 Then
                    WtWFe = WtWFe + ThisWt
                End If
                WtFe = WtFe + aSfcDataSprd(RowIdx).Fe2O3 * ThisWt

                If aSfcDataSprd(RowIdx).Al2O3 <> 0 Then
                    WtWAl = WtWAl + ThisWt
                End If
                WtAl = WtAl + aSfcDataSprd(RowIdx).Al2O3 * ThisWt

                If aSfcDataSprd(RowIdx).FeAl <> 0 Then
                    WtWFeAl = WtWFeAl + ThisWt
                End If
                WtFeAl = WtFeAl + aSfcDataSprd(RowIdx).FeAl * ThisWt

                If aSfcDataSprd(RowIdx).Cd <> 0 Then
                    WtWCd = WtWCd + ThisWt
                End If
                WtCd = WtCd + aSfcDataSprd(RowIdx).Cd * ThisWt
            End If
        Next RowIdx

        With aSfcData
            .Weight = WtSum
            If WtWBpl <> 0 Then
                .Bpl = Round(WtBpl / WtWBpl, 1)
            Else
                .Bpl = 0
            End If
            '--
            If WtWIns <> 0 Then
                .Insol = Round(WtIns / WtWIns, 1)
            Else
                .Insol = 0
            End If
            '--
            If WtWCa <> 0 Then
                .CaO = Round(WtCa / WtWCa, 2)
            Else
                .CaO = 0
            End If
            '--
            If WtWMg <> 0 Then
                .MgO = Round(WtMg / WtWMg, 2)
            Else
                .MgO = 0
            End If
            '--
            If WtWFe <> 0 Then
                .Fe2O3 = Round(WtFe / WtWFe, 2)
            Else
                .Fe2O3 = 0
            End If
            '--
            If WtWAl <> 0 Then
                .Al2O3 = Round(WtAl / WtWAl, 2)
            Else
                .Al2O3 = 0
            End If
            '--
            If WtWFeAl <> 0 Then
                .FeAl = Round(WtFeAl / WtWFeAl, 2)
            Else
                .FeAl = 0
            End If
            '--
            If WtWCd <> 0 Then
                .Cd = Round(WtCd / WtWCd, 1)
            Else
                .Cd = 0
            End If
        End With
    End Function

    Public Function gCalcMoist(ByVal aWetWt As Single,
                               ByVal aDryWt As Single,
                               ByVal aTareWt As Single) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo CalcMoistError

        Dim PctSolids As Single

        PctSolids = 0

        '02/11/2010, lss
        'Added this functionality!  Previously it would have returned 100%
        'If aWetWt or aDryWt is <= 0 then return zero (not 100!)
        'If aTareWt < 0 then return 0 (Tare weight can be zero).
        If aWetWt <= 0 Or aDryWt <= 0 Or aTareWt < 0 Then
            gCalcMoist = 0
            Exit Function
        End If

        If aWetWt - aTareWt > 0 Then
            PctSolids = Round((aDryWt - aTareWt) /
                        (aWetWt - aTareWt), 4)
        Else
            PctSolids = 0
        End If

        gCalcMoist = 100 - Round(100 * PctSolids, 1)

        Exit Function

CalcMoistError:
        MsgBox("Error calculating %moisture." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Calculate %Moisture Error")
    End Function

    Public Function gCalcMoist2(ByVal aWetWt As Single,
                                ByVal aDryWt As Single,
                                ByVal aTareWt As Single,
                                ByVal aWetWt2 As Single,
                                ByVal aDryWt2 As Single,
                                ByVal aTareWt2 As Single) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo CalcMoist2Error

        Dim PctSolids As Single
        Dim PctSolids1 As Single
        Dim PctSolids2 As Single

        PctSolids = 0
        PctSolids1 = 0
        PctSolids2 = 0

        'If problems with #2 sample measure data then process as if a 1 measure
        'mtx %moisture situation
        'If aWetWt2 or aDryWt2 is <= 0 then use gCalcMoist.
        'If aTareWt2 < 0 then use gCalcMoist.
        If aWetWt2 <= 0 Or aDryWt2 <= 0 Or aTareWt2 < 0 Then
            gCalcMoist2 = gCalcMoist(aWetWt, aDryWt, aTareWt)
            Exit Function
        End If

        'Assume that we have a full set of #1 and #2 mtx %moisture data values!
        'Process #1 Mtx moisture data values
        If aWetWt - aTareWt > 0 Then
            PctSolids1 = Round((aDryWt - aTareWt) /
                        (aWetWt - aTareWt), 4)
        Else
            PctSolids1 = 0
        End If

        'Process #2 Mtx moisture data values
        If aWetWt2 - aTareWt2 > 0 Then
            PctSolids2 = Round((aDryWt2 - aTareWt2) /
                        (aWetWt2 - aTareWt2), 4)
        Else
            PctSolids2 = 0
        End If

        'Now average the two percents appropriately
        PctSolids = Round((PctSolids1 + PctSolids2) / 2, 4)

        gCalcMoist2 = 100 - Round(100 * PctSolids, 1)

        Exit Function

CalcMoist2Error:
        MsgBox("Error calculating 2 sample %moisture." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Calculate 2 Sample %Moisture Error")
    End Function

    Public Function gGetSfcPctPlusData(ByRef aSfcPlusMinusData As gSfcPlusMinusDataType,
                                       ByVal aMatl As String,
                                       ByVal aPlusTarg As String,
                                       ByRef aSfcDataSprd() As gRawProspSfcSprdType) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim RowIdx As Integer
        Dim ThisMatl As String
        Dim ThisDesc As String
        Dim ThisWt As Single
        Dim ThisBpl As Single
        Dim PlusValFound As Boolean
        Dim MeshTarg As String
        Dim SampWtPlus As Single
        Dim SampWtTot As Single
        Dim SumBplWtPlus As Single
        Dim SumWtWbplPlus As Single
        Dim SumBplWtMinus As Single
        Dim SumWtWbplMinus As Single

        MeshTarg = aPlusTarg & "M"
        SampWtPlus = 0
        SampWtTot = 0
        SumBplWtPlus = 0
        SumWtWbplPlus = 0
        SumBplWtMinus = 0
        SumWtWbplMinus = 0

        PlusValFound = False
        gGetSfcPctPlusData = 0

        For RowIdx = 1 To UBound(aSfcDataSprd)
            ThisMatl = aSfcDataSprd(RowIdx).SfcMatlName
            ThisDesc = aSfcDataSprd(RowIdx).SfcDescription
            ThisWt = aSfcDataSprd(RowIdx).SizeFrctnWt
            ThisBpl = aSfcDataSprd(RowIdx).Bpl

            'Example -- Feed + 35M
            'Will have something like this in aSfcDataSprd()
            'SFC   Description  Matl   Weight    BPL
            '---   -----------  ----   ------    ----
            '051   -16M +20M    Fd     6         65.5
            '052   -20M +35M    Fd     45        44.2
            '070   -35M +150M   Fd     764       19.6

            'SFC codes will be in order of decreasing size
            'Sum up and including when we find +35M in the description

            If ThisMatl = aMatl Then
                If InStr(ThisDesc, MeshTarg) <> 0 Then
                    'Add to plus mesh stuff & to total stuff
                    SampWtPlus = SampWtPlus + ThisWt
                    SampWtTot = SampWtTot + ThisWt
                    '----
                    SumBplWtPlus = SumBplWtPlus + ThisWt * ThisBpl
                    If ThisBpl <> 0 Then
                        SumWtWbplPlus = SumWtWbplPlus + ThisWt
                    End If
                    PlusValFound = True
                Else
                    If PlusValFound = False Then
                        'Add to plus mesh stuff & to total stuff
                        SampWtPlus = SampWtPlus + ThisWt
                        SampWtTot = SampWtTot + ThisWt
                        '----
                        SumBplWtPlus = SumBplWtPlus + ThisWt * ThisBpl
                        If ThisBpl <> 0 Then
                            SumWtWbplPlus = SumWtWbplPlus + ThisWt
                        End If
                    Else
                        'Add to total stuff only
                        SampWtTot = SampWtTot + ThisWt
                        '----
                        SumBplWtMinus = SumBplWtMinus + ThisWt * ThisBpl
                        If ThisBpl <> 0 Then
                            SumWtWbplMinus = SumWtWbplMinus + ThisWt
                        End If
                    End If
                End If
            End If
        Next RowIdx

        If SampWtTot <> 0 Then
            gGetSfcPctPlusData = Round(SampWtPlus / SampWtTot * 100, 2)
        Else
            gGetSfcPctPlusData = 0
        End If

        With aSfcPlusMinusData
            .PlusWt = SampWtPlus
            If SumWtWbplPlus <> 0 Then
                .PlusBpl = SumBplWtPlus / SumWtWbplPlus
            Else
                .PlusBpl = 0
            End If

            .MinusWt = SampWtTot - SampWtPlus
            If SumWtWbplMinus <> 0 Then
                .MinusBpl = SumBplWtMinus / SumWtWbplMinus
            Else
                .MinusBpl = 0
            End If
        End With
    End Function

    Public Function gGetMatlWtPct(ByVal aMatl As String,
                                  ByRef aRawProspBase As gRawProspBaseType,
                                  ByRef aSfcDataSprd() As gRawProspSfcSprdType) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim MtxPctMoist As Single
        Dim MtxTotWetWt As Single
        Dim MtxTotDryWt As Single
        Dim FdPctMoist As Single
        Dim FdTotWetWt As Single
        Dim FdTotDryWt As Single
        Dim PbWtLbsAdj As Single
        Dim SfcData As gSfcDataType
        Dim FdWtPct As Single
        Dim PbWtPct As Single
        Dim ClWtPct As Single

        'MtxPctMoist = gCalcMoist(aRawProspBase.MtxMoistWetWt, _
        '                         aRawProspBase.MtxMoistDryWt, _
        '                         aRawProspBase.MtxMoistTareWt)

        '02/16/2010, lss
        'Need to get matrix %moisture based on 2 sets of matrix moisture samples!
        'The second set of matrix moisture samples (wet, dry, tare) will all be
        'zeros for older holes -- gCalcMoist2 will handle this correctly!
        MtxPctMoist = gCalcMoist2(aRawProspBase.MtxMoistWetWt,
                                  aRawProspBase.MtxMoistDryWt,
                                  aRawProspBase.MtxMoistTareWt,
                                  aRawProspBase.MtxMoistWetWt2,
                                  aRawProspBase.MtxMoistDryWt2,
                                  aRawProspBase.MtxMoistTareWt2)

        MtxTotDryWt = Round((100 - MtxPctMoist) / 100 * aRawProspBase.MtxTotWetWt, 1)

        FdPctMoist = gCalcMoist(aRawProspBase.FdMoistWetWt,
                                aRawProspBase.FdMoistDryWt,
                                aRawProspBase.FdMoistTareWt)
        FdTotDryWt = Round((100 - FdPctMoist) / 100 * aRawProspBase.FdTotWetWt, 1)

        gGetSfcDataMatl(SfcData,
                        "Pb",
                        aSfcDataSprd)
        'Pebble weight returned is the "adjusted" value!
        PbWtLbsAdj = Round(SfcData.Weight / 453.6, 1)

        If MtxTotDryWt <> 0 Then
            PbWtPct = Round(PbWtLbsAdj / MtxTotDryWt * 100, 1)
        Else
            PbWtPct = 0
        End If

        If MtxTotDryWt <> 0 Then
            FdWtPct = Round(FdTotDryWt / MtxTotDryWt * 100, 1)
        Else
            FdWtPct = 0
        End If

        ClWtPct = 100 - FdWtPct - PbWtPct

        If ClWtPct < 0 Then
            ClWtPct = 0
        End If

        Select Case aMatl
            Case Is = "Pb"
                gGetMatlWtPct = PbWtPct
            Case Is = "Fd"
                gGetMatlWtPct = FdWtPct
            Case Is = "Cl"
                gGetMatlWtPct = ClWtPct
            Case Else
                gGetMatlWtPct = 0
        End Select
    End Function

    Public Function gCalcDryDensity(ByRef aRawProspBase As gRawProspBaseType) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim MtxPctSolids As Single

        With aRawProspBase
            'Will assume that if there are valid matrix %moisture samples they
            'will be in the 1st set!
            If .MtxMoistWetWt = 0 And .MtxMoistDryWt = 0 Then
                gCalcDryDensity = 0
                Exit Function
            End If

            '04/15/2008, lss  Fix for 0.33 errors!
            If .DensCylSize = 0.33 Then
                .DensCylSize = 0.3
            End If

            '02/19/2010, lss
            'May have 2 sets of matrix %moisture samples.
            'If we have 2 valid sets then need to average the MtxMoistWetWt's and the
            'MtxMoistDryWt's.

            MtxPctSolids = gCalcSolids2(.MtxMoistWetWt,
                                        .MtxMoistDryWt,
                                        .MtxMoistTareWt,
                                        .MtxMoistWetWt2,
                                        .MtxMoistDryWt2,
                                        .MtxMoistTareWt2,
                                        4)

            'Old 1 set of matrix %moisture samples method.
            'If (.DensCylSize - (.DensCylH2oWt / 62.43)) <> 0 Then
            '    gCalcDryDensity = Round(.DensCylWetWt * ((.MtxMoistDryWt - .MtxMoistTareWt) / _
            '                     (.MtxMoistWetWt - .MtxMoistTareWt)) / _
            '                     (.DensCylSize - (.DensCylH2oWt / 62.43)), 2)
            'Else
            '    gCalcDryDensity = 0
            'End If

            'New 1 set of matrix %moisture samples method.
            If (.DensCylSize - (.DensCylH2oWt / 62.43)) <> 0 Then
                gCalcDryDensity = Round(.DensCylWetWt * (MtxPctSolids / 100) /
                                 (.DensCylSize - (.DensCylH2oWt / 62.43)), 2)
            Else
                gCalcDryDensity = 0
            End If
        End With
    End Function

    Public Function gCalcFdTotWetWtAdj(ByRef aRawProspBase As gRawProspBaseType) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Will not calculate this for holes from mainframe.
        'Will only calculate for new holes added to the new system.

        With aRawProspBase
            If .MtxTotWetWt <> 0 Then
                gCalcFdTotWetWtAdj = Round(.MtxTotWetWt * .FdTotWetWtMsr /
                                    .MtxProcWetWt, 1)
            Else
                gCalcFdTotWetWtAdj = 0
            End If
        End With
    End Function

    'Public Sub gGetSfcDataFromInputs(ByRef aSfcDataSprd() As gRawProspSfcSprdType, _
    '                                 ByRef aSsSizeFrctnData As vaSpread)

    '    '**********************************************************************
    '    '
    '    '
    '    '
    '    '**********************************************************************

    '    Dim RowIdx As Integer

    '    'Get the data from ssSizeFrctnData
    '    With aSsSizeFrctnData
    '        ReDim aSfcDataSprd(.MaxRows)

    '        For RowIdx = 1 To .MaxRows
    '            .Row = RowIdx
    '            .Col = 1
    '            aSfcDataSprd(RowIdx).SizeFrctnCode = .Value

    '            .Col = 3
    '            aSfcDataSprd(RowIdx).SfcDescription = .Value

    '            .Col = 4
    '            aSfcDataSprd(RowIdx).SfcMatlName = .Value

    '            'Either Col5 or Col6 should have a weight in it!
    '            .Col = 5
    '            If IsNumeric(.Value) Then
    '                aSfcDataSprd(RowIdx).SizeFrctnWtMsr = .Value
    '            End If

    '            .Col = 6
    '            If IsNumeric(.Value) Then
    '                aSfcDataSprd(RowIdx).SizeFrctnWtMsr = .Value
    '            End If

    '            .Col = 7
    '            If IsNumeric(.Value) Then
    '                aSfcDataSprd(RowIdx).Bpl = .Value
    '            Else
    '                aSfcDataSprd(RowIdx).Bpl = 0
    '            End If

    '            .Col = 8
    '            If IsNumeric(.Value) Then
    '                aSfcDataSprd(RowIdx).Insol = .Value
    '            Else
    '                aSfcDataSprd(RowIdx).Insol = 0
    '            End If

    '            .Col = 9
    '            If IsNumeric(.Value) Then
    '                aSfcDataSprd(RowIdx).CaO = .Value
    '            Else
    '                aSfcDataSprd(RowIdx).CaO = 0
    '            End If

    '            .Col = 10
    '            If IsNumeric(.Value) Then
    '                aSfcDataSprd(RowIdx).MgO = .Value
    '            Else
    '                aSfcDataSprd(RowIdx).MgO = 0
    '            End If

    '            .Col = 11
    '            If IsNumeric(.Value) Then
    '                aSfcDataSprd(RowIdx).Fe2O3 = .Value
    '            Else
    '                aSfcDataSprd(RowIdx).Fe2O3 = 0
    '            End If

    '            .Col = 12
    '            If IsNumeric(.Value) Then
    '                aSfcDataSprd(RowIdx).Al2O3 = .Value
    '            Else
    '                aSfcDataSprd(RowIdx).Al2O3 = 0
    '            End If

    '            .Col = 13
    '            If IsNumeric(.Value) Then
    '                aSfcDataSprd(RowIdx).FeAl = .Value
    '            Else
    '                aSfcDataSprd(RowIdx).FeAl = 0
    '            End If

    '            .Col = 14
    '            If IsNumeric(.Value) Then
    '                aSfcDataSprd(RowIdx).Cd = .Value
    '            Else
    '                aSfcDataSprd(RowIdx).Cd = 0
    '            End If

    '            .Col = 15
    '            If IsNumeric(.Value) Then
    '                aSfcDataSprd(RowIdx).SizeFrctnWtAdj = .Value
    '            Else
    '                aSfcDataSprd(RowIdx).SizeFrctnWtAdj = 0
    '            End If

    '            .Col = 16
    '            aSfcDataSprd(RowIdx).SizeFrctnType = .Value

    '            aSfcDataSprd(RowIdx).SizeFrctnWt = 0
    '            If aSfcDataSprd(RowIdx).SizeFrctnType = "T" Then
    '                aSfcDataSprd(RowIdx).SizeFrctnWt = aSfcDataSprd(RowIdx).SizeFrctnWtAdj
    '            End If
    '            If aSfcDataSprd(RowIdx).SizeFrctnType = "P" Then
    '                aSfcDataSprd(RowIdx).SizeFrctnWt = aSfcDataSprd(RowIdx).SizeFrctnWtMsr
    '            End If
    '        Next RowIdx
    '    End With
    'End Sub

    Public Function gGetMatlNameAbbrv(ByVal aMatlName As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Materials associated with size fraction codes are:
        'Pebble
        'Concentrate
        'Tails
        'Head
        'Feed
        'Clay
        'Pan

        gGetMatlNameAbbrv = "??"
        aMatlName = StrConv(aMatlName, vbUpperCase)

        Select Case aMatlName
            Case Is = "PEBBLE"
                gGetMatlNameAbbrv = "Pb"
            Case Is = "CONCENTRATE"
                gGetMatlNameAbbrv = "Cn"
            Case Is = "TAILS"
                gGetMatlNameAbbrv = "Tl"
            Case Is = "HEAD"
                gGetMatlNameAbbrv = "Hd"
            Case Is = "FEED"
                gGetMatlNameAbbrv = "Fd"
            Case Is = "CLAY"
                gGetMatlNameAbbrv = "Cl"
            Case Is = "PAN"
                gGetMatlNameAbbrv = "Pn"
        End Select
    End Function

    Public Function gGetSplitsForHole(ByVal aTownship As Integer,
                                      ByVal aRange As Integer,
                                      ByVal aSection As Integer,
                                      ByVal aHoleLocation As String,
                                      ByVal aProspDate As String,
                                      ByRef AllSplitBaseData() As gRawProspLoctnType) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetSplitsForHoleError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer
        Dim ProspRawDynaset As OraDynaset
        Dim SplCnt As Integer

        gGetSplitsForHole = -1

        'Get all of the splits for this hole from PROSP_RAW_SPLIT
        params = gDBParams

        params.Add("pTownship", aTownship, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRange, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSection, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHoleLocation, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", CDate(aProspDate), ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pResult", "", ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_splits_for_hole
        'pTownship                  IN     NUMBER,
        'pRange                     IN     NUMBER,
        'pSection                   IN     NUMBER,
        'pHoleLocation              IN     VARCHAR2,
        'pProspDate                 IN     DATE,
        'pResult                    IN OUT c_prosprawbase)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_splits_for_hole(:pTownship, " +
                  ":pRange, :pSection, :pHoleLocation, :pProspDate, :pResult);end;", ORASQL_FAILEXEC)

        ProspRawDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = ProspRawDynaset.RecordCount
        'Return the number of splits that were found
        gGetSplitsForHole = RecordCount

        If RecordCount <> 0 Then
            ReDim AllSplitBaseData(RecordCount)
            SplCnt = 0
            ProspRawDynaset.MoveFirst()

            Do While Not ProspRawDynaset.EOF
                SplCnt = SplCnt + 1
                With AllSplitBaseData(SplCnt)
                    .SampleId = ProspRawDynaset.Fields("sample_id").Value
                    .Township = ProspRawDynaset.Fields("township").Value
                    .Range = ProspRawDynaset.Fields("range").Value
                    .Section = ProspRawDynaset.Fields("section").Value
                    .HoleLocation = ProspRawDynaset.Fields("hole_location").Value
                    .ProspDate = ProspRawDynaset.Fields("prosp_date").Value
                    .SplitNumber = ProspRawDynaset.Fields("split_number").Value
                End With
                ProspRawDynaset.MoveNext()
            Loop
        End If

        Exit Function

gGetSplitsForHoleError:
        MsgBox("Error getting splits for hole." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Splits for Hole Error")

        On Error Resume Next
        gGetSplitsForHole = -1
        ClearParams(params)
    End Function

    Public Function gGetProspRawHoleDataOnly(ByVal aTwp As Integer,
                                             ByVal aRge As Integer,
                                             ByVal aSec As Integer,
                                             ByVal aHloc As String,
                                             ByVal aProspDate As Date,
                                             ByRef aProspRawHoleData As gRawProspBaseHoleType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspRawHoleDataOnlyError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ProspRawDynaset As OraDynaset
        Dim RecordCount As Integer

        'PROCEDURE get_prosp_raw_base
        'pTownship           IN     NUMBER,
        'pRange              IN     NUMBER,
        'pSection            IN     NUMBER,
        'pHoleLocation       IN     VARCHAR2,
        'pProspDate          IN     DATE,
        'pResult             IN OUT c_prosprawbase)

        params = gDBParams

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHloc, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_base(" &
                  ":pTownship, :pRange, :pSection, :pHoleLocation, " &
                  ":pProspDate, :pResult);end;", ORASQL_FAILEXEC)

        ProspRawDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = ProspRawDynaset.RecordCount

        If RecordCount = 1 Then
            gGetProspRawHoleDataOnly = True

            ProspRawDynaset.MoveFirst()
            With aProspRawHoleData
                .Township = ProspRawDynaset.Fields("township").Value
                .Range = ProspRawDynaset.Fields("range").Value
                .Section = ProspRawDynaset.Fields("section").Value
                .HoleLocation = ProspRawDynaset.Fields("hole_location").Value
                .Forty = ProspRawDynaset.Fields("forty").Value
                .State = ProspRawDynaset.Fields("state").Value
                .Quadrant = ProspRawDynaset.Fields("quadrant").Value

                If Not IsDBNull(ProspRawDynaset.Fields("mine_name").Value) Then
                    .MineName = ProspRawDynaset.Fields("mine_name").Value
                Else
                    .MineName = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("spec_area_name").Value) Then
                    .SpecAreaName = ProspRawDynaset.Fields("spec_area_name").Value
                Else
                    .SpecAreaName = ""
                End If

                .ExpDrill = ProspRawDynaset.Fields("exp_drill").Value
                .SplitTotalNum = ProspRawDynaset.Fields("split_total_num").Value
                .Xcoord = ProspRawDynaset.Fields("x_coord").Value
                .Ycoord = ProspRawDynaset.Fields("y_coord").Value
                .FtlDepth = ProspRawDynaset.Fields("ftl_depth").Value
                .OvbCored = ProspRawDynaset.Fields("ovb_cored").Value
                .Ownership = ProspRawDynaset.Fields("ownership").Value

                .ProspDate = Format(ProspRawDynaset.Fields("prosp_date").Value, "MM/dd/yyyy")

                .MinedStatus = ProspRawDynaset.Fields("mined_status").Value
                .Elevation = ProspRawDynaset.Fields("elevation").Value
                .TotDepth = ProspRawDynaset.Fields("tot_depth").Value
                .Aoi = ProspRawDynaset.Fields("aoi").Value
                .CoordSurveyed = ProspRawDynaset.Fields("coord_surveyed").Value

                If Not IsDBNull(ProspRawDynaset.Fields("long_comment").Value) Then
                    .HoleComment = ProspRawDynaset.Fields("long_comment").Value
                Else
                    .HoleComment = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("hole_location_char").Value) Then
                    .HoleLocationChar = ProspRawDynaset.Fields("hole_location_char").Value
                Else
                    .HoleLocationChar = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("who_modified").Value) Then
                    .WhoModifiedHole = ProspRawDynaset.Fields("who_modified").Value
                Else
                    .WhoModifiedHole = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("when_modified").Value) Then
                    'Want date and time!
                    .WhenModifiedHole = Format(ProspRawDynaset.Fields("when_modified").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .WhenModifiedHole = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("log_date").Value) Then
                    .LogDate = Format(ProspRawDynaset.Fields("log_date").Value, "MM/dd/yyyy")
                Else
                    .LogDate = ""
                End If

                .Released = ProspRawDynaset.Fields("released").Value
                .Redrilled = ProspRawDynaset.Fields("redrilled").Value

                If Not IsDBNull(ProspRawDynaset.Fields("redrill_date").Value) Then
                    .RedrillDate = Format(ProspRawDynaset.Fields("redrill_date").Value, "MM/dd/yyyy")
                Else
                    .RedrillDate = ""
                End If

                .UseForReduction = ProspRawDynaset.Fields("use_for_reduction").Value

                .QaQcHole = ProspRawDynaset.Fields("qaqc_hole").Value

                '-----
                'New columns added 10/17/2011, lss
                If Not IsDBNull(ProspRawDynaset.Fields("hardpan_from").Value) Then
                    .HardpanFrom = ProspRawDynaset.Fields("hardpan_from").Value
                Else
                    .HardpanFrom = 0
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("hardpan_to").Value) Then
                    .HardpanTo = ProspRawDynaset.Fields("hardpan_to").Value
                Else
                    .HardpanTo = 0
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("hardpan_code").Value) Then
                    .HardpanCode = ProspRawDynaset.Fields("hardpan_code").Value
                Else
                    .HardpanCode = "0"
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("hardpan_thck").Value) Then
                    .HardpanThck = ProspRawDynaset.Fields("hardpan_thck").Value
                Else
                    .HardpanThck = 0
                End If
            End With
        Else
            gGetProspRawHoleDataOnly = False
        End If

        ProspRawDynaset.Close()

        Exit Function

gGetProspRawHoleDataOnlyError:
        gGetProspRawHoleDataOnly = False

        MsgBox("Error accessing raw prospect hole only data." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Raw Prospect Hole Only Data Access Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        ProspRawDynaset.Close()
    End Function

    Public Function gGetProspHoleCount(ByVal aTwp As Integer,
                                       ByVal aRge As Integer,
                                       ByVal aSec As Integer,
                                       ByVal aHloc As String) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspHoleCountError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        'PROCEDURE prosp_hole_count
        'pSection               IN     NUMBER,
        'pTownship              IN     NUMBER,
        'pRange                 IN     NUMBER,
        'pHoleLocation          IN     VARCHAR2,
        'pResult                IN OUT NUMBER);

        params = gDBParams

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHloc, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.prosp_hole_count(" &
                  ":pSection, :pTownship, :pRange, :pHoleLocation, " &
                  ":pResult);end;", ORASQL_FAILEXEC)

        gGetProspHoleCount = params("pResult").Value
        ClearParams(params)

        Exit Function

gGetProspHoleCountError:
        gGetProspHoleCount = 0

        MsgBox("Error accessing raw prospect hole count." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Raw Prospect Hole Count Error")

        On Error Resume Next
        ClearParams(params)
    End Function

    Public Sub gGetSizeFrctnCodes(ByRef aDynaset As OraDynaset,
                                  ByVal aOrder As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetSizeFrctnCodesError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecCount As Integer

        params = gDBParams

        params.Add("pOrderMode", aOrder, ORAPARM_INPUT)
        params("pOrderMode").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_all_sfrctn_codes
        'pOrderMode          IN     VARCHAR2,
        'pResult             IN OUT c_sfrctncodes);
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospectnew.get_all_sfrctn_codes(" +
                  ":pOrderMode, :pResult);end;", ORASQL_FAILEXEC)
        aDynaset = params("pResult").Value
        ClearParams(params)

        RecCount = aDynaset.RecordCount

        Exit Sub

gGetSizeFrctnCodesError:
        MsgBox("Error getting size fraction codes." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Size Fraction Codes Get Error")

        On Error Resume Next
        ClearParams(params)
    End Sub

    Public Function gGetSfcFromAlphaCode(ByVal aAlphaCode As String,
                                         ByVal aMatlAbbrv As String,
                                         ByVal aExpStatus As String,
                                         ByRef aSfcDynaset As OraDynaset)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ThisRegAlphaCode As String
        Dim ThisRegAnalysisCode As String
        Dim ThisExpAlphaCode As String
        Dim ThisExpAnalysisCode As String
        Dim ThisSizeFrctnCode As String
        Dim ThisMatlAbbrv As String
        Dim ThisCode As String

        gGetSfcFromAlphaCode = "?"

        aAlphaCode = StrConv(aAlphaCode, vbUpperCase)
        aMatlAbbrv = StrConv(aMatlAbbrv, vbUpperCase)

        aSfcDynaset.MoveFirst()
        'Looking at all of the size fraction codes!
        Do While Not aSfcDynaset.EOF
            If Not IsDBNull(aSfcDynaset.Fields("reg_alpha_code").Value) Then
                ThisRegAlphaCode = StrConv(aSfcDynaset.Fields("reg_alpha_code").Value, vbUpperCase)
            Else
                ThisRegAlphaCode = ""
            End If
            If Not IsDBNull(aSfcDynaset.Fields("reg_analysis_code").Value) Then
                ThisRegAnalysisCode = StrConv(aSfcDynaset.Fields("reg_analysis_code").Value, vbUpperCase)
            Else
                ThisRegAnalysisCode = ""
            End If

            If Not IsDBNull(aSfcDynaset.Fields("exp_alpha_code").Value) Then
                ThisExpAlphaCode = StrConv(aSfcDynaset.Fields("exp_alpha_code").Value, vbUpperCase)
            Else
                ThisExpAlphaCode = ""
            End If
            If Not IsDBNull(aSfcDynaset.Fields("exp_analysis_code").Value) Then
                ThisExpAnalysisCode = StrConv(aSfcDynaset.Fields("exp_analysis_code").Value, vbUpperCase)
            Else
                ThisExpAnalysisCode = ""
            End If

            ThisSizeFrctnCode = aSfcDynaset.Fields("size_frctn_code").Value
            ThisMatlAbbrv = StrConv(aSfcDynaset.Fields("matl_abbrv").Value, vbUpperCase)

            'Regular analysis hole
            If StrConv(aExpStatus, vbUpperCase) = "NO" Then
                'Need to look at regular alpha code
                If ThisRegAlphaCode = aAlphaCode Or
                    ThisRegAnalysisCode = aAlphaCode Then
                    If ThisMatlAbbrv = aMatlAbbrv Then
                        gGetSfcFromAlphaCode = ThisSizeFrctnCode
                    Else
                        gGetSfcFromAlphaCode = "?"
                    End If
                    Exit Function
                End If
            End If

            'Expanded analysis hole
            If StrConv(aExpStatus, vbUpperCase) = "YES" Then
                'Need to look at expanded alpha code
                If ThisExpAlphaCode = aAlphaCode Or
                   ThisExpAnalysisCode = aAlphaCode Then
                    If ThisMatlAbbrv = aMatlAbbrv Then
                        gGetSfcFromAlphaCode = ThisSizeFrctnCode
                    Else
                        gGetSfcFromAlphaCode = "?"
                    End If
                    Exit Function
                End If
            End If

            aSfcDynaset.MoveNext()
        Loop

        '11/23/2009
        'Special temporary fox for Ona-Pioneer drilling
        'Regular Analyis
        '---------------
        '
        '   SFC   Material   Alpha Code   Analysis Code   Alpha Code Old
        '   ---   --------   ----------   -------------   --------------
        '1) 004      Pb          P1             1                A
        '2) 04G      Pb          P2             2                B
        '3) 021      Pb          P3             3                C
        '4) 041      Cn          C1             9                G
        '5) 042      Tl          T1             0                H
        '6) 051      Fd          F1             5                D
        '7) 052      Fd          F2             6                E
        '8) 070      Fd          F3             7                F
        '9) 090      Cl          W1             --               --

        '11/23/2009 -- Now doing +3/8 & -3/8 +6M instead of
        '                        +1/2 & -1/2 +6M
        '              +3/8 & -3/8 are the new P1 & P2!
        '
        '1) 009      Pb          P1             1                A
        '2) 011      Pb          P2             2                B


    End Function

    Public Sub gGetProspCodesToCbo(ByRef aCboBox As ComboBox,
                                   ByVal aProspCodeTypeName As String,
                                   ByVal aAddBlankSelection As Boolean,
                                   ByVal aBlankSelectionStr As String,
                                   ByVal aDescriptions As Boolean)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspCodesToCboError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim CodeDynaset As OraDynaset
        Dim ThisCode As String
        Dim ThisCodeDesc As String
        Dim ItemIdx As Integer

        'Clear the by reference combo box
        For ItemIdx = 0 To aCboBox.Items.Count - 1
            aCboBox.Items.RemoveAt(0)
        Next ItemIdx

        params = gDBParams

        params.Add("pProspCodeTypeName", aProspCodeTypeName, ORAPARM_INPUT)
        params("pProspCodeTypeName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_codes
        'pProspCodeTypeName   IN     VARCHAR2,
        'pResult              IN OUT c_prospcodes)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospectnew.get_prosp_codes(" +
                      ":pProspCodeTypeName, :pResult);end;", ORASQL_FAILEXEC)
        CodeDynaset = params("pResult").Value
        ClearParams(params)

        If aAddBlankSelection = True Then
            aCboBox.Items.Add(aBlankSelectionStr)
        End If

        CodeDynaset.MoveFirst()
        Do While Not CodeDynaset.EOF
            ThisCode = CodeDynaset.Fields("prosp_code").Value
            ThisCodeDesc = CodeDynaset.Fields("prosp_code_desc").Value

            If aDescriptions = True Then
                aCboBox.Items.Add(ThisCode & "-" & ThisCodeDesc)
            Else
                aCboBox.Items.Add(ThisCode)
            End If

            CodeDynaset.MoveNext()
        Loop

        CodeDynaset.Close()

        aCboBox.Text = aCboBox.Items(0)

        Exit Sub

gGetProspCodesToCboError:
        MsgBox("Error getting prospect codes." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Prospect Codes Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        CodeDynaset.Close()
    End Sub

    Public Function gGetProspCode(ByVal aProspCode As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim DashPos As Integer

        DashPos = InStr(aProspCode, "-")

        If DashPos <> 0 Then
            gGetProspCode = Mid(aProspCode, 1, DashPos - 1)
        Else
            If aProspCode = "(Select...)" Then
                gGetProspCode = ""
            Else
                gGetProspCode = aProspCode
            End If
        End If
    End Function

    Public Function gGetProspCodePlusDesc(ByVal aProspCodeTypeName As String,
                                          ByVal aProspCode As String,
                                          ByRef aProspCodeDynaset As OraDynaset) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ThisType As String
        Dim ThisCode As String
        Dim ThisDesc As String

        gGetProspCodePlusDesc = ""

        ' 1) Prospect code types
        ' 2) Clay settling characteristics
        ' 3) Degree of consolidation
        ' 4) Digging characteristics
        ' 5) Hardness code
        ' 6) Lithology
        ' 7) Matrix color
        ' 8) Ownership
        ' 9) Phosphate color
        '10) Pumping characteristics
        '11) Quadrant
        '12) State

        aProspCodeDynaset.MoveFirst()
        Do While Not aProspCodeDynaset.EOF
            ThisType = aProspCodeDynaset.Fields("prosp_code_type_name").Value
            ThisCode = aProspCodeDynaset.Fields("prosp_code").Value
            ThisDesc = aProspCodeDynaset.Fields("prosp_code_desc").Value

            If StrConv(ThisType, vbUpperCase) = StrConv(aProspCodeTypeName, vbUpperCase) And
                StrConv(ThisCode, vbUpperCase) = StrConv(aProspCode, vbUpperCase) Then

                gGetProspCodePlusDesc = ThisCode & "-" & ThisDesc
            End If

            aProspCodeDynaset.MoveNext()
        Loop
    End Function

    Public Sub gGetBankCodesToCbo(ByRef aCboBox As ComboBox,
                                  ByVal aMineName As String,
                                  ByVal aAddBlankSelection As Boolean,
                                  ByVal aBlankSelectionStr As String,
                                  ByVal aDescriptions As Boolean)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetBankCodesToCboError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim CodeDynaset As OraDynaset
        Dim ThisCode As String
        Dim ThisCodeDesc As String
        Dim ItemIdx As Integer
        Dim RecordCount As Integer

        'Clear the by reference combo box
        For ItemIdx = 0 To aCboBox.Items.Count - 1
            aCboBox.Items.RemoveAt(0)
        Next ItemIdx

        'Get all existing mining royalty areas for mine
        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_all_mine_royaltyarea
        'pMineName            IN     VARCHAR2,
        'pResult              IN OUT c_royaltyarea)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_survey.get_all_mine_royaltyarea(:pMineName," +
                  ":pResult);end;", ORASQL_FAILEXEC)
        CodeDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = CodeDynaset.RecordCount

        If aAddBlankSelection = True Then
            aCboBox.Items.Add(aBlankSelectionStr)
        End If

        CodeDynaset.MoveFirst()
        Do While Not CodeDynaset.EOF
            ThisCode = CodeDynaset.Fields("area_code").Value   'Bank code
            ThisCodeDesc = CodeDynaset.Fields("description").Value

            If aDescriptions = True Then
                aCboBox.Items.Add(ThisCode & "-" & ThisCodeDesc)
            Else
                aCboBox.Items.Add(ThisCode)
            End If

            CodeDynaset.MoveNext()
        Loop

        CodeDynaset.Close()

        aCboBox.Text = aCboBox.Items(0)

        Exit Sub

gGetBankCodesToCboError:
        MsgBox("Error getting bank codes." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Bank Codes Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        CodeDynaset.Close()
    End Sub

    Public Function gGetBankCodePlusDesc(ByVal aBankCode As String,
                                         ByRef aBankCodeDynaset As OraDynaset) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ThisCode As String
        Dim ThisDesc As String

        gGetBankCodePlusDesc = ""

        aBankCodeDynaset.MoveFirst()
        Do While Not aBankCodeDynaset.EOF
            ThisCode = aBankCodeDynaset.Fields("area_code").Value
            ThisDesc = aBankCodeDynaset.Fields("description").Value

            If StrConv(ThisCode, vbUpperCase) = StrConv(aBankCode, vbUpperCase) Then
                gGetBankCodePlusDesc = ThisCode & "-" & ThisDesc
            End If

            aBankCodeDynaset.MoveNext()
        Loop
    End Function

    Public Function gGetExpandedStatus(ByVal aSampleId As String) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetExpandedStatusError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ProspRawDynaset As OraDynaset
        Dim RecordCount As Integer

        'Need to determine if the hole associated with this sample ID is
        'an expanded hole.
        ' gGetExpandedStatus = 1  --> Expanded hole
        ' gGetExpandedStatus = 0  --> Regular hole
        ' gGetExpandedStatus = -1 --> Can't find sample ID in MOIS!

        gGetExpandedStatus = 0

        'PROCEDURE get_expanded_hole_status
        'pSampleId           IN     VARCHAR2,
        'pResult             IN OUT NUMBER)
        params = gDBParams

        params.Add("pSampleId", aSampleId, ORAPARM_INPUT)
        params("pSampleId").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_expanded_hole_status(" &
                  ":pSampleId, :pResult);end;", ORASQL_FAILEXEC)

        gGetExpandedStatus = params("pResult").Value
        ClearParams(params)

        Exit Function

gGetExpandedStatusError:
        gGetExpandedStatus = -1

        On Error Resume Next
        ClearParams(params)
        'Don't send a message to the user!
        'MsgBox "Error determining expanded status." & vbCrLf & _
        '       Err.Description, _
        '       vbOKOnly + vbExclamation, _
        '       "Expanded Status Error"
    End Function

    Public Function gGetAlphaCodeSfcDesc(ByVal aProcessMode As String,
                                         ByVal aAlphaCode As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetAlphaCodeSfcDescError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ProspRawSfcDynaset As OraDynaset
        Dim RecordCount As Integer
        Dim Description As String

        params = gDBParams

        params.Add("pProcessMode", aProcessMode, ORAPARM_INPUT)
        params("pProcessMode").serverType = ORATYPE_VARCHAR2

        params.Add("pAlphaCode", aAlphaCode, ORAPARM_INPUT)
        params("pAlphaCode").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_alpha_code_sfrctn
        'pProcessMode        IN     VARCHAR2,
        'pAlphaCode          IN     VARCHAR2,
        'pResult             IN OUT c_sfrctncodes);
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospectnew.get_alpha_code_sfrctn(" &
                  ":pProcessMode, :pAlphaCode, :pResult);end;", ORASQL_FAILEXEC)

        ProspRawSfcDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = ProspRawSfcDynaset.RecordCount

        If RecordCount <> 1 Then
            gGetAlphaCodeSfcDesc = "?"
        Else
            ProspRawSfcDynaset.MoveFirst()
            Description = ProspRawSfcDynaset.Fields("description").Value
            gGetAlphaCodeSfcDesc = Description
        End If

        ProspRawSfcDynaset.Close()

        Exit Function

gGetAlphaCodeSfcDescError:
        gGetAlphaCodeSfcDesc = "?"

        MsgBox("Error accessing alpha code sfc description." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Alpha Code Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        ProspRawSfcDynaset.Close()
    End Function

    Public Function gGetProspRawCoordsExist(ByVal aTwp As Integer,
                                            ByVal aRge As Integer,
                                            ByVal aSec As Integer,
                                            ByVal aHloc As String,
                                            ByRef aMoisXcoord As Double,
                                            ByRef aMoisYcoord As Double,
                                            ByRef aMoisElev As Double) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspRawCoordsExistError

        'Check the coordinate and elevation status of a hole in the MOIS raw prospect data.

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ProspRawDynaset As OraDynaset
        Dim RecordCount As Integer
        Dim ThisXcoord As Double
        Dim ThisYcoord As Double
        Dim ThisElev As Single

        gGetProspRawCoordsExist = 0
        aMoisXcoord = 0
        aMoisYcoord = 0
        aMoisElev = 0

        'Will return:
        '1 = Hole exists -- Coordinates & elevation exist (both X & Y must be there!)
        '2 = Hole exists -- No coordinates or elevation
        '3 = Hole exists -- Coordinates only exist (both X & Y must be there!)
        '4 = Hole exists -- Elevation only exists
        '5 = Hole does not exist
        '0 = Error

        'PROCEDURE get_prosp_raw_coord
        'pTownship           IN     NUMBER,
        'pRange              IN     NUMBER,
        'pSection            IN     NUMBER,
        'pHoleLocation       IN     VARCHAR2,
        'pProspDate          IN     DATE,
        'pResult             IN OUT c_prosprawbase)

        params = gDBParams

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHloc, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", DBNull.Value, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_coord(" &
                  ":pTownship, :pRange, :pSection, :pHoleLocation, " &
                  ":pProspDate, :pResult);end;", ORASQL_FAILEXEC)

        ProspRawDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = ProspRawDynaset.RecordCount

        'Note:  This procedure will return all occurrences of this hole location in the raw
        '       prospect data (there may be redrills).  It will use the first occurrence of
        '       the hole that it finds in MOIS.

        If RecordCount = 0 Then
            gGetProspRawCoordsExist = 5     'Hole does not exist!
            aMoisXcoord = 0
            aMoisYcoord = 0
            aMoisElev = 0
        Else
            ProspRawDynaset.MoveFirst()
            Do While Not ProspRawDynaset.EOF
                ThisXcoord = ProspRawDynaset.Fields("x_coord").Value
                ThisYcoord = ProspRawDynaset.Fields("y_coord").Value
                ThisElev = ProspRawDynaset.Fields("elevation").Value

                'Will return:
                '1 = Hole exists -- Coordinates & elevation exist (both X & Y must be there!)
                '2 = Hole exists -- No coordinates or elevation
                '3 = Hole exists -- Coordinates only exist (both X & Y must be there!)
                '4 = Hole exists -- Elevation only exists
                '5 = Hole does not exist
                '0 = Error

                'If a coordinate or an elevation is less than zero in MOIS then this
                'counts as missing data.
                If ThisXcoord > 0 And ThisYcoord > 0 And ThisElev > 0 Then
                    gGetProspRawCoordsExist = 1     'Hole exists -- Coordinates & elevation exist
                    aMoisXcoord = ThisXcoord
                    aMoisYcoord = ThisYcoord
                    aMoisElev = ThisElev
                    Exit Do
                End If

                If ThisXcoord <= 0 And ThisYcoord <= 0 And ThisElev <= 0 Then
                    gGetProspRawCoordsExist = 2     'Hole exists -- No coordinates or elevation
                    aMoisXcoord = ThisXcoord
                    aMoisYcoord = ThisYcoord
                    aMoisElev = ThisElev
                    Exit Do
                End If

                If ThisXcoord > 0 And ThisYcoord > 0 And ThisElev <= 0 Then
                    gGetProspRawCoordsExist = 3     'Hole exists -- Coordinates only exist
                    aMoisXcoord = ThisXcoord
                    aMoisYcoord = ThisYcoord
                    aMoisElev = ThisElev
                    Exit Do
                End If

                If ThisXcoord <= 0 And ThisYcoord <= 0 And ThisElev > 0 Then
                    gGetProspRawCoordsExist = 4     'Hole exists -- Elevation only exists
                    aMoisXcoord = ThisXcoord
                    aMoisYcoord = ThisYcoord
                    aMoisElev = ThisElev
                    Exit Do
                End If

                ProspRawDynaset.MoveNext()
            Loop
        End If

        ProspRawDynaset.Close()

        Exit Function

gGetProspRawCoordsExistError:
        gGetProspRawCoordsExist = 0

        MsgBox("Error with coordinate check." & vbCrLf &
            Err.Description,
            vbOKOnly + vbExclamation,
            "Coordinate Check Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        ProspRawDynaset.Close()
    End Function

    Public Function gGetIntervalsNew(ByVal aSection As Integer,
                                     ByVal aTownship As Integer,
                                     ByVal aRange As Integer,
                                     ByVal aHoleLocation As String,
                                     ByVal aSplitNum As Integer,
                                     ByVal aProspDate As Date,
                                     ByRef aAllSplits() As gHoleIntervalType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetIntervalsNewError

        Dim SampleDynaset As OraDynaset
        Dim SplitCount As Integer
        Dim RowIdx As Integer
        Dim RecordCount As Integer
        Dim SampsOk As Boolean

        'This is what type gHoleIntervalType is:
        'Public Type gHoleIntervalType
        '    TosDepth As Single
        '    BosDepth As Single
        '    SampNum As String
        '    DrillDate As Date
        '    Split As Integer
        'End Type

        SampsOk = gGetDrillHoleSamplesNew(aSection,
                                          aTownship,
                                          aRange,
                                          aHoleLocation,
                                          CStr(aProspDate),
                                          SampleDynaset)

        If SampsOk = False Then
            gGetIntervalsNew = False
            Exit Function
        End If

        RecordCount = SampleDynaset.RecordCount

        If RecordCount = 0 Then
            gGetIntervalsNew = False
            Exit Function
        Else
            gGetIntervalsNew = True
        End If

        SplitCount = 0

        ReDim aAllSplits(RecordCount)

        For RowIdx = 1 To RecordCount
            aAllSplits(RowIdx).TosDepth = 0
            aAllSplits(RowIdx).BosDepth = 0
            aAllSplits(RowIdx).SampNum = ""
            aAllSplits(RowIdx).DrillDate = #12/31/8888#
            aAllSplits(RowIdx).Split = 0
        Next RowIdx

        SampleDynaset.MoveFirst()

        Do While Not SampleDynaset.EOF
            SplitCount = SplitCount + 1
            aAllSplits(SplitCount).TosDepth = SampleDynaset.Fields("split_depth_top").Value
            aAllSplits(SplitCount).BosDepth = SampleDynaset.Fields("split_depth_bot").Value
            aAllSplits(SplitCount).SampNum = SampleDynaset.Fields("sample_id").Value
            aAllSplits(SplitCount).DrillDate = SampleDynaset.Fields("prosp_date").Value
            aAllSplits(SplitCount).Split = SampleDynaset.Fields("split_number").Value

            SampleDynaset.MoveNext()
        Loop

        SampleDynaset.Close()

        Exit Function

gGetIntervalsNewError:
        MsgBox("Error getting all sample#'s for this hole." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "All Hole Sample#'s Access Error")

        On Error Resume Next
        gGetIntervalsNew = False
        On Error Resume Next
        SampleDynaset.Close()
    End Function

    Public Function gGetHoleDatesInMois(ByVal aTwp As Integer,
                                        ByVal aRge As Integer,
                                        ByVal aSec As Integer,
                                        ByVal aHloc As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetHoleDatesInMoisError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim HoleDynaset As OraDynaset
        Dim RecordCount As Integer
        Dim ThisProspDate As String
        Dim ThisProspStandard As String
        Dim DateCnt As Integer
        Dim ThisMineName As String
        Dim ProcessIdx As Integer
        Dim MaxProcess As Integer
        Dim AlphaHoleLoc As String
        Dim TargHole As String

        'The hole location that is passed into this function will be a numeric
        'hole location.  We will have to check the numeric hole location and
        'its alpha-numeric translation if there is one.

        AlphaHoleLoc = gGetHoleLoc2(aHloc, "Char")

        If AlphaHoleLoc <> "???" Then
            MaxProcess = 2
        Else
            MaxProcess = 1
        End If

        DateCnt = 0

        For ProcessIdx = 1 To MaxProcess
            If ProcessIdx = 1 Then
                TargHole = aHloc
            Else
                TargHole = AlphaHoleLoc
            End If

            params = gDBParams

            params.Add("pTownship", aTwp, ORAPARM_INPUT)
            params("pTownship").serverType = ORATYPE_NUMBER

            params.Add("pRange", aRge, ORAPARM_INPUT)
            params("pRange").serverType = ORATYPE_NUMBER

            params.Add("pSection", aSec, ORAPARM_INPUT)
            params("pSection").serverType = ORATYPE_NUMBER

            params.Add("pHoleLocation", TargHole, ORAPARM_INPUT)
            params("pHoleLocation").serverType = ORATYPE_VARCHAR2

            params.Add("pResult", 0, ORAPARM_OUTPUT)
            params("pResult").serverType = ORATYPE_CURSOR

            'PROCEDURE get_hole_prospect_comp_base
            'pTownship      IN     NUMBER,
            'pRange         IN     NUMBER,
            'pSection       IN     NUMBER,
            'pHoleLocation  IN     VARCHAR2,
            'pResult        IN OUT c_composite)
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect2.get_hole_prospect_comp_base(" &
                      ":pTownship, :pRange, :pSection, :pHoleLocation, " &
                      ":pResult);end;", ORASQL_FAILEXEC)

            HoleDynaset = params("pResult").Value
            ClearParams(params)

            RecordCount = HoleDynaset.RecordCount

            If RecordCount <> 0 Then
                HoleDynaset.MoveFirst()
                Do While Not HoleDynaset.EOF
                    ThisProspDate = HoleDynaset.Fields("drill_cdate").Value
                    If Trim(ThisProspDate) = "" Then
                        ThisProspDate = "??/??/????"
                    End If
                    ThisProspStandard = HoleDynaset.Fields("prosp_standard").Value
                    ThisMineName = HoleDynaset.Fields("mine_name").Value
                    DateCnt = DateCnt + 1
                    If DateCnt = 1 Then
                        gGetHoleDatesInMois = "Prosp date = " & ThisProspDate & "   " &
                                              ThisMineName & "   " & ThisProspStandard
                    Else

                        gGetHoleDatesInMois = gGetHoleDatesInMois & vbCrLf &
                                              "Prosp date = " & ThisProspDate & "   " &
                                              ThisMineName & "   " & ThisProspStandard
                    End If
                    HoleDynaset.MoveNext()
                Loop
            End If
        Next ProcessIdx

        If DateCnt = 0 Then
            gGetHoleDatesInMois = "Hole not in MOIS!"
        End If

        HoleDynaset.Close()

        Exit Function

gGetHoleDatesInMoisError:
        MsgBox("Error accessing hole data." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Process Status")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        HoleDynaset.Close()
    End Function

    Public Function gUpdateRawProspect(ByRef aRawProspBase As gRawProspBaseType,
                                       ByVal aUpdateMode As String,
                                       ByVal aUserName As String,
                                       ByRef aSsSizeFrctnData As AxvaSpread,
                                       ByRef aSfcDataSprd() As gRawProspSfcSprdType,
                                       ByVal aUseSpread As Boolean) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gUpdateRawProspectError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ErrString As String
        Dim DateVal As Object
        Dim NumVal As Object
        Dim DryDensity As Single

        params = gDBParams

        With aRawProspBase
            'May need to assign some dates and stuff.
            If aUpdateMode = "Adding" Then
                .WhoModifiedHole = StrConv(aUserName, vbUpperCase)
                .WhenModifiedHole = Format(Now, "MM/dd/yyyy hh:mm tt")

                .WhoModifiedSplit = StrConv(aUserName, vbUpperCase)
                .WhenModifiedSplit = Format(Now, "MM/dd/yyyy hh:mm tt")

                .WhoMetLab = StrConv(aUserName, vbUpperCase)
                .DateMetLab = Format(Now, "MM/dd/yyyy hh:mm tt")

                .WhoChemLab = ""
                .DateChemLab = ""

                'Need to get the alpha-numeric hole location (if there is
                'one) for the numeric hole location.  For example:
                '2062 --> J15.
                .HoleLocationChar = gGetHoleLoc2(.HoleLocation, "Char")
            End If

            If aUpdateMode = "Editing" Then
                .WhoModifiedHole = StrConv(aUserName, vbUpperCase)
                .WhenModifiedHole = Format(Now, "MM/dd/yyyy hh:mm tt")

                .WhoModifiedSplit = StrConv(aUserName, vbUpperCase)
                .WhenModifiedSplit = Format(Now, "MM/dd/yyyy hh:mm tt")

                .WhoMetLab = StrConv(aUserName, vbUpperCase)
                .DateMetLab = Format(Now, "MM/dd/yyyy hh:mm tt")
            End If
        End With

        With aRawProspBase
            '1
            params.Add("pSampleId", .SampleId, ORAPARM_INPUT)
            params("pSampleId").serverType = ORATYPE_VARCHAR2

            '2
            params.Add("pTownShip", .Township, ORAPARM_INPUT)
            params("pTownShip").serverType = ORATYPE_NUMBER

            '3
            params.Add("pRange", .Range, ORAPARM_INPUT)
            params("pRange").serverType = ORATYPE_NUMBER

            '4
            params.Add("pSection", .Section, ORAPARM_INPUT)
            params("pSection").serverType = ORATYPE_NUMBER

            '5
            params.Add("pHoleLocation", .HoleLocation, ORAPARM_INPUT)
            params("pHoleLocation").serverType = ORATYPE_VARCHAR2

            '6
            params.Add("pForty", .Forty, ORAPARM_INPUT)
            params("pForty").serverType = ORATYPE_NUMBER

            '7
            params.Add("pState", .State, ORAPARM_INPUT)
            params("pState").serverType = ORATYPE_NUMBER

            '8
            params.Add("pQuadrant", .Quadrant, ORAPARM_INPUT)
            params("pQuadrant").serverType = ORATYPE_NUMBER

            '9
            params.Add("pMineName", .MineName, ORAPARM_INPUT)
            params("pMineName").serverType = ORATYPE_VARCHAR2

            '10
            params.Add("pExpDrill", .ExpDrill, ORAPARM_INPUT)
            params("pExpDrill").serverType = ORATYPE_NUMBER

            '11
            params.Add("pSplitTotalNum", .SplitTotalNum, ORAPARM_INPUT)
            params("pSplitTotalNum").serverType = ORATYPE_NUMBER

            '12
            params.Add("pXcoord", .Xcoord, ORAPARM_INPUT)
            params("pXcoord").serverType = ORATYPE_NUMBER

            '13
            params.Add("pYcoord", .Ycoord, ORAPARM_INPUT)
            params("pYcoord").serverType = ORATYPE_NUMBER

            '14
            params.Add("pFtlDepth", .FtlDepth, ORAPARM_INPUT)
            params("pFtlDepth").serverType = ORATYPE_NUMBER

            '15
            params.Add("pOvbCored", .OvbCored, ORAPARM_INPUT)
            params("pOvbCored").serverType = ORATYPE_NUMBER

            '16
            params.Add("pOwnership", .Ownership, ORAPARM_INPUT)
            params("pOwnership").serverType = ORATYPE_VARCHAR2

            '17
            'Have checked this date already -- it will not be null.
            params.Add("pProspDate", CDate(.ProspDate), ORAPARM_INPUT)
            params("pProspDate").serverType = ORATYPE_DATE

            '18
            params.Add("pMinedStatus", .MinedStatus, ORAPARM_INPUT)
            params("pMinedStatus").serverType = ORATYPE_NUMBER

            '19
            params.Add("pElevation", .Elevation, ORAPARM_INPUT)
            params("pElevation").serverType = ORATYPE_NUMBER

            '20
            params.Add("pTotDepth", .TotDepth, ORAPARM_INPUT)
            params("pTotDepth").serverType = ORATYPE_NUMBER

            '21
            params.Add("pAoi", .Aoi, ORAPARM_INPUT)
            params("pAoi").serverType = ORATYPE_NUMBER

            '22
            params.Add("pCoordSurveyed", .CoordSurveyed, ORAPARM_INPUT)
            params("pCoordSurveyed").serverType = ORATYPE_NUMBER

            '23
            params.Add("pLongComment", .HoleComment, ORAPARM_INPUT)
            params("pLongComment").serverType = ORATYPE_VARCHAR2

            '24
            params.Add("pHoleLocationChar", .HoleLocationChar, ORAPARM_INPUT)
            params("pHoleLocationChar").serverType = ORATYPE_VARCHAR2

            '25
            params.Add("pWhoModified", .WhoModifiedHole, ORAPARM_INPUT)
            params("pWhoModified").serverType = ORATYPE_VARCHAR2

            '26
            'This date may be null.
            If Not IsDBNull(.WhenModifiedHole) And
                Trim(.WhenModifiedHole) <> "" Then
                DateVal = CDate(.WhenModifiedHole)
            Else
                DateVal = DBNull.Value
            End If
            params.Add("pWhenModified", DateVal, ORAPARM_INPUT)
            params("pWhenModified").serverType = ORATYPE_DATE

            '27
            'This date may be null.
            If Not IsDBNull(.LogDate) And
                Trim(.LogDate) <> "" Then
                DateVal = CDate(.LogDate)
            Else
                DateVal = DBNull.Value
            End If
            params.Add("pLogDate", DateVal, ORAPARM_INPUT)
            params("pLogDate").serverType = ORATYPE_DATE

            '28
            params.Add("pReleased", .Released, ORAPARM_INPUT)
            params("pReleased").serverType = ORATYPE_NUMBER

            '29
            params.Add("pReDrilled", .Redrilled, ORAPARM_INPUT)
            params("pReDrilled").serverType = ORATYPE_NUMBER

            '30
            'This date may be null.
            If Not IsDBNull(.RedrillDate) And
                Trim(.RedrillDate) <> "" Then
                DateVal = CDate(.RedrillDate)
            Else
                DateVal = DBNull.Value
            End If
            params.Add("pRedrillDate", DateVal, ORAPARM_INPUT)
            params("pRedrillDate").serverType = ORATYPE_DATE

            '31
            params.Add("pUseForReduction", .UseForReduction, ORAPARM_INPUT)
            params("pUseForReduction").serverType = ORATYPE_NUMBER

            '32
            params.Add("pSplitNumber", .SplitNumber, ORAPARM_INPUT)
            params("pSplitNumber").serverType = ORATYPE_NUMBER

            '33
            params.Add("pBarren", .Barren, ORAPARM_INPUT)
            params("pBarren").serverType = ORATYPE_NUMBER

            '34
            params.Add("pSplitFtlBottom", .SplitFtlBottom, ORAPARM_INPUT)
            params("pSplitFtlBottom").serverType = ORATYPE_NUMBER

            '35
            params.Add("pMtxTotWetWt", .MtxTotWetWt, ORAPARM_INPUT)
            params("pMtxTotWetWt").serverType = ORATYPE_NUMBER

            '36
            params.Add("pMtxMoistWetWt", .MtxMoistWetWt, ORAPARM_INPUT)
            params("pMtxMoistWetWt").serverType = ORATYPE_NUMBER

            '37
            params.Add("pMtxMoistDryWt", .MtxMoistDryWt, ORAPARM_INPUT)
            params("pMtxMoistDryWt").serverType = ORATYPE_NUMBER

            '38
            params.Add("pMtxMoistTareWt", .MtxMoistTareWt, ORAPARM_INPUT)
            params("pMtxMoistTareWt").serverType = ORATYPE_NUMBER

            '39
            params.Add("pFdTotWetWt", .FdTotWetWt, ORAPARM_INPUT)
            params("pFdTotWetWt").serverType = ORATYPE_NUMBER

            '40
            params.Add("pFdTotWetWtMsr", .FdTotWetWtMsr, ORAPARM_INPUT)
            params("pFdTotWetWtMsr").serverType = ORATYPE_NUMBER

            '41
            params.Add("pFdMoistWetWt", .FdMoistWetWt, ORAPARM_INPUT)
            params("pFdMoistWetWt").serverType = ORATYPE_NUMBER

            '42
            params.Add("pFdMoistDryWt", .FdMoistDryWt, ORAPARM_INPUT)
            params("pFdMoistDryWt").serverType = ORATYPE_NUMBER

            '43
            params.Add("pFdMoistTareWt", .FdMoistTareWt, ORAPARM_INPUT)
            params("pFdMoistTareWt").serverType = ORATYPE_NUMBER

            '44
            params.Add("pFdScrnSampWt", .FdScrnSampWt, ORAPARM_INPUT)
            params("pFdScrnSampWt").serverType = ORATYPE_NUMBER

            '45
            params.Add("pDensCylSize", .DensCylSize, ORAPARM_INPUT)
            params("pDensCylSize").serverType = ORATYPE_NUMBER

            '46
            params.Add("pDensCylWetWt", .DensCylWetWt, ORAPARM_INPUT)
            params("pDensCylWetWt").serverType = ORATYPE_NUMBER

            '47
            params.Add("pDensCylH2oWt", .DensCylH2oWt, ORAPARM_INPUT)
            params("pDensCylH2oWt").serverType = ORATYPE_NUMBER

            '48
            DryDensity = .DryDensity
            If .DryDensityOverride > 0 Or .DryDensityOverride = -1 Then
                DryDensity = .DryDensityOverride
            End If
            params.Add("pDryDensity", DryDensity, ORAPARM_INPUT)
            params("pDryDensity").serverType = ORATYPE_NUMBER

            '49
            params.Add("pFlotWetWt", .FlotFdWetWt, ORAPARM_INPUT)
            params("pFlotWetWt").serverType = ORATYPE_NUMBER

            '50
            params.Add("pMtxProcWetWt", .MtxProcWetWt, ORAPARM_INPUT)
            params("pMtxProcWetWt").serverType = ORATYPE_NUMBER

            '51
            params.Add("pExpExcessWt", .ExpExcessWt, ORAPARM_INPUT)
            params("pExpExcessWt").serverType = ORATYPE_NUMBER

            '52
            params.Add("pMtxColor", .MtxColor, ORAPARM_INPUT)
            params("pMtxColor").serverType = ORATYPE_VARCHAR2

            '53
            params.Add("pDegConsol", .DegConsol, ORAPARM_INPUT)
            params("pDegConsol").serverType = ORATYPE_VARCHAR2

            '54
            params.Add("pDigChar", .DigChar, ORAPARM_INPUT)
            params("pDigChar").serverType = ORATYPE_VARCHAR2

            '55
            params.Add("pPumpChar", .PumpChar, ORAPARM_INPUT)
            params("pPumpChar").serverType = ORATYPE_VARCHAR2

            '56
            params.Add("pLithology", .Lithology, ORAPARM_INPUT)
            params("pLithology").serverType = ORATYPE_VARCHAR2

            '57
            params.Add("pPhosphColor", .PhosphColor, ORAPARM_INPUT)
            params("pPhosphColor").serverType = ORATYPE_VARCHAR2

            '58
            params.Add("pPhysMineable", .PhysMineable, ORAPARM_INPUT)
            params("pPhysMineable").serverType = ORATYPE_NUMBER

            '59
            params.Add("pClaySettChar", .ClaySettChar, ORAPARM_INPUT)
            params("pClaySettChar").serverType = ORATYPE_VARCHAR2

            '60
            params.Add("pFdScrnSampWtComp", .FdScrnSampWtComp, ORAPARM_INPUT)
            params("pFdScrnSampWtComp").serverType = ORATYPE_NUMBER

            '61
            params.Add("pRecordLocked", .RecordLocked, ORAPARM_INPUT)
            params("pRecordLocked").serverType = ORATYPE_NUMBER

            '62
            'This date may be null.
            If Not IsDBNull(.DateChemLab) And
                Trim(.DateChemLab) <> "" Then
                DateVal = CDate(.DateChemLab)
            Else
                DateVal = DBNull.Value
            End If
            params.Add("pDateChemLab", DateVal, ORAPARM_INPUT)
            params("pDateChemLab").serverType = ORATYPE_DATE

            '63
            params.Add("pWhoChemLab", .WhoChemLab, ORAPARM_INPUT)
            params("pWhoChemLab").serverType = ORATYPE_VARCHAR2

            '64
            params.Add("pRerunStatus", .RerunStatus, ORAPARM_INPUT)
            params("pRerunStatus").serverType = ORATYPE_NUMBER

            '65
            'This date may be null.
            If Not IsDBNull(.DateRerun) And
                Trim(.DateRerun) <> "" Then
                DateVal = CDate(.DateRerun)
            Else
                DateVal = DBNull.Value
            End If
            params.Add("pDateRerun", DateVal, ORAPARM_INPUT)
            params("pDateRerun").serverType = ORATYPE_DATE

            '66
            params.Add("pMetLabComment", .MetLabComment, ORAPARM_INPUT)
            params("pMetLabComment").serverType = ORATYPE_VARCHAR2

            '67
            params.Add("pChemLabComment", .ChemLabComment, ORAPARM_INPUT)
            params("pChemLabComment").serverType = ORATYPE_VARCHAR2

            '68
            'This date may be null.
            If Not IsDBNull(.DateMetLab) And
                Trim(.DateMetLab) <> "" Then
                DateVal = CDate(.DateMetLab)
            Else
                DateVal = DBNull.Value
            End If
            params.Add("pDateMetLab", DateVal, ORAPARM_INPUT)
            params("pDateMetLab").serverType = ORATYPE_DATE

            '69
            params.Add("pWhoMetLab", .WhoMetLab, ORAPARM_INPUT)
            params("pWhoMetLab").serverType = ORATYPE_VARCHAR2

            '70
            params.Add("pSplitDepthTop", .SplitDepthTop, ORAPARM_INPUT)
            params("pSplitDepthTop").serverType = ORATYPE_NUMBER

            '71
            params.Add("pSplitDepthBot", .SplitDepthBot, ORAPARM_INPUT)
            params("pSplitDepthBot").serverType = ORATYPE_NUMBER

            '72
            params.Add("pSplitThck", .SplitThck, ORAPARM_INPUT)
            params("pSplitThck").serverType = ORATYPE_NUMBER

            '73
            'This date may be null.
            If Not IsDBNull(.WashDate) And
                Trim(.WashDate) <> "" Then
                DateVal = CDate(.WashDate)
            Else
                DateVal = DBNull.Value
            End If
            params.Add("pWashDate", DateVal, ORAPARM_INPUT)
            params("pWashDate").serverType = ORATYPE_DATE

            '74
            params.Add("pWhoModifiedSplit", .WhoModifiedSplit, ORAPARM_INPUT)
            params("pWhoModifiedSplit").serverType = ORATYPE_VARCHAR2

            '75
            'This date may be null.
            If Not IsDBNull(.WhenModifiedSplit) And
                Trim(.WhenModifiedSplit) <> "" Then
                DateVal = CDate(.WhenModifiedSplit)
            Else
                DateVal = DBNull.Value
            End If
            params.Add("pWhenModifiedSplit", DateVal, ORAPARM_INPUT)
            params("pWhenModifiedSplit").serverType = ORATYPE_DATE

            '76
            params.Add("pOrigData", .OrigData, ORAPARM_INPUT)
            params("pOrigData").serverType = ORATYPE_NUMBER

            '77
            params.Add("pCounty", .County, ORAPARM_INPUT)
            params("pCounty").serverType = ORATYPE_VARCHAR2

            '78
            params.Add("pBankCode", .BankCode, ORAPARM_INPUT)
            params("pBankCode").serverType = ORATYPE_VARCHAR2

            '-----

            '79
            If .HoleMinable = -1 Then
                NumVal = DBNull.Value
            Else
                NumVal = .HoleMinable   '0 or 1
            End If
            params.Add("pHoleMinable", NumVal, ORAPARM_INPUT)
            params("pHoleMinable").serverType = ORATYPE_NUMBER

            '80
            'This date may be null.
            If Not IsDBNull(.HoleMinableWhen) And
                Trim(.HoleMinableWhen) <> "" Then
                DateVal = CDate(.HoleMinableWhen)
            Else
                DateVal = DBNull.Value
            End If
            params.Add("pHoleMinableWhen", DateVal, ORAPARM_INPUT)
            params("pHoleMinableWhen").serverType = ORATYPE_DATE

            '81
            params.Add("pHoleMinableWho", Trim(.HoleMinableWho), ORAPARM_INPUT)
            params("pHoleMinableWho").serverType = ORATYPE_VARCHAR2

            '82
            params.Add("pSpecAreaName", Trim(.SpecAreaName), ORAPARM_INPUT)
            params("pSpecAreaName").serverType = ORATYPE_VARCHAR2

            '83
            params.Add("pManufacturedData", .ManufacturedData, ORAPARM_INPUT)
            params("pManufacturedData").serverType = ORATYPE_NUMBER

            '84
            If .SplitMinable = -1 Then
                NumVal = DBNull.Value
            Else
                NumVal = .SplitMinable   '0 or 1
            End If
            params.Add("pSplitMinable", NumVal, ORAPARM_INPUT)
            params("pSplitMinable").serverType = ORATYPE_NUMBER

            '85
            'This date may be null.
            If Not IsDBNull(.SplitMinableWhen) And
                Trim(.SplitMinableWhen) <> "" Then
                DateVal = CDate(.SplitMinableWhen)
            Else
                DateVal = DBNull.Value
            End If
            params.Add("pSplitMinableWhen", DateVal, ORAPARM_INPUT)
            params("pSplitMinableWhen").serverType = ORATYPE_DATE

            '86
            params.Add("pSplitMinableWho", Trim(.SplitMinableWho), ORAPARM_INPUT)
            params("pSplitMinableWho").serverType = ORATYPE_VARCHAR2

            '87
            params.Add("pSampleIdCargill", Trim(.SampleIdCargill), ORAPARM_INPUT)
            params("pSampleIdCargill").serverType = ORATYPE_VARCHAR2

            '88
            params.Add("pBedCode", Trim(.BedCode), ORAPARM_INPUT)
            params("pBedCode").serverType = ORATYPE_VARCHAR2

            '89
            params.Add("pClaySettlingLvl", .ClaySettlingLvl, ORAPARM_INPUT)
            params("pClaySettlingLvl").serverType = ORATYPE_NUMBER

            '90
            params.Add("pPbClayPct", .PbClayPct, ORAPARM_INPUT)
            params("pPbClayPct").serverType = ORATYPE_NUMBER

            '-----

            '91
            params.Add("pMtxMoistWetWt2", .MtxMoistWetWt2, ORAPARM_INPUT)
            params("pMtxMoistWetWt2").serverType = ORATYPE_NUMBER

            '92
            params.Add("pMtxMoistDryWt2", .MtxMoistDryWt2, ORAPARM_INPUT)
            params("pMtxMoistDryWt2").serverType = ORATYPE_NUMBER

            '93
            params.Add("pMtxMoistTareWt2", .MtxMoistTareWt2, ORAPARM_INPUT)
            params("pMtxMoistTareWt2").serverType = ORATYPE_NUMBER

            '-----

            '94
            params.Add("pQaQcHole", .QaQcHole, ORAPARM_INPUT)
            params("pQaQcHole").serverType = ORATYPE_NUMBER

            '-----
            '95
            params.Add("pHardpanFrom", .HardpanFrom, ORAPARM_INPUT)
            params("pHardpanFrom").serverType = ORATYPE_NUMBER

            '96
            params.Add("pHardpanTo", .HardpanTo, ORAPARM_INPUT)
            params("pHardpanTo").serverType = ORATYPE_NUMBER

            '97
            params.Add("pHardpanCode", .HardpanCode, ORAPARM_INPUT)
            params("pHardpanCode").serverType = ORATYPE_VARCHAR2

            '98
            params.Add("pHardpanThck", .HardpanThck, ORAPARM_INPUT)
            params("pHardpanThck").serverType = ORATYPE_NUMBER

            '99
            params.Add("pResult", 0, ORAPARM_OUTPUT)
            params("pResult").serverType = ORATYPE_NUMBER
        End With

        'Procedure update_raw_prosp_split
        'pSampleId              IN     VARCHAR2,   --1
        'pTownship              IN     NUMBER,     --2
        'pRange                 IN     NUMBER,     --3
        'pSection               IN     NUMBER,     --4
        'pHoleLocation          IN     VARCHAR2,   --5
        'pForty                 IN     NUMBER,     --6
        'pState                 IN     VARCHAR2,   --7
        'pQuadrant              IN     NUMBER,     --8
        'pMineName              IN     VARCHAR2,   --9
        'pExpDrill              IN     NUMBER,     --10
        'pSplitTotalNum         IN     NUMBER,     --11
        'pXCoord                IN     NUMBER,     --12
        'pYCoord                IN     NUMBER,     --13
        'pFtlDepth              IN     NUMBER,     --14
        'pOvbCored              IN     NUMBER,     --15
        'pOwnership             IN     VARCHAR2,   --16
        'pProspDate             IN     DATE,       --17
        'pMinedStatus           IN     NUMBER,     --18
        'pElevation             IN     NUMBER,     --19
        'pTotDepth              IN     NUMBER,     --20
        'pAoi                   IN     NUMBER,     --21
        'pCoordSurveyed         IN     NUMBER,     --22
        'pLongComment           IN     VARCHAR2,   --23
        'pHoleLocationChar      IN     VARCHAR2,   --24
        'pWhoModified           IN     VARCHAR2,   --25
        'pWhenModified          IN     DATE,       --26
        'pLogDate               IN     DATE,       --27
        'pReleased              IN     NUMBER,     --28
        'pRedrilled             IN     NUMBER,     --29
        'pRedrillDate           IN     DATE,       --30
        'pUseForReduction       IN     NUMBER,     --31
        '--
        'pSplitNumber           IN     NUMBER,     --32
        'pBarren                IN     NUMBER,     --33
        'pSplitFtlBottom        IN     NUMBER,     --34
        'pMtxTotWetWt           IN     NUMBER,     --35
        'pMtxMoistWetWt         IN     NUMBER,     --36
        'pMtxMoistDryWt         IN     NUMBER,     --37
        'pMtxMoistTareWt        IN     NUMBER,     --38
        'pFdTotWetWt            IN     NUMBER,     --39
        'pFdTotWetWtMsr         IN     NUMBER,     --40
        'pFdMoistWetWt          IN     NUMBER,     --41
        'pFdMoistDryWt          IN     NUMBER,     --42
        'pFdMoistTareWt         IN     NUMBER,     --43
        'pFdScrnSampWt          IN     NUMBER,     --44
        'pDensCylSize           IN     NUMBER,     --45
        'pDensCylWetWt          IN     NUMBER,     --46
        'pDensCylH2oWt          IN     NUMBER,     --47
        'pDryDensity            IN     NUMBER,     --48
        'pFlotWetWt             IN     NUMBER,     --49
        'pMtxProcWetWt          IN     NUMBER,     --50
        'pExpExcessWt           IN     NUMBER,     --51
        'pMtxColor              IN     VARCHAR2,   --52
        'pDegConsol             IN     VARCHAR2,   --53
        'pDigChar               IN     VARCHAR2,   --54
        'pPumpChar              IN     VARCHAR2,   --55
        'pLithology             IN     VARCHAR2,   --56
        'pPhosphColor           IN     VARCHAR2,   --57
        'pPhysMineable          IN     NUMBER,     --58
        'pClaySettChar          IN     VARCHAR2,   --59
        'pFdScrnSampWtComp      IN     NUMBER,     --60
        'pRecordLocked          IN     NUMBER,     --61
        'pDateChemLab           IN     DATE,       --62
        'pWhoChemLab            IN     VARCHAR2,   --63
        'pRerunStatus           IN     NUMBER,     --64
        'pDateRerun             IN     DATE,       --65
        'pMetLabComment         IN     VARCHAR2,   --66
        'pChemLabComment        IN     VARCHAR2,   --67
        'pDateMetLab            IN     DATE,       --68
        'pWhoMetLab             IN     VARCHAR2,   --69
        'pSplitDepthTop         IN     NUMBER,     --70
        'pSplitDepthBot         IN     NUMBER,     --71
        'pSplitThck             IN     NUMBER,     --72
        'pWashDate              IN     DATE,       --73
        'pWhoModifiedSplit      IN     VARCHAR2,   --74
        'pWhenModifiedSplit     IN     DATE,       --75
        'pOrigData              IN     NUMBER,     --76
        '--
        'pCounty                IN     VARCHAR2,   --77
        'pBankCode              IN     VARCHAR2,   --78
        '--
        'pHoleMinable           IN     NUMBER,     --79
        'pHoleMinableWhen       IN     DATE,       --80
        'pHoleMinableWho        IN     VARCHAR2,   --81
        'pSpecAreaName          IN     VARCHAR2,   --82
        'pManufacturedData      IN     NUMBER,     --83
        '--
        'pSplitMinable          IN     NUMBER,     --84
        'pSplitMinableWhen      IN     DATE,       --85
        'pSplitMinableWho       IN     VARCHAR2,   --86
        'pSampleIdCargill       IN     VARCHAR2,   --87
        'pBedCode               IN     VARCHAR2,   --88
        '--
        'pClaySettlingLvl       IN     NUMBER,     --89
        'pPbClayPct             IN     NUMBER,     --90
        '--
        'pMtxMoistWetWt2        IN     NUMBER,     --91
        'pMtxMoistDryWt2        IN     NUMBER,     --92
        'pMtxMoistTareWt2       IN     NUMBER,     --93
        '--
        'pQaQcHoleData          IN     NUMBER,     --94
        '--
        'pHardpanFrom           IN     NUMBER,     --95
        'pHardpanTo             IN     NUMBER,     --96
        'pHardpanCode           IN     VARCHAR2,   --97
        'pHardpanThck           IN     NUMBER,     --98
        '--
        'pResult                IN OUT NUMBER)     --99

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.update_raw_prosp_split2(:pSampleId," +
                  ":pTownship, :pRange, :pSection, :pHoleLocation, :pForty, :pState, :pQuadrant, :pMineName," +
                  ":pExpDrill, :pSplitTotalNum, :pXcoord, :pYcoord, :pFtlDepth, :pOvbCored, :pOwnership," +
                  ":pProspDate, :pMinedStatus, :pElevation, :pTotDepth, :pAoi, :pCoordSurveyed, :pLongComment," +
                  ":pHoleLocationChar, :pWhoModified, :pWhenModified, :pLogDate, :pReleased, :pRedrilled," +
                  ":pRedrillDate, :pUseForReduction, :pSplitNumber, :pBarren, :pSplitFtlBottom, :pMtxTotWetWt," +
                  ":pMtxMoistWetWt, :pMtxMoistDryWt, :pMtxMoistTareWt, :pFdTotWetWt, :pFdTotWetWtMsr, :pFdMoistWetWt," +
                  ":pFdMoistDryWt, :pFdMoistTareWt, :pFdScrnSampWt, :pDensCylSize, :pDensCylWetWt, :pDensCylH2oWt," +
                  ":pDryDensity, :pFlotWetWt, :pMtxProcWetWt, :pExpExcessWt, :pMtxColor, :pDegConsol, :pDigChar," +
                  ":pPumpChar, :pLithology, :pPhosphColor, :pPhysMineable, :pClaySettChar, :pFdScrnSampWtComp," +
                  ":pRecordLocked, :pDateChemLab, :pWhoChemLab, :pRerunStatus, :pDateRerun, :pMetLabComment," +
                  ":pChemLabComment, :pDateMetLab, :pWhoMetLab, :pSplitDepthTop, :pSplitDepthBot," +
                  ":pSplitThck, :pWashDate, :pWhoModifiedSplit, :pWhenModifiedSplit, :pOrigData, :pCounty, :pBankCode," +
                  ":pHoleMinable, :pHoleMinableWhen, :pHoleMinableWho, :pSpecAreaName, :pManufacturedData, " +
                  ":pSplitMinable, :pSplitMinableWhen, :pSplitMinableWho, :pSampleIdCargill, :pBedCode, " +
                  ":pClaySettlingLvl, :pPbClayPct, :pMtxMoistWetWt2, :pMtxMoistDryWt2, :pMtxMoistTareWt2, :pQaQcHole, " +
                  ":pHardpanFrom, :pHardpanTo, :pHardpanCode, :pHardpanThck, :pResult);end;", ORASQL_FAILEXEC)

        ClearParams(params)

        'Update the size fraction code stuff for this sample.
        If aUseSpread = True Then
            gUpdateProspRawSizeFrctnSprd(aRawProspBase,
                                         aUpdateMode,
                                         aUserName,
                                         aSsSizeFrctnData)
        Else
            gUpdateProspRawSizeFrctnArray(aRawProspBase,
                                          aUpdateMode,
                                          aUserName,
                                          aSfcDataSprd)
        End If

        gUpdateRawProspect = True

        Exit Function

gUpdateRawProspectError:

        With aRawProspBase
            ErrString = "Twp = " + Trim(Str(.Township)) + ", " +
                        "Rge = " + Trim(Str(.Range)) + ", " +
                        "Sec = " + Trim(Str(.Section)) + ", " +
                        "Hole location = " + .HoleLocation
        End With

        MsgBox("Oracle returned an error while attempting to add the data." + Str(Err.Number) + Chr(10) + Chr(10) +
               ErrString + Chr(10) + Chr(10) +
               Err.Description, vbExclamation, "Error Adding Raw Prospect Data")

        gUpdateRawProspect = False

        On Error Resume Next
        ClearParams(params)
    End Function

    Public Sub gUpdateProspRawSizeFrctnSprd(ByRef aProspRawBase As gRawProspBaseType,
                                            ByVal aUpdateMode As String,
                                            ByVal aUserName As String,
                                            ByRef aSsSizeFrctnData As AxvaSpread)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gUpdateProspRawSizeFrctnSprdError

        Dim SizeFrctnCode() As String
        Dim SizeFrctnWt() As Single
        Dim SizeFrctnWtMsr() As Single
        Dim SizeFrctnType() As String
        Dim Bpl() As Single
        Dim FeAl() As Single
        Dim Insol() As Single
        Dim CaO() As Single
        Dim MgO() As Single
        Dim Fe2O3() As Single
        Dim Al2O3() As Single
        Dim Cd() As Single
        Dim OrderNum() As Integer

        Dim ItemCount As Integer
        Dim InsertSQL As String
        Dim RowIdx As Integer
        Dim WhoModified As String
        Dim WhenModified As Date

        'Need to assign some dates and stuff.
        If aUpdateMode = "Adding" Then
            WhoModified = StrConv(aUserName, vbUpperCase)
            WhenModified = CDate(Format(Now, "MM/dd/yyyy hh:mm tt"))
        End If

        ItemCount = aSsSizeFrctnData.MaxRows

        'If data exists then redimension transfer arrays
        If ItemCount <> 0 Then
            ReDim SizeFrctnCode(ItemCount - 1)
            ReDim SizeFrctnWt(ItemCount - 1)
            ReDim SizeFrctnWtMsr(ItemCount - 1)
            ReDim SizeFrctnType(ItemCount - 1)
            ReDim Bpl(ItemCount - 1)
            ReDim FeAl(ItemCount - 1)
            ReDim Insol(ItemCount - 1)
            ReDim CaO(ItemCount - 1)
            ReDim MgO(ItemCount - 1)
            ReDim Fe2O3(ItemCount - 1)
            ReDim Al2O3(ItemCount - 1)
            ReDim Cd(ItemCount - 1)
            ReDim OrderNum(ItemCount - 1)
        Else
            'Nothing to update
            Exit Sub
        End If

        'Now place the data into the transfer arrays.
        'Non-numeric entries ie. this measure does not apply to this
        'size fraction code will be marked with -1's.
        ItemCount = 0
        With aSsSizeFrctnData
            For RowIdx = 1 To .MaxRows
                .Row = RowIdx
                .Col = 1   'Size fraction code
                SizeFrctnCode(ItemCount) = .Text

                .Col = 5    'Total size fraction weight -- measured
                If IsNumeric(.Text) Then
                    SizeFrctnType(ItemCount) = "T"
                    SizeFrctnWtMsr(ItemCount) = .Value
                End If

                .Col = 6    'Sample weight
                'Need to enter this value for both SizeFrctnWtMsr and
                'SizeFrctnWt since it is size fraction type = "P"
                If IsNumeric(.Text) Then
                    SizeFrctnType(ItemCount) = "P"
                    'Measured and Adjusted are the same for size fraction type = "P"
                    SizeFrctnWtMsr(ItemCount) = .Value  'Measured value
                    SizeFrctnWt(ItemCount) = .Value     'Adjusted value
                End If

                .Col = 7    'BPL
                If IsNumeric(.Text) Then
                    Bpl(ItemCount) = Val(.Text)
                Else
                    Bpl(ItemCount) = -1
                End If

                .Col = 8    'Insol
                If IsNumeric(.Text) Then
                    Insol(ItemCount) = Val(.Text)
                Else
                    Insol(ItemCount) = -1
                End If

                .Col = 9    'CaO
                If IsNumeric(.Text) Then
                    CaO(ItemCount) = Val(.Text)
                Else
                    CaO(ItemCount) = -1
                End If

                .Col = 10   'MgO
                If IsNumeric(.Text) Then
                    MgO(ItemCount) = Val(.Text)
                Else
                    MgO(ItemCount) = -1
                End If

                .Col = 11   'Fe2O3
                If IsNumeric(.Text) Then
                    Fe2O3(ItemCount) = Val(.Text)
                Else
                    Fe2O3(ItemCount) = -1
                End If

                .Col = 12   'Al2O3
                If IsNumeric(.Text) Then
                    Al2O3(ItemCount) = Val(.Text)
                Else
                    Al2O3(ItemCount) = -1
                End If

                .Col = 13   'Fe&Al
                If IsNumeric(.Text) Then
                    FeAl(ItemCount) = Val(.Text)
                Else
                    FeAl(ItemCount) = -1
                End If

                'Don't have a Cd measure for the user right now since
                'they don't measure this in the Chem lab -- will set to 0
                'Cd .Col = 14
                Cd(ItemCount) = -1

                .Col = 15    'Total size fraction weight -- adjusted
                'This really applies only to size fraction type = "T"
                'not size fraction type = "P"
                If SizeFrctnType(ItemCount) = "T" Then
                    If IsNumeric(.Text) Then
                        SizeFrctnWt(ItemCount) = .Value
                    Else
                        SizeFrctnWt(ItemCount) = -1
                    End If
                End If

                'Fix the size fraction type if necessary
                If SizeFrctnType(ItemCount) = "" Then
                    SizeFrctnType(ItemCount) = " "
                End If

                OrderNum(ItemCount) = RowIdx

                ItemCount = ItemCount + 1
            Next RowIdx
        End With

        'Procedure update_prosp_raw_size_frctn
        'pArraySize        IN     INTEGER,
        'pSampleId         IN     VARCHAR2,
        'pTownship         IN     NUMBER,
        'pRange            IN     NUMBER,
        'pSection          IN     NUMBER,
        'pHoleLocation     IN     VARCHAR2,
        'pProspDate        IN     DATE,
        'pSplitNumber      IN     NUMBER,
        'pWhoModified      IN     VARCHAR2,
        'pWhenModified     IN     DATE,
        'pSizeFrctnCode    IN     VCHAR2ARRAY3,
        'pBpl              IN     NUMBERARRAY,
        'pFeAl             IN     NUMBERARRAY,
        'pInsol            IN     NUMBERARRAY,
        'pCaO              IN     NUMBERARRAY,
        'pMgO              IN     NUMBERARRAY,
        'pFe2O3            IN     NUMBERARRAY,
        'pAl2O3            IN     NUMBERARRAY,
        'pCd               IN     NUMBERARRAY,
        'pSzeFrctnWt       IN     NUMBERARRAY,
        'pSzeFrctnWtMsr    IN     NUMBERARRAY,
        'pSzeFrctnType     IN     VCHAR2ARRAY1,
        'pOrderNum         IN     NUMBERARRAY,
        'pResult           IN OUT NUMBER)

        InsertSQL = "Begin mois.mois_raw_prospectnew.update_prosp_raw_size_frctn(" &
        "   :pArraySize, " &
        "   :pSampleId, " &
        "   :pTownship, " &
        "   :pRange, " &
        "   :pSection, " &
        "   :pHoleLocation, " &
        "   :pProspDate, " &
        "   :pSplitNumber, " &
        "   :pWhoModified, " &
        "   :pWhenModified, " &
        "   :pSizeFrctnCode, " &
        "   :pBpl, " &
        "   :pFeAl, " &
        "   :pInsol, " &
        "   :pCaO, " &
        "   :pMgO, " &
        "   :pFe2O3, " &
        "   :pAl2O3, " &
        "   :pCd, " &
        "   :pSizeFrctnWt, " & "   :pSizeFrctnWtMsr, " &
        "   :pSizeFrctnType, " &
        "   :pOrderNum, " &
        "   :pResult); " &
        "end;"
        Dim arA1() As Object = {"pArraySize", ItemCount, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA2() As Object = {"pSampleId", aProspRawBase.SampleId, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA3() As Object = {"pTownship", aProspRawBase.Township, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA4() As Object = {"pRange", aProspRawBase.Range, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA5() As Object = {"pSection", aProspRawBase.Section, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA6() As Object = {"pHoleLocation", aProspRawBase.HoleLocation, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA7() As Object = {"pProspDate", aProspRawBase.ProspDate, ORAPARM_INPUT, ORATYPE_DATE}
        Dim arA8() As Object = {"pSplitNumber", aProspRawBase.SplitNumber, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA9() As Object = {"pWhoModified", WhoModified, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA10() As Object = {"pWhenModified", WhenModified, ORAPARM_INPUT, ORATYPE_DATE}
        Dim arA11() As Object = {"pSizeFrctnCode", SizeFrctnCode, ORAPARM_INPUT, ORATYPE_VARCHAR2, 3}
        Dim arA12() As Object = {"pBpl", Bpl, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}

        Dim arA13() As Object = {"pFeAl", FeAl, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA14() As Object = {"pInsol", Insol, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA15() As Object = {"pCaO", CaO, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA16() As Object = {"pMgO", MgO, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA17() As Object = {"pFe2O3", Fe2O3, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA18() As Object = {"pAl2O3", Al2O3, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA19() As Object = {"pCd", Cd, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA20() As Object = {"pSizeFrctnWt", SizeFrctnWt, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA21() As Object = {"pSizeFrctnWtMsr", SizeFrctnWtMsr, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA22() As Object = {"pSizeFrctnType", SizeFrctnType, ORAPARM_INPUT, ORATYPE_VARCHAR2, 1}
        Dim arA23() As Object = {"pOrderNum", OrderNum, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}

        Dim arA24() As Object = {"pResult", 0, ORAPARM_OUTPUT, ORATYPE_NUMBER}
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
        arA13,
        arA14,
        arA15,
        arA16,
        arA17,
        arA18,
        arA19,
        arA20,
        arA21,
        arA22,
        arA23,
        arA24)
        'RunBatchSP(InsertSQL, _
        '    Array("pArraySize", ItemCount, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pSampleId", aProspRawBase.SampleId, ORAPARM_INPUT, ORATYPE_VARCHAR2), _
        '    Array("pTownship", aProspRawBase.Township, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pRange", aProspRawBase.Range, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pSection", aProspRawBase.Section, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pHoleLocation", aProspRawBase.HoleLocation, ORAPARM_INPUT, ORATYPE_VARCHAR2), _
        '    Array("pProspDate", aProspRawBase.ProspDate, ORAPARM_INPUT, ORATYPE_DATE), _
        '    Array("pSplitNumber", aProspRawBase.SplitNumber, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pWhoModified", WhoModified, ORAPARM_INPUT, ORATYPE_VARCHAR2), _
        '    Array("pWhenModified", WhenModified, ORAPARM_INPUT, ORATYPE_DATE), _
        '    Array("pSizeFrctnCode", SizeFrctnCode(), ORAPARM_INPUT, ORATYPE_VARCHAR2, 3), _
        '    Array("pBpl", Bpl(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pFeAl", FeAl(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pResult", 0, ORAPARM_OUTPUT, ORATYPE_NUMBER))
        '    Array("pInsol", Insol(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pCaO", CaO(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pMgO", MgO(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pFe2O3", Fe2O3(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pAl2O3", Al2O3(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pCd", Cd(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pSizeFrctnWt", SizeFrctnWt(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pSizeFrctnWtMsr", SizeFrctnWtMsr(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pSizeFrctnType", SizeFrctnType(), ORAPARM_INPUT, ORATYPE_VARCHAR2, 1), _
        '    Array("pOrderNum", OrderNum(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _

        Exit Sub

gUpdateProspRawSizeFrctnSprdError:
        MsgBox("Error while saving." & Str(Err.Number) &
               Chr(10) & Chr(10) &
               Err.Description, vbExclamation,
               "Update Error")
    End Sub

    Public Sub gUpdateProspRawSizeFrctnArray(ByRef aProspRawBase As gRawProspBaseType,
                                             ByVal aUpdateMode As String,
                                             ByVal aUserName As String,
                                             ByRef aSfcDataSprd() As gRawProspSfcSprdType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gUpdateProspRawSizeFrctnArrayError

        Dim SizeFrctnCode() As String
        Dim SizeFrctnWt() As Single
        Dim SizeFrctnWtMsr() As Single
        Dim SizeFrctnType() As String
        Dim Bpl() As Single
        Dim FeAl() As Single
        Dim Insol() As Single
        Dim CaO() As Single
        Dim MgO() As Single
        Dim Fe2O3() As Single
        Dim Al2O3() As Single
        Dim Cd() As Single
        Dim OrderNum() As Integer

        Dim ItemCount As Integer
        Dim InsertSQL As String
        Dim RowIdx As Integer
        Dim WhoModified As String
        Dim WhenModified As Date

        'Need to assign some dates and stuff.
        WhoModified = StrConv(aUserName, vbUpperCase)
        WhenModified = CDate(Format(Now, "MM/dd/yyyy hh:mm tt"))

        ItemCount = UBound(aSfcDataSprd)

        'If data exists then redimension transfer arrays
        If ItemCount <> 0 Then
            ReDim SizeFrctnCode(ItemCount - 1)
            ReDim SizeFrctnWt(ItemCount - 1)
            ReDim SizeFrctnWtMsr(ItemCount - 1)
            ReDim SizeFrctnType(ItemCount - 1)
            ReDim Bpl(ItemCount - 1)
            ReDim FeAl(ItemCount - 1)
            ReDim Insol(ItemCount - 1)
            ReDim CaO(ItemCount - 1)
            ReDim MgO(ItemCount - 1)
            ReDim Fe2O3(ItemCount - 1)
            ReDim Al2O3(ItemCount - 1)
            ReDim Cd(ItemCount - 1)
            ReDim OrderNum(ItemCount - 1)
        Else
            'Nothing to update
            Exit Sub
        End If

        'Size fraction codes for Cargill raw prospect:
        '1)  004 [+1/2 Pb]
        '2)  04F [-1/2 +16m Pb]
        '3)  041 [-16m + 150m Cn]
        '4)  042 [-16m + 150m Tl]
        '5)  050 [-16m + 35m Fd]
        '6)  070 [-35m + 150m Fd]
        '7)  090 [-150m Cl]

        'Now place the data into the transfer arrays.
        'Non-numeric entries ie. this measure does not apply to this
        'size fraction code will be marked with -1's.
        ItemCount = 0
        For RowIdx = 1 To UBound(aSfcDataSprd)
            'Size fraction code
            SizeFrctnCode(ItemCount) = aSfcDataSprd(RowIdx).SizeFrctnCode

            SizeFrctnType(ItemCount) = aSfcDataSprd(RowIdx).SizeFrctnType
            SizeFrctnWtMsr(ItemCount) = aSfcDataSprd(RowIdx).SizeFrctnWtMsr
            SizeFrctnWt(ItemCount) = aSfcDataSprd(RowIdx).SizeFrctnWtAdj

            '1)  004 [+1/2 Pb]
            If aSfcDataSprd(RowIdx).SizeFrctnCode = "004" Then
                Bpl(ItemCount) = aSfcDataSprd(RowIdx).Bpl
                Insol(ItemCount) = aSfcDataSprd(RowIdx).Insol
                CaO(ItemCount) = aSfcDataSprd(RowIdx).CaO
                MgO(ItemCount) = aSfcDataSprd(RowIdx).MgO
                Fe2O3(ItemCount) = aSfcDataSprd(RowIdx).Fe2O3
                Al2O3(ItemCount) = aSfcDataSprd(RowIdx).Al2O3
                FeAl(ItemCount) = aSfcDataSprd(RowIdx).FeAl
                Cd(ItemCount) = -1
                OrderNum(ItemCount) = RowIdx
            End If

            '2)  04F [-1/2 +16m Pb]
            If aSfcDataSprd(RowIdx).SizeFrctnCode = "04F" Then
                Bpl(ItemCount) = aSfcDataSprd(RowIdx).Bpl
                Insol(ItemCount) = aSfcDataSprd(RowIdx).Insol
                CaO(ItemCount) = aSfcDataSprd(RowIdx).CaO
                MgO(ItemCount) = aSfcDataSprd(RowIdx).MgO
                Fe2O3(ItemCount) = aSfcDataSprd(RowIdx).Fe2O3
                Al2O3(ItemCount) = aSfcDataSprd(RowIdx).Al2O3
                FeAl(ItemCount) = aSfcDataSprd(RowIdx).FeAl
                Cd(ItemCount) = -1
                OrderNum(ItemCount) = RowIdx
            End If

            '3)  041 [-16m + 150m Cn]
            If aSfcDataSprd(RowIdx).SizeFrctnCode = "041" Then
                Bpl(ItemCount) = aSfcDataSprd(RowIdx).Bpl
                Insol(ItemCount) = aSfcDataSprd(RowIdx).Insol
                CaO(ItemCount) = aSfcDataSprd(RowIdx).CaO
                MgO(ItemCount) = aSfcDataSprd(RowIdx).MgO
                Fe2O3(ItemCount) = aSfcDataSprd(RowIdx).Fe2O3
                Al2O3(ItemCount) = aSfcDataSprd(RowIdx).Al2O3
                FeAl(ItemCount) = aSfcDataSprd(RowIdx).FeAl
                Cd(ItemCount) = -1
                OrderNum(ItemCount) = RowIdx
            End If

            '4)  042 [-16m + 150m Tl]
            If aSfcDataSprd(RowIdx).SizeFrctnCode = "042" Then
                Bpl(ItemCount) = aSfcDataSprd(RowIdx).Bpl
                Insol(ItemCount) = -1
                CaO(ItemCount) = -1
                MgO(ItemCount) = -1
                Fe2O3(ItemCount) = -1
                Al2O3(ItemCount) = -1
                FeAl(ItemCount) = -1
                Cd(ItemCount) = -1
                OrderNum(ItemCount) = RowIdx
            End If

            '5)  050 [-16m + 35m Fd]
            If aSfcDataSprd(RowIdx).SizeFrctnCode = "050" Then
                Bpl(ItemCount) = aSfcDataSprd(RowIdx).Bpl
                Insol(ItemCount) = -1
                CaO(ItemCount) = -1
                MgO(ItemCount) = -1
                Fe2O3(ItemCount) = -1
                Al2O3(ItemCount) = -1
                FeAl(ItemCount) = -1
                Cd(ItemCount) = -1
                OrderNum(ItemCount) = RowIdx
            End If

            '6)  070 [-35m + 150m Fd]
            If aSfcDataSprd(RowIdx).SizeFrctnCode = "070" Then
                Bpl(ItemCount) = aSfcDataSprd(RowIdx).Bpl
                Insol(ItemCount) = -1
                CaO(ItemCount) = -1
                MgO(ItemCount) = -1
                Fe2O3(ItemCount) = -1
                Al2O3(ItemCount) = -1
                FeAl(ItemCount) = -1
                Cd(ItemCount) = -1
                OrderNum(ItemCount) = RowIdx
            End If

            '7)  090 [-150m Cl]
            If aSfcDataSprd(RowIdx).SizeFrctnCode = "090" Then
                Bpl(ItemCount) = aSfcDataSprd(RowIdx).Bpl
                Insol(ItemCount) = -1
                CaO(ItemCount) = -1
                MgO(ItemCount) = -1
                Fe2O3(ItemCount) = -1
                Al2O3(ItemCount) = -1
                FeAl(ItemCount) = -1
                Cd(ItemCount) = -1
                OrderNum(ItemCount) = RowIdx
            End If

            ItemCount = ItemCount + 1
        Next RowIdx

        'Procedure update_prosp_raw_size_frctn
        'pArraySize        IN     INTEGER,
        'pSampleId         IN     VARCHAR2,
        'pTownship         IN     NUMBER,
        'pRange            IN     NUMBER,
        'pSection          IN     NUMBER,
        'pHoleLocation     IN     VARCHAR2,
        'pProspDate        IN     DATE,
        'pSplitNumber      IN     NUMBER,
        'pWhoModified      IN     VARCHAR2,
        'pWhenModified     IN     DATE,
        'pSizeFrctnCode    IN     VCHAR2ARRAY3,
        'pBpl              IN     NUMBERARRAY,
        'pFeAl             IN     NUMBERARRAY,
        'pInsol            IN     NUMBERARRAY,
        'pCaO              IN     NUMBERARRAY,
        'pMgO              IN     NUMBERARRAY,
        'pFe2O3            IN     NUMBERARRAY,
        'pAl2O3            IN     NUMBERARRAY,
        'pCd               IN     NUMBERARRAY,
        'pSzeFrctnWt       IN     NUMBERARRAY,
        'pSzeFrctnWtMsr    IN     NUMBERARRAY,
        'pSzeFrctnType     IN     VCHAR2ARRAY1,
        'pOrderNum         IN     NUMBERARRAY,
        'pResult           IN OUT NUMBER)

        InsertSQL = "Begin mois.mois_raw_prospectnew.update_prosp_raw_size_frctn(" &
        "   :pArraySize, " &
        "   :pSampleId, " &
        "   :pTownship, " &
        "   :pRange, " &
        "   :pSection, " &
        "   :pHoleLocation, " &
        "   :pProspDate, " &
        "   :pSplitNumber, " &
        "   :pWhoModified, " &
        "   :pWhenModified, " &
        "   :pSizeFrctnCode, " &
        "   :pBpl, " &
        "   :pFeAl, " &
        "   :pInsol, " &
        "   :pCaO, " &
        "   :pMgO, " &
        "   :pFe2O3, " &
        "   :pAl2O3, " &
        "   :pCd, " &
        "   :pSizeFrctnWt, " & "   :pSizeFrctnWtMsr, " &
        "   :pSizeFrctnType, " &
        "   :pOrderNum, " &
        "   :pResult); " &
        "end;"
        Dim arA1() As Object = {"pArraySize", ItemCount, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA2() As Object = {"pSampleId", aProspRawBase.SampleId, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA3() As Object = {"pTownship", aProspRawBase.Township, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA4() As Object = {"pRange", aProspRawBase.Range, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA5() As Object = {"pSection", aProspRawBase.Section, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA6() As Object = {"pHoleLocation", aProspRawBase.HoleLocation, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA7() As Object = {"pProspDate", aProspRawBase.ProspDate, ORAPARM_INPUT, ORATYPE_DATE}
        Dim arA8() As Object = {"pSplitNumber", aProspRawBase.SplitNumber, ORAPARM_INPUT, ORATYPE_NUMBER}
        Dim arA9() As Object = {"pWhoModified", WhoModified, ORAPARM_INPUT, ORATYPE_VARCHAR2}
        Dim arA10() As Object = {"pWhenModified", WhenModified, ORAPARM_INPUT, ORATYPE_DATE}
        Dim arA11() As Object = {"pSizeFrctnCode", SizeFrctnCode, ORAPARM_INPUT, ORATYPE_VARCHAR2, 3}
        Dim arA12() As Object = {"pBpl", Bpl, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA13() As Object = {"pFeAl", FeAl, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA14() As Object = {"pInsol", Insol, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA15() As Object = {"pCaO", CaO, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA16() As Object = {"pMgO", MgO, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA17() As Object = {"pFe2O3", Fe2O3, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA18() As Object = {"pAl2O3", Al2O3, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA19() As Object = {"pCd", Cd, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA20() As Object = {"pSizeFrctnWt", SizeFrctnWt, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA21() As Object = {"pSizeFrctnWtMsr", SizeFrctnWtMsr, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA22() As Object = {"pSizeFrctnType", SizeFrctnType, ORAPARM_INPUT, ORATYPE_VARCHAR2, 1}
        Dim arA23() As Object = {"pOrderNum", OrderNum, ORAPARM_INPUT, ORATYPE_NUMBER, vbNull}
        Dim arA24() As Object = {"pResult", 0, ORAPARM_OUTPUT, ORATYPE_NUMBER}

        'RunBatchSP(InsertSQL, _
        '    Array("pArraySize", ItemCount, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pSampleId", aProspRawBase.SampleId, ORAPARM_INPUT, ORATYPE_VARCHAR2), _
        '    Array("pTownship", aProspRawBase.Township, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pRange", aProspRawBase.Range, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pSection", aProspRawBase.Section, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pHoleLocation", aProspRawBase.HoleLocation, ORAPARM_INPUT, ORATYPE_VARCHAR2), _
        '    Array("pProspDate", aProspRawBase.ProspDate, ORAPARM_INPUT, ORATYPE_DATE), _
        '    Array("pSplitNumber", aProspRawBase.SplitNumber, ORAPARM_INPUT, ORATYPE_NUMBER), _
        '    Array("pWhoModified", WhoModified, ORAPARM_INPUT, ORATYPE_VARCHAR2), _
        '    Array("pWhenModified", WhenModified, ORAPARM_INPUT, ORATYPE_DATE), _
        '    Array("pSizeFrctnCode", SizeFrctnCode(), ORAPARM_INPUT, ORATYPE_VARCHAR2, 3), _
        '    Array("pBpl", Bpl(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pFeAl", FeAl(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pInsol", Insol(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pCaO", CaO(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pMgO", MgO(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pFe2O3", Fe2O3(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pAl2O3", Al2O3(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pCd", Cd(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pSizeFrctnWt", SizeFrctnWt(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pSizeFrctnWtMsr", SizeFrctnWtMsr(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
        '    Array("pSizeFrctnType", SizeFrctnType(), ORAPARM_INPUT, ORATYPE_VARCHAR2, 1), _
        '    Array("pOrderNum", OrderNum(), ORAPARM_INPUT, ORATYPE_NUMBER, Null), _
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
        arA12,
        arA13,
        arA14,
        arA15,
        arA16,
        arA17,
        arA18,
        arA19,
        arA20,
        arA21,
        arA22,
        arA23,
        arA24)


        Exit Sub

gUpdateProspRawSizeFrctnArrayError:
        MsgBox("Error while saving." & Str(Err.Number) &
               Chr(10) & Chr(10) &
               Err.Description, vbExclamation,
               "Update Error")
    End Sub

    Public Function gGetRawHoleSplCnt(ByVal aSec As Integer,
                                      ByVal aTwp As Integer,
                                      ByVal aRge As Integer,
                                      ByVal aHoleLoc As String,
                                      ByVal aProspDate As Date) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetRawHoleSplCntError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim SplCount As Integer

        SplCount = 0

        params = gDBParams

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHoleLoc, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pResult", "", ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        'PROCEDURE get_raw_hole_spl_cnt
        'pTownship                  IN     NUMBER,
        'pRange                     IN     NUMBER,
        'pSection                   IN     NUMBER,
        'pHoleLocation              IN     VARCHAR2,
        'pProspDate                 IN     DATE,
        'pResult                    IN OUT NUMBER)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_raw_hole_spl_cnt(:pTownship, " +
                  ":pRange, :pSection, :pHoleLocation, " +
                  ":pProspDate, :pResult);end;", ORASQL_FAILEXEC)

        SplCount = params("pResult").Value
        ClearParams(params)

        gGetRawHoleSplCnt = SplCount

        Exit Function

gGetRawHoleSplCntError:
        MsgBox("Error getting count." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Count Error")

        On Error Resume Next
        ClearParams(params)
        gGetRawHoleSplCnt = 0
    End Function

    Public Function gGetMineForSampleId(ByVal aSampleId As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetMineForSampleIdError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim SplCount As Integer
        Dim DataDynaset As OraDynaset
        Dim ThisMine As String

        gGetMineForSampleId = ""

        params = gDBParams

        params.Add("pSampleId", aSampleId, ORAPARM_INPUT)
        params("pSampleId").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_mine_for_sampleid
        'pSampleId                  IN     VARCHAR2,
        'pResult                    IN OUT c_prosprawsplit)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_mine_for_sampleid(" +
                  ":pSampleId, :pResult);end;", ORASQL_FAILEXEC)

        DataDynaset = params("pResult").Value
        ClearParams(params)

        'Should be only one row returned.
        If DataDynaset.RecordCount = 1 Then
            DataDynaset.MoveFirst()
            If Not IsDBNull(DataDynaset.Fields("mine_name").Value) Then
                ThisMine = DataDynaset.Fields("mine_name").Value
            Else
                ThisMine = ""
            End If
        Else
            ThisMine = ""
        End If

        gGetMineForSampleId = ThisMine

        DataDynaset.Close()
        Exit Function

gGetMineForSampleIdError:
        MsgBox("Error getting mine for Sample IDt." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Process Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        DataDynaset.Close()
        gGetMineForSampleId = ""
    End Function

    '    Public Function gDispMinabilities(ByVal aSection As Integer, _
    '                                      ByVal aTownship As Integer, _
    '                                      ByVal aRange As Integer, _
    '                                      ByVal aHoleLocation As String, _
    '                                      ByVal aSplitNum As Integer, _
    '                                      ByVal aProspDate As Date, _
    '                                      ByRef aDispSpread As vaSpread) As Boolean

    '        '**********************************************************************
    '        '
    '        '
    '        '
    '        '**********************************************************************

    '        On Error GoTo DisplayMinabilitiesError

    '        Dim SampleDynaset As OraDynaset
    '        Dim SplitThk As Single
    '        Dim MetComment As String
    '        Dim ChemComment As String
    '        Dim DisplayedSplit As Integer
    '        Dim RecordCount As Integer
    '        Dim SampsOk As Boolean
    '        Dim Minability As String
    '        Dim ThisSplit As Integer

    '        DisplayedSplit = aSplitNum

    '        'Intervals will be displayed in aDispSpread.
    '        aDispSpread.MaxRows = 0

    '        SampsOk = gGetDrillHoleSamplesNew(aSection, _
    '                                          aTownship, _
    '                                          aRange, _
    '                                          aHoleLocation, _
    '                                          aProspDate, _
    '                                          SampleDynaset)

    '        If SampsOk = False Then
    '            gDispMinabilities = False
    '            Exit Function
    '        End If

    '        RecordCount = SampleDynaset.RecordCount

    '        If RecordCount = 0 Then
    '            gDispMinabilities = False
    '            Exit Function
    '        Else
    '            gDispMinabilities = True
    '        End If

    '        SampleDynaset.MoveFirst()

    '        Do While Not SampleDynaset.EOF
    '            With aDispSpread
    '                .MaxRows = .MaxRows + 1
    '                .Row = .MaxRows

    '                ThisSplit = SampleDynaset.Fields("split_number").Value
    '                .Col = 0
    '                .Text = "Spl" & CStr(ThisSplit)

    '                'Col1   Minability
    '                .Col = 1
    '                If Not isdbnull(SampleDynaset.Fields("split_minable").Value) Then
    '                    Minability = IIf(SampleDynaset.Fields("split_minable").Value = 1, "M", "U")
    '                Else
    '                    Minability = "Not assigned"
    '                End If
    '                .Text = Minability

    '                'Col2   When was the minability set?
    '                .Col = 2
    '                If Not isdbnull(SampleDynaset.Fields("split_minable_when").Value) Then
    '                    .Text = Format(SampleDynaset.Fields("split_minable_when").Value, "mm/dd/yyyy hh:mm AM/PM")
    '                Else
    '                    .Text = ""
    '                End If

    '                'Col3   Who set the minability?
    '                .Col = 3
    '                If Not isdbnull(SampleDynaset.Fields("split_minable_who").Value) Then
    '                    .Text = Format(SampleDynaset.Fields("split_minable_who").Value, "mm/dd/yyyy hh:mm AM/PM")
    '                Else
    '                    .Text = ""
    '                End If

    '                If ThisSplit = DisplayedSplit Then
    '                    .BlockMode = True
    '                    .Row = .MaxRows
    '                    .Row2 = .MaxRows
    '                    .Col = 1
    '                    .Col2 = .MaxCols
    '                    .BackColor = &HC0FFC0   'Light green
    '                    .BlockMode = False
    '                End If

    '                SampleDynaset.MoveNext()
    '            End With
    '        Loop

    '        'Get the hole minability too
    '        SampsOk = gGetDrillHoleNew(aSection, _
    '                                   aTownship, _
    '                                   aRange, _
    '                                   aHoleLocation, _
    '                                   aProspDate, _
    '                                   SampleDynaset)

    '        If SampsOk = False Then
    '            gDispMinabilities = False
    '            Exit Function
    '        End If

    '        RecordCount = SampleDynaset.RecordCount

    '        If RecordCount <> 1 Then
    '            gDispMinabilities = False
    '            Exit Function
    '        Else
    '            gDispMinabilities = True
    '        End If

    '        'Add a horizontal divider.
    '        With aDispSpread
    '            .MaxRows = .MaxRows + 1
    '            .Row = .MaxRows
    '            .action = SS_ACTION_INSERT_ROW
    '            .RowHeight(.Row) = 0.4
    '            .BlockMode = True
    '            .Row = .MaxRows
    '            .Row2 = .MaxRows
    '            .Col = 0
    '            .Col2 = .MaxCols
    '            .CellType = SS_CELL_TYPE_STATIC_TEXT
    '            .Text = " "
    '            .TypeTextShadow = False
    '            .BackColor = vbBlack
    '            .BlockMode = False
    '        End With

    '        SampleDynaset.MoveFirst()
    '        With aDispSpread
    '            .MaxRows = .MaxRows + 1
    '            .Row = .MaxRows

    '            .Col = 0
    '            .Text = "Hole"

    '            'Col1   Minability
    '            .Col = 1
    '            If Not isdbnull(SampleDynaset.Fields("hole_minable").Value) Then
    '                Minability = IIf(SampleDynaset.Fields("hole_minable").Value = 1, "M", "U")
    '            Else
    '                Minability = "Not assigned"
    '            End If
    '            .Text = Minability

    '            'Col2   When was the minability set?
    '            .Col = 2
    '            If Not isdbnull(SampleDynaset.Fields("hole_minable_when").Value) Then
    '                .Text = Format(SampleDynaset.Fields("hole_minable_when").Value, "mm/dd/yyyy hh:mm AM/PM")
    '            Else
    '                .Text = ""
    '            End If

    '            'Col3   Who set the minability?
    '            .Col = 3
    '            If Not isdbnull(SampleDynaset.Fields("hole_minable_who").Value) Then
    '                .Text = Format(SampleDynaset.Fields("hole_minable_who").Value, "mm/dd/yyyy hh:mm AM/PM")
    '            Else
    '                .Text = ""
    '            End If
    '        End With

    '        With aDispSpread
    '            .BlockMode = True
    '            .Row = 1
    '            .Row2 = .MaxRows
    '            .Col = 0
    '            .Col2 = 0
    '            .TypeTextWordWrap = False
    '            .TypeHAlign = SS_CELL_H_ALIGN_LEFT
    '            .BlockMode = False
    '        End With

    '        SampleDynaset.Close()

    '        Exit Function

    'DisplayMinabilitiesError:
    '        MsgBox("Error getting all sample#'s for this hole." & vbCrLf & _
    '               Err.Description, _
    '               vbOKOnly + vbExclamation, _
    '               "All Hole Sample#'s Access Error")

    '        On Error Resume Next
    '        gDispMinabilities = False
    '        SampleDynaset.Close()
    '    End Function

    Public Function gGetProspRawStatus(ByVal aTwp As Integer,
                                       ByVal aRge As Integer,
                                       ByVal aSec As Integer,
                                       ByVal aHloc As String,
                                       ByVal aProspDate As Date,
                                       ByRef aRedrilled As Integer,
                                       ByRef aReleased As Integer,
                                       ByRef aUseForReduction As Integer,
                                       ByRef aQaQcHole As Integer) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspRawStatusError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ProspRawDynaset As OraDynaset
        Dim RecordCount As Integer

        'PROCEDURE get_prosp_raw_status
        'pTownship           IN     NUMBER,
        'pRange              IN     NUMBER,
        'pSection            IN     NUMBER,
        'pHoleLocation       IN     VARCHAR2,
        'pProspDate          IN     DATE,
        'pResult             IN OUT c_prosprawbase)
        params = gDBParams

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHloc, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_status(" &
                  ":pTownship, :pRange, :pSection, :pHoleLocation, " &
                  ":pProspDate, :pResult);end;", ORASQL_FAILEXEC)

        ProspRawDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = ProspRawDynaset.RecordCount

        If RecordCount = 1 Then
            aReleased = ProspRawDynaset.Fields("released").Value
            aRedrilled = ProspRawDynaset.Fields("redrilled").Value
            aUseForReduction = ProspRawDynaset.Fields("use_for_reduction").Value
            aQaQcHole = ProspRawDynaset.Fields("qaqc_hole").Value
            gGetProspRawStatus = True
        Else
            aReleased = -1
            aRedrilled = -1
            aUseForReduction = -1
            aQaQcHole = -1
            gGetProspRawStatus = False
        End If

        ProspRawDynaset.Close()

        Exit Function

gGetProspRawStatusError:
        gGetProspRawStatus = False

        MsgBox("Error accessing raw prospect hole status data." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Process Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        ProspRawDynaset.Close()
    End Function

    Public Function gGetProspCodeDesc(ByVal aProspCodeTypeName As String,
                                      ByVal aProspCode As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspCodeDescError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim CodeDynaset As OraDynaset
        Dim ThisCodeDesc As String

        gGetProspCodeDesc = ""

        params = gDBParams

        params.Add("pProspCodeTypeName", aProspCodeTypeName, ORAPARM_INPUT)
        params("pProspCodeTypeName").serverType = ORATYPE_VARCHAR2

        params.Add("pProspCode", aProspCode, ORAPARM_INPUT)
        params("pProspCode").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_code
        'pProspCodeTypeName  IN     VARCHAR2,
        'pProspCode          IN     VARCHAR2,
        'pResult             IN OUT c_prospcodes);
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospectnew.get_prosp_code(" +
                      ":pProspCodeTypeName, :pProspCode, :pResult);end;", ORASQL_FAILEXEC)
        CodeDynaset = params("pResult").Value
        ClearParams(params)

        'Should be only one row returned!
        If CodeDynaset.RecordCount = 1 Then
            CodeDynaset.MoveFirst()
            ThisCodeDesc = CodeDynaset.Fields("prosp_code_desc").Value
        Else
            ThisCodeDesc = ""
        End If

        CodeDynaset.Close()
        gGetProspCodeDesc = ThisCodeDesc

        Exit Function

gGetProspCodeDescError:
        MsgBox("Error getting prospect codes." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Prospect Codes Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        CodeDynaset.Close()
        gGetProspCodeDesc = ""
    End Function

    Public Sub gSaveRdctnWhoAndWhen(ByVal aTwp As Integer,
                                    ByVal aRge As Integer,
                                    ByVal aSec As Integer,
                                    ByVal aHole As String,
                                    ByVal aProspDate As Date,
                                    ByVal aSetNull As Boolean,
                                    ByVal aWho As String,
                                    ByVal aWhen As Date)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo SaveRdctnWhoAndWhenError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim WhoReduced As String
        Dim WhenReduced As Date

        'Need to assign some dates and stuff.
        If Trim(aWho) = "" Then
            WhoReduced = StrConv(gUserName, vbUpperCase)
            WhenReduced = CDate(Format(Now, "MM/dd/yyyy hh:mm tt"))
        Else
            WhoReduced = StrConv(aWho, vbUpperCase)
            WhenReduced = aWhen
        End If

        'PROCEDURE update_rdctn_info
        'pTownship         IN     NUMBER,
        'pRange            IN     NUMBER,
        'pSection          IN     NUMBER,
        'pHoleLocation     IN     VARCHAR2,
        'pProspDate        IN     DATE,
        'pSavedMoisWhen    IN     DATE,
        'pSavedMoisWho     IN     VARCHAR2,
        'pResult           IN OUT NUMBER)
        params = gDBParams

        params.Add("pTownShip", aTwp, ORAPARM_INPUT)
        params("pTownShip").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHole, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        If aSetNull = True Then
            params.Add("pSavedMoisWhen", DBNull.Value, ORAPARM_INPUT)
            params("pSavedMoisWhen").serverType = ORATYPE_DATE

            params.Add("pSavedMoisWho", "", ORAPARM_INPUT)
            params("pSavedMoisWho").serverType = ORATYPE_VARCHAR2
        Else
            params.Add("pSavedMoisWhen", WhenReduced, ORAPARM_INPUT)
            params("pSavedMoisWhen").serverType = ORATYPE_DATE

            params.Add("pSavedMoisWho", WhoReduced, ORAPARM_INPUT)
            params("pSavedMoisWho").serverType = ORATYPE_VARCHAR2
        End If

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.update_rdctn_info(:pTownship, " +
                  ":pRange, :pSection, :pHoleLocation, :pProspDate, " +
                  ":pSavedMoisWhen, :pSavedMoisWho, :pResult);end;", ORASQL_FAILEXEC)
        ClearParams(params)

        Exit Sub

SaveRdctnWhoAndWhenError:
        MsgBox("Error while saving." & Str(Err.Number) &
               vbCrLf &
               Err.Description, vbExclamation,
               "Update Error")
    End Sub

    Public Function gGetRawProspSampleId(ByVal aTownship As Integer,
                                         ByVal aRange As Integer,
                                         ByVal aSection As Integer,
                                         ByVal aHoleLocation As String,
                                         ByVal aSplitNumber As Integer,
                                         ByVal aProspDate As Date) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetRawProspSampleIdError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim CodeDynaset As OraDynaset
        Dim ThisSampleId As String

        gGetRawProspSampleId = ""

        params = gDBParams

        params.Add("pTownship", aTownship, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRange, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSection, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHoleLocation, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pSplitNumber", aSplitNumber, ORAPARM_INPUT)
        params("pSplitNumber").serverType = ORATYPE_NUMBER

        params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_raw_sample_id
        'pTownship              IN     NUMBER,
        'pRange                 IN     NUMBER,
        'pSection               IN     NUMBER,
        'pHoleLocation          IN     VARCHAR2,
        'pSplitNumber           IN     NUMBER,
        'pProspDate             IN     DATE,
        'pResult                IN OUT c_prosprawbase);
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_sample_id(" +
                      ":pTownship, :pRange, :pSection, :pHoleLocation, " +
                      ":pSplitNumber, :pProspDate, :pResult);end;", ORASQL_FAILEXEC)
        CodeDynaset = params("pResult").Value
        ClearParams(params)

        'Should be only one row returned!
        If CodeDynaset.RecordCount = 1 Then
            CodeDynaset.MoveFirst()
            ThisSampleId = CodeDynaset.Fields("sample_id").Value
        Else
            ThisSampleId = "?"
        End If

        CodeDynaset.Close()
        gGetRawProspSampleId = ThisSampleId

        Exit Function

gGetRawProspSampleIdError:
        MsgBox("Error getting Sample ID." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Process Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        CodeDynaset.Close()
        gGetRawProspSampleId = "?"
    End Function

    '    Public Function gDispLithologies(ByVal aSection As Integer, _
    '                                     ByVal aTownship As Integer, _
    '                                     ByVal aRange As Integer, _
    '                                     ByVal aHoleLocation As String, _
    '                                     ByVal aSplitNum As Integer, _
    '                                     ByVal aProspDate As Date, _
    '                                     ByRef aDispSpread As vaSpread, _
    '                                     ByVal aShowRawMinabilities As Boolean, _
    '                                     ByVal aGeocompDate As String, _
    '                                     ByVal aGeocompHole As String, _
    '                                     ByVal aProspStandard As String, _
    '                                     ByVal aMineName As String) As Boolean

    '        '**********************************************************************
    '        '
    '        '
    '        '
    '        '**********************************************************************

    '        On Error GoTo gDispLithologiesError

    '        Dim SampleDynaset As OraDynaset
    '        Dim SplitThk As Single
    '        Dim MetComment As String
    '        Dim ChemComment As String
    '        Dim DisplayedSplit As Integer
    '        Dim RecordCount As Integer
    '        Dim SampsOk As Boolean
    '        Dim Minability As String
    '        Dim ThisSplit As Integer
    '        Dim HasRawMinability As Boolean
    '        Dim RowIdx As Integer
    '        Dim ColIdx As Integer
    '        Dim MaxColWidth As Single
    '        Dim FtlDepth As Single
    '        Dim OvbCored As Single
    '        Dim TotDepth As Single
    '        Dim Code As String
    '        Dim CodeDesc As String
    '        Dim ThisMinability As String

    '        DisplayedSplit = aSplitNum

    '        aDispSpread.MaxRows = 0

    '        'This procedure will return raw sample data with minabilities
    '        'that have been set in the raw prospect data (or not).
    '        If aShowRawMinabilities = True Then
    '            SampsOk = gGetDrillHoleSamplesNewLithRaw(aTownship, _
    '                                                     aRange, _
    '                                                     aSection, _
    '                                                     aHoleLocation, _
    '                                                     aProspDate, _
    '                                                     SampleDynaset)
    '        Else
    '            SampsOk = gGetDrillHoleSamplesNewLithSpl(aTownship, _
    '                                                     aRange, _
    '                                                     aSection, _
    '                                                     aHoleLocation, _
    '                                                     aProspDate, _
    '                                                     aProspStandard, _
    '                                                     aGeocompDate, _
    '                                                     aGeocompHole, _
    '                                                     SampleDynaset, _
    '                                                     aMineName)

    '        End If

    '        If SampsOk = False Then
    '            gDispLithologies = False
    '            Exit Function
    '        End If

    '        RecordCount = SampleDynaset.RecordCount

    '        If RecordCount = 0 Then
    '            gDispLithologies = False
    '            Exit Function
    '        Else
    '            gDispLithologies = True
    '        End If

    '        aDispSpread.Redraw = False

    '        'Add a couple of special rows first.
    '        With aDispSpread
    '            .MaxRows = .MaxRows + 1
    '            .Row = .MaxRows
    '            .Col = 0
    '            .Text = "Fishtail"
    '            .MaxRows = .MaxRows + 1
    '            .Row = .MaxRows
    '            .Col = 0
    '            .Text = "Ovb cored"
    '        End With

    '        SampleDynaset.MoveFirst()

    '        Do While Not SampleDynaset.EOF
    '            With aDispSpread
    '                .MaxRows = .MaxRows + 1
    '                .Row = .MaxRows

    '                ThisSplit = SampleDynaset.Fields("split_number").Value
    '                .Col = 0
    '                .Text = "Spl" & CStr(ThisSplit)

    '                'Col1   From
    '                .Col = 1
    '                .Value = SampleDynaset.Fields("split_depth_top").Value

    '                'Col2   To
    '                .Col = 2
    '                .Value = SampleDynaset.Fields("split_depth_bot").Value

    '                'Col3   Thick
    '                .Col = 3
    '                .Value = SampleDynaset.Fields("split_thck").Value

    '                'Col4   Minability -- Minability marked in raw prospect -- may not be there!
    '                .Col = 4
    '                If aShowRawMinabilities = True Then
    '                    If Not isdbnull(SampleDynaset.Fields("split_minable_when").Value) Then
    '                        HasRawMinability = True
    '                    Else
    '                        HasRawMinability = False
    '                    End If

    '                    If HasRawMinability = False Then
    '                        .BackColor = &H80FFFF         'Light yellow
    '                    Else
    '                        If SampleDynaset.Fields("split_minable").Value = 0 Then
    '                            .BackColor = &HC0C0FF     'Light red
    '                        Else
    '                            .BackColor = &HC0FFC0     'Light green
    '                        End If
    '                    End If
    '                Else
    '                    'GEOCOMP split mineability.
    '                    'GEOCOMP minability codes are:
    '                    'A = Active split
    '                    'M = Mined out
    '                    'I = Unminable split
    '                    'B = Bottom split
    '                    'O = Unminable hole
    '                    If Not isdbnull(SampleDynaset.Fields("minable_status").Value) Then
    '                        ThisMinability = SampleDynaset.Fields("minable_status").Value
    '                    Else
    '                        ThisMinability = ""
    '                    End If

    '                    If ThisMinability = "A" Then
    '                        .BackColor = &HC0FFC0     'Light green
    '                    Else
    '                        .BackColor = &HC0C0FF     'Light red
    '                    End If
    '                End If

    '                'Col6   Matrix color
    '                If Not isdbnull(SampleDynaset.Fields("mtx_color").Value) Then
    '                    Code = SampleDynaset.Fields("mtx_color").Value
    '                Else
    '                    Code = ""
    '                End If
    '                If Not isdbnull(SampleDynaset.Fields("mtx_color_desc").Value) Then
    '                    CodeDesc = SampleDynaset.Fields("mtx_color_desc").Value
    '                Else
    '                    CodeDesc = ""
    '                End If
    '                .Col = 6
    '                .Text = Code & "-" & CodeDesc

    '                'Col7   Matrix hardness -- Degree of consolidation
    '                If Not isdbnull(SampleDynaset.Fields("deg_consol").Value) Then
    '                    Code = SampleDynaset.Fields("deg_consol").Value
    '                Else
    '                    Code = ""
    '                End If
    '                If Not isdbnull(SampleDynaset.Fields("deg_consol_desc").Value) Then
    '                    CodeDesc = SampleDynaset.Fields("deg_consol_desc").Value
    '                Else
    '                    CodeDesc = ""
    '                End If
    '                .Col = 7
    '                .Text = Code & "-" & CodeDesc

    '                'Col8   Digging characteristics
    '                If Not isdbnull(SampleDynaset.Fields("dig_char").Value) Then
    '                    Code = SampleDynaset.Fields("dig_char").Value
    '                Else
    '                    Code = ""
    '                End If
    '                If Not isdbnull(SampleDynaset.Fields("dig_char_desc").Value) Then
    '                    CodeDesc = SampleDynaset.Fields("dig_char_desc").Value
    '                Else
    '                    CodeDesc = ""
    '                End If
    '                .Col = 8
    '                .Text = Code & "-" & CodeDesc

    '                'Col9   Pumping characteristics
    '                If Not isdbnull(SampleDynaset.Fields("pump_char").Value) Then
    '                    Code = SampleDynaset.Fields("pump_char").Value
    '                Else
    '                    Code = ""
    '                End If
    '                If Not isdbnull(SampleDynaset.Fields("pump_char_desc").Value) Then
    '                    CodeDesc = SampleDynaset.Fields("pump_char_desc").Value
    '                Else
    '                    CodeDesc = ""
    '                End If
    '                .Col = 9
    '                .Text = Code & "-" & CodeDesc

    '                'Col10  Lithology
    '                If Not isdbnull(SampleDynaset.Fields("lithology").Value) Then
    '                    Code = SampleDynaset.Fields("lithology").Value
    '                Else
    '                    Code = ""
    '                End If
    '                If Not isdbnull(SampleDynaset.Fields("lithology_desc").Value) Then
    '                    CodeDesc = SampleDynaset.Fields("lithology_desc").Value
    '                Else
    '                    CodeDesc = ""
    '                End If
    '                .Col = 10
    '                .Text = Code & "-" & CodeDesc

    '                'Col11  Physically mineable
    '                .Col = 11
    '                If SampleDynaset.Fields("phys_mineable").Value = 1 Then
    '                    .Text = "Yes"
    '                Else
    '                    .Text = "No"
    '                End If

    '                'Col12  Phosphate color
    '                If Not isdbnull(SampleDynaset.Fields("phosph_color").Value) Then
    '                    Code = SampleDynaset.Fields("phosph_color").Value
    '                Else
    '                    Code = ""
    '                End If
    '                If Not isdbnull(SampleDynaset.Fields("phosph_color_desc").Value) Then
    '                    CodeDesc = SampleDynaset.Fields("phosph_color_desc").Value
    '                Else
    '                    CodeDesc = ""
    '                End If
    '                .Col = 12
    '                .Text = Code & "-" & CodeDesc

    '                'Col13  Bed code
    '                .Col = 13
    '                If Not isdbnull(SampleDynaset.Fields("bed_code").Value) Then
    '                    .Text = SampleDynaset.Fields("bed_code").Value
    '                Else
    '                    .Text = ""
    '                End If

    '                If DisplayedSplit <> 0 Then
    '                    If ThisSplit = DisplayedSplit Then
    '                        .BlockMode = True
    '                        .Row = .MaxRows
    '                        .Row2 = .MaxRows
    '                        .Col = 1
    '                        .Col2 = 3
    '                        .BackColor = &HC0FFC0   'Light green
    '                        .BlockMode = False
    '                    End If
    '                End If

    '                If ThisSplit = 1 Then
    '                    FtlDepth = SampleDynaset.Fields("ftl_depth").Value
    '                    OvbCored = SampleDynaset.Fields("ovb_cored").Value
    '                    TotDepth = SampleDynaset.Fields("tot_depth").Value
    '                    .Row = 1
    '                    .Col = 1
    '                    .Value = 0
    '                    .Col = 2
    '                    .Value = FtlDepth
    '                    .Col = 3
    '                    .Value = FtlDepth
    '                    .Row = 2
    '                    .Col = 1
    '                    .Value = FtlDepth
    '                    .Col = 2
    '                    .Value = FtlDepth + OvbCored
    '                    .Col = 3
    '                    .Value = OvbCored
    '                End If

    '                SampleDynaset.MoveNext()
    '            End With
    '        Loop

    '        'Add a horizontal divider plus the "Total hole depth" row.
    '        With aDispSpread
    '            .MaxRows = .MaxRows + 1
    '            .Row = .MaxRows
    '            .action = SS_ACTION_INSERT_ROW
    '            .RowHeight(.Row) = 0.4
    '            .BlockMode = True
    '            .Row = .MaxRows
    '            .Row2 = .MaxRows
    '            .Col = 0
    '            .Col2 = .MaxCols
    '            .CellType = SS_CELL_TYPE_STATIC_TEXT
    '            .Text = " "
    '            .TypeTextShadow = False
    '            .BackColor = vbBlack
    '            .BlockMode = False

    '            .MaxRows = .MaxRows + 1
    '            .Row = .MaxRows
    '            .Col = 0
    '            .Text = " "
    '            .Col = 2
    '            .Value = TotDepth
    '        End With

    '        With aDispSpread
    '            .BlockMode = True
    '            .Row = 1
    '            .Row2 = .MaxRows
    '            .Col = 0
    '            .Col2 = 0
    '            .TypeTextWordWrap = False
    '            .TypeHAlign = SS_CELL_H_ALIGN_LEFT
    '            .BlockMode = False
    '        End With

    '        With aDispSpread
    '            .Row = 0
    '            .Col = 5
    '            .Text = " "
    '            .ColWidth(.Col) = 0.17
    '            For RowIdx = 0 To .MaxRows
    '                .Row = RowIdx
    '                .CellType = SS_CELL_TYPE_STATIC_TEXT
    '                .BackColor = vbBlack
    '            Next
    '        End With

    '        'Adjust column widths
    '        With aDispSpread
    '            For ColIdx = 6 To .MaxCols
    '                MaxColWidth = .MaxTextColWidth(ColIdx)
    '                .ColWidth(ColIdx) = MaxColWidth
    '            Next ColIdx
    '        End With

    '        aDispSpread.Redraw = True

    '        SampleDynaset.Close()

    '        Exit Function

    'gDispLithologiesError:
    '        MsgBox("Error getting lithologies." & vbCrLf & _
    '               Err.Description, _
    '               vbOKOnly + vbExclamation, _
    '               "Process Error")

    '        On Error Resume Next
    '        gDispLithologies = False
    '        SampleDynaset.Close()
    '    End Function

    Public Function gGetDrillHoleSamplesNewLithRaw(ByVal aTownship As Integer,
                                                   ByVal aRange As Integer,
                                                   ByVal aSection As Integer,
                                                   ByVal aHoleLocation As String,
                                                   ByVal aProspDate As Date,
                                                   ByRef aSampleDynaset As OraDynaset) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetDrillHoleSamplesNewLithRawError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer

        gGetDrillHoleSamplesNewLithRaw = False

        params = gDBParams

        params.Add("pSection", aSection, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTownship, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRange, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHoleLocation, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_raw_hole_lith
        'pTownship           IN     NUMBER,
        'pRange              IN     NUMBER,
        'pSection            IN     NUMBER,
        'pHoleLocation       IN     VARCHAR2,
        'pProspDate          IN     DATE,
        'pResult             IN OUT c_prosprawsplit)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_hole_lith(" +
                  ":pTownship, :pRange, :pSection, :pHoleLocation, :pProspDate, " +
                  ":pResult);end;", ORASQL_FAILEXEC)
        aSampleDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = aSampleDynaset.RecordCount
        gGetDrillHoleSamplesNewLithRaw = True

        Exit Function

gGetDrillHoleSamplesNewLithRawError:
        MsgBox("Error getting all sample#'s for this prospect hole." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "All Hole Sample#'s Access Error")

        On Error Resume Next
        ClearParams(params)
        gGetDrillHoleSamplesNewLithRaw = False
    End Function

    Public Function gGetDrillHoleSamplesNewLithSpl(ByVal aTownship As Integer,
                                                   ByVal aRange As Integer,
                                                   ByVal aSection As Integer,
                                                   ByVal aHoleLocation As String,
                                                   ByVal aProspDate As Date,
                                                   ByVal aProspStandard As String,
                                                   ByVal aGeocompDate As String,
                                                   ByVal aGeocompHole As String,
                                                   ByRef aSampleDynaset As OraDynaset,
                                                   ByVal aMineName As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetDrillHoleSamplesNewLithSplError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer

        'Assume that aProspStandard will be "100%PROSPECT" or "CATALOG".

        gGetDrillHoleSamplesNewLithSpl = False

        params = gDBParams

        params.Add("pSection", aSection, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTownship, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRange, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHoleLocation, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pProspStandard", aProspStandard, ORAPARM_INPUT)
        params("pProspStandard").serverType = ORATYPE_VARCHAR2

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pGeocompDate", aGeocompDate, ORAPARM_INPUT)
        params("pGeocompDate").serverType = ORATYPE_VARCHAR2

        params.Add("pGeocompHole", aGeocompHole, ORAPARM_INPUT)
        params("pGeocompHole").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_raw_hole_lith3
        'pTownship           IN     NUMBER,
        'pRange              IN     NUMBER,
        'pSection            IN     NUMBER,
        'pHoleLocation       IN     VARCHAR2,
        'pProspDate          IN     DATE,
        'pProspStandard      IN     VARCHAR2,
        'pMineName           IN     VARCHAR2,
        'pGeocompHole        IN     VARCHAR2,
        'pResult             IN OUT c_prosprawsplit)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_hole_lith3(" +
                  ":pTownship, :pRange, :pSection, :pHoleLocation, :pProspDate, " +
                  ":pProspStandard, :pMineName, :pGeocompDate, :pGeocompHole, :pResult);end;", ORASQL_FAILEXEC)
        aSampleDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = aSampleDynaset.RecordCount
        gGetDrillHoleSamplesNewLithSpl = True

        Exit Function

gGetDrillHoleSamplesNewLithSplError:
        MsgBox("Error getting all sample#'s for this prospect hole." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "All Hole Sample#'s Access Error")

        On Error Resume Next
        ClearParams(params)
        gGetDrillHoleSamplesNewLithSpl = False
    End Function

    Public Function gGetSampIdSpec(ByVal aMine As String,
                                   ByVal aSec As Integer,
                                   ByVal aTwp As Integer,
                                   ByVal aRge As Integer,
                                   ByVal aHoleLoc As String,
                                   ByVal aSplit As Integer) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetSampIdSpecError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        gGetSampIdSpec = ""

        params = gDBParams

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHoleLoc, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pSplitNumber", aSplit, ORAPARM_INPUT)
        params("pSplitNumber").serverType = ORATYPE_NUMBER

        params.Add("pSampleId", "", ORAPARM_OUTPUT)
        params("pSampleId").serverType = ORATYPE_VARCHAR2

        'PROCEDURE get_sample_id_spec
        'pTownship                  IN     NUMBER,
        'pRange                     IN     NUMBER,
        'pSection                   IN     NUMBER,
        'pHoleLocation              IN     VARCHAR2,
        'pSplitNumber               IN     NUMBER,
        'pSampleId                  IN OUT VARCHAR2)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_sample_id_spec(:pTownship, " +
                  ":pRange, :pSection, :pHoleLocation, " +
                  ":pSplitNumber, :pSampleId);end;", ORASQL_FAILEXEC)

        gGetSampIdSpec = params("pSampleId").Value
        ClearParams(params)

        Exit Function

gGetSampIdSpecError:
        'Let's not show an error message here!
        'MsgBox "Error getting Sample ID." & vbCrLf & _
        'Err.Description, _
        'vbOKOnly + vbExclamation, _
        '"Sample ID Access Error"

        On Error Resume Next
        gGetSampIdSpec = ""
        On Error Resume Next
        ClearParams(params)
    End Function

    Public Function gGetSampIdSpec2(ByVal aSec As Integer,
                                    ByVal aTwp As Integer,
                                    ByVal aRge As Integer,
                                    ByVal aHoleLoc As String,
                                    ByVal aSplit As Integer) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetSampIdSpec2Error

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        gGetSampIdSpec2 = ""

        params = gDBParams

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHoleLoc, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pSplitNumber", aSplit, ORAPARM_INPUT)
        params("pSplitNumber").serverType = ORATYPE_NUMBER

        params.Add("pSampleId", "", ORAPARM_OUTPUT)
        params("pSampleId").serverType = ORATYPE_VARCHAR2

        'PROCEDURE get_sample_id_spec2
        'pTownship                  IN     NUMBER,
        'pRange                     IN     NUMBER,
        'pSection                   IN     NUMBER,
        'pHoleLocation              IN     VARCHAR2,
        'pSplitNumber               IN     NUMBER,
        'pSampleId                  IN OUT VARCHAR2)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_sample_id_spec2(:pTownship, " +
                  ":pRange, :pSection, :pHoleLocation, " +
                  ":pSplitNumber, :pSampleId);end;", ORASQL_FAILEXEC)

        gGetSampIdSpec2 = params("pSampleId").Value
        ClearParams(params)

        Exit Function

gGetSampIdSpec2Error:
        'Let's not show an error message here!
        'MsgBox "Error getting Sample ID." & vbCrLf & _
        'Err.Description, _
        'vbOKOnly + vbExclamation, _
        '"Sample ID Access Error"

        On Error Resume Next
        gGetSampIdSpec2 = ""
        On Error Resume Next
        ClearParams(params)
    End Function

    '    Public Sub gPrintRawProspRpt(ByVal aSampNum As String, _
    '                                 ByRef aReport As CrystalReport)

    '        '**********************************************************************
    '        '
    '        '
    '        '
    '        '**********************************************************************

    '        On Error GoTo PrintRawProspRptError

    '        Dim ConnectString As String

    '        'Reporting application = Seagate Crystal Reports Professional
    '        aReport.ReportFileName = gPath + "\Reports\" + _
    '                                 "IndSampRawProsp2.rpt"

    '        'Connect to Oracle database
    '        ConnectString = "DSN = " + gDataSource + ";UID = " + gOracleUserName + _
    '                        ";PWD = " + gOracleUserPassword + ";DSQ = "

    '        aReport.Connect = ConnectString

    '        'Need to pass the company name into the report
    '        aReport.ParameterFields(0) = "pCompanyName;" & gCompanyName & ";TRUE"
    '        aReport.ParameterFields(1) = "pSampleNum;" & aSampNum & ";TRUE"
    '        aReport.ParameterFields(2) = "pResult;" & " " & ";TRUE"

    '        'Report window maximized
    '        aReport.WindowState = crptMaximized

    '        aReport.WindowTitle = "Raw Prospect Hole Split"

    '        'User allowed to minimize report window
    '        aReport.WindowMinButton = True

    '        'Start Crystal Reports
    '        aReport.action = 1

    '        Exit Sub

    'PrintRawProspRptError:
    '        MsgBox("Error printing report." & vbCrLf & _
    '               Err.Description, _
    '               vbOKOnly + vbExclamation, _
    '               "Report Printing Error")
    '    End Sub

    Public Function gGetSampIdProspDate(ByVal aSampleId As String) As Date

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetSampIdProspDateError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        gGetSampIdProspDate = #12/31/5555#

        params = gDBParams

        params.Add("pSampleId", aSampleId, ORAPARM_INPUT)
        params("pSampleId").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", "", ORAPARM_OUTPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        'PROCEDURE get_sample_prosp_date
        'pSampleId                  IN     VARCHAR2,
        'pProspDate                 IN OUT DATE)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_sample_prosp_date(:pSampleId, " +
                  ":pProspDate);end;", ORASQL_FAILEXEC)

        gGetSampIdProspDate = params("pProspDate").Value
        ClearParams(params)

        Exit Function

gGetSampIdProspDateError:
        'Let's not show an error message here!
        'MsgBox "Error getting Sample ID." & vbCrLf & _
        'Err.Description, _
        'vbOKOnly + vbExclamation, _
        '"Sample ID Access Error"

        On Error Resume Next
        gGetSampIdProspDate = #12/31/5555#
        On Error Resume Next
        ClearParams(params)
    End Function

    Public Function gGetProspRawDataAllNew(ByVal aTwp As Integer,
                                           ByVal aRge As Integer,
                                           ByVal aSec As Integer,
                                           ByVal aHloc As String,
                                           ByVal aProspDate As Date,
                                           ByRef aProspRawData() As gRawProspBaseType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspRawDataAllNewError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ProspRawDynaset As OraDynaset
        Dim RecordCount As Integer
        Dim SplCnt As Integer

        'Get all of the raw split data for a prospect hole -- all of the splits assigned
        'to a prospect hole.

        params = gDBParams

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_VARCHAR2

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHloc, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspDate", aProspDate, ORAPARM_INPUT)
        params("pProspDate").serverType = ORATYPE_DATE

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_raw_split_hole
        'pTownship           IN     NUMBER,
        'pRange              IN     NUMBER,
        'pSection            IN     NUMBER,
        'pHoleLocation       IN     VARCHAR2,
        'pProspDate          IN     DATE,
        'pResult             IN OUT c_prosprawsplit)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_split_hole(" &
                  ":pTownship, :pRange, :pSection, :pHoleLocation, " &
                  ":pProspDate, :pResult);end;", ORASQL_FAILEXEC)

        ProspRawDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = ProspRawDynaset.RecordCount

        If RecordCount = 0 Then
            gGetProspRawDataAllNew = False
            Exit Function
        End If

        ReDim aProspRawData(RecordCount)

        ProspRawDynaset.MoveFirst()
        SplCnt = 1
        Do While Not ProspRawDynaset.EOF
            With aProspRawData(SplCnt)
                .Township = ProspRawDynaset.Fields("township").Value
                .Range = ProspRawDynaset.Fields("range").Value
                .Section = ProspRawDynaset.Fields("section").Value
                .HoleLocation = ProspRawDynaset.Fields("hole_location").Value
                .ProspDate = Format(ProspRawDynaset.Fields("prosp_date").Value, "MM/dd/yyyy")
                '--
                .SampleId = ProspRawDynaset.Fields("sample_id").Value
                .SplitNumber = ProspRawDynaset.Fields("split_number").Value
                .Barren = ProspRawDynaset.Fields("barren").Value
                .SplitFtlBottom = ProspRawDynaset.Fields("split_ftl_bottom").Value
                .MtxTotWetWt = ProspRawDynaset.Fields("mtx_tot_wet_wt").Value
                .MtxMoistWetWt = ProspRawDynaset.Fields("mtx_moist_wet_wt").Value
                .MtxMoistDryWt = ProspRawDynaset.Fields("mtx_moist_dry_wt").Value
                .MtxMoistTareWt = ProspRawDynaset.Fields("mtx_moist_tare_wt").Value
                .FdTotWetWt = ProspRawDynaset.Fields("fd_tot_wet_wt").Value
                .FdTotWetWtMsr = ProspRawDynaset.Fields("fd_tot_wet_wt_msr").Value
                .FdMoistWetWt = ProspRawDynaset.Fields("fd_moist_wet_wt").Value
                .FdMoistDryWt = ProspRawDynaset.Fields("fd_moist_dry_wt").Value
                .FdMoistTareWt = ProspRawDynaset.Fields("fd_moist_tare_wt").Value
                .FdScrnSampWt = ProspRawDynaset.Fields("fd_scrn_samp_wt").Value
                .DensCylSize = ProspRawDynaset.Fields("dens_cyl_size").Value
                .DensCylWetWt = ProspRawDynaset.Fields("dens_cyl_wet_wt").Value
                .DensCylH2oWt = ProspRawDynaset.Fields("dens_cyl_h2o_wt").Value
                .DryDensity = ProspRawDynaset.Fields("dry_density").Value
                .FlotFdWetWt = ProspRawDynaset.Fields("flot_wet_wt").Value
                .MtxProcWetWt = ProspRawDynaset.Fields("mtx_proc_wet_wt").Value
                .ExpExcessWt = ProspRawDynaset.Fields("exp_excess_wt").Value

                If Not IsDBNull(ProspRawDynaset.Fields("mtx_color").Value) Then
                    .MtxColor = ProspRawDynaset.Fields("mtx_color").Value
                Else
                    .MtxColor = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("deg_consol").Value) Then
                    .DegConsol = ProspRawDynaset.Fields("deg_consol").Value
                Else
                    .DegConsol = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("dig_char").Value) Then
                    .DigChar = ProspRawDynaset.Fields("dig_char").Value
                Else
                    .DigChar = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("pump_char").Value) Then
                    .PumpChar = ProspRawDynaset.Fields("pump_char").Value
                Else
                    .PumpChar = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("lithology").Value) Then
                    .Lithology = ProspRawDynaset.Fields("lithology").Value
                Else
                    .Lithology = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("phosph_color").Value) Then
                    .PhosphColor = ProspRawDynaset.Fields("phosph_color").Value
                Else
                    .PhosphColor = ""
                End If

                .PhysMineable = ProspRawDynaset.Fields("phys_mineable").Value

                If Not IsDBNull(ProspRawDynaset.Fields("clay_sett_char").Value) Then
                    .ClaySettChar = ProspRawDynaset.Fields("clay_sett_char").Value
                Else
                    .ClaySettChar = ""
                End If

                .FdScrnSampWtComp = ProspRawDynaset.Fields("fd_scrn_samp_wt_comp").Value
                .RecordLocked = ProspRawDynaset.Fields("record_locked").Value

                If Not IsDBNull(ProspRawDynaset.Fields("date_chem_lab").Value) Then
                    'Want date and time!
                    .DateChemLab = Format(ProspRawDynaset.Fields("date_chem_lab").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .DateChemLab = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("who_chem_lab").Value) Then
                    .WhoChemLab = ProspRawDynaset.Fields("who_chem_lab").Value
                Else
                    .WhoChemLab = ""
                End If

                .RerunStatus = ProspRawDynaset.Fields("rerun_status").Value

                If Not IsDBNull(ProspRawDynaset.Fields("date_rerun").Value) Then
                    .DateRerun = Format(ProspRawDynaset.Fields("date_rerun").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .DateRerun = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("metlab_comment").Value) Then
                    .MetLabComment = ProspRawDynaset.Fields("metlab_comment").Value
                Else
                    .MetLabComment = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("chemlab_comment").Value) Then
                    .ChemLabComment = ProspRawDynaset.Fields("chemlab_comment").Value
                Else
                    .ChemLabComment = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("date_met_lab").Value) Then
                    'Want date and time!
                    .DateMetLab = Format(ProspRawDynaset.Fields("date_met_lab").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .DateMetLab = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("who_met_lab").Value) Then
                    .WhoMetLab = ProspRawDynaset.Fields("who_met_lab").Value
                Else
                    .WhoMetLab = ""
                End If

                .SplitDepthTop = ProspRawDynaset.Fields("split_depth_top").Value
                .SplitDepthBot = ProspRawDynaset.Fields("split_depth_bot").Value
                .SplitThck = ProspRawDynaset.Fields("split_thck").Value

                If Not IsDBNull(ProspRawDynaset.Fields("wash_date").Value) Then
                    .WashDate = Format(ProspRawDynaset.Fields("wash_date").Value, "MM/dd/yyyy")
                Else
                    .WashDate = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("who_modified").Value) Then
                    .WhoModifiedSplit = ProspRawDynaset.Fields("who_modified").Value
                Else
                    .WhoModifiedSplit = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("when_modified").Value) Then
                    .WhenModifiedSplit = Format(ProspRawDynaset.Fields("when_modified").Value, "MM/dd/yyyy")
                Else
                    .WhenModifiedSplit = ""
                End If

                .OrigData = ProspRawDynaset.Fields("orig_data").Value

                '-----
                'New columns added 10/30/2007, lss

                If Not IsDBNull(ProspRawDynaset.Fields("split_minable").Value) Then
                    .SplitMinable = ProspRawDynaset.Fields("split_minable").Value
                Else
                    'A null hole minable value will be represented with -1.
                    'It will be displayed as "NA" = Not assigned.
                    .SplitMinable = -1
                End If
                If IsDate(ProspRawDynaset.Fields("split_minable_when").Value) Then
                    'Want date and time!
                    .SplitMinableWhen = Format(ProspRawDynaset.Fields("split_minable_when").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .SplitMinableWhen = ""
                End If
                If Not IsDBNull(ProspRawDynaset.Fields("split_minable_who").Value) Then
                    .SplitMinableWho = ProspRawDynaset.Fields("split_minable_who").Value
                Else
                    .SplitMinableWho = ""
                End If
                If Not IsDBNull(ProspRawDynaset.Fields("sample_id_cargill").Value) Then
                    .SampleIdCargill = ProspRawDynaset.Fields("sample_id_cargill").Value
                Else
                    .SampleIdCargill = ""
                End If
                If Not IsDBNull(ProspRawDynaset.Fields("bed_code").Value) Then
                    .BedCode = ProspRawDynaset.Fields("bed_code").Value
                Else
                    .BedCode = ""
                End If

                .ClaySettlingLvl = ProspRawDynaset.Fields("clay_settling_lvl").Value
                .PbClayPct = ProspRawDynaset.Fields("pb_clay_pct").Value
            End With
            ProspRawDynaset.MoveNext()
            SplCnt = SplCnt + 1
        Loop

        ProspRawDynaset.Close()

        gGetProspRawDataAllNew = True

        Exit Function

gGetProspRawDataAllNewError:
        gGetProspRawDataAllNew = False

        MsgBox("Error accessing raw prospect data." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Raw Prospect Data Access Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        ProspRawDynaset.Close()
    End Function

    Public Sub gGetMineAreaRawProsp(ByRef aRawDynaset As OraDynaset,
                                    ByVal aMineName As String,
                                    ByVal aSpecAreaName As String,
                                    ByVal aTwp As Integer,
                                    ByVal aRge As Integer,
                                    ByVal aSec As Integer,
                                    ByVal aUseRawProspMineDesignation As Integer,
                                    ByVal aExpDrillOnly As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetMineAreaRawProspError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Long

        'Get SFC data
        If aTwp <> 0 And aRge <> 0 And aSec <> 0 Then
            params = gDBParams

            params.Add("pTownship", aTwp, ORAPARM_INPUT)
            params("pTownship").serverType = ORATYPE_NUMBER

            params.Add("pRange", aRge, ORAPARM_INPUT)
            params("pRange").serverType = ORATYPE_NUMBER

            params.Add("pSection", aSec, ORAPARM_INPUT)
            params("pSection").serverType = ORATYPE_NUMBER

            'May only want the expanded drilling.
            params.Add("pExpDrillOnly", aExpDrillOnly, ORAPARM_INPUT)
            params("pExpDrillOnly").serverType = ORATYPE_NUMBER

            params.Add("pResult", 0, ORAPARM_OUTPUT)
            params("pResult").serverType = ORATYPE_CURSOR

            'PROCEDURE get_prosp_raw_sec_sfc3
            'pTownship           IN     NUMBER,
            'pRange              IN     NUMBER,
            'pSection            IN     NUMBER,
            'pExpDrillOnly       IN     NUMBER,
            'pResult             IN OUT c_prosprawsplit)
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_sec_sfc3(:pTownship, " +
                      ":pRange, :pSection, :pExpDrillOnly, :pResult);end;", ORASQL_FAILEXEC)
            aRawDynaset = params("pResult").Value
            ClearParams(params)
        Else
            params = gDBParams

            params.Add("pMineName", aMineName, ORAPARM_INPUT)
            params("pMineName").serverType = ORATYPE_VARCHAR2

            params.Add("pAreaName", aSpecAreaName, ORAPARM_INPUT)
            params("pAreaName").serverType = ORATYPE_VARCHAR2

            'May only want the expanded drilling.
            params.Add("pExpDrillOnly", aExpDrillOnly, ORAPARM_INPUT)
            params("pExpDrillOnly").serverType = ORATYPE_NUMBER

            params.Add("pUseRawProspMine", aUseRawProspMineDesignation, ORAPARM_INPUT)
            params("pUseRawProspMine").serverType = ORATYPE_NUMBER

            params.Add("pResult", 0, ORAPARM_OUTPUT)
            params("pResult").serverType = ORATYPE_CURSOR

            'PROCEDURE get_prosp_raw_sec_sfc2
            'pMineName           IN     VARCHAR2,
            'pExpDrillOnly       IN     NUMBER,
            'pUseRawProspMine    IN     NUMBER,
            'pResult             IN OUT c_prosprawsplit)
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_sec_sfc2(:pMineName, " +
                      ":pAreaName, :pExpDrillOnly, :pUseRawProspMine, :pResult);end;", ORASQL_FAILEXEC)
            aRawDynaset = params("pResult").Value
            ClearParams(params)
        End If

        RecordCount = aRawDynaset.RecordCount

        ClearParams(params)

        Exit Sub

GetMineAreaRawProspError:
        On Error Resume Next

        MsgBox("Error getting raw prospect data." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Process Error")

        On Error Resume Next
        ClearParams(params)
    End Sub

    Public Function gGetProspRawHoleDataOnly2(ByVal aTwp As Integer,
                                              ByVal aRge As Integer,
                                              ByVal aSec As Integer,
                                              ByVal aHloc As String,
                                              ByRef aProspRawHoleData As gRawProspBaseHoleType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspRawHoleDataOnly2Error

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ProspRawDynaset As OraDynaset
        Dim RecordCount As Integer

        'Will get only a hole that hs not been marked as redrilled in
        'PROSP_RAW_BASE (redrilled = 0).  Ideally there should be only one such
        'prospect hole for a location -- any additional holes (if Earnest Terry is
        'doing his job will be marked as a redrill).

        'PROCEDURE get_prosp_raw_base_redrill
        'pTownship              IN     NUMBER,
        'pRange                 IN     NUMBER,
        'pSection               IN     NUMBER,
        'pHoleLocation          IN     VARCHAR2,
        'pResult                IN OUT c_prosprawbase);
        params = gDBParams

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHloc, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew.get_prosp_raw_base_redrill(" &
                  ":pTownship, :pRange, :pSection, :pHoleLocation, " &
                  ":pResult);end;", ORASQL_FAILEXEC)

        ProspRawDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = ProspRawDynaset.RecordCount

        'Ideally only one record should be returned!

        If RecordCount = 1 Then
            gGetProspRawHoleDataOnly2 = True

            ProspRawDynaset.MoveFirst()
            With aProspRawHoleData
                .Township = ProspRawDynaset.Fields("township").Value
                .Range = ProspRawDynaset.Fields("range").Value
                .Section = ProspRawDynaset.Fields("section").Value
                .HoleLocation = ProspRawDynaset.Fields("hole_location").Value
                .Forty = ProspRawDynaset.Fields("forty").Value
                .State = ProspRawDynaset.Fields("state").Value
                .Quadrant = ProspRawDynaset.Fields("quadrant").Value

                If Not IsDBNull(ProspRawDynaset.Fields("mine_name").Value) Then
                    .MineName = ProspRawDynaset.Fields("mine_name").Value
                Else
                    .MineName = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("spec_area_name").Value) Then
                    .SpecAreaName = ProspRawDynaset.Fields("spec_area_name").Value
                Else
                    .SpecAreaName = ""
                End If

                .ExpDrill = ProspRawDynaset.Fields("exp_drill").Value
                .SplitTotalNum = ProspRawDynaset.Fields("split_total_num").Value
                .Xcoord = ProspRawDynaset.Fields("x_coord").Value
                .Ycoord = ProspRawDynaset.Fields("y_coord").Value
                .FtlDepth = ProspRawDynaset.Fields("ftl_depth").Value
                .OvbCored = ProspRawDynaset.Fields("ovb_cored").Value
                .Ownership = ProspRawDynaset.Fields("ownership").Value

                If Not IsDBNull(ProspRawDynaset.Fields("ownership_desc").Value) Then
                    .Ownership = .Ownership & " - " & ProspRawDynaset.Fields("ownership_desc").Value
                Else
                    .Ownership = .Ownership & " - ??"
                End If

                .ProspDate = Format(ProspRawDynaset.Fields("prosp_date").Value, "MM/dd/yyyy")

                .MinedStatus = ProspRawDynaset.Fields("mined_status").Value
                .Elevation = ProspRawDynaset.Fields("elevation").Value
                .TotDepth = ProspRawDynaset.Fields("tot_depth").Value
                .Aoi = ProspRawDynaset.Fields("aoi").Value
                .CoordSurveyed = ProspRawDynaset.Fields("coord_surveyed").Value

                If Not IsDBNull(ProspRawDynaset.Fields("long_comment").Value) Then
                    .HoleComment = ProspRawDynaset.Fields("long_comment").Value
                Else
                    .HoleComment = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("hole_location_char").Value) Then
                    .HoleLocationChar = ProspRawDynaset.Fields("hole_location_char").Value
                Else
                    .HoleLocationChar = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("who_modified").Value) Then
                    .WhoModifiedHole = ProspRawDynaset.Fields("who_modified").Value
                Else
                    .WhoModifiedHole = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("when_modified").Value) Then
                    'Want date and time!
                    .WhenModifiedHole = Format(ProspRawDynaset.Fields("when_modified").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .WhenModifiedHole = ""
                End If

                If Not IsDBNull(ProspRawDynaset.Fields("log_date").Value) Then
                    .LogDate = Format(ProspRawDynaset.Fields("log_date").Value, "MM/dd/yyyy")
                Else
                    .LogDate = ""
                End If

                .Released = ProspRawDynaset.Fields("released").Value
                .Redrilled = ProspRawDynaset.Fields("redrilled").Value

                If Not IsDBNull(ProspRawDynaset.Fields("redrill_date").Value) Then
                    .RedrillDate = Format(ProspRawDynaset.Fields("redrill_date").Value, "MM/dd/yyyy")
                Else
                    .RedrillDate = ""
                End If

                .UseForReduction = ProspRawDynaset.Fields("use_for_reduction").Value

                If Not IsDBNull(ProspRawDynaset.Fields("hole_minable").Value) Then
                    .HoleMinable = ProspRawDynaset.Fields("hole_minable").Value
                Else
                    'A null hole minable value will be represented with -1.
                    'It will be displayed as "NA" = Not assigned.
                    .HoleMinable = -1
                End If
                If IsDate(ProspRawDynaset.Fields("hole_minable_when").Value) Then
                    'Want date and time!
                    .HoleMinableWhen = Format(ProspRawDynaset.Fields("hole_minable_when").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .HoleMinableWhen = ""
                End If
                If Not IsDBNull(ProspRawDynaset.Fields("hole_minable_who").Value) Then
                    .HoleMinableWho = ProspRawDynaset.Fields("hole_minable_who").Value
                Else
                    .HoleMinableWho = ""
                End If
                .ManufacturedData = ProspRawDynaset.Fields("manufactured_data").Value

                If IsDate(ProspRawDynaset.Fields("saved_mois_when").Value) Then
                    'Want date and time!
                    .SavedMoisWhen = Format(ProspRawDynaset.Fields("saved_mois_when").Value, "MM/dd/yyyy hh:mm tt")
                Else
                    .SavedMoisWhen = ""
                End If
                If Not IsDBNull(ProspRawDynaset.Fields("saved_mois_who").Value) Then
                    .SavedMoisWho = ProspRawDynaset.Fields("saved_mois_who").Value
                Else
                    .SavedMoisWho = ""
                End If
            End With
        Else
            gGetProspRawHoleDataOnly2 = False
        End If

        ProspRawDynaset.Close()

        Exit Function

gGetProspRawHoleDataOnly2Error:
        gGetProspRawHoleDataOnly2 = False

        MsgBox("Error accessing raw prospect hole only data." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Raw Prospect Hole Only Data Access Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        ProspRawDynaset.Close()
    End Function

    'Public Sub gGetProspCodesToGrid(ByRef aSpread As vaSpread, _
    '                                ByVal aProspCodeTypeName As String)

    ''**********************************************************************
    ''
    ''
    ''
    ''**********************************************************************

    '    On Error GoTo gGetProspCodesToGridError

    '    Dim params As OraParameters
    '    Dim SQLStmt As OraSqlStmt
    '    Dim CodeDynaset As OraDynaset
    '    Dim ThisCode As String
    '    Dim ThisCodeDesc As String

    '    aSpread.MaxRows = 0

    '        params = gDBParams

    '    params.Add "pProspCodeTypeName", aProspCodeTypeName, ORAPARM_INPUT
    '    params("pProspCodeTypeName").serverType = ORATYPE_VARCHAR2

    '    params.Add "pResult", 0, ORAPARM_OUTPUT
    '    params("pResult").serverType = ORATYPE_CURSOR

    '    'PROCEDURE get_prosp_codes
    '    'pProspCodeTypeName   IN     VARCHAR2,
    '    'pResult              IN OUT c_prospcodes)
    '        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospectnew.get_prosp_codes(" + _
    '                      ":pProspCodeTypeName, :pResult);end;", ORASQL_FAILEXEC)
    '        CodeDynaset = params("pResult").Value
    '    ClearParams params

    '    CodeDynaset.MoveFirst
    '    Do While Not CodeDynaset.EOF
    '        ThisCode = CodeDynaset.Fields("prosp_code").Value
    '        ThisCodeDesc = CodeDynaset.Fields("prosp_code_desc").Value

    '        With aSpread
    '            .MaxRows = .MaxRows + 1
    '            .Row = .MaxRows
    '            .Col = 1
    '            .Text = ThisCode
    '            .Col = 2
    '            .Text = ThisCodeDesc
    '        End With

    '        CodeDynaset.MoveNext
    '    Loop

    '    CodeDynaset.Close

    '    Exit Sub

    'gGetProspCodesToGridError:
    '    MsgBox "Error getting prospect codes." & vbCrLf & _
    '           Err.Description, _
    '           vbOKOnly + vbExclamation, _
    '           "Prospect Codes Error"

    '    On Error Resume Next
    '    ClearParams params
    '    On Error Resume Next
    '    CodeDynaset.Close
    'End Sub

    'Public Sub gFixHoleCol(ByVal aFixMode As String, _
    '                       ByRef aSpread As vaSpread)

    ''**********************************************************************
    ''
    ''
    ''
    ''**********************************************************************

    '    Dim RowIdx As Integer
    '    Dim LenSplTxt As Integer
    '    'FixMode will be "On" or "Off".

    '    With aSpread
    '        For RowIdx = 1 To .MaxRows
    '            .Row = RowIdx
    '            .Col = 1
    '            If .TypeButtonColor = &HC0C0FF Then      'Light red -- unminable hole
    '                If aFixMode = "On" Then
    '                    .TypeButtonText = "x " & .TypeButtonText & " x"
    '                Else
    '                    LenSplTxt = Len(.TypeButtonText) - 4
    '                    .TypeButtonText = Mid(.TypeButtonText, 3, LenSplTxt)
    '                End If
    '            End If
    '        Next RowIdx
    '    End With
    'End Sub

    'Public Sub gFixLithologyCol(ByVal aFixMode As String, _
    '                            ByRef aSpread As vaSpread)

    ''**********************************************************************
    ''
    ''
    ''
    ''**********************************************************************

    '    Dim RowIdx As Integer

    '    'FixMode will be "On" or "Off".

    '    With aSpread
    '        For RowIdx = 1 To .MaxRows
    '            .Row = RowIdx
    '            .Col = 4
    '            If .BackColor = &HC0C0FF Then      'Light red -- unminable hole
    '                If aFixMode = "On" Then
    '                    .TypeHAlign = SS_CELL_H_ALIGN_CENTER
    '                    .Text = "x"
    '                Else
    '                    .Text = " "
    '                End If
    '            End If
    '        Next RowIdx
    '    End With
    'End Sub

    'Public Sub gAppendLithology(ByRef aHoleDataSpread As vaSpread, _
    '                            ByRef aLithologySpread As vaSpread)

    ''**********************************************************************
    ''
    ''
    ''
    ''**********************************************************************

    '    'Append lithology to hole data if possible!

    '    Dim RowIdx As Integer
    '    Dim ColIdx As Integer
    '    Dim HoleDataSplThk As Single
    '    Dim ThisTxt As String

    '    If aLithologySpread.MaxRows <> aHoleDataSpread.MaxRows + 3 Then
    '        Exit Sub
    '    End If

    '    With aHoleDataSpread
    '        For RowIdx = 2 To .MaxRows
    '            .Row = RowIdx
    '            .Col = 2
    '            HoleDataSplThk = .Value

    '            With aLithologySpread
    '                .Row = RowIdx + 1
    '                .Col = 3
    '                If .Value <> HoleDataSplThk Then
    '                    Exit Sub
    '                End If
    '            End With
    '        Next RowIdx
    '    End With

    '    'Everything is OK -- append the data
    '    With aLithologySpread
    '        For RowIdx = 3 To .MaxRows - 2
    '            .Row = RowIdx
    '            For ColIdx = 6 To .MaxCols
    '                .Col = ColIdx
    '                ThisTxt = .Text

    '                aHoleDataSpread.Row = RowIdx - 1
    '                aHoleDataSpread.Col = ColIdx + 8
    '                aHoleDataSpread.Text = ThisTxt
    '            Next ColIdx
    '        Next RowIdx
    '    End With
    'End Sub

    Public Function gGetMineAreaSpecAbbrv(ByVal aName As String,
                                          ByVal aMode As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'This stuff is hard-coded here.  In the .Net version it will not
        'be.

        If aMode = "Mine name" Then
            Select Case aName
                Case Is = "Hookers Prairie"
                    gGetMineAreaSpecAbbrv = "HP"
                Case Is = "Hopewell"
                    gGetMineAreaSpecAbbrv = "HW"
                Case Is = "Ona-Pioneer"
                    gGetMineAreaSpecAbbrv = "OP"
                Case Is = "South Fort Meade"
                    gGetMineAreaSpecAbbrv = "SF"
                Case Is = "Wingate"
                    gGetMineAreaSpecAbbrv = "WG"
                Case Is = "Four Corners"
                    gGetMineAreaSpecAbbrv = "FC"
                Case Is = "Fort Green"
                    gGetMineAreaSpecAbbrv = "FG"
                Case Is = "Exploration"
                    gGetMineAreaSpecAbbrv = "EX"
                Case Is = "Ona"   'Eventually should be none of these!
                    gGetMineAreaSpecAbbrv = "ON"
                Case Is = ""
                    gGetMineAreaSpecAbbrv = "  "
                Case Else
                    gGetMineAreaSpecAbbrv = "?"
            End Select
        End If

        If aMode = "Area name" Then
            Select Case aName
                Case Is = "Altman"
                    gGetMineAreaSpecAbbrv = "Al"
                Case Is = "Brewster Remnants"
                    gGetMineAreaSpecAbbrv = "Br"
                Case Is = "Carlton"
                    gGetMineAreaSpecAbbrv = "Ca"
                Case Is = "Cytec"
                    gGetMineAreaSpecAbbrv = "Cy"
                Case Is = "Debris"
                    gGetMineAreaSpecAbbrv = "De"
                Case Is = "Fort Meade Hardee"
                    gGetMineAreaSpecAbbrv = "FH"
                Case Is = "Fort Meade Polk"
                    gGetMineAreaSpecAbbrv = "FP"
                Case Is = "HKP Proper"
                    gGetMineAreaSpecAbbrv = "Hp"
                Case Is = "Keys Desoto"
                    gGetMineAreaSpecAbbrv = "Kd"
                Case Is = "Keys Manatee"
                    gGetMineAreaSpecAbbrv = "Km"
                Case Is = "Lonesome"
                    gGetMineAreaSpecAbbrv = "Lo"
                Case Is = "Manson Jenkins"
                    gGetMineAreaSpecAbbrv = "Mj"
                Case Is = "MissChem"
                    gGetMineAreaSpecAbbrv = "Mc"
                Case Is = "NE Manatee"
                    gGetMineAreaSpecAbbrv = "Nm"
                Case Is = "Ona"
                    gGetMineAreaSpecAbbrv = "On"
                Case Is = "Ona Extension"
                    gGetMineAreaSpecAbbrv = "Oe"
                Case Is = "Ona test"
                    gGetMineAreaSpecAbbrv = "Ot"
                Case Is = "PIO 20 X 20"
                    gGetMineAreaSpecAbbrv = "P2"
                Case Is = "PNL Desoto"
                    gGetMineAreaSpecAbbrv = "Pd"
                Case Is = "PNL Manatee"
                    gGetMineAreaSpecAbbrv = "Pm"
                Case Is = "Payne Creek"
                    gGetMineAreaSpecAbbrv = "Pc"
                Case Is = "Pioneer"
                    gGetMineAreaSpecAbbrv = "Pi"
                Case Is = "Pioneer West"
                    gGetMineAreaSpecAbbrv = "Pw"
                Case Is = "S-1"
                    gGetMineAreaSpecAbbrv = "S1"
                Case Is = "SE Hillsborough"
                    gGetMineAreaSpecAbbrv = "Sh"
                Case Is = "SFM Fee Hardee"
                    gGetMineAreaSpecAbbrv = "Fh"
                Case Is = "SFM Fee Polk"
                    gGetMineAreaSpecAbbrv = "Fp"
                Case Is = "SFM Lease Hardee"
                    gGetMineAreaSpecAbbrv = "Lh"
                Case Is = "SFM Lease Polk"
                    gGetMineAreaSpecAbbrv = "Lp"
                Case Is = "SFTG Proper"
                    gGetMineAreaSpecAbbrv = "Sp"
                Case Is = "Texaco"
                    gGetMineAreaSpecAbbrv = "Tx"
                Case Is = "WIN"
                    gGetMineAreaSpecAbbrv = "Wi"
                Case Is = ""
                    gGetMineAreaSpecAbbrv = "  "
                Case Else
                    gGetMineAreaSpecAbbrv = "?"
            End Select
        End If
    End Function

    'Public Sub gPrintImcDb2RawProsp(ByVal aTwp As Integer, _
    '                                ByVal aRge As Integer, _
    '                                ByVal aSec As Integer, _
    '                                ByVal aHole As String, _
    '                                ByVal aSplNum As Integer, _
    '                                ByRef aReport As CrystalReport)

    ''**********************************************************************
    ''
    ''
    ''
    ''**********************************************************************

    '    On Error GoTo gPrintImcDb2RawProspError

    '    Dim ConnectString As String
    '    Dim HoleId As String
    '    Dim SplNumStr As String
    '    Dim Forty As Integer
    '    Dim SubNum As Integer
    '    Dim SubIdx As Integer

    '    Forty = gGetForty(aHole, "NUM")

    '    SplNumStr = Format(aSplNum, "0#")

    '    HoleId = "013" & Format(aTwp, "0#") & Format(aRge, "0#") & _
    '             Format(aSec, "0#") & Format(Forty, "0#") & aHole


    '    'Reporting application = Seagate Crystal Reports Professional
    '    aReport.ReportFileName = gPath + "\Reports\" + _
    '                             "IndSampRawProspImcRs.rpt"

    '    'Connect to Oracle database
    '    ConnectString = "DSN = " + gDataSource + ";UID = " + gOracleUserName + _
    '                    ";PWD = " + gOracleUserPassword + ";DSQ = "

    '    aReport.Connect = ConnectString

    '    'Need to pass the company name into the report
    '    aReport.ParameterFields(0) = "pCompanyName;" & gCompanyName & ";TRUE"
    '    aReport.ParameterFields(1) = "pHcd2Hole;" & HoleId & ";TRUE"
    '    aReport.ParameterFields(2) = "pScd2Seqn;" & SplNumStr & ";TRUE"

    '    SubNum = aReport.GetNSubreports
    '    If SubNum > 0 Then
    '        For SubIdx = 0 To SubNum - 1
    '            aReport.SubreportToChange = aReport.GetNthSubreportName(SubIdx)
    '            aReport.Connect = ConnectString
    '        Next
    '        aReport.SubreportToChange = ""
    '    End If

    '    'Report window maximized
    '    aReport.WindowState = crptMaximized

    '    aReport.WindowTitle = "IMC/DB2 Raw Prospect Data"

    '    'User allowed to minimize report window
    '    aReport.WindowMinButton = True

    '    'Start Crystal Reports
    '    aReport.action = 1

    '    Exit Sub

    'gPrintImcDb2RawProspError:
    '    MsgBox "Error printing report." & vbCrLf & _
    '           Err.Description, _
    '           vbOKOnly + vbExclamation, _
    '           "Report Printing Error"
    'End Sub

    Public Function gGetDrillHoleDateSpec(ByVal aSection As Integer,
                                          ByVal aTownship As Integer,
                                          ByVal aRange As Integer,
                                          ByVal aHoleLocation As String,
                                          ByVal aOmitRedrills As Integer,
                                          ByRef aSampleDynaset As OraDynaset) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetDrillHoleDateSpecError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer

        gGetDrillHoleDateSpec = False

        params = gDBParams

        params.Add("pSection", aSection, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTownship, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRange, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHoleLocation, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pOmitRedrills", aOmitRedrills, ORAPARM_INPUT)
        params("pOmitRedrills").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_raw_hole_data
        'pTownship              IN     NUMBER,
        'pRange                 IN     NUMBER,
        'pSection               IN     NUMBER,
        'pHoleLocation          IN     VARCHAR2,
        'pOmitRedrills          IN     NUMBER,
        'pResult                IN OUT c_prosprawsplit);
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospectnew2.get_prosp_raw_hole_data(:pTownship," +
                  ":pRange, :pSection, :pHoleLocation, :pOmitRedrills, " +
                  ":pResult);end;", ORASQL_FAILEXEC)
        aSampleDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = aSampleDynaset.RecordCount
        gGetDrillHoleDateSpec = True

        Exit Function

gGetDrillHoleDateSpecError:
        MsgBox("Error getting data for this prospect hole." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Data Access Error")

        On Error Resume Next
        ClearParams(params)
        gGetDrillHoleDateSpec = False
    End Function

    Public Function gCalcSolids(ByVal aWetWt As Single,
                                ByVal aDryWt As Single,
                                ByVal aTareWt As Single,
                                ByVal aRoundVal As Integer) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo CalcSolidsError

        Dim PctSolids As Single

        PctSolids = 0

        '02/11/2010, lss
        'Added this functionality!  Previously it would have returned 100%
        'If aWetWt or aDryWt is <= 0 then return zero (not 100!)
        'If aTareWt < 0 then return 0 (Tare weight can be zero).
        If aWetWt <= 0 Or aDryWt <= 0 Or aTareWt < 0 Then
            gCalcSolids = 0
            Exit Function
        End If

        If aWetWt - aTareWt > 0 Then
            PctSolids = Round((aDryWt - aTareWt) /
                        (aWetWt - aTareWt), 4)
        Else
            PctSolids = 0
        End If

        gCalcSolids = Round(100 * PctSolids, aRoundVal)

        Exit Function

CalcSolidsError:
        MsgBox("Error calculating %solids." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Calculate %Solid Error")
    End Function

    Public Function gCalcSolids2(ByVal aWetWt As Single,
                                 ByVal aDryWt As Single,
                                 ByVal aTareWt As Single,
                                 ByVal aWetWt2 As Single,
                                 ByVal aDryWt2 As Single,
                                 ByVal aTareWt2 As Single,
                                 ByVal aRoundVal As Integer) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo CalcSolids2Error

        Dim PctSolids As Single
        Dim PctSolids1 As Single
        Dim PctSolids2 As Single

        PctSolids = 0
        PctSolids1 = 0
        PctSolids2 = 0

        'If problems with #2 sample measure data then process as if a 1 measure
        'mtx %moisture situation
        'If aWetWt2 or aDryWt2 is <= 0 then use gCalcSolids.
        'If aTareWt2 < 0 then use gCalcSolids.
        If aWetWt2 <= 0 Or aDryWt2 <= 0 Or aTareWt2 < 0 Then
            gCalcSolids2 = gCalcSolids(aWetWt, aDryWt, aTareWt, 4)
            Exit Function
        End If

        'Assume that we have a full set of #1 and #2 mtx %moisture data values!
        'Process #1 Mtx moisture data values
        If aWetWt - aTareWt > 0 Then
            PctSolids1 = Round((aDryWt - aTareWt) /
                        (aWetWt - aTareWt), 4)
        Else
            PctSolids1 = 0
        End If

        'Process #2 Mtx moisture data values
        If aWetWt2 - aTareWt2 > 0 Then
            PctSolids2 = Round((aDryWt2 - aTareWt2) /
                        (aWetWt2 - aTareWt2), 4)
        Else
            PctSolids2 = 0
        End If

        'Now average the two percents appropriately
        PctSolids = Round((PctSolids1 + PctSolids2) / 2, 4)

        gCalcSolids2 = Round(100 * PctSolids, aRoundVal)

        Exit Function

CalcSolids2Error:
        MsgBox("Error calculating 2 sample %solids." & vbCrLf &
               Err.Description,
               vbOKOnly + vbExclamation,
               "Calculate 2 Sample %Solids Error")
    End Function


End Module
