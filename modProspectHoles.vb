Option Explicit On
Imports OracleInProcServer
Imports System.Math
Module modProspectHoles
    'Attribute VB_Name = "modProspectHoles"
    '**********************************************************************
    'PROSPECT HOLES MODULE
    '
    'Special Comments
    '----------------
    '1) Prospect hole location display:
    '   gGetHoleLocation
    '   gGetHoleLocationTitled
    '   gGetHoleLocationShort
    '   gGetHoleLocationShortDot
    '
    '**********************************************************************
    '   Maintenance Log
    '
    '   01/25/2001, lss
    '       Added this module.
    '   01/29/2001, lss
    '       Continued working on this module.
    '   10/17/2001, lss
    '       Added .MtxDryDensity to gProspectComposite
    '       Added .MtxConcentrateX to gProspectComposite
    '       Added .TotalConcentrateX to gProspectComposite
    '   04/25/2002, lss
    '       Added gGetHoleLocation.
    '   05/16/2002, lss
    '       Added gCompHoleExists, gGetHoleLocationShort.
    '   05/20/2002, lss
    '       Added gGetHoleLocationShortDot.
    '   05/28/2002, lss
    '       Added gProductX.
    '   06/19/2002, lss
    '       Added gHoleLocationOk.
    '   07/02/2002, lss
    '       Added gGetNumSplits.
    '   07/23/2003, lss
    '       Added gGetDrillHole.
    '   07/24/2003, lss
    '       Added gGetAnalTotal.
    '   10/22/2003, lss
    '       Fixed gCompHoleExists.
    '   12/11/2003, lss
    '       Modified gGetCompositeData to handle null .SplitsSummarized
    '       and null .ProspectorCode.
    '   12/12/2003, lss
    '       Modified gGetCompositeData for new Pioneer and Wingate items.
    '   01/19/2004, lss
    '       Added CfWtp, FfWtp, TfWtp correction to gGetSplitData.
    '   02/02/2004, lss
    '       Added gGetTotalValue.
    '   03/10/2004, lss
    '       Added isnull() check for AnalysisCdate for composites.
    '   06/16/2004, lss
    '      Modified gGetDrillHole for .SplitSum = Null.
    '   06/17/2004, lss
    '      Fixed divide by zero problem in gGetCompositeData().
    '   07/14/2004, lss
    '      Added gGetDrillDate().
    '   09/15/2004, lss
    '      Added gLoadProspCombos().
    '      Added gParseHoleLoc().
    '   10/22/2004, lss
    '      Added gGetHoleLoc2().
    '   12/13/2004, lss
    '      Added Cd and Hardpan code to gProspectComposite.
    '   04/12/2005, lss
    '      Changed Round in gGetTotalValue to gRound.
    '   05/05/2005, lss
    '      Added Function GetNumSplitsInHole.
    '   05/11/2005, lss
    '      Added Function gGetAllHoleLocList().
    '   05/16/2005, lss
    '      Added Function gHoleLocationOkNew().
    '      Added Function gGetClosestHole().
    '   05/17/2005, lss
    '      Added Public Type gCompBase.
    '   07/06/2005, lss
    '      Modified gGetCompositeData and gGetSplitData -- added
    '      aProspStandard.
    '   07/07/2005, lss
    '      Added aProspStandard to gCompHoleExists.
    '   07/08/2005, lss
    '      Added aProspStandard to Function gGetNumSplits.
    '      Added aProspStandard to Function gGetNumSplitsInHole.
    '      Added aProspStandard to Function gGetClosestHole.
    '      Added aProspStandard to Sub gGetDrillHole.
    '   07/15/2005, lss
    '      Modified gGetSplit for pTfBpl and pTfBplCalc.
    '   10/26/2005, lss
    '      Added WstThck As Double to gProspectCompositeType2.
    '      Added WstThck assignment to gGetDrillHole.
    '   05/23/2006, lss
    '      Added Public Function gGetMtxx.
    '      Added Public Function gGetMer.
    '      Added Public Type gCalcSplitType.
    '   06/28/2006, lss
    '      Added aDisplayError to Function gGetNumSplitsInHole.
    '   09/25/2006, lss
    '      Added Function gNumHoleLocIsValid.
    '   09/27/2006, lss
    '      Modified Case statement in Function gGetForty -- it was not
    '      correct!
    '   10/31/2006, lss
    '      Added Public Function gGetHoleLocationTitled.
    '   11/02/2006, lss
    '      Modified gGetDrillHole -- will now not average any 0 values!
    '   12/19/2006, lss
    '      Added Public Function gGetHoleLocationTrs.
    '   04/20/2007, lss
    '      Added Public Function gGetTotalValue4 -- will handle 4 values.
    '      Added Public Function gGetTotalValue2 -- will handle 2 values.
    '   07/24/2007, lss
    '      Added Function gGetHoleLocationTitled2.
    '   08/22/2007, lss
    '      Added "Special parameter fix" to Function gGetNumSplits.
    '   08/27/2007, lss
    '      Fixed Function gGetNumSplits -- had OUTPUT for some of the
    '      parameters when it should have been INPUT.
    '      Changed Types gProspectComposite and gProspectSplit from
    '      Private to Public.
    '   08/27/2007, lss
    '      Added Public Type gProspCompType.
    '      Added Public Type gCompCalcType.
    '   08/28/2007, lss
    '      Added matrix BPL correction in Function gGetSplitData.  The
    '      matrix "X" BPL from GEOCOMP is not ususally correct.  The lab
    '      does not run a waste clay BPL.
    '      Added WstThk to gCompCalcType.
    '   10/15/2007, lss
    '      Added Function gGetDepth1stSplitRaw.
    '      Added Sub gGetHoleCoordElev.
    '   11/16/2007, lss
    '      Added Public Sub gGetHoleMoisExist.
    '   12/18/2007, lss
    '      Handled Null AnalysisCdate in gGetDrillHole.
    '   12/21/2007, lss
    '      Added Public Function gGetAvgValue2.
    '   03/31/2008, lss
    '      Added Public Sub gGetHoleDateAndNumSplits.
    '   04/10/2008, lss
    '      Remarked out Public Type gProspCompType -- not really used
    '      anymore.
    '   04/15/2008, lss
    '      gProspectComposite, gProspectSplit -- changed some Doubles to
    '      Singles.  Also gProspectCompositeType2.
    '   04/15/2008, lss
    '      Added Public Sub gGetTonsFromProspSplits
    '      Added Private Function GetSplitsWclMtxTpa
    '      Added Private Sub GetSplTpasForThk
    '   05/20/2008, lss
    '      Changed Type gCompBase to Type gComBaseType.
    '   05/21/2008, lss
    '      Added Public Sub gGetNeighborSection and related stuff.
    '   08/20/2008, lss
    '      Added Public Function gGetCompositeBase.
    '   08/21/2008, lss
    '      Added Public Function gGetSplitBase.
    '   08/28/2008, lss
    '      Modified Function gGetDrillDate to handle mm/yy type dates.
    '   09/03/2008, lss
    '      Added Public Sub gPrintComposite.
    '   04/15/2010, lss
    '      Added Function gGetHoleLocationTitled3.
    '   04/27/2010, lss
    '      Added Public Function gIsHalfHole.
    '   06/21/2010, lss
    '      Added Public Function gGetTotalValue3.
    '   10/20/2010, lss
    '      Added Public Function gGetHoleInSfmHardee.
    '   01/19/2011, lss
    '      Added Public Function gGetMerAt.
    '   10/02/2013, lss
    '      Fixed the top of seam elevation gGetSplitData.
    '
    '**********************************************************************


    Public Structure gProspectComposite
        Public Mine As String                      '1
        Public Section As Integer                  '2
        Public Township As Integer                 '3
        Public Range As Integer                    '4
        Public HoleLocation As String              '5
        Public DrillCdate As String                '6
        Public AnalysisCdate As String             '7
        Public AreaOfInfluence As Single           '8
        Public HoleElevation As Single             '9
        Public PitBottomElevation As Single        '10
        Public XSPCoordinate As Double             '11
        Public YSPCoordinate As Double             '12
        Public TriangleCode As String              '13
        Public TotalNumberSplits As Integer        '14
        Public SplitsSummarized As String          '15
        Public ProspectorCode As String            '16
        Public OvbThickness As Single              '17
        Public OvbX As Single                      '18
        Public MtxThickness As Single              '19
        Public MtxPebbleX As Single                '20
        Public MtxX As Single                      '21
        Public TotalThickness As Single            '22
        Public TotalPebbleX As Single              '23
        Public TotalX As Single                    '24
        Public MtxPercentSolids As Single          '25
        Public MtxWetDensity As Single             '26
        Public CoarsePebbleWtp As Single           '27
        Public FinePebbleWtp As Single             '28
        Public TotalPebbleWtp As Single            '29
        Public ConcentrateWtp As Single            '30
        Public TotalProductWtp As Single           '31
        Public TotalTailWtp As Single              '32
        Public WasteClayWtp As Single              '33
        Public GrossConcentrateWtp As Single       '34
        Public GrossProductWtp As Single           '35
        Public CoarseFeedWtp As Single             '36
        Public FineFeedWtp As Single               '37
        Public TotalFeedWtp As Single              '38
        Public MtxTons As Double                   '39
        Public CoarsePebbleTPA As Double           '40
        Public FinePebbleTPA As Double             '41
        Public TotalPebbleTpa As Double            '42
        Public ConcentrateTPA As Double            '43
        Public TotalProductTpa As Double           '44
        Public TotalTailTpa As Double              '45
        Public WasteClayTpa As Double              '46
        Public GrossConcentrateTpa As Double       '47
        Public GrossProductTpa As Double           '48
        Public CoarseFeedTpa As Double             '49
        Public FineFeedTpa As Double               '50
        Public TotalFeedTpa As Double              '51
        Public MtxBPL As Single                    '52
        Public CoarsePebbleBPL As Single           '53
        Public FinePebbleBPL As Single             '54
        Public TotalPebbleBpl As Single            '55
        Public ConcentrateBPL As Single            '56
        Public TotalProductBpl As Single           '57
        Public TotalTailBPL As Single              '58
        Public WasteClayBPL As Single              '59
        Public GrossConcentrateBpl As Single       '60
        Public GrossProductBpl As Single           '61
        Public CoarseFeedBpl As Single             '62
        Public FineFeedBpl As Single               '63
        Public TotalFeedBpl As Single              '64
        Public FinePebbleFe2O3 As Single           '65
        Public FinePebbleAl2O3 As Single           '66
        Public FinePebbleMgO As Single             '67
        Public FinePebbleCaO As Single             '68
        Public FinePebbleInsol As Single           '69
        Public FinePebbleIa As Single              '70
        Public CoarsePebbleFe2O3 As Single         '71
        Public CoarsePebbleAl2O3 As Single         '72
        Public CoarsePebbleMgO As Single           '73
        Public CoarsePebbleCaO As Single           '74
        Public CoarsePebbleInsol As Single         '75
        Public CoarsePebbleIa As Single            '76
        Public TotalPebbleFe2O3 As Single          '77
        Public TotalPebbleAl2O3 As Single          '78
        Public TotalPebbleMgO As Single            '79
        Public TotalPebbleCaO As Single            '80
        Public TotalPebbleInsol As Single          '81
        Public TotalPebbleIa As Single             '82
        Public ConcentrateFe2O3 As Single          '83
        Public ConcentrateAl2O3 As Single          '84
        Public ConcentrateMgO As Single            '85
        Public ConcentrateCaO As Single            '86
        Public ConcentrateInsol As Single          '87
        Public ConcentrateIA As Single             '88
        Public TotalProductFe2O3 As Single         '89
        Public TotalProductAl2O3 As Single         '90
        Public TotalProductMgO As Single           '91
        Public TotalProductCaO As Single           '92
        Public TotalProductInsol As Single         '93
        Public TotalProductIA As Single            '94
        Public GrossConcentrateInsol As Single     '95
        Public GrossProductInsol As Single         '96
        '--------
        Public MtxDryDensity As Single             '97  -- calculated in this proc
        Public MtxConcentrateX As Single           '98  -- calculated in this proc
        Public TotalConcentrateX As Single         '99  -- calculated in this proc
        '----------
        'The following were added 12/12/2003 for Pioneer, lss
        Public WstThck As Single                   '100
        Public TotX As Single                      '101
        Public MinableSplits As String             '102
        Public HoleMinable As String               '103
        Public CpbMinable As String                '104
        Public FpbMinable As String                '105
        Public CpbFeTpaWt As Double                '106
        Public CpbAlTpaWt As Double                '107
        Public CpbIaTpaWt As Double                '108
        Public CpbCaTpaWt As Double                '109
        Public FpbFeTpaWt As Double                '110
        Public FpbAlTpaWt As Double                '111
        Public FpbIaTpaWt As Double                '112
        Public FpbCaTpaWt As Double                '113
        Public CnFeTpaWt As Double                 '114
        Public CnAlTpaWt As Double                 '115
        Public CnIaTpaWt As Double                 '116
        Public CnCaTpaWt As Double                 '117
        Public CpIa As Single                      '118
        Public FpIa As Single                      '119
        Public CnIa As Single                      '120
        Public TpIA As Single                      '121
        Public TpbIA As Single                     '122
        Public FltBplRcvryCalc As Single           '123
        Public MtxYdsPerAcre As Single             '124
        Public Rc As Single                        '125
        Public HasExtraData As Boolean             '126
        Public WstPbWtp As Single                  '127
        Public WstPbTpa As Double                  '128
        Public WstPbBpl As Single                  '129
        Public WstPbIns As Single                  '130
        Public WstPbFe As Single                   '131
        Public WstPbAl As Single                   '132
        Public WstPbMg As Single                   '133
        Public WstPbCa As Single                   '134
        Public WstPbIa1 As Single                  '135
        Public WstPbIa2 As Single                  '136
        '----------
        'The following were added 12/13/2004, lss
        Public CpbCd As Single                     '137
        Public FpbCd As Single                     '138
        Public TcnCd As Single                     '139
        Public TprCd As Single                     '140
        Public TpbCd As Single                     '141
        Public HardpanCode As Integer              '142
        '----------
        Public ProspStandard As String             '143
    End Structure
    Public gComposite As gProspectComposite

    Public Structure gProspectSplit
        Public Mine As String                     '1
        Public Section As Integer                 '2
        Public Township As Integer                '3
        Public Range As Integer                   '4
        Public HoleLocation As String             '5
        Public Split As Integer                   '6
        Public MinableStatus As String            '7
        Public SplitThickness As Single           '8
        Public DrillCdate As String               '9
        Public WashCdate As String                '10
        Public AreaOfInfluence As Single          '11
        Public ProspectorCode As String           '12
        Public TopOfSplitDepth As Single          '13
        Public BotOfSplitDepth As Single          '14
        Public SampleNumber As String             '15
        Public TotalNumberSplits As Integer       '16
        Public RatioOfConc As Single              '17
        Public CountyCode As String               '18
        Public MiningCode As String               '19
        Public PumpingCode As String              '20
        Public MetLabID As String                 '21
        Public ChemLabID As String                '22
        Public Color As String                    '23
        Public SplitElevation As Single           '24
        Public HoleNumber As Integer              '25
        Public SampleNumber2 As String            '26
        Public MtxX As Single                     '27
        Public CalcLossPercent As Double          '28
        Public CalcLossTPA As Double              '29
        Public CalcLossBPL As Double              '30
        Public WetMtxLbs As Double                '31
        Public MtxGmsWet As Double                '32
        Public MtxGmsDry As Double                '33
        Public PercentSolidsMtx As Double         '34
        Public WetFeedLbs As Double               '35
        Public FeedMoistWetGms As Double          '36
        Public FeedMoistDryGms As Double          '37
        Public TriangleCode As String             '38
        Public CpWtp As Single                    '39
        Public FpWtp As Single                    '40
        Public TfWtp As Single                    '41
        Public WcWtp As Single                    '42
        Public CnWtp As Single                    '43
        Public TpWtp As Single                    '44
        Public FfWtp As Single                    '45
        Public CfWtp As Single                    '46
        Public CpTpa As Double                    '47
        Public FpTpa As Double                    '48
        Public TfTPA As Double                    '49
        Public WcTpa As Double                    '50
        Public CnTpa As Double                    '51
        Public TpTpa As Double                    '52
        Public FfTpa As Double                    '53
        Public CfTpa As Double                    '54
        Public FatTpa As Double                   '55
        Public AtTpa As Double                    '56
        Public AcnTpa As Double                   '57
        Public MtxTPA As Double                   '58
        Public CpBPL As Single                    '59
        Public FpBpl As Single                    '60
        Public LcnBpl As Single                   '61
        Public TpBpl As Single                    '62
        Public MtxBPL As Single                   '63
        Public CnBpl As Single                    '64
        Public CpInsol As Single                  '65
        Public CpFe2O3 As Single                  '66
        Public CpAl2O3 As Single                  '67
        Public CpMgO As Single                    '68
        Public CpCaO As Single                    '69
        Public FpInsol As Single                  '70
        Public FpFe2O3 As Single                  '71
        Public FpAl2O3 As Single                  '72
        Public FpMgO As Single                    '73
        Public FpCaO As Single                    '74
        Public FpFeAl As Single                   '75
        Public FpCaOP2O5 As Single                '76
        Public LcnInsol As Single                 '77
        Public LcnFe2O3 As Single                 '78
        Public LcnAl2O3 As Single                 '79
        Public LcnMgO As Single                   '80
        Public LcnCaO As Single                   '81
        Public LcnFeAl As Single                  '82
        Public TpInsol As Single                  '83
        Public TpFe2O3 As Single                  '84
        Public TpAl2O3 As Single                  '85
        Public TpMgO As Single                    '86
        Public TpCaO As Single                    '87
        Public TpFeAl As Single                   '88
        Public TpCaOP2O5 As Single                '89
        Public MtxInsol As Single                 '90
        Public CnInsol As Single                  '91
        Public CnCaO As Single                    '92
        Public CnFeAl As Single                   '93
        Public CpGrams As Double                  '94
        Public FpGrams As Double                  '95
        Public TfGrams As Double                  '96
        Public FatGrams As Double                 '97
        Public AtGrams As Double                  '98
        Public LcnGrams As Double                 '99
        Public WetDensityVolume As Double         '100
        Public WetDensityWeight As Double         '101
        Public WetDensity As Single               '102
        Public DryDensityVolume As Double         '103
        Public DryDensityWeight As Double         '104
        Public DryDensity As Single               '105
        Public CalcHeadFeedBpl As Single          '106
        Public GrossConcentrateTpa As Double      '107
        Public GrossConcentrateBpl As Single      '108
        Public GrossConcentrateInsol As Single    '109
        Public GrossPebbleWtp As Single           '110
        Public GrossPebbleTPA As Double           '111
        Public GrossPebbleBPL As Single           '112
        Public GrossPebbleInsol As Single         '113
        Public GrossPebbleFe2O3 As Single         '114
        Public GrossPebbleAl2O3 As Single         '115
        Public GrossPebbleMgO As Single           '116
        Public ConcentrateWtp As Single           '117
        Public ConcentrateTPA As Double           '118
        Public ConcentrateBPL As Single           '119
        Public ConcentrateInsol As Single         '120
        Public ConcentrateFe2O3 As Single         '121
        Public ConcentrateAl2O3 As Single         '122
        Public ConcentrateMgO As Single           '123
        Public TotalProductWtp As Single          '124
        Public TotalProductTpa As Double          '125
        Public TotalProductBpl As Single          '126
        Public TotalProductInsol As Single        '127
        Public TotalProductFe2O3 As Single        '128
        Public TotalProductAl2O3 As Single        '129
        Public TotalProductMgO As Single          '130
        Public TotalTailWtp As Single             '131
        Public TotalTailTpa As Double             '132
        Public TotalTailBPL As Single             '133
        Public ColorTrans As String               '134
        Public MinableStatusTrans As String       '135
        Public CalcOvbX As Single                 '136
        Public CalcMtxX As Single                 '137
        Public CalcTotalX As Single               '138
        Public FlotationRC As Single              '139
        Public FlotationRecovery As Single        '140
        Public TopOfSeamElevation As Single       '141
        Public FfBpl As Single                    '142
        Public CfBpl As Single                    '143
        Public TfBPL As Single                    '144   Measured value
        Public FatBpl As Single                   '145
        Public AtBpl As Single                    '146
        Public WcBpl As Single                    '147
        '----------
        Public HardpanCode As Integer             '148
        Public GrossPebbleCaO As Single           '149
        Public ConcentrateCaO As Single           '150
        Public TotalProductCaO As Single          '151
        Public CpCd As Single                     '152
        Public FpCd As Single                     '153
        Public LcnCd As Single                    '154
        Public ConcentrateCd As Single            '155
        Public TotalProductCd As Single           '156
        Public TpCd As Single                     '157
        Public GrossPebbleCd As Single            '158
        '----------
        Public TfBplCalc As Single                '159
        Public ProspStandard As String            '160
    End Structure
    Public gSplit As gProspectSplit

    'Composite2 -- uses get_composite2
    Public Structure gProspectCompositeType2
        Public MineName As String              '1
        Public HoleLocation As String          '2
        Public Section As Integer              '3
        Public Township As Integer             '4
        Public Range As Integer                '5
        Public XSpCdnt As Double               '6
        Public YSpCdnt As Double               '7
        Public DrillCdate As String            '8
        Public AnalysisCdate As String         '9
        Public AreaInfluence As Double         '10
        Public OvbThck As Double               '11
        Public MtxThck As Double               '12
        Public MtxWetDensity As Double         '13
        Public MtxPctSolids As Double          '14
        Public MtxX As Double                  '15
        Public HoleElevation As Double         '16
        Public SplitTotalNum As Double         '17
        Public PitBottomElevation As Double    '18
        Public TriangleCode As String          '19
        Public ProspCode As String             '20
        Public MtxTons As Double               '21
        Public SplitSum As String              '22
        Public CpbBpl As Single                '23
        Public FpbBpl As Single                '24
        Public TfdBpl As Single                '25
        Public CfdBpl As Single                '26
        Public FfdBpl As Single                '27
        Public CncBpl As Single                '28
        Public TpbBpl As Single                '29
        Public TprBpl As Single                '30
        Public MtxBPL As Single                '31
        Public CpbTpa As Double                '32
        Public FpbTpa As Double                '33
        Public TfdTpa As Double                '34
        Public CfdTpa As Double                '35
        Public FfdTpa As Double                '36
        Public CncTpa As Double                '37
        Public TpbTpa As Double                '38
        Public TprTpa As Double                '39
        Public WclTpa As Double                '40
        Public CpbWtp As Single                '41
        Public FpbWtp As Single                '42
        Public TfdWtp As Single                '43
        Public CncWtp As Single                '44
        Public TpbWtp As Single                '45
        Public TprWtp As Single                '46
        Public WclWtp As Single                '47
        Public CpbIns As Single                '48
        Public FpbIns As Single                '49
        Public TpbIns As Single                '50
        Public CncIns As Single                '51
        Public TprIns As Single                '52
        Public CpbFe As Single                 '53
        Public FpbFe As Single                 '54
        Public TpbFe As Single                 '55
        Public CncFe As Single                 '56
        Public TprFe As Single                 '57
        Public CpbAl As Single                 '58
        Public FpbAl As Single                 '59
        Public TpbAl As Single                 '60
        Public CncAl As Single                 '61
        Public TprAl As Single                 '62
        Public CpbMg As Single                 '63
        Public FpbMg As Single                 '64
        Public TpbMg As Single                 '65
        Public CncMg As Single                 '66
        Public TprMg As Single                 '67
        Public CpbCa As Single                 '68
        Public FpbCa As Single                 '69
        Public TpbCa As Single                 '70
        Public CncCa As Single                 '71
        Public TprCa As Single                 '72
        '----------
        Public TfdBpl2 As Single               '72
        Public TfdTpa2 As Single               '73
        Public TfdWtp2 As Single               '74
        '----------
        Public TpbBpl2 As Single               '75
        Public TpbFe2 As Single                '76
        Public TpbAl2 As Single                '77
        Public TpbMg2 As Single                '78
        Public TpbIns2 As Single               '79
        Public TpbCa2 As Single                '80   '10/08/2012, lss  New
        '----------
        Public ProspStandard As String         '81
        Public WstThck As Single               '82
    End Structure

    Public Structure gCompBaseType
        Public MineName As String
        Public Section As Integer
        Public Township As Integer
        Public Range As Integer
        Public HoleLoc As String
        Public Xcoord As Double
        Public Ycoord As Double
        Public Elevation As Single
        Public ProspDate As String
        Public OvbThk As Single
        Public MtxThk As Single
        Public WstThk As Single
        Public TotNumSplits As Integer
    End Structure

    Public Structure gSplitBaseType
        Public MineName As String
        Public Section As Integer
        Public Township As Integer
        Public Range As Integer
        Public HoleLoc As String
        Public ProspDate As String
        Public Split As Integer
        Public SplitDepthTop As Single
        Public SplitDepthBot As Single
        Public SplitThk As Single
        Public MinableStatus As String
    End Structure

    Public Structure gCalcSplitType
        Public LcnBpl As Single
        Public LcnIns As Single
        Public LcnFe As Single
        Public LcnAl As Single
        Public LcnMg As Single
        Public LcnCa As Single

        Public CfBpl As Single
        Public CfTpa As Single
        Public FfBpl As Single
        Public FfTpa As Single

        Public TfBPL As Single
        Public TfTPA As Single

        Public FpbBpl As Single
        Public FpbIns As Single
        Public FpbFe As Single
        Public FpbAl As Single
        Public FpbMg As Single
        Public FpbCa As Single
        Public FpbTpa As Single

        Public CpbBpl As Single
        Public CpbIns As Single
        Public CpbFe As Single
        Public CpbAl As Single
        Public CpbMg As Single
        Public CpbCa As Single
        Public CpbTpa As Single

        Public AdjInsTarg As Single
        Public AdjIns As Single

        Public CnBpl As Single
        Public CnBplChnge As Single
        Public TotTlBpl As Single
        Public FlotRC As Single

        Public CnTpa As Single
        Public CnIns As Single
        Public CnFe As Single
        Public CnAl As Single
        Public CnMg As Single
        Public CnCa As Single

        Public GpbTpa As Single
        Public GpbBpl As Single
        Public GpbIns As Single
        Public GpbFe As Single
        Public GpbAl As Single
        Public GpbMg As Single
        Public GpbCa As Single

        Public TpTpa As Single
        Public TpBpl As Single
        Public TpIns As Single
        Public TpFe As Single
        Public TpAl As Single
        Public TpMg As Single
        Public TpCa As Single

        Public TfBplCalc As Single
        Public AcnTpa As Single
    End Structure

    Dim mProspHoleInfo As OraDynaset

    '08/27/207, lss
    'Moved this type from frmPdHoleAdd to here -- changed it from
    'fProspCompType to gProspCompType.
    '04/10/2008, lss
    'Remarked this type out -- not used anymore.
    'Public Type gProspCompType
    '    MineName As String                  '1
    '    Township As Integer                 '2
    '    Range As Integer                    '3
    '    Section As Integer                  '4
    '    XSPCoordinate As Double             '5
    '    YSPCoordinate As Double             '6
    '    HoleLocation As String              '7
    '    DrillDate As String                 '8
    '    WashDate As String                  '9
    '    AreaOfInfluence As Double           '10
    '    OvbThickness As Single              '11
    '    MtxThickness As Single              '12
    '    MtxWetDensity As Double             '13
    '    MtxPercentSolids As Double          '14
    '    MtxTons As Double                   '15
    '    MtxBPL As Single                    '16
    '    CpWeightPercent As Single           '17
    '    CpTonsPerAcre As Double             '18
    '    CpBPL As Single                     '19
    '    CpInsol As String                   '20
    '    CpFe2O3 As Single                   '21
    '    CpAl2O3 As Single                   '22
    '    CpMgO As Single                     '23
    '    CpCaO As Single                     '24
    '    FpWeightPercent As Single           '25
    '    FpTonsPerAcre As Double             '26
    '    FpBPL As Single                     '27
    '    FpInsol As String                   '28
    '    FpFe2O3 As Single                   '29
    '    FpAl2O3 As Single                   '30
    '    FpMgO As Single                     '31
    '    FpCaO As Single                     '32
    '    TfWeightPercent As Single           '33
    '    TfTonsPerAcre As Double             '34
    '    TfBPL As Single                     '35
    '    WcWeightPercent As Single           '36
    '    WcTonsPerAcre As Double             '37
    '    CfBpl As Single                     '38
    '    FfBpl As Single                     '39
    '    CfTonsPerAcre As Double             '40
    '    FfTonsPerAcre As Double             '41
    '    CnWeightPercent As Single           '42
    '    CnTonsPerAcre As Double             '43
    '    CnBpl As Single                     '44
    '    CnInsol As String                   '45
    '    CnFe2O3 As Single                   '46
    '    CnAl2O3 As Single                   '47
    '    CnMgO As Single                     '48
    '    CnCaO As Single                     '49
    '    TpWeightPercent As Single           '50
    '    TpTonsPerAcre As Double             '51
    '    TpBpl As Single                     '52
    '    TpInsol As String                   '53
    '    TpFe2O3 As Single                   '54
    '    TpAl2O3 As Single                   '55
    '    TpMgO As Single                     '56
    '    TpCaO As Single                     '57
    '    MtxX As Double                      '58
    '    HoleElevation As Single             '59
    '    TotalNumberSplits As Single         '60
    '    PitBottomElevation As Single        '61
    '    TriangleCode As String              '62
    '    ProspectorCode As String            '63
    '    SplitsSummarized As String          '64
    '    TpbWeightPercent As Single          '65
    '    TpbTonsPerAcre As Double            '66
    '    TpbBpl As Single                    '67
    '    TpbInsol As String                  '68
    '    TpbFe2O3 As Single                  '69
    '    TpbAl2O3 As Single                  '70
    '    TpbMgO As Single                    '71
    '    TpbCaO As Single                    '72
    'End Type

    '08/27/2007, lss
    'Move this type from frmCompositeSplits to here -- changed it from
    'fCompCalcType to gCompCalcType.
    Public Structure gCompCalcType
        Public OvbThk As Single
        Public MtxThk As Single
        Public WstThk As Single
        '----------
        Public MtxTPA As Double
        '----------
        Public CrsPbWtPct As Single
        Public CrsPbTpa As Double
        Public CrsPbBpl As Single
        Public CrsPbIns As Single
        Public CrsPbFe As Single
        Public CrsPbAl As Single
        Public CrsPbMg As Single
        Public CrsPbCa As Single
        Public SumCrsPbTpa As Double
        Public SumCrsPbBplTpa As Double
        Public SumCrsPbInsTpa As Double
        Public SumCrsPbFeTpa As Double
        Public SumCrsPbAlTpa As Double
        Public SumCrsPbMgTpa As Double
        Public SumCrsPbCaTpa As Double
        Public SumCrsPbTpaWithBpl As Double
        Public SumCrsPbTpaWithIns As Double
        Public SumCrsPbTpaWithFe As Double
        Public SumCrsPbTpaWithAl As Double
        Public SumCrsPbTpaWithMg As Double
        Public SumCrsPbTpaWithCa As Double
        '----------
        Public FnePbWtPct As Single
        Public FnePbTpa As Double
        Public FnePbBpl As Single
        Public FnePbIns As Single
        Public FnePbFe As Single
        Public FnePbAl As Single
        Public FnePbMg As Single
        Public FnePbCa As Single
        Public SumFnePbTpa As Double
        Public SumFnePbBplTpa As Double
        Public SumFnePbInsTpa As Double
        Public SumFnePbFeTpa As Double
        Public SumFnePbAlTpa As Double
        Public SumFnePbMgTpa As Double
        Public SumFnePbCaTpa As Double
        Public SumFnePbTpaWithBpl As Double
        Public SumFnePbTpaWithIns As Double
        Public SumFnePbTpaWithFe As Double
        Public SumFnePbTpaWithAl As Double
        Public SumFnePbTpaWithMg As Double
        Public SumFnePbTpaWithCa As Double
        '----------
        Public CrsFdTpa As Double
        Public CrsFdBpl As Single
        Public SumCrsFdTpa As Double
        Public SumCrsFdBplTpa As Double
        Public SumCrsFdTpaWithBpl As Double
        '----------
        Public FneFdTpa As Double
        Public FneFdBpl As Single
        Public SumFneFdTpa As Double
        Public SumFneFdBplTpa As Double
        Public SumFneFdTpaWithBpl As Double
        '----------
        Public WclWtPct As Single
        Public WclTpa As Double
        '----------
        Public CnWtPct As Single
        Public CnTpa As Double
        Public CnBpl As Single
        Public CnIns As Single
        Public CnFe As Single
        Public CnAl As Single
        Public CnMg As Single
        Public CnCa As Single
        Public SumCnTpa As Double
        Public SumCnBplTpa As Double
        Public SumCnInsTpa As Double
        Public SumCnFeTpa As Double
        Public SumCnAlTpa As Double
        Public SumCnMgTpa As Double
        Public SumCnCaTpa As Double
        Public SumCnTpaWithBpl As Double
        Public SumCnTpaWithIns As Double
        Public SumCnTpaWithFe As Double
        Public SumCnTpaWithAl As Double
        Public SumCnTpaWithMg As Double
        Public SumCnTpaWithCa As Double
        '----------
        Public TotFdWtPct As Single
        Public TotFdTpa As Double
        Public TotFdBpl As Single
        '----------
        Public TotProdWtPct As Single
        Public TotProdTpa As Double
        Public TotProdBpl As Single
        Public TotProdIns As Single
        Public TotProdFe As Single
        Public TotProdAl As Single
        Public TotProdMg As Single
        Public TotProdCa As Single
        Public SumTpTpa As Double
        Public SumTpBplTpa As Double
        Public SumTpInsTpa As Double
        Public SumTpFeTpa As Double
        Public SumTpAlTpa As Double
        Public SumTpMgTpa As Double
        Public SumTpCaTpa As Double
        Public SumTpTpaWithBpl As Double
        Public SumTpTpaWithIns As Double
        Public SumTpTpaWithFe As Double
        Public SumTpTpaWithAl As Double
        Public SumTpTpaWithMg As Double
        Public SumTpTpaWithCa As Double
        '----------
        Public MtxX As Single
        Public TotX As Single
        Public TotNumSplits As Integer
        Public ListOfSplits As String
        '----------
        Public TotPbWtPct As Single
        Public TotPbTpa As Single
        Public TotPbBpl As Single
        Public TotPbIns As Single
        Public TotPbFe As Single
        Public TotPbAl As Single
        Public TotPbMg As Single
        Public TotPbCa As Single
        '----------
        Public SumCubicFtMtx As Double
        Public SumWetLbsMtx As Double
        Public SumDryLbsMtx As Double
        Public MtxWetDens As Single
        Public MtxDryDens As Single
        Public MtxPctSolids As Single
        '----------
        Public CrsPbIa As Single
        Public FnePbIa As Single
        Public CnIa As Single
        Public TotProdIA As Single
        Public TotPbIa As Single
    End Structure

    Public Function gGetCompositeData(ByVal aMineName As String, _
                                      ByVal aSec As Integer, _
                                      ByVal aTwp As Integer, _
                                      ByVal aRge As Integer, _
                                      ByVal aHole As String, _
                                      ByVal aDisplayMissingError As Boolean, _
                                      ByVal aProspStandard As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        Dim TonsWithAnalysis As Double
        Dim AnalysisTons As Double

        On Error GoTo gGetCompositeDataError

        ZeroCompositeData()

        'Get prospect data
        params = gDBParams

        '1
        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        '2
        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        '3
        params.Add("pTownShip", aTwp, ORAPARM_INPUT)
        params("pTownShip").serverType = ORATYPE_NUMBER

        '4
        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        '5  VARCHAR2
        params.Add("pHoleLocation", aHole, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        '6
        params.Add("pProspStandard", aProspStandard, ORAPARM_INPUT)
        params("pProspStandard").serverType = ORATYPE_VARCHAR2

        '7  VARCHAR2
        params.Add("pDrillCdate", "", ORAPARM_OUTPUT)
        params("pDrillCdate").serverType = ORATYPE_VARCHAR2

        '8  VARCHAR2
        params.Add("pAnalysisCdate", "", ORAPARM_OUTPUT)
        params("pAnalysisCdate").serverType = ORATYPE_VARCHAR2

        '9
        params.Add("pAreaOfInfluence", 0, ORAPARM_OUTPUT)
        params("pAreaOfInfluence").serverType = ORATYPE_NUMBER

        '10
        params.Add("pHoleElevation", 0, ORAPARM_OUTPUT)
        params("pHoleElevation").serverType = ORATYPE_NUMBER

        '11
        params.Add("pPitBottomElevation", 0, ORAPARM_OUTPUT)
        params("pPitBottomElevation").serverType = ORATYPE_NUMBER

        '12
        params.Add("pXSPCoordinate", 0, ORAPARM_OUTPUT)
        params("pXSPCoordinate").serverType = ORATYPE_NUMBER

        '13
        params.Add("pYSPCoordinate", 0, ORAPARM_OUTPUT)
        params("pYSPCoordinate").serverType = ORATYPE_NUMBER

        '14  VARCHAR2
        params.Add("pTriangleCode", "", ORAPARM_OUTPUT)
        params("pTriangleCode").serverType = ORATYPE_VARCHAR2

        '15
        params.Add("pTotalNumberSplits", 0, ORAPARM_OUTPUT)
        params("pTotalNumberSplits").serverType = ORATYPE_NUMBER

        '16  VARCHAR2
        params.Add("pSplitsSummarized", "", ORAPARM_OUTPUT)
        params("pSplitsSummarized").serverType = ORATYPE_VARCHAR2

        '17  VARCHAR2
        params.Add("pProspectorCode", "", ORAPARM_OUTPUT)
        params("pProspectorCode").serverType = ORATYPE_VARCHAR2

        '18
        params.Add("pOvbThickness", 0, ORAPARM_OUTPUT)
        params("pOvbThickness").serverType = ORATYPE_NUMBER

        '19
        params.Add("pOvbX", 0, ORAPARM_OUTPUT)
        params("pOvbX").serverType = ORATYPE_NUMBER

        '20
        params.Add("pMtxThickness", 0, ORAPARM_OUTPUT)
        params("pMtxThickness").serverType = ORATYPE_NUMBER

        '21
        params.Add("pMtxPebbleX", 0, ORAPARM_OUTPUT)
        params("pMtxPebbleX").serverType = ORATYPE_NUMBER

        '22
        params.Add("pMtxX", 0, ORAPARM_OUTPUT)
        params("pMtxX").serverType = ORATYPE_NUMBER

        '23
        params.Add("pTotalThickness", 0, ORAPARM_OUTPUT)
        params("pTotalThickness").serverType = ORATYPE_NUMBER

        '24
        params.Add("pTotalPebbleX", 0, ORAPARM_OUTPUT)
        params("pTotalPebbleX").serverType = ORATYPE_NUMBER

        '25
        params.Add("pTotalX", 0, ORAPARM_OUTPUT)
        params("pTotalX").serverType = ORATYPE_NUMBER

        '26
        params.Add("pMtxPercentSolids", 0, ORAPARM_OUTPUT)
        params("pMtxPercentSolids").serverType = ORATYPE_NUMBER

        '27
        params.Add("pMtxWetDensity", 0, ORAPARM_OUTPUT)
        params("pMtxWetDensity").serverType = ORATYPE_NUMBER

        '28
        params.Add("pCoarsePebbleWtp", 0, ORAPARM_OUTPUT)
        params("pCoarsePebbleWtp").serverType = ORATYPE_NUMBER

        '29
        params.Add("pFinePebbleWtp", 0, ORAPARM_OUTPUT)
        params("pFinePebbleWtp").serverType = ORATYPE_NUMBER

        '30
        params.Add("pTotalPebbleWtp", 0, ORAPARM_OUTPUT)
        params("pTotalPebbleWtp").serverType = ORATYPE_NUMBER

        '31
        params.Add("pConcentrateWtp", 0, ORAPARM_OUTPUT)
        params("pConcentrateWtp").serverType = ORATYPE_NUMBER

        '32
        params.Add("pTotalProductWtp", 0, ORAPARM_OUTPUT)
        params("pTotalProductWtp").serverType = ORATYPE_NUMBER

        '33
        params.Add("pTotalTailWtp", 0, ORAPARM_OUTPUT)
        params("pTotalTailWtp").serverType = ORATYPE_NUMBER

        '34
        params.Add("pWasteClayWtp", 0, ORAPARM_OUTPUT)
        params("pWasteClayWtp").serverType = ORATYPE_NUMBER

        '35
        params.Add("pGrossConcentrateWtp", 0, ORAPARM_OUTPUT)
        params("pGrossConcentrateWtp").serverType = ORATYPE_NUMBER

        '36
        params.Add("pGrossProductWtp", 0, ORAPARM_OUTPUT)
        params("pGrossProductWtp").serverType = ORATYPE_NUMBER

        '37
        params.Add("pCoarseFeedWtp", 0, ORAPARM_OUTPUT)
        params("pCoarseFeedWtp").serverType = ORATYPE_NUMBER

        '38
        params.Add("pFineFeedWtp", 0, ORAPARM_OUTPUT)
        params("pFineFeedWtp").serverType = ORATYPE_NUMBER

        '39
        params.Add("pTotalFeedWtp", 0, ORAPARM_OUTPUT)
        params("pTotalFeedWtp").serverType = ORATYPE_NUMBER

        '40
        params.Add("pMtxTons", 0, ORAPARM_OUTPUT)
        params("pMtxTons").serverType = ORATYPE_NUMBER

        '41
        params.Add("pCoarsePebbleTPA", 0, ORAPARM_OUTPUT)
        params("pCoarsePebbleTPA").serverType = ORATYPE_NUMBER

        '42
        params.Add("pFinePebbleTPA", 0, ORAPARM_OUTPUT)
        params("pFinePebbleTPA").serverType = ORATYPE_NUMBER

        '43
        params.Add("pTotalPebbleTPA", 0, ORAPARM_OUTPUT)
        params("pTotalPebbleTPA").serverType = ORATYPE_NUMBER

        '44
        params.Add("pConcentrateTPA", 0, ORAPARM_OUTPUT)
        params("pConcentrateTPA").serverType = ORATYPE_NUMBER

        '45
        params.Add("pTotalProductTPA", 0, ORAPARM_OUTPUT)
        params("pTotalProductTPA").serverType = ORATYPE_NUMBER

        '46
        params.Add("pTotalTailTPA", 0, ORAPARM_OUTPUT)
        params("pTotalTailTPA").serverType = ORATYPE_NUMBER

        '47
        params.Add("pWasteClayTPA", 0, ORAPARM_OUTPUT)
        params("pWasteClayTPA").serverType = ORATYPE_NUMBER

        '48
        params.Add("pGrossConcentrateTPA", 0, ORAPARM_OUTPUT)
        params("pGrossConcentrateTPA").serverType = ORATYPE_NUMBER

        '49
        params.Add("pGrossProductTPA", 0, ORAPARM_OUTPUT)
        params("pGrossProductTPA").serverType = ORATYPE_NUMBER

        '50
        params.Add("pCoarseFeedTPA", 0, ORAPARM_OUTPUT)
        params("pCoarseFeedTPA").serverType = ORATYPE_NUMBER

        '51
        params.Add("pFineFeedTPA", 0, ORAPARM_OUTPUT)
        params("pFineFeedTPA").serverType = ORATYPE_NUMBER

        '52
        params.Add("pTotalFeedTPA", 0, ORAPARM_OUTPUT)
        params("pTotalFeedTPA").serverType = ORATYPE_NUMBER

        '53
        params.Add("pMtxBPL", 0, ORAPARM_OUTPUT)
        params("pMtxBPL").serverType = ORATYPE_NUMBER

        '54
        params.Add("pCoarsePebbleBPL", 0, ORAPARM_OUTPUT)
        params("pCoarsePebbleBPL").serverType = ORATYPE_NUMBER

        '55
        params.Add("pFinePebbleBPL", 0, ORAPARM_OUTPUT)
        params("pFinePebbleBPL").serverType = ORATYPE_NUMBER

        '56
        params.Add("pTotalPebbleBPL", 0, ORAPARM_OUTPUT)
        params("pTotalPebbleBPL").serverType = ORATYPE_NUMBER

        '57
        params.Add("pConcentrateBPL", 0, ORAPARM_OUTPUT)
        params("pConcentrateBPL").serverType = ORATYPE_NUMBER

        '58
        params.Add("pTotalProductBPL", 0, ORAPARM_OUTPUT)
        params("pTotalProductBPL").serverType = ORATYPE_NUMBER

        '59
        params.Add("pTotalTailBPL", 0, ORAPARM_OUTPUT)
        params("pTotalTailBPL").serverType = ORATYPE_NUMBER

        '60
        params.Add("pWasteClayBPL", 0, ORAPARM_OUTPUT)
        params("pWasteClayBPL").serverType = ORATYPE_NUMBER

        '61
        params.Add("pGrossConcentrateBPL", 0, ORAPARM_OUTPUT)
        params("pGrossConcentrateBPL").serverType = ORATYPE_NUMBER

        '62
        params.Add("pGrossProductBPL", 0, ORAPARM_OUTPUT)
        params("pGrossProductBPL").serverType = ORATYPE_NUMBER

        '63
        params.Add("pCoarseFeedBPL", 0, ORAPARM_OUTPUT)
        params("pCoarseFeedBPL").serverType = ORATYPE_NUMBER

        '64
        params.Add("pFineFeedBPL", 0, ORAPARM_OUTPUT)
        params("pFineFeedBPL").serverType = ORATYPE_NUMBER

        '65
        params.Add("pTotalFeedBPL", 0, ORAPARM_OUTPUT)
        params("pTotalFeedBPL").serverType = ORATYPE_NUMBER

        '66
        params.Add("pFinePebbleFe2O3", 0, ORAPARM_OUTPUT)
        params("pFinePebbleFe2O3").serverType = ORATYPE_NUMBER

        '67
        params.Add("pFinePebbleAl2O3", 0, ORAPARM_OUTPUT)
        params("pFinePebbleAl2O3").serverType = ORATYPE_NUMBER

        '68
        params.Add("pFinePebbleMgO", 0, ORAPARM_OUTPUT)
        params("pFinePebbleMgO").serverType = ORATYPE_NUMBER

        '69
        params.Add("pFinePebbleCaO", 0, ORAPARM_OUTPUT)
        params("pFinePebbleCaO").serverType = ORATYPE_NUMBER

        '70
        params.Add("pFinePebbleInsol", 0, ORAPARM_OUTPUT)
        params("pFinePebbleInsol").serverType = ORATYPE_NUMBER

        '71
        params.Add("pFinePebbleIA", 0, ORAPARM_OUTPUT)
        params("pFinePebbleIA").serverType = ORATYPE_NUMBER

        '72
        params.Add("pCoarsePebbleFe2O3", 0, ORAPARM_OUTPUT)
        params("pCoarsePebbleFe2O3").serverType = ORATYPE_NUMBER

        '73
        params.Add("pCoarsePebbleAl2O3", 0, ORAPARM_OUTPUT)
        params("pCoarsePebbleAl2O3").serverType = ORATYPE_NUMBER

        '74
        params.Add("pCoarsePebbleMgO", 0, ORAPARM_OUTPUT)
        params("pCoarsePebbleMgO").serverType = ORATYPE_NUMBER

        '75
        params.Add("pCoarsePebbleCaO", 0, ORAPARM_OUTPUT)
        params("pCoarsePebbleCaO").serverType = ORATYPE_NUMBER

        '76
        params.Add("pCoarsePebbleInsol", 0, ORAPARM_OUTPUT)
        params("pCoarsePebbleInsol").serverType = ORATYPE_NUMBER

        '77
        params.Add("pCoarsePebbleIA", 0, ORAPARM_OUTPUT)
        params("pCoarsePebbleIA").serverType = ORATYPE_NUMBER

        '78
        params.Add("pTotalPebbleFe2O3", 0, ORAPARM_OUTPUT)
        params("pTotalPebbleFe2O3").serverType = ORATYPE_NUMBER

        '79
        params.Add("pTotalPebbleAl2O3", 0, ORAPARM_OUTPUT)
        params("pTotalPebbleAl2O3").serverType = ORATYPE_NUMBER

        '80
        params.Add("pTotalPebbleMgO", 0, ORAPARM_OUTPUT)
        params("pTotalPebbleMgO").serverType = ORATYPE_NUMBER

        '81
        params.Add("pTotalPebbleCaO", 0, ORAPARM_OUTPUT)
        params("pTotalPebbleCaO").serverType = ORATYPE_NUMBER

        '82
        params.Add("pTotalPebbleInsol", 0, ORAPARM_OUTPUT)
        params("pTotalPebbleInsol").serverType = ORATYPE_NUMBER

        '83
        params.Add("pTotalPebbleIA", 0, ORAPARM_OUTPUT)
        params("pTotalPebbleIA").serverType = ORATYPE_NUMBER

        '84
        params.Add("pConcentrateFe2O3", 0, ORAPARM_OUTPUT)
        params("pConcentrateFe2O3").serverType = ORATYPE_NUMBER

        '85
        params.Add("pConcentrateAl2O3", 0, ORAPARM_OUTPUT)
        params("pConcentrateAl2O3").serverType = ORATYPE_NUMBER

        '86
        params.Add("pConcentrateMgO", 0, ORAPARM_OUTPUT)
        params("pConcentrateMgO").serverType = ORATYPE_NUMBER

        '87
        params.Add("pConcentrateCaO", 0, ORAPARM_OUTPUT)
        params("pConcentrateCaO").serverType = ORATYPE_NUMBER

        '88
        params.Add("pConcentrateInsol", 0, ORAPARM_OUTPUT)
        params("pConcentrateInsol").serverType = ORATYPE_NUMBER

        '89
        params.Add("pConcentrateIA", 0, ORAPARM_OUTPUT)
        params("pConcentrateIA").serverType = ORATYPE_NUMBER

        '90
        params.Add("pTotalProductFe2O3", 0, ORAPARM_OUTPUT)
        params("pTotalProductFe2O3").serverType = ORATYPE_NUMBER

        '91
        params.Add("pTotalProductAl2O3", 0, ORAPARM_OUTPUT)
        params("pTotalProductAl2O3").serverType = ORATYPE_NUMBER

        '92
        params.Add("pTotalProductMgO", 0, ORAPARM_OUTPUT)
        params("pTotalProductMgO").serverType = ORATYPE_NUMBER

        '93
        params.Add("pTotalProductCaO", 0, ORAPARM_OUTPUT)
        params("pTotalProductCaO").serverType = ORATYPE_NUMBER

        '94
        params.Add("pTotalProductInsol", 0, ORAPARM_OUTPUT)
        params("pTotalProductInsol").serverType = ORATYPE_NUMBER

        '95
        params.Add("pTotalProductIA", 0, ORAPARM_OUTPUT)
        params("pTotalProductIA").serverType = ORATYPE_NUMBER

        '96
        params.Add("pGrossConcentrateInsol", 0, ORAPARM_OUTPUT)
        params("pGrossConcentrateInsol").serverType = ORATYPE_NUMBER

        '97
        params.Add("pGrossProductInsol", 0, ORAPARM_OUTPUT)
        params("pGrossProductinsol").serverType = ORATYPE_NUMBER

        'New items added 12/12/2003, lss

        '98
        params.Add("pWstThck", 0, ORAPARM_OUTPUT)
        params("pWstThck").serverType = ORATYPE_NUMBER

        '99
        params.Add("pTotX", 0, ORAPARM_OUTPUT)
        params("pTotX").serverType = ORATYPE_NUMBER

        '100
        params.Add("pMinableSplits", 0, ORAPARM_OUTPUT)
        params("pMinableSplits").serverType = ORATYPE_VARCHAR2

        '101
        params.Add("pHoleMinable", 0, ORAPARM_OUTPUT)
        params("pHoleMinable").serverType = ORATYPE_VARCHAR2

        '102
        params.Add("pCpbMinable", 0, ORAPARM_OUTPUT)
        params("pCpbMinable").serverType = ORATYPE_VARCHAR2

        '103
        params.Add("pFpbMinable", 0, ORAPARM_OUTPUT)
        params("pFpbMinable").serverType = ORATYPE_VARCHAR2

        '104
        params.Add("pCpbFeTpaWt", 0, ORAPARM_OUTPUT)
        params("pCpbFeTpaWt").serverType = ORATYPE_NUMBER

        '105
        params.Add("pCpbAlTpaWt", 0, ORAPARM_OUTPUT)
        params("pCpbAlTpaWt").serverType = ORATYPE_NUMBER

        '106
        params.Add("pCpbIaTpaWt", 0, ORAPARM_OUTPUT)
        params("pCpbIaTpaWt").serverType = ORATYPE_NUMBER

        '107
        params.Add("pCpbCaTpaWt", 0, ORAPARM_OUTPUT)
        params("pCpbCaTpaWt").serverType = ORATYPE_NUMBER

        '108
        params.Add("pFpbFeTpaWt", 0, ORAPARM_OUTPUT)
        params("pFpbFeTpaWt").serverType = ORATYPE_NUMBER

        '109
        params.Add("pFpbAlTpaWt", 0, ORAPARM_OUTPUT)
        params("pFpbAlTpaWt").serverType = ORATYPE_NUMBER

        '110
        params.Add("pFpbIaTpaWt", 0, ORAPARM_OUTPUT)
        params("pFpbIaTpaWt").serverType = ORATYPE_NUMBER

        '111
        params.Add("pFpbCaTpaWt", 0, ORAPARM_OUTPUT)
        params("pFpbCaTpaWt").serverType = ORATYPE_NUMBER

        '112
        params.Add("pCnFeTpaWt", 0, ORAPARM_OUTPUT)
        params("pCnFeTpaWt").serverType = ORATYPE_NUMBER

        '113
        params.Add("pCnAlTpaWt", 0, ORAPARM_OUTPUT)
        params("pCnAlTpaWt").serverType = ORATYPE_NUMBER

        '114
        params.Add("pCnIaTpaWt", 0, ORAPARM_OUTPUT)
        params("pCnIaTpaWt").serverType = ORATYPE_NUMBER

        '115
        params.Add("pCnCaTpaWt", 0, ORAPARM_OUTPUT)
        params("pCnCaTpaWt").serverType = ORATYPE_NUMBER

        '116
        params.Add("pCpIa", 0, ORAPARM_OUTPUT)
        params("pCpIa").serverType = ORATYPE_NUMBER

        '117
        params.Add("pFpIa", 0, ORAPARM_OUTPUT)
        params("pCpIa").serverType = ORATYPE_NUMBER

        '118
        params.Add("pCnIa", 0, ORAPARM_OUTPUT)
        params("pCpIa").serverType = ORATYPE_NUMBER

        '119
        params.Add("pTpIa", 0, ORAPARM_OUTPUT)
        params("pCpIa").serverType = ORATYPE_NUMBER

        '120
        params.Add("pTpbIa", 0, ORAPARM_OUTPUT)
        params("pCpIa").serverType = ORATYPE_NUMBER

        '121
        params.Add("pCpbCd", 0, ORAPARM_OUTPUT)
        params("pCpbCd").serverType = ORATYPE_NUMBER

        '122
        params.Add("pFpbCd", 0, ORAPARM_OUTPUT)
        params("pFpbCd").serverType = ORATYPE_NUMBER

        '123
        params.Add("pTcnCd", 0, ORAPARM_OUTPUT)
        params("pTcnCd").serverType = ORATYPE_NUMBER

        '124
        params.Add("pTprCd", 0, ORAPARM_OUTPUT)
        params("pTprCd").serverType = ORATYPE_NUMBER

        '125
        params.Add("pTpbCd", 0, ORAPARM_OUTPUT)
        params("pTpbCd").serverType = ORATYPE_NUMBER

        '126
        params.Add("pHardpanCode", 0, ORAPARM_OUTPUT)
        params("pHardpanCode").serverType = ORATYPE_NUMBER

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect.get_composite(:pMineName, :pSection," + _
                      ":pTownship, :pRange, :pHoleLocation, :pProspStandard, :pDrillCdate, :pAnalysisCdate, :pAreaOfInfluence, :pHoleElevation," + _
                      ":pPitBottomElevation, :pXSPCoordinate, :pYSPCoordinate, :pTriangleCode, :pTotalNumberSplits, :pSplitsSummarized," + _
                      ":pProspectorCode, :pOvbThickness, :pOvbX, :pMtxThickness, :pMtxPebbleX, :pMtxX, :pTotalThickness, :pTotalPebbleX," + _
                      ":pTotalX, :pMtxPercentSolids, :pMtxWetDensity, :pCoarsePebbleWtp, :pFinePebbleWtp, :pTotalPebbleWtp," + _
                      ":pConcentrateWtp, :pTotalProductWtp, :pTotalTailWtp, :pWasteClayWtp, :pGrossConcentrateWtp, :pGrossProductWtp," + _
                      ":pCoarseFeedWtp, :pFineFeedWtp, :pTotalFeedWtp, :pMtxTons, :pCoarsePebbleTPA, :pFinePebbleTPA, :pTotalPebbleTPA," + _
                      ":pConcentrateTPA, :pTotalProductTPA, :pTotalTailTPA, :pWasteClayTPA, :pGrossConcentrateTPA, :pGrossProductTPA," + _
                      ":pCoarseFeedTPA, :pFineFeedTPA, :pTotalFeedTPA, :pMtxBPL, :pCoarsePebbleBPL, :pFinePebbleBPL, :pTotalPebbleBPL," + _
                      ":pConcentrateBPL, :pTotalProductBPL, :pTotalTailBPL, :pWasteClayBPL, :pGrossConcentrateBPL, :pGrossProductBPL, :pCoarseFeedBPL," + _
                      ":pFineFeedBPL, :pTotalFeedBPL, :pFinePebbleFe2O3, :pFinePebbleAl2O3, :pFinePebbleMgO," + _
                      ":pFinePebbleCaO, :pFinePebbleInsol, :pFinePebbleIA, :pCoarsePebbleFe2O3, :pCoarsePebbleAl2O3, :pCoarsePebbleMgO," + _
                      ":pCoarsePebbleCaO, :pCoarsePebbleInsol, :pCoarsePebbleIA, :pTotalPebbleFe2O3, :pTotalPebbleAl2O3, :pTotalPebbleMgO," + _
                      ":pTotalPebbleCaO, :pTotalPebbleInsol, :pTotalPebbleIA, :pConcentrateFe2O3, :pConcentrateAl2O3, :pConcentrateMgO," + _
                      ":pConcentrateCaO, :pConcentrateInsol, :pConcentrateIA, :pTotalProductFe2O3, :pTotalProductAl2O3, :pTotalProductMgO," + _
                      ":pTotalProductCaO, :pTotalProductInsol, :pTotalProductIA, :pGrossConcentrateInsol, :pGrossProductInsol," + _
                      ":pWstThck, :pTotX, :pMinableSplits, :pHoleMinable, :pCpbMinable, :pFpbMinable, " + _
                      ":pCpbFeTpaWt, :pCpbAlTpaWt, :pCpbIaTpaWt, :pCpbCaTpaWt, :pFpbFeTpaWt, :pFpbAlTpaWt, :pFpbIaTpaWt, :pFpbCaTpaWt, " + _
                      ":pCnFeTpaWt, :pCnAlTpaWt, :pCnIaTpaWt, :pCnCaTpaWt, :pCpIa, :pFpIa, :pCnIa, :pTpIa, :pTpbIa, " + _
                      ":pCpbCd, :pFpbCd, :pTcnCd, :pTprCd, :pTpbCd, :pHardpanCode);end;", ORASQL_FAILEXEC)

        With gComposite
            .Mine = params("pMineName").Value                            '0
            .Section = params("pSection").Value                          '1
            .Township = params("pTownship").Value                        '2
            .Range = params("pRange").Value                              '3
            .HoleLocation = params("pHoleLocation").Value                '5
            .DrillCdate = params("pDrillCdate").Value                    '6

            If Not IsDBNull(params("pAnalysisCdate").Value) Then
                .AnalysisCdate = params("pAnalysisCdate").Value          '7
            Else
                .AnalysisCdate = ""
            End If

            .AreaOfInfluence = params("pAreaOfInfluence").Value          '8
            .HoleElevation = params("pHoleElevation").Value              '9
            .PitBottomElevation = params("pPitBottomElevation").Value    '10
            .XSPCoordinate = params("pXSPCoordinate").Value              '11
            .YSPCoordinate = params("pYSPCoordinate").Value              '12

            If Not IsDBNull(params("pTriangleCode").Value) Then
                .TriangleCode = params("pTriangleCode").Value            '13
            Else
                .TriangleCode = ""
            End If

            .TotalNumberSplits = params("pTotalNumberSplits").Value      '14

            If Not IsDBNull(params("pSplitsSummarized").Value) Then
                .SplitsSummarized = params("pSplitsSummarized").Value    '15
            Else
                .SplitsSummarized = ""
            End If

            If Not IsDBNull(params("pProspectorCode").Value) Then
                .ProspectorCode = params("pProspectorCode").Value        '16
            Else
                .ProspectorCode = ""
            End If

            .OvbThickness = params("pOvbThickness").Value                '17
            .OvbX = params("pOvbX").Value                                '18
            .MtxThickness = params("pMtxThickness").Value                '19
            .MtxPebbleX = params("pMtxPebbleX").Value                    '20
            .MtxX = params("pMtxX").Value                                '21
            .TotalThickness = params("pTotalThickness").Value            '22
            .TotalPebbleX = params("pTotalPebbleX").Value                '23
            .TotalX = params("pTotalX").Value                            '24
            .MtxPercentSolids = params("pMtxPercentSolids").Value        '25
            .MtxWetDensity = params("pMtxWetDensity").Value              '26
            .CoarsePebbleWtp = params("pCoarsePebbleWtp").Value          '27
            .FinePebbleWtp = params("pFinePebbleWtp").Value              '28
            .TotalPebbleWtp = params("pTotalPebbleWtp").Value            '29
            .ConcentrateWtp = params("pConcentrateWtp").Value            '30
            .TotalProductWtp = params("pTotalProductWtp").Value          '31
            .TotalTailWtp = params("pTotalTailWtp").Value                '32
            .WasteClayWtp = params("pWasteClayWtp").Value                '33
            .GrossConcentrateWtp = params("pGrossConcentrateWtp").Value  '34
            .GrossProductWtp = params("pGrossProductWtp").Value          '35
            .CoarseFeedWtp = params("pCoarseFeedWtp").Value              '36
            .FineFeedWtp = params("pFineFeedWtp").Value                  '37
            .TotalFeedWtp = params("pTotalFeedWtp").Value                '38
            .MtxTons = params("pMtxTons").Value                          '39
            .CoarsePebbleTPA = params("pCoarsePebbleTPA").Value          '40
            .FinePebbleTPA = params("pFinePebbleTPA").Value              '41
            .TotalPebbleTpa = params("pTotalPebbleTPA").Value            '42
            .ConcentrateTPA = params("pConcentrateTPA").Value            '43
            .TotalProductTpa = params("pTotalProductTPA").Value          '44
            .TotalTailTpa = params("pTotalTailTPA").Value                '45
            .WasteClayTpa = params("pWasteClayTPA").Value                '46
            .GrossConcentrateTpa = params("pGrossConcentrateTPA").Value  '47
            .GrossProductTpa = params("pGrossProductTPA").Value          '48
            .CoarseFeedTpa = params("pCoarseFeedTPA").Value              '49
            .FineFeedTpa = params("pFineFeedTPA").Value                  '50
            .TotalFeedTpa = params("pTotalFeedTPA").Value                '51
            .MtxBPL = params("pMtxBPL").Value                            '52
            .CoarsePebbleBPL = params("pCoarsePebbleBPL").Value          '53
            .FinePebbleBPL = params("pFinePebbleBPL").Value              '54
            .TotalPebbleBpl = params("pTotalPebbleBPL").Value            '55
            .ConcentrateBPL = params("pConcentrateBPL").Value            '56
            .TotalProductBpl = params("pTotalProductBPL").Value          '57
            .TotalTailBPL = params("pTotalTailBPL").Value                '58
            .WasteClayBPL = params("pWasteClayBPL").Value                '59
            .GrossConcentrateBpl = params("pGrossConcentrateBPL").Value  '60
            .GrossProductBpl = params("pGrossProductBPL").Value          '61
            .CoarseFeedBpl = params("pCoarseFeedBPL").Value              '62
            .FineFeedBpl = params("pFineFeedBPL").Value                  '63
            .TotalFeedBpl = params("pTotalFeedBPL").Value                '64
            .FinePebbleFe2O3 = params("pFinePebbleFe2O3").Value          '65
            .FinePebbleAl2O3 = params("pFinePebbleAl2O3").Value          '66
            .FinePebbleMgO = params("pFinePebbleMgO").Value              '67
            .FinePebbleCaO = params("pFinePebbleCaO").Value              '68
            .FinePebbleInsol = params("pFinePebbleInsol").Value          '69
            .FinePebbleIa = params("pFinePebbleIA").Value                '70
            .CoarsePebbleFe2O3 = params("pCoarsePebbleFe2O3").Value      '71
            .CoarsePebbleAl2O3 = params("pCoarsePebbleAl2O3").Value      '72
            .CoarsePebbleMgO = params("pCoarsePebbleMgO").Value          '73
            .CoarsePebbleCaO = params("pCoarsePebbleCaO").Value          '74
            .CoarsePebbleInsol = params("pCoarsePebbleInsol").Value      '75
            .CoarsePebbleIa = params("pCoarsePebbleIA").Value            '76
            .TotalPebbleFe2O3 = params("pTotalPebbleFe2O3").Value        '77
            .TotalPebbleAl2O3 = params("pTotalPebbleAl2O3").Value        '78
            .TotalPebbleMgO = params("pTotalPebbleMgO").Value            '79
            .TotalPebbleCaO = params("pTotalPebbleCaO").Value            '80
            .TotalPebbleInsol = params("pTotalPebbleInsol").Value        '81
            .TotalPebbleIa = params("pTotalPebbleIA").Value              '82
            .ConcentrateFe2O3 = params("pConcentrateFe2O3").Value        '83
            .ConcentrateAl2O3 = params("pConcentrateAl2O3").Value        '84
            .ConcentrateMgO = params("pConcentrateMgO").Value            '85
            .ConcentrateCaO = params("pConcentrateCaO").Value            '86
            .ConcentrateInsol = params("pConcentrateInsol").Value        '87
            .ConcentrateIA = params("pConcentrateIA").Value              '88
            .TotalProductFe2O3 = params("pTotalProductFe2O3").Value      '89
            .TotalProductAl2O3 = params("pTotalProductAl2O3").Value      '90
            .TotalProductMgO = params("pTotalProductMgO").Value          '91
            .TotalProductCaO = params("pTotalProductCaO").Value          '92
            .TotalProductInsol = params("pTotalProductInsol").Value      '93
            .TotalProductIA = params("pTotalProductIA").Value            '94
            .GrossConcentrateInsol = params("pGrossConcentrateInsol").Value  '95
            .GrossProductInsol = params("pGrossProductInsol").Value          '96

            'The following added 12/12/2003, lss
            .WstThck = params("pWstThck").Value                              '97
            .TotX = params("pTotX").Value                                    '98

            If Not IsDBNull(params("pMinableSplits").Value) Then
                .MinableSplits = params("pMinableSplits").Value              '99
            Else
                .MinableSplits = ""
            End If

            If Not IsDBNull(params("pHoleMinable").Value) Then
                .HoleMinable = params("pHoleMinable").Value                  '100
            Else
                .HoleMinable = ""
            End If

            If Not IsDBNull(params("pCpbMinable").Value) Then
                .CpbMinable = params("pCpbMinable").Value                    '101
            Else
                .CpbMinable = ""
            End If

            If Not IsDBNull(params("pFpbMinable").Value) Then
                .FpbMinable = params("pFpbMinable").Value                    '102
            Else
                .FpbMinable = ""
            End If

            .CpbFeTpaWt = params("pCpbFeTpaWt").Value                    '103
            .CpbAlTpaWt = params("pCpbAlTpaWt").Value                    '104
            .CpbIaTpaWt = params("pCpbIaTpaWt").Value                    '105
            .CpbCaTpaWt = params("pCpbCaTpaWt").Value                    '106
            .FpbFeTpaWt = params("pFpbFeTpaWt").Value                    '107
            .FpbAlTpaWt = params("pFpbAlTpaWt").Value                    '108
            .FpbIaTpaWt = params("pFpbIaTpaWt").Value                    '109
            .FpbCaTpaWt = params("pFpbCaTpaWt").Value                    '110
            .CnFeTpaWt = params("pCnFeTpaWt").Value                      '111
            .CnAlTpaWt = params("pCnAlTpaWt").Value                      '112
            .CnIaTpaWt = params("pCnIaTpaWt").Value                      '113
            .CnCaTpaWt = params("pCnCaTpaWt").Value                      '114
            .CpIa = params("pCpIa").Value                                '115
            .FpIa = params("pFpIa").Value                                '116
            .CnIa = params("pCnIa").Value                                '117
            .TpIA = params("pTpIa").Value                                '118
            .TpbIA = params("pTpbIa").Value                              '119

            .CpbCd = params("pCpbCd").Value                              '120
            .FpbCd = params("pFpbCd").Value                              '121
            .TcnCd = params("pTcnCd").Value                              '122
            .TprCd = params("pTprCd").Value                              '123
            .TpbCd = params("pTpbCd").Value                              '124
            .HardpanCode = params("pHardpanCode").Value                  '125

            .ProspStandard = params("pProspStandard").Value              '126
        End With

        'Do some additional calculations
        With gComposite
            If .CpbMinable = "M" Or .CpbMinable = "U" Then
                .HasExtraData = True
            Else
                .HasExtraData = False
            End If

            If .ConcentrateTPA <> 0 Then
                .Rc = Round(.TotalFeedTpa / .ConcentrateTPA, 2)
            Else
                .Rc = 0
            End If

            If .AreaOfInfluence <> 0 Then
                .MtxYdsPerAcre = Round((.MtxThickness * .AreaOfInfluence * 43560 / 27) / _
                                 .AreaOfInfluence, 0)
            Else
                .MtxYdsPerAcre = 0
            End If

            If .TotalFeedBpl * .TotalFeedTpa <> 0 Then
                .FltBplRcvryCalc = Round((.ConcentrateBPL * .ConcentrateTPA) / _
                                   (.TotalFeedBpl * .TotalFeedTpa) * 100, 1)
            Else
                .FltBplRcvryCalc = 0
            End If

            If .HasExtraData = False Then
                .TotX = .TotalX
            End If
            '----------

            .MtxDryDensity = Round(.MtxWetDensity * (.MtxPercentSolids / 100), 2)

            If .ConcentrateTPA <> 0 Then
                .MtxConcentrateX = Round((.MtxThickness * 43560 / 27) / _
                                   .ConcentrateTPA, 2)
            Else
                .MtxConcentrateX = 0
            End If

            If .ConcentrateTPA <> 0 Then
                .TotalConcentrateX = Round((.TotalThickness * 43560 / 27) / _
                                   .ConcentrateTPA, 2)
            Else
                .TotalConcentrateX = 0
            End If

            If .Mine = "Pioneer" Or .Mine = "WingateX" Then
                'Calculate the waste pebble -- for Pioneer & Wingate
                .WstPbTpa = .CoarsePebbleTPA + .FinePebbleTPA
                .WstPbWtp = .CoarsePebbleWtp + .FinePebbleWtp

                'Waste pebble BPL
                TonsWithAnalysis = 0
                AnalysisTons = 0
                AnalysisTons = .CoarsePebbleBPL * .CoarsePebbleTPA + _
                               .FinePebbleBPL * .FinePebbleTPA
                If .CoarsePebbleBPL <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .CoarsePebbleTPA
                End If
                If .FinePebbleBPL <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .FinePebbleTPA
                End If
                If TonsWithAnalysis <> 0 Then
                    .WstPbBpl = Round(AnalysisTons / TonsWithAnalysis, 2)
                Else
                    .WstPbBpl = 0
                End If

                'Waste pebble Insol
                TonsWithAnalysis = 0
                AnalysisTons = 0
                AnalysisTons = .CoarsePebbleInsol * .CoarsePebbleTPA + _
                               .FinePebbleInsol * .FinePebbleTPA
                If .CoarsePebbleInsol <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .CoarsePebbleTPA
                End If
                If .FinePebbleInsol <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .FinePebbleTPA
                End If
                If TonsWithAnalysis <> 0 Then
                    .WstPbIns = Round(AnalysisTons / TonsWithAnalysis, 2)
                Else
                    .WstPbIns = 0
                End If

                'Waste pebble Al2O3
                TonsWithAnalysis = 0
                AnalysisTons = 0
                AnalysisTons = .CoarsePebbleAl2O3 * .CoarsePebbleTPA + _
                               .FinePebbleAl2O3 * .FinePebbleTPA
                If .CoarsePebbleAl2O3 <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .CoarsePebbleTPA
                End If
                If .FinePebbleAl2O3 <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .FinePebbleTPA
                End If
                If TonsWithAnalysis <> 0 Then
                    .WstPbAl = Round(AnalysisTons / TonsWithAnalysis, 2)
                Else
                    .WstPbAl = 0
                End If

                'Waste pebble Fe2O3
                TonsWithAnalysis = 0
                AnalysisTons = 0
                AnalysisTons = .CoarsePebbleFe2O3 * .CoarsePebbleTPA + _
                               .FinePebbleFe2O3 * .FinePebbleTPA
                If .CoarsePebbleFe2O3 <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .CoarsePebbleTPA
                End If
                If .FinePebbleFe2O3 <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .FinePebbleTPA
                End If
                If TonsWithAnalysis <> 0 Then
                    .WstPbFe = Round(AnalysisTons / TonsWithAnalysis, 2)
                Else
                    .WstPbFe = 0
                End If

                'Waste pebble MgO
                TonsWithAnalysis = 0
                AnalysisTons = 0
                AnalysisTons = .CoarsePebbleMgO * .CoarsePebbleTPA + _
                               .FinePebbleMgO * .FinePebbleTPA
                If .CoarsePebbleMgO <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .CoarsePebbleTPA
                End If
                If .FinePebbleMgO <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .FinePebbleTPA
                End If
                If TonsWithAnalysis <> 0 Then
                    .WstPbMg = Round(AnalysisTons / TonsWithAnalysis, 2)
                Else
                    .WstPbMg = 0
                End If

                'Waste pebble CaO
                TonsWithAnalysis = 0
                AnalysisTons = 0
                AnalysisTons = .CoarsePebbleCaO * .CoarsePebbleTPA + _
                               .FinePebbleCaO * .FinePebbleTPA
                If .CoarsePebbleCaO <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .CoarsePebbleTPA
                End If
                If .FinePebbleCaO <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .FinePebbleTPA
                End If
                If TonsWithAnalysis <> 0 Then
                    .WstPbCa = Round(AnalysisTons / TonsWithAnalysis, 2)
                Else
                    .WstPbCa = 0
                End If

                'Waste pebble IA1
                TonsWithAnalysis = 0
                AnalysisTons = 0
                AnalysisTons = .CoarsePebbleIa * .CoarsePebbleTPA + _
                               .FinePebbleIa * .FinePebbleTPA
                If .CoarsePebbleIa <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .CoarsePebbleTPA
                End If
                If .FinePebbleIa <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .FinePebbleTPA
                End If
                If TonsWithAnalysis <> 0 Then
                    .WstPbIa1 = Round(AnalysisTons / TonsWithAnalysis, 2)
                Else
                    .WstPbIa1 = 0
                End If

                'Waste pebble IA2
                TonsWithAnalysis = 0
                AnalysisTons = 0
                AnalysisTons = .CpIa * .CoarsePebbleTPA + _
                               .FpIa * .FinePebbleTPA
                If .CpIa <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .CoarsePebbleTPA
                End If
                If .FpIa <> 0 Then
                    TonsWithAnalysis = TonsWithAnalysis + .FinePebbleTPA
                End If
                If TonsWithAnalysis <> 0 Then
                    .WstPbIa2 = Round(AnalysisTons / TonsWithAnalysis, 2)
                Else
                    .WstPbIa2 = 0
                End If

                'Revisit the IA's
                'Coarse pebble
                If .CoarsePebbleIa = 0 And .CpIa <> 0 Then
                    .CoarsePebbleIa = .CpIa
                End If

                'Fine pebble
                If .FinePebbleIa = 0 And .FpIa <> 0 Then
                    .FinePebbleIa = .FpIa
                End If

                'Concentrate
                If .ConcentrateIA = 0 And .CnIa <> 0 Then
                    .ConcentrateIA = .CnIa
                End If

                'Total product
                If .TotalProductIA = 0 And .TpIA <> 0 Then
                    .TotalProductIA = .TpIA
                End If

                'Total pebble
                If .TotalPebbleIa = 0 And .TpbIA <> 0 Then
                    .TotalPebbleIa = .CpIa
                End If

                'Waste pebble
                If .WstPbIa1 = 0 And .WstPbIa2 <> 0 Then
                    .WstPbIa1 = .WstPbIa2
                End If
            End If
        End With

        gGetCompositeData = True
        ClearParams(params)

        Exit Function

gGetCompositeDataError:
        If aDisplayMissingError = True Then
            MsgBox("Error loading composite hole." & vbCrLf & _
                Err.Description, _
                vbOKOnly + vbExclamation, _
                "Composite Hole Loading Error")
        End If

        gGetCompositeData = False
        ClearParams(params)
    End Function

    Private Sub ZeroCompositeData()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        With gComposite
            .Mine = ""                   '1
            .Section = 0                 '2
            .Township = 0                '3
            .Range = 0                   '4
            .HoleLocation = ""           '5
            .DrillCdate = ""             '6
            .AnalysisCdate = ""          '7
            .AreaOfInfluence = 0         '8
            .HoleElevation = 0           '9
            .PitBottomElevation = 0      '10
            .XSPCoordinate = 0           '11
            .YSPCoordinate = 0           '12
            .TriangleCode = ""           '13
            .TotalNumberSplits = 0       '14
            .SplitsSummarized = ""       '15
            .ProspectorCode = ""         '16
            .OvbThickness = 0            '17
            .OvbX = 0                    '18
            .MtxThickness = 0            '19
            .MtxPebbleX = 0              '20
            .MtxX = 0                    '21
            .TotalThickness = 0          '22
            .TotalPebbleX = 0            '23
            .TotalX = 0                  '24
            .MtxPercentSolids = 0        '25
            .MtxWetDensity = 0           '26
            .CoarsePebbleWtp = 0         '27
            .FinePebbleWtp = 0           '28
            .TotalPebbleWtp = 0          '29
            .ConcentrateWtp = 0          '30
            .TotalProductWtp = 0         '31
            .TotalTailWtp = 0            '32
            .WasteClayWtp = 0            '33
            .GrossConcentrateWtp = 0     '34
            .GrossProductWtp = 0         '35
            .CoarseFeedWtp = 0           '36
            .FineFeedWtp = 0             '37
            .TotalFeedWtp = 0            '38
            .MtxTons = 0                 '39
            .CoarsePebbleTPA = 0         '40
            .FinePebbleTPA = 0           '41
            .TotalPebbleTpa = 0          '42
            .ConcentrateTPA = 0          '43
            .TotalProductTpa = 0         '44
            .TotalTailTpa = 0            '45
            .WasteClayTpa = 0            '46
            .GrossConcentrateTpa = 0     '47
            .GrossProductTpa = 0         '48
            .CoarseFeedTpa = 0           '49
            .FineFeedTpa = 0             '50
            .TotalFeedTpa = 0            '51
            .MtxBPL = 0                  '52
            .CoarsePebbleBPL = 0         '53
            .FinePebbleBPL = 0           '54
            .TotalPebbleBpl = 0          '55
            .ConcentrateBPL = 0          '56
            .TotalProductBpl = 0         '57
            .TotalTailBPL = 0            '58
            .WasteClayBPL = 0            '59
            .GrossConcentrateBpl = 0     '60
            .GrossProductBpl = 0         '61
            .CoarseFeedBpl = 0           '62
            .FineFeedBpl = 0             '63
            .TotalFeedBpl = 0            '64
            .FinePebbleFe2O3 = 0         '65
            .FinePebbleAl2O3 = 0         '66
            .FinePebbleMgO = 0           '67
            .FinePebbleCaO = 0           '68
            .FinePebbleInsol = 0         '69
            .FinePebbleIa = 0            '70
            .CoarsePebbleFe2O3 = 0       '71
            .CoarsePebbleAl2O3 = 0       '72
            .CoarsePebbleMgO = 0         '73
            .CoarsePebbleCaO = 0         '74
            .CoarsePebbleInsol = 0       '75
            .CoarsePebbleIa = 0          '76
            .TotalPebbleFe2O3 = 0        '77
            .TotalPebbleAl2O3 = 0        '78
            .TotalPebbleMgO = 0          '79
            .TotalPebbleCaO = 0          '80
            .TotalPebbleInsol = 0        '81
            .TotalPebbleIa = 0           '82
            .ConcentrateFe2O3 = 0        '83
            .ConcentrateAl2O3 = 0        '84
            .ConcentrateMgO = 0          '85
            .ConcentrateCaO = 0          '86
            .ConcentrateInsol = 0        '87
            .ConcentrateIA = 0           '88
            .TotalProductFe2O3 = 0       '89
            .TotalProductAl2O3 = 0       '90
            .TotalProductMgO = 0         '91
            .TotalProductCaO = 0         '92
            .TotalProductInsol = 0       '93
            .TotalProductIA = 0          '94
            .GrossConcentrateInsol = 0   '95
            .GrossProductInsol = 0       '96

            .MtxDryDensity = 0           '97
            .MtxConcentrateX = 0         '98
            .TotalConcentrateX = 0       '99

            .WstThck = 0                 '100
            .TotX = 0                    '101
            .MinableSplits = 0           '102
            .HoleMinable = 0             '103
            .CpbMinable = 0              '104
            .FpbMinable = 0              '105
            .CpbFeTpaWt = 0              '106
            .CpbAlTpaWt = 0              '107
            .CpbIaTpaWt = 0              '108
            .CpbCaTpaWt = 0              '109
            .FpbFeTpaWt = 0              '110
            .FpbAlTpaWt = 0              '111
            .FpbIaTpaWt = 0              '112
            .FpbCaTpaWt = 0              '113
            .CnFeTpaWt = 0               '114
            .CnAlTpaWt = 0               '115
            .CnIaTpaWt = 0               '116
            .CnCaTpaWt = 0               '117
            .CpIa = 0                    '118
            .FpIa = 0                    '119
            .CnIa = 0                    '120
            .TpIA = 0                    '121
            .TpbIA = 0                   '122
            .FltBplRcvryCalc = 0         '123
            .MtxYdsPerAcre = 0           '124
            .Rc = 0                      '125
            .HasExtraData = 0            '126
            .WstPbWtp = 0                '127
            .WstPbTpa = 0                '128
            .WstPbBpl = 0                '129
            .WstPbFe = 0                 '130
            .WstPbAl = 0                 '131
            .WstPbMg = 0                 '132
            .WstPbCa = 0                 '133
            .WstPbIa1 = 0                '134
            .WstPbIa2 = 0                '135
        End With
    End Sub

    Public Function gGetSplitData(ByVal aMineName As String, _
                                  ByVal aSec As Object, _
                                  ByVal aTwp As Object, _
                                  ByVal aRge As Object, _
                                  ByVal aHole As Object, _
                                  ByVal aSplit As Object, _
                                  ByVal aProspStandard As Object) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetSplitDataError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim DontAdjustAnything As Integer
        Dim CalcSplitConc As Boolean

        'Cannot use gCalcSplitConc here since a mine name may not
        'be gActiveMineNameLong.
        CalcSplitConc = gGetCalcSplitConc(aMineName)

        If CalcSplitConc = True Then
            DontAdjustAnything = 0
        Else
            DontAdjustAnything = 1
        End If

        ZeroSplitData()

        'Cp    Coarse pebble
        'Fp    Fine pebble
        'Tf    Total feed
        'Ff    Fine feed
        'Cf    Coarse feed
        'Lcn   Lab concentrate
        'Cn    Concentrate
        'Tp    Total product
        'Wc    Waste clay
        'Mtx   Matrix
        'Fat   Fine amine tails
        'At    Amine tails
        'Acn   Amine concentrate

        'Get prospect data
        params = gDBParams

        '1
        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        '2
        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        '3
        params.Add("pTownShip", aTwp, ORAPARM_INPUT)
        params("pTownShip").serverType = ORATYPE_NUMBER

        '4
        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        '5
        params.Add("pHoleLocation", aHole, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        '6
        params.Add("pProspStandard", aProspStandard, ORAPARM_INPUT)
        params("pProspStandard").serverType = ORATYPE_VARCHAR2

        '7
        params.Add("pSplit", aSplit, ORAPARM_INPUT)
        params("pSplit").serverType = ORATYPE_NUMBER

        'PROSPECT_SPLIT_BASE

        '8  VARCHAR2    2
        params.Add("pMinableStatus", " ", ORAPARM_OUTPUT)
        params("pMinableStatus").serverType = ORATYPE_VARCHAR2

        '9
        params.Add("pSplitThickness", 0, ORAPARM_OUTPUT)
        params("pSplitThickness").serverType = ORATYPE_NUMBER

        '10   VARCHAR2    3
        params.Add("pDrillCdate", " ", ORAPARM_OUTPUT)
        params("pDrillCdate").serverType = ORATYPE_VARCHAR2

        '10  VARCHAR2   4
        params.Add("pWashCdate", " ", ORAPARM_OUTPUT)
        params("pWashCdate").serverType = ORATYPE_VARCHAR2

        '12
        params.Add("pAreaOfInfluence", 0, ORAPARM_OUTPUT)
        params("pAreaOfInfluence").serverType = ORATYPE_NUMBER

        '13  VARCHAR2   5
        params.Add("pProspectorCode", " ", ORAPARM_OUTPUT)
        params("pProspectorCode").serverType = ORATYPE_VARCHAR2

        '14
        params.Add("pTopOfSplitDepth", 0, ORAPARM_OUTPUT)
        params("pTopOfSplitDepth").serverType = ORATYPE_NUMBER

        '15
        params.Add("pBotOfSplitDepth", 0, ORAPARM_OUTPUT)
        params("pBotOfSplitDepth").serverType = ORATYPE_NUMBER

        '16  VARCHAR2   6
        params.Add("pSampleNumber", " ", ORAPARM_OUTPUT)
        params("pSampleNumber").serverType = ORATYPE_VARCHAR2

        '17
        params.Add("pTotalNumberSplits", 0, ORAPARM_OUTPUT)
        params("pTotalNumberSplits").serverType = ORATYPE_NUMBER

        '18
        params.Add("pRatioOfConc", 0, ORAPARM_OUTPUT)
        params("pRatioOfConc").serverType = ORATYPE_NUMBER

        '19  VARCHAR2   7
        params.Add("pCountyCode", " ", ORAPARM_OUTPUT)
        params("pCountyCode").serverType = ORATYPE_VARCHAR2

        '20  VARCHAR2   8
        params.Add("pMiningCode", " ", ORAPARM_OUTPUT)
        params("pMiningCode").serverType = ORATYPE_VARCHAR2

        '20  VARCHAR2   9
        params.Add("pPumpingCode", " ", ORAPARM_OUTPUT)
        params("pPumpingCode").serverType = ORATYPE_VARCHAR2

        '21  VARCHAR2   10
        params.Add("pMetLabID", " ", ORAPARM_OUTPUT)
        params("pMetLabID").serverType = ORATYPE_VARCHAR2

        '22  VARCHAR2   11
        params.Add("pChemLabID", " ", ORAPARM_OUTPUT)
        params("pChemLabID").serverType = ORATYPE_VARCHAR2

        '23  VARCHAR2   12
        params.Add("pColor", " ", ORAPARM_OUTPUT)
        params("pColor").serverType = ORATYPE_VARCHAR2

        '24
        params.Add("pSplitElevation", 0, ORAPARM_OUTPUT)
        params("pSplitElevation").serverType = ORATYPE_NUMBER

        '25
        params.Add("pHoleNumber", 0, ORAPARM_OUTPUT)
        params("pHoleNumber").serverType = ORATYPE_NUMBER

        '26  VARCHAR2   13
        params.Add("pSampleNumber2", " ", ORAPARM_OUTPUT)
        params("pSampleNumber2").serverType = ORATYPE_VARCHAR2

        '27
        params.Add("pMtxX", 0, ORAPARM_OUTPUT)
        params("pMtxX").serverType = ORATYPE_NUMBER

        '28
        params.Add("pCalcLossPercent", 0, ORAPARM_OUTPUT)
        params("pCalcLossPercent").serverType = ORATYPE_NUMBER

        '29
        params.Add("pCalcLossTPA", 0, ORAPARM_OUTPUT)
        params("pCalcLossTPA").serverType = ORATYPE_NUMBER

        '30
        params.Add("pCalcLossBPL", 0, ORAPARM_OUTPUT)
        params("pCalcLossBPL").serverType = ORATYPE_NUMBER

        '31
        params.Add("pWetMtxLbs", 0, ORAPARM_OUTPUT)
        params("pWetMtxLbs").serverType = ORATYPE_NUMBER

        '32
        params.Add("pMtxGmsWet", 0, ORAPARM_OUTPUT)
        params("pMtxGmsWet").serverType = ORATYPE_NUMBER

        '33
        params.Add("pMtxGmsDry", 0, ORAPARM_OUTPUT)
        params("pMtxGmsDry").serverType = ORATYPE_NUMBER

        '34
        params.Add("pPercentSolidsMtx", 0, ORAPARM_OUTPUT)
        params("pPercentSolidsMtx").serverType = ORATYPE_NUMBER

        '35
        params.Add("pWetFeedLbs", 0, ORAPARM_OUTPUT)
        params("pWetFeedLbs").serverType = ORATYPE_NUMBER

        '36
        params.Add("pFeedMoistWetGms", 0, ORAPARM_OUTPUT)
        params("pFeedMoistWetGms").serverType = ORATYPE_NUMBER

        '37
        params.Add("pFeedMoistDryGms", 0, ORAPARM_OUTPUT)
        params("pFeedMoistDryGms").serverType = ORATYPE_NUMBER

        '38    VARCHAR2    14
        params.Add("pTriangleCode", " ", ORAPARM_OUTPUT)
        params("pTriangleCode").serverType = ORATYPE_VARCHAR2

        'PROSPECT_SPLIT_WT_PERCENT

        '39     1
        params.Add("pCpWtp", 0, ORAPARM_OUTPUT)
        params("pCpWtp").serverType = ORATYPE_NUMBER

        '40     2
        params.Add("pFpWtp", 0, ORAPARM_OUTPUT)
        params("pFpWtp").serverType = ORATYPE_NUMBER

        '41     3
        params.Add("pTfWtp", 0, ORAPARM_OUTPUT)
        params("pTfWtp").serverType = ORATYPE_NUMBER

        '42     4
        params.Add("pWcWtp", 0, ORAPARM_OUTPUT)
        params("pWcWtp").serverType = ORATYPE_NUMBER

        '43     5
        params.Add("pCnWtp", 0, ORAPARM_OUTPUT)
        params("pCnWtp").serverType = ORATYPE_NUMBER

        '44     6
        params.Add("pTpWtp", 0, ORAPARM_OUTPUT)
        params("pTpWtp").serverType = ORATYPE_NUMBER

        '45     7
        params.Add("pFfWtp", 0, ORAPARM_OUTPUT)
        params("pFfWtp").serverType = ORATYPE_NUMBER

        '46     8
        params.Add("pCfWtp", 0, ORAPARM_OUTPUT)
        params("pCfWtp").serverType = ORATYPE_NUMBER

        ' PROSPECT_SPLIT_TPA

        '47     1
        params.Add("pCpTPA", 0, ORAPARM_OUTPUT)
        params("pCpTPA").serverType = ORATYPE_NUMBER

        '48     2
        params.Add("pFpTPA", 0, ORAPARM_OUTPUT)
        params("pFpTPA").serverType = ORATYPE_NUMBER

        '49     3
        params.Add("pTfTPA", 0, ORAPARM_OUTPUT)
        params("pTfTPA").serverType = ORATYPE_NUMBER

        '50     4
        params.Add("pWcTPA", 0, ORAPARM_OUTPUT)
        params("pWcTPA").serverType = ORATYPE_NUMBER

        '51     5
        params.Add("pCnTPA", 0, ORAPARM_OUTPUT)
        params("pCnTPA").serverType = ORATYPE_NUMBER

        '52     6
        params.Add("pTpTPA", 0, ORAPARM_OUTPUT)
        params("pTpTPA").serverType = ORATYPE_NUMBER

        '53     7
        params.Add("pFfTPA", 0, ORAPARM_OUTPUT)
        params("pFfTPA").serverType = ORATYPE_NUMBER

        '54     8
        params.Add("pCfTPA", 0, ORAPARM_OUTPUT)
        params("pCfTPA").serverType = ORATYPE_NUMBER

        '55     9
        params.Add("pFatTPA", 0, ORAPARM_OUTPUT)
        params("pFatTPA").serverType = ORATYPE_NUMBER

        '56     10
        params.Add("pAtTPA", 0, ORAPARM_OUTPUT)
        params("pAtTPA").serverType = ORATYPE_NUMBER

        '57     11
        params.Add("pAcnTPA", 0, ORAPARM_OUTPUT)
        params("pAcnTPA").serverType = ORATYPE_NUMBER

        '58     12
        params.Add("pMtxTPA", 0, ORAPARM_OUTPUT)
        params("pMtxTPA").serverType = ORATYPE_NUMBER

        'PROSPECT_BASE_BPL

        '59     1
        params.Add("pCpBPL", 0, ORAPARM_OUTPUT)
        params("pCpBPL").serverType = ORATYPE_NUMBER

        '60     2
        params.Add("pFpBPL", 0, ORAPARM_OUTPUT)
        params("pFpBPL").serverType = ORATYPE_NUMBER

        '61     3
        params.Add("pLcnBPL", 0, ORAPARM_OUTPUT)
        params("pLcnBPL").serverType = ORATYPE_NUMBER

        '62     4
        params.Add("pTpBPL", 0, ORAPARM_OUTPUT)
        params("pTpBPL").serverType = ORATYPE_NUMBER

        '63     5
        params.Add("pMtxBPL", 0, ORAPARM_OUTPUT)
        params("pMtxBPL").serverType = ORATYPE_NUMBER

        '64     6
        params.Add("pCnBPL", 0, ORAPARM_OUTPUT)
        params("pCnBPL").serverType = ORATYPE_NUMBER

        'PROSPECT_SPLIT_IMP

        '65     1
        params.Add("pCpInsol", 0, ORAPARM_OUTPUT)
        params("pCpInsol").serverType = ORATYPE_NUMBER

        '66     2
        params.Add("pCpFe2O3", 0, ORAPARM_OUTPUT)
        params("pCpFe2O3").serverType = ORATYPE_NUMBER

        '67     3
        params.Add("pCpAl2O3", 0, ORAPARM_OUTPUT)
        params("pCpAl2O3").serverType = ORATYPE_NUMBER

        '68     4
        params.Add("pCpMgO", 0, ORAPARM_OUTPUT)
        params("pCpMgO").serverType = ORATYPE_NUMBER

        '69     5
        params.Add("pCpCaO", 0, ORAPARM_OUTPUT)
        params("pCpCaO").serverType = ORATYPE_NUMBER

        '70     6
        params.Add("pFpInsol", 0, ORAPARM_OUTPUT)
        params("pFpInsol").serverType = ORATYPE_NUMBER

        '71     7
        params.Add("pFpFe2O3", 0, ORAPARM_OUTPUT)
        params("pFpFe2O3").serverType = ORATYPE_NUMBER

        '72     8
        params.Add("pFpAl2O3", 0, ORAPARM_OUTPUT)
        params("pFpAl2O3").serverType = ORATYPE_NUMBER

        '73     9
        params.Add("pFpMgO", 0, ORAPARM_OUTPUT)
        params("pFpMgO").serverType = ORATYPE_NUMBER

        '74     10
        params.Add("pFpCaO", 0, ORAPARM_OUTPUT)
        params("pFpCaO").serverType = ORATYPE_NUMBER

        '75     11
        params.Add("pFpFeAl", 0, ORAPARM_OUTPUT)
        params("pFpFeAl").serverType = ORATYPE_NUMBER

        '76     12
        params.Add("pFpCaOP2O5", 0, ORAPARM_OUTPUT)
        params("pFpCaOP2O5").serverType = ORATYPE_NUMBER

        '77     13
        params.Add("pLcnInsol", 0, ORAPARM_OUTPUT)
        params("pLcnInsol").serverType = ORATYPE_NUMBER

        '78     14
        params.Add("pLcnFe2O3", 0, ORAPARM_OUTPUT)
        params("pLcnFe2O3").serverType = ORATYPE_NUMBER

        '79     15
        params.Add("pLcnAl2O3", 0, ORAPARM_OUTPUT)
        params("pLcnAl2O3").serverType = ORATYPE_NUMBER

        '80     16
        params.Add("pLcnMgO", 0, ORAPARM_OUTPUT)
        params("pLcnMgO").serverType = ORATYPE_NUMBER

        '81     17
        params.Add("pLcnCaO", 0, ORAPARM_OUTPUT)
        params("pLcnCaO").serverType = ORATYPE_NUMBER

        '82     18
        params.Add("pLcnFeAl", 0, ORAPARM_OUTPUT)
        params("pLcnFeAl").serverType = ORATYPE_NUMBER

        '83     19
        params.Add("pTpInsol", 0, ORAPARM_OUTPUT)
        params("pTpInsol").serverType = ORATYPE_NUMBER

        '84     20
        params.Add("pTpFe2O3", 0, ORAPARM_OUTPUT)
        params("pTpFe2O3").serverType = ORATYPE_NUMBER

        '85     21
        params.Add("pTpAl2O3", 0, ORAPARM_OUTPUT)
        params("pTpAl2O3").serverType = ORATYPE_NUMBER

        '86     22
        params.Add("pTpMgO", 0, ORAPARM_OUTPUT)
        params("pTpMgO").serverType = ORATYPE_NUMBER

        '87     23
        params.Add("pTpCaO", 0, ORAPARM_OUTPUT)
        params("pTpCaO").serverType = ORATYPE_NUMBER

        '88     24
        params.Add("pTpFeAl", 0, ORAPARM_OUTPUT)
        params("pTpFeAl").serverType = ORATYPE_NUMBER

        '89     25
        params.Add("pTpCaOP2O5", 0, ORAPARM_OUTPUT)
        params("pTpCaOP2O5").serverType = ORATYPE_NUMBER

        '90     26
        params.Add("pMtxInsol", 0, ORAPARM_OUTPUT)
        params("pMtxInsol").serverType = ORATYPE_NUMBER

        '91     27
        params.Add("pCnInsol", 0, ORAPARM_OUTPUT)
        params("pCnInsol").serverType = ORATYPE_NUMBER

        '92     28
        params.Add("pCnCaO", 0, ORAPARM_OUTPUT)
        params("pCnCaO").serverType = ORATYPE_NUMBER

        '93     29
        params.Add("pCnFeAl", 0, ORAPARM_OUTPUT)
        params("pCnFeAl").serverType = ORATYPE_NUMBER

        'PROSPECT_SPLIT_GRAMS

        '94     1
        params.Add("pCpGrams", 0, ORAPARM_OUTPUT)
        params("pCpGrams").serverType = ORATYPE_NUMBER

        '95     2
        params.Add("pFpGrams", 0, ORAPARM_OUTPUT)
        params("pFpGrams").serverType = ORATYPE_NUMBER

        '96     3
        params.Add("pTfGrams", 0, ORAPARM_OUTPUT)
        params("pTfGrams").serverType = ORATYPE_NUMBER

        '97     4
        params.Add("pFatGrams", 0, ORAPARM_OUTPUT)
        params("pFatGrams").serverType = ORATYPE_NUMBER

        '98     5
        params.Add("pAtGrams", 0, ORAPARM_OUTPUT)
        params("pAtGrams").serverType = ORATYPE_NUMBER

        '99     6
        params.Add("pLcnGrams", 0, ORAPARM_OUTPUT)
        params("pLcnGrams").serverType = ORATYPE_NUMBER

        'PROSPECT_SPLIT_DENSITY

        '100    1
        params.Add("pWetDensityVolume", 0, ORAPARM_OUTPUT)
        params("pWetDensityVolume").serverType = ORATYPE_NUMBER

        '101    2
        params.Add("pWetDensityWeight", 0, ORAPARM_OUTPUT)
        params("pWetDensityWeight").serverType = ORATYPE_NUMBER

        '102    3
        params.Add("pWetDensity", 0, ORAPARM_OUTPUT)
        params("pWetDensity").serverType = ORATYPE_NUMBER

        '103    4
        params.Add("pDryDensityVolume", 0, ORAPARM_OUTPUT)
        params("pDryDensityVolume").serverType = ORATYPE_NUMBER

        '104    5
        params.Add("pDryDensityWeight", 0, ORAPARM_OUTPUT)
        params("pDryDensityWeight").serverType = ORATYPE_NUMBER

        '105    6
        params.Add("pDryDensity", 0, ORAPARM_OUTPUT)
        params("pDryDensity").serverType = ORATYPE_NUMBER

        'ALL CALCULATIONS

        '106    1
        params.Add("pCalcHeadFeedBPL", 0, ORAPARM_OUTPUT)
        params("pCalcHeadFeedBPL").serverType = ORATYPE_NUMBER

        '107    2
        params.Add("pGrossConcentrateTPA", 0, ORAPARM_OUTPUT)
        params("pGrossConcentrateTPA").serverType = ORATYPE_NUMBER

        '108    3
        params.Add("pGrossConcentrateBPL", 0, ORAPARM_OUTPUT)
        params("pGrossConcentrateBPL").serverType = ORATYPE_NUMBER

        '109    4
        params.Add("pGrossConcentrateInsol", 0, ORAPARM_OUTPUT)
        params("pGrossConcentrateInsol").serverType = ORATYPE_NUMBER

        '110    5
        params.Add("pGrossPebbleWtp", 0, ORAPARM_OUTPUT)
        params("pGrossPebbleWtp").serverType = ORATYPE_NUMBER

        '111    6
        params.Add("pGrossPebbleTPA", 0, ORAPARM_OUTPUT)
        params("pGrossPebbleTPA").serverType = ORATYPE_NUMBER

        '112    7
        params.Add("pGrossPebbleBPL", 0, ORAPARM_OUTPUT)
        params("pGrossPebbleBPL").serverType = ORATYPE_NUMBER

        '113    8
        params.Add("pGrossPebbleInsol", 0, ORAPARM_OUTPUT)
        params("pGrossPebbleInsol").serverType = ORATYPE_NUMBER

        '114    9
        params.Add("pGrossPebbleFe2O3", 0, ORAPARM_OUTPUT)
        params("pGrossPebbleFe2O3").serverType = ORATYPE_NUMBER

        '115    10
        params.Add("pGrossPebbleAl2O3", 0, ORAPARM_OUTPUT)
        params("pGrossPebbleAl2O3").serverType = ORATYPE_NUMBER

        '116    11
        params.Add("pGrossPebbleMgO", 0, ORAPARM_OUTPUT)
        params("pGrossPebbleMgO").serverType = ORATYPE_NUMBER

        '117    12
        params.Add("pConcentrateWtp", 0, ORAPARM_OUTPUT)
        params("pConcentrateWtp").serverType = ORATYPE_NUMBER

        '118    13
        params.Add("pConcentrateTPA", 0, ORAPARM_OUTPUT)
        params("pConcentrateTPA").serverType = ORATYPE_NUMBER

        '119    14
        params.Add("pConcentrateBPL", 0, ORAPARM_OUTPUT)
        params("pConcentrateBPL").serverType = ORATYPE_NUMBER

        '120    15
        params.Add("pConcentrateInsol", 0, ORAPARM_OUTPUT)
        params("pConcentrateInsol").serverType = ORATYPE_NUMBER

        '121    16
        params.Add("pConcentrateFe2O3", 0, ORAPARM_OUTPUT)
        params("pConcentrateFe2O3").serverType = ORATYPE_NUMBER

        '122    17
        params.Add("pConcentrateAl2O3", 0, ORAPARM_OUTPUT)
        params("pConcentrateAl2O3").serverType = ORATYPE_NUMBER

        '123    18
        params.Add("pConcentrateMgO", 0, ORAPARM_OUTPUT)
        params("pConcentrateMgO").serverType = ORATYPE_NUMBER

        '124    19
        params.Add("pTotalProductWtp", 0, ORAPARM_OUTPUT)
        params("pTotalProductWtp").serverType = ORATYPE_NUMBER

        '125    20
        params.Add("pTotalProductTPA", 0, ORAPARM_OUTPUT)
        params("pTotalProductTPA").serverType = ORATYPE_NUMBER

        '126    21
        params.Add("pTotalProductBPL", 0, ORAPARM_OUTPUT)
        params("pTotalProductBPL").serverType = ORATYPE_NUMBER

        '127    22
        params.Add("pTotalProductInsol", 0, ORAPARM_OUTPUT)
        params("pTotalProductInsol").serverType = ORATYPE_NUMBER

        '128    23
        params.Add("pTotalProductFe2O3", 0, ORAPARM_OUTPUT)
        params("pTotalProductFe2O3").serverType = ORATYPE_NUMBER

        '129    24
        params.Add("pTotalProductAl2O3", 0, ORAPARM_OUTPUT)
        params("pTotalProductAl2O3").serverType = ORATYPE_NUMBER

        '130    25
        params.Add("pTotalProductMgO", 0, ORAPARM_OUTPUT)
        params("pTotalProductMgO").serverType = ORATYPE_NUMBER

        '131    26
        params.Add("pTotalTailWtp", 0, ORAPARM_OUTPUT)
        params("pTotalTailWtp").serverType = ORATYPE_NUMBER

        '132    27
        params.Add("pTotalTailTPA", 0, ORAPARM_OUTPUT)
        params("pTotalTailTPA").serverType = ORATYPE_NUMBER

        '133    28
        params.Add("pTotalTailBPL", 0, ORAPARM_OUTPUT)
        params("pTotalTailBPL").serverType = ORATYPE_NUMBER

        '134    29      VARCHAR2  15
        params.Add("pColorTrans", " ", ORAPARM_OUTPUT)
        params("pColorTrans").serverType = ORATYPE_VARCHAR2

        '135    30      VARCHAR2  16
        params.Add("pMinableStatusTrans", " ", ORAPARM_OUTPUT)
        params("pMinableStatusTrans").serverType = ORATYPE_VARCHAR2

        '136    31
        params.Add("pCalcOvbX", 0, ORAPARM_OUTPUT)
        params("pCalcOvbX").serverType = ORATYPE_NUMBER

        '137    32
        params.Add("pCalcMtxX", 0, ORAPARM_OUTPUT)
        params("pCalcMtxX").serverType = ORATYPE_NUMBER

        '138    33
        params.Add("pCalcTotalX", 0, ORAPARM_OUTPUT)
        params("pCalcTotalX").serverType = ORATYPE_NUMBER

        '139    34
        params.Add("pFlotationRC", 0, ORAPARM_OUTPUT)
        params("pFlotationRC").serverType = ORATYPE_NUMBER

        '140    35
        params.Add("pFlotationRecovery", 0, ORAPARM_OUTPUT)
        params("pFlotationRecovery").serverType = ORATYPE_NUMBER

        '141    36
        params.Add("pTopOfSeamElevation", 0, ORAPARM_OUTPUT)
        params("pTopOfSeamElevation").serverType = ORATYPE_NUMBER

        '142    37
        params.Add("pFfBPL", 0, ORAPARM_OUTPUT)
        params("pFfBPL").serverType = ORATYPE_NUMBER

        '143    38
        params.Add("pCfBPL", 0, ORAPARM_OUTPUT)
        params("pCfBPL").serverType = ORATYPE_NUMBER

        '144    39
        params.Add("pTfBPL", 0, ORAPARM_OUTPUT)
        params("pTfBPL").serverType = ORATYPE_NUMBER

        '145    40
        params.Add("pFatBPL", 0, ORAPARM_OUTPUT)
        params("pFatBPL").serverType = ORATYPE_NUMBER

        '146    41
        params.Add("pAtBPL", 0, ORAPARM_OUTPUT)
        params("pAtBPL").serverType = ORATYPE_NUMBER

        '147    42
        params.Add("pWcBPL", 0, ORAPARM_OUTPUT)
        params("pWcBPL").serverType = ORATYPE_NUMBER

        '148    43
        params.Add("pHardpanCode", 0, ORAPARM_OUTPUT)
        params("pHardpanCode").serverType = ORATYPE_NUMBER

        '149    44
        params.Add("pGrossPebbleCaO", 0, ORAPARM_OUTPUT)
        params("pGrossPebbleCaO").serverType = ORATYPE_NUMBER

        '150    45
        params.Add("pConcentrateCaO", 0, ORAPARM_OUTPUT)
        params("pConcentrateCaO").serverType = ORATYPE_NUMBER

        '151    46
        params.Add("pTotalProductCaO", 0, ORAPARM_OUTPUT)
        params("pTotalProductCaO").serverType = ORATYPE_NUMBER

        '152    47
        params.Add("pCpCd", 0, ORAPARM_OUTPUT)
        params("pCpCd").serverType = ORATYPE_NUMBER

        '153    48
        params.Add("pFpCd", 0, ORAPARM_OUTPUT)
        params("pFpCd").serverType = ORATYPE_NUMBER

        '154    49
        params.Add("pLcnCd", 0, ORAPARM_OUTPUT)
        params("pLcnCd").serverType = ORATYPE_NUMBER

        '155    50
        params.Add("pConcentrateCd", 0, ORAPARM_OUTPUT)
        params("pConcentrateCd").serverType = ORATYPE_NUMBER

        '156    51
        params.Add("pTotalProductCd", 0, ORAPARM_OUTPUT)
        params("pTotalProductCd").serverType = ORATYPE_NUMBER

        '158    52
        params.Add("pTpCd", 0, ORAPARM_OUTPUT)
        params("pTpCd").serverType = ORATYPE_NUMBER

        '159    53
        params.Add("pGrossPebbleCd", 0, ORAPARM_OUTPUT)
        params("pGrossPebbleCd").serverType = ORATYPE_NUMBER

        '160
        params.Add("pTfBplCalc", 0, ORAPARM_OUTPUT)
        params("pTfBplCalc").serverType = ORATYPE_NUMBER

        '161
        params.Add("pDontAdjustAnything", DontAdjustAnything, ORAPARM_INPUT)
        params("pDontAdjustAnything").serverType = ORATYPE_NUMBER

        'Procedure get_split
        'pMineName              IN      VARCHAR2,    -- 1
        'pSection               IN      NUMBER,      -- 2
        'pTownShip              IN      NUMBER,      -- 3
        'pRange                 IN      NUMBER,      -- 4
        'pHoleLocation          IN      VARCHAR2,    -- 5
        'pProspStandard         IN      VARCHAR2,    -- 6
        'pSplit                 IN      NUMBER,      -- 7
        'pMinableStatus         OUT     VARCHAR2,    -- 8
        'pSplitThickness        OUT     NUMBER,      -- 9
        'pDrillCdate            OUT     VARCHAR2,    -- 10
        'pWashCdate             OUT     VARCHAR2,    -- 11
        'pAreaOfInfluence       OUT     NUMBER,      -- 12
        'pProspectorCode        OUT     VARCHAR2,    -- 13
        'pTopOfSplitDepth       OUT     NUMBER,      -- 14
        'pBotOfSplitDepth       OUT     NUMBER,      -- 15
        'pSampleNumber          OUT     VARCHAR2,    -- 16
        'pTotalNumberSplits     OUT     NUMBER,      -- 17
        'pRatioOfConc           OUT     NUMBER,      -- 18
        'pCountyCode            OUT     VARCHAR2,    -- 19
        'pMiningCode            OUT     VARCHAR2,    -- 20
        'pPumpingCode           OUT     VARCHAR2,    -- 21
        'pMetLabID              OUT     VARCHAR2,    -- 22
        'pChemLabID             OUT     VARCHAR2,    -- 23
        'pColor                 OUT     VARCHAR2,    -- 24
        'pSplitElevation        OUT     NUMBER,      -- 25
        'pHoleNumber            OUT     NUMBER,      -- 26
        'pSampleNumber2         OUT     VARCHAR2,    -- 27
        'pMtxX                  OUT     NUMBER,      -- 28
        'pCalcLossPercent       OUT     NUMBER,      -- 29
        'pCalcLossTPA           OUT     NUMBER,      -- 30
        'pCalcLossBPL           OUT     NUMBER,      -- 31
        'pWetMtxLbs             OUT     NUMBER,      -- 32
        'pMtxGmsWet             OUT     NUMBER,      -- 33
        'pMtxGmsDry             OUT     NUMBER,      -- 34
        'pPercentSolidsMtx      OUT     NUMBER,      -- 35
        'pWetFeedLbs            OUT     NUMBER,      -- 36
        'pFeedMoistWetGms       OUT     NUMBER,      -- 37
        'pFeedMoistDryGms       OUT     NUMBER,      -- 38
        'pTriangleCode          OUT     VARCHAR2,    -- 39
        'pCpWtp                 OUT     NUMBER,      -- 40
        'pFpWtp                 OUT     NUMBER,      -- 41
        'pTfWtp                 OUT     NUMBER,      -- 42
        'pWcWtp                 OUT     NUMBER,      -- 43
        'pCnWtp                 OUT     NUMBER,      -- 44
        'pTpWtp                 OUT     NUMBER,      -- 45
        'pFfWtp                 OUT     NUMBER,      -- 46
        'pCfWtp                 OUT     NUMBER,      -- 47
        'pCpTPA                 OUT     NUMBER,      -- 48
        'pFpTPA                 OUT     NUMBER,      -- 49
        'pTfTPA                 OUT     NUMBER,      -- 50
        'pWcTPA                 OUT     NUMBER,      -- 51
        'pCnTPA                 OUT     NUMBER,      -- 52
        'pTpTPA                 OUT     NUMBER,      -- 53
        'pFfTPA                 OUT     NUMBER,      -- 54
        'pCfTPA                 OUT     NUMBER,      -- 55
        'pFatTPA                OUT     NUMBER,      -- 56
        'pAtTPA                 OUT     NUMBER,      -- 57
        'pAcnTPA                OUT     NUMBER,      -- 58
        'pMtxTPA                OUT     NUMBER,      -- 59
        'pCpBPL                 OUT     NUMBER,      -- 60
        'pFpBPL                 OUT     NUMBER,      -- 61
        'pLcnBPL                OUT     NUMBER,      -- 62
        'pTpBPL                 OUT     NUMBER,      -- 63
        'pMtxBPL                OUT     NUMBER,      -- 64
        'pCnBPL                 OUT     NUMBER,      -- 65
        '--
        'pCpInsol               OUT     NUMBER,      -- 66
        'pCpFe2O3               OUT     NUMBER,      -- 67
        'pCpAl2O3               OUT     NUMBER,      -- 68
        'pCpMgO                 OUT     NUMBER,      -- 69
        'pCpCaO                 OUT     NUMBER,      -- 70
        '--
        'pFpInsol               OUT     NUMBER,      -- 71
        'pFpFe2O3               OUT     NUMBER,      -- 72
        'pFpAl2O3               OUT     NUMBER,      -- 73
        'pFpMgO                 OUT     NUMBER,      -- 74
        'pFpCaO                 OUT     NUMBER,      -- 75
        'pFpFeAl                OUT     NUMBER,      -- 76
        'pFpCaOP2O5             OUT     NUMBER,      -- 77
        '--
        'pLcnInsol              OUT     NUMBER,      -- 78
        'pLcnFe2O3              OUT     NUMBER,      -- 79
        'pLcnAl2O3              OUT     NUMBER,      -- 80
        'pLcnMgO                OUT     NUMBER,      -- 81
        'pLcnCaO                OUT     NUMBER,      -- 82
        'pLcnFeAl               OUT     NUMBER,      -- 83
        '--
        'pTpInsol               OUT     NUMBER,      -- 84
        'pTpFe2O3               OUT     NUMBER,      -- 85
        'pTpAl2O3               OUT     NUMBER,      -- 86
        'pTpMgO                 OUT     NUMBER,      -- 87
        'pTpCaO                 OUT     NUMBER,      -- 88
        'pTpFeAl                OUT     NUMBER,      -- 89
        'pTpCaOP2O5             OUT     NUMBER,      -- 90
        '--
        'pMtxInsol              OUT     NUMBER,      -- 91
        'pCnInsol               OUT     NUMBER,      -- 92
        'pCnCaO                 OUT     NUMBER,      -- 93
        'pCnFeAl                OUT     NUMBER,      -- 94
        'pCpGrams               OUT     NUMBER,      -- 95
        'pFpGrams               OUT     NUMBER,      -- 96
        'pTfGrams               OUT     NUMBER,      -- 97
        'pFatGrams              OUT     NUMBER,      -- 98
        'pAtGrams               OUT     NUMBER,      -- 99
        'pLcnGrams              OUT     NUMBER,      -- 100
        'pWetDensityVolume      OUT     NUMBER,      -- 101
        'pWetDensityWeight      OUT     NUMBER,      -- 102
        'pWetDensity            OUT     NUMBER,      -- 103
        'pDryDensityVolume      OUT     NUMBER,      -- 104
        'pDryDensityWeight      OUT     NUMBER,      -- 105
        'pDryDensity            OUT     NUMBER,      -- 106
        'pCalcHeadFeedBPL       OUT     NUMBER,      -- 107
        'pGrossConcentrateTPA   OUT     NUMBER,      -- 108
        'pGrossConcentrateBPL   OUT     NUMBER,      -- 109
        'pGrossConcentrateInsol OUT     NUMBER,      -- 110
        '--
        'pGrossPebbleWtp        OUT     NUMBER,      -- 111
        'pGrossPebbleTPA        OUT     NUMBER,      -- 112
        'pGrossPebbleBPL        OUT     NUMBER,      -- 113
        'pGrossPebbleInsol      OUT     NUMBER,      -- 114
        'pGrossPebbleFe2O3      OUT     NUMBER,      -- 115
        'pGrossPebbleAl2O3      OUT     NUMBER,      -- 116
        'pGrossPebbleMgO        OUT     NUMBER,      -- 117
        '--
        'pConcentrateWtp        OUT     NUMBER,      -- 118
        'pConcentrateTPA        OUT     NUMBER,      -- 119
        'pConcentrateBPL        OUT     NUMBER,      -- 120
        'pConcentrateInsol      OUT     NUMBER,      -- 121
        'pConcentrateFe2O3      OUT     NUMBER,      -- 122
        'pConcentrateAl2O3      OUT     NUMBER,      -- 123
        'pConcentrateMgO        OUT     NUMBER,      -- 124
        '--
        'pTotalProductWtp       OUT     NUMBER,      -- 125
        'pTotalProductTPA       OUT     NUMBER,      -- 126
        'pTotalProductBPL       OUT     NUMBER,      -- 127
        'pTotalProductInsol     OUT     NUMBER,      -- 128
        'pTotalProductFe2O3     OUT     NUMBER,      -- 129
        'pTotalProductAl2O3     OUT     NUMBER,      -- 130
        'pTotalProductMgO       OUT     NUMBER,      -- 131
        '--
        'pTotalTailWtp          OUT     NUMBER,      -- 132
        'pTotalTailTPA          OUT     NUMBER,      -- 133
        'pTotalTailBPL          OUT     NUMBER,      -- 134
        'pColorTrans            OUT   VARCHAR2,      -- 135
        'pMinableStatusTrans    OUT   VARCHAR2,      -- 136
        'pCalcOvbX              OUT     NUMBER,      -- 137
        'pCalcMtxX              OUT     NUMBER,      -- 138
        'pCalcTotalX            OUT     NUMBER,      -- 139
        'pFlotationRC           OUT     NUMBER,      -- 140
        'pFlotationRecovery     OUT     NUMBER,      -- 141
        'pTopOfSeamElevation    OUT     NUMBER,      -- 142
        'pFfBPL                 OUT     NUMBER,      -- 143
        'pCfBPL                 OUT     NUMBER,      -- 144
        'pTfBPL                 OUT     NUMBER,      -- 145
        'pFatBPL                OUT     NUMBER,      -- 146
        'pAtBPL                 OUT     NUMBER,      -- 147
        'pWcBPL                 OUT     NUMBER,      -- 148
        '--
        'pHardpanCode           OUT     NUMBER,      -- 149
        'pGrossPebbleCaO        OUT     NUMBER,      -- 150
        'pConcentrateCaO        OUT     NUMBER,      -- 151
        'pTotalProductCaO       OUT     NUMBER,      -- 152
        'pCpCd                  OUT     NUMBER,      -- 153
        'pFpCd                  OUT     NUMBER,      -- 154
        'pLcnCd                 OUT     NUMBER,      -- 155
        'pConcentrateCd         OUT     NUMBER,      -- 156
        'pTotalProductCd        OUT     NUMBER,      -- 157
        'pTpCd                  OUT     NUMBER,      -- 158
        'pGrossPebbleCd         OUT     NUMBER,      -- 159
        'pTfdBplCalc            OUT     NUMBER,      -- 160
        'pDontAdjustAnything    IN      NUMBER)      -- 161

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect.get_split(:pMineName, :pSection," + _
                      ":pTownShip, :pRange, :pHoleLocation, :pProspStandard, :pSplit, :pMinableStatus, :pSplitThickness, :pDrillCdate, :pWashCdate, :pAreaOfInfluence, :pProspectorCode," + _
                      ":pTopOfSplitDepth, :pBotOfSplitDepth, :pSampleNumber, :pTotalNumberSplits, :pRatioOfConc, :pCountyCode," + _
                      ":pMiningCode, :pPumpingCode, :pMetLabID, :pChemLabID, :pColor, :pSplitElevation, :pHoleNumber, :pSampleNumber2," + _
                      ":pMtxX, :pCalcLossPercent, :pCalcLossTPA, :pCalcLossBPL, :pWetMtxLbs, :pMtxGmsWet," + _
                      ":pMtxGmsDry, :pPercentSolidsMtx, :pWetFeedLbs, :pFeedMoistWetGms, :pFeedMoistDryGms, :pTriangleCode," + _
                      ":pCpWtp, :pFpWtp, :pTfWtp, :pWcWtp, :pCnWtp, :pTpWtp, :pFfWtp, :pCfWtp," + _
                      ":pCpTPA, :pFpTPA, :pTfTPA, :pWcTPA, :pCnTPA, :pTpTPA, :pFfTPA, :pCfTPA, :pFatTPA, :pAtTPA, :pAcnTPA, :pMtxTPA," + _
                      ":pCpBPL, :pFpBPL, :pLcnBPL, :pTpBPL, :pMtxBPL, :pCnBPL, :pCpInsol, :pCpFe2O3, :pCpAl2O3, :pCpMgO, :pCpCaO," + _
                      ":pFpInsol, :pFpFe2O3, :pFpAl2O3, :pFpMgO, :pFpCaO, :pFpFeAl, :pFpCaOP2O5," + _
                      ":pLcnInsol, :pLcnFe2O3, :pLcnAl2O3, :pLcnMgO, :pLcnCaO, :pLcnFeAl," + _
                      ":pTpInsol, :pTpFe2O3, :pTpAl2O3, :pTpMgO, :pTpCaO, :pTpFeAl, :pTpCaOP2O5," + _
                      ":pMtxInsol, :pCnInsol, :pCnCaO, :pCnFeAl, :pCpGrams, :pFpGrams, :pTfGrams, :pFatGrams, :pAtGrams," + _
                      ":pLcnGrams, :pWetDensityVolume, :pWetDensityWeight, :pWetDensity," + _
                      ":pDryDensityVolume, :pDryDensityWeight, :pDryDensity, :pCalcHeadFeedBPL," + _
                      ":pGrossConcentrateTPA, :pGrossConcentrateBPL, :pGrossConcentrateInsol, :pGrossPebbleWtp, :pGrossPebbleTPA," + _
                      ":pGrossPebbleBPL, :pGrossPebbleInsol, :pGrossPebbleFe2O3, :pGrossPebbleAl2O3," + _
                      ":pGrossPebbleMgO, :pConcentrateWtp, :pConcentrateTPA, :pConcentrateBPL, :pConcentrateInsol," + _
                      ":pConcentrateFe2O3, :pConcentrateAl2O3, :pConcentrateMgO, :pTotalProductWtp, :pTotalProductTPA," + _
                      ":pTotalProductBPL, :pTotalProductInsol, :pTotalProductFe2O3, :pTotalProductAl2O3," + _
                      ":pTotalProductMgO, :pTotalTailWtp, :pTotalTailTPA, :pTotalTailBPL, :pColorTrans," + _
                      ":pMinableStatusTrans, :pCalcOvbX, :pCalcMtxX, :pCalcTotalX, :pFlotationRC, :pFlotationRecovery," + _
                      ":pTopOfSeamElevation, :pFfBPL, :pCfBPL, :pTfBPL, :pFatBPL, :pAtBPL, :pWcBPL," + _
                      ":pHardpanCode,:pGrossPebbleCaO, :pConcentrateCaO,:pTotalProductCaO,:pCpCd, " + _
                      ":pFpCd, :pLcnCd, :pConcentrateCd, :pTotalProductCd, :pTpCd, :pGrossPebbleCd, :pTfBplCalc, :pDontAdjustAnything);end;", ORASQL_FAILEXEC)

        With gSplit
            .Mine = params("pMineName").Value                            '1
            .Section = params("pSection").Value                          '2
            .Township = params("pTownship").Value                        '3
            .Range = params("pRange").Value                              '4
            .HoleLocation = params("pHoleLocation").Value                '5
            .Split = params("pSplit").Value                              '6
            .MinableStatus = params("pMinableStatus").Value              '7
            .SplitThickness = params("pSplitThickness").Value            '8
            .DrillCdate = params("pDrillCdate").Value                    '9

            If Not IsDBNull(params("pWashCdate").Value) Then
                .WashCdate = params("pWashCdate").Value                  '10
            Else
                .WashCdate = ""
            End If

            .AreaOfInfluence = params("pAreaOfInfluence").Value          '11
            .ProspectorCode = params("pProspectorCode").Value            '12
            .TopOfSplitDepth = params("pTopOfSplitDepth").Value          '13
            .BotOfSplitDepth = params("pBotOfSplitDepth").Value          '14
            .SampleNumber = params("pSampleNumber").Value                '15
            .TotalNumberSplits = params("pTotalNumberSplits").Value      '16          '12
            .RatioOfConc = params("pRatioOfConc").Value                  '17
            .CountyCode = params("pCountyCode").Value                    '18
            .MiningCode = params("pMiningCode").Value                    '19
            .PumpingCode = params("pPumpingCode").Value                  '20
            .MetLabID = params("pMetLabID").Value                        '21
            .ChemLabID = params("pChemLabID").Value                      '22
            .Color = params("pColor").Value                              '23
            .SplitElevation = params("pSplitElevation").Value            '24
            .HoleNumber = params("pHoleNumber").Value                    '25
            .SampleNumber2 = params("pSampleNumber2").Value              '26
            .MtxX = params("pMtxX").Value                                '27
            .CalcLossPercent = params("pCalcLossPercent").Value          '28
            .CalcLossTPA = params("pCalcLossTPA").Value                  '30
            .CalcLossBPL = params("pCalclossBPL").Value                  '31
            .WetMtxLbs = params("pWetMtxLbs").Value                      '32
            .MtxGmsWet = params("pMtxGmsWet").Value                      '33
            .MtxGmsDry = params("pMtxGmsDry").Value                      '34
            .PercentSolidsMtx = params("pPercentSolidsMtx").Value        '35
            .WetFeedLbs = params("pWetFeedLbs").Value                    '36
            .FeedMoistWetGms = params("pFeedMoistWetGms").Value          '37
            .FeedMoistDryGms = params("pFeedMoistDryGms").Value          '38

            If Not IsDBNull(params("pTriangleCode").Value) Then
                .TriangleCode = params("pTriangleCode").Value            '39
            Else
                .TriangleCode = ""
            End If

            .CpWtp = params("pCpWtp").Value                              '40
            .FpWtp = params("pFpWtp").Value                              '41
            .TfWtp = params("pTfWtp").Value                              '42
            .WcWtp = params("pWcWtp").Value                              '43
            .CnWtp = params("pCnWtp").Value                              '44
            .TpWtp = params("pTpWtp").Value                              '45
            .FfWtp = params("pFfWtp").Value                              '46
            .CfWtp = params("pCfWtp").Value                              '47
            .CpTpa = params("pCpTPA").Value                              '48
            .FpTpa = params("pFpTPA").Value                              '49
            .TfTPA = params("pTfTPA").Value                              '50
            .WcTpa = params("pWcTPA").Value                              '51
            .CnTpa = params("pCnTPA").Value                              '52
            .TpTpa = params("pTpTPA").Value                              '53
            .FfTpa = params("pFfTPA").Value                              '54
            .CfTpa = params("pCfTPA").Value                              '55
            .FatTpa = params("pFatTPA").Value                            '56
            .AtTpa = params("pAtTPA").Value                              '57
            .AcnTpa = params("pAcnTPA").Value                            '58
            .MtxTPA = params("pMtxTPA").Value                            '59
            .CpBPL = params("pCpBPL").Value                              '60
            .FpBpl = params("pFpBPL").Value                              '61
            .LcnBpl = params("pLcnBPL").Value                            '62
            .TpBpl = params("pTpBPL").Value                              '63
            .MtxBPL = params("pMtxBPL").Value                            '64
            .CnBpl = params("pCnBPL").Value                              '65
            .CpInsol = params("pCpInsol").Value                          '66
            .CpFe2O3 = params("pCpFe2O3").Value                          '67
            .CpAl2O3 = params("pCpAl2O3").Value                          '68
            .CpMgO = params("pCpMgO").Value                              '69
            .CpCaO = params("pCpCaO").Value                              '70
            .FpInsol = params("pFpInsol").Value                          '71
            .FpFe2O3 = params("pFpFe2O3").Value                          '72
            .FpAl2O3 = params("pFpAl2O3").Value                          '73
            .FpMgO = params("pFpMgO").Value                              '74
            .FpCaO = params("pFpCaO").Value                              '75
            .FpFeAl = params("pFpFeAl").Value                            '76
            .FpCaOP2O5 = params("pFpCaOP2O5").Value                      '77
            .LcnInsol = params("pLcnInsol").Value                        '78
            .LcnFe2O3 = params("pLcnFe2O3").Value                        '79
            .LcnAl2O3 = params("pLcnAl2O3").Value                        '80
            .LcnMgO = params("pLcnMgO").Value                            '81
            .LcnCaO = params("pLcnCaO").Value                            '82
            .LcnFeAl = params("pLcnFeAl").Value                          '83
            .TpInsol = params("pTpInsol").Value                          '84
            .TpFe2O3 = params("pTpFe2O3").Value                          '85
            .TpAl2O3 = params("pTpAl2O3").Value                          '86
            .TpMgO = params("pTpMgO").Value                              '87
            .TpCaO = params("pTpCaO").Value                              '88
            .TpFeAl = params("pTpFeAl").Value                            '89
            .TpCaOP2O5 = params("pTpCaOP2O5").Value                      '90
            .MtxInsol = params("pMtxInsol").Value                        '91
            .CnInsol = params("pCnInsol").Value                          '92
            .CnCaO = params("pCnCaO").Value                              '93
            .CnFeAl = params("pCnFeAl").Value                            '94
            .CpGrams = params("pCpGrams").Value                          '95
            .FpGrams = params("pFpGrams").Value                          '96
            .TfGrams = params("pTfGrams").Value                          '97
            .FatGrams = params("pFatGrams").Value                        '98
            .AtGrams = params("pAtGrams").Value                          '99
            .LcnGrams = params("pLcnGrams").Value                        '100
            .WetDensityVolume = params("pWetDensityVolume").Value        '101
            .WetDensityWeight = params("pWetDensityWeight").Value        '102
            .WetDensity = params("pWetDensity").Value                    '103
            .DryDensityVolume = params("pDryDensityVolume").Value        '104
            .DryDensityWeight = params("pDryDensityWeight").Value        '105
            .DryDensity = params("pDryDensity").Value                    '106
            .CalcHeadFeedBpl = params("pCalcHeadFeedBPL").Value          '107
            .GrossConcentrateTpa = params("pGrossConcentrateTPA").Value  '108
            .GrossConcentrateBpl = params("pGrossConcentrateBPL").Value  '109
            .GrossConcentrateInsol = params("pGrossConcentrateInsol").Value '110
            .GrossPebbleWtp = params("pGrossPebbleWtp").Value            '111
            .GrossPebbleTPA = params("pGrossPebbleTPA").Value            '112
            .GrossPebbleBPL = params("pGrossPebbleBPL").Value            '113
            .GrossPebbleInsol = params("pGrossPebbleInsol").Value        '114
            .GrossPebbleFe2O3 = params("pGrossPebbleFe2O3").Value        '115
            .GrossPebbleAl2O3 = params("pGrossPebbleAl2O3").Value        '116
            .GrossPebbleMgO = params("pGrossPebbleMgO").Value            '117
            .ConcentrateWtp = params("pConcentrateWtp").Value            '118
            .ConcentrateTPA = params("pConcentrateTPA").Value            '119
            .ConcentrateBPL = params("pConcentrateBPL").Value            '120
            .ConcentrateInsol = params("pConcentrateInsol").Value        '121
            .ConcentrateFe2O3 = params("pConcentrateFe2O3").Value        '122
            .ConcentrateAl2O3 = params("pConcentrateAl2O3").Value        '123
            .ConcentrateMgO = params("pConcentrateMgO").Value            '124
            .TotalProductWtp = params("pTotalProductWtp").Value          '125
            .TotalProductTpa = params("pTotalProductTPA").Value          '126
            .TotalProductBpl = params("pTotalProductBPL").Value          '127
            .TotalProductInsol = params("pTotalProductInsol").Value      '128
            .TotalProductFe2O3 = params("pTotalProductFe2O3").Value      '129
            .TotalProductAl2O3 = params("pTotalProductAl2O3").Value      '130
            .TotalProductMgO = params("pTotalProductMgO").Value          '131
            .TotalTailWtp = params("pTotalTailWtp").Value                '132
            .TotalTailTpa = params("pTotalTailTPA").Value                '133
            .TotalTailBPL = params("pTotalTailBPL").Value                '134
            .ColorTrans = params("pColorTrans").Value                    '135
            .MinableStatusTrans = params("pMinableStatusTrans").Value    '136
            .CalcOvbX = params("pCalcOvbX").Value                        '137
            .CalcMtxX = params("pCalcMtxX").Value                        '138
            .CalcTotalX = params("pCalcTotalX").Value                    '139
            .FlotationRC = params("pFlotationRC").Value                  '140
            .FlotationRecovery = params("pFlotationRecovery").Value      '141
            .TopOfSeamElevation = params("pTopOfSeamElevation").Value    '142
            .FfBpl = params("pFfBPL").Value                              '143
            .CfBpl = params("pCfBPL").Value                              '144
            .TfBPL = params("pTfBPL").Value                              '145
            .FatBpl = params("pFatBPL").Value                            '146
            .AtBpl = params("pAtBPL").Value                              '147
            .WcBpl = params("pWcBPL").Value                              '148
            '---------
            .HardpanCode = params("pHardpanCode").Value                  '148
            .GrossPebbleCaO = params("pGrossPebbleCaO").Value            '149
            .ConcentrateCaO = params("pConcentrateCaO").Value            '150
            .TotalProductCaO = params("pTotalProductCaO").Value          '151
            .CpCd = params("pCpCd").Value                                '152
            .FpCd = params("pFpCd").Value                                '153
            .LcnCd = params("pLcnCd").Value                              '154
            .ConcentrateCd = params("pConcentrateCd").Value              '155
            .TotalProductCd = params("pTotalProductCd").Value            '156
            .TpCd = params("pTpCd").Value                                '157
            .GrossPebbleCd = params("pGrossPebbleCd").Value              '158
            .TfBplCalc = params("pTfBplCalc").Value                              '159
            '----------
            .ProspStandard = params("pProspStandard").Value              '159
        End With

        'CfWtp, FfWtp, TfWtp are wrong -- we will fix them here.
        'Coarse feed TPA = .CfTpa
        'Fine feed TPA   = .FfTpa
        'Total feed TPA  = .TfTpa
        'Total feed Wtp  = .TfWtp
        With gSplit
            If .TfTPA <> 0 Then
                .FfWtp = Round((.TfWtp * .FfTpa) / .TfTPA, 2)
            Else
                .FfWtp = 0
            End If
            If .TfTPA <> 0 Then
                .CfWtp = .TfWtp - .FfWtp
            Else
                .CfWtp = 0
            End If
        End With

        'For completeness sake...
        'The matrix BPL from GEOCOMP is not really correct!
        With gSplit
            If .WcBpl = 0 Then
                'There is no way a matrix BPL could be calculated for this
                'split since we only have the pebble BPL and the feed BPL!
                .MtxBPL = 0
            End If
        End With

        'Let's fix the top of seam elevation if necessary if we can
        '10/02/2013, lss
        'Holes with "M-123456" (example) will have the .SplitElevation = Top of seam elevation
        With gSplit
            If Len(Trim(.SampleNumber)) = 8 And Mid(.SampleNumber, 1, 2) = "M-" Then
                .TopOfSeamElevation = .SplitElevation
            End If
        End With

        ClearParams(params)

        gGetSplitData = True
        Exit Function

gGetSplitDataError:
        If InStr(Err.Description, "no data found") = 0 Then
            MsgBox("Error loading split." & vbCrLf & _
                Err.Description, _
                vbOKOnly + vbExclamation, _
                "Split Loading Error")
        Else
            'Error will be handled later!
        End If

        On Error Resume Next
        ClearParams(params)
        gGetSplitData = False
    End Function

    Private Sub ZeroSplitData()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        With gSplit
            .Mine = ""                  '1
            .Section = 0                '2
            .Township = 0               '3
            .Range = 0                  '4
            .HoleLocation = ""          '5
            .Split = 0                  '6
            .MinableStatus = ""         '7
            .SplitThickness = 0         '8
            .DrillCdate = ""            '9
            .WashCdate = ""             '10
            .AreaOfInfluence = 0        '11
            .ProspectorCode = ""        '12
            .TopOfSplitDepth = 0        '13
            .BotOfSplitDepth = 0        '14
            .SampleNumber = ""          '15
            .TotalNumberSplits = 0      '16
            .RatioOfConc = 0            '17
            .CountyCode = ""            '18
            .MiningCode = ""            '19
            .PumpingCode = ""           '20
            .MetLabID = ""              '21
            .ChemLabID = ""             '22
            .Color = ""                 '23
            .SplitElevation = 0         '24
            .HoleNumber = 0             '25
            .SampleNumber2 = ""         '26
            .MtxX = 0                   '27
            .CalcLossPercent = 0        '28
            .CalcLossTPA = 0            '29
            .CalcLossBPL = 0            '30
            .WetMtxLbs = 0              '31
            .MtxGmsWet = 0              '32
            .MtxGmsDry = 0              '33
            .PercentSolidsMtx = 0       '34
            .WetFeedLbs = 0             '35
            .FeedMoistWetGms = 0        '36
            .FeedMoistDryGms = 0        '37
            .TriangleCode = ""          '38
            .CpWtp = 0                  '39
            .FpWtp = 0                  '40
            .TfWtp = 0                  '41
            .WcWtp = 0                  '42
            .CnWtp = 0                  '43
            .TpWtp = 0                  '44
            .FfWtp = 0                  '45
            .CfWtp = 0                  '46
            .CpTpa = 0                  '47
            .FpTpa = 0                  '48
            .TfTPA = 0                  '49
            .WcTpa = 0                  '50
            .CnTpa = 0                  '51
            .TpTpa = 0                  '52
            .FfTpa = 0                  '53
            .CfTpa = 0                  '54
            .FatTpa = 0                 '55
            .AtTpa = 0                  '56
            .AcnTpa = 0                 '57
            .MtxTPA = 0                 '58
            .CpBPL = 0                  '59
            .FpBpl = 0                  '60
            .LcnBpl = 0                 '61
            .TpBpl = 0                  '62
            .MtxBPL = 0                 '63
            .CnBpl = 0                  '64
            .CpInsol = 0                '65
            .CpFe2O3 = 0                '66
            .CpAl2O3 = 0                '67
            .CpMgO = 0                  '68
            .CpCaO = 0                  '69
            .FpInsol = 0                '70
            .FpFe2O3 = 0                '71
            .FpAl2O3 = 0                '72
            .FpMgO = 0                  '73
            .FpCaO = 0                  '74
            .FpFeAl = 0                 '75
            .FpCaOP2O5 = 0              '76
            .LcnInsol = 0               '77
            .LcnFe2O3 = 0               '78
            .LcnAl2O3 = 0               '79
            .LcnMgO = 0                 '80
            .LcnCaO = 0                 '81
            .LcnFeAl = 0                '82
            .TpInsol = 0                '83
            .TpFe2O3 = 0                '84
            .TpAl2O3 = 0                '85
            .TpMgO = 0                  '86
            .TpCaO = 0                  '87
            .TpFeAl = 0                 '88
            .TpCaOP2O5 = 0              '89
            .MtxInsol = 0               '90
            .CnInsol = 0                '91
            .CnCaO = 0                  '92
            .CnFeAl = 0                 '93
            .CpGrams = 0                '94
            .FpGrams = 0                '95
            .TfGrams = 0                '96
            .FatGrams = 0               '97
            .AtGrams = 0                '98
            .LcnGrams = 0               '99
            .WetDensityVolume = 0       '100
            .WetDensityWeight = 0       '101
            .WetDensity = 0             '102
            .DryDensityVolume = 0       '103
            .DryDensityWeight = 0       '104
            .DryDensity = 0             '105
            .CalcHeadFeedBpl = 0        '106
            .GrossConcentrateTpa = 0    '107
            .GrossConcentrateBpl = 0    '108
            .GrossConcentrateInsol = 0  '109
            .GrossPebbleWtp = 0         '110
            .GrossPebbleTPA = 0         '111
            .GrossPebbleBPL = 0         '112
            .GrossPebbleInsol = 0       '113
            .GrossPebbleFe2O3 = 0       '114
            .GrossPebbleAl2O3 = 0       '115
            .GrossPebbleMgO = 0         '116
            .ConcentrateWtp = 0         '117
            .ConcentrateTPA = 0         '118
            .ConcentrateBPL = 0         '119
            .ConcentrateInsol = 0       '120
            .ConcentrateFe2O3 = 0       '121
            .ConcentrateAl2O3 = 0       '122
            .ConcentrateMgO = 0         '123
            .TotalProductWtp = 0        '124
            .TotalProductTpa = 0        '125
            .TotalProductBpl = 0        '126
            .TotalProductInsol = 0      '127
            .TotalProductFe2O3 = 0      '128
            .TotalProductAl2O3 = 0      '129
            .TotalProductMgO = 0        '130
            .TotalTailWtp = 0           '131
            .TotalTailTpa = 0           '132
            .TotalTailBPL = 0           '133
            .ColorTrans = ""            '134
            .MinableStatusTrans = ""    '135
            .CalcOvbX = 0               '136
            .CalcMtxX = 0               '137
            .CalcTotalX = 0             '138
            .FlotationRC = 0            '139
            .FlotationRecovery = 0      '140
            .TopOfSeamElevation = 0     '141
            .FfBpl = 0                  '142
            .CfBpl = 0                  '143
            .TfBPL = 0                  '144
            .FatBpl = 0                 '145
            .AtBpl = 0                  '146
            .WcBpl = 0                  '147
        End With
    End Sub

    Public Function gGetHoleLocation(ByVal aSection As Integer, _
                                      ByVal aTownship As Integer, _
                                      ByVal aRange As Integer, _
                                      ByVal aHoleLocation As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SecStr As String
        Dim TwpStr As String
        Dim RgeStr As String

        If aSection > 9 Then
            SecStr = Trim(CStr(aSection))
        Else
            SecStr = "0" + Trim(CStr(aSection))
        End If
        If aTownship > 9 Then
            TwpStr = Trim(CStr(aTownship))
        Else
            TwpStr = "0" + Trim(CStr(aTownship))
        End If
        If aRange > 9 Then
            RgeStr = Trim(CStr(aRange))
        Else
            RgeStr = "0" + Trim(CStr(aRange))
        End If

        gGetHoleLocation = SecStr & "-" & TwpStr & "-" & RgeStr & _
                           " " & aHoleLocation
    End Function

    Public Function gGetHoleLocationTrs(ByVal aSection As Integer, _
                                        ByVal aTownship As Integer, _
                                        ByVal aRange As Integer, _
                                        ByVal aHoleLocation As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SecStr As String
        Dim TwpStr As String
        Dim RgeStr As String

        If aSection > 9 Then
            SecStr = Trim(CStr(aSection))
        Else
            SecStr = "0" + Trim(CStr(aSection))
        End If
        If aTownship > 9 Then
            TwpStr = Trim(CStr(aTownship))
        Else
            TwpStr = "0" + Trim(CStr(aTownship))
        End If
        If aRange > 9 Then
            RgeStr = Trim(CStr(aRange))
        Else
            RgeStr = "0" + Trim(CStr(aRange))
        End If

        gGetHoleLocationTrs = TwpStr & "-" & RgeStr & "-" & SecStr & _
                              " " & aHoleLocation
    End Function

    Public Function gGetHoleLocationTitled(ByVal aSection As Integer, _
                                           ByVal aTownship As Integer, _
                                           ByVal aRange As Integer, _
                                           ByVal aHoleLocation As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SecStr As String
        Dim TwpStr As String
        Dim RgeStr As String

        If aSection > 9 Then
            SecStr = Trim(CStr(aSection))
        Else
            SecStr = "0" + Trim(CStr(aSection))
        End If
        If aTownship > 9 Then
            TwpStr = Trim(CStr(aTownship))
        Else
            TwpStr = "0" + Trim(CStr(aTownship))
        End If
        If aRange > 9 Then
            RgeStr = Trim(CStr(aRange))
        Else
            RgeStr = "0" + Trim(CStr(aRange))
        End If

        gGetHoleLocationTitled = "Sec " & SecStr & "  Twp " & TwpStr & "  Rge " & _
                                 RgeStr & "  Hole " + aHoleLocation
    End Function

    Public Function gGetHoleLocationTitled2(ByVal aSection As Integer, _
                                            ByVal aTownship As Integer, _
                                            ByVal aRange As Integer, _
                                            ByVal aHoleLocation As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SecStr As String
        Dim TwpStr As String
        Dim RgeStr As String

        If aSection > 9 Then
            SecStr = Trim(CStr(aSection))
        Else
            SecStr = "0" + Trim(CStr(aSection))
        End If
        If aTownship > 9 Then
            TwpStr = Trim(CStr(aTownship))
        Else
            TwpStr = "0" + Trim(CStr(aTownship))
        End If
        If aRange > 9 Then
            RgeStr = Trim(CStr(aRange))
        Else
            RgeStr = "0" + Trim(CStr(aRange))
        End If

        gGetHoleLocationTitled2 = "Twp " & TwpStr & "  Rge " & RgeStr & "  Sec " & _
                                  SecStr & "  Hole " + aHoleLocation
    End Function

    Public Function gGetHoleLocationTitled3(ByVal aSection As Integer, _
                                            ByVal aTownship As Integer, _
                                            ByVal aRange As Integer, _
                                            ByVal aHoleLocation As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SecStr As String
        Dim TwpStr As String
        Dim RgeStr As String

        If aSection > 9 Then
            SecStr = Trim(CStr(aSection))
        Else
            SecStr = "0" + Trim(CStr(aSection))
        End If
        If aTownship > 9 Then
            TwpStr = Trim(CStr(aTownship))
        Else
            TwpStr = "0" + Trim(CStr(aTownship))
        End If
        If aRange > 9 Then
            RgeStr = Trim(CStr(aRange))
        Else
            RgeStr = "0" + Trim(CStr(aRange))
        End If

        gGetHoleLocationTitled3 = "T" & TwpStr & " R" & RgeStr & " S" & _
                                  SecStr & "  " + aHoleLocation
    End Function

    Public Function gGetHoleLocationShort(ByVal aSection As Integer, _
                                          ByVal aTownship As Integer, _
                                          ByVal aRange As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SecStr As String
        Dim TwpStr As String
        Dim RgeStr As String

        If aSection > 9 Then
            SecStr = Trim(CStr(aSection))
        Else
            SecStr = "0" + Trim(CStr(aSection))
        End If
        If aTownship > 9 Then
            TwpStr = Trim(CStr(aTownship))
        Else
            TwpStr = "0" + Trim(CStr(aTownship))
        End If
        If aRange > 9 Then
            RgeStr = Trim(CStr(aRange))
        Else
            RgeStr = "0" + Trim(CStr(aRange))
        End If

        gGetHoleLocationShort = SecStr & "-" & TwpStr & "-" & RgeStr
    End Function

    Public Function gGetHoleLocationShortDot(ByVal aSection As Integer, _
                                             ByVal aTownship As Integer, _
                                             ByVal aRange As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SecStr As String
        Dim TwpStr As String
        Dim RgeStr As String

        If aSection > 9 Then
            SecStr = Trim(CStr(aSection))
        Else
            SecStr = "0" + Trim(CStr(aSection))
        End If
        If aTownship > 9 Then
            TwpStr = Trim(CStr(aTownship))
        Else
            TwpStr = "0" + Trim(CStr(aTownship))
        End If
        If aRange > 9 Then
            RgeStr = Trim(CStr(aRange))
        Else
            RgeStr = "0" + Trim(CStr(aRange))
        End If

        gGetHoleLocationShortDot = SecStr & "." & TwpStr & "." & RgeStr
    End Function

    Public Function gCompHoleExists(ByVal aMine As String, _
                                    ByVal aSec As Integer, _
                                    ByVal aTwp As Integer, _
                                    ByVal aRge As Integer, _
                                    ByVal aHole As String, _
                                    ByVal aProspStandard As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim HoleCount As Integer

        On Error GoTo gCompHoleExistsError

        params = gDBParams

        params.Add("pMineName", aMine, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHole, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        'Assum that a legitimate prospect standard has been passed
        params.Add("pProspStandard", aProspStandard, ORAPARM_INPUT)
        params("pProspStandard").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        'PROCEDURE hole_exists
        'pMineName           IN VARCHAR2,
        'pSection            IN NUMBER,
        'pTownship           IN NUMBER,
        'pRange              IN NUMBER,
        'pHoleLocation       IN VARCHAR2,
        'pProspStandard      IN VARCHAR2,
        'pResult             IN OUT NUMBER)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect.hole_exists(:pMineName, " + _
                      ":pSection, :pTownship, :pRange, :pHoleLocation, :pProspStandard, :pResult);end;", ORASQL_FAILEXEC)
        HoleCount = params("pResult").Value
        ClearParams(params)

        'HoleCount should be 0 or 1.
        'If HoleCount > 1 then will return False (Hole does not exist).
        If HoleCount = 0 Then
            gCompHoleExists = False
        Else
            If HoleCount = 1 Then
                gCompHoleExists = True
            Else
                gCompHoleExists = False
            End If
        End If

        Exit Function

gCompHoleExistsError:
        MsgBox("Error determining if composite hole exists." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Composite Hole Exists Error")

        gCompHoleExists = False
        ClearParams(params)
    End Function

    Public Function gProductX(ByVal aPebbTpa As Long, _
                              ByVal aConcTpa As Long, _
                              ByVal aMtxThk As Single, _
                              ByVal aOvbThk As Single, _
                              ByVal aMode As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim TotThk As Single
        Dim TotTpa As Double

        'Mode = Total, Matrix
        gProductX = 0

        Select Case aMode
            Case Is = "Total"
                TotThk = aMtxThk + aOvbThk
                TotTpa = aPebbTpa + aConcTpa
                If TotTpa <> 0 Then
                    gProductX = Round(((TotThk * 43560) / 27) / TotTpa, 2)
                Else
                    gProductX = 0
                End If

            Case Is = "Matrix"
                TotTpa = aPebbTpa + aConcTpa
                If TotTpa <> 0 Then
                    gProductX = Round(((aMtxThk * 43560) / 27) / TotTpa, 2)
                Else
                    gProductX = 0
                End If
        End Select
    End Function

    Public Function gHoleLocationOk(ByVal aHole As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ThisHole As String
        Dim ThisChar As String
        Dim CharString As String
        Dim ThisNum As Integer

        CharString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

        ThisHole = Trim(aHole)

        gHoleLocationOk = True

        If Len(ThisHole) <> 3 Then
            gHoleLocationOk = False
            Exit Function
        End If

        ThisChar = Mid(ThisHole, 1, 1)

        If InStr(CharString, ThisChar) = 0 Then
            gHoleLocationOk = False
            Exit Function
        End If

        ThisNum = Val(Mid(ThisHole, 2))
        If ThisNum < 1 Or ThisNum > 16 Then
            gHoleLocationOk = False
        End If
    End Function

    Public Function gHoleLocationOkNew(ByVal aHole As String, _
                                       ByVal aProspGridType As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ThisHole As String
        Dim ThisChar As String
        Dim CharString As String
        Dim ThisNum As Integer

        Dim Num1st As Integer
        Dim Num2nd As Integer

        gHoleLocationOkNew = False

        If aProspGridType = "Alpha-numeric" Then
            CharString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

            ThisHole = Trim(aHole)

            gHoleLocationOkNew = True

            If Len(ThisHole) <> 3 Then
                gHoleLocationOkNew = False
                Exit Function
            End If

            ThisChar = Mid(ThisHole, 1, 1)

            If InStr(CharString, ThisChar) = 0 Then
                gHoleLocationOkNew = False
                Exit Function
            End If

            ThisNum = Val(Mid(ThisHole, 2))
            If ThisNum < 1 Or ThisNum > 16 Then
                gHoleLocationOkNew = False
            End If
        End If

        If aProspGridType = "Numeric" Then
            If Not IsNumeric(aHole) Then
                gHoleLocationOkNew = False
                Exit Function
            End If

            gHoleLocationOkNew = False

            ThisHole = Format(Val(aHole), "000#")

            Num1st = Val(Mid(ThisHole, 1, 2))
            Num2nd = Val(Mid(ThisHole, 3))

            If ((Num1st >= 2 And Num1st <= 33) Or _
               (Num1st >= 73 And Num1st <= 78)) And _
               (Num2nd >= 33 And Num2nd <= 72) Then
                gHoleLocationOkNew = True
            End If
        End If
    End Function

    Public Function gGetNumSplits(ByVal aMine As String, _
                                  ByVal aSec As Integer, _
                                  ByVal aTwp As Integer, _
                                  ByVal aRge As Integer, _
                                  ByVal aHole As String, _
                                  ByVal aProspStandard As String) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetNumSplitsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim SplitCount As Integer

        'Special parameter fix!
        If StrConv(aProspStandard, vbUpperCase) = "100% PROSPECT" Then
            aProspStandard = "100%PROSPECT"
        End If
        If StrConv(aProspStandard, vbUpperCase) = "CATALOG" Then
            aProspStandard = "CATALOG"
        End If

        params = gDBParams

        params.Add("pMineName", aMine, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHole, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspStandard", aProspStandard, ORAPARM_INPUT)
        params("pProspStandard").serverType = ORATYPE_VARCHAR2

        params.Add("pNumSplits", 0, ORAPARM_OUTPUT)
        params("pNumSplits").serverType = ORATYPE_NUMBER

        'PROCEDURE split_count
        'pMineName           IN     VARCHAR2,
        'pSection            IN     NUMBER,
        'pTownship           IN     NUMBER,
        'pRange              IN     NUMBER,
        'pHoleLocation       IN     VARCHAR2,
        'pProspStandard      IN     VARCHAR2,
        'pNumSplits          IN OUT NUMBER)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect.split_count(:pMineName, " + _
                      ":pSection, :pTownship, :pRange, :pHoleLocation, :pProspStandard, :pNumSplits);end;", ORASQL_FAILEXEC)
        SplitCount = params("pNumSplits").Value
        ClearParams(params)

        gGetNumSplits = SplitCount

        Exit Function

gGetNumSplitsError:
        MsgBox("Error determining number of splits." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Number Splits Count Error")

        gGetNumSplits = 0
        ClearParams(params)
    End Function

    Public Sub gGetDrillHole(ByVal aMineName As String, _
                             ByVal aProspStandard As String, _
                             ByVal aSec As Integer, _
                             ByVal aTwp As Integer, _
                             ByVal aRge As Integer, _
                             ByVal aHloc As String, _
                             ByRef aProspHole As gProspectCompositeType2)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'NOTE:
        'As of 07/23/2003, get_composite2 returns:
        ' 1) mine_name
        ' 2) hole_location
        ' 3) section
        ' 4) township
        ' 5) range
        ' 6) x_sp_cdnt
        ' 7) y_sp_cdnt
        ' 8) drill_cdate
        ' 9) analysis_cdate
        '10) area_influence
        '11) ovb_thck
        '12) mtx_thck
        '13) mtx_wet_density
        '14) mtx_pct_solids
        '15) mtx_x
        '16) hole_elevation
        '17) split_total_num
        '18) pit_bottom_elevation
        '19) triangle_code
        '20) prosp_code
        '21) mtx_tons
        '22) split_sum
        '23) CPB_BPL
        '24) FPB_BPL
        '25) TFD_BPL
        '26) CFD_BPL
        '27) FFD_BPL
        '28) CNC_BPL
        '29) TPB_BPL
        '30) TPR_BPL
        '31) MTX_BPL
        '32) CPB_TPA
        '33) FPB_TPA
        '34) TFD_TPA
        '35) CFD_TPA
        '36) FFD_TPA
        '37) CNC_TPA
        '38) TPB_TPA
        '39) TPR_TPA
        '40) WCL_TPA
        '41) CPB_WTP
        '42) FPB_WTP
        '43) TFD_WTP
        '44) CNC_WTP
        '45) TPB_WTP
        '46) TPR_WTP
        '47) WCL_WTP
        '48) CPB_INS
        '49) FPB_INS
        '50) TPB_INS
        '51) CNC_INS
        '52) TPR_INS
        '53) CPB_FE
        '54) FPB_FE
        '55) TPB_FE
        '56) CNC_FE
        '57) TPR_FE
        '58) CPB_AL
        '59) FPB_AL
        '60) TPB_AL
        '61) CNC_AL
        '62) TPR_AL
        '63) CPB_MG
        '64) FPB_MG
        '65) TPB_MG
        '66) CNC_MG
        '67) TPR_MG
        '68) CPB_CA
        '69) FPB_CA
        '70) TPB_CA
        '71) CNC_CA
        '72) TPR_CA

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        params = gDBParams

        '1
        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        '2
        params.Add("pHoleLocation", aHloc, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        '3
        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        '4
        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        '5
        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        '6
        params.Add("pProspStandard", aProspStandard, ORAPARM_INPUT)
        params("pProspStandard").serverType = ORATYPE_VARCHAR2

        '7
        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_composite2
        'pMineName      IN     VARCHAR2,
        'pHoleLocation  IN     VARCHAR2,
        'pSection       IN     NUMBER,
        'pTownship      IN     NUMBER,
        'pRange         IN     NUMBER,
        'pProspStandard IN     VARCHAR2,
        'pResult        IN OUT c_composite)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect.get_composite2(:pMineName, " + _
                                       ":pHoleLocation, :pSection, :pTownship, " + _
                                       ":pRange, :pProspStandard, :pResult);end;", ORASQL_FAILEXEC)

        mProspHoleInfo = params("pResult").Value
        ClearParams(params)

        ZeroDrillHole(aProspHole)

        If mProspHoleInfo.RecordCount = 1 Then
            With aProspHole
                .MineName = mProspHoleInfo("mine_name").Value
                .ProspStandard = mProspHoleInfo("prosp_standard").Value
                .HoleLocation = mProspHoleInfo("hole_location").Value               '2
                .Section = mProspHoleInfo("section").Value                          '3
                .Township = mProspHoleInfo("township").Value                        '4
                .Range = mProspHoleInfo("range").Value                              '5
                .XSpCdnt = mProspHoleInfo("x_sp_cdnt").Value                        '6
                .YSpCdnt = mProspHoleInfo("y_sp_cdnt").Value                        '7
                .DrillCdate = mProspHoleInfo("drill_cdate").Value                   '8

                If Not IsDBNull(mProspHoleInfo("analysis_cdate").Value) Then
                    .AnalysisCdate = mProspHoleInfo("analysis_cdate").Value         '9
                Else
                    .AnalysisCdate = ""
                End If

                .AreaInfluence = mProspHoleInfo("area_influence").Value             '10
                .OvbThck = mProspHoleInfo("ovb_thck").Value                         '11
                .MtxThck = mProspHoleInfo("mtx_thck").Value                         '12
                .WstThck = mProspHoleInfo("wst_thck").Value                         '
                .MtxWetDensity = mProspHoleInfo("mtx_wet_density").Value            '13
                .MtxPctSolids = mProspHoleInfo("mtx_pct_solids").Value              '14
                .MtxX = mProspHoleInfo("mtx_x").Value                               '15
                .HoleElevation = mProspHoleInfo("hole_elevation").Value             '16
                .SplitTotalNum = mProspHoleInfo("split_total_num").Value            '17
                .PitBottomElevation = mProspHoleInfo("pit_bottom_elevation").Value  '18

                If Not IsDBNull(mProspHoleInfo("triangle_code").Value) Then
                    .TriangleCode = mProspHoleInfo("triangle_code").Value           '19
                Else
                    .TriangleCode = ""
                End If

                If Not IsDBNull(mProspHoleInfo("prosp_code").Value) Then
                    .ProspCode = mProspHoleInfo("prosp_code").Value                 '20
                Else
                    .ProspCode = ""
                End If

                .MtxTons = mProspHoleInfo("mtx_tons").Value                         '21

                If Not IsDBNull(mProspHoleInfo("split_sum").Value) Then
                    .SplitSum = mProspHoleInfo("split_sum").Value                   '22
                Else
                    .SplitSum = ""
                End If

                '----------
                .CpbBpl = mProspHoleInfo("cpb_bpl").Value         '23
                .FpbBpl = mProspHoleInfo("fpb_bpl").Value         '24
                .TfdBpl = mProspHoleInfo("tfd_bpl").Value         '25
                .CfdBpl = mProspHoleInfo("cfd_bpl").Value         '26
                .FfdBpl = mProspHoleInfo("ffd_bpl").Value         '27
                .CncBpl = mProspHoleInfo("cnc_bpl").Value         '28
                .TpbBpl = mProspHoleInfo("tpb_bpl").Value         '29
                .TprBpl = mProspHoleInfo("tpr_bpl").Value         '30
                .MtxBPL = mProspHoleInfo("mtx_bpl").Value         '31
                '----------
                .CpbTpa = mProspHoleInfo("cpb_tpa").Value         '32
                .FpbTpa = mProspHoleInfo("fpb_tpa").Value         '33
                .TfdTpa = mProspHoleInfo("tfd_tpa").Value         '34
                .CfdTpa = mProspHoleInfo("cfd_tpa").Value         '35
                .FfdTpa = mProspHoleInfo("ffd_tpa").Value         '36
                .CncTpa = mProspHoleInfo("cnc_tpa").Value         '37
                .TpbTpa = mProspHoleInfo("tpb_tpa").Value         '38
                .TprTpa = mProspHoleInfo("tpr_tpa").Value         '39
                .WclTpa = mProspHoleInfo("wcl_tpa").Value         '40
                '----------
                .CpbWtp = mProspHoleInfo("cpb_wtp").Value         '41
                .FpbWtp = mProspHoleInfo("fpb_wtp").Value         '42
                .TfdWtp = mProspHoleInfo("tfd_wtp").Value         '43
                .CncWtp = mProspHoleInfo("cnc_wtp").Value         '44
                .TpbWtp = mProspHoleInfo("tpb_wtp").Value         '45
                .TprWtp = mProspHoleInfo("tpr_wtp").Value         '46
                .WclWtp = mProspHoleInfo("wcl_wtp").Value         '47
                '----------
                .CpbIns = mProspHoleInfo("cpb_ins").Value         '48
                .FpbIns = mProspHoleInfo("fpb_ins").Value         '49
                .TpbIns = mProspHoleInfo("tpb_ins").Value         '50
                .CncIns = mProspHoleInfo("cnc_ins").Value         '51
                .TprIns = mProspHoleInfo("tpr_ins").Value         '52
                '----------
                .CpbFe = mProspHoleInfo("cpb_fe").Value           '53
                .FpbFe = mProspHoleInfo("fpb_fe").Value           '54
                .TpbFe = mProspHoleInfo("tpb_fe").Value           '55
                .CncFe = mProspHoleInfo("cnc_fe").Value           '56
                .TprFe = mProspHoleInfo("tpr_fe").Value           '57
                '----------
                .CpbAl = mProspHoleInfo("cpb_al").Value           '58
                .FpbAl = mProspHoleInfo("fpb_al").Value           '59
                .TpbAl = mProspHoleInfo("tpb_al").Value           '60
                .CncAl = mProspHoleInfo("cnc_al").Value           '61
                .TprAl = mProspHoleInfo("tpr_al").Value           '62
                '----------
                .CpbMg = mProspHoleInfo("cpb_mg").Value           '63
                .FpbMg = mProspHoleInfo("fpb_mg").Value           '64
                .TpbMg = mProspHoleInfo("tpb_mg").Value           '65
                .CncMg = mProspHoleInfo("cnc_mg").Value           '66
                .TprMg = mProspHoleInfo("tpr_mg").Value           '67
                '----------
                .CpbCa = mProspHoleInfo("cpb_ca").Value           '68
                .FpbCa = mProspHoleInfo("fpb_ca").Value           '69
                .TpbCa = mProspHoleInfo("tpb_ca").Value           '70
                .CncCa = mProspHoleInfo("cnc_ca").Value           '71
                .TprCa = mProspHoleInfo("tpr_ca").Value           '72

                'Need to calculate total feed
                'Don not want to average in any 0 values!
                .TfdTpa2 = .CfdTpa + .FfdTpa
                .TfdWtp2 = .TfdWtp  'Fine feed Wtp & Coarse feed Wtp are not passed by
                'get_composite2 so will not calculate at this time
                If IIf(.CfdBpl <> 0, .CfdTpa, 0) + IIf(.FfdBpl <> 0, .FfdTpa, 0) <> 0 Then
                    .TfdBpl2 = gRound(((.CfdTpa * .CfdBpl) + (.FfdTpa * .FfdBpl)) / _
                               (IIf(.CfdBpl <> 0, .CfdTpa, 0) + _
                               IIf(.FfdBpl <> 0, .FfdTpa, 0)), 1)
                Else
                    .TfdBpl2 = 0
                End If

                'Recalculate total pebble BPL & impurities also
                'Do not want to average in any 0 values!
                If IIf(.CpbBpl <> 0, .CpbTpa, 0) + IIf(.FpbBpl <> 0, .FpbTpa, 0) <> 0 Then
                    .TpbBpl2 = gRound(((.CpbTpa * .CpbBpl) + (.FpbTpa * .FpbBpl)) / _
                                     (IIf(.CpbBpl <> 0, .CpbTpa, 0) + _
                                      IIf(.FpbBpl <> 0, .FpbTpa, 0)), 1)
                Else
                    .TpbBpl2 = 0
                End If

                If IIf(.CpbFe <> 0, .CpbTpa, 0) + IIf(.FpbFe <> 0, .FpbTpa, 0) <> 0 Then
                    .TpbFe2 = gRound(((.CpbTpa * .CpbFe) + (.FpbTpa * .FpbFe)) / _
                                     (IIf(.CpbFe <> 0, .CpbTpa, 0) + _
                                      IIf(.FpbFe <> 0, .FpbTpa, 0)), 2)
                Else
                    '10/08/2012, lss  Was .TpbFe = 0
                    .TpbFe2 = 0
                End If

                If IIf(.CpbAl <> 0, .CpbTpa, 0) + IIf(.FpbAl <> 0, .FpbTpa, 0) <> 0 Then
                    .TpbAl2 = gRound(((.CpbTpa * .CpbAl) + (.FpbTpa * .FpbAl)) / _
                                     (IIf(.CpbAl <> 0, .CpbTpa, 0) + _
                                      IIf(.FpbAl <> 0, .FpbTpa, 0)), 2)
                Else
                    .TpbAl2 = 0
                End If

                If IIf(.CpbMg <> 0, .CpbTpa, 0) + IIf(.FpbMg <> 0, .FpbTpa, 0) <> 0 Then
                    .TpbMg2 = gRound(((.CpbTpa * .CpbMg) + (.FpbTpa * .FpbMg)) / _
                                     (IIf(.CpbMg <> 0, .CpbTpa, 0) + _
                                      IIf(.FpbMg <> 0, .FpbTpa, 0)), 2)
                Else
                    .TpbMg2 = 0
                End If

                If IIf(.CpbIns <> 0, .CpbTpa, 0) + IIf(.FpbIns <> 0, .FpbTpa, 0) <> 0 Then
                    .TpbIns2 = gRound(((.CpbTpa * .CpbIns) + (.FpbTpa * .FpbIns)) / _
                                     (IIf(.CpbIns <> 0, .CpbTpa, 0) + _
                                      IIf(.FpbIns <> 0, .FpbTpa, 0)), 2)
                Else
                    .TpbIns2 = 0
                End If

                '10/08/2012, lss
                If IIf(.CpbCa <> 0, .CpbTpa, 0) + IIf(.FpbCa <> 0, .FpbTpa, 0) <> 0 Then
                    .TpbCa2 = gRound(((.CpbTpa * .CpbCa) + (.FpbTpa * .FpbCa)) / _
                                     (IIf(.CpbCa <> 0, .CpbTpa, 0) + _
                                      IIf(.FpbCa <> 0, .FpbTpa, 0)), 2)
                Else
                    .TpbCa2 = 0
                End If
            End With
        End If
    End Sub

    Public Sub ZeroDrillHole(ByRef aProspHole As gProspectCompositeType2)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        With aProspHole
            .MineName = ""                     '1
            .HoleLocation = ""                 '2
            .Section = 0                       '3
            .Township = 0                      '4
            .Range = 0                         '5
            .XSpCdnt = 0                       '6
            .YSpCdnt = 0                       '7
            .DrillCdate = ""                   '8
            .AnalysisCdate = ""                '9
            .AreaInfluence = 0                 '10
            .OvbThck = 0                       '11
            .MtxThck = 0                       '12
            .MtxWetDensity = 0                 '13
            .MtxPctSolids = 0                  '14
            .MtxX = 0                          '15
            .HoleElevation = 0                 '16
            .SplitTotalNum = 0                 '17
            .PitBottomElevation = 0            '18
            .TriangleCode = ""                 '19
            .ProspCode = ""                    '20
            .MtxTons = 0                       '21
            .SplitSum = ""                     '22
            '----------
            .CpbBpl = 0                        '23
            .FpbBpl = 0                        '24
            .TfdBpl = 0                        '25
            .CfdBpl = 0                        '26
            .FfdBpl = 0                        '27
            .CncBpl = 0                        '28
            .TpbBpl = 0                        '29
            .TprBpl = 0                        '30
            .MtxBPL = 0                        '31
            '----------
            .CpbTpa = 0                        '32
            .FpbTpa = 0                        '33
            .TfdTpa = 0                        '34
            .CfdTpa = 0                        '35
            .FfdTpa = 0                        '36
            .CncTpa = 0                        '37
            .TpbTpa = 0                        '38
            .TprTpa = 0                        '39
            .WclTpa = 0                        '40
            '----------
            .CpbWtp = 0                        '41
            .FpbWtp = 0                        '42
            .TfdWtp = 0                        '43
            .CncWtp = 0                        '44
            .TpbWtp = 0                        '45
            .TprWtp = 0                        '46
            .WclWtp = 0                        '47
            '----------
            .CpbIns = 0                        '48
            .FpbIns = 0                        '49
            .TpbIns = 0                        '50
            .CncIns = 0                        '51
            .TprIns = 0                        '52
            '----------
            .CpbFe = 0                         '53
            .FpbFe = 0                         '54
            .TpbFe = 0                         '55
            .CncFe = 0                         '56
            .TprFe = 0                         '57
            '----------
            .CpbAl = 0                         '58
            .FpbAl = 0                         '59
            .TpbAl = 0                         '60
            .CncAl = 0                         '61
            .TprAl = 0                         '62
            '----------
            .CpbMg = 0                         '63
            .FpbMg = 0                         '64
            .TpbMg = 0                         '65
            .CncMg = 0                         '66
            .TprMg = 0                         '67
            '----------
            .CpbCa = 0                         '68
            .FpbCa = 0                         '69
            .TpbCa = 0                         '70
            .CncCa = 0                         '71
            .TprCa = 0                         '72
        End With
    End Sub

    Public Function gGetAnalTotal(ByVal aValue1 As Single, _
                                  ByVal aTons1 As Long, _
                                  ByVal aValue2 As Single, _
                                  ByVal aTons2 As Long, _
                                  ByVal aRound As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim TonsWithAnal As Double
        Dim TonAnals As Double

        TonsWithAnal = 0
        TonAnals = 0

        If aTons1 <> 0 Then
            TonsWithAnal = aTons1
        End If
        If aTons2 <> 0 Then
            TonsWithAnal = TonsWithAnal + aTons2
        End If

        TonAnals = (aValue1 * aTons1) + (aValue2 * aTons2)

        If TonsWithAnal <> 0 Then
            gGetAnalTotal = gRound(TonAnals / TonsWithAnal, aRound)
        Else
            gGetAnalTotal = 0
        End If
    End Function

    Public Function gGetTotalValue(ByVal aValue1 As Single, _
                                   ByVal aTpa1 As Single, _
                                   ByVal aValue2 As Single, _
                                   ByVal aTpa2 As Single, _
                                   ByVal aRound As Integer) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Averages the two values together -- doesn't care if one of the
        'values is zero!
        gGetTotalValue = 0

        If aTpa1 + aTpa2 <> 0 Then
            gGetTotalValue = gRound((aValue1 * aTpa1 + aValue2 * aTpa2) / _
                             (aTpa1 + aTpa2), aRound)
        Else
            gGetTotalValue = 0
        End If
    End Function

    Public Function gGetTotalValue4(ByVal aValue1 As Single, _
                                    ByVal aTpa1 As Single, _
                                    ByVal aValue2 As Single, _
                                    ByVal aTpa2 As Single, _
                                    ByVal aValue3 As Single, _
                                    ByVal aTpa3 As Single, _
                                    ByVal aValue4 As Single, _
                                    ByVal aTpa4 As Single, _
                                    ByVal aRound As Integer) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Will not average in zero values!

        Dim TpaWVal As Double
        Dim TpaVal As Double

        gGetTotalValue4 = 0

        TpaWVal = 0
        TpaVal = 0

        TpaVal = aTpa1 * aValue1 + aTpa2 * aValue2 + _
                 aTpa3 * aValue3 + aTpa4 * aValue4

        If aValue1 > 0 Then
            TpaWVal = TpaWVal + aTpa1
        End If
        If aValue2 > 0 Then
            TpaWVal = TpaWVal + aTpa2
        End If
        If aValue3 > 0 Then
            TpaWVal = TpaWVal + aTpa3
        End If
        If aValue4 > 0 Then
            TpaWVal = TpaWVal + aTpa4
        End If

        If TpaWVal <> 0 Then
            gGetTotalValue4 = gRound(TpaVal / TpaWVal, aRound)
        Else
            gGetTotalValue4 = 0
        End If
    End Function

    Public Function gGetTotalValue2(ByVal aValue1 As Single, _
                                    ByVal aTpa1 As Single, _
                                    ByVal aValue2 As Single, _
                                    ByVal aTpa2 As Single, _
                                    ByVal aRound As Integer) As Single

        'Will not average in zero values!

        Dim TpaWVal As Double
        Dim TpaVal As Double

        gGetTotalValue2 = 0

        TpaWVal = 0
        TpaVal = 0

        TpaVal = aTpa1 * aValue1 + aTpa2 * aValue2

        If aValue1 > 0 Then
            TpaWVal = TpaWVal + aTpa1
        End If
        If aValue2 > 0 Then
            TpaWVal = TpaWVal + aTpa2
        End If

        If TpaWVal <> 0 Then
            gGetTotalValue2 = gRound(TpaVal / TpaWVal, aRound)
        Else
            gGetTotalValue2 = 0
        End If
    End Function

    Public Function gGetDrillDate(ByVal aDateStr As String) As Date

        '**********************************************************************
        ' Try to make a mm/dd/yyyy date out of the date string that is passed
        ' in.  If a valid date cannot be determined then will return
        ' 12/31/8888.  The function will always return a date (it may be
        ' 12/31/8888 indicating the function has failed).
        '**********************************************************************

        Dim TempString As String
        Dim TempDate As Date
        Dim ThisDate As String

        aDateStr = Trim(aDateStr)
        ThisDate = ""

        Select Case Len(aDateStr)
            Case 0
                ThisDate = "12/31/8888"

            Case 2
                'Assume that this is an older date and it must be
                'in the 1900's
                ThisDate = "01/01/19" + aDateStr

            Case 4
                ThisDate = "01/01/" & aDateStr

            Case 5
                'May be dmmyy or may be mm/yy!
                'dmmyy
                If Mid(aDateStr, 3, 1) = "/" Then
                    If Mid(aDateStr, 1, 2) = "01" Or Mid(aDateStr, 1, 2) = "02" Or _
                       Mid(aDateStr, 1, 2) = "03" Or Mid(aDateStr, 1, 2) = "04" Or _
                       Mid(aDateStr, 1, 2) = "05" Or Mid(aDateStr, 1, 2) = "06" Or _
                       Mid(aDateStr, 1, 2) = "07" Or Mid(aDateStr, 1, 2) = "08" Or _
                       Mid(aDateStr, 1, 2) = "09" Or Mid(aDateStr, 1, 2) = "10" Or _
                       Mid(aDateStr, 1, 2) = "11" Or Mid(aDateStr, 1, 2) = "12" Then
                        If Val(Mid(aDateStr, 4)) >= 0 And Val(Mid(aDateStr, 4)) <= 20 Then
                            'mm/01/20yy
                            ThisDate = Mid(aDateStr, 1, 3) & "01/20" & Mid(aDateStr, 4, 3)
                        Else
                            If Val(Mid(aDateStr, 4)) >= 21 And Val(Mid(aDateStr, 4)) <= 99 Then
                                'mm/01/19yy
                                ThisDate = Mid(aDateStr, 1, 3) & "01/19" & Mid(aDateStr, 4, 3)
                            Else
                                ThisDate = ""
                            End If
                        End If
                    Else
                        ThisDate = ""
                    End If
                Else
                    If Mid(aDateStr, 4, 2) <> "00" And Mid(aDateStr, 4, 2) <> "01" And _
                        Mid(aDateStr, 4, 2) <> "02" And Mid(aDateStr, 4, 2) <> "03" And _
                        Mid(aDateStr, 4, 2) <> "04" And Mid(aDateStr, 4, 2) <> "05" And _
                        Mid(aDateStr, 4, 2) <> "06" And Mid(aDateStr, 4, 2) <> "07" And _
                        Mid(aDateStr, 4, 2) <> "08" And Mid(aDateStr, 4, 2) <> "09" Then
                        TempString = Mid(aDateStr, 2, 2) + "/" + "0" + Mid(aDateStr, 1, 1) + _
                             "/" + "19" + Mid(aDateStr, 4, 2)
                    Else
                        TempString = Mid(aDateStr, 2, 2) + "/" + "0" + Mid(aDateStr, 1, 1) + _
                             "/" + "20" + Mid(aDateStr, 4, 2)
                    End If
                    ThisDate = TempString
                End If

            Case 6
                'ddmmyy
                If Mid(aDateStr, 5, 2) <> "00" And Mid(aDateStr, 5, 2) <> "01" And _
                    Mid(aDateStr, 5, 2) <> "02" And Mid(aDateStr, 5, 2) <> "03" And _
                    Mid(aDateStr, 5, 2) <> "04" And Mid(aDateStr, 5, 2) <> "05" And _
                    Mid(aDateStr, 5, 2) <> "06" And Mid(aDateStr, 5, 2) <> "07" And _
                    Mid(aDateStr, 5, 2) <> "08" And Mid(aDateStr, 5, 2) <> "09" Then
                    TempString = Mid(aDateStr, 3, 2) + "/" + Mid(aDateStr, 1, 2) + _
                                 "/" + "19" + Mid(aDateStr, 5, 2)
                Else
                    TempString = Mid(aDateStr, 3, 2) + "/" + Mid(aDateStr, 1, 2) + _
                                 "/" + "20" + Mid(aDateStr, 5, 2)
                End If

                ThisDate = TempString

            Case 10
                'dd/mm/yyyy
                ThisDate = aDateStr

            Case Else
                ThisDate = ""
        End Select

        If IsDate(ThisDate) Then
            gGetDrillDate = CDate(ThisDate)
        Else
            gGetDrillDate = #12/31/8888#
        End If
    End Function

    Public Sub gLoadProspCombos(ByVal aMineName As String, _
                                ByRef aSecCbo As ComboBox, _
                                ByRef aTwpCbo As ComboBox, _
                                ByRef aRgeCbo As ComboBox, _
                                ByRef aHlocCbo As ComboBox, _
                                ByVal AddtlItem As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo LoadProspCombosError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ComboBoxChoiceDynaset As OraDynaset
        Dim RowIdx As Integer

        'Make sure the combo boxes that are 'passed in' are empty
        For RowIdx = 0 To aSecCbo.Items.Count - 1
            aSecCbo.Items.RemoveAt(0)
        Next RowIdx
        For RowIdx = 0 To aTwpCbo.Items.Count - 1
            aTwpCbo.Items.RemoveAt(0)
        Next RowIdx
        For RowIdx = 0 To aRgeCbo.Items.Count - 1
            aRgeCbo.Items.RemoveAt(0)
        Next RowIdx
        For RowIdx = 0 To aHlocCbo.Items.Count - 1
            aHlocCbo.Items.RemoveAt(0)
        Next RowIdx

        '----------------------------------------------------------------
        'Sections  Sections  Sections  Sections  Sections  Sections
        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pCategoryName", "Section number", ORAPARM_INPUT)
        params("pCategoryName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_all_choices_num
        'pMineName               IN     VARCHAR2,
        'pCategoryName           IN     VARCHAR2,
        'pResult                 IN OUT c_choices
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_choices_num(:pMineName, " + _
                      ":pCategoryName, :pResult);end;", ORASQL_FAILEXEC)
        ComboBoxChoiceDynaset = params("pResult").Value
        ClearParams(params)

        If AddtlItem <> "" Then
            If AddtlItem = "()" Then
                aSecCbo.Items.Add("(Select sec...)")
            Else
                aSecCbo.Items.Add(AddtlItem)
            End If
        End If

        ComboBoxChoiceDynaset.MoveFirst()
        Do While Not ComboBoxChoiceDynaset.EOF
            aSecCbo.Items.Add(ComboBoxChoiceDynaset.Fields("combo_box_choice_text").Value)
            ComboBoxChoiceDynaset.MoveNext()
        Loop
        aSecCbo.Text = aSecCbo.Items(0)

        '----------------------------------------------------------------
        'Townships  Townships  Townships  Tonwships  Townships  Townships
        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pCategoryName", "Township", ORAPARM_INPUT)
        params("pCategoryName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_all_choices_num
        'pMineName               IN     VARCHAR2,
        'pCategoryName           IN     VARCHAR2,
        'pResult                 IN OUT c_choices
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_choices_num(:pMineName, " + _
                      ":pCategoryName, :pResult);end;", ORASQL_FAILEXEC)
        ComboBoxChoiceDynaset = params("pResult").Value
        ClearParams(params)

        If AddtlItem <> "" Then
            If AddtlItem = "()" Then
                aTwpCbo.Items.Add("(Select twp...)")
            Else
                aTwpCbo.Items.Add(AddtlItem)
            End If
        End If

        ComboBoxChoiceDynaset.MoveFirst()
        Do While Not ComboBoxChoiceDynaset.EOF
            aTwpCbo.Items.Add(ComboBoxChoiceDynaset.Fields("combo_box_choice_text").Value)
            ComboBoxChoiceDynaset.MoveNext()
        Loop
        aTwpCbo.Text = aTwpCbo.Items(0)

        '----------------------------------------------------------------
        'Ranges  Ranges  Ranges  Ranges  Ranges  Ranges  Ranges  Ranges
        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pCategoryName", "Range", ORAPARM_INPUT)
        params("pCategoryName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_all_choices_num
        'pMineName               IN     VARCHAR2,
        'pCategoryName           IN     VARCHAR2,
        'pResult                 IN OUT c_choices
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_choices_num(:pMineName, " + _
                      ":pCategoryName, :pResult);end;", ORASQL_FAILEXEC)
        ComboBoxChoiceDynaset = params("pResult").Value
        ClearParams(params)

        If AddtlItem <> "" Then
            If AddtlItem = "()" Then
                aRgeCbo.Items.Add("(Select rge...)")
            Else
                aRgeCbo.Items.Add(AddtlItem)
            End If
        End If

        ComboBoxChoiceDynaset.MoveFirst()
        Do While Not ComboBoxChoiceDynaset.EOF
            aRgeCbo.Items.Add(ComboBoxChoiceDynaset.Fields("combo_box_choice_text").Value)
            ComboBoxChoiceDynaset.MoveNext()
        Loop
        aRgeCbo.Text = aRgeCbo.Items(0)

        '----------------------------------------------------------------
        'Hole locations  Hole locations  Hole locations  Hole locations
        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pCategoryName", "Hole locations", ORAPARM_INPUT)
        params("pCategoryName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_all_choices
        'pMineName               IN     VARCHAR2,
        'pCategoryName           IN     VARCHAR2,
        'pResult                 IN OUT c_choices
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_choices(:pMineName, " + _
                      ":pCategoryName, :pResult);end;", ORASQL_FAILEXEC)
        ComboBoxChoiceDynaset = params("pResult").Value
        ClearParams(params)

        If AddtlItem <> "" Then
            If AddtlItem = "()" Then
                aHlocCbo.Items.Add("(Select loctn...)")
            Else
                aHlocCbo.Items.Add(AddtlItem)
            End If
        End If

        ComboBoxChoiceDynaset.MoveFirst()
        Do While Not ComboBoxChoiceDynaset.EOF
            aHlocCbo.Items.Add(ComboBoxChoiceDynaset.Fields("combo_box_choice_text").Value)
            ComboBoxChoiceDynaset.MoveNext()
        Loop
        aHlocCbo.Text = aHlocCbo.Items(0)

        ComboBoxChoiceDynaset.Close()
        Exit Sub

LoadProspCombosError:
        MsgBox("Error loading prospect combos." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Prospect Combo Loading Error")

        On Error Resume Next
        ClearParams(params)
        ComboBoxChoiceDynaset.Close()
    End Sub

    Public Sub gParseHoleLoc(ByVal aHoleDesc As String, _
                             ByRef aSec As Integer, _
                             ByRef aTwp As Integer, _
                             ByRef aRge As Integer, _
                             ByRef aHoleLoc As String)

        '**********************************************************************
        '  Parse sec-twp-rge hloc into its parts (ex. 13-32-25 A16).
        '
        '
        '**********************************************************************

        Dim DashLoc As Integer
        Dim BlankLoc As Integer
        Dim CurrStr As String

        DashLoc = InStr(aHoleDesc, "-")
        aSec = Mid(aHoleDesc, 1, DashLoc - 1)
        CurrStr = Mid(aHoleDesc, DashLoc + 1)

        DashLoc = InStr(CurrStr, "-")
        aTwp = Mid(CurrStr, 1, DashLoc - 1)
        CurrStr = Mid(CurrStr, DashLoc + 1)

        BlankLoc = InStr(CurrStr, " ")
        aRge = Mid(CurrStr, 1, BlankLoc - 1)
        CurrStr = Mid(CurrStr, BlankLoc + 1)

        aHoleLoc = CurrStr
    End Sub

    Public Function gGetHoleLoc2(ByVal aHoleLoc As String, _
                                 ByVal aMode As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'aMode = Char -- will change from 0234 to A01, etc.
        '        Num  -- will change from A01 to 0234, etc.

        Dim AlphaStr As String
        Dim TempStr As String
        Dim Num1 As Integer
        Dim Num2 As Integer
        Dim Char1 As String

        AlphaStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        aHoleLoc = Trim(StrConv(aHoleLoc, vbUpperCase))
        aMode = Trim(StrConv(aMode, vbUpperCase))
        gGetHoleLoc2 = ""

        'May have passed 234 instead of 0234 -- correct that here
        If IsNumeric(aHoleLoc) And Len(aHoleLoc) = 3 Then
            aHoleLoc = "0" & aHoleLoc
        End If

        'Change to numeric format  Change to numeric format
        'Change to numeric format  Change to numeric format
        'Change to numeric format  Change to numeric format
        If Len(aHoleLoc) = 3 Then
            If aMode <> "NUM" Then
                'Error -- should be A03, etc
                gGetHoleLoc2 = ""
                Exit Function
            End If

            Char1 = Mid(aHoleLoc, 1, 1)

            If InStr(AlphaStr, Char1) = 0 Then
                'Error -- should be A03, etc
                gGetHoleLoc2 = ""
                Exit Function
            End If

            TempStr = Mid(aHoleLoc, 2)

            If IsNumeric(TempStr) = True Then
                Num1 = Val(TempStr)
                If Num1 >= 1 And Num1 <= 16 Then
                    'Valid A01 type hole location has been entered
                Else
                    'Error -- should be A03, etc
                    gGetHoleLoc2 = ""
                    Exit Function
                End If
            Else
                'Error -- should be A03, etc
                gGetHoleLoc2 = ""
                Exit Function
            End If

            'A valid A01 type hole location has been entered!
            'Need to convert it to a 0234 type hole location.
            Num1 = InStr(AlphaStr, Char1) * 2
            Num2 = 34 + Val(Mid(aHoleLoc, 2)) * 2 - 2

            gGetHoleLoc2 = Format(Num1, "00") & Format(Num2, "00")
            Exit Function
        End If

        'Change to character format  Change to character format
        'Change to character format  Change to character format
        'Change to character format  Change to character format
        If Len(aHoleLoc) = 4 Then
            If aMode <> "CHAR" Then
                'Error -- should be 0234, etc
                gGetHoleLoc2 = ""
                Exit Function
            End If

            TempStr = Mid(aHoleLoc, 1, 2)

            If IsNumeric(TempStr) = True Then
                Num1 = Val(TempStr)
                If Num1 >= 2 And Num1 <= 32 And gIsEvenNumber(Num1) = True Then
                    'Valid 1st half of 0234 type hole location has been entered
                Else
                    'Error -- should be 0234, etc
                    gGetHoleLoc2 = "???"
                    Exit Function
                End If
            Else
                'Error -- should be 0234, etc
                gGetHoleLoc2 = "???"
                Exit Function
            End If

            TempStr = Mid(aHoleLoc, 3)

            If IsNumeric(TempStr) = True Then
                Num2 = Val(TempStr)
                If Num2 >= 34 And Num2 <= 64 And gIsEvenNumber(Num2) = True Then
                    'Valid 2nd half of 0234 type hole location has been entered
                Else
                    'Error -- should be 0234, etc
                    gGetHoleLoc2 = "???"
                    Exit Function
                End If
            Else
                'Error -- should be 0234, etc
                gGetHoleLoc2 = "???"
                Exit Function
            End If

            'A valid 0234 type hole location has been entered!
            'Need to convert it to a A01 type hole location.
            Char1 = Mid(AlphaStr, Num1 / 2, 1)

            Num1 = InStr(AlphaStr, Char1) * 2
            Num2 = (Num2 - 32) / 2

            gGetHoleLoc2 = Char1 & Format(Num2, "00")
            Exit Function
        End If
    End Function

    Public Function gGetForty(ByVal aHoleLoc As String, _
                              ByVal aMode As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'aMode = Char -- A01, B12, P08, etc.
        '        Num  -- 0234, 2850, 1262, etc.

        Dim Num1 As Integer
        Dim Num2 As Integer
        Dim Char1 As String

        gGetForty = 0

        If aMode = "CHAR" Then
            Char1 = Mid(aHoleLoc, 1, 1)
            Num1 = Val(Mid(aHoleLoc, 2))
            Select Case Num1
                Case Char1 >= "A" And Char1 <= "D"
                    Select Case Num1
                        Case 1 To 4
                            gGetForty = 12
                        Case 5 To 8
                            gGetForty = 14
                        Case 9 To 12
                            gGetForty = 16
                        Case 13 To 16
                            gGetForty = 18
                    End Select

                Case Char1 >= "E" And Char1 <= "H"
                    Select Case Num1
                        Case 1 To 4
                            gGetForty = 32
                        Case 5 To 8
                            gGetForty = 34
                        Case 9 To 12
                            gGetForty = 36
                        Case 13 To 16
                            gGetForty = 38
                    End Select

                Case Char1 >= "I" And Char1 <= "L"
                    Select Case Num1
                        Case 1 To 4
                            gGetForty = 52
                        Case 5 To 8
                            gGetForty = 54
                        Case 9 To 12
                            gGetForty = 56
                        Case 13 To 16
                            gGetForty = 58
                    End Select

                Case Char1 >= "M" And Char1 <= "P"
                    Select Case Num1
                        Case 1 To 4
                            gGetForty = 62
                        Case 5 To 8
                            gGetForty = 64
                        Case 9 To 12
                            gGetForty = 66
                        Case 13 To 16
                            gGetForty = 68
                    End Select

                Case Else
                    gGetForty = 0
            End Select
        End If

        If aMode = "NUM" Then
            Num1 = Val(Mid(aHoleLoc, 1, 2))
            Num2 = Val(Mid(aHoleLoc, 3))

            Select Case Num1
                Case 1 To 8
                    Select Case Num2
                        Case 33 To 40
                            gGetForty = 12
                        Case 41 To 48
                            gGetForty = 14
                        Case 49 To 56
                            gGetForty = 16
                        Case 57 To 72
                            gGetForty = 18
                    End Select

                Case 9 To 16
                    Select Case Num2
                        Case 33 To 40
                            gGetForty = 32
                        Case 41 To 48
                            gGetForty = 34
                        Case 49 To 56
                            gGetForty = 36
                        Case 57 To 64
                            gGetForty = 38
                    End Select

                Case 17 To 24
                    Select Case Num2
                        Case 33 To 40
                            gGetForty = 52
                        Case 41 To 48
                            gGetForty = 54
                        Case 49 To 56
                            gGetForty = 56
                        Case 57 To 64
                            gGetForty = 58
                    End Select

                Case 25 To 33
                    Select Case Num2
                        Case 33 To 40
                            gGetForty = 72
                        Case 41 To 48
                            gGetForty = 74
                        Case 49 To 56
                            gGetForty = 76
                        Case 57 To 64
                            gGetForty = 78
                    End Select

                Case Else
                    gGetForty = 0
            End Select
        End If
    End Function

    Public Function gHoleLocFitsAlpha(ByVal aHoleLocNumeric As Integer) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Some numeric hole locations will translate to alpha-numeric hole
        'locations (1636 = H02, 0258 = A13, etc.)
        'Some numeric hole locations will NOT translate to alpha-numeric
        'hole locations (0340 = ???, 1551 = ???, 1635 = ???, etc.)
        'If either of the two numbers in a numeric hole location is odd
        'then it will NOT translate to an alpha-numeric hole location.

        Dim FirstNumber As Integer
        Dim SecondNumber As Integer

        gHoleLocFitsAlpha = False

        FirstNumber = Val(Mid(Format(aHoleLocNumeric, "0000"), 1, 2))
        SecondNumber = Val(Mid(Format(aHoleLocNumeric, "0000"), 3))

        If gIsEvenNumber(FirstNumber) And gIsEvenNumber(SecondNumber) Then
            gHoleLocFitsAlpha = True
        Else
            gHoleLocFitsAlpha = False
        End If
    End Function

    Public Function gGetNumSplitsInHole(ByVal aMineName As String, _
                                        ByVal aSec As Integer, _
                                        ByVal aTwp As Integer, _
                                        ByVal aRge As Integer, _
                                        ByVal aHole As String, _
                                        ByVal aProspStandard As String, _
                                        ByVal aDisplayError As Boolean) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetNumSplitsInHoleError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        gGetNumSplitsInHole = 0

        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownShip", aTwp, ORAPARM_INPUT)
        params("pTownShip").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHole, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pProspStandard", aProspStandard, ORAPARM_INPUT)
        params("pProspStandard").serverType = ORATYPE_VARCHAR2

        params.Add("pSplitTotalNum", 0, ORAPARM_OUTPUT)
        params("pSplitTotalNum").serverType = ORATYPE_NUMBER

        'PROCEDURE get_hole_split_total
        'pMineName      IN     VARCHAR2,
        'pSection       IN     NUMBER,
        'pTownship      IN     NUMBER,
        'pRange         IN     NUMBER,
        'pHoleLocation  IN     VARCHAR2,
        'pProspStandard IN     VARCHAR2,
        'pSplitTotalNum IN OUT NUMBER)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect2.get_hole_split_total(:pMineName, " + _
                      ":pSection, :pTownship, :pRange, :pHoleLocation, " + _
                      ":pProspStandard, :pSplitTotalNum);end;", ORASQL_FAILEXEC)
        gGetNumSplitsInHole = params("pSplitTotalNum").Value
        ClearParams(params)

        Exit Function

gGetNumSplitsInHoleError:
        If aDisplayError = True Then
            MsgBox("Error getting number of splits in hole." & vbCrLf & _
                   Err.Description, _
                   vbOKOnly + vbExclamation, _
                   "Split Count Get Error")
        End If
        On Error Resume Next
        ClearParams(params)
        gGetNumSplitsInHole = 0
    End Function

    Public Function gGetAllHoleLocList(ByVal aMineName As String, _
                                       ByVal aIncludeBlank As Boolean) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetAllHoleLocListError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ChoiceDynaset As OraDynaset
        Dim HoleStr As String
        Dim ChoiceCount As Integer

        If gProspGridType <> "Alpha-numeric" And _
            gProspGridType <> "Numeric" Then
            gGetAllHoleLocList = ""
            Exit Function
        End If

        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pCategoryName", "Hole locations", ORAPARM_INPUT)
        params("pCategoryName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        If gProspGridType = "Alpha-numeric" Then
            'PROCEDURE get_all_choices
            'pMineName           IN     VARCHAR2,
            'pCategoryName       IN     VARCHAR2,
            'pResult             IN OUT c_choices)
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_choices(:pMineName, " + _
                          ":pCategoryName, :pResult);end;", ORASQL_FAILEXEC)
            ChoiceDynaset = params("pResult").Value
        End If

        If gProspGridType = "Numeric" Then
            'PROCEDURE get_all_choices_num
            'pMineName           IN     VARCHAR2,
            'pCategoryName       IN     VARCHAR2,
            'pResult             IN OUT c_choices)
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_choices_num(:pMineName, " + _
                          ":pCategoryName, :pResult);end;", ORASQL_FAILEXEC)
            ChoiceDynaset = params("pResult").Value
        End If

        'Create spreadsheet combo-box string from the dynaset
        HoleStr = ""
        ChoiceCount = 0

        ChoiceDynaset.MoveFirst()
        Do While Not ChoiceDynaset.EOF
            ChoiceCount = ChoiceCount + 1

            If aIncludeBlank = True Then
                If ChoiceCount = 1 Then
                    HoleStr = " " + Chr(9) + _
                              ChoiceDynaset.Fields("combo_box_choice_text").Value
                Else
                    HoleStr = HoleStr + Chr(9) + _
                              ChoiceDynaset.Fields("combo_box_choice_text").Value
                End If
            Else
                If ChoiceCount = 1 Then
                    HoleStr = ChoiceDynaset.Fields("combo_box_choice_text").Value
                Else
                    HoleStr = HoleStr + Chr(9) + _
                              ChoiceDynaset.Fields("combo_box_choice_text").Value
                End If
            End If

            ChoiceDynaset.MoveNext()
        Loop
        gGetAllHoleLocList = HoleStr

        ClearParams(params)
        ChoiceDynaset.Close()

        Exit Function

gGetAllHoleLocListError:
        MsgBox("Error getting hole locations." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Hole Locations Error")

        On Error Resume Next
        ClearParams(params)
        gGetAllHoleLocList = ""
        ChoiceDynaset.Close()
    End Function

    Public Function gGetClosestHole(ByVal aMineName As String, _
                                    ByVal aProspStandard As String, _
                                    ByVal aDirection As String, _
                                    ByVal aXcoord As Double, _
                                    ByVal aYcoord As Double, _
                                    ByVal aDiffMax As Single) As gCompBaseType

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetClosestHoleError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim HoleDynaset As OraDynaset
        Dim RecordCount As Long
        Dim ClosestDistance As Single
        Dim ThisXcoord As Double
        Dim ThisYcoord As Double
        Dim ThisDistance As Single
        Dim ClosestHole As gCompBaseType
        Dim ThisHole As String

        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pProspStandard", aProspStandard, ORAPARM_INPUT)
        params("pProspStandard").serverType = ORATYPE_VARCHAR2

        params.Add("pDirection", aDirection, ORAPARM_INPUT)
        params("pDirection").serverType = ORATYPE_VARCHAR2

        params.Add("pXcoord", aXcoord, ORAPARM_INPUT)
        params("pXcoord").serverType = ORATYPE_NUMBER

        params.Add("pYcoord", aYcoord, ORAPARM_INPUT)
        params("pYcoord").serverType = ORATYPE_NUMBER

        params.Add("pDiffMax", aDiffMax, ORAPARM_INPUT)
        params("pDiffMax").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'Procedure get_holes_closest
        'pMineName      IN     VARCHAR2,
        'pProspStandard IN     VARCHAR2,
        'pDirection     IN     VARCHAR2,
        'pXcoord        IN     NUMBER,
        'pYcoord        IN     NUMBER,
        'pDiffMax       IN     NUMBER,
        'pResult        IN OUT c_composite)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect2.get_holes_closest(:pMineName, " + _
                      ":pProspStandard, :pDirection, :pXcoord, :pYcoord, :pDiffMax, :pResult);end;", ORASQL_FAILEXEC)
        HoleDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = HoleDynaset.RecordCount
        ClosestDistance = 999999
        With ClosestHole
            .MineName = ""
            .HoleLoc = ""
            .Section = 0
            .Township = 0
            .Range = 0
            .Xcoord = 0
            .Ycoord = 0
        End With

        If RecordCount <> 0 Then
            'Search through the holes that have been returned and get the closest
            HoleDynaset.MoveFirst()
            Do While Not HoleDynaset.EOF
                With ClosestHole
                    ThisXcoord = HoleDynaset.Fields("x_sp_cdnt").Value
                    ThisYcoord = HoleDynaset.Fields("y_sp_cdnt").Value
                    ThisHole = HoleDynaset.Fields("hole_location").Value

                    ThisDistance = gGetDistBtwTwoPnts(ThisXcoord, ThisYcoord, _
                                                      aXcoord, aYcoord)

                    If HoleIsCloser(ThisDistance, _
                                    ClosestDistance, _
                                    ThisXcoord, _
                                    ThisYcoord, _
                                    aXcoord, _
                                    aYcoord, _
                                    aDirection, _
                                    ThisHole) = True Then

                        ClosestDistance = ThisDistance
                        .Xcoord = HoleDynaset.Fields("x_sp_cdnt").Value
                        .Ycoord = HoleDynaset.Fields("y_sp_cdnt").Value
                        .HoleLoc = HoleDynaset.Fields("hole_location").Value
                        .Section = HoleDynaset.Fields("section").Value
                        .Township = HoleDynaset.Fields("township").Value
                        .Range = HoleDynaset.Fields("range").Value
                        .MineName = HoleDynaset.Fields("mine_name").Value
                    End If
                End With
                HoleDynaset.MoveNext()
            Loop
        End If

        gGetClosestHole = ClosestHole

        HoleDynaset.Close()
        Exit Function

gGetClosestHoleError:
        MsgBox("Error getting closest hole." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Closest Hole Error")

        On Error Resume Next
        ClearParams(params)
        HoleDynaset.Close()
    End Function

    Public Function gGetDistBtwTwoPnts(ByVal aX1 As Double, _
                                       ByVal aY1 As Double, _
                                       ByVal aX2 As Double, _
                                       ByVal aY2 As Double) As Double

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Distance = SQRT((x2 - x1)^2 + (y2 - y1)^2)

        gGetDistBtwTwoPnts = Sqrt((aX2 - aX1) * (aX2 - aX1) + _
                               (aY2 - aY1) * (aY2 - aY1))
    End Function

    Public Function HoleIsCloser(ByVal aThisDistance As Single,
                                 ByVal aClosestDistance As Single,
                                 ByVal aXcoord2 As Double,
                                 ByVal aYcoord2 As Double,
                                 ByVal aXcoord1 As Double,
                                 ByVal aYcoord1 As Double,
                                 ByVal aDirection As String,
                                 ByVal aHole As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim Xdiff As Double
        Dim Ydiff As Double
        Dim IsHoleCloser As Boolean = False
        If aThisDistance >= aClosestDistance Then
            Return IsHoleCloser
        End If

        Xdiff = aXcoord2 - aXcoord1
        Ydiff = aYcoord2 - aYcoord1

        Select Case aDirection
            Case Is = "North"
                'Hole must be at least 250' North and no more than 400'
                'East or West of hole
                If Ydiff > 250 And Abs(Xdiff) < 400 Then
                    IsHoleCloser = True
                Else
                    HoleIsCloser = False
                End If

            Case Is = "South"
                'Hole must be at least 250' South and no more than 400'
                'East or West of hole
                If Ydiff < -250 And Abs(Xdiff) < 400 Then
                    IsHoleCloser = True
                Else
                    IsHoleCloser = False
                End If

            Case Is = "East"
                'Hole must be at least 250' East and no more than 400'
                'North or South of hole
                If Xdiff > 250 And Abs(Ydiff) < 400 Then
                    IsHoleCloser = True
                Else
                    IsHoleCloser = False
                End If

            Case Is = "West"
                'Hole must be at least 250' West and no more than 400'
                'North or South of hole
                If Xdiff < -250 And Abs(Ydiff) < 400 Then
                    IsHoleCloser = True
                Else
                    IsHoleCloser = False
                End If
        End Select
        Return IsHoleCloser
    End Function

    Public Function gGetMer(ByVal aTpBpl As Single,
                            ByVal aTpFe As Single,
                            ByVal aTpAl As Single,
                            ByVal aTpMg As Single,
                            ByVal aRoundVal As Integer) As Single

        Dim P2O5 As Double
        Dim result As Double = 0
        P2O5 = Round(aTpBpl / 2.185, 1)
        If P2O5 <> 0 Then
            result = Round((aTpFe + aTpAl + aTpMg) / P2O5 * 100, aRoundVal)
        End If
        Return result
    End Function


    Public Function gGetMerAt(ByVal aTpBpl As Single, _
                              ByVal aTpFe As Single, _
                              ByVal aTpAl As Single, _
                              ByVal aTpMg As Single, _
                              ByVal aRoundVal As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'gGetMerAt  At = Allen Truesdell  (Don't multiply by 100!)

        Dim P2O5 As Double

        P2O5 = Round(aTpBpl / 2.185, 1)
        If P2O5 <> 0 Then
            gGetMerAt = Round((aTpFe + aTpAl + aTpMg) / P2O5, aRoundVal)
        Else
            gGetMerAt = 0
        End If
    End Function

    Public Function gNumHoleLocIsValid(ByVal aHoleLoc As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'aMode = Char -- will change from 0234 to A01, etc.
        '        Num  -- will change from A01 to 0234, etc.

        Dim TempStr As String
        Dim Num1 As Integer
        Dim Num2 As Integer

        gNumHoleLocIsValid = True

        If Not IsNumeric(aHoleLoc) Or _
            (Len(Trim(aHoleLoc)) <> 3 And _
            Len(Trim(aHoleLoc)) <> 4) Then
            gNumHoleLocIsValid = False
            Exit Function
        End If

        'May have passed 234 instead of 0234 -- correct that here
        If IsNumeric(aHoleLoc) And Len(aHoleLoc) = 3 Then
            aHoleLoc = "0" & aHoleLoc
        End If

        TempStr = Mid(aHoleLoc, 1, 2)

        If IsNumeric(TempStr) = True Then
            Num1 = Val(TempStr)
            If Num1 >= 1 And Num1 <= 33 Then
                'Valid 1st half of 0234 type hole location has been entered
            Else
                'Error -- should be 0234, etc
                gNumHoleLocIsValid = False
                Exit Function
            End If
        Else
            'Error -- should be 0234, etc
            gNumHoleLocIsValid = False
            Exit Function
        End If

        TempStr = Mid(aHoleLoc, 3)

        If IsNumeric(TempStr) = True Then
            Num2 = Val(TempStr)
            If Num2 >= 33 And Num2 <= 72 Then
                'Valid 2nd half of 0234 type hole location has been entered
            Else
                'Error -- should be 0234, etc
                gNumHoleLocIsValid = False
                Exit Function
            End If
        Else
            'Error -- should be 0234, etc
            gNumHoleLocIsValid = False
            Exit Function
        End If
    End Function

    Public Sub gGetHoleCoordElev(ByVal aMineName As String, _
                                 ByVal aTownship As Single, _
                                 ByVal aRange As Single, _
                                 ByVal aSection As Single, _
                                 ByVal aHoleLocation As String, _
                                 ByVal aProspStandard As String, _
                                 ByRef aXcoord As Double, _
                                 ByRef aYcoord As Double, _
                                 ByRef aHoleElevation As Single, _
                                 ByRef aDrillCdate As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetHoleCoordElevError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim HoleDynaset As OraDynaset
        Dim RecordCount As Long

        aXcoord = 0
        aYcoord = 0
        aHoleElevation = 0
        aDrillCdate = ""

        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pHoleLocation", aHoleLocation, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pSection", aSection, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTownship, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRange, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pProspStandard", "100%PROSPECT", ORAPARM_INPUT)
        params("pProspStandard").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_hole_coords_elev
        'pMineName      IN     VARCHAR2,
        'pHoleLocation  IN     VARCHAR2,
        'pSection       IN     NUMBER,
        'pTownship      IN     NUMBER,
        'pRange         IN     NUMBER,
        'pProspStandard IN     VARCHAR2,
        'pResult        IN OUT c_composite)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect2.get_hole_coords_elev(:pMineName, " + _
                      ":pHoleLocation, :pSection, :pTownship, :pRange, " + _
                      ":pProspStandard, :pResult);end;", ORASQL_FAILEXEC)
        HoleDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = HoleDynaset.RecordCount

        'Will assume that at most on record will be returned.
        If RecordCount = 1 Then
            HoleDynaset.MoveFirst()
            aXcoord = HoleDynaset.Fields("x_sp_cdnt").Value
            aYcoord = HoleDynaset.Fields("y_sp_cdnt").Value
            aHoleElevation = HoleDynaset.Fields("hole_elevation").Value
            aDrillCdate = HoleDynaset.Fields("drill_cdate").Value
        End If

        HoleDynaset.Close()
        Exit Sub

gGetHoleCoordElevError:
        MsgBox("Error getting hole coordinates and elevation." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Process Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        HoleDynaset.Close()
    End Sub

    Public Function gGetDepth1stSplitRaw(ByVal aMineName As String, _
                                         ByVal aTownship As Single, _
                                         ByVal aRange As Single, _
                                         ByVal aSection As Single, _
                                         ByVal aHoleLocation As String) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetDepth1stSplitRawError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim SplitDynaset As OraDynaset
        Dim RecordCount As Long

        gGetDepth1stSplitRaw = 0

        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pHoleLocation", aHoleLocation, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pSection", aSection, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTownship, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRange, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_hole_raw_depth_1st
        'pMineName      IN     VARCHAR2,
        'pHoleLocation  IN     VARCHAR2,
        'pSection       IN     NUMBER,
        'pTownship      IN     NUMBER,
        'pRange         IN     NUMBER,
        'pResult        IN OUT c_split)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect2.get_hole_raw_depth_1st(:pMineName, " + _
                      ":pHoleLocation, :pSection, :pTownship, :pRange, " + _
                      ":pResult);end;", ORASQL_FAILEXEC)
        SplitDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = SplitDynaset.RecordCount

        'Will assume that at most on record will be returned.
        If RecordCount = 1 Then
            SplitDynaset.MoveFirst()
            gGetDepth1stSplitRaw = SplitDynaset.Fields("split_depth_top").Value
        Else
            gGetDepth1stSplitRaw = 0
        End If

        SplitDynaset.Close()
        Exit Function

gGetDepth1stSplitRawError:
        MsgBox("Error getting depth 1st split." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Process Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        SplitDynaset.Close()
    End Function

    Public Function gGetDepthBotSplitRaw(ByVal aMineName As String, _
                                         ByVal aTownship As Single, _
                                         ByVal aRange As Single, _
                                         ByVal aSection As Single, _
                                         ByVal aHoleLocation As String, _
                                         ByVal aSplitTotalNum As Integer) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetDepthBotSplitRawError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim SplitDynaset As OraDynaset
        Dim RecordCount As Long

        gGetDepthBotSplitRaw = 0

        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pHoleLocation", aHoleLocation, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pSection", aSection, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTownship, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRange, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSplitTotalNum", aSplitTotalNum, ORAPARM_INPUT)
        params("pSplitTotalNum").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_hole_raw_depth_bottom
        'pMineName      IN     VARCHAR2,
        'pHoleLocation  IN     VARCHAR2,
        'pSection       IN     NUMBER,
        'pTownship      IN     NUMBER,
        'pRange         IN     NUMBER,
        'pSplitTotalNum IN     NUMBER,
        'pResult        IN OUT c_split)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect2.get_hole_raw_depth_bottom(:pMineName, " + _
                      ":pHoleLocation, :pSection, :pTownship, :pRange, " + _
                      ":pSplitTotalNum, :pResult);end;", ORASQL_FAILEXEC)
        SplitDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = SplitDynaset.RecordCount

        'Will assume that at most on record will be returned.
        If RecordCount = 1 Then
            SplitDynaset.MoveFirst()
            gGetDepthBotSplitRaw = SplitDynaset.Fields("split_depth_bot").Value
        Else
            gGetDepthBotSplitRaw = 0
        End If

        SplitDynaset.Close()
        Exit Function

gGetDepthBotSplitRawError:
        MsgBox("Error getting depth bottom split." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Process Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        SplitDynaset.Close()
    End Function

    Public Sub gGetHoleMoisExist(ByVal aTownship As Integer, _
                                 ByVal aRange As Integer, _
                                 ByVal aSection As Integer, _
                                 ByVal aHoleLocation As String, _
                                 ByRef aHoleDynaset As OraDynaset)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetHoleMoisExistError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim HoleCount As Integer
        Dim HoleLocAlpha As String
        Dim RecordCount As Integer

        'The hole location that is passed in here will be numeric!
        HoleLocAlpha = gGetHoleLoc2(aHoleLocation, "Char")
        If HoleLocAlpha = "???" Then
            HoleLocAlpha = ""
        End If

        params = gDBParams

        params.Add("pTownship", aTownship, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRange, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSection, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        'Numeric hole location.
        params.Add("pHoleLocation1", aHoleLocation, ORAPARM_INPUT)
        params("pHoleLocation1").serverType = ORATYPE_VARCHAR2

        'Alpha-numeric hole location if one is possible.
        params.Add("pHoleLocation2", HoleLocAlpha, ORAPARM_INPUT)
        params("pHoleLocation2").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_hole_prospect_comp_base2
        'pTownship       IN     NUMBER,
        'pRange          IN     NUMBER,
        'pSection        IN     NUMBER,
        'pHoleLocation1  IN     VARCHAR2,
        'pHoleLocation2  IN     VARCHAR2,
        'pResult         IN OUT c_composite)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect2.get_hole_prospect_comp_base2(" + _
                      ":pTownship, :pRange, :pSection, " + _
                      ":pHoleLocation1, :pHoleLocation2, :pResult);end;", ORASQL_FAILEXEC)
        aHoleDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = aHoleDynaset.RecordCount

        Exit Sub

gGetHoleMoisExistError:
        MsgBox("Error getting hole data." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Process Error")

        On Error Resume Next
        ClearParams(params)
    End Sub

    Public Function gGetAvgValue2(ByVal aValue1 As Single, _
                                  ByVal aValue2 As Single, _
                                  ByVal aRound As Integer) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Will not average in zero values!

        If aValue1 > 0 And aValue2 = 0 Then
            gGetAvgValue2 = Round(aValue1, aRound)
            Exit Function
        End If

        If aValue2 > 0 And aValue1 = 0 Then
            gGetAvgValue2 = Round(aValue2, aRound)
            Exit Function
        End If

        gGetAvgValue2 = Round((aValue1 + aValue2) / 2, aRound)
    End Function

    Public Sub gGetHoleDateAndNumSplits(ByVal aMine As String, _
                                        ByVal aSec As Integer, _
                                        ByVal aTwp As Integer, _
                                        ByVal aRge As Integer, _
                                        ByVal aHole As String, _
                                        ByVal aProspStandard As String, _
                                        ByRef aProspDate As Date, _
                                        ByRef aSplCnt As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetHoleDateAndNumSplitsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim DrillCdate As String
        Dim HoleDynaset As OraDynaset
        Dim ProspGridType As String
        Dim HoleLocAlpha As String

        'Make sure we have the correct hole location type for MOIS!
        ProspGridType = gGetProspGridType(aMine)

        'Add new composite prospect data
        If ProspGridType = "Alpha-numeric" Then
            'May need to get the alpha-numeric hole location.
            'If it won't translate will get "???".
            If IsNumeric(aHole) Then
                aHole = gGetHoleLoc2(aHole, "Char")

                If aHole = "???" Then
                    'Numeric holelocation will not translate to alpha-numeric!
                    aProspDate = #12/31/8888#
                    aSplCnt = 0
                    Exit Sub
                End If
            End If
        End If
        If ProspGridType = "Numeric" Then
            'May need to get the numeric hole location.
            'If it won't translate will get "???".
            If Not IsNumeric(aHole) Then
                aHole = gGetHoleLoc2(aHole, "Num")

                If aHole = "???" Then
                    'Numeric holelocation will not translate to alpha-numeric!
                    aProspDate = #12/31/8888#
                    aSplCnt = 0
                    Exit Sub
                End If
            End If
        End If

        params = gDBParams

        params.Add("pMineName", aMine, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pHoleLocation", aHole, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pProspStandard", aProspStandard, ORAPARM_INPUT)
        params("pProspStandard").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_hole_date_numsplits
        'pMineName        IN     VARCHAR2,
        'pHoleLocation    IN     VARCHAR2,
        'pTownship        IN     NUMBER,
        'pRange           IN     NUMBER,
        'pSection         IN     NUMBER,
        'pProspStandard   IN     VARCHAR2,
        'pResult          IN OUT c_composite)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect2.get_hole_date_numsplits(:pMineName, " + _
                      ":pHoleLocation, :pTownship, :pRange, :pSection, :pProspStandard, :pResult);end;", ORASQL_FAILEXEC)
        HoleDynaset = params("pResult").Value
        ClearParams(params)

        If HoleDynaset.RecordCount = 1 Then
            DrillCdate = HoleDynaset("drill_cdate").Value
            aSplCnt = HoleDynaset("split_total_num").Value

            'Let's try to get a real date for the character date!
            'DRILL_CDATE in PROSPECT_COMP_BASE may be some goofy date from GEOCOMP!
            aProspDate = gGetDrillDate(DrillCdate)
        Else
            aProspDate = #12/31/8888#
            aSplCnt = 0
        End If

        HoleDynaset.Close()

        Exit Sub

gGetHoleDateAndNumSplitsError:
        MsgBox("Error getting hole data." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Process Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        HoleDynaset.Close()
    End Sub

    Public Sub gGetTonsFromProspSplits(ByVal aMineName As String, _
                                       ByVal aProspStandard As String, _
                                       ByVal aTwp As Integer, _
                                       ByVal aRge As Integer, _
                                       ByVal aSec As Integer, _
                                       ByVal aHloc As String, _
                                       ByVal aMtxThkAct As Single, _
                                       ByVal aOvbThkAct As Single, _
                                       ByVal aItbThkAct As Single, _
                                       ByRef aPbTpaUm As Integer, _
                                       ByRef aCnTpaUm As Integer, _
                                       ByRef aFdTpaUm As Integer, _
                                       ByRef aMtxTpaUm As Long, _
                                       ByRef aWclTpaUm As Long, _
                                       ByRef aPctWclUm As Single, _
                                       ByRef aMtxDryDensUm As Single)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SplData() As gProspectSplit
        Dim SplCnt As Integer
        Dim MtxCuFt As Single
        Dim MtxAcres As Single
        Dim MtxDryDensity As Double
        Dim MtxTPA As Long

        'Want to get all of the splits for this hole and then using aMtxThkAct and
        'aOvbThkAct to determine some prospect data for the mtx thickness passed
        'in through aMtxThkAct.  Really only concerned with Mtx TPA, Clay TPA, %Clay,
        'and Mtx dry density (aMtxTpa, aWclTpa, aPctWcl, aMtxDryDens) at this time.

        aPbTpaUm = 0
        aCnTpaUm = 0
        aFdTpaUm = 0
        aMtxTpaUm = 0
        aWclTpaUm = 0
        aPctWclUm = 0
        aMtxDryDensUm = 0

        'Need to get the split data for this hole -- Wcl & Mtx TPA.
        SplCnt = GetSplitsWclMtxTpa(aMineName, _
                                    aSec, _
                                    aTwp, _
                                    aRge, _
                                    aHloc, _
                                    aProspStandard, _
                                    SplData)

        'Split data is now in SplData()
        GetSplTpasForThk(aMtxThkAct, _
                         aOvbThkAct, _
                         aItbThkAct, _
                         aPbTpaUm, _
                         aCnTpaUm, _
                         aFdTpaUm, _
                         aWclTpaUm, _
                         aMtxTpaUm, _
                         SplData)

        'Need to calculate Clay% and Mtx dry density.
        If aMtxTpaUm <> 0 Then
            aPctWclUm = Round(aWclTpaUm / aMtxTpaUm * 100, 2)
        Else
            aPctWclUm = 0
        End If

        'Determine the matrix dry density in lbs per cubic foot.
        MtxAcres = 1
        MtxCuFt = aMtxThkAct * MtxAcres * 43560
        MtxTPA = aMtxTpaUm

        If MtxCuFt <> 0 Then
            MtxDryDensity = Round((MtxTPA * 2000) / MtxCuFt, 2)
        Else
            MtxDryDensity = 0
        End If
        aMtxDryDensUm = MtxDryDensity
    End Sub

    Private Function GetSplitsWclMtxTpa(ByVal aMineName As String, _
                                        ByVal aSec As Integer, _
                                        ByVal aTwp As Integer, _
                                        ByVal aRge As Integer, _
                                        ByVal aHole As String, _
                                        ByVal aProspStandard As String, _
                                        ByRef aSplData() As gProspectSplit) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetSplitsWclMtxTpaError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim SplDynaset As OraDynaset
        Dim SplCnt As Integer

        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pHoleLocation", aHole, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pProspStandard", aProspStandard, ORAPARM_INPUT)
        params("pProspStandard").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_split_wcl_tpa
        'pMineName      IN     VARCHAR2,
        'pHoleLocation  IN     VARCHAR2,
        'pSection       IN     NUMBER,
        'pTownship      IN     NUMBER,
        'pRange         IN     NUMBER,
        'pProspStandard IN     VARCHAR2,
        'pResult        IN OUT c_split)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect2.get_split_wcl_tpa(:pMineName, " + _
                                       ":pHoleLocation, :pSection, :pTownship, " + _
                                       ":pRange, :pProspStandard, :pResult);end;", ORASQL_FAILEXEC)

        SplDynaset = params("pResult").Value
        ClearParams(params)

        SplDynaset.MoveFirst()
        SplCnt = 0
        GetSplitsWclMtxTpa = SplDynaset.RecordCount

        If SplDynaset.RecordCount <> 0 Then
            ReDim aSplData(SplDynaset.RecordCount)

            Do While Not SplDynaset.EOF
                SplCnt = SplCnt + 1

                With aSplData(SplCnt)
                    .Mine = SplDynaset("mine_name").Value
                    .ProspStandard = SplDynaset("prosp_standard").Value
                    .HoleLocation = SplDynaset("hole_location").Value
                    .Section = SplDynaset("section").Value
                    .Township = SplDynaset("township").Value
                    .Range = SplDynaset("range").Value
                    .Split = SplDynaset("split").Value
                    .SplitThickness = SplDynaset("split_thck").Value
                    .MinableStatus = SplDynaset("minable_status").Value
                    .TopOfSplitDepth = SplDynaset("split_depth_top").Value
                    .BotOfSplitDepth = SplDynaset("bot_split_depth").Value
                    '----------
                    .WcTpa = SplDynaset("wcl_tpa").Value
                    .MtxTPA = SplDynaset("mtx_tpa").Value
                End With
                SplDynaset.MoveNext()
            Loop
        End If

        SplDynaset.Close()

        Exit Function

GetSplitsWclMtxTpaError:
        MsgBox("Error getting split data." & vbCrLf & _
                Err.Description, _
                vbOKOnly + vbExclamation, _
                "Process Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        SplDynaset.Close()
        GetSplitsWclMtxTpa = 0
    End Function

    Private Sub GetSplTpasForThk(ByVal aMtxThkAct As Single, _
                                 ByVal aOvbThkAct As Single, _
                                 ByVal aItbThkAct As Single, _
                                 ByRef aPbTpaUm As Integer, _
                                 ByRef aCnTpaUm As Integer, _
                                 ByRef aFdTpaUm As Integer, _
                                 ByRef aWclTpaUm As Long, _
                                 ByRef aMtxTpaUm As Long, _
                                 ByRef aSplData() As gProspectSplit)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim SplIdx As Integer
        Dim MtxTopAct As Single
        Dim MtxBotAct As Single
        Dim BotSpl As Boolean
        Dim TopSpl As Boolean
        Dim ThisSplThk As Single
        Dim ThisSplTop As Single
        Dim ThisSplBot As Single
        Dim ExtraThk As Single
        Dim SplTopIncl As Single
        Dim SplBotIncl As Single
        Dim SumWclTpa As Single
        Dim SumMtxTpa As Single

        'This Sub is called from gGetTonsFromProspSplits.
        'This Sub processes prospect split data.

        'It is used to determine clay and matrix TPA's that would
        'have been realized if the "actual" matrix footage is applied to the prospect hole.
        'Only concerned with "actual" ovb and "actual" mtx footages right now (not interburdens).

        'This is useful when evaluating "unmineable" prospect holes that were actually mined.
        'At some point I may add Pb, Cn and Fd to this subroutine.  The Cn TPA is not "readily"
        'available in the stored split data in MOIS (it is basically a copy of what GEOCOMP did
        'and the Cn TPA needs to be calculated.

        aPbTpaUm = 0
        aCnTpaUm = 0
        aFdTpaUm = 0
        aWclTpaUm = 0
        aMtxTpaUm = 0

        'If the actual overburden is less than the prospect overburden then add the difference
        'to the 1st split in the hole.

        'If the actual mining depth is greater than the depth of the prospect hole then add the
        'additional footage to the bottom split.

        MtxTopAct = aOvbThkAct
        MtxBotAct = aMtxThkAct + aOvbThkAct

        For SplIdx = 1 To UBound(aSplData)
            If SplIdx = 1 Then
                TopSpl = True
            Else
                TopSpl = False
            End If

            If SplIdx = UBound(aSplData) Then
                BotSpl = True
            Else
                BotSpl = False
            End If

            ThisSplThk = aSplData(SplIdx).SplitThickness
            ThisSplTop = aSplData(SplIdx).TopOfSplitDepth
            ThisSplBot = aSplData(SplIdx).BotOfSplitDepth

            If BotSpl = True Then
                'If necessary increase the thickness of the 1st split if the
                'actual overburden is less than the prospect overburden.
                If aOvbThkAct < ThisSplTop Then
                    ThisSplThk = ThisSplThk + ThisSplTop - aOvbThkAct
                End If
            End If

            If BotSpl = True Then
                'If necessary increase the thickness of the bottom split to match
                'the depth of actual mining.
                If MtxBotAct > ThisSplBot Then
                    ThisSplThk = ThisSplThk + MtxBotAct - ThisSplBot
                End If
            End If

            'Is any of this split between MtxTopAct and MtxBotAct?
            If (ThisSplTop > MtxTopAct And ThisSplTop < MtxBotAct) Or _
                (ThisSplBot > MtxTopAct And ThisSplBot < MtxBotAct) Then
                'We have some split to process.
                If ThisSplTop < MtxTopAct Then
                    SplTopIncl = MtxTopAct
                Else
                    SplTopIncl = ThisSplTop
                End If

                If ThisSplBot > MtxBotAct Then
                    SplBotIncl = MtxBotAct
                Else
                    SplBotIncl = ThisSplBot
                End If

                ExtraThk = SplBotIncl - SplTopIncl

                If ThisSplThk <> 0 Then
                    SumWclTpa = SumWclTpa + Round((ExtraThk / ThisSplThk) * aSplData(SplIdx).WcTpa, 0)
                End If
                If ThisSplThk <> 0 Then
                    SumMtxTpa = SumMtxTpa + Round((ExtraThk / ThisSplThk) * aSplData(SplIdx).MtxTPA, 0)
                End If
            End If
        Next SplIdx

        aWclTpaUm = SumWclTpa
        aMtxTpaUm = SumMtxTpa
    End Sub

    Public Sub gGetNeighborSection(ByVal aDirection As String, _
                                   ByVal aTwp As Integer, _
                                   ByVal aRge As Integer, _
                                   ByVal aSec As Integer, _
                                   ByRef aAdjSec As gCompBaseType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        SectionMove(aDirection, aTwp, aRge, aSec)

        With aAdjSec
            .Township = aTwp
            .Range = aRge
            .Section = aSec
        End With
    End Sub

    Private Sub SectionMove(ByVal aDirection As String, _
                            ByRef aTwp As Integer, _
                            ByRef aRge As Integer, _
                            ByRef aSec As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ThisTwp As Integer
        Dim ThisRge As Integer
        Dim ThisSec As Integer

        ThisTwp = aTwp
        ThisRge = aRge
        ThisSec = aSec

        'aDirection
        '1) North
        '2) South
        '3) East
        '4) West
        '5) NE
        '6) NW
        '7) SW
        '8) SE

        Select Case aDirection
            Case Is = "North"
                Select Case ThisSec
                    Case Is = 1
                        aSec = 36
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 2
                        aSec = 35
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 3
                        aSec = 34
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 4
                        aSec = 33
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 5
                        aSec = 32
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 6
                        aSec = 31
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 7
                        aSec = 6
                    Case Is = 8
                        aSec = 5
                    Case Is = 9
                        aSec = 4
                    Case Is = 10
                        aSec = 3
                    Case Is = 11
                        aSec = 2
                    Case Is = 12
                        aSec = 1
                    Case Is = 13
                        aSec = 12
                    Case Is = 14
                        aSec = 11
                    Case Is = 15
                        aSec = 10
                    Case Is = 16
                        aSec = 9
                    Case Is = 17
                        aSec = 8
                    Case Is = 18
                        aSec = 7
                    Case Is = 19
                        aSec = 18
                    Case Is = 20
                        aSec = 17
                    Case Is = 21
                        aSec = 16
                    Case Is = 22
                        aSec = 15
                    Case Is = 23
                        aSec = 14
                    Case Is = 24
                        aSec = 13
                    Case Is = 25
                        aSec = 24
                    Case Is = 26
                        aSec = 23
                    Case Is = 27
                        aSec = 22
                    Case Is = 28
                        aSec = 21
                    Case Is = 29
                        aSec = 20
                    Case Is = 30
                        aSec = 19
                    Case Is = 31
                        aSec = 30
                    Case Is = 32
                        aSec = 29
                    Case Is = 33
                        aSec = 28
                    Case Is = 34
                        aSec = 27
                    Case Is = 35
                        aSec = 26
                    Case Is = 36
                        aSec = 25
                End Select

            Case Is = "South"
                Select Case aSec
                    Case Is = 1
                        aSec = 12
                    Case Is = 2
                        aSec = 11
                    Case Is = 3
                        aSec = 10
                    Case Is = 4
                        aSec = 9
                    Case Is = 5
                        aSec = 8
                    Case Is = 6
                        aSec = 7
                    Case Is = 7
                        aSec = 18
                    Case Is = 8
                        aSec = 17
                    Case Is = 9
                        aSec = 16
                    Case Is = 10
                        aSec = 15
                    Case Is = 11
                        aSec = 14
                    Case Is = 12
                        aSec = 13
                    Case Is = 13
                        aSec = 24
                    Case Is = 14
                        aSec = 23
                    Case Is = 15
                        aSec = 22
                    Case Is = 16
                        aSec = 21
                    Case Is = 17
                        aSec = 20
                    Case Is = 18
                        aSec = 19
                    Case Is = 19
                        aSec = 30
                    Case Is = 20
                        aSec = 29
                    Case Is = 21
                        aSec = 28
                    Case Is = 22
                        aSec = 27
                    Case Is = 23
                        aSec = 26
                    Case Is = 24
                        aSec = 25
                    Case Is = 25
                        aSec = 36
                    Case Is = 26
                        aSec = 35
                    Case Is = 27
                        aSec = 34
                    Case Is = 28
                        aSec = 33
                    Case Is = 29
                        aSec = 32
                    Case Is = 30
                        aSec = 31
                    Case Is = 31
                        aSec = 6
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 32
                        aSec = 5
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 33
                        aSec = 4
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 34
                        aSec = 3
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 35
                        aSec = 2
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 36
                        aSec = 1
                        aTwp = TownshipMove("South", ThisTwp)
                End Select

            Case Is = "West"
                Select Case aSec
                    Case Is = 1
                        aSec = 2
                    Case Is = 2
                        aSec = 3
                    Case Is = 3
                        aSec = 4
                    Case Is = 4
                        aSec = 5
                    Case Is = 5
                        aSec = 6
                    Case Is = 6
                        aSec = 1
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 7
                        aSec = 12
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 8
                        aSec = 7
                    Case Is = 9
                        aSec = 8
                    Case Is = 10
                        aSec = 9
                    Case Is = 11
                        aSec = 10
                    Case Is = 12
                        aSec = 11
                    Case Is = 13
                        aSec = 14
                    Case Is = 14
                        aSec = 15
                    Case Is = 15
                        aSec = 16
                    Case Is = 16
                        aSec = 17
                    Case Is = 17
                        aSec = 18
                    Case Is = 18
                        aSec = 13
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 19
                        aSec = 24
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 20
                        aSec = 19
                    Case Is = 21
                        aSec = 20
                    Case Is = 22
                        aSec = 21
                    Case Is = 23
                        aSec = 22
                    Case Is = 24
                        aSec = 23
                    Case Is = 25
                        aSec = 26
                    Case Is = 26
                        aSec = 27
                    Case Is = 27
                        aSec = 28
                    Case Is = 28
                        aSec = 29
                    Case Is = 29
                        aSec = 30
                    Case Is = 30
                        aSec = 25
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 31
                        aSec = 36
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 32
                        aSec = 31
                    Case Is = 33
                        aSec = 32
                    Case Is = 34
                        aSec = 33
                    Case Is = 35
                        aSec = 34
                    Case Is = 36
                        aSec = 35
                End Select

            Case Is = "East"
                Select Case aSec
                    Case Is = 1
                        aSec = 6
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 2
                        aSec = 1
                    Case Is = 3
                        aSec = 2
                    Case Is = 4
                        aSec = 3
                    Case Is = 5
                        aSec = 4
                    Case Is = 6
                        aSec = 5
                    Case Is = 7
                        aSec = 8
                    Case Is = 8
                        aSec = 9
                    Case Is = 9
                        aSec = 10
                    Case Is = 10
                        aSec = 11
                    Case Is = 11
                        aSec = 12
                    Case Is = 12
                        aSec = 7
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 13
                        aSec = 18
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 14
                        aSec = 13
                    Case Is = 15
                        aSec = 14
                    Case Is = 16
                        aSec = 15
                    Case Is = 17
                        aSec = 16
                    Case Is = 18
                        aSec = 17
                    Case Is = 19
                        aSec = 20
                    Case Is = 20
                        aSec = 21
                    Case Is = 21
                        aSec = 22
                    Case Is = 22
                        aSec = 23
                    Case Is = 23
                        aSec = 24
                    Case Is = 24
                        aSec = 19
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 25
                        aSec = 30
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 26
                        aSec = 25
                    Case Is = 27
                        aSec = 26
                    Case Is = 28
                        aSec = 27
                    Case Is = 29
                        aSec = 28
                    Case Is = 30
                        aSec = 29
                    Case Is = 31
                        aSec = 32
                    Case Is = 32
                        aSec = 33
                    Case Is = 33
                        aSec = 34
                    Case Is = 34
                        aSec = 35
                    Case Is = 35
                        aSec = 36
                    Case Is = 36
                        aSec = 31
                        aRge = RangeMove("East", ThisRge)
                End Select

            Case Is = "NE"
                Select Case ThisSec
                    Case Is = 1
                        aSec = 31
                        aTwp = TownshipMove("North", ThisTwp)
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 2
                        aSec = 36
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 3
                        aSec = 35
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 4
                        aSec = 34
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 5
                        aSec = 33
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 6
                        aSec = 32
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 7
                        aSec = 5
                    Case Is = 8
                        aSec = 4
                    Case Is = 9
                        aSec = 3
                    Case Is = 10
                        aSec = 2
                    Case Is = 11
                        aSec = 1
                    Case Is = 12
                        aSec = 6
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 13
                        aSec = 7
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 14
                        aSec = 12
                    Case Is = 15
                        aSec = 11
                    Case Is = 16
                        aSec = 10
                    Case Is = 17
                        aSec = 9
                    Case Is = 18
                        aSec = 8
                    Case Is = 19
                        aSec = 17
                    Case Is = 20
                        aSec = 16
                    Case Is = 21
                        aSec = 15
                    Case Is = 22
                        aSec = 14
                    Case Is = 23
                        aSec = 13
                    Case Is = 24
                        aSec = 18
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 25
                        aSec = 19
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 26
                        aSec = 24
                    Case Is = 27
                        aSec = 23
                    Case Is = 28
                        aSec = 22
                    Case Is = 29
                        aSec = 21
                    Case Is = 30
                        aSec = 20
                    Case Is = 31
                        aSec = 29
                    Case Is = 32
                        aSec = 28
                    Case Is = 33
                        aSec = 27
                    Case Is = 34
                        aSec = 26
                    Case Is = 35
                        aSec = 25
                    Case Is = 36
                        aSec = 30
                        aRge = RangeMove("East", ThisRge)
                End Select

            Case Is = "SE"
                Select Case ThisSec
                    Case Is = 1
                        aSec = 7
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 2
                        aSec = 12
                    Case Is = 3
                        aSec = 11
                    Case Is = 4
                        aSec = 10
                    Case Is = 5
                        aSec = 9
                    Case Is = 6
                        aSec = 8
                    Case Is = 7
                        aSec = 17
                    Case Is = 8
                        aSec = 16
                    Case Is = 9
                        aSec = 15
                    Case Is = 10
                        aSec = 14
                    Case Is = 11
                        aSec = 13
                    Case Is = 12
                        aSec = 18
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 13
                        aSec = 19
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 14
                        aSec = 24
                    Case Is = 15
                        aSec = 23
                    Case Is = 16
                        aSec = 22
                    Case Is = 17
                        aSec = 21
                    Case Is = 18
                        aSec = 20
                    Case Is = 19
                        aSec = 29
                    Case Is = 20
                        aSec = 28
                    Case Is = 21
                        aSec = 27
                    Case Is = 22
                        aSec = 26
                    Case Is = 23
                        aSec = 25
                    Case Is = 24
                        aSec = 30
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 25
                        aSec = 31
                        aRge = RangeMove("East", ThisRge)
                    Case Is = 26
                        aSec = 36
                    Case Is = 27
                        aSec = 35
                    Case Is = 28
                        aSec = 34
                    Case Is = 29
                        aSec = 33
                    Case Is = 30
                        aSec = 32
                    Case Is = 31
                        aSec = 5
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 32
                        aSec = 4
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 33
                        aSec = 3
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 34
                        aSec = 2
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 35
                        aSec = 1
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 36
                        aSec = 6
                        aTwp = TownshipMove("South", ThisTwp)
                        aRge = RangeMove("East", ThisRge)
                End Select

            Case Is = "SW"
                Select Case ThisSec
                    Case Is = 1
                        aSec = 11
                    Case Is = 2
                        aSec = 10
                    Case Is = 3
                        aSec = 9
                    Case Is = 4
                        aSec = 8
                    Case Is = 5
                        aSec = 7
                    Case Is = 6
                        aSec = 12
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 7
                        aSec = 13
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 8
                        aSec = 18
                    Case Is = 9
                        aSec = 17
                    Case Is = 10
                        aSec = 16
                    Case Is = 11
                        aSec = 15
                    Case Is = 12
                        aSec = 14
                    Case Is = 13
                        aSec = 23
                    Case Is = 14
                        aSec = 22
                    Case Is = 15
                        aSec = 21
                    Case Is = 16
                        aSec = 20
                    Case Is = 17
                        aSec = 19
                    Case Is = 18
                        aSec = 24
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 19
                        aSec = 25
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 20
                        aSec = 30
                    Case Is = 21
                        aSec = 29
                    Case Is = 22
                        aSec = 28
                    Case Is = 23
                        aSec = 27
                    Case Is = 24
                        aSec = 26
                    Case Is = 25
                        aSec = 35
                    Case Is = 26
                        aSec = 34
                    Case Is = 27
                        aSec = 33
                    Case Is = 28
                        aSec = 32
                    Case Is = 29
                        aSec = 31
                    Case Is = 30
                        aSec = 36
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 31
                        aSec = 1
                        aRge = RangeMove("West", ThisRge)
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 32
                        aSec = 6
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 33
                        aSec = 5
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 34
                        aSec = 4
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 35
                        aSec = 3
                        aTwp = TownshipMove("South", ThisTwp)
                    Case Is = 36
                        aSec = 2
                        aTwp = TownshipMove("South", ThisTwp)
                End Select

            Case Is = "NW"
                Select Case ThisSec
                    Case Is = 1
                        aSec = 35
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 2
                        aSec = 34
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 3
                        aSec = 33
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 4
                        aSec = 32
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 5
                        aSec = 31
                        aTwp = TownshipMove("North", ThisTwp)
                    Case Is = 6
                        aSec = 36
                        aTwp = TownshipMove("North", ThisTwp)
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 7
                        aSec = 1
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 8
                        aSec = 6
                    Case Is = 9
                        aSec = 5
                    Case Is = 10
                        aSec = 4
                    Case Is = 11
                        aSec = 3
                    Case Is = 12
                        aSec = 2
                    Case Is = 13
                        aSec = 11
                    Case Is = 14
                        aSec = 10
                    Case Is = 15
                        aSec = 9
                    Case Is = 16
                        aSec = 8
                    Case Is = 17
                        aSec = 7
                    Case Is = 18
                        aSec = 12
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 19
                        aSec = 13
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 20
                        aSec = 18
                    Case Is = 21
                        aSec = 17
                    Case Is = 22
                        aSec = 16
                    Case Is = 23
                        aSec = 15
                    Case Is = 24
                        aSec = 14
                    Case Is = 25
                        aSec = 23
                    Case Is = 26
                        aSec = 22
                    Case Is = 27
                        aSec = 21
                    Case Is = 28
                        aSec = 20
                    Case Is = 29
                        aSec = 19
                    Case Is = 30
                        aSec = 24
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 31
                        aSec = 25
                        aRge = RangeMove("West", ThisRge)
                    Case Is = 32
                        aSec = 30
                    Case Is = 33
                        aSec = 29
                    Case Is = 34
                        aSec = 28
                    Case Is = 35
                        aSec = 27
                    Case Is = 36
                        aSec = 26
                End Select
        End Select
    End Sub

    Private Function TownshipMove(ByVal aDirection As String, _
                                  ByVal aTwp As Integer) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Select Case aDirection
            Case Is = "North"
                TownshipMove = aTwp - 1

            Case Is = "South"
                TownshipMove = aTwp + 1
        End Select
    End Function

    Private Function RangeMove(ByVal aDirection As String, _
                               ByVal aRge As Integer) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Select Case aDirection
            Case Is = "West"
                RangeMove = aRge - 1

            Case Is = "East"
                RangeMove = aRge + 1
        End Select
    End Function

    Public Function gGetCompositeBase(ByVal aMineName As String, _
                                      ByVal aSec As Integer, _
                                      ByVal aTwp As Integer, _
                                      ByVal aRge As Integer, _
                                      ByVal aHloc As String, _
                                      ByVal aProspStandard As String, _
                                      ByRef aCompBase As gCompBaseType, _
                                      ByRef aRecCnt As Integer) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetCompositeBaseError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ProspDynaset As OraDynaset
        Dim RecordCount As Integer
        Dim HlocAlpha As String
        Dim HlocNum As String
        Dim Hloc As String
        Dim HoleIdx As Integer

        'aHloc will be a numeric hole location.
        HlocNum = aHloc
        HlocAlpha = gGetHoleLoc2(aHloc, "Char")

        gGetCompositeBase = False
        aRecCnt = 0

        'aMineName may be "".

        With aCompBase
            .MineName = ""
            .Section = 0
            .Township = 0
            .Range = 0
            .HoleLoc = ""
            .Xcoord = 0
            .Ycoord = 0
            .Elevation = 0
            .ProspDate = ""
            .OvbThk = 0
            .MtxThk = 0
            .WstThk = 0
            .TotNumSplits = 0
        End With

        For HoleIdx = 1 To 2
            If HoleIdx = 1 Then
                Hloc = HlocNum
            Else
                Hloc = HlocAlpha
            End If

            params = gDBParams

            params.Add("pMineName", aMineName, ORAPARM_INPUT)
            params("pMineName").serverType = ORATYPE_VARCHAR2

            params.Add("pHoleLocation", Hloc, ORAPARM_INPUT)
            params("pHoleLocation").serverType = ORATYPE_VARCHAR2

            params.Add("pSection", aSec, ORAPARM_INPUT)
            params("pSection").serverType = ORATYPE_VARCHAR2

            params.Add("pTownship", aTwp, ORAPARM_INPUT)
            params("pTownship").serverType = ORATYPE_VARCHAR2

            params.Add("pRange", aRge, ORAPARM_INPUT)
            params("pRange").serverType = ORATYPE_VARCHAR2

            params.Add("pProspStandard", aProspStandard, ORAPARM_INPUT)
            params("pProspStandard").serverType = ORATYPE_VARCHAR2

            params.Add("pResult", 0, ORAPARM_OUTPUT)
            params("pResult").serverType = ORATYPE_CURSOR

            'PROCEDURE get_composite_base
            'pMineName      IN     VARCHAR2,
            'pHoleLocation  IN     VARCHAR2,
            'pSection       IN     NUMBER,
            'pTownship      IN     NUMBER,
            'pRange         IN     NUMBER,
            'pProspStandard IN    VARCHAR2,
            'pResult        IN OUT c_composite)

            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect.get_composite_base(:pMineName, " + _
                          ":pHoleLocation, :pSection, :pTownship, :pRange, :pProspStandard, :pResult);end;", ORASQL_FAILEXEC)
            ProspDynaset = params("pResult").Value
            ClearParams(params)

            'Should be one row returned -- but possibly none!
            RecordCount = ProspDynaset.RecordCount
            aRecCnt = ProspDynaset.RecordCount

            If RecordCount = 1 Then
                ProspDynaset.MoveFirst()

                With aCompBase
                    .MineName = ProspDynaset.Fields("mine_name").Value
                    .Section = ProspDynaset.Fields("section").Value
                    .Township = ProspDynaset.Fields("township").Value
                    .Range = ProspDynaset.Fields("range").Value
                    .HoleLoc = ProspDynaset.Fields("hole_location").Value
                    .Xcoord = ProspDynaset.Fields("x_sp_cdnt").Value
                    .Ycoord = ProspDynaset.Fields("y_sp_cdnt").Value
                    .Elevation = ProspDynaset.Fields("hole_elevation").Value
                    .ProspDate = ProspDynaset.Fields("drill_cdate").Value
                    .OvbThk = ProspDynaset.Fields("ovb_thck").Value
                    .MtxThk = ProspDynaset.Fields("mtx_thck").Value
                    .WstThk = ProspDynaset.Fields("wst_thck").Value
                    .TotNumSplits = ProspDynaset.Fields("split_total_num").Value
                End With

                gGetCompositeBase = True
                Exit For
            End If
        Next HoleIdx

        ProspDynaset.Close()
        Exit Function

gGetCompositeBaseError:
        MsgBox("Error getting composite base data." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Process Error")
        gGetCompositeBase = False
        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        ProspDynaset.Close()
    End Function

    Public Function gGetSplitBase(ByVal aMineName As String, _
                                  ByVal aSec As Integer, _
                                  ByVal aTwp As Integer, _
                                  ByVal aRge As Integer, _
                                  ByVal aHloc As String, _
                                  ByVal aProspStandard As String, _
                                  ByRef aSplitBase() As gSplitBaseType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetSplitBaseError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ProspDynaset As OraDynaset
        Dim RecordCount As Integer
        Dim HlocAlpha As String
        Dim HlocNum As String
        Dim Hloc As String
        Dim SplIdx As Integer
        Dim SplCnt As Integer

        'aHloc will be a numeric hole location.
        HlocNum = aHloc
        HlocAlpha = gGetHoleLoc2(aHloc, "Char")

        'aMineName may be "".

        'Will check to see if splits exist for this hole for both the numeric hole location
        'and the alpha hole location -- it will be one or the other, never both.

        For SplIdx = 1 To 2
            If SplIdx = 1 Then
                Hloc = HlocNum
            Else
                Hloc = HlocAlpha
            End If

            params = gDBParams

            params.Add("pMineName", aMineName, ORAPARM_INPUT)
            params("pMineName").serverType = ORATYPE_VARCHAR2

            params.Add("pHoleLocation", Hloc, ORAPARM_INPUT)
            params("pHoleLocation").serverType = ORATYPE_VARCHAR2

            params.Add("pSection", aSec, ORAPARM_INPUT)
            params("pSection").serverType = ORATYPE_VARCHAR2

            params.Add("pTownship", aTwp, ORAPARM_INPUT)
            params("pTownship").serverType = ORATYPE_VARCHAR2

            params.Add("pRange", aRge, ORAPARM_INPUT)
            params("pRange").serverType = ORATYPE_VARCHAR2

            params.Add("pProspStandard", aProspStandard, ORAPARM_INPUT)
            params("pProspStandard").serverType = ORATYPE_VARCHAR2

            params.Add("pResult", 0, ORAPARM_OUTPUT)
            params("pResult").serverType = ORATYPE_CURSOR

            'PROCEDURE get_split_base
            'pMineName      IN     VARCHAR2,
            'pHoleLocation  IN     VARCHAR2,
            'pSection       IN     NUMBER,
            'pTownship      IN     NUMBER,
            'pRange         IN     NUMBER,
            'pProspStandard IN     VARCHAR2,
            'pResult        IN OUT c_splits)
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prospect.get_split_base(:pMineName, " + _
                          ":pHoleLocation, :pSection, :pTownship, :pRange, :pProspStandard, :pResult);end;", ORASQL_FAILEXEC)
            ProspDynaset = params("pResult").Value
            ClearParams(params)

            'Should be multiple rows returned -- but possibly none!
            RecordCount = ProspDynaset.RecordCount

            If RecordCount >= 1 Then
                ReDim aSplitBase(RecordCount)
                SplCnt = 0

                ProspDynaset.MoveFirst()
                Do While Not ProspDynaset.EOF
                    SplCnt = SplCnt + 1
                    aSplitBase(SplCnt).MineName = ProspDynaset.Fields("mine_name").Value
                    aSplitBase(SplCnt).Section = ProspDynaset.Fields("section").Value
                    aSplitBase(SplCnt).Township = ProspDynaset.Fields("township").Value
                    aSplitBase(SplCnt).Range = ProspDynaset.Fields("range").Value
                    aSplitBase(SplCnt).HoleLoc = ProspDynaset.Fields("hole_location").Value
                    aSplitBase(SplCnt).ProspDate = ProspDynaset.Fields("drill_cdate").Value
                    aSplitBase(SplCnt).Split = ProspDynaset.Fields("split").Value
                    aSplitBase(SplCnt).SplitDepthTop = ProspDynaset.Fields("split_depth_top").Value
                    aSplitBase(SplCnt).SplitDepthBot = ProspDynaset.Fields("bot_split_depth").Value
                    aSplitBase(SplCnt).SplitThk = ProspDynaset.Fields("split_thck").Value
                    aSplitBase(SplCnt).MinableStatus = ProspDynaset.Fields("minable_status").Value
                    ProspDynaset.MoveNext()
                Loop

                gGetSplitBase = True
                Exit For
            End If
        Next SplIdx

        ProspDynaset.Close()
        Exit Function

gGetSplitBaseError:
        MsgBox("Error getting split base data." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Process Error")
        gGetSplitBase = False
        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        ProspDynaset.Close()
    End Function

    '    Public Sub gPrintComposite(ByRef aComposite As gProspectComposite, _
    '                               ByRef aRptProspectComposite As CrystalReport, _
    '                               ByVal aCompanyName As String, _
    '                               ByVal aProspgridMineDefault As String)

    '        '**********************************************************************
    '        '
    '        '
    '        '
    '        '**********************************************************************

    '        On Error GoTo gPrintCompositeError

    '        Dim ReportString As String
    '        Dim ConnectString As String
    '        Dim HoleLocOther As String

    '        'Connect to Oracle database
    '        aRptProspectComposite.Reset()
    '        aRptProspectComposite.ReportFileName = gPath + "\Reports\" + "ProspectComposite.rpt"

    '        ConnectString = "DSN = " + gDataSource + ";UID = " + gOracleUserName + _
    '            ";PWD = " + gOracleUserPassword + ";DSQ = "

    '        aRptProspectComposite.Connect = ConnectString

    '        With aComposite
    '            aRptProspectComposite.Formulas(0) = "Mine = '" & .Mine & "'"
    '            aRptProspectComposite.Formulas(1) = "Section = " & .Section & ""
    '            aRptProspectComposite.Formulas(2) = "Township = " & .Township & ""
    '            aRptProspectComposite.Formulas(3) = "Range = " & .Range & ""

    '            'This will be the "Main" hole location -- it may be Alpha-numeric or
    '            'Numeric depending on the mine.
    '            aRptProspectComposite.Formulas(4) = "HoleLocation = '" & .HoleLocation & "'"

    '            aRptProspectComposite.Formulas(5) = "DrillCdate = '" & .DrillCdate & "'"
    '            aRptProspectComposite.Formulas(6) = "AnalysisCdate = '" & .AnalysisCdate & "'"
    '            aRptProspectComposite.Formulas(7) = "AreaOfInfluence = '" & .AreaOfInfluence & "'"
    '            aRptProspectComposite.Formulas(8) = "HoleElevation = " & .HoleElevation & ""
    '            aRptProspectComposite.Formulas(9) = "PitBottomElevation = " & .PitBottomElevation & ""
    '            aRptProspectComposite.Formulas(10) = "XSPCoordinate = " & .XSPCoordinate & ""
    '            aRptProspectComposite.Formulas(11) = "YSPCoordinate = " & .YSPCoordinate & ""
    '            aRptProspectComposite.Formulas(12) = "TriangleCode = '" & Trim(.TriangleCode) & "'"
    '            aRptProspectComposite.Formulas(13) = "TotalNumberSplits = " & .TotalNumberSplits & ""
    '            aRptProspectComposite.Formulas(14) = "SplitsSummarized = '" & Trim(.SplitsSummarized) & "'"
    '            'arptprospectcomposite.Formulas(15) = "ProspectorCode = '" & .ProspectorCode & "'"
    '            aRptProspectComposite.Formulas(16) = "OvbThickness = " & .OvbThickness & ""
    '            aRptProspectComposite.Formulas(17) = "OvbX = " & .OvbX & ""
    '            aRptProspectComposite.Formulas(18) = "MtxThickness = " & .MtxThickness & ""
    '            aRptProspectComposite.Formulas(19) = "MtxPebbleX = " & .MtxPebbleX & ""
    '            aRptProspectComposite.Formulas(20) = "MtxX = " & .MtxX & ""
    '            aRptProspectComposite.Formulas(21) = "TotalThickness = " & .TotalThickness & ""
    '            aRptProspectComposite.Formulas(22) = "TotalPebbleX = " & .TotalPebbleX & ""
    '            aRptProspectComposite.Formulas(23) = "TotalX = " & .TotalX & ""
    '            aRptProspectComposite.Formulas(24) = "MtxPercentSolids = " & .MtxPercentSolids & ""
    '            aRptProspectComposite.Formulas(25) = "MtxWetDensity = " & .MtxWetDensity & ""
    '            aRptProspectComposite.Formulas(26) = "CoarsePebbleWtp = " & .CoarsePebbleWtp & ""
    '            aRptProspectComposite.Formulas(27) = "FinePebbleWtp = " & .FinePebbleWtp & ""
    '            aRptProspectComposite.Formulas(28) = "TotalPebbleWtp = " & .TotalPebbleWtp & ""
    '            aRptProspectComposite.Formulas(29) = "ConcentrateWtp = " & .ConcentrateWtp & ""
    '            aRptProspectComposite.Formulas(30) = "TotalProductWtp = " & .TotalProductWtp & ""
    '            aRptProspectComposite.Formulas(31) = "TotalTailWtp = " & .TotalTailWtp & ""
    '            aRptProspectComposite.Formulas(32) = "WasteClayWtp = " & .WasteClayWtp & ""
    '            aRptProspectComposite.Formulas(33) = "GrossConcentrateWtp = " & .GrossConcentrateWtp & ""
    '            aRptProspectComposite.Formulas(34) = "GrossProductWtp = " & .GrossProductWtp & ""
    '            aRptProspectComposite.Formulas(35) = "CoarseFeedWtp = " & .CoarseFeedWtp & ""
    '            aRptProspectComposite.Formulas(36) = "FineFeedWtp = " & .FineFeedWtp & ""
    '            aRptProspectComposite.Formulas(37) = "TotalFeedWtp = " & .TotalFeedWtp & ""
    '            aRptProspectComposite.Formulas(38) = "MtxTons = " & .MtxTons & ""
    '            aRptProspectComposite.Formulas(39) = "CoarsePebbleTPA = " & .CoarsePebbleTPA & ""
    '            aRptProspectComposite.Formulas(40) = "FinePebbleTPA = " & .FinePebbleTPA & ""
    '            aRptProspectComposite.Formulas(41) = "TotalPebbleTPA = " & .TotalPebbleTpa & ""
    '            aRptProspectComposite.Formulas(42) = "ConcentrateTPA = " & .ConcentrateTPA & ""
    '            aRptProspectComposite.Formulas(43) = "TotalProductTPA = " & .TotalProductTpa & ""
    '            aRptProspectComposite.Formulas(44) = "TotalTailTPA = " & .TotalTailTpa & ""
    '            aRptProspectComposite.Formulas(45) = "WasteClayTPA = " & .WasteClayTpa & ""
    '            aRptProspectComposite.Formulas(46) = "GrossConcentrateTPA = " & .GrossConcentrateTpa & ""
    '            aRptProspectComposite.Formulas(47) = "GrossProductTPA = " & .GrossProductTpa & ""
    '            aRptProspectComposite.Formulas(48) = "CoarseFeedTPA = " & .CoarseFeedTpa & ""
    '            aRptProspectComposite.Formulas(49) = "FineFeedTPA = " & .FineFeedTpa & ""
    '            aRptProspectComposite.Formulas(50) = "TotalFeedTPA = " & .TotalFeedTpa & ""
    '            'arptprospectcomposite.Formulas(51) = "MtxBPL = " & .MtxBPL & ""
    '            aRptProspectComposite.Formulas(52) = "CoarsePebbleBPL = " & .CoarsePebbleBPL & ""
    '            aRptProspectComposite.Formulas(53) = "FinePebbleBPL = " & .FinePebbleBPL & ""
    '            aRptProspectComposite.Formulas(54) = "TotalPebbleBPL = " & .TotalPebbleBpl & ""
    '            aRptProspectComposite.Formulas(55) = "ConcentrateBPL = " & .ConcentrateBPL & ""
    '            aRptProspectComposite.Formulas(56) = "TotalProductBPL = " & .TotalProductBpl & ""
    '            aRptProspectComposite.Formulas(57) = "TotalTailBPL = " & .TotalTailBPL & ""
    '            aRptProspectComposite.Formulas(58) = "WasteClayBPL = " & .WasteClayBPL & ""
    '            aRptProspectComposite.Formulas(59) = "GrossConcentrateBPL = " & .GrossConcentrateBpl & ""
    '            aRptProspectComposite.Formulas(60) = "GrossProductBPL = " & .GrossProductBpl & ""
    '            aRptProspectComposite.Formulas(61) = "CoarseFeedBPL = " & .CoarseFeedBpl & ""
    '            aRptProspectComposite.Formulas(62) = "FineFeedBPL = " & .FineFeedBpl & ""
    '            aRptProspectComposite.Formulas(63) = "TotalFeedBPL = " & .TotalFeedBpl & ""
    '            aRptProspectComposite.Formulas(64) = "FinePebbleFe2O3 = " & .FinePebbleFe2O3 & ""
    '            aRptProspectComposite.Formulas(65) = "FinePebbleAl2O3 = " & .FinePebbleAl2O3 & ""
    '            aRptProspectComposite.Formulas(66) = "FinePebbleMgO = " & .FinePebbleMgO & ""
    '            aRptProspectComposite.Formulas(67) = "FinePebbleCaO = " & .FinePebbleCaO & ""
    '            aRptProspectComposite.Formulas(68) = "FinePebbleInsol = " & .FinePebbleInsol & ""
    '            aRptProspectComposite.Formulas(69) = "FinePebbleIA = " & .FinePebbleIa & ""
    '            aRptProspectComposite.Formulas(70) = "CoarsePebbleFe2O3 = " & .CoarsePebbleFe2O3 & ""
    '            aRptProspectComposite.Formulas(71) = "CoarsePebbleAl2O3 = " & .CoarsePebbleAl2O3 & ""
    '            aRptProspectComposite.Formulas(72) = "CoarsePebbleMgO = " & .CoarsePebbleMgO & ""
    '            aRptProspectComposite.Formulas(73) = "CoarsePebbleCaO = " & .CoarsePebbleCaO & ""
    '            aRptProspectComposite.Formulas(74) = "CoarsePebbleInsol = " & .CoarsePebbleInsol & ""
    '            aRptProspectComposite.Formulas(75) = "CoarsePebbleIA = " & .CoarsePebbleIa & ""
    '            aRptProspectComposite.Formulas(76) = "TotalPebbleFe2O3 = " & .TotalPebbleFe2O3 & ""
    '            aRptProspectComposite.Formulas(77) = "TotalPebbleAl2O3 = " & .TotalPebbleAl2O3 & ""
    '            aRptProspectComposite.Formulas(78) = "TotalPebbleMgO = " & .TotalPebbleMgO & ""
    '            aRptProspectComposite.Formulas(79) = "TotalPebbleCaO = " & .TotalPebbleCaO & ""
    '            aRptProspectComposite.Formulas(80) = "TotalPebbleInsol = " & .TotalPebbleInsol & ""
    '            aRptProspectComposite.Formulas(81) = "TotalPebbleIA = " & .TotalPebbleIa & ""
    '            aRptProspectComposite.Formulas(82) = "ConcentrateFe2O3 = " & .ConcentrateFe2O3 & ""
    '            aRptProspectComposite.Formulas(83) = "ConcentrateAl2O3 = " & .ConcentrateAl2O3 & ""
    '            aRptProspectComposite.Formulas(84) = "ConcentrateMgO = " & .ConcentrateMgO & ""
    '            aRptProspectComposite.Formulas(85) = "ConcentrateCaO = " & .ConcentrateCaO & ""
    '            aRptProspectComposite.Formulas(86) = "ConcentrateInsol = " & .ConcentrateInsol & ""
    '            aRptProspectComposite.Formulas(87) = "ConcentrateIA = " & .ConcentrateIA & ""
    '            aRptProspectComposite.Formulas(88) = "TotalProductFe2O3 = " & .TotalProductFe2O3 & ""
    '            aRptProspectComposite.Formulas(89) = "TotalProductAl2O3 = " & .TotalProductAl2O3 & ""
    '            aRptProspectComposite.Formulas(90) = "TotalProductMgO = " & .TotalProductMgO & ""
    '            aRptProspectComposite.Formulas(91) = "TotalProductCaO = " & .TotalProductCaO & ""
    '            aRptProspectComposite.Formulas(92) = "TotalProductInsol = " & .TotalProductInsol & ""
    '            aRptProspectComposite.Formulas(93) = "TotalProductIA = " & .TotalProductIA & ""
    '            aRptProspectComposite.Formulas(94) = "GrossConcentrateInsol = " & .GrossConcentrateInsol & ""
    '            aRptProspectComposite.Formulas(95) = "GrossProductInsol = " & .GrossProductInsol & ""
    '            '----------
    '            aRptProspectComposite.Formulas(96) = "MtxDryDensity = " & .MtxDryDensity & ""
    '            aRptProspectComposite.Formulas(97) = "MtxConcentrateX = " & .MtxConcentrateX & ""
    '            aRptProspectComposite.Formulas(98) = "TotalConcentrateX = " & .TotalConcentrateX & ""
    '            '----------
    '            aRptProspectComposite.Formulas(99) = "WstThck = " & .WstThck & ""
    '            aRptProspectComposite.Formulas(100) = "TotX = " & .TotX & ""
    '            aRptProspectComposite.Formulas(101) = "MinableSplits = '" & Trim(.MinableSplits) & "'"
    '            aRptProspectComposite.Formulas(102) = "HoleMinable = '" & Trim(.HoleMinable) & "'"
    '            aRptProspectComposite.Formulas(103) = "CpbMinable = '" & Trim(.CpbMinable) & "'"
    '            aRptProspectComposite.Formulas(104) = "FpbMinable = '" & Trim(.FpbMinable) & "'"
    '            aRptProspectComposite.Formulas(105) = "FltBplRcvryCalc = " & .FltBplRcvryCalc & ""
    '            aRptProspectComposite.Formulas(106) = "Rc = " & .Rc & ""
    '            aRptProspectComposite.Formulas(107) = "MtxYdsPerAcre = " & .MtxYdsPerAcre & ""
    '            aRptProspectComposite.Formulas(108) = "CpIa = " & .CpIa & ""
    '            aRptProspectComposite.Formulas(109) = "FpIa = " & .FpIa & ""
    '            aRptProspectComposite.Formulas(110) = "CnIa = " & .CnIa & ""
    '            aRptProspectComposite.Formulas(111) = "TpIa = " & .TpIA & ""
    '            aRptProspectComposite.Formulas(112) = "TpbIa = " & .TpbIA & ""
    '            aRptProspectComposite.Formulas(113) = "WstPbWtp = " & .WstPbWtp & ""
    '            aRptProspectComposite.Formulas(114) = "WstPbTpa = " & .WstPbTpa & ""
    '            aRptProspectComposite.Formulas(115) = "WstPbBpl = " & .WstPbBpl & ""
    '            aRptProspectComposite.Formulas(116) = "WstPbFe = " & .WstPbFe & ""
    '            aRptProspectComposite.Formulas(117) = "WstPbAl = " & .WstPbAl & ""
    '            aRptProspectComposite.Formulas(118) = "WstPbMg = " & .WstPbMg & ""
    '            aRptProspectComposite.Formulas(119) = "WstPbCa = " & .WstPbCa & ""
    '            aRptProspectComposite.Formulas(120) = "WstPbIa1 = " & .WstPbIa1 & ""
    '            aRptProspectComposite.Formulas(121) = "WstPbIa2 = " & .WstPbIa2 & ""
    '            aRptProspectComposite.Formulas(122) = "WstPbIns = " & .WstPbIns & ""
    '            '----------
    '            aRptProspectComposite.Formulas(123) = "CpbCd = " & .CpbCd & ""
    '            aRptProspectComposite.Formulas(124) = "FpbCd = " & .FpbCd & ""
    '            aRptProspectComposite.Formulas(125) = "TcnCd = " & .TcnCd & ""
    '            aRptProspectComposite.Formulas(126) = "TprCd = " & .TprCd & ""
    '            aRptProspectComposite.Formulas(127) = "TpbCd = " & .TpbCd & ""
    '            aRptProspectComposite.Formulas(128) = "HardpanCode = " & .HardpanCode & ""
    '            '----------
    '            'This will be the "other" hole location (alpha-numeric if the "main
    '            'location" is numeric and numeric if the main location is
    '            'alpha-numeric).

    '            If aProspgridMineDefault = "Alpha-numeric" Then
    '                If IsNumeric(.HoleLocation) = False Then
    '                    HoleLocOther = gGetHoleLoc2(.HoleLocation, "Num")
    '                Else
    '                    'Actually probably have a numeric hole location passed in.
    '                    HoleLocOther = gGetHoleLoc2(.HoleLocation, "Char")
    '                End If
    '            Else    'fProspGridMineDefault = "Numeric"
    '                If IsNumeric(.HoleLocation) = True Then
    '                    HoleLocOther = gGetHoleLoc2(.HoleLocation, "Char")
    '                Else
    '                    'Actually probably have a alpha-numeric hole location passed in.
    '                    HoleLocOther = gGetHoleLoc2(.HoleLocation, "Num")
    '                End If
    '            End If
    '            If HoleLocOther = "???" Then
    '                HoleLocOther = " "
    '            End If
    '            aRptProspectComposite.Formulas(129) = "HoleLocOther = '" & HoleLocOther & "'"
    '            '----------
    '            aRptProspectComposite.Formulas(130) = "ProspStandard = '" & .ProspStandard & "'"
    '        End With

    '        'Need to pass the company name into the report
    '        aRptProspectComposite.ParameterFields(0) = "pCompanyName;" & aCompanyName & ";TRUE"

    '        'Report window maximized
    '        aRptProspectComposite.WindowState = crptMaximized

    '        aRptProspectComposite.WindowTitle = "Prospect Composite Hole"

    '        'User allowed to minimize report window
    '        aRptProspectComposite.WindowMinButton = True

    '        'Start Crystal Reports
    '        aRptProspectComposite.action = 1
    '        aRptProspectComposite.Reset()

    '        Exit Sub

    'gPrintCompositeError:
    '        MsgBox("Error printing composite." & vbCrLf & _
    '               Err.Description, _
    '               vbOKOnly + vbExclamation, _
    '               "Composite Print Error")
    '    End Sub

    Public Function gIsHalfHole(ByVal aHole As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim FirstNumber As Integer
        Dim SecondNumber As Integer

        'Numeric holes that are not half-holes are divisible by two (they are even).

        If IsNumeric(aHole) = False Then
            gIsHalfHole = False
            Exit Function
        End If

        If Len(aHole) <> 4 Then
            gIsHalfHole = False
            Exit Function
        End If

        FirstNumber = Val(Mid(aHole, 1, 2))
        SecondNumber = Val(Mid(aHole, 3))

        If gIsEvenNumber(FirstNumber) = True And _
            gIsEvenNumber(SecondNumber) = True Then
            gIsHalfHole = False
        Else
            gIsHalfHole = True
        End If
    End Function

    Public Function gGetTotalValue3(ByVal aValue1 As Single, _
                                    ByVal aTpa1 As Single, _
                                    ByVal aValue2 As Single, _
                                    ByVal aTpa2 As Single, _
                                    ByVal aValue3 As Single, _
                                    ByVal aTpa3 As Single, _
                                    ByVal aRound As Integer) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Will not average in zero values!

        Dim TpaWVal As Double
        Dim TpaVal As Double

        gGetTotalValue3 = 0

        TpaWVal = 0
        TpaVal = 0

        TpaVal = aTpa1 * aValue1 + aTpa2 * aValue2 + aTpa3 * aValue3

        If aValue1 > 0 Then
            TpaWVal = TpaWVal + aTpa1
        End If
        If aValue2 > 0 Then
            TpaWVal = TpaWVal + aTpa2
        End If
        If aValue3 > 0 Then
            TpaWVal = TpaWVal + aTpa3
        End If

        If TpaWVal <> 0 Then
            gGetTotalValue3 = gRound(TpaVal / TpaWVal, aRound)
        Else
            gGetTotalValue3 = 0
        End If
    End Function

    Public Function gGetHoleInSfmHardee(ByVal aSection As Integer, _
                                        ByVal aTownship As Integer, _
                                        ByVal aRange As Integer) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        gGetHoleInSfmHardee = False

        If aTownship = 33 Then
            If aRange = 25 Then
                If aSection = 1 Or aSection = 2 Or aSection = 3 Or aSection = 4 Or _
                    aSection = 9 Or aSection = 10 Or aSection = 11 Or aSection = 12 Or _
                    aSection = 13 Or aSection = 14 Or aSection = 15 Or aSection = 22 Or _
                    aSection = 23 Or aSection = 24 Or aSection = 25 Or aSection = 26 Or _
                    aSection = 27 Or aSection = 34 Or aSection = 35 Or aSection = 36 Then
                    gGetHoleInSfmHardee = True
                End If
            End If

            If aRange = 26 Then
                If aSection = 5 Or aSection = 6 Or aSection = 7 Or aSection = 8 Or _
                    aSection = 18 Or aSection = 19 Or aSection = 30 Then
                    gGetHoleInSfmHardee = True
                End If
            End If
        End If
    End Function


End Module
