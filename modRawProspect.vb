Option Explicit On
Imports OracleInProcServer
Imports System.Math
Module modRawProspect
    'Attribute VB_Name = "modRawProspect"
    '**********************************************************************
    'RAW PROSPECT MODULE
    '
    '
    '**********************************************************************
    '   Maintenance Log
    '
    '   12/30/2004, lss
    '       Added this module.
    '   01/03/2005, lss
    '       Added Function gSampNumExists().
    '       Added Function Function gGetDrillHoleSamples().
    '       Added Function gDispIntervals().
    '   01/17/2005, lss
    '       Added Function gGetCodeDesc().
    '   01/20/2005, lss
    '       Added Function gCodeFromCodeDesc().
    '   03/06/07, lss
    '       Will not check for Cd -- Cd's are not run on prospect
    '       anymore! (per Earnest Terry).
    '   08/10/2007, lss
    '       Added IsNull check for .HoleLocation in Function
    '       gGetProspRawData.
    '   10/20/2010, lss
    '       Added Public Function gGetOldRawProspect.
    '
    '**********************************************************************


    'Public type gProspRawDataType -- used by Metallurgical Lab (raw
    'prospect data)
    Public Structure gProspRawDataType
        Public MineName As String                  '1
        Public SampNum As String                   '2
        Public DrillDate As Date                   '3
        Public WashDate As Date                    '4
        Public LogDate As Date                     '5
        Public HoleLocation As String              '6
        Public Split As Integer                    '7
        Public Section As Integer                  '8
        Public Township As Integer                 '9
        Public Range As Integer                    '10
        Public SplitTotalNum As Integer            '11
        Public SplitDepthTop As Double             '12
        Public SplitDepthBot As Double             '13
        Public NetWeight As Double                 '14
        Public MtxWetWt As Double                  '15
        Public MtxDryWt As Double                  '16
        Public MtxTareWt As Double                 '17
        Public WetCoreWasher As Double             '18
        Public MinutesMixed As Integer             '19
        '----
        Public CrsPbDryLbs As Double               '20
        Public CrsPbBpl As Double                  '21
        Public CrsPbIns As Double                  '22
        Public CrsPbFe As Double                   '23
        Public CrsPbAl As Double                   '24
        Public CrsPbMg As Double                   '25
        Public CrsPbCa As Double                   '26
        '----
        Public FnePbDryLbs As Double               '27
        Public FnePbBpl As Double                  '28
        Public FnePbIns As Double                  '29
        Public FnePbFe As Double                   '30
        Public FnePbAl As Double                   '31
        Public FnePbMg As Double                   '32
        Public FnePbCa As Double                   '33
        '-----
        Public FdM16P150WetLbs As Double           '34
        Public FdM16P150Bpl As Double              '35
        '-----
        Public WasteClayBPL As Double              '36
        '-----
        Public FdWetGms As Double                  '37
        Public FdDryGms As Double                  '38
        Public FdTareGms As Double                 '39
        Public FdFlotWetGms As Double              '40
        '-----
        Public FdP35WtGms As Double                '41
        Public FdP35Bpl As Double                  '42
        Public FdM35WtGms As Double                '43
        Public FdM35Bpl As Double                  '44
        '-----
        Public FdRghrTailsWt As Double             '45
        Public FdRghrTailsBpl As Double            '46
        '-----
        Public FdClnrTailsWt As Double             '47
        Public FdClnrTailsBpl As Double            '48
        '-----
        Public FdAmineCnWt As Double               '49
        Public FdAmineCnBpl As Double              '50
        Public FdAmineCnIns As Double              '51
        Public FdAmineCnFe As Double               '52
        Public FdAmineCnAl As Double               '53
        Public FdAmineCnMg As Double               '54
        Public FdAmineCnCa As Double               '55
        Public FdAmineCnColor As String            '56
        '-----
        Public FdM35WetGms As Double               '57
        Public FdM35DryGms As Double               '58
        Public FdM35TareGms As Double              '59
        Public FdM35FlotWetGms As Double           '60
        '-----
        Public FdM35AmineCnWt As Double            '61
        Public FdM35AmineCnBpl As Double           '62
        Public FdM35AmineCnIns As Double           '63
        Public FdM35AmineCnFe As Double            '64
        Public FdM35AmineCnAl As Double            '65
        Public FdM35AmineCnMg As Double            '66
        Public FdM35AmineCnCa As Double            '67
        Public FdM35AmineCnColor As String         '68
        '-----
        Public FdM35RghrTailsWt As Double          '69
        Public FdM35RghrTailsBpl As Double         '70
        '-----
        Public FdM35ClnrTailsWt As Double          '71
        Public FdM35ClnrTailsBpl As Double         '72
        '-----
        Public CrsPbCd As Double                   '73
        Public FnePbCd As Double                   '74
        Public FdAmineCnCd As Double               '75
        Public FdM35AmineCnCd As Double            '76
        '-----
        Public MetLabComment As String             '77
        Public ChemLabComment As String            '78
        '-----
        Public DateChemLab As Date                 '79
        Public WhoChemLab As String                '80
        Public RecordLocked As Integer             '81
        Public RerunStatus As Integer              '82
        Public DateRerun As Date                   '83
        '-----
        Public HardpanCode As Integer              '84
    End Structure

    Public Structure gHoleIntervalType
        Public TosDepth As Single
        Public BosDepth As Single
        Public SampNum As String
        Public DrillDate As Date
        Public Split As Integer
    End Structure

    Public Function gGetProspRawData(ByVal aSampleId As String, _
                                     ByRef aProspRawData As gProspRawDataType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim params As OraParameter
        Dim SQLStmt As OraSqlStmt

        Dim NumOfSplits As Integer
        Dim SplitCount As Integer
        Dim ThisRerunStatus As Integer

        On Error GoTo gGetProspRawDataError

        ZeroProspRaw(aProspRawData)

        'Procedure get_raw_prospect
        'pSampleId                  IN OUT VARCHAR2,    -- 1
        'pMineName                  OUT    VARCHAR2,    -- 2
        'pTownship                  OUT    NUMBER,      -- 3
        'pRange                     OUT    NUMBER,      -- 4
        'pSection                   OUT    NUMBER,      -- 5
        'pHoleLocation              OUT    VARCHAR2,    -- 6
        'pSplit                     OUT    NUMBER,      -- 7
        'pSplitTotalNum             OUT    NUMBER,      -- 8
        'pDrillDate                 OUT    DATE,        -- 9
        'pWashDate                  OUT    DATE,        -- 10
        'pLogDate                   OUT    DATE,        -- 11
        'pSplitDepthTop             OUT    NUMBER,      -- 12
        'pSplitDepthBot             OUT    NUMBER,      -- 13
        'pNetWeight                 OUT    NUMBER,      -- 14
        'pMtxWetWt                  OUT    NUMBER,      -- 15
        'pMtxDryWt                  OUT    NUMBER,      -- 16
        'pMtxTareWt                 OUT    NUMBER,      -- 17
        'pWetCoreWasher             OUT    NUMBER,      -- 18
        'pMinutesMixed              OUT    NUMBER,      -- 19
        'pHalfPbDryLbs              OUT    NUMBER,      -- 20
        'pHalfPbBpl                 OUT    NUMBER,      -- 21
        'pHalfPbInsol               OUT    NUMBER,      -- 22
        'pHalfPbFe2O3               OUT    NUMBER,      -- 23
        'pHalfPbAl2O3               OUT    NUMBER,      -- 24
        'pHalfPbMgO                 OUT    NUMBER,      -- 25
        'pP16mPbDryLbs              OUT    NUMBER,      -- 26
        'pP16mPbBpl                 OUT    NUMBER,      -- 27
        'pP16mPbInsol               OUT    NUMBER,      -- 28
        'pP16mPbFe2O3               OUT    NUMBER,      -- 29
        'pP16mPbAl2O3               OUT    NUMBER,      -- 30
        'pP16mPbMgO                 OUT    NUMBER,      -- 31
        'pP150mFdWetLbs             OUT    NUMBER,      -- 32
        'pP150mFdBpl                OUT    NUMBER,      -- 33
        'pM150mWasteClayBpl         OUT    NUMBER,      -- 34
        'pTorP35mFdWetGrams         OUT    NUMBER,      -- 35
        'pTorP35mFdDryGrams         OUT    NUMBER,      -- 36
        'pTorP35mFdTareGrams        OUT    NUMBER,      -- 37
        'pP35mFdWt                  OUT    NUMBER,      -- 38
        'pP35mFdBpl                 OUT    NUMBER,      -- 39
        'pM35mFdWt                  OUT    NUMBER,      -- 40
        'pM35mFdBpl                 OUT    NUMBER,      -- 41
        'pTorP35mFdFlotSampWetGrams OUT    NUMBER,      -- 42
        'pTorP35mFdRghrTlngsWt      OUT    NUMBER,      -- 43
        'pTorP35mFdRghrTlngsBpl     OUT    NUMBER,      -- 44
        'pTorP35mFdClnrTlngsWt      OUT    NUMBER,      -- 45
        'pTorP35mFdClnrTlngsBpl     OUT    NUMBER,      -- 46
        'pTorP35mFdAmineCnWt        OUT    NUMBER,      -- 47
        'pTorP35mFdAmineCnBpl       OUT    NUMBER,      -- 48
        'pTorP35mFdAmineCnInsol     OUT    NUMBER,      -- 49
        'pTorP35mFdAmineCnFe2O3     OUT    NUMBER,      -- 50
        'pTorP35mFdAmineCnAl2O3     OUT    NUMBER,      -- 51
        'pTorP35mFdAmineCnMgO       OUT    NUMBER,      -- 52
        'pTorP35mFdAmineCnColor     OUT    VARCHAR2,    -- 53
        'pM35mFdAmineCnWt           OUT    NUMBER,      -- 54
        'pM35mFdAmineCnBpl          OUT    NUMBER,      -- 55
        'pM35mFdAmineCnInsol        OUT    NUMBER,      -- 56
        'pM35mFdAmineCnFe2O3        OUT    NUMBER,      -- 57
        'pM35mFdAmineCnAl2O3        OUT    NUMBER,      -- 58
        'pM35mFdAmineCnMgO          OUT    NUMBER,      -- 59
        'pM35mFdAmineCnColor        OUT    VARCHAR2,    -- 60
        'pM35mFdWetGrams            OUT    NUMBER,      -- 61
        'pM35mFdDryGrams            OUT    NUMBER,      -- 62
        'pM35mFdTareGrams           OUT    NUMBER,      -- 63
        'pM35mFdFlotSampWetGrams    OUT    NUMBER,      -- 64
        'pM35mFdRghrTlngsWt         OUT    NUMBER,      -- 65
        'pM35mFdRghrTlngsBpl        OUT    NUMBER,      -- 66
        'pM35mFdClnrTlngsWt         OUT    NUMBER,      -- 67
        'pM35mFdClnrTlngsBpl        OUT    NUMBER,      -- 68
        'pHalfPbCaO                 OUT    NUMBER,      -- 69
        'pP16mPbCaO                 OUT    NUMBER,      -- 70
        'pTorP35mFdAmineCnCaO       OUT    NUMBER,      -- 71
        'pM35mFdAmineCnCaO          OUT    NUMBER,      -- 72
        'pDateChemLab               OUT    DATE,        -- 73
        'pWhoChemLab                OUT    VARCHAR2,    -- 74
        'pRecordLocked              OUT    NUMBER,      -- 75
        'pRerunStatus               OUT    NUMBER,      -- 76
        'pDateRerun                 OUT    DATE,        -- 77
        'pMetLabComment             OUT    VARCHAR2,    -- 78
        'pChemLabComment            OUT    VARCHAR2,    -- 79
        'pHalfPbCd                  OUT    NUMBER,      -- 80
        'pP16mPbCd                  OUT    NUMBER,      -- 81
        'pTorP35mFdAmineCnCd        OUT    NUMBER,      -- 82
        'pM35mFdAmineCnCd           OUT    NUMBER,      -- 83
        'pHardpanCode               OUT    NUMBER);     -- 84

        params = gDBParams

        '1 Sample ID
        params.Add("pSampleId", aSampleId, ORAPARM_OUTPUT)
        params("pSampleId").serverType = ORATYPE_VARCHAR2

        '2  Mine name
        params.Add("pMineName", "", ORAPARM_OUTPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        '3  Township
        params.Add("pTownShip", 0, ORAPARM_OUTPUT)
        params("pTownShip").serverType = ORATYPE_NUMBER

        '4  Range
        params.Add("pRange", 0, ORAPARM_OUTPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        '5  Section
        params.Add("pSection", 0, ORAPARM_OUTPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        '6  Hole location
        params.Add("pHoleLocation", "", ORAPARM_OUTPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        '7  Split
        params.Add("pSplit", 0, ORAPARM_OUTPUT)
        params("pSplit").serverType = ORATYPE_NUMBER

        '8  Total number of splits in hole
        params.Add("pSplitTotalNum", 0, ORAPARM_OUTPUT)
        params("pSplitTotalNum").serverType = ORATYPE_NUMBER

        '9  Drill date
        params.Add("pDrillDate", "", ORAPARM_OUTPUT)
        params("pDrillDate").serverType = ORATYPE_DATE

        '10 Wash date
        params.Add("pWashDate", "", ORAPARM_OUTPUT)
        params("pWashDate").serverType = ORATYPE_DATE

        '11 Log date
        params.Add("pLogDate", "", ORAPARM_OUTPUT)
        params("pLogDate").serverType = ORATYPE_DATE

        '12 Depth to top of split
        params.Add("pSplitDepthTop", 0, ORAPARM_OUTPUT)
        params("pSplitDepthTop").serverType = ORATYPE_NUMBER

        '13 Depth to bottom of split
        params.Add("pSplitDepthBot", 0, ORAPARM_OUTPUT)
        params("pSplitDepthBot").serverType = ORATYPE_NUMBER

        '14 Net weight
        params.Add("pNetWeight", 0, ORAPARM_OUTPUT)
        params("pNetWeight").serverType = ORATYPE_NUMBER

        '15 Matrix wet weight
        params.Add("pMtxWetWt", 0, ORAPARM_OUTPUT)
        params("pMtxWetWt").serverType = ORATYPE_NUMBER

        '16 Matrix dry weight
        params.Add("pMtxDryWt", 0, ORAPARM_OUTPUT)
        params("pMtxDryWt").serverType = ORATYPE_NUMBER

        '17 Matrix tare weight
        params.Add("pMtxTareWt", 0, ORAPARM_OUTPUT)
        params("pMtxTareWt").serverType = ORATYPE_NUMBER

        '18 Wet core to washer
        params.Add("pWetCoreWasher", 0, ORAPARM_OUTPUT)
        params("pWetCoreWasher").serverType = ORATYPE_NUMBER

        '19 Minutes mixed
        params.Add("pMinutesMixed", 0, ORAPARM_OUTPUT)
        params("pMinutesMixed").serverType = ORATYPE_NUMBER

        '20 +1/2 inch pebble dry pounds
        params.Add("pHalfPbDryLbs", 0, ORAPARM_OUTPUT)
        params("pHalfPbDryLbs").serverType = ORATYPE_NUMBER

        '21 +1/2 inch pebble BPL
        params.Add("pHalfPbBpl", 0, ORAPARM_OUTPUT)
        params("pHalfPbBpl").serverType = ORATYPE_NUMBER

        '22 +1/2 inch pebble Insol
        params.Add("pHalfPbInsol", 0, ORAPARM_OUTPUT)
        params("pHalfPbInsol").serverType = ORATYPE_NUMBER

        '23 +1/2 inch pebble Fe2O3
        params.Add("pHalfPbFe2O3", 0, ORAPARM_OUTPUT)
        params("pHalfPbFe2O3").serverType = ORATYPE_NUMBER

        '24 +1/2 inch pebble Al2O3
        params.Add("pHalfPbAl2O3", 0, ORAPARM_OUTPUT)
        params("pHalfPbAl2O3").serverType = ORATYPE_NUMBER

        '25 +1/2 inch pebble MgO
        params.Add("pHalfPbMgO", 0, ORAPARM_OUTPUT)
        params("pHalfPbMgO").serverType = ORATYPE_NUMBER

        '26 -1/2 inch +16 mesh pebble dry pounds
        params.Add("pP16mPbDryLbs", 0, ORAPARM_OUTPUT)
        params("pP16mPbDryLbs").serverType = ORATYPE_NUMBER

        '27 -1/2 inch +16 mesh pebble BPL
        params.Add("pP16mPbBpl", 0, ORAPARM_OUTPUT)
        params("pP16mPbBpl").serverType = ORATYPE_NUMBER

        '28 -1/2 inch +16 mesh pebble Insol
        params.Add("pP16mPbInsol", 0, ORAPARM_OUTPUT)
        params("pP16mPbInsol").serverType = ORATYPE_NUMBER

        '29 -1/2 inch +16 mesh pebble Fe2O3
        params.Add("pP16mPbFe2O3", 0, ORAPARM_OUTPUT)
        params("pP16mPbFe2O3").serverType = ORATYPE_NUMBER

        '30 -1/2 inch +16 mesh pebble Al2O3
        params.Add("pP16mPbAl2O3", 0, ORAPARM_OUTPUT)
        params("pP16mPbAl2O3").serverType = ORATYPE_NUMBER

        '31 -1/2 inch +16 mesh pebble MgO
        params.Add("pP16mPbMgO", 0, ORAPARM_OUTPUT)
        params("pP16mPbMgO").serverType = ORATYPE_NUMBER

        '32 -16 mesh +150 mesh feed wet pounds
        params.Add("pP150mFdWetLbs", 0, ORAPARM_OUTPUT)
        params("pP150mFdWetLbs").serverType = ORATYPE_NUMBER

        '33 -16 mesh +150 mesh feed BPL
        params.Add("pP150mFdBpl", 0, ORAPARM_OUTPUT)
        params("pP150mFdBpl").serverType = ORATYPE_NUMBER

        '34 -150 mesh waste clay BPL
        params.Add("pM150mWasteClayBpl", 0, ORAPARM_OUTPUT)
        params("pM150mWasteClayBpl").serverType = ORATYPE_NUMBER

        '35 Total feed or +35 mesh feed wet grams
        params.Add("pTorP35mFdWetGrams", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdWetGrams").serverType = ORATYPE_NUMBER

        '36 Total feed or +35 mesh feed dry grams
        params.Add("pTorP35mFdDryGrams", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdDryGrams").serverType = ORATYPE_NUMBER

        '37 Total feed or +35 mesh feed tare grams
        params.Add("pTorP35mFdTareGrams", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdTareGrams").serverType = ORATYPE_NUMBER

        '38 +35 mesh feed weight
        params.Add("pP35mFdWt", 0, ORAPARM_OUTPUT)
        params("pP35mFdWt").serverType = ORATYPE_NUMBER

        '39 +35 mesh feed BPL
        params.Add("pP35mFdBpl", 0, ORAPARM_OUTPUT)
        params("pP35mFdBpl").serverType = ORATYPE_NUMBER

        '40 -35 mesh feed weight
        params.Add("pM35mFdWt", 0, ORAPARM_OUTPUT)
        params("pM35mFdWt").serverType = ORATYPE_NUMBER

        '41 -35 mesh feed BPL
        params.Add("pM35mFdBpl", 0, ORAPARM_OUTPUT)
        params("pM35mFdBpl").serverType = ORATYPE_NUMBER

        '42 Total feed or +35m feed flotation sample wet grams
        params.Add("pTorP35mFdFlotSampWetGrams", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdFlotSampWetGrams").serverType = ORATYPE_NUMBER

        '43 Total feed or +35m feed rougher tailings weight
        params.Add("pTorP35mFdRghrTlngsWt", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdRghrTlngsWt").serverType = ORATYPE_NUMBER

        '44 Total feed or +35m feed rougher tailings BPL
        params.Add("pTorP35mFdRghrTlngsBpl", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdRghrTlngsBpl").serverType = ORATYPE_NUMBER

        '45 Total feed or +35m feed cleaner tailings weight
        params.Add("pTorP35mFdClnrTlngsWt", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdClnrTlngsWt").serverType = ORATYPE_NUMBER

        '46 Total feed or +35m feed cleaner tailings BPL
        params.Add("pTorP35mFdClnrTlngsBpl", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdClnrTlngsBpl").serverType = ORATYPE_NUMBER

        '47 Total feed or +35m feed amine concentrate weight
        params.Add("pTorP35mFdAmineCnWt", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdAmineCnWt").serverType = ORATYPE_NUMBER

        '48 Total feed or +35m feed amine concentrate BPL
        params.Add("pTorP35mFdAmineCnBpl", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdAmineCnBpl").serverType = ORATYPE_NUMBER

        '49 Total feed or +35m feed amine concentrate Insol
        params.Add("pTorP35mFdAmineCnInsol", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdAmineCnInsol").serverType = ORATYPE_NUMBER

        '50 Total feed or +35m feed amine concentrate Fe2O3
        params.Add("pTorP35mFdAmineCnFe2O3", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdAmineCnFe2O3").serverType = ORATYPE_NUMBER

        '51 Total feed or +35m feed amine concentrate Al2O3
        params.Add("pTorP35mFdAmineCnAl2O3", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdAmineCnAl2O3").serverType = ORATYPE_NUMBER

        '52 Total feed or +35m feed amine concentrate MgO
        params.Add("pTorP35mFdAmineCnMgO", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdAmineCnMgO").serverType = ORATYPE_NUMBER

        '53 Total feed or +35m feed amine concentrate color
        params.Add("pTorP35mFdAmineCnColor", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdAmineCnColor").serverType = ORATYPE_VARCHAR2

        '54 -35m feed amine concentrate weight
        params.Add("pM35mFdAmineCnWt", 0, ORAPARM_OUTPUT)
        params("pM35mFdAmineCnWt").serverType = ORATYPE_NUMBER

        '55 -35m feed amine concentrate BPL
        params.Add("pM35mFdAmineCnBpl", 0, ORAPARM_OUTPUT)
        params("pM35mFdAmineCnBpl").serverType = ORATYPE_NUMBER

        '56 -35m feed amine concentrate Insol
        params.Add("pM35mFdAmineCnInsol", 0, ORAPARM_OUTPUT)
        params("pM35mFdAmineCnInsol").serverType = ORATYPE_NUMBER

        '57 -35m feed amine concentrate Fe2O3
        params.Add("pM35mFdAmineCnFe2O3", 0, ORAPARM_OUTPUT)
        params("pM35mFdAmineCnFe2O3").serverType = ORATYPE_NUMBER

        '58 -35m feed amine concentrate Al2O3
        params.Add("pM35mFdAmineCnAl2O3", 0, ORAPARM_OUTPUT)
        params("pM35mFdAmineCnAl2O3").serverType = ORATYPE_NUMBER

        '59 -35m feed amine concentrate MgO
        params.Add("pM35mFdAmineCnMgO", 0, ORAPARM_OUTPUT)
        params("pM35mFdAmineCnMgO").serverType = ORATYPE_NUMBER

        '60 -35m feed amine concentrate color
        params.Add("pM35mFdAmineCnColor", 0, ORAPARM_OUTPUT)
        params("pM35mFdAmineCnColor").serverType = ORATYPE_VARCHAR2

        '61 -35m feed wet grams
        params.Add("pM35mFdWetGrams", 0, ORAPARM_OUTPUT)
        params("pM35mFdWetGrams").serverType = ORATYPE_NUMBER

        '62 -35m feed dry grams
        params.Add("pM35mFdDryGrams", 0, ORAPARM_OUTPUT)
        params("pM35mFdDryGrams").serverType = ORATYPE_NUMBER

        '63 -35m feed tare grams
        params.Add("pM35mFdTareGrams", 0, ORAPARM_OUTPUT)
        params("pM35mFdTareGrams").serverType = ORATYPE_NUMBER

        '64 -35m feed flotation sample wet grams
        params.Add("pM35mFdFlotSampWetGrams", 0, ORAPARM_OUTPUT)
        params("pM35mFdFlotSampWetGrams").serverType = ORATYPE_NUMBER

        '65 -35m feed rougher tailings weight
        params.Add("pM35mFdRghrTlngsWt", 0, ORAPARM_OUTPUT)
        params("pM35mFdRghrTlngsWt").serverType = ORATYPE_NUMBER

        '66 -35m feed rougher tailings BPL
        params.Add("pM35mFdRghrTlngsBpl", 0, ORAPARM_OUTPUT)
        params("pM35mFdRghrTlngsBpl").serverType = ORATYPE_NUMBER

        '67 -35m feed cleaner tailings weight
        params.Add("pM35mFdClnrTlngsWt", 0, ORAPARM_OUTPUT)
        params("pM35mFdClnrTlngsWt").serverType = ORATYPE_NUMBER

        '68 -35m feed cleaner tailings BPL
        params.Add("pM35mFdClnrTlngsBpl", 0, ORAPARM_OUTPUT)
        params("pM35mFdClnrTlngsBpl").serverType = ORATYPE_NUMBER

        '69 +1/2 inch pebble CaO
        params.Add("pHalfPbCaO", 0, ORAPARM_OUTPUT)
        params("pHalfPbCaO").serverType = ORATYPE_NUMBER

        '70 -1/2 inch +16 mesh pebble CaO
        params.Add("pP16mPbCaO", 0, ORAPARM_OUTPUT)
        params("pP16mPbCaO").serverType = ORATYPE_NUMBER

        '71 Total feed or +35m feed amine concentrate CaO
        params.Add("pTorP35mFdAmineCnCaO", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdAmineCnCaO").serverType = ORATYPE_NUMBER

        '72 -35m feed amine concentrate CaO
        params.Add("pM35mFdAmineCnCaO", 0, ORAPARM_OUTPUT)
        params("pM35mFdAmineCnCaO").serverType = ORATYPE_NUMBER

        '73 Date chem lab
        params.Add("pDateChemLab", "", ORAPARM_OUTPUT)
        params("pDateChemLab").serverType = ORATYPE_DATE

        '74 Who chem lab
        params.Add("pWhoChemLab", "", ORAPARM_OUTPUT)
        params("pWhoChemLab").serverType = ORATYPE_VARCHAR2

        '75 Record locked
        params.Add("pRecordLocked", 0, ORAPARM_OUTPUT)
        params("pRecordLocked").serverType = ORATYPE_NUMBER

        '76 Rerun status
        params.Add("pRerunStatus", 0, ORAPARM_OUTPUT)
        params("pRerunStatus").serverType = ORATYPE_NUMBER

        '77 Date rerun
        params.Add("pDateRerun", "", ORAPARM_OUTPUT)
        params("pDateRerun").serverType = ORATYPE_DATE

        '78 Met lab comment
        params.Add("pMetLabComment", "", ORAPARM_OUTPUT)
        params("pMetLabComment").serverType = ORATYPE_VARCHAR2

        '79 Chem lab comment
        params.Add("pChemLabComment", "", ORAPARM_OUTPUT)
        params("pChemLabComment").serverType = ORATYPE_VARCHAR2

        '----------

        '80 +1/2 inch pebble Cd
        params.Add("pHalfPbCd", 0, ORAPARM_OUTPUT)
        params("pHalfPbCd").serverType = ORATYPE_NUMBER

        '81 -1/2 inch +16 mesh pebble Cd
        params.Add("pP16mPbCd", 0, ORAPARM_OUTPUT)
        params("pP16mPbCd").serverType = ORATYPE_NUMBER

        '82 Total feed or +35m feed amine concentrate Cd
        params.Add("pTorP35mFdAmineCnCd", 0, ORAPARM_OUTPUT)
        params("pTorP35mFdAmineCnCd").serverType = ORATYPE_NUMBER

        '83 -35m feed amine concentrate Cd
        params.Add("pM35mFdAmineCnCd", 0, ORAPARM_OUTPUT)
        params("pM35mFdAmineCnCd").serverType = ORATYPE_NUMBER

        '84 Hardpan code
        params.Add("pHardpanCode", 0, ORAPARM_OUTPUT)
        params("pHardpanCode").serverType = ORATYPE_NUMBER

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospect.get_raw_prospect(:pSampleId," + _
                      ":pMineName, :pTownship, :pRange, :pSection, :pHoleLocation," + _
                      ":pSplit, :pSplitTotalNum, :pDrillDate, :pWashDate," + _
                      ":pLogDate, :pSplitDepthTop, :pSplitDepthBot, :pNetWeight, :pMtxWetWt," + _
                      ":pMtxDryWt, :pMtxTareWt, :pWetCoreWasher, :pMinutesMixed, :pHalfPbDryLbs," + _
                      ":pHalfPbBpl, :pHalfPbInsol, :pHalfPbFe2O3, :pHalfPbAl2O3, :pHalfPbMgO," + _
                      ":pP16mPbDryLbs, :pP16mPbBpl, :pP16mPbInsol, :pP16mPbFe2O3, :pP16mPbAl2O3," + _
                      ":pP16mPbMgO, :pP150mFdWetLbs, :pP150mFdBpl, :pM150mWasteClayBpl, :pTorP35mFdWetGrams," + _
                      ":pTorP35mFdDryGrams, :pTorP35mFdTareGrams, :pP35mFdWt, :pP35mFdBpl, :pM35mFdWt," + _
                      ":pM35mFdBpl, :pTorP35mFdFlotSampWetGrams, :pTorP35mFdRghrTlngsWt, :pTorP35mFdRghrTlngsBpl," + _
                      ":pTorP35mFdClnrTlngsWt, :pTorP35mFdClnrTlngsBpl, :pTorP35mFdAmineCnWt, :pTorP35mFdAmineCnBpl," + _
                      ":pTorP35mFdAmineCnInsol, :pTorP35mFdAmineCnFe2O3, :pTorP35mFdAmineCnAl2O3, :pTorP35mFdAmineCnMgO," + _
                      ":pTorP35mFdAmineCnColor, :pM35mFdAmineCnWt," + _
                      ":pM35mFdAmineCnBpl, :pM35mFdAmineCnInsol, :pM35mFdAmineCnFe2O3, :pM35mFdAmineCnAl2O3," + _
                      ":pM35mFdAmineCnMgO, :pM35mFdAmineCnColor, :pM35mFdWetGrams, :pM35mFdDryGrams," + _
                      ":pM35mFdTareGrams, :pM35mFdFlotSampWetGrams, :pM35mFdRghrTlngsWt, :pM35mFdRghrTlngsBpl," + _
                      ":pM35mFdClnrTlngsWt, :pM35mFdClnrTlngsBpl, :pHalfPbCaO, :pP16mPbCaO, :pTorP35mFdAmineCnCaO," + _
                      ":pM35mFdAmineCnCaO, :pDateChemLab, :pWhoChemLab, :pRecordLocked, :pRerunStatus," + _
                      ":pDateRerun, :pMetLabComment, :pChemLabComment," + _
                      ":pHalfPbCd, :pP16mPbCd, :pTorP35mFdAmineCnCd, :pM35mFdAmineCnCd, :pHardpanCode);end;", ORASQL_FAILEXEC)

        NumOfSplits = params("pSplitTotalNum").Value

        With aProspRawData
            .MineName = params("pMineName").Value
            .SampNum = params("pSampleId").Value

            .Section = params("pSection").Value
            .Township = params("pTownship").Value
            .Range = params("pRange").Value

            If Not IsDBNull(params("pHoleLocation").Value) Then
                .HoleLocation = params("pHoleLocation").Value
            Else
                .HoleLocation = ""
            End If

            .Split = params("pSplit").Value
            .SplitTotalNum = params("pSplitTotalNum").Value

            If Not IsDBNull(params("pDrillDate").Value) And _
                params("pDrillDate").Value <> "" And _
                params("pDrillDate").Value <> #1/4/1970# Then
                .DrillDate = params("pDrillDate").Value
            Else
                '12/31/8888 indicates a missing date
                .DrillDate = #12/31/8888#
            End If

            If Not IsDBNull(params("pWashDate").Value) And _
                params("pWashDate").Value <> "" And _
                params("pWashDate").Value <> #1/4/1970# Then
                .WashDate = params("pWashDate").Value
            Else
                '12/31/8888 indicates a missing date
                .WashDate = #12/31/8888#
            End If

            If Not IsDBNull(params("pLogDate").Value) And _
                params("pLogDate").Value <> "" And _
                params("pLogDate").Value <> #1/4/1970# Then
                .LogDate = params("pLogDate").Value
            Else
                '12/31/8888 indicates a missing date
                .LogDate = #12/31/8888#
            End If

            .SplitDepthTop = params("pSplitDepthTop").Value
            .SplitDepthBot = params("pSplitDepthBot").Value
            .NetWeight = params("pNetWeight").Value

            If Not IsDBNull(params("pMtxWetWt").Value) And _
                params("pMtxWetWt").Value <> "" Then
                .MtxWetWt = params("pMtxWetWt").Value
            Else
                .MtxWetWt = 0
            End If

            If Not IsDBNull(params("pMtxDryWt").Value) And _
                params("pMtxDryWt").Value <> "" Then
                .MtxDryWt = params("pMtxDryWt").Value
            Else
                .MtxDryWt = 0
            End If

            If Not IsDBNull(params("pMtxTareWt").Value) And _
                params("pMtxTareWt").Value <> "" Then
                .MtxTareWt = params("pMtxTareWt").Value
            Else
                .MtxTareWt = 0
            End If

            .WetCoreWasher = params("pWetCoreWasher").Value
            .MinutesMixed = params("pMinutesMixed").Value

            If Not IsDBNull(params("pHalfPbDryLbs").Value) And _
                params("pHalfPbDryLbs").Value <> "" Then
                .CrsPbDryLbs = params("pHalfPbDryLbs").Value
            Else
                .CrsPbDryLbs = 0
            End If

            If Not IsDBNull(params("pHalfPbBpl").Value) And _
                params("pHalfPbBpl").Value <> "" Then
                .CrsPbBpl = params("pHalfPbBpl").Value
            Else
                .CrsPbBpl = 0
            End If

            If Not IsDBNull(params("pHalfPbInsol").Value) And _
                params("pHalfPbInsol").Value <> "" Then
                .CrsPbIns = params("pHalfPbInsol").Value
            Else
                .CrsPbIns = 0
            End If

            If Not IsDBNull(params("pHalfPbFe2O3").Value) And _
                params("pHalfPbFe2O3").Value <> "" Then
                .CrsPbFe = params("pHalfPbFe2O3").Value
            Else
                .CrsPbFe = 0
            End If

            If Not IsDBNull(params("pHalfPbAl2O3").Value) And _
                params("pHalfPbAl2O3").Value <> "" Then
                .CrsPbAl = params("pHalfPbAl2O3").Value
            Else
                .CrsPbAl = 0
            End If

            If Not IsDBNull(params("pHalfPbMgO").Value) And _
                params("pHalfPbMgO").Value <> "" Then
                .CrsPbMg = params("pHalfPbMgO").Value
            Else
                .CrsPbMg = 0
            End If

            If Not IsDBNull(params("pHalfPbCaO").Value) And _
                params("pHalfPbCaO").Value <> "" Then
                .CrsPbCa = params("pHalfPbCaO").Value
            Else
                .CrsPbCa = 0
            End If

            If Not IsDBNull(params("pHalfPbCd").Value) And _
                params("pHalfPbCd").Value <> "" Then
                .CrsPbCd = params("pHalfPbCd").Value
            Else
                .CrsPbCd = 0
            End If

            If Not IsDBNull(params("pP16mPbDryLbs").Value) And _
                params("pP16mPbDryLbs").Value <> "" Then
                .FnePbDryLbs = params("pP16mPbDryLbs").Value
            Else
                .FnePbDryLbs = 0
            End If

            .FnePbBpl = params("pP16mPbBpl").Value
            .FnePbIns = params("pP16mPbInsol").Value
            .FnePbFe = params("pP16mPbFe2O3").Value
            .FnePbAl = params("pP16mPbAl2O3").Value
            .FnePbMg = params("pP16mPbMgO").Value
            .FnePbCa = params("pP16mPbCaO").Value
            .FnePbCd = params("pP16mPbCd").Value

            .FdM16P150WetLbs = params("pP150mFdWetLbs").Value
            .FdM16P150Bpl = params("pP150mFdBpl").Value
            .WasteClayBPL = params("pM150mWasteClayBpl").Value

            .FdP35WtGms = params("pP35mFdWt").Value
            .FdP35Bpl = params("pP35mFdBpl").Value
            .FdM35WtGms = params("pM35mFdWt").Value
            .FdM35Bpl = params("pM35mFdBpl").Value

            .FdWetGms = params("pTorP35mFdWetGrams").Value
            .FdDryGms = params("pTorP35mFdDryGrams").Value
            .FdTareGms = params("pTorP35mFdTareGrams").Value
            .FdFlotWetGms = params("pTorP35mFdFlotSampWetGrams").Value

            .FdRghrTailsWt = params("pTorP35mFdRghrTlngsWt").Value
            .FdRghrTailsBpl = params("pTorP35mFdRghrTlngsBpl").Value
            .FdClnrTailsWt = params("pTorP35mFdClnrTlngsWt").Value
            .FdClnrTailsBpl = params("pTorP35mFdClnrTlngsBpl").Value

            .FdAmineCnWt = params("pTorP35mFdAmineCnWt").Value
            .FdAmineCnBpl = params("pTorP35mFdAmineCnBpl").Value
            .FdAmineCnIns = params("pTorP35mFdAmineCnInsol").Value
            .FdAmineCnFe = params("pTorP35mFdAmineCnFe2O3").Value
            .FdAmineCnAl = params("pTorP35mFdAmineCnAl2O3").Value
            .FdAmineCnMg = params("pTorP35mFdAmineCnMgO").Value
            .FdAmineCnCa = params("pTorP35mFdAmineCnCaO").Value
            .FdAmineCnCd = params("pTorP35mFdAmineCnCd").Value
            If Not IsDBNull(params("pTorP35mFdAmineCnColor").Value) Then
                .FdAmineCnColor = params("pTorP35mFdAmineCnColor").Value
            Else
                .FdAmineCnColor = ""
            End If

            .FdM35WetGms = params("pM35mFdWetGrams").Value
            .FdM35DryGms = params("pM35mFdDryGrams").Value
            .FdM35TareGms = params("pM35mFdTareGrams").Value

            If Not IsDBNull(params("pM35mFdFlotSampWetGrams").Value) And _
                params("pM35mFdFlotSampWetGrams").Value <> "" Then
                .FdM35FlotWetGms = params("pM35mFdFlotSampWetGrams").Value
            Else
                .FdM35FlotWetGms = 0
            End If

            .FdM35RghrTailsWt = params("pM35mFdRghrTlngsWt").Value
            .FdM35RghrTailsBpl = params("pM35mFdRghrTlngsBpl").Value
            .FdM35ClnrTailsWt = params("pM35mFdClnrTlngsWt").Value
            .FdM35ClnrTailsBpl = params("pM35mFdClnrTlngsBpl").Value
            .FdM35AmineCnWt = params("pM35mFdAmineCnWt").Value
            .FdM35AmineCnBpl = params("pM35mFdAmineCnBpl").Value
            .FdM35AmineCnIns = params("pM35mFdAmineCnInsol").Value
            .FdM35AmineCnFe = params("pM35mFdAmineCnFe2O3").Value
            .FdM35AmineCnAl = params("pM35mFdAmineCnAl2O3").Value
            .FdM35AmineCnMg = params("pM35mFdAmineCnMgO").Value
            .FdM35AmineCnCa = params("pM35mFdAmineCnCaO").Value
            .FdM35AmineCnCd = params("pM35mFdAmineCnCd").Value

            If Not IsDBNull(params("pM35mFdAmineCnColor").Value) Then
                .FdM35AmineCnColor = params("pM35mFdAmineCnColor").Value
            Else
                .FdM35AmineCnColor = ""
            End If

            If Not IsDBNull(params("pDateChemLab").Value) And _
                params("pDateChemLab").Value <> "" And _
                params("pDateChemLab").Value <> #1/4/1970# Then
                .DateChemLab = params("pDateChemLab").Value
            Else
                '12/31/8888 indicates a missing date
                .DateChemLab = #12/31/8888#
            End If

            If Not IsDBNull(params("pWhoChemLab").Value) Then
                .WhoChemLab = params("pWhoChemLab").Value
            Else
                .WhoChemLab = ""
            End If

            .RecordLocked = params("pRecordLocked").Value
            .RerunStatus = params("pRerunStatus").Value

            If Not IsDBNull(params("pDateRerun").Value) And _
                params("pDateRerun").Value <> "" And _
                params("pDateRerun").Value <> #1/4/1970# Then
                .DateRerun = params("pDateRerun").Value
            Else
                '12/31/8888 indicates a missing date
                .DateRerun = #12/31/8888#
            End If

            If Not IsDBNull(params("pMetLabComment").Value) Then
                .MetLabComment = params("pMetLabComment").Value
            Else
                .MetLabComment = ""
            End If

            If Not IsDBNull(params("pChemLabComment").Value) Then
                .ChemLabComment = params("pChemLabComment").Value
            Else
                .ChemLabComment = ""
            End If

            .HardpanCode = params("pHardpanCode").Value
        End With

        ClearParams(params)
        gGetProspRawData = True

        Exit Function

gGetProspRawDataError:
        gGetProspRawData = False

        MsgBox("Error accessing raw prospect data." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Raw Prospect Data Access Error")

        On Error Resume Next
        ClearParams(params)
    End Function

    Private Sub ZeroProspRaw(ByRef aProspRawData As gProspRawDataType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        With aProspRawData
            .MineName = ""
            .SampNum = ""
            '-----
            .DrillDate = #12/31/8888#
            .WashDate = #12/31/8888#
            .LogDate = #12/31/8888#
            .HoleLocation = ""
            .Split = 0
            .Section = 0
            .Township = 0
            .Range = 0
            .SplitTotalNum = 0
            .SplitDepthTop = 0
            .SplitDepthBot = 0
            .NetWeight = 0
            .MtxWetWt = 0
            .MtxDryWt = 0
            .MtxTareWt = 0
            .WetCoreWasher = 0
            .MinutesMixed = 0
            '-----
            .CrsPbDryLbs = 0
            .CrsPbBpl = 0
            .CrsPbIns = 0
            .CrsPbFe = 0
            .CrsPbAl = 0
            .CrsPbMg = 0
            .CrsPbCa = 0
            .CrsPbCd = 0
            '-----
            .FnePbDryLbs = 0
            .FnePbBpl = 0
            .FnePbIns = 0
            .FnePbFe = 0
            .FnePbAl = 0
            .FnePbMg = 0
            .FnePbCa = 0
            .FnePbCd = 0
            '-----
            .FdM16P150WetLbs = 0
            .FdM16P150Bpl = 0
            '-----
            .WasteClayBPL = 0
            '-----
            .FdWetGms = 0
            .FdDryGms = 0
            .FdTareGms = 0
            .FdFlotWetGms = 0
            '-----
            .FdP35WtGms = 0
            .FdP35Bpl = 0
            .FdM35WtGms = 0
            .FdM35Bpl = 0
            '-----
            .FdRghrTailsWt = 0
            .FdRghrTailsBpl = 0
            '-----
            .FdClnrTailsWt = 0
            .FdClnrTailsBpl = 0
            '-----
            .FdAmineCnWt = 0
            .FdAmineCnBpl = 0
            .FdAmineCnIns = 0
            .FdAmineCnFe = 0
            .FdAmineCnAl = 0
            .FdAmineCnMg = 0
            .FdAmineCnCa = 0
            .FdAmineCnCd = 0
            .FdAmineCnColor = ""
            '-----
            .FdM35WetGms = 0
            .FdM35DryGms = 0
            .FdM35TareGms = 0
            .FdM35FlotWetGms = 0
            '-----
            .FdM35AmineCnWt = 0
            .FdM35AmineCnBpl = 0
            .FdM35AmineCnIns = 0
            .FdM35AmineCnFe = 0
            .FdM35AmineCnAl = 0
            .FdM35AmineCnMg = 0
            .FdM35AmineCnCa = 0
            .FdM35AmineCnCd = 0
            .FdM35AmineCnColor = ""
            '----
            .MetLabComment = ""
            .ChemLabComment = ""
            '----
            .FdM35RghrTailsWt = 0
            .FdM35RghrTailsBpl = 0
            '-----
            .FdM35ClnrTailsWt = 0
            .FdM35ClnrTailsBpl = 0
            '-----
            .HardpanCode = 0
        End With
    End Sub

    Public Function gSampNumExists(ByVal aSampleId As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer

        On Error GoTo SampNumExistsError

        gSampNumExists = False

        'Does this sample number exist?
        params = gDBParams

        params.Add("pSampleId", aSampleId, ORAPARM_INPUT)
        params("pSampleId").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", "", ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        'PROCEDURE sample_num_exists
        'pSampleId                  IN     VARCHAR2,
        'pResult                    IN OUT NUMBER
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospect.sample_num_exists(:pSampleId, " + _
                      ":pResult);end;", ORASQL_FAILEXEC)

        RecordCount = params("pResult").Value

        ClearParams(params)

        Select Case RecordCount
            Case Is = 0
                gSampNumExists = False

            Case Is = 1
                gSampNumExists = True

            Case Else
                gSampNumExists = False
        End Select

        Exit Function

SampNumExistsError:
        MsgBox("Error checking if sample# exists." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Check if Sample# Exists Error")

        On Error Resume Next
        gSampNumExists = False
        ClearParams(params)
    End Function

    Public Function gGetDrillHoleSamples(ByVal aSection As Integer, _
                                         ByVal aTownship As Integer, _
                                         ByVal aRange As Integer, _
                                         ByVal aHoleLocation As String, _
                                         ByVal aDrillHoleDate As String, _
                                         ByRef aSampleDynaset As OraDynaset) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        gGetDrillHoleSamples = False

        On Error GoTo gGetDrillHoleSamplesError

        params = gDBParams

        params.Add("pSection", aSection, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTownship, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRange, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHoleLocation, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pDrillDate", aDrillHoleDate, ORAPARM_INPUT)
        params("pDrillDate").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_all_loc_prosprawbase
        'pSection             IN     NUMBER,
        'pTownship            IN     NUMBER,
        'pRange               IN     NUMBER,
        'pHoleLocation        IN     VARCHAR2,
        'pDrillDate           IN     VARCHAR2,
        'pResult              IN OUT c_prosprawbase);

        'Note -- This proc will return all of the samples for the drill hole
        '        regardless of the drill date.  If the hole was redrilled and
        '        both sets of splits are in the database then there will be
        '        duplicates.

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospect.get_all_loc_prosprawbase(:pSection," + _
                      ":pTownship, :pRange, :pHoleLocation, :pDrillDate, " + _
                      ":pResult);end;", ORASQL_FAILEXEC)
        aSampleDynaset = params("pResult").Value
        ClearParams(params)

        gGetDrillHoleSamples = True

        Exit Function

gGetDrillHoleSamplesError:
        MsgBox("Error getting all sample#'s for this prospect hole." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "All Hole Sample#'s Access Error")

        On Error Resume Next
        ClearParams(params)
        gGetDrillHoleSamples = False
    End Function

    '    Public Function gDispIntervals(ByVal aSection As Integer, _
    '                                   ByVal aTownship As Integer, _
    '                                   ByVal aRange As Integer, _
    '                                   ByVal aHoleLocation As String, _
    '                                   ByVal aSplitNum As Integer, _
    '                                   ByVal aDrillDate As Date, _
    '                                   ByRef aDispSpread As vaSpread) As Boolean

    '        '**********************************************************************
    '        '
    '        '
    '        '
    '        '**********************************************************************

    '        On Error GoTo DisplayIntervalsError

    '        Dim SampleDynaset As OraDynaset
    '        Dim SplitThk As Single
    '        Dim CurrRow As Integer
    '        Dim MetComment As String
    '        Dim ChemComment As String
    '        Dim DisplayedSplit As Integer
    '        Dim RecordCount As Integer
    '        Dim SampsOk As Boolean

    '        DisplayedSplit = aSplitNum

    '        'Intervals will be displayed in ssInterval
    '        With aDispSpread
    '            .BlockMode = True
    '            .Row = 1
    '            .Row2 = .MaxRows
    '            .Col = 1
    '            .Col2 = .MaxCols
    '            .action = 12

    '            .BlockMode = False
    '        End With
    '        aDispSpread.MaxRows = 0

    '        SampsOk = gGetDrillHoleSamples(aSection, aTownship, _
    '                                       aRange, aHoleLocation, _
    '                                       CStr(aDrillDate), SampleDynaset)

    '        If SampsOk = False Then
    '            gDispIntervals = False
    '            Exit Function
    '        End If

    '        RecordCount = SampleDynaset.RecordCount

    '        If RecordCount = 0 Then
    '            gDispIntervals = False
    '            Exit Function
    '        Else
    '            gDispIntervals = True
    '        End If

    '        CurrRow = 0
    '        SampleDynaset.MoveFirst()

    '        Do While Not SampleDynaset.EOF
    '            With aDispSpread
    '                CurrRow = CurrRow + 1
    '                .MaxRows = .MaxRows + 1

    '                .Row = CurrRow

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
    '                .Text = SampleDynaset.Fields("drill_date").Value

    '                'Col6   Comments
    '                If Not IsDBNull(SampleDynaset.Fields("metlab_comment").Value) Then
    '                    MetComment = SampleDynaset.Fields("metlab_comment").Value
    '                Else
    '                    MetComment = ""
    '                End If
    '                If Not IsDBNull(SampleDynaset.Fields("chemlab_comment").Value) Then
    '                    ChemComment = SampleDynaset.Fields("chemlab_comment").Value
    '                Else
    '                    ChemComment = ""
    '                End If

    '                .Col = 6
    '                .Text = Trim(MetComment) + vbCrLf + Trim(ChemComment)

    '                If CurrRow = DisplayedSplit Then
    '                    .BlockMode = True
    '                    .Row = CurrRow
    '                    .Row2 = CurrRow
    '                    .Col = 1
    '                    .Col2 = .MaxCols
    '                    .BackColor = &HC0FFC0   'Light green
    '                    .BlockMode = False
    '                End If

    '                SampleDynaset.MoveNext()
    '            End With
    '        Loop

    '        SampleDynaset.Close()

    '        Exit Function

    'DisplayIntervalsError:
    '        MsgBox("Error getting all sample#'s for this hole." & vbCrLf & _
    '            Err.Description, _
    '            vbOKOnly + vbExclamation, _
    '            "All Hole Sample#'s Access Error")

    '        On Error Resume Next
    '        gDispIntervals = False
    '        SampleDynaset.Close()
    '    End Function

    Public Function gGetMetLabErrors(ByRef aProspRawData As gProspRawDataType, _
                                     ByRef aErrComms() As String, _
                                     ByVal aForMultiSampleRpt As Boolean, _
                                     ByVal aDispMtxWtsOk As Boolean, _
                                     ByVal aDispFdWtsOk As Boolean, _
                                     ByRef aSplitFeed As Boolean) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetMetLabErrorsError

        Dim MtxMoist As Single
        Dim FeedMoist1 As Single
        Dim FeedMoist2 As Single
        Dim MtxDry As Single
        Dim DryFdLbs As Single
        Dim TotDryLbs As Single
        Dim DryFdLbs2 As Single
        Dim TotDryLbs2 As Single
        Dim FeedDiff As Single
        Dim PctClay As Single
        Dim TotFdBpl As Single
        Dim CompFdBpl As Single
        Dim IntvProbs As Boolean
        Dim DateProbs As Boolean
        Dim CurrDate As Date
        Dim PrevDate As Date

        Dim Top As Single
        Dim Bottom As Single
        Dim RowIdx As Integer
        Dim MissingSplit As Boolean
        Dim DoubledSplit As Boolean
        Dim ZeroSplit As Boolean
        Dim TotalSplit As Boolean
        Dim PrevSplit As Integer
        Dim CurrSplit As Integer
        Dim ErrCount As Integer
        Dim IntervalStat As Boolean
        Dim SplitFeed As Boolean
        Dim AllSplits() As gHoleIntervalType

        ReDim aErrComms(50)

        If aForMultiSampleRpt = True Then
            For RowIdx = 1 To 50
                aErrComms(RowIdx) = "---"
            Next RowIdx
        Else
            For RowIdx = 1 To 50
                aErrComms(RowIdx) = ""
            Next RowIdx
        End If
        ErrCount = 0

        'For multi-sample report
        'Pb+Fd>DryMtx           1
        'Loss>150gms            2
        'DryTotTooHi            3
        'FdBpl>2                4
        'Hl#                    5
        'Missng Split           6
        'Dbld Split             7
        'Ftge Lgnth             8
        'Wst Cly Problem        9
        '+35 Fd Anlsys          10
        '-35 Fd Anlsys          11
        '-1/2+16 Anlsys         12
        '+1/2 Anlsys            13
        'Conc Anlsys            14
        'I&A>4 +1/2             15
        'I&A>4 -1/2+16          16
        'I&A>4 Conc             17
        'MgO>4 +1/2             18
        'MgO>4 -1/2+16          19
        'MgO>4 Conc             20
        '%Rcvry                 21
        'Split dates<>          22
        'CdProb +1/2            23  New 12/08/2003
        'CdProb -1/2+16         24  New 12/08/2003
        'CdProb Conc            25  New 12/08/2003
        'Not used               26
        'Not used               27
        'Not used               28
        'Not used               29
        'Not used               30
        '-----
        'Loss>150gms            31
        'DryTotTooHi            32
        'Dbld Split             33
        'Conc Anlsys            34
        'I&A>4 Conc             35
        'MgO>4 Conc             36
        '%Rcvry                 37
        'CdProb Conc            38  New 12/08/2003
        'Not used               39
        'Not used               40
        'Not used               41
        'Not used               42
        'Not used               43
        'Not used               44
        'Not used               45
        'Not used               46
        'Not used               47
        'Not used               48
        'Not used               49
        'Not used               50

        'Uses these functions/procedures in modRawProspect:
        '1) CalcPctClay
        '2) CalcTotFdBpl
        '3) CalcRcvry
        '4) gGetIntervals
        '5) CalcFeedMoist

        With aProspRawData
            SplitFeed = HasSplitFeed(aProspRawData)

            '----
            'These lines of code were in ReptMetLabErrors when it was in
            'frmProspectRawData
            ''.FdM35RghrTailsWt = 0
            ''.FdM35RghrTailsBpl = 0
            '-----
            ''.FdM35ClnrTailsWt = 0
            ''.FdM35ClnrTailsBpl = 0

            'Check matrix weights
            If .MtxWetWt - .MtxTareWt <> 0 And _
                .MtxDryWt > 0 Then
                MtxMoist = gRound((.MtxDryWt - .MtxTareWt) / _
                           (.MtxWetWt - .MtxTareWt), 4)
            Else
                MtxMoist = 0
            End If

            FeedMoist1 = 1 - (CalcFeedMoist(1, aProspRawData) / 100)
            FeedMoist2 = 1 - (CalcFeedMoist(2, aProspRawData) / 100)

            MtxDry = gRound(.WetCoreWasher * MtxMoist, 4)
            DryFdLbs = gRound(FeedMoist1 * .FdM16P150WetLbs, 4)
            TotDryLbs = .CrsPbDryLbs + .FnePbDryLbs + _
                        DryFdLbs

            If TotDryLbs >= MtxDry Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(1) = "MtxWt"       'Pb+Fd>DryMtx
                Else
                    aErrComms(ErrCount) = "Pebble + feed weight > dry matrix weight."
                End If
            Else
                'Matrix weight is OK
                If aDispMtxWtsOk = True Then
                    ErrCount = ErrCount + 1
                    aErrComms(ErrCount) = "Matrix weights are OK."
                End If
            End If

            If SplitFeed = False Then
                DryFdLbs2 = gRound(FeedMoist1 * .FdFlotWetGms, 4)
                TotDryLbs2 = .FdRghrTailsWt + .FdClnrTailsWt + _
                             .FdAmineCnWt
                FeedDiff = DryFdLbs2 - TotDryLbs2

                If FeedDiff > 150 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(2) = "FdLoss"      'Loss>150gms
                    Else
                        aErrComms(ErrCount) = "There is a loss greater " & _
                                              "than 150 grams (-16m +150m Fd)."
                    End If
                Else
                    If FeedDiff < 0 Then
                        ErrCount = ErrCount + 1
                        If aForMultiSampleRpt = True Then
                            aErrComms(3) = "DryTot"  'DryTotTooHi
                        Else
                            aErrComms(ErrCount) = "Dry total is too high (-16m +150m Fd), " & _
                            Format(DryFdLbs2, "#####0.0") & " vs " & _
                            Format(TotDryLbs2, "#####0.0")
                        End If
                    Else
                        'Feed weights are OK
                        If aDispFdWtsOk = True Then
                            ErrCount = ErrCount + 1
                            aErrComms(ErrCount) = "Feed weights are OK (-16m +150m Fd)."
                        End If
                    End If
                End If
            Else
                DryFdLbs2 = gRound(FeedMoist1 * .FdFlotWetGms, 4)
                TotDryLbs2 = .FdRghrTailsWt + .FdClnrTailsWt + _
                             .FdAmineCnWt
                FeedDiff = DryFdLbs2 - TotDryLbs2

                If FeedDiff > 150 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(2) = "FdLoss"      'Loss>150gms
                    Else
                        aErrComms(ErrCount) = "There is a loss greater than 150 grams (-16m +35m Fd)."
                    End If
                Else
                    If FeedDiff < 0 Then
                        ErrCount = ErrCount + 1
                        If aForMultiSampleRpt = True Then
                            aErrComms(3) = "DryTot"  'DryTotTooHi
                        Else
                            aErrComms(ErrCount) = "Dry total is too high (-16m +35m Fd), " & _
                                                  Format(DryFdLbs2, "#####0.0") & " vs " & _
                                                  Format(TotDryLbs2, "#####0.0")
                        End If
                    Else
                        'Feed weights are OK
                        If aDispFdWtsOk = True Then
                            ErrCount = ErrCount + 1
                            aErrComms(ErrCount) = "Feed weights are OK (-16m +35m Fd)."
                        End If
                    End If
                End If

                DryFdLbs2 = gRound(FeedMoist2 * .FdM35FlotWetGms, 4)
                TotDryLbs2 = .FdM35RghrTailsWt + .FdM35ClnrTailsWt + _
                             .FdM35AmineCnWt
                FeedDiff = DryFdLbs2 - TotDryLbs2

                If FeedDiff > 150 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(31) = "FdLoss"      'Loss>150gms
                    Else
                        aErrComms(ErrCount) = "There is a loss greater " & _
                                              "than 150 grams (-35m +150m Fd)."
                    End If
                Else
                    If FeedDiff < 0 Then
                        ErrCount = ErrCount + 1
                        If aForMultiSampleRpt = True Then
                            aErrComms(32) = "DryTot"  'DryTotTooHi
                        Else
                            aErrComms(ErrCount) = "Dry total is too high (-35m +150m Fd), " & _
                                                  Format(DryFdLbs2, "#####0.0") & " vs " & _
                                                  Format(TotDryLbs2, "#####0.0")
                        End If
                    Else
                        If aDispFdWtsOk = True Then
                            ErrCount = ErrCount + 1
                            aErrComms(ErrCount) = "Feed weights are OK (-35m +150m Fd)."
                        End If
                    End If
                End If
            End If

            'Check MgO's
            If .CrsPbMg > 4 Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(18) = "Mg1/2"
                Else
                    aErrComms(ErrCount) = "+1/2 Pebble MgO > 4.00"
                End If
            End If
            If .FnePbMg > 4 Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(19) = "Mg1/2+16"
                Else
                    aErrComms(ErrCount) = "-1/2 + 16m Pebble MgO > 4.00"
                End If
            End If

            If SplitFeed = False Then
                If .FdAmineCnMg > 4 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(20) = "MgCnc"
                    Else
                        aErrComms(ErrCount) = "Amine concentrate MgO (-16m +150m) > 4.00"
                    End If
                End If
            Else
                If .FdAmineCnMg > 4 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(20) = "MgCnc"
                    Else
                        aErrComms(ErrCount) = "Amine concentrate MgO (-16m +35m) > 4.00"
                    End If
                End If

                If .FdM35AmineCnMg > 4 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(36) = "MgCnc"
                    Else
                        aErrComms(ErrCount) = "Amine concentrate MgO (-35m +150m) > 4.00"
                    End If
                End If
            End If

            'Check Cd's  (07/14/04 -- Changed from 9)
            If .CrsPbCd > 15 Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(23) = "Cd1/2"
                Else
                    aErrComms(ErrCount) = "+1/2 Pebble Cd > 15.00"
                End If
            End If
            If .FnePbCd > 15 Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(24) = "Cd1/2+16"
                Else
                    aErrComms(ErrCount) = "-1/2 +16m Pebble Cd > 15.00"
                End If
            End If

            If SplitFeed = False Then
                If .FdAmineCnCd > 15 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(25) = "CdCnc"
                    Else
                        aErrComms(ErrCount) = "Amine concentrate Cd (-16m +150m) Cd > 15.00"
                    End If
                End If
            Else
                If .FdAmineCnCd > 15 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(25) = "CdCnc"
                    Else
                        aErrComms(ErrCount) = "Amine concentrate Cd (-16m +35m) Cd > 15.00"
                    End If
                End If

                If .FdM35AmineCnCd > 15 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(38) = "CdCnc"
                    Else
                        aErrComms(ErrCount) = "Amine concentrate Cd (-35m +150m) Cd > 15.00"
                    End If
                End If
            End If

            'Check Ca's
            'Only check coarse pebble CaO if there actually is
            'coarse pebble!
            If .CrsPbDryLbs <> 0 Then
                If (.CrsPbCa > 59 Or .CrsPbCa < 10) And .CrsPbCa <> 0 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(26) = "Ca1/2"
                    Else
                        aErrComms(ErrCount) = "+1/2 Pebble Ca > 59 or < 10"
                    End If
                End If
            End If
            If (.FnePbCa > 59 Or .FnePbCa < 10) And .FnePbCa <> 0 Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(27) = "Ca1/2+16"
                Else
                    aErrComms(ErrCount) = "-1/2 +16m Pebble Ca > 59 or < 10"
                End If
            End If

            If SplitFeed = False Then
                If (.FdAmineCnCa > 59 Or .FdAmineCnCa < 10) And .FdAmineCnCa <> 0 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(28) = "CaCnc"
                    Else
                        aErrComms(ErrCount) = "Amine concentrate (-16m +150m) Ca > 59 or < 10"
                    End If
                End If
            Else
                If (.FdAmineCnCa > 59 Or .FdAmineCnCa < 10) And .FdAmineCnCa <> 0 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(28) = "CaCnc"
                    Else
                        aErrComms(ErrCount) = "Amine concentrate (-16m + 35m) Ca > 59 or < 10"
                    End If
                End If

                If (.FdM35AmineCnCa > 59 Or .FdM35AmineCnCa < 10) And .FdM35AmineCnCa <> 0 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(39) = "CaCnc"
                    Else
                        aErrComms(ErrCount) = "Amine concentrate (-35m +150m) Ca > 59 or < 10"
                    End If
                End If
            End If

            'Check I&A's
            If .CrsPbAl + .CrsPbFe > 4 Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(15) = "IA1/2"          'I&A>4 +1/2
                Else
                    aErrComms(ErrCount) = "+1/2 Pebble I&A > 4.00"
                End If
            End If
            If .FnePbAl + .FnePbFe > 4 Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(16) = "IA1/2+16"       'I&A>4 -1/2+16
                Else
                    aErrComms(ErrCount) = "-1/2 + 16m Pebble I&A > 4.00"
                End If
            End If

            If SplitFeed = False Then
                If .FdAmineCnAl + .FdAmineCnFe > 4 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(17) = "IACnc"      'I&A>4 Conc
                    Else
                        aErrComms(ErrCount) = "Amine concentrate (-16m +150m) I&A > 4.00"
                    End If
                End If
            Else
                If .FdAmineCnAl + .FdAmineCnFe > 4 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(17) = "IACnc"      'I&A>4 Conc
                    Else
                        aErrComms(ErrCount) = "Amine concentrate (-16m +35m) I&A > 4.00"
                    End If
                End If

                If .FdM35AmineCnAl + .FdM35AmineCnFe > 4 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(35) = "IACnc"      'I&A>4 Conc
                    Else
                        aErrComms(ErrCount) = "Amine concentrate (-35m +150m) I&A > 4.00"
                    End If
                End If
            End If

            'Check waste clay
            PctClay = CalcPctClay(aProspRawData)
            If PctClay < 10 Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(9) = "Cly<10%"
                Else
                    aErrComms(ErrCount) = "Waste clay < 10%, " & _
                                          Format(PctClay, "##0.0") & "%"
                End If
            End If
            If PctClay > 85 Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(9) = "Cly>85%"
                Else
                    aErrComms(ErrCount) = "Waste clay > 85%, " & _
                                          Format(PctClay, "##0.0") & "%"
                End If
            End If

            'Check feed difference
            'gtmFdWetBpl

            TotFdBpl = .FdM16P150Bpl
            CompFdBpl = CalcTotFdBpl(aProspRawData)

            If TotFdBpl <> 0 Then
                If Abs(TotFdBpl - CompFdBpl) > 2 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(4) = "FdBPL"
                    Else
                        aErrComms(ErrCount) = "Feed BPL problem, " & _
                                              Format(TotFdBpl, "#0.0") & _
                                              " vs " + Format(CompFdBpl, "#0.0")
                    End If
                End If
            End If
            If TotFdBpl = 0 Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(4) = "FdBPL"
                Else
                    aErrComms(ErrCount) = "-16m +150m Feed BPL is missing."
                End If
            End If

            'Check for missing +35 feed bpl
            If .FdP35WtGms <> 0 And .FdP35Bpl = 0 Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(10) = "+35miss"
                Else
                    aErrComms(ErrCount) = "+35m Feed BPL missing."
                End If
            End If

            'Check for missing -35 feed bpl
            If .FdM35WtGms <> 0 And .FdM35Bpl = 0 Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(11) = "-35miss"
                Else
                    aErrComms(ErrCount) = "-35m Feed BPL missing."
                End If
            End If

            'fixed to here

            'Check for incomplete -1/2+16 Pebb chemical analysis    (fine pebble)
            'Note CaO is not needed for a complete pebble chemical analysis
            'Complete fine pebble chemical analysis includes:
            '1) BPL
            '2) Insol
            '3) Fe2O3
            '4) Al2O3
            '5) MgO
            '6) CaO
            '7) Cd

            '03/06/07, lss
            'Will not check for Cd -- Cd's are not run on prospect
            'anymore!
            If .FnePbDryLbs <> 0 And (.FnePbBpl = 0 Or _
                .FnePbIns = 0 Or .FnePbFe = 0 Or _
                .FnePbAl = 0 Or .FnePbMg = 0 Or _
                .FnePbCa = 0) Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(12) = "1/2+16miss"
                Else
                    aErrComms(ErrCount) = "-1/2 +16m Pebble -- " & _
                                          "incomplete chem analysis"
                End If
            End If

            'Check for incomplete +1/2 Pebb chemical analysis   (coarse pebble)
            'Complete coarse pebble chemical analysis includes:
            '1) BPL
            '2) Insol
            '3) Fe2O3
            '4) Al2O3
            '5) MgO
            '6) CaO
            '7) Cd
            '03/06/07, lss
            'Will not check for Cd -- Cd's are not run on prospect
            'anymore!
            If .CrsPbDryLbs <> 0 And (.CrsPbBpl = 0 Or _
                .CrsPbIns = 0 Or .CrsPbFe = 0 Or _
                .CrsPbAl = 0 Or .CrsPbMg = 0 Or _
                .CrsPbCa = 0) Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(13) = "1/2miss"
                Else
                    aErrComms(ErrCount) = "+1/2 Pebble -- " & _
                                          "incomplete chem analysis."
                End If
            End If

            'Check for incomplete concentrate chemical analysis
            '1) BPL
            '2) Insol
            '3) Fe2O3
            '4) Al2O3
            '5) MgO
            '6) CaO
            '7) Cd
            '03/06/07, lss
            'Will not check for Cd -- Cd's are not run on prospect
            'anymore!
            If SplitFeed = False Then
                If .FdAmineCnWt <> 0 And (.FdAmineCnBpl = 0 Or _
                    .FdAmineCnIns = 0 Or .FdAmineCnFe = 0 Or _
                    .FdAmineCnAl = 0 Or .FdAmineCnMg = 0 Or _
                    .FdAmineCnCa = 0) Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(14) = "CncMiss"
                    Else
                        aErrComms(ErrCount) = "Amine concentrate (-16m +150m) -- " & _
                                              "incomplete chem analysis."
                    End If
                End If
            Else
                If .FdAmineCnWt <> 0 And (.FdAmineCnBpl = 0 Or _
                    .FdAmineCnIns = 0 Or .FdAmineCnFe = 0 Or _
                    .FdAmineCnAl = 0 Or .FdAmineCnMg = 0 Or _
                    .FdAmineCnCa = 0) Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(14) = "CncMiss"
                    Else
                        aErrComms(ErrCount) = "Amine concentrate (-16m +35m) -- " & _
                                              "incomplete chem analysis."
                    End If
                End If

                If .FdM35AmineCnWt <> 0 And (.FdM35AmineCnBpl = 0 Or _
                    .FdM35AmineCnIns = 0 Or .FdM35AmineCnFe = 0 Or _
                    .FdM35AmineCnAl = 0 Or .FdM35AmineCnMg = 0 Or _
                    .FdM35AmineCnCa = 0) Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(34) = "CncMiss"
                    Else
                        aErrComms(ErrCount) = "Amine concentrate (-35m +150m) -- " & _
                                              "incomplete chem analysis."
                    End If
                End If
            End If

            'Check for high %recovery problems
            If SplitFeed = False Then
                If CalcRcvry(1, SplitFeed, aProspRawData) > 95 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(21) = ">95%"
                    Else
                        aErrComms(ErrCount) = "%Recovery (-16m +150m Fd) > 95%."
                    End If
                End If
            Else
                If CalcRcvry(1, SplitFeed, aProspRawData) > 95 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(21) = ">95%"
                    Else
                        aErrComms(ErrCount) = "%Recovery (-16m +35m Fd) > 95%."
                    End If
                End If

                If CalcRcvry(2, SplitFeed, aProspRawData) > 95 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(37) = ">95%"
                    Else
                        aErrComms(ErrCount) = "%Recovery (-35m +150m Fd) > 95%."
                    End If
                End If
            End If

            'Check for low %recovery problems (test #1 -- 75%)
            If SplitFeed = False Then
                If CalcRcvry(1, SplitFeed, aProspRawData) < 75 Then
                    If .FdM16P150Bpl >= 10 Then
                        ErrCount = ErrCount + 1
                        If aForMultiSampleRpt = True Then
                            aErrComms(21) = "<75%"
                        Else
                            aErrComms(ErrCount) = "%Recovery (-16m +150m Fd) < 75% " & _
                                                  "(+10% feed BPL cutoff)."
                        End If
                    End If
                End If
            Else
                If CalcRcvry(1, SplitFeed, aProspRawData) < 75 Then
                    If .FdP35Bpl >= 10 Then
                        ErrCount = ErrCount + 1
                        If aForMultiSampleRpt = True Then
                            aErrComms(21) = "<75%"
                        Else
                            aErrComms(ErrCount) = "%Recovery (-16m +35m Fd) < 75% " & _
                                                  "(+10% feed BPL cutoff)."
                        End If
                    End If
                End If

                If CalcRcvry(2, SplitFeed, aProspRawData) < 75 Then
                    If .FdM35Bpl >= 10 Then
                        ErrCount = ErrCount + 1
                        If aForMultiSampleRpt = True Then
                            aErrComms(37) = "<75%"
                        Else
                            aErrComms(ErrCount) = "%Recovery (-35m +150m Fd) < 75% " & _
                                                  "(+10% feed BPL cutoff)."
                        End If
                    End If
                End If
            End If

            'Check for low %recovery problems (test #2 -- 40%)
            If SplitFeed = False Then
                If CalcRcvry(1, SplitFeed, aProspRawData) < 40 And _
                    CalcRcvry(1, SplitFeed, aProspRawData) <> 0 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(21) = "<40%"
                    Else
                        aErrComms(ErrCount) = "%Recovery (-16m +150m Fd) < 40%"
                    End If
                End If
            Else
                If CalcRcvry(1, SplitFeed, aProspRawData) < 40 And _
                    CalcRcvry(1, SplitFeed, aProspRawData) <> 0 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(21) = "<40%"
                    Else
                        aErrComms(ErrCount) = "%Recovery (-16m +35m Fd) < 40%"
                    End If
                End If

                If CalcRcvry(2, SplitFeed, aProspRawData) < 40 And _
                    CalcRcvry(2, SplitFeed, aProspRawData) <> 0 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        aErrComms(37) = "<40%"
                    Else
                        aErrComms(ErrCount) = "%Recovery (-35m +150m Fd) < 40%"
                    End If
                End If
            End If

            If SplitFeed = False Then
                'Check for low concentrate BPL
                If .FdAmineCnBpl < 50 And .FdAmineCnBpl <> 0 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        'Not included in report
                    Else
                        aErrComms(ErrCount) = "Conc BPL (-16m +150m Fd) < 50."
                    End If
                End If

                'Check for low concentrate Insol
                If .FdAmineCnIns < 2 And .FdAmineCnIns <> 0 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        'Not included in report
                    Else
                        aErrComms(ErrCount) = "Conc Insol (-16m +150m Fd) < 2.0"
                    End If
                End If
            Else
                'Check for low concentrate BPL
                If .FdAmineCnBpl < 50 And .FdAmineCnBpl <> 0 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        'Not included in report
                    Else
                        aErrComms(ErrCount) = "Conc BPL (-16m +35m Fd) < 50."
                    End If
                End If

                'Check for low concentrate Insol
                If .FdAmineCnIns < 2 And .FdAmineCnIns <> 0 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        'Not included in report
                    Else
                        aErrComms(ErrCount) = "Conc Insol (-16m +35m Fd) < 2.0"
                    End If
                End If

                'Check for low concentrate BPL
                If .FdM35AmineCnBpl < 50 And .FdM35AmineCnBpl <> 0 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        'Not included in report
                    Else
                        aErrComms(ErrCount) = "Conc BPL (-35m +150m Fd) < 50"
                    End If
                End If

                'Check for low concentrate Insol
                If .FdM35AmineCnIns < 2 And .FdM35AmineCnIns <> 0 Then
                    ErrCount = ErrCount + 1
                    If aForMultiSampleRpt = True Then
                        'Not included in report
                    Else
                        aErrComms(ErrCount) = "Conc Insol (-35m +150m Fd) < 2.0"
                    End If
                End If
            End If

            'Check for interval footage problems
            IntervalStat = gGetIntervals(.Section, .Township, _
                                         .Range, .HoleLocation, .Split, _
                                         .DrillDate, AllSplits)

            'Check for drill date problems -- should be the same for all splits for a given
            'prospect hole
            DateProbs = False
            PrevDate = AllSplits(1).DrillDate
            For RowIdx = 2 To UBound(AllSplits)
                If Not IsDBNull(AllSplits(RowIdx).DrillDate) Then
                    CurrDate = AllSplits(RowIdx).DrillDate
                Else
                    CurrDate = #12/31/8888#
                End If
                If CurrDate <> PrevDate Then
                    DateProbs = True
                End If
                PrevDate = CurrDate
            Next RowIdx

            If DateProbs = True Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(22) = "Split Dates<>"
                Else
                    aErrComms(ErrCount) = "Drill date problems for this hole"
                End If
            End If

            'fAllSplits     Row 1   Top of seam depth
            '               Row 2   Bottom of seam depth
            '               Row 3   Sample#
            '               Row 4   Drill date
            '               Row 5   Split#

            IntvProbs = False
            Top = AllSplits(1).TosDepth
            Bottom = AllSplits(1).BosDepth

            If Bottom <= Top Then
                IntvProbs = True
            End If

            If IntvProbs = False Then
                For RowIdx = 2 To UBound(AllSplits)
                    Top = AllSplits(RowIdx).TosDepth

                    If Top <> Bottom Then
                        IntvProbs = True
                        Exit For
                    End If

                    Bottom = AllSplits(RowIdx).BosDepth

                    If Bottom <= Top Then
                        IntvProbs = True
                        Exit For
                    End If
                Next RowIdx
            End If

            If IntvProbs = True Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(8) = "FtLgh"
                Else
                    aErrComms(ErrCount) = "Interval footage problems for this hole."
                End If
            End If

            'Check for missing splits (& doubled splits)
            MissingSplit = False
            DoubledSplit = False
            ZeroSplit = False
            TotalSplit = False

            For RowIdx = 1 To UBound(AllSplits)
                If AllSplits(RowIdx).Split = 0 Then
                    ZeroSplit = True
                End If
            Next RowIdx

            PrevSplit = AllSplits(1).Split
            For RowIdx = 2 To UBound(AllSplits)
                CurrSplit = AllSplits(RowIdx).Split

                If CurrSplit <> PrevSplit + 1 Then
                    MissingSplit = True
                    Exit For
                End If

                PrevSplit = CurrSplit
            Next RowIdx

            PrevSplit = AllSplits(1).Split
            For RowIdx = 2 To UBound(AllSplits)
                CurrSplit = AllSplits(RowIdx).Split

                If CurrSplit = PrevSplit Then
                    DoubledSplit = True
                    Exit For
                End If

                PrevSplit = CurrSplit
            Next RowIdx

            If UBound(AllSplits) <> aProspRawData.SplitTotalNum Then
                TotalSplit = True
            End If

            If MissingSplit = True Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(6) = "MsngSplit"
                Else
                    aErrComms(ErrCount) = "Split is missing for this hole."
                End If
            End If
            If DoubledSplit = True Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(7) = "DblSplit"
                Else
                    aErrComms(ErrCount) = "Split is used 2X for this hole."
                End If
            End If
            If ZeroSplit = True Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    'Not included in report
                Else
                    aErrComms(ErrCount) = "Split with zero value for this hole."
                End If
            End If
            If TotalSplit = True Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(33) = "TotSplit"
                Else
                    aErrComms(ErrCount) = "Total split error."
                End If
            End If

            'Check hole numbers.
            If Trim(.HoleLocation) = "" Then
                ErrCount = ErrCount + 1
                If aForMultiSampleRpt = True Then
                    aErrComms(5) = "Hl#"
                Else
                    aErrComms(ErrCount) = "Hole# missing."
                End If
            End If
        End With

        gGetMetLabErrors = ErrCount
        aSplitFeed = SplitFeed

        Exit Function

gGetMetLabErrorsError:
        MsgBox("Error getting raw prospect errors." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Raw Prospect Error")

        On Error Resume Next
        gGetMetLabErrors = 0
        aSplitFeed = False
    End Function

    Private Function CalcFeedMoist(ByVal aFeedType As Integer, _
                                   ByRef aProspRawData As gProspRawDataType) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo CalcFeedMoistError

        Dim PctMoist As Single
        PctMoist = 0

        With aProspRawData
            Select Case aFeedType
                Case Is = 1
                    If .FdWetGms - .FdTareGms > 0 Then
                        PctMoist = Round((.FdDryGms - .FdTareGms) / _
                                   (.FdWetGms - .FdTareGms), 4)
                    Else
                        PctMoist = 0
                    End If

                Case Is = 2
                    If .FdM35WetGms - .FdM35TareGms > 0 Then
                        PctMoist = Round((.FdM35DryGms - .FdM35TareGms) / _
                                   (.FdM35WetGms - .FdM35TareGms), 4)
                    Else
                        PctMoist = 0
                    End If

            End Select
        End With

        CalcFeedMoist = 100 - Round(100 * PctMoist, 1)

        Exit Function

CalcFeedMoistError:
        MsgBox("Error calculating feed %moisture." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Calculate Feed %Moisture Error")
    End Function

    Private Function CalcPctClay(ByRef aProspRawData As gProspRawDataType) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo CalcPctClayError

        Dim PctDry As Single
        Dim DryFdLbs As Single
        Dim TotDryLbs As Single
        Dim PctDry2 As Single
        Dim DryCrewShr As Single
        Dim DiffLb As Single

        With aProspRawData
            If .FdWetGms - .FdTareGms <> 0 Then
                PctDry = gRound((.FdDryGms - .FdTareGms) / _
                         (.FdWetGms - .FdTareGms), 4)
            Else
                PctDry = 0
            End If

            DryFdLbs = PctDry * .FdM16P150WetLbs
            TotDryLbs = DryFdLbs + .CrsPbDryLbs + .FnePbDryLbs

            If .MtxWetWt - .MtxTareWt <> 0 Then
                PctDry2 = gRound((.MtxDryWt - .MtxTareWt) / _
                         (.MtxWetWt - .MtxTareWt), 4)
            Else
                PctDry2 = 0
            End If

            DryCrewShr = PctDry2 * .WetCoreWasher
            DiffLb = DryCrewShr - TotDryLbs
            If DiffLb < 0 Then
                DiffLb = 0
            End If

            If DryCrewShr <> 0 Then
                CalcPctClay = gRound(DiffLb / DryCrewShr * 100, 1)
            Else
                CalcPctClay = 0
            End If
        End With

        Exit Function

CalcPctClayError:
        MsgBox("Error calculating %clay." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Calculate %Clay Error")
    End Function

    Private Function CalcTotFdBpl(ByRef aProspRawData As gProspRawDataType) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo CalcTotFdBplError

        'fProspRawData.FdP35WtGms
        'fProspRawData.FdP35Bpl
        'fProspRawData.FdM35WtGms
        'fProspRawData.FdM35Bpl

        With aProspRawData
            If .FdP35WtGms + .FdM35WtGms <> 0 Then
                CalcTotFdBpl = gRound((.FdP35WtGms * .FdP35Bpl + _
                                   .FdM35WtGms * .FdM35Bpl) / _
                                   (.FdP35WtGms + .FdM35WtGms), 1)
            Else
                CalcTotFdBpl = 0
            End If
        End With

        Exit Function

CalcTotFdBplError:
        MsgBox("Error calculating total feed BPL." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Calculate Total Feed BPL Error")
    End Function

    Private Function CalcRcvry(ByVal aFeedType As Integer, _
                               ByVal aSplitFeed As Boolean, _
                               ByRef aProspRawData As gProspRawDataType) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo CalcRcvryError

        Dim PctRcvry As Single
        Dim PctDry As Single
        Dim FlotDryGrams As Single

        PctRcvry = 0

        'gtmFdWetGms.Value    gtmFdWetGms2.Value
        'gtmFdDryGms.Value    gtmFdDryGms2.Value
        'gtmFdTareGms.Value   gtmFdTareGms2.Value

        With aProspRawData
            Select Case aFeedType
                Case Is = 1
                    If .FdWetGms - .FdTareGms <> 0 Then
                        PctDry = gRound((.FdDryGms - .FdTareGms) / _
                                 (.FdWetGms - .FdTareGms), 4)
                    Else
                        PctDry = 0
                    End If

                    FlotDryGrams = gRound(PctDry * .FdFlotWetGms, 0)

                    If aSplitFeed = False Then
                        If FlotDryGrams * .FdM16P150Bpl <> 0 Then
                            PctRcvry = gRound((.FdAmineCnWt * .FdAmineCnBpl) / _
                                       (FlotDryGrams * .FdM16P150Bpl) * 100, 1)
                        Else
                            PctRcvry = 0
                        End If
                    Else
                        If FlotDryGrams * .FdP35Bpl <> 0 Then
                            PctRcvry = gRound((.FdAmineCnWt * .FdAmineCnBpl) / _
                                       (FlotDryGrams * .FdP35Bpl) * 100, 1)
                        Else
                            PctRcvry = 0
                        End If
                    End If

                Case Is = 2
                    If .FdM35WetGms - .FdM35TareGms <> 0 Then
                        PctDry = gRound((.FdM35DryGms - .FdM35TareGms) / _
                                 (.FdM35WetGms - .FdM35TareGms), 4)
                    Else
                        PctDry = 0
                    End If

                    FlotDryGrams = gRound(PctDry * .FdM35FlotWetGms, 0)

                    If FlotDryGrams * .FdM35Bpl <> 0 Then
                        PctRcvry = gRound((.FdM35AmineCnWt * .FdM35AmineCnBpl) / _
                                   (FlotDryGrams * .FdM35Bpl) * 100, 1)
                    Else
                        PctRcvry = 0
                    End If

            End Select
        End With

        CalcRcvry = PctRcvry

        Exit Function

CalcRcvryError:
        MsgBox("Error calculating concentrate recovery." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Calculate Concentrate Recovery Error")
    End Function

    Public Function gGetIntervals(ByVal aSection As Integer, _
                                  ByVal aTownship As Integer, _
                                  ByVal aRange As Integer, _
                                  ByVal aHoleLocation As String, _
                                  ByVal aSplitNum As Integer, _
                                  ByVal aDrillDate As Date, _
                                  ByRef aAllSplits() As gHoleIntervalType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetIntervalsError

        Dim SampleDynaset As OraDynaset
        Dim SplitCount As Integer
        Dim RowIdx As Integer
        Dim RecordCount As Integer
        Dim SampsOk As Boolean

        SampsOk = gGetDrillHoleSamples(aSection, aTownship, _
                                       aRange, aHoleLocation, _
                                       CStr(aDrillDate), SampleDynaset)

        If SampsOk = False Then
            gGetIntervals = False
            Exit Function
        End If

        RecordCount = SampleDynaset.RecordCount

        If RecordCount = 0 Then
            gGetIntervals = False
            Exit Function
        Else
            gGetIntervals = True
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
            aAllSplits(SplitCount).DrillDate = SampleDynaset.Fields("drill_date").Value
            aAllSplits(SplitCount).Split = SampleDynaset.Fields("split").Value

            SampleDynaset.MoveNext()
        Loop

        SampleDynaset.Close()

        Exit Function

gGetIntervalsError:
        MsgBox("Error getting all sample#'s for this hole." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "All Hole Sample#'s Access Error")

        On Error Resume Next
        gGetIntervals = False
        SampleDynaset.Close()
    End Function

    Private Function HasSplitFeed(ByRef aProspRawData As gProspRawDataType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        With aProspRawData
            If .FdM35WetGms <> 0 Or .FdM35DryGms <> 0 Or .FdM35TareGms <> 0 Or _
                .FdM35FlotWetGms <> 0 Or .FdM35AmineCnWt <> 0 Or _
                .FdM35AmineCnBpl <> 0 Or .FdM35AmineCnIns <> 0 Or _
                .FdM35AmineCnFe <> 0 Or .FdM35AmineCnAl <> 0 Or _
                .FdM35AmineCnMg <> 0 Or .FdM35AmineCnCa <> 0 Or _
                .FdM35AmineCnCd <> 0 Or .FdM35AmineCnCd <> 0 Or _
                .FdM35AmineCnColor <> "" Then
                HasSplitFeed = True
            Else
                HasSplitFeed = False
            End If
        End With
    End Function

    Public Function gGetCodeDesc(ByVal aCode As String, _
                                 ByVal aMinusLimit As String, _
                                 ByVal aPlusLimit As String, _
                                 ByVal aDetailed As Boolean) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ThisMinusLimit As String
        Dim ThisPlusLimit As String
        Dim ThisCode As String

        ThisMinusLimit = aMinusLimit
        ThisPlusLimit = aPlusLimit
        ThisCode = aCode
        gGetCodeDesc = ""

        If StrConv(ThisMinusLimit, vbUpperCase) = "NONE" Or _
            StrConv(ThisMinusLimit, vbUpperCase) = "ROUGHER" Or _
            StrConv(ThisMinusLimit, vbUpperCase) = "AMINE" Then
            ThisMinusLimit = ""
        Else
            ThisMinusLimit = "-" & ThisMinusLimit
        End If

        If StrConv(ThisPlusLimit, vbUpperCase) = "NONE" Or _
            StrConv(ThisPlusLimit, vbUpperCase) = "ROUGHER" Or _
            StrConv(ThisPlusLimit, vbUpperCase) = "AMINE" Then
            ThisPlusLimit = ""
        Else
            ThisPlusLimit = "+" & ThisPlusLimit
        End If

        If aDetailed = True Then
            gGetCodeDesc = ThisCode & "  (" & Trim(ThisMinusLimit & _
                           " " & ThisPlusLimit) & ")"
        Else
            gGetCodeDesc = ThisMinusLimit & " " & ThisPlusLimit
        End If
    End Function

    Public Function gCodeFromCodeDesc(ByVal aCodeDesc As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ParentPos As Integer

        gCodeFromCodeDesc = ""
        ParentPos = InStr(aCodeDesc, "(")

        If ParentPos <> 0 Then
            gCodeFromCodeDesc = Trim(Mid(aCodeDesc, 1, ParentPos - 1))
        Else
            gCodeFromCodeDesc = ""
        End If
    End Function

    Public Function gGetOldRawProspect(ByVal aTwp As Integer, _
                                       ByVal aRge As Integer, _
                                       ByVal aSec As Integer, _
                                       ByVal aHole As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetOldRawProspectError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RawProspDynaset As OraDynaset
        Dim ThisDrillDate As Date
        Dim RecordCount As Integer
        Dim ThisSampleId As String
        Dim ThisSplit As Integer

        gGetOldRawProspect = ""

        params = gDBParams

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pHoleLocation", aHole, ORAPARM_INPUT)
        params("pHoleLocation").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_prosp_raw_base_spec
        'pTownship               IN     NUMBER,
        'pRange                  IN     NUMBER,
        'pSection                IN     NUMBER,
        'pHoleLocation           IN     VARCHAR2,
        'pResult                 IN OUT c_prospraw)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_raw_prospect.get_prosp_raw_base_spec(:pTownship, " + _
                      ":pRange, :pSection, :pHoleLocation, :pResult);end;", ORASQL_FAILEXEC)
        RawProspDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = RawProspDynaset.RecordCount

        RawProspDynaset.MoveFirst()

        Do While Not RawProspDynaset.EOF
            ThisDrillDate = RawProspDynaset.Fields("drill_date").Value
            ThisSampleId = RawProspDynaset.Fields("sample_id").Value
            ThisSplit = RawProspDynaset.Fields("split").Value

            If ThisSplit = 1 Then
                gGetOldRawProspect = gGetOldRawProspect & Format(ThisDrillDate, "mm/dd/yy") & _
                                     "-" & Mid(ThisSampleId, 3) & " "
            End If

            RawProspDynaset.MoveNext()
        Loop

        RawProspDynaset.Close()

        gGetOldRawProspect = Trim(gGetOldRawProspect)

        Exit Function

GetOldRawProspectError:
        MsgBox("Error getting old raw prospect data." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Get Error")

        On Error Resume Next
        ClearParams(params)
    End Function


End Module
