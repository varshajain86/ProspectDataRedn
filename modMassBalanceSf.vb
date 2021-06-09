Option Explicit On
Imports System.Math
Imports OracleInProcServer
Module modMassBalanceSf
    'Attribute VB_Name = "modMassBalanceSf"
    '**********************************************************************
    'South Fort Meade Mass Balance Module
    '
    'Special Comments
    '----------------
    'This module handles mass balance and metallurgical reports for
    'South Fort Meade only -- gMassBalanceSf and gMetallurgicalSf.
    '
    'Types in this module:
    'Private Type MassBalanceSfRptType
    'Dim MassBalanceSfRpt As MassBalanceSfRptType
    '
    'Private Type MetallurgicalSfRptType
    'Dim MetallurgicalSfRpt As MetallurgicalSfRptType
    '
    'Private Type MassBalanceSfShiftType
    'Dim MbSfShift As MassBalanceSfShiftType
    '
    'Private Type MassBalanceSfTotalType
    'Dim MbSfTotal As MassBalanceSfTotalType
    '
    'Procedures/Functions in this module:
    '1) gMassBalanceSf
    '2) gMetallurgicalSf
    '3) gGetSfFloatPlantBalanceData
    '4) ZeroSfSummingData
    '5) ProcessSfMassBalanceData
    '6) ProcessSfMassBalanceTotals
    '7) gAdjustedFeedTonsSf
    '8) gGetMetReagentDataSf

    '**********************************************************************
    '   Maintenance Log
    '
    '   10/14/2004, lss
    '       Set up this module -- moved functionality from modFloatPlant.
    '       See the maintenance log in modFloatPlant for any changes made
    '       to this code prior to 10/14/2004.
    '   12/15/2004, lss
    '       Added rounding for BPL's in gMetallurgicalSF -- extra digits
    '       for small number (ex. .05834349, 5.834349E-02) was causing
    '       problems in the Crystal report.
    '   05/23/2005, lss
    '       Modified for SR to CA transfer.
    '   03/14/2006, lss
    '       Added Public Function gFltPltRcvrySF.
    '   10/31/2011, lss
    '       Wasn't handling #1FA, #2FA, #3FA concentrate BPL's correctly.
    '       All changes marked with 10/31/2011!
    '
    '**********************************************************************


    'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
    'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
    'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade

    Public gCcnTonsEst As Boolean

    Private Enum mSfFloatPlantCircRowEnum
        sfCrN1FineRghr = 1
        sfCrN2FineRghr = 2
        sfCrN3FineRghr = 3
        sfCrN4FineRghr = 4
        sfCrN5FineRghr = 5
        sfCrSwingRghr = 6
        sfCrCoarseRghr = 7
        sfCrUltraCoarseRghr = 8
        sfCrTotalRghr = 9
        sfCrN1FineAmine = 10
        sfCrN2FineAmine = 11
        sfCrN3FineAmine = 12
        sfCrCoarseAmine = 13
        sfCrTotalClnr = 14
        sfCrGrandTotal = 15
        sfCrFcnProduct = 16
        sfCrCcnProduct = 17
        sfCrUccnProduct = 18
        sfCrCnProduct = 19
    End Enum

    Private Enum mSfFloatPlantCircColEnum
        sfCcOperHrs = 1
        sfCcFdTonsRpt = 2
        sfCcFdTonsAdj = 3
        sfCcFdBpl = 4
        sfCcCnBpl = 5
        sfCcTlBpl = 6
        sfCcRc = 7
        sfCcPctActRcvry = 8
        sfCcPctStdRcvry = 9
        sfCcCnTonsAdj = 10
        sfCcTlTonsAdj = 11
        sfCcFdTph = 12
        sfCcCnTph = 13
        sfCcTlTph = 14
    End Enum

    Private Enum mSfFloatPlantGmtRowEnum
        sfGrAsReportedGmtBpl = 1
        sfGrCalculatedGmtBpl = 2
        sfGrReportedFdTons = 3
        sfGrGmtBplFromCircuits = 4
    End Enum

    Private Enum mSfFloatPlantGmtColEnum
        sfGcFdTons = 1
        sfGcCnTons = 2
        sfGcFdBpl = 3
        sfGcCnBpl = 4
        sfGcTlBpl = 5
        sfGcRc = 6
        sfGcPctRcvry = 7
    End Enum

    Private Enum mSfFloatPlantShiftColEnum
        eN1frOperHours = 1
        eN1frFdTonsRpt = 2
        eN1frFdBpl = 3
        eN1frCnBpl = 4
        eN1frTlBpl = 5

        eN2FrOperHours = 6
        eN2FrFdTonsRpt = 7
        eN2FrFdBpl = 8
        eN2FrCnBpl = 9
        eN2FrTlBpl = 10

        eN3FrOperHours = 11
        eN3FrFdTonsRpt = 12
        eN3FrFdBpl = 13
        eN3FrCnBpl = 14
        eN3FrTlBpl = 15

        eN4FrOperHours = 16
        eN4FrFdTonsRpt = 17
        eN4FrFdBpl = 18
        eN4FrCnBpl = 19
        eN4FrTlBpl = 20

        eN5FrOperHours = 21
        eN5FrFdTonsRpt = 22
        eN5FrFdBpl = 23
        eN5FrCnBpl = 24
        eN5FrTlBpl = 25

        eCrOperHours = 26
        eCrFdTonsRpt = 27
        eCrFdBpl = 28
        eCrCnBpl = 29
        eCrTlBpl = 30

        eSrOperHours = 31
        eSrFdTonsRpt = 32
        eSrFdBpl = 33
        eSrCnBpl = 34
        eSrTlBpl = 35

        eUcrOperHours = 36
        eUcrFdTonsRpt = 37
        eUcrFdBpl = 38
        eUcrCnBpl = 39
        eUcrTlBpl = 40

        eN1FaOperHours = 41
        eN1FaCnTons = 42
        eN1FaCnBpl = 43
        eN1FaFdBpl = 44
        eN1FaTlBpl = 45

        eN2FaOperHours = 46
        eN2FaCnTons = 47
        eN2FaCnBpl = 48
        eN2FaFdBpl = 49
        eN2FaTlBpl = 50

        eN3FaOperHours = 46
        eN3FaCnTons = 47
        eN3FaCnBpl = 48
        eN3FaFdBpl = 49
        eN3FaTlBpl = 50
    End Enum

    Public Structure MassBalanceSfRptType
        Public PrdPbTons As Long
        Public PrdPbBpl As Single
        Public PrdCnTons As Long
        Public PrdCnBpl As Single
        Public PrdTotTons As Long
        Public PrdTotBpl As Single
        Public PrdFcnTons As Long
        Public PrdFcnBpl As Single
        Public PrdCcnTons As Long
        Public PrdCcnBpl As Single
        Public PrdUccnTons As Long
        Public PrdUccnBpl As Single

        Public N1frHrs As Single
        Public N2frHrs As Single
        Public N3frHrs As Single
        Public N4frHrs As Single
        Public N5frHrs As Single
        Public SrHrs As Single
        Public CrHrs As Single
        Public UcrHrs As Single
        Public N1faHrs As Single
        Public N2faHrs As Single
        Public N3faHrs As Single
        Public CaHrs As Single

        Public N1frFdTonsRpt As Long
        Public N2frFdTonsRpt As Long
        Public N3frFdTonsRpt As Long
        Public N4frFdTonsRpt As Long
        Public N5frFdTonsRpt As Long
        Public SrFdTonsRpt As Long
        Public CrFdTonsRpt As Long
        Public UcrFdTonsRpt As Long
        Public N1faFdTonsRpt As Long
        Public N2faFdTonsRpt As Long
        Public N3faFdTonsRpt As Long
        Public CaFdTonsRpt As Long

        Public N1frFdTonsAdj As Long
        Public N2frFdTonsAdj As Long
        Public N3frFdTonsAdj As Long
        Public N4frFdTonsAdj As Long
        Public N5frFdTonsAdj As Long
        Public SrFdTonsAdj As Long
        Public CrFdTonsAdj As Long
        Public UcrFdTonsAdj As Long
        Public N1faFdTonsAdj As Long
        Public N2faFdTonsAdj As Long
        Public N3faFdTonsAdj As Long
        Public CaFdTonsAdj As Long

        Public N1frFdBpl As Single
        Public N2frFdBpl As Single
        Public N3frFdBpl As Single
        Public N4frFdBpl As Single
        Public N5frFdBpl As Single
        Public SrFdBpl As Single
        Public CrFdBpl As Single
        Public UcrFdBpl As Single
        Public N1faFdBpl As Single
        Public N2faFdBpl As Single
        Public N3faFdBpl As Single
        Public CaFdBpl As Single

        Public N1frCnBpl As Single
        Public N2frCnBpl As Single
        Public N3frCnBpl As Single
        Public N4frCnBpl As Single
        Public N5frCnBpl As Single
        Public SrCnBpl As Single
        Public CrCnBpl As Single
        Public UcrCnBpl As Single
        Public N1faCnBpl As Single
        Public N2faCnBpl As Single
        Public N3faCnBpl As Single
        Public CaCnBpl As Single

        Public N1frTlBpl As Single
        Public N2frTlBpl As Single
        Public N3frTlBpl As Single
        Public N4frTlBpl As Single
        Public N5frTlBpl As Single
        Public SrTlBpl As Single
        Public CrTlBpl As Single
        Public UcrTlBpl As Single
        Public N1faTlBpl As Single
        Public N2faTlBpl As Single
        Public N3faTlBpl As Single
        Public CaTlBpl As Single

        Public N1frRc As Single
        Public N2frRc As Single
        Public N3frRc As Single
        Public N4frRc As Single
        Public N5frRc As Single
        Public SrRc As Single
        Public CrRc As Single
        Public UcrRc As Single
        Public N1faRc As Single
        Public N2faRc As Single
        Public N3faRc As Single
        Public CaRc As Single

        Public N1frAr As Single
        Public N2frAr As Single
        Public N3frAr As Single
        Public N4frAr As Single
        Public N5frAr As Single
        Public SrAr As Single
        Public CrAr As Single
        Public UcrAr As Single
        Public N1faAr As Single
        Public N2faAr As Single
        Public N3faAr As Single
        Public CaAr As Single

        Public N1frSr As Single
        Public N2frSr As Single
        Public N3frSr As Single
        Public N4frSr As Single
        Public N5frSr As Single
        Public SrSr As Single
        Public CrSr As Single
        Public UcrSr As Single
        Public N1faSr As Single
        Public N2faSr As Single
        Public N3faSr As Single
        Public CaSr As Single

        Public N1frCnTonsAdj As Long
        Public N2FrCnTonsAdj As Long
        Public N3FrCnTonsAdj As Long
        Public N4FrCnTonsAdj As Long
        Public N5FrCnTonsAdj As Long
        Public SrCnTonsAdj As Long
        Public CrCnTonsAdj As Long
        Public UcrCnTonsAdj As Long
        Public N1faCnTonsAdj As Long
        Public N2faCnTonsAdj As Long
        Public N3faCnTonsAdj As Long
        Public CaCnTonsAdj As Long

        Public N1frTlTonsAdj As Long
        Public N2frTlTonsAdj As Long
        Public N3frTlTonsAdj As Long
        Public N4frTlTonsAdj As Long
        Public N5frTlTonsAdj As Long
        Public SrTlTonsAdj As Long
        Public CrTlTonsAdj As Long
        Public UcrTlTonsAdj As Long
        Public N1faTlTonsAdj As Long
        Public N2faTlTonsAdj As Long
        Public N3faTlTonsAdj As Long
        Public CaTlTonsAdj As Long

        Public SumRghrFdBpl As Single
        Public SumRghrFdTonsRpt As Long
        Public SumRghrFdTonsAdj As Long
        Public SumRghrCnTonsAdj As Long
        Public SumRghrTlTonsAdj As Long
        Public SumClnrFdTonsRpt As Long
        Public SumClnrFdTonsAdj As Long
        Public SumClnrCnTonsAdj As Long
        Public SumClnrTlTonsAdj As Long
        Public SumAllFdTonsAdj As Long
        Public SumAllCnTonsAdj As Long
        Public SumAllTlTonsAdj As Long

        Public ArGmtFdBpl As Single
        Public ArGmtCnBpl As Single
        Public ArGmtTlBpl As Single
        Public ArGmtRc As Single
        Public ArGmtRcvry As Single

        Public ClcGmtFdTons As Long
        Public ClcGmtCnTons As Long
        Public ClcGmtFdBpl As Single
        Public ClcGmtCnBpl As Single
        Public ClcGmtTlBpl As Single
        Public ClcGmtRc As Single
        Public ClcGmtRcvry As Single

        Public ArFdTonsFdTons As Long
        Public ArFdTonsCnTons As Long
        Public ArFdTonsFdBpl As Single
        Public ArFdTonsCnBpl As Single
        Public ArFdTonsTlBpl As Single
        Public ArFdTonsRc As Single
        Public ArFdTonsRcvry As Single

        Public GmtFromCircs As Single

        Public CcnTonsEst As Boolean
    End Structure
    Dim MassBalanceSfRpt As MassBalanceSfRptType

    Private Structure MetallurgicalSfRptType
        Public N1frFdBpl As Single
        Public N1frFdTons As Long
        Public N1frCnBpl As Single
        Public N1frCnTons As Long
        Public N1frTlBpl As Single
        Public N1frTlTons As Long
        Public N1frRc As Single
        Public N1frRcvry As Single

        Public N2frFdBpl As Single
        Public N2frFdTons As Long
        Public N2frCnBpl As Single
        Public N2frCnTons As Long
        Public N2frTlBpl As Single
        Public N2frTlTons As Long
        Public N2frRc As Single
        Public N2frRcvry As Single

        Public N3frFdBpl As Single
        Public N3frFdTons As Long
        Public N3frCnBpl As Single
        Public N3frCnTons As Long
        Public N3frTlBpl As Single
        Public N3frTlTons As Long
        Public N3frRc As Single
        Public N3frRcvry As Single

        Public N4frFdBpl As Single
        Public N4frFdTons As Long
        Public N4frCnBpl As Single
        Public N4frCnTons As Long
        Public N4frTlBpl As Single
        Public N4frTlTons As Long
        Public N4frRc As Single
        Public N4frRcvry As Single

        Public N5frFdBpl As Single
        Public N5frFdTons As Long
        Public N5frCnBpl As Single
        Public N5frCnTons As Long
        Public N5frTlBpl As Single
        Public N5frTlTons As Long
        Public N5frRc As Single
        Public N5frRcvry As Single

        Public FrFdBplAvg As Single
        Public FrFdTonsAvg As Long
        Public FrCnBplAvg As Single
        Public FrCnTonsAvg As Long
        Public FrTlBplAvg As Single
        Public FrTlTonsAvg As Long
        Public FrRcAvg As Single
        Public FrRcvryAvg As Single

        Public SrFdBpl As Single
        Public SrFdTons As Long
        Public SrCnBpl As Single
        Public SrCnTons As Long
        Public SrTlBpl As Single
        Public SrTlTons As Long
        Public SrRc As Single
        Public SrRcvry As Single

        Public CrFdBpl As Single
        Public CrFdTons As Long
        Public CrCnBpl As Single
        Public CrCnTons As Long
        Public CrTlBpl As Single
        Public CrTlTons As Long
        Public CrRc As Single
        Public CrRcvry As Single

        Public CrFdBplAvg As Single
        Public CrFdTonsAvg As Long
        Public CrCnBplAvg As Single
        Public CrCnTonsAvg As Long
        Public CrTlBplAvg As Single
        Public CrTlTonsAvg As Long
        Public CrRcAvg As Single
        Public CrRcvryAvg As Single

        Public FrSrCrFdBplAvg As Single
        Public FrSrCrFdTonsAvg As Long
        Public FrSrCrCnBplAvg As Single
        Public FrSrCrCnTonsAvg As Long
        Public FrSrCrTlBplAvg As Single
        Public FrSrCrTlTonsAvg As Long
        Public FrSrCrRcAvg As Single
        Public FrSrCrRcvryAvg As Single

        Public N1faFdBpl As Single
        Public N1faFdTons As Long
        Public N1faCnBpl As Single
        Public N1faCnTons As Long
        Public N1faTlBpl As Single
        Public N1faTlTons As Long
        Public N1faRc As Single
        Public N1faRcvry As Single

        Public N2faFdBpl As Single
        Public N2faFdTons As Long
        Public N2faCnBpl As Single
        Public N2faCnTons As Long
        Public N2faTlBpl As Single
        Public N2faTlTons As Long
        Public N2faRc As Single
        Public N2faRcvry As Single

        Public N3faFdBpl As Single
        Public N3faFdTons As Long
        Public N3faCnBpl As Single
        Public N3faCnTons As Long
        Public N3faTlBpl As Single
        Public N3faTlTons As Long
        Public N3faRc As Single
        Public N3faRcvry As Single

        Public FaFdBplAvg As Single
        Public FaFdTonsAvg As Long
        Public FaCnBplAvg As Single
        Public FaCnTonsAvg As Long
        Public FaTlBplAvg As Single
        Public FaTlTonsAvg As Long
        Public FaRcAvg As Single
        Public FaRcvryAvg As Single

        Public CaFdBpl As Single
        Public CaFdTons As Long
        Public CaCnBpl As Single
        Public CaCnTons As Long
        Public CaTlBpl As Single
        Public CaTlTons As Long
        Public CaRc As Single
        Public CaRcvry As Single

        Public ClnrFdBplAvg As Single
        Public ClnrFdTonsAvg As Long
        Public ClnrCnBplAvg As Single
        Public ClnrCnTonsAvg As Long
        Public ClnrTlBplAvg As Single
        Public ClnrTlTonsAvg As Long
        Public ClnrRcAvg As Single
        Public ClnrRcvryAvg As Single

        Public PltFdBplAvg As Single
        Public PltFdTonsAvg As Long
        Public PltCnBplAvg As Single
        Public PltCnTonsAvg As Long
        Public PltTlBplAvg As Single
        Public PltTlTonsAvg As Long
        Public PltRcAvg As Single
        Public PltRcvryAvg As Single

        Public UcrFdBpl As Single
        Public UcrFdTons As Long
        Public UcrCnBpl As Single
        Public UcrCnTons As Long
        Public UcrTlBpl As Single
        Public UcrTlTons As Long
        Public UcrRc As Single
        Public UcrRcvry As Single

        Public CombFdBplAvg As Single
        Public CombFdTonsAvg As Long
        Public CombCnBplAvg As Single
        Public CombCnTonsAvg As Long
        Public CombTlBplAvg As Single
        Public CombTlTonsAvg As Long
        Public CombRcAvg As Single
        Public CombRcvryAvg As Single

        Public ReportedGmtBpl As Single

        Public RgSuTotUnits As Long
        Public RgAmTotUnits As Long
        Public RgSaTotUnits As Long
        Public RgSoTotUnits As Long
        Public RgFaTotUnits As Long
        Public RgFoTotUnits As Long
        Public RgDeTotUnits As Long
        Public RgAllTotUnits As Long

        Public RgSuTotCost As Long
        Public RgAmTotCost As Long
        Public RgSaTotCost As Long
        Public RgSoTotCost As Long
        Public RgFaTotCost As Long
        Public RgFoTotCost As Long
        Public RgDeTotCost As Long
        Public RgAllTotCost As Long

        Public RgSuAdjFdDpt As Single
        Public RgAmAdjFdDpt As Single
        Public RgSaAdjFdDpt As Single
        Public RgSoAdjFdDpt As Single
        Public RgFaAdjFdDpt As Single
        Public RgFoAdjFdDpt As Single
        Public RgDeAdjFdDpt As Single
        Public RgAllAdjFdDpt As Single

        Public RgSuRptFdDpt As Single
        Public RgAmRptFdDpt As Single
        Public RgSaRptFdDpt As Single
        Public RgSoRptFdDpt As Single
        Public RgFaRptFdDpt As Single
        Public RgFoRptFdDpt As Single
        Public RgDeRptFdDpt As Single
        Public RgAllRptFdDpt As Single

        Public RgSuCnDpt As Single
        Public RgAmCnDpt As Single
        Public RgSaCnDpt As Single
        Public RgSoCnDpt As Single
        Public RgFaCnDpt As Single
        Public RgFoCnDpt As Single
        Public RgDeCnDpt As Single
        Public RgAllCnDpt As Single

        Public RgSuAdjFdUpt As Single
        Public RgAmAdjFdUpt As Single
        Public RgSaAdjFdUpt As Single
        Public RgSoAdjFdUpt As Single
        Public RgFaAdjFdUpt As Single
        Public RgFoAdjFdUpt As Single
        Public RgDeAdjFdUpt As Single
        Public RgAllAdjFdUpt As Single

        Public RgSuRptFdUpt As Single
        Public RgAmRptFdUpt As Single
        Public RgSaRptFdUpt As Single
        Public RgSoRptFdUpt As Single
        Public RgFaRptFdUpt As Single
        Public RgFoRptFdUpt As Single
        Public RgDeRptFdUpt As Single
        Public RgAllRptFdUpt As Single

        Public RgSuCnUpt As Single
        Public RgAmCnUpt As Single
        Public RgSaCnUpt As Single
        Public RgSoCnUpt As Single
        Public RgFaCnUpt As Single
        Public RgFoCnUpt As Single
        Public RgDeCnUpt As Single
        Public RgAllCnUpt As Single

        Public RgTotRptFdTons As Long
        Public RgTotCnTons As Long
        Public RgTotAdjFdTons As Long

        Public CcnTonsEst As Boolean
    End Structure
    Dim MetallurgicalSfRpt As MetallurgicalSfRptType

    Private Structure MassBalanceSfShiftType
        Public PrdFcnTons As Long
        Public PrdFcnBpl As Single

        Public PrdCcnTons As Long
        Public PrdCcnBpl As Single

        Public PrdUccnTons As Long
        Public PrdUccnBpl As Single

        Public PrdTotCnTons As Long

        Public Tr1FrTo1Fa As Single
        Public Tr1FrTo2Fa As Single
        Public Tr2FrTo2Fa As Single
        Public Tr2FrTo3Fa As Single
        Public Tr3FrTo2Fa As Single
        Public Tr3FrTo3Fa As Single
        Public Tr4FrTo1Fa As Single
        Public Tr4FrTo2Fa As Single
        Public Tr5FrTo1Fa As Single
        Public Tr5FrTo2Fa As Single
        Public TrSrTo1Fa As Single
        Public TrSrTo2Fa As Single
        Public TrSrTo3Fa As Single
        Public TrCrToCa As Single
        Public TrCrTo1Fa As Single

        Public N1frHrs As Single
        Public N2frHrs As Single
        Public N3frHrs As Single
        Public N4frHrs As Single
        Public N5frHrs As Single
        Public SrHrs As Single
        Public CrHrs As Single
        Public UcrHrs As Single

        Public N1frFdBpl As Single
        Public N2frFdBpl As Single
        Public N3frFdBpl As Single
        Public N4frFdBpl As Single
        Public N5frFdBpl As Single
        Public SrFdBpl As Single
        Public CrFdBpl As Single
        Public UcrFdBpl As Single
        Public N1faFdBpl As Single
        Public N2faFdBpl As Single
        Public N3faFdBpl As Single
        Public CaFdBpl As Single

        Public N1frCnBpl As Single
        Public N2frCnBpl As Single
        Public N3frCnBpl As Single
        Public N4frCnBpl As Single
        Public N5frCnBpl As Single
        Public SrCnBpl As Single
        Public CrCnBpl As Single
        Public UcrCnBpl As Single
        Public N1faCnBpl As Single
        Public N2faCnBpl As Single
        Public N3faCnBpl As Single

        Public N1frTlBpl As Single
        Public N2frTlBpl As Single
        Public N3frTlBpl As Single
        Public N4frTlBpl As Single
        Public N5frTlBpl As Single
        Public SrTlBpl As Single
        Public CrTlBpl As Single
        Public UcrTlBpl As Single
        Public N1faTlBpl As Single
        Public N2faTlBpl As Single
        Public N3faTlBpl As Single
        Public CaTlBpl As Single
        Public GmtBpl As Single

        Public N1frRc As Single
        Public N2frRc As Single
        Public N3frRc As Single
        Public N4frRc As Single
        Public N5frRc As Single
        Public SrRc As Single
        Public CrRc As Single
        Public UcrRc As Single

        Public N1frFdTonsRpt As Long
        Public N2frFdTonsRpt As Long
        Public N3frFdTonsRpt As Long
        Public N4frFdTonsRpt As Long
        Public N5frFdTonsRpt As Long
        Public SrFdTonsRpt As Long
        Public CrFdTonsRpt As Long
        Public UcrFdTonsRpt As Long

        Public N1frFdTonsRptW As Long    'Tons with BPL
        Public N2frFdTonsRptW As Long
        Public N3frFdTonsRptW As Long
        Public N4frFdTonsRptW As Long
        Public N5frFdTonsRptW As Long
        Public SrFdTonsRptW As Long
        Public CrFdTonsRptW As Long
        Public UcrFdTonsRptW As Long

        Public N1frFdBtRpt As Double     'Tons X BPL  (BPL Tons)
        Public N2frFdBtRpt As Double
        Public N3frFdBtRpt As Double
        Public N4frFdBtRpt As Double
        Public N5frFdBtRpt As Double
        Public SrFdBtRpt As Double
        Public CrFdBtRpt As Double
        Public UcrFdBtRpt As Double

        Public N1frCnTonsExp As Long     'Concentrate tons expected
        Public N2frCnTonsExp As Long     'based on ratio of concentration
        Public N3frCnTonsExp As Long
        Public N4frCnTonsExp As Long
        Public N5frCnTonsExp As Long
        Public SrCnTonsExp As Long
        Public CrCnTonsExp As Long
        Public UcrCnTonsExp As Long

        Public N1faFdTonsRpt As Long
        Public N1faFdTonsAdj As Long
        Public N1faFdBt As Double
        Public N1faRc As Single
        Public N1faCnTonsExp As Long
        Public N1faPct3 As Single
        Public N1faCnTons As Long

        Public N2faFdTonsRpt As Long
        Public N2faFdTonsAdj As Long
        Public N2faFdBt As Double
        Public N2faRc As Single
        Public N2faCnTonsExp As Long
        Public N2faPct3 As Single
        Public N2faCnTons As Long

        Public N3faFdTonsRpt As Long
        Public N3faFdTonsAdj As Long
        Public N3faFdBt As Double
        Public N3faRc As Single
        Public N3faCnTonsExp As Long
        Public N3faPct3 As Single
        Public N3faCnTons As Long

        Public N1faCnBt As Double      '10/31/2011, lss  New
        Public N2faCnBt As Double      '10/31/2011, lss  New
        Public N3faCnBt As Double      '10/31/2011, lss  New

        Public N123faCnTonsExp As Long

        Public CaFdTonsRpt As Long
        Public CaFdTonsAdj As Long
        Public CaFdBt As Double
        Public CaRc As Double
        Public CaCnTons As Long

        Public N1frCnBt As Double
        Public N2frCnBt As Double
        Public N3frCnBt As Double
        Public N4frCnBt As Double
        Public N5frCnBt As Double
        Public CrCnBt As Double
        Public SrCnBt As Double

        Public N1FaN1frBt As Double
        Public N1FaN4frBt As Double
        Public N1FaN5frBt As Double
        Public N1FaSrBt As Double
        Public N1FaCrBt As Double

        Public N1FaTotBt As Double

        Public N1FaN1frBtPct As Double
        Public N1FaN4frBtPct As Double
        Public N1FaN5frBtPct As Double
        Public N1FaSrBtPct As Double
        Public N1FaCrBtPct As Double

        Public N2FaN1frBt As Double
        Public N2FaN2frBt As Double
        Public N2FaN3frBt As Double
        Public N2FaN4frBt As Double
        Public N2FaN5frBt As Double
        Public N2FaSrBt As Double

        Public N2FaTotBt As Double

        Public N2FaN1frBtPct As Double
        Public N2FaN2frBtPct As Double
        Public N2FaN3frBtPct As Double
        Public N2FaN4frBtPct As Double
        Public N2FaN5frBtPct As Double
        Public N2FaSrBtPct As Double

        Public N3FaN2frBt As Double
        Public N3FaN3frBt As Double
        Public N3FaSrBt As Double

        Public N3FaTotBt As Double

        Public N3FaN2frBtPct As Double
        Public N3FaN3frBtPct As Double
        Public N3FaSrBtPct As Double

        '---------- 01/23/2005, lss
        Public CaSrBt As Double
        Public CaCrBt As Double
        Public CaTotBt As Double
        Public CaSrBtPct As Double
        Public CaCrBtPct As Double
        '----------

        Public N1frCnTonsAdj As Long
        Public N2FrCnTonsAdj As Long
        Public N3FrCnTonsAdj As Long
        Public N4FrCnTonsAdj As Long
        Public N5FrCnTonsAdj As Long
        Public CrCnTonsAdj As Long
        Public SrCnTonsAdj As Long
        Public UcrCnTonsAdj As Long

        Public N1frCnBtAdj As Double
        Public N2FrCnBtAdj As Double
        Public N3FrCnBtAdj As Double
        Public N4FrCnBtAdj As Double
        Public N5FrCnBtAdj As Double
        Public CrCnBtAdj As Double
        Public SrCnBtAdj As Double

        Public N1faTlTonsAdj As Long
        Public N2faTlTonsAdj As Long
        Public N3faTlTonsAdj As Long
        Public CaTlTonsAdj As Long

        Public TotFaFdTonsAdj As Long
        Public TotFaRc As Long

        Public N1frFdTonsAdj As Long
        Public N2frFdTonsAdj As Long
        Public N3frFdTonsAdj As Long
        Public N4frFdTonsAdj As Long
        Public N5frFdTonsAdj As Long
        Public CrFdTonsAdj As Long
        Public SrFdTonsAdj As Long
        Public UcrFdTonsAdj As Long
        Public TotRghrFdTonsAdj As Long

        Public N1frTlTonsAdj As Long
        Public N2frTlTonsAdj As Long
        Public N3frTlTonsAdj As Long
        Public N4frTlTonsAdj As Long
        Public N5frTlTonsAdj As Long
        Public SrTlTonsAdj As Long
        Public CrTlTonsAdj As Long
        Public UcrTlTonsAdj As Long

        Public TotTlTons As Long

        Public CcnTonCorr As Long
        Public FcnTonCorr As Long
    End Structure

    Dim MbSfShift As MassBalanceSfShiftType

    Private Structure MassBalanceSfTotalType
        Public PrdFcnTons As Long
        Public PrdFcnTonsW As Long
        Public PrdFcnBt As Double
        Public PrdFcnBpl As Single

        Public PrdCcnTons As Long
        Public PrdCcnTonsW As Long
        Public PrdCcnBt As Double
        Public PrdCcnBpl As Single

        Public PrdUccnTons As Long
        Public PrdUccnTonsW As Long
        Public PrdUccnBt As Double
        Public PrdUccnBpl As Single

        Public PrdCnTonsWuc As Long
        Public PrdCnBplWuc As Single
        Public PrdCnTonsWouc As Long
        Public PrdCnBplWouc As Single

        Public PrdCnBt As Double

        Public N1frHrs As Single
        Public N2frHrs As Single
        Public N3frHrs As Single
        Public N4frHrs As Single
        Public N5frHrs As Single
        Public SrHrs As Single
        Public CrHrs As Single
        Public UcrHrs As Single

        Public N1frFdTonsRpt As Long
        Public N2frFdTonsRpt As Long
        Public N3frFdTonsRpt As Long
        Public N4frFdTonsRpt As Long
        Public N5frFdTonsRpt As Long
        Public SrFdTonsRpt As Long
        Public CrFdTonsRpt As Long
        Public UcrFdTonsRpt As Long

        Public N1frFdTonsRptW As Long          'Tons with BPL
        Public N2frFdTonsRptW As Long
        Public N3frFdTonsRptW As Long
        Public N4frFdTonsRptW As Long
        Public N5frFdTonsRptW As Long
        Public SrFdTonsRptW As Long
        Public CrFdTonsRptW As Long
        Public UcrFdTonsRptW As Long

        Public N1frFdBtRpt As Double         'Tons X BPL  (BPL Tons)
        Public N2frFdBtRpt As Double
        Public N3frFdBtRpt As Double
        Public N4frFdBtRpt As Double
        Public N5frFdBtRpt As Double
        Public SrFdBtRpt As Double
        Public CrFdBtRpt As Double
        Public UcrFdBtRpt As Double

        Public N1faFdTonsRpt As Long
        Public N2faFdTonsRpt As Long
        Public N3faFdTonsRpt As Long
        Public CaFdTonsRpt As Long

        Public N1faFdTonsAdj As Long
        Public N1faFdBtAdj As Double
        Public N2faFdTonsAdj As Long
        Public N2faFdBtAdj As Double
        Public N3faFdTonsAdj As Long
        Public N3faFdBtAdj As Double
        Public CaFdTonsAdj As Long
        Public CaFdBtAdj As Double

        Public N1faCnBtAdj As Double      '10/31/2011, lss  New
        Public N2faCnBtAdj As Double      '10/31/2011, lss  New
        Public N3faCnBtAdj As Double      '10/31/2011, lss  New

        Public N1frCnTonsAdj As Long
        Public N2FrCnTonsAdj As Long
        Public N3FrCnTonsAdj As Long
        Public N4FrCnTonsAdj As Long
        Public N5FrCnTonsAdj As Long
        Public CrCnTonsAdj As Long
        Public SrCnTonsAdj As Long
        Public UcrCnTonsAdj As Long

        Public N1frCnBtAdj As Double
        Public N2FrCnBtAdj As Double
        Public N3FrCnBtAdj As Double
        Public N4FrCnBtAdj As Double
        Public N5FrCnBtAdj As Double
        Public CrCnBtAdj As Double
        Public SrCnBtAdj As Double
        Public UcrCnBtAdj As Double

        Public N1faTlTonsAdj As Long
        Public N2faTlTonsAdj As Long
        Public N3faTlTonsAdj As Long
        Public CaTlTonsAdj As Long

        Public N1faTlBtAdj As Double
        Public N2faTlBtAdj As Double
        Public N3faTlBtAdj As Double
        Public CaTlBtAdj As Double

        Public TotFaFdTonsAdj As Long

        Public N1frFdTonsAdj As Long
        Public N2frFdTonsAdj As Long
        Public N3frFdTonsAdj As Long
        Public N4frFdTonsAdj As Long
        Public N5frFdTonsAdj As Long
        Public CrFdTonsAdj As Long
        Public SrFdTonsAdj As Long
        Public UcrFdTonsAdj As Long

        Public N1frFdBtAdj As Double
        Public N2frFdBtAdj As Double
        Public N3frFdBtAdj As Double
        Public N4frFdBtAdj As Double
        Public N5frFdBtAdj As Double
        Public CrFdBtAdj As Double
        Public SrFdBtAdj As Double
        Public UcrFdBtAdj As Double

        Public N1frTlTonsAdj As Long
        Public N2frTlTonsAdj As Long
        Public N3frTlTonsAdj As Long
        Public N4frTlTonsAdj As Long
        Public N5frTlTonsAdj As Long
        Public SrTlTonsAdj As Long
        Public CrTlTonsAdj As Long
        Public UcrTlTonsAdj As Long

        Public N1frTlBtAdj As Double
        Public N2frTlBtAdj As Double
        Public N3frTlBtAdj As Double
        Public N4frTlBtAdj As Double
        Public N5frTlBtAdj As Double
        Public SrTlBtAdj As Double
        Public CrTlBtAdj As Double
        Public UcrTlBtAdj As Double

        Public TotGmtTlTonsW As Long
        Public TotGmtTlBt As Double

        Public N1frFdBpl As Single
        Public N2frFdBpl As Single
        Public N3frFdBpl As Single
        Public N4frFdBpl As Single
        Public N5frFdBpl As Single
        Public SrFdBpl As Single
        Public CrFdBpl As Single
        Public UcrFdBpl As Single

        Public N1frArFdBpl As Single
        Public N2frArFdBpl As Single
        Public N3frArFdBpl As Single
        Public N4frArFdBpl As Single
        Public N5frArFdBpl As Single
        Public SrArFdBpl As Single
        Public CrArFdBpl As Single
        Public UcrArFdBpl As Single

        Public N1frCnBpl As Single
        Public N2frCnBpl As Single
        Public N3frCnBpl As Single
        Public N4frCnBpl As Single
        Public N5frCnBpl As Single
        Public SrCnBpl As Single
        Public CrCnBpl As Single
        Public UcrCnBpl As Single

        Public N1frTlBpl As Single
        Public N2frTlBpl As Single
        Public N3frTlBpl As Single
        Public N4frTlBpl As Single
        Public N5frTlBpl As Single
        Public SrTlBpl As Single
        Public CrTlBpl As Single
        Public UcrTlBpl As Single

        Public N1faFdBpl As Single
        Public N2faFdBpl As Single
        Public N3faFdBpl As Single
        Public CaFdBpl As Single

        Public N1faCnBpl As Single
        Public N2faCnBpl As Single
        Public N3faCnBpl As Single
        Public CaCnBpl As Single

        Public N1faTlBpl As Single
        Public N2faTlBpl As Single
        Public N3faTlBpl As Single
        Public CaTlBpl As Single

        Public N1frRc As Single
        Public N2frRc As Single
        Public N3frRc As Single
        Public N4frRc As Single
        Public N5frRc As Single
        Public SrRc As Single
        Public CrRc As Single
        Public UcrRc As Single
        Public N1faRc As Single
        Public N2faRc As Single
        Public N3faRc As Single
        Public CaRc As Single

        Public N1frFdTph As Double
        Public N1frCnTph As Double
        Public N1frTlTph As Double
        Public N1frActPctRcvry As Double
        Public N1frStdRc As Double
        Public N1frStdCnTons As Long
        Public N1frStdPctRcvry As Double

        Public N2frFdTph As Double
        Public N2frCnTph As Double
        Public N2frTlTph As Double
        Public N2frActPctRcvry As Double
        Public N2frStdRc As Double
        Public N2frStdCnTons As Long
        Public N2frStdPctRcvry As Double

        Public N3frFdTph As Double
        Public N3frCnTph As Double
        Public N3frTlTph As Double
        Public N3frActPctRcvry As Double
        Public N3frStdRc As Double
        Public N3frStdCnTons As Long
        Public N3frStdPctRcvry As Double

        Public N4frFdTph As Double
        Public N4frCnTph As Double
        Public N4frTlTph As Double
        Public N4frActPctRcvry As Double
        Public N4frStdRc As Double
        Public N4frStdCnTons As Long
        Public N4frStdPctRcvry As Double

        Public N5frFdTph As Double
        Public N5frCnTph As Double
        Public N5frTlTph As Double
        Public N5frActPctRcvry As Double
        Public N5frStdRc As Double
        Public N5frStdCnTons As Long
        Public N5frStdPctRcvry As Double

        Public CrFdTph As Double
        Public CrCnTph As Double
        Public CrTlTph As Double
        Public CrActPctRcvry As Double
        Public CrStdRc As Double
        Public CrStdCnTons As Long
        Public CrStdPctRcvry As Double

        Public SrFdTph As Double
        Public SrCnTph As Double
        Public SrTlTph As Double
        Public SrActPctRcvry As Double
        Public SrStdRc As Double
        Public SrStdCnTons As Long
        Public SrStdPctRcvry As Double

        Public UcrFdTph As Double
        Public UcrCnTph As Double
        Public UcrTlTph As Double
        Public UcrActPctRcvry As Double
        Public UcrStdRc As Double
        Public UcrStdCnTons As Long
        Public UcrStdPctRcvry As Double

        Public N1faCnTons As Long
        Public N2faCnTons As Long
        Public N3faCnTons As Long
        Public CaCnTons As Long

        Public N1faActPctRcvry As Double
        Public N1faStdRc As Double
        Public N1faStdCnTons As Long
        Public N1faStdPctRcvry As Double

        Public N2faActPctRcvry As Double
        Public N2faStdRc As Double
        Public N2faStdCnTons As Long
        Public N2faStdPctRcvry As Double

        Public N3faActPctRcvry As Double
        Public N3faStdRc As Double
        Public N3faStdCnTons As Long
        Public N3faStdPctRcvry As Double

        Public CaActPctRcvry As Double
        Public CaStdRc As Double
        Public CaStdCnTons As Long
        Public CaStdPctRcvry As Double

        Public TotFdBplAdj As Single
        Public TotFdTonsAdj As Long
        Public TotTlBtAdjMeth1 As Double
        Public TotTlTonsAdjMeth1 As Long
        Public TotTlTonsAdjMeth2 As Long
        Public TotTlBplMsrd As Single
        Public TotFdTonsMsrd As Long
        Public TotFdBtAdj As Long
        Public TotTlBplAdjFd As Single
        Public TotTlBplRptFd As Single
        Public TotFdTonsRpt As Long
        Public TotFdTonsRptW As Long
        Public TotFdTonsRptBt As Double
        Public TotTlTonsRpt As Long
        Public TotFdBplRpt As Single
        Public TotTlBplFromCircs As Single

        Public TotTlBplMsrdRc As Single
        Public TotTlBPlMsrdRcvry As Single

        Public TotTlBplAdjFdRc As Single
        Public TotTlBplAdjFdRcvry As Single

        Public TotTlBplRptFdRc As Single
        Public TotTlBplRptFdRcvry As Single

        Public TotFineRghrFdTons As Long
        Public TotFineRghrFdBt As Double
        Public TotFineRghrFdTonsW As Long
        Public TotFineRghrFdBpl As Single
        Public TotFineRghrCnTons As Long
        Public TotFineRghrCnBt As Double
        Public TotFineRghrCnTonsW As Long
        Public TotFineRghrCnBpl As Single
        Public TotFineRghrTlTons As Long
        Public TotFineRghrTlBt As Double
        Public TotFineRghrTlTonsW As Long
        Public TotFineRghrTlBpl As Single
        Public TotFineRghrRc As Single
        Public TotFineRghrRcvry As Single

        Public TotCrsRghrFdTons As Long
        Public TotCrsRghrFdBt As Double
        Public TotCrsRghrFdTonsW As Long
        Public TotCrsRghrFdBpl As Single
        Public TotCrsRghrCnTons As Long
        Public TotCrsRghrCnBt As Double
        Public TotCrsRghrCnTonsW As Long
        Public TotCrsRghrCnBpl As Single
        Public TotCrsRghrTlTons As Long
        Public TotCrsRghrTlBt As Double
        Public TotCrsRghrTlTonsW As Long
        Public TotCrsRghrTlBpl As Single
        Public TotCrsRghrRc As Single
        Public TotCrsRghrRcvry As Single

        Public TotRghrFdTons As Long
        Public TotRghrFdTonsRpt As Long
        Public TotRghrFdBt As Double
        Public TotRghrFdTonsW As Long
        Public TotRghrFdBpl As Single
        Public TotRghrCnTons As Long
        Public TotRghrCnBt As Double
        Public TotRghrCnTonsW As Long
        Public TotRghrCnBpl As Single
        Public TotRghrTlTons As Long
        Public TotRghrTlBt As Double
        Public TotRghrTlTonsW As Long
        Public TotRghrTlBpl As Single
        Public TotRghrRc As Single
        Public TotRghrRcvry As Single

        Public TotRghr2FdTons As Long
        Public TotRghr2FdBt As Double
        Public TotRghr2FdTonsW As Long
        Public TotRghr2FdBpl As Single
        Public TotRghr2CnTons As Long
        Public TotRghr2CnBt As Double
        Public TotRghr2CnTonsW As Long
        Public TotRghr2CnBpl As Single
        Public TotRghr2TlTons As Long
        Public TotRghr2TlBt As Double
        Public TotRghr2TlTonsW As Long
        Public TotRghr2TlBpl As Single
        Public TotRghr2Rc As Single
        Public TotRghr2Rcvry As Single

        Public TotClnrFdTons As Long
        Public TotClnrFdTonsRpt As Long
        Public TotClnrFdBt As Double
        Public TotClnrFdTonsW As Long
        Public TotClnrFdBpl As Single
        Public TotClnrCnTons As Long
        Public TotClnrCnBt As Double
        Public TotClnrCnTonsW As Long
        Public TotClnrCnBpl As Single
        Public TotClnrTlTons As Long
        Public TotClnrTlBt As Double
        Public TotClnrTlTonsW As Long
        Public TotClnrTlBpl As Single
        Public TotClnrRc As Single
        Public TotClnrRcvry As Single

        Public TotFineClnrFdTons As Long
        Public TotFineClnrFdBt As Double
        Public TotFineClnrFdTonsW As Long
        Public TotFineClnrFdBpl As Single
        Public TotFineClnrCnTons As Long
        Public TotFineClnrCnBt As Double
        Public TotFineClnrCnTonsW As Long
        Public TotFineClnrCnBpl As Single
        Public TotFineClnrTlTons As Long
        Public TotFineClnrTlBt As Double
        Public TotFineClnrTlTonsW As Long
        Public TotFineClnrTlBpl As Single
        Public TotFineClnrRc As Single
        Public TotFineClnrRcvry As Single

        Public TotPlantFdTons As Long
        Public TotPlantFdTonsRpt As Long
        Public TotPlantCnTons As Long
        Public TotPlantTlTons As Long
        Public TotPlantFdBpl As Single
        Public TotPlantCnBpl As Single
        Public TotPlantTlBpl As Single
        Public TotPlantRc As Single
        Public TotPlantRcvry As Single

        Public TotCombFdTons As Long
        Public TotCombFdTonsRpt As Long
        Public TotCombCnTons As Long
        Public TotCombTlTons As Long
        Public TotCombFdBpl As Single
        Public TotCombCnBpl As Single
        Public TotCombTlBpl As Single
        Public TotCombRc As Single
        Public TotCombRcvry As Single

    End Structure
    Dim MbSfTotal As MassBalanceSfTotalType

    'Transfer possibilities between circuits:
    ' 1)  1FR to 1FA
    ' 2)  1FR to 2FA

    ' 3)  2FR to 2FA
    ' 4)  2FR to 3FA

    ' 5)  3FR to 2FA
    ' 6)  3FR to 3FA

    ' 7)  4FR to 1FA
    ' 8)  4FR to 2FA

    ' 9)  5FR to 1FA
    '10)  5FR to 2FA

    '11)  SR to 1FA
    '12)  SR to 2FA
    '13)  SR to 3FA

    '14)  CR to CA
    '15)  CR to 1FA

    '16)  SR to CA    New -- 05/23/2005 ,lss

    'Sources of feed for 1FA
    '1)  1FR
    '2)  4FR
    '3)  5FR
    '4)  SR
    '5)  CR

    'Sources of feed for 2FA
    '1)  1FR
    '2)  2FR
    '3)  3FR
    '4)  4FR
    '5)  5FR
    '6)  SR

    'Sources of feed for 3FA
    '1)  2FR
    '2)  3FR
    '3)  SR

    'Sources of feed for CA
    '1)  CR
    '2)  SR     New -- 05/23/2005 ,lss

    'Sort of a transfer type added on 01/13/03, lss
    '1)  "CA to FC"  some or all of the product from the
    '    coarse amine circuit may be directed to the
    '    fine concentrate bins rather than the coarse
    '    concentrate bins.

    Private Structure TransferType
        Public N1FRtoN1FA As Single    ' 1)
        Public N1FRtoN2FA As Single    ' 2)
        Public N2FRtoN2FA As Single    ' 3)
        Public N2FRtoN3FA As Single    ' 4)
        Public N3FRtoN2FA As Single    ' 5)
        Public N3FRtoN3FA As Single    ' 6)
        Public N4FRtoN1FA As Single    ' 7)
        Public N4FRtoN2FA As Single    ' 8)
        Public N5FRtoN1FA As Single    ' 9)
        Public N5FRtoN2FA As Single    '10)
        Public SRtoN1FA As Single    '11)
        Public SRtoN2FA As Single    '12)
        Public SRtoN3FA As Single    '13)
        Public CRtoCA As Single    '14)
        Public CRtoN1FA As Single    '15)
        Public CAtoFC As Single    '16)
        Public SRtoCA As Single    '17) -- new 05/23/2005, lss
    End Structure
    Dim Transfer As TransferType

    Private Structure N1faFdTonsType
        Public N1fr As Long
        Public N4fr As Long
        Public N5fr As Long
        Public Sr As Long
        Public Cr As Long
    End Structure

    Private Structure N2faFdTonsType
        Public N1fr As Long
        Public N2fr As Long
        Public N3fr As Long
        Public N4fr As Long
        Public N5fr As Long
        Public Sr As Long
    End Structure

    Private Structure N3faFdTonsType
        Public N2fr As Long
        Public N3fr As Long
        Public Sr As Long
    End Structure

    Private Structure CaFdTonsType
        Public Cr As Long
        Public Sr As Long
    End Structure

    Dim N1faFdTons As N1faFdTonsType
    Dim N2faFdTons As N2faFdTonsType
    Dim N3faFdTons As N3faFdTonsType
    Dim CaFdTons As CaFdTonsType

    Dim mMassBalanceDynaset As OraDynaset

    Dim mRoundVal As Integer
    Dim mUseFaCnChange As Boolean

    Public Function gGetSfFloatPlantBalanceData(ByRef FloatPlantCirc As Object, _
                                                ByRef FloatPlantGmt As Object, _
                                                ByVal aBeginDate As Date, _
                                                ByVal aBeginShift As String, _
                                                ByVal aEndDate As Date, _
                                                ByVal aEndShift As String, _
                                                ByVal aCrewNumber As String, _
                                                ByVal aSkipDownMonths As Integer) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************


        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade

        'This function will return the number of shifts processed.
        'It will also "return" data through the fFloatPlantData array.
        On Error GoTo gGetSfFloatPlantBalanceDataError

        Dim RowIdx As Integer
        Dim ColIdx As Integer

        Dim CalcNumShifts As Integer
        Dim ActualNumshifts As Integer

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        Dim RecCount As Long

        Dim CurrentDate As Date
        Dim CurrentShift As String
        Dim ThisDate As Date
        Dim ThisShift As String
        Dim ThisEqpt As String

        Dim w_percent As Single
        Dim per1_2fa As Single
        Dim per1_1fa As Single
        Dim per23_1fa As Single
        Dim pers_1fr As Single
        Dim pers_crs As Single
        Dim sw_to1fr As Single
        Dim sw_tocr As Single
        Dim pers_4fr As Single
        Dim perc_1fa As Single
        Dim perc_crs As Single
        Dim persw_2fa As Single
        Dim pers_s3fa As Single
        Dim per23_3fa As Single

        Dim NumShifts As Integer

        Dim IsPlantDownDay As Boolean

        mUseFaCnChange = True

        CalcNumShifts = (DateDiff(DateInterval.Day, aEndDate, aBeginDate) + 1) * 2
        If aBeginShift = aEndShift Then
            CalcNumShifts = CalcNumShifts - 1
        End If

        'aCrewNumber will be "All", "1", "2", "3", or "4"
        If aCrewNumber = "All" Then
            NumShifts = CalcNumShifts
        Else
            'NumShifts = gGetCrewShiftCount("South Fort Meade", _
            '                               aBeginDate, _
            '                               aBeginShift, _
            '                               aEndDate, _
            '                               aEndShift, _
            '                               Val(aCrewNumber))
        End If

        ReDim FloatPlantCirc(0 To 19, 0 To 14)
        ReDim FloatPlantGmt(0 To 4, 0 To 7)

        mRoundVal = 4

        'fFloatPlantCirc
        '---------------
        '
        '       Rows                Columns
        '       --------------      ----------------
        ' 1)    #1FR                Hours
        ' 2)    #2FR                Feed tons reported
        ' 3)    #3FR                Feed tons adjusted
        ' 4)    #4FR                Feed BPL
        ' 5)    #5FR                Conc BPL
        ' 6)    SR                  Tail BPL
        ' 7)    CR                  Ratio of concentration
        ' 8)    UCR                 %Actual recovery
        ' 9)    Total roughers      %Standard recovery
        '10)    #1FA                Concentrate tons adjusted
        '11)    #2FA                Tail tons adjusted
        '12)    #3FA                Feed TPH
        '13)    CA                  Concentrate TPH
        '14)    Total amine         Tail TPH
        '15)    Grand totals

        'fFloatPlantGmt
        '---------------
        '
        '       Rows                            Columns
        '       --------------                  ----------------
        ' 1)    Based on as reported GMT BPL    Feed tons
        ' 2)    Based on calculated GMT BPL     Concentrate tons
        ' 3)    Based on reported feed tons     Feed BPL
        ' 4)    GMT from circuits               Concentrate BPL
        ' 5)                                    Tail BPL
        ' 6)                                    Ratio of concentration
        ' 7)                                    %Recovery

        For RowIdx = 1 To 19
            For ColIdx = 1 To 14
                FloatPlantCirc(RowIdx, ColIdx) = 0
            Next ColIdx
        Next RowIdx

        For RowIdx = 1 To 4
            For ColIdx = 1 To 7
                FloatPlantGmt(RowIdx, ColIdx) = 0
            Next ColIdx
        Next RowIdx

        ZeroSfSummingData()

        'Get basic floatplant data from EQPT_MSRMNT, EQPT_EXT_MSRMNT, EQPT_CALC

        params = gDBParams

        params.Add("pMineName", "South Fort Meade", ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pBeginDate", aBeginDate, ORAPARM_INPUT)
        params("pBeginDate").serverType = ORATYPE_DATE

        params.Add("pBeginShift", StrConv(aBeginShift, vbUpperCase), ORAPARM_INPUT)
        params("pBeginShift").serverType = ORATYPE_VARCHAR2

        params.Add("pEndDate", aEndDate, ORAPARM_INPUT)
        params("pEndDate").serverType = ORATYPE_DATE

        params.Add("pEndShift", StrConv(aEndShift, vbUpperCase), ORAPARM_INPUT)
        params("pEndShift").serverType = ORATYPE_VARCHAR2

        params.Add("pCrewNumber", aCrewNumber, ORAPARM_INPUT)
        params("pCrewNumber").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_floatplant.get_sf_mass_balance_data(:pMineName," + _
                      ":pBeginDate, :pBeginShift, :pEndDate, :pEndShift, :pCrewNumber, :pResult);end;", ORASQL_FAILEXEC)
        mMassBalanceDynaset = params("pResult").Value
        RecCount = mMassBalanceDynaset.RecordCount

        If RecCount = 0 Then
            ClearParams(params)
            Exit Function
        End If

        mMassBalanceDynaset.MoveFirst()
        CurrentDate = mMassBalanceDynaset.Fields("prod_date").Value
        CurrentShift = mMassBalanceDynaset.Fields("shift").Value

        Do While Not mMassBalanceDynaset.EOF
            ThisDate = mMassBalanceDynaset.Fields("prod_date").Value
            ThisShift = mMassBalanceDynaset.Fields("shift").Value

            If ThisDate = CurrentDate And ThisShift = CurrentShift Then
                ThisEqpt = mMassBalanceDynaset.Fields("eqpt_name").Value

                With MbSfShift
                    Select Case ThisEqpt
                        Case Is = "#1 Fine rougher"
                            .N1frTlBpl = mMassBalanceDynaset.Fields("tail_bpl").Value
                            .N1frFdBpl = mMassBalanceDynaset.Fields("feed_bpl").Value
                            .N1frCnBpl = mMassBalanceDynaset.Fields("concentrate_bpl").Value
                            .N1frFdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value
                            .N1frHrs = mMassBalanceDynaset.Fields("operating_hours").Value

                        Case Is = "#2 Fine rougher"
                            .N2frTlBpl = mMassBalanceDynaset.Fields("tail_bpl").Value
                            .N2frFdBpl = mMassBalanceDynaset.Fields("feed_bpl").Value
                            .N2frCnBpl = mMassBalanceDynaset.Fields("concentrate_bpl").Value
                            .N2frHrs = mMassBalanceDynaset.Fields("operating_hours").Value
                            .N2frFdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value

                        Case Is = "#3 Fine rougher"
                            .N3frTlBpl = mMassBalanceDynaset.Fields("tail_bpl").Value
                            .N3frFdBpl = mMassBalanceDynaset.Fields("feed_bpl").Value
                            .N3frCnBpl = mMassBalanceDynaset.Fields("concentrate_bpl").Value
                            .N3frHrs = mMassBalanceDynaset.Fields("operating_hours").Value
                            .N3frFdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value

                        Case Is = "#4 Fine rougher"
                            .N4frTlBpl = mMassBalanceDynaset.Fields("tail_bpl").Value
                            .N4frFdBpl = mMassBalanceDynaset.Fields("feed_bpl").Value
                            .N4frCnBpl = mMassBalanceDynaset.Fields("concentrate_bpl").Value
                            .N4frHrs = mMassBalanceDynaset.Fields("operating_hours").Value
                            .N4frFdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value

                        Case Is = "#5 Fine rougher"
                            .N5frTlBpl = mMassBalanceDynaset.Fields("tail_bpl").Value
                            .N5frFdBpl = mMassBalanceDynaset.Fields("feed_bpl").Value
                            .N5frCnBpl = mMassBalanceDynaset.Fields("concentrate_bpl").Value
                            .N5frHrs = mMassBalanceDynaset.Fields("operating_hours").Value
                            .N5frFdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value

                        Case Is = "Coarse rougher"
                            .CrTlBpl = mMassBalanceDynaset.Fields("tail_bpl").Value
                            .CrFdBpl = mMassBalanceDynaset.Fields("feed_bpl").Value
                            .CrCnBpl = mMassBalanceDynaset.Fields("concentrate_bpl").Value
                            .CrHrs = mMassBalanceDynaset.Fields("operating_hours").Value
                            .CrFdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value

                        Case Is = "Swing rougher"
                            .SrTlBpl = mMassBalanceDynaset.Fields("tail_bpl").Value
                            .SrFdBpl = mMassBalanceDynaset.Fields("feed_bpl").Value
                            .SrCnBpl = mMassBalanceDynaset.Fields("concentrate_bpl").Value
                            .SrHrs = mMassBalanceDynaset.Fields("operating_hours").Value
                            .SrFdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value

                        Case Is = "Ultra-coarse rougher"
                            .UcrTlBpl = mMassBalanceDynaset.Fields("tail_bpl").Value
                            .UcrFdBpl = mMassBalanceDynaset.Fields("feed_bpl").Value
                            .UcrHrs = mMassBalanceDynaset.Fields("operating_hours").Value
                            .UcrFdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value

                        Case Is = "#1 Fine amine"
                            .N1faTlBpl = mMassBalanceDynaset.Fields("tail_bpl").Value
                            .N1faFdBpl = mMassBalanceDynaset.Fields("feed_bpl").Value

                        Case Is = "#1 Fine amine"
                            .N1faTlBpl = mMassBalanceDynaset.Fields("tail_bpl").Value
                            .N1faFdBpl = mMassBalanceDynaset.Fields("feed_bpl").Value

                        Case Is = "#2 Fine amine"
                            .N2faTlBpl = mMassBalanceDynaset.Fields("tail_bpl").Value
                            .N2faFdBpl = mMassBalanceDynaset.Fields("feed_bpl").Value

                        Case Is = "#3 Fine amine"
                            .N3faTlBpl = mMassBalanceDynaset.Fields("tail_bpl").Value
                            .N3faFdBpl = mMassBalanceDynaset.Fields("feed_bpl").Value

                        Case Is = "Coarse amine"
                            .CaTlBpl = mMassBalanceDynaset.Fields("tail_bpl").Value
                            .CaFdBpl = mMassBalanceDynaset.Fields("feed_bpl").Value

                        Case Is = "Float plant"
                            .PrdCcnTons = mMassBalanceDynaset.Fields("coarse_concentrate_tons").Value
                            .PrdCcnBpl = mMassBalanceDynaset.Fields("coarse_concentrate_bpl").Value

                            .PrdFcnTons = mMassBalanceDynaset.Fields("fine_concentrate_tons").Value
                            .PrdFcnBpl = mMassBalanceDynaset.Fields("fine_concentrate_bpl").Value

                            .PrdUccnTons = mMassBalanceDynaset.Fields("ultracoarse_concentrate_tons").Value
                            .PrdUccnBpl = mMassBalanceDynaset.Fields("ultracoarse_concentrate_bpl").Value

                            .GmtBpl = mMassBalanceDynaset.Fields("tail_bpl").Value

                            'Transfer factors, 1997 vintage are:
                            ' 1)  1FR to 1FA
                            ' 2)  1FR to 2FA
                            ' 3)  2FR to 2FA
                            ' 4)  2FR to 3FA
                            ' 5)  3FR to 2FA
                            ' 6)  3FR to 3FA
                            ' 7)  4FR to 1FA
                            ' 8)  4FR to 2FA
                            ' 9)  5FR to 1FA
                            '10)  5FR to 2FA
                            '11)  SR to 1FA
                            '12)  SR to 2FA
                            '13)  SR to 3FA
                            '14)  CR to CA
                            '15)  CR to 1FA

                            '16)  CA to FC      'Added 01/13/2003, lss
                            '17)  SR to CA      'Added 05/23/2005, lss

                            '1)  1FR to 1FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_1fr_to_1fa").Value) Then
                                Transfer.N1FRtoN1FA = 0
                            Else
                                Transfer.N1FRtoN1FA = mMassBalanceDynaset.Fields("flow_1fr_to_1fa").Value
                            End If

                            '2)  1FR to 2FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_1fr_to_2fa").Value) Then
                                Transfer.N1FRtoN2FA = 0
                            Else
                                Transfer.N1FRtoN2FA = mMassBalanceDynaset.Fields("flow_1fr_to_2fa").Value
                            End If

                            '3)  2FR to 2FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_2fr_to_2fa").Value) Then
                                Transfer.N2FRtoN2FA = 0
                            Else
                                Transfer.N2FRtoN2FA = mMassBalanceDynaset.Fields("flow_2fr_to_2fa").Value
                            End If

                            '4)  2FR to 3FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_2fr_to_3fa").Value) Then
                                Transfer.N2FRtoN3FA = 0
                            Else
                                Transfer.N2FRtoN3FA = mMassBalanceDynaset.Fields("flow_2fr_to_3fa").Value
                            End If

                            '5)  3FR to 2FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_3fr_to_2fa").Value) Then
                                Transfer.N3FRtoN2FA = 0
                            Else
                                Transfer.N3FRtoN2FA = mMassBalanceDynaset.Fields("flow_3fr_to_2fa").Value
                            End If

                            '6)  3FR to 3FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_3fr_to_3fa").Value) Then
                                Transfer.N3FRtoN3FA = 0
                            Else
                                Transfer.N3FRtoN3FA = mMassBalanceDynaset.Fields("flow_3fr_to_3fa").Value
                            End If

                            '7)  4FR to 1FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_4fr_to_1fa").Value) Then
                                Transfer.N4FRtoN1FA = 0
                            Else
                                Transfer.N4FRtoN1FA = mMassBalanceDynaset.Fields("flow_4fr_to_1fa").Value
                            End If

                            '8)  4FR to 2FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_4fr_to_2fa").Value) Then
                                Transfer.N4FRtoN2FA = 0
                            Else
                                Transfer.N4FRtoN2FA = mMassBalanceDynaset.Fields("flow_4fr_to_2fa").Value
                            End If

                            '9)  5FR to 1FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_5fr_to_1fa").Value) Then
                                Transfer.N5FRtoN1FA = 0
                            Else
                                Transfer.N5FRtoN1FA = mMassBalanceDynaset.Fields("flow_5fr_to_1fa").Value
                            End If

                            '10)  5FR to 2FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_5fr_to_2fa").Value) Then
                                Transfer.N5FRtoN2FA = 0
                            Else
                                Transfer.N5FRtoN2FA = mMassBalanceDynaset.Fields("flow_5fr_to_2fa").Value
                            End If

                            '11)  SR to 1FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_sr_to_1fa").Value) Then
                                Transfer.SRtoN1FA = 0
                            Else
                                Transfer.SRtoN1FA = mMassBalanceDynaset.Fields("flow_sr_to_1fa").Value
                            End If

                            '12)  SR to 2FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_sr_to_2fa").Value) Then
                                Transfer.SRtoN2FA = 0
                            Else
                                Transfer.SRtoN2FA = mMassBalanceDynaset.Fields("flow_sr_to_2fa").Value
                            End If

                            '13)  SR to 3FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_sr_to_3fa").Value) Then
                                Transfer.SRtoN3FA = 0
                            Else
                                Transfer.SRtoN3FA = mMassBalanceDynaset.Fields("flow_sr_to_3fa").Value
                            End If

                            '14)  CR to CA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_cr_to_ca").Value) Then
                                Transfer.CRtoCA = 0
                            Else
                                Transfer.CRtoCA = mMassBalanceDynaset.Fields("flow_cr_to_ca").Value
                            End If

                            '15)  CR to 1FA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_cr_to_1fa").Value) Then
                                Transfer.CRtoN1FA = 0
                            Else
                                Transfer.CRtoN1FA = mMassBalanceDynaset.Fields("flow_cr_to_1fa").Value
                            End If

                            '16)  CA to FC
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_ca_to_fc").Value) Then
                                Transfer.CAtoFC = 0
                            Else
                                Transfer.CAtoFC = mMassBalanceDynaset.Fields("flow_ca_to_fc").Value
                            End If

                            '17)  SR to CA
                            If IsDBNull(mMassBalanceDynaset.Fields("flow_sr_to_ca").Value) Then
                                Transfer.SRtoCA = 0
                            Else
                                Transfer.SRtoCA = mMassBalanceDynaset.Fields("flow_sr_to_ca").Value
                            End If
                    End Select
                End With
                mMassBalanceDynaset.MoveNext()

            Else    'ThisDate <> CurrentDate Or ThisShift <> CurrentShift
                IsPlantDownDay = False
                If aSkipDownMonths = 1 Then
                    If CurrentDate >= #9/1/2010# And CurrentDate <= #11/30/2010# Then
                        IsPlantDownDay = True
                    Else
                        IsPlantDownDay = False
                    End If
                End If

                If IsPlantDownDay = False Then    'Added!
                    ProcessSfMassBalanceData()
                End If                            'Added!
                CurrentDate = ThisDate
                CurrentShift = ThisShift
            End If
        Loop

        'Process last shift's worth of data
        IsPlantDownDay = False
        If aSkipDownMonths = 1 Then
            If CurrentDate >= #9/1/2010# And CurrentDate <= #11/30/2010# Then
                IsPlantDownDay = True
            Else
                IsPlantDownDay = False
            End If
        End If

        If IsPlantDownDay = False Then    'Added!
            ProcessSfMassBalanceData()
        End If                            'Added!

        'Summing of mass balance shift data completed

        ProcessSfMassBalanceTotals()

        'Place data in array  Place data in array  Place data in array
        'Place data in array  Place data in array  Place data in array
        'Place data in array  Place data in array  Place data in array

        'FloatPlantCirc()

        'Rows in the array                Columns in the array
        'sfCrN1FineRghr = 1               sfCcOperHrs = 1
        'sfCrN2FineRghr = 2               sfCcFdTonsRpt = 2
        'sfCrN3FineRghr = 3               sfCcFdTonsAdj = 3
        'sfCrN4FineRghr = 4               sfCcFdBpl = 4
        'sfCrN5FineRghr = 5               sfCcCnBpl = 5
        'sfCrSwingRghr = 6                sfCcTlBpl = 6
        'sfCrCoarseRghr = 7               sfCcRc = 7
        'sfCrUltraCoarseRghr = 8          sfCcPctActRcvry = 8
        'sfCrTotalRghr = 9                sfCcPctStdRcvry = 9
        'sfCrN1FineAmine = 10             sfCcCnTonsAdj = 10
        'sfCrN2FineAmine = 11             sfCcTlTonsAdj = 11
        'sfCrN3FineAmine = 12             sfCcFdTph = 12
        'sfCrCoarseAmine = 13             sfCcCnTph = 13
        'sfCrTotalAmine = 14              sfCcTlTph = 14
        'sfCrGrandTotal = 15
        'sfCrFcnProduct = 16
        'sfCrCcnProduct = 17
        'sfCrUccnProduct = 18
        'sfCrCnProduct = 19

        'Product tons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrFcnProduct, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.PrdFcnTons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrFcnProduct, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.PrdFcnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCcnProduct, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.PrdCcnTons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCcnProduct, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.PrdCcnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUccnProduct, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.PrdUccnTons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUccnProduct, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.PrdUccnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCnProduct, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.PrdCnTonsWuc
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCnProduct, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.PrdCnBplWuc

        'Operating hours
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcOperHrs) = MbSfTotal.N1frHrs
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcOperHrs) = MbSfTotal.N2frHrs
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcOperHrs) = MbSfTotal.N3frHrs
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcOperHrs) = MbSfTotal.N4frHrs
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcOperHrs) = MbSfTotal.N5frHrs
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcOperHrs) = MbSfTotal.SrHrs
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcOperHrs) = MbSfTotal.CrHrs
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcOperHrs) = MbSfTotal.UcrHrs

        'Feed tons reported
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.N1frFdTonsRpt
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.N2frFdTonsRpt
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.N3frFdTonsRpt
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.N4frFdTonsRpt
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.N5frFdTonsRpt
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.SrFdTonsRpt
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.CrFdTonsRpt
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.UcrFdTonsRpt
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.N1faFdTonsRpt
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.N2faFdTonsRpt
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.N3faFdTonsRpt
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.CaFdTonsRpt
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.TotRghrFdTonsRpt
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalClnr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt) = MbSfTotal.TotClnrFdTonsRpt

        'Feed tons adjusted
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.N1frFdTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.N2frFdTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.N3frFdTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.N4frFdTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.N5frFdTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.SrFdTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.CrFdTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.UcrFdTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.N1faFdTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.N2faFdTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.N3faFdTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.CaFdTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.TotRghrFdTons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalClnr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.TotClnrFdTons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrGrandTotal, mSfFloatPlantCircColEnum.sfCcFdTonsAdj) = MbSfTotal.TotPlantFdTons

        'Feed BPL
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcFdBpl) = MbSfTotal.N1frFdBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcFdBpl) = MbSfTotal.N2frFdBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcFdBpl) = MbSfTotal.N3frFdBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcFdBpl) = MbSfTotal.N4frFdBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcFdBpl) = MbSfTotal.N5frFdBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcFdBpl) = MbSfTotal.SrFdBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdBpl) = MbSfTotal.CrFdBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdBpl) = MbSfTotal.UcrFdBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcFdBpl) = MbSfTotal.N1faFdBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcFdBpl) = MbSfTotal.N2faFdBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcFdBpl) = MbSfTotal.N3faFdBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcFdBpl) = MbSfTotal.CaFdBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalRghr, mSfFloatPlantCircColEnum.sfCcFdBpl) = MbSfTotal.TotRghrFdBpl

        'Concentrate BPL
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.N1frCnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.N2frCnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.N3frCnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.N4frCnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.N5frCnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.SrCnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.CrCnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.UcrCnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.N1faCnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.N2faCnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.N3faCnBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcCnBpl) = MbSfTotal.CaCnBpl

        'Tail BPL
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcTlBpl) = MbSfTotal.N1frTlBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcTlBpl) = MbSfTotal.N2frTlBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcTlBpl) = MbSfTotal.N3frTlBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcTlBpl) = MbSfTotal.N4frTlBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcTlBpl) = MbSfTotal.N5frTlBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcTlBpl) = MbSfTotal.SrTlBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcTlBpl) = MbSfTotal.CrTlBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcTlBpl) = MbSfTotal.UcrTlBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcTlBpl) = MbSfTotal.N1faTlBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcTlBpl) = MbSfTotal.N2faTlBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcTlBpl) = MbSfTotal.N3faTlBpl
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcTlBpl) = MbSfTotal.CaTlBpl

        'Ratio of concentration
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcRc) = MbSfTotal.N1frRc
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcRc) = MbSfTotal.N2frRc
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcRc) = MbSfTotal.N3frRc
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcRc) = MbSfTotal.N4frRc
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcRc) = MbSfTotal.N5frRc
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcRc) = MbSfTotal.SrRc
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcRc) = MbSfTotal.CrRc
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcRc) = MbSfTotal.UcrRc
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcRc) = MbSfTotal.N1faRc
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcRc) = MbSfTotal.N2faRc
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcRc) = MbSfTotal.N3faRc
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcRc) = MbSfTotal.CaRc

        'Actual recovery
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry) = MbSfTotal.N1frActPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry) = MbSfTotal.N2frActPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry) = MbSfTotal.N3frActPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry) = MbSfTotal.N4frActPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry) = MbSfTotal.N5frActPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry) = MbSfTotal.SrActPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry) = MbSfTotal.CrActPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry) = MbSfTotal.UcrActPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcPctActRcvry) = MbSfTotal.N1faActPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcPctActRcvry) = MbSfTotal.N2faActPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcPctActRcvry) = MbSfTotal.N3faActPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcPctActRcvry) = MbSfTotal.CaActPctRcvry

        'Standard recovery
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry) = MbSfTotal.N1frStdPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry) = MbSfTotal.N2frStdPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry) = MbSfTotal.N3frStdPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry) = MbSfTotal.N4frStdPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry) = MbSfTotal.N5frStdPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry) = MbSfTotal.SrStdPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry) = MbSfTotal.CrStdPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry) = MbSfTotal.UcrStdPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcPctStdRcvry) = MbSfTotal.N1faStdPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcPctStdRcvry) = MbSfTotal.N2faStdPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcPctStdRcvry) = MbSfTotal.N3faStdPctRcvry
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcPctStdRcvry) = MbSfTotal.CaStdPctRcvry

        'Concentrate tons adjusted
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.N1frCnTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.N2FrCnTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.N3FrCnTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.N4FrCnTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.N5FrCnTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.SrCnTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.CrCnTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.UcrCnTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.N1faCnTons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.N2faCnTons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.N3faCnTons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.CaCnTons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.TotRghrCnTons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalClnr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.TotClnrCnTons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrGrandTotal, mSfFloatPlantCircColEnum.sfCcCnTonsAdj) = MbSfTotal.TotPlantCnTons

        'Tail tons adjusted
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.N1frTlTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.N2frTlTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.N3frTlTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.N4frTlTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.N5frTlTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.SrTlTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.CrTlTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.UcrTlTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.N1faTlTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.N2faTlTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.N3faTlTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.CaTlTonsAdj
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.TotRghrTlTons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalClnr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.TotClnrTlTons
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrGrandTotal, mSfFloatPlantCircColEnum.sfCcTlTonsAdj) = MbSfTotal.TotPlantTlTons

        'Feed TPH adjusted
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcFdTph) = MbSfTotal.N1frFdTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcFdTph) = MbSfTotal.N2frFdTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcFdTph) = MbSfTotal.N3frFdTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcFdTph) = MbSfTotal.N4frFdTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcFdTph) = MbSfTotal.N5frFdTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcFdTph) = MbSfTotal.SrFdTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdTph) = MbSfTotal.CrFdTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdTph) = MbSfTotal.UcrFdTph

        'Concentrate TPH adjusted
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcCnTph) = MbSfTotal.N1frCnTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcCnTph) = MbSfTotal.N2frCnTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcCnTph) = MbSfTotal.N3frCnTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcCnTph) = MbSfTotal.N4frCnTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcCnTph) = MbSfTotal.N5frCnTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcCnTph) = MbSfTotal.SrCnTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcCnTph) = MbSfTotal.CrCnTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcCnTph) = MbSfTotal.UcrCnTph

        'Tail TPH adjusted
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcTlTph) = MbSfTotal.N1frTlTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcTlTph) = MbSfTotal.N2frTlTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcTlTph) = MbSfTotal.N3frTlTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcTlTph) = MbSfTotal.N4frTlTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcTlTph) = MbSfTotal.N5frTlTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcTlTph) = MbSfTotal.SrTlTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcTlTph) = MbSfTotal.CrTlTph
        FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcTlTph) = MbSfTotal.UcrTlTph

        'FloatPlantGmt

        'Rows in the array              Columns in the array
        'grAsReportedGmtBpl = 1         gcFdTons = 1
        'grCalculatedGmtBpl = 2         gcCnTons = 2
        'grReportedFdTons = 3           gcFdBpl = 3
        'grGmtBplFromCircuits = 4       gcCnBpl = 4
        '                               gcTlBpl = 5
        '                               gcRC = 6
        '                               gcPctRcvry = 7

        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrAsReportedGmtBpl, mSfFloatPlantGmtColEnum.sfGcFdBpl) = MbSfTotal.TotFdBplAdj
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrAsReportedGmtBpl, mSfFloatPlantGmtColEnum.sfGcCnBpl) = MbSfTotal.PrdCnBplWuc
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrAsReportedGmtBpl, mSfFloatPlantGmtColEnum.sfGcTlBpl) = MbSfTotal.TotTlBplMsrd
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrAsReportedGmtBpl, mSfFloatPlantGmtColEnum.sfGcRc) = MbSfTotal.TotTlBplMsrdRc
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrAsReportedGmtBpl, mSfFloatPlantGmtColEnum.sfGcPctRcvry) = MbSfTotal.TotTlBPlMsrdRcvry

        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcFdTons) = MbSfTotal.TotFdTonsAdj
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcCnTons) = MbSfTotal.PrdCnTonsWuc
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcFdBpl) = MbSfTotal.TotFdBplAdj
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcCnBpl) = MbSfTotal.PrdCnBplWuc
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcTlBpl) = MbSfTotal.TotTlBplAdjFd
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcRc) = MbSfTotal.TotTlBplAdjFdRc
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcPctRcvry) = MbSfTotal.TotTlBplAdjFdRcvry

        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcFdTons) = MbSfTotal.TotFdTonsRpt
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcCnTons) = MbSfTotal.PrdCnTonsWuc
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcFdBpl) = MbSfTotal.TotFdBplRpt
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcCnBpl) = MbSfTotal.PrdCnBplWuc
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcTlBpl) = MbSfTotal.TotTlBplRptFd
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcRc) = MbSfTotal.TotTlBplRptFdRc
        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcPctRcvry) = MbSfTotal.TotTlBplRptFdRcvry

        FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrGmtBplFromCircuits, mSfFloatPlantGmtColEnum.sfGcTlBpl) = MbSfTotal.TotTlBplFromCircs

        ClearParams(params)

        gGetSfFloatPlantBalanceData = NumShifts

        Exit Function

gGetSfFloatPlantBalanceDataError:

        MsgBox("Error calculating South Fort Meade Mass Balance." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "South Fort Meade Mass Balance Calculation Error")

        On Error Resume Next
        ClearParams(params)
    End Function

    Public Function gMassBalanceSF(ByVal aBeginDate As Date, _
                                   ByVal aBeginShift As String, _
                                   ByVal aEndDate As Date, _
                                   ByVal aEndShift As String, _
                                   ByVal aCrewNumber As String, _
                                   ByVal aSkipDownMonths As Integer) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade

        'This function will return the number of shifts processed.
        'It will also "return" data through the fFloatPlantData array.

        On Error GoTo gMassBalanceSFError

        Dim ConnectString As String
        Dim RowIdx As Integer
        Dim TimeFrame As String
        Dim NumShifts As Integer

        Dim FloatPlantCirc(20, 15) As Object
        Dim FloatPlantGmt(5, 8) As Object
        'ReDim FloatPlantCirc(0 To 19, 0 To 14)
        'ReDim FloatPlantGmt(0 To 4, 0 To 7)

        Dim CcnTonsEst As Integer

        ' frmViewData.rptInputData.Reset()

        ZeroSfSummingData()

        'Miscellaneous data setup
        ' frmViewData.rptInputData.Formulas(0) = "MineName = '" & "South Fort Meade" & "'"
        If aBeginDate = aEndDate And aBeginShift = aEndShift Then
            TimeFrame = aBeginDate & " " & _
                        StrConv(aBeginShift, vbProperCase) & " Shift"
        Else
            TimeFrame = aBeginDate & " " & _
                        StrConv(aBeginShift, vbProperCase) & _
                        " Shift" & " thru " & _
                        aEndDate & " " & _
                        StrConv(aEndShift, vbProperCase) & " Shift"
        End If
        'frmViewData.rptInputData.Formulas(1) = "TimeFrame = '" & TimeFrame & "'"
        ' frmViewData.rptInputData.Formulas(2) = "CrewNumber = '" & aCrewNumber & "'"

        'Get data for float plant mass balance

        NumShifts = gGetSfFloatPlantBalanceData(FloatPlantCirc, _
                                                FloatPlantGmt, _
                                                aBeginDate, _
                                                StrConv(aBeginShift, vbUpperCase), _
                                                aEndDate, _
                                                StrConv(aEndShift, vbUpperCase), _
                                                aCrewNumber, _
                                                aSkipDownMonths)

        'gCCnTonsEst -- if True then coarse concentrate tons have been estimated.

        With MassBalanceSfRpt
            .N1frHrs = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcOperHrs)
            .N2frHrs = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcOperHrs)
            .N3frHrs = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcOperHrs)
            .N4frHrs = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcOperHrs)
            .N5frHrs = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcOperHrs)
            .SrHrs = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcOperHrs)
            .CrHrs = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcOperHrs)
            .UcrHrs = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcOperHrs)
            .N1faHrs = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcOperHrs)
            .N2faHrs = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcOperHrs)
            .N3faHrs = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcOperHrs)
            .CaHrs = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcOperHrs)

            .N1frFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)
            .N2frFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)
            .N3frFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)
            .N4frFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)
            .N5frFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)
            .SrFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)
            .CrFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)
            .UcrFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)
            .N1faFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)
            .N2faFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)
            .N3faFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)
            .CaFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)

            .N1frFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .N2frFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .N3frFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .N4frFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .N5frFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .SrFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .CrFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .UcrFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .N1faFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .N2faFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .N3faFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .CaFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)

            .N1frFdBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcFdBpl)
            .N2frFdBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcFdBpl)
            .N3frFdBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcFdBpl)
            .N4frFdBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcFdBpl)
            .N5frFdBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcFdBpl)
            .SrFdBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcFdBpl)
            .CrFdBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdBpl)
            .UcrFdBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcFdBpl)
            .N1faFdBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcFdBpl)
            .N2faFdBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcFdBpl)
            .N3faFdBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcFdBpl)
            .CaFdBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcFdBpl)

            .N1frCnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .N2frCnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .N3frCnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .N4frCnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .N5frCnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .SrCnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .CrCnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .UcrCnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .N1faCnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .N2faCnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .N3faCnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .CaCnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcCnBpl)

            .N1frTlBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcTlBpl)
            .N2frTlBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcTlBpl)
            .N3frTlBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcTlBpl)
            .N4frTlBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcTlBpl)
            .N5frTlBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcTlBpl)
            .SrTlBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcTlBpl)
            .CrTlBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcTlBpl)
            .UcrTlBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcTlBpl)
            .N1faTlBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcTlBpl)
            .N2faTlBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcTlBpl)
            .N3faTlBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcTlBpl)
            .CaTlBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcTlBpl)

            .N1frRc = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcRc)
            .N2frRc = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcRc)
            .N3frRc = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcRc)
            .N4frRc = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcRc)
            .N5frRc = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcRc)
            .SrRc = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcRc)
            .CrRc = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcRc)
            .UcrRc = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcRc)
            .N1faRc = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcRc)
            .N2faRc = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcRc)
            .N3faRc = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcRc)
            .CaRc = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcRc)

            .N1frAr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry)
            .N2frAr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry)
            .N3frAr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry)
            .N4frAr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry)
            .N5frAr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry)
            .SrAr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry)
            .CrAr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry)
            .UcrAr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcPctActRcvry)
            .N1faAr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcPctActRcvry)
            .N2faAr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcPctActRcvry)
            .N3faAr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcPctActRcvry)
            .CaAr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcPctActRcvry)

            .N1frSr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry)
            .N2frSr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry)
            .N3frSr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry)
            .N4frSr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry)
            .N5frSr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry)
            .SrSr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry)
            .CrSr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry)
            .UcrSr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcPctStdRcvry)
            .N1faSr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcPctStdRcvry)
            .N2faSr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcPctStdRcvry)
            .N3faSr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcPctStdRcvry)
            .CaSr = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcPctStdRcvry)

            .N1frCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .N2FrCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .N3FrCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .N4FrCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .N5FrCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .SrCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .CrCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .UcrCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .N1faCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .N2faCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .N3faCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .CaCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)

            .N1frTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)
            .N2frTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)
            .N3frTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)
            .N4frTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN4FineRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)
            .N5frTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN5FineRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)
            .SrTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrSwingRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)
            .CrTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)
            .UcrTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUltraCoarseRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)
            .N1faTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN1FineAmine, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)
            .N2faTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN2FineAmine, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)
            .N3faTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrN3FineAmine, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)
            .CaTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCoarseAmine, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)

            .SumRghrFdBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalRghr, mSfFloatPlantCircColEnum.sfCcFdBpl)
            .SumRghrFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)
            .SumRghrFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .SumRghrCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalRghr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .SumRghrTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalRghr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)
            .SumClnrFdTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalClnr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)
            .SumClnrFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalClnr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .SumClnrCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalClnr, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .SumClnrTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalClnr, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)
            .SumAllFdTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrGrandTotal, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
            .SumAllCnTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrGrandTotal, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .SumAllTlTonsAdj = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrGrandTotal, mSfFloatPlantCircColEnum.sfCcTlTonsAdj)

            .PrdFcnTons = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrFcnProduct, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .PrdFcnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrFcnProduct, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .PrdCcnTons = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCcnProduct, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .PrdCcnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCcnProduct, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .PrdUccnTons = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUccnProduct, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .PrdUccnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrUccnProduct, mSfFloatPlantCircColEnum.sfCcCnBpl)
            .PrdCnTons = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCnProduct, mSfFloatPlantCircColEnum.sfCcCnTonsAdj)
            .PrdCnBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrCnProduct, mSfFloatPlantCircColEnum.sfCcCnBpl)

            .ArGmtFdBpl = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrAsReportedGmtBpl, mSfFloatPlantGmtColEnum.sfGcFdBpl)
            .ArGmtCnBpl = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrAsReportedGmtBpl, mSfFloatPlantGmtColEnum.sfGcCnBpl)
            .ArGmtTlBpl = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrAsReportedGmtBpl, mSfFloatPlantGmtColEnum.sfGcTlBpl)
            .ArGmtRc = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrAsReportedGmtBpl, mSfFloatPlantGmtColEnum.sfGcRc)
            .ArGmtRcvry = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrAsReportedGmtBpl, mSfFloatPlantGmtColEnum.sfGcPctRcvry)

            .ClcGmtFdTons = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcFdTons)
            .ClcGmtCnTons = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcCnTons)
            .ClcGmtFdBpl = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcFdBpl)
            .ClcGmtCnBpl = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcCnBpl)
            .ClcGmtTlBpl = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcTlBpl)
            .ClcGmtRc = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcRc)
            .ClcGmtRcvry = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcPctRcvry)

            .ArFdTonsFdTons = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcFdTons)
            .ArFdTonsCnTons = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcCnTons)
            .ArFdTonsFdBpl = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcFdBpl)
            .ArFdTonsCnBpl = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcCnBpl)

            If FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcTlBpl) >= 0 Then
                .ArFdTonsTlBpl = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcTlBpl)
            Else
                .ArFdTonsTlBpl = 0
            End If
            .ArFdTonsRc = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcRc)
            .ArFdTonsRcvry = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrReportedFdTons, mSfFloatPlantGmtColEnum.sfGcPctRcvry)

            .GmtFromCircs = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrGmtBplFromCircuits, mSfFloatPlantGmtColEnum.sfGcTlBpl)

            .CcnTonsEst = gCcnTonsEst
        End With

        With MassBalanceSfRpt
            'frmViewData.rptInputData.Formulas(3) = "N1frHrs = " & .N1frHrs & ""
            'frmViewData.rptInputData.Formulas(4) = "N2frHrs = " & .N2frHrs & ""
            'frmViewData.rptInputData.Formulas(5) = "N3frHrs = " & .N3frHrs & ""
            'frmViewData.rptInputData.Formulas(6) = "N4frHrs = " & .N4frHrs & ""
            'frmViewData.rptInputData.Formulas(7) = "N5frHrs = " & .N5frHrs & ""
            'frmViewData.rptInputData.Formulas(8) = "SrHrs = " & .SrHrs & ""
            'frmViewData.rptInputData.Formulas(9) = "CrHrs = " & .CrHrs & ""
            'frmViewData.rptInputData.Formulas(10) = "UcrHrs = " & .UcrHrs & ""
            'frmViewData.rptInputData.Formulas(11) = "N1faHrs = " & .N1faHrs & ""
            'frmViewData.rptInputData.Formulas(12) = "N2faHrs = " & .N2faHrs & ""
            'frmViewData.rptInputData.Formulas(13) = "N3faHrs = " & .N3faHrs & ""
            'frmViewData.rptInputData.Formulas(14) = "CaHrs = " & .CaHrs & ""

            'frmViewData.rptInputData.Formulas(15) = "N1frFdTonsRpt = " & .N1frFdTonsRpt & ""
            'frmViewData.rptInputData.Formulas(16) = "N2frFdTonsRpt = " & .N2frFdTonsRpt & ""
            'frmViewData.rptInputData.Formulas(17) = "N3frFdTonsRpt = " & .N3frFdTonsRpt & ""
            'frmViewData.rptInputData.Formulas(18) = "N4frFdTonsRpt = " & .N4frFdTonsRpt & ""
            'frmViewData.rptInputData.Formulas(19) = "N5frFdTonsRpt = " & .N5frFdTonsRpt & ""
            'frmViewData.rptInputData.Formulas(20) = "SrFdTonsRpt = " & .SrFdTonsRpt & ""
            'frmViewData.rptInputData.Formulas(21) = "CrFdTonsRpt = " & .CrFdTonsRpt & ""
            'frmViewData.rptInputData.Formulas(22) = "UcrFdTonsRpt = " & .UcrFdTonsRpt & ""
            'frmViewData.rptInputData.Formulas(23) = "N1faFdTonsRpt = " & .N1faFdTonsRpt & ""
            'frmViewData.rptInputData.Formulas(24) = "N2faFdTonsRpt = " & .N2faFdTonsRpt & ""
            'frmViewData.rptInputData.Formulas(25) = "N3faFdTonsRpt = " & .N3faFdTonsRpt & ""
            'frmViewData.rptInputData.Formulas(26) = "CaFdTonsRpt = " & .CaFdTonsRpt & ""

            'frmViewData.rptInputData.Formulas(27) = "N1frFdTonsAdj = " & .N1frFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(28) = "N2frFdTonsAdj = " & .N2frFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(29) = "N3frFdTonsAdj = " & .N3frFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(30) = "N4frFdTonsAdj = " & .N4frFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(31) = "N5frFdTonsAdj = " & .N5frFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(32) = "SrFdTonsAdj = " & .SrFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(33) = "CrFdTonsAdj = " & .CrFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(34) = "UcrFdTonsAdj = " & .UcrFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(35) = "N1faFdTonsAdj = " & .N1faFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(36) = "N2faFdTonsAdj = " & .N2faFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(37) = "N3faFdTonsAdj = " & .N3faFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(38) = "CaFdTonsAdj = " & .CaFdTonsAdj & ""

            'frmViewData.rptInputData.Formulas(39) = "N1frFdBpl = " & .N1frFdBpl & ""
            'frmViewData.rptInputData.Formulas(40) = "N2frFdBpl = " & .N2frFdBpl & ""
            'frmViewData.rptInputData.Formulas(41) = "N3frFdBpl = " & .N3frFdBpl & ""
            'frmViewData.rptInputData.Formulas(42) = "N4frFdBpl = " & .N4frFdBpl & ""
            'frmViewData.rptInputData.Formulas(43) = "N5frFdBpl = " & .N5frFdBpl & ""
            'frmViewData.rptInputData.Formulas(44) = "SrFdBpl = " & .SrFdBpl & ""
            'frmViewData.rptInputData.Formulas(45) = "CrFdBpl = " & .CrFdBpl & ""
            'frmViewData.rptInputData.Formulas(46) = "UcrFdBpl = " & .UcrFdBpl & ""
            'frmViewData.rptInputData.Formulas(47) = "N1faFdBpl = " & .N1faFdBpl & ""
            'frmViewData.rptInputData.Formulas(48) = "N2faFdBpl = " & .N2faFdBpl & ""
            'frmViewData.rptInputData.Formulas(49) = "N3faFdBpl = " & .N3faFdBpl & ""
            'frmViewData.rptInputData.Formulas(50) = "CaFdBpl = " & .CaFdBpl & ""

            'frmViewData.rptInputData.Formulas(51) = "N1frCnBpl = " & .N1frCnBpl & ""
            'frmViewData.rptInputData.Formulas(52) = "N2frCnBpl = " & .N2frCnBpl & ""
            'frmViewData.rptInputData.Formulas(53) = "N3frCnBpl = " & .N3frCnBpl & ""
            'frmViewData.rptInputData.Formulas(54) = "N4frCnBpl = " & .N4frCnBpl & ""
            'frmViewData.rptInputData.Formulas(55) = "N5frCnBpl = " & .N5frCnBpl & ""
            'frmViewData.rptInputData.Formulas(56) = "SrCnBpl = " & .SrCnBpl & ""
            'frmViewData.rptInputData.Formulas(57) = "CrCnBpl = " & .CrCnBpl & ""
            'frmViewData.rptInputData.Formulas(58) = "UcrCnBpl = " & .UcrCnBpl & ""
            'frmViewData.rptInputData.Formulas(59) = "N1faCnBpl = " & .N1faCnBpl & ""
            'frmViewData.rptInputData.Formulas(60) = "N2faCnBpl = " & .N2faCnBpl & ""
            'frmViewData.rptInputData.Formulas(61) = "N3faCnBpl = " & .N3faCnBpl & ""
            'frmViewData.rptInputData.Formulas(62) = "CaCnBpl = " & .CaCnBpl & ""

            'frmViewData.rptInputData.Formulas(63) = "N1frTlBpl = " & .N1frTlBpl & ""
            'frmViewData.rptInputData.Formulas(64) = "N2frTlBpl = " & .N2frTlBpl & ""
            'frmViewData.rptInputData.Formulas(65) = "N3frTlBpl = " & .N3frTlBpl & ""
            'frmViewData.rptInputData.Formulas(66) = "N4frTlBpl = " & .N4frTlBpl & ""
            'frmViewData.rptInputData.Formulas(67) = "N5frTlBpl = " & .N5frTlBpl & ""
            'frmViewData.rptInputData.Formulas(68) = "SrTlBpl = " & .SrTlBpl & ""
            'frmViewData.rptInputData.Formulas(69) = "CrTlBpl = " & .CrTlBpl & ""
            'frmViewData.rptInputData.Formulas(70) = "UcrTlBpl = " & .UcrTlBpl & ""
            'frmViewData.rptInputData.Formulas(71) = "N1faTlBpl = " & .N1faTlBpl & ""
            'frmViewData.rptInputData.Formulas(72) = "N2faTlBpl = " & .N2faTlBpl & ""
            'frmViewData.rptInputData.Formulas(73) = "N3faTlBpl = " & .N3faTlBpl & ""
            'frmViewData.rptInputData.Formulas(74) = "CaTlBpl = " & .CaTlBpl & ""

            'frmViewData.rptInputData.Formulas(75) = "N1frCnTonsAdj = " & .N1frCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(76) = "N2frCnTonsAdj = " & .N2FrCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(77) = "N3frCnTonsAdj = " & .N3FrCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(78) = "N4frCnTonsAdj = " & .N4FrCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(79) = "N5frCnTonsAdj = " & .N5FrCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(80) = "SrCnTonsAdj = " & .SrCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(81) = "CrCnTonsAdj = " & .CrCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(82) = "UcrCnTonsAdj = " & .UcrCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(83) = "N1faCnTonsAdj = " & .N1faCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(84) = "N2faCnTonsAdj = " & .N2faCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(85) = "N3faCnTonsAdj = " & .N3faCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(86) = "CaCnTonsAdj = " & .CaCnTonsAdj & ""

            'frmViewData.rptInputData.Formulas(87) = "N1frTlTonsAdj = " & .N1frTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(88) = "N2frTlTonsAdj = " & .N2frTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(89) = "N3frTlTonsAdj = " & .N3frTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(90) = "N4frTlTonsAdj = " & .N4frTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(91) = "N5frTlTonsAdj = " & .N5frTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(92) = "SrTlTonsAdj = " & .SrTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(93) = "CrTlTonsAdj = " & .CrTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(94) = "UcrTlTonsAdj = " & .UcrTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(95) = "N1faTlTonsAdj = " & .N1faTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(96) = "N2faTlTonsAdj = " & .N2faTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(97) = "N3faTlTonsAdj = " & .N3faTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(98) = "CaTlTonsAdj = " & .CaTlTonsAdj & ""

            'frmViewData.rptInputData.Formulas(99) = "SumRghrFdBpl = " & .SumRghrFdBpl & ""
            'frmViewData.rptInputData.Formulas(100) = "SumRghrFdTonsRpt = " & .SumRghrFdTonsRpt & ""
            'frmViewData.rptInputData.Formulas(101) = "SumRghrFdTonsAdj = " & .SumRghrFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(102) = "SumRghrCnTonsAdj = " & .SumRghrCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(103) = "SumRghrTlTonsAdj = " & .SumRghrTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(104) = "SumClnrFdTonsRpt = " & .SumClnrFdTonsRpt & ""
            'frmViewData.rptInputData.Formulas(105) = "SumClnrFdTonsAdj = " & .SumClnrFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(106) = "SumClnrCnTonsAdj = " & .SumClnrCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(107) = "SumClnrTlTonsAdj = " & .SumClnrTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(108) = "SumAllFdTonsAdj = " & .SumAllFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(109) = "SumAllCnTonsAdj = " & .SumAllCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(110) = "SumAllTlTonsAdj = " & .SumAllTlTonsAdj & ""

            'frmViewData.rptInputData.Formulas(111) = "ArGmtFdBpl = " & .ArGmtFdBpl & ""
            'frmViewData.rptInputData.Formulas(112) = "ArGmtCnBpl = " & .ArGmtCnBpl & ""
            'frmViewData.rptInputData.Formulas(113) = "ArGmtTlBpl = " & .ArGmtTlBpl & ""
            'frmViewData.rptInputData.Formulas(114) = "ArGmtRc = " & .ArGmtRc & ""
            'frmViewData.rptInputData.Formulas(115) = "ArGmtRcvry = " & .ArGmtRcvry & ""

            'frmViewData.rptInputData.Formulas(116) = "ClcGmtFdTons = " & .ClcGmtFdTons & ""
            'frmViewData.rptInputData.Formulas(117) = "ClcGmtCnTons = " & .ClcGmtCnTons & ""
            'frmViewData.rptInputData.Formulas(118) = "ClcGmtFdBpl = " & .ClcGmtFdBpl & ""
            'frmViewData.rptInputData.Formulas(119) = "ClcGmtCnBpl = " & .ClcGmtCnBpl & ""
            'frmViewData.rptInputData.Formulas(120) = "ClcGmtTlBpl = " & .ClcGmtTlBpl & ""
            'frmViewData.rptInputData.Formulas(121) = "ClcGmtRc = " & .ClcGmtRc & ""
            'frmViewData.rptInputData.Formulas(122) = "ClcGmtRcvry = " & .ClcGmtRcvry & ""

            'frmViewData.rptInputData.Formulas(123) = "ArFdTonsFdTons = " & .ArFdTonsFdTons & ""
            'frmViewData.rptInputData.Formulas(124) = "ArFdTonsCnTons = " & .ArFdTonsCnTons & ""
            'frmViewData.rptInputData.Formulas(125) = "ArFdTonsFdBpl = " & .ArFdTonsFdBpl & ""
            'frmViewData.rptInputData.Formulas(126) = "ArFdTonsCnBpl = " & .ArFdTonsCnBpl & ""
            'frmViewData.rptInputData.Formulas(127) = "ArFdTonsTlBpl = " & .ArFdTonsTlBpl & ""
            'frmViewData.rptInputData.Formulas(128) = "ArFdTonsRc = " & .ArFdTonsRc & ""
            'frmViewData.rptInputData.Formulas(129) = "ArFdTonsRcvry = " & .ArFdTonsRcvry & ""

            'frmViewData.rptInputData.Formulas(130) = "GmtFromCircs = " & .GmtFromCircs & ""

            'frmViewData.rptInputData.Formulas(131) = "N1frRc = " & .N1frRc & ""
            'frmViewData.rptInputData.Formulas(132) = "N2frRc = " & .N2frRc & ""
            'frmViewData.rptInputData.Formulas(133) = "N3frRc = " & .N3frRc & ""
            'frmViewData.rptInputData.Formulas(134) = "N4frRc = " & .N4frRc & ""
            'frmViewData.rptInputData.Formulas(135) = "N5frRc = " & .N5frRc & ""
            'frmViewData.rptInputData.Formulas(136) = "SrRc = " & .SrRc & ""
            'frmViewData.rptInputData.Formulas(137) = "CrRc = " & .CrRc & ""
            'frmViewData.rptInputData.Formulas(138) = "UcrRc = " & .UcrRc & ""
            'frmViewData.rptInputData.Formulas(139) = "N1faRc = " & .N1faRc & ""
            'frmViewData.rptInputData.Formulas(140) = "N2faRc = " & .N2faRc & ""
            'frmViewData.rptInputData.Formulas(141) = "N3faRc = " & .N3faRc & ""
            'frmViewData.rptInputData.Formulas(142) = "CaRc = " & .CaRc & ""

            'frmViewData.rptInputData.Formulas(143) = "N1frAr = " & .N1frAr & ""
            'frmViewData.rptInputData.Formulas(144) = "N2frAr = " & .N2frAr & ""
            'frmViewData.rptInputData.Formulas(145) = "N3frAr = " & .N3frAr & ""
            'frmViewData.rptInputData.Formulas(146) = "N4frAr = " & .N4frAr & ""
            'frmViewData.rptInputData.Formulas(147) = "N5frAr = " & .N5frAr & ""
            'frmViewData.rptInputData.Formulas(148) = "SrAr = " & .SrAr & ""
            'frmViewData.rptInputData.Formulas(149) = "CrAr = " & .CrAr & ""
            'frmViewData.rptInputData.Formulas(150) = "UcrAr = " & .UcrAr & ""
            'frmViewData.rptInputData.Formulas(151) = "N1faAr = " & .N1faAr & ""
            'frmViewData.rptInputData.Formulas(152) = "N2faAr = " & .N2faAr & ""
            'frmViewData.rptInputData.Formulas(153) = "N3faAr = " & .N3faAr & ""
            'frmViewData.rptInputData.Formulas(154) = "CaAr = " & .CaAr & ""

            'frmViewData.rptInputData.Formulas(155) = "N1frSr = " & .N1frSr & ""
            'frmViewData.rptInputData.Formulas(156) = "N2frSr = " & .N2frSr & ""
            'frmViewData.rptInputData.Formulas(157) = "N3frSr = " & .N3frSr & ""
            'frmViewData.rptInputData.Formulas(158) = "N4frSr = " & .N4frSr & ""
            'frmViewData.rptInputData.Formulas(159) = "N5frSr = " & .N5frSr & ""
            'frmViewData.rptInputData.Formulas(160) = "SrSr = " & .SrSr & ""
            'frmViewData.rptInputData.Formulas(161) = "CrSr = " & .CrSr & ""
            'frmViewData.rptInputData.Formulas(162) = "UcrSr = " & .UcrSr & ""
            'frmViewData.rptInputData.Formulas(163) = "N1faSr = " & .N1faSr & ""
            'frmViewData.rptInputData.Formulas(164) = "N2faSr = " & .N2faSr & ""
            'frmViewData.rptInputData.Formulas(165) = "N3faSr = " & .N3faSr & ""
            'frmViewData.rptInputData.Formulas(166) = "CaSr = " & .CaSr & ""

            'frmViewData.rptInputData.Formulas(167) = "PrdFcnTons = " & .PrdFcnTons & ""
            'frmViewData.rptInputData.Formulas(168) = "PrdFcnBpl = " & .PrdFcnBpl & ""
            'frmViewData.rptInputData.Formulas(169) = "PrdCcnTons = " & .PrdCcnTons & ""
            'frmViewData.rptInputData.Formulas(170) = "PrdCcnBpl = " & .PrdCcnBpl & ""
            'frmViewData.rptInputData.Formulas(171) = "PrdUccnTons = " & .PrdUccnTons & ""
            'frmViewData.rptInputData.Formulas(172) = "PrdUccnBpl = " & .PrdUccnBpl & ""
            'frmViewData.rptInputData.Formulas(173) = "PrdCnTons = " & .PrdCnTons & ""
            'frmViewData.rptInputData.Formulas(174) = "PrdCnBpl = " & .PrdCnBpl & ""

            If .CcnTonsEst = True Then
                CcnTonsEst = 1
            Else
                CcnTonsEst = 0
            End If
            'frmViewData.rptInputData.Formulas(175) = "PrdCcnTonsEst = " & CcnTonsEst & ""
        End With

        'Need to pass the company name into the report
        'frmViewData.rptInputData.ParameterFields(0) = "pCompanyName;" & gCompanyName & ";TRUE"

        'Have all the needed data -- start the report
        'frmViewData.rptInputData.ReportFileName = gPath + "\Reports\" + _
        '"MassBalanceSF.rpt"

        'Connect to Oracle database
        ConnectString = "DSN = " + gDataSource + ";UID = " + gOracleUserName + _
            ";PWD = " + gOracleUserPassword + ";DSQ = "

        'frmViewData.rptInputData.Connect = ConnectString
        ''Report window maximized
        'frmViewData.rptInputData.WindowState = crptMaximized

        'frmViewData.rptInputData.WindowTitle = "Prospect Split Summary"

        ''User not allowed to minimize report window
        'frmViewData.rptInputData.WindowMinButton = False

        ''Start Crystal Reports
        'frmViewData.rptInputData.action = 1

        'frmViewData.rptInputData.ReportFileName = ""
        'frmViewData.rptInputData.Reset()

        Exit Function

gMassBalanceSFError:

        MsgBox("Error printing South Fort Meade Mass Balance report." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "South Fort Meade Mass Balance Reporting Error")

    End Function

    Public Function gMetallurgicalSF(ByVal aBeginDate As Date, _
                                     ByVal aBeginShift As String, _
                                     ByVal aEndDate As Date, _
                                     ByVal aEndShift As String, _
                                     ByVal aCrewNumber As String, _
                                     ByVal aSkipDownMonths As Integer) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade

        'This function will return the number of shifts processed.
        'It will also "return" data through the fFloatPlantData array.
        On Error GoTo gMetallurgicalSFError

        Dim ConnectString As String
        Dim RowIdx As Integer
        Dim TimeFrame As String
        Dim NumShifts As Integer

        Dim FloatPlantCirc() As Object
        Dim FloatPlantGmt() As Object

        Dim GotReagents As Boolean

        Dim TotCnTons As Long
        Dim TotAdjFdTons As Long
        Dim TotRptFdTons As Long

        Dim CcnTonsEst As Double

        'frmViewData.rptInputData.Reset()

        ZeroSfSummingData()

        'Miscellaneous data setup
        'frmViewData.rptInputData.Formulas(0) = "MineName = '" & "South Fort Meade" & "'"
        If aBeginDate = aEndDate And aBeginShift = aEndShift Then
            TimeFrame = aBeginDate & " " & _
                        StrConv(aBeginShift, vbProperCase) & " Shift"
        Else
            TimeFrame = aBeginDate & " " & _
                        StrConv(aBeginShift, vbProperCase) & _
                        " Shift" & " thru " & _
                        aEndDate & " " & _
                        StrConv(aEndShift, vbProperCase) & " Shift"
        End If
        'frmViewData.rptInputData.Formulas(1) = "TimeFrame = '" & TimeFrame & "'"
        'frmViewData.rptInputData.Formulas(2) = "CrewNumber = '" & aCrewNumber & "'"

        'Get data for float plant metallurgical (mass balance data)
        'Get data for float plant metallurgical (mass balance data)
        'Get data for float plant metallurgical (mass balance data)

        NumShifts = gGetSfFloatPlantBalanceData(FloatPlantCirc, _
                                                FloatPlantGmt, _
                                                aBeginDate, _
                                                StrConv(aBeginShift, vbUpperCase), _
                                                aEndDate, _
                                                StrConv(aEndShift, vbUpperCase), _
                                                aCrewNumber, _
                                                aSkipDownMonths)

        'gCCnTonsEst -- if True then coarse concentrate tons have been estimated.

        With MetallurgicalSfRpt
            .N1frFdBpl = Round(MbSfTotal.N1frFdBpl, 2)
            .N1frFdTons = MbSfTotal.N1frFdTonsAdj
            .N1frCnBpl = Round(MbSfTotal.N1frCnBpl, 2)
            .N1frCnTons = MbSfTotal.N1frCnTonsAdj
            .N1frTlBpl = Round(MbSfTotal.N1frTlBpl, 2)
            .N1frTlTons = MbSfTotal.N1frTlTonsAdj
            .N1frRc = MbSfTotal.N1frRc
            .N1frRcvry = MbSfTotal.N1frActPctRcvry

            .N2frFdBpl = Round(MbSfTotal.N2frFdBpl, 2)
            .N2frFdTons = MbSfTotal.N2frFdTonsAdj
            .N2frCnBpl = Round(MbSfTotal.N2frCnBpl, 2)
            .N2frCnTons = MbSfTotal.N2FrCnTonsAdj
            .N2frTlBpl = Round(MbSfTotal.N2frTlBpl, 2)
            .N2frTlTons = MbSfTotal.N2frTlTonsAdj
            .N2frRc = MbSfTotal.N2frRc
            .N2frRcvry = MbSfTotal.N2frActPctRcvry

            .N3frFdBpl = Round(MbSfTotal.N3frFdBpl, 2)
            .N3frFdTons = MbSfTotal.N3frFdTonsAdj
            .N3frCnBpl = Round(MbSfTotal.N3frCnBpl, 2)
            .N3frCnTons = MbSfTotal.N3FrCnTonsAdj
            .N3frTlBpl = Round(MbSfTotal.N3frTlBpl, 2)
            .N3frTlTons = MbSfTotal.N3frTlTonsAdj
            .N3frRc = MbSfTotal.N3frRc
            .N3frRcvry = MbSfTotal.N3frActPctRcvry

            .N4frFdBpl = Round(MbSfTotal.N4frFdBpl, 2)
            .N4frFdTons = MbSfTotal.N4frFdTonsAdj
            .N4frCnBpl = Round(MbSfTotal.N4frCnBpl, 2)
            .N4frCnTons = MbSfTotal.N4FrCnTonsAdj
            .N4frTlBpl = Round(MbSfTotal.N4frTlBpl, 2)
            .N4frTlTons = MbSfTotal.N4frTlTonsAdj
            .N4frRc = MbSfTotal.N4frRc
            .N4frRcvry = MbSfTotal.N4frActPctRcvry

            .N5frFdBpl = Round(MbSfTotal.N5frFdBpl, 2)
            .N5frFdTons = MbSfTotal.N5frFdTonsAdj
            .N5frCnBpl = Round(MbSfTotal.N5frCnBpl, 2)
            .N5frCnTons = MbSfTotal.N5FrCnTonsAdj
            .N5frTlBpl = Round(MbSfTotal.N5frTlBpl, 2)
            .N5frTlTons = MbSfTotal.N5frTlTonsAdj
            .N5frRc = MbSfTotal.N5frRc
            .N5frRcvry = MbSfTotal.N5frActPctRcvry

            .FrFdBplAvg = Round(MbSfTotal.TotFineRghrFdBpl, 2)
            .FrFdTonsAvg = MbSfTotal.TotFineRghrFdTons
            .FrCnBplAvg = Round(MbSfTotal.TotFineRghrCnBpl, 2)
            .FrCnTonsAvg = MbSfTotal.TotFineRghrCnTons
            .FrTlBplAvg = Round(MbSfTotal.TotFineRghrTlBpl, 2)
            .FrTlTonsAvg = MbSfTotal.TotFineRghrTlTons
            .FrRcAvg = MbSfTotal.TotFineRghrRc
            .FrRcvryAvg = MbSfTotal.TotFineRghrRcvry

            .SrFdBpl = Round(MbSfTotal.SrFdBpl, 2)
            .SrFdTons = MbSfTotal.SrFdTonsAdj
            .SrCnBpl = Round(MbSfTotal.SrCnBpl, 2)
            .SrCnTons = MbSfTotal.SrCnTonsAdj
            .SrTlBpl = Round(MbSfTotal.SrTlBpl, 2)
            .SrTlTons = MbSfTotal.SrTlTonsAdj
            .SrRc = MbSfTotal.SrRc
            .SrRcvry = MbSfTotal.SrActPctRcvry

            .CrFdBpl = Round(MbSfTotal.CrFdBpl, 2)
            .CrFdTons = MbSfTotal.CrFdTonsAdj
            .CrCnBpl = Round(MbSfTotal.CrCnBpl, 2)
            .CrCnTons = MbSfTotal.CrCnTonsAdj
            .CrTlBpl = Round(MbSfTotal.CrTlBpl, 2)
            .CrTlTons = MbSfTotal.CrTlTonsAdj
            .CrRc = MbSfTotal.CrRc
            .CrRcvry = MbSfTotal.CrActPctRcvry

            .CrFdBplAvg = Round(MbSfTotal.TotCrsRghrFdBpl, 2)
            .CrFdTonsAvg = MbSfTotal.TotCrsRghrFdTons
            .CrCnBplAvg = Round(MbSfTotal.TotCrsRghrCnBpl, 2)
            .CrCnTonsAvg = MbSfTotal.TotCrsRghrCnTons
            .CrTlBplAvg = Round(MbSfTotal.TotCrsRghrTlBpl, 2)
            .CrTlTonsAvg = MbSfTotal.TotCrsRghrTlTons
            .CrRcAvg = MbSfTotal.TotCrsRghrRc
            .CrRcvryAvg = MbSfTotal.TotCrsRghrRcvry

            .FrSrCrFdBplAvg = Round(MbSfTotal.TotRghr2FdBpl, 2)
            .FrSrCrFdTonsAvg = MbSfTotal.TotRghr2FdTons
            .FrSrCrCnBplAvg = Round(MbSfTotal.TotRghr2CnBpl, 2)
            .FrSrCrCnTonsAvg = MbSfTotal.TotRghr2CnTons
            .FrSrCrTlBplAvg = Round(MbSfTotal.TotRghr2TlBpl, 2)
            .FrSrCrTlTonsAvg = MbSfTotal.TotRghr2TlTons
            .FrSrCrRcAvg = MbSfTotal.TotRghr2Rc
            .FrSrCrRcvryAvg = MbSfTotal.TotRghr2Rcvry

            .N1faFdBpl = Round(MbSfTotal.N1faFdBpl, 2)
            .N1faFdTons = MbSfTotal.N1faFdTonsAdj
            .N1faCnBpl = Round(MbSfTotal.N1faCnBpl, 2)
            .N1faCnTons = MbSfTotal.N1faCnTons
            .N1faTlBpl = Round(MbSfTotal.N1faTlBpl, 2)
            .N1faTlTons = MbSfTotal.N1faTlTonsAdj
            .N1faRc = MbSfTotal.N1faRc
            .N1faRcvry = MbSfTotal.N1faActPctRcvry

            .N2faFdBpl = Round(MbSfTotal.N2faFdBpl, 2)
            .N2faFdTons = MbSfTotal.N2faFdTonsAdj
            .N2faCnBpl = Round(MbSfTotal.N2faCnBpl, 2)
            .N2faCnTons = MbSfTotal.N2faCnTons
            .N2faTlBpl = Round(MbSfTotal.N2faTlBpl, 2)
            .N2faTlTons = MbSfTotal.N2faTlTonsAdj
            .N2faRc = MbSfTotal.N2faRc
            .N2faRcvry = MbSfTotal.N2faActPctRcvry

            .N3faFdBpl = Round(MbSfTotal.N3faFdBpl, 2)
            .N3faFdTons = MbSfTotal.N3faFdTonsAdj
            .N3faCnBpl = Round(MbSfTotal.N3faCnBpl, 2)
            .N3faCnTons = MbSfTotal.N3faCnTons
            .N3faTlBpl = Round(MbSfTotal.N3faTlBpl, 2)
            .N3faTlTons = MbSfTotal.N3faTlTonsAdj
            .N3faRc = MbSfTotal.N3faRc
            .N3faRcvry = MbSfTotal.N3faActPctRcvry

            .FaFdBplAvg = Round(MbSfTotal.TotFineClnrFdBpl, 2)
            .FaFdTonsAvg = MbSfTotal.TotFineClnrFdTons
            .FaCnBplAvg = Round(MbSfTotal.TotFineClnrCnBpl, 2)
            .FaCnTonsAvg = MbSfTotal.TotFineClnrCnTons
            .FaTlBplAvg = Round(MbSfTotal.TotFineClnrTlBpl, 2)
            .FaTlTonsAvg = MbSfTotal.TotFineClnrTlTons
            .FaRcAvg = MbSfTotal.TotFineClnrRc
            .FaRcvryAvg = MbSfTotal.TotFineClnrRcvry

            .CaFdBpl = Round(MbSfTotal.CaFdBpl, 2)
            .CaFdTons = MbSfTotal.CaFdTonsAdj
            .CaCnBpl = Round(MbSfTotal.CaCnBpl, 2)
            .CaCnTons = MbSfTotal.CaCnTons
            .CaTlBpl = Round(MbSfTotal.CaTlBpl, 2)
            .CaTlTons = MbSfTotal.CaTlTonsAdj
            .CaRc = MbSfTotal.CaRc
            .CaRcvry = MbSfTotal.CaActPctRcvry

            .ClnrFdBplAvg = Round(MbSfTotal.TotClnrFdBpl, 2)
            .ClnrFdTonsAvg = MbSfTotal.TotClnrFdTons
            .ClnrCnBplAvg = Round(MbSfTotal.TotClnrCnBpl, 2)
            .ClnrCnTonsAvg = MbSfTotal.TotClnrCnTons
            .ClnrTlBplAvg = Round(MbSfTotal.TotClnrTlBpl, 2)
            .ClnrTlTonsAvg = MbSfTotal.TotClnrTlTons
            .ClnrRcAvg = MbSfTotal.TotClnrRc
            .ClnrRcvryAvg = MbSfTotal.TotClnrRcvry

            .PltFdBplAvg = Round(MbSfTotal.TotPlantFdBpl, 2)
            .PltFdTonsAvg = MbSfTotal.TotPlantFdTons
            .PltCnBplAvg = Round(MbSfTotal.TotPlantCnBpl, 2)
            .PltCnTonsAvg = MbSfTotal.TotPlantCnTons
            .PltTlBplAvg = Round(MbSfTotal.TotPlantTlBpl, 2)
            .PltTlTonsAvg = MbSfTotal.TotPlantTlTons
            .PltRcAvg = MbSfTotal.TotPlantRc
            .PltRcvryAvg = MbSfTotal.TotPlantRcvry

            .UcrFdBpl = Round(MbSfTotal.UcrFdBpl, 2)
            .UcrFdTons = MbSfTotal.UcrFdTonsAdj
            .UcrCnBpl = Round(MbSfTotal.UcrCnBpl, 2)
            .UcrCnTons = MbSfTotal.UcrCnTonsAdj
            .UcrTlBpl = Round(MbSfTotal.UcrTlBpl, 2)
            .UcrTlTons = MbSfTotal.UcrTlTonsAdj
            .UcrRc = MbSfTotal.UcrRc
            .UcrRcvry = MbSfTotal.UcrActPctRcvry

            .CombFdBplAvg = Round(MbSfTotal.TotCombFdBpl, 2)
            .CombFdTonsAvg = MbSfTotal.TotCombFdTons
            .CombCnBplAvg = Round(MbSfTotal.TotCombCnBpl, 2)
            .CombCnTonsAvg = MbSfTotal.TotCombCnTons
            .CombTlBplAvg = Round(MbSfTotal.TotCombTlBpl, 2)
            .CombTlTonsAvg = MbSfTotal.TotCombTlTons
            .CombRcAvg = MbSfTotal.TotCombRc
            .CombRcvryAvg = MbSfTotal.TotCombRcvry

            .RgTotRptFdTons = MbSfTotal.TotFdTonsRpt
            .RgTotCnTons = MbSfTotal.TotPlantCnTons
            .RgTotAdjFdTons = MbSfTotal.TotPlantFdTons

            .ReportedGmtBpl = Round(MbSfTotal.TotTlBplMsrd, 2)

            .CcnTonsEst = gCcnTonsEst
        End With

        'Now get the reagent data for the Metallurgical Report
        'Now get the reagent data for the Metallurgical Report
        'Now get the reagent data for the Metallurgical Report

        TotAdjFdTons = MetallurgicalSfRpt.CombFdTonsAvg
        TotCnTons = MetallurgicalSfRpt.CombCnTonsAvg
        TotRptFdTons = MbSfTotal.TotFdTonsRpt

        GotReagents = gGetMetReagentDataSf(aBeginDate, aBeginShift, _
                                           aEndDate, aEndShift, _
                                           aCrewNumber, TotAdjFdTons, _
                                           TotRptFdTons, TotCnTons)

        With MetallurgicalSfRpt
            'frmViewData.rptInputData.Formulas(3) = "N1frFdBpl = " & .N1frFdBpl & ""
            'frmViewData.rptInputData.Formulas(4) = "N1frFdTons = " & .N1frFdTons & ""
            'frmViewData.rptInputData.Formulas(5) = "N1frCnBpl = " & .N1frCnBpl & ""
            'frmViewData.rptInputData.Formulas(6) = "N1frCnTons = " & .N1frCnTons & ""
            'frmViewData.rptInputData.Formulas(7) = "N1frTlBpl = " & .N1frTlBpl & ""
            'frmViewData.rptInputData.Formulas(8) = "N1frTlTons = " & .N1frTlTons & ""
            'frmViewData.rptInputData.Formulas(9) = "N1frRc = " & .N1frRc & ""
            'frmViewData.rptInputData.Formulas(10) = "N1frRcvry = " & .N1frRcvry & ""

            'frmViewData.rptInputData.Formulas(11) = "N2frFdBpl = " & .N2frFdBpl & ""
            'frmViewData.rptInputData.Formulas(12) = "N2frFdTons = " & .N2frFdTons & ""
            'frmViewData.rptInputData.Formulas(13) = "N2frCnBpl = " & .N2frCnBpl & ""
            'frmViewData.rptInputData.Formulas(14) = "N2frCnTons = " & .N2frCnTons & ""
            'frmViewData.rptInputData.Formulas(15) = "N2frTlBpl = " & .N2frTlBpl & ""
            'frmViewData.rptInputData.Formulas(16) = "N2frTlTons = " & .N2frTlTons & ""
            'frmViewData.rptInputData.Formulas(17) = "N2frRc = " & .N2frRc & ""
            'frmViewData.rptInputData.Formulas(18) = "N2frRcvry = " & .N2frRcvry & ""

            'frmViewData.rptInputData.Formulas(19) = "N3frFdBpl = " & .N3frFdBpl & ""
            'frmViewData.rptInputData.Formulas(20) = "N3frFdTons = " & .N3frFdTons & ""
            'frmViewData.rptInputData.Formulas(21) = "N3frCnBpl = " & .N3frCnBpl & ""
            'frmViewData.rptInputData.Formulas(22) = "N3frCnTons = " & .N3frCnTons & ""
            'frmViewData.rptInputData.Formulas(23) = "N3frTlBpl = " & .N3frTlBpl & ""
            'frmViewData.rptInputData.Formulas(24) = "N3frTlTons = " & .N3frTlTons & ""
            'frmViewData.rptInputData.Formulas(25) = "N3frRc = " & .N3frRc & ""
            'frmViewData.rptInputData.Formulas(26) = "N3frRcvry = " & .N3frRcvry & ""

            'frmViewData.rptInputData.Formulas(27) = "N4frFdBpl = " & .N4frFdBpl & ""
            'frmViewData.rptInputData.Formulas(28) = "N4frFdTons = " & .N4frFdTons & ""
            'frmViewData.rptInputData.Formulas(29) = "N4frCnBpl = " & .N4frCnBpl & ""
            'frmViewData.rptInputData.Formulas(30) = "N4frCnTons = " & .N4frCnTons & ""
            'frmViewData.rptInputData.Formulas(31) = "N4frTlBpl = " & .N4frTlBpl & ""
            'frmViewData.rptInputData.Formulas(32) = "N4frTlTons = " & .N4frTlTons & ""
            'frmViewData.rptInputData.Formulas(33) = "N4frRc = " & .N4frRc & ""
            'frmViewData.rptInputData.Formulas(34) = "N4frRcvry = " & .N4frRcvry & ""

            'frmViewData.rptInputData.Formulas(35) = "N5frFdBpl = " & .N5frFdBpl & ""
            'frmViewData.rptInputData.Formulas(36) = "N5frFdTons = " & .N5frFdTons & ""
            'frmViewData.rptInputData.Formulas(37) = "N5frCnBpl = " & .N5frCnBpl & ""
            'frmViewData.rptInputData.Formulas(38) = "N5frCnTons = " & .N5frCnTons & ""
            'frmViewData.rptInputData.Formulas(39) = "N5frTlBpl = " & .N5frTlBpl & ""
            'frmViewData.rptInputData.Formulas(40) = "N5frTlTons = " & .N5frTlTons & ""
            'frmViewData.rptInputData.Formulas(41) = "N5frRc = " & .N5frRc & ""
            'frmViewData.rptInputData.Formulas(42) = "N5frRcvry = " & .N5frRcvry & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(43) = "FrFdBplAvg = " & .FrFdBplAvg & ""
            'frmViewData.rptInputData.Formulas(44) = "FrFdTonsAvg = " & .FrFdTonsAvg & ""
            'frmViewData.rptInputData.Formulas(45) = "FrCnBplAvg = " & .FrCnBplAvg & ""
            'frmViewData.rptInputData.Formulas(46) = "FrCnTonsAvg = " & .FrCnTonsAvg & ""
            'frmViewData.rptInputData.Formulas(47) = "FrTlBplAvg = " & .FrTlBplAvg & ""
            'frmViewData.rptInputData.Formulas(48) = "FrTlTonsAvg = " & .FrTlTonsAvg & ""
            'frmViewData.rptInputData.Formulas(49) = "FrRcAvg = " & .FrRcAvg & ""
            'frmViewData.rptInputData.Formulas(50) = "FrRcvryAvg = " & .FrRcvryAvg & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(51) = "SrFdBpl = " & .SrFdBpl & ""
            'frmViewData.rptInputData.Formulas(52) = "SrFdTons = " & .SrFdTons & ""
            'frmViewData.rptInputData.Formulas(53) = "SrCnBpl = " & .SrCnBpl & ""
            'frmViewData.rptInputData.Formulas(54) = "SrCnTons = " & .SrCnTons & ""
            'frmViewData.rptInputData.Formulas(55) = "SrTlBpl = " & .SrTlBpl & ""
            'frmViewData.rptInputData.Formulas(56) = "SrTlTons = " & .SrTlTons & ""
            'frmViewData.rptInputData.Formulas(57) = "SrRc = " & .SrRc & ""
            'frmViewData.rptInputData.Formulas(58) = "SrRcvry = " & .SrRcvry & ""

            'frmViewData.rptInputData.Formulas(59) = "CrFdBpl = " & .CrFdBpl & ""
            'frmViewData.rptInputData.Formulas(60) = "CrFdTons = " & .CrFdTons & ""
            'frmViewData.rptInputData.Formulas(61) = "CrCnBpl = " & .CrCnBpl & ""
            'frmViewData.rptInputData.Formulas(62) = "CrCnTons = " & .CrCnTons & ""
            'frmViewData.rptInputData.Formulas(63) = "CrTlBpl = " & .CrTlBpl & ""
            'frmViewData.rptInputData.Formulas(64) = "CrTlTons = " & .CrTlTons & ""
            'frmViewData.rptInputData.Formulas(65) = "CrRc = " & .CrRc & ""
            'frmViewData.rptInputData.Formulas(66) = "CrRcvry = " & .CrRcvry & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(67) = "CrFdBplAvg = " & .CrFdBplAvg & ""
            'frmViewData.rptInputData.Formulas(68) = "CrFdTonsAvg = " & .CrFdTonsAvg & ""
            'frmViewData.rptInputData.Formulas(69) = "CrCnBplAvg = " & .CrCnBplAvg & ""
            'frmViewData.rptInputData.Formulas(70) = "CrCnTonsAvg = " & .CrCnTonsAvg & ""
            'frmViewData.rptInputData.Formulas(71) = "CrTlBplAvg = " & .CrTlBplAvg & ""
            'frmViewData.rptInputData.Formulas(72) = "CrTlTonsAvg = " & .CrTlTonsAvg & ""
            'frmViewData.rptInputData.Formulas(73) = "CrRcAvg = " & .CrRcAvg & ""
            'frmViewData.rptInputData.Formulas(74) = "CrRcvryAvg = " & .CrRcvryAvg & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(75) = "FrSrCrFdBplAvg = " & .FrSrCrFdBplAvg & ""
            'frmViewData.rptInputData.Formulas(76) = "FrSrCrFdTonsAvg = " & .FrSrCrFdTonsAvg & ""
            'frmViewData.rptInputData.Formulas(77) = "FrSrCrCnBplAvg = " & .FrSrCrCnBplAvg & ""
            'frmViewData.rptInputData.Formulas(78) = "FrSrCrCnTonsAvg = " & .FrSrCrCnTonsAvg & ""
            'frmViewData.rptInputData.Formulas(79) = "FrSrCrTlBplAvg = " & .FrSrCrTlBplAvg & ""
            'frmViewData.rptInputData.Formulas(80) = "FrSrCrTlTonsAvg = " & .FrSrCrTlTonsAvg & ""
            'frmViewData.rptInputData.Formulas(81) = "FrSrCrRcAvg = " & .FrSrCrRcAvg & ""
            'frmViewData.rptInputData.Formulas(82) = "FrSrCrRcvryAvg = " & .FrSrCrRcvryAvg & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(83) = "N1faFdBpl = " & .N1faFdBpl & ""
            'frmViewData.rptInputData.Formulas(84) = "N1faFdTons = " & .N1faFdTons & ""
            'frmViewData.rptInputData.Formulas(85) = "N1faCnBpl = " & .N1faCnBpl & ""
            'frmViewData.rptInputData.Formulas(86) = "N1faCnTons = " & .N1faCnTons & ""
            'frmViewData.rptInputData.Formulas(87) = "N1faTlBpl = " & .N1faTlBpl & ""
            'frmViewData.rptInputData.Formulas(88) = "N1faTlTons = " & .N1faTlTons & ""
            'frmViewData.rptInputData.Formulas(89) = "N1faRc = " & .N1faRc & ""
            'frmViewData.rptInputData.Formulas(90) = "N1faRcvry = " & .N1faRcvry & ""

            'frmViewData.rptInputData.Formulas(91) = "N2faFdBpl = " & .N2faFdBpl & ""
            'frmViewData.rptInputData.Formulas(92) = "N2faFdTons = " & .N2faFdTons & ""
            'frmViewData.rptInputData.Formulas(93) = "N2faCnBpl = " & .N2faCnBpl & ""
            'frmViewData.rptInputData.Formulas(94) = "N2faCnTons = " & .N2faCnTons & ""
            'frmViewData.rptInputData.Formulas(95) = "N2faTlBpl = " & .N2faTlBpl & ""
            'frmViewData.rptInputData.Formulas(96) = "N2faTlTons = " & .N2faTlTons & ""
            'frmViewData.rptInputData.Formulas(97) = "N2faRc = " & .N2faRc & ""
            'frmViewData.rptInputData.Formulas(98) = "N2faRcvry = " & .N2faRcvry & ""

            'frmViewData.rptInputData.Formulas(99) = "N3faFdBpl = " & .N3faFdBpl & ""
            'frmViewData.rptInputData.Formulas(100) = "N3faFdTons = " & .N3faFdTons & ""
            'frmViewData.rptInputData.Formulas(101) = "N3faCnBpl = " & .N3faCnBpl & ""
            'frmViewData.rptInputData.Formulas(102) = "N3faCnTons = " & .N3faCnTons & ""
            'frmViewData.rptInputData.Formulas(103) = "N3faTlBpl = " & .N3faTlBpl & ""
            'frmViewData.rptInputData.Formulas(104) = "N3faTlTons = " & .N3faTlTons & ""
            'frmViewData.rptInputData.Formulas(105) = "N3faRc = " & .N3faRc & ""
            'frmViewData.rptInputData.Formulas(106) = "N3faRcvry = " & .N3faRcvry & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(107) = "FaFdBplAvg = " & .FaFdBplAvg & ""
            'frmViewData.rptInputData.Formulas(108) = "FaFdTonsAvg = " & .FaFdTonsAvg & ""
            'frmViewData.rptInputData.Formulas(109) = "FaCnBplAvg = " & .FaCnBplAvg & ""
            'frmViewData.rptInputData.Formulas(110) = "FaCnTonsAvg = " & .FaCnTonsAvg & ""
            'frmViewData.rptInputData.Formulas(111) = "FaTlBplAvg = " & .FaTlBplAvg & ""
            'frmViewData.rptInputData.Formulas(112) = "FaTlTonsAvg = " & .FaTlTonsAvg & ""
            'frmViewData.rptInputData.Formulas(113) = "FaRcAvg = " & .FaRcAvg & ""
            'frmViewData.rptInputData.Formulas(114) = "FaRcvryAvg = " & .FaRcvryAvg & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(115) = "CaFdBpl = " & .CaFdBpl & ""
            'frmViewData.rptInputData.Formulas(116) = "CaFdTons = " & .CaFdTons & ""
            'frmViewData.rptInputData.Formulas(117) = "CaCnBpl = " & .CaCnBpl & ""
            'frmViewData.rptInputData.Formulas(118) = "CaCnTons = " & .CaCnTons & ""
            'frmViewData.rptInputData.Formulas(119) = "CaTlBpl = " & .CaTlBpl & ""
            'frmViewData.rptInputData.Formulas(120) = "CaTlTons = " & .CaTlTons & ""
            'frmViewData.rptInputData.Formulas(121) = "CaRc = " & .CaRc & ""
            'frmViewData.rptInputData.Formulas(122) = "CaRcvry = " & .CaRcvry & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(123) = "ClnrFdBplAvg = " & .ClnrFdBplAvg & ""
            'frmViewData.rptInputData.Formulas(124) = "ClnrFdTonsAvg = " & .ClnrFdTonsAvg & ""
            'frmViewData.rptInputData.Formulas(125) = "ClnrCnBplAvg = " & .ClnrCnBplAvg & ""
            'frmViewData.rptInputData.Formulas(126) = "ClnrCnTonsAvg = " & .ClnrCnTonsAvg & ""
            'frmViewData.rptInputData.Formulas(127) = "ClnrTlBplAvg = " & .ClnrTlBplAvg & ""
            'frmViewData.rptInputData.Formulas(128) = "ClnrTlTonsAvg = " & .ClnrTlTonsAvg & ""
            'frmViewData.rptInputData.Formulas(129) = "ClnrRcAvg = " & .ClnrRcAvg & ""
            'frmViewData.rptInputData.Formulas(130) = "ClnrRcvryAvg = " & .ClnrRcvryAvg & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(131) = "PltFdBplAvg = " & .PltFdBplAvg & ""
            'frmViewData.rptInputData.Formulas(132) = "PltFdTonsAvg = " & .PltFdTonsAvg & ""
            'frmViewData.rptInputData.Formulas(133) = "PltCnBplAvg = " & .PltCnBplAvg & ""
            'frmViewData.rptInputData.Formulas(134) = "PltCnTonsAvg = " & .PltCnTonsAvg & ""
            'frmViewData.rptInputData.Formulas(135) = "PltTlBplAvg = " & .PltTlBplAvg & ""
            'frmViewData.rptInputData.Formulas(136) = "PltTlTonsAvg = " & .PltTlTonsAvg & ""
            'frmViewData.rptInputData.Formulas(137) = "PltRcAvg = " & .PltRcAvg & ""
            'frmViewData.rptInputData.Formulas(138) = "PltRcvryAvg = " & .PltRcvryAvg & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(139) = "UcrFdBpl = " & .UcrFdBpl & ""
            'frmViewData.rptInputData.Formulas(140) = "UcrFdTons = " & .UcrFdTons & ""
            'frmViewData.rptInputData.Formulas(141) = "UcrCnBpl = " & .UcrCnBpl & ""
            'frmViewData.rptInputData.Formulas(142) = "UcrCnTons = " & .UcrCnTons & ""
            'frmViewData.rptInputData.Formulas(143) = "UcrTlBpl = " & .UcrTlBpl & ""
            'frmViewData.rptInputData.Formulas(144) = "UcrTlTons = " & .UcrTlTons & ""
            'frmViewData.rptInputData.Formulas(145) = "UcrRc = " & .UcrRc & ""
            'frmViewData.rptInputData.Formulas(146) = "UcrRcvry = " & .UcrRcvry & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(147) = "CombFdBplAvg = " & .CombFdBplAvg & ""
            'frmViewData.rptInputData.Formulas(148) = "CombFdTonsAvg = " & .CombFdTonsAvg & ""
            'frmViewData.rptInputData.Formulas(149) = "CombCnBplAvg = " & .CombCnBplAvg & ""
            'frmViewData.rptInputData.Formulas(150) = "CombCnTonsAvg = " & .CombCnTonsAvg & ""
            'frmViewData.rptInputData.Formulas(151) = "CombTlBplAvg = " & .CombTlBplAvg & ""
            'frmViewData.rptInputData.Formulas(152) = "CombTlTonsAvg = " & .CombTlTonsAvg & ""
            'frmViewData.rptInputData.Formulas(153) = "CombRcAvg = " & .CombRcAvg & ""
            'frmViewData.rptInputData.Formulas(154) = "CombRcvryAvg = " & .CombRcvryAvg & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(155) = "RgSuTotUnits = " & .RgSuTotUnits & ""
            'frmViewData.rptInputData.Formulas(156) = "RgAmTotUnits = " & .RgAmTotUnits & ""
            'frmViewData.rptInputData.Formulas(157) = "RgSaTotUnits = " & .RgSaTotUnits & ""
            'frmViewData.rptInputData.Formulas(158) = "RgSoTotUnits = " & .RgSoTotUnits & ""
            'frmViewData.rptInputData.Formulas(159) = "RgFaTotUnits = " & .RgFaTotUnits & ""
            'frmViewData.rptInputData.Formulas(160) = "RgFoTotUnits = " & .RgFoTotUnits & ""
            'frmViewData.rptInputData.Formulas(161) = "RgDeTotUnits = " & .RgDeTotUnits & ""
            'frmViewData.rptInputData.Formulas(162) = "RgAllTotUnits = " & .RgAllTotUnits & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(163) = "RgSuTotCost = " & .RgSuTotCost & ""
            'frmViewData.rptInputData.Formulas(164) = "RgAmTotCost = " & .RgAmTotCost & ""
            'frmViewData.rptInputData.Formulas(165) = "RgSaTotCost = " & .RgSaTotCost & ""
            'frmViewData.rptInputData.Formulas(166) = "RgSoTotCost = " & .RgSoTotCost & ""
            'frmViewData.rptInputData.Formulas(167) = "RgFaTotCost = " & .RgFaTotCost & ""
            'frmViewData.rptInputData.Formulas(168) = "RgFoTotCost = " & .RgFoTotCost & ""
            'frmViewData.rptInputData.Formulas(169) = "RgDeTotCost = " & .RgDeTotCost & ""
            'frmViewData.rptInputData.Formulas(170) = "RgAllTotCost = " & .RgAllTotCost & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(171) = "RgSuAdjFdDpt = " & .RgSuAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(172) = "RgAmAdjFdDpt = " & .RgAmAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(173) = "RgSaAdjFdDpt = " & .RgSaAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(174) = "RgSoAdjFdDpt = " & .RgSoAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(175) = "RgFaAdjFdDpt = " & .RgFaAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(176) = "RgFoAdjFdDpt = " & .RgFoAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(177) = "RgDeAdjFdDpt = " & .RgDeAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(178) = "RgAllAdjFdDpt = " & .RgAllAdjFdDpt & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(179) = "RgSuRptFdDpt = " & .RgSuRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(180) = "RgAmRptFdDpt = " & .RgAmRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(181) = "RgSaRptFdDpt = " & .RgSaRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(182) = "RgSoRptFdDpt = " & .RgSoRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(183) = "RgFaRptFdDpt = " & .RgFaRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(184) = "RgFoRptFdDpt = " & .RgFoRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(185) = "RgDeRptFdDpt = " & .RgDeRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(186) = "RgAllRptFdDpt = " & .RgAllRptFdDpt & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(187) = "RgSuCnDpt = " & .RgSuCnDpt & ""
            'frmViewData.rptInputData.Formulas(188) = "RgAmCnDpt = " & .RgAmCnDpt & ""
            'frmViewData.rptInputData.Formulas(189) = "RgSaCnDpt = " & .RgSaCnDpt & ""
            'frmViewData.rptInputData.Formulas(190) = "RgSoCnDpt = " & .RgSoCnDpt & ""
            'frmViewData.rptInputData.Formulas(191) = "RgFaCnDpt = " & .RgFaCnDpt & ""
            'frmViewData.rptInputData.Formulas(192) = "RgFoCnDpt = " & .RgFoCnDpt & ""
            'frmViewData.rptInputData.Formulas(193) = "RgDeCnDpt = " & .RgDeCnDpt & ""
            'frmViewData.rptInputData.Formulas(194) = "RgAllCnDpt = " & .RgAllCnDpt & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(195) = "RgSuAdjFdUpt = " & .RgSuAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(196) = "RgAmAdjFdUpt = " & .RgAmAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(197) = "RgSaAdjFdUpt = " & .RgSaAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(198) = "RgSoAdjFdUpt = " & .RgSoAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(199) = "RgFaAdjFdUpt = " & .RgFaAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(200) = "RgFoAdjFdUpt = " & .RgFoAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(201) = "RgDeAdjFdUpt = " & .RgDeAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(202) = "RgAllAdjFdUpt = " & .RgAllAdjFdUpt & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(203) = "RgSuRptFdUpt = " & .RgSuRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(204) = "RgAmRptFdUpt = " & .RgAmRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(205) = "RgSaRptFdUpt = " & .RgSaRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(206) = "RgSoRptFdUpt = " & .RgSoRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(207) = "RgFaRptFdUpt = " & .RgFaRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(208) = "RgFoRptFdUpt = " & .RgFoRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(209) = "RgDeRptFdUpt = " & .RgDeRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(210) = "RgAllRptFdUpt = " & .RgAllRptFdUpt & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(211) = "RgSuCnUpt = " & .RgSuCnUpt & ""
            'frmViewData.rptInputData.Formulas(212) = "RgAmCnUpt = " & .RgAmCnUpt & ""
            'frmViewData.rptInputData.Formulas(213) = "RgSaCnUpt = " & .RgSaCnUpt & ""
            'frmViewData.rptInputData.Formulas(214) = "RgSoCnUpt = " & .RgSoCnUpt & ""
            'frmViewData.rptInputData.Formulas(215) = "RgFaCnUpt = " & .RgFaCnUpt & ""
            'frmViewData.rptInputData.Formulas(216) = "RgFoCnUpt = " & .RgFoCnUpt & ""
            'frmViewData.rptInputData.Formulas(217) = "RgDeCnUpt = " & .RgDeCnUpt & ""
            'frmViewData.rptInputData.Formulas(218) = "RgAllCnUpt = " & .RgAllCnUpt & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(219) = "RgTotRptFdTons = " & .RgTotRptFdTons & ""

            ''------------------------------

            'frmViewData.rptInputData.Formulas(220) = "ReportedGmtBpl = " & .ReportedGmtBpl & ""

            ''------------------------------

            'If .CcnTonsEst = True Then
            '    CcnTonsEst = 1
            'Else
            '    CcnTonsEst = 0
            'End If
            'frmViewData.rptInputData.Formulas(221) = "PrdCcnTonsEst = " & CcnTonsEst & ""
        End With

        ''Need to pass the company name into the report
        'frmViewData.rptInputData.ParameterFields(0) = "pCompanyName;" & gCompanyName & ";TRUE"

        ''Have all the needed data -- start the report
        'frmViewData.rptInputData.ReportFileName = gPath + "\Reports\" + _
        '                                          "MetallurgicalSF.rpt"

        'Connect to Oracle database
        ConnectString = "DSN = " + gDataSource + ";UID = " + gOracleUserName + _
            ";PWD = " + gOracleUserPassword + ";DSQ = "

        'frmViewData.rptInputData.Connect = ConnectString
        ''Report window maximized
        'frmViewData.rptInputData.WindowState = crptMaximized

        'frmViewData.rptInputData.WindowTitle = "Prospect Split Summary"

        ''User not allowed to minimize report window
        'frmViewData.rptInputData.WindowMinButton = False

        ''Start Crystal Reports
        'frmViewData.rptInputData.action = 1

        Exit Function

gMetallurgicalSFError:

        MsgBox("Error printing South Fort Meade Metallurgical report." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "South Fort Meade Metallurgical Report Printing Error")

    End Function

    Private Sub ProcessSfMassBalanceData()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade

        'Have all data for this shift -- process it!
        On Error GoTo ProcessSfMassBalanceError

        gCcnTonsEst = False

        'Make sure Transfer.CRtoCA is OK
        If Transfer.CRtoN1FA = 0 Then
            Transfer.CRtoCA = 100
        Else
            Transfer.CRtoCA = 100 - Transfer.CRtoN1FA
        End If

        'Transfer.CAtoFC -- 01/13/2003, lss
        'Some or all of the coarse concentrate may have been sent to
        'the fine concentrate bins -- thus we don't have a measure of
        'what the coarse concentrate production for the shift was.
        'The coarse concentrate tons in the coarse concentrate bins
        'is in MbSfShift.PrdCcnTons.
        'Need to determine MbSfShift.CcnTonCorr and MbSfShift.FcnTonCorr.
        gCcnTonsEst = False
        If Transfer.CAtoFC Then
            DetermineMissingCcnTons()

            If MbSfShift.CcnTonCorr <> 0 Then
                gCcnTonsEst = True
            End If

            'Adjust coarse concentrate product tons
            MbSfShift.PrdCcnTons = MbSfShift.PrdCcnTons + MbSfShift.CcnTonCorr

            'Adjust fine concentrate product tons
            MbSfShift.PrdFcnTons = MbSfShift.PrdFcnTons + MbSfShift.FcnTonCorr
        End If

        'Total operating hours for the shift
        'Rougher circuits only at this time
        MbSfTotal.N1frHrs = MbSfTotal.N1frHrs + MbSfShift.N1frHrs
        MbSfTotal.N2frHrs = MbSfTotal.N2frHrs + MbSfShift.N2frHrs
        MbSfTotal.N3frHrs = MbSfTotal.N3frHrs + MbSfShift.N3frHrs
        MbSfTotal.N4frHrs = MbSfTotal.N4frHrs + MbSfShift.N4frHrs
        MbSfTotal.N5frHrs = MbSfTotal.N5frHrs + MbSfShift.N5frHrs
        MbSfTotal.CrHrs = MbSfTotal.CrHrs + MbSfShift.CrHrs
        MbSfTotal.SrHrs = MbSfTotal.SrHrs + MbSfShift.SrHrs
        MbSfTotal.UcrHrs = MbSfTotal.UcrHrs + MbSfShift.UcrHrs

        'Product tons & BPL
        'Fine concentrate tons & BPL
        MbSfTotal.PrdFcnTons = MbSfTotal.PrdFcnTons + MbSfShift.PrdFcnTons
        If MbSfShift.PrdFcnBpl <> 0 Then
            MbSfTotal.PrdFcnTonsW = MbSfTotal.PrdFcnTonsW + MbSfShift.PrdFcnTons
        End If
        MbSfTotal.PrdFcnBt = MbSfTotal.PrdFcnBt + MbSfShift.PrdFcnTons * _
                                  MbSfShift.PrdFcnBpl

        'Coarse concentrate tons & BPL
        MbSfTotal.PrdCcnTons = MbSfTotal.PrdCcnTons + MbSfShift.PrdCcnTons
        If MbSfShift.PrdCcnBpl <> 0 Then
            MbSfTotal.PrdCcnTonsW = MbSfTotal.PrdCcnTonsW + MbSfShift.PrdCcnTons
        End If
        MbSfTotal.PrdCcnBt = MbSfTotal.PrdCcnBt + MbSfShift.PrdCcnTons * _
                                  MbSfShift.PrdCcnBpl

        'Ultra-coarse concentrate tons & BPL
        MbSfTotal.PrdUccnTons = MbSfTotal.PrdUccnTons + MbSfShift.PrdUccnTons
        If MbSfShift.PrdUccnBpl <> 0 Then
            MbSfTotal.PrdUccnTonsW = MbSfTotal.PrdUccnTonsW + MbSfShift.PrdUccnTons
        End If
        MbSfTotal.PrdUccnBt = MbSfTotal.PrdUccnBt + MbSfShift.PrdUccnTons * _
                                  MbSfShift.PrdUccnBpl

        'Total concentrate tons for the shift
        MbSfShift.PrdTotCnTons = MbSfShift.PrdFcnTons + MbSfShift.PrdCcnTons + _
                               MbSfShift.PrdUccnTons

        'Calculate ratio of concentrations for rougher circuits for shift
        '(Concentrate BPL - Tail BPL) / (Feed BPL - Tail BPL)
        'Ratio of concentrations are rounded to two decimal places
        With MbSfShift
            'N1FR ratio of concentration
            If (.N1frFdBpl - .N1frTlBpl) <> 0 Then
                .N1frRc = Round((.N1frCnBpl - .N1frTlBpl) / (.N1frFdBpl - .N1frTlBpl), mRoundVal)
            Else
                .N1frRc = 0
            End If

            'N2FR ratio of concentration
            If (.N2frFdBpl - .N2frTlBpl) <> 0 Then
                .N2frRc = Round((.N2frCnBpl - .N2frTlBpl) / (.N2frFdBpl - .N2frTlBpl), mRoundVal)
            Else
                .N2frRc = 0
            End If

            'N3FR ratio of concentration
            If (.N3frFdBpl - .N3frTlBpl) <> 0 Then
                .N3frRc = Round((.N3frCnBpl - .N3frTlBpl) / (.N3frFdBpl - .N3frTlBpl), mRoundVal)
            Else
                .N3frRc = 0
            End If

            'N4FR ratio of concentration
            If (.N4frFdBpl - .N4frTlBpl) <> 0 Then
                .N4frRc = Round((.N4frCnBpl - .N4frTlBpl) / (.N4frFdBpl - .N4frTlBpl), mRoundVal)
            Else
                .N4frRc = 0
            End If

            'N5FR ratio of concentration
            If (.N5frFdBpl - .N5frTlBpl) <> 0 Then
                .N5frRc = Round((.N5frCnBpl - .N5frTlBpl) / (.N5frFdBpl - .N5frTlBpl), mRoundVal)
            Else
                .N5frRc = 0
            End If

            'CR ratio of concentration
            If (.CrFdBpl - .CrTlBpl) <> 0 Then
                .CrRc = Round((.CrCnBpl - .CrTlBpl) / (.CrFdBpl - .CrTlBpl), mRoundVal)
            Else
                .CrRc = 0
            End If

            'SR ratio of concentration
            If (.SrFdBpl - .SrTlBpl) <> 0 Then
                .SrRc = Round((.SrCnBpl - .SrTlBpl) / (.SrFdBpl - .SrTlBpl), mRoundVal)
            Else
                .SrRc = 0
            End If

            'UCR ratio of concentration
            .UcrCnBpl = .PrdUccnBpl
            If (.UcrFdBpl - .UcrTlBpl) <> 0 Then
                .UcrRc = Round((.UcrCnBpl - .UcrTlBpl) / (.UcrFdBpl - .UcrTlBpl), mRoundVal)
            Else
                .UcrRc = 0
            End If
        End With

        'Sum total reported feed tons, reported feed tons with BPL,
        'and reported feed BPL tons

        With MbSfTotal
            'Sum reported feed tons for rougher circuits based
            'on totalizer reads
            .N1frFdTonsRpt = .N1frFdTonsRpt + MbSfShift.N1frFdTonsRpt     'N1FR
            .N2frFdTonsRpt = .N2frFdTonsRpt + MbSfShift.N2frFdTonsRpt     'N2FR
            .N3frFdTonsRpt = .N3frFdTonsRpt + MbSfShift.N3frFdTonsRpt     'N3FR
            .N4frFdTonsRpt = .N4frFdTonsRpt + MbSfShift.N4frFdTonsRpt     'N4FR
            .N5frFdTonsRpt = .N5frFdTonsRpt + MbSfShift.N5frFdTonsRpt     'N5FR
            .SrFdTonsRpt = .SrFdTonsRpt + MbSfShift.SrFdTonsRpt           'SR
            .CrFdTonsRpt = .CrFdTonsRpt + MbSfShift.CrFdTonsRpt           'CR
            .UcrFdTonsRpt = .UcrFdTonsRpt + MbSfShift.UcrFdTonsRpt        'UCR

            'Reported feed tons with BPL
            'N1FR reported feed tons with BPL
            If MbSfShift.N1frFdBpl <> 0 Then
                .N1frFdTonsRptW = .N1frFdTonsRptW + MbSfShift.N1frFdTonsRpt
            Else
                .N1frFdTonsRptW = .N1frFdTonsRptW
            End If

            'N2FR reported feed tons with BPL
            If MbSfShift.N2frFdBpl <> 0 Then
                .N2frFdTonsRptW = .N2frFdTonsRptW + MbSfShift.N2frFdTonsRpt
            Else
                .N2frFdTonsRptW = .N2frFdTonsRptW
            End If

            'N3FR reported feed tons with BPL
            If MbSfShift.N3frFdBpl <> 0 Then
                .N3frFdTonsRptW = .N3frFdTonsRptW + MbSfShift.N3frFdTonsRpt
            Else
                .N3frFdTonsRptW = .N3frFdTonsRptW
            End If

            'N4FR reported feed tons with BPL
            If MbSfShift.N4frFdBpl <> 0 Then
                .N4frFdTonsRptW = .N4frFdTonsRptW + MbSfShift.N4frFdTonsRpt
            Else
                .N4frFdTonsRptW = .N4frFdTonsRptW
            End If

            'N5FR reported feed tons with BPL
            If MbSfShift.N5frFdBpl <> 0 Then
                .N5frFdTonsRptW = .N5frFdTonsRptW + MbSfShift.N5frFdTonsRpt
            Else
                .N5frFdTonsRptW = .N5frFdTonsRptW
            End If

            'SR reported feed tons with BPL
            If MbSfShift.SrFdBpl <> 0 Then
                .SrFdTonsRptW = .SrFdTonsRptW + MbSfShift.SrFdTonsRpt
            Else
                .SrFdTonsRptW = .SrFdTonsRptW
            End If

            'CR reported feed tons with BPL
            If MbSfShift.CrFdBpl <> 0 Then
                .CrFdTonsRptW = .CrFdTonsRptW + MbSfShift.CrFdTonsRpt
            Else
                .CrFdTonsRptW = .CrFdTonsRptW
            End If

            'UCR reported feed tons with BPL
            If MbSfShift.UcrFdBpl <> 0 Then
                .UcrFdTonsRptW = .UcrFdTonsRptW + MbSfShift.UcrFdTonsRpt
            Else
                .UcrFdTonsRptW = .UcrFdTonsRptW
            End If

            'Sum reported feed BPL * reported feed tons  (BPL tons)
            'N1FR reported feed BPL tons
            'Feed BPL-tons are rounded to 1 decimal place
            .N1frFdBtRpt = .N1frFdBtRpt + Round(MbSfShift.N1frFdTonsRpt * _
                               MbSfShift.N1frFdBpl, 1)

            'N2FR reported feed BPL tons
            .N2frFdBtRpt = .N2frFdBtRpt + Round(MbSfShift.N2frFdTonsRpt * _
                               MbSfShift.N2frFdBpl, 1)

            'N3FR reported feed BPL tons
            .N3frFdBtRpt = .N3frFdBtRpt + Round(MbSfShift.N3frFdTonsRpt * _
                               MbSfShift.N3frFdBpl, 1)

            'N4FR reported feed BPL tons
            .N4frFdBtRpt = .N4frFdBtRpt + Round(MbSfShift.N4frFdTonsRpt * _
                               MbSfShift.N4frFdBpl, 1)

            'N5FR reported feed BPL tons
            .N5frFdBtRpt = .N5frFdBtRpt + Round(MbSfShift.N5frFdTonsRpt * _
                               MbSfShift.N5frFdBpl, 1)

            'SR reported feed BPL tons
            .SrFdBtRpt = .SrFdBtRpt + Round(MbSfShift.SrFdTonsRpt * _
                             MbSfShift.SrFdBpl, 1)

            'CR reported feed BPL tons
            .CrFdBtRpt = .CrFdBtRpt + Round(MbSfShift.CrFdTonsRpt * _
                             MbSfShift.CrFdBpl, 1)

            'UCR reported feed BPL tons
            .UcrFdBtRpt = .UcrFdBtRpt + Round(MbSfShift.UcrFdTonsRpt * _
                             MbSfShift.UcrFdBpl, 1)

        End With

        'Determine concentrate tons expected from totalizer feed tons
        'for the rougher circuits.
        'Use totalizer feed tons and previously calculated ratio of
        'concentrations to determine concentrate tons expected from
        'each rougher circuit.
        'Concentrate tons expected are rounded to 0 decimal places

        With MbSfShift
            'N1FR concentrate tons expected
            If .N1frRc <> 0 Then
                .N1frCnTonsExp = Round(.N1frFdTonsRpt / .N1frRc, 0)
            Else
                .N1frCnTonsExp = 0
            End If

            'N2FR concentrate tons expected
            If .N2frRc <> 0 Then
                .N2frCnTonsExp = Round(.N2frFdTonsRpt / .N2frRc, 0)
            Else
                .N2frCnTonsExp = 0
            End If

            'N3FR concentrate tons expected
            If .N3frRc <> 0 Then
                .N3frCnTonsExp = Round(.N3frFdTonsRpt / .N3frRc, 0)
            Else
                .N3frCnTonsExp = 0
            End If

            'N4FR concentrate tons expected
            If .N4frRc <> 0 Then
                .N4frCnTonsExp = Round(.N4frFdTonsRpt / .N4frRc, 0)
            Else
                .N4frCnTonsExp = 0
            End If

            'N5FR concentrate tons expected
            If .N5frRc <> 0 Then
                .N5frCnTonsExp = Round(.N5frFdTonsRpt / .N5frRc, 0)
            Else
                .N5frCnTonsExp = 0
            End If

            'CR concentrate tons expected
            If .CrRc <> 0 Then
                .CrCnTonsExp = Round(.CrFdTonsRpt / .CrRc, 0)
            Else
                .CrCnTonsExp = 0
            End If

            'SR concentrate tons expected
            If .SrRc <> 0 Then
                .SrCnTonsExp = Round(.SrFdTonsRpt / .SrRc, 0)
            Else
                .SrCnTonsExp = 0
            End If

        End With

        'Shift fine concentrate BPL
        'Equals #1 Fine amine concentrate BPL, #2 Fine amine concentrate BPL,
        'and #3 Fine amine concentrate BPL
        'The N1FA, N2FA, & N3FA concentrate BPL's are equal to the shift production
        'fine concentrate BPL

        '10/31/2011, lss
        'The #1FA, #2FA, #3FA may not all have run!
        'Will check .N1faFdBpl, .N2faFdBpl, .N3faFdBpl to see if the circuits ran!

        With MbSfShift
            If .N1faFdBpl > 0 Then
                .N1faCnBpl = .PrdFcnBpl
            Else
                .N1faCnBpl = 0
            End If

            If .N2faFdBpl > 0 Then
                .N2faCnBpl = .PrdFcnBpl
            Else
                .N2faCnBpl = 0
            End If

            If .N3faFdBpl > 0 Then
                .N3faCnBpl = .PrdFcnBpl
            Else
                .N3faCnBpl = 0
            End If
        End With

        'Transfer factors, 1997 vintage are:
        ' 1)  1FR to 1FA  100       Sources of feed for 1FA
        ' 2)  1FR to 2FA            1)  1FR         2)  4FR
        ' 3)  2FR to 2FA            3)  5FR         4)  SR
        ' 4)  2FR to 3FA  100       5)  CR
        ' 5)  3FR to 2FA
        ' 6)  3FR to 3FA  100       Sources of feed for 2FA
        ' 7)  4FR to 1FA  100       1)  1FR         2)  2FR
        ' 8)  4FR to 2FA            3)  3FR         4)  4FR
        ' 9)  5FR to 1FA            5)  5FR         6)  SR
        '10)  5FR to 2FA  100
        '11)  SR to 1FA             Sources of feed for 3FA
        '12)  SR to 2FA   100       1)  2FR         2)  3FR
        '13)  SR to 3FA             3)  SR
        '14)  CR to CA    100
        '15)  CR to 1FA             Sources of feed for CA
        '16)  SR to CA              1)  CR
        '                           2)  SR

        '#1 Fine amine circuit  #1 Fine amine circuit
        '#1 Fine amine circuit  #1 Fine amine circuit
        '#1 Fine amine circuit  #1 Fine amine circuit

        'N1FA feed may come from these rougher circuits:
        'N1FR, N4FR, N5FR, SR, CR
        'Usually it comes from N1FR & N4FR
        N1faFdTons.N1fr = 0
        N1faFdTons.N4fr = 0
        N1faFdTons.N5fr = 0
        N1faFdTons.Sr = 0
        N1faFdTons.Cr = 0

        '#1FA -- rougher conc tons from #1FR
        N1faFdTons.N1fr = Round(Transfer.N1FRtoN1FA / 100 * MbSfShift.N1frCnTonsExp, 0)

        '#1FA -- rougher conc tons from #4FR
        N1faFdTons.N4fr = Round(Transfer.N4FRtoN1FA / 100 * MbSfShift.N4frCnTonsExp, 0)

        '#1FA -- rougher conc tons from #5FR
        N1faFdTons.N5fr = Round(Transfer.N5FRtoN1FA / 100 * MbSfShift.N5frCnTonsExp, 0)

        '#1FA -- rougher conc tons from SR
        N1faFdTons.Sr = Round(Transfer.SRtoN1FA / 100 * MbSfShift.SrCnTonsExp, 0)

        '#1FA -- rougher conc tons from CR
        N1faFdTons.Cr = Round(Transfer.CRtoN1FA / 100 * MbSfShift.CrCnTonsExp, 0)

        MbSfShift.N1faFdTonsRpt = N1faFdTons.N1fr + N1faFdTons.N4fr + _
                                N1faFdTons.N5fr + N1faFdTons.Sr + N1faFdTons.Cr

        MbSfTotal.N1faFdTonsRpt = MbSfTotal.N1faFdTonsRpt + MbSfShift.N1faFdTonsRpt

        MbSfShift.N1faFdBt = Round((N1faFdTons.N1fr * MbSfShift.N1frCnBpl) + _
                           (N1faFdTons.N4fr * MbSfShift.N4frCnBpl) + _
                           (N1faFdTons.N5fr * MbSfShift.N5frCnBpl) + _
                           (N1faFdTons.Sr * MbSfShift.SrCnBpl) + _
                           (N1faFdTons.Cr * MbSfShift.CrCnBpl), 1)

        If MbSfShift.N1faFdTonsRpt <> 0 Then
            MbSfShift.N1faFdBpl = Round(MbSfShift.N1faFdBt / MbSfShift.N1faFdTonsRpt, 1)
        Else
            MbSfShift.N1faFdBpl = 0
        End If

        '#2 Fine amine circuit  #2 Fine amine circuit
        '#2 Fine amine circuit  #2 Fine amine circuit
        '#2 Fine amine circuit  #2 Fine amine circuit

        'N2FA feed may come from these rougher circuits:
        'N1FR, N2FR, N3FR, N4FR, N5FR, SR
        'Usually it comes from N5FR & SR
        N2faFdTons.N1fr = 0
        N2faFdTons.N2fr = 0
        N2faFdTons.N3fr = 0
        N2faFdTons.N4fr = 0
        N2faFdTons.N5fr = 0
        N2faFdTons.Sr = 0

        '#2FA -- rougher conc tons from #1FR
        N2faFdTons.N1fr = Round(Transfer.N1FRtoN2FA / 100 * MbSfShift.N1frCnTonsExp, 0)

        '#2FA -- rougher conc tons from #2FR
        N2faFdTons.N2fr = Round(Transfer.N2FRtoN2FA / 100 * MbSfShift.N2frCnTonsExp, 0)

        '#2FA -- rougher conc tons from #3FR
        N2faFdTons.N3fr = Round(Transfer.N3FRtoN2FA / 100 * MbSfShift.N3frCnTonsExp, 0)

        '#2FA -- rougher conc tons from #4FR
        N2faFdTons.N4fr = Round(Transfer.N4FRtoN2FA / 100 * MbSfShift.N4frCnTonsExp, 0)

        '#2FA -- rougher conc tons from #5FR
        N2faFdTons.N5fr = Round(Transfer.N5FRtoN2FA / 100 * MbSfShift.N5frCnTonsExp, 0)

        '#2FA -- rougher conc tons from SR
        N2faFdTons.Sr = Round(Transfer.SRtoN2FA / 100 * MbSfShift.SrCnTonsExp, 0)

        MbSfShift.N2faFdTonsRpt = N2faFdTons.N1fr + N2faFdTons.N2fr + _
                                N2faFdTons.N3fr + N2faFdTons.N4fr + N2faFdTons.N5fr + _
                                N2faFdTons.Sr

        MbSfTotal.N2faFdTonsRpt = MbSfTotal.N2faFdTonsRpt + MbSfShift.N2faFdTonsRpt

        MbSfShift.N2faFdBt = Round((N2faFdTons.N1fr * MbSfShift.N1frCnBpl) + _
                           (N2faFdTons.N2fr * MbSfShift.N2frCnBpl) + _
                           (N2faFdTons.N3fr * MbSfShift.N3frCnBpl) + _
                           (N2faFdTons.N4fr * MbSfShift.N4frCnBpl) + _
                           (N2faFdTons.N5fr * MbSfShift.N5frCnBpl) + _
                           (N2faFdTons.Sr * MbSfShift.SrCnBpl), 1)

        If MbSfShift.N2faFdTonsRpt <> 0 Then
            MbSfShift.N2faFdBpl = Round(MbSfShift.N2faFdBt / MbSfShift.N2faFdTonsRpt, 1)
        Else
            MbSfShift.N2faFdBpl = 0
        End If

        '#3 Fine amine circuit  #3 Fine amine circuit
        '#3 Fine amine circuit  #3 Fine amine circuit
        '#3 Fine amine circuit  #3 Fine amine circuit

        'N3FA feed may come from these rougher circuits:
        'N2FR, N3FR, SR
        'Usually it comes from the N2FR & N3FR
        N3faFdTons.N2fr = 0
        N3faFdTons.N3fr = 0
        N3faFdTons.Sr = 0

        '#2FA -- rougher conc tons from #2FR
        N3faFdTons.N2fr = Round(Transfer.N2FRtoN3FA / 100 * MbSfShift.N2frCnTonsExp, 0)

        '#2FA -- rougher conc tons from #3FR
        N3faFdTons.N3fr = Round(Transfer.N3FRtoN3FA / 100 * MbSfShift.N3frCnTonsExp, 0)

        '#2FA -- rougher conc tons from SR
        N3faFdTons.Sr = Round(Transfer.SRtoN3FA / 100 * MbSfShift.SrCnTonsExp, 0)

        MbSfShift.N3faFdTonsRpt = N3faFdTons.N2fr + N3faFdTons.N3fr + _
                                N3faFdTons.Sr

        MbSfTotal.N3faFdTonsRpt = MbSfTotal.N3faFdTonsRpt + MbSfShift.N3faFdTonsRpt

        MbSfShift.N3faFdBt = Round((N3faFdTons.N2fr * MbSfShift.N2frCnBpl) + _
                           (N3faFdTons.N3fr * MbSfShift.N3frCnBpl) + _
                           (N3faFdTons.Sr * MbSfShift.SrCnBpl), 1)

        If MbSfShift.N3faFdTonsRpt <> 0 Then
            MbSfShift.N3faFdBpl = Round(MbSfShift.N3faFdBt / MbSfShift.N3faFdTonsRpt, 1)
        Else
            MbSfShift.N3faFdBpl = 0
        End If

        'Coarse amine circuit  Coarse amine circuit
        'Coarse amine circuit  Coarse amine circuit
        'Coarse amine circuit  Coarse amine circuit

        'CA feed may come from these rougher circuits:
        'CR
        'SR -- added 05/25/2005, lss

        CaFdTons.Cr = 0
        CaFdTons.Sr = 0

        'CA -- rougher conc tons from CR
        CaFdTons.Cr = Round(Transfer.CRtoCA / 100 * MbSfShift.CrCnTonsExp, 0)

        'CA -- rougher conc tons from SR (added 05/23/2005, lss)
        CaFdTons.Sr = Round(Transfer.SRtoCA / 100 * MbSfShift.SrCnTonsExp, 0)

        MbSfShift.CaFdTonsRpt = CaFdTons.Cr + CaFdTons.Sr
        MbSfTotal.CaFdTonsRpt = MbSfTotal.CaFdTonsRpt + MbSfShift.CaFdTonsRpt

        MbSfShift.CaFdBt = Round(CaFdTons.Cr * MbSfShift.CrCnBpl, 1) + _
                           Round(CaFdTons.Sr * MbSfShift.SrCnBpl, 1)

        If MbSfShift.CaFdTonsRpt <> 0 Then
            MbSfShift.CaFdBpl = Round(MbSfShift.CaFdBt / MbSfShift.CaFdTonsRpt, 1)
        Else
            MbSfShift.CaFdBpl = 0
        End If

        'Divide fine concentrate tons  Divide fine concentrate tons
        'Divide fine concentrate tons  Divide fine concentrate tons
        'Divide fine concentrate tons  Divide fine concentrate tons

        'Need to divide fine concentrate tons into #1FA amine concentrate
        'tons, #2FA concentrate tons, & #3Fa concentrate tons.
        'Will ratio by using feed BPL-tons.

        '#1FA, #2FA, #3FA feed BPL-tons -- have already calculated above.
        '#1FA   MbSfShift.N1faFdBt
        '#2FA   MbSfShift.N2faFdBt
        '#3FA   MbSfShift.N3faFdBt

        '#1FA ratio of concentration
        'MbSfShift.N1faFdBpl      Calculated above
        'MbSfShift.N1faCnBpl      Total fine concentrate product BPL  (measured)
        'MbSfShift.N1faTlBpl      Measured
        'Ratio of concentrations are rounded to two decimal places
        If MbSfShift.N1faFdBpl - MbSfShift.N1faTlBpl <> 0 Then
            MbSfShift.N1faRc = Round((MbSfShift.N1faCnBpl - MbSfShift.N1faTlBpl) / _
                            (MbSfShift.N1faFdBpl - MbSfShift.N1faTlBpl), 2)
        Else
            MbSfShift.N1faRc = 0
        End If

        '#2FA ratio of concentration
        'MbSfShift.N2faFdBpl      Calculated above
        'MbSfShift.N2faCnBpl      Total fine concentrate product BPL  (measured)
        'MbSfShift.N2faTlBpl      Measured
        'Ratio of concentrations are rounded to two decimal places
        If MbSfShift.N2faFdBpl - MbSfShift.N2faTlBpl <> 0 Then
            MbSfShift.N2faRc = Round((MbSfShift.N2faCnBpl - MbSfShift.N2faTlBpl) / _
                            (MbSfShift.N2faFdBpl - MbSfShift.N2faTlBpl), 2)
        Else
            MbSfShift.N2faRc = 0
        End If

        '#3FA ratio of concentration
        'MbSfShift.N3faFdBpl      Calculated above
        'MbSfShift.N3faCnBpl      Total fine concentrate product BPL  (measured)
        'MbSfShift.N3faTlBpl      Measured
        'Ratio of concentrations are rounded to two decimal places
        If MbSfShift.N3faFdBpl - MbSfShift.N3faTlBpl <> 0 Then
            MbSfShift.N3faRc = Round((MbSfShift.N3faCnBpl - MbSfShift.N3faTlBpl) / _
                            (MbSfShift.N3faFdBpl - MbSfShift.N3faTlBpl), 2)
        Else
            MbSfShift.N3faRc = 0
        End If

        '#1FA fine concentrate tons expected
        If MbSfShift.N1faRc <> 0 Then
            MbSfShift.N1faCnTonsExp = Round(MbSfShift.N1faFdTonsRpt / MbSfShift.N1faRc, 0)
        Else
            MbSfShift.N1faCnTonsExp = 0
        End If

        '#2FA fine concentrate tons expected
        If MbSfShift.N2faRc <> 0 Then
            MbSfShift.N2faCnTonsExp = Round(MbSfShift.N2faFdTonsRpt / MbSfShift.N2faRc, 0)
        Else
            MbSfShift.N2faCnTonsExp = 0
        End If

        '#3FA fine concentrate tons expected
        If MbSfShift.N3faRc <> 0 Then
            MbSfShift.N3faCnTonsExp = Round(MbSfShift.N3faFdTonsRpt / MbSfShift.N3faRc, 0)
        Else
            MbSfShift.N3faCnTonsExp = 0
        End If

        'Total concentrate tons expected -- N1FA, N2FA, & N3FA
        MbSfShift.N123faCnTonsExp = MbSfShift.N1faCnTonsExp + MbSfShift.N2faCnTonsExp + _
                                MbSfShift.N3faCnTonsExp

        'Method 3 from the old Ops Module
        'Determine percent of total concentrate tons to fine amine circuit that is
        'assigned to the #1FA circuit
        'Determine percent of total concentrate tons to fine amine circuit that is
        'assigned to the #2FA circuit
        'Determine percent of total concentrate tons to fine amine circuit that is
        'assigned to the #3FA circuit
        If MbSfShift.N123faCnTonsExp <> 0 Then
            MbSfShift.N1faPct3 = Round(MbSfShift.N1faCnTonsExp / MbSfShift.N123faCnTonsExp, 4)
        Else
            MbSfShift.N1faPct3 = 0
        End If

        If MbSfShift.N123faCnTonsExp <> 0 Then
            MbSfShift.N2faPct3 = Round(MbSfShift.N2faCnTonsExp / MbSfShift.N123faCnTonsExp, 4)
        Else
            MbSfShift.N2faPct3 = 0
        End If

        MbSfShift.N3faPct3 = 1 - MbSfShift.N1faPct3 - MbSfShift.N2faPct3

        'Divide actual fine concentrate tons (Product) between the N1FA, N2FA, &
        'N3FA circuits
        MbSfShift.N1faCnTons = Round(MbSfShift.N1faPct3 * MbSfShift.PrdFcnTons, 0)
        MbSfTotal.N1faCnTons = MbSfTotal.N1faCnTons + MbSfShift.N1faCnTons

        MbSfShift.N2faCnTons = Round(MbSfShift.N2faPct3 * MbSfShift.PrdFcnTons, 0)
        MbSfTotal.N2faCnTons = MbSfTotal.N2faCnTons + MbSfShift.N2faCnTons

        MbSfShift.N3faCnTons = MbSfShift.PrdFcnTons - MbSfShift.N1faCnTons - _
                             MbSfShift.N2faCnTons
        MbSfTotal.N3faCnTons = MbSfTotal.N3faCnTons + MbSfShift.N3faCnTons

        '10/31/2011, lss
        'Need Cn Bpl-Tons
        MbSfShift.N1faCnBt = MbSfShift.N1faCnBpl * MbSfShift.N1faCnTons
        MbSfTotal.N1faCnBtAdj = MbSfTotal.N1faCnBtAdj + MbSfShift.N1faCnBt
        '--
        MbSfShift.N2faCnBt = MbSfShift.N2faCnBpl * MbSfShift.N2faCnTons
        MbSfTotal.N2faCnBtAdj = MbSfTotal.N2faCnBtAdj + MbSfShift.N2faCnBt
        '--
        MbSfShift.N3faCnBt = MbSfShift.N3faCnBpl * MbSfShift.N3faCnTons
        MbSfTotal.N3faCnBtAdj = MbSfTotal.N3faCnBtAdj + MbSfShift.N3faCnBt

        'Coarse amine circuit miscellanous
        'Coarse amine circuit miscellanous
        'Coarse amine circuit miscellanous

        'Only one coarse amine circuit -- do not have to divvy up coarse
        'concentrate tons the way the fine concentrate tons were divided.

        MbSfShift.CaCnTons = MbSfShift.PrdCcnTons
        MbSfTotal.CaCnTons = MbSfTotal.CaCnTons + MbSfShift.CaCnTons

        'Determine ratio of concentration for four amine circuits
        '#1FA ratio of concentration
        'MbSfShift.N1faRc -- already calculated
        'MbSfShift.N2faRc -- already calculated
        'MbSfShift.N3frRc -- already calculated
        'MbSfShift.CaRc
        'Ratio of concentrations are rounded to two decimal places
        If MbSfShift.CaFdBpl - MbSfShift.CaTlBpl <> 0 Then
            MbSfShift.CaRc = Round((MbSfShift.PrdCcnBpl - MbSfShift.CaTlBpl) / _
                            (MbSfShift.CaFdBpl - MbSfShift.CaTlBpl), 2)
        Else
            MbSfShift.CaRc = 0
        End If

        'Adjust the amine circuit feed tons
        'Adjusted #1FA feed tons (Rougher concentrate tons to #1FA)
        MbSfShift.N1faFdTonsAdj = Round(MbSfShift.N1faRc * MbSfShift.N1faCnTons, 0)
        MbSfTotal.N1faFdTonsAdj = MbSfTotal.N1faFdTonsAdj + MbSfShift.N1faFdTonsAdj
        'Total #1FA feed BPL-tons (Rougher concentrate BPL-tons to #1FA)
        MbSfTotal.N1faFdBtAdj = MbSfTotal.N1faFdBtAdj + Round(MbSfShift.N1faFdTonsAdj * _
                              MbSfShift.N1faFdBpl, 1)

        'Adjusted #2FA feed tons (Rougher concentrate tons to #2FA)
        MbSfShift.N2faFdTonsAdj = Round(MbSfShift.N2faRc * MbSfShift.N2faCnTons, 0)
        MbSfTotal.N2faFdTonsAdj = MbSfTotal.N2faFdTonsAdj + MbSfShift.N2faFdTonsAdj
        'Total #2FA feed BPL-tons (Rougher concentrate BPL-tons to #2FA)
        MbSfTotal.N2faFdBtAdj = MbSfTotal.N2faFdBtAdj + Round(MbSfShift.N2faFdTonsAdj * _
                              MbSfShift.N2faFdBpl, 1)

        'Adjusted #3FA feed tons (Rougher concentrate tons to #3FA)
        MbSfShift.N3faFdTonsAdj = Round(MbSfShift.N3faRc * MbSfShift.N3faCnTons, 0)
        MbSfTotal.N3faFdTonsAdj = MbSfTotal.N3faFdTonsAdj + MbSfShift.N3faFdTonsAdj
        'Total #3FA feed BPL-tons (Rougher concntrate BPL-tons to #3FA)
        MbSfTotal.N3faFdBtAdj = MbSfTotal.N3faFdBtAdj + Round(MbSfShift.N3faFdTonsAdj * _
                              MbSfShift.N3faFdBpl, 1)

        'Adjusted #CA feed tons (Rougher concentrate tons to #CA)
        MbSfShift.CaFdTonsAdj = MbSfShift.CaRc * MbSfShift.PrdCcnTons
        MbSfTotal.CaFdTonsAdj = MbSfTotal.CaFdTonsAdj + MbSfShift.CaFdTonsAdj
        'Total #CA feed BPL-tons (Rougher concntrate BPL-tons to #CA)
        MbSfTotal.CaFdBtAdj = MbSfTotal.CaFdBtAdj + Round(MbSfShift.CaFdTonsAdj * _
                              MbSfShift.CaFdBpl, 1)


        'Rougher circuit rougher concentrate BPL-tons
        MbSfShift.N1frCnBt = Round(MbSfShift.N1frCnTonsExp * MbSfShift.N1frCnBpl, 1)
        MbSfShift.N2frCnBt = Round(MbSfShift.N2frCnTonsExp * MbSfShift.N2frCnBpl, 1)
        MbSfShift.N3frCnBt = Round(MbSfShift.N3frCnTonsExp * MbSfShift.N3frCnBpl, 1)
        MbSfShift.N4frCnBt = Round(MbSfShift.N4frCnTonsExp * MbSfShift.N4frCnBpl, 1)
        MbSfShift.N5frCnBt = Round(MbSfShift.N5frCnTonsExp * MbSfShift.N5frCnBpl, 1)
        MbSfShift.CrCnBt = Round(MbSfShift.CrCnTonsExp * MbSfShift.CrCnBpl, 1)
        MbSfShift.SrCnBt = Round(MbSfShift.SrCnTonsExp * MbSfShift.SrCnBpl, 1)

        'Divide #1FA circuit  Divide #1FA circuit
        'Divide #1FA circuit  Divide #1FA circuit
        'Divide #1FA circuit  Divide #1FA circuit

        'Sources of feed for #1FA
        'N1FR       Transfer.N1FRtoN1FA     MbSfShift.N1frCnTonsExp
        'N4FR       Transfer.N4FRtoN1FA     MbSfShift.N4frCnTonsExp
        'N5FR       Transfer.N5FRtoN1FA     MbSfShift.N5frCnTonsExp
        'SR         Transfer.SRtoN1FA       MbSfShift.SrCnTonsExp
        'CR         Transfer.CRtoN1FA       MbSfShift.CrCnTonsExp
        'Usually from N1FR & N4FR
        MbSfShift.N1FaN1frBt = Round((Transfer.N1FRtoN1FA / 100) * _
                             MbSfShift.N1frCnTonsExp * MbSfShift.N1frCnBpl, 1)
        MbSfShift.N1FaN4frBt = Round((Transfer.N4FRtoN1FA / 100) * _
                             MbSfShift.N4frCnTonsExp * MbSfShift.N4frCnBpl, 1)
        MbSfShift.N1FaN5frBt = Round((Transfer.N5FRtoN1FA / 100) * _
                             MbSfShift.N5frCnTonsExp * MbSfShift.N5frCnBpl, 1)
        MbSfShift.N1FaSrBt = Round((Transfer.SRtoN1FA / 100) * _
                             MbSfShift.SrCnTonsExp * MbSfShift.SrCnBpl, 1)
        MbSfShift.N1FaCrBt = Round((Transfer.CRtoN1FA / 100) * _
                             MbSfShift.CrCnTonsExp * MbSfShift.CrCnBpl, 1)

        MbSfShift.N1FaTotBt = MbSfShift.N1FaN1frBt + MbSfShift.N1FaN4frBt + _
                            MbSfShift.N1FaN5frBt + MbSfShift.N1FaSrBt + _
                            MbSfShift.N1FaCrBt

        If MbSfShift.N1FaTotBt <> 0 Then
            MbSfShift.N1FaN1frBtPct = Round(MbSfShift.N1FaN1frBt / _
                                    MbSfShift.N1FaTotBt, 4)
        Else
            MbSfShift.N1FaN1frBtPct = 0
        End If

        If MbSfShift.N1FaTotBt <> 0 Then
            MbSfShift.N1FaN4frBtPct = Round(MbSfShift.N1FaN4frBt / _
                                    MbSfShift.N1FaTotBt, 4)
        Else
            MbSfShift.N1FaN4frBtPct = 0
        End If

        If MbSfShift.N1FaTotBt <> 0 Then
            MbSfShift.N1FaN5frBtPct = Round(MbSfShift.N1FaN5frBt / _
                                    MbSfShift.N1FaTotBt, 4)
        Else
            MbSfShift.N1FaN5frBtPct = 0
        End If

        If MbSfShift.N1FaTotBt <> 0 Then
            MbSfShift.N1FaSrBtPct = Round(MbSfShift.N1FaSrBt / _
                                    MbSfShift.N1FaTotBt, 4)
        Else
            MbSfShift.N1FaSrBtPct = 0
        End If

        If MbSfShift.N1FaTotBt <> 0 Then
            MbSfShift.N1FaCrBtPct = Round(MbSfShift.N1FaCrBt / _
                                    MbSfShift.N1FaTotBt, 4)
        Else
            MbSfShift.N1FaCrBtPct = 0
        End If

        'Divide #2FA circuit  Divide #2FA circuit
        'Divide #2FA circuit  Divide #2FA circuit
        'Divide #2FA circuit  Divide #2FA circuit

        'Sources of feed for #2FA
        'N1FR       Transfer.N1FRtoN1FA     MbSfShift.N1frCnTonsExp
        'N2FR       Transfer.N2FRtoN1FA     MbSfShift.N2frCnTonsExp
        'N3FR       Transfer.N3FRtoN1FA     MbSfShift.N3frCnTonsExp
        'N4FR       Transfer.N4FRtoN1FA     MbSfShift.N4frCnTonsExp
        'N5FR       Transfer.N5FRtoN1FA     MbSfShift.N5frCnTonsExp
        'SR         Transfer.SRtoN1FA       MbSfShift.SrCnTonsExp
        'Usually from N5FR & SR
        MbSfShift.N2FaN1frBt = Round((Transfer.N1FRtoN2FA / 100) * _
                             MbSfShift.N1frCnTonsExp * MbSfShift.N1frCnBpl, 1)
        MbSfShift.N2FaN2frBt = Round((Transfer.N2FRtoN2FA / 100) * _
                             MbSfShift.N2frCnTonsExp * MbSfShift.N2frCnBpl, 1)
        MbSfShift.N2FaN3frBt = Round((Transfer.N3FRtoN2FA / 100) * _
                             MbSfShift.N3frCnTonsExp * MbSfShift.N3frCnBpl, 1)
        MbSfShift.N2FaN4frBt = Round((Transfer.N4FRtoN2FA / 100) * _
                             MbSfShift.N4frCnTonsExp * MbSfShift.N4frCnBpl, 1)
        MbSfShift.N2FaN5frBt = Round((Transfer.N5FRtoN2FA / 100) * _
                             MbSfShift.N5frCnTonsExp * MbSfShift.N5frCnBpl, 1)
        MbSfShift.N2FaSrBt = Round((Transfer.SRtoN2FA / 100) * _
                             MbSfShift.SrCnTonsExp * MbSfShift.SrCnBpl, 1)

        MbSfShift.N2FaTotBt = MbSfShift.N2FaN1frBt + MbSfShift.N2FaN2frBt + _
                            MbSfShift.N2FaN3frBt + MbSfShift.N2FaN4frBt + _
                            MbSfShift.N2FaN5frBt + MbSfShift.N2FaSrBt

        If MbSfShift.N2FaTotBt <> 0 Then
            MbSfShift.N2FaN1frBtPct = Round(MbSfShift.N2FaN1frBt / _
                                    MbSfShift.N2FaTotBt, 4)
        Else
            MbSfShift.N2FaN1frBtPct = 0
        End If

        If MbSfShift.N2FaTotBt <> 0 Then
            MbSfShift.N2FaN2frBtPct = Round(MbSfShift.N2FaN2frBt / _
                                    MbSfShift.N2FaTotBt, 4)
        Else
            MbSfShift.N2FaN2frBtPct = 0
        End If

        If MbSfShift.N2FaTotBt <> 0 Then
            MbSfShift.N2FaN3frBtPct = Round(MbSfShift.N2FaN3frBt / _
                                    MbSfShift.N2FaTotBt, 4)
        Else
            MbSfShift.N2FaN3frBtPct = 0
        End If

        If MbSfShift.N2FaTotBt <> 0 Then
            MbSfShift.N2FaN4frBtPct = Round(MbSfShift.N2FaN4frBt / _
                                    MbSfShift.N2FaTotBt, 4)
        Else
            MbSfShift.N2FaN4frBtPct = 0
        End If

        If MbSfShift.N2FaTotBt <> 0 Then
            MbSfShift.N2FaN5frBtPct = Round(MbSfShift.N2FaN5frBt / _
                                    MbSfShift.N2FaTotBt, 4)
        Else
            MbSfShift.N2FaN5frBtPct = 0
        End If

        If MbSfShift.N2FaTotBt <> 0 Then
            MbSfShift.N2FaSrBtPct = Round(MbSfShift.N2FaSrBt / _
                                    MbSfShift.N2FaTotBt, 4)
        Else
            MbSfShift.N2FaSrBtPct = 0
        End If

        'Divide #3FA circuit  Divide #3FA circuit
        'Divide #3FA circuit  Divide #3FA circuit
        'Divide #3FA circuit  Divide #3FA circuit

        'Sources of feed for #2FA
        'N2FR       Transfer.N2FRtoN3FA     MbSfShift.N2frCnTonsExp
        'N3FR       Transfer.N3FRtoN3FA     MbSfShift.N3frCnTonsExp
        'SR         Transfer.SRtoN3FA       MbSfShift.SrCnTonsExp
        MbSfShift.N3FaN2frBt = Round((Transfer.N2FRtoN3FA / 100) * _
                             MbSfShift.N2frCnTonsExp * MbSfShift.N2frCnBpl, 1)
        MbSfShift.N3FaN3frBt = Round((Transfer.N3FRtoN3FA / 100) * _
                             MbSfShift.N3frCnTonsExp * MbSfShift.N3frCnBpl, 1)
        MbSfShift.N3FaSrBt = Round((Transfer.SRtoN3FA / 100) * _
                             MbSfShift.SrCnTonsExp * MbSfShift.SrCnBpl, 1)

        MbSfShift.N3FaTotBt = MbSfShift.N3FaN2frBt + MbSfShift.N3FaN3frBt + _
                            MbSfShift.N3FaSrBt

        If MbSfShift.N3FaTotBt <> 0 Then
            MbSfShift.N3FaN2frBtPct = Round(MbSfShift.N3FaN2frBt / _
                                    MbSfShift.N3FaTotBt, 4)
        Else
            MbSfShift.N3FaN2frBtPct = 0
        End If

        If MbSfShift.N3FaTotBt <> 0 Then
            MbSfShift.N3FaN3frBtPct = Round(MbSfShift.N3FaN3frBt / _
                                    MbSfShift.N3FaTotBt, 4)
        Else
            MbSfShift.N3FaN3frBtPct = 0
        End If

        If MbSfShift.N3FaTotBt <> 0 Then
            MbSfShift.N3FaSrBtPct = Round(MbSfShift.N3FaSrBt / _
                                    MbSfShift.N3FaTotBt, 4)
        Else
            MbSfShift.N3FaSrBtPct = 0
        End If

        'Sources of feed for CA -- added 05/23/2005, lss
        'CR         Transfer.CRtoCA       MbSfShift.CrCnTonsExp
        'SR         Transfer.SRtoCA       MbSfShift.SrCnTonsExp
        'As of 05/2005 the coarse amine circuit can receive rougher
        'concentrate from both the SR and the CR
        MbSfShift.CaCrBt = Round((Transfer.CRtoCA / 100) * _
                           MbSfShift.CrCnTonsExp * MbSfShift.CrCnBpl, 1)
        MbSfShift.CaSrBt = Round((Transfer.SRtoCA / 100) * _
                           MbSfShift.SrCnTonsExp * MbSfShift.SrCnBpl, 1)

        MbSfShift.CaTotBt = MbSfShift.CaCrBt + MbSfShift.CaSrBt

        If MbSfShift.CaTotBt <> 0 Then
            MbSfShift.CaCrBtPct = Round(MbSfShift.CaCrBt / _
                                    MbSfShift.CaTotBt, 4)
        Else
            MbSfShift.CaCrBtPct = 0
        End If

        If MbSfShift.CaTotBt <> 0 Then
            MbSfShift.CaSrBtPct = Round(MbSfShift.CaSrBt / _
                                    MbSfShift.CaTotBt, 4)
        Else
            MbSfShift.CaSrBtPct = 0
        End If

        'Recalculations  Recalculations  Recalculations
        'Recalculations  Recalculations  Recalculations
        'Recalculations  Recalculations  Recalculations

        '#1FR rougher concentrate tons -- recalculated
        '#1FR rougher concentrate tons -- recalculated
        '#1FR rougher concentrate tons -- recalculated

        'From #1FA & #2Fa
        MbSfShift.N1frCnTonsAdj = Round(MbSfShift.N1faFdTonsAdj * MbSfShift.N1FaN1frBtPct + _
                                MbSfShift.N2faFdTonsAdj * MbSfShift.N2FaN1frBtPct, 0)
        MbSfTotal.N1frCnTonsAdj = MbSfTotal.N1frCnTonsAdj + MbSfShift.N1frCnTonsAdj
        MbSfTotal.N1frCnBtAdj = MbSfTotal.N1frCnBtAdj + Round(MbSfShift.N1frCnTonsAdj * _
                              MbSfShift.N1frCnBpl, 1)

        '#2FR rougher concentrate tons -- recalculated
        '#2FR rougher concentrate tons -- recalculated
        '#2FR rougher concentrate tons -- recalculated

        'From #2FA & #3Fa
        MbSfShift.N2FrCnTonsAdj = Round(MbSfShift.N2faFdTonsAdj * MbSfShift.N2FaN2frBtPct + _
                                MbSfShift.N3faFdTonsAdj * MbSfShift.N3FaN2frBtPct, 0)
        MbSfTotal.N2FrCnTonsAdj = MbSfTotal.N2FrCnTonsAdj + MbSfShift.N2FrCnTonsAdj
        MbSfTotal.N2FrCnBtAdj = MbSfTotal.N2FrCnBtAdj + Round(MbSfShift.N2FrCnTonsAdj * _
                              MbSfShift.N2frCnBpl, 1)

        '#3FR rougher concentrate tons -- recalculated
        '#3FR rougher concentrate tons -- recalculated
        '#3FR rougher concentrate tons -- recalculated

        'From #2FA & #3Fa
        MbSfShift.N3FrCnTonsAdj = Round(MbSfShift.N2faFdTonsAdj * MbSfShift.N2FaN3frBtPct + _
                                MbSfShift.N3faFdTonsAdj * MbSfShift.N3FaN3frBtPct, 0)
        MbSfTotal.N3FrCnTonsAdj = MbSfTotal.N3FrCnTonsAdj + MbSfShift.N3FrCnTonsAdj
        MbSfTotal.N3FrCnBtAdj = MbSfTotal.N3FrCnBtAdj + Round(MbSfShift.N3FrCnTonsAdj * _
                              MbSfShift.N3frCnBpl, 1)

        '#4FR rougher concentrate tons -- recalculated
        '#4FR rougher concentrate tons -- recalculated
        '#4FR rougher concentrate tons -- recalculated

        'From #1FA & #2Fa
        MbSfShift.N4FrCnTonsAdj = Round(MbSfShift.N1faFdTonsAdj * MbSfShift.N1FaN4frBtPct + _
                                MbSfShift.N2faFdTonsAdj * MbSfShift.N2FaN4frBtPct, 0)
        MbSfTotal.N4FrCnTonsAdj = MbSfTotal.N4FrCnTonsAdj + MbSfShift.N4FrCnTonsAdj
        MbSfTotal.N4FrCnBtAdj = MbSfTotal.N4FrCnBtAdj + Round(MbSfShift.N4FrCnTonsAdj * _
                              MbSfShift.N4frCnBpl, 1)

        '#5FR rougher concentrate tons -- recalculated
        '#5FR rougher concentrate tons -- recalculated
        '#5FR rougher concentrate tons -- recalculated

        'From #1FA & #2Fa
        MbSfShift.N5FrCnTonsAdj = Round(MbSfShift.N1faFdTonsAdj * MbSfShift.N1FaN5frBtPct + _
                                MbSfShift.N2faFdTonsAdj * MbSfShift.N2FaN5frBtPct, 0)
        MbSfTotal.N5FrCnTonsAdj = MbSfTotal.N5FrCnTonsAdj + MbSfShift.N5FrCnTonsAdj
        MbSfTotal.N5FrCnBtAdj = MbSfTotal.N5FrCnBtAdj + Round(MbSfShift.N5FrCnTonsAdj * _
                              MbSfShift.N5frCnBpl, 1)

        'SR rougher concentrate tons -- recalculated
        'SR rougher concentrate tons -- recalculated
        'SR rougher concentrate tons -- recalculated

        'From #1FA, #2FA & #3Fa
        'Also from the Ca -- 05/23/2005, lss
        'As of 05/2005 the coarse amine circuit can receive rougher
        'concentrate from both the SR and the CR
        MbSfShift.SrCnTonsAdj = Round(MbSfShift.N1faFdTonsAdj * MbSfShift.N1FaSrBtPct + _
                              MbSfShift.N2faFdTonsAdj * MbSfShift.N2FaSrBtPct + _
                              MbSfShift.N3faFdTonsAdj * MbSfShift.N3FaSrBtPct + _
                              MbSfShift.CaFdTonsAdj * MbSfShift.CaSrBtPct, 0)
        MbSfTotal.SrCnTonsAdj = MbSfTotal.SrCnTonsAdj + MbSfShift.SrCnTonsAdj
        MbSfTotal.SrCnBtAdj = MbSfTotal.SrCnBtAdj + Round(MbSfShift.SrCnTonsAdj * _
                              MbSfShift.SrCnBpl, 1)

        'CR rougher concentrate tons -- recalculated
        'CR rougher concentrate tons -- recalculated
        'CR rougher concentrate tons -- recalculated
        'From CA only -- 01/23/2005, lss
        'As of 05/2005 the coarse amine circuit can receive rougher
        'concentrate from both the SR and the CR.  Thus the CR may not get
        'all of the adjusted tons from the CA -- some or all of it may
        'actually go to the SR.
        'MbSfShift.CrCnTonsAdj = MbSfShift.CaFdTonsAdj
        MbSfShift.CrCnTonsAdj = Round(MbSfShift.CaFdTonsAdj * _
                                MbSfShift.CaCrBtPct, 0)

        MbSfTotal.CrCnTonsAdj = MbSfTotal.CrCnTonsAdj + MbSfShift.CrCnTonsAdj
        MbSfTotal.CrCnBtAdj = MbSfTotal.CrCnBtAdj + Round(MbSfShift.CrCnTonsAdj * _
                              MbSfShift.CrCnBpl, 1)

        'Tails  Tails  Tails  Tails  Tails  Tails
        'Tails  Tails  Tails  Tails  Tails  Tails
        'Tails  Tails  Tails  Tails  Tails  Tails
        '#1FA tail tons & tail BPL-tons
        'MbSfShift.N1faFdTonsAdj -- adjusted 1FA feed tons
        'MbSfShift.N1faCnTons -- #1FA concentrate tons
        MbSfShift.N1faTlTonsAdj = MbSfShift.N1faFdTonsAdj - MbSfShift.N1faCnTons
        MbSfTotal.N1faTlTonsAdj = MbSfTotal.N1faTlTonsAdj + MbSfShift.N1faTlTonsAdj
        MbSfTotal.N1faTlBtAdj = MbSfTotal.N1faTlBtAdj + Round(MbSfShift.N1faTlTonsAdj * _
                              MbSfShift.N1faTlBpl, 1)

        '#2FA tail tons & tail BPL-tons
        MbSfShift.N2faTlTonsAdj = MbSfShift.N2faFdTonsAdj - MbSfShift.N2faCnTons
        MbSfTotal.N2faTlTonsAdj = MbSfTotal.N2faTlTonsAdj + MbSfShift.N2faTlTonsAdj
        MbSfTotal.N2faTlBtAdj = MbSfTotal.N2faTlBtAdj + Round(MbSfShift.N2faTlTonsAdj * _
                              MbSfShift.N2faTlBpl, 1)

        '#3FA tail tons & tail BPL-tons
        MbSfShift.N3faTlTonsAdj = MbSfShift.N3faFdTonsAdj - MbSfShift.N3faCnTons
        MbSfTotal.N3faTlTonsAdj = MbSfTotal.N3faTlTonsAdj + MbSfShift.N3faTlTonsAdj
        MbSfTotal.N3faTlBtAdj = MbSfTotal.N3faTlBtAdj + Round(MbSfShift.N3faTlTonsAdj * _
                              MbSfShift.N3faTlBpl, 1)

        'CA tail tons & tail BPL-tons
        MbSfShift.CaTlTonsAdj = MbSfShift.CaFdTonsAdj - MbSfShift.CaCnTons
        MbSfTotal.CaTlTonsAdj = MbSfTotal.CaTlTonsAdj + MbSfShift.CaTlTonsAdj
        MbSfTotal.CaTlBtAdj = MbSfTotal.CaTlBtAdj + Round(MbSfShift.CaTlTonsAdj * _
                              MbSfShift.CaTlBpl, 1)

        'Total fine amine feed tons (#1FA, #2FA, #3FA)
        MbSfShift.TotFaFdTonsAdj = MbSfShift.N1faFdTonsAdj + MbSfShift.N2faFdTonsAdj + _
                                 MbSfShift.N3faFdTonsAdj
        MbSfTotal.TotFaFdTonsAdj = MbSfTotal.TotFaFdTonsAdj + MbSfShift.TotFaFdTonsAdj

        'Total fine amine ratio of concentration
        'MbSfShift.PrdFcnTons -- measured shift total fine concentrate tons
        'Ratio of concentrations are rounded to two decimal places
        If MbSfShift.PrdFcnTons <> 0 Then
            MbSfShift.TotFaRc = Round(MbSfShift.TotFaFdTonsAdj / MbSfShift.PrdFcnTons, 2)
        Else
            MbSfShift.TotFaRc = 0
        End If

        'Determine adjusted #4FR feed tons
        'MbSfShift.N4FrCnTonsAdj -- rougher concentrate tons from the #4FR circuit
        MbSfShift.N4frFdTonsAdj = Round(MbSfShift.N4frRc * MbSfShift.N4FrCnTonsAdj, 0)
        MbSfTotal.N4frFdTonsAdj = MbSfTotal.N4frFdTonsAdj + MbSfShift.N4frFdTonsAdj
        MbSfTotal.N4frFdBtAdj = MbSfTotal.N4frFdBtAdj + Round(MbSfShift.N4frFdTonsAdj * _
                              MbSfShift.N4frFdBpl, 1)

        'Determine adjusted #5FR feed tons
        'MbSfShift.N5frCnTonsAdj -- rougher concentrate tons from the #5FR circuit
        MbSfShift.N5frFdTonsAdj = Round(MbSfShift.N5frRc * MbSfShift.N5FrCnTonsAdj, 0)
        MbSfTotal.N5frFdTonsAdj = MbSfTotal.N5frFdTonsAdj + MbSfShift.N5frFdTonsAdj
        MbSfTotal.N5frFdBtAdj = MbSfTotal.N5frFdBtAdj + Round(MbSfShift.N5frFdTonsAdj * _
                              MbSfShift.N5frFdBpl, 1)

        'Determine adjusted #1FR feed tons
        'MbSfShift.N1frCnTonsAdj -- rougher concentrate tons from the #1FR circuit
        MbSfShift.N1frFdTonsAdj = Round(MbSfShift.N1frRc * MbSfShift.N1frCnTonsAdj, 0)
        MbSfTotal.N1frFdTonsAdj = MbSfTotal.N1frFdTonsAdj + MbSfShift.N1frFdTonsAdj
        MbSfTotal.N1frFdBtAdj = MbSfTotal.N1frFdBtAdj + Round(MbSfShift.N1frFdTonsAdj * _
                              MbSfShift.N1frFdBpl, 1)

        'Determine adjusted #2FR feed tons
        'MbSfShift.N2frCnTonsAdj -- rougher concentrate tons from the #2FR circuit
        MbSfShift.N2frFdTonsAdj = Round(MbSfShift.N2frRc * MbSfShift.N2FrCnTonsAdj, 0)
        MbSfTotal.N2frFdTonsAdj = MbSfTotal.N2frFdTonsAdj + MbSfShift.N2frFdTonsAdj
        MbSfTotal.N2frFdBtAdj = MbSfTotal.N2frFdBtAdj + Round(MbSfShift.N2frFdTonsAdj * _
                              MbSfShift.N2frFdBpl, 1)

        'Determine adjusted #3FR feed tons
        'MbSfShift.N3frCnTonsAdj -- rougher concentrate tons from the #3FR circuit
        MbSfShift.N3frFdTonsAdj = Round(MbSfShift.N3frRc * MbSfShift.N3FrCnTonsAdj, 0)
        MbSfTotal.N3frFdTonsAdj = MbSfTotal.N3frFdTonsAdj + MbSfShift.N3frFdTonsAdj
        MbSfTotal.N3frFdBtAdj = MbSfTotal.N3frFdBtAdj + Round(MbSfShift.N3frFdTonsAdj * _
                              MbSfShift.N3frFdBpl, 1)

        'Determine adjusted SR feed tons
        'MbSfShift.SrCnTonsAdj -- rougher concentrate tons from the SR circuit
        MbSfShift.SrFdTonsAdj = Round(MbSfShift.SrRc * MbSfShift.SrCnTonsAdj, 0)
        MbSfTotal.SrFdTonsAdj = MbSfTotal.SrFdTonsAdj + MbSfShift.SrFdTonsAdj
        MbSfTotal.SrFdBtAdj = MbSfTotal.SrFdBtAdj + Round(MbSfShift.SrFdTonsAdj * _
                              MbSfShift.SrFdBpl, 1)

        'Determine adjusted CR feed tons
        'MbSfShift.CrCnTonsAdj -- rougher concentrate tons from the CR circuit
        MbSfShift.CrFdTonsAdj = Round(MbSfShift.CrRc * MbSfShift.CrCnTonsAdj, 0)
        MbSfTotal.CrFdTonsAdj = MbSfTotal.CrFdTonsAdj + MbSfShift.CrFdTonsAdj
        MbSfTotal.CrFdBtAdj = MbSfTotal.CrFdBtAdj + Round(MbSfShift.CrFdTonsAdj * _
                              MbSfShift.CrFdBpl, 1)

        '#1 fine rougher  #1 fine rougher  #1 fine rougher
        '#1 fine rougher  #1 fine rougher  #1 fine rougher
        '#1 fine rougher  #1 fine rougher  #1 fine rougher
        MbSfShift.N1frTlTonsAdj = MbSfShift.N1frFdTonsAdj - MbSfShift.N1frCnTonsAdj
        MbSfTotal.N1frTlTonsAdj = MbSfTotal.N1frTlTonsAdj + MbSfShift.N1frTlTonsAdj
        MbSfTotal.N1frTlBtAdj = MbSfTotal.N1frTlBtAdj + Round(MbSfShift.N1frTlTonsAdj * _
                              MbSfShift.N1frTlBpl, 1)

        '#2 fine rougher  #1 fine rougher  #1 fine rougher
        '#2 fine rougher  #1 fine rougher  #1 fine rougher
        '#2 fine rougher  #1 fine rougher  #1 fine rougher
        MbSfShift.N2frTlTonsAdj = MbSfShift.N2frFdTonsAdj - MbSfShift.N2FrCnTonsAdj
        MbSfTotal.N2frTlTonsAdj = MbSfTotal.N2frTlTonsAdj + MbSfShift.N2frTlTonsAdj
        MbSfTotal.N2frTlBtAdj = MbSfTotal.N2frTlBtAdj + Round(MbSfShift.N2frTlTonsAdj * _
                              MbSfShift.N2frTlBpl, 1)

        '#3 fine rougher  #1 fine rougher  #1 fine rougher
        '#3 fine rougher  #1 fine rougher  #1 fine rougher
        '#3 fine rougher  #1 fine rougher  #1 fine rougher
        MbSfShift.N3frTlTonsAdj = MbSfShift.N3frFdTonsAdj - MbSfShift.N3FrCnTonsAdj
        MbSfTotal.N3frTlTonsAdj = MbSfTotal.N3frTlTonsAdj + MbSfShift.N3frTlTonsAdj
        MbSfTotal.N3frTlBtAdj = MbSfTotal.N3frTlBtAdj + Round(MbSfShift.N3frTlTonsAdj * _
                              MbSfShift.N3frTlBpl, 1)

        '#4 fine rougher  #4 fine rougher  #4 fine rougher
        '#4 fine rougher  #4 fine rougher  #4 fine rougher
        '#4 fine rougher  #4 fine rougher  #4 fine rougher
        '#4FR tail tons (shift & total)
        MbSfShift.N4frTlTonsAdj = MbSfShift.N4frFdTonsAdj - MbSfShift.N4FrCnTonsAdj
        MbSfTotal.N4frTlTonsAdj = MbSfTotal.N4frTlTonsAdj + MbSfShift.N4frTlTonsAdj
        MbSfTotal.N4frTlBtAdj = MbSfTotal.N4frTlBtAdj + Round(MbSfShift.N4frTlTonsAdj * _
                              MbSfShift.N4frTlBpl, 1)

        '#5 fine rougher  #5 fine rougher  #5 fine rougher
        '#5 fine rougher  #5 fine rougher  #5 fine rougher
        '#5 fine rougher  #5 fine rougher  #5 fine rougher
        MbSfShift.N5frTlTonsAdj = MbSfShift.N5frFdTonsAdj - MbSfShift.N5FrCnTonsAdj
        MbSfTotal.N5frTlTonsAdj = MbSfTotal.N5frTlTonsAdj + MbSfShift.N5frTlTonsAdj
        MbSfTotal.N5frTlBtAdj = MbSfTotal.N5frTlBtAdj + Round(MbSfShift.N5frTlTonsAdj * _
                              MbSfShift.N5frTlBpl, 1)

        '#Coarse rougher  Coarse rougher  Coarse rougher
        '#Coarse rougher  Coarse rougher  Coarse rougher
        '#Coarse rougher  Coarse rougher  Coarse rougher
        MbSfShift.CrTlTonsAdj = MbSfShift.CrFdTonsAdj - MbSfShift.CrCnTonsAdj
        MbSfTotal.CrTlTonsAdj = MbSfTotal.CrTlTonsAdj + MbSfShift.CrTlTonsAdj
        MbSfTotal.CrTlBtAdj = MbSfTotal.CrTlBtAdj + Round(MbSfShift.CrTlTonsAdj * _
                              MbSfShift.CrTlBpl, 1)

        '#Swing rougher  Swing rougher  Swing rougher
        '#Swing rougher  Swing rougher  Swing rougher
        '#Swing rougher  Swing rougher  Swing rougher
        MbSfShift.SrTlTonsAdj = MbSfShift.SrFdTonsAdj - MbSfShift.SrCnTonsAdj
        MbSfTotal.SrTlTonsAdj = MbSfTotal.SrTlTonsAdj + MbSfShift.SrTlTonsAdj
        MbSfTotal.SrTlBtAdj = MbSfTotal.SrTlBtAdj + Round(MbSfShift.SrTlTonsAdj * _
                              MbSfShift.SrTlBpl, 1)

        'Ultra-coarse rougher  Ultra-coarse rougher
        'Ultra-coarse rougher  Ultra-coarse rougher
        'Ultra-coarse rougher  Ultra-coarse rougher
        'Adjusted total ultra coarse feed tons
        'Ratio of concentration from circuit BPL's * ultra-coarse
        'production tons
        MbSfShift.UcrFdTonsAdj = Round(MbSfShift.PrdUccnTons * MbSfShift.UcrRc, 0)
        MbSfTotal.UcrFdTonsAdj = MbSfTotal.UcrFdTonsAdj + MbSfShift.UcrFdTonsAdj
        MbSfTotal.UcrFdBtAdj = MbSfTotal.UcrFdBtAdj + Round(MbSfShift.UcrFdTonsAdj * _
                             MbSfShift.UcrFdBpl, 1)
        MbSfShift.UcrCnTonsAdj = MbSfShift.PrdUccnTons
        MbSfTotal.UcrCnTonsAdj = MbSfTotal.UcrCnTonsAdj + MbSfShift.UcrCnTonsAdj
        MbSfShift.UcrTlTonsAdj = MbSfShift.UcrFdTonsAdj - MbSfShift.UcrCnTonsAdj
        MbSfTotal.UcrTlTonsAdj = MbSfTotal.UcrTlTonsAdj + MbSfShift.UcrTlTonsAdj
        MbSfTotal.UcrTlBtAdj = MbSfTotal.UcrTlBtAdj + Round(MbSfShift.UcrTlTonsAdj * _
                             MbSfShift.UcrTlBpl, 1)

        '#1 fine amine  #1 fine amine  #1 fine amine
        '#1 fine amine  #1 fine amine  #1 fine amine
        '#1 fine amine  #1 fine amine  #1 fine amine
        'No further calculations necessary

        '#2 fine amine  #2 fine amine  #2 fine amine
        '#2 fine amine  #2 fine amine  #2 fine amine
        '#2 fine amine  #2 fine amine  #2 fine amine
        'No further calculations necessary

        '#3 fine amine  #3 fine amine  #3 fine amine
        '#3 fine amine  #3 fine amine  #3 fine amine
        '#3 fine amine  #3 fine amine  #3 fine amine
        'No further calculations necessary

        'Coarse amine  Coarse amine  Coarse amine
        'Coarse amine  Coarse amine  Coarse amine
        'Coarse amine  Coarse amine  Coarse amine
        'No further calculations necessary

        'Miscellaneous  Miscellaneous  Miscellaneous
        'Miscellaneous  Miscellaneous  Miscellaneous
        'Miscellaneous  Miscellaneous  Miscellaneous

        'GMT BPL

        'Total adjusted feed tons for shift
        MbSfShift.TotRghrFdTonsAdj = MbSfShift.N1frFdTonsAdj + MbSfShift.N2frFdTonsAdj + _
                                   MbSfShift.N3frFdTonsAdj + MbSfShift.N4frFdTonsAdj + _
                                   MbSfShift.N5frFdTonsAdj + MbSfShift.SrFdTonsAdj + _
                                   MbSfShift.CrFdTonsAdj + MbSfShift.UcrFdTonsAdj

        'Total final concentrate tons for shift
        MbSfShift.TotTlTons = MbSfShift.TotRghrFdTonsAdj - MbSfShift.PrdTotCnTons
        If MbSfShift.GmtBpl <> 0 Then
            MbSfTotal.TotGmtTlTonsW = MbSfTotal.TotGmtTlTonsW + MbSfShift.TotTlTons
        End If
        MbSfTotal.TotGmtTlBt = MbSfTotal.TotGmtTlBt + _
                             Round(MbSfShift.GmtBpl * MbSfShift.TotTlTons, 1)

        Exit Sub

ProcessSfMassBalanceError:

        MsgBox("Error in South Fort Meade mass balance." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "South Fort Meade Mass Balance Computation Error")
    End Sub

    Private Sub ProcessSfMassBalanceTotals()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************
        Dim TonsWithBpl As Long
        Dim BplTons As Double

        On Error GoTo ProcessSfMassBalanceTotalsError

        With MbSfTotal
            'Product tons and analysis  Product tons and analysis
            'Product tons and analysis  Product tons and analysis
            'Product tons and analysis  Product tons and analysis
            'Fine concentrate tons & BPL
            'Fine concentrate tons = MbTotal.PrdFcnTons
            If .PrdFcnTonsW <> 0 Then
                .PrdFcnBpl = Round(.PrdFcnBt / .PrdFcnTonsW, 1)
            Else
                .PrdFcnBpl = 0
            End If

            'Coarse concentrate tons & BPL
            'Coarse concentrate tons & BPL
            'Coarse concentrate tons = MbTotal.PrdCcnTons
            If .PrdCcnTonsW <> 0 Then
                .PrdCcnBpl = Round(.PrdCcnBt / .PrdCcnTonsW, 1)
            Else
                .PrdCcnBpl = 0
            End If

            'Ultra-coarse concentrate tons & BPL
            'Ultra-coarse concentrate tons & BPL
            'Ultra-coarse concentrate tons = MbTotal.PrdFcnTons
            If .PrdUccnTonsW <> 0 Then
                .PrdUccnBpl = Round(.PrdUccnBt / .PrdUccnTonsW, 1)
            Else
                .PrdUccnBpl = 0
            End If

            'Total concentrate tons with Ultra-coarse & BPL
            .PrdCnTonsWuc = .PrdFcnTons + .PrdCcnTons + .PrdUccnTons
            TonsWithBpl = 0
            If .PrdFcnBpl <> 0 Then
                TonsWithBpl = TonsWithBpl + .PrdFcnTons
            End If
            If .PrdCcnBpl <> 0 Then
                TonsWithBpl = TonsWithBpl + .PrdCcnTons
            End If
            If .PrdUccnBpl <> 0 Then
                TonsWithBpl = TonsWithBpl + .PrdUccnTons
            End If
            BplTons = .PrdFcnTons * .PrdFcnBpl + _
                      .PrdCcnTons * .PrdCcnBpl + _
                      .PrdUccnTons * .PrdUccnBpl
            If TonsWithBpl <> 0 Then
                .PrdCnBplWuc = Round(BplTons / TonsWithBpl, 1)
            Else
                .PrdCnBplWuc = 0
            End If

            'Total concentrate tons without Ultra-coarse & BPL
            .PrdCnTonsWouc = .PrdFcnTons + .PrdCcnTons
            TonsWithBpl = 0
            If .PrdFcnBpl <> 0 Then
                TonsWithBpl = TonsWithBpl + .PrdFcnTons
            End If
            If .PrdCcnBpl <> 0 Then
                TonsWithBpl = TonsWithBpl + .PrdCcnTons
            End If
            BplTons = .PrdFcnTons * .PrdFcnBpl + _
                      .PrdCcnTons * .PrdCcnBpl
            If TonsWithBpl <> 0 Then
                .PrdCnBplWouc = Round(BplTons / TonsWithBpl, 1)
            Else
                .PrdCnBplWouc = 0
            End If

            '#1 fine rougher  #1 fine rougher  #1 fine rougher
            '#1 fine rougher  #1 fine rougher  #1 fine rougher
            '#1 fine rougher  #1 fine rougher  #1 fine rougher

            If .N1frFdTonsAdj <> 0 Then
                .N1frFdBpl = Round(.N1frFdBtAdj / .N1frFdTonsAdj, 1)
            Else
                .N1frFdBpl = 0
            End If
            If .N1frCnTonsAdj <> 0 Then
                .N1frCnBpl = Round(.N1frCnBtAdj / .N1frCnTonsAdj, 1)
            Else
                .N1frCnBpl = 0
            End If
            If .N1frTlTonsAdj <> 0 Then
                .N1frTlBpl = Round(.N1frTlBtAdj / .N1frTlTonsAdj, 1)
            Else
                .N1frTlBpl = 0
            End If
            If .N1frFdTonsRpt <> 0 Then
                .N1frArFdBpl = Round(.N1frFdBtRpt / .N1frFdTonsRpt, 1)
            Else
                .N1frArFdBpl = 0
            End If

            '#2 fine rougher  #2 fine rougher  #2 fine rougher
            '#2 fine rougher  #2 fine rougher  #2 fine rougher
            '#2 fine rougher  #2 fine rougher  #2 fine rougher

            If .N2frFdTonsAdj <> 0 Then
                .N2frFdBpl = Round(.N2frFdBtAdj / .N2frFdTonsAdj, 1)
            Else
                .N2frFdBpl = 0
            End If
            If .N2FrCnTonsAdj <> 0 Then
                .N2frCnBpl = Round(.N2FrCnBtAdj / .N2FrCnTonsAdj, 1)
            Else
                .N2frCnBpl = 0
            End If
            If .N2frTlTonsAdj <> 0 Then
                .N2frTlBpl = Round(.N2frTlBtAdj / .N2frTlTonsAdj, 1)
            Else
                .N2frTlBpl = 0
            End If
            If .N2frFdTonsRpt <> 0 Then
                .N2frArFdBpl = Round(.N2frFdBtRpt / .N2frFdTonsRpt, 1)
            Else
                .N2frArFdBpl = 0
            End If

            '#3 fine rougher  #3 fine rougher  #3 fine rougher
            '#3 fine rougher  #3 fine rougher  #3 fine rougher
            '#3 fine rougher  #3 fine rougher  #3 fine rougher

            If .N3frFdTonsAdj <> 0 Then
                .N3frFdBpl = Round(.N3frFdBtAdj / .N3frFdTonsAdj, 1)
            Else
                .N3frFdBpl = 0
            End If
            If .N3FrCnTonsAdj <> 0 Then
                .N3frCnBpl = Round(.N3FrCnBtAdj / .N3FrCnTonsAdj, 1)
            Else
                .N3frCnBpl = 0
            End If
            If .N3frTlTonsAdj <> 0 Then
                .N3frTlBpl = Round(.N3frTlBtAdj / .N3frTlTonsAdj, 1)
            Else
                .N3frTlBpl = 0
            End If
            If .N3frFdTonsRpt <> 0 Then
                .N3frArFdBpl = Round(.N3frFdBtRpt / .N3frFdTonsRpt, 1)
            Else
                .N3frArFdBpl = 0
            End If

            '#4 fine rougher  #4 fine rougher  #4 fine rougher
            '#4 fine rougher  #4 fine rougher  #4 fine rougher
            '#4 fine rougher  #4 fine rougher  #4 fine rougher

            If .N4frFdTonsAdj <> 0 Then
                .N4frFdBpl = Round(.N4frFdBtAdj / .N4frFdTonsAdj, 1)
            Else
                .N4frFdBpl = 0
            End If
            If .N4FrCnTonsAdj <> 0 Then
                .N4frCnBpl = Round(.N4FrCnBtAdj / .N4FrCnTonsAdj, 1)
            Else
                .N4frCnBpl = 0
            End If
            If .N4frTlTonsAdj <> 0 Then
                .N4frTlBpl = Round(.N4frTlBtAdj / .N4frTlTonsAdj, 1)
            Else
                .N4frTlBpl = 0
            End If
            If .N4frFdTonsRpt <> 0 Then
                .N4frArFdBpl = Round(.N4frFdBtRpt / .N4frFdTonsRpt, 1)
            Else
                .N4frArFdBpl = 0
            End If

            '#5 fine rougher  #5 fine rougher  #5 fine rougher
            '#5 fine rougher  #5 fine rougher  #5 fine rougher
            '#5 fine rougher  #5 fine rougher  #5 fine rougher

            If .N5frFdTonsAdj <> 0 Then
                .N5frFdBpl = Round(.N5frFdBtAdj / .N5frFdTonsAdj, 1)
            Else
                .N5frFdBpl = 0
            End If
            If .N5FrCnTonsAdj <> 0 Then
                .N5frCnBpl = Round(.N5FrCnBtAdj / .N5FrCnTonsAdj, 1)
            Else
                .N5frCnBpl = 0
            End If
            If .N5frTlTonsAdj <> 0 Then
                .N5frTlBpl = Round(.N5frTlBtAdj / .N5frTlTonsAdj, 1)
            Else
                .N5frTlBpl = 0
            End If
            If .N5frFdTonsRpt <> 0 Then
                .N5frArFdBpl = Round(.N5frFdBtRpt / .N5frFdTonsRpt, 1)
            Else
                .N5frArFdBpl = 0
            End If

            'Coarse rougher  Coarse rougher  Coarse rougher
            'Coarse rougher  Coarse rougher  Coarse rougher
            'Coarse rougher  Coarse rougher  Coarse rougher

            If .CrFdTonsAdj <> 0 Then
                .CrFdBpl = Round(.CrFdBtAdj / .CrFdTonsAdj, 1)
            Else
                .CrFdBpl = 0
            End If
            If .CrCnTonsAdj <> 0 Then
                .CrCnBpl = Round(.CrCnBtAdj / .CrCnTonsAdj, 1)
            Else
                .CrCnBpl = 0
            End If
            If .CrTlTonsAdj <> 0 Then
                .CrTlBpl = Round(.CrTlBtAdj / .CrTlTonsAdj, 1)
            Else
                .CrTlBpl = 0
            End If
            If .CrFdTonsRpt <> 0 Then
                .CrArFdBpl = Round(.CrFdBtRpt / .CrFdTonsRpt, 1)
            Else
                .CrArFdBpl = 0
            End If

            'Swing rougher  Swing rougher  Swing rougher
            'Swing rougher  Swing rougher  Swing rougher
            'Swing rougher  Swing rougher  Swing rougher

            If .SrFdTonsAdj <> 0 Then
                .SrFdBpl = Round(.SrFdBtAdj / .SrFdTonsAdj, 1)
            Else
                .SrFdBpl = 0
            End If
            If .SrCnTonsAdj <> 0 Then
                .SrCnBpl = Round(.SrCnBtAdj / .SrCnTonsAdj, 1)
            Else
                .SrCnBpl = 0
            End If
            If .SrTlTonsAdj <> 0 Then
                .SrTlBpl = Round(.SrTlBtAdj / .SrTlTonsAdj, 1)
            Else
                .SrTlBpl = 0
            End If
            If .SrFdTonsRpt <> 0 Then
                .SrArFdBpl = Round(.SrFdBtRpt / .SrFdTonsRpt, 1)
            Else
                .SrArFdBpl = 0
            End If

            'Ultra-coarse rougher  Ultra-coarse rougher  Ultra-coarse rougher
            'Ultra-coarse rougher  Ultra-coarse rougher  Ultra-coarse rougher
            'Ultra-coarse rougher  Ultra-coarse rougher  Ultra-coarse rougher

            If .UcrFdTonsAdj <> 0 Then
                .UcrFdBpl = Round(.UcrFdBtAdj / .UcrFdTonsAdj, 1)
            Else
                .UcrFdBpl = 0
            End If

            .UcrCnBpl = .PrdUccnBpl

            If .UcrTlTonsAdj <> 0 Then
                .UcrTlBpl = Round(.UcrTlBtAdj / .UcrTlTonsAdj, 1)
            Else
                .UcrTlBpl = 0
            End If
            If .UcrFdTonsRpt <> 0 Then
                .UcrArFdBpl = Round(.UcrFdBtRpt / .UcrFdTonsRpt, 1)
            Else
                .UcrArFdBpl = 0
            End If

            '#1 fine amine  #1 fine amine  #1 fine amine
            '#1 fine amine  #1 fine amine  #1 fine amine
            '#1 fine amine  #1 fine amine  #1 fine amine
            If .N1faFdTonsAdj <> 0 Then
                .N1faFdBpl = Round(.N1faFdBtAdj / .N1faFdTonsAdj, 1)
            Else
                .N1faFdBpl = 0
            End If

            '10/30/2011, lss
            If mUseFaCnChange = False Then
                .N1faCnBpl = .PrdFcnBpl
            Else
                If .N1faCnTons <> 0 Then
                    .N1faCnBpl = Round(.N1faCnBtAdj / .N1faCnTons, 1)
                Else
                    .N1faCnBpl = 0
                End If
            End If

            If .N1faTlTonsAdj <> 0 Then
                .N1faTlBpl = Round(.N1faTlBtAdj / .N1faTlTonsAdj, 1)
            Else
                .N1faTlBpl = 0
            End If

            '#2 fine amine  #2 fine amine  #2 fine amine
            '#2 fine amine  #2 fine amine  #2 fine amine
            '#2 fine amine  #2 fine amine  #2 fine amine
            If .N2faFdTonsAdj <> 0 Then
                .N2faFdBpl = Round(.N2faFdBtAdj / .N2faFdTonsAdj, 1)
            Else
                .N2faFdBpl = 0
            End If

            '10/30/2011, lss
            If mUseFaCnChange = False Then
                .N2faCnBpl = .PrdFcnBpl
            Else
                If .N2faCnTons <> 0 Then
                    .N2faCnBpl = Round(.N2faCnBtAdj / .N2faCnTons, 1)
                Else
                    .N2faCnBpl = 0
                End If
            End If

            If .N2faTlTonsAdj <> 0 Then
                .N2faTlBpl = Round(.N2faTlBtAdj / .N2faTlTonsAdj, 1)
            Else
                .N2faTlBpl = 0
            End If

            '#3 fine amine  #3 fine amine  #3 fine amine
            '#3 fine amine  #3 fine amine  #3 fine amine
            '#3 fine amine  #3 fine amine  #3 fine amine
            If .N3faFdTonsAdj <> 0 Then
                .N3faFdBpl = Round(.N3faFdBtAdj / .N3faFdTonsAdj, 1)
            Else
                .N3faFdBpl = 0
            End If

            '10/30/2011, lss
            If mUseFaCnChange = False Then
                .N3faCnBpl = .PrdFcnBpl
            Else
                If .N3faCnTons <> 0 Then
                    .N3faCnBpl = Round(.N3faCnBtAdj / .N3faCnTons, 1)
                Else
                    .N3faCnBpl = 0
                End If
            End If

            If .N3faTlTonsAdj <> 0 Then
                .N3faTlBpl = Round(.N3faTlBtAdj / .N3faTlTonsAdj, 1)
            Else
                .N3faTlBpl = 0
            End If

            'Coarse amine  Coarse amine  Coarse amine
            'Coarse amine  Coarse amine  Coarse amine
            'Coarse amine  Coarse amine  Coarse amine
            If .CaFdTonsAdj <> 0 Then
                .CaFdBpl = Round(.CaFdBtAdj / .CaFdTonsAdj, 1)
            Else
                .CaFdBpl = 0
            End If

            .CaCnBpl = .PrdCcnBpl

            If .CaTlTonsAdj <> 0 Then
                .CaTlBpl = Round(.CaTlBtAdj / .CaTlTonsAdj, 1)
            Else
                .CaTlBpl = 0
            End If

            'Ratios of concentration  Ratios of concentration
            'Ratios of concentration  Ratios of concentration
            'Ratios of concentration  Ratios of concentration
            '#4 fine rougher ratio of concentration
            If .N4frFdBpl - .N4frTlBpl <> 0 Then
                .N4frRc = Round((.N4frCnBpl - .N4frTlBpl) / _
                                 (.N4frFdBpl - .N4frTlBpl), 2)
            Else
                .N4frRc = 0
            End If

            '#5 fine rougher ratio of concentration
            If .N5frFdBpl - .N5frTlBpl <> 0 Then
                .N5frRc = Round((.N5frCnBpl - .N5frTlBpl) / _
                                 (.N5frFdBpl - .N5frTlBpl), 2)
            Else
                .N5frRc = 0
            End If

            '#1 fine rougher ratio of concentration
            If .N1frFdBpl - .N1frTlBpl <> 0 Then
                .N1frRc = Round((.N1frCnBpl - .N1frTlBpl) / _
                                 (.N1frFdBpl - .N1frTlBpl), 2)
            Else
                .N1frRc = 0
            End If

            '#2 fine rougher ratio of concentration
            If .N2frFdBpl - .N2frTlBpl <> 0 Then
                .N2frRc = Round((.N2frCnBpl - .N2frTlBpl) / _
                                 (.N2frFdBpl - .N2frTlBpl), 2)
            Else
                .N2frRc = 0
            End If

            '#3 fine rougher ratio of concentration
            If .N3frFdBpl - .N3frTlBpl <> 0 Then
                .N3frRc = Round((.N3frCnBpl - .N3frTlBpl) / _
                                 (.N3frFdBpl - .N3frTlBpl), 2)
            Else
                .N3frRc = 0
            End If

            'Swing rougher ratio of concentration
            If .SrFdBpl - .SrTlBpl <> 0 Then
                .SrRc = Round((.SrCnBpl - .SrTlBpl) / _
                                 (.SrFdBpl - .SrTlBpl), 2)
            Else
                .SrRc = 0
            End If

            'Coarse rougher ratio of concentration
            If .CrFdBpl - .CrTlBpl <> 0 Then
                .CrRc = Round((.CrCnBpl - .CrTlBpl) / _
                                 (.CrFdBpl - .CrTlBpl), 2)
            Else
                .CrRc = 0
            End If

            'Ultra-coarse rougher ratio of concentration
            If .UcrFdBpl - .UcrTlBpl <> 0 Then
                .UcrRc = Round((.UcrCnBpl - .UcrTlBpl) / _
                                 (.UcrFdBpl - .UcrTlBpl), 2)
            Else
                .UcrRc = 0
            End If

            '#1 fine amine ratio of concentration
            If .N1faFdBpl - .N1faTlBpl <> 0 Then
                .N1faRc = Round((.N1faCnBpl - .N1faTlBpl) / _
                                 (.N1faFdBpl - .N1faTlBpl), 2)
            Else
                .N1faRc = 0
            End If

            '#2 fine amine ratio of concentration
            If .N2faFdBpl - .N2faTlBpl <> 0 Then
                .N2faRc = Round((.N2faCnBpl - .N2faTlBpl) / _
                                 (.N2faFdBpl - .N2faTlBpl), 2)
            Else
                .N2faRc = 0
            End If

            '#3 fine amine ratio of concentration
            If .N3faFdBpl - .N3faTlBpl <> 0 Then
                .N3faRc = Round((.N3faCnBpl - .N3faTlBpl) / _
                                 (.N3faFdBpl - .N3faTlBpl), 2)
            Else
                .N3faRc = 0
            End If

            'Coarse amine ratio of concentration
            If .CaFdBpl - .CaTlBpl <> 0 Then
                .CaRc = Round((.CaCnBpl - .CaTlBpl) / _
                                 (.CaFdBpl - .CaTlBpl), 2)
            Else
                .CaRc = 0
            End If

            '#1 fine rougher  #1 fine rougher  #1 fine rougher
            '#1 fine rougher  #1 fine rougher  #1 fine rougher
            '#1 fine rougher  #1 fine rougher  #1 fine rougher
            If .N1frHrs <> 0 Then
                .N1frFdTph = Round(.N1frFdTonsAdj / .N1frHrs, 0)
            Else
                .N1frFdTph = 0
            End If
            If .N1frHrs <> 0 Then
                .N1frCnTph = Round(.N1frCnTonsAdj / .N1frHrs, 0)
            Else
                .N1frCnTph = 0
            End If
            If .N1frHrs <> 0 Then
                .N1frTlTph = Round(.N1frTlTonsAdj / .N1frHrs, 0)
            Else
                .N1frTlTph = 0
            End If
            If .N1frFdBpl * .N1frFdTonsAdj <> 0 Then
                .N1frActPctRcvry = Round((.N1frCnBpl * .N1frCnTonsAdj) / _
                                      (.N1frFdBpl * .N1frFdTonsAdj) * 100, 1)
            Else
                .N1frActPctRcvry = 0
            End If

            If .N1frFdBpl >= 0 Then
                If .N1frFdBpl - Sqrt(.N1frFdBpl) <> 0 Then
                    .N1frStdRc = Round((.N1frCnBpl - Sqrt(.N1frFdBpl)) / _
                               (.N1frFdBpl - Sqrt(.N1frFdBpl)), 1)
                Else
                    .N1frStdRc = 0
                End If
            Else
                .N1frStdRc = 0
            End If

            If .N1frStdRc <> 0 Then
                .N1frStdCnTons = Round(.N1frFdTonsAdj / .N1frStdRc, 0)
            Else
                .N1frStdCnTons = 0
            End If
            If .N1frFdBpl * .N1frFdTonsAdj <> 0 Then
                .N1frStdPctRcvry = Round((.N1frCnBpl * .N1frStdCnTons) / _
                                      (.N1frFdBpl * .N1frFdTonsAdj) * 100, 1)
            Else
                .N1frStdPctRcvry = 0
            End If

            '#2 fine rougher  #2 fine rougher  #2 fine rougher
            '#2 fine rougher  #2 fine rougher  #2 fine rougher
            '#2 fine rougher  #2 fine rougher  #2 fine rougher
            If .N2frHrs <> 0 Then
                .N2frFdTph = Round(.N2frFdTonsAdj / .N2frHrs, 0)
            Else
                .N2frFdTph = 0
            End If
            If .N2frHrs <> 0 Then
                .N2frCnTph = Round(.N2FrCnTonsAdj / .N2frHrs, 0)
            Else
                .N2frCnTph = 0
            End If
            If .N2frHrs <> 0 Then
                .N2frTlTph = Round(.N2frTlTonsAdj / .N2frHrs, 0)
            Else
                .N2frTlTph = 0
            End If
            If .N2frFdBpl * .N2frFdTonsAdj <> 0 Then
                .N2frActPctRcvry = Round((.N2frCnBpl * .N2FrCnTonsAdj) / _
                                      (.N2frFdBpl * .N2frFdTonsAdj) * 100, 1)
            Else
                .N2frActPctRcvry = 0
            End If

            If .N2frFdBpl >= 0 Then
                If .N2frFdBpl - Sqrt(.N2frFdBpl) <> 0 Then
                    .N2frStdRc = Round((.N2frCnBpl - Sqrt(.N2frFdBpl)) / _
                               (.N2frFdBpl - Sqrt(.N2frFdBpl)), 1)
                Else
                    .N2frStdRc = 0
                End If
            Else
                .N2frStdRc = 0
            End If

            If .N2frStdRc <> 0 Then
                .N2frStdCnTons = Round(.N2frFdTonsAdj / .N2frStdRc, 0)
            Else
                .N2frStdCnTons = 0
            End If
            If .N2frFdBpl * .N2frFdTonsAdj <> 0 Then
                .N2frStdPctRcvry = Round((.N2frCnBpl * .N2frStdCnTons) / _
                                      (.N2frFdBpl * .N2frFdTonsAdj) * 100, 1)
            Else
                .N2frStdPctRcvry = 0
            End If

            '#3 fine rougher  #3 fine rougher  #3 fine rougher
            '#3 fine rougher  #3 fine rougher  #3 fine rougher
            '#3 fine rougher  #3 fine rougher  #3 fine rougher
            If .N3frHrs <> 0 Then
                .N3frFdTph = Round(.N3frFdTonsAdj / .N3frHrs, 0)
            Else
                .N3frFdTph = 0
            End If
            If .N3frHrs <> 0 Then
                .N3frCnTph = Round(.N3FrCnTonsAdj / .N3frHrs, 0)
            Else
                .N3frCnTph = 0
            End If
            If .N3frHrs <> 0 Then
                .N3frTlTph = Round(.N3frTlTonsAdj / .N3frHrs, 0)
            Else
                .N3frTlTph = 0
            End If
            If .N3frFdBpl * .N3frFdTonsAdj <> 0 Then
                .N3frActPctRcvry = Round((.N3frCnBpl * .N3FrCnTonsAdj) / _
                                      (.N3frFdBpl * .N3frFdTonsAdj) * 100, 1)
            Else
                .N3frActPctRcvry = 0
            End If

            If .N3frFdBpl >= 0 Then
                If .N3frFdBpl - Sqrt(.N3frFdBpl) <> 0 Then
                    .N3frStdRc = Round((.N3frCnBpl - Sqrt(.N3frFdBpl)) / _
                               (.N3frFdBpl - Sqrt(.N3frFdBpl)), 1)
                Else
                    .N3frStdRc = 0
                End If
            Else
                .N3frStdRc = 0
            End If

            If .N3frStdRc <> 0 Then
                .N3frStdCnTons = Round(.N3frFdTonsAdj / .N3frStdRc, 0)
            Else
                .N3frStdCnTons = 0
            End If
            If .N3frFdBpl * .N3frFdTonsAdj <> 0 Then
                .N3frStdPctRcvry = Round((.N3frCnBpl * .N3frStdCnTons) / _
                                      (.N3frFdBpl * .N3frFdTonsAdj) * 100, 1)
            Else
                .N3frStdPctRcvry = 0
            End If

            '#4 fine rougher  #4 fine rougher  #4 fine rougher
            '#4 fine rougher  #4 fine rougher  #4 fine rougher
            '#4 fine rougher  #4 fine rougher  #4 fine rougher
            If .N4frHrs <> 0 Then
                .N4frFdTph = Round(.N4frFdTonsAdj / .N4frHrs, 0)
            Else
                .N4frFdTph = 0
            End If
            If .N4frHrs <> 0 Then
                .N4frCnTph = Round(.N4FrCnTonsAdj / .N4frHrs, 0)
            Else
                .N4frCnTph = 0
            End If
            If .N4frHrs <> 0 Then
                .N4frTlTph = Round(.N4frTlTonsAdj / .N4frHrs, 0)
            Else
                .N4frTlTph = 0
            End If
            If .N4frFdBpl * .N4frFdTonsAdj <> 0 Then
                .N4frActPctRcvry = Round((.N4frCnBpl * .N4FrCnTonsAdj) / _
                                      (.N4frFdBpl * .N4frFdTonsAdj) * 100, 1)
            Else
                .N4frActPctRcvry = 0
            End If

            If .N4frFdBpl >= 0 Then
                If .N4frFdBpl - Sqrt(.N4frFdBpl) <> 0 Then
                    .N4frStdRc = Round((.N4frCnBpl - Sqrt(.N4frFdBpl)) / _
                               (.N4frFdBpl - Sqrt(.N4frFdBpl)), 1)
                Else
                    .N4frStdRc = 0
                End If
            Else
                .N4frStdRc = 0
            End If

            If .N4frStdRc <> 0 Then
                .N4frStdCnTons = Round(.N4frFdTonsAdj / .N4frStdRc, 0)
            Else
                .N4frStdCnTons = 0
            End If
            If .N4frFdBpl * .N4frFdTonsAdj <> 0 Then
                .N4frStdPctRcvry = Round((.N4frCnBpl * .N4frStdCnTons) / _
                                      (.N4frFdBpl * .N4frFdTonsAdj) * 100, 1)
            Else
                .N4frStdPctRcvry = 0
            End If

            '#5 fine rougher  #5 fine rougher  #5 fine rougher
            '#5 fine rougher  #5 fine rougher  #5 fine rougher
            '#5 fine rougher  #5 fine rougher  #5 fine rougher
            If .N5frHrs <> 0 Then
                .N5frFdTph = Round(.N5frFdTonsAdj / .N5frHrs, 0)
            Else
                .N5frFdTph = 0
            End If
            If .N5frHrs <> 0 Then
                .N5frCnTph = Round(.N5FrCnTonsAdj / .N5frHrs, 0)
            Else
                .N5frCnTph = 0
            End If
            If .N5frHrs <> 0 Then
                .N5frTlTph = Round(.N5frTlTonsAdj / .N5frHrs, 0)
            Else
                .N5frTlTph = 0
            End If
            If .N5frFdBpl * .N5frFdTonsAdj <> 0 Then
                .N5frActPctRcvry = Round((.N5frCnBpl * .N5FrCnTonsAdj) / _
                                      (.N5frFdBpl * .N5frFdTonsAdj) * 100, 1)
            Else
                .N5frActPctRcvry = 0
            End If

            If .N5frFdBpl >= 0 Then
                If .N5frFdBpl - Sqrt(.N5frFdBpl) <> 0 Then
                    .N5frStdRc = Round((.N5frCnBpl - Sqrt(.N5frFdBpl)) / _
                               (.N5frFdBpl - Sqrt(.N5frFdBpl)), 1)
                Else
                    .N5frStdRc = 0
                End If
            Else
                .N5frStdRc = 0
            End If

            If .N5frStdRc <> 0 Then
                .N5frStdCnTons = Round(.N5frFdTonsAdj / .N5frStdRc, 0)
            Else
                .N5frStdCnTons = 0
            End If
            If .N5frFdBpl * .N5frFdTonsAdj <> 0 Then
                .N5frStdPctRcvry = Round((.N5frCnBpl * .N5frStdCnTons) / _
                                      (.N5frFdBpl * .N5frFdTonsAdj) * 100, 1)
            Else
                .N5frStdPctRcvry = 0
            End If

            'Swing rougher  Swing rougher  Swing rougher
            'Swing rougher  Swing rougher  Swing rougher
            'Swing rougher  Swing rougher  Swing rougher
            If .SrHrs <> 0 Then
                .SrFdTph = Round(.SrFdTonsAdj / .SrHrs, 0)
            Else
                .SrFdTph = 0
            End If
            If .SrHrs <> 0 Then
                .SrCnTph = Round(.SrCnTonsAdj / .SrHrs, 0)
            Else
                .SrCnTph = 0
            End If
            If .SrHrs <> 0 Then
                .SrTlTph = Round(.SrTlTonsAdj / .SrHrs, 0)
            Else
                .SrTlTph = 0
            End If
            If .SrFdBpl * .SrFdTonsAdj <> 0 Then
                .SrActPctRcvry = Round((.SrCnBpl * .SrCnTonsAdj) / _
                                      (.SrFdBpl * .SrFdTonsAdj) * 100, 1)
            Else
                .SrActPctRcvry = 0
            End If

            If .SrFdBpl >= 0 Then
                If .SrFdBpl - Sqrt(.SrFdBpl) <> 0 Then
                    .SrStdRc = Round((.SrCnBpl - Sqrt(.SrFdBpl)) / _
                               (.SrFdBpl - Sqrt(.SrFdBpl)), 1)
                Else
                    .SrStdRc = 0
                End If
            Else
                .SrStdRc = 0
            End If

            If .SrStdRc <> 0 Then
                .SrStdCnTons = Round(.SrFdTonsAdj / .SrStdRc, 0)
            Else
                .SrStdCnTons = 0
            End If
            If .SrFdBpl * .SrFdTonsAdj <> 0 Then
                .SrStdPctRcvry = Round((.SrCnBpl * .SrStdCnTons) / _
                                      (.SrFdBpl * .SrFdTonsAdj) * 100, 1)
            Else
                .SrStdPctRcvry = 0
            End If

            'Coarse rougher  Coarse rougher  Coarse rougher
            'Coarse rougher  Coarse rougher  Coarse rougher
            'Coarse rougher  Coarse rougher  Coarse rougher
            If .CrHrs <> 0 Then
                .CrFdTph = Round(.CrFdTonsAdj / .CrHrs, 0)
            Else
                .CrFdTph = 0
            End If
            If .CrHrs <> 0 Then
                .CrCnTph = Round(.CrCnTonsAdj / .CrHrs, 0)
            Else
                .CrCnTph = 0
            End If
            If .CrHrs <> 0 Then
                .CrTlTph = Round(.CrTlTonsAdj / .CrHrs, 0)
            Else
                .CrTlTph = 0
            End If
            If .CrFdBpl * .CrFdTonsAdj <> 0 Then
                .CrActPctRcvry = Round((.CrCnBpl * .CrCnTonsAdj) / _
                                      (.CrFdBpl * .CrFdTonsAdj) * 100, 1)
            Else
                .CrActPctRcvry = 0
            End If

            If .CrFdBpl >= 0 Then
                If .CrFdBpl - Sqrt(.CrFdBpl) <> 0 Then
                    .CrStdRc = Round((.CrCnBpl - Sqrt(.CrFdBpl)) / _
                               (.CrFdBpl - Sqrt(.CrFdBpl)), 1)
                Else
                    .CrStdRc = 0
                End If
            Else
                .CrStdRc = 0
            End If

            If .CrStdRc <> 0 Then
                .CrStdCnTons = Round(.CrFdTonsAdj / .CrStdRc, 0)
            Else
                .CrStdCnTons = 0
            End If
            If .CrFdBpl * .CrFdTonsAdj <> 0 Then
                .CrStdPctRcvry = Round((.CrCnBpl * .CrStdCnTons) / _
                                      (.CrFdBpl * .CrFdTonsAdj) * 100, 1)
            Else
                .CrStdPctRcvry = 0
            End If

            'Ultra-coarse rougher  Ultra-coarse rougher  Ultra-coarse rougher
            'Ultra-coarse rougher  Ultra-coarse rougher  Ultra-coarse rougher
            'Ultra-coarse rougher  Ultra-coarse rougher  Ultra-coarse rougher
            If .UcrHrs <> 0 Then
                .UcrFdTph = Round(.UcrFdTonsAdj / .UcrHrs, 0)
            Else
                .UcrFdTph = 0
            End If
            If .UcrHrs <> 0 Then
                .UcrCnTph = Round(.UcrCnTonsAdj / .UcrHrs, 0)
            Else
                .UcrCnTph = 0
            End If
            If .UcrHrs <> 0 Then
                .UcrTlTph = Round(.UcrTlTonsAdj / .UcrHrs, 0)
            Else
                .UcrTlTph = 0
            End If
            If .UcrFdBpl * .UcrFdTonsAdj <> 0 Then
                .UcrActPctRcvry = Round((.UcrCnBpl * .UcrCnTonsAdj) / _
                                      (.UcrFdBpl * .UcrFdTonsAdj) * 100, 1)
            Else
                .UcrActPctRcvry = 0
            End If

            If .UcrFdBpl >= 0 Then
                If .UcrFdBpl - Sqrt(.UcrFdBpl) <> 0 Then
                    .UcrStdRc = Round((.UcrCnBpl - Sqrt(.UcrFdBpl)) / _
                               (.UcrFdBpl - Sqrt(.UcrFdBpl)), 1)
                Else
                    .UcrStdRc = 0
                End If
            Else
                .UcrStdRc = 0
            End If

            If .UcrStdRc <> 0 Then
                .UcrStdCnTons = Round(.UcrFdTonsAdj / .UcrStdRc, 0)
            Else
                .UcrStdCnTons = 0
            End If
            If .UcrFdBpl * .UcrFdTonsAdj <> 0 Then
                .UcrStdPctRcvry = Round((.UcrCnBpl * .UcrStdCnTons) / _
                                      (.UcrFdBpl * .UcrFdTonsAdj) * 100, 1)
            Else
                .UcrStdPctRcvry = 0
            End If

            '#1 fine amine  #1 fine amine  #1 fine amine
            '#1 fine amine  #1 fine amine  #1 fine amine
            '#1 fine amine  #1 fine amine  #1 fine amine

            If .N1faFdBpl * .N1faFdTonsAdj <> 0 Then
                .N1faActPctRcvry = Round((.N1faCnBpl * .N1faCnTons) / _
                                      (.N1faFdBpl * .N1faFdTonsAdj) * 100, 1)
            Else
                .N1faActPctRcvry = 0
            End If

            If .N1faFdBpl >= 0 Then
                If .N1faFdBpl - Sqrt(.N1faFdBpl) <> 0 Then
                    .N1faStdRc = Round((.N1faCnBpl - Sqrt(.N1faFdBpl)) / _
                               (.N1faFdBpl - Sqrt(.N1faFdBpl)), 1)
                Else
                    .N1faStdRc = 0
                End If
            Else
                .N1faStdRc = 0
            End If

            If .N1faStdRc <> 0 Then
                .N1faStdCnTons = Round(.N1faFdTonsAdj / .N1faStdRc, 0)
            Else
                .N1faStdCnTons = 0
            End If
            If .N1faFdBpl * .N1faFdTonsAdj <> 0 Then
                .N1faStdPctRcvry = Round((.N1faCnBpl * .N1faStdCnTons) / _
                                      (.N1faFdBpl * .N1faFdTonsAdj) * 100, 1)
            Else
                .N1faStdPctRcvry = 0
            End If

            '#2 fine amine  #2 fine amine  #2 fine amine
            '#2 fine amine  #2 fine amine  #2 fine amine
            '#2 fine amine  #2 fine amine  #2 fine amine

            If .N2faFdBpl * .N2faFdTonsAdj <> 0 Then
                .N2faActPctRcvry = Round((.N2faCnBpl * .N2faCnTons) / _
                                      (.N2faFdBpl * .N2faFdTonsAdj) * 100, 1)
            Else
                .N2faActPctRcvry = 0
            End If

            If .N2faFdBpl >= 0 Then
                If .N2faFdBpl - Sqrt(.N2faFdBpl) <> 0 Then
                    .N2faStdRc = Round((.N2faCnBpl - Sqrt(.N2faFdBpl)) / _
                               (.N2faFdBpl - Sqrt(.N2faFdBpl)), 1)
                Else
                    .N2faStdRc = 0
                End If
            Else
                .N2faStdRc = 0
            End If

            If .N2faStdRc <> 0 Then
                .N2faStdCnTons = Round(.N2faFdTonsAdj / .N2faStdRc, 0)
            Else
                .N2faStdCnTons = 0
            End If
            If .N2faFdBpl * .N2faFdTonsAdj <> 0 Then
                .N2faStdPctRcvry = Round((.N2faCnBpl * .N2faStdCnTons) / _
                                      (.N2faFdBpl * .N2faFdTonsAdj) * 100, 1)
            Else
                .N2faStdPctRcvry = 0
            End If

            '#3 fine amine  #3 fine amine  #3 fine amine
            '#3 fine amine  #3 fine amine  #3 fine amine
            '#3 fine amine  #3 fine amine  #3 fine amine

            If .N3faFdBpl * .N3faFdTonsAdj <> 0 Then
                .N3faActPctRcvry = Round((.N3faCnBpl * .N3faCnTons) / _
                                      (.N3faFdBpl * .N3faFdTonsAdj) * 100, 1)
            Else
                .N3faActPctRcvry = 0
            End If

            If .N3faFdBpl >= 0 Then
                If .N3faFdBpl - Sqrt(.N3faFdBpl) <> 0 Then
                    .N3faStdRc = Round((.N3faCnBpl - Sqrt(.N3faFdBpl)) / _
                               (.N3faFdBpl - Sqrt(.N3faFdBpl)), 1)
                Else
                    .N3faStdRc = 0
                End If
            Else
                .N3faStdRc = 0
            End If

            If .N3faStdRc <> 0 Then
                .N3faStdCnTons = Round(.N3faFdTonsAdj / .N3faStdRc, 0)
            Else
                .N3faStdCnTons = 0
            End If
            If .N3faFdBpl * .N3faFdTonsAdj <> 0 Then
                .N3faStdPctRcvry = Round((.N3faCnBpl * .N3faStdCnTons) / _
                                      (.N3faFdBpl * .N3faFdTonsAdj) * 100, 1)
            Else
                .N3faStdPctRcvry = 0
            End If

            'Coarse amine  Coarse amine  Coarse amine
            'Coarse amine  Coarse amine  Coarse amine
            'Coarse amine  Coarse amine  Coarse amine

            If .CaFdBpl * .CaFdTonsAdj <> 0 Then
                .CaActPctRcvry = Round((.CaCnBpl * .CaCnTons) / _
                                      (.CaFdBpl * .CaFdTonsAdj) * 100, 1)
            Else
                .CaActPctRcvry = 0
            End If

            If .CaFdBpl >= 0 Then
                If .CaFdBpl - Sqrt(.CaFdBpl) <> 0 Then
                    .CaStdRc = Round((.CaCnBpl - Sqrt(.CaFdBpl)) / _
                               (.CaFdBpl - Sqrt(.CaFdBpl)), 1)
                Else
                    .CaStdRc = 0
                End If
            Else
                .CaStdRc = 0
            End If

            If .CaStdRc <> 0 Then
                .CaStdCnTons = Round(.CaFdTonsAdj / .CaStdRc, 0)
            Else
                .CaStdCnTons = 0
            End If
            If .CaFdBpl * .CaFdTonsAdj <> 0 Then
                .CaStdPctRcvry = Round((.CaCnBpl * .CaStdCnTons) / _
                                      (.CaFdBpl * .CaFdTonsAdj) * 100, 1)
            Else
                .CaStdPctRcvry = 0
            End If

            'Miscellaneous  Miscellaneous  Miscellaneous  Miscellaneous
            'Miscellaneous  Miscellaneous  Miscellaneous  Miscellaneous
            'Miscellaneous  Miscellaneous  Miscellaneous  Miscellaneous

            'Total adjusted feed BPL
            If (.N1frFdTonsAdj + .N2frFdTonsAdj + .N3frFdTonsAdj + _
                .N4frFdTonsAdj + .N5frFdTonsAdj + .SrFdTonsAdj + _
                .CrFdTonsAdj + .UcrFdTonsAdj) <> 0 Then
                .TotFdBplAdj = Round((.N1frFdTonsAdj * .N1frFdBpl + _
                               .N2frFdTonsAdj * .N2frFdBpl + _
                               .N3frFdTonsAdj * .N3frFdBpl + _
                               .N4frFdTonsAdj * .N4frFdBpl + _
                               .N5frFdTonsAdj * .N5frFdBpl + _
                               .SrFdTonsAdj * .SrFdBpl + _
                               .CrFdTonsAdj * .CrFdBpl + _
                               .UcrFdTonsAdj * .UcrFdBpl) / _
                               (.N1frFdTonsAdj + .N2frFdTonsAdj + _
                               .N3frFdTonsAdj + .N4frFdTonsAdj + _
                               .N5frFdTonsAdj + .SrFdTonsAdj + _
                               .CrFdTonsAdj + .UcrFdTonsAdj), 1)
            Else
                .TotFdBplAdj = 0
            End If

            'Total feed tons adjusted
            .TotFdTonsAdj = .N1frFdTonsAdj + .N2frFdTonsAdj + .N3frFdTonsAdj + _
                .N4frFdTonsAdj + .N5frFdTonsAdj + .SrFdTonsAdj + _
                .CrFdTonsAdj + .UcrFdTonsAdj

            'Method 1 -- sum tails from each circuit
            .TotTlTonsAdjMeth1 = .N1frTlTonsAdj + .N2frTlTonsAdj + _
                                 .N3frTlTonsAdj + .N4frTlTonsAdj + _
                                 .N5frTlTonsAdj + .SrTlTonsAdj + _
                                 .CrTlTonsAdj + .UcrTlTonsAdj + _
                                 .N1faTlTonsAdj + .N2faTlTonsAdj + _
                                 .N3faTlTonsAdj + .CaTlTonsAdj

            'Method2 -- total feed tons - total concentrate product tons
            'Total concentrate product tons -- includes ultra-coarse tons
            .TotTlTonsAdjMeth2 = .TotFdTonsAdj - .PrdCnTonsWuc

            'Total tail BPL, actually measured -- using GMT samples
            'Use method #2 from above
            'Tail tons are based on adjusted feed tons
            If .TotGmtTlTonsW <> 0 Then
                .TotTlBplMsrd = Round(.TotGmtTlBt / .TotGmtTlTonsW, 1)
            Else
                .TotTlBplMsrd = 0
            End If

            'Total adjusted feed BPL-tons
            .TotFdBtAdj = .N1frFdBtAdj + .N2frFdBtAdj + .N3frFdBtAdj + _
                          .N4frFdBtAdj + .N5frFdBtAdj + .SrFdBtAdj + _
                          .CrFdBtAdj + .UcrFdBtAdj

            'Total concentrate BPL-tons
            .PrdCnBt = .PrdCnTonsWuc * .PrdCnBplWuc

            'Back calculate tail BPL (with adjusted feed tons)
            '((feed tons & feed BPL) - (conc tons * conc BPL)) / tail tons
            '09/20/2006, lss -- added Round()
            If .TotTlTonsAdjMeth2 <> 0 Then
                .TotTlBplAdjFd = Round((.TotFdBtAdj - .PrdCnBt) / .TotTlTonsAdjMeth2, 1)
            Else
                .TotTlBplAdjFd = 0
            End If

            'Back calculate tail BPL (with as reported feed tons)
            '((feed tons & feed BPL) - (conc tons * conc BPL)) / tail tons
            'Total as-reported feed BPL
            .TotFdTonsRptW = .N1frFdTonsRptW + .N2frFdTonsRptW + _
                             .N3frFdTonsRptW + .N4frFdTonsRptW + _
                             .N5frFdTonsRptW + .SrFdTonsRptW + _
                             .CrFdTonsRptW + .UcrFdTonsRptW

            .TotFdTonsRptBt = .N1frFdBtRpt + .N2frFdBtRpt + _
                             .N3frFdBtRpt + .N4frFdBtRpt + _
                             .N5frFdBtRpt + .SrFdBtRpt + _
                             .CrFdBtRpt + .UcrFdBtRpt

            If .TotFdTonsRptW <> 0 Then
                .TotFdBplRpt = Round(.TotFdTonsRptBt / .TotFdTonsRptW, 1)
            Else
                .TotFdBplRpt = 0
            End If

            'Total as-reported feed tons
            .TotFdTonsRpt = .N1frFdTonsRpt + .N2frFdTonsRpt + _
                            .N3frFdTonsRpt + .N4frFdTonsRpt + _
                            .N5frFdTonsRpt + .SrFdTonsRpt + _
                            .CrFdTonsRpt + .UcrFdTonsRpt

            'Total as-reported tail tons
            .TotTlTonsRpt = .TotFdTonsRpt - .PrdCnTonsWuc

            '09/20/2006, lss -- added Round()
            If .TotTlTonsRpt <> 0 Then
                .TotTlBplRptFd = Round((.TotFdTonsRptBt - .PrdCnBt) / .TotTlTonsRpt, 1)
            Else
                .TotTlBplRptFd = 0
            End If

            'Calculated gmt bpl -- from weighted average of all circuit tails
            'Total tail tons -- MbTotal.TotTlTonsAdjMeth1
            .TotTlBtAdjMeth1 = Round(.N1frTlTonsAdj * .N1frTlBpl + _
                               .N2frTlTonsAdj * .N2frTlBpl + _
                               .N3frTlTonsAdj * .N3frTlBpl + _
                               .N4frTlTonsAdj * .N4frTlBpl + _
                               .N5frTlTonsAdj * .N5frTlBpl + _
                               .SrTlTonsAdj * .SrTlBpl + _
                               .CrTlTonsAdj * .CrTlBpl + _
                               .N1faTlTonsAdj * .N1faTlBpl + _
                               .N1faTlTonsAdj * .N1faTlBpl + _
                               .N1faTlTonsAdj * .N1faTlBpl + _
                               .N1faTlTonsAdj * .N1faTlBpl, 1)

            If .TotTlTonsAdjMeth1 <> 0 Then
                .TotTlBplFromCircs = Round(.TotTlBtAdjMeth1 / .TotTlTonsAdjMeth1, 1)
            Else
                .TotTlBplFromCircs = 0
            End If

            'RC and Rcvry -- Based on reported GMT BPL
            'MbTotal.TotFdBplAdj
            'MbTotal.PrdCnBplWuc
            'mbTotal.TotTlBplMsrd

            If .TotFdBplAdj - .TotTlBplMsrd <> 0 Then
                .TotTlBplMsrdRc = Round((.PrdCnBplWuc - .TotTlBplMsrd) / _
                                  (.TotFdBplAdj - .TotTlBplMsrd), 2)
            Else
                .TotTlBplMsrdRc = 0
            End If

            .TotFdTonsMsrd = Round(.TotTlBplMsrdRc * .PrdCnTonsWuc, 0)

            If .TotFdBplAdj * .TotFdTonsAdj <> 0 Then
                .TotTlBPlMsrdRcvry = Round((.PrdCnBplWuc * .PrdCnTonsWuc) / _
                                      (.TotFdBplAdj * .TotFdTonsMsrd) * 100, 1)
            Else
                .TotTlBPlMsrdRcvry = 0
            End If

            'RC and Rcvry -- Based on calculated GMT BPL
            'MbTotal.TotFdTonsAdj
            'MbTotal.PrdCnTonsWuc
            'MbTotal.TotFdBplAdj
            'MbTotal.PrdCnBplWuc
            'MbTotal.TotTlBplAdjFd
            If .TotFdBplAdj - .TotTlBplAdjFd <> 0 Then
                .TotTlBplAdjFdRc = Round((.PrdCnBplWuc - .TotTlBplAdjFd) / _
                                   (.TotFdBplAdj - .TotTlBplAdjFd), 2)
            Else
                .TotTlBplAdjFdRc = 0
            End If

            If .TotFdBplAdj * .TotFdTonsAdj <> 0 Then
                .TotTlBplAdjFdRcvry = Round((.PrdCnBplWuc * .PrdCnTonsWuc) / _
                                      (.TotFdBplAdj * .TotFdTonsAdj) * 100, 1)
            Else
                .TotTlBplAdjFdRcvry = 0
            End If

            'RC and Rcvry -- Based on reported feed tons
            'MbTotal.TotFdTonsRpt
            'MbTotal.PrdCnTonsWuc
            'MbTotal.TotFdBplRpt
            'MbTotal.PrdCnBplWuc
            'MbTotal.TotTlBplRptFd
            If .TotFdBplRpt - .TotTlBplRptFd <> 0 Then
                .TotTlBplRptFdRc = Round((.PrdCnBplWuc - .TotTlBplRptFd) / _
                                   (.TotFdBplRpt - .TotTlBplRptFd), 2)
            Else
                .TotTlBplRptFdRc = 0
            End If

            If .TotFdBplRpt * .TotFdTonsRpt <> 0 Then
                .TotTlBplRptFdRcvry = Round((.PrdCnBplWuc * .PrdCnTonsWuc) / _
                                      (.TotFdBplRpt * .TotFdTonsRpt) * 100, 1)
            Else
                .TotTlBplRptFdRcvry = 0
            End If

            'Total fine roughers -- #1FR, #2FR, #3FR, #4FR, #5FR
            'Total fine roughers -- #1FR, #2FR, #3FR, #4FR, #5FR
            'Total fine roughers -- #1FR, #2FR, #3FR, #4FR, #5FR
            'Feeds
            .TotFineRghrFdTons = .N1frFdTonsAdj + .N2frFdTonsAdj + _
                                 .N3frFdTonsAdj + .N4frFdTonsAdj + _
                                 .N5frFdTonsAdj

            .TotFineRghrFdBt = .N1frFdTonsAdj * .N1frFdBpl + _
                               .N2frFdTonsAdj * .N2frFdBpl + _
                               .N3frFdTonsAdj * .N3frFdBpl + _
                               .N4frFdTonsAdj * .N4frFdBpl + _
                               .N5frFdTonsAdj * .N5frFdBpl

            .TotFineRghrFdTonsW = IIf(.N1frFdBpl <> 0, .N1frFdTonsAdj, 0) + _
                                  IIf(.N2frFdBpl <> 0, .N2frFdTonsAdj, 0) + _
                                  IIf(.N3frFdBpl <> 0, .N3frFdTonsAdj, 0) + _
                                  IIf(.N4frFdBpl <> 0, .N4frFdTonsAdj, 0) + _
                                  IIf(.N5frFdBpl <> 0, .N5frFdTonsAdj, 0)

            'Concentrates
            .TotFineRghrCnTons = .N1frCnTonsAdj + .N2FrCnTonsAdj + _
                                 .N3FrCnTonsAdj + .N4FrCnTonsAdj + _
                                 .N5FrCnTonsAdj

            .TotFineRghrCnBt = .N1frCnTonsAdj * .N1frCnBpl + _
                               .N2FrCnTonsAdj * .N2frCnBpl + _
                               .N3FrCnTonsAdj * .N3frCnBpl + _
                               .N4FrCnTonsAdj * .N4frCnBpl + _
                               .N5FrCnTonsAdj * .N5frCnBpl

            .TotFineRghrCnTonsW = IIf(.N1frCnBpl <> 0, .N1frCnTonsAdj, 0) + _
                                  IIf(.N2frCnBpl <> 0, .N2FrCnTonsAdj, 0) + _
                                  IIf(.N3frCnBpl <> 0, .N3FrCnTonsAdj, 0) + _
                                  IIf(.N4frCnBpl <> 0, .N4FrCnTonsAdj, 0) + _
                                  IIf(.N5frCnBpl <> 0, .N5FrCnTonsAdj, 0)

            'Tails
            .TotFineRghrTlTons = .N1frTlTonsAdj + .N2frTlTonsAdj + _
                                 .N3frTlTonsAdj + .N4frTlTonsAdj + _
                                 .N5frTlTonsAdj

            .TotFineRghrTlBt = .N1frTlTonsAdj * .N1frTlBpl + _
                               .N2frTlTonsAdj * .N2frTlBpl + _
                               .N3frTlTonsAdj * .N3frTlBpl + _
                               .N4frTlTonsAdj * .N4frTlBpl + _
                               .N5frTlTonsAdj * .N5frTlBpl

            .TotFineRghrTlTonsW = IIf(.N1frTlBpl <> 0, .N1frTlTonsAdj, 0) + _
                                  IIf(.N2frTlBpl <> 0, .N2frTlTonsAdj, 0) + _
                                  IIf(.N3frTlBpl <> 0, .N3frTlTonsAdj, 0) + _
                                  IIf(.N4frTlBpl <> 0, .N4frTlTonsAdj, 0) + _
                                  IIf(.N5frTlBpl <> 0, .N5frTlTonsAdj, 0)

            If .TotFineRghrFdTonsW <> 0 Then
                .TotFineRghrFdBpl = Round(.TotFineRghrFdBt / .TotFineRghrFdTonsW, 1)
            Else
                .TotFineRghrFdBpl = 0
            End If

            If .TotFineRghrCnTonsW <> 0 Then
                .TotFineRghrCnBpl = Round(.TotFineRghrCnBt / .TotFineRghrCnTonsW, 1)
            Else
                .TotFineRghrCnBpl = 0
            End If

            If .TotFineRghrTlTonsW <> 0 Then
                .TotFineRghrTlBpl = Round(.TotFineRghrTlBt / .TotFineRghrTlTonsW, 1)
            Else
                .TotFineRghrTlBpl = 0
            End If

            If .TotFineRghrCnTons <> 0 Then
                .TotFineRghrRc = Round(.TotFineRghrFdTons / .TotFineRghrCnTons, 2)
            Else
                .TotFineRghrRc = 0
            End If

            If .TotFineRghrFdBt <> 0 Then
                .TotFineRghrRcvry = Round(.TotFineRghrCnBt / .TotFineRghrFdBt * 100, 1)
            Else
                .TotFineRghrRcvry = 0
            End If

            'Total coarse roughers -- SR, CR
            'Total coarse roughers -- SR, CR
            'Total coarse roughers -- SR, CR
            'Feeds
            .TotCrsRghrFdTons = .SrFdTonsAdj + .CrFdTonsAdj

            .TotCrsRghrFdBt = .SrFdTonsAdj * .SrFdBpl + .CrFdTonsAdj * .CrFdBpl

            .TotCrsRghrFdTonsW = IIf(.SrFdBpl <> 0, .SrFdTonsAdj, 0) + _
                              IIf(.CrFdBpl <> 0, .CrFdTonsAdj, 0)

            'Concentrates
            .TotCrsRghrCnTons = .SrCnTonsAdj + .CrCnTonsAdj

            .TotCrsRghrCnBt = .SrCnTonsAdj * .SrCnBpl + .CrCnTonsAdj * .CrCnBpl

            .TotCrsRghrCnTonsW = IIf(.SrCnBpl <> 0, .SrCnTonsAdj, 0) + _
                              IIf(.CrCnBpl <> 0, .CrCnTonsAdj, 0)

            'Tails
            .TotCrsRghrTlTons = .SrTlTonsAdj + .CrTlTonsAdj

            .TotCrsRghrTlBt = .SrTlTonsAdj * .SrTlBpl + .CrTlTonsAdj * .CrTlBpl

            .TotCrsRghrTlTonsW = IIf(.SrTlBpl <> 0, .SrTlTonsAdj, 0) + _
                              IIf(.CrTlBpl <> 0, .CrTlTonsAdj, 0)

            If .TotCrsRghrFdTonsW <> 0 Then
                .TotCrsRghrFdBpl = Round(.TotCrsRghrFdBt / .TotCrsRghrFdTonsW, 1)
            Else
                .TotCrsRghrFdBpl = 0
            End If

            If .TotCrsRghrCnTonsW <> 0 Then
                .TotCrsRghrCnBpl = Round(.TotCrsRghrCnBt / .TotCrsRghrCnTonsW, 1)
            Else
                .TotCrsRghrCnBpl = 0
            End If

            If .TotCrsRghrTlTonsW <> 0 Then
                .TotCrsRghrTlBpl = Round(.TotCrsRghrTlBt / .TotCrsRghrTlTonsW, 1)
            Else
                .TotCrsRghrTlBpl = 0
            End If

            If .TotCrsRghrCnTons <> 0 Then
                .TotCrsRghrRc = Round(.TotCrsRghrFdTons / .TotCrsRghrCnTons, 2)
            Else
                .TotCrsRghrRc = 0
            End If

            If .TotCrsRghrFdBt <> 0 Then
                .TotCrsRghrRcvry = Round(.TotCrsRghrCnBt / .TotCrsRghrFdBt * 100, 1)
            Else
                .TotCrsRghrRcvry = 0
            End If

            'Total roughers -- #1FR, #2FR, #3FR, #4FR, #5FR, SR, CR, UCR
            'Total roughers -- #1FR, #2FR, #3FR, #4FR, #5FR, SR, CR, UCR
            'Total roughers -- #1FR, #2FR, #3FR, #4FR, #5FR, SR, CR, UCR

            'Feeds
            .TotRghrFdTons = .N1frFdTonsAdj + .N2frFdTonsAdj + _
                             .N3frFdTonsAdj + .N4frFdTonsAdj + _
                             .N5frFdTonsAdj + .SrFdTonsAdj + _
                             .CrFdTonsAdj + .UcrFdTonsAdj

            .TotRghrFdTonsRpt = .N1frFdTonsRpt + .N2frFdTonsRpt + _
                                .N3frFdTonsRpt + .N4frFdTonsRpt + _
                                .N5frFdTonsRpt + .SrFdTonsRpt + _
                                .CrFdTonsRpt + .UcrFdTonsRpt

            .TotRghrFdBt = .N1frFdTonsAdj * .N1frFdBpl + _
                           .N2frFdTonsAdj * .N2frFdBpl + _
                           .N3frFdTonsAdj * .N3frFdBpl + _
                           .N4frFdTonsAdj * .N4frFdBpl + _
                           .N5frFdTonsAdj * .N5frFdBpl + _
                           .SrFdTonsAdj * .SrFdBpl + _
                           .CrFdTonsAdj * .CrFdBpl + _
                           .UcrFdTonsAdj * .UcrFdBpl

            .TotRghrFdTonsW = IIf(.N1frFdBpl <> 0, .N1frFdTonsAdj, 0) + _
                                  IIf(.N2frFdBpl <> 0, .N2frFdTonsAdj, 0) + _
                                  IIf(.N3frFdBpl <> 0, .N3frFdTonsAdj, 0) + _
                                  IIf(.N4frFdBpl <> 0, .N4frFdTonsAdj, 0) + _
                                  IIf(.N5frFdBpl <> 0, .N5frFdTonsAdj, 0) + _
                                  IIf(.SrFdBpl <> 0, .SrFdTonsAdj, 0) + _
                                  IIf(.CrFdBpl <> 0, .CrFdTonsAdj, 0) + _
                                  IIf(.UcrFdBpl <> 0, .UcrFdTonsAdj, 0)

            'Concentrates
            .TotRghrCnTons = .N1frCnTonsAdj + .N2FrCnTonsAdj + _
                             .N3FrCnTonsAdj + .N4FrCnTonsAdj + _
                             .N5FrCnTonsAdj + .SrCnTonsAdj + _
                             .CrCnTonsAdj + .UcrCnTonsAdj

            .TotRghrCnBt = .N1frCnTonsAdj * .N1frCnBpl + _
                           .N2FrCnTonsAdj * .N2frCnBpl + _
                           .N3FrCnTonsAdj * .N3frCnBpl + _
                           .N4FrCnTonsAdj * .N4frCnBpl + _
                           .N5FrCnTonsAdj * .N5frCnBpl + _
                           .SrCnTonsAdj * .SrCnBpl + _
                           .CrCnTonsAdj * .CrCnBpl + _
                           .UcrCnTonsAdj * .UcrCnBpl

            .TotRghrCnTonsW = IIf(.N1frCnBpl <> 0, .N1frCnTonsAdj, 0) + _
                                  IIf(.N2frCnBpl <> 0, .N2FrCnTonsAdj, 0) + _
                                  IIf(.N3frCnBpl <> 0, .N3FrCnTonsAdj, 0) + _
                                  IIf(.N4frCnBpl <> 0, .N4FrCnTonsAdj, 0) + _
                                  IIf(.N5frCnBpl <> 0, .N5FrCnTonsAdj, 0) + _
                                  IIf(.SrCnBpl <> 0, .SrCnTonsAdj, 0) + _
                                  IIf(.CrCnBpl <> 0, .CrCnTonsAdj, 0) + _
                                  IIf(.UcrCnBpl <> 0, .UcrCnTonsAdj, 0)

            'Tails
            .TotRghrTlTons = .N1frTlTonsAdj + .N2frTlTonsAdj + _
                             .N3frTlTonsAdj + .N4frTlTonsAdj + _
                             .N5frTlTonsAdj + .SrTlTonsAdj + _
                             .CrTlTonsAdj + .UcrTlTonsAdj

            .TotRghrTlBt = .N1frTlTonsAdj * .N1frTlBpl + _
                           .N2frTlTonsAdj * .N2frTlBpl + _
                           .N3frTlTonsAdj * .N3frTlBpl + _
                           .N4frTlTonsAdj * .N4frTlBpl + _
                           .N5frTlTonsAdj * .N5frTlBpl + _
                           .SrTlTonsAdj * .SrTlBpl + _
                           .CrTlTonsAdj * .CrTlBpl + _
                           .UcrTlTonsAdj * .UcrTlBpl

            .TotRghrTlTonsW = IIf(.N1frTlBpl <> 0, .N1frTlTonsAdj, 0) + _
                                  IIf(.N2frTlBpl <> 0, .N2frTlTonsAdj, 0) + _
                                  IIf(.N3frTlBpl <> 0, .N3frTlTonsAdj, 0) + _
                                  IIf(.N4frTlBpl <> 0, .N4frTlTonsAdj, 0) + _
                                  IIf(.N5frTlBpl <> 0, .N5frTlTonsAdj, 0) + _
                                  IIf(.SrTlBpl <> 0, .SrTlTonsAdj, 0) + _
                                  IIf(.CrTlBpl <> 0, .CrTlTonsAdj, 0) + _
                                  IIf(.UcrTlBpl <> 0, .UcrTlTonsAdj, 0)

            If .TotRghrFdTonsW <> 0 Then
                .TotRghrFdBpl = Round(.TotRghrFdBt / .TotRghrFdTonsW, 1)
            Else
                .TotRghrFdBpl = 0
            End If

            If .TotRghrCnTonsW <> 0 Then
                .TotRghrCnBpl = Round(.TotRghrCnBt / .TotRghrCnTonsW, 1)
            Else
                .TotRghrCnBpl = 0
            End If

            If .TotRghrTlTonsW <> 0 Then
                .TotRghrTlBpl = Round(.TotRghrTlBt / .TotRghrTlTonsW, 1)
            Else
                .TotRghrTlBpl = 0
            End If

            If .TotRghrCnTons <> 0 Then
                .TotRghrRc = Round(.TotRghrFdTons / .TotRghrCnTons, 2)
            Else
                .TotRghrRc = 0
            End If

            If .TotRghrFdBt <> 0 Then
                .TotRghrRcvry = Round(.TotRghrCnBt / .TotRghrFdBt * 100, 1)
            Else
                .TotRghrRcvry = 0
            End If

            'Total roughers #2 -- #1FR, #2FR, #3FR, #4FR, #5FR, SR, CR
            'Total roughers #2 -- #1FR, #2FR, #3FR, #4FR, #5FR, SR, CR
            'Total roughers #2 -- #1FR, #2FR, #3FR, #4FR, #5FR, SR, CR

            'Feeds
            .TotRghr2FdTons = .N1frFdTonsAdj + .N2frFdTonsAdj + _
                             .N3frFdTonsAdj + .N4frFdTonsAdj + _
                             .N5frFdTonsAdj + .SrFdTonsAdj + _
                             .CrFdTonsAdj

            .TotRghr2FdBt = .N1frFdTonsAdj * .N1frFdBpl + _
                           .N2frFdTonsAdj * .N2frFdBpl + _
                           .N3frFdTonsAdj * .N3frFdBpl + _
                           .N4frFdTonsAdj * .N4frFdBpl + _
                           .N5frFdTonsAdj * .N5frFdBpl + _
                           .SrFdTonsAdj * .SrFdBpl + _
                           .CrFdTonsAdj * .CrFdBpl

            .TotRghr2FdTonsW = IIf(.N1frFdBpl <> 0, .N1frFdTonsAdj, 0) + _
                                  IIf(.N2frFdBpl <> 0, .N2frFdTonsAdj, 0) + _
                                  IIf(.N3frFdBpl <> 0, .N3frFdTonsAdj, 0) + _
                                  IIf(.N4frFdBpl <> 0, .N4frFdTonsAdj, 0) + _
                                  IIf(.N5frFdBpl <> 0, .N5frFdTonsAdj, 0) + _
                                  IIf(.SrFdBpl <> 0, .SrFdTonsAdj, 0) + _
                                  IIf(.CrFdBpl <> 0, .CrFdTonsAdj, 0)
            'Concentrates
            .TotRghr2CnTons = .N1frCnTonsAdj + .N2FrCnTonsAdj + _
                             .N3FrCnTonsAdj + .N4FrCnTonsAdj + _
                             .N5FrCnTonsAdj + .SrCnTonsAdj + _
                             .CrCnTonsAdj

            .TotRghr2CnBt = .N1frCnTonsAdj * .N1frCnBpl + _
                           .N2FrCnTonsAdj * .N2frCnBpl + _
                           .N3FrCnTonsAdj * .N3frCnBpl + _
                           .N4FrCnTonsAdj * .N4frCnBpl + _
                           .N5FrCnTonsAdj * .N5frCnBpl + _
                           .SrCnTonsAdj * .SrCnBpl + _
                           .CrCnTonsAdj * .CrCnBpl

            .TotRghr2CnTonsW = IIf(.N1frCnBpl <> 0, .N1frCnTonsAdj, 0) + _
                                  IIf(.N2frCnBpl <> 0, .N2FrCnTonsAdj, 0) + _
                                  IIf(.N3frCnBpl <> 0, .N3FrCnTonsAdj, 0) + _
                                  IIf(.N4frCnBpl <> 0, .N4FrCnTonsAdj, 0) + _
                                  IIf(.N5frCnBpl <> 0, .N5FrCnTonsAdj, 0) + _
                                  IIf(.SrCnBpl <> 0, .SrCnTonsAdj, 0) + _
                                  IIf(.CrCnBpl <> 0, .CrCnTonsAdj, 0)

            'Tails
            .TotRghr2TlTons = .N1frTlTonsAdj + .N2frTlTonsAdj + _
                             .N3frTlTonsAdj + .N4frTlTonsAdj + _
                             .N5frTlTonsAdj + .SrTlTonsAdj + _
                             .CrTlTonsAdj

            .TotRghr2TlBt = .N1frTlTonsAdj * .N1frTlBpl + _
                           .N2frTlTonsAdj * .N2frTlBpl + _
                           .N3frTlTonsAdj * .N3frTlBpl + _
                           .N4frTlTonsAdj * .N4frTlBpl + _
                           .N5frTlTonsAdj * .N5frTlBpl + _
                           .SrTlTonsAdj * .SrTlBpl + _
                           .CrTlTonsAdj * .CrTlBpl

            .TotRghr2TlTonsW = IIf(.N1frTlBpl <> 0, .N1frTlTonsAdj, 0) + _
                                  IIf(.N2frTlBpl <> 0, .N2frTlTonsAdj, 0) + _
                                  IIf(.N3frTlBpl <> 0, .N3frTlTonsAdj, 0) + _
                                  IIf(.N4frTlBpl <> 0, .N4frTlTonsAdj, 0) + _
                                  IIf(.N5frTlBpl <> 0, .N5frTlTonsAdj, 0) + _
                                  IIf(.SrTlBpl <> 0, .SrTlTonsAdj, 0) + _
                                  IIf(.CrTlBpl <> 0, .CrTlTonsAdj, 0)

            If .TotRghr2FdTonsW <> 0 Then
                .TotRghr2FdBpl = Round(.TotRghr2FdBt / .TotRghr2FdTonsW, 1)
            Else
                .TotRghr2FdBpl = 0
            End If

            If .TotRghr2CnTonsW <> 0 Then
                .TotRghr2CnBpl = Round(.TotRghr2CnBt / .TotRghr2CnTonsW, 1)
            Else
                .TotRghr2CnBpl = 0
            End If

            If .TotRghr2TlTonsW <> 0 Then
                .TotRghr2TlBpl = Round(.TotRghr2TlBt / .TotRghr2TlTonsW, 1)
            Else
                .TotRghr2TlBpl = 0
            End If

            If .TotRghr2CnTons <> 0 Then
                .TotRghr2Rc = Round(.TotRghr2FdTons / .TotRghr2CnTons, 2)
            Else
                .TotRghr2Rc = 0
            End If

            If .TotRghr2FdBt <> 0 Then
                .TotRghr2Rcvry = Round(.TotRghr2CnBt / .TotRghr2FdBt * 100, 1)
            Else
                .TotRghr2Rcvry = 0
            End If

            'Total cleaners -- #1FA, #2FA, #3FA, CA
            'Total cleaners -- #1FA, #2FA, #3FA, CA
            'Total cleaners -- #1FA, #2FA, #3FA, CA
            'Feeds
            .TotClnrFdTons = .N1faFdTonsAdj + .N2faFdTonsAdj + _
                             .N3faFdTonsAdj + .CaFdTonsAdj

            .TotClnrFdTonsRpt = .N1faFdTonsRpt + .N2faFdTonsRpt + _
                                .N3faFdTonsRpt + .CaFdTonsRpt

            .TotClnrFdBt = .N1faFdTonsAdj * .N1faFdBpl + _
                           .N2faFdTonsAdj * .N2faFdBpl + _
                           .N3faFdTonsAdj * .N3faFdBpl + _
                           .CaFdTonsAdj * .CaFdBpl

            .TotClnrFdTonsW = IIf(.N1faFdBpl <> 0, .N1faFdTonsAdj, 0) + _
                                  IIf(.N2faFdBpl <> 0, .N2faFdTonsAdj, 0) + _
                                  IIf(.N3faFdBpl <> 0, .N3faFdTonsAdj, 0) + _
                                  IIf(.CaFdBpl <> 0, .CaFdTonsAdj, 0)

            'Concentrates
            .TotClnrCnTons = .N1faCnTons + .N2faCnTons + _
                             .N3faCnTons + .CaCnTons

            .TotClnrCnBt = .N1faCnTons * .N1faCnBpl + _
                           .N2faCnTons * .N2faCnBpl + _
                           .N3faCnTons * .N3faCnBpl + _
                           .CaCnTons * .CaCnBpl

            .TotClnrCnTonsW = IIf(.N1faCnBpl <> 0, .N1faCnTons, 0) + _
                                  IIf(.N2faCnBpl <> 0, .N2faCnTons, 0) + _
                                  IIf(.N3faCnBpl <> 0, .N3faCnTons, 0) + _
                                  IIf(.CaCnBpl <> 0, .CaCnTons, 0)

            'Tails
            .TotClnrTlTons = .N1faTlTonsAdj + .N2faTlTonsAdj + _
                             .N3faTlTonsAdj + .CaTlTonsAdj

            .TotClnrTlBt = .N1faTlTonsAdj * .N1faTlBpl + _
                           .N2faTlTonsAdj * .N2faTlBpl + _
                           .N3faTlTonsAdj * .N3faTlBpl + _
                           .CaTlTonsAdj * .CaTlBpl

            .TotClnrTlTonsW = IIf(.N1faTlBpl <> 0, .N1faTlTonsAdj, 0) + _
                                  IIf(.N2faTlBpl <> 0, .N2faTlTonsAdj, 0) + _
                                  IIf(.N3faTlBpl <> 0, .N3faTlTonsAdj, 0) + _
                                  IIf(.CaTlBpl <> 0, .CaTlTonsAdj, 0)

            If .TotClnrFdTonsW <> 0 Then
                .TotClnrFdBpl = Round(.TotClnrFdBt / .TotClnrFdTonsW, 1)
            Else
                .TotClnrFdBpl = 0
            End If

            If .TotClnrCnTonsW <> 0 Then
                .TotClnrCnBpl = Round(.TotClnrCnBt / .TotClnrCnTonsW, 1)
            Else
                .TotClnrCnBpl = 0
            End If

            If .TotClnrTlTonsW <> 0 Then
                .TotClnrTlBpl = Round(.TotClnrTlBt / .TotClnrTlTonsW, 1)
            Else
                .TotClnrTlBpl = 0
            End If

            If .TotClnrCnTons <> 0 Then
                .TotClnrRc = Round(.TotClnrFdTons / .TotClnrCnTons, 2)
            Else
                .TotClnrRc = 0
            End If

            If .TotClnrFdBt <> 0 Then
                .TotClnrRcvry = Round(.TotClnrCnBt / .TotClnrFdBt * 100, 1)
            Else
                .TotClnrRcvry = 0
            End If

            'Total cleaners -- #1FA, #2FA, #3FA, CA
            'Total cleaners -- #1FA, #2FA, #3FA, CA
            'Total cleaners -- #1FA, #2FA, #3FA, CA
            'Feeds
            .TotClnrFdTons = .N1faFdTonsAdj + .N2faFdTonsAdj + _
                             .N3faFdTonsAdj + .CaFdTonsAdj

            .TotClnrFdBt = .N1faFdTonsAdj * .N1faFdBpl + _
                           .N2faFdTonsAdj * .N2faFdBpl + _
                           .N3faFdTonsAdj * .N3faFdBpl + _
                           .CaFdTonsAdj * .CaFdBpl

            .TotClnrFdTonsW = IIf(.N1faFdBpl <> 0, .N1faFdTonsAdj, 0) + _
                                  IIf(.N2faFdBpl <> 0, .N2faFdTonsAdj, 0) + _
                                  IIf(.N3faFdBpl <> 0, .N3faFdTonsAdj, 0) + _
                                  IIf(.CaFdBpl <> 0, .CaFdTonsAdj, 0)

            'Concentrates
            .TotClnrCnTons = .N1faCnTons + .N2faCnTons + _
                             .N3faCnTons + .CaCnTons

            .TotClnrCnBt = .N1faCnTons * .N1faCnBpl + _
                           .N2faCnTons * .N2faCnBpl + _
                           .N3faCnTons * .N3faCnBpl + _
                           .CaCnTons * .CaCnBpl

            .TotClnrCnTonsW = IIf(.N1faCnBpl <> 0, .N1faCnTons, 0) + _
                                  IIf(.N2faCnBpl <> 0, .N2faCnTons, 0) + _
                                  IIf(.N3faCnBpl <> 0, .N3faCnTons, 0) + _
                                  IIf(.CaCnBpl <> 0, .CaCnTons, 0)

            'Tails
            .TotClnrTlTons = .N1faTlTonsAdj + .N2faTlTonsAdj + _
                             .N3faTlTonsAdj + .CaTlTonsAdj

            .TotClnrTlBt = .N1faTlTonsAdj * .N1faTlBpl + _
                           .N2faTlTonsAdj * .N2faTlBpl + _
                           .N3faTlTonsAdj * .N3faTlBpl + _
                           .CaTlTonsAdj * .CaTlBpl

            .TotClnrTlTonsW = IIf(.N1faTlBpl <> 0, .N1faTlTonsAdj, 0) + _
                                  IIf(.N2faTlBpl <> 0, .N2faTlTonsAdj, 0) + _
                                  IIf(.N3faTlBpl <> 0, .N3faTlTonsAdj, 0) + _
                                  IIf(.CaTlBpl <> 0, .CaTlTonsAdj, 0)

            If .TotClnrFdTonsW <> 0 Then
                .TotClnrFdBpl = Round(.TotClnrFdBt / .TotClnrFdTonsW, 1)
            Else
                .TotClnrFdBpl = 0
            End If

            If .TotClnrCnTonsW <> 0 Then
                .TotClnrCnBpl = Round(.TotClnrCnBt / .TotClnrCnTonsW, 1)
            Else
                .TotClnrCnBpl = 0
            End If

            If .TotClnrTlTonsW <> 0 Then
                .TotClnrTlBpl = Round(.TotClnrTlBt / .TotClnrTlTonsW, 1)
            Else
                .TotClnrTlBpl = 0
            End If

            If .TotClnrCnTons <> 0 Then
                .TotClnrRc = Round(.TotClnrFdTons / .TotClnrCnTons, 2)
            Else
                .TotClnrRc = 0
            End If

            If .TotClnrFdBt <> 0 Then
                .TotClnrRcvry = Round(.TotClnrCnBt / .TotClnrFdBt * 100, 1)
            Else
                .TotClnrRcvry = 0
            End If

            'Total fine cleaners -- #1FA, #2FA, #3FA
            'Total fine cleaners -- #1FA, #2FA, #3FA
            'Total fine cleaners -- #1FA, #2FA, #3FA
            'Feeds
            .TotFineClnrFdTons = .N1faFdTonsAdj + .N2faFdTonsAdj + _
                             .N3faFdTonsAdj

            .TotFineClnrFdBt = .N1faFdTonsAdj * .N1faFdBpl + _
                           .N2faFdTonsAdj * .N2faFdBpl + _
                           .N3faFdTonsAdj * .N3faFdBpl

            .TotFineClnrFdTonsW = IIf(.N1faFdBpl <> 0, .N1faFdTonsAdj, 0) + _
                                  IIf(.N2faFdBpl <> 0, .N2faFdTonsAdj, 0) + _
                                  IIf(.N3faFdBpl <> 0, .N3faFdTonsAdj, 0)

            'Concentrates
            .TotFineClnrCnTons = .N1faCnTons + .N2faCnTons + _
                             .N3faCnTons

            .TotFineClnrCnBt = .N1faCnTons * .N1faCnBpl + _
                           .N2faCnTons * .N2faCnBpl + _
                           .N3faCnTons * .N3faCnBpl

            .TotFineClnrCnTonsW = IIf(.N1faCnBpl <> 0, .N1faCnTons, 0) + _
                                  IIf(.N2faCnBpl <> 0, .N2faCnTons, 0) + _
                                  IIf(.N3faCnBpl <> 0, .N3faCnTons, 0)

            'Tails
            .TotFineClnrTlTons = .N1faTlTonsAdj + .N2faTlTonsAdj + _
                             .N3faTlTonsAdj

            .TotFineClnrTlBt = .N1faTlTonsAdj * .N1faTlBpl + _
                           .N2faTlTonsAdj * .N2faTlBpl + _
                           .N3faTlTonsAdj * .N3faTlBpl

            .TotFineClnrTlTonsW = IIf(.N1faTlBpl <> 0, .N1faTlTonsAdj, 0) + _
                                  IIf(.N2faTlBpl <> 0, .N2faTlTonsAdj, 0) + _
                                  IIf(.N3faTlBpl <> 0, .N3faTlTonsAdj, 0)

            If .TotFineClnrFdTonsW <> 0 Then
                .TotFineClnrFdBpl = Round(.TotFineClnrFdBt / .TotFineClnrFdTonsW, 1)
            Else
                .TotFineClnrFdBpl = 0
            End If

            If .TotFineClnrCnTonsW <> 0 Then
                .TotFineClnrCnBpl = Round(.TotFineClnrCnBt / .TotFineClnrCnTonsW, 1)
            Else
                .TotFineClnrCnBpl = 0
            End If

            If .TotFineClnrTlTonsW <> 0 Then
                .TotFineClnrTlBpl = Round(.TotFineClnrTlBt / .TotFineClnrTlTonsW, 1)
            Else
                .TotFineClnrTlBpl = 0
            End If

            If .TotFineClnrCnTons <> 0 Then
                .TotFineClnrRc = Round(.TotFineClnrFdTons / .TotFineClnrCnTons, 2)
            Else
                .TotFineClnrRc = 0
            End If

            If .TotFineClnrFdBt <> 0 Then
                .TotFineClnrRcvry = Round(.TotFineClnrCnBt / .TotFineClnrFdBt * 100, 1)
            Else
                .TotFineClnrRcvry = 0
            End If

            'Total plant
            'Total plant
            'Total plant
            'Feed
            .TotPlantFdTons = .TotRghr2FdTons
            .TotPlantFdBpl = .TotRghr2FdBpl

            'Concentrate
            .TotPlantCnTons = .TotClnrCnTons
            .TotPlantCnBpl = .TotClnrCnBpl

            'Tails
            .TotPlantTlTons = .TotRghr2TlTons + .TotClnrTlTons
            If .TotRghr2TlTons * .TotRghr2TlBpl <> 0 Then
                .TotPlantTlBpl = Round((.TotRghr2TlTons * .TotRghr2TlBpl + _
                                 .TotClnrTlTons * .TotClnrTlBpl) / _
                                 (.TotRghr2TlTons * .TotRghr2TlBpl), 1)
            Else
                .TotPlantTlBpl = 0
            End If

            If .TotPlantCnTons <> 0 Then
                .TotPlantRc = Round(.TotPlantFdTons / .TotPlantCnTons, 2)
            Else
                .TotPlantRc = 0
            End If

            If .TotPlantFdBpl * .TotPlantFdTons <> 0 Then
                .TotPlantRcvry = Round(((.TotPlantCnBpl * .TotPlantCnTons) / _
                                 (.TotPlantFdBpl * .TotPlantFdTons)) * 100, 1)
            Else
                .TotPlantRcvry = 0
            End If

            'Total combined
            'Total combined
            'Total combined
            'Feed
            .TotCombFdTons = .TotFdTonsAdj
            .TotCombFdBpl = .TotFdBplAdj

            'Concentrate
            .TotCombCnTons = .PrdCnTonsWuc
            .TotCombCnBpl = .PrdCnBplWuc

            'Tails
            .TotCombTlTons = .TotTlTonsAdjMeth2
            .TotCombTlBpl = .TotTlBplAdjFd

            .TotCombRc = .TotTlBplAdjFdRc
            .TotCombRcvry = .TotTlBplAdjFdRcvry
        End With

        Exit Sub

ProcessSfMassBalanceTotalsError:

        MsgBox("Error in processing South Fort Meade mass balance totals." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "South Fort Meade Mass Balance Totals Computation Error")
    End Sub

    Private Sub DetermineMissingCcnTons()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade

        'Added 01/13/2003, lss
        'Some or all of the coarse concentrate may have been sent to
        'the fine concentrate bins -- thus we don't have a measure of
        'what the coarse concentrate production for the shift was.
        'The coarse concentrate tons in the coarse concentrate bins
        'is in MbSfShift.PrdCcnTons.
        'Need to determine MbSfShift.CcnTonCorr and MbSfShift.FcnTonCorr.

        Dim CrRc As Single
        Dim CaRc As Single
        Dim CrCnTonsExp As Double
        Dim CcnTonsExp As Double

        MbSfShift.CcnTonCorr = 0
        MbSfShift.FcnTonCorr = 0

        With MbSfShift
            'Coarse rougher ratio of concentration
            If (.CrFdBpl - .CrTlBpl) <> 0 Then
                CrRc = Round((.CrCnBpl - .CrTlBpl) / (.CrFdBpl - .CrTlBpl), 2)
            Else
                CrRc = 0
            End If

            'CR rougher concentrate tons expected -- feed to the coarse amine circuit
            If CrRc <> 0 Then
                CrCnTonsExp = Round(.CrFdTonsRpt / CrRc, 0)
            Else
                CrCnTonsExp = 0
            End If

            'These coarse rougher tons will now be sent to the coarse amine
            'circuit -- how many coarse concentrate tons will be expected?

            'Have coarse amine feed and coarse amine tail BPL's -- however
            'may not have a coarse concentrate BPL.

            'If the transfer factor was not 100 then some coarse concentrate
            'was sent to the coarse concentrate bins and should have a
            'legitimate coarse concentrate BPL associated with it.
            If .PrdCcnBpl <> 0 And Transfer.CAtoFC <> 100 Then
                If .CaFdBpl - .CaTlBpl <> 0 Then
                    CaRc = Round((.PrdCcnBpl - .CaTlBpl) / (.CaFdBpl - .CaTlBpl), 2)
                Else
                    CaRc = 0
                End If
            Else
                'Don't have a legitimate coarse concentrate BPL -- will have to use the
                'fine concentrate BPL for the shift instead.
                'If Transfer.CAtoFC = 100 and .PrdCcnBpl <> 0 then will still use
                '.PrdFcnBpl instead.

                .PrdCcnBpl = .PrdFcnBpl

                If .CaFdBpl - .CaTlBpl <> 0 Then
                    CaRc = Round((.PrdFcnBpl - .CaTlBpl) / (.CaFdBpl - .CaTlBpl), 2)
                Else
                    CaRc = 0
                End If
            End If

            'Expected coarse concentrate product tons
            If CaRc <> 0 Then
                CcnTonsExp = Round(CrCnTonsExp / CaRc, 0)
            Else
                CcnTonsExp = 0
            End If

            .CcnTonCorr = gRoundFifty((Transfer.CAtoFC / 100) * CcnTonsExp)
            .FcnTonCorr = -1 * .CcnTonCorr
        End With
    End Sub

    Private Sub ZeroSfSummingData()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade

        With MbSfTotal
            .PrdFcnTons = 0
            .PrdFcnTonsW = 0
            .PrdFcnBt = 0
            .PrdFcnBpl = 0

            .PrdCcnTons = 0
            .PrdCcnTonsW = 0
            .PrdCcnBt = 0
            .PrdCcnBpl = 0

            .PrdUccnTons = 0
            .PrdUccnTonsW = 0
            .PrdUccnBt = 0
            .PrdUccnBpl = 0

            .PrdCnTonsWuc = 0
            .PrdCnBplWuc = 0
            .PrdCnTonsWouc = 0
            .PrdCnBplWouc = 0

            .PrdCnBt = 0

            .N1frHrs = 0
            .N2frHrs = 0
            .N3frHrs = 0
            .N4frHrs = 0
            .N5frHrs = 0
            .SrHrs = 0
            .CrHrs = 0
            .UcrHrs = 0

            .N1frFdTonsRpt = 0
            .N2frFdTonsRpt = 0
            .N3frFdTonsRpt = 0
            .N4frFdTonsRpt = 0
            .N5frFdTonsRpt = 0
            .SrFdTonsRpt = 0
            .CrFdTonsRpt = 0
            .UcrFdTonsRpt = 0

            .N1frFdTonsRptW = 0
            .N2frFdTonsRptW = 0
            .N3frFdTonsRptW = 0
            .N4frFdTonsRptW = 0
            .N5frFdTonsRptW = 0
            .SrFdTonsRptW = 0
            .CrFdTonsRptW = 0
            .UcrFdTonsRptW = 0

            .N1frFdBtRpt = 0
            .N2frFdBtRpt = 0
            .N3frFdBtRpt = 0
            .N4frFdBtRpt = 0
            .N5frFdBtRpt = 0
            .SrFdBtRpt = 0
            .CrFdBtRpt = 0
            .UcrFdBtRpt = 0

            .N1faFdTonsRpt = 0
            .N2faFdTonsRpt = 0
            .N3faFdTonsRpt = 0
            .CaFdTonsRpt = 0

            .N1faFdTonsAdj = 0
            .N1faFdBtAdj = 0
            .N2faFdTonsAdj = 0
            .N2faFdBtAdj = 0
            .N3faFdTonsAdj = 0
            .N3faFdBtAdj = 0
            .CaFdTonsAdj = 0
            .CaFdBtAdj = 0

            .N1faCnBtAdj = 0     '10/31/2011, lss  New
            .N2faCnBtAdj = 0     '10/31/2011, lss  New
            .N3faCnBtAdj = 0     '10/31/2011, lss  New

            .N1frCnTonsAdj = 0
            .N2FrCnTonsAdj = 0
            .N3FrCnTonsAdj = 0
            .N4FrCnTonsAdj = 0
            .N5FrCnTonsAdj = 0
            .CrCnTonsAdj = 0
            .SrCnTonsAdj = 0
            .UcrCnTonsAdj = 0

            .N1frCnBtAdj = 0
            .N2FrCnBtAdj = 0
            .N3FrCnBtAdj = 0
            .N4FrCnBtAdj = 0
            .N5FrCnBtAdj = 0
            .CrCnBtAdj = 0
            .SrCnBtAdj = 0
            .UcrCnBtAdj = 0

            .N1faTlTonsAdj = 0
            .N2faTlTonsAdj = 0
            .N3faTlTonsAdj = 0
            .CaTlTonsAdj = 0

            .N1faTlBtAdj = 0
            .N2faTlBtAdj = 0
            .N3faTlBtAdj = 0
            .CaTlBtAdj = 0

            .TotFaFdTonsAdj = 0

            .N1frFdTonsAdj = 0
            .N2frFdTonsAdj = 0
            .N3frFdTonsAdj = 0
            .N4frFdTonsAdj = 0
            .N5frFdTonsAdj = 0
            .CrFdTonsAdj = 0
            .SrFdTonsAdj = 0
            .UcrFdTonsAdj = 0

            .N1frFdBtAdj = 0
            .N2frFdBtAdj = 0
            .N3frFdBtAdj = 0
            .N4frFdBtAdj = 0
            .N5frFdBtAdj = 0
            .CrFdBtAdj = 0
            .SrFdBtAdj = 0
            .UcrFdBtAdj = 0

            .N1frTlTonsAdj = 0
            .N2frTlTonsAdj = 0
            .N3frTlTonsAdj = 0
            .N4frTlTonsAdj = 0
            .N5frTlTonsAdj = 0
            .SrTlTonsAdj = 0
            .CrTlTonsAdj = 0
            .UcrTlTonsAdj = 0

            .N1frTlBtAdj = 0
            .N2frTlBtAdj = 0
            .N3frTlBtAdj = 0
            .N4frTlBtAdj = 0
            .N5frTlBtAdj = 0
            .SrTlBtAdj = 0
            .CrTlBtAdj = 0
            .UcrTlBtAdj = 0

            .TotGmtTlTonsW = 0
            .TotGmtTlBt = 0

            .N1frFdBpl = 0
            .N2frFdBpl = 0
            .N3frFdBpl = 0
            .N4frFdBpl = 0
            .N5frFdBpl = 0
            .SrFdBpl = 0
            .CrFdBpl = 0
            .UcrFdBpl = 0

            .N1frArFdBpl = 0
            .N2frArFdBpl = 0
            .N3frArFdBpl = 0
            .N4frArFdBpl = 0
            .N5frArFdBpl = 0
            .SrArFdBpl = 0
            .CrArFdBpl = 0
            .UcrArFdBpl = 0

            .N1frCnBpl = 0
            .N2frCnBpl = 0
            .N3frCnBpl = 0
            .N4frCnBpl = 0
            .N5frCnBpl = 0
            .SrCnBpl = 0
            .CrCnBpl = 0
            .UcrCnBpl = 0

            .N1frTlBpl = 0
            .N2frTlBpl = 0
            .N3frTlBpl = 0
            .N4frTlBpl = 0
            .N5frTlBpl = 0
            .SrTlBpl = 0
            .CrTlBpl = 0
            .UcrTlBpl = 0

            .N1faFdBpl = 0
            .N2faFdBpl = 0
            .N3faFdBpl = 0
            .CaFdBpl = 0

            .N1faCnBpl = 0
            .N2faCnBpl = 0
            .N3faCnBpl = 0
            .CaCnBpl = 0

            .N1faTlBpl = 0
            .N2faTlBpl = 0
            .N3faTlBpl = 0
            .CaTlBpl = 0

            .N1frRc = 0
            .N2frRc = 0
            .N3frRc = 0
            .N4frRc = 0
            .N5frRc = 0
            .SrRc = 0
            .CrRc = 0
            .UcrRc = 0
            .N1faRc = 0
            .N2faRc = 0
            .N3faRc = 0
            .CaRc = 0

            .N1frFdTph = 0
            .N1frCnTph = 0
            .N1frTlTph = 0
            .N1frActPctRcvry = 0
            .N1frStdRc = 0
            .N1frStdCnTons = 0
            .N1frStdPctRcvry = 0

            .N2frFdTph = 0
            .N2frCnTph = 0
            .N2frTlTph = 0
            .N2frActPctRcvry = 0
            .N2frStdRc = 0
            .N2frStdCnTons = 0
            .N2frStdPctRcvry = 0

            .N3frFdTph = 0
            .N3frCnTph = 0
            .N3frTlTph = 0
            .N3frActPctRcvry = 0
            .N3frStdRc = 0
            .N3frStdCnTons = 0
            .N3frStdPctRcvry = 0

            .N4frFdTph = 0
            .N4frCnTph = 0
            .N4frTlTph = 0
            .N4frActPctRcvry = 0
            .N4frStdRc = 0
            .N4frStdCnTons = 0
            .N4frStdPctRcvry = 0

            .N5frFdTph = 0
            .N5frCnTph = 0
            .N5frTlTph = 0
            .N5frActPctRcvry = 0
            .N5frStdRc = 0
            .N5frStdCnTons = 0
            .N5frStdPctRcvry = 0

            .CrFdTph = 0
            .CrCnTph = 0
            .CrTlTph = 0
            .CrActPctRcvry = 0
            .CrStdRc = 0
            .CrStdCnTons = 0
            .CrStdPctRcvry = 0

            .SrFdTph = 0
            .SrCnTph = 0
            .SrTlTph = 0
            .SrActPctRcvry = 0
            .SrStdRc = 0
            .SrStdCnTons = 0
            .SrStdPctRcvry = 0

            .UcrFdTph = 0
            .UcrCnTph = 0
            .UcrTlTph = 0
            .UcrActPctRcvry = 0
            .UcrStdRc = 0
            .UcrStdCnTons = 0
            .UcrStdPctRcvry = 0

            .N1faCnTons = 0
            .N2faCnTons = 0
            .N3faCnTons = 0
            .CaCnTons = 0

            .N1faActPctRcvry = 0
            .N1faStdRc = 0
            .N1faStdCnTons = 0
            .N1faStdPctRcvry = 0

            .N2faActPctRcvry = 0
            .N2faStdRc = 0
            .N2faStdCnTons = 0
            .N2faStdPctRcvry = 0

            .N3faActPctRcvry = 0
            .N3faStdRc = 0
            .N3faStdCnTons = 0
            .N3faStdPctRcvry = 0

            .CaActPctRcvry = 0
            .CaStdRc = 0
            .CaStdCnTons = 0
            .CaStdPctRcvry = 0

            .TotFdBplAdj = 0
            .TotFdTonsAdj = 0
            .TotTlBtAdjMeth1 = 0
            .TotTlTonsAdjMeth1 = 0
            .TotTlTonsAdjMeth2 = 0
            .TotTlBplMsrd = 0
            .TotFdTonsMsrd = 0
            .TotFdBtAdj = 0
            .TotTlBplAdjFd = 0
            .TotTlBplRptFd = 0
            .TotFdTonsRpt = 0
            .TotFdTonsRptW = 0
            .TotFdTonsRptBt = 0
            .TotTlTonsRpt = 0
            .TotFdBplRpt = 0
            .TotTlBplFromCircs = 0

            .TotTlBplMsrdRc = 0
            .TotTlBPlMsrdRcvry = 0

            .TotTlBplAdjFdRc = 0
            .TotTlBplAdjFdRcvry = 0

            .TotTlBplRptFdRc = 0
            .TotTlBplRptFdRcvry = 0

            .TotFineRghrFdTons = 0
            .TotFineRghrFdBt = 0
            .TotFineRghrFdTonsW = 0
            .TotFineRghrFdBpl = 0
            .TotFineRghrCnTons = 0
            .TotFineRghrCnBt = 0
            .TotFineRghrCnTonsW = 0
            .TotFineRghrCnBpl = 0
            .TotFineRghrTlTons = 0
            .TotFineRghrTlBt = 0
            .TotFineRghrTlTonsW = 0
            .TotFineRghrTlBpl = 0
            .TotFineRghrRc = 0
            .TotFineRghrRcvry = 0

            .TotCrsRghrFdTons = 0
            .TotCrsRghrFdBt = 0
            .TotCrsRghrFdTonsW = 0
            .TotCrsRghrFdBpl = 0
            .TotCrsRghrCnTons = 0
            .TotCrsRghrCnBt = 0
            .TotCrsRghrCnTonsW = 0
            .TotCrsRghrCnBpl = 0
            .TotCrsRghrTlTons = 0
            .TotCrsRghrTlBt = 0
            .TotCrsRghrTlTonsW = 0
            .TotCrsRghrTlBpl = 0
            .TotCrsRghrRc = 0
            .TotCrsRghrRcvry = 0

            .TotRghrFdTons = 0
            .TotRghrFdTonsRpt = 0
            .TotRghrFdBt = 0
            .TotRghrFdTonsW = 0
            .TotRghrFdBpl = 0
            .TotRghrCnTons = 0
            .TotRghrCnBt = 0
            .TotRghrCnTonsW = 0
            .TotRghrCnBpl = 0
            .TotRghrTlTons = 0
            .TotRghrTlBt = 0
            .TotRghrTlTonsW = 0
            .TotRghrTlBpl = 0
            .TotRghrRc = 0
            .TotRghrRcvry = 0

            .TotRghr2FdTons = 0
            .TotRghr2FdBt = 0
            .TotRghr2FdTonsW = 0
            .TotRghr2FdBpl = 0
            .TotRghr2CnTons = 0
            .TotRghr2CnBt = 0
            .TotRghr2CnTonsW = 0
            .TotRghr2CnBpl = 0
            .TotRghr2TlTons = 0
            .TotRghr2TlBt = 0
            .TotRghr2TlTonsW = 0
            .TotRghr2TlBpl = 0
            .TotRghr2Rc = 0
            .TotRghr2Rcvry = 0

            .TotClnrFdTons = 0
            .TotClnrFdTonsRpt = 0
            .TotClnrFdBt = 0
            .TotClnrFdTonsW = 0
            .TotClnrFdBpl = 0
            .TotClnrCnTons = 0
            .TotClnrCnBt = 0
            .TotClnrCnTonsW = 0
            .TotClnrCnBpl = 0
            .TotClnrTlTons = 0
            .TotClnrTlBt = 0
            .TotClnrTlTonsW = 0
            .TotClnrTlBpl = 0
            .TotClnrRc = 0
            .TotClnrRcvry = 0

            .TotFineClnrFdTons = 0
            .TotFineClnrFdBt = 0
            .TotFineClnrFdTonsW = 0
            .TotFineClnrFdBpl = 0
            .TotFineClnrCnTons = 0
            .TotFineClnrCnBt = 0
            .TotFineClnrCnTonsW = 0
            .TotFineClnrCnBpl = 0
            .TotFineClnrTlTons = 0
            .TotFineClnrTlBt = 0
            .TotFineClnrTlTonsW = 0
            .TotFineClnrTlBpl = 0
            .TotFineClnrRc = 0
            .TotFineClnrRcvry = 0

            .TotPlantFdTons = 0
            .TotPlantFdTonsRpt = 0
            .TotPlantCnTons = 0
            .TotPlantTlTons = 0
            .TotPlantFdBpl = 0
            .TotPlantCnBpl = 0
            .TotPlantTlBpl = 0
            .TotPlantRc = 0
            .TotPlantRcvry = 0

            .TotCombFdTons = 0
            .TotCombFdTonsRpt = 0
            .TotCombCnTons = 0
            .TotCombTlTons = 0
            .TotCombFdBpl = 0
            .TotCombCnBpl = 0
            .TotCombTlBpl = 0
            .TotCombRc = 0
            .TotCombRcvry = 0
        End With
    End Sub

    Public Function gAdjustedFeedTonsSF(ByVal aBeginDate As Date, _
                                        ByVal aBeginShift As String, _
                                        ByVal aEndDate As Date, _
                                        ByVal aEndShift As String, _
                                        ByVal aCrewNumber As String, _
                                        ByRef rFeedBpl As Double, _
                                        ByRef rFeedTonsRpt As Long) As Long

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade

        'This function will return the total adjusted feed tons for
        'any given time frame.

        On Error GoTo gAdjustedFeedTonsSFError

        Dim RowIdx As Integer
        Dim NumShifts As Integer

        Dim FloatPlantCirc(20, 15) As Object
        Dim FloatPlantGmt(5, 15) As Object

        'Get data for float plant mass balance

        NumShifts = gGetSfFloatPlantBalanceData(FloatPlantCirc, _
                                                FloatPlantGmt, _
                                                aBeginDate, _
                                                StrConv(aBeginShift, vbUpperCase), _
                                                aEndDate, _
                                                StrConv(aEndShift, vbUpperCase), _
                                                aCrewNumber, _
                                                1)

        gAdjustedFeedTonsSF = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalRghr, mSfFloatPlantCircColEnum.sfCcFdTonsAdj)
        rFeedBpl = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalRghr, mSfFloatPlantCircColEnum.sfCcFdBpl)
        rFeedTonsRpt = FloatPlantCirc(mSfFloatPlantCircRowEnum.sfCrTotalRghr, mSfFloatPlantCircColEnum.sfCcFdTonsRpt)

        Exit Function

gAdjustedFeedTonsSFError:

        MsgBox("Error summing South Fort Meade adjusted feed tons." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "South Fort Meade Adjusted Feed Tons Error")

    End Function

    Public Function gGetMetReagentDataSf(ByVal aBeginDate As Date, _
                                         ByVal aBeginShift As String, _
                                         ByVal aEndDate As Date, _
                                         ByVal aEndShift As String, _
                                         ByVal aCrewNumber As String, _
                                         ByVal TotAdjFdTons As Long, _
                                         ByVal TotRptFdTons As Long, _
                                         ByVal TotCnTons As Long) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade
        'South Fort Meade  South Fort Meade  South Fort Meade  South Fort Meade

        On Error GoTo gGetMetReagentDataErrorSf

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        Dim MetReagentDynaset As OraDynaset
        Dim RecCount As Long

        Dim ThisMatl As String

        Dim TotCost As Long
        Dim TotUnits As Long

        TotCost = 0
        TotUnits = 0

        'Get reagent data from EQPT_CALC

        params = gDBParams

        params.Add("pMineName", "South Fort Meade", ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pBeginDate", aBeginDate, ORAPARM_INPUT)
        params("pBeginDate").serverType = ORATYPE_DATE

        params.Add("pBeginShift", StrConv(aBeginShift, vbUpperCase), ORAPARM_INPUT)
        params("pBeginShift").serverType = ORATYPE_VARCHAR2

        params.Add("pEndDate", aEndDate, ORAPARM_INPUT)
        params("pEndDate").serverType = ORATYPE_DATE

        params.Add("pEndShift", StrConv(aEndShift, vbUpperCase), ORAPARM_INPUT)
        params("pEndShift").serverType = ORATYPE_VARCHAR2

        params.Add("pCrewNumber", aCrewNumber, ORAPARM_INPUT)
        params("pCrewNumber").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_reagents.get_reagent_usage(:pMineName," + _
                      ":pBeginDate, :pBeginShift, :pEndDate, :pEndShift, :pCrewNumber, :pResult);end;", ORASQL_FAILEXEC)
        MetReagentDynaset = params("pResult").Value
        RecCount = MetReagentDynaset.RecordCount

        If RecCount = 0 Then
            gGetMetReagentDataSf = False
            ClearParams(params)
            Exit Function
        End If

        gGetMetReagentDataSf = True
        MetallurgicalSfRpt.RgAllTotCost = 0
        MetallurgicalSfRpt.RgAllTotUnits = 0

        MetReagentDynaset.MoveFirst()

        Do While Not MetReagentDynaset.EOF
            ThisMatl = MetReagentDynaset.Fields("matl_name").Value
            With MetallurgicalSfRpt
                Select Case ThisMatl
                    Case Is = "Amine"
                        .RgAmTotUnits = MetReagentDynaset.Fields("pound_usage").Value
                        .RgAmTotCost = MetReagentDynaset.Fields("cost").Value
                        .RgAllTotUnits = .RgAllTotUnits + .RgAmTotUnits
                        .RgAllTotCost = .RgAllTotCost + .RgAmTotCost

                    Case Is = "Depressant"
                        .RgDeTotUnits = MetReagentDynaset.Fields("pound_usage").Value
                        .RgDeTotCost = MetReagentDynaset.Fields("cost").Value
                        .RgAllTotUnits = .RgAllTotUnits + .RgDeTotUnits
                        .RgAllTotCost = .RgAllTotCost + .RgDeTotCost

                    Case Is = "Fatty acid"
                        .RgFaTotUnits = MetReagentDynaset.Fields("pound_usage").Value
                        .RgFaTotCost = MetReagentDynaset.Fields("cost").Value
                        .RgAllTotUnits = .RgAllTotUnits + .RgFaTotUnits
                        .RgAllTotCost = .RgAllTotCost + .RgFaTotCost

                    Case Is = "Fuel oil"
                        .RgFoTotUnits = MetReagentDynaset.Fields("pound_usage").Value
                        .RgFoTotCost = MetReagentDynaset.Fields("cost").Value
                        .RgAllTotUnits = .RgAllTotUnits + .RgFoTotUnits
                        .RgAllTotCost = .RgAllTotCost + .RgFoTotCost

                    Case Is = "Soda ash"
                        .RgSoTotUnits = MetReagentDynaset.Fields("pound_usage").Value
                        .RgSoTotCost = MetReagentDynaset.Fields("cost").Value
                        .RgAllTotUnits = .RgAllTotUnits + .RgSoTotUnits
                        .RgAllTotCost = .RgAllTotCost + .RgSoTotCost

                    Case Is = "Sulfuric acid"
                        .RgSaTotUnits = MetReagentDynaset.Fields("pound_usage").Value
                        .RgSaTotCost = MetReagentDynaset.Fields("cost").Value
                        .RgAllTotUnits = .RgAllTotUnits + .RgSaTotUnits
                        .RgAllTotCost = .RgAllTotCost + .RgSaTotCost

                    Case Is = "Surfactant"
                        .RgSuTotUnits = MetReagentDynaset.Fields("pound_usage").Value
                        .RgSuTotCost = MetReagentDynaset.Fields("cost").Value
                        .RgAllTotUnits = .RgAllTotUnits + .RgSuTotUnits
                        .RgAllTotCost = .RgAllTotCost + .RgSuTotCost

                End Select
            End With

            MetReagentDynaset.MoveNext()
        Loop

        ClearParams(params)

        With MetallurgicalSfRpt
            'Adjusted feed ton calculations
            'Adjusted feed ton calculations
            'Adjusted feed ton calculations

            If TotAdjFdTons <> 0 Then
                .RgAmAdjFdDpt = Round(.RgAmTotCost / TotAdjFdTons, 4)
            Else
                .RgAmAdjFdDpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgDeAdjFdDpt = Round(.RgDeTotCost / TotAdjFdTons, 4)
            Else
                .RgDeAdjFdDpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgFaAdjFdDpt = Round(.RgFaTotCost / TotAdjFdTons, 4)
            Else
                .RgFaAdjFdDpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgFoAdjFdDpt = Round(.RgFoTotCost / TotAdjFdTons, 4)
            Else
                .RgFoAdjFdDpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgSoAdjFdDpt = Round(.RgSoTotCost / TotAdjFdTons, 4)
            Else
                .RgSoAdjFdDpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgSaAdjFdDpt = Round(.RgSaTotCost / TotAdjFdTons, 4)
            Else
                .RgSaAdjFdDpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgSuAdjFdDpt = Round(.RgSuTotCost / TotAdjFdTons, 4)
            Else
                .RgSuAdjFdDpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgAllAdjFdDpt = Round(.RgAllTotCost / TotAdjFdTons, 4)
            Else
                .RgAllAdjFdDpt = 0
            End If

            '--------------------

            If TotAdjFdTons <> 0 Then
                .RgAmAdjFdUpt = Round(.RgAmTotUnits / TotAdjFdTons, 4)
            Else
                .RgAmAdjFdUpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgDeAdjFdUpt = Round(.RgDeTotUnits / TotAdjFdTons, 4)
            Else
                .RgDeAdjFdUpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgFaAdjFdUpt = Round(.RgFaTotUnits / TotAdjFdTons, 4)
            Else
                .RgFaAdjFdUpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgFoAdjFdUpt = Round(.RgFoTotUnits / TotAdjFdTons, 4)
            Else
                .RgFoAdjFdUpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgSoAdjFdUpt = Round(.RgSoTotUnits / TotAdjFdTons, 4)
            Else
                .RgSoAdjFdUpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgSaAdjFdUpt = Round(.RgSaTotUnits / TotAdjFdTons, 4)
            Else
                .RgSaAdjFdUpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgSuAdjFdUpt = Round(.RgSuTotUnits / TotAdjFdTons, 4)
            Else
                .RgSuAdjFdUpt = 0
            End If

            If TotAdjFdTons <> 0 Then
                .RgAllAdjFdUpt = Round(.RgAllTotUnits / TotAdjFdTons, 4)
            Else
                .RgAllAdjFdUpt = 0
            End If

            'Reported feed ton calculations
            'Reported feed ton calculations
            'Reported feed ton calculations

            If TotRptFdTons <> 0 Then
                .RgAmRptFdDpt = Round(.RgAmTotCost / TotRptFdTons, 4)
            Else
                .RgAmRptFdDpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgDeRptFdDpt = Round(.RgDeTotCost / TotRptFdTons, 4)
            Else
                .RgDeRptFdDpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgFaRptFdDpt = Round(.RgFaTotCost / TotRptFdTons, 4)
            Else
                .RgFaRptFdDpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgFoRptFdDpt = Round(.RgFoTotCost / TotRptFdTons, 4)
            Else
                .RgFoRptFdDpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgSoRptFdDpt = Round(.RgSoTotCost / TotRptFdTons, 4)
            Else
                .RgSoRptFdDpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgSaRptFdDpt = Round(.RgSaTotCost / TotRptFdTons, 4)
            Else
                .RgSaRptFdDpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgSuRptFdDpt = Round(.RgSuTotCost / TotRptFdTons, 4)
            Else
                .RgSuRptFdDpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgAllRptFdDpt = Round(.RgAllTotCost / TotRptFdTons, 4)
            Else
                .RgAllRptFdDpt = 0
            End If

            '--------------------

            If TotRptFdTons <> 0 Then
                .RgAmRptFdUpt = Round(.RgAmTotUnits / TotRptFdTons, 4)
            Else
                .RgAmRptFdUpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgDeRptFdUpt = Round(.RgDeTotUnits / TotRptFdTons, 4)
            Else
                .RgDeRptFdUpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgFaRptFdUpt = Round(.RgFaTotUnits / TotRptFdTons, 4)
            Else
                .RgFaRptFdUpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgFoRptFdUpt = Round(.RgFoTotUnits / TotRptFdTons, 4)
            Else
                .RgFoRptFdUpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgSoRptFdUpt = Round(.RgSoTotUnits / TotRptFdTons, 4)
            Else
                .RgSoRptFdUpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgSaRptFdUpt = Round(.RgSaTotUnits / TotRptFdTons, 4)
            Else
                .RgSaRptFdUpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgSuRptFdUpt = Round(.RgSuTotUnits / TotRptFdTons, 4)
            Else
                .RgSuRptFdUpt = 0
            End If

            If TotRptFdTons <> 0 Then
                .RgAllRptFdUpt = Round(.RgAllTotUnits / TotRptFdTons, 4)
            Else
                .RgAllRptFdUpt = 0
            End If

            'Concentrate ton calculations
            'Concentrate ton calculations
            'Concentrate ton calculations

            If TotCnTons <> 0 Then
                .RgAmCnDpt = Round(.RgAmTotCost / TotCnTons, 4)
            Else
                .RgAmCnDpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgDeCnDpt = Round(.RgDeTotCost / TotCnTons, 4)
            Else
                .RgDeCnDpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgFaCnDpt = Round(.RgFaTotCost / TotCnTons, 4)
            Else
                .RgFaCnDpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgFoCnDpt = Round(.RgFoTotCost / TotCnTons, 4)
            Else
                .RgFoCnDpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgSoCnDpt = Round(.RgSoTotCost / TotCnTons, 4)
            Else
                .RgSoCnDpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgSaCnDpt = Round(.RgSaTotCost / TotCnTons, 4)
            Else
                .RgSaCnDpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgSuCnDpt = Round(.RgSuTotCost / TotCnTons, 4)
            Else
                .RgSuCnDpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgAllCnDpt = Round(.RgAllTotCost / TotCnTons, 4)
            Else
                .RgAllCnDpt = 0
            End If

            '--------------------

            If TotCnTons <> 0 Then
                .RgAmCnUpt = Round(.RgAmTotUnits / TotCnTons, 4)
            Else
                .RgAmCnUpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgDeCnUpt = Round(.RgDeTotUnits / TotCnTons, 4)
            Else
                .RgDeCnUpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgFaCnUpt = Round(.RgFaTotUnits / TotCnTons, 4)
            Else
                .RgFaCnUpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgFoCnUpt = Round(.RgFoTotUnits / TotCnTons, 4)
            Else
                .RgFoCnUpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgSoCnUpt = Round(.RgSoTotUnits / TotCnTons, 4)
            Else
                .RgSoCnUpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgSaCnUpt = Round(.RgSaTotUnits / TotCnTons, 4)
            Else
                .RgSaCnUpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgSuCnUpt = Round(.RgSuTotUnits / TotCnTons, 4)
            Else
                .RgSuCnUpt = 0
            End If

            If TotCnTons <> 0 Then
                .RgAllCnUpt = Round(.RgAllTotUnits / TotCnTons, 4)
            Else
                .RgAllCnUpt = 0
            End If

        End With

        Exit Function

gGetMetReagentDataErrorSf:

        MsgBox("Error getting South Fort Meade reagent data." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "South Fort Meade Reagent Data Error")

        On Error Resume Next
        ClearParams(params)
    End Function

    Public Function gFltPltRcvrySF(ByVal aBeginDate As Date, _
                                   ByVal aBeginShift As String, _
                                   ByVal aEndDate As Date, _
                                   ByVal aEndShift As String, _
                                   ByVal aCrewNumber As String, _
                                   ByRef rFeedBpl As Single, _
                                   ByVal aBplRound As Integer) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'This function will return the plant recovery for
        'any given time frame.  It will also return the adjusted feed BPL through
        'rFeedBpl.

        On Error GoTo gFltPltRcvrySfError

        Dim RowIdx As Integer
        Dim NumShifts As Integer

        Dim FloatPlantCirc(20, 15) As Object
        Dim FloatPlantGmt(5, 15) As Object

        'Get data for float plant mass balance
        NumShifts = gGetSfFloatPlantBalanceData(FloatPlantCirc, _
                                                FloatPlantGmt, _
                                                aBeginDate, _
                                                StrConv(aBeginShift, vbUpperCase), _
                                                aEndDate, _
                                                StrConv(aEndShift, vbUpperCase), _
                                                aCrewNumber, _
                                                1)

        gFltPltRcvrySF = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcPctRcvry)

        rFeedBpl = FloatPlantGmt(mSfFloatPlantGmtRowEnum.sfGrCalculatedGmtBpl, mSfFloatPlantGmtColEnum.sfGcFdBpl)

        Exit Function

gFltPltRcvrySfError:

        MsgBox("Error getting South Fort Meade float plant recovery." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "South Fort Meade Plant Recovery Error")
    End Function




End Module
