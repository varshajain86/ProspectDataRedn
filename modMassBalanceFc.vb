Option Explicit On
Imports OracleInProcServer
Imports System.Math
Module modMassBalanceFc
    'Attribute VB_Name = "modMassBalanceFc"

    '**********************************************************************
    'Four Corners Mass Balance Module
    '
    'Special Comments
    '----------------
    'This module handles mass balance and metallurgical reports for
    'Four Corners only -- gMassBalanceFC and gMetallurgicalFC.
    '
    'Types in this module:
    'Private Type MassBalanceFcRptType
    'Dim MassBalanceFcRpt As MassBalanceFcRptType
    '
    'Private Type MetallurgicalFcRptType
    'Dim MetallurgicalFcRpt As MetallurgicalFcRptType
    '
    'Private Type MassBalanceFcShiftType
    'Dim fMbFcShift As MassBalanceFcShiftType
    '
    'Private Type MassBalanceFcTotalType
    'Dim fMbFcTotal As MassBalanceFcTotalType
    '
    'Procedures/Functions in this module:
    '1) gMassBalanceFc
    '2) gMetallurgicalFc
    '3) gGetFcFloatPlantBalanceData
    '4) ZeroFcSummingData
    '5) ProcessFcMassBalanceData
    '6) ProcessFcMassBalanceTotals
    '7) gAdjustedFeedTonsFc
    '8) gGetMetReagentDataFc

    'IMPORTANT -- This module handles the mass balance for Glen Oswald's
    '             original GMT based mass balance.
    '
    'Glen Oswald's original GMT based mass balance!
    'Glen Oswald's original GMT based mass balance!
    'Glen Oswald's original GMT based mass balance!
    '
    '**********************************************************************
    '   Maintenance Log
    '
    '   04/19/2005, lss
    '       Set up this module.
    '   09/07/2005, lss
    '       Set up MassBalance report for Four Corners.  There are issues
    '       to be discussed!
    '   01/24/2006, lss
    '       Completed smoothing out the mass balance for Four Corners.  It
    '       is based on a method that Glen Oswald provided.
    '   03/13/2006, lss
    '       Added Public Function gFltPltRcvryFC.
    '   03/31/2006, lss
    '       Added Function gGetBalanceDistribution2.
    '   08/10/2006, lss
    '       Modified Function gGetBalanceDistribution2 -- uses
    '       gGetPeriodicEqptMsrAvg3 instead of gGetPeriodicEqptMsrAvg.
    '       This change was made for 3 -> 2 shift change.
    '   08/14/2006, lss
    '       Added BalDistFrPct, BalDistCrPct, BalDistFaPct and BalDistCaPct
    '       to mMbFcTotal.  These items will be displayed in the FCO mass
    '       balance report = MassBalanceFc2.rpt
    '   08/22/2006, lss
    '       Form is OK for gFirstShift and gLastShift.
    '   08/23/2006, lss
    '       Removed SSIEBER so that report runs correctly!  Fixed reagent
    '       part of the report!
    '   12/04/2006, lss
    '       Changed TotPltTlTonsAdj and TotPltGmtTlBtRpt from Single to
    '       Double -- was causing an error in the mass balance report
    '       MassBalanceFc2.rpt).
    '   02/25/2008, lss
    '       Added this formula to the report:
    '       frmViewData.rptInputData.Formulas(153) =
    '       "TotPltTlBplRpt = " & .TotPltTlBplRpt & ""
    '   04/09/2008, lss
    '       Added Fatty acid Custaflot 109 stuff.
    '
    '**********************************************************************


    Dim mMassBalanceDynaset As OraDynaset

    'Four Corners  Four Corners  Four Corners  Four Corners  Four Corners
    'Four Corners  Four Corners  Four Corners  Four Corners  Four Corners
    'Four Corners  Four Corners  Four Corners  Four Corners  Four Corners

    Private Enum mFcFloatPlantCircRowEnum
        fcCrFneRghr = 1       'Fine rougher
        fcCrFneAmine = 2      'Fine amine
        fcCrTotFne = 3        'Total fine
        fcCrCrsRghr = 4       'Coarse rougher
        fcCrCrsAmine = 5      'Coarse amine
        fcCrTotCrs = 6        'Total coarse
        fcCrTotAmine = 7      'Total amine
        fcCrTotPlant = 8      'Total plant
        fcCrCnProduct = 9
    End Enum

    Private Enum mFcFloatPlantCircColEnum
        fcCcOperHrs = 1
        fcCcFdTonsRpt = 2
        fcCcFdTonsAdj = 3
        fcCcFdBpl = 4
        fcCcCnBpl = 5
        fcCcTlBpl = 6
        fcCcRc = 7
        fcCcPctActRcvry = 8
        fcCcPctRptRcvry = 9
        fcCcCnTonsAdj = 10
        fcCcTlTonsAdj = 11
        fcCcFdTph = 12
        fcCcCnTph = 13
        fcCcTlTph = 14
    End Enum

    Private Enum mFcFloatPlantGmtRowEnum
        fcGrAsReportedGmtBpl = 1
        fcGrCalculatedGmtBpl = 2
        fcGrReportedFdTons = 3
        fcGrGmtBplFromCircuits = 4
    End Enum

    Private Enum mWgFloatPlantGmtColEnum
        fcGcFdTons = 1
        fcGcCnTons = 2
        fcGcFdBpl = 3
        fcGcCnBpl = 4
        fcGcTlBpl = 5
        fcGcRc = 6
        fcGcPctRcvry = 7
    End Enum

    'Mass Balance  Mass Balance  Mass Balance  Mass Balance
    'Mass Balance  Mass Balance  Mass Balance  Mass Balance
    'Mass Balance  Mass Balance  Mass Balance  Mass Balance

    '12/17/2007, lss
    'Added:
    '1) Nfr1TlBplRpt As Single
    '2) Nfr2TlBplRpt As Single
    '3) Sfr1TlBplRpt As Single
    '4) Sfr2TlBplRpt As Single

    Public Structure mMassBalanceFcShiftType
        Public Nfr1Hrs As Single
        Public Nfr2Hrs As Single
        Public Sfr1Hrs As Single
        Public Sfr2Hrs As Single
        Public NcrHrs As Single
        Public ScrHrs As Single
        Public NcsHrs As Single
        Public ScsHrs As Single
        Public NcrCaHrs As Single
        Public ScrCaHrs As Single
        Public CrHrs As Single    '!!!
        Public CaHrs As Single    '!!!
        Public CaPct As Single    '!!!
        '-----
        Public NfrRc As Single
        Public SfrRc As Single
        Public FaRc As Single
        Public NcrRc As Single
        Public ScrRc As Single
        Public CaRc As Single
        '-----
        Public PrdCnBpl As Single          'Actual concentrate production tons
        Public PrdCnTons As Double         'Actual concentrate production BPL
        '-----
        Public NfrFdBplRpt As Single
        Public SfrFdBplRpt As Single
        Public NcrFdBplRpt As Single
        Public ScrFdBplRpt As Single
        Public CrsColFdBplRpt As Single
        '-----
        Public NfrCnBplRpt As Single
        Public SfrCnBplRpt As Single
        Public CrCnBplRpt As Single
        Public CrsColCnBplRpt As Single
        '-----
        Public CrsColTlBplRpt As Single
        '-----
        Public NfrTlBplRpt As Single
        Public SfrTlBplRpt As Single
        Public FaTlBplRpt As Single
        Public NcrTlBplRpt As Single
        Public ScrTlBplRpt As Single
        Public CaTlBplRpt As Single
        Public GmtBplRpt As Single
        '-----
        Public CaFdBplRpt As Single
        '-----
        Public NfrTlBplAdj As Single
        Public SfrTlBplAdj As Single
        Public FaTlBplAdj As Single
        Public NcrTlBplAdj As Single
        Public ScrTlBplAdj As Single
        Public CaTlBplAdj As Single
        Public GmtBplAdj As Single
        Public FrTlBplAdj As Single
        Public CrTlBplAdj As Single
        '-----
        Public Nfr1FdTonsRpt As Double
        Public Nfr2FdTonsRpt As Double
        Public Sfr1FdTonsRpt As Double
        Public Sfr2FdTonsRpt As Double
        Public NfrFdTonsRpt As Double
        Public SfrFdTonsRpt As Double
        Public NcrFdTonsRpt As Double
        Public ScrFdTonsRpt As Double
        Public FaFdTonsRpt As Double
        Public CaFdTonsRpt As Double
        Public FrFdTonsRpt As Double
        Public CrFdTonsRpt As Double
        '-----
        Public NfrFdTonsAdj As Double
        Public SfrFdTonsAdj As Double
        Public NcrFdTonsAdj As Double
        Public ScrFdTonsAdj As Double
        Public FaFdTonsAdj As Double
        Public CaFdTonsAdj As Double
        Public FrFdTonsAdj As Double
        Public CrFdTonsAdj As Double
        '-----
        Public AvgFrCnBpl As Single
        Public AvgCrCnBpl As Single
        '-----
        Public NfrTlBtRpt As Double
        Public SfrTlBtRpt As Double
        Public FaTlBtRpt As Double
        Public NcrTlBtRpt As Double
        Public ScrTlBtRpt As Double
        Public CaTlBtRpt As Double
        '-----
        Public NfrTlBtAdj As Double
        Public SfrTlBtAdj As Double
        Public FaTlBtAdj As Double
        Public NcrTlBtAdj As Double
        Public ScrTlBtAdj As Double
        Public CaTlBtAdj As Double
        '-----
        Public NfrCnTonsRpt As Double
        Public SfrCnTonsRpt As Double
        Public FaCnTonsRpt As Double
        Public NcrCnTonsRpt As Double
        Public ScrCnTonsRpt As Double
        Public CaCnTonsRpt As Double
        '-----
        Public NfrCnTonsAdj As Double
        Public SfrCnTonsAdj As Double
        Public FaCnTonsAdj As Double
        Public NcrCnTonsAdj As Double
        Public ScrCnTonsAdj As Double
        Public CaCnTonsAdj As Double
        '-----
        Public NfrTlTonsRpt As Double
        Public SfrTlTonsRpt As Double
        Public FaTlTonsRpt As Double
        Public NcrTlTonsRpt As Double
        Public ScrTlTonsRpt As Double
        Public CaTlTonsRpt As Double
        Public FrTlTonsRpt As Double
        Public CrTlTonsRpt As Double
        '-----
        Public NfrTlTonsAdj As Double
        Public SfrTlTonsAdj As Double
        Public FaTlTonsAdj As Double
        Public NcrTlTonsAdj As Double
        Public ScrTlTonsAdj As Double
        Public CaTlTonsAdj As Double
        Public FrTlTonsAdj As Double
        Public CrTlTonsAdj As Double
    End Structure
    Dim mMbFcShift As mMassBalanceFcShiftType

    Private Structure MassBalanceFcTotalType
        Public PrdCnTons As Long
        Public PrdCnTonsW As Long
        Public PrdCnBt As Double
        Public PrdCnBpl As Single
        '----------
        Public Nfr1Hrs As Double
        Public Nfr2Hrs As Double
        Public Sfr1Hrs As Double
        Public Sfr2Hrs As Double
        Public NcrHrs As Double
        Public ScrHrs As Double
        Public NcsHrs As Double
        Public ScsHrs As Double
        Public FrHrs As Double
        Public CrHrs As Double
        Public FaHrs As Double
        Public CaHrs As Double
        Public TotPltHrs As Double
        Public NcrCaHrs As Double
        Public ScrCaHrs As Double
        '----------
        Public FrFdTonsRpt As Long
        Public FrFdBtRpt As Double
        Public FrFdBplRpt As Single
        Public FrCnTonsRpt As Long
        Public FrCnBtRpt As Double
        Public FrCnBplRpt As Single
        Public FrTlTonsRpt As Long
        Public FrTlBtRpt As Double
        Public FrTlBplRpt As Single
        '----------
        Public FaFdTonsRpt As Long
        Public FaFdBtRpt As Double
        Public FaFdBplRpt As Single
        Public FaCnTonsRpt As Long
        Public FaCnBtRpt As Double
        Public FaCnBplRpt As Single
        Public FaTlTonsRpt As Long
        Public FaTlBtRpt As Double
        Public FaTlBplRpt As Single
        '----------
        Public CrFdTonsRpt As Long
        Public CrFdBtRpt As Double
        Public CrFdBplRpt As Single
        Public CrCnTonsRpt As Long
        Public CrCnBtRpt As Double
        Public CrCnBplRpt As Single
        Public CrTlTonsRpt As Long
        Public CrTlBtRpt As Double
        Public CrTlBplRpt As Single
        '----------
        Public CaFdTonsRpt As Long
        Public CaFdBtRpt As Double
        Public CaFdBplRpt As Single
        Public CaCnTonsRpt As Long
        Public CaCnBtRpt As Double
        Public CaCnBplRpt As Single
        Public CaTlTonsRpt As Long
        Public CaTlBtRpt As Double
        Public CaTlBplRpt As Single
        '----------
        Public FrFdTonsAdj As Long
        Public FrFdBtAdj As Double
        Public FrFdBplAdj As Single
        Public FrCnTonsAdj As Long
        Public FrCnBtAdj As Double
        Public FrCnBplAdj As Single
        Public FrTlTonsAdj As Long
        Public FrTlBtAdj As Double
        Public FrTlBplAdj As Single
        '----------
        Public FaFdTonsAdj As Long
        Public FaFdBtAdj As Double
        Public FaFdBplAdj As Single
        Public FaCnTonsAdj As Long
        Public FaCnBtAdj As Double
        Public FaCnBplAdj As Single
        Public FaTlTonsAdj As Long
        Public FaTlBtAdj As Double
        Public FaTlBplAdj As Single
        '----------
        Public CrFdTonsAdj As Long
        Public CrFdBtAdj As Double
        Public CrFdBplAdj As Single
        Public CrCnTonsAdj As Long
        Public CrCnBtAdj As Double
        Public CrCnBplAdj As Single
        Public CrTlTonsAdj As Long
        Public CrTlBtAdj As Double
        Public CrTlBplAdj As Single
        '----------
        Public CaFdTonsAdj As Long
        Public CaFdBtAdj As Double
        Public CaFdBplAdj As Single
        Public CaCnTonsAdj As Long
        Public CaCnBtAdj As Double
        Public CaCnBplAdj As Single
        Public CaTlTonsAdj As Long
        Public CaTlBtAdj As Double
        Public CaTlBplAdj As Single
        '----------
        Public FrRcAdj As Single
        Public FaRcAdj As Single
        Public CrRcAdj As Single
        Public CaRcAdj As Single
        '----------
        Public FrFdTphRpt As Double
        Public FrCnTphRpt As Double
        Public FrTlTphRpt As Double
        Public CrFdTphRpt As Double
        Public CrCnTphRpt As Double
        Public CrTlTphRpt As Double
        Public FaFdTphRpt As Double
        Public FaCnTphRpt As Double
        Public FaTlTphRpt As Double
        Public CaFdTphRpt As Double
        Public CaCnTphRpt As Double
        Public CaTlTphRpt As Double
        '----------
        Public FrFdTphAdj As Double
        Public FrCnTphAdj As Double
        Public FrTlTphAdj As Double
        Public CrFdTphAdj As Double
        Public CrCnTphAdj As Double
        Public CrTlTphAdj As Double
        Public FaFdTphAdj As Double
        Public FaCnTphAdj As Double
        Public FaTlTphAdj As Double
        Public CaFdTphAdj As Double
        Public CaCnTphAdj As Double
        Public CaTlTphAdj As Double
        '----------
        Public TotPltFdBplAdj As Single
        Public TotPltCnBplAdj As Single
        Public TotPltTlBplAdj As Single
        '----------
        Public TotPltTlTonsAdj As Double
        Public TotPltGmtTlBtRpt As Double
        Public TotPltTlBplRpt As Single
        Public TotPltTlBplRpt2 As Single
        Public TotPltFdBplRpt As Single
        '----------
        Public TotPltFdTonsRpt As Double
        Public TotPltFdTonsAdj As Double
        Public TotPltTlTonsRpt As Double
        '----------
        Public FrPctAdjRcvry As Single
        Public FaPctAdjRcvry As Single
        Public CrPctAdjRcvry As Single
        Public CaPctAdjRcvry As Single
        Public TotPltPctAdjRcvry As Single
        '----------
        Public FrPctRptRcvry As Single
        Public FaPctRptRcvry As Single
        Public CrPctRptRcvry As Single
        Public CaPctRptRcvry As Single
        Public TotPltPctRptRcvry As Single
        '----------
        Public TotPltRcAdj As Single
        Public TotPltCnTonsAdj As Double
        Public TotPltFdTphAdj As Single
        Public TotPltCnTphAdj As Single
        Public TotPltTlTphAdj As Single
        Public TotPltCnBplRpt As Single
        '-----------
        Public RgTotRptFdTons As Long
        Public RgTotCnTons As Long
        Public RgTotAdjFdTons As Long

        Public BalDistFrPct As Single
        Public BalDistCrPct As Single
        Public BalDistFaPct As Single
        Public BalDistCaPct As Single
    End Structure
    Dim mMbFcTotal As MassBalanceFcTotalType

    Public Structure gMassBalanceFcReagDataType
        Public RgSuTotUnits As Long
        Public RgAmTotUnits As Long
        Public RgSaTotUnits As Long
        Public RgSoTotUnits As Long
        Public RgFaTotUnits As Long
        Public RgFa2TotUnits As Long
        Public RgFoTotUnits As Long
        Public RgDeTotUnits As Long
        Public RgSiTotUnits As Long
        Public RgAllTotUnits As Long

        Public RgSuTotCost As Long
        Public RgAmTotCost As Long
        Public RgSaTotCost As Long
        Public RgSoTotCost As Long
        Public RgFaTotCost As Long
        Public RgFa2TotCost As Long
        Public RgFoTotCost As Long
        Public RgDeTotCost As Long
        Public RgSiTotCost As Long
        Public RgAllTotCost As Long

        Public RgSuAdjFdDpt As Single
        Public RgAmAdjFdDpt As Single
        Public RgSaAdjFdDpt As Single
        Public RgSoAdjFdDpt As Single
        Public RgFaAdjFdDpt As Single
        Public RgFa2AdjFdDpt As Single
        Public RgFoAdjFdDpt As Single
        Public RgDeAdjFdDpt As Single
        Public RgSiAdjFdDpt As Single
        Public RgAllAdjFdDpt As Single

        Public RgSuRptFdDpt As Single
        Public RgAmRptFdDpt As Single
        Public RgSaRptFdDpt As Single
        Public RgSoRptFdDpt As Single
        Public RgFaRptFdDpt As Single
        Public RgFa2RptFdDpt As Single
        Public RgFoRptFdDpt As Single
        Public RgDeRptFdDpt As Single
        Public RgSiRptFdDpt As Single
        Public RgAllRptFdDpt As Single

        Public RgSuCnDpt As Single
        Public RgAmCnDpt As Single
        Public RgSaCnDpt As Single
        Public RgSoCnDpt As Single
        Public RgFaCnDpt As Single
        Public RgFa2CnDpt As Single
        Public RgFoCnDpt As Single
        Public RgDeCnDpt As Single
        Public RgSiCnDpt As Single
        Public RgAllCnDpt As Single

        Public RgSuAdjFdUpt As Single
        Public RgAmAdjFdUpt As Single
        Public RgSaAdjFdUpt As Single
        Public RgSoAdjFdUpt As Single
        Public RgFaAdjFdUpt As Single
        Public RgFa2AdjFdUpt As Single
        Public RgFoAdjFdUpt As Single
        Public RgDeAdjFdUpt As Single
        Public RgSiAdjFdUpt As Single
        Public RgAllAdjFdUpt As Single

        Public RgSuRptFdUpt As Single
        Public RgAmRptFdUpt As Single
        Public RgSaRptFdUpt As Single
        Public RgSoRptFdUpt As Single
        Public RgFaRptFdUpt As Single
        Public RgFa2RptFdUpt As Single
        Public RgFoRptFdUpt As Single
        Public RgDeRptFdUpt As Single
        Public RgSiRptFdUpt As Single
        Public RgAllRptFdUpt As Single

        Public RgSuCnUpt As Single
        Public RgAmCnUpt As Single
        Public RgSaCnUpt As Single
        Public RgSoCnUpt As Single
        Public RgFaCnUpt As Single
        Public RgFa2CnUpt As Single
        Public RgFoCnUpt As Single
        Public RgDeCnUpt As Single
        Public RgSiCnUpt As Single
        Public RgAllCnUpt As Single
    End Structure
    Dim mMbFcReag As gMassBalanceFcReagDataType

    'As of 01/23/2006 Four Corners does not have separate "Mass Balance"
    'and "Metallurgical" reports.  It has a single combined report =
    'MassBalanceF2.rpt.  This report includes reagent data which is
    'normally included on the "Metallurgical" report for other mines.
    'Metallurgical  Metallurgical  Metallurgical  Metallurgical
    'Metallurgical  Metallurgical  Metallurgical  Metallurgical
    'Metallurgical  Metallurgical  Metallurgical  Metallurgical

    'Private Type mMetallurgicalFcRptType

    'End Type
    'Dim mMetallurgicalFcRpt As mMetallurgicalFcRptType

    'Miscellaneous  Miscellaneous  Miscellaneous  Miscellaneous
    'Miscellaneous  Miscellaneous  Miscellaneous  Miscellaneous
    'Miscellaneous  Miscellaneous  Miscellaneous  Miscellaneous

    Public Structure mBalanceDistributionType
        Public FrPct As Single
        Public CrPct As Single
        Public FaPct As Single
        Public CaPct As Single
    End Structure

    Dim mBeginDate As Date
    Dim mBeginShift As String
    Dim mEndDate As Date
    Dim mEndShift As String

    Public Function gMassBalanceFC(ByVal aBeginDate As Date, _
                                   ByVal aBeginShift As String, _
                                   ByVal aEndDate As Date, _
                                   ByVal aEndShift As String, _
                                   ByVal aCrewNumber As String, _
                                   ByVal aBplRound As Integer, _
                                   ByVal aMassBalMode As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gMassBalanceFcError

        Dim ConnectString As String
        Dim RowIdx As Integer
        Dim TimeFrame As String
        Dim NumShifts As Integer

        Dim FloatPlantCirc() As Object
        Dim FloatPlantGmt() As Object
        Dim GotReagents As Boolean
        Dim ReagentTitle As String

        Dim ReagBeginDate As Date
        Dim ReagBeginShift As String
        Dim ReagEndDate As Date
        Dim ReagEndShift As String

        Dim RerunBalance As Boolean

        '   frmViewData.rptInputData.Reset()
        '
        ZeroFcSummingData()

        'Miscellaneous data setup
        ' frmViewData.rptInputData.Formulas(0) = "MineName = '" & "Four Corners" & "'"

        If aBeginDate = aEndDate And aBeginShift = aEndShift Then
            TimeFrame = aBeginDate & "  " & _
                        StrConv(aBeginShift, vbProperCase) & " Shift"
        Else
            TimeFrame = aBeginDate & " " & _
                        StrConv(aBeginShift, vbProperCase) & _
                        " Shift" & " thru " & _
                        aEndDate & " " & _
                        StrConv(aEndShift, vbProperCase) & " Shift"
        End If
        '  frmViewData.rptInputData.Formulas(1) = "TimeFrame = '" & TimeFrame & "'"
        '  frmViewData.rptInputData.Formulas(2) = "CrewNumber = '" & aCrewNumber & "'"

        'Get data for float plant mass balance
        NumShifts = gGetFcFloatPlantBalanceData(FloatPlantCirc, _
                                                FloatPlantGmt, _
                                                aBeginDate, _
                                                StrConv(aBeginShift, vbUpperCase), _
                                                aEndDate, _
                                                StrConv(aEndShift, vbUpperCase), _
                                                aCrewNumber, _
                                                aBplRound, _
                                                aMassBalMode)
        With mMbFcTotal
            'frmViewData.rptInputData.Formulas(3) = "FrHrs = " & .FrHrs & ""
            'frmViewData.rptInputData.Formulas(4) = "FaHrs = " & .FaHrs & ""
            'frmViewData.rptInputData.Formulas(5) = "CrHrs = " & .CrHrs & ""
            'frmViewData.rptInputData.Formulas(6) = "CaHrs = " & .CaHrs & ""

            'frmViewData.rptInputData.Formulas(7) = "FrFdTonsAdj = " & .FrFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(8) = "FaFdTonsAdj = " & .FaFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(9) = "CrFdTonsAdj = " & .CrFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(10) = "CaFdTonsAdj = " & .CaFdTonsAdj & ""

            'frmViewData.rptInputData.Formulas(11) = "FrFdBplAdj = " & .FrFdBplAdj & ""
            'frmViewData.rptInputData.Formulas(12) = "FrCnBplAdj = " & .FrCnBplAdj & ""
            'frmViewData.rptInputData.Formulas(13) = "FrTlBplAdj = " & .FrTlBplAdj & ""

            'frmViewData.rptInputData.Formulas(14) = "FaFdBplAdj = " & .FaFdBplAdj & ""
            'frmViewData.rptInputData.Formulas(15) = "FaCnBplAdj = " & .FaCnBplAdj & ""
            'frmViewData.rptInputData.Formulas(16) = "FaTlBplAdj = " & .FaTlBplAdj & ""

            'frmViewData.rptInputData.Formulas(17) = "CrFdBplAdj = " & .CrFdBplAdj & ""
            'frmViewData.rptInputData.Formulas(18) = "CrCnBplAdj = " & .CrCnBplAdj & ""
            'frmViewData.rptInputData.Formulas(19) = "CrTlBplAdj = " & .CrTlBplAdj & ""

            'frmViewData.rptInputData.Formulas(20) = "CaFdBplAdj = " & .CaFdBplAdj & ""
            'frmViewData.rptInputData.Formulas(21) = "CaCnBplAdj = " & .CaCnBplAdj & ""
            'frmViewData.rptInputData.Formulas(22) = "CaTlBplAdj = " & .CaTlBplAdj & ""

            'frmViewData.rptInputData.Formulas(23) = "FrCnTonsAdj = " & .FrCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(24) = "FaCnTonsAdj = " & .FaCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(25) = "CrCnTonsAdj = " & .CrCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(26) = "CaCnTonsAdj = " & .CaCnTonsAdj & ""

            'frmViewData.rptInputData.Formulas(27) = "FrTlTonsAdj = " & .FrTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(28) = "FaTlTonsAdj = " & .FaTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(29) = "CrTlTonsAdj = " & .CrTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(30) = "CaTlTonsAdj = " & .CaTlTonsAdj & ""

            'frmViewData.rptInputData.Formulas(31) = "FrFdTphAdj = " & .FrFdTphAdj & ""
            'frmViewData.rptInputData.Formulas(32) = "FrCnTphAdj = " & .FrCnTphAdj & ""
            'frmViewData.rptInputData.Formulas(33) = "FrTlTphAdj = " & .FrTlTphAdj & ""

            'frmViewData.rptInputData.Formulas(34) = "FaFdTphAdj = " & .FaFdTphAdj & ""
            'frmViewData.rptInputData.Formulas(35) = "FaCnTphAdj = " & .FaCnTphAdj & ""
            'frmViewData.rptInputData.Formulas(36) = "FaTlTphAdj = " & .FaTlTphAdj & ""

            'frmViewData.rptInputData.Formulas(37) = "CrFdTphAdj = " & .CrFdTphAdj & ""
            'frmViewData.rptInputData.Formulas(38) = "CrCnTphAdj = " & .CrCnTphAdj & ""
            'frmViewData.rptInputData.Formulas(39) = "CrTlTphAdj = " & .CrTlTphAdj & ""

            'frmViewData.rptInputData.Formulas(40) = "CaFdTphAdj = " & .CaFdTphAdj & ""
            'frmViewData.rptInputData.Formulas(41) = "CaCnTphAdj = " & .CaCnTphAdj & ""
            'frmViewData.rptInputData.Formulas(42) = "CaTlTphAdj = " & .CaTlTphAdj & ""

            'frmViewData.rptInputData.Formulas(43) = "FrRcAdj = " & .FrRcAdj & ""
            'frmViewData.rptInputData.Formulas(44) = "FaRcAdj = " & .FaRcAdj & ""
            'frmViewData.rptInputData.Formulas(45) = "CrRcAdj = " & .CrRcAdj & ""
            'frmViewData.rptInputData.Formulas(46) = "CaRcAdj = " & .CaRcAdj & ""

            'frmViewData.rptInputData.Formulas(47) = "FrAr = " & .FrPctAdjRcvry & ""
            'frmViewData.rptInputData.Formulas(48) = "FaAr = " & .FaPctAdjRcvry & ""
            'frmViewData.rptInputData.Formulas(49) = "CrAr = " & .CrPctAdjRcvry & ""
            'frmViewData.rptInputData.Formulas(50) = "CaAr = " & .CaPctAdjRcvry & ""

            'frmViewData.rptInputData.Formulas(51) = "FrRr = " & .FrPctRptRcvry & ""
            'frmViewData.rptInputData.Formulas(52) = "FaRr = " & .FaPctRptRcvry & ""
            'frmViewData.rptInputData.Formulas(53) = "CrRr = " & .CrPctRptRcvry & ""
            'frmViewData.rptInputData.Formulas(54) = "CaRr = " & .CaPctRptRcvry & ""

            'frmViewData.rptInputData.Formulas(55) = "TotPltHrs = " & .TotPltHrs & ""
            'frmViewData.rptInputData.Formulas(56) = "TotPltFdTonsAdj = " & .TotPltFdTonsAdj & ""
            'frmViewData.rptInputData.Formulas(57) = "TotPltFdBplAdj = " & .TotPltFdBplAdj & ""
            'frmViewData.rptInputData.Formulas(58) = "TotPltCnBplAdj = " & .TotPltCnBplAdj & ""
            'frmViewData.rptInputData.Formulas(59) = "TotPltTlBplAdj = " & .TotPltTlBplAdj & ""
            'frmViewData.rptInputData.Formulas(60) = "TotPltRcAdj = " & .TotPltRcAdj & ""
            'frmViewData.rptInputData.Formulas(61) = "TotPltRr = " & .TotPltPctRptRcvry & ""
            'frmViewData.rptInputData.Formulas(62) = "TotPltAr = " & .TotPltPctAdjRcvry & ""

            'frmViewData.rptInputData.Formulas(63) = "TotPltCnTonsAdj = " & .TotPltCnTonsAdj & ""
            'frmViewData.rptInputData.Formulas(64) = "TotPltTlTonsAdj = " & .TotPltTlTonsAdj & ""
            'frmViewData.rptInputData.Formulas(65) = "TotPltFdTphAdj = " & .TotPltFdTphAdj & ""
            'frmViewData.rptInputData.Formulas(66) = "TotPltCnTphAdj = " & .TotPltCnTphAdj & ""
            'frmViewData.rptInputData.Formulas(67) = "TotPltTlTphAdj = " & .TotPltTlTphAdj & ""

            'frmViewData.rptInputData.Formulas(68) = "FrTlBplRpt = " & .FrTlBplRpt & ""
            'frmViewData.rptInputData.Formulas(69) = "CrTlBplRpt = " & .CrTlBplRpt & ""
            'frmViewData.rptInputData.Formulas(70) = "FaTlBplRpt = " & .FaTlBplRpt & ""
            'frmViewData.rptInputData.Formulas(71) = "CaTlBplRpt = " & .CaTlBplRpt & ""

            'frmViewData.rptInputData.Formulas(72) = "PrdCnTons = " & .PrdCnTons & ""
            'frmViewData.rptInputData.Formulas(73) = "PrdCnBpl = " & Round(.PrdCnBpl, aBplRound) & ""

            ''Extra formulas added 08/15/2006, lss
            'frmViewData.rptInputData.Formulas(149) = "BalDistFrPct = " & .BalDistFrPct & ""
            'frmViewData.rptInputData.Formulas(150) = "BalDistCrPct = " & .BalDistCrPct & ""
            'frmViewData.rptInputData.Formulas(151) = "BalDistFaPct = " & .BalDistFaPct & ""
            'frmViewData.rptInputData.Formulas(152) = "BalDistCaPct = " & .BalDistCaPct & ""

            ''Extra formula added 02/25/2008s
            'frmViewData.rptInputData.Formulas(153) = "TotPltTlBplRpt = " & .TotPltTlBplRpt & ""
        End With

        'Four Corners only has a single "Mass Balance/Metallurgical" report at this
        'time so we need to get reagent data here also.

        'Reagent data is available by day totals only -- not by shift!

        'If we are displaying a "Mass Balance" report for 1 shift then we need
        'to get mass balance data for the entire day for use with the reagent data
        'since the reagent data is available only for day total (not 1 shift).

        ReagentTitle = ""

        ReagBeginDate = aBeginDate
        ReagBeginShift = aBeginShift
        ReagEndDate = aEndDate
        ReagEndShift = aEndShift

        RerunBalance = False

        If aBeginDate = aEndDate And aBeginShift = aEndShift Then
            'This is a shift report!
            ReagBeginDate = aBeginDate
            ReagBeginShift = StrConv(gFirstShift, vbUpperCase)
            ReagEndDate = aEndDate
            ReagEndShift = StrConv(gLastShift, vbUpperCase)
            RerunBalance = True
        End If

        'For before the 3 -> 2 shift change
        If aBeginDate = aEndDate.AddDays(-1) And StrConv(aBeginShift, vbUpperCase) = "3RD" _
            And StrConv(aEndShift, vbUpperCase) = "2ND" Then
            'This is an offset shift report!
            'We want the reagent data for the second date
            ReagBeginDate = aEndDate
            ReagBeginShift = StrConv(gFirstShift, vbUpperCase)
            ReagEndDate = aEndDate
            ReagEndShift = StrConv(gLastShift, vbUpperCase)
            RerunBalance = True
        End If

        'For after the 3 -> 2 shift change
        If aBeginDate = aEndDate.AddDays(-1) And StrConv(aBeginShift, vbUpperCase) = "NIGHT" _
            And StrConv(aEndShift, vbUpperCase) = "DAY" Then
            'This is an offset shift report!
            'We want the reagent data for the second date
            ReagBeginDate = aEndDate
            ReagBeginShift = StrConv(gFirstShift, vbUpperCase)
            ReagEndDate = aEndDate
            ReagEndShift = StrConv(gLastShift, vbUpperCase)
            RerunBalance = True
        End If

        If RerunBalance = True Then
            NumShifts = gGetFcFloatPlantBalanceData(FloatPlantCirc, _
                                                    FloatPlantGmt, _
                                                    ReagBeginDate, _
                                                    ReagBeginShift, _
                                                    ReagEndDate, _
                                                    ReagEndShift, _
                                                    aCrewNumber, _
                                                    aBplRound, _
                                                    aMassBalMode)
        End If

        With mMbFcTotal
            .RgTotRptFdTons = mMbFcTotal.TotPltFdTonsRpt
            .RgTotCnTons = mMbFcTotal.PrdCnTons
            .RgTotAdjFdTons = mMbFcTotal.TotPltFdTonsAdj
        End With

        'Now get the reagent data for the Mass Balance/Metallurgical report
        'Now get the reagent data for the Mass Balance/Metallurgical report
        'Now get the reagent data for the Mass Balance/Metallurgical report

        GotReagents = gGetMetReagentDataFc(ReagBeginDate, _
                                           ReagBeginShift, _
                                           ReagEndDate, _
                                           ReagEndShift, _
                                           aCrewNumber, _
                                           mMbFcTotal.RgTotAdjFdTons, _
                                           mMbFcTotal.RgTotRptFdTons, _
                                           mMbFcTotal.RgTotCnTons, _
                                           mMbFcReag)

        'Add the reagent data to the report formulas.
        With mMbFcReag
            'frmViewData.rptInputData.Formulas(74) = "RgSuTotUnits = " & .RgSuTotUnits & ""
            'frmViewData.rptInputData.Formulas(75) = "RgAmTotUnits = " & .RgAmTotUnits & ""
            'frmViewData.rptInputData.Formulas(76) = "RgSaTotUnits = " & .RgSaTotUnits & ""
            'frmViewData.rptInputData.Formulas(77) = "RgSoTotUnits = " & .RgSoTotUnits & ""
            'frmViewData.rptInputData.Formulas(78) = "RgFaTotUnits = " & .RgFaTotUnits & ""
            'frmViewData.rptInputData.Formulas(79) = "RgFoTotUnits = " & .RgFoTotUnits & ""
            'frmViewData.rptInputData.Formulas(80) = "RgDeTotUnits = " & .RgDeTotUnits & ""
            'frmViewData.rptInputData.Formulas(81) = "RgSiTotUnits = " & .RgSiTotUnits & ""
            'frmViewData.rptInputData.Formulas(82) = "RgAllTotUnits = " & .RgAllTotUnits & ""
            ''-----
            'frmViewData.rptInputData.Formulas(83) = "RgSuTotCost = " & .RgSuTotCost & ""
            'frmViewData.rptInputData.Formulas(84) = "RgAmTotCost = " & .RgAmTotCost & ""
            'frmViewData.rptInputData.Formulas(85) = "RgSaTotCost = " & .RgSaTotCost & ""
            'frmViewData.rptInputData.Formulas(86) = "RgSoTotCost = " & .RgSoTotCost & ""
            'frmViewData.rptInputData.Formulas(87) = "RgFaTotCost = " & .RgFaTotCost & ""
            'frmViewData.rptInputData.Formulas(88) = "RgFoTotCost = " & .RgFoTotCost & ""
            'frmViewData.rptInputData.Formulas(89) = "RgDeTotCost = " & .RgDeTotCost & ""
            'frmViewData.rptInputData.Formulas(90) = "RgSiTotCost = " & .RgSiTotCost & ""
            'frmViewData.rptInputData.Formulas(91) = "RgAllTotCost = " & .RgAllTotCost & ""
            ''-----
            'frmViewData.rptInputData.Formulas(92) = "RgSuAdjFdDpt = " & .RgSuAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(93) = "RgAmAdjFdDpt = " & .RgAmAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(94) = "RgSaAdjFdDpt = " & .RgSaAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(95) = "RgSoAdjFdDpt = " & .RgSoAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(96) = "RgFaAdjFdDpt = " & .RgFaAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(97) = "RgFoAdjFdDpt = " & .RgFoAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(98) = "RgDeAdjFdDpt = " & .RgDeAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(99) = "RgSiAdjFdDpt = " & .RgSiAdjFdDpt & ""
            'frmViewData.rptInputData.Formulas(100) = "RgAllAdjFdDpt = " & .RgAllAdjFdDpt & ""
            ''-----
            'frmViewData.rptInputData.Formulas(101) = "RgSuRptFdDpt = " & .RgSuRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(102) = "RgAmRptFdDpt = " & .RgAmRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(103) = "RgSaRptFdDpt = " & .RgSaRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(104) = "RgSoRptFdDpt = " & .RgSoRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(105) = "RgFaRptFdDpt = " & .RgFaRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(106) = "RgFoRptFdDpt = " & .RgFoRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(107) = "RgDeRptFdDpt = " & .RgDeRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(108) = "RgSiRptFdDpt = " & .RgSiRptFdDpt & ""
            'frmViewData.rptInputData.Formulas(109) = "RgAllRptFdDpt = " & .RgAllRptFdDpt & ""
            ''-----
            'frmViewData.rptInputData.Formulas(110) = "RgSuCnDpt = " & .RgSuCnDpt & ""
            'frmViewData.rptInputData.Formulas(111) = "RgAmCnDpt = " & .RgAmCnDpt & ""
            'frmViewData.rptInputData.Formulas(112) = "RgSaCnDpt = " & .RgSaCnDpt & ""
            'frmViewData.rptInputData.Formulas(113) = "RgSoCnDpt = " & .RgSoCnDpt & ""
            'frmViewData.rptInputData.Formulas(114) = "RgFaCnDpt = " & .RgFaCnDpt & ""
            'frmViewData.rptInputData.Formulas(115) = "RgFoCnDpt = " & .RgFoCnDpt & ""
            'frmViewData.rptInputData.Formulas(116) = "RgDeCnDpt = " & .RgDeCnDpt & ""
            'frmViewData.rptInputData.Formulas(117) = "RgSiCnDpt = " & .RgSiCnDpt & ""
            'frmViewData.rptInputData.Formulas(118) = "RgAllCnDpt = " & .RgAllCnDpt & ""
            ''-----
            'frmViewData.rptInputData.Formulas(119) = "RgSuAdjFdUpt = " & .RgSuAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(120) = "RgAmAdjFdUpt = " & .RgAmAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(121) = "RgSaAdjFdUpt = " & .RgSaAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(122) = "RgSoAdjFdUpt = " & .RgSoAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(123) = "RgFaAdjFdUpt = " & .RgFaAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(124) = "RgFoAdjFdUpt = " & .RgFoAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(125) = "RgDeAdjFdUpt = " & .RgDeAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(126) = "RgSiAdjFdUpt = " & .RgSiAdjFdUpt & ""
            'frmViewData.rptInputData.Formulas(127) = "RgAllAdjFdUpt = " & .RgAllAdjFdUpt & ""
            ''-----
            'frmViewData.rptInputData.Formulas(128) = "RgSuRptFdUpt = " & .RgSuRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(129) = "RgAmRptFdUpt = " & .RgAmRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(130) = "RgSaRptFdUpt = " & .RgSaRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(131) = "RgSoRptFdUpt = " & .RgSoRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(132) = "RgFaRptFdUpt = " & .RgFaRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(133) = "RgFoRptFdUpt = " & .RgFoRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(134) = "RgDeRptFdUpt = " & .RgDeRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(135) = "RgSiRptFdUpt = " & .RgSiRptFdUpt & ""
            'frmViewData.rptInputData.Formulas(136) = "RgAllRptFdUpt = " & .RgAllRptFdUpt & ""
            ''-----
            'frmViewData.rptInputData.Formulas(137) = "RgSuCnUpt = " & .RgSuCnUpt & ""
            'frmViewData.rptInputData.Formulas(138) = "RgAmCnUpt = " & .RgAmCnUpt & ""
            'frmViewData.rptInputData.Formulas(139) = "RgSaCnUpt = " & .RgSaCnUpt & ""
            'frmViewData.rptInputData.Formulas(140) = "RgSoCnUpt = " & .RgSoCnUpt & ""
            'frmViewData.rptInputData.Formulas(141) = "RgFaCnUpt = " & .RgFaCnUpt & ""
            'frmViewData.rptInputData.Formulas(142) = "RgFoCnUpt = " & .RgFoCnUpt & ""
            'frmViewData.rptInputData.Formulas(143) = "RgDeCnUpt = " & .RgDeCnUpt & ""
            'frmViewData.rptInputData.Formulas(144) = "RgSiCnUpt = " & .RgSiCnUpt & ""
            'frmViewData.rptInputData.Formulas(145) = "RgAllCnUpt = " & .RgAllCnUpt & ""
            '-----
        End With

        'With mMbFcTotal
        '    frmViewData.rptInputData.Formulas(146) = "RgTotRptFdTons = " & .RgTotRptFdTons & ""
        '    frmViewData.rptInputData.Formulas(147) = "RgTotCnTons = " & .RgTotCnTons & ""
        '    frmViewData.rptInputData.Formulas(148) = "RgTotAdjFdTons = " & .RgTotAdjFdTons & ""
        'End With

        'Need to create the reagent title!
        If aBeginDate = aEndDate And aBeginShift = aEndShift Then
            'This is a shift report!
            ReagentTitle = "Reagent Usage & Cost   (Day Totals for " & _
                           Format(aEndDate, "MM/dd/yyyy") & ")"
        Else
            If aBeginDate = aEndDate.AddDays(-1) And StrConv(aBeginShift, vbUpperCase) = "3RD" And _
                StrConv(aEndShift, vbUpperCase) = "2ND" Then
                'This is an offset shift report!
                ReagentTitle = "Reagent Usage & Cost   (Day Totals for " & _
                               Format(aEndDate, "MM/dd/yyyy") & ")"
            Else
                ReagentTitle = "Reagent Usage & Cost"
            End If
        End If

        'Need to pass the company name into the report
        'frmViewData.rptInputData.ParameterFields(0) = "pCompanyName;" & gCompanyName & ";TRUE"

        ''Need to pass the reagent title into the report
        'frmViewData.rptInputData.ParameterFields(1) = "pReagentTitle;" & ReagentTitle & ";TRUE"

        ''Have all the needed data -- start the report
        'frmViewData.rptInputData.ReportFileName = gPath + "\Reports\" + _
        '                                          "MassBalanceFc2.rpt"

        'Connect to Oracle database
        ConnectString = "DSN = " + gDataSource + ";UID = " + gOracleUserName + _
            ";PWD = " + gOracleUserPassword + ";DSQ = "

        'frmViewData.rptInputData.Connect = ConnectString
        ''Report window maximized
        'frmViewData.rptInputData.WindowState = crptMaximized

        'frmViewData.rptInputData.WindowTitle = "Four Corners Mass Balance"

        ''User not allowed to minimize report window
        'frmViewData.rptInputData.WindowMinButton = False

        ''Start Crystal Reports
        'frmViewData.rptInputData.action = 1

        'frmViewData.rptInputData.ReportFileName = ""
        'frmViewData.rptInputData.Reset()

        Exit Function

gMassBalanceFcError:

        MsgBox("Error printing Four Corners Mass Balance report." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Four Corners Mass Balance Reporting Error")
    End Function

    Public Function gMetallurgicalFC(ByVal aBeginDate As Date, _
                                     ByVal aBeginShift As String, _
                                     ByVal aEndDate As Date, _
                                     ByVal aEndShift As String, _
                                     ByVal aCrewNumber As String, _
                                     ByVal aBplRound As Integer) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        '01/23/2006, lss -- Don't have a separate Metallurgical report at this time.
    End Function

    Public Function gGetFcFloatPlantBalanceData(ByRef FloatPlantCirc As Object, _
                                                ByRef FloatPlantGmt As Object, _
                                                ByVal aBeginDate As Date, _
                                                ByVal aBeginShift As String, _
                                                ByVal aEndDate As Date, _
                                                ByVal aEndShift As String, _
                                                ByVal aCrewNumber As String, _
                                                ByVal aBplRound As Integer, _
                                                ByVal aMassBalMode As String) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'This function will return the number of shifts processed.
        'It will also "return" data through the FloatPlantCirc() and
        'FloatPlantGmt() arrays.

        On Error GoTo gGetFcFloatPlantBalanceDataError

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

        Dim NumShifts As Integer
        Dim SumShifts As Boolean

        mBeginDate = aBeginDate
        mBeginShift = aBeginShift
        mEndDate = aEndDate
        mEndShift = aEndShift

        '03/23/2006, lss
        'Will not sum shifts per Glen Oswald.  Will average the circuit BPL's
        'and sum the reported feed and concentrate tons for the period
        'instead.
        SumShifts = False

        CalcNumShifts = gGetNumShiftsRge2("Four Corners", _
                                          aBeginDate, _
                                          aEndDate)

        'Mass balances are only run for either one shift or for
        'a range of complete days.
        If aBeginDate = aEndDate And aBeginShift = aEndShift Then
            CalcNumShifts = 1
        End If

        'aCrewNumber will be "All", "A", "B", "C", or "D"
        If aCrewNumber = "All" Then
            NumShifts = CalcNumShifts
        Else
            'NumShifts = gGetCrewShiftCount("Four Corners", _
            '                               aBeginDate, _
            '                               aBeginShift, _
            '                               aEndDate, _
            '                               aEndShift, _
            '                               aCrewNumber)
        End If

        ReDim FloatPlantCirc(0 To 9, 0 To 14)
        ReDim FloatPlantGmt(0 To 4, 0 To 7)

        'fFloatPlantCirc
        '---------------
        '
        '       Rows                 Columns
        '       --------------       ----------------
        ' 1)    Fine rougher         Hours
        ' 2)    Fine amine           Feed tons reported
        ' 3)    Total fine           Feed tons adjusted
        ' 4)    Coarse rougher       Feed BPL
        ' 5)    Coarse amine         Conc BPL
        ' 6)    Total coarse         Tail BPL
        ' 7)    Total amine          Ratio of concentration
        ' 8)    Grand totals         %Actual recovery
        ' 9)    Concentrate product  %Standard recovery
        '10)                         Concentrate tons adjusted
        '11)                         Tail tons adjusted
        '12)                         Feed TPH
        '13)                         Concentrate TPH
        '14)                         Tail TPH

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

        For RowIdx = 1 To 9
            For ColIdx = 1 To 14
                FloatPlantCirc(RowIdx, ColIdx) = 0
            Next ColIdx
        Next RowIdx

        For RowIdx = 1 To 4
            For ColIdx = 1 To 7
                FloatPlantGmt(RowIdx, ColIdx) = 0
            Next ColIdx
        Next RowIdx

        ZeroFcSummingData()

        If SumShifts = True Then
            ZeroFcShiftData()

            'Get basic floatplant data from EQPT_MSRMNT, EQPT_EXT_MSRMNT, EQPT_CALC
            params = gDBParams

            params.Add("pMineName", "Four Corners", ORAPARM_INPUT)
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

            'PROCEDURE get_mass_balance_data
            'pMineName           IN     VARCHAR2,
            'pBeginDate          IN     DATE,
            'pBeginShift         IN     VARCHAR2,
            'pEndDate            IN     DATE,
            'pEndShift           IN     VARCHAR2,
            'pCrewNumber         IN     VARCHAR2,
            'pResult             IN OUT c_massbalance);
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_floatplant.get_mass_balance_data(:pMineName," + _
                          ":pBeginDate, :pBeginShift, :pEndDate, :pEndShift, :pCrewNumber, :pResult);end;", ORASQL_FAILEXEC)
            mMassBalanceDynaset = params("pResult").Value
            RecCount = mMassBalanceDynaset.RecordCount
            ClearParams(params)

            If RecCount = 0 Then
                Exit Function
            End If
        Else
            'Will average the circuit BPL's and sum the reported feed and
            'concentrate tons for the period instead.
            'The data will be in mMbFcShift
            ZeroFcShiftData()

            GetPeriodAvgsAndSums(aBeginDate, _
                                 aBeginShift, _
                                 aEndDate, _
                                 aEndShift, _
                                 aBplRound, _
                                 aMassBalMode)
        End If

        If SumShifts = True Then
            mMassBalanceDynaset.MoveFirst()
            CurrentDate = mMassBalanceDynaset.Fields("prod_date").Value
            CurrentShift = mMassBalanceDynaset.Fields("shift").Value

            Do While Not mMassBalanceDynaset.EOF
                ThisDate = mMassBalanceDynaset.Fields("prod_date").Value
                ThisShift = mMassBalanceDynaset.Fields("shift").Value

                If ThisDate = CurrentDate And ThisShift = CurrentShift Then
                    ThisEqpt = mMassBalanceDynaset.Fields("eqpt_name").Value

                    With mMbFcShift
                        Select Case ThisEqpt
                            Case Is = "North fine rougher"
                                'Feed BPL, Concentrate BPL and Tail BPL
                                .NfrFdBplRpt = mMassBalanceDynaset.Fields("feed_bpl").Value
                                .NfrCnBplRpt = mMassBalanceDynaset.Fields("concentrate_bpl").Value
                                .NfrTlBplRpt = mMassBalanceDynaset.Fields("tail_bpl").Value

                            Case Is = "South fine rougher"
                                'Feed BPL, Concentrate BPL and Tail BPL
                                .SfrFdBplRpt = mMassBalanceDynaset.Fields("feed_bpl").Value
                                .SfrCnBplRpt = mMassBalanceDynaset.Fields("concentrate_bpl").Value
                                .SfrTlBplRpt = mMassBalanceDynaset.Fields("tail_bpl").Value

                            Case Is = "North coarse rougher"
                                'Feed BPL, Tail BPL, Operating hours and Reported feed tons
                                .NcrFdBplRpt = mMassBalanceDynaset.Fields("feed_bpl").Value
                                .NcrTlBplRpt = mMassBalanceDynaset.Fields("tail_bpl").Value
                                .NcrHrs = mMassBalanceDynaset.Fields("operating_hours").Value
                                .NcrFdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value

                            Case Is = "South coarse rougher"
                                'Feed BPL, Tail BPL, Operating hours and Reported feed tons
                                .ScrFdBplRpt = mMassBalanceDynaset.Fields("feed_bpl").Value
                                .ScrTlBplRpt = mMassBalanceDynaset.Fields("tail_bpl").Value
                                .ScrHrs = mMassBalanceDynaset.Fields("operating_hours").Value
                                .ScrFdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value

                            Case Is = "Coarse rougher"
                                'Concentrate BPL
                                .CrCnBplRpt = mMassBalanceDynaset.Fields("concentrate_bpl").Value

                            Case Is = "Coarse column"
                                'Feed BPL, Concentrate BPL and Tail BPL
                                .CrsColFdBplRpt = mMassBalanceDynaset.Fields("feed_bpl").Value
                                .CrsColCnBplRpt = mMassBalanceDynaset.Fields("concentrate_bpl").Value
                                .CrsColTlBplRpt = mMassBalanceDynaset.Fields("tail_bpl").Value

                            Case Is = "Fine amine"
                                'Tail BPL
                                .FaTlBplRpt = mMassBalanceDynaset.Fields("tail_bpl").Value

                            Case Is = "Coarse amine"
                                'Tail BPL
                                .CaTlBplRpt = mMassBalanceDynaset.Fields("tail_bpl").Value

                            Case Is = "Float plant"
                                .PrdCnTons = mMassBalanceDynaset.Fields("concentrate_product_tons").Value
                                .PrdCnBpl = mMassBalanceDynaset.Fields("concentrate_product_bpl").Value
                                .GmtBplRpt = mMassBalanceDynaset.Fields("tail_bpl").Value

                            Case Is = "North fine rougher 1"
                                'Operating hours and Reported feed tons
                                .Nfr1Hrs = mMassBalanceDynaset.Fields("operating_hours").Value
                                .Nfr1FdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value

                            Case Is = "North fine rougher 2"
                                'Operating hours and Reported feed tons
                                .Nfr2Hrs = mMassBalanceDynaset.Fields("operating_hours").Value
                                .Nfr2FdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value

                            Case Is = "South fine rougher 1"
                                'Operating hours and Reported feed tons
                                .Sfr1Hrs = mMassBalanceDynaset.Fields("operating_hours").Value
                                .Sfr1FdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value

                            Case Is = "South fine rougher 2"
                                'Operating hours and Reported feed tons
                                .Sfr2Hrs = mMassBalanceDynaset.Fields("operating_hours").Value
                                .Sfr2FdTonsRpt = mMassBalanceDynaset.Fields("reported_feed_tons").Value

                            Case Is = "North coarse scalp"
                                'Operating hours
                                .NcsHrs = mMassBalanceDynaset.Fields("operating_hours").Value

                            Case Is = "South coarse scalp"
                                'Operating hours
                                .ScsHrs = mMassBalanceDynaset.Fields("operating_hours").Value
                        End Select
                    End With
                    mMassBalanceDynaset.MoveNext()
                Else    'ThisDate <> CurrentDate Or ThisShift <> CurrentShift
                    'Have all data for this shift -- process it!
                    ProcessFcMassBalanceData(CurrentDate, _
                                             CurrentShift, _
                                             aBplRound, _
                                             SumShifts)

                    CurrentDate = ThisDate
                    CurrentShift = ThisShift
                    ZeroFcShiftData()
                End If
            Loop
        End If

        'NOTE: If SumShifts were false then the data will be processed here
        '      in the next ProcessFcMassBalanceData
        If SumShifts = False Then
            CurrentDate = aBeginDate
            CurrentShift = aBeginShift
        End If

        'Process last shift's worth of data if necessary
        ProcessFcMassBalanceData(CurrentDate, _
                                 CurrentShift, _
                                 aBplRound, _
                                 SumShifts)

        'Summing of mass balance shift data completed
        ProcessFcMassBalanceTotals(aBplRound)

        'Place data in array  Place data in array  Place data in array
        'Place data in array  Place data in array  Place data in array
        'Place data in array  Place data in array  Place data in array

        'FloatPlantCirc()

        'Rows in the array                Columns in the array
        'fcFneRghr = 1                    fcCcOperHrs = 1
        'fcFneAmine = 2                   fcCcFdTonsRpt = 2
        'fcTotFne = 3                     fcCcFdTonsAdj = 3
        'fcCrsRghr = 4                    fcCcFdBpl = 4
        'fcCrsAmine = 5                   fcCcCnBpl = 5
        'fcTotCrs = 6                     fcCcTlBpl = 6
        'fcTotAmine = 7                   fcCcRC = 7
        'fcTotPlant = 8                   fcCcPctActRcvry = 8
        'fcCrCnProduct = 9                fcCcPctStdRcvry = 9
        '                                 fcCcCnTonsAdj = 10
        '                                 fcCcTlTonsAdj = 11
        '                                 fcCcFdTph = 12
        '                                 fcCcCnTph = 13
        '                                 fcCcTlTph = 14

        'Product tons
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCnProduct, mFcFloatPlantCircColEnum.fcCcCnTonsAdj) = mMbFcTotal.PrdCnTons
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCnProduct, mFcFloatPlantCircColEnum.fcCcCnBpl) = mMbFcTotal.PrdCnBpl

        'Operating hours
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcOperHrs) = mMbFcTotal.FrHrs
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcOperHrs) = mMbFcTotal.CrHrs
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcOperHrs) = mMbFcTotal.FaHrs
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcOperHrs) = mMbFcTotal.CaHrs
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcOperHrs) = mMbFcTotal.TotPltHrs

        'Feed tons reported
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcFdTonsRpt) = mMbFcTotal.FrFdTonsRpt
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcFdTonsRpt) = mMbFcTotal.CrFdTonsRpt
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcFdTonsRpt) = mMbFcTotal.FaFdTonsRpt
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcFdTonsRpt) = mMbFcTotal.CaFdTonsRpt
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcFdTonsRpt) = mMbFcTotal.TotPltFdTonsRpt

        'Feed tons adjusted
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcFdTonsAdj) = mMbFcTotal.FrFdTonsAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcFdTonsAdj) = mMbFcTotal.CrFdTonsAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcFdTonsAdj) = mMbFcTotal.FaFdTonsAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcFdTonsAdj) = mMbFcTotal.CaFdTonsAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcFdTonsAdj) = mMbFcTotal.TotPltFdTonsAdj

        'Feed BPL
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcFdBpl) = mMbFcTotal.FrFdBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcFdBpl) = mMbFcTotal.CrFdBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcFdBpl) = mMbFcTotal.FaFdBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcFdBpl) = mMbFcTotal.CaFdBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcFdBpl) = mMbFcTotal.TotPltFdBplAdj

        'Concentrate BPL
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcCnBpl) = mMbFcTotal.FrCnBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcCnBpl) = mMbFcTotal.CrCnBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcCnBpl) = mMbFcTotal.FaCnBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcCnBpl) = mMbFcTotal.CaCnBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcCnBpl) = mMbFcTotal.TotPltCnBplAdj

        'Tail BPL
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcTlBpl) = mMbFcTotal.FrTlBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcTlBpl) = mMbFcTotal.CrTlBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcTlBpl) = mMbFcTotal.FaTlBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcTlBpl) = mMbFcTotal.CaTlBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcTlBpl) = mMbFcTotal.TotPltTlBplAdj

        'Ratio of concentration
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcRc) = mMbFcTotal.FrRcAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcRc) = mMbFcTotal.CrRcAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcRc) = mMbFcTotal.FaRcAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcRc) = mMbFcTotal.CaRcAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcRc) = mMbFcTotal.TotPltRcAdj

        'Actual recovery
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcPctActRcvry) = mMbFcTotal.FrPctAdjRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcPctActRcvry) = mMbFcTotal.CrPctAdjRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcPctActRcvry) = mMbFcTotal.FaPctAdjRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcPctActRcvry) = mMbFcTotal.CaPctAdjRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcPctActRcvry) = mMbFcTotal.TotPltPctAdjRcvry

        'Reported recovery
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcPctRptRcvry) = mMbFcTotal.FrPctRptRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcPctRptRcvry) = mMbFcTotal.CrPctRptRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcPctRptRcvry) = mMbFcTotal.FaPctRptRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcPctRptRcvry) = mMbFcTotal.CaPctRptRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcPctRptRcvry) = mMbFcTotal.TotPltPctRptRcvry

        'Concentrate tons adjusted
        'Nothing to put in here right now.

        'Tail tons adjusted
        'Nothing to put in here right now.

        'Feed TPH adjusted
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcFdTph) = mMbFcTotal.TotPltFdTphAdj

        'Concentrate TPH adjusted
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcCnTph) = mMbFcTotal.TotPltCnTphAdj

        'Tail TPH adjusted
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcTlTph) = mMbFcTotal.TotPltTlTphAdj

        'FloatPlantGmt

        'Rows in the array              Columns in the array
        'grAsReportedGmtBpl = 1         gcFdTons = 1
        'grCalculatedGmtBpl = 2         gcCnTons = 2
        'grReportedFdTons = 3           gcFdBpl = 3
        'grGmtBplFromCircuits = 4       gcCnBpl = 4
        '                               gcTlBpl = 5
        '                               gcRC = 6
        '                               gcPctRcvry = 7

        'Based on as reported GMT BPL
        'Only interested in an as reported Gmt BPL value here
        'We have two total plant tail BPL's we could use here
        '1) mMbFcTotal.TotPltTlBplRpt    Based on measured circuit tail BPL's
        '2) mMbFcTotal.TotPltTlBplRpt2   Based on total plant measured GMT

        'Will use mMbFcTotal.TotPltTlBplRpt for now.

        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrAsReportedGmtBpl, mWgFloatPlantGmtColEnum.fcGcFdBpl) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrAsReportedGmtBpl, mWgFloatPlantGmtColEnum.fcGcCnBpl) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrAsReportedGmtBpl, mWgFloatPlantGmtColEnum.fcGcTlBpl) = mMbFcTotal.TotPltTlBplRpt
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrAsReportedGmtBpl, mWgFloatPlantGmtColEnum.fcGcRc) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrAsReportedGmtBpl, mWgFloatPlantGmtColEnum.fcGcPctRcvry) = 0

        'Based on Adjusted or Calculated feed tons
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcFdTons) = mMbFcTotal.TotPltFdTonsAdj
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcCnTons) = mMbFcTotal.TotPltCnTonsAdj
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcFdBpl) = mMbFcTotal.TotPltFdBplAdj
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcCnBpl) = mMbFcTotal.TotPltCnBplAdj
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcTlBpl) = mMbFcTotal.TotPltTlBplAdj
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcRc) = mMbFcTotal.TotPltRcAdj
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcPctRcvry) = mMbFcTotal.TotPltPctAdjRcvry

        'Based on reported feed tons.
        'Will not put anything here for Four Corners right now (will just
        'fill in zeros).
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcFdTons) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcCnTons) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcFdBpl) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcCnBpl) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcTlBpl) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcRc) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcPctRcvry) = 0

        'This will not really apply at Four Corners at this time (will just
        'assign a zero).
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrGmtBplFromCircuits, mWgFloatPlantGmtColEnum.fcGcTlBpl) = 0

        gGetFcFloatPlantBalanceData = NumShifts

        Exit Function

gGetFcFloatPlantBalanceDataError:

        MsgBox("Error calculating Four Corners Mass Balance." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Four Corners Mass Balance Calculation Error")

        On Error Resume Next
        ClearParams(params)
    End Function

    Private Sub ZeroFcSummingData()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        With mMbFcTotal
            .PrdCnTons = 0
            .PrdCnTonsW = 0
            .PrdCnBt = 0
            .PrdCnBpl = 0
            '----------
            .Nfr1Hrs = 0
            .Nfr2Hrs = 0
            .Sfr1Hrs = 0
            .Sfr2Hrs = 0
            .NcrHrs = 0
            .ScrHrs = 0
            .NcsHrs = 0
            .ScsHrs = 0
            .FrHrs = 0
            .CrHrs = 0
            .FaHrs = 0
            .CaHrs = 0
            .TotPltHrs = 0
            .NcrCaHrs = 0
            .ScrCaHrs = 0
            '----------
            .FrFdTonsRpt = 0
            .FrFdBtRpt = 0
            .FrFdBplRpt = 0
            .FrCnTonsRpt = 0
            .FrCnBtRpt = 0
            .FrCnBplRpt = 0
            .FrTlTonsRpt = 0
            .FrTlBtRpt = 0
            .FrTlBplRpt = 0
            '----------
            .FaFdTonsRpt = 0
            .FaFdBtRpt = 0
            .FaFdBplRpt = 0
            .FaCnTonsRpt = 0
            .FaCnBtRpt = 0
            .FaCnBplRpt = 0
            .FaTlTonsRpt = 0
            .FaTlBtRpt = 0
            .FaTlBplRpt = 0
            '----------
            .CrFdTonsRpt = 0
            .CrFdBtRpt = 0
            .CrFdBplRpt = 0
            .CrCnTonsRpt = 0
            .CrCnBtRpt = 0
            .CrCnBplRpt = 0
            .CrTlTonsRpt = 0
            .CrTlBtRpt = 0
            .CrTlBplRpt = 0
            '----------
            .CaFdTonsRpt = 0
            .CaFdBtRpt = 0
            .CaFdBplRpt = 0
            .CaCnTonsRpt = 0
            .CaCnBtRpt = 0
            .CaCnBplRpt = 0
            .CaTlTonsRpt = 0
            .CaTlBtRpt = 0
            .CaTlBplRpt = 0
            '----------
            .FrFdTonsAdj = 0
            .FrFdBtAdj = 0
            .FrFdBplAdj = 0
            .FrCnTonsAdj = 0
            .FrCnBtAdj = 0
            .FrCnBplAdj = 0
            .FrTlTonsAdj = 0
            .FrTlBtAdj = 0
            .FrTlBplAdj = 0
            '----------
            .FaFdTonsAdj = 0
            .FaFdBtAdj = 0
            .FaFdBplAdj = 0
            .FaCnTonsAdj = 0
            .FaCnBtAdj = 0
            .FaCnBplAdj = 0
            .FaTlTonsAdj = 0
            .FaTlBtAdj = 0
            .FaTlBplAdj = 0
            '----------
            .CrFdTonsAdj = 0
            .CrFdBtAdj = 0
            .CrFdBplAdj = 0
            .CrCnTonsAdj = 0
            .CrCnBtAdj = 0
            .CrCnBplAdj = 0
            .CrTlTonsAdj = 0
            .CrTlBtAdj = 0
            .CrTlBplAdj = 0
            '----------
            .CaFdTonsAdj = 0
            .CaFdBtAdj = 0
            .CaFdBplAdj = 0
            .CaCnTonsAdj = 0
            .CaCnBtAdj = 0
            .CaCnBplAdj = 0
            .CaTlTonsAdj = 0
            .CaTlBtAdj = 0
            .CaTlBplAdj = 0
            '----------
            .FrRcAdj = 0
            .FaRcAdj = 0
            .CrRcAdj = 0
            .CaRcAdj = 0
            '----------
            .FrFdTphRpt = 0
            .FrCnTphRpt = 0
            .FrTlTphRpt = 0
            .CrFdTphRpt = 0
            .CrCnTphRpt = 0
            .CrTlTphRpt = 0
            .FaFdTphRpt = 0
            .FaCnTphRpt = 0
            .FaTlTphRpt = 0
            .CaFdTphRpt = 0
            .CaCnTphRpt = 0
            .CaTlTphRpt = 0
            '----------
            .FrFdTphAdj = 0
            .FrCnTphAdj = 0
            .FrTlTphAdj = 0
            .CrFdTphAdj = 0
            .CrCnTphAdj = 0
            .CrTlTphAdj = 0
            .FaFdTphAdj = 0
            .FaCnTphAdj = 0
            .FaTlTphAdj = 0
            .CaFdTphAdj = 0
            .CaCnTphAdj = 0
            .CaTlTphAdj = 0
            '----------
            .TotPltFdBplAdj = 0
            .TotPltCnBplAdj = 0
            .TotPltTlBplAdj = 0
            '----------
            .TotPltTlTonsAdj = 0
            .TotPltGmtTlBtRpt = 0
            .TotPltTlBplRpt = 0
            .TotPltTlBplRpt2 = 0
            .TotPltFdBplRpt = 0
            '----------
            .TotPltFdTonsRpt = 0
            .TotPltFdTonsAdj = 0
            .TotPltTlTonsRpt = 0
            '----------
            .FrPctAdjRcvry = 0
            .FaPctAdjRcvry = 0
            .CrPctAdjRcvry = 0
            .CaPctAdjRcvry = 0
            .TotPltPctAdjRcvry = 0
            '----------
            .FrPctRptRcvry = 0
            .FaPctRptRcvry = 0
            .CrPctRptRcvry = 0
            .CaPctRptRcvry = 0
            .TotPltPctRptRcvry = 0
            '----------
            .TotPltRcAdj = 0
            .TotPltCnTonsAdj = 0
            .TotPltFdTphAdj = 0
            .TotPltCnTphAdj = 0
            .TotPltTlTphAdj = 0
            .TotPltCnBplRpt = 0
            '----------
            .RgTotRptFdTons = 0
            .RgTotCnTons = 0
            .RgTotAdjFdTons = 0
            '----------
            .BalDistFrPct = 0
            .BalDistCrPct = 0
            .BalDistFaPct = 0
            .BalDistCaPct = 0
        End With

        With mMbFcReag
            .RgSuTotUnits = 0
            .RgAmTotUnits = 0
            .RgSaTotUnits = 0
            .RgSoTotUnits = 0
            .RgFaTotUnits = 0
            .RgFa2TotUnits = 0
            .RgFoTotUnits = 0
            .RgDeTotUnits = 0
            .RgSiTotUnits = 0
            .RgAllTotUnits = 0
            '----------
            .RgSuTotCost = 0
            .RgAmTotCost = 0
            .RgSaTotCost = 0
            .RgSoTotCost = 0
            .RgFaTotCost = 0
            .RgFa2TotCost = 0
            .RgFoTotCost = 0
            .RgDeTotCost = 0
            .RgSiTotCost = 0
            .RgAllTotCost = 0
            '----------
            .RgSuAdjFdDpt = 0
            .RgAmAdjFdDpt = 0
            .RgSaAdjFdDpt = 0
            .RgSoAdjFdDpt = 0
            .RgFaAdjFdDpt = 0
            .RgFa2AdjFdDpt = 0
            .RgFoAdjFdDpt = 0
            .RgDeAdjFdDpt = 0
            .RgSiAdjFdDpt = 0
            .RgAllAdjFdDpt = 0
            '----------
            .RgSuRptFdDpt = 0
            .RgAmRptFdDpt = 0
            .RgSaRptFdDpt = 0
            .RgSoRptFdDpt = 0
            .RgFaRptFdDpt = 0
            .RgFa2RptFdDpt = 0
            .RgFoRptFdDpt = 0
            .RgDeRptFdDpt = 0
            .RgSiRptFdDpt = 0
            .RgAllRptFdDpt = 0
            '----------
            .RgSuCnDpt = 0
            .RgAmCnDpt = 0
            .RgSaCnDpt = 0
            .RgSoCnDpt = 0
            .RgFaCnDpt = 0
            .RgFa2CnDpt = 0
            .RgFoCnDpt = 0
            .RgDeCnDpt = 0
            .RgSiCnDpt = 0
            .RgAllCnDpt = 0
            '----------
            .RgSuAdjFdUpt = 0
            .RgAmAdjFdUpt = 0
            .RgSaAdjFdUpt = 0
            .RgSoAdjFdUpt = 0
            .RgFaAdjFdUpt = 0
            .RgFa2AdjFdUpt = 0
            .RgFoAdjFdUpt = 0
            .RgDeAdjFdUpt = 0
            .RgSiAdjFdUpt = 0
            .RgAllAdjFdUpt = 0
            '----------
            .RgSuRptFdUpt = 0
            .RgAmRptFdUpt = 0
            .RgSaRptFdUpt = 0
            .RgSoRptFdUpt = 0
            .RgFaRptFdUpt = 0
            .RgFa2RptFdUpt = 0
            .RgFoRptFdUpt = 0
            .RgDeRptFdUpt = 0
            .RgSiRptFdUpt = 0
            .RgAllRptFdUpt = 0
            '----------
            .RgSuCnUpt = 0
            .RgAmCnUpt = 0
            .RgSaCnUpt = 0
            .RgSoCnUpt = 0
            .RgFaCnUpt = 0
            .RgFa2CnUpt = 0
            .RgFoCnUpt = 0
            .RgDeCnUpt = 0
            .RgSiCnUpt = 0
            .RgAllCnUpt = 0
        End With
    End Sub

    Private Sub ProcessFcMassBalanceData(ByVal aDate As Date, _
                                         ByVal aShift As String, _
                                         ByVal aBplRound As Integer, _
                                         ByVal aSumShifts As Boolean)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo ProcessFcMassBalanceDataError

        Dim TotPct As Double
        Dim BalDist As mBalanceDistributionType
        Dim CloseEnough As Integer
        Dim IterationCount As Integer

        Dim TlTonsBalSum As Double
        Dim TlBtBalSum As Double
        Dim GmtBplDiff As Single
        Dim GmtBplTonDiff As Double

        Dim FrTlBplPrev As Single
        Dim CrTlBplPrev As Single
        Dim FaTlBplPrev As Single
        Dim CaTlBplPrev As Single
        Dim FrTlBplBal As Single
        Dim CrTlBplBal As Single
        Dim FaTlBplBal As Single
        Dim CaTlBplBal As Single
        Dim FrTlBplResid As Single
        Dim CrTlBplResid As Single
        Dim FaTlBplResid As Single
        Dim CaTlBplResid As Single
        Dim TotBplCheck As Single
        Dim TotBplCheckDiff As Single
        Dim CloseEnoughVal As Single

        Dim FrFdBplRpt As Single
        Dim FrCnBplRpt As Single
        Dim FrTlBplRpt As Single
        Dim FrFdBplAdj As Single
        Dim FrCnBplAdj As Single
        Dim FrTlBplAdj As Single
        Dim CrFdBplRpt As Single
        Dim CrCnBplRpt As Single
        Dim CrTlBplRpt As Single
        Dim CrFdBplAdj As Single
        Dim CrCnBplAdj As Single
        Dim CrTlBplAdj As Single

        Dim TotPltTlTonsRpt As Single

        Dim ShiftFrHrs As Single
        Dim ShiftFaHrs As Single
        Dim ShiftCrHrs As Single
        Dim ShiftCaHrs As Single
        Dim ShiftPltHrs As Single
        Dim ShiftTons As Single

        'Dim's for handling really screwed up data from the NARS at Four Corners
        Dim FrTlTonsAdjChk As Single
        Dim CrTlTonsAdjChk As Single
        Dim FaTlTonsAdjChk As Single
        Dim CaTlTonsAdjChk As Single
        Dim LeaveFrAlone As Boolean
        Dim LeaveCrAlone As Boolean
        Dim LeaveFaAlone As Boolean
        Dim LeaveCaAlone As Boolean
        Dim BalDistFrPct As Single
        Dim BalDistCrPct As Single
        Dim BalDistFaPct As Single
        Dim BalDistCaPct As Single
        Dim AmtToDistribute As Single
        Dim NewPct As Single

        'Need the "balance distributions" (periodic equipment measurements)
        'for Fine rougher, Coarse rougher, Fine amine and Coarse amine.
        'BalDist.FrPct    45 for 06/01/2005 per Glen Oswald
        'BalDist.CrPct    27 for 06/01/2005 per Glen Oswald
        'BalDist.FaPct    26 for 06/01/2005 per Glen Oswald
        'BalDist.CaPct    2  for 06/01/2005 per Glen Oswald
        'The 4 value should sum to 100!
        'These values are "estimates to be changed by the plant process
        'engineer as sampling conditions change" per Glen Oswald.
        If aSumShifts = True Then
            BalDist = gGetBalanceDistribution("Four Corners", _
                                              aDate, _
                                              aShift)
        Else
            BalDist = gGetBalanceDistribution2("Four Corners", _
                                               mBeginDate, _
                                               mBeginShift, _
                                               mEndDate, _
                                               mEndShift)
        End If

        CloseEnough = False
        IterationCount = 0
        CloseEnoughVal = 0.02

        'IMPORTANT:  We are making the assumption that there are no
        '            missing BPL's.  If tons exist then there should be
        '            corresponding BPL's

        'Set up the adjusted tail BPL's that we will be adjusting.
        'The feed BPL's and concentrate BPL's will not be changing.
        'These 6 tail BPL values represent circuit samples for the shift
        '1) North fine rougher     .NfrTlBplRpt
        '2) South fine rougher     .SfrTlBplRpt
        '3) North coarse rougher   .NcrTlBplRpt
        '4) South coarse rougher   .ScrTlBplRpt
        '5) Fine amine             .FaTlBplRpt
        '6) Coarse amine           .CaTlBplRpt

        With mMbFcShift
            .NfrTlBplAdj = .NfrTlBplRpt
            .SfrTlBplAdj = .SfrTlBplRpt
            .FaTlBplAdj = .FaTlBplRpt

            .NcrTlBplAdj = .NcrTlBplRpt
            .ScrTlBplAdj = .ScrTlBplRpt
            .CaTlBplAdj = .CaTlBplRpt
        End With

        Do Until CloseEnough = True Or IterationCount > 10
            IterationCount = IterationCount + 1

            'The tail BPL's (.NfrTlBplAdj, .SfrTlBplAdj, etc.) will change
            'as we go through each iteration.

            With mMbFcShift
                'Need to determine some information about when the
                'coarse amine circuit ran!
                ShiftCrHrs = Round((mMbFcShift.NcrHrs + mMbFcShift.ScrHrs) / 2, 2)
                .CrHrs = .CrHrs + ShiftCrHrs
                ShiftCaHrs = Round((mMbFcShift.NcrCaHrs + mMbFcShift.ScrCaHrs) / 2, 2)
                .CaHrs = .CaHrs + ShiftCaHrs
                If .CrHrs <> 0 Then
                    .CaPct = Round(.CaHrs / .CrHrs, 4)
                Else
                    .CaPct = 0
                End If

                'Get the ratio of concentrations
                'Get the ratio of concentrations
                'Get the ratio of concentrations

                'Since the tail BPL's will change as we go through each
                'iteration, the ratio of concentrations will also change
                'as we got through each iteration.

                'Fine Roughers and Fine Amine
                'North fine rougher ratio of concentration
                '.NfrTlBplAdj starts off as .NfrTlBplRpt and then changes
                'with each iteration.
                If .NfrFdBplRpt - .NfrTlBplAdj <> 0 Then
                    .NfrRc = Round((.NfrCnBplRpt - .NfrTlBplAdj) / _
                             (.NfrFdBplRpt - .NfrTlBplAdj), 2)
                Else
                    .NfrRc = 0
                End If

                'South fine rougher ratio of concentration
                '.SfrTlBplAdj starts off as .SfrTlBplRpt and then changes
                'with each iteration.
                If .SfrFdBplRpt - .SfrTlBplAdj <> 0 Then
                    .SfrRc = Round((.SfrCnBplRpt - .SfrTlBplAdj) / _
                             (.SfrFdBplRpt - .SfrTlBplAdj), 2)
                Else
                    .SfrRc = 0
                End If

                'gGet2NumAvg will return a "special" average -- it will not
                'average in a zero value if one exists.
                'The fine amine feed BPL will be the strait average of the
                'North fine rougher concentrate BPL and the South fine rougher
                'concentrate BPL.
                .AvgFrCnBpl = gGet2NumAvg(.NfrCnBplRpt, .SfrCnBplRpt, 2)

                'The fine amine concentrate BPL will be the total concentrate
                'product BPL for the shift mMbFcShift.PrdCnBpl.

                'Fine amine ratio of concentration
                '.FaTlBplAdj starts off as .FaTlBplRpt and then changes
                'with each iteration.
                If .AvgFrCnBpl - .FaTlBplAdj <> 0 Then
                    .FaRc = Round((.PrdCnBpl - .FaTlBplAdj) / _
                            (.AvgFrCnBpl - .FaTlBplAdj), 2)
                Else
                    .FaRc = 0
                End If

                'Coarse roughers and Coarse amine
                'We only have a Coarse rougher concentrate BPL (no North coarse
                'rougher and South coarse rougher concentrate BPL's).
                'We will use the Coarse rougher concentrate BPL for both the
                'North coarse rougher and the South coarse rougher.
                'North coarse rougher ratio of concentraion
                If .NcrFdBplRpt - .NcrTlBplAdj <> 0 Then
                    .NcrRc = Round((.CrCnBplRpt - .NcrTlBplAdj) / _
                             (.NcrFdBplRpt - .NcrTlBplAdj), 2)
                Else
                    .NcrRc = 0
                End If

                'South coarse rougher ratio of concentration
                If .ScrFdBplRpt - .ScrTlBplAdj <> 0 Then
                    .ScrRc = Round((.CrCnBplRpt - .ScrTlBplAdj) / _
                             (.ScrFdBplRpt - .ScrTlBplAdj), 2)
                Else
                    .ScrRc = 0
                End If

                'The "average" coarse rougher concentrate BPL will be the
                'Coarse rougher concentrate BPL since we don't have North
                'rougher concentrate BPL and South concentrate BPL values!
                'This value will be the Coarse amine feed BPL.
                .AvgCrCnBpl = .CrCnBplRpt

                'Coarse amine ratio of concentration
                If .AvgCrCnBpl - .CaTlBplAdj <> 0 Then
                    .CaRc = Round((.PrdCnBpl - .CaTlBplAdj) / _
                            (.AvgCrCnBpl - .CaTlBplAdj), 2)
                Else
                    .CaRc = 0
                End If

                'Correct for if the coarse amine circuit did not run!
                If .CaPct = 0 Then
                    .CaRc = 0
                End If

                'Get the concentrate tons, tail tons, and tail BPL tons
                'Get the concentrate tons, tail tons, and tail BPL tons
                'Get the concentrate tons, tail tons, and tail BPL tons

                'NOTE: During this Tail BPL balancing procedure the rougher
                '      feed tons (both fine and coarse) stay the same!

                'Fine Roughers and Fine Amine
                'North fine rougher
                'North fine rougher
                'North fine rougher

                .NfrFdTonsRpt = .Nfr1FdTonsRpt + .Nfr2FdTonsRpt
                .NfrFdTonsAdj = .Nfr1FdTonsRpt + .Nfr2FdTonsRpt

                'North fine rougher -- Rougher concentrate tons
                If .NfrRc <> 0 Then
                    If IterationCount = 1 Then
                        'If this is the 1st iteration then will save as
                        '"reported" concentrate tons.
                        .NfrCnTonsRpt = Round(.NfrFdTonsRpt / .NfrRc, 0)
                    End If
                    .NfrCnTonsAdj = Round(.NfrFdTonsAdj / .NfrRc, 0)
                Else
                    If IterationCount = 1 Then
                        .NfrCnTonsRpt = 0
                    End If
                    .NfrCnTonsAdj = 0
                End If

                'North fine rougher -- Tail tons
                If IterationCount = 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" tail tons.
                    .NfrTlTonsRpt = .NfrFdTonsRpt - .NfrCnTonsRpt
                End If
                .NfrTlTonsAdj = .NfrFdTonsAdj - .NfrCnTonsAdj

                'North fine rougher -- Tail BPL tons
                If IterationCount = 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" tail BPL tons.
                    .NfrTlBtRpt = Round(.NfrTlTonsRpt * .NfrTlBplRpt, 2)
                End If
                .NfrTlBtAdj = Round(.NfrTlTonsAdj * .NfrTlBplAdj, 2)

                'South fine rougher
                'South fine rougher
                'South fine rougher

                'Note: Fine rougher feed tons are not adjusted in this mass balance
                '      algorithm
                .SfrFdTonsRpt = .Sfr1FdTonsRpt + .Sfr2FdTonsRpt
                .SfrFdTonsAdj = .Sfr1FdTonsRpt + .Sfr2FdTonsRpt

                'South fine rougher -- Rougher concentrate tons
                If .SfrRc <> 0 Then
                    If IterationCount = 1 Then
                        'If this is the 1st iteration then will save as
                        '"reported" concentrate tons.
                        .SfrCnTonsRpt = Round(.SfrFdTonsRpt / .SfrRc, 0)
                    End If
                    .SfrCnTonsAdj = Round(.SfrFdTonsAdj / .SfrRc, 0)
                Else
                    If IterationCount = 1 Then
                        .SfrCnTonsRpt = 0
                    End If
                    .SfrCnTonsAdj = 0
                End If

                'South fine rougher -- Tail tons
                If IterationCount = 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" tail tons.
                    .SfrTlTonsRpt = .SfrFdTonsRpt - .SfrCnTonsRpt
                End If
                .SfrTlTonsAdj = .SfrFdTonsAdj - .SfrCnTonsAdj

                'South fine rougher -- Tail BPL tons
                If IterationCount = 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" tail BPL tons.
                    .SfrTlBtRpt = Round(.SfrTlTonsRpt * .SfrTlBplRpt, 2)
                End If
                .SfrTlBtAdj = Round(.SfrTlTonsAdj * .SfrTlBplAdj, 2)

                'Fine amine
                'Fine amine
                'Fine amine

                'The fine amine feed tons will be the North fine rougher
                'concentrate tons + South fine rougher concentrate tons.
                If IterationCount = 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" feed tons.
                    .FaFdTonsRpt = .NfrCnTonsRpt + .SfrCnTonsRpt
                End If
                .FaFdTonsAdj = .NfrCnTonsAdj + .SfrCnTonsAdj

                'Fine amine -- Concentrate tons (part of the final product
                'concentrate tons).
                If .FaRc <> 0 Then
                    If IterationCount = 1 Then
                        'If this is the 1st iteration then will save as
                        '"reported" concentrate tons.
                        .FaCnTonsRpt = Round(.FaFdTonsRpt / .FaRc, 0)
                    End If
                    .FaCnTonsAdj = Round(.FaFdTonsAdj / .FaRc, 0)
                Else
                    If IterationCount = 1 Then
                        .FaCnTonsRpt = 0
                    End If
                    .FaCnTonsAdj = 0
                End If

                'Fine amine -- Tail tons
                If IterationCount = 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" tail tons.
                    .FaTlTonsRpt = .FaFdTonsRpt - .FaCnTonsRpt
                End If
                .FaTlTonsAdj = .FaFdTonsAdj - .FaCnTonsAdj

                'Fine amine -- Tail BPL tons
                If IterationCount = 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" tail BPL tons.
                    .FaTlBtRpt = Round(.FaTlTonsRpt * .FaTlBplRpt, 2)
                End If
                .FaTlBtAdj = Round(.FaTlTonsAdj * .FaTlBplAdj, 2)

                '----------

                'Coarse Roughers and Coarse Amine
                'North coarse rougher
                'North coarse rougher
                'North coarse rougher

                'Note: Coarse rougher feed tons are not adjusted
                'Have .NcrFdTonsRpt only (No .Ncr1FdTonsRpt or .Ncr2FdTonsRpt)
                .NcrFdTonsAdj = .NcrFdTonsRpt

                'North coarse rougher -- Rougher concentrate tons
                If .NcrRc <> 0 Then
                    If IterationCount = 1 Then
                        'If this is the 1st iteration then will save as
                        '"reported" concentrate tons.
                        .NcrCnTonsRpt = Round(.NcrFdTonsRpt / .NcrRc, 0)
                    End If
                    .NcrCnTonsAdj = Round(.NcrFdTonsAdj / .NcrRc, 0)
                Else
                    .NcrCnTonsAdj = 0
                End If

                'North coarse rougher -- Tail tons
                If IterationCount = 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" tail tons.
                    .NcrTlTonsRpt = .NcrFdTonsRpt - .NcrCnTonsRpt
                End If
                .NcrTlTonsAdj = .NcrFdTonsAdj - .NcrCnTonsAdj

                'North coarse rougher -- Tail BPL tons
                If IterationCount = 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" tail BPL tons.
                    .NcrTlBtRpt = Round(.NcrTlTonsRpt * .NcrTlBplRpt, 2)
                End If
                .NcrTlBtAdj = Round(.NcrTlTonsAdj * .NcrTlBplAdj, 2)

                'South coarse rougher
                'South coarse rougher
                'South coarse rougher

                'Note: Coarse rougher feed tons are not adjusted
                'Have .ScrFdTonsRpt only (No .Scr1FdTonsRpt or .Scr2FdTonsRpt)
                .ScrFdTonsAdj = .ScrFdTonsRpt

                'South coarse rougher -- Rougher concentrate tons
                If .ScrRc <> 0 Then
                    If IterationCount = 1 Then
                        'If this is the 1st iteration then will save as
                        '"reported" concentrate tons.
                        .ScrCnTonsRpt = Round(.ScrFdTonsRpt / .ScrRc, 0)
                    End If
                    .ScrCnTonsAdj = Round(.ScrFdTonsAdj / .ScrRc, 0)
                Else
                    If IterationCount = 1 Then
                        .ScrCnTonsRpt = 0
                    End If
                    .ScrCnTonsAdj = 0
                End If

                'South fine rougher -- Tail tons
                If IterationCount = 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" tail tons.
                    .ScrTlTonsRpt = .ScrFdTonsRpt - .ScrCnTonsRpt
                End If
                .ScrTlTonsAdj = .ScrFdTonsAdj - .ScrCnTonsAdj

                'South fine rougher -- Tail BPL tons
                If IterationCount <> 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" tail BPL tons.
                    .ScrTlBtRpt = Round(.ScrTlTonsRpt * .ScrTlBplRpt, 2)
                End If
                .ScrTlBtAdj = Round(.ScrTlTonsAdj * .ScrTlBplAdj, 2)

                'Coarse amine
                'Coarse amine
                'Coarse amine

                'Coarse amine -- Feed tons
                'Coarse amine feed tons = North coarse rougher concentrate
                'tons + South coarse rougher concentrate tons
                If IterationCount = 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" feed tons.
                    .CaFdTonsRpt = .NcrCnTonsRpt + .ScrCnTonsRpt

                    'Correct for the amount that the coarse amine circuit
                    'actually ran.
                    .CaFdTonsRpt = Round(.CaPct * .CaFdTonsRpt, 0)
                End If
                .CaFdTonsAdj = .NcrCnTonsAdj + .ScrCnTonsAdj

                'Correct for the amount that the coarse amine circuit
                'actually ran.
                .CaFdTonsAdj = Round(.CaPct * .CaFdTonsAdj, 0)

                'Coarse amine -- Concentrate tons (part of final concentrate
                'product)
                If .CaRc <> 0 Then
                    If IterationCount = 1 Then
                        'If this is the 1st iteration then will save as
                        '"reported" concentrate tons.
                        .CaCnTonsRpt = Round(.CaFdTonsRpt / .CaRc, 0)

                        'Correct for the amount that the coarse amine circuit
                        'actually ran (.CaRc should be zero but just in case
                        'it is not).
                        If .CaPct = 0 Then
                            .CaCnTonsRpt = 0
                        End If
                    End If
                    .CaCnTonsAdj = Round(.CaFdTonsAdj / .CaRc, 0)

                    'Correct for the amount that the coarse amine circuit
                    'actually ran (.CaRc should be zero but just in case
                    'it is not).
                    If .CaPct = 0 Then
                        .CaCnTonsAdj = 0
                    End If
                Else
                    If IterationCount = 1 Then
                        .CaCnTonsRpt = 0
                    End If
                    .CaCnTonsAdj = 0
                End If

                'Coarse amine -- Tail tons
                If IterationCount = 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" tail tons.
                    .CaTlTonsRpt = .CaFdTonsRpt - .CaCnTonsRpt

                    'Correct for the amount that the coarse amine circuit
                    'actually ran (.CaRc should be zero but just in case
                    'it is not).
                    If .CaPct = 0 Then
                        .CaTlTonsRpt = 0
                    End If
                End If
                .CaTlTonsAdj = .CaFdTonsAdj - .CaCnTonsAdj

                'Correct for the amount that the coarse amine circuit
                'actually ran (.CaRc should be zero but just in case
                'it is not).
                If .CaPct = 0 Then
                    .CaTlTonsAdj = 0
                End If

                'Coarse amine -- Tail BPL tons
                If IterationCount = 1 Then
                    'If this is the 1st iteration then will save as
                    '"reported" tail BPL tons.
                    .CaTlBtRpt = Round(.CaTlTonsRpt * .CaTlBplRpt, 2)

                    'Correct for the amount that the coarse amine circuit
                    'actually ran.
                    If .CaPct = 0 Then
                        .CaTlBtRpt = 0
                    End If
                End If
                .CaTlBtAdj = Round(.CaTlTonsAdj * .CaTlBplAdj, 2)

                'Correct for the amount that the coarse amine circuit
                'actually ran.
                If .CaPct = 0 Then
                    .CaTlBtAdj = 0
                End If
                '----------

                'Fine rougher (North fine rougher + South fine rougher)
                'Fine rougher (North fine rougher + South fine rougher)
                'Fine rougher (North fine rougher + South fine rougher)

                'Need fine rougher as a whole (North fine rougher +
                'South fine rougher)
                .FrFdTonsRpt = .NfrFdTonsRpt + .SfrFdTonsRpt
                .FrFdTonsAdj = .NfrFdTonsAdj + .SfrFdTonsAdj

                If .NfrTlTonsAdj + .SfrTlTonsAdj <> 0 Then
                    .FrTlBplAdj = Round((.NfrTlBtAdj + .SfrTlBtAdj) / _
                              (.NfrTlTonsAdj + .SfrTlTonsAdj), 2)
                Else
                    .FrTlBplAdj = 0
                End If
                .FrTlTonsAdj = .NfrTlTonsAdj + .SfrTlTonsAdj

                'Coarse rougher (North coarse rougher + South coarse rougher)
                'Coarse rougher (North coarse rougher + South coarse rougher)
                'Coarse rougher (North coarse rougher + South coarse rougher)

                'Need Coarse rougher as a whole (North coarse rougher +
                'South coarse rougher)
                .CrFdTonsRpt = .NcrFdTonsRpt + .ScrFdTonsRpt
                .CrFdTonsAdj = .NcrFdTonsAdj + .ScrFdTonsAdj

                If .NcrTlTonsAdj + .ScrTlTonsAdj <> 0 Then
                    .CrTlBplAdj = Round((.NcrTlBtAdj + .ScrTlBtAdj) / _
                                  (.NcrTlTonsAdj + .ScrTlTonsAdj), 2)
                Else
                    .CrTlBplAdj = 0
                End If
                .CrTlTonsAdj = .NcrTlTonsAdj + .ScrTlTonsAdj

                '----------

                TlTonsBalSum = .NfrTlTonsAdj + .SfrTlTonsAdj + .FaTlTonsAdj + _
                               .NcrTlTonsAdj + .ScrTlTonsAdj + .CaTlTonsAdj

                TlBtBalSum = .NfrTlBtAdj + .SfrTlBtAdj + .FaTlBtAdj + _
                              .NcrTlBtAdj + .ScrTlBtAdj + .CaTlBtAdj

                'Recalculate the GMT BPL
                'Recalculate the GMT BPL
                'Recalculate the GMT BPL

                If TlTonsBalSum <> 0 Then
                    .GmtBplAdj = Round(TlBtBalSum / TlTonsBalSum, 2)
                Else
                    .GmtBplAdj = 0
                End If

                GmtBplDiff = .GmtBplRpt - .GmtBplAdj

                GmtBplTonDiff = Round(GmtBplDiff * TlTonsBalSum, 2)

                If Abs(GmtBplDiff) < CloseEnoughVal Then
                    CloseEnough = True
                    Exit Do
                End If

                'Make the "crude" adjustments
                'Make the "crude" adjustments
                'Make the "crude" adjustments

                'Distribute extra BPL tons appropriately

                'Let's look at something first
                'Handle Really Screwed Up Data  Handle Really Screwed Up Data
                'Handle Really Screwed Up Data  Handle Really Screwed Up Data
                'Handle Really Screwed Up Data  Handle Really Screwed Up Data

                LeaveFrAlone = True
                LeaveCrAlone = True
                LeaveFaAlone = True
                LeaveCaAlone = True

                BalDistFrPct = BalDist.FrPct
                BalDistCrPct = BalDist.CrPct
                BalDistFaPct = BalDist.FaPct
                BalDistCaPct = BalDist.CaPct

                FrTlTonsAdjChk = (.FrTlBplAdj * .FrTlTonsAdj) + _
                                 (GmtBplTonDiff * BalDist.FrPct / 100)
                CrTlTonsAdjChk = (.CrTlBplAdj * .CrTlTonsAdj) + _
                                 (GmtBplTonDiff * BalDist.CrPct / 100)
                FaTlTonsAdjChk = (.FaTlBplAdj * .FaTlTonsAdj) + _
                                 (GmtBplTonDiff * BalDist.FaPct / 100)
                CaTlTonsAdjChk = (.CaTlBplAdj * .CaTlTonsAdj) + _
                                 (GmtBplTonDiff * BalDist.CaPct / 100)

                If FrTlTonsAdjChk < 0 Then
                    LeaveFrAlone = True
                Else
                    LeaveFrAlone = False
                End If

                If CrTlTonsAdjChk < 0 Then
                    LeaveCrAlone = True
                Else
                    LeaveCrAlone = False
                End If

                If FaTlTonsAdjChk < 0 Then
                    LeaveFaAlone = True
                Else
                    LeaveFaAlone = False
                End If

                If CaTlTonsAdjChk < 0 Then
                    LeaveCaAlone = True
                Else
                    LeaveCaAlone = False
                End If

                If LeaveFrAlone = True Or LeaveCrAlone = True Or LeaveFaAlone = True Or _
                    LeaveCaAlone = True Then
                    AmtToDistribute = 0

                    If LeaveFrAlone = True Then
                        AmtToDistribute = AmtToDistribute + BalDistFrPct
                        BalDistFrPct = 0
                    End If
                    If LeaveCrAlone = True Then
                        AmtToDistribute = AmtToDistribute + BalDistCrPct
                        BalDistCrPct = 0
                    End If
                    If LeaveFaAlone = True Then
                        AmtToDistribute = AmtToDistribute + BalDistFaPct
                        BalDistFaPct = 0
                    End If
                    If LeaveCaAlone = True Then
                        AmtToDistribute = AmtToDistribute + BalDistCaPct
                        BalDistCaPct = 0
                    End If

                    'Now redistribute the percent distribution!
                    If LeaveFrAlone = False Then
                        NewPct = Round((BalDistFrPct * 100) / _
                                 (100 - AmtToDistribute), 0)
                        BalDistFrPct = BalDistFrPct + _
                                      Round((NewPct / 100 * AmtToDistribute), 0)
                    End If
                    If LeaveCrAlone = False Then
                        NewPct = Round((BalDistCrPct * 100) / _
                                 (100 - AmtToDistribute), 0)
                        BalDistCrPct = BalDistCrPct + _
                                      Round((NewPct / 100 * AmtToDistribute), 0)
                    End If
                    If LeaveFaAlone = False Then
                        NewPct = Round((BalDistFaPct * 100) / _
                                 (100 - AmtToDistribute), 0)
                        BalDistFaPct = BalDistFaPct + _
                                      Round((NewPct / 100 * AmtToDistribute), 0)
                    End If
                    If LeaveCaAlone = False Then
                        NewPct = Round((BalDistCaPct * 100) / _
                                 (100 - AmtToDistribute), 0)
                        BalDistCaPct = BalDistCaPct + _
                                      Round((NewPct / 100 * AmtToDistribute), 0)
                    End If
                End If

                'End of Handle Really Screwed Up Data  End of Handle Really Screwed Up Data
                'End of Handle Really Screwed Up Data  End of Handle Really Screwed Up Data
                'End of Handle Really Screwed Up Data  End of Handle Really Screwed Up Data

                'Fine rougher
                If .FrTlTonsAdj <> 0 Then
                    FrTlBplBal = Round(((.FrTlBplAdj * .FrTlTonsAdj) + _
                                 (GmtBplTonDiff * BalDistFrPct / 100)) / _
                                 .FrTlTonsAdj, aBplRound)
                Else
                    FrTlBplBal = 0
                End If
                FrTlBplResid = FrTlBplBal - .FrTlBplAdj

                'Coarse rougher
                If .CrTlTonsAdj <> 0 Then
                    CrTlBplBal = Round(((.CrTlBplAdj * .CrTlTonsAdj) + _
                                 (GmtBplTonDiff * BalDistCrPct / 100)) / _
                                 .CrTlTonsAdj, aBplRound)
                Else
                    CrTlBplBal = 0
                End If
                CrTlBplResid = CrTlBplBal - .CrTlBplAdj

                'Fine amine
                If .FaTlTonsAdj <> 0 Then
                    FaTlBplBal = Round(((.FaTlBplAdj * .FaTlTonsAdj) + _
                                 (GmtBplTonDiff * BalDistFaPct / 100)) / _
                                 .FaTlTonsAdj, aBplRound)
                Else
                    FaTlBplBal = 0
                End If
                FaTlBplResid = FaTlBplBal - .FaTlBplAdj

                'Coarse amine
                If .CaTlTonsAdj <> 0 Then
                    CaTlBplBal = Round(((.CaTlBplAdj * .CaTlTonsAdj) + _
                                 (GmtBplTonDiff * BalDistCaPct / 100)) / _
                                 .CaTlTonsAdj, aBplRound)
                Else
                    CaTlBplBal = 0
                End If
                CaTlBplResid = CaTlBplBal - .CaTlBplAdj

                If .NfrTlTonsAdj + .SfrTlTonsAdj + _
                   .NcrTlTonsAdj + .ScrTlTonsAdj + _
                   .FaTlTonsAdj + .CaTlTonsAdj <> 0 Then
                    TotBplCheck = Round((.FrTlTonsAdj * FrTlBplBal + _
                                  .CrTlTonsAdj * CrTlBplBal + _
                                  .FaTlTonsAdj * FaTlBplBal + _
                                  .CaTlTonsAdj * CaTlBplBal) / _
                                  (.NfrTlTonsAdj + .SfrTlTonsAdj + _
                                  .NcrTlTonsAdj + .ScrTlTonsAdj + _
                                  .FaTlTonsAdj + .CaTlTonsAdj), 2)
                Else
                    TotBplCheck = 0
                End If

                'If everything is going "OK" then this value should be zero.
                TotBplCheckDiff = TotBplCheck - .GmtBplRpt

                'Set the adjusted tail BPL values to the balanced tail BPL values
                'and continue on with the next iteration.
                '1) North fine rougher tail BPL
                '2) South fine rougher tail BPL
                '3) North coarse rougher tail BPL
                '4) South coarse rougher tail BPL
                '5) Fine amine tail BPL
                '6) Coarse amine tail BPL

                .NfrTlBplAdj = FrTlBplBal
                .SfrTlBplAdj = FrTlBplBal

                .NcrTlBplAdj = CrTlBplBal
                .ScrTlBplAdj = CrTlBplBal

                .FaTlBplAdj = FaTlBplBal
                .CaTlBplAdj = CaTlBplBal
            End With
        Loop

        'Get some data together for summing purposes
        With mMbFcTotal
            'Concentrate product tons (final concentrate product)
            .PrdCnTons = .PrdCnTons + mMbFcShift.PrdCnTons
            If mMbFcShift.PrdCnBpl <> 0 Then
                .PrdCnTonsW = .PrdCnTonsW + mMbFcShift.PrdCnTons
            End If
            .PrdCnBt = .PrdCnBt + mMbFcShift.PrdCnTons * mMbFcShift.PrdCnBpl

            'Operating hours
            'Fine rougher hours
            .Nfr1Hrs = .Nfr1Hrs + mMbFcShift.Nfr1Hrs    'North fine rougher 1
            .Nfr2Hrs = .Nfr2Hrs + mMbFcShift.Nfr2Hrs    'North fine rougher 2
            .Sfr1Hrs = .Sfr1Hrs + mMbFcShift.Sfr1Hrs    'South fine rougher 1
            .Sfr2Hrs = .Sfr2Hrs + mMbFcShift.Sfr2Hrs    'South fine rougher 2

            'Average all 4 fine rougher sections to get the fine rougher
            'operating hours
            ShiftFrHrs = Round((mMbFcShift.Nfr1Hrs + _
                               mMbFcShift.Nfr2Hrs + _
                               mMbFcShift.Sfr1Hrs + _
                               mMbFcShift.Sfr2Hrs) / 4, 2)
            .FrHrs = .FrHrs + ShiftFrHrs

            'Assume that the fine amine operating hours are the same as the
            'fine rougher operating hours
            ShiftFaHrs = Round((mMbFcShift.Nfr1Hrs + _
                               mMbFcShift.Nfr2Hrs + _
                               mMbFcShift.Sfr1Hrs + _
                               mMbFcShift.Sfr2Hrs) / 4, 2)
            .FaHrs = .FaHrs + ShiftFaHrs

            'Coarse rougher hours
            .NcrHrs = .NcrHrs + mMbFcShift.NcrHrs       'North coarse rougher
            .ScrHrs = .ScrHrs + mMbFcShift.ScrHrs       'South coarse rougher

            'Coarse rougher hrs where the coarse amine ran
            .NcrCaHrs = .NcrCaHrs + mMbFcShift.NcrCaHrs   'North coarse rougher
            .ScrCaHrs = .ScrCaHrs + mMbFcShift.ScrCaHrs   'South coarse rougher

            'Average the two coarse rougher sections to get the coarse rougher
            'operating hours
            ShiftCrHrs = Round((mMbFcShift.NcrHrs + mMbFcShift.ScrHrs) / 2, 2)
            .CrHrs = .CrHrs + ShiftCrHrs

            'Assume that the coarse amine operating hours are the same as the
            'coarse rougher operating hours
            'ShiftCaHrs = Round((mMbFcShift.NcrHrs + mMbFcShift.ScrHrs) / 2, 2)
            '.CaHrs = .CaHrs + ShiftCaHrs

            '04/05/2006, lss
            'Will no longer assume that the coarse amine operating hours are
            'the same as the coarse rougher operating hours.  There are shifts
            'where the concentrate from the coarse roughers is considered
            'final product and is not run through the coarse amine circuit.
            'I have remarked out the ShiftCaHrs and .CaHrs assignments above.
            ShiftCaHrs = Round((mMbFcShift.NcrCaHrs + mMbFcShift.ScrCaHrs) / 2, 2)
            .CaHrs = .CaHrs + ShiftCaHrs

            .NcsHrs = .NcsHrs + mMbFcShift.NcsHrs       'North coarse scalp -- Not really used here
            .ScsHrs = .ScsHrs + mMbFcShift.ScsHrs       'South coarse scalp -- Not really used here

            'Plant operating hours = (2 * fine rougher hours + coarse rougher hours) / 3
            ShiftPltHrs = Round((2 * ShiftFrHrs + ShiftCrHrs) / 3, 2)
            .TotPltHrs = .TotPltHrs + ShiftPltHrs

            'Fine rougher  Fine rougher  Fine rougher  Fine rougher  Fine rougher
            'Fine rougher  Fine rougher  Fine rougher  Fine rougher  Fine rougher
            'Fine rougher  Fine rougher  Fine rougher  Fine rougher  Fine rougher

            'Feed -- reported
            'Fine rougher feed tons reported = North fine rougher feed tons reported +
            'South fine rougher feed tons reported
            ShiftTons = mMbFcShift.NfrFdTonsRpt + mMbFcShift.SfrFdTonsRpt
            .FrFdTonsRpt = .FrFdTonsRpt + ShiftTons
            'Need the FR Fd BPL from the NFR and SFR Fd BPL's
            If mMbFcShift.NfrFdTonsRpt + mMbFcShift.SfrFdTonsRpt <> 0 Then
                FrFdBplRpt = Round((mMbFcShift.NfrFdTonsRpt * mMbFcShift.NfrFdBplRpt + _
                              mMbFcShift.SfrFdTonsRpt * mMbFcShift.SfrFdBplRpt) / _
                              (mMbFcShift.NfrFdTonsRpt + mMbFcShift.SfrFdTonsRpt), aBplRound)
            Else
                FrFdBplRpt = 0
            End If
            .FrFdBtRpt = .FrFdBtRpt + ShiftTons * FrFdBplRpt

            'Concentrate -- reported
            'Fine rougher concentrate tons reported = North fine rougher
            'concentrate tons reported + South fine rougher concentrate tons reported
            ShiftTons = mMbFcShift.NfrCnTonsRpt + mMbFcShift.SfrCnTonsRpt
            .FrCnTonsRpt = .FrCnTonsRpt + ShiftTons
            'Need the FR Cn BPL from the NFR and SFR Cn BPL's
            If mMbFcShift.NfrCnTonsRpt + mMbFcShift.SfrCnTonsRpt <> 0 Then
                FrCnBplRpt = Round((mMbFcShift.NfrCnTonsRpt * mMbFcShift.NfrCnBplRpt + _
                             mMbFcShift.SfrCnTonsRpt * mMbFcShift.SfrCnBplRpt) / _
                             (mMbFcShift.NfrCnTonsRpt + mMbFcShift.SfrCnTonsRpt), aBplRound)
            Else
                FrCnBplRpt = 0
            End If
            .FrCnBtRpt = .FrCnBtRpt + ShiftTons * FrCnBplRpt

            'Tails -- reported
            'Fine rougher tail tons reported = North fine rougher tail tons reported +
            'South fine rougher tail tons reported
            ShiftTons = mMbFcShift.NfrTlTonsRpt + mMbFcShift.SfrTlTonsRpt
            .FrTlTonsRpt = .FrTlTonsRpt + ShiftTons
            'Need the FR Tl BPL from the NFR and SFR Tl BPL's
            If mMbFcShift.NfrTlTonsRpt + mMbFcShift.SfrTlTonsRpt <> 0 Then
                FrTlBplRpt = Round((mMbFcShift.NfrTlTonsRpt * mMbFcShift.NfrTlBplRpt + _
                             mMbFcShift.SfrTlTonsRpt * mMbFcShift.SfrTlBplRpt) / _
                             (mMbFcShift.NfrTlTonsRpt + mMbFcShift.SfrTlTonsRpt), aBplRound)
            Else
                FrTlBplRpt = 0
            End If
            .FrTlBtRpt = .FrTlBtRpt + ShiftTons * FrTlBplRpt

            'Feed -- adjusted
            'Fine rougher feed tons adjusted = North fine rougher feed tons adjusted +
            'South fine rougher feed tons adjusted
            'Actually the way this mass balance reduction works the adjusted feed tons
            'are the same as the reported feed tons for the circuits.
            ShiftTons = mMbFcShift.NfrFdTonsAdj + mMbFcShift.SfrFdTonsAdj
            .FrFdTonsAdj = .FrFdTonsAdj + ShiftTons
            'Need the FR Fd BPL from the NFR and SFR Fd BPL's
            If mMbFcShift.NfrFdTonsAdj + mMbFcShift.SfrFdTonsAdj <> 0 Then
                FrFdBplAdj = Round((mMbFcShift.NfrFdTonsAdj * mMbFcShift.NfrFdBplRpt + _
                             mMbFcShift.SfrFdTonsAdj * mMbFcShift.SfrFdBplRpt) / _
                             (mMbFcShift.NfrFdTonsAdj + mMbFcShift.SfrFdTonsAdj), aBplRound)
            Else
                FrFdBplAdj = 0
            End If
            .FrFdBtAdj = .FrFdBtAdj + ShiftTons * FrFdBplAdj

            'Concentrate -- adjusted
            'Fine rougher concentrate tons adjusted = North fine rougher
            'concentrate tons adjusted + South fine rougher concentrate tons adjusted
            ShiftTons = mMbFcShift.NfrCnTonsAdj + mMbFcShift.SfrCnTonsAdj
            .FrCnTonsAdj = .FrCnTonsAdj + ShiftTons
            'Need the FR Cn BPL from the NFR and SFR Cn BPL's
            If mMbFcShift.NfrCnTonsAdj + mMbFcShift.SfrCnTonsAdj <> 0 Then
                FrCnBplAdj = Round((mMbFcShift.NfrCnTonsAdj * mMbFcShift.NfrCnBplRpt + _
                             mMbFcShift.SfrCnTonsAdj * mMbFcShift.SfrCnBplRpt) / _
                             (mMbFcShift.NfrCnTonsAdj + mMbFcShift.SfrCnTonsAdj), aBplRound)
            Else
                FrCnBplAdj = 0
            End If
            .FrCnBtAdj = .FrCnBtAdj + ShiftTons * FrCnBplAdj

            'Tails -- adjusted
            'Fine rougher tail tons adjusted = North fine rougher tail tons adjusted +
            'South fine rougher tail tons adjusted
            ShiftTons = mMbFcShift.NfrTlTonsAdj + mMbFcShift.SfrTlTonsAdj
            .FrTlTonsAdj = .FrTlTonsAdj + ShiftTons
            'Need the FR Tl BPL from the NFR and SFR Tl BPL's
            If mMbFcShift.NfrTlTonsAdj + mMbFcShift.SfrTlTonsAdj <> 0 Then
                FrTlBplAdj = Round((mMbFcShift.NfrTlTonsAdj * mMbFcShift.NfrTlBplAdj + _
                             mMbFcShift.SfrTlTonsAdj * mMbFcShift.SfrTlBplAdj) / _
                             (mMbFcShift.NfrTlTonsAdj + mMbFcShift.SfrTlTonsAdj), aBplRound)
            Else
                FrTlBplAdj = 0
            End If
            .FrTlBtAdj = .FrTlBtAdj + ShiftTons * FrTlBplAdj

            'Coarse rougher  Coarse rougher  Coarse rougher  Coarse rougher
            'Coarse rougher  Coarse rougher  Coarse rougher  Coarse rougher
            'Coarse rougher  Coarse rougher  Coarse rougher  Coarse rougher

            'Feed -- reported
            'Coarse rougher feed tons reported = North coarse rougher feed tons reported +
            'South coarse rougher feed tons reported
            ShiftTons = mMbFcShift.NcrFdTonsRpt + mMbFcShift.ScrFdTonsRpt
            .CrFdTonsRpt = .CrFdTonsRpt + ShiftTons
            'Need the CR Fd BPL from the NCR and SCR Fd BPL's
            If mMbFcShift.NcrFdTonsRpt + mMbFcShift.ScrFdTonsRpt <> 0 Then
                CrFdBplRpt = Round((mMbFcShift.NcrFdTonsRpt * mMbFcShift.NcrFdBplRpt + _
                             mMbFcShift.ScrFdTonsRpt * mMbFcShift.ScrFdBplRpt) / _
                             (mMbFcShift.NcrFdTonsRpt + mMbFcShift.ScrFdTonsRpt), aBplRound)
            Else
                CrFdBplRpt = 0
            End If
            .CrFdBtRpt = .CrFdBtRpt + ShiftTons * CrFdBplRpt

            'Concentrate -- reported
            'Coarse rougher concentrate tons reported = North coarse rougher
            'concentrate tons reported + South coarse rougher concentrate tons reported
            ShiftTons = mMbFcShift.NcrCnTonsRpt + mMbFcShift.ScrCnTonsRpt
            .CrCnTonsRpt = .CrCnTonsRpt + ShiftTons
            'Need the CR Cn BPL from the NCR and SCR Cn BPL's
            If mMbFcShift.NcrCnTonsRpt + mMbFcShift.ScrCnTonsRpt <> 0 Then
                CrCnBplRpt = Round((mMbFcShift.NcrCnTonsRpt * mMbFcShift.CrCnBplRpt + _
                             mMbFcShift.ScrCnTonsRpt * mMbFcShift.CrCnBplRpt) / _
                             (mMbFcShift.NcrCnTonsRpt + mMbFcShift.ScrCnTonsRpt), aBplRound)
            Else
                CrCnBplRpt = 0
            End If
            .CrCnBtRpt = .CrCnBtRpt + ShiftTons * CrCnBplRpt

            'Tails -- reported
            'Coarse rougher tail tons reported = North coarse rougher tail tons reported +
            'South coarse rougher tail tons reported
            ShiftTons = mMbFcShift.NcrTlTonsRpt + mMbFcShift.ScrTlTonsRpt
            .CrTlTonsRpt = .CrTlTonsRpt + ShiftTons
            'Need the CR Tl BPL from the NCR and SCR Tl BPL's
            If mMbFcShift.NcrTlTonsRpt + mMbFcShift.ScrTlTonsRpt <> 0 Then
                CrTlBplRpt = Round((mMbFcShift.NcrTlTonsRpt * mMbFcShift.NcrTlBplRpt + _
                             mMbFcShift.ScrTlTonsRpt * mMbFcShift.ScrTlBplRpt) / _
                             (mMbFcShift.NcrTlTonsRpt + mMbFcShift.ScrTlTonsRpt), aBplRound)
            Else
                CrTlBplRpt = 0
            End If
            .CrTlBtRpt = .CrTlBtRpt + ShiftTons * CrTlBplRpt

            'Feed -- adjusted
            'Coarse rougher feed tons adjusted = Coarse fine rougher feed tons adjusted +
            'Coarse fine rougher feed tons adjusted
            ShiftTons = mMbFcShift.NcrFdTonsAdj + mMbFcShift.ScrFdTonsAdj
            .CrFdTonsAdj = .CrFdTonsAdj + ShiftTons
            'Need the CR Fd BPL from the NCR and SCR Fd BPL's
            If mMbFcShift.NcrFdTonsAdj + mMbFcShift.ScrFdTonsAdj <> 0 Then
                CrFdBplAdj = Round((mMbFcShift.NcrFdTonsAdj * mMbFcShift.NcrFdBplRpt + _
                             mMbFcShift.ScrFdTonsAdj * mMbFcShift.ScrFdBplRpt) / _
                             (mMbFcShift.NcrFdTonsAdj + mMbFcShift.ScrFdTonsAdj), aBplRound)
            Else
                CrFdBplAdj = 0
            End If
            .CrFdBtAdj = .CrFdBtAdj + ShiftTons * CrFdBplAdj

            'Concentrate -- adjusted
            'Coarse rougher concentrate tons adjusted = North coarse rougher
            'concentrate tons adjusted + South coarse rougher concentrate tons adjusted
            ShiftTons = mMbFcShift.NcrCnTonsAdj + mMbFcShift.ScrCnTonsAdj
            .CrCnTonsAdj = .CrCnTonsAdj + ShiftTons
            'Need the CR Cn BPL from the NCR and SCR Cn BPL's
            If mMbFcShift.NcrCnTonsAdj + mMbFcShift.ScrCnTonsAdj <> 0 Then
                CrCnBplAdj = Round((mMbFcShift.NcrCnTonsAdj * mMbFcShift.CrCnBplRpt + _
                             mMbFcShift.ScrCnTonsAdj * mMbFcShift.CrCnBplRpt) / _
                             (mMbFcShift.NcrCnTonsAdj + mMbFcShift.ScrCnTonsAdj), aBplRound)
            Else
                CrCnBplAdj = 0
            End If
            .CrCnBtAdj = .CrCnBtAdj + ShiftTons * CrCnBplAdj

            'Tails -- adjusted
            'Coarse rougher tail tons adjusted = North coarse rougher tail tons adjusted +
            'South coarse rougher tail tons adjusted
            ShiftTons = mMbFcShift.NcrTlTonsAdj + mMbFcShift.ScrTlTonsAdj
            .CrTlTonsAdj = .CrTlTonsAdj + ShiftTons
            'Need the CR Tl BPL from the NCR and SCR Tl BPL's
            If mMbFcShift.NcrTlTonsAdj + mMbFcShift.ScrTlTonsAdj <> 0 Then
                CrTlBplAdj = Round((mMbFcShift.NcrTlTonsAdj * mMbFcShift.NcrTlBplAdj + _
                             mMbFcShift.ScrTlTonsAdj * mMbFcShift.ScrTlBplAdj) / _
                             (mMbFcShift.NcrTlTonsAdj + mMbFcShift.ScrTlTonsAdj), aBplRound)
            Else
                CrTlBplAdj = 0
            End If
            .CrTlBtAdj = .CrTlBtAdj + ShiftTons * CrTlBplAdj

            'Fine amine  Fine amine  Fine amine  Fine amine  Fine amine
            'Fine amine  Fine amine  Fine amine  Fine amine  Fine amine
            'Fine amine  Fine amine  Fine amine  Fine amine  Fine amine

            'Feed -- reported
            .FaFdTonsRpt = .FaFdTonsRpt + mMbFcShift.FaFdTonsRpt
            'The reported FA Fd BPL is the average FR Cn BPL
            .FaFdBtRpt = .FaFdBtRpt + mMbFcShift.FaFdTonsRpt * mMbFcShift.AvgFrCnBpl

            'Concentrate -- reported
            .FaCnTonsRpt = .FaCnTonsRpt + mMbFcShift.FaCnTonsRpt
            'The reported FA Cn BPL is the final Cn product BPL
            .FaCnBtRpt = .FaCnBtRpt + mMbFcShift.FaCnTonsRpt * mMbFcShift.PrdCnBpl


            'Tails -- reported
            .FaTlTonsRpt = .FaTlTonsRpt + mMbFcShift.FaTlTonsRpt
            .FaTlBtRpt = .FaTlBtRpt + mMbFcShift.FaTlTonsRpt * mMbFcShift.FaTlBplRpt

            'Feed -- adjusted
            .FaFdTonsAdj = .FaFdTonsAdj + mMbFcShift.FaFdTonsAdj
            .FaFdBtAdj = .FaFdBtAdj + mMbFcShift.FaFdTonsAdj * mMbFcShift.AvgFrCnBpl

            'Concentrate -- adjusted
            .FaCnTonsAdj = .FaCnTonsAdj + mMbFcShift.FaCnTonsAdj
            .FaCnBtAdj = .FaCnBtAdj + mMbFcShift.FaCnTonsAdj * mMbFcShift.PrdCnBpl

            'Tails -- adjusted
            .FaTlTonsAdj = .FaTlTonsAdj + mMbFcShift.FaTlTonsAdj
            .FaTlBtAdj = .FaTlBtAdj + mMbFcShift.FaTlTonsAdj * mMbFcShift.FaTlBplAdj

            'Coarse amine  Coarse amine  Coarse amine  Coarse amine
            'Coarse amine  Coarse amine  Coarse amine  Coarse amine
            'Coarse amine  Coarse amine  Coarse amine  Coarse amine

            'Feed -- reported
            .CaFdTonsRpt = .CaFdTonsRpt + mMbFcShift.CaFdTonsRpt
            .CaFdBtRpt = .CaFdBtRpt + mMbFcShift.CaFdTonsRpt * mMbFcShift.AvgCrCnBpl

            'Concentrate -- reported
            .CaCnTonsRpt = .CaCnTonsRpt + mMbFcShift.CaCnTonsRpt
            .CaCnBtRpt = .CaCnBtRpt + mMbFcShift.CaCnTonsRpt * mMbFcShift.PrdCnBpl

            'Tails -- reported
            .CaTlTonsRpt = .CaTlTonsRpt + mMbFcShift.CaTlTonsRpt
            .CaTlBtRpt = .CaTlBtRpt + mMbFcShift.CaTlTonsRpt * mMbFcShift.CaTlBplRpt

            'Feed -- adjusted
            .CaFdTonsAdj = .CaFdTonsAdj + mMbFcShift.CaFdTonsAdj
            .CaFdBtAdj = .CaFdBtAdj + mMbFcShift.CaFdTonsAdj * mMbFcShift.AvgCrCnBpl

            'Concentrate -- adjusted
            .CaCnTonsAdj = .CaCnTonsAdj + mMbFcShift.CaCnTonsAdj
            .CaCnBtAdj = .CaCnBtAdj + mMbFcShift.CaCnTonsAdj * mMbFcShift.PrdCnBpl

            'Tails -- adjusted
            .CaTlTonsAdj = .CaTlTonsAdj + mMbFcShift.CaTlTonsAdj
            .CaTlBtAdj = .CaTlBtAdj + mMbFcShift.CaTlTonsAdj * mMbFcShift.CaTlBplAdj

            'Miscellaneous  Miscellaneous  Miscellaneous
            'Miscellaneous  Miscellaneous  Miscellaneous
            'Miscellaneous  Miscellaneous  Miscellaneous
            .TotPltFdTonsRpt = .TotPltFdTonsRpt + _
                               mMbFcShift.FrFdTonsRpt + mMbFcShift.CrFdTonsRpt

            TotPltTlTonsRpt = mMbFcShift.NfrTlTonsRpt + mMbFcShift.SfrTlTonsRpt + _
                              mMbFcShift.FaTlTonsRpt + _
                              mMbFcShift.NcrTlTonsRpt + mMbFcShift.ScrTlTonsRpt + _
                              mMbFcShift.CaTlTonsRpt

            .TotPltTlTonsRpt = .TotPltTlTonsRpt + TotPltTlTonsRpt

            .TotPltGmtTlBtRpt = .TotPltGmtTlBtRpt + TotPltTlTonsRpt * _
                                mMbFcShift.GmtBplRpt

            '----------
            .BalDistFrPct = BalDist.FrPct
            .BalDistCrPct = BalDist.CrPct
            .BalDistFaPct = BalDist.FaPct
            .BalDistCaPct = BalDist.CaPct
        End With

        Exit Sub

ProcessFcMassBalanceDataError:

        MsgBox("Error in Four Corners mass balance." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Four Corners Mass Balance Computation Error")
    End Sub

    Private Sub ProcessFcMassBalanceTotals(ByVal aBplRound As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo ProcessFcMassBalanceTotalsError

        Dim AvgFrHrs As Double
        Dim AvgCrHrs As Double

        With mMbFcTotal
            'Product analysis -- final concentrate
            If .PrdCnTonsW <> 0 Then
                .PrdCnBpl = Round(.PrdCnBt / .PrdCnTonsW, 2)
            Else
                .PrdCnBpl = 0
            End If

            'TPH's  TPH's  TPH's  TPH's  TPH's  TPH's  TPH's
            'TPH's  TPH's  TPH's  TPH's  TPH's  TPH's  TPH's
            'TPH's  TPH's  TPH's  TPH's  TPH's  TPH's  TPH's

            AvgFrHrs = Round((.Nfr1Hrs + .Nfr2Hrs + .Sfr1Hrs + .Sfr2Hrs) / 4, 2)
            AvgCrHrs = Round((.NcrHrs + .ScrHrs) / 4, 2)

            'Could use the calculated values AvgFrHrs and AvgCrHrs here but
            'since we have been summing Fr hours and Cr hours we will use that
            'instead (.FrHrs, .FaHrs, .CrHrs, and .CaHrs already have the values
            'that we want).

            'Reported feed TPH's for the circuits -- based on reported tons
            '1) Fine rougher reported feed TPH
            '2) Fine amine reported feed TPH
            '3) Coarse rougher reported feed TPH
            '4) Coarse amine reported feed TPH
            If .FrHrs <> 0 Then
                .FrFdTphRpt = Round(.FrFdTonsRpt / .FrHrs, 0)
            Else
                .FrFdTphRpt = 0
            End If
            If .FaHrs <> 0 Then
                .FaFdTphRpt = Round(.FaFdTonsRpt / .FaHrs, 0)
            Else
                .FaFdTphRpt = 0
            End If
            If .CrHrs <> 0 Then
                .CrFdTphRpt = Round(.CrFdTonsRpt / .CrHrs, 0)
            Else
                .CrFdTphRpt = 0
            End If
            If .CaHrs <> 0 Then
                .CaFdTphRpt = Round(.CaFdTonsRpt / .CaHrs, 0)
            Else
                .CaFdTphRpt = 0
            End If

            'Reported concentrate TPH's for the circuits -- based on reported tons
            '1) Fine rougher reported concentrate TPH
            '2) Fine amine reported concentrate TPH
            '3) Coarse rougher reported concentrate TPH
            '4) Coarse amine reported concentrate TPH
            If .FrHrs <> 0 Then
                .FrCnTphRpt = Round(.FrCnTonsRpt / .FrHrs, 0)
            Else
                .FrCnTphRpt = 0
            End If
            If .FaHrs <> 0 Then
                .FaCnTphRpt = Round(.FaCnTonsRpt / .FaHrs, 0)
            Else
                .FaCnTphRpt = 0
            End If
            If .CrHrs <> 0 Then
                .CrCnTphRpt = Round(.CrCnTonsRpt / .CrHrs, 0)
            Else
                .CrCnTphRpt = 0
            End If
            If .CaHrs <> 0 Then
                .CaCnTphRpt = Round(.CaCnTonsRpt / .CaHrs, 0)
            Else
                .CaCnTphRpt = 0
            End If

            'Reported tails TPH's for the circuits -- based on reported tons
            '1) Fine rougher reported tail TPH
            '2) Fine amine reported tail TPH
            '3) Coarse rougher reported tail TPH
            '4) Coarse amine reported tail TPH
            If .FrHrs <> 0 Then
                .FrTlTphRpt = Round(.FrTlTonsRpt / .FrHrs, 0)
            Else
                .FrTlTphRpt = 0
            End If
            If .FaHrs <> 0 Then
                .FaTlTphRpt = Round(.FaTlTonsRpt / .FaHrs, 0)
            Else
                .FaTlTphRpt = 0
            End If
            If .CrHrs <> 0 Then
                .CrTlTphRpt = Round(.CrTlTonsRpt / .CrHrs, 0)
            Else
                .CrTlTphRpt = 0
            End If
            If .CaHrs <> 0 Then
                .CaTlTphRpt = Round(.CaTlTonsRpt / .CaHrs, 0)
            Else
                .CaTlTphRpt = 0
            End If

            'Reported feed TPH's for the circuits -- based on reported tons
            '1) Fine rougher reported feed TPH
            '2) Fine amine reported feed TPH
            '3) Coarse rougher reported feed TPH
            '4) Coarse amine reported feed TPH
            If .FrHrs <> 0 Then
                .FrFdTphAdj = Round(.FrFdTonsAdj / .FrHrs, 0)
            Else
                .FrFdTphAdj = 0
            End If
            If .FaHrs <> 0 Then
                .FaFdTphAdj = Round(.FaFdTonsAdj / .FaHrs, 0)
            Else
                .FaFdTphAdj = 0
            End If
            If .CrHrs <> 0 Then
                .CrFdTphAdj = Round(.CrFdTonsAdj / .CrHrs, 0)
            Else
                .CrFdTphAdj = 0
            End If
            If .CaHrs <> 0 Then
                .CaFdTphAdj = Round(.CaFdTonsAdj / .CaHrs, 0)
            Else
                .CaFdTphAdj = 0
            End If

            'Reported concentrate TPH's for the circuits -- based on reported tons
            '1) Fine rougher reported concentrate TPH
            '2) Fine amine reported concentrate TPH
            '3) Coarse rougher reported concentrate TPH
            '4) Coarse amine reported concentrate TPH
            If .FrHrs <> 0 Then
                .FrCnTphAdj = Round(.FrCnTonsAdj / .FrHrs, 0)
            Else
                .FrCnTphAdj = 0
            End If
            If .FaHrs <> 0 Then
                .FaCnTphAdj = Round(.FaCnTonsAdj / .FaHrs, 0)
            Else
                .FaCnTphAdj = 0
            End If
            If .CrHrs <> 0 Then
                .CrCnTphAdj = Round(.CrCnTonsAdj / .CrHrs, 0)
            Else
                .CrCnTphAdj = 0
            End If
            If .CaHrs <> 0 Then
                .CaCnTphAdj = Round(.CaCnTonsAdj / .CaHrs, 0)
            Else
                .CaCnTphAdj = 0
            End If

            'Reported tails TPH's for the circuits -- based on reported tons
            '1) Fine rougher reported tail TPH
            '2) Fine amine reported tail TPH
            '3) Coarse rougher reported tail TPH
            '4) Coarse amine reported tail TPH
            If .FrHrs <> 0 Then
                .FrTlTphAdj = Round(.FrTlTonsAdj / .FrHrs, 0)
            Else
                .FrTlTphAdj = 0
            End If
            If .FaHrs <> 0 Then
                .FaTlTphAdj = Round(.FaTlTonsAdj / .FaHrs, 0)
            Else
                .FaTlTphAdj = 0
            End If
            If .CrHrs <> 0 Then
                .CrTlTphAdj = Round(.CrTlTonsAdj / .CrHrs, 0)
            Else
                .CrTlTphAdj = 0
            End If
            If .CaHrs <> 0 Then
                .CaTlTphAdj = Round(.CaTlTonsAdj / .CaHrs, 0)
            Else
                .CaTlTphAdj = 0
            End If

            'BPL's  BPL's  BPL's  BPL's  BPL's  BPL's  BPL's
            'BPL's  BPL's  BPL's  BPL's  BPL's  BPL's  BPL's
            'BPL's  BPL's  BPL's  BPL's  BPL's  BPL's  BPL's

            'Adjusted  Adjusted  Adjusted  Adjusted  Adjusted

            'Adjusted Fine rougher BPL's
            If .FrFdTonsAdj <> 0 Then
                .FrFdBplAdj = Round(.FrFdBtAdj / .FrFdTonsAdj, aBplRound)
            Else
                .FrFdBplAdj = 0
            End If
            If .FrCnTonsAdj <> 0 Then
                .FrCnBplAdj = Round(.FrCnBtAdj / .FrCnTonsAdj, aBplRound)
            Else
                .FrCnBplAdj = 0
            End If
            If .FrTlTonsAdj <> 0 Then
                .FrTlBplAdj = Round(.FrTlBtAdj / .FrTlTonsAdj, aBplRound)
            Else
                .FrTlBplAdj = 0
            End If

            'Adjusted Fine amine BPL's
            If .FaFdTonsAdj <> 0 Then
                .FaFdBplAdj = Round(.FaFdBtAdj / .FaFdTonsAdj, aBplRound)
            Else
                .FaFdBplAdj = 0
            End If
            If .FaCnTonsAdj <> 0 Then
                .FaCnBplAdj = Round(.FaCnBtAdj / .FaCnTonsAdj, aBplRound)
            Else
                .FaCnBplAdj = 0
            End If
            If .FaTlTonsAdj <> 0 Then
                .FaTlBplAdj = Round(.FaTlBtAdj / .FaTlTonsAdj, aBplRound)
            Else
                .FaTlBplAdj = 0
            End If

            'Adjusted Coarse rougher BPL's
            If .CrFdTonsAdj <> 0 Then
                .CrFdBplAdj = Round(.CrFdBtAdj / .CrFdTonsAdj, aBplRound)
            Else
                .CrFdBplAdj = 0
            End If
            If .CrCnTonsAdj <> 0 Then
                .CrCnBplAdj = Round(.CrCnBtAdj / .CrCnTonsAdj, aBplRound)
            Else
                .CrCnBplAdj = 0
            End If
            If .CrTlTonsAdj <> 0 Then
                .CrTlBplAdj = Round(.CrTlBtAdj / .CrTlTonsAdj, aBplRound)
            Else
                .CrTlBplAdj = 0
            End If

            'Adjusted Coarse amine BPL's
            If .CaFdTonsAdj <> 0 Then
                .CaFdBplAdj = Round(.CaFdBtAdj / .CaFdTonsAdj, aBplRound)
            Else
                .CaFdBplAdj = 0
            End If
            If .CaCnTonsAdj <> 0 Then
                .CaCnBplAdj = Round(.CaCnBtAdj / .CaCnTonsAdj, aBplRound)
            Else
                .CaCnBplAdj = 0
            End If
            If .CaTlTonsAdj <> 0 Then
                .CaTlBplAdj = Round(.CaTlBtAdj / .CaTlTonsAdj, aBplRound)
            Else
                .CaTlBplAdj = 0
            End If

            'Reported  Reported  Reported  Reported  Reported

            'Reported Fine rougher BPL's
            If .FrFdTonsRpt <> 0 Then
                .FrFdBplRpt = Round(.FrFdBtRpt / .FrFdTonsRpt, aBplRound)
            Else
                .FrFdBplRpt = 0
            End If
            If .FrCnTonsRpt <> 0 Then
                .FrCnBplRpt = Round(.FrCnBtRpt / .FrCnTonsRpt, aBplRound)
            Else
                .FrCnBplRpt = 0
            End If
            If .FrTlTonsRpt <> 0 Then
                .FrTlBplRpt = Round(.FrTlBtRpt / .FrTlTonsRpt, aBplRound)
            Else
                .FrTlBplRpt = 0
            End If

            'Reported Fine amine BPL's
            If .FaFdTonsRpt <> 0 Then
                .FaFdBplRpt = Round(.FaFdBtRpt / .FaFdTonsRpt, aBplRound)
            Else
                .FaFdBplRpt = 0
            End If
            If .FaCnTonsRpt <> 0 Then
                .FaCnBplRpt = Round(.FaCnBtRpt / .FaCnTonsRpt, aBplRound)
            Else
                .FaCnBplRpt = 0
            End If
            If .FaTlTonsRpt <> 0 Then
                .FaTlBplRpt = Round(.FaTlBtRpt / .FaTlTonsRpt, aBplRound)
            Else
                .FaTlBplRpt = 0
            End If

            'Reported Coarse rougher BPL's
            If .CrFdTonsRpt <> 0 Then
                .CrFdBplRpt = Round(.CrFdBtRpt / .CrFdTonsRpt, aBplRound)
            Else
                .CrFdBplRpt = 0
            End If
            If .CrCnTonsRpt <> 0 Then
                .CrCnBplRpt = Round(.CrCnBtRpt / .CrCnTonsRpt, aBplRound)
            Else
                .CrCnBplRpt = 0
            End If
            If .CrTlTonsRpt <> 0 Then
                .CrTlBplRpt = Round(.CrTlBtRpt / .CrTlTonsRpt, aBplRound)
            Else
                .CrTlBplRpt = 0
            End If

            'Reported Coarse amine BPL's
            If .CaFdTonsRpt <> 0 Then
                .CaFdBplRpt = Round(.CaFdBtRpt / .CaFdTonsRpt, aBplRound)
            Else
                .CaFdBplRpt = 0
            End If
            If .CaCnTonsRpt <> 0 Then
                .CaCnBplRpt = Round(.CaCnBtRpt / .CaCnTonsRpt, aBplRound)
            Else
                .CaCnBplRpt = 0
            End If
            If .CaTlTonsRpt <> 0 Then
                .CaTlBplRpt = Round(.CaTlBtRpt / .CaTlTonsRpt, aBplRound)
            Else
                .CaTlBplRpt = 0
            End If

            'Total plant  Total plant  Total plant  Total plant

            'Adjusted  Adjusted  Adjusted  Adjusted  Adjusted

            'Total plant Fd BPL adjusted
            If .FrFdTonsAdj + .CrFdTonsAdj <> 0 Then
                .TotPltFdBplAdj = Round((.FrFdBtAdj + .CrFdBtAdj) / _
                                  (.FrFdTonsAdj + .CrFdTonsAdj), aBplRound)
            Else
                .TotPltFdBplAdj = 0
            End If

            'Total plant Cn BPL adjusted
            .TotPltCnBplAdj = .PrdCnBpl

            'Total plant Tl BPL adjusted
            If .FrTlTonsAdj + .CrTlTonsAdj + .FaTlTonsAdj + .CaTlTonsAdj <> 0 Then
                .TotPltTlBplAdj = Round((.FrTlBtAdj + .CrTlBtAdj + .FaTlBtAdj + .CaTlBtAdj) / _
                                  (.FrTlTonsAdj + .CrTlTonsAdj + .FaTlTonsAdj + .CaTlTonsAdj), aBplRound)
            Else
                .TotPltTlBplAdj = 0
            End If

            'Total plant  Total plant  Total plant  Total plant

            'Reported  Reported  Reported  Reported  Reported

            'Total plant Fd BPL reported
            If .FrFdTonsRpt + .CrFdTonsRpt <> 0 Then
                .TotPltFdBplRpt = Round((.FrFdBtRpt + .CrFdBtRpt) / _
                                  (.FrFdTonsRpt + .CrFdTonsRpt), aBplRound)
            Else
                .TotPltFdBplRpt = 0
            End If

            'Total plant Cn BPL reported
            .TotPltCnBplRpt = .PrdCnBpl

            'Total plant Tl BPL reported
            If .FrTlTonsRpt + .CrTlTonsRpt + .FaTlTonsRpt + .CaTlTonsRpt <> 0 Then
                .TotPltTlBplRpt = Round((.FrTlBtRpt + .CrTlBtRpt + .FaTlBtRpt + .CaTlBtRpt) / _
                                  (.FrTlTonsRpt + .CrTlTonsRpt + .FaTlTonsRpt + .CaTlTonsRpt), aBplRound)
            Else
                .TotPltTlBplRpt = 0
            End If

            'Total plant Tl BPL reported -- using GMT
            If .TotPltTlTonsRpt <> 0 Then
                .TotPltTlBplRpt2 = Round(.TotPltGmtTlBtRpt / .TotPltTlTonsRpt, aBplRound)
            Else
                .TotPltTlBplRpt2 = 0
            End If

            'Ratio of concentrations  Ratio of concentrations
            'Ratio of concentrations  Ratio of concentrations
            'Ratio of concentrations  Ratio of concentrations

            'Fine rougher ratio of concentration adjusted
            If .FrFdBplAdj - .FrTlBplAdj <> 0 Then
                .FrRcAdj = Round((.FrCnBplAdj - .FrTlBplAdj) / _
                           (.FrFdBplAdj - .FrTlBplAdj), 2)
            Else
                .FrRcAdj = 0
            End If

            'Fine amine ratio of concentration adjusted
            If .FaFdBplAdj - .FaTlBplAdj <> 0 Then
                .FaRcAdj = Round((.FaCnBplAdj - .FaTlBplAdj) / _
                           (.FaFdBplAdj - .FaTlBplAdj), 2)
            Else
                .FaRcAdj = 0
            End If

            'Coarse rougher ratio of concentration adjusted
            If .CrFdBplAdj - .CrTlBplAdj <> 0 Then
                .CrRcAdj = Round((.CrCnBplAdj - .CrTlBplAdj) / _
                           (.CrFdBplAdj - .CrTlBplAdj), 2)
            Else
                .CrRcAdj = 0
            End If

            'Coarse amine ratio of concentration adjusted
            If .CaFdBplAdj - .CaTlBplAdj <> 0 Then
                .CaRcAdj = Round((.CaCnBplAdj - .CaTlBplAdj) / _
                           (.CaFdBplAdj - .CaTlBplAdj), 2)
            Else
                .CaRcAdj = 0
            End If

            'Recoveries  Recoveries  Recoveries  Recoveries
            'Recoveries  Recoveries  Recoveries  Recoveries
            'Recoveries  Recoveries  Recoveries  Recoveries

            'Fine rougher adjusted recovery  (Actual recovery)
            If .FrCnBplAdj - .FrTlBplAdj <> 0 And .FrFdBplAdj <> 0 Then
                .FrPctAdjRcvry = Round(100 * (.FrCnBplAdj / .FrFdBplAdj) * _
                                 (.FrFdBplAdj - .FrTlBplAdj) / _
                                 (.FrCnBplAdj - .FrTlBplAdj), 1)
            Else
                .FrPctAdjRcvry = 0
            End If

            'Fine rougher reported recovery
            If .FrCnBplRpt - .FrTlBplRpt <> 0 And .FrFdBplRpt <> 0 Then
                .FrPctRptRcvry = Round(100 * (.FrCnBplRpt / .FrFdBplRpt) * _
                                 (.FrFdBplRpt - .FrTlBplRpt) / _
                                 (.FrCnBplRpt - .FrTlBplRpt), 1)
            Else
                .FrPctRptRcvry = 0
            End If

            'Fine amine adjusted recovery  (Actual recovery)
            If .FaCnBplAdj - .FaTlBplAdj <> 0 And .FaFdBplAdj <> 0 Then
                .FaPctAdjRcvry = Round(100 * (.FaCnBplAdj / .FaFdBplAdj) * _
                                 (.FaFdBplAdj - .FaTlBplAdj) / _
                                 (.FaCnBplAdj - .FaTlBplAdj), 1)
            Else
                .FaPctAdjRcvry = 0
            End If

            'Fine amine reported recovery
            If .FaCnBplRpt - .FaTlBplRpt <> 0 And .FaFdBplRpt <> 0 Then
                .FaPctRptRcvry = Round(100 * (.FaCnBplRpt / .FaFdBplRpt) * _
                                 (.FaFdBplRpt - .FaTlBplRpt) / _
                                 (.FaCnBplRpt - .FaTlBplRpt), 1)
            Else
                .FaPctRptRcvry = 0
            End If

            'Coarse rougher adjusted recovery  (Actual recovery)
            If .CrCnBplAdj - .CrTlBplAdj <> 0 And .CrFdBplAdj <> 0 Then
                .CrPctAdjRcvry = Round(100 * (.CrCnBplAdj / .CrFdBplAdj) * _
                                 (.CrFdBplAdj - .CrTlBplAdj) / _
                                 (.CrCnBplAdj - .CrTlBplAdj), 1)
            Else
                .CrPctAdjRcvry = 0
            End If

            If .CrCnBplRpt - .CrTlBplRpt <> 0 And .CrFdBplRpt <> 0 Then
                .CrPctRptRcvry = Round(100 * (.CrCnBplRpt / .CrFdBplRpt) * _
                                 (.CrFdBplRpt - .CrTlBplRpt) / _
                                 (.CrCnBplRpt - .CrTlBplRpt), 1)
            Else
                .CrPctRptRcvry = 0
            End If

            'Coarse amine adjusted recovery  (Actual recovery)
            If .CaCnBplAdj - .CaTlBplAdj <> 0 And .CaFdBplAdj <> 0 Then
                .CaPctAdjRcvry = Round(100 * (.CaCnBplAdj / .CaFdBplAdj) * _
                                 (.CaFdBplAdj - .CaTlBplAdj) / _
                                 (.CaCnBplAdj - .CaTlBplAdj), 1)
            Else
                .CaPctAdjRcvry = 0
            End If

            If .CaCnBplRpt - .CaTlBplRpt <> 0 And .CaFdBplRpt <> 0 Then
                .CaPctRptRcvry = Round(100 * (.CaCnBplRpt / .CaFdBplRpt) * _
                                 (.CaFdBplRpt - .CaTlBplRpt) / _
                                 (.CaCnBplRpt - .CaTlBplRpt), 1)
            Else
                .CaPctRptRcvry = 0
            End If

            'Total plant miscellaneous  Total plant miscellaneous
            'Total plant miscellaneous  Total plant miscellaneous
            'Total plant miscellaneous  Total plant miscellaneous

            'Total plant ratio of concentration adjusted
            If .TotPltFdBplAdj - .TotPltTlBplAdj <> 0 Then
                .TotPltRcAdj = Round((.TotPltCnBplAdj - .TotPltTlBplAdj) / _
                               (.TotPltFdBplAdj - .TotPltTlBplAdj), 2)
            Else
                .TotPltRcAdj = 0
            End If

            'Total plant feed tons adjusted
            .TotPltFdTonsAdj = Round(.TotPltRcAdj * .PrdCnTons, 0)

            'Total plant concentrate tons adjusted
            .TotPltCnTonsAdj = .PrdCnTons

            'Total plant tail tons adjusted
            .TotPltTlTonsAdj = .TotPltFdTonsAdj - .TotPltCnTonsAdj

            'Total plant adjusted recovery  (Actual recovery)
            If .TotPltCnBplAdj - .TotPltTlBplAdj <> 0 And .TotPltFdBplAdj <> 0 Then
                .TotPltPctAdjRcvry = Round(100 * (.TotPltCnBplAdj / .TotPltFdBplAdj) * _
                                     (.TotPltFdBplAdj - .TotPltTlBplAdj) / _
                                     (.TotPltCnBplAdj - .TotPltTlBplAdj), 2)
            Else
                .TotPltPctAdjRcvry = 0
            End If

            'Total plant reported recovery
            'Use the total plant tail BPL based on GMT tail BPL
            If .TotPltCnBplRpt - .TotPltTlBplRpt2 <> 0 And .TotPltFdBplRpt <> 0 Then
                .TotPltPctRptRcvry = Round(100 * (.TotPltCnBplRpt / .TotPltFdBplRpt) * _
                                     (.TotPltFdBplRpt - .TotPltTlBplRpt2) / _
                                     (.TotPltCnBplRpt - .TotPltTlBplRpt2), 2)
            Else
                .TotPltPctRptRcvry = 0
            End If

            If .TotPltHrs <> 0 Then
                .TotPltFdTphAdj = Round(.TotPltFdTonsAdj / .TotPltHrs, 0)
            Else
                .TotPltFdTphAdj = 0
            End If

            If .TotPltHrs <> 0 Then
                .TotPltCnTphAdj = Round(.TotPltCnTonsAdj / .TotPltHrs, 0)
            Else
                .TotPltCnTphAdj = 0
            End If

            If .TotPltHrs <> 0 Then
                .TotPltTlTphAdj = Round(.TotPltTlTonsAdj / .TotPltHrs, 0)
            Else
                .TotPltTlTphAdj = 0
            End If

            'Calculate an "as reported" total plant tail BPL

            'heyhey
        End With

        Exit Sub

ProcessFcMassBalanceTotalsError:

        MsgBox("Error in processing Four Corners mass balance totals." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Four Corners Mass Balance Totals Computation Error")
    End Sub

    Public Function gAdjustedFeedTonsFC(ByVal aBeginDate As Date, _
                                        ByVal aBeginShift As String, _
                                        ByVal aEndDate As Date, _
                                        ByVal aEndShift As String, _
                                        ByVal aCrewNumber As String, _
                                        ByRef rFeedBpl As Double, _
                                        ByRef rFeedTonsRpt As Long, _
                                        ByVal aBplRound As Integer, _
                                        ByVal aMassBalMode As String) As Long

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'This function will return the total adjusted feed tons for
        'any given time frame.

        On Error GoTo gAdjustedFeedTonsFcError

        Dim RowIdx As Integer
        Dim NumShifts As Integer

        Dim FloatPlantCirc(9, 14) As Object
        Dim FloatPlantGmt(4, 7) As Object
        'ReDim FloatPlantCirc(0 To 9, 0 To 14)
        'ReDim FloatPlantGmt(0 To 4, 0 To 7)

        'Get data for float plant mass balance
        NumShifts = gGetFcFloatPlantBalanceData(FloatPlantCirc, _
                                                FloatPlantGmt, _
                                                aBeginDate, _
                                                StrConv(aBeginShift, vbUpperCase), _
                                                aEndDate, _
                                                StrConv(aEndShift, vbUpperCase), _
                                                aCrewNumber, _
                                                aBplRound, _
                                                aMassBalMode)

        gAdjustedFeedTonsFC = FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcFdTonsAdj)

        rFeedBpl = FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcFdBpl)
        rFeedTonsRpt = FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcFdTonsRpt) + _
                       FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcFdTonsRpt)

        Exit Function

gAdjustedFeedTonsFcError:

        MsgBox("Error summing Four Corners adjusted feed tons." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Four Corners Adjusted Feed Tons Error")
    End Function

    Public Function gGetBalanceDistribution(ByVal aMineName As String, _
                                            ByVal aDate As Date, _
                                            ByVal aShift As String) As mBalanceDistributionType

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        With gGetBalanceDistribution
            .FrPct = gGetPeriodicEqptMsrmnt(aMineName, _
                                            aDate, _
                                            StrConv(aShift, vbUpperCase), _
                                            "Float plant rougher circuit", _
                                            "Fine rougher", _
                                            "Balance distribution")

            .CrPct = gGetPeriodicEqptMsrmnt(aMineName, _
                                            aDate, _
                                            StrConv(aShift, vbUpperCase), _
                                            "Float plant rougher circuit", _
                                            "Coarse rougher", _
                                            "Balance distribution")

            .FaPct = gGetPeriodicEqptMsrmnt(aMineName, _
                                            aDate, _
                                            StrConv(aShift, vbUpperCase), _
                                            "Float plant cleaner circuit", _
                                            "Fine amine", _
                                            "Balance distribution")

            .CaPct = gGetPeriodicEqptMsrmnt(aMineName, _
                                            aDate, _
                                            StrConv(aShift, vbUpperCase), _
                                            "Float plant cleaner circuit", _
                                            "Coarse amine", _
                                            "Balance distribution")
        End With
    End Function

    Public Function gGetBalanceDistribution2(ByVal aMineName As String, _
                                             ByVal aBeginDate As Date, _
                                             ByVal aBeginShift As String, _
                                             ByVal aEndDate As Date, _
                                             ByVal aEndShift As String) As mBalanceDistributionType

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim CorrAmt As Single

        With gGetBalanceDistribution2
            .FrPct = gGetPeriodicEqptMsrAvg3(aMineName, _
                                             aBeginDate, _
                                             StrConv(aBeginShift, vbUpperCase), _
                                             aEndDate, _
                                             StrConv(aEndShift, vbUpperCase), _
                                             "Float plant rougher circuit", _
                                             "Fine rougher", _
                                             "Balance distribution")

            .FrPct = Round(.FrPct, 0)

            .CrPct = gGetPeriodicEqptMsrAvg3(aMineName, _
                                             aBeginDate, _
                                             StrConv(aBeginShift, vbUpperCase), _
                                             aEndDate, _
                                             StrConv(aEndShift, vbUpperCase), _
                                             "Float plant rougher circuit", _
                                             "Coarse rougher", _
                                             "Balance distribution")
            .CrPct = Round(.CrPct)

            .FaPct = gGetPeriodicEqptMsrAvg3(aMineName, _
                                             aBeginDate, _
                                             StrConv(aBeginShift, vbUpperCase), _
                                             aEndDate, _
                                             StrConv(aEndShift, vbUpperCase), _
                                             "Float plant cleaner circuit", _
                                             "Fine amine", _
                                             "Balance distribution")
            .FaPct = Round(.FaPct)

            .CaPct = gGetPeriodicEqptMsrAvg3(aMineName, _
                                             aBeginDate, _
                                             StrConv(aBeginShift, vbUpperCase), _
                                             aEndDate, _
                                             StrConv(aEndShift, vbUpperCase), _
                                             "Float plant cleaner circuit", _
                                             "Coarse amine", _
                                             "Balance distribution")
            .CaPct = Round(.CaPct)


            'The above four average values shoud really add to 100!
            'Will make the assumption that .FrPCt will never be zero!
            CorrAmt = 100 - .FaPct - .CaPct - .FrPct - .CrPct
            .FrPct = .FrPct + CorrAmt

            'Now the four average values will add to 100!
        End With
    End Function

    Public Function gGetMetReagentDataFc(ByVal aBeginDate As Date, _
                                         ByVal aBeginShift As String, _
                                         ByVal aEndDate As Date, _
                                         ByVal aEndShift As String, _
                                         ByVal aCrewNumber As String, _
                                         ByVal TotAdjFdTons As Long, _
                                         ByVal TotRptFdTons As Long, _
                                         ByVal TotCnTons As Long, _
                                         ByRef mMbFcReag As gMassBalanceFcReagDataType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetMetReagentDataFcError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        Dim MetReagentDynaset As OraDynaset
        Dim RecCount As Long

        Dim ThisMatl As String

        Dim TotCost As Long
        Dim TotUnits As Long

        Dim BeginDate As Date
        Dim BeginShift As String
        Dim EndDate As Date
        Dim EndShift As String

        BeginDate = aBeginDate
        BeginShift = StrConv(gGetFirstShift2("Four Corners", _
                     aBeginDate), vbUpperCase)

        EndDate = aEndDate
        EndShift = StrConv(gGetLastShift2("Four Corners", _
                   aEndDate), vbUpperCase)

        TotCost = 0
        TotUnits = 0

        'Get reagent data from EQPT_CALC
        params = gDBParams

        params.Add("pMineName", "Four Corners", ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pBeginDate", BeginDate, ORAPARM_INPUT)
        params("pBeginDate").serverType = ORATYPE_DATE

        params.Add("pBeginShift", StrConv(BeginShift, vbUpperCase), ORAPARM_INPUT)
        params("pBeginShift").serverType = ORATYPE_VARCHAR2

        params.Add("pEndDate", EndDate, ORAPARM_INPUT)
        params("pEndDate").serverType = ORATYPE_DATE

        params.Add("pEndShift", StrConv(EndShift, vbUpperCase), ORAPARM_INPUT)
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
            gGetMetReagentDataFc = False
            ClearParams(params)
            Exit Function
        End If

        gGetMetReagentDataFc = True
        mMbFcReag.RgAllTotCost = 0
        mMbFcReag.RgAllTotUnits = 0

        MetReagentDynaset.MoveFirst()

        Do While Not MetReagentDynaset.EOF
            ThisMatl = MetReagentDynaset.Fields("matl_name").Value
            With mMbFcReag
                Select Case ThisMatl
                    Case Is = "Amine"
                        .RgAmTotUnits = MetReagentDynaset.Fields("pound_usage").Value
                        .RgAmTotCost = MetReagentDynaset.Fields("cost").Value
                        .RgAllTotUnits = .RgAllTotUnits + .RgAmTotUnits
                        .RgAllTotCost = .RgAllTotCost + .RgAmTotCost

                    Case Is = "Defoamer"
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

                    Case Is = "Sodium silicate"
                        .RgSiTotUnits = MetReagentDynaset.Fields("pound_usage").Value
                        .RgSiTotCost = MetReagentDynaset.Fields("cost").Value
                        .RgAllTotUnits = .RgAllTotUnits + .RgSiTotUnits
                        .RgAllTotCost = .RgAllTotCost + .RgSiTotCost

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

                    Case Is = "Fatty acid CF109"
                        .RgFa2TotUnits = MetReagentDynaset.Fields("pound_usage").Value
                        .RgFa2TotCost = MetReagentDynaset.Fields("cost").Value
                        .RgAllTotUnits = .RgAllTotUnits + .RgFa2TotUnits
                        .RgAllTotCost = .RgAllTotCost + .RgFa2TotCost
                End Select
            End With

            MetReagentDynaset.MoveNext()
        Loop

        ClearParams(params)

        With mMbFcReag
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
                .RgFa2AdjFdDpt = Round(.RgFa2TotCost / TotAdjFdTons, 4)
            Else
                .RgFa2AdjFdDpt = 0
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
                .RgSiAdjFdDpt = Round(.RgSiTotCost / TotAdjFdTons, 4)
            Else
                .RgSiAdjFdDpt = 0
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
            '-----
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
                .RgFa2AdjFdUpt = Round(.RgFa2TotUnits / TotAdjFdTons, 4)
            Else
                .RgFa2AdjFdUpt = 0
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
                .RgSiAdjFdUpt = Round(.RgSiTotUnits / TotAdjFdTons, 4)
            Else
                .RgSiAdjFdUpt = 0
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
                .RgFa2RptFdDpt = Round(.RgFa2TotCost / TotRptFdTons, 4)
            Else
                .RgFa2RptFdDpt = 0
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
                .RgSiRptFdDpt = Round(.RgSiTotCost / TotRptFdTons, 4)
            Else
                .RgSiRptFdDpt = 0
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
            '-----
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
                .RgFa2RptFdUpt = Round(.RgFa2TotUnits / TotRptFdTons, 4)
            Else
                .RgFa2RptFdUpt = 0
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
                .RgSiRptFdUpt = Round(.RgSiTotUnits / TotRptFdTons, 4)
            Else
                .RgSiRptFdUpt = 0
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
                .RgFa2CnDpt = Round(.RgFa2TotCost / TotCnTons, 4)
            Else
                .RgFa2CnDpt = 0
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
                .RgSiCnDpt = Round(.RgSiTotCost / TotCnTons, 4)
            Else
                .RgSiCnDpt = 0
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
            '-----
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
                .RgFa2CnUpt = Round(.RgFa2TotUnits / TotCnTons, 4)
            Else
                .RgFa2CnUpt = 0
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
                .RgSiCnUpt = Round(.RgSiTotUnits / TotCnTons, 4)
            Else
                .RgSiCnUpt = 0
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

gGetMetReagentDataFcError:

        MsgBox("Error getting Four Corners reagent data." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Four Corners Reagent Data Error")

        On Error Resume Next
        ClearParams(params)
    End Function

    Public Function gFltPltRcvryFC(ByVal aBeginDate As Date, _
                                   ByVal aBeginShift As String, _
                                   ByVal aEndDate As Date, _
                                   ByVal aEndShift As String, _
                                   ByVal aCrewNumber As String, _
                                   ByRef rFeedBpl As Single, _
                                   ByVal aBplRound As Integer, _
                                   ByVal aMassBalanceMode As String) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'This function will return the plant recovery for
        'any given time frame.  It will also return the adjusted feed BPL through
        'rFeedBpl.

        On Error GoTo gFltPltRcvryFcError

        Dim RowIdx As Integer
        Dim NumShifts As Integer

        Dim FloatPlantCirc(9, 14) As Object
        Dim FloatPlantGmt(4, 7) As Object

        'Get data for float plant mass balance
        NumShifts = gGetFcFloatPlantBalanceData(FloatPlantCirc, _
                                                FloatPlantGmt, _
                                                aBeginDate, _
                                                StrConv(aBeginShift, vbUpperCase), _
                                                aEndDate, _
                                                StrConv(aEndShift, vbUpperCase), _
                                                aCrewNumber, _
                                                aBplRound, _
                                                aMassBalanceMode)

        gFltPltRcvryFC = FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcPctActRcvry)

        rFeedBpl = FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcFdBpl)

        Exit Function

gFltPltRcvryFcError:

        MsgBox("Error getting Four Corners float plant recovery." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Four Corners Plant Recovery Error")
    End Function

    Private Sub GetPeriodAvgsAndSums(ByVal aBeginDate As Date, _
                                     ByVal aBeginShift As String, _
                                     ByVal aEndDate As Date, _
                                     ByVal aEndShift As String, _
                                     ByVal aBplRound As Integer, _
                                     ByVal aMassBalMode As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Get the circuit BPL averages

        If aMassBalMode = "Circuit" Then
            GetCircBplAvgs(aBeginDate, _
                            aBeginShift, _
                            aEndDate, _
                            aEndShift, _
                            aBplRound)
        End If

        If aMassBalMode = "GMT" Then
            GetCircBplAvgs2(aBeginDate, _
                            aBeginShift, _
                            aEndDate, _
                            aEndShift, _
                            aBplRound)
        End If

        'Get the period concentrate tons
        GetPeriodCnTons(aBeginDate, _
                        aBeginShift, _
                        aEndDate, _
                        aEndShift)

        'Get the period reported feed tons and operating hours
        GetCircTonsHrs(aBeginDate, _
                       aBeginShift, _
                       aEndDate, _
                       aEndShift)

        'Get some special operating hours -- need to know the coarse rougher
        'operating hours where the coarse amine was operating.
        GetCrsAmineOperHrs(aBeginDate, _
                           aBeginShift, _
                           aEndDate, _
                           aEndShift)
    End Sub

    Private Sub GetCircBplAvgs(ByVal aBeginDate As Date, _
                               ByVal aBeginShift As String, _
                               ByVal aEndDate As Date, _
                               ByVal aEndShift As String, _
                               ByVal aBplRound As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetCircBplAvgsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecCount As Long
        Dim CircAvgDynaset As OraDynaset
        Dim ThisEqpt As String
        Dim ThisMsrName As String
        Dim ThisAvgVal As Single

        'Get the circuit bpl averages
        params = gDBParams

        params.Add("pMineName", "Four Corners", ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pStartDate", aBeginDate, ORAPARM_INPUT)
        params("pStartDate").serverType = ORATYPE_DATE

        params.Add("pStartShift", StrConv(aBeginShift, vbUpperCase), ORAPARM_INPUT)
        params("pStartShift").serverType = ORATYPE_VARCHAR2

        params.Add("pStopDate", aEndDate, ORAPARM_INPUT)
        params("pStopDate").serverType = ORATYPE_DATE

        params.Add("pStopShift", StrConv(aEndShift, vbUpperCase), ORAPARM_INPUT)
        params("pStopShift").serverType = ORATYPE_VARCHAR2

        params.Add("pRoundVal", aBplRound, ORAPARM_INPUT)
        params("pRoundVal").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_massbal_avg_ext_msr
        'pMineName          IN     VARCHAR2,
        'pStartDate         IN     DATE,
        'pStartShift        IN     VARCHAR2,
        'pStopDate          IN     DATE,
        'pStopShift         IN     VARCHAR2,
        'pRoundVal          IN     NUMBER,
        'pResult            IN OUT c_massbalavg)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_floatplant2.get_massbal_avg_ext_msr(:pMineName, " + _
                      ":pStartDate, :pStartShift, :pStopDate, :pStopShift, :pRoundVal, :pResult);end;", ORASQL_FAILEXEC)
        CircAvgDynaset = params("pResult").Value
        ClearParams(params)

        RecCount = CircAvgDynaset.RecordCount

        'Assign the average data to mMbFcShift
        Do While Not CircAvgDynaset.EOF
            ThisEqpt = CircAvgDynaset.Fields("eqpt_name").Value
            ThisMsrName = CircAvgDynaset.Fields("measure_name").Value
            ThisAvgVal = CircAvgDynaset.Fields("avg_val").Value

            With mMbFcShift
                Select Case ThisEqpt
                    Case Is = "North fine rougher"
                        'Feed BPL, Concentrate BPL and Tail BPL
                        If ThisMsrName = "Feed BPL" Then
                            .NfrFdBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Concentrate BPL" Then
                            .NfrCnBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Tail BPL" Then
                            .NfrTlBplRpt = ThisAvgVal
                        End If

                    Case Is = "South fine rougher"
                        'Feed BPL, Concentrate BPL and Tail BPL
                        If ThisMsrName = "Feed BPL" Then
                            .SfrFdBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Concentrate BPL" Then
                            .SfrCnBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Tail BPL" Then
                            .SfrTlBplRpt = ThisAvgVal
                        End If

                    Case Is = "Coarse rougher"
                        'Concentrate BPL
                        If ThisMsrName = "Concentrate BPL" Then
                            .CrCnBplRpt = ThisAvgVal
                        End If

                    Case Is = "Coarse column"
                        'Feed BPL, Concentrate BPL and Tail BPL
                        If ThisMsrName = "Feed BPL" Then
                            .CrsColFdBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Concentrate BPL" Then
                            .CrsColCnBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Tail BPL" Then
                            .CrsColTlBplRpt = ThisAvgVal
                        End If

                    Case Is = "Fine amine"
                        'Tail BPL
                        If ThisMsrName = "Tail BPL" Then
                            .FaTlBplRpt = ThisAvgVal
                        End If

                    Case Is = "Coarse amine"
                        'Tail BPL
                        If ThisMsrName = "Tail BPL" Then
                            .CaTlBplRpt = ThisAvgVal
                        End If

                    Case Is = "Float plant"
                        'Tail BPL
                        If ThisMsrName = "Tail BPL" Then
                            .GmtBplRpt = ThisAvgVal
                        End If
                End Select
            End With
            CircAvgDynaset.MoveNext()
        Loop

        CircAvgDynaset.Close()

        Exit Sub

GetCircBplAvgsError:

        MsgBox("Error getting period average data." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Four Corners Period Average Error")

        On Error Resume Next
        CircAvgDynaset.Close()
        ClearParams(params)
    End Sub

    Private Sub GetCircTonsHrs(ByVal aBeginDate As Date, _
                               ByVal aBeginShift As String, _
                               ByVal aEndDate As Date, _
                               ByVal aEndShift As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetCircTonsHrsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecCount As Long
        Dim CircSumDynaset As OraDynaset
        Dim ThisEqpt As String
        Dim ThisMsrName As String
        Dim ThisSumVal As Single

        'Get the circuit ton and hour sums
        params = gDBParams

        params.Add("pMineName", "Four Corners", ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pStartDate", aBeginDate, ORAPARM_INPUT)
        params("pStartDate").serverType = ORATYPE_DATE

        params.Add("pStartShift", StrConv(aBeginShift, vbUpperCase), ORAPARM_INPUT)
        params("pStartShift").serverType = ORATYPE_VARCHAR2

        params.Add("pStopDate", aEndDate, ORAPARM_INPUT)
        params("pStopDate").serverType = ORATYPE_DATE

        params.Add("pStopShift", StrConv(aEndShift, vbUpperCase), ORAPARM_INPUT)
        params("pStopShift").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'Procedure get_massbal_sums
        'pMineName           IN     VARCHAR2,
        'pStartDate          IN     DATE,
        'pStartShift         IN     VARCHAR2,
        'pStopDate           IN     DATE,
        'pStopShift          IN     VARCHAR2,
        'pResult             IN OUT c_massbalsum);
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_floatplant2.get_massbal_sums(:pMineName, " + _
                      ":pStartDate, :pStartShift, :pStopDate, :pStopShift, :pResult);end;", ORASQL_FAILEXEC)
        CircSumDynaset = params("pResult").Value
        ClearParams(params)

        RecCount = CircSumDynaset.RecordCount

        'Assign the summed data to mMbFcShift
        Do While Not CircSumDynaset.EOF
            ThisEqpt = CircSumDynaset.Fields("eqpt_name").Value
            ThisMsrName = CircSumDynaset.Fields("measure_name").Value
            ThisSumVal = CircSumDynaset.Fields("sum_val").Value

            With mMbFcShift
                Select Case ThisEqpt
                    Case Is = "North coarse rougher"
                        'Operating hours and Reported feed tons
                        If ThisMsrName = "Operating hours" Then
                            .NcrHrs = ThisSumVal
                        End If
                        If ThisMsrName = "Sized feed tons" Then
                            .NcrFdTonsRpt = ThisSumVal
                        End If

                    Case Is = "South coarse rougher"
                        'Operating hours and Reported feed tons
                        If ThisMsrName = "Operating hours" Then
                            .ScrHrs = ThisSumVal
                        End If
                        If ThisMsrName = "Sized feed tons" Then
                            .ScrFdTonsRpt = ThisSumVal
                        End If

                    Case Is = "North fine rougher 1"
                        'Operating hours and Reported feed tons
                        If ThisMsrName = "Operating hours" Then
                            .Nfr1Hrs = ThisSumVal
                        End If
                        If ThisMsrName = "Sized feed tons" Then
                            .Nfr1FdTonsRpt = ThisSumVal
                        End If

                    Case Is = "North fine rougher 2"
                        'Operating hours and Reported feed tons
                        If ThisMsrName = "Operating hours" Then
                            .Nfr2Hrs = ThisSumVal
                        End If
                        If ThisMsrName = "Sized feed tons" Then
                            .Nfr2FdTonsRpt = ThisSumVal
                        End If

                    Case Is = "South fine rougher 1"
                        'Operating hours and Reported feed tons
                        If ThisMsrName = "Operating hours" Then
                            .Sfr1Hrs = ThisSumVal
                        End If
                        If ThisMsrName = "Sized feed tons" Then
                            .Sfr1FdTonsRpt = ThisSumVal
                        End If

                    Case Is = "South fine rougher 2"
                        'Operating hours and Reported feed tons
                        If ThisMsrName = "Operating hours" Then
                            .Sfr2Hrs = ThisSumVal
                        End If
                        If ThisMsrName = "Sized feed tons" Then
                            .Sfr2FdTonsRpt = ThisSumVal
                        End If

                    Case Is = "North coarse scalp"
                        'Operating hours
                        If ThisMsrName = "Operating hours" Then
                            .NcsHrs = ThisSumVal
                        End If

                    Case Is = "South coarse scalp"
                        'Operating hours
                        If ThisMsrName = "Operating hours" Then
                            .ScsHrs = ThisSumVal
                        End If
                End Select
            End With
            CircSumDynaset.MoveNext()
        Loop

        CircSumDynaset.Close()

        Exit Sub

GetCircTonsHrsError:

        MsgBox("Error getting period sum data." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Four Corners Period Sum Error")

        On Error Resume Next
        CircSumDynaset.Close()
        ClearParams(params)
    End Sub

    Private Sub GetCrsAmineOperHrs(ByVal aBeginDate As Date, _
                                   ByVal aBeginShift As String, _
                                   ByVal aEndDate As Date, _
                                   ByVal aEndShift As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetCrsAmineOperHrsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecCount As Long
        Dim CircSumDynaset As OraDynaset
        Dim ThisSumVal As Single
        Dim RghrIdx As Integer
        Dim RghrCirc As String

        For RghrIdx = 1 To 2
            If RghrIdx = 1 Then
                RghrCirc = "North coarse rougher"
            Else
                RghrCirc = "South coarse rougher"
            End If

            'Get the circuit ton and hour sums
            params = gDBParams

            params.Add("pMineName", "Four Corners", ORAPARM_INPUT)
            params("pMineName").serverType = ORATYPE_VARCHAR2

            params.Add("pStartDate", aBeginDate, ORAPARM_INPUT)
            params("pStartDate").serverType = ORATYPE_DATE

            params.Add("pStartShift", StrConv(aBeginShift, vbUpperCase), ORAPARM_INPUT)
            params("pStartShift").serverType = ORATYPE_VARCHAR2

            params.Add("pStopDate", aEndDate, ORAPARM_INPUT)
            params("pStopDate").serverType = ORATYPE_DATE

            params.Add("pStopShift", StrConv(aEndShift, vbUpperCase), ORAPARM_INPUT)
            params("pStopShift").serverType = ORATYPE_VARCHAR2

            params.Add("pEqptTypeName", "Float plant rougher circuit", ORAPARM_INPUT)
            params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

            params.Add("pEqptName", RghrCirc, ORAPARM_INPUT)
            params("pEqptName").serverType = ORATYPE_VARCHAR2

            params.Add("pEqptTypeNameBalDist", "Float plant cleaner circuit", ORAPARM_INPUT)
            params("pEqptTypeNameBalDist").serverType = ORATYPE_VARCHAR2

            params.Add("pEqptNameBalDist", "Coarse amine", ORAPARM_INPUT)
            params("pEqptNameBalDist").serverType = ORATYPE_VARCHAR2

            params.Add("pResult", 0, ORAPARM_OUTPUT)
            params("pResult").serverType = ORATYPE_CURSOR

            'PROCEDURE get_ophrs_special
            'pMineName            IN     VARCHAR2,
            'pStartDate           IN     DATE,
            'pStartShift          IN     VARCHAR2,
            'pStopDate            IN     DATE,
            'pStopShift           IN     VARCHAR2,
            'pEqptTypeName        IN     VARCHAR2,
            'pEqptName            IN     VARCHAR2,
            'pEqptTypeNameBalDist IN     VARCHAR2,
            'pEqptNameBalDist     IN     VARCHAR2,
            'pResult              IN OUT c_massbalsum)
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_floatplant2.get_ophrs_special(:pMineName, " + _
                          ":pStartDate, :pStartShift, :pStopDate, :pStopShift, " + _
                          ":pEqptTypeName, :pEqptName, " + _
                          ":pEqptTypeNameBalDist, :pEqptNameBalDist, :pResult);end;", ORASQL_FAILEXEC)
            CircSumDynaset = params("pResult").Value
            ClearParams(params)

            RecCount = CircSumDynaset.RecordCount

            'Assign the summed data to mMbFcShift
            If RecCount = 1 Then
                CircSumDynaset.MoveFirst()
                ThisSumVal = CircSumDynaset.Fields("sum_val").Value

                If RghrIdx = 1 Then
                    mMbFcShift.NcrCaHrs = ThisSumVal
                Else
                    mMbFcShift.ScrCaHrs = ThisSumVal
                End If
            Else
                If RghrIdx = 1 Then
                    mMbFcShift.NcrCaHrs = 0
                Else
                    mMbFcShift.ScrCaHrs = 0
                End If
            End If
        Next RghrIdx


        CircSumDynaset.Close()

        Exit Sub

GetCrsAmineOperHrsError:

        MsgBox("Error getting coarse amine oper hrs." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Operating Hours Error")

        On Error Resume Next
        CircSumDynaset.Close()
        ClearParams(params)
    End Sub

    Private Sub ZeroFcShiftData()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        With mMbFcShift
            .Nfr1Hrs = 0
            .Nfr2Hrs = 0
            .Sfr1Hrs = 0
            .Sfr2Hrs = 0
            .NcrHrs = 0
            .ScrHrs = 0
            .NcsHrs = 0
            .ScsHrs = 0
            .NcrCaHrs = 0
            .ScrCaHrs = 0
            .CrHrs = 0   '!!!
            .CaHrs = 0   '!!!
            .CaPct = 0   '!!!
            '-----
            .NfrRc = 0
            .SfrRc = 0
            .FaRc = 0
            .NcrRc = 0
            .ScrRc = 0
            .CaRc = 0
            '-----
            .PrdCnBpl = 0
            .PrdCnTons = 0
            '-----
            .NfrFdBplRpt = 0
            .SfrFdBplRpt = 0
            .NcrFdBplRpt = 0
            .ScrFdBplRpt = 0
            .CrsColFdBplRpt = 0
            '-----
            .NfrCnBplRpt = 0
            .SfrCnBplRpt = 0
            .CrCnBplRpt = 0
            .CrsColCnBplRpt = 0
            '-----
            .CrsColTlBplRpt = 0
            '-----
            .NfrTlBplRpt = 0
            .SfrTlBplRpt = 0
            .FaTlBplRpt = 0
            .NcrTlBplRpt = 0
            .ScrTlBplRpt = 0
            .CaTlBplRpt = 0
            .GmtBplRpt = 0
            '-----
            .NfrTlBplAdj = 0
            .SfrTlBplAdj = 0
            .FaTlBplAdj = 0
            .NcrTlBplAdj = 0
            .ScrTlBplAdj = 0
            .CaTlBplAdj = 0
            .GmtBplAdj = 0
            .FrTlBplAdj = 0
            .CrTlBplAdj = 0
            '-----
            .Nfr1FdTonsRpt = 0
            .Nfr2FdTonsRpt = 0
            .Sfr1FdTonsRpt = 0
            .Sfr2FdTonsRpt = 0
            .NfrFdTonsRpt = 0
            .SfrFdTonsRpt = 0
            .NcrFdTonsRpt = 0
            .ScrFdTonsRpt = 0
            .FaFdTonsRpt = 0
            .CaFdTonsRpt = 0
            .FrFdTonsRpt = 0
            .CrFdTonsRpt = 0
            '-----
            .NfrFdTonsAdj = 0
            .SfrFdTonsAdj = 0
            .NcrFdTonsAdj = 0
            .ScrFdTonsAdj = 0
            .FaFdTonsAdj = 0
            .CaFdTonsAdj = 0
            .FrFdTonsAdj = 0
            .CrFdTonsAdj = 0
            '-----
            .AvgFrCnBpl = 0
            .AvgCrCnBpl = 0
            '-----
            .NfrTlBtRpt = 0
            .SfrTlBtRpt = 0
            .FaTlBtRpt = 0
            .NcrTlBtRpt = 0
            .ScrTlBtRpt = 0
            .CaTlBtRpt = 0
            '-----
            .NfrTlBtAdj = 0
            .SfrTlBtAdj = 0
            .FaTlBtAdj = 0
            .NcrTlBtAdj = 0
            .ScrTlBtAdj = 0
            .CaTlBtAdj = 0
            '-----
            .NfrCnTonsRpt = 0
            .SfrCnTonsRpt = 0
            .FaCnTonsRpt = 0
            .NcrCnTonsRpt = 0
            .ScrCnTonsRpt = 0
            .CaCnTonsRpt = 0
            '-----
            .NfrCnTonsAdj = 0
            .SfrCnTonsAdj = 0
            .FaCnTonsAdj = 0
            .NcrCnTonsAdj = 0
            .ScrCnTonsAdj = 0
            .CaCnTonsAdj = 0
            '-----
            .NfrTlTonsRpt = 0
            .SfrTlTonsRpt = 0
            .FaTlTonsRpt = 0
            .NcrTlTonsRpt = 0
            .ScrTlTonsRpt = 0
            .CaTlTonsRpt = 0
            .FrTlTonsRpt = 0
            .CrTlTonsRpt = 0
            '-----
            .NfrTlTonsAdj = 0
            .SfrTlTonsAdj = 0
            .FaTlTonsAdj = 0
            .NcrTlTonsAdj = 0
            .ScrTlTonsAdj = 0
            .CaTlTonsAdj = 0
            .FrTlTonsAdj = 0
            .CrTlTonsAdj = 0
        End With
    End Sub

    Private Sub GetPeriodCnTons(ByVal aBeginDate As Date, _
                                ByVal aBeginShift As String, _
                                ByVal aEndDate As Date, _
                                ByVal aEndShift As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim NumberOfDetailRows As Integer
        Dim fProdResults() As Object
        Dim ActiveMineNameLongSave As String
        Dim CnRow As Integer

        ActiveMineNameLongSave = gActiveMineNameLong
        gActiveMineNameLong = "Four Corners"

        'NumberOfDetailRows = gGetProductionArray(aBeginDate, _
        '                                         aBeginShift, _
        '                                         aEndDate, _
        '                                         aEndShift, _
        '                                         fProdResults)

        gActiveMineNameLong = ActiveMineNameLongSave

        'Columns in ProdResults() will be:
        'Col 1   "BPL"
        'Col 2   "Insol"
        'Col 3   "Fe2O3"
        'Col 4   "Al2O3"
        'Col 5   "MgO"
        'Col 6   "CaO"
        'Col 7   "Cd"
        'Col 8   "Tons"
        'Col 9   "P2O5"
        'Col 10  "CaO/ P2O5"
        'Col 11  "MER"
        'Col 12  "MgO/ P2O5"

        'Don't depend on a particular product to be in a particular row.
        'CnRow = gGetProductionRow(fProdResults, "Total concentrate")

        'mMbFcShift.PrdCnTons = fProdResults(CnRow, 8)
        'mMbFcShift.PrdCnBpl = fProdResults(CnRow, 1)
    End Sub

    Private Sub GetCircBplAvgsAll(ByVal aBeginDate As Date, _
                                  ByVal aBeginShift As String, _
                                  ByVal aEndDate As Date, _
                                  ByVal aEndShift As String, _
                                  ByVal aBplRound As Integer)

        '**********************************************************************
        '  This procedure just gets all of the data back and does the averaging
        '  here in VB.
        '
        '**********************************************************************

        On Error GoTo GetCircBplAvgsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecCount As Long
        Dim CircAvgDynaset As OraDynaset
        Dim ThisEqpt As String
        Dim ThisMsrName As String
        Dim ThisAvgVal As Single

        'Get the circuit bpl averages
        params = gDBParams

        params.Add("pMineName", "Four Corners", ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pStartDate", aBeginDate, ORAPARM_INPUT)
        params("pStartDate").serverType = ORATYPE_DATE

        params.Add("pStartShift", StrConv(aBeginShift, vbUpperCase), ORAPARM_INPUT)
        params("pStartShift").serverType = ORATYPE_VARCHAR2

        params.Add("pStopDate", aEndDate, ORAPARM_INPUT)
        params("pStopDate").serverType = ORATYPE_DATE

        params.Add("pStopShift", StrConv(aEndShift, vbUpperCase), ORAPARM_INPUT)
        params("pStopShift").serverType = ORATYPE_VARCHAR2

        params.Add("pRoundVal", aBplRound, ORAPARM_INPUT)
        params("pRoundVal").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_massbal_all_ext_msr
        'pMineName            IN     VARCHAR2,
        'pStartDate           IN     DATE,
        'pStartShift          IN     VARCHAR2,
        'pStopDate            IN     DATE,
        'pStopShift           IN     VARCHAR2,
        'pCrewNumber          IN     VARCHAR2,
        'pResult              IN OUT c_massbalance);
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_floatplant2.get_massbal_all_ext_msr(:pMineName, " + _
                      ":pStartDate, :pStartShift, :pStopDate, :pStopShift, :pRoundVal, :pResult);end;", ORASQL_FAILEXEC)
        CircAvgDynaset = params("pResult").Value
        ClearParams(params)

        RecCount = CircAvgDynaset.RecordCount

        'Need to add averaging code here if I decide to go this route!

        CircAvgDynaset.Close()

        Exit Sub

GetCircBplAvgsError:

        MsgBox("Error getting period average data." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Four Corners Period Average Error")

        On Error Resume Next
        CircAvgDynaset.Close()
        ClearParams(params)
    End Sub

    Private Sub GetCircBplAvgs2(ByVal aBeginDate As Date, _
                                ByVal aBeginShift As String, _
                                ByVal aEndDate As Date, _
                                ByVal aEndShift As String, _
                                ByVal aBplRound As Integer)

        '**********************************************************************
        '  This proc makes use of Bob's procedure get_massbal_ext_msr_spec
        '  which uses his view mois.circuit_daily_avg_v.
        '
        '**********************************************************************

        On Error GoTo GetCircBplAvgsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecCount As Long
        Dim CircAvgDynaset As OraDynaset
        Dim ThisEqpt As String
        Dim ThisMsrName As String
        Dim ThisAvgVal As Single

        'Get the circuit bpl averages
        params = gDBParams

        params.Add("pMineName", "Four Corners", ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pStartDate", aBeginDate, ORAPARM_INPUT)
        params("pStartDate").serverType = ORATYPE_DATE

        params.Add("pStartShift", StrConv(aBeginShift, vbUpperCase), ORAPARM_INPUT)
        params("pStartShift").serverType = ORATYPE_VARCHAR2

        params.Add("pStopDate", aEndDate, ORAPARM_INPUT)
        params("pStopDate").serverType = ORATYPE_DATE

        params.Add("pStopShift", StrConv(aEndShift, vbUpperCase), ORAPARM_INPUT)
        params("pStopShift").serverType = ORATYPE_VARCHAR2

        params.Add("pRoundVal", aBplRound, ORAPARM_INPUT)
        params("pRoundVal").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_massbal_ext_msr_spec
        'pMineName          IN     VARCHAR2,
        'pStartDate         IN     DATE,
        'pStartShift        IN     VARCHAR2,
        'pStopDate          IN     DATE,
        'pStopShift         IN     VARCHAR2,
        'pRoundVal          IN     NUMBER,
        'pResult            IN OUT c_massbalavg)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_floatplant2.get_massbal_ext_msr_spec(:pMineName, " + _
                      ":pStartDate, :pStartShift, :pStopDate, :pStopShift, :pRoundVal, :pResult);end;", ORASQL_FAILEXEC)
        CircAvgDynaset = params("pResult").Value
        ClearParams(params)

        RecCount = CircAvgDynaset.RecordCount

        'Assign the average data to mMbFcShift
        Do While Not CircAvgDynaset.EOF
            ThisEqpt = CircAvgDynaset.Fields("eqpt_name").Value
            ThisMsrName = CircAvgDynaset.Fields("measure_name").Value
            ThisAvgVal = CircAvgDynaset.Fields("avg_val").Value

            With mMbFcShift
                Select Case ThisEqpt
                    Case Is = "North fine rougher"
                        'Feed BPL, Concentrate BPL and Tail BPL
                        If ThisMsrName = "Feed BPL" Then
                            .NfrFdBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Concentrate BPL" Then
                            .NfrCnBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Tail BPL" Then
                            .NfrTlBplRpt = ThisAvgVal
                        End If

                    Case Is = "South fine rougher"
                        'Feed BPL, Concentrate BPL and Tail BPL
                        If ThisMsrName = "Feed BPL" Then
                            .SfrFdBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Concentrate BPL" Then
                            .SfrCnBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Tail BPL" Then
                            .SfrTlBplRpt = ThisAvgVal
                        End If

                    Case Is = "North coarse rougher"
                        'Feed BPL, Tail BPL
                        If ThisMsrName = "Feed BPL" Then
                            .NcrFdBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Tail BPL" Then
                            .NcrTlBplRpt = ThisAvgVal
                        End If

                    Case Is = "South coarse rougher"
                        'Feed BPL, Tail BPL
                        If ThisMsrName = "Feed BPL" Then
                            .ScrFdBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Tail BPL" Then
                            .ScrTlBplRpt = ThisAvgVal
                        End If

                    Case Is = "Coarse rougher"
                        'Concentrate BPL
                        If ThisMsrName = "Concentrate BPL" Then
                            .CrCnBplRpt = ThisAvgVal
                        End If

                    Case Is = "Coarse column"
                        'Feed BPL, Concentrate BPL and Tail BPL
                        If ThisMsrName = "Feed BPL" Then
                            .CrsColFdBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Concentrate BPL" Then
                            .CrsColCnBplRpt = ThisAvgVal
                        End If
                        If ThisMsrName = "Tail BPL" Then
                            .CrsColTlBplRpt = ThisAvgVal
                        End If

                    Case Is = "Fine amine"
                        'Tail BPL
                        If ThisMsrName = "Tail BPL" Then
                            .FaTlBplRpt = ThisAvgVal
                        End If

                    Case Is = "Coarse amine"
                        'Tail BPL
                        If ThisMsrName = "Tail BPL" Then
                            .CaTlBplRpt = ThisAvgVal
                        End If

                    Case Is = "Float plant"
                        'Tail BPL
                        If ThisMsrName = "Tail BPL" Then
                            .GmtBplRpt = ThisAvgVal
                        End If
                End Select
            End With
            CircAvgDynaset.MoveNext()
        Loop

        CircAvgDynaset.Close()

        Exit Sub

GetCircBplAvgsError:

        MsgBox("Error getting period average data." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Four Corners Period Average Error")

        On Error Resume Next
        CircAvgDynaset.Close()
        ClearParams(params)
    End Sub

    Public Function gGetFcFloatPlantBalanceData2(ByRef FloatPlantCirc As Object, _
                                                 ByRef FloatPlantGmt As Object, _
                                                 ByVal aBeginDate As Date, _
                                                 ByVal aBeginShift As String, _
                                                 ByVal aEndDate As Date, _
                                                 ByVal aEndShift As String, _
                                                 ByVal aCrewNumber As String, _
                                                 ByVal aBplRound As Integer, _
                                                 ByVal aMassBalMode As String) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetFcFloatPlantBalanceDataError

        'This function will return the number of shifts processed.
        'It will also "return" data through the FloatPlantCirc() and
        'FloatPlantGmt() arrays.

        'Circuit  Circuit  Circuit  Circuit  Circuit  Circuit
        'Circuit  Circuit  Circuit  Circuit  Circuit  Circuit
        'Circuit  Circuit  Circuit  Circuit  Circuit  Circuit

        'aMassBalMode should be "Circuit"

        Dim RowIdx As Integer
        Dim ColIdx As Integer

        Dim CalcNumShifts As Integer
        Dim ActualNumshifts As Integer

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        Dim RecCount As Long

        Dim ThisDate As Date
        Dim ThisShift As String
        Dim ThisEqpt As String

        Dim NumShifts As Integer

        mBeginDate = aBeginDate
        mBeginShift = aBeginShift
        mEndDate = aEndDate
        mEndShift = aEndShift

        '12/17/2007, lss
        'This procedure wil not sum shifts -- it will average the circuit BPL's
        'and sum the reported feed and concentrate tons for the period
        'instead.

        CalcNumShifts = gGetNumShiftsRge2("Four Corners", _
                                          aBeginDate, _
                                          aEndDate)

        'Mass balances are only run for either one shift or for
        'a range of complete days.
        If aBeginDate = aEndDate And aBeginShift = aEndShift Then
            CalcNumShifts = 1
        End If

        'aCrewNumber will be "All", "A", "B", "C", or "D"
        If aCrewNumber = "All" Then
            NumShifts = CalcNumShifts
        Else
            'NumShifts = gGetCrewShiftCount("Four Corners", _
            '                               aBeginDate, _
            '                               aBeginShift, _
            '                               aEndDate, _
            '                               aEndShift, _
            '                               aCrewNumber)
        End If

        ReDim FloatPlantCirc(0 To 9, 0 To 14)
        ReDim FloatPlantGmt(0 To 4, 0 To 7)

        'fFloatPlantCirc
        '---------------
        '
        '       Rows                 Columns
        '       --------------       ----------------
        ' 1)    Fine rougher         Hours
        ' 2)    Fine amine           Feed tons reported
        ' 3)    Total fine           Feed tons adjusted
        ' 4)    Coarse rougher       Feed BPL
        ' 5)    Coarse amine         Conc BPL
        ' 6)    Total coarse         Tail BPL
        ' 7)    Total amine          Ratio of concentration
        ' 8)    Grand totals         %Actual recovery
        ' 9)    Concentrate product  %Standard recovery
        '10)                         Concentrate tons adjusted
        '11)                         Tail tons adjusted
        '12)                         Feed TPH
        '13)                         Concentrate TPH
        '14)                         Tail TPH

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

        For RowIdx = 1 To 9
            For ColIdx = 1 To 14
                FloatPlantCirc(RowIdx, ColIdx) = 0
            Next ColIdx
        Next RowIdx

        For RowIdx = 1 To 4
            For ColIdx = 1 To 7
                FloatPlantGmt(RowIdx, ColIdx) = 0
            Next ColIdx
        Next RowIdx

        ZeroFcSummingData()

        'Will average the circuit BPL's and sum the reported feed and
        'concentrate tons for the period instead.
        'The data will be in mMbFcShift
        ZeroFcShiftData()

        GetPeriodAvgsAndSums(aBeginDate, _
                             aBeginShift, _
                             aEndDate, _
                             aEndShift, _
                             aBplRound, _
                             aMassBalMode)

        ProcessFcMassBalanceData2(aBplRound)

        'Summing of mass balance shift data completed
        ProcessFcMassBalanceTotals(aBplRound)

        'Place data in array  Place data in array  Place data in array
        'Place data in array  Place data in array  Place data in array
        'Place data in array  Place data in array  Place data in array

        'FloatPlantCirc()

        'Rows in the array                Columns in the array
        'fcFneRghr = 1                    fcCcOperHrs = 1
        'fcFneAmine = 2                   fcCcFdTonsRpt = 2
        'fcTotFne = 3                     fcCcFdTonsAdj = 3
        'fcCrsRghr = 4                    fcCcFdBpl = 4
        'fcCrsAmine = 5                   fcCcCnBpl = 5
        'fcTotCrs = 6                     fcCcTlBpl = 6
        'fcTotAmine = 7                   fcCcRC = 7
        'fcTotPlant = 8                   fcCcPctActRcvry = 8
        'fcCrCnProduct = 9                fcCcPctStdRcvry = 9
        '                                 fcCcCnTonsAdj = 10
        '                                 fcCcTlTonsAdj = 11
        '                                 fcCcFdTph = 12
        '                                 fcCcCnTph = 13
        '                                 fcCcTlTph = 14

        'Product tons
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCnProduct, mFcFloatPlantCircColEnum.fcCcCnTonsAdj) = mMbFcTotal.PrdCnTons
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCnProduct, mFcFloatPlantCircColEnum.fcCcCnBpl) = mMbFcTotal.PrdCnBpl

        'Operating hours
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcOperHrs) = mMbFcTotal.FrHrs
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcOperHrs) = mMbFcTotal.CrHrs
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcOperHrs) = mMbFcTotal.FaHrs
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcOperHrs) = mMbFcTotal.CaHrs
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcOperHrs) = mMbFcTotal.TotPltHrs

        'Feed tons reported
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcFdTonsRpt) = mMbFcTotal.FrFdTonsRpt
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcFdTonsRpt) = mMbFcTotal.CrFdTonsRpt
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcFdTonsRpt) = mMbFcTotal.FaFdTonsRpt
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcFdTonsRpt) = mMbFcTotal.CaFdTonsRpt
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcFdTonsRpt) = mMbFcTotal.TotPltFdTonsRpt

        'Feed tons adjusted
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcFdTonsAdj) = mMbFcTotal.FrFdTonsAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcFdTonsAdj) = mMbFcTotal.CrFdTonsAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcFdTonsAdj) = mMbFcTotal.FaFdTonsAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcFdTonsAdj) = mMbFcTotal.CaFdTonsAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcFdTonsAdj) = mMbFcTotal.TotPltFdTonsAdj

        'Feed BPL
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcFdBpl) = mMbFcTotal.FrFdBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcFdBpl) = mMbFcTotal.CrFdBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcFdBpl) = mMbFcTotal.FaFdBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcFdBpl) = mMbFcTotal.CaFdBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcFdBpl) = mMbFcTotal.TotPltFdBplAdj

        'Concentrate BPL
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcCnBpl) = mMbFcTotal.FrCnBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcCnBpl) = mMbFcTotal.CrCnBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcCnBpl) = mMbFcTotal.FaCnBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcCnBpl) = mMbFcTotal.CaCnBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcCnBpl) = mMbFcTotal.TotPltCnBplAdj

        'Tail BPL
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcTlBpl) = mMbFcTotal.FrTlBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcTlBpl) = mMbFcTotal.CrTlBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcTlBpl) = mMbFcTotal.FaTlBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcTlBpl) = mMbFcTotal.CaTlBplAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcTlBpl) = mMbFcTotal.TotPltTlBplAdj

        'Ratio of concentration
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcRc) = mMbFcTotal.FrRcAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcRc) = mMbFcTotal.CrRcAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcRc) = mMbFcTotal.FaRcAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcRc) = mMbFcTotal.CaRcAdj
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcRc) = mMbFcTotal.TotPltRcAdj

        'Actual recovery
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcPctActRcvry) = mMbFcTotal.FrPctAdjRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcPctActRcvry) = mMbFcTotal.CrPctAdjRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcPctActRcvry) = mMbFcTotal.FaPctAdjRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcPctActRcvry) = mMbFcTotal.CaPctAdjRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcPctActRcvry) = mMbFcTotal.TotPltPctAdjRcvry

        'Reported recovery
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneRghr, mFcFloatPlantCircColEnum.fcCcPctRptRcvry) = mMbFcTotal.FrPctRptRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsRghr, mFcFloatPlantCircColEnum.fcCcPctRptRcvry) = mMbFcTotal.CrPctRptRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrFneAmine, mFcFloatPlantCircColEnum.fcCcPctRptRcvry) = mMbFcTotal.FaPctRptRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrCrsAmine, mFcFloatPlantCircColEnum.fcCcPctRptRcvry) = mMbFcTotal.CaPctRptRcvry
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcPctRptRcvry) = mMbFcTotal.TotPltPctRptRcvry

        'Concentrate tons adjusted
        'Nothing to put in here right now.

        'Tail tons adjusted
        'Nothing to put in here right now.

        'Feed TPH adjusted
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcFdTph) = mMbFcTotal.TotPltFdTphAdj

        'Concentrate TPH adjusted
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcCnTph) = mMbFcTotal.TotPltCnTphAdj

        'Tail TPH adjusted
        FloatPlantCirc(mFcFloatPlantCircRowEnum.fcCrTotPlant, mFcFloatPlantCircColEnum.fcCcTlTph) = mMbFcTotal.TotPltTlTphAdj

        'FloatPlantGmt

        'Rows in the array              Columns in the array
        'grAsReportedGmtBpl = 1         gcFdTons = 1
        'grCalculatedGmtBpl = 2         gcCnTons = 2
        'grReportedFdTons = 3           gcFdBpl = 3
        'grGmtBplFromCircuits = 4       gcCnBpl = 4
        '                               gcTlBpl = 5
        '                               gcRC = 6
        '                               gcPctRcvry = 7

        'Based on as reported GMT BPL
        'Only interested in an as reported Gmt BPL value here
        'We have two total plant tail BPL's we could use here
        '1) mMbFcTotal.TotPltTlBplRpt    Based on measured circuit tail BPL's
        '2) mMbFcTotal.TotPltTlBplRpt2   Based on total plant measured GMT

        'Will use mMbFcTotal.TotPltTlBplRpt for now.

        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrAsReportedGmtBpl, mWgFloatPlantGmtColEnum.fcGcFdBpl) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrAsReportedGmtBpl, mWgFloatPlantGmtColEnum.fcGcCnBpl) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrAsReportedGmtBpl, mWgFloatPlantGmtColEnum.fcGcTlBpl) = mMbFcTotal.TotPltTlBplRpt
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrAsReportedGmtBpl, mWgFloatPlantGmtColEnum.fcGcRc) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrAsReportedGmtBpl, mWgFloatPlantGmtColEnum.fcGcPctRcvry) = 0

        'Based on Adjusted or Calculated feed tons
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcFdTons) = mMbFcTotal.TotPltFdTonsAdj
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcCnTons) = mMbFcTotal.TotPltCnTonsAdj
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcFdBpl) = mMbFcTotal.TotPltFdBplAdj
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcCnBpl) = mMbFcTotal.TotPltCnBplAdj
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcTlBpl) = mMbFcTotal.TotPltTlBplAdj
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcRc) = mMbFcTotal.TotPltRcAdj
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrCalculatedGmtBpl, mWgFloatPlantGmtColEnum.fcGcPctRcvry) = mMbFcTotal.TotPltPctAdjRcvry

        'Based on reported feed tons.
        'Will not put anything here for Four Corners right now (will just
        'fill in zeros).
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcFdTons) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcCnTons) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcFdBpl) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcCnBpl) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcTlBpl) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcRc) = 0
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrReportedFdTons, mWgFloatPlantGmtColEnum.fcGcPctRcvry) = 0

        'This will not really apply at Four Corners at this time (will just
        'assign a zero).
        FloatPlantGmt(mFcFloatPlantGmtRowEnum.fcGrGmtBplFromCircuits, mWgFloatPlantGmtColEnum.fcGcTlBpl) = 0

        gGetFcFloatPlantBalanceData2 = NumShifts

        Exit Function

gGetFcFloatPlantBalanceDataError:

        MsgBox("Error calculating Four Corners Mass Balance." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Four Corners Mass Balance Calculation Error")

        On Error Resume Next
        ClearParams(params)
    End Function

    Private Sub ProcessFcMassBalanceData2(ByVal aBplRound As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo ProcessFcMassBalanceDataError

        Dim ShiftFrHrs As Single
        Dim Rc As Single

        'IMPORTANT:  We are making the assumption that there are no
        '            missing BPL's.  If tons exist then there should be
        '            corresponding BPL's

        'Compute the Coarse amine feed BPL -- average the North fine amine and
        'the South fine amine feed BPL's.

        'In mMbFcShift have:

        ' 1) PrdCnBpl
        ' 2) PrdCnTons
        '------
        ' 3) Nfr1FdTonsRpt
        ' 4) Nfr2FdTonsRpt
        ' 5) Sfr1FdTonsRpt
        ' 6) Sfr2FdTonsRpt
        ' 7) NcrFdTonsRpt
        ' 8) ScrFdTonsRpt
        '-----
        ' 9) NfrFdBplRpt
        '10) NfrCnBplRpt
        '11) Nfr1TlBplRpt
        '12) Nfr2TlBplRpt
        '-----
        '13) SfrFdBplRpt
        '14) SfrCnBplRpt
        '15) Sfr1TlBplRpt
        '16) Sfr2TlBplRpt
        '-----
        '17) NcrFdBplRpt
        '18) NcrCnBplRpt
        '19) NcrTlBplRpt
        '-----
        '20) ScrFdBplRpt
        '21) ScrCnBplRpt
        '22) ScrTlBplRpt
        '-----
        '23) NfaTlBplRpt
        '24) SfaTlBplRpt
        '25) NcaTlBplRpt
        '26) ScaTlBplRpt

        With mMbFcTotal
            'Concentrate product tons (final concentrate product)
            .PrdCnTons = mMbFcShift.PrdCnTons
            .PrdCnBpl = mMbFcShift.PrdCnBpl

            'Operating hours
            'Fine rougher hours
            .Nfr1Hrs = mMbFcShift.Nfr1Hrs     'North fine rougher 1
            .Nfr2Hrs = mMbFcShift.Nfr2Hrs     'North fine rougher 2
            .Sfr1Hrs = mMbFcShift.Sfr1Hrs     'South fine rougher 1
            .Sfr2Hrs = mMbFcShift.Sfr2Hrs     'South fine rougher 2

            'Average all 4 fine rougher sections to get the fine rougher
            'operating hours
            .FrHrs = Round((mMbFcShift.Nfr1Hrs + _
                            mMbFcShift.Nfr2Hrs + _
                            mMbFcShift.Sfr1Hrs + _
                            mMbFcShift.Sfr2Hrs) / 4, 2)

            'Assume that the fine amine operating hours are the same as the
            'fine rougher operating hours
            .FaHrs = .FrHrs

            'Coarse rougher hours
            .NcrHrs = mMbFcShift.NcrHrs       'North coarse rougher
            .ScrHrs = mMbFcShift.ScrHrs       'South coarse rougher

            'Coarse rougher hrs where the coarse amine ran
            .NcrCaHrs = mMbFcShift.NcrCaHrs   'North coarse rougher
            .ScrCaHrs = mMbFcShift.ScrCaHrs   'South coarse rougher

            'Average the two coarse rougher sections to get the coarse rougher
            'operating hours
            .CrHrs = Round((mMbFcShift.NcrHrs + mMbFcShift.ScrHrs) / 2, 2)

            .CaHrs = Round((mMbFcShift.NcrCaHrs + mMbFcShift.ScrCaHrs) / 2, 2)

            .NcsHrs = mMbFcShift.NcsHrs       'North coarse scalp -- Not really used here
            .ScsHrs = mMbFcShift.ScsHrs       'South coarse scalp -- Not really used here

            'Plant operating hours = (2 * fine rougher hours + coarse rougher hours) / 3
            .TotPltHrs = Round((2 * .FrHrs + .CrHrs) / 3, 2)
        End With

        Exit Sub

ProcessFcMassBalanceDataError:

        MsgBox("Error in Four Corners mass balance." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Four Corners Mass Balance Computation Error")
    End Sub

End Module
