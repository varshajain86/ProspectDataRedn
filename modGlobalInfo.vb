Option Explicit On
Imports System.Math
Imports OracleInProcServer
Imports Microsoft.VisualBasic
Imports AxFPSpread
Imports FPSpread
Module modGlobalInfo
    ' Attribute VB_Name = "modGlobalInfo"
'**********************************************************************
'GLOBAL VARIABLES MODULE
'
'
    '**********************************************************************
    Public Const ORASQL_NO_AUTOBIND = &H1& 'Do not perform automatic binding of database parameters. 
    Public Const ORASQL_FAILEXEC = &H2& 'Raise error and do not create SQL statement object. 
    Public Const ORASQL_NONBLK = &H4& 'Execute SQL in a nonblocking state. 
    Public Const ORADYN_DEFAULT = 0
    Public Const ORAPARM_INPUT = 1 'Used for input variables only 
    Public Const ORAPARM_OUTPUT = 2 'Used for output variables only 
    Public Const ORAPARM_BOTH = 3 'Used for variables that are both input and output 
    Public Const ORATYPE_VARCHAR2 = 1 'VARCHAR2 
    Public Const ORATYPE_NUMBER = 2 'NUMBER 
    Public Const ORATYPE_SINT = 3 'SIGNED INTEGER 
    Public Const ORATYPE_FLOAT = 4 'FLOAT 
    Public Const ORATYPE_STRING = 5 'Null Terminated STRING 
    Public Const ORATYPE_LONG = 8 'LONG 
    Public Const ORATYPE_VARCHAR = 9 'VARCHAR 
    Public Const ORATYPE_DATE = 12 'DATE 
    Public Const ORATYPE_RAW = 23 'RAW 
    Public Const ORATYPE_LONGRAW = 24 'LONG RAW 
    Public Const ORATYPE_UINT = 68 'UNSIGNED INTEGER 
    Public Const ORATYPE_CHAR = 96 'CHAR 
    Public Const ORATYPE_CHARZ = 97 'Null Terminated CHAR 
    Public Const ORATYPE_BFLOAT = 100 'BINARY_FLOAT 
    Public Const ORATYPE_BDOUBLE = 101 ' BINARY_DOUBLE 
    Public Const ORATYPE_CURSOR = 102 ' PLSQL CURSOR 
    Public Const ORATYPE_MLSLABEL = 105 'MLSLABEL 
    Public Const ORATYPE_OBJECT = 108 'OBJECT 
    Public Const ORATYPE_REF = 110 'REF 
    Public Const ORATYPE_CLOB = 112 'CLOB 
    Public Const ORATYPE_BLOB = 113 'BLOB 
    Public Const ORATYPE_BFILE = 114 'BFILE 
    Public Const ORATYPE_TIMESTAMP = 187 'TIMESTAMP 
    Public Const ORATYPE_TIMESTAMPTZ = 188 ' TIMESTAMP WITH TIMEZONE 
    Public Const ORATYPE_INTERVALYM = 189 'INTERVAL YEAR TO MONTH 
    Public Const ORATYPE_INTERVALDS = 190 'INTERVAL DAY TO SECOND 
    Public Const ORATYPE_TIMESTAMPLTZ = 232 'TIMESTAMP WITH LOCAL TIME ZONE 
    Public Const ORATYPE_VARRAY = 247 'VARRAY 
    Public Const ORATYPE_TABLE = 248 'NESTED TABLE 
    Public Const ORATYPE_RAW_BIN = 2000 'RAW 
'
'Special Comments
'----------------
'This is a partial listing of the functions in this module:
'
'    Function Name           Parameters                             What it returns                           Type Returned
'    ----------------------  -------------------------------------  ----------------------------------------  ------------
' 1) GetDaysInMonth          aMonth, aYear                          Days in month                             Integer
' 2) gGetDaysInMonth2        aDate                                  Days in month                             Integer
' 3) gGetDOW                 aDate                                  Name of the day of the week               String
' 4) gRoundFifty             aNumber                                Number rounded to fifty                   Long
' 5) gGetFirstOfMonth        aDate                                  First day of month                        Date
' 6) gGetLastOfMonth         aDate                                  Last day of month                         Date
' 7) gGetFrozenStatus        aDate                                  Frozen status ("Hard", "Soft", "None")    String
' 8) gGetCalYrBegin          aDate                                  Calendar year begin date                  Date
' 9) gGetCalYrEnd            aDate                                  Calendar year end date                    Date
'10) gGetFiscYrBegin         aDate                                  Fiscal year begin date                    Date
'11) gGetFiscYrEnd           aDate                                  Fiscal year end date                      Date
'12) gGetFirstDayOfWeek      aDate                                  First date in week                        Date
'13) gGetLastDayOfWeek       aDate                                  Last date in week                         Date
'14) LastDayInMonth          aDate                                  True/False -- last day in month?          Boolean
'15) gLastDayInCalendarYear  aDate                                  True/False -- last day in calendar year?  Boolean
'16) gLastDayInFiscalYear    aDate                                  True/False -- last day in fiscal year?    Boolean
'17) gConvertToShiftDate     sDate, sShift                          Shift date                                Date
'18) gGetEndOfFiscalYear     aDate                                  Last date in fiscal year                  Date
'19) gGetBeginOfFiscalYear   aDate                                  First date in fiscal year                 Date
'20) gGetSupervisor          aDate, aShift                          Field supervisor's name                   String
'21) gGetNumUnits            aMineName, aReport, aProdDate, aShift  Number of shift report units              Integer
'22) gGetMonthAndYear        aDate, aIncludeSpace                   Month & year (January, 1995, etc.)        String
'23) gPadLeft                aString, aLength                       Left blank padded string                  String
'24) gPadRight               aString, aLength                       Right blank padded string                 String
'25) gIsEvenNumber           aValue                                 True/False -- number is even?             Boolean
'26) gPasswordIsValid        aPassword, aUserName                   Checks password for validity              String
'27) gGetMonthDate           aDate                                  Creates "mm/yyyy" from date               String
'28) gUserExists             aMineName, aUserName                   True/False -- user set up for mine?       Boolean
'29) gGetBeginOfCalendarYear aDate                                  Calendar year begin date                  Date
'30) gUserIsAdministrator    aMineName, aUserName                   True/False -- user is an Administrator    Boolean
'31) gIsLegalShift           aShiftName                             True/False -- shift name is legal         Boolean
'32) gGetMilitaryTime        aAmPmTime                              Military time                             String
'33) gGetMoisBeginDate       aMineName                              MOIS start date for a mine                Date
'34) gGetMonthsInRange       aBeginDate, aEndDate                   Number of months in range                 Integer
'35) gGetWeeksInRange        aBeginDate, aEndDate                   Number of weeks in range                  Integer
'36) gGetDateTime            aDate, aTime                           Date time                                 Date
'37) gGetFirstShift          aMineName                              Name of first shift for a mine            String
'38) gQtrBeginDateCurr       aDate                                  First date in the quarter                 Date
'39) gQtrEndDateCurr         aDate                                  Last date in the quarter                  Date
'40) gQtrBeginDate           aDate, aQuarter                        Begin date for the quarter                Date
'41) gGetShiftBegDtime       aDate, aBeginHour                      Creates shift date                        Date
'42) gGetFirstShiftBegDtime  aDate                                  First shift shift date                    Date
'43) gGetLastShiftBegDtime   aDate                                  Last shift shift date                     Date
'44) gGetShiftStartHour      aShift                                 Hour that shift begins                    Integer
'45) gLastDayInQuarter       aDate                                  Last date in quarter (1st,2nd,3rd,4th)   String
'
'**********************************************************************
'   Maintenance Log
'
'   10/20/2000, lss
'       Modified gRound -- wasn't handling rounding correctly.
'   10/23/2000, lss
'       Added gUserShiftSelectedDate.
'   10/30/2000, lss
'       Added gLastUserSelectedDate.
'   10/30/2000, lss
'       Added gHole2SplitsDynaset and gHole2CompositeDynaset.
'   11/20/2000, lss
'       Added maps to gUserPermissions.
'   06/18/2001, lss
'       Added gActiveDragline.
'   08/17/2001, lss
'       Added Utility Operator Report & DL Inspection Report to
'       gUserPermissions.
'   08/30/2001, lss
'       Added Pump Inspection Report to gUserPermissions.
'   08/31/2001, lss
'       Added gSrptUserName and gSrptPassword.
'   09/03/2001, lss
'       Added gSrptUorOk, gSrptDlOk, gSrptPumpOk, gSrptCancel.
'   09/24/2001, lss
'       Added Piezometer stuff to gUserPermissions.
'   11/21/2001, lss
'       Modified gGetNumUnits.
'   11/28/2001, lss
'       Added Reclamation Activity stuff to gUserPermissions.
'   12/18/2001, lss
'       Changed GetDOW to gGetDOW.
'   01/09/2002, lss
'       Added gGetDaysInMonth2.
'   02/18/2002, lss
'       Modified gGetMonthAndYear.
'   03/22/2002, lss
'       Added global "grid to text" file variables.
'   03/26/2002, lss
'       gPrintGridSubHeader1 & gPrintGridSubHeader2.
'   04/02/2002, lss
'       Added global variables for "One number form".
'       Added global variables for "One string form".
'   04/04/2002, lss
'       Added gPrintGridMineName.
'   05/20/2002, lss
'       Added gAcadFileName.
'   05/29/2002, lss
'       Added Surfer related items.
'   06/05/2002, lss
'       Added gOrientSubHeaders.
'       Added Grid to graph related items.
'       Added gPadLeft.
'   06/11/2002, lss
'       Added Product to BigCalendarGraph.
'       Added gByVoloView, gByInViso, gByAutoCad.
'   06/13/2002, lss
'       Added gPadRight.   Grid to graph related items
'   08/15/2002, lss
'       Added more Grid to graph related items  **NEW**.
'   08/27/2000, lss
'       Added Water Samples and Pipe Thickness to gUserPermissions.
'   10/30/2002, lss
'       Added gLegendFontSize, gAxisTitleFontSize, gLabelFontSize to
'       grid to graph related items.
'   11/27/2002, lss
'       Added gGraphMean for grid to graph.
'   12/16/2002, lss
'       Added items to grid to graph.
'   01/03/2003, lss
'       Added dragline prospect stuff.
'   02/21/2003, lss
'       Added decision grid to gSetUserPermissions.
'   04/15/2003, lss
'       Added gOracleUserName, gOraclePassword.
'   05/29/2003, lss
'       Modified gUserPermission for WasherShiftReport,
'       FloatPlantShiftReport, SizingShiftReport, ReagentShiftReport,
'       ShippingShiftReport and SrptSpvsrChkOff.
'   06/10/2003, lss
'       Added function gSrptSetup() and gSrptSetupOrRead().
'   06/11/2003, lss
'       Added gSrptThisRptWriteOk() and gSrptThisCtgryName().
'   06/20/2003, lss
'       Added gSrptWasherOk, gSrptFltPltOk, gSrptSizingOk,
'       gSrptReagentOk, gSrptShippingOk.
'       Added gSetSrptGlobalPermissions().
'   06/30/2003, lss
'       Added function gGetSrptSupervisor().
'   07/01/2003, lss
'       Added function gGetSftyCommCount().
'   08/19/2003, lss
'       Fixed gGetSrptSupervisor for "Day" and "Day shift".
'   10/02/2003, lss
'       Added gTxtDefaultFile.
'   11/04/2003, lss
'       Changed gRoundFifty -- TempNumberFra <= 99 to
'                              TempNumberFra <= 99.99
'       Changed gRoundFifty(aNumber As Single) As Long to
'               gRoundFifty(aNumber As Variant) As Long
'   12/16/2003, lss
'       Added gGetMonthDate.
'   01/05/2004, lss
'       Added gUserExists.
'   01/13/2004, lss
'       Added gGetSpecialPath.
'   01/20/2004, lss
'       Added gGetBeginOfCalendarYear.
'   01/28/2004, lss
'       Added gAcadMacroPath.
'   02/17/2004, lss
'       Modified gRound -- parameters passed ByVal.
'       Added gPrintMarginLeft, gPrintMarginRight, gPrintMarginTop and
'       gPrintMarginBottom.
'   02/18/2004, lss
'       Added gClearGridPrint.
'   03/24/2004, lss
'       Added Absentees functionality for user permissions.
'   05/17/2004, lss
'       Added gLogonMine, gPipeTracking, gMineProspect, gCompanyName.
'   05/20/2004, lss
'       Added gProdDateAccessTime, gProdDateCreateTime,
'       gSampDayShiftBeginTime and gSampDayShiftEndTime.
'   06/10/2004, lss
'       Added gSzdFdTpohAsWhole.
'   06/14/2004, lss
'       Changed gDlDetail(7, 16) to gDlDetail(10, 16).
'   06/12/2004, lss
'       Added gWasherPctOperMode, gSizingPctOperMode,
'       gFltPltPctOperMode.
'   06/15/2004, lss
'       Modified type RockBooks -- added ActSocFltPltOh,
'       and ActSocWasherOh.
'   06/16/2004, lss
'       Added gSetActiveDateAndShift() -- moved it from frmMain.
'   07/28/2004, lss
'       Added gUserIsAdministrator().
'   07/30/2004, lss
'       Modified gSetActiveDateAndShift so that it is not Day, Night
'       shift dependent.  Added gGetShiftNames().  Added
'       SetActiveShiftAndDate().
'   08/04/2004, lss
'       Added gGetPrevShift() and gIsLegalShift().
'   08/09/2004, lss
'       Removed default mine stuff from gSetUserPermissions()
'       (SFdefault, HPdefault, FMdefault).
'   08/11/2004, lss
'       Added gGoToShift1, gGoToShift2 and gGoToShift3.
'   08/17/2004, lss
'       Added Function gGetMilitaryTime(aAmPmTime As String).
'   08/30/2004, lss
'       Added gMoisBeginDate, gHasDredges and gReagentByDay.
'       Added gGetMoisBeginDate().
'   09/01/2004, lss
'       Added gGetMonthsInRange() and gGetWeeksInRange().
'   09/08/2004, lss
'       Added gGetDateTime().
'   09/09/2004, lss
'       Modified type RockBooks -- removed ActSocFltPltOh,
'       and ActSocWasherOh.  Added ActNumWasherSides and
'       ActWasherSidesOh.
'   09/09/2004, lss
'       Added gNumWasherSides.
'   09/21/2004, lss
'       Added gGetFirstShift().
'   09/30/2004, lss
'       Added gGetMiscMineGlobals().
'   10/06/2004, lss
'       Modified gSrptThisCtgryName for dredges.
'       Modified gSrptThisRptWriteOk for dredges.
'       Modified gSetSrptGlobalPermissions for dredges.
'   10/25/2004, lss
'       Added gQtrBeginDateCurr(), gQtrBeginDate() and
'       gQtrEndDateCurr().  Added gGetFirstShiftBegDtime() and
'       gGetFirstShiftBegDtime().
'   10/26/2004, lss
'       Added gGetShiftBegDtime() and gGetShiftStartHour().
'       Changed LastDayInFiscalYear to gLastDayInFiscalYear().
'       Added gLastDayInQuarter().
'       Changed CalculateTimeFrame() to gCalculateTimeFrame().
'       Changed LastDayInCalendarYear to gLastDayInCalendarYear().
'   10/28/2004, ls
'       Added "Public gFyrBeginDate As Integer".
'       Removed references to "06/01" -- gGetFiscYrBegin,
'       gGetFiscYrEnd, gGetBeginOfFiscalYear, gGetEndOfFiscalYear.
'   11/10/2004, lss
'       Added gGetLastShift().
'   11/15/2004, lss
'       Added gGetAllShiftsCbo().
'   11/19/2004, lss
'       Added gPrintGridSubHeader3, gPrintGridFooter2, gOrientFooter2,
'       gOrientFooter.
'   12/02/2004, lss
'       Added "Setup" to Pipe Thickness Tracking in gUserPermissions.
'   12/05/2004, lss
'       Added gResizeSsCols().
'   12/08/2004, lss
'       Added  gEveryOtherGreen().
'   12/16/2004, lss
'       Added gPrintOrientation, added it to gClearGridPrint.
'       Added gLabelEvery and gGraphLblData.
'   12/30/2004, lss
'       Changed gGraphLoctnName to gGraphLoctnName().
'   01/13/2005, lss
'       Added function gGetChoices().
'   01/14/2005, lss
'       Added function gGetAllMineNamesDyn() and gGetAllMineNamesCbo().
'       Added function gGetTypes().
'   02/09/2005, lss
'       Added gPct100ProspDesc.
'   02/10/2005, lss
'       Added gGetFirstOfNextMonth.
'   02/24/2005, lss
'       Added gShiftLength.
'   02/25/2005, lss
'       Added gDlMtxBktFillFctr.
'   03/07/2005, lss
'       Added gDlHasMtxTons and gDlMtxYdsSource.
'   03/10/2005, lss
'       Added gGetEqptTypeName and gGetProdMatlTypeName.
'   03/18/2005, lss
'       Added Subroutine gSaveVbError.
'   03/28/2005, lss
'       Added gProspGridType, gWasherInput, gProspCnInsAdj and
'       gCalcSplitConc.
'   04/04/2005, lss
'       Added Function gGetDlName().
'   04/06/2005, lss
'       Added gFltPltFdTonsMsr and gSzgFdTonsMsr.
'       Added gMineHasPb, gMineHasCn and gMineHasIp.
'       Added Sub gGetMineProducts().
'   04/11/2005, lss
'       Added gShowIpPlusPb.
'   04/12/2005, lss
'       Added gProdInfoType.
'       Added gMineHasSubProducts Sub.
'       Added gMineHasSubProds Global
'   04/13/2005, lss
'       Added Function gGetMineHasIp().
'   04/19/2005, lss
'       Added Function gGetNumShifts().
'       Changed gGetShiftNames() to gSetGlobalShiftNames().
'       Added new Function gGetShiftNames().
'       Added Function gGetDlHasMtxTons and Function
'       gGetDlMtxYdsSource().
'   04/21/2005, lss
'       Added gProspPbIsPbIp.
'   04/25/2005, lss
'       Added Function gGetProspGridType().
'   05/06/2005, lss
'       Added Function gGetCalcSplitConc().
'       Added Functions gRoundTen() and gRoundFive().
'   05/09/2005, lss
'       Added gProdTonsRound.
'   05/12/2005, lss
'       Added Function gGetShiftForDateTime().
'   05/18/2005, lss
'       Added Sub SetActiveShiftAndDate().
'   06/06/2005, lss
'       Modified gSetUserPermissions for
'       gUserPermission.SafetyMeetings.
'   07/01/2005, lss
'       Added "prosp pb is pbip" to gGetMiscMineGlobals.
'   07/06/2005, lss
'       Added Function gGetProspStandard.
'       Added gDragProspStandard.
'   07/12/2005, lss
'       Added CATALOG items to type RockBooks.
'       Added gHasCtlgReserves and gInvAdjAppliedShifts.
'       Added Function gGetHasCtlgReserves.
'       Added Function gGetInvAdjAppliedShifts.
'   07/13/2005, lss
'       100%MINEABLE --> 100%PROSPECT.
'   07/18/2005, lss
'       Added gProspStandard.
'   07/19/2005, lss
'       Added gCatalogDesc.
'   07/27/2005, lss
'       Added PumpPackShiftReport to Type UserPermission.
'       Modified Function gSetUserPermissions for PumpPackShiftReport.
'       Modified Function gSrptSetup() for PumpPackShiftReport.
'       Modified Function gSrptSetupOrRead() for PumpPackShiftReport.
'       Modified Function gSrptThisRptWriteOk() for PumpPackShiftReport.
'       Modified Function gSrptThisCtgryName() for
'       "Field Pump Pack Shift Report"
'       Modified gSetSrptGlobalPermissions() for gSrptPumpPackOk.
'   08/08/2005, lss
'       Added gGenExplainCaption and gGenExplainComment.
'       Added gGenDelayComment.
'       Added gMiscMoisSetup.
'   08/11/2005, lss
'       Added Function gGetBinNum.
'       Added Function gGetDateFromDateTime.
'   08/12/2005, lss
'       Added Function gGetMatlCorrection.
'   08/15/2005, lss
'       Added Overlay graph related items.
'   08/24/2005, lss
'       Added Function gGetPeriodicEqptMsrmnt.
'       Added Function gGetEqptCalcSum.
'       Added gPrtBeginShift and gPrtEndshift.  Added them to
'       Sub gCalculateTimeFrame.
'   08/25/2005, lss
'       Changed gCatalogDesc to gCatalogProspDesc.
'   08/26/2005, lss
'       Added gMassBalanceMode.
'   08/29/2005, lss
'       Added gViewShift and gViewCrewNum.
'       Added gMassBalRpt.
'   09/01/2005, lss
'       Added Function gGet2NumAvg.
'   09/09/2005, lss
'       Added Function gGetDlNum.  Added Function gGetDlNameFromNum.
'   09/19/2005, lss
'       Changed gDlDetail(10, 16) to gDlDetail(10, 23).  Need to
'       capture interburden and prestrip acres and feet.
'   10/04/2005, lss
'       Added aProspectMinesOnly to gGetAllMineNamesCbo.
'   10/05/2005, lss
'       Added Function gGetMonthNum.
'   10/17/2005, lss
'       Added Function gGetYearsBackDate.
'       Added gGphAutoPrint.
'   10/21/2005, lss
'       Added gOlayRightTitle and gOlayRightTitleStyle.
'   10/25/2005, lss
'       Added "has catalog reserves" to gGetMiscMineGlobals.
'   10/26/2005, lss
'       Added gConvertToShiftDate2 -- it is the same as
'       gConvertToShiftDate only it receives the mine name as a
'       parameter (gConvertToShiftDate uses gActiveMineNameLong).
'   10/27/2005, lss
'       Changed gDlDetail(10, 23) to gDlDetail(10, 26).
'   11/28/2005, lss
'       Added gUseNewProdAnalCalc.
'   12/07/2005, lss
'       Added Public Function gGetMonthNum2.
'   12/14/2005, lss
'       Added "prod tons round" to gGetMiscMineGlobals.
'   12/15/2005, lss
'       Added "hard freeze date" and "soft freeze date"
'       to gGetMiscMineGlobals.
'       Added "has ctlg reserves" to gGetMiscMineGlobals.
'   12/27/2005, lss
'       Added Public Function gGetNextShift.
'   12/29/2005, lss
'       Added gProdAnalNumDaysBack, gProdAnalNumDaysForward
'   01/04/2006, lss
'       Added gUserPermission.InventoryAdjust.
'   01/19/2006, lss
'       Added IP to Function gGetProdMatlTypeName.
'   01/23/2006, lss
'       Added Public Function gGetMineNameShort.
'   01/25/2006, lss
'       Added Public Function gGetMassBalanceMethod.
'   02/20/2006, lss
'       Added Public Function gStrCharCount.
'   03/14/2006, lss
'       Added Public Type gProdMonthAdjType.
'   03/16/2006, lss
'       Added Public Function gGetFirstOfNextWeek.
'   03/20/2006, lss
'       Added Public Function gGetCircBplSpec.
'   03/24/2006, lss
'       Added Public Function gHasMultifosPotential.
'       Added gMultifosMinBpl, gMultifosTargIns, gMultifosMaxAl.
'   03/31/2006, lss
'       Added gGetPeriodicEqptMsrAvg.
'   04/04/2006, lss
'       Added Public Function gMoisProdToLimsProd.
'   04/04/2006, lss
'       Added Public Function gGetModifiedHour.
'       Added Public Function gGetInsAdjBpl.
'       Added Public Function gGetInsAdjAl.
'   04/06/2006, lss
'       Added Public Function gShiftGreater.
'   04/11/2006, lss
'       Added "catalog prosp desc" & "100% prosp desc" to
'       Function gGetMiscMineGlobals.
'   04/18/2006, lss
'       Added Function gGetProspStandardRev.
'   04/21/2006, lss
'       Added gHasDraglines.
'   04/27/2006, lss
'       Added Public Function gTransSapphireProd.
'       Added Public Function gTransSapphireEqpt.
'   04/28/2006, lss
'       Added Public Function gGetMoisDateForExtDate.
'   05/05/2006, lss
'       Modified Public Function gTransSapphireEqpt for 1 or 2 MOIS
'       equipment items for 1 Sapphire equipment item
'       ("SFM Bin 7 and 8", "SFM Bin 5 and 6")
'   05/15/2006, lss
'       Added "Pump yardages setup" to gSetUserPermissions.
'   06/06/2006, lss
'       Added Sub gAddEqptMsrmnt.
'       Added Function gCboBoxChoiceIsThere.
'       Added Function gGetShiftLengthForDate
'   06/08/2006, lss
'       Added Function gGetDlPitName.
'   06/26/2006, lss
'       Added "Absentees read" to gSetUserPermissions.
'   06/27/2006, lss
'       Added Function gGetMineUserPermissions.
'   06/28/2006, lss
'       Added Public Function gGetFirstShift2.
'       Added Public Function gGetLastShift2.
'   06/29/2006, lss
'       Added Public Function gGetEqptYds.
'   07/20/2006, lss
'       Added Public Sub gGetShiftDataForDate.
'       Added Public Function gGetFirstShiftBegDtime2.
'       Added Public Function gGetLastShiftBegDtime2.
'       Modified Public Sub gCalculateTimeFrame for 3 -> 2 shifts.
'   07/25/2006, lss
'       Added Public Sub gAddGridAvgs.
'       Added Public Sub gGetAllShiftsCbo2.
'   07/27/2006, lss
'       Public Function gGetShiftLengthForDate now compensates for
'       12/31/9999.
'   07/30/2006, lss
'       Added gGetPrevShift2 and gGetNextShift2.
'   08/07/2006, lss
'       Added Function gGetNumShifts2.
'   08/07/2006, lss
'       Added Public Function gGetNumShiftsRge().
'   08/09/2006, lss
'       Added Public gShiftChangeDate As Date
'       Added Public gHasShiftChangeDate As Boolean
'       Added "shift change date" to Function gGetMiscMineGlobals.
'   08/10/2006, lss
'       Added Public Function gGetNumShiftsRge2().
'   08/28/2006, lss
'       Added Global Const gMassBalanceCutOffDateHw = #1/1/2004#.
'   09/06/2006, lss
'       Added Public Function gGetMilitaryTimeSec.
'   09/07/2006, lss
'       Added Public Function gGetAllMatl.
'   10/10/2006, lss
'       Added Public Function gGetDateFromMonthAndYear.
'       Added Public Function gGetMonthAndYearAbbrv.
'   10/20/2006, lss
'       Added aNoComma to Function gGetMonthAndYearAbbrv.
'   10/27/2006, lss
'       Added gAutoPrint.
'   10/31/2006, lss
'       Added "IP Product" for Hopewell in Function
'       gTransSapphireProd.
'   11/10/2006, lss
'       Added gPadLeftChar and gPadRightChar.
'   12/11/2006, lss
'       Added Public Sub gGetMassBalDataAbbrv.
'   01/15/2007, lss
'       Added Raw prospect reduction to gUserPermissions.
'   03/07/2007, lss
'       Added Function gNeedToChangeShiftNames().
'   03/08/2007, lss
'       Added Function gGetLastShiftHardCode and Function
'       gGetFirstShiftHardCode.
'   03/09/2007, lss
'       Added RockBookRecalc and NoFrillsMois to gUserPermission.
'   03/16/2007, lss
'       Added gFormLoad and gWriteOk.
'   03/22/2007, lss
'       Added MoisTester to gUserPermission.
'   04/26/2007, lss
'       Added gHaveRawProspData.
'   05/08/2007, lss
'       Added gFormMode.
'   07/06/2007, lss
'       Modified Function gTransSapphireProd
'       -- gActiveMineNameLong = "Wingate"
'   01/03/2007, lss
'       Added Public gMassBalSelectMode As String.
'   02/25/2008, lss
'       Modified Function gGetMassBalanceMethod()
'       Added "mass balance mode" to Function gGetMiscMineGlobals.
'   03/17/2008, lss
'       Changed gDlDetail(10, 26) to gDlDetail(10, 27).  Need to
'       capture matrix tons.
'   04/24/2008, lss
'       Modified Sub gGetAllShiftsCbo2.
'   05/13/2008, lss
'       Added Function gRoundHalf.
'   11/04/2008, lss
'       Added "has draglines" to Function gGetMiscMineGlobals.
'   11/10/2008, lss
'       Added gActiveViewArea.
'   03/02/2009, lss
'       Added Case Is = "Raw prospect reduction admin"
'                 gUserPermission.RawProspectReduction = "Admin"
'       to gSetUserPermissions.
'   06/23/2009, lss
'       Added Public Function gGetChoicesActive.
'   07/25/2010, lss
'       Fixed Sub gPrintSpecialShiftReport -- added
'       aReportControl.GetNSubreports.
'   08/12/2010, lss
'       Added Function gGetEqptMsrmntValue.
'   10/08/2010, lss
'       Added Public gActiveMineNameRkBkRpts.
'
'**********************************************************************


    Public gActiveDate As Date
    Public gUserShiftSelectedDate As Date
    Public gLastUserSelectedDate As Date
    Public gActiveShift As String
    Public gActiveMode As String

    Public gActiveMineNameLong As String
    Public gActiveMineNameRkBkRpts As String

    Public gActiveMineNameShort As String
    Public gActiveInputArea As String
    Public gActiveViewArea As String
    Public gResize As Boolean
    Public gRptTimeFrame As String
    Public gRptDateRange As String
    Public gRptTimeFrame2 As String
    Public gActiveDragline As String
    '--
    Public gViewBegDate As String
    Public gViewEndDate As String
    Public gViewBegShift As String
    Public gViewEndShift As String
    Public gViewCrewNum As String
    Public gMassBalRpt As String
    '--
    Public gProdAnalNumDaysBack
    Public gProdAnalNumDaysForward
    '--
    'Login stuff
    Public gUserName As String
    Public gPassword As String
    Public gConnected As Boolean
    Public gDataSource As String
    Public gDBParams As OraParameters
    Public gOraSession As OraSession
    Public gOradatabase As OraDatabase
    Public gPath As String
    Public gReportPath As String
    Public gAcadMacroPath As String
    Public gOracleUserName As String
    Public gOracleUserPassword As String

    Public gLogonMine As Boolean
    Public gPipeTracking As Boolean
    Public gMineProspect As Boolean
    Public gCompanyName As String
    Public gProdDateAccessTime As String
    Public gProdDateCreateTime As String
    Public gSampDayShiftBeginTime As String
    Public gSampDayShiftEndTime As String
    Public gSzdFdTpohAsWhole As Boolean
    Public gWasherPctOperMode As String
    Public gSizingPctOperMode As String
    Public gFltPltPctOperMode As String
    Public gMoisBeginDate As Date
    Public gHasDredges As Boolean
    Public gHasDraglines As Boolean
    Public gReagentByDay As Boolean
    Public gNumWasherSides As Integer
    Public gFyrBeginDate As Date
    Public gPct100ProspDesc As String
    Public gDlMtxBktFillFctr As Boolean
    Public gDlHasMtxTons As Boolean
    Public gDlMtxYdsSource As String    '"Bucket count", "Matrix tons",
    'or "Not applicable"
    Public gProspGridType As String     '"Alpha-numeric", "Numeric"
    Public gWasherInput As Boolean
    Public gProspCnInsAdj As Integer
    Public gCalcSplitConc As Integer
    Public gFltPltFdTonsMsr As String   '"NA", "Tons", "TPOH", "Totalizer reads"
    Public gSzgFdTonsMsr As String      '"NA", "Tons", "TPOH", "Totalizer reads"
    Public gMineHasPb As Boolean
    Public gMineHasCn As Boolean
    Public gMineHasIp As Boolean
    Public gShowIpPlusPb As Boolean
    Public gMineHasSubProds As Boolean
    Public gProspPbIsPbIp As Boolean
    Public gProdTonsRound As Integer
    Public gHasCtlgReserves As Boolean
    Public gInvAdjAppliedShifts As Boolean
    Public gCatalogProspDesc As String
    Public gMassBalanceMode As String
    Public gUseNewProdAnalCalc As Boolean

    Public gShiftChangeDate As Date
    Public gHasShiftChangeDate As Boolean
    Public gFcoChangeDate As Date

'Special shift report login stuff
    Private mSrptPermissionsDynaset As OraDynaset
    Public gSrptUserName As String
    Public gSrptPassword As String
    Public gSrptUorOk As Boolean
    Public gSrptDlOk As Boolean
    Public gSrptPumpOk As Boolean
    Public gSrptWasherOk As Boolean
    Public gSrptFltPltOk As Boolean
    Public gSrptSizingOk As Boolean
    Public gSrptReagentOk As Boolean
    Public gSrptShippingOk As Boolean
    Public gSrptCancel As Boolean
    Public gSrptOperatorName As String
    Public gSrptPumpPackOk As Boolean

'Transfer between View1, View2, Input stuff
    Public gGoToView2 As Boolean
    Public gGoToView1 As Boolean
    Public gGoToInput As Boolean

'Transfer between Shifts (Day shift, Night shift, etc)
'and, Total day for View2 stuff
    Public gGoToTotalDay As Boolean
    Public gGoToShift1 As Boolean
    Public gGoToShift2 As Boolean
    Public gGoToShift3 As Boolean

'Big calendar graphing related items
    Public gSmallNumberTest(31) As Object
    Public gBigNumberTest(31) As Object
    Public g3rdNumberTest(31) As Object
    Public Structure BigCalendarGraph
        Public DataSource As String
        Public GraphTitle As String
        Public NumPoints As Integer
        Public LeftTitle As String
        Public BottomTitle As String
        Public GraphType As String
        Public Analysis As String
        Public Selected As String
        Public Product As String
    End Structure
Public gBigCalendarGraph As BigCalendarGraph

'User permission related items
    Public Structure UserPermission
        Public UserId As String       '1
        Public Field As String       '2  Input
        Public Washer As String       '3  Input
        Public Sizing As String       '4  Input
        Public FloatPlant As String       '5  Input
        Public Misc As String       '6  Input
        Public Analysis As String       '7  Input
        Public Shipping As String       '8  Input
        Public Production As String       '9  Input
        Public Reagent As String       '10 Input
        Public Prospect As String       '11
        Public Survey As String       '12
        Public MinePlan As String       '13
        Public Utilities As String       '14
        Public PumpYardages As String       '15
        Public StackerPosition As String       '16
        Public TrainShipping As String       '17
        Public DraglineCables As String       '18
        Public RawProspectChem As String       '19
        Public RawProspectMet As String       '20
        Public Views As Boolean      '21
        Public ProspectLoad As Boolean      '22
        Public SuperUser As Boolean      '23
        Public MultiMine As Boolean      '24
        Public Administrator As Boolean      '25
        Public CircuitAnalysisLoad As Boolean      '26
        Public BinAnalysisLoad As Boolean      '27
        Public TrainAnalysisLoad As Boolean      '28
        Public MaxInputScreens As Boolean      '29
        Public Graphs As Boolean      '30
        Public Maps As String       '31
        Public UtilityOperatorReport As String       '32
        Public DlInspectionReport As String       '33
        Public PumpInspectionReport As String       '34
        Public Piezometers As String       '35
        Public ReclamationActivity As String       '36
        Public WebReports As String       '37
        Public WaterSamples As String       '38
        Public PipeThickness As String       '39
        Public DecisionGrid As String       '40
        Public WasherShiftReport As String       '41
        Public FloatPlantShiftReport As String       '42
        Public SizingShiftReport As String       '43
        Public ReagentShiftReport As String       '44
        Public ShippingShiftReport As String       '45
        Public SrptSpvsrChkOff As Boolean      '46
        Public Absentees As String       '47
        Public SafetyMeetings As String       '48
        Public PumpPackShiftReport As String       '49
        Public InventoryAdjust As Boolean      '50
        Public RawProspectReduction As String       '51
        Public RockBookRecalc As Boolean      '52
        Public NoFrillsMois As Boolean      '53
        Public MoisTester As Boolean      '54
    End Structure
Public gUserPermission As UserPermission

'User permission related items
    Public Structure UserInfo
        Public UserName As String       '1
        Public MailName As String       '2
        Public MailServer As String       '3
        Public FirstName As String       '4
        Public MiddleInit As String       '5
        Public LastName As String       '6
        Public UserLoctn As String       '7
        Public DefaultMine As String       '8
End Structure
Public gUserInfo As UserInfo

    Public Structure RockBooks
        Public MineName As String   '1
        Public PitTakeupDate As Date     '2
        Public BeginDate As Date     '3
        Public EndDate As Date     '4
        Public MonthBeginDate As Date     '5
        Public ShortComment As String   '6
    '--
        Public ActAdjustedFeedTons As Long     '7
        Public ActAdjustedFeedBpl As Single   '8
        Public ActSizerRockTons As Single   '9
        Public ActTailBplGmt As Single   '10
        Public ActTailBplAdjustedFeed As Single   '11
        Public ActTailBplCircuits As Single   '12
        Public ActReportedFeedTons As Single   '13
        Public ActReportedFeedBpl As Single   '14
        Public ActTailBplReportedFeed As Single   '15
        Public ActWasherHours As Single   '16
        Public ActFloatPlantHours As Single   '17
        Public ActPbTons As Long     '18
        Public ActPbBpl As Single   '19
        Public ActPbFe2O3 As Single   '20
        Public ActPbAl2O3 As Single   '21
        Public ActPbMgO As Single   '22
        Public ActPbInsol As Single   '23
        Public ActCnTons As Long     '24
        Public ActCnBpl As Single   '25
        Public ActCnFe2O3 As Single   '26
        Public ActCnAl2O3 As Single   '27
        Public ActCnMgO As Single   '28
        Public ActCnInsol As Single   '29
    '--
        Public PrFeedTons As Long     '30 (100%PROSPECT)
        Public PrFeedBpl As Single   '31 (100%PROSPECT)
        Public PrClayTons As Single   '32 (100%PROSPECT)
        Public PrPbTons As Single   '33 (100%PROSPECT)
        Public PrPbBPl As Single   '34 (100%PROSPECT)
        Public PrPbFe2O3 As Single   '35 (100%PROSPECT)
        Public PrPbAl2O3 As Single   '36 (100%PROSPECT)
        Public PrPbMgO As Single   '37 (100%PROSPECT)
        Public PrPbInsol As Single   '38 (100%PROSPECT)
        Public PrCnTons As Single   '39 (100%PROSPECT)
        Public PrCnBpl As Single   '40 (100%PROSPECT)
        Public PrCnFe2O3 As Single   '41 (100%PROSPECT)
        Public PrCnAl2O3 As Single   '42 (100%PROSPECT)
        Public PrCnMgO As Single   '43 (100%PROSPECT)
        Public PrCnInsol As Single   '44 (100%PROSPECT)
        Public PrOvbAvgThickness As Single   '45 (100%PROSPECT)
        Public PrMtxAvgThickness As Single   '46 (100%PROSPECT)
        Public PrItbAvgThickness As Single   '47 (100%PROSPECT) new
    '--
        Public ActNumWasherSides As Single   '48
        Public ActWasherSidesOh As Single   '49
    '--
        Public PrcFeedTons As Long     '50 (CATALOG)
        Public PrcFeedBpl As Single   '51 (CATALOG)
        Public PrcClayTons As Single   '52 (CATALOG)
        Public PrcPbTons As Single   '53 (CATALOG)
        Public PrcPbBpl As Single   '54 (CATALOG)
        Public PrcPbFe2O3 As Single   '55 (CATALOG)
        Public PrcPbAl2O3 As Single   '56 (CATALOG)
        Public PrcPbMgO As Single   '57 (CATALOG)
        Public PrcPbInsol As Single   '58 (CATALOG)
        Public PrcCnTons As Single   '59 (CATALOG)
        Public PrcCnBpl As Single   '60 (CATALOG)
        Public PrcCnFe2O3 As Single   '61 (CATALOG)
        Public PrcCnAl2O3 As Single   '62 (CATALOG)
        Public PrcCnMgO As Single   '63 (CATALOG)
        Public PrcCnInsol As Single   '64 (CATALOG)
        Public PrcOvbAvgThickness As Single   '65 (CATALOG)
        Public PrcMtxAvgThickness As Single   '66 (CATALOG)
        Public PrcItbAvgThickness As Single   '67 (CATALOG) new
    '--
        Public ActIpTons As Long     '66
        Public ActIpBpl As Single   '67
        Public ActIpFe2O3 As Single   '68
        Public ActIpAl2O3 As Single   '69
        Public ActIpMgO As Single   '70
        Public ActIpInsol As Single   '71
    '--
        Public ProspPbIsPbIp As Boolean  '72
    '--
    '10/08/2012, lss new
        Public ActPbCaO As Single   '73
        Public ActCnCaO As Single   '74
        Public PrPbCaO As Single   '75
        Public PrCnCaO As Single   '76
        Public PrcPbCaO As Single   '77
        Public PrcCnCaO As Single   '78
        Public ActIpCaO As Single   '79
End Structure
Public gRockBook As RockBooks

Public gDlDetail(10, 28) As String

'Prospect data related items
Public gHole1SplitsDynaset As OraDynaset
Public gHole1CompositeDynaset As OraDynaset

Public gHole2SplitsDynaset As OraDynaset
Public gHole2CompositeDynaset As OraDynaset

'Print selection related items
'gPrtSelection      0 = Cancel print
'                   1 = Regular print Screen
'                   2 = Bob's special print Screen
'                   3 = Print spreadsheet #1
'                   4 = Print spreadsheet #2
'                   3 = Print spreadsheet #3
'                   4 = Print spreadsheet #4
Public gPrtSelection As Integer
Public gPrtPrintScreenOK As Boolean
Public gPrtSpecPrintScreenOK As Boolean
Public gPrtSSname1 As String
Public gPrtSSname2 As String
Public gPrtSSname3 As String
Public gPrtSSname4 As String

'Stuff for selecting 2 dates for printing reports
Public gPrtBeginDate As Date
Public gCalBeginDate As Date
Public gPrtEndDate As Date
Public gCalEndDate As Date
Public gPrtOK As Boolean
Public gPrtBeginShift As String     'Added 08/24/2005, lss
Public gPrtEndShift As String       'Added 08/24/2005, lss

'Stuff for hard and soft freeze dates
Public gHardFreezeDate As Date
Public gSoftFreezeDate As Date

'Stuff for dragline prospect
    Public gSplitPebb(20, 8) As Object
    Public gSplitConc(20, 8) As Object
    Public gSplitTotProd(20, 8) As Object

'Resizing related items
    Public Structure gControlProperties
        Public WidthProportions As Single
        Public HeightProportions As Single
        Public TopProportions As Single
        Public LeftProportions As Single
        Public FontProportions As Single
    End Structure

'Grid to text file transfers related items
    Public gGridObject As AxvaSpread                  '1
Public gPrintGridHeader As String               '2
Public gPrintGridSubHeader1 As String           '3
Public gPrintGridSubHeader2 As String           '4
Public gPrintGridSubHeader3 As String           '5
    Public gPrintGridDefaultTxtFname As String = String.Empty      '6
Public gPrintGridFooter As String               '7
Public gPrintGridFooter2 As String              '8
Public gOrientHeader As String                  '9
Public gOrientSubHeader1 As String              '10
Public gOrientSubHeader2 As String              '11
Public gOrientSubHeader3 As String              '12
Public gOrientFooter As String                  '13
Public gOrientFooter2 As String                 '14
Public gSubHead2IsHeader As Boolean             '15
Public gPrintMarginLeft As Long                 '16
Public gPrintMarginRight As Long                '17
Public gPrintMarginTop As Long                  '18
Public gPrintMarginBottom As Long               '19
Public gPrintOrientation As Integer             '20
Public gProspStandard As String                 '21
Public gAutoPrint As Boolean                    '22

'Grid to graph related items
    Public gGraphGridObject As AxvaSpread     '1
Public gGraphCol As Integer             '2
Public gGraphRowSkip As Integer         '3
Public gGraphOmitZeros As Boolean       '4
Public gGraphMineName As String         '5
Public gGraphLoctnName() As String      '6
Public gGraphFullMsrName As String      '7
Public gGraphXaxisLabel As String       '8
Public gLegendFontSize As String        '9
Public gAxisTitleFontSize As String     '10
Public gLabelFontSize As String         '11
Public gGraphMean As Boolean            '12
Public gGraphLblEvery As Integer        '13
Public gGraphLblData As String          '14

'Overlay graph related items
Public gOlayNeeded As Boolean
Public gOlayData() As Single
Public gOlayGraphStyle As Integer
Public gOlayThickLines As Integer
Public gOlayPatternedLines As Integer
Public gOlayPattern As Integer
Public gOlayYaxisUse As Integer
Public gOlayYaxisPos As Integer
Public gOlayYaxisStyle As Integer
Public gOlayYaxisMax As Single
Public gOlayYaxisTicks As Integer
Public gOlaySymbol As Integer
Public gOlayLegendText As String
Public gOlayRightTitle As String
Public gOlayRightTitleStyle As Integer
Public gOlayAdjust As Boolean
Public gOlayColor As Integer

'Graph module related items
    Public gGraphData() As Object
    Public gGphNumPoints As Integer
    Public gGphNumSets As Integer
    Public gGphDecimal As Single
    Public gGphSetInterval As Double
    Public gGphLblFormat As String
    Public gGphRoundValue As Integer
    Public gGphMinimum As Double
    Public gGphAutoPrint As Boolean

'Grid to dxf file related items
Public gPrintGridMineName As String
Public gByVoloView As Boolean
Public gByInViso As Boolean
Public gByAutoCad As Boolean

'Stuff for "One number form"
Public gOneNumFrmCaption As String
Public gOneNumLine1 As String
Public gOneNumLine2 As String
Public gOneNumLine3 As String
Public gOneNumValue As Double
Public gOneNumDefault As Double
Public gOneNumStatus As Boolean

'Stuff for "One string form"
Public gOneStrFrmCaption As String
Public gOneStrLine1 As String
Public gOneStrLine2 As String
Public gOneStrLine3 As String
Public gOneStrValue As String
Public gOneStrDefault As String
Public gOneStrMaxLen As Integer
Public gOneStrStatus As Boolean

'AutoCAD file view form related items
Public gAcadFileName As String

'Surfer file view form related items
Public gSurferDataFileName As String
Public gSurferMineName As String
Public gSurferParameter As String
Public gSurferBySection As Boolean

'Miscellaneous default files
Public gDxfDefaultFile As String
Public gDatDefaultFile As String
Public gTxtDefaultFile As String

'Dragline prospect stuff
Public gDragMine As String
Public gDragSec As Integer
Public gDragTwp As Integer
Public gDragRge As Integer
Public gDragHole As String
Public gDragProspStandard As String

'Special shift report stuff
    Public mMineNameDynaset As OraDynaset
    Public mRptNameDynaset As OraDynaset
    Public mSftyComm1Lbl As String
    Public mSftyComm2Lbl As String
    Public mSafeAreaComm1Lbl As String
    Public mSafeAreaComm2Lbl As String
   
'Shift names
    Public Structure gShiftNamesType
        Public ShiftName As String
        Public BeginTime As String
        Public EndTime As String
        Public ShiftLength As Integer
        Public ShiftOrder As Integer
        Public SampBeginTime As String
        Public SampEndTime As String
    '-----
        Public BeginHour As Integer
        Public BeginMinute As Integer
        Public EndHour As Integer
        Public EndMinute As Integer
    End Structure
Public gShiftNames() As gShiftNamesType
Public gNumShifts As Integer
Public gFirstShift As String
Public gLastShift As String
Public gShiftLength As Single

    Public Structure ShiftAccessType
        Public ShiftName As String
        Public BeginHour As Integer
        Public BeginMinute As Integer
        Public EndHour As Integer
        Public EndMinute As Integer
    End Structure
    Private mShiftAccess() As ShiftAccessType

    Public Structure gShiftInfoType
        Public dDate As Date
        Public Shift As String
    End Structure

    Public Structure gProdInfoType
        Public Tons As Double
        Public Bpl As Single
        Public Insol As Single
        Public Fe2O3 As Single
        Public Al2O3 As Single
        Public MgO As Single
        Public CaO As Single
        Public Cd As Single
        Public P2O5 As Single
        Public CaOP2O5 As Single
        Public Mer As Single
        Public MgoP2O5 As Single
    End Structure

    Public Structure gProdMonthAdjType
        Public ActPbTons As Single
        Public ActCnTons As Single
        Public ActIpTons As Single
    '-----
        Public ActPbBpl As Single
        Public ActCnBpl As Single
        Public ActIpBpl As Single
    '-----
        Public AdjPbTons As Single
        Public AdjCnTons As Single
        Public AdjIpTons As Single
    '-----
        Public PbInv As Double
        Public CnInv As Double
        Public IpInv As Double
    '-----
        Public FlyoverDate As Date
        Public AppliedToShifts As Boolean
    End Structure

'General explain related
Public gGenExplainCaption As String
Public gGenExplainComment As String
Public gGenDelayComment As String

    Public Structure gSzdFdTonsType
        Public FneFdTons As Double
        Public CrsFdTons As Double
        Public FneFdHrs As Single
        Public CrsFdHrs As Single
        Public FneFdTph As Single
        Public CrsFdTph As Single
        Public NumFneCircs As Integer
        Public NumCrsCircs As Integer
    End Structure

    Public Const gMassBalanceCutOffDateSf = #11/1/1997#
    Public Const gMassBalanceCutOffDateHp = #6/1/1994#
    Public Const gMassBalanceCutOffDateWg = #9/22/2004#
    Public Const gMassBalanceCutOffDateFc = #1/1/2004#
    Public Const gMassBalanceCutOffDateHw = #1/1/2004#

'Global variables for Four Corners Multifos concentrate
Public gMultifosMinBpl As Single
Public gMultifosTargIns As Single
Public gMultifosMaxAl As Single

    Public Structure gMassBalDataAbbrvType
        Public AdjFdTons As Double
        Public FdBpl As Single
        Public GmtBpl As Single
        Public PctRcvry As Single
    End Structure

    Public Enum fFloatPlantGmtRowEnum
        GrAsReportedGmtBpl = 1
        GrCalculatedGmtBpl = 2
        GrReportedFdTons = 3
        GrGmtBplFromCircuits = 4
    End Enum

public Enum fFloatPlantGmtColEnum
    GcFdTons = 1
    GcCnTons = 2
    GcFdBpl = 3
    GcCnBpl = 4
    GcTlBpl = 5
    GcRc = 6
    GcPctRcvry = 7
End Enum

    Public gFormLoad As Boolean
    Public gWriteOk As Boolean
    Public gFormMode As String
    Public gRawProspDynaset As OraDynaset
    Public gHaveRawProspData As Boolean
    Public gMassBalSelectMode As String

    Public gFileNumber As Integer
    Public gOutputLines As List(Of String)

    Public Function GetDaysInMonth(ByVal aMonth As Integer, _
                               ByVal aYear As Integer) As Integer

'**********************************************************************
'
'
'
'**********************************************************************

    'Determine the number of days in the month
    
    Dim NextMonth As Integer
    Dim NextYear As Integer
    Dim FirstOfNextMonth As Date
    Dim LastOfThisMonth As Date
    
    NextMonth = aMonth + 1
    NextYear = aYear
    
    If NextMonth = 13 Then
        NextMonth = 1
        NextYear = NextYear + 1
    End If
    
    FirstOfNextMonth = CDate(Str(NextMonth) + "/01/" + Str(NextYear))
        LastOfThisMonth = CDate(FirstOfNextMonth).AddDays(-1)
    
    GetDaysInMonth = DatePart("d", LastOfThisMonth)
End Function

Public Function gGetDaysInMonth2(ByVal aDate As Date) As Integer

'**********************************************************************
'
'
'
'**********************************************************************

    'Determine the number of days in the month
    
    Dim NextMonth As Integer
    Dim NextYear As Integer
    Dim FirstOfNextMonth As Date
    Dim LastOfThisMonth As Date
    Dim ThisMonth As Integer
    Dim ThisYear As Integer
    
    ThisMonth = DatePart("m", aDate)
    ThisYear = DatePart("yyyy", aDate)
    
    NextMonth = ThisMonth + 1
    NextYear = ThisYear
    
    If NextMonth = 13 Then
        NextMonth = 1
        NextYear = NextYear + 1
    End If
    
    FirstOfNextMonth = Str(NextMonth) + "/01/" + Str(NextYear)
        LastOfThisMonth = FirstOfNextMonth.AddDays(-1)
    
    gGetDaysInMonth2 = DatePart("d", LastOfThisMonth)
End Function

Public Function gGetDOW(ByVal aDate As Date) As String

'**********************************************************************
'
'
'
'**********************************************************************

    'Determine the name of the day for a given date
    
    Dim ThisWeekDay As Integer
    
    ThisWeekDay = Weekday(aDate, vbSunday)
        gGetDOW = String.Empty
    Select Case ThisWeekDay
        Case Is = 1
            gGetDOW = "Sunday"
        Case Is = 2
            gGetDOW = "Monday"
        Case Is = 3
            gGetDOW = "Tuesday"
        Case Is = 4
            gGetDOW = "Wednesday"
        Case Is = 5
            gGetDOW = "Thursday"
        Case Is = 6
            gGetDOW = "Friday"
        Case Is = 7
            gGetDOW = "Saturday"
    End Select
    
End Function

Sub ClearParams(aParams As OraParameters)

'**********************************************************************
'
'
'
'**********************************************************************

    Dim a As Integer
    
    If Not aParams Is Nothing Then
        For a = 0 To aParams.Count - 1
                aParams.Remove(0)
        Next a
    End If
End Sub

Function gSetUserPermissions(ByVal aUserId As String, _
                             ByVal aMineName As String) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    Dim UserPermissionDynaset As OraDynaset
    Dim MaxRows As Integer
    Dim CurrentRow As Integer
    Dim SettingUp As Boolean
    Dim RecordCount As Integer
    
    SettingUp = False
    gSetUserPermissions = True
    
    'aMineName -- South Fort Meade, Fort Meade, Hookers Prairie, etc.
    
    If Not SettingUp Then
        'Load user permissions from Oracle table"

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
 
        'Get all permissions for user
            'Set 
            params = gDBParams
    
            params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2
     
            params.Add("pUserID", StrConv(aUserId, vbUpperCase), ORAPARM_INPUT)
        params("pUserID").serverType = ORATYPE_NUMBER
    
            params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR
    
        'PROCEDURE get_user_permissions
        'pMineName        IN     VARCHAR2,
        'pUserID          IN     VARCHAR2,
        'pResult          IN OUT c_users)
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_user_permissions(:pMineName," + _
                      ":pUserID, :pResult);end;", ORASQL_FAILEXEC)
            UserPermissionDynaset = params("pResult").Value
                        
        RecordCount = UserPermissionDynaset.RecordCount
        
        gUserPermission.UserId = StrConv(aUserId, vbUpperCase)
        gUserPermission.Field = "None"
        gUserPermission.Washer = "None"
        gUserPermission.Sizing = "None"
        gUserPermission.FloatPlant = "None"
        gUserPermission.Misc = "None"
        gUserPermission.Analysis = "None"
        gUserPermission.Shipping = "None"
        gUserPermission.Production = "None"
        gUserPermission.Reagent = "None"
        gUserPermission.Prospect = "None"
        gUserPermission.Survey = "None"
        gUserPermission.MinePlan = "None"
        gUserPermission.Utilities = "None"
        gUserPermission.PumpYardages = "None"
        gUserPermission.StackerPosition = "None"
        gUserPermission.TrainShipping = "None"
        gUserPermission.DraglineCables = "None"
        gUserPermission.RawProspectChem = "None"
        gUserPermission.RawProspectMet = "None"
        gUserPermission.Views = False
        gUserPermission.ProspectLoad = False
        gUserPermission.MultiMine = False
        gUserPermission.SuperUser = False
        gUserPermission.Administrator = False
        gUserPermission.CircuitAnalysisLoad = False
        gUserPermission.BinAnalysisLoad = False
        gUserPermission.TrainAnalysisLoad = False
        gUserPermission.MaxInputScreens = False
        gUserPermission.Graphs = False
        gUserPermission.Maps = "None"
        gUserPermission.UtilityOperatorReport = "None"
        gUserPermission.DlInspectionReport = "None"
        gUserPermission.PumpInspectionReport = "None"
        gUserPermission.Piezometers = "None"
        gUserPermission.ReclamationActivity = "None"
        gUserPermission.WebReports = "None"
        gUserPermission.WaterSamples = "None"
        gUserPermission.PipeThickness = "None"
        gUserPermission.DecisionGrid = "None"
        gUserPermission.WasherShiftReport = "None"
        gUserPermission.FloatPlantShiftReport = "None"
        gUserPermission.SizingShiftReport = "None"
        gUserPermission.ReagentShiftReport = "None"
        gUserPermission.ShippingShiftReport = "None"
        gUserPermission.SrptSpvsrChkOff = False
        gUserPermission.Absentees = "None"
        gUserPermission.SafetyMeetings = "None"
        gUserPermission.PumpPackShiftReport = "None"
        gUserPermission.InventoryAdjust = False
        gUserPermission.RawProspectReduction = "None"
        gUserPermission.RockBookRecalc = False
        gUserPermission.NoFrillsMois = False
        gUserPermission.MoisTester = False
        
        MaxRows = UserPermissionDynaset.RecordCount + 1
        
        UserPermissionDynaset.MoveFirst
 
        For CurrentRow = 2 To MaxRows
            Select Case UserPermissionDynaset.Fields("permission_type_name").Value
                '1
                Case Is = "Field input read"
                    gUserPermission.Field = "Read"
                 
                Case Is = "Field input write"
                    gUserPermission.Field = "Write"

                '2
                Case Is = "Washer input read"
                    gUserPermission.Washer = "Read"
                 
                Case Is = "Washer input write"
                    gUserPermission.Washer = "Write"
                
                '3
                Case Is = "Sizing input read"
                    gUserPermission.Sizing = "Read"
                
                Case Is = "Sizing input write"
                    gUserPermission.Sizing = "Write"
                
                '4
                Case Is = "Float plant input read"
                    gUserPermission.FloatPlant = "Read"
                 
                Case Is = "Float plant input write"
                    gUserPermission.FloatPlant = "Write"
                
                '5
                Case Is = "Misc input read"
                    gUserPermission.Misc = "Read"
                  
                Case Is = "Misc input write"
                    gUserPermission.Misc = "Write"
                 
                '6
                Case Is = "Analysis input read"
                    gUserPermission.Analysis = "Read"
                  
                Case Is = "Analysis input write"
                    gUserPermission.Analysis = "Write"
                
                '7
                Case Is = "Shipping input read"
                    gUserPermission.Shipping = "Read"
                
                Case Is = "Shipping input write"
                    gUserPermission.Shipping = "Write"
                 
                '8
                Case Is = "Production input read"
                    gUserPermission.Production = "Read"
                
                Case Is = "Production input write"
                    gUserPermission.Production = "Write"
                
                '9
                Case Is = "Reagent input read"
                    gUserPermission.Reagent = "Read"
                 
                Case Is = "Reagent input write"
                    gUserPermission.Reagent = "Write"
                
                '10
                Case Is = "Prospect read"
                    gUserPermission.Prospect = "Read"
                 
                Case Is = "Prospect write"
                    gUserPermission.Prospect = "Write"
               
                '11
                Case Is = "Survey read"
                    gUserPermission.Survey = "Read"
                     
                Case Is = "Survey write"
                    gUserPermission.Survey = "Write"
                    
                '12
                Case Is = "Mine plan read"
                    gUserPermission.MinePlan = "Read"
                    
                Case Is = "Mine plan write"
                    gUserPermission.MinePlan = "Write"
                    
                '13
                Case Is = "Utilities read"
                    gUserPermission.Utilities = "Read"
                     
                Case Is = "Utilities write"
                    gUserPermission.Utilities = "Write"
                     
                '14
                Case Is = "Pump yardages read"
                    gUserPermission.PumpYardages = "Read"
                     
                Case Is = "Pump yardages write"
                    gUserPermission.PumpYardages = "Write"
                    
                Case Is = "Pump yardages setup"
                    gUserPermission.PumpYardages = "Setup"
                    
                '15
                Case Is = "Stacker position read"
                    gUserPermission.StackerPosition = "Read"
                     
                Case Is = "Stacker position write"
                    gUserPermission.StackerPosition = "Write"
                    
                '16
                Case Is = "Train shipping read"
                    gUserPermission.TrainShipping = "Read"
                    
                Case Is = "Train shipping write"
                    gUserPermission.TrainShipping = "Write"
                    
                '17
                Case Is = "Dragline cables read"
                    gUserPermission.DraglineCables = "Read"
                    
                Case Is = "Dragline cables write"
                    gUserPermission.DraglineCables = "Write"
                     
                '18
                Case Is = "Prospect load"
                    gUserPermission.ProspectLoad = True
                     
                '19
                Case Is = "Superuser"
                    gUserPermission.SuperUser = True
    
                '20
                Case Is = "Views and reports"
                    gUserPermission.Views = True
                    
                '21
                Case Is = "Multi-mine"
                    gUserPermission.MultiMine = True
                    
                '22
                Case Is = "Raw prospect chem lab read"
                    gUserPermission.RawProspectChem = "Read"
                     
                Case Is = "Raw prospect chem lab write"
                    gUserPermission.RawProspectChem = "Write"
                     
                '23
                Case Is = "Raw prospect met lab read"
                    gUserPermission.RawProspectMet = "Read"
                    
                Case Is = "Raw prospect met lab write"
                    gUserPermission.RawProspectMet = "Write"
                                                  
                '24
                Case Is = "Administrator"
                    gUserPermission.Administrator = True
                 
                '25
                Case Is = "Circuit analysis load"
                    gUserPermission.CircuitAnalysisLoad = True
      
                '26
                Case Is = "Bin analysis load"
                    gUserPermission.BinAnalysisLoad = True
                    
                '27
                Case Is = "Train analysis load"
                    gUserPermission.TrainAnalysisLoad = True
              
                '28
                Case Is = "Max input screens"
                    gUserPermission.MaxInputScreens = True
                    
                '29
                Case Is = "Graphs"
                    gUserPermission.Graphs = True
       
                '30
                Case Is = "Maps read"
                    gUserPermission.Maps = "Read"
                    
                Case Is = "Maps write"
                    gUserPermission.Maps = "Write"
                    
                '31
                Case Is = "Utility Operator Report read"
                    gUserPermission.UtilityOperatorReport = "Read"
                    
                Case Is = "Utility Operator Report write"
                    gUserPermission.UtilityOperatorReport = "Write"
                    
                Case Is = "Utility Operator Report setup"
                    gUserPermission.UtilityOperatorReport = "Setup"
                    
                '32
                Case Is = "DL Inspection Report read"
                    gUserPermission.DlInspectionReport = "Read"
                    
                Case Is = "DL Inspection Report write"
                    gUserPermission.DlInspectionReport = "Write"
                    
                Case Is = "DL Inspection Report setup"
                    gUserPermission.DlInspectionReport = "Setup"
                    
                '33
                Case Is = "Pump Inspection Report read"
                    gUserPermission.PumpInspectionReport = "Read"
                    
                Case Is = "Pump Inspection Report write"
                    gUserPermission.PumpInspectionReport = "Write"
                    
                Case Is = "Pump Inspection Report setup"
                    gUserPermission.PumpInspectionReport = "Setup"
                    
                '34
                Case Is = "Piezometers read"
                    gUserPermission.Piezometers = "Read"
                    
                Case Is = "Piezometers write"
                    gUserPermission.Piezometers = "Write"
                    
                '35
                Case Is = "Reclamation activity read"
                    gUserPermission.ReclamationActivity = "Read"
                    
                Case Is = "Reclamation activity write"
                    gUserPermission.ReclamationActivity = "Write"
                    
                '36
                Case Is = "Web reports read"
                    gUserPermission.WebReports = "Read"
                    
                Case Is = "Web reports write"
                    gUserPermission.WebReports = "Write"
                    
                '37
                Case Is = "Water samples read"
                    gUserPermission.WaterSamples = "Read"
                    
                Case Is = "Water samples write"
                    gUserPermission.WaterSamples = "Write"
                    
                '38
                Case Is = "Pipe thickness read"
                    gUserPermission.PipeThickness = "Read"
                    
                Case Is = "Pipe thickness write"
                    gUserPermission.PipeThickness = "Write"
                    
                Case Is = "Pipe thickness setup"
                    gUserPermission.PipeThickness = "Setup"
                    
                '39
                Case Is = "Decision grid read"
                    gUserPermission.DecisionGrid = "Read"
                    
                Case Is = "Decision grid write"
                    gUserPermission.DecisionGrid = "Write"
                    
                '40
                Case Is = "Washer Shift Report read"
                    gUserPermission.WasherShiftReport = "Read"
                    
                Case Is = "Washer Shift Report write"
                    gUserPermission.WasherShiftReport = "Write"
                    
                Case Is = "Washer Shift Report setup"
                    gUserPermission.WasherShiftReport = "Setup"
                    
                '41
                Case Is = "Float Plant Shift Report read"
                    gUserPermission.FloatPlantShiftReport = "Read"
                    
                Case Is = "Float Plant Shift Report write"
                    gUserPermission.FloatPlantShiftReport = "Write"
                    
                Case Is = "Float Plant Shift Report setup"
                    gUserPermission.FloatPlantShiftReport = "Setup"
                                        
                '42
                Case Is = "Sizing Shift Report read"
                    gUserPermission.SizingShiftReport = "Read"
                    
                Case Is = "Sizing Shift Report write"
                    gUserPermission.SizingShiftReport = "Write"
                    
                Case Is = "Sizing Shift Report setup"
                    gUserPermission.SizingShiftReport = "Setup"
                    
                '43
                Case Is = "Reagent Shift Report read"
                    gUserPermission.ReagentShiftReport = "Read"
                    
                Case Is = "Reagent Shift Report write"
                    gUserPermission.ReagentShiftReport = "Write"
                    
                Case Is = "Reagent Shift Report setup"
                    gUserPermission.ReagentShiftReport = "Setup"
                                        
                '44
                Case Is = "Shipping Shift Report read"
                    gUserPermission.ShippingShiftReport = "Read"
                    
                Case Is = "Shipping Shift Report write"
                    gUserPermission.ShippingShiftReport = "Write"
                    
                Case Is = "Shipping Shift Report setup"
                    gUserPermission.ShippingShiftReport = "Setup"
                    
                '45
                Case Is = "Srpt spvsr chk-off"
                    gUserPermission.SrptSpvsrChkOff = True
      
                '46
                Case Is = "Absentees input"
                    gUserPermission.Absentees = "Input"
                    
                Case Is = "Absentees review"
                    gUserPermission.Absentees = "Review"
                    
                Case Is = "Absentees setup"
                    gUserPermission.Absentees = "Setup"
                    
                Case Is = "Absentees read"
                    gUserPermission.Absentees = "Read"
                    
                '47
                Case Is = "Safety meetings read"
                    gUserPermission.SafetyMeetings = "Read"
                    
                Case Is = "Safety meetings write"
                    gUserPermission.SafetyMeetings = "Write"
                    
                Case Is = "Safety meetings setup"
                    gUserPermission.SafetyMeetings = "Setup"
                    
                '48
                Case Is = "Pump Pack Shift Report read"
                    gUserPermission.PumpPackShiftReport = "Read"
                    
                Case Is = "Pump Pack Shift Report write"
                    gUserPermission.PumpPackShiftReport = "Write"
                    
                Case Is = "Pump Pack Shift Report setup"
                    gUserPermission.PumpPackShiftReport = "Setup"
                    
                '49
                Case Is = "Inventory adjust"
                    gUserPermission.InventoryAdjust = True
                    
                '50
                Case Is = "Raw prospect reduction read"
                    gUserPermission.RawProspectReduction = "Read"
                    
                Case Is = "Raw prospect reduction write"
                    gUserPermission.RawProspectReduction = "Write"
                    
                Case Is = "Raw prospect reduction setup"
                    gUserPermission.RawProspectReduction = "Setup"
                    
                Case Is = "Raw prospect reduction admin"
                    gUserPermission.RawProspectReduction = "Admin"
                    
                '51
                Case Is = "Rock book recalc"
                    gUserPermission.RockBookRecalc = True
                    
                '52
                Case Is = "No frills MOIS"
                    gUserPermission.NoFrillsMois = True
                    
                '53
                Case Is = "MOIS tester"
                    gUserPermission.MoisTester = True
            End Select
            
            UserPermissionDynaset.MoveNext
        Next CurrentRow
                 
            ClearParams(params)
    Else
        gUserPermission.UserId = StrConv(aUserId, vbUpperCase)
        gUserPermission.Field = "Write"
        gUserPermission.Washer = "Write"
        gUserPermission.Sizing = "Write"
        gUserPermission.FloatPlant = "Write"
        gUserPermission.Misc = "Write"
        gUserPermission.Analysis = "Write"
        gUserPermission.Shipping = "Write"
        gUserPermission.Production = "Write"
        gUserPermission.Reagent = "Write"
        gUserPermission.Prospect = "Write"
        gUserPermission.Survey = "Write"
        gUserPermission.MinePlan = "Write"
        gUserPermission.Utilities = "Write"
        gUserPermission.PumpYardages = "Write"
        gUserPermission.StackerPosition = "Write"
        gUserPermission.TrainShipping = "Write"
        gUserPermission.DraglineCables = "Write"
        gUserPermission.RawProspectChem = "Write"
        gUserPermission.RawProspectMet = "Write"
        gUserPermission.Views = True
        gUserPermission.ProspectLoad = True
        gUserPermission.MultiMine = True
        gUserPermission.SuperUser = True
        gUserPermission.Administrator = True
        gUserPermission.CircuitAnalysisLoad = True
        gUserPermission.BinAnalysisLoad = True
        gUserPermission.TrainAnalysisLoad = True
        gUserPermission.MaxInputScreens = True
        gUserPermission.Graphs = True
        gUserPermission.Maps = "Write"
        gUserPermission.UtilityOperatorReport = "Setup"
        gUserPermission.DlInspectionReport = "Setup"
        gUserPermission.PumpInspectionReport = "Setup"
        gUserPermission.Piezometers = "Write"
        gUserPermission.ReclamationActivity = "Write"
        gUserPermission.WebReports = "Write"
        gUserPermission.WaterSamples = "Write"
        gUserPermission.PipeThickness = "Write"
        gUserPermission.DecisionGrid = "Write"
        gUserPermission.WasherShiftReport = "Setup"
        gUserPermission.FloatPlantShiftReport = "Setup"
        gUserPermission.SizingShiftReport = "Setup"
        gUserPermission.ReagentShiftReport = "Setup"
        gUserPermission.ShippingShiftReport = "Setup"
        gUserPermission.SrptSpvsrChkOff = True
        gUserPermission.Absentees = "Setup"
        gUserPermission.SafetyMeetings = "Setup"
        gUserPermission.PumpPackShiftReport = "Setup"
        gUserPermission.InventoryAdjust = True
        gUserPermission.RawProspectReduction = "Setup"
        gUserPermission.RockBookRecalc = True
        gUserPermission.NoFrillsMois = True
        gUserPermission.MoisTester = True
    End If
       
    If RecordCount = 0 Then
        gSetUserPermissions = False
    End If
End Function

Sub gGiveMessage(ByVal aMessage As String, ByVal aTitle As String)

'**********************************************************************
'
'
'
'**********************************************************************

        MsgBox(aMessage, vbExclamation, aTitle)
End Sub

Sub gSetActionStatus(ByVal aStatus As String)

'**********************************************************************
'
'
'
'**********************************************************************

    On Error Resume Next
        'frmMain.sbrMain.Panels(1).Text = aStatus
End Sub

Sub gSetActionMode(ByVal aMode As String)

'**********************************************************************
'
'
'
'**********************************************************************

    On Error Resume Next
        'frmMain.sbrMain.Panels(2).Text = aMode
End Sub

Sub gShowMainExit(ByVal aStatus As String)

'**********************************************************************
'
'
'
'**********************************************************************

    On Error Resume Next
    
        'If aStatus = "On" Then
        '    frmMain.cmdExit.Visible = True
        'Else
        '    frmMain.cmdExit.Visible = False
        'End If
End Sub

Public Function gGetFirstOfMonth(ByVal aDate As Date) As Date

'**********************************************************************
'
'
'
'**********************************************************************
    Dim MonthStr As String
    Dim YearStr As String
    Dim DateStr As String
     
    MonthStr = DatePart("m", aDate)
    YearStr = DatePart("yyyy", aDate)
    DateStr = MonthStr & "/01/" & YearStr
        
    gGetFirstOfMonth = CDate(DateStr)
End Function

Public Function gGetLastOfMonth(ByVal aDate As Date) As Date

'**********************************************************************
'
'
'
'**********************************************************************

    Dim MonthStr As String
    Dim YearStr As String
    Dim DateStr As String
    Dim NumDaysInMonth As Integer
    Dim FirstDayInMonth As Date
    
    MonthStr = DatePart("m", aDate)
    YearStr = DatePart("yyyy", aDate)
    DateStr = MonthStr & "/01/" & YearStr
        
    FirstDayInMonth = CDate(DateStr)
    
    NumDaysInMonth = GetDaysInMonth(Val(MonthStr), Val(YearStr))
        
        gGetLastOfMonth = FirstDayInMonth.AddDays((NumDaysInMonth - 1))
End Function

Public Function gGetFirstOfNextMonth(ByVal aDate As Date) As Date

'**********************************************************************
'
'
'
'**********************************************************************
    Dim MonthStr As String
    Dim YearStr As String
    Dim DateStr As String
     
    MonthStr = DatePart("m", aDate)
    YearStr = DatePart("yyyy", aDate)
    
    MonthStr = CStr(Val(DatePart("m", aDate)) + 1)
    If Val(MonthStr) = 13 Then
        MonthStr = "1"
        YearStr = CStr(Val(YearStr) + 1)
    End If
    
    DateStr = MonthStr & "/01/" & YearStr
        
    gGetFirstOfNextMonth = CDate(DateStr)
End Function

Function gGetCurrentFreezeDates(aMineName As String) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo gGetCurrentFreezDatesError
    
    Dim MaxRows As Integer
    Dim CurrentRow As Integer
    Dim SettingUp As Boolean
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    
    SettingUp = False
    
    'aMineName -- South Fort Meade, Fort Meade, Hookers Prairie, etc.
    
    'Load freeze dates from Oracle table"
     
        params = gDBParams
    
        params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pHardFreezeDate", DBNull.Value, ORAPARM_OUTPUT)
        params("pHardFreezeDate").serverType = ORATYPE_DATE

        params.Add("pSoftFreezeDate", DBNull.Value, ORAPARM_OUTPUT)
        params("pSoftFreezeDate").serverType = ORATYPE_DATE
    
    'pMineName
    'pResult
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_freeze_dates(:pMineName," + _
                  ":pHardFreezeDate, :pSoftFreezeDate);end;", ORASQL_FAILEXEC)
                       
        If Not IsDBNull(params("pHardFreezeDate").Value) Then
            gHardFreezeDate = params("pHardFreezeDate").Value
        Else
            gHardFreezeDate = Today
        End If
    
        If Not IsDBNull(params("pSoftFreezeDate").Value) Then
            gSoftFreezeDate = params("pSoftFreezeDate").Value
        Else
            gSoftFreezeDate = Today
        End If
 
        ClearParams(params)
    gGetCurrentFreezeDates = True
    
    Exit Function
    
gGetCurrentFreezDatesError:
        MsgBox("Error getting freeze dates." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Freeze Dates Error")
        
    On Error Resume Next
        ClearParams(params)
    gGetCurrentFreezeDates = False
End Function

Public Function gGetFrozenStatus(ByVal aDate As Date) As String

'**********************************************************************
'
'
'
'**********************************************************************

    'Returns "Hard", "Soft", "None"
    
    'gHardFreezeDate As Date
    'gSoftFreezeDate As Date
    
    'Frozen status = "Hard" -- noone may make changes, NOT EVEN
    '                          AN ADMINISTRATOR!
    
    'Frozen status = "Soft" -- Only an Administrator may make changes!
    
    Select Case aDate
        Case Is <= gHardFreezeDate
            gGetFrozenStatus = "Hard"
            
        Case Is <= gSoftFreezeDate
            gGetFrozenStatus = "Soft"
        
        Case Else
            gGetFrozenStatus = "None"
    End Select
End Function

Public Sub gCalculateTimeFrame()

'**********************************************************************
'
'
'
'**********************************************************************

    'This procedure originally had items like gActiveDate + 0.25,
    'and gActiveDate + 0.75  They were removed on 10/26/2004, lss
    
    '1) Shifts           OK for 2 -> 3 shifts  07/20/2006, lss
    '2) Total day        OK for 2 -> 3 shifts  07/20/2006, lss
    '3) Week-to-date     OK for 2 -> 3 shifts  07/20/2006, lss
    '4) Month-to-date    OK for 2 -> 3 shifts  07/20/2006, lss
    '5) Calendar YTD     OK for 2 -> 3 shifts  07/20/2006, lss
    '6) Fiscal YTD       OK for 2 -> 3 shifts  07/20/2006, lss
    '7) Date range       OK for 2 -> 3 shifts  07/20/2006, lss
    '8) Quarter-to-date  OK for 2 -> 3 shifts  07/20/2006, lss
    
    '03/08/2007, lss -- modified to go faster (added gNeedToChangeShiftNames).
 
    Dim ShiftPos As Integer
    Dim ShiftBegHour As Integer
    Dim Fybd As Date
    Dim PeriodBegDate As Date
    
    ShiftPos = InStr(StrConv(gRptTimeFrame, vbUpperCase), "SHIFT")
    
    If ShiftPos <> 0 Then
        'We have something like "Day shift", "Night shift", "1st shift",
        '"2nd shift" or "3rd shift"
        'Will always be on gActiveDate thus can use gGetShiftBegDtime
        
        ShiftBegHour = gGetShiftStartHour(gRptTimeFrame)
        gPrtBeginDate = gGetShiftBegDtime(gActiveDate, ShiftBegHour)
                
        gPrtEndDate = gPrtBeginDate
            gRptDateRange = "( " & Format(gPrtBeginDate, "MM/dd/yyyy") & ")"
        gRptTimeFrame2 = gRptTimeFrame
        
        gPrtBeginShift = StrConv(Mid(gRptTimeFrame, 1, ShiftPos - 2), vbUpperCase)
        gPrtEndShift = StrConv(Mid(gRptTimeFrame, 1, ShiftPos - 2), vbUpperCase)
    Else
        Select Case gRptTimeFrame
            Case "Total day"
                'Will always be on gActiveDate thus can use gGetShiftBegDtime
                
                gPrtBeginDate = gGetFirstShiftBegDtime(gActiveDate)
                gPrtEndDate = gGetLastShiftBegDtime(gActiveDate)
                
                    gRptDateRange = "( " & Format(gPrtBeginDate, "MM/dd/yyyy") & ")"
                gRptTimeFrame2 = gRptTimeFrame
                       
                gPrtBeginShift = gFirstShift
                gPrtEndShift = gLastShift
                
            Case "Week-to-date"
                'The begin date in this case will not be gActiveDate
                'thus we will have to treat it specially -- will have
                'to use gGetFirstShiftBegDtime2 instead of gGetFirstShiftBegDtime
                
                    PeriodBegDate = CDate(gActiveDate.AddDays(-DatePart("w", gActiveDate, vbMonday))).AddDays(1)
                
                gPrtBeginDate = gGetFirstShiftBegDtime2(gActiveMineNameLong, _
                                                        PeriodBegDate)
                gPrtEndDate = gGetLastShiftBegDtime(gActiveDate)
                
                    gRptDateRange = "( " & Format(gPrtBeginDate, "MM/dd/yyyy") _
                               & "  to  " _
                               & Format(gPrtEndDate, "MM/dd/yyyy") & " )"
                               
                gRptTimeFrame2 = gRptTimeFrame
                
                'Is this date a Sunday?
                If Weekday(gPrtEndDate, vbMonday) = 7 Then
                    gRptTimeFrame2 = "Week Totals"
                End If
                
                If gNeedToChangeShiftNames(gActiveMineNameLong, _
                                           PeriodBegDate, _
                                           gActiveDate) = True Then
                    gPrtBeginShift = gGetFirstShiftHardCode(gActiveMineNameLong, _
                                                            PeriodBegDate)
                Else
                    gPrtBeginShift = gFirstShift
                End If
                gPrtEndShift = gLastShift
                
            Case "Month-to-date"
                'The begin date in this case will not be gActiveDate
                'thus we will have to treat it specially -- will have
                'to use gGetFirstShiftBegDtime2 instead of gGetFirstShiftBegDtime
                
                PeriodBegDate = CDate(DatePart("m", gActiveDate) & _
                                "/01/" & DatePart("yyyy", gActiveDate))
                                
                gPrtBeginDate = gGetFirstShiftBegDtime2(gActiveMineNameLong, _
                                                        PeriodBegDate)
                gPrtEndDate = gGetLastShiftBegDtime(gActiveDate)
                
                    gRptDateRange = "( " & Format(gPrtBeginDate, "MM/dd/yyyy") _
                               & "  to  " _
                               & Format(gPrtEndDate, "MM/dd/yyyy") & " )"
                               
                gRptTimeFrame2 = gRptTimeFrame
                
                'Is this last day in month?
                If LastDayInMonth(gPrtEndDate) = True Then
                    gRptTimeFrame2 = "Month Totals"
                End If
                
                If gNeedToChangeShiftNames(gActiveMineNameLong, _
                                           PeriodBegDate, _
                                           gActiveDate) = True Then
                    gPrtBeginShift = gGetFirstShiftHardCode(gActiveMineNameLong, _
                                                            PeriodBegDate)
                Else
                    gPrtBeginShift = gFirstShift
                End If
                gPrtEndShift = gLastShift
                
            Case "Calendar YTD"
                'The begin date in this case will not be gActiveDate
                'thus we will have to treat it specially -- will have
                'to use gGetFirstShiftBegDtime2 instead of gGetFirstShiftBegDtime
                
                PeriodBegDate = CDate("01/01/" & DatePart("yyyy", gActiveDate))
                                                        
                gPrtBeginDate = gGetFirstShiftBegDtime2(gActiveMineNameLong, _
                                                        PeriodBegDate)
                gPrtEndDate = gGetLastShiftBegDtime(gActiveDate)
                
                    gRptDateRange = "( " & Format(gPrtBeginDate, "MM/dd/yyyy") _
                               & "  to  " _
                               & Format(gPrtEndDate, "MM/dd/yyyy") & " )"
                               
                gRptTimeFrame2 = gRptTimeFrame
                
                'Is this the last day of the calendar year?
                If gLastDayInCalendarYear(gPrtEndDate) Then
                    gRptTimeFrame2 = "Calendar Year Totals"
                End If
                
                If gNeedToChangeShiftNames(gActiveMineNameLong, _
                                           PeriodBegDate, _
                                           gActiveDate) = True Then
                    gPrtBeginShift = gGetFirstShiftHardCode(gActiveMineNameLong, _
                                                            PeriodBegDate)
                Else
                    gPrtBeginShift = gFirstShift
                End If
                gPrtEndShift = gLastShift
                
            Case "Fiscal YTD"
                'The begin date in this case will not be gActiveDate
                'thus we will have to treat it specially -- will have
                'to use gGetFirstShiftBegDtime2 instead of gGetFirstShiftBegDtime
                
                Fybd = gGetBeginOfFiscalYear(gActiveDate)
                gPrtBeginDate = gGetFirstShiftBegDtime2(gActiveMineNameLong, _
                                                        Fybd)
                gPrtEndDate = gGetLastShiftBegDtime(gActiveDate)
                
                    gRptDateRange = "( " & Format(gPrtBeginDate, "MM/dd/yyyy") _
                               & "  to  " _
                               & Format(gPrtEndDate, "MM/dd/yyyy") & " )"
                               
                gRptTimeFrame2 = gRptTimeFrame
                
                'Is this the last day of the fiscal year?
                If gLastDayInFiscalYear(gPrtEndDate) Then
                    gRptTimeFrame2 = "Fiscal Year Totals"
                End If
                
                If gNeedToChangeShiftNames(gActiveMineNameLong, _
                                           Fybd, _
                                           gActiveDate) = True Then
                    gPrtBeginShift = gGetFirstShiftHardCode(gActiveMineNameLong, _
                                                            Fybd)
                Else
                    gPrtBeginShift = gFirstShift
                End If
                gPrtEndShift = gLastShift
                
            Case "Date range"
                'The begin and end dates in this case may not be gActiveDate
                'thus we will have to treat them specially -- will have
                'to use gGetFirstShiftBegDtime2 instead of gGetFirstShiftBegDtime
            
                gPrtBeginDate = gGetFirstShiftBegDtime2(gActiveMineNameLong, _
                                                        gCalBeginDate)
                gPrtEndDate = gGetLastShiftBegDtime2(gActiveMineNameLong, _
                                                     gCalEndDate)
                
                    gRptDateRange = "( " & Format(gPrtBeginDate, "MM/dd/yyyy") _
                               & "  to  " _
                               & Format(gPrtEndDate, "MM/dd/yyyy") & " )"
                               
                gRptTimeFrame2 = ""
                
                If gNeedToChangeShiftNames(gActiveMineNameLong, _
                                           gCalBeginDate, _
                                           gActiveDate) = True Then
                    gPrtBeginShift = gGetFirstShiftHardCode(gActiveMineNameLong, _
                                                            gCalBeginDate)
                Else
                    gPrtBeginShift = gFirstShift
                End If
                
                If gNeedToChangeShiftNames(gActiveMineNameLong, _
                                           gCalEndDate, _
                                           gActiveDate) = True Then
                    gPrtEndShift = gGetLastShiftHardCode(gActiveMineNameLong, _
                                                         gCalEndDate)
                Else
                    gPrtEndShift = gLastShift
                End If
                
            Case "Quarter-to-date"
                'The begin date in this case will not be gActiveDate
                'thus we will have to treat it specially -- will have
                'to use gGetFirstShiftBegDtime2 instead of gGetFirstShiftBegDtime
                
                PeriodBegDate = gQtrBeginDateCurr(gActiveDate)
                
                gPrtBeginDate = gGetFirstShiftBegDtime2(gActiveMineNameLong, _
                                                        PeriodBegDate)
                gPrtEndDate = gGetLastShiftBegDtime(gActiveDate)
                
                    gRptDateRange = "( " & Format(gPrtBeginDate, "MM/dd/yyyy") _
                               & "  to  " _
                               & Format(gPrtEndDate, "MM/dd/yyyy") & " )"
                               
                gRptTimeFrame2 = gRptTimeFrame
                
                'Is this the last day of a quarter?
                If gLastDayInQuarter(gPrtEndDate) <> "" Then
                    gRptTimeFrame2 = gLastDayInQuarter(gPrtEndDate) & _
                                     " Quarter Totals"
                End If
                
                If gNeedToChangeShiftNames(gActiveMineNameLong, _
                                           PeriodBegDate, _
                                           gActiveDate) = True Then
                    gPrtBeginShift = gGetFirstShiftHardCode(gActiveMineNameLong, _
                                                            PeriodBegDate)
                Else
                    gPrtBeginShift = gFirstShift
                End If
                gPrtEndShift = gLastShift
        End Select
    End If
End Sub

Public Function gGetCalYrBegin(ByVal aDate As Date) As Date

'**********************************************************************
'
'
'
'**********************************************************************

    Dim MonthStr As String
    Dim YearStr As String
    Dim DateStr As String
     
    MonthStr = DatePart("m", aDate)
    YearStr = DatePart("yyyy", aDate)
    DateStr = "01/01/" & YearStr
        
    gGetCalYrBegin = CDate(DateStr)
End Function

Public Function gGetCalYrEnd(ByVal aDate As Date) As Date

'**********************************************************************
'
'
'
'**********************************************************************

    Dim MonthStr As String
    Dim YearStr As String
    Dim DateStr As String
    
    MonthStr = DatePart("m", aDate)
    YearStr = DatePart("yyyy", aDate)
    DateStr = "12/31/" & YearStr
        
    gGetCalYrEnd = CDate(DateStr)
End Function

Public Function gGetFiscYrBegin(ByVal aDate As Date) As Date

'**********************************************************************
'
'
'
'**********************************************************************

    Dim MonthStr As String
    Dim YearStr As String
    Dim DateStr As String
    Dim MonthVal As Integer
    Dim YearVal As Integer
    
    Dim MonthDayStr As String
    Dim FyrMonthBegin As Integer
    
    'A fiscal year begin date is in gFyrBeginDate
    'This fiscal year begin date will be a date like 06/01/2004.
    'All we really want to use is the month and the day from this
    'date.
    
        MonthDayStr = Format(gFyrBeginDate, "MM/dd")
    FyrMonthBegin = DatePart("m", gFyrBeginDate)
    
    MonthStr = DatePart("m", aDate)
    MonthVal = Val(MonthStr)
    YearStr = DatePart("yyyy", aDate)
    YearVal = Val(YearStr)
        
    If MonthVal >= 1 And MonthVal <= FyrMonthBegin - 1 Then
        YearVal = YearVal - 1
    End If
    
    YearStr = Str(YearVal)
    DateStr = MonthDayStr & "/" & YearStr
        
    gGetFiscYrBegin = CDate(DateStr)
End Function

Public Function gGetFiscYrEnd(ByVal aDate As Date) As Date

'**********************************************************************
'
'
'
'**********************************************************************

    Dim MonthStr As String
    Dim YearStr As String
    Dim DateStr As String
    Dim MonthVal As Integer
    Dim YearVal As Integer
    
    Dim MonthDayBegStr As String
    Dim FyrMonthBegin As Integer
    Dim MonthDayEndStr As String
    Dim FyrMonthEnd As Integer
    Dim FyrEndDate As Date
    
    'A fiscal year begin date is in gFyrBeginDate
    'This fiscal year begin date will be a date like 06/01/2004.
    'All we really want to use is the month and the day from this
    'date.
    
        MonthDayBegStr = Format(gFyrBeginDate, "MM/dd") '"06/31"
    FyrMonthBegin = DatePart("m", gFyrBeginDate)    '6
        FyrEndDate = gFyrBeginDate.AddDays(-1)
        MonthDayEndStr = Format(FyrEndDate, "MM/dd")    '"05/31"
    FyrMonthEnd = DatePart("m", FyrEndDate)         '5
      
    MonthStr = DatePart("m", aDate)
    MonthVal = Val(MonthStr)
    YearStr = DatePart("yyyy", aDate)
    YearVal = Val(YearStr)
        
    If MonthVal >= FyrMonthBegin And MonthVal <= 12 Then
        YearVal = YearVal + 1
    End If
    
    YearStr = Str(YearVal)
    DateStr = MonthDayEndStr & "/" & YearStr
        
    gGetFiscYrEnd = CDate(DateStr)
End Function

Public Function gGetFirstDayOfWeek(ByVal aDate As Date) As Date

'**********************************************************************
'
'
'
'**********************************************************************

        'Dim MonthStr As String
        'Dim YearStr As String
        'Dim DateStr As String
 
        gGetFirstDayOfWeek = CDate(aDate.AddDays(-DatePart("w", aDate, vbMonday))).AddDays(1)
End Function

Public Function gGetLastDayOfWeek(ByVal aDate As Date) As Date

'**********************************************************************
'
'
'
'**********************************************************************

        'Dim MonthStr As String
        'Dim YearStr As String
        'Dim DateStr As String
    
    Dim FirstDayOfWeek As Date
    
        FirstDayOfWeek = CDate(aDate.AddDays(-DatePart("w", aDate, vbMonday))).AddDays(1)
        gGetLastDayOfWeek = FirstDayOfWeek.AddDays(6)
End Function

Public Function LastDayInMonth(ByVal aDate As Date) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    Dim LastDay As Date
    
    LastDay = gGetLastOfMonth(aDate)
    
    If DatePart("m", LastDay) = DatePart("m", aDate) And _
       DatePart("d", LastDay) = DatePart("d", aDate) And _
       DatePart("yyyy", LastDay) = DatePart("yyyy", aDate) Then
        LastDayInMonth = True
    Else
        LastDayInMonth = False
    End If
End Function
 
Public Function gLastDayInCalendarYear(ByVal aDate As Date) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    Dim MonthNum As Integer
    Dim DayNum As Integer
 
    MonthNum = DatePart("m", aDate)
    DayNum = DatePart("d", aDate)
     
    If DayNum = 31 And MonthNum = 12 Then
        gLastDayInCalendarYear = True
    Else
        gLastDayInCalendarYear = False
    End If
End Function

Public Function gLastDayInFiscalYear(ByVal aDate As Date) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    Dim MonthNum As Integer
    Dim DayNum As Integer
 
    MonthNum = DatePart("m", aDate)
    DayNum = DatePart("d", aDate)
     
    If DayNum = 30 And MonthNum = 6 Then
        gLastDayInFiscalYear = True
    Else
        gLastDayInFiscalYear = False
    End If
End Function
 
    Public Sub gArrayInitialize(ByRef aControlProportionsArray() As gControlProperties, _
                                ByRef aControls As Object, _
                               ByVal aScaleWidth As Integer, ByVal aScaleHeight As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim I As Integer

        On Error Resume Next

        For I = 0 To aControls.Count - 1
            If TypeOf aControls(I) Is Timer Then
                'Do nothing
            ElseIf TypeOf aControls(I) Is CommonDialog Then
                'Do nothing
            Else
                With aControlProportionsArray(I)
                    .WidthProportions = aControls(I).Width / aScaleWidth
                    .HeightProportions = aControls(I).Height / aScaleHeight
                    .LeftProportions = aControls(I).Left / aScaleWidth
                    .TopProportions = aControls(I).Top / aScaleHeight
                    .FontProportions = aControls(I).FontSize / aScaleHeight
                End With
            End If
        Next I
    End Sub

    Public Sub gFormResize(ByRef aControlProportionsArray() As gControlProperties, _
                           ByRef aControls As Object, _
                           ByVal aScaleWidth As Integer, _
                           ByVal aScaleHeight As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim I As Integer

        On Error Resume Next

        For I = 0 To aControls.Count - 1
            If TypeOf aControls(I) Is Timer Then
                'Do nothing
            ElseIf TypeOf aControls(I) Is CommonDialog Then
                'Do nothing

            ElseIf TypeOf aControls(I) Is ComboBox Then
                aControls(I).Move(aControlProportionsArray(I).LeftProportions * _
                aScaleWidth, _
                aControlProportionsArray(I).TopProportions * aScaleHeight, _
                aControlProportionsArray(I).WidthProportions * aScaleWidth)

            Else
                aControls(I).Move(aControlProportionsArray(I).LeftProportions * _
                aScaleWidth, _
                aControlProportionsArray(I).TopProportions * aScaleHeight, _
                aControlProportionsArray(I).WidthProportions * aScaleWidth, _
                aControlProportionsArray(I).HeightProportions * aScaleHeight)

                If TypeOf aControls(I) Is AxvaSpread Then
                    'Do nothing
                Else
                    aControls(I).FontSize = Round((aControlProportionsArray(I).FontProportions * _
                                            aScaleHeight) * 0.95, 2)
                End If
            End If
        Next I
    End Sub

Public Function gConvertToShiftDate(sDate As Date, _
                                    sShift As String) As Date

'**********************************************************************
'   This routine converts a date and shift to a shiftdate using the
'   mois.check_shifts Oracle function. This function assumes the mine
'   name is gActiveMineNameLong.
'
'   This function is not tied to the 2 shift ("Day", "Night") scheme.
'   It will work for 3 shift mines also.
'**********************************************************************

    Dim SQLString As String
   'Set up SQL command.
   
   sShift = StrConv(sShift, vbUpperCase)
   SQLString = _
        "Begin :pShiftDate := mois.check_shifts" & _
        "(:pMineName, " & _
        ":pDate, " & _
        ":pShift);  " & _
        "End;"
    
    'Execute command with parameter information passed in as an array
        'of arrays.
        Dim arA1() As Object = {"pMineName", gActiveMineNameLong, ORAPARM_INPUT, ORATYPE_VARCHAR2} ') As Array
        Dim arA2() As Object = {"pDate", sDate.ToString, ORAPARM_INPUT, ORATYPE_DATE} ') As Array
        Dim arA3() As Object = {"pShift", sShift, ORAPARM_INPUT, ORATYPE_VARCHAR2} ') As Array
        Dim arA4() As Object = {"pShiftDate", 0, ORAPARM_OUTPUT, ORATYPE_DATE} ') As Array
        gConvertToShiftDate = RunSPReturnDate _
            ( _
                SQLString, _
                 arA1, _
                arA2, _
                arA3, _
                arA4 _
            )
        '         Array[], _
        'Array("pDate", sDate, ORAPARM_INPUT, ORATYPE_DATE), _
        'Array("pShift", sShift, ORAPARM_INPUT, ORATYPE_VARCHAR2), _
        'Array("pShiftDate", 0, ORAPARM_OUTPUT, ORATYPE_DATE) _


    End Function

Public Function gRound3(sNum As Double, _
                        sPrecision As Integer) As Double

'**********************************************************************
'
'
'
'**********************************************************************

    Dim Z1 As Long
    
    Z1 = 10 ^ (sPrecision + 2) * Abs(sNum)
    If Z1 Mod 100 < 50 Then
        Z1 = Z1 - Z1 Mod 100
    Else
        Z1 = Z1 + (100 - Z1 Mod 100)
    End If
    If sNum > 0 Then
        gRound3 = Z1 / 10 ^ (sPrecision + 2)
    Else
        gRound3 = -1 * (Z1 / 10 ^ (sPrecision + 2))
    End If
End Function

Public Function gRound2(sNum As Double, _
                        sPrecision As Integer) As Double

'**********************************************************************
'
'
'
'**********************************************************************

    If sNum > 0 Then
        gRound2 = Int(Round(10 ^ sPrecision * Abs(sNum), 1) + 0.5) / 10 ^ sPrecision
    Else
        gRound2 = -1 * (Int(Round(10 ^ sPrecision * Abs(sNum), 1) + 0.5) / 10 ^ sPrecision)
    End If
End Function

Public Function gGetEndOfFiscalYear(ByVal aDate As Date) As Date

'**********************************************************************
'This function returns the last day of the fiscal year that aDate is
'in.
'
'**********************************************************************
    
    Dim FyrEndDate As Date
    Dim MonthDayBegStr As String
    Dim MonthDayEndStr As String
    
    'A fiscal year begin date is in gFyrBeginDate
    'This fiscal year begin date will be a date like 06/01/2004.
    'All we really want to use is the month and the day from this
    'date.
    
        FyrEndDate = gFyrBeginDate.AddDays(-1)
        MonthDayBegStr = Format(gFyrBeginDate, "MM/dd")
        MonthDayEndStr = Format(FyrEndDate, "MM/dd")
    
    If aDate <= CDate("12/31/" & DatePart("yyyy", aDate)) And _
        aDate >= CDate(MonthDayBegStr & "/" & DatePart("yyyy", aDate)) Then
        gGetEndOfFiscalYear = CDate(MonthDayEndStr & "/" & DatePart("yyyy", DateAdd("yyyy", 1, aDate)))
    Else
        gGetEndOfFiscalYear = CDate(MonthDayEndStr & "/" & DatePart("yyyy", aDate))
    End If
End Function

Public Function gGetBeginOfFiscalYear(ByVal aDate As Date) As Date

'**********************************************************************
'This function returns the first day of the fiscal year that aDate is
'in.
'
'**********************************************************************

    Dim MonthDayStr As String
    
    'A fiscal year begin date is in gFyrBeginDate
    'This fiscal year begin date will be a date like 06/01/2004.
    'All we really want to use is the month and the day from this
    'date.
    
        MonthDayStr = Format(gFyrBeginDate, "MM/dd")
    
    If aDate >= CDate(MonthDayStr & "/" & DatePart("yyyy", aDate)) Then
        gGetBeginOfFiscalYear = CDate(MonthDayStr & "/" & DatePart("yyyy", aDate))
    Else
        gGetBeginOfFiscalYear = CDate(MonthDayStr & "/" & DatePart("yyyy", DateAdd("yyyy", -1, aDate)))
    End If
End Function

Public Sub gPrintSpecialShiftReport(ByVal aMineName As String, _
                                    ByVal aActiveDate As Date, _
                                    ByVal aActiveShift As String, _
                                    ByVal aPath As String, _
                                    ByVal aReport As String, _
                                    ByRef aReportControl As Object, _
                                    ByVal aEqptName As String)
                                        
'**********************************************************************
'
'
'
'**********************************************************************
                                        
    On Error GoTo gPrintSpecialShiftReportError
    
    Dim ConnectString As String
    Dim ShiftValue As String
    Dim ThisSupervisor As String
    Dim ReportString As String
    Dim NumUnits As Integer
    Dim UsesTypes As Boolean
    Dim SubRptCount As Integer
    
    UsesTypes = SrptUsesTypes(aReport)
    ThisSupervisor = gGetSrptSupervisor(aActiveDate, _
                                        StrConv(aActiveShift, vbUpperCase), _
                                        aReport)

    Select Case aReport
        Case Is = "Utility Operator Shift Report"
            SubRptCount = 0
            
        Case Is = "DL Inspection Shift Report"
            SubRptCount = 0
            
        Case Is = "Pump Inspection Shift Report"
            SubRptCount = 0
            
        Case Is = "Washer Shift Report"
            SubRptCount = 1
            
        Case Is = "Float Plant Shift Report"
            SubRptCount = 1
            
        Case Is = "Sizing Shift Report"
            SubRptCount = 1
            
        Case Is = "Reagent Area Shift Report"
            SubRptCount = 0
            
        Case Is = "Shipping Shift Report"
            SubRptCount = 0
    End Select
        
    aReportControl.Reset
    
    If UsesTypes = True Then
        aReportControl.ReportFileName = gPath + "\Reports\" + _
                                       "DlInspect.rpt"
    Else
        aReportControl.ReportFileName = gPath + "\Reports\" + _
                                        "UtilOpers.rpt"
    End If
            
    'Connect to Oracle database
    ConnectString = "DSN = " + gDataSource + ";UID = " + gOracleUserName + _
        ";PWD = " + gOracleUserPassword + ";DSQ = "
        
    'Use SelectionFormula property to select for Unit
    If aEqptName <> "All" Then
        ReportString = "{GET_SPECIAL_SHIFT_RPT_ROWS.EQPT_NAME} = '" + aEqptName + "'"
        NumUnits = 1
    Else
        ReportString = ""
        NumUnits = gGetNumUnits(aMineName, aReport, _
                                aActiveDate, aActiveShift)
    End If
    aReportControl.SelectionFormula = ReportString
         
    aReportControl.Connect = ConnectString
    'Report window maximized
        ' aReportControl.WindowState =   crptMaximized
         
    aReportControl.WindowTitle = aReport
           
    'User not allowed to minimize report window
    aReportControl.WindowMinButton = False
        
    ShiftValue = StrConv(aActiveShift, vbUpperCase)
           
    'Report needs number of units
    aReportControl.Formulas(1) = "NumUnits = '" & NumUnits & "'"
        
    'pMineName
    'pRptTypeName
    'pRptName
    'pEqptTypeName
    'pEqptName
    'pShiftDate
    'pShift
     
    aReportControl.ParameterFields(0) = "pMineName;" & aMineName & ";TRUE"
    aReportControl.ParameterFields(1) = "pRptTypeName;" & "Special shift report" & ";TRUE"
    aReportControl.ParameterFields(2) = "pRptName;" & aReport & ";TRUE"
        aReportControl.ParameterFields(3) = "pShiftDate;" & _
                                             Format(aActiveDate, "MM/dd/yyyy") & ";TRUE"
    aReportControl.ParameterFields(4) = "pShift;" & ShiftValue & ";TRUE"
    aReportControl.ParameterFields(5) = "pResult;" & " " & ";TRUE"
    aReportControl.ParameterFields(6) = "pSupervisor;" & ThisSupervisor & ";TRUE"

    'New special shift report parameters -- 06/30/2003, lss
    aReportControl.ParameterFields(7) = "pSftyCommLbl1;" & mSftyComm1Lbl & ";TRUE"
    aReportControl.ParameterFields(8) = "pSftyCommLbl2;" & mSftyComm2Lbl & ";TRUE"
    aReportControl.ParameterFields(9) = "pWorkAreaSafe1;" & mSafeAreaComm1Lbl & ";TRUE"
    aReportControl.ParameterFields(10) = "pWorkAreaSafe2;" & mSafeAreaComm2Lbl & ";TRUE"
    aReportControl.ParameterFields(11) = "pNumSubRpts;" & SubRptCount & ";TRUE"

    'Need to pass the company name into the report
    aReportControl.ParameterFields(12) = "pCompanyName;" & gCompanyName & ";TRUE"
    
    '07/19/2010, lss  Added this stuff...
    Dim SubNum As Integer
    Dim SubIdx As Integer
    If SubRptCount = 1 Then
        SubNum = aReportControl.GetNSubreports
        If SubNum > 0 Then
            For SubIdx = 0 To SubNum - 1
                aReportControl.SubreportToChange = aReportControl.GetNthSubreportName(SubIdx)
                aReportControl.Connect = ConnectString
            Next
            aReportControl.SubreportToChange = ""
        End If
    End If
    
    'Start Crystal Reports
    aReportControl.action = 1
       
    aReportControl.Reset
    
    Exit Sub
    
gPrintSpecialShiftReportError:
        MsgBox("Error printing special shift report." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Special Shift Report Print Error")
End Sub

Private Function SrptUsesTypes(ByVal aRptName As String) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo SrptUsesTypesError
    
    'Need to get SFTY_COMM1_LBL, SFTY_COMM2_LBL, SAFEAREA_COMM1_LBL,
    'SAFEAREA_COMM2_LBL.
    'Place data in mSftyComm1Lbl, mSftyComm2Lbl, mSafeAreaComm1Lbl,
    'mSafeAreaComm2Lbl
        
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    
    Dim ThisUsesLoctnTypes As Integer
        SrptUsesTypes = False
    mSftyComm1Lbl = ""
    mSftyComm2Lbl = ""
    mSafeAreaComm1Lbl = ""
    mSafeAreaComm2Lbl = ""

    'Get basic special shift report info
        'Set 
        params = gDBParams
    
        params.Add("pMineName", gActiveMineNameLong, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pRptTypeName", "Special shift report", ORAPARM_INPUT)
    params("pRptTypeName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pRptName", aRptName, ORAPARM_INPUT)
    params("pRptName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
       
    'PROCEDURE get_srpt_rpts
    'pMineName      IN     VARCHAR2,
    'pRptTypeName   IN     VARCHAR2,
    'pRptName       IN     VARCHAR2,
    'pResult        IN OUT c_reports
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_srpt.get_srpt_rpts(:pMineName," + _
                  ":pRptTypename, :pRptName, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        mRptNameDynaset = params("pResult").Value
        ClearParams(params)
            
    If mRptNameDynaset.RecordCount = 1 Then
        mRptNameDynaset.MoveFirst
        
        ThisUsesLoctnTypes = mRptNameDynaset.Fields("uses_types").Value
        If ThisUsesLoctnTypes = 1 Then
            SrptUsesTypes = True
        Else
            SrptUsesTypes = False
        End If
        
            If Not IsDBNull(mRptNameDynaset.Fields("sfty_comm1_lbl").Value) Then
                mSftyComm1Lbl = mRptNameDynaset.Fields("sfty_comm1_lbl").Value
            Else
                mSftyComm1Lbl = ""
            End If
            If Not IsDBNull(mRptNameDynaset.Fields("sfty_comm2_lbl").Value) Then
                mSftyComm2Lbl = mRptNameDynaset.Fields("sfty_comm2_lbl").Value
            Else
                mSftyComm2Lbl = ""
            End If
            If Not IsDBNull(mRptNameDynaset.Fields("sfty_comm1_lbl").Value) Then
                mSafeAreaComm1Lbl = mRptNameDynaset.Fields("safearea_comm1_lbl").Value
            Else
                mSafeAreaComm1Lbl = ""
            End If
            If Not IsDBNull(mRptNameDynaset.Fields("sfty_comm2_lbl").Value) Then
                mSafeAreaComm2Lbl = mRptNameDynaset.Fields("safearea_comm2_lbl").Value
            Else
                mSafeAreaComm2Lbl = ""
            End If
    End If
     
    Exit Function
    
SrptUsesTypesError:
        MsgBox("Error getting special shift report info" & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Special Shift Report Info Get Error")
        
    On Error Resume Next
        ClearParams(params)
End Function

Public Function gGetSupervisor(ByVal aDate As Date, _
                               ByVal aShift As String) As String

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo GetSupervisorError
    
    Dim ThisMeasureName As String
    Dim RecCount As Integer
 
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim MineInfoDynaset As OraDynaset
    
    Dim ThisMineName As String
    
    'get all existing mine data (Eqpt type = "Mine", Eqpt = "Mine")
    
        'Set 
        params = gDBParams
 
        params.Add("pMineName", gActiveMineNameLong, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
 
    params.Add("pShiftDate", aDate, ORAPARM_INPUT)
    params("pShiftDate").serverType = ORATYPE_DATE
    
    params.Add("pShift", aShift, ORAPARM_INPUT)
    params("pShift").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_all_mine_data
    'pMineName
    'pShiftDate
    'pShift
    'pResult
                                         
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_mine_data(:pMineName, " + _
                  ":pShiftDate, :pShift, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MineInfoDynaset = params("pResult").Value
     
    MineInfoDynaset.MoveFirst
    RecCount = MineInfoDynaset.RecordCount
    
    gGetSupervisor = "Not entered"
    
    Do While Not MineInfoDynaset.EOF
        ThisMeasureName = MineInfoDynaset.Fields("measure_name").Value
        
        If ThisMeasureName = "Field supervisor's name" Then
                If Not IsDBNull(MineInfoDynaset.Fields("value").Value) Then
                    gGetSupervisor = MineInfoDynaset.Fields("value").Value
                Else
                    gGetSupervisor = "Not entered"
                End If
        End If
        MineInfoDynaset.MoveNext
    Loop
          
    If Trim(gGetSupervisor) = "" Then
        gGetSupervisor = "Not entered"
    End If
            
        ClearParams(params)
    MineInfoDynaset.Close
    
    Exit Function
    
GetSupervisorError:
        MsgBox("Error getting supervisor." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Supervisor Get Error")
        
    On Error Resume Next
        ClearParams(params)
    MineInfoDynaset.Close
End Function

Public Function gGetNumUnits(ByVal aMineName As String, _
                             ByVal aReport As String, _
                             ByVal aProdDate As Date, _
                             ByVal aShift As String) As Integer

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo GetNumUnitsError
    
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
        Dim ShiftDateValue As Object
        Dim ShiftValue As Object
    Dim RecordCount As Integer
    Dim UnitShiftDynaset As OraDynaset
    
    'Get number of units
    
        'Set 
        params = gDBParams
     
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
     
    params.Add("pRptTypeName", "Special shift report", ORAPARM_INPUT)
    params("pRptTypeName").serverType = ORATYPE_VARCHAR2
        
    params.Add("pRptName", aReport, ORAPARM_INPUT)
    params("pRptName").serverType = ORATYPE_VARCHAR2
        
    ShiftDateValue = CDate(aProdDate)
    params.Add("pShiftDate", ShiftDateValue, ORAPARM_INPUT)
    params("pShiftDate").serverType = ORATYPE_DATE
    
    ShiftValue = aShift
    
    If ShiftValue = "Day" Then
        ShiftValue = "DAY"
    End If
    If ShiftValue = "Night" Then
        ShiftValue = "NIGHT"
    End If
    
    params.Add("pShift", ShiftValue, ORAPARM_INPUT)
    params("pShift").serverType = ORATYPE_VARCHAR2
        
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_shift_units
    'pMineName
    'pRptTypeName
    'pRptName
    'pShiftDate
    'pShift
    'pResult
    
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_srpt.get_shift_units(:pMineName," + _
                     ":pRptTypeName, :pRptName, :pShiftDate," + _
                     ":pShift, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        UnitShiftDynaset = params("pResult").Value
    RecordCount = UnitShiftDynaset.RecordCount
    
    gGetNumUnits = RecordCount
    
        ClearParams(params)
    UnitShiftDynaset.Close
    
    Exit Function
    
GetNumUnitsError:
        MsgBox("Error getting number of units." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Number of Units Get Error")
        
    On Error Resume Next
        ClearParams(params)
    UnitShiftDynaset.Close
End Function

Public Function gGetMonthAndYear(ByVal aDate As Date, _
                                 ByVal aIncludeSpace As Boolean) As String

'**********************************************************************
'
'
'
'**********************************************************************

    Dim YearValue As Integer
    Dim MonthValue As String
    Dim MonthName As String
    
    YearValue = Year(aDate)
    MonthValue = Month(aDate)
    
    Select Case MonthValue
        Case Is = 1
            MonthName = "January"
        Case Is = 2
            MonthName = "February"
        Case Is = 3
            MonthName = "March"
        Case Is = 4
            MonthName = "April"
        Case Is = 5
            MonthName = "May"
        Case Is = 6
            MonthName = "June"
        Case Is = 7
            MonthName = "July"
        Case Is = 8
            MonthName = "August"
        Case Is = 9
            MonthName = "September"
        Case Is = 10
            MonthName = "October"
        Case Is = 11
            MonthName = "November"
        Case Is = 12
            MonthName = "December"
    End Select
    
    If aIncludeSpace = True Then
        gGetMonthAndYear = MonthName + ", " + Str(YearValue)
    Else
        gGetMonthAndYear = MonthName + "," + Str(YearValue)
    End If
End Function

Public Function gGetMonthAndYearAbbrv(ByVal aDate As Date, _
                                      ByVal aIncludeSpace As Boolean, _
                                      ByVal aNoComma As Boolean) As String

'**********************************************************************
'
'
'
'**********************************************************************

    Dim YearValue As Integer
    Dim MonthValue As String
    Dim MonthName As String
    
    YearValue = Year(aDate)
    MonthValue = Month(aDate)
    
    Select Case MonthValue
        Case Is = 1
            MonthName = "Jan"
        Case Is = 2
            MonthName = "Feb"
        Case Is = 3
            MonthName = "Mar"
        Case Is = 4
            MonthName = "Apr"
        Case Is = 5
            MonthName = "May"
        Case Is = 6
            MonthName = "Jun"
        Case Is = 7
            MonthName = "Jul"
        Case Is = 8
            MonthName = "Aug"
        Case Is = 9
            MonthName = "Sep"
        Case Is = 10
            MonthName = "Oct"
        Case Is = 11
            MonthName = "Nov"
        Case Is = 12
            MonthName = "Dec"
    End Select
     
    If aNoComma = False Then
        If aIncludeSpace = True Then
            gGetMonthAndYearAbbrv = MonthName + ", " + Str(YearValue)
        Else
            gGetMonthAndYearAbbrv = MonthName + "," + Str(YearValue)
        End If
    Else
        gGetMonthAndYearAbbrv = Trim(MonthName) + " " + Trim(Str(YearValue))
    End If
End Function

Public Function gPadLeft(ByVal aString As String, _
                         ByVal aLength As Integer) As String

'**********************************************************************
'
'
'
'**********************************************************************
    
    gPadLeft = Trim(aString)
    
    Do While Len(gPadLeft) < aLength
        gPadLeft = " " + gPadLeft
    Loop
End Function

Public Function gPadRight(ByVal aString As String, _
                          ByVal aLength As Integer) As String

'**********************************************************************
'
'
'
'**********************************************************************
    
    gPadRight = Trim(aString)
    
    Do While Len(gPadRight) < aLength
        gPadRight = gPadRight + " "
    Loop
End Function

Public Function gIsEvenNumber(ByVal aValue As Double) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************
    
    gIsEvenNumber = False
    
    If aValue Mod 2 = 0 Then
        gIsEvenNumber = True
    End If
End Function

Public Function gPasswordIsValid(ByVal aPassword As String, _
                                 ByVal aUserName As String) As String

'**********************************************************************
'
'
'
'**********************************************************************
    
    Dim BlankCnt As Integer
    Dim DigitCnt As Integer
    Dim AlphaCnt As Integer
    Dim CharCnt As Integer
    Dim BogusCnt As Integer
    Dim FirstCharIsDigit As Boolean
    Dim AlphaStr As String
    Dim ThisChar As String
    
    AlphaStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    gPasswordIsValid = ""
     
    'Currently enforced
    '------------------
    '1) Password must be at least 6 characters.
    '2) Password must contain at least 2 digits.
    '3) Password must not begin with a digit -- it must begin with a character.
    '4) Password must not contain any punctuations.
    '5) Password must not contain any blanks.
    '6) Password may not contain the word "Cargill".
    '7) Password may not contain the word "Password".
    '8) Password may not contain the word "System".
    '9) Password may not contain User Name.
    
    'May be enforced at a later date
    '-------------------------------
    'Password must not be the same as the last password.
    
    'See if there are any blanks on the password
    
    BlankCnt = 0
    DigitCnt = 0
    AlphaCnt = 0
    CharCnt = 0
    FirstCharIsDigit = False
    
    For CharCnt = 1 To Len(aPassword)
        ThisChar = Mid(aPassword, CharCnt, 1)
        
        If IsNumeric(ThisChar) Then
            DigitCnt = DigitCnt + 1
            
            If CharCnt = 1 Then
                FirstCharIsDigit = True
            End If
        End If
        
        If ThisChar = " " Then
            BlankCnt = BlankCnt + 1
        End If
        
        If InStr(AlphaStr, StrConv(ThisChar, vbUpperCase)) <> 0 Then
            AlphaCnt = AlphaCnt + 1
        End If
    Next CharCnt
        
    '5) Password must not contain any blanks.
    If BlankCnt <> 0 Then
        gPasswordIsValid = "No blanks allowed in password!"
        Exit Function
    End If
    
    '1) Password must be at least 6 characters.
    If Len(aPassword) < 6 Then
        gPasswordIsValid = "Password length must be >= 6"
    End If
    
    '2) Password must contain at least 2 digits.
    If DigitCnt < 2 Then
        gPasswordIsValid = "Must have at least 2 digits in password!"
        Exit Function
    End If
    
    '3) Password must not begin with a digit -- it must begin with a character.
    If FirstCharIsDigit = True Then
        gPasswordIsValid = "Password may not start with a digit!"
        Exit Function
    End If
    
    '4) Password must not contain any punctuations.
    If AlphaCnt + DigitCnt <> Len(aPassword) Then
        gPasswordIsValid = "Password must not contain any punctuations."
        Exit Function
    End If
    
    '6) Password may not contain the word "Cargill".
    If InStr(StrConv(aPassword, vbUpperCase), "CARGILL") <> 0 Then
        gPasswordIsValid = "Password may not contain the word 'Cargill'."
        Exit Function
    End If
    
    '7) Password may not contain the word "Password".
    If InStr(StrConv(aPassword, vbUpperCase), "PASSWORD") <> 0 Then
        gPasswordIsValid = "Password may not contain the word 'Password'."
        Exit Function
    End If
    
    '8) Password may not contain the word "System".
    If InStr(StrConv(aPassword, vbUpperCase), "SYSTEM") <> 0 Then
        gPasswordIsValid = "Password may not contain the word 'System'."
        Exit Function
    End If
    
    '9) Password may not contain User Name.
    If InStr(StrConv(aPassword, vbUpperCase), _
        StrConv(aUserName, vbUpperCase)) <> 0 Then
        gPasswordIsValid = "Password may not contain your User Name."
        Exit Function
    End If
End Function

Public Function gSrptSetup() As Boolean

'**********************************************************************
'
'
'
'**********************************************************************
    
    If gUserPermission.DlInspectionReport = "Setup" Or _
        gUserPermission.WasherShiftReport = "Setup" Or _
        gUserPermission.FloatPlantShiftReport = "Setup" Or _
        gUserPermission.SizingShiftReport = "Setup" Or _
        gUserPermission.ReagentShiftReport = "Setup" Or _
        gUserPermission.ShippingShiftReport = "Setup" Or _
        gUserPermission.UtilityOperatorReport = "Setup" Or _
        gUserPermission.PumpInspectionReport = "Setup" Or _
        gUserPermission.PumpPackShiftReport = "Setup" Then
        gSrptSetup = True
    Else
        gSrptSetup = False
    End If
End Function

Public Function gSrptSetupOrRead() As Boolean

'**********************************************************************
'
'
'
'**********************************************************************
    
    If gUserPermission.DlInspectionReport = "Setup" Or _
        gUserPermission.WasherShiftReport = "Setup" Or _
        gUserPermission.FloatPlantShiftReport = "Setup" Or _
        gUserPermission.SizingShiftReport = "Setup" Or _
        gUserPermission.ReagentShiftReport = "Setup" Or _
        gUserPermission.ShippingShiftReport = "Setup" Or _
        gUserPermission.UtilityOperatorReport = "Setup" Or _
        gUserPermission.PumpInspectionReport = "Setup" Or _
        gUserPermission.PumpPackShiftReport = "Setup" Or _
        gUserPermission.DlInspectionReport = "Read" Or _
        gUserPermission.WasherShiftReport = "Read" Or _
        gUserPermission.FloatPlantShiftReport = "Read" Or _
        gUserPermission.SizingShiftReport = "Read" Or _
        gUserPermission.ReagentShiftReport = "Read" Or _
        gUserPermission.ShippingShiftReport = "Read" Or _
        gUserPermission.UtilityOperatorReport = "Read" Or _
        gUserPermission.PumpInspectionReport = "Read" Or _
        gUserPermission.PumpPackShiftReport = "Read" Then
        gSrptSetupOrRead = True
    Else
        gSrptSetupOrRead = False
    End If
End Function

Public Function gSrptThisRptWriteOk(ByVal aSrpt As String) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************
    
    '1) gUserPermission.DlInspectionReport   (This covers the Dredge
    '                                         Inspection Report too.)
    '2) gUserPermission.WasherShiftReport
    '3) gUserPermission.FloatPlantShiftReport
    '4) gUserPermission.SizingShiftReport
    '5) gUserPermission.ReagentShiftReport
    '6) gUserPermission.ShippingShiftReport
    '7) gUserPermission.UtilityOperatorReport
    '8) gUserPermission.PumpInspectionReport
    '9) gUserPermission.PumpPackShiftReport
       
    gSrptThisRptWriteOk = False
   
    Select Case aSrpt
        Case Is = "DL Inspection Shift Report"
            If gUserPermission.DlInspectionReport = "Setup" Then
                gSrptThisRptWriteOk = True
            End If
                
        Case Is = "Dredge Inspection Shift Report"
            If gUserPermission.DlInspectionReport = "Setup" Then
                gSrptThisRptWriteOk = True
            End If
            
        Case Is = "Washer Shift Report"
            If gUserPermission.WasherShiftReport = "Setup" Then
                gSrptThisRptWriteOk = True
            End If
                
        Case Is = "Float Plant Shift Report"
            If gUserPermission.FloatPlantShiftReport = "Setup" Then
                gSrptThisRptWriteOk = True
            End If
                
        Case Is = "Sizing Shift Report"
            If gUserPermission.SizingShiftReport = "Setup" Then
                gSrptThisRptWriteOk = True
            End If
                
        Case Is = "Reagent Area Shift Report"
            If gUserPermission.ReagentShiftReport = "Setup" Then
                gSrptThisRptWriteOk = True
            End If
                        
        Case Is = "Shipping Shift Report"
            If gUserPermission.ShippingShiftReport = "Setup" Then
                gSrptThisRptWriteOk = True
            End If
                
        Case Is = "Utility Operator Shift Report"
            If gUserPermission.UtilityOperatorReport = "Setup" Then
                gSrptThisRptWriteOk = True
            End If
                          
        Case Is = "Pump Inspection Shift report"
            If gUserPermission.PumpInspectionReport = "Setup" Then
                gSrptThisRptWriteOk = True
            End If
            
        Case Is = "Field Pump Pack Shift Report"
            If gUserPermission.PumpPackShiftReport = "Setup" Then
                gSrptThisRptWriteOk = True
            End If
    End Select
End Function

Public Function gSrptThisCtgryName(ByVal aSrpt As String) As String

'**********************************************************************
'
'
'
'**********************************************************************

    gSrptThisCtgryName = ""
    
    Select Case aSrpt
        Case Is = "DL Inspection Shift Report"
            gSrptThisCtgryName = "Dragline inspection"
            
        Case Is = "Dredge Inspection Shift Report"
            gSrptThisCtgryName = "Dredge inspection"
            
        Case Is = "Washer Shift Report"
            gSrptThisCtgryName = "Washer shift report"
            
        Case Is = "Float Plant Shift Report"
            gSrptThisCtgryName = "Float plant shift report"
            
        Case Is = "Sizing Shift Report"
            gSrptThisCtgryName = "Sizing shift report"
            
        Case Is = "Reagent Area Shift Report"
            gSrptThisCtgryName = "Reagent area shift report"
    
        Case Is = "Shipping Shift Report"
            gSrptThisCtgryName = "Shipping shift report"
            
        Case Is = "Utility Operator Shift Report"
            gSrptThisCtgryName = "Location"
            
        Case Is = "Pump Inspection Shift Report"
            gSrptThisCtgryName = "Pump inspection"
            
        Case Is = "Field Pump Pack Shift Report"
            gSrptThisCtgryName = "Pump pack shift report"
    End Select
End Function

Public Sub gSetSrptGlobalPermissions()

'**********************************************************************
'
'
'
'**********************************************************************

    'Need to check if user is in SRPT_USERS_RPT
    
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim RecordCount As Integer
    Dim ThisReport As String
       
    On Error GoTo gSetSrptGlobalPermissionsError
    
    'Get all special shift report permissions for this user
    'Special Shift Report user names are all uppercase
    
        'Set 
        params = gDBParams
    
    params.Add("pMineName", gActiveMineNameLong, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
    
    params.Add("pUserName", StrConv(gSrptUserName, vbUpperCase), ORAPARM_INPUT)
    params("pUserName").serverType = ORATYPE_VARCHAR2
   
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR

    'PROCEDURE get_srpt_user_rpt
    'pMineName      IN     VARCHAR2,
    'pUserName      IN     VARCHAR2,
    'pResult        IN OUT c_users)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_srpt.get_srpt_user_rpt(:pMineName," + _
                  ":pUserName, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        mSrptPermissionsDynaset = params("pResult").Value
        ClearParams(params)
    
    RecordCount = mSrptPermissionsDynaset.RecordCount
        
    mSrptPermissionsDynaset.MoveFirst
  
    gSrptUorOk = False
    gSrptDlOk = False
    gSrptPumpOk = False
    gSrptWasherOk = False
    gSrptFltPltOk = False
    gSrptSizingOk = False
    gSrptReagentOk = False
    gSrptShippingOk = False
    gSrptPumpPackOk = False

    Do While Not mSrptPermissionsDynaset.EOF
        ThisReport = mSrptPermissionsDynaset.Fields("rpt_name").Value
        
        Select Case ThisReport
            Case Is = "Utility Operator Shift Report"
                gSrptUorOk = True
                
            Case Is = "DL Inspection Shift Report"
                gSrptDlOk = True
                
            Case Is = "Dredge Inspection Shift Report"
                gSrptDlOk = True
                
            Case Is = "Pump Inspection Shift Report"
                gSrptPumpOk = True
                
            Case Is = "Washer Shift Report"
                gSrptWasherOk = True
                
            Case Is = "Float Plant Shift Report"
                gSrptFltPltOk = True
                
            Case Is = "Sizing Shift Report"
                gSrptSizingOk = True
                
            Case Is = "Reagent Area Shift Report"
                gSrptReagentOk = True
                
            Case Is = "Shipping Shift Report"
                gSrptShippingOk = True
                
            Case Is = "Field Pump Pack Shift Report"
                gSrptPumpPackOk = True
        End Select
        
        mSrptPermissionsDynaset.MoveNext
    Loop

    Exit Sub
    
gSetSrptGlobalPermissionsError:
        MsgBox(Err.Description)
    
    On Error Resume Next
        ClearParams(params)
End Sub
 
Public Function gGetSrptSupervisor(ByVal aDate, _
                                   ByVal aShift, _
                                   ByVal aRptName) As String

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo gGetSupervisorError
    
    Dim ThisMeasureName As String
    Dim ShiftDateValue As Date
    Dim ShiftValue As String
    Dim RecCount As Integer
    
    Dim FieldSuper As String
    Dim PlantSuper As String
    
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    
    Dim ThisMineName As String
    
    'get all existing mine data (Eqpt type = "Mine", Eqpt = "Mine")
    
        'Set 
        params = gDBParams
 
    params.Add("pMineName", gActiveMineNameLong, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
        
    ShiftDateValue = CDate(aDate)
    params.Add("pShiftDate", ShiftDateValue, ORAPARM_INPUT)
    params("pShiftDate").serverType = ORATYPE_DATE
    
    If StrConv(aShift, vbUpperCase) = "DAY SHIFT" Or _
        StrConv(aShift, vbUpperCase) = "DAY" Then
        ShiftValue = "DAY"
    Else
        ShiftValue = "NIGHT"
    End If
    
    params.Add("pShift", ShiftValue, ORAPARM_INPUT)
    params("pShift").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_all_mine_data
    'pMineName
    'pShiftDate
    'pShift
    'pResult
                                         
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_mine_data(:pMineName, " + _
                  ":pShiftDate, :pShift, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        mMineNameDynaset = params("pResult").Value
     
    mMineNameDynaset.MoveFirst
    RecCount = mMineNameDynaset.RecordCount
    
    gGetSrptSupervisor = "Not entered"
    
    Do While Not mMineNameDynaset.EOF
        ThisMeasureName = mMineNameDynaset.Fields("measure_name").Value
        
        If ThisMeasureName = "Field supervisor's name" Then
                If Not IsDBNull(mMineNameDynaset.Fields("value").Value) Then
                    FieldSuper = mMineNameDynaset.Fields("value").Value
                Else
                    FieldSuper = "Not entered"
                End If
        End If
        
        If ThisMeasureName = "Plant supervisor's name" Then
                If Not IsDBNull(mMineNameDynaset.Fields("value").Value) Then
                    PlantSuper = mMineNameDynaset.Fields("value").Value
                Else
                    PlantSuper = "Not entered"
                End If
        End If
        mMineNameDynaset.MoveNext
    Loop
       
    'Now need to assign gGetSrptSupervisor.
    Select Case aRptName
        Case Is = "Utility Operator Shift Report"
            gGetSrptSupervisor = FieldSuper
                              
        Case Is = "DL Inspection Shift Report"
            gGetSrptSupervisor = FieldSuper

        Case Is = "Pump Inspection Shift Report"
            gGetSrptSupervisor = FieldSuper
            
        Case Is = "Washer Shift Report"
            gGetSrptSupervisor = PlantSuper
            
        Case Is = "Float Plant Shift Report"
            gGetSrptSupervisor = PlantSuper
            
        Case Is = "Sizing Shift Report"
            gGetSrptSupervisor = PlantSuper
            
        Case Is = "Reagent Area Shift Report"
            gGetSrptSupervisor = PlantSuper
            
        Case Is = "Shipping Shift Report"
           gGetSrptSupervisor = PlantSuper
    End Select
    
    If Trim(gGetSrptSupervisor) = "" Then
        gGetSrptSupervisor = "Not entered"
    End If
            
        ClearParams(params)

    Exit Function
    
gGetSupervisorError:
        MsgBox("Error getting special shift report supervisor." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Special Shift Report Supervisor Get Error")
        
    On Error Resume Next
        ClearParams(params)
End Function

Public Function gGetSftyCommCount(ByVal aRptName As String) As Integer

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo gGetSftyCommCountError
      
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim CommCount As Integer

    CommCount = 0
    
    'Get info on special shift report
        'Set 
        params = gDBParams
    
    params.Add("pMineName", gActiveMineNameLong, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
    
    params.Add("pRptTypeName", "Special shift report", ORAPARM_INPUT)
    params("pRptTypeName").serverType = ORATYPE_VARCHAR2
    
    params.Add("pRptName", aRptName, ORAPARM_INPUT)
    params("pRptName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
       
    'PROCEDURE get_srpt_rpts
    'pMineName      IN     VARCHAR2,
    'pRptTypeName   IN     VARCHAR2,
    'pRptName       IN     VARCHAR2,
    'pResult        IN OUT c_reports
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_srpt.get_srpt_rpts(:pMineName," + _
                  ":pRptTypename, :pRptName, :pResult);end;", ORASQL_FAILEXEC)
        'et 
        mRptNameDynaset = params("pResult").Value
        ClearParams(params)
            
    If mRptNameDynaset.RecordCount = 1 Then
        mRptNameDynaset.MoveFirst
           
            If Not IsDBNull(mRptNameDynaset.Fields("sfty_comm1_lbl").Value) Then
                CommCount = 1
            End If
            If Not IsDBNull(mRptNameDynaset.Fields("sfty_comm2_lbl").Value) Then
                CommCount = CommCount + 1
            End If
    End If
    
    gGetSftyCommCount = CommCount
    
    Exit Function
    
gGetSftyCommCountError:
        MsgBox("Error getting special shift report safety comment count" & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Special Shift Reports Get Error")
        
    On Error Resume Next
        ClearParams(params)
End Function

Public Function gSpecRptExists(ByVal aMineName As String, _
                               ByVal aRptName As String, _
                               ByVal aEqptName As String, _
                               ByVal aDate As Date, _
                               ByVal aShift As String, _
                               ByVal aEqptType As String) As Boolean
 
'**********************************************************************
'
'
'
'**********************************************************************
            
    On Error GoTo gSpecRptExistsError

    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim RowCount As Integer
    
        'Set 
        params = gDBParams
    
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pRptTypeName", "Special shift report", ORAPARM_INPUT)
    params("pRptTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pRptName", aRptName, ORAPARM_INPUT)
    params("pRptName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptTypeName", aEqptType, ORAPARM_INPUT)
    params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptName", aEqptName, ORAPARM_INPUT)
    params("pEqptName").serverType = ORATYPE_VARCHAR2

    params.Add("pShiftDate", aDate, ORAPARM_INPUT)
    params("pShiftDate").serverType = ORATYPE_DATE
    
    params.Add("pShift", aShift, ORAPARM_INPUT)
    params("pShift").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_NUMBER

    'PROCEDURE srpt_unit_shift_exists
    'pMineName
    'pRptTypeName
    'pRptName
    'pEqptTypeName
    'pEqptName
    'pShiftDate
    'pShift
    'pResult
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_srpt.srpt_unit_shift_exists(:pMineName," + _
                 ":pRptTypeName, :pRptName, :pEqptTypeName, :pEqptName, :pShiftDate," + _
                 ":pShift, :pResult);end;", ORASQL_FAILEXEC)
    RowCount = params("pResult").Value
        ClearParams(params)
        
    If RowCount <> 0 Then
        gSpecRptExists = True
    Else
        gSpecRptExists = False
    End If
        
    Exit Function
    
gSpecRptExistsError:
        MsgBox("Error determining if report exists." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Special Shift Report Exist Error")
        
    On Error Resume Next
        ClearParams(params)
End Function

    Public Sub ggViewInExcel(ByVal aCommaDelimFile As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        ''''Dim appexcel As Application
        Dim AppExcel As Object

        On Error Resume Next
        'Set 
        AppExcel = GetObject("Excel.Application")

        If AppExcel Is Nothing Then
            'Start AutoCAD if it is not running.
            'Set 
            AppExcel = CreateObject("Excel.Application")

            If AppExcel Is Nothing Then
                MsgBox(Err.Description)
                Exit Sub
            Else
                AppExcel.Visible = True
            End If
        Else
            AppExcel.Visible = True
        End If

        On Error GoTo gViewInExcelError

        AppExcel.Workbooks.Open(aCommaDelimFile, , , 2)
        AppExcel.Visible = True
        AppActivate(AppExcel.Caption)

        Exit Sub

gViewInExcelError:

        MsgBox("Error accessing Excel." + Str(Err.Number) + Chr(10) + Chr(10) + _
            Err.Description, vbExclamation, "Error Accessing Excel")
    End Sub

Public Function gRoundFiftyOld(aNumber As Single) As Long

'**********************************************************************
'
'
'
'**********************************************************************

    Dim IsNegative As Boolean
    Dim TempNumber As Single
    Dim TempNumberInt As Single
    Dim TempNumberFra As Single
    Dim haha As String
    Dim bozo As Long
    
    haha = "AA"
    If aNumber < 0 Then
        IsNegative = True
    Else
        IsNegative = False
    End If
    
    TempNumber = Abs(aNumber)
    TempNumber = Int(TempNumber)
    TempNumberInt = TempNumber / 100 - Int(TempNumber / 100)
    TempNumberFra = TempNumberInt * 100
    
    If TempNumberFra <= 25 Then
        gRoundFiftyOld = TempNumber - TempNumberFra
    End If
    
    If TempNumberFra >= 25 And TempNumberFra <= 50 Then
        gRoundFiftyOld = TempNumber + (50 - TempNumberFra)
    End If
    
    If TempNumberFra >= 50 And TempNumberFra < 75 Then
        gRoundFiftyOld = TempNumber - TempNumberFra + 50
    End If
    
    If TempNumberFra >= 75 And TempNumberFra <= 99 Then
        gRoundFiftyOld = TempNumber + (100 - TempNumberFra)
    End If

    If IsNegative Then
        gRoundFiftyOld = gRoundFiftyOld * -1
    End If
End Function

Public Function gGetMonthDate(ByVal aDate As Date) As String

'**********************************************************************
'
'
'
'**********************************************************************
    
    Dim DateStr As String
    Dim Year As Integer

        DateStr = Format(aDate, "MM/dd/yyyy")
    
    gGetMonthDate = Mid(DateStr, 1, 2) & "/" & Mid(DateStr, 7, 4)
End Function

Public Function gUserExists(ByVal aMineName As String, _
                            ByVal aUserName As String) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo gUserExistsError
    
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim Result As Integer
  
        ' Set 
        params = gDBParams
    
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
   
    params.Add("pUserName", aUserName, ORAPARM_INPUT)
    params("pUserName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_NUMBER
    
    'PROCEDURE user_exists
    'pMineName               IN     VARCHAR2,
    'pUserName               IN     VARCHAR2,
    'pResult                 IN OUT NUMBER)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.user_exists(:pMineName," + _
                  ":pUserName, :pResult);end;", ORASQL_FAILEXEC)
    Result = params("pResult").Value
        ClearParams(params)
    
    If Result = 1 Then
        gUserExists = True
    Else
        gUserExists = False
    End If
    
    Exit Function
    
gUserExistsError:
        MsgBox("Error checking MOIS user status" & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "MOIS User Check Status Error")
        
    On Error Resume Next
        ClearParams(params)
End Function

Public Function gGetSpecialPath(ByVal aMineName As String, _
                                ByVal aPathType As String) As String

'**********************************************************************
'
'
'
'**********************************************************************
           
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim MineInfoDynaset As OraDynaset

    'Get mine information from MINES_MOIS
    
        'Set 
        params = gDBParams
 
    params.Add("pMine", aMineName, ORAPARM_INPUT)
    params("pMine").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_one_mine_info(:pMine, " + _
                 ":pResult);end;", ORASQL_FAILEXEC)
        ' Set 
        MineInfoDynaset = params("pResult").Value
        ClearParams(params)
    
    Select Case StrConv(aPathType, vbUpperCase)
        Case Is = "PROSGRIDPATH"
                If Not IsDBNull(MineInfoDynaset.Fields("prosgrid_path").Value) Then
                    gGetSpecialPath = MineInfoDynaset.Fields("prosgrid_path").Value
                Else
                    gGetSpecialPath = ""
                End If
            
        Case Is = "PLINEDWGPATH"
                If Not IsDBNull(MineInfoDynaset.Fields("plinedwg_path").Value) Then
                    gGetSpecialPath = MineInfoDynaset.Fields("plinedwg_path").Value
                Else
                    gGetSpecialPath = ""
                End If
            
        Case Else
            gGetSpecialPath = ""
    End Select
End Function

Public Function gGetBeginOfCalendarYear(ByVal aDate As Date) As Date

'**********************************************************************
'
'
'
'**********************************************************************

    gGetBeginOfCalendarYear = CDate("01/01/" & DatePart("yyyy", aDate))
End Function

Public Sub gClearGridPrint()

'**********************************************************************
'
'
'
'**********************************************************************

    'Grid to text file transfers related items
    'gPrintGridHeader As String             2
    'gPrintGridSubHeader1 As String         3
    'gPrintGridSubHeader2 As String         4
    'gPrintGridSubHeader3 As String         5
    'gPrintGridDefaultTxtFname As String    6
    'gPrintGridFooter As String             7
    'gOrientHeader As String                8
    'gOrientSubHeader1 As String            9
    'gOrientSubHeader2 As String            10
    'gOrientSubHeader3 As String            11
    'gOrientFooter As String                12
    'gOrientFooter2 As String               13
    'gSubHead2IsHeader As Boolean           14
    'gPrintMarginLeft As Long               15
    'gPrintMarginRight As Long              16
    'gPrintMarginTop As Long                17
    'gPrintMarginBottom As Long             18
    'gPrintOrientation As Integer           19

    gPrintGridHeader = ""
    gPrintGridSubHeader1 = ""
    gPrintGridSubHeader2 = ""
    gPrintGridSubHeader3 = ""
    gPrintGridDefaultTxtFname = ""
    gPrintGridFooter = ""
    gPrintGridFooter2 = ""
    gOrientHeader = ""
    gOrientSubHeader1 = ""
    gOrientSubHeader2 = ""
    gOrientSubHeader3 = ""
    gOrientFooter = ""
    gOrientFooter2 = ""
    gSubHead2IsHeader = False
    gPrintMarginLeft = 0
    gPrintMarginRight = 0
    gPrintMarginTop = 0
    gPrintMarginBottom = 0
    gPrintOrientation = 0
    gAutoPrint = False
End Sub

Public Sub gSetActiveDateAndShift()

'**********************************************************************
'
'
'
'**********************************************************************
     
    Dim ShiftIdx As Integer
    Dim HourOffset As Integer
    Dim NewHour As Integer
    
    'Shift access time will be set to 2 hours plus the defined
    'begin hour for the shift!
    
    'For the two shift mines then the defined start time is 6:00 AM
    'for the Day Shift and 6:00 PM for the Night Shift, thus the
    'day shift access time is 8:00 AM.
    
    ReDim mShiftAccess(gNumShifts)
    HourOffset = 2
    
    'Need the access time begin and end for each shift for this mine
    For ShiftIdx = 1 To gNumShifts
        mShiftAccess(ShiftIdx).ShiftName = gShiftNames(ShiftIdx).ShiftName
        
        'Access time begin hour and minute
        NewHour = gShiftNames(ShiftIdx).BeginHour + HourOffset
        If NewHour > 24 Then
            NewHour = NewHour - 24
        End If
        
        mShiftAccess(ShiftIdx).BeginHour = NewHour
        mShiftAccess(ShiftIdx).BeginMinute = gShiftNames(ShiftIdx).BeginMinute
        
        'Access time end hour and minute
        NewHour = gShiftNames(ShiftIdx).EndHour + HourOffset
        If NewHour > 24 Then
            NewHour = NewHour - 24
        End If
        
        mShiftAccess(ShiftIdx).EndHour = NewHour
        mShiftAccess(ShiftIdx).EndMinute = gShiftNames(ShiftIdx).EndMinute
    Next ShiftIdx
    
    'Now determine which shift we are currently in.
        SetActiveShiftAndDate(mShiftAccess)
End Sub

Public Sub SetActiveShiftAndDate(ByRef aShiftAccess() As ShiftAccessType)

'**********************************************************************
'
'
'
'**********************************************************************

    Dim CurrHr As Integer
    Dim CurrMin As Integer
    Dim RowIdx As Integer
    
    Dim ThisBegHr As Integer
    Dim ThisBegMin As Integer
    Dim ThisEndHr As Integer
    Dim ThisEndMin As Integer
    Dim LastShift As String
    
    'Need to have military times
    CurrHr = DatePart("h", Now)
    CurrMin = DatePart("n", Now)

    LastShift = gGetLastShift(gActiveMineNameLong)
    
    'Process through aShiftAccess to see which shift we are in
        For RowIdx = 1 To UBound(aShiftAccess)
            ThisBegHr = aShiftAccess(RowIdx).BeginHour
            ThisBegMin = aShiftAccess(RowIdx).BeginMinute
            ThisEndHr = aShiftAccess(RowIdx).EndHour
            ThisEndMin = aShiftAccess(RowIdx).EndMinute

            If ThisEndHr > ThisBegHr Then
                If CurrHr >= ThisBegHr And CurrMin >= ThisBegMin And _
                    CurrHr <= ThisEndHr And CurrMin <= ThisEndMin Then

                    If StrConv(aShiftAccess(RowIdx).ShiftName, vbProperCase) = _
                        StrConv(LastShift, vbProperCase) Then
                        gActiveDate = Today.AddDays(-1)
                    Else
                        gActiveDate = Today
                    End If
                    gActiveShift = StrConv(aShiftAccess(RowIdx).ShiftName, vbProperCase)
                    Exit Sub
                End If
            Else
                If CurrHr >= ThisBegHr And CurrMin >= ThisBegMin And _
                    CurrHr <= 24 And CurrMin <= 59 Then
                    gActiveDate = Today
                    gActiveShift = StrConv(aShiftAccess(RowIdx).ShiftName, vbProperCase)
                    Exit Sub
                Else
                    gActiveDate = Today.AddDays(-1)
                    gActiveShift = StrConv(aShiftAccess(RowIdx).ShiftName, vbProperCase)
                    Exit Sub
                End If
            End If

            'If ThisEndHr > ThisBegHr Then
            '    If CurrHr >= ThisBegHr And CurrMin >= ThisBegMin And _
            '        CurrHr <= ThisEndHr And CurrMin <= ThisEndMin Then
            '        gActiveDate = Date
            '        gActiveShift = StrConv(aShiftAccess(RowIdx).ShiftName, vbProperCase)
            '        Exit Sub
            '    End If
            'Else
            '    If CurrHr >= ThisBegHr And CurrMin >= ThisBegMin And _
            '        CurrHr <= 24 And CurrMin <= 59 Then
            '        gActiveDate = Date
            '        gActiveShift = StrConv(aShiftAccess(RowIdx).ShiftName, vbProperCase)
            '        Exit Sub
            '    End If
            '    If CurrHr <= ThisEndHr And CurrMin <= ThisEndMin Then
            '        gActiveDate = Date - 1
            '        gActiveShift = StrConv(aShiftAccess(RowIdx).ShiftName, vbProperCase)
            '        Exit Sub
            '    End If
            'End If
        Next RowIdx
End Sub

Public Sub gSetActiveDateAndShiftOld()

'**********************************************************************
'
'
'
'**********************************************************************
     
    'This is hardcoded to 8:00 AM as a shift access begin time.
    
        If Now.Hour >= 0.83333 Then '8pm to Midnight
            gActiveDate = Today
            gActiveShift = "Night"
        Else
            If Now.Hour <= 0.3333 Then 'Midnight to 8am
                gActiveDate = Today.AddDays(-1)
                gActiveShift = "Night"
            Else
                gActiveDate = Today
                gActiveShift = "Day"
            End If
        End If
End Sub

Public Sub gSetActiveDateAndShiftNew()

'**********************************************************************
'
'
'
'**********************************************************************
        
    Dim CurrHr As Single
    Dim CurrMin As Single
    Dim AmCutOffHr As Single
    Dim AmCutOffMin As Single
    Dim PmCutOffHr As Single
    Dim PmCutOffMin As Single
    
    'Get the hour part of the date -- if 6:35 AM then returns 6.
    '                                 if 6:35 PM then returns 18.
    CurrHr = DatePart("h", Now)
    CurrMin = DatePart("n", Now)
    
    'MOIS has the production date access time in table MINES_MOIS
    'which is available in gProdDateAccessTime.
    'This time is in military time.
    AmCutOffHr = DatePart("h", gProdDateAccessTime)
    AmCutOffMin = DatePart("n", gProdDateAccessTime)
    PmCutOffHr = AmCutOffHr + 12
    PmCutOffMin = AmCutOffMin
    
    'Will make the assumption that the time in gProdDateAccessTime
    'is in the AM.  Therefore PmCutOffHr will be in the PM.
        
    If CurrHr >= PmCutOffHr And CurrMin >= PmCutOffMin Then
        'Example for AmCutOffHr = 08:00 -- 8pm to Midnight
            gActiveDate = Today
        gActiveShift = "Night"
    Else
        If CurrHr <= AmCutOffHr And AmCutOffMin <= AmCutOffMin Then
            'Example for AmCutOffHr = 08:00 -- Midnight to 8am
                gActiveDate = Today.AddDays(-1)
            gActiveShift = "Night"
        Else
            'It must be the day shift!
                gActiveDate = Today
            gActiveShift = "Day"
        End If
    End If
End Sub

Public Function gUserMineSetupCount(ByVal aUserId As String)

'**********************************************************************
'
'
'
'**********************************************************************
        
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim MineSetupCount As Integer

    'Get mine setup count. (How many mines is this user set up for in
    'MOIS?).
    
        'Set 
        params = gDBParams
 
    'PROCEDURE get_user_minesetup_count
    'pUserId                 IN     VARCHAR2,
    'pResult                 IN OUT NUMBER)
    
    params.Add("pUserId", aUserId, ORAPARM_INPUT)
    params("pUserId").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_NUMBER
    
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_user_minesetup_count(:pUserId, " + _
                 ":pResult);end;", ORASQL_FAILEXEC)
    MineSetupCount = params("pResult").Value
        ClearParams(params)

    gUserMineSetupCount = MineSetupCount
End Function

Public Function gUserIsAdministrator(ByVal aMineName As String, _
                                     ByVal aUserName As String) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo UserIsAdministratorError
    
    Dim UserPermsDynaset As OraDynaset
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt

    Dim ThisUserName As String
    
        'Set 
        params = gDBParams
      
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
    
    params.Add("pPermissionTypeName", "Administrator", ORAPARM_INPUT)
    params("pPermissionTypeName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_perm_type_users
    'pMineName           IN     VARCHAR2,
    'pPermissionTypeName IN     VARCHAR2,
    'pResult             IN OUT c_minenames)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_perm_type_users(:pMineName, " & _
                  ":pPermissionTypeName, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        UserPermsDynaset = params("pResult").Value
        ClearParams(params)

    UserPermsDynaset.MoveFirst
            
    gUserIsAdministrator = False
    
    Do While Not UserPermsDynaset.EOF
        ThisUserName = UserPermsDynaset.Fields("user_id").Value
                
        If StrConv(ThisUserName, vbUpperCase) = _
           StrConv(aUserName, vbUpperCase) Then
           gUserIsAdministrator = True
           Exit Do
        End If
        
        UserPermsDynaset.MoveNext
    Loop
  
    UserPermsDynaset.Close
    Exit Function
    
UserIsAdministratorError:
        MsgBox("Error getting users with permission type." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Users With Permission Types Get Error")
        
    On Error Resume Next
        ClearParams(params)
    UserPermsDynaset.Close
End Function

Public Sub gSetGlobalShiftNames(ByVal aMineName As String, _
                                ByVal aDate As Date)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gSetGlobalShiftNamesError

        Dim ShiftNamesDynaset As OraDynaset
        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        Dim ThisShiftName As String
        Dim ThisBeginTime As String
        Dim ThisEndTime As String
        Dim ThisShiftLength As Integer
        Dim ThisShiftOrder As Integer
        Dim ThisSampBeginTime As String
        Dim ThisSampEndTime As String
        Dim ThisBegHr As Integer
        Dim ThisBegMin As Integer
        Dim ThisEndHr As Integer
        Dim ThisEndMin As Integer
        Dim ShiftNameCnt As Integer
        Dim RecordCnt As Integer
        Dim ShiftLength As Integer
        Dim ShiftFound As Boolean
        Dim ShiftIdx As Integer

        'Over time a mine may change the number of shifts that
        'it has per day (typically 2 shifts to 3 shifts or 3 shifts
        'to 2 shifts).  Thus we need to be careful getting the shifts
        'for a mine -- it is date dependent.

        'First need to get the shift length for this date!
        ShiftLength = gGetShiftLengthForDate(aMineName, aDate)

        'Need to assign the following:
        '1) gShiftNames() As String
        '2) gNumShifts As Integer
        '3) gFirstShift As String
        '4) gLastShift as String
        '5) BeginHour As Integer
        '6) BeginMinute As Integer
        '7) EndHour As Integer
        '8) EndMinute as Integer

        'Set 
        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pShiftLengthHrs", ShiftLength, ORAPARM_INPUT)
        params("pShiftLengthHrs").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_shift_names_bylength
        'pMineName           IN     VARCHAR2,
        'pShiftLengthHrs     IN     NUMBER,
        'pResult             IN OUT c_minenames)
        ' Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_shift_names_bylength(:pMineName, " &
                 ":pShiftLengthHrs, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        ShiftNamesDynaset = params("pResult").Value
        ClearParams(params)
        RecordCnt = ShiftNamesDynaset.RecordCount
        ReDim gShiftNames(RecordCnt)

        ShiftNameCnt = 0

        ShiftNamesDynaset.MoveFirst()

        Do While Not ShiftNamesDynaset.EOF
            ThisShiftName = ShiftNamesDynaset.Fields("shift").Value
            ThisBeginTime = ShiftNamesDynaset.Fields("begin_time").Value
            ThisEndTime = ShiftNamesDynaset.Fields("end_time").Value
            ThisShiftLength = ShiftNamesDynaset.Fields("shift_length_hrs").Value
            ThisShiftOrder = 0
            ThisSampBeginTime = ""
            ThisSampEndTime = ""

            ShiftNameCnt = ShiftNameCnt + 1

            If ShiftNameCnt = 1 Then
                gFirstShift = ThisShiftName
            End If

            'Need to get hour and minute for ThisBeginTime
            ThisBegHr = Val(Mid(ThisBeginTime, 1, 2))
            ThisBegMin = Val(Mid(ThisBeginTime, 4))

            'Need to get hour and minute for ThisBeginTime
            ThisEndHr = Val(Mid(ThisEndTime, 1, 2))
            ThisEndMin = Val(Mid(ThisEndTime, 4))

            gShiftNames(ShiftNameCnt).ShiftName = ThisShiftName
            gShiftNames(ShiftNameCnt).BeginTime = ThisBeginTime
            gShiftNames(ShiftNameCnt).EndTime = ThisEndTime
            gShiftNames(ShiftNameCnt).ShiftLength = ThisShiftLength
            gShiftNames(ShiftNameCnt).ShiftOrder = ThisShiftOrder
            gShiftNames(ShiftNameCnt).SampBeginTime = ThisSampBeginTime
            gShiftNames(ShiftNameCnt).SampEndTime = ThisSampEndTime

            gShiftNames(ShiftNameCnt).BeginHour = ThisBegHr
            gShiftNames(ShiftNameCnt).BeginMinute = ThisBegMin

            gShiftNames(ShiftNameCnt).EndHour = ThisEndHr
            gShiftNames(ShiftNameCnt).EndMinute = ThisEndMin

            ShiftNamesDynaset.MoveNext()
        Loop

        gLastShift = ThisShiftName
        gNumShifts = ShiftNameCnt

        'Assume that all shifts have the same length at any given mine (for
        'any given date)!
        gShiftLength = gShiftNames(1).ShiftLength

        'We may have created a problem with gActiveShift -- it may no longer
        'be an option in gShiftNames!
        ShiftFound = False
        For ShiftIdx = 1 To UBound(gShiftNames)
            If StrConv(gShiftNames(ShiftIdx).ShiftName, vbUpperCase) =
          StrConv(gActiveShift, vbUpperCase) Then
                ShiftFound = True
            End If
        Next ShiftIdx

        If ShiftFound = False Then
            'Set gActiveShift to the first shift option
            gActiveShift = gShiftNames(1).ShiftName
        End If

        ShiftNamesDynaset.Close()
        Exit Sub

gSetGlobalShiftNamesError:
        MsgBox("Error setting global shift names." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Shift Names Set Error")
        
    On Error Resume Next
        ClearParams(params)
    ShiftNamesDynaset.Close
End Sub

Public Function gGetPrevShift(ByVal aDate As Date, _
                              ByVal aShift As String) As gShiftInfoType

'**********************************************************************
'This function returns the previous shift from the date and shift
'that it receives as input parameters.
'
'**********************************************************************
    
    Dim RowIdx As Integer
    Dim ShiftInfo As gShiftInfoType
    
    'Find the shift we are in (gShiftNames)
    'Assume that the shifts in gShiftInfo are in the correct order
    'timewise.
    
    For RowIdx = 1 To UBound(gShiftNames)
        If StrConv(gShiftNames(RowIdx).ShiftName, vbUpperCase) = _
            StrConv(aShift, vbUpperCase) Then
            'We have found the shift, we want the shift that is previous
            'to this one
            
            If RowIdx <> 1 Then
                'Just need to go back one shift for the current date.
                    ShiftInfo.dDate = aDate
                ShiftInfo.Shift = gShiftNames(RowIdx - 1).ShiftName
            Else
                'Need the last shift from the previous day.
                    ShiftInfo.dDate = aDate.AddDays(-1)
                ShiftInfo.Shift = gShiftNames(UBound(gShiftNames)).ShiftName
            End If
        End If
    Next RowIdx
    
    gGetPrevShift = ShiftInfo
End Function

Public Function gGetNextShift(ByVal aDate As Date, _
                              ByVal aShift As String) As gShiftInfoType

'**********************************************************************
'This function returns the next shift from the date and shift
'that it receives as input parameters.
'
'**********************************************************************
    
    Dim RowIdx As Integer
    Dim ShiftInfo As gShiftInfoType
    
    'Find the shift we are in (gShiftNames)
    'Assume that the shifts in gShiftInfo are in the correct order
    'timewise.
    
    For RowIdx = 1 To UBound(gShiftNames)
        If StrConv(gShiftNames(RowIdx).ShiftName, vbUpperCase) = _
            StrConv(aShift, vbUpperCase) Then
            'We have found the shift, we want the shift that is next
            'to this one
            
            If RowIdx <> UBound(gShiftNames) Then
                'Just need to go forward one shift for the current date.
                    ShiftInfo.dDate = aDate
                ShiftInfo.Shift = gShiftNames(RowIdx + 1).ShiftName
            Else
                'Need the first shift from the next day.
                    ShiftInfo.dDate = aDate.AddDays(1)
                ShiftInfo.Shift = gShiftNames(1).ShiftName
            End If
        End If
    Next RowIdx
    
    gGetNextShift = ShiftInfo
End Function

Public Function gIsLegalShift(ByVal aShiftName As String) As Boolean

'**********************************************************************
'This function determines whether a shift name is legal.
'
'
'**********************************************************************

    Dim RowIdx As Integer
    
    gIsLegalShift = False
    
    For RowIdx = 1 To UBound(gShiftNames)
        If StrConv(aShiftName, vbUpperCase) = _
            StrConv(gShiftNames(RowIdx).ShiftName, vbUpperCase) Then
            gIsLegalShift = True
            Exit For
        End If
    Next RowIdx
End Function

Public Sub gGetUserInfo(aUserId As String)

'**********************************************************************
'
'
'
'**********************************************************************
    
    On Error GoTo gGetUserInfoError
    
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim MoisPwordDynaset As OraDynaset
    
    Dim RecordCount As Long
             
        'Set 
        params = gDBParams
    
    params.Add("pUserId", aUserId, ORAPARM_INPUT)
    params("pUserId").serverType = ORATYPE_VARCHAR2
         
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
        
    'Procedure get_mois_pwords
    'pUserId          IN     VARCHAR2,
    'pResult          IN OUT c_pwords)
    
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_pwords.get_mois_pwords(:pUserId," + _
                  ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MoisPwordDynaset = params("pResult").Value
        ClearParams(params)
    
    'RecordCount should be one!
    RecordCount = MoisPwordDynaset.RecordCount
    
    'Now add data to input fields
        If Not IsDBNull(MoisPwordDynaset.Fields("user_name").Value) Then
            gUserInfo.UserName = MoisPwordDynaset.Fields("user_name").Value
        Else
            gUserInfo.UserName = ""
        End If
    '----------
        If Not IsDBNull(MoisPwordDynaset.Fields("mail_name").Value) Then
            gUserInfo.MailName = MoisPwordDynaset.Fields("mail_name").Value
        Else
            gUserInfo.MailName = ""
        End If
    '----------
        If Not IsDBNull(MoisPwordDynaset.Fields("mail_server").Value) Then
            gUserInfo.MailServer = MoisPwordDynaset.Fields("mail_server").Value
        Else
            gUserInfo.MailServer = ""
        End If
    '----------
        If Not IsDBNull(MoisPwordDynaset.Fields("first_name").Value) Then
            gUserInfo.FirstName = MoisPwordDynaset.Fields("first_name").Value
        Else
            gUserInfo.FirstName = ""
        End If
    '----------
        If Not IsDBNull(MoisPwordDynaset.Fields("middle_init").Value) Then
            gUserInfo.MiddleInit = MoisPwordDynaset.Fields("middle_init").Value
        Else
            gUserInfo.MiddleInit = ""
        End If
    '----------
        If Not IsDBNull(MoisPwordDynaset.Fields("last_name").Value) Then
            gUserInfo.LastName = MoisPwordDynaset.Fields("last_name").Value
        Else
            gUserInfo.LastName = ""
        End If
    '----------
        If Not IsDBNull(MoisPwordDynaset.Fields("user_loctn").Value) Then
            gUserInfo.UserLoctn = MoisPwordDynaset.Fields("user_loctn").Value
        Else
            gUserInfo.UserLoctn = ""
        End If
    '----------
        If Not IsDBNull(MoisPwordDynaset.Fields("default_mine").Value) Then
            gUserInfo.DefaultMine = MoisPwordDynaset.Fields("default_mine").Value
        Else
            gUserInfo.DefaultMine = ""
        End If
    '----------
    
    MoisPwordDynaset.Close
    Exit Sub
    
gGetUserInfoError:
        MsgBox("Error getting MOIS user info." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "MOIS User Info Get Error")
        
    On Error Resume Next
        ClearParams(params)
    MoisPwordDynaset.Close
End Sub

Public Function gGetMilitaryTime(aAmPmTime As String) As String

'**********************************************************************
'Expects AM PM time string like 08:05 AM, 12:30 PM, etc.
'
'
'**********************************************************************
                 
    '01:00 AM   0100            01:00 PM   1300
    '02:00 AM   0200            02:00 PM   1400
    '03:00 AM   0300            03:00 PM   1500
    '04:00 AM   0400            04:00 PM   1600
    '05:00 AM   0500            05:00 PM   1700
    '06:00 AM   0600            06:00 PM   1800
    '07:00 AM   0700            07:00 PM   1900
    '08:00 AM   0800            08:00 PM   2000
    '09:00 AM   0900            09:00 PM   2100
    '10:00 AM   1000            10:00 PM   2200
    '11:00 AM   1100            11:00 PM   2300
    '12:00 AM   1200            12:00 PM   2400
    
    Dim DateStr As String
    Dim DateTemp As Date
    Dim ThisHr As Integer
    Dim ThisMin As Integer
     
        DateStr = CStr(Today) & " " & aAmPmTime
    
    If IsDate(DateStr) = True Then
        DateTemp = CDate(DateStr)
        ThisHr = DatePart("h", DateTemp)    'Returns military time
        ThisMin = DatePart("n", DateTemp)
        gGetMilitaryTime = Format(ThisHr, "0#") & ":" & Format(ThisMin, "0#")
    Else
        gGetMilitaryTime = "??:??"
    End If
End Function

Public Function gGetMilitaryTimeSec(aAmPmTime As String) As String

'**********************************************************************
'Expects AM PM time string like 08:05:23 AM, 12:30:05 PM, etc.
'
'
'**********************************************************************
                 
    '01:00 AM   0100            01:00 PM   1300
    '02:00 AM   0200            02:00 PM   1400
    '03:00 AM   0300            03:00 PM   1500
    '04:00 AM   0400            04:00 PM   1600
    '05:00 AM   0500            05:00 PM   1700
    '06:00 AM   0600            06:00 PM   1800
    '07:00 AM   0700            07:00 PM   1900
    '08:00 AM   0800            08:00 PM   2000
    '09:00 AM   0900            09:00 PM   2100
    '10:00 AM   1000            10:00 PM   2200
    '11:00 AM   1100            11:00 PM   2300
    '12:00 AM   1200            12:00 PM   2400
    
    Dim DateStr As String
    Dim DateTemp As Date
    Dim ThisHr As Integer
    Dim ThisMin As Integer
    Dim ThisSec As Integer
    
        DateStr = CStr(Today) & " " & aAmPmTime
    
    If IsDate(DateStr) = True Then
        DateTemp = CDate(DateStr)
        ThisHr = DatePart("h", DateTemp)    'Returns military time
        ThisMin = DatePart("n", DateTemp)
        ThisSec = DatePart("s", DateTemp)
        
        gGetMilitaryTimeSec = Format(ThisHr, "0#") & ":" & _
                           Format(ThisMin, "0#") & ":" & _
                           Format(ThisSec, "0#")
    Else
        gGetMilitaryTimeSec = "??:??:??"
    End If
End Function

Public Function gGetMoisBeginDate(ByVal aMineName As String) As Date

'**********************************************************************
'
'
'
'**********************************************************************
           
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim MoisBeginDate As Date

    'Get MOIS begin date from MINES_MOIS
        'Set 
        params = gDBParams
 
    params.Add("pMine", aMineName, ORAPARM_INPUT)
    params("pMine").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_DATE
    
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_mois_begin_date(:pMine, " + _
                 ":pResult);end;", ORASQL_FAILEXEC)
    MoisBeginDate = params("pResult").Value
        ClearParams(params)
    
    gGetMoisBeginDate = MoisBeginDate
End Function

Public Function gGetMonthsInRange(ByVal aBeginDate As Date, _
                                  ByVal aEndDate As Date) As Integer

'**********************************************************************
'
'
'
'**********************************************************************
           
    Dim MonthCount As Integer
    Dim CurrMo As Integer
    Dim CurrYr As Integer
    Dim ThisMo As Integer
    Dim ThisYr As Integer
    Dim ThisDate As Date
      
    CurrMo = 0
    CurrYr = 0
    MonthCount = 1
    gGetMonthsInRange = 0
    
        'For ThisDate = aBeginDate To aEndDate
        '    ThisMo = DatePart("m", ThisDate)
        '    ThisYr = DatePart("yyyy", ThisDate)

        '    If (ThisMo <> CurrMo Or ThisYr <> CurrYr) And CurrMo <> 0 Then
        '        'New month has been encountered
        '        MonthCount = MonthCount + 1
        '    End If

        '    CurrMo = ThisMo
        '    CurrYr = ThisYr
        'Next ThisDate
        MonthCount = Abs(DateDiff(DateInterval.Month, aBeginDate, aEndDate))
    gGetMonthsInRange = MonthCount
End Function

Public Function gGetWeeksInRange(ByVal aBeginDate As Date, _
                                 ByVal aEndDate As Date) As Integer

'**********************************************************************
'
'
'
'**********************************************************************
           
    Dim WeekCount As Integer
    Dim CurrWk As Integer
    Dim ThisWk As Integer
    Dim ThisDate As Date
      
    CurrWk = 0
    WeekCount = 1
    gGetWeeksInRange = 0
    
    'Assumes first day of week is Monday.
    
        'For ThisDate = aBeginDate To aEndDate
        '    ThisWk = DatePart("ww", ThisDate, vbMonday)

        '    If ThisWk <> CurrWk And CurrWk <> 0 Then
        '        'New week has been encountered
        '        WeekCount = WeekCount + 1
        '    End If

        '    CurrWk = ThisWk
        'Next ThisDate
        WeekCount = Abs(DateDiff(DateInterval.Day, aBeginDate, aEndDate) / 7)

    gGetWeeksInRange = WeekCount
End Function

Public Function gGetDateTime(ByVal aDate As Date, _
                             ByVal aTime As Date) As Date
'**********************************************************************
'
'
'
'**********************************************************************

        gGetDateTime = CDate(Format(aDate, "MM/dd/yyyy") & " " & _
                        Format(aTime, "hh:mm AM/PM"))
End Function

Public Function gGetFirstShift(ByVal aMineName As String) As String

'**********************************************************************
' This function has been replaced with gGetFirstShift2.  The first
' shift is date dependent -- gGetFirstShift2 has aDate as an
' additional parameter 06/28/2006, lss.  This function should not
' typically be used anymore (be careful if you use it)!
'**********************************************************************
           
    On Error GoTo gGetFirstShiftError
    
    Dim ShiftNamesDynaset As OraDynaset
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt

    Dim ThisShiftName As String
     
    gGetFirstShift = ""
    
        'Set 
        params = gDBParams
      
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_shift_names
    'pMineName           IN     VARCHAR2,
    'pResult             IN OUT c_minenames)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_shift_names(:pMineName, " & _
                  ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        ShiftNamesDynaset = params("pResult").Value
        ClearParams(params)
   
    ShiftNamesDynaset.MoveFirst
    gGetFirstShift = ShiftNamesDynaset.Fields("shift").Value
    
    ShiftNamesDynaset.Close
    Exit Function
    
gGetFirstShiftError:
        MsgBox("Error getting first shift name." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "First Shift Name Get Error")
        
    On Error Resume Next
    gGetFirstShift = ""
    On Error Resume Next
        ClearParams(params)
    On Error Resume Next
    ShiftNamesDynaset.Close
End Function

Public Function gGetFirstShift2(ByVal aMineName As String, _
                                ByVal aDate As Date) As String

'**********************************************************************
' This function replaced gGetFirstShift.  When FCO changed from 3 to
' 2 shifts (August, 2006), the first shift value became date
' dependent.  All uses of gGetFirstShift in MOIS were then replaced
' with gGetFirstShift2 06/28/2006, lss.
'**********************************************************************
           
    On Error GoTo gGetFirstShift2Error
    
    Dim ShiftNamesDynaset As OraDynaset
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim ShiftLengthForDate As Integer
    
    gGetFirstShift2 = ""
    ShiftLengthForDate = gGetShiftLengthForDate(aMineName, aDate)
    
        ' Set 
        params = gDBParams
      
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
    
    params.Add("pShiftLengthHrs", ShiftLengthForDate, ORAPARM_INPUT)
    params("pshiftLengthHrs").serverType = ORATYPE_NUMBER
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_shift_names_bylength
    'pMineName           IN     VARCHAR2,
    'pShiftLengthHrs     IN     NUMBER,
    'pResult             IN OUT c_shiftnames)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_shift_names_bylength(:pMineName, " & _
                  ":pShiftLengthHrs, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        ShiftNamesDynaset = params("pResult").Value
        ClearParams(params)
   
    ShiftNamesDynaset.MoveFirst
    gGetFirstShift2 = ShiftNamesDynaset.Fields("shift").Value
    
    ShiftNamesDynaset.Close
    Exit Function
    
gGetFirstShift2Error:
        MsgBox("Error getting first shift name." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "First Shift Name Get Error")
        
    On Error Resume Next
    gGetFirstShift2 = ""
    On Error Resume Next
        ClearParams(params)
    On Error Resume Next
    ShiftNamesDynaset.Close
End Function

Public Function gGetLastShift(ByVal aMineName As String) As String

'**********************************************************************
' This function has been replaced with gGetLastShift2.  The last
' shift is date dependent -- gGetLastShift2 has aDate as an
' additional parameter 06/28/2006, lss.  This function should not
' typically be used anymore (be careful if you use it)!
'**********************************************************************
           
    On Error GoTo gGetLastShiftError
    
    Dim ShiftNamesDynaset As OraDynaset
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
     
    gGetLastShift = ""
    
        'Set 
        params = gDBParams
      
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_shift_names
    'pMineName           IN     VARCHAR2,
    'pResult             IN OUT c_minenames)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_shift_names(:pMineName, " & _
                  ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        ShiftNamesDynaset = params("pResult").Value
        ClearParams(params)
   
    ShiftNamesDynaset.MoveLast
    gGetLastShift = ShiftNamesDynaset.Fields("shift").Value
    
    ShiftNamesDynaset.Close
    Exit Function
    
gGetLastShiftError:
        MsgBox("Error getting last shift name." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Last Shift Name Get Error")
        
    On Error Resume Next
    gGetLastShift = ""
    On Error Resume Next
        ClearParams(params)
    On Error Resume Next
    ShiftNamesDynaset.Close
End Function

Public Function gGetLastShift2(ByVal aMineName As String, _
                               ByVal aDate As Date) As String

'**********************************************************************
' This function replaced gGetLastShift.  When FCO changed from 3 to
' 2 shifts (August, 2006), the last shift value became date
' dependent.  All uses of gGetLastShift in MOIS were then replaced
' with gGetLastShift2 06/28/2006, lss.
'**********************************************************************
           
    On Error GoTo gGetLastShift2Error
    
    Dim ShiftNamesDynaset As OraDynaset
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim ShiftLengthForDate As Integer
    
    gGetLastShift2 = ""
    ShiftLengthForDate = gGetShiftLengthForDate(aMineName, aDate)
    
        'Set 
        params = gDBParams
      
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
    
    params.Add("pShiftLengthHrs", ShiftLengthForDate, ORAPARM_INPUT)
    params("pShiftLengthHrs").serverType = ORATYPE_NUMBER
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_shift_names_bylength
    'pMineName           IN     VARCHAR2,
    'pShiftLengthHrs     IN     NUMBER,
    'pResult             IN OUT c_shiftnames)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_shift_names_bylength(:pMineName, " & _
                  ":pShiftLengthHrs, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        ShiftNamesDynaset = params("pResult").Value
        ClearParams(params)
   
    ShiftNamesDynaset.MoveLast
    gGetLastShift2 = ShiftNamesDynaset.Fields("shift").Value
    
    ShiftNamesDynaset.Close
    Exit Function
    
gGetLastShift2Error:
        MsgBox("Error getting last shift name." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Last Shift Name Get Error")
        
    On Error Resume Next
    gGetLastShift2 = ""
    On Error Resume Next
        ClearParams(params)
    On Error Resume Next
    ShiftNamesDynaset.Close
End Function

Public Function gGetNumShifts(ByVal aMineName As String) As Integer

'**********************************************************************
'
'
'
'**********************************************************************
           
    On Error GoTo gGetNumShiftsError
    
    Dim ShiftNamesDynaset As OraDynaset
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt

    Dim ThisShiftName As String
     
    gGetNumShifts = 0
    
        ' Set 
        params = gDBParams
      
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_shift_names
    'pMineName           IN     VARCHAR2,
    'pResult             IN OUT c_minenames)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_shift_names(:pMineName, " & _
                  ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        ShiftNamesDynaset = params("pResult").Value
        ClearParams(params)
 
    gGetNumShifts = ShiftNamesDynaset.RecordCount
    
    ShiftNamesDynaset.Close
    Exit Function
    
gGetNumShiftsError:
        MsgBox("Error getting number of shifts." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Shift Number Error")
        
    On Error Resume Next
    gGetNumShifts = 0
        ClearParams(params)
    ShiftNamesDynaset.Close
End Function

Public Function gGetMiscMineGlobals(ByVal aMineName As String, _
                                    ByVal aGlobal As String) As String

'**********************************************************************
'
'
'
'**********************************************************************
    
    'Global that this function can return:
    
    ' 1) "company name"
    ' 2) "prod date access time"
    ' 3) "prod date create time"
    ' 4) "samp day shift begin time"
    ' 5) "samp day shift end time"
    ' 6) "szdfd tpoh aswhole"
    ' 7) "washer poper mode"
    ' 8) "sizing poper mode"
    ' 9) "float_plant_poper_mode"
    '10) "logon mine"
    '11) "pipe tracking"
    '12) "mine prospect"
    '13) "has dredges"
    '14) "reagent by day"
    '15) "mois begin date"
    '16) "num washer sides"
    '17) "prosp pb is pbip"
    '18) "has catalog reserves"
    '19) "prod tons round"
    '20) "hard freeze date"
    '21) "soft freeze date"
    '22) "has ctlg reserves"
    '23) "catalog prosp desc"
    '24) "100% prosp desc"
    '25) "shift change date"
    '26) "mass balance mode"
    '27) "has draglines"
                
    On Error GoTo SetMiscMineGlobalsError
    
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim MineInfoDynaset As OraDynaset

    'Get global mine information from MINES_MOIS
        gGetMiscMineGlobals = String.Empty
        'Set 
        params = gDBParams
 
    params.Add("pMine", aMineName, ORAPARM_INPUT)
    params("pMine").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
        'et 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_one_mine_info(:pMine, " + _
                 ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MineInfoDynaset = params("pResult").Value
        ClearParams(params)
    
    'Should only be one row returned.
    MineInfoDynaset.MoveFirst
    
    Select Case aGlobal
        Case Is = "company name"
                If Not IsDBNull(MineInfoDynaset.Fields("company_name").Value) Then
                    gGetMiscMineGlobals = MineInfoDynaset.Fields("company_name").Value
                Else
                    gGetMiscMineGlobals = "Unknown Company Name!!!"
                End If
    
        Case Is = "prod date access time"
                If Not IsDBNull(MineInfoDynaset.Fields("prod_date_access_time").Value) Then
                    gGetMiscMineGlobals = MineInfoDynaset.Fields("prod_date_access_time").Value
                Else
                    gGetMiscMineGlobals = ""
                End If
    
        Case Is = "prod date create time"
                If Not IsDBNull(MineInfoDynaset.Fields("prod_date_create_time").Value) Then
                    gGetMiscMineGlobals = MineInfoDynaset.Fields("prod_date_create_time").Value
                Else
                    gGetMiscMineGlobals = ""
                End If
    
        Case Is = "samp day shift begin time"
                If Not IsDBNull(MineInfoDynaset.Fields("samp_day_shift_begin_time").Value) Then
                    gGetMiscMineGlobals = MineInfoDynaset.Fields("samp_day_shift_begin_time").Value
                Else
                    gGetMiscMineGlobals = ""
                End If
    
        Case Is = "samp day shift end time"
                If Not IsDBNull(MineInfoDynaset.Fields("samp_day_shift_end_time").Value) Then
                    gGetMiscMineGlobals = MineInfoDynaset.Fields("samp_day_shift_end_time").Value
                Else
                    gGetMiscMineGlobals = ""
                End If
    
        Case Is = "szdfd tpoh aswhole"
                If Not IsDBNull(MineInfoDynaset.Fields("szdfd_tpoh_aswhole").Value) Then
                    If MineInfoDynaset.Fields("szdfd_tpoh_aswhole").Value = 1 Then
                        gGetMiscMineGlobals = "Yes"
                    Else
                        gGetMiscMineGlobals = "No"
                    End If
                Else
                    gGetMiscMineGlobals = "No"
                End If
    
        Case Is = "washer poper mode"
                If Not IsDBNull(MineInfoDynaset.Fields("washer_poper_mode").Value) Then
                    gGetMiscMineGlobals = MineInfoDynaset.Fields("washer_poper_mode").Value
                Else
                    gGetMiscMineGlobals = ""
                End If
    
        Case Is = "sizing poper mode"
                If Not IsDBNull(MineInfoDynaset.Fields("sizing_poper_mode").Value) Then
                    gGetMiscMineGlobals = MineInfoDynaset.Fields("sizing_poper_mode").Value
                Else
                    gGetMiscMineGlobals = ""
                End If
    
        Case Is = "float_plant_poper_mode"
                If Not IsDBNull(MineInfoDynaset.Fields("float_plant_poper_mode").Value) Then
                    gGetMiscMineGlobals = MineInfoDynaset.Fields("float_plant_poper_mode").Value
                Else
                    gGetMiscMineGlobals = ""
                End If
    
        Case Is = "logon mine"
                If Not IsDBNull(MineInfoDynaset.Fields("logon_mine").Value) Then
                    If MineInfoDynaset.Fields("logon_mine").Value = 1 Then
                        gGetMiscMineGlobals = "Yes"
                    Else
                        gGetMiscMineGlobals = "No"
                    End If
                Else
                    gGetMiscMineGlobals = "No"
                End If
    
        Case Is = "pipe tracking"
                If Not IsDBNull(MineInfoDynaset.Fields("pipe_tracking").Value) Then
                    If MineInfoDynaset.Fields("pipe_tracking").Value = 1 Then
                        gGetMiscMineGlobals = "Yes"
                    Else
                        gGetMiscMineGlobals = "No"
                    End If
                Else
                    gGetMiscMineGlobals = "No"
                End If
    
        Case Is = "mine prospect"
                If Not IsDBNull(MineInfoDynaset.Fields("mine_prospect").Value) Then
                    If MineInfoDynaset.Fields("mine_prospect").Value = 1 Then
                        gGetMiscMineGlobals = "Yes"
                    Else
                        gGetMiscMineGlobals = "No"
                    End If
                Else
                    gGetMiscMineGlobals = "No"
                End If
    
        Case Is = "has dredges"
                If Not IsDBNull(MineInfoDynaset.Fields("has_dredges").Value) Then
                    If MineInfoDynaset.Fields("has_dredges").Value = 1 Then
                        gGetMiscMineGlobals = "Yes"
                    Else
                        gGetMiscMineGlobals = "No"
                    End If
                Else
                    gGetMiscMineGlobals = "No"
                End If
    
        Case Is = "reagent by day"
                If Not IsDBNull(MineInfoDynaset.Fields("reagent_by_day").Value) Then
                    If MineInfoDynaset.Fields("reagent_by_day").Value = 1 Then
                        gGetMiscMineGlobals = "Yes"
                    Else
                        gGetMiscMineGlobals = "No"
                    End If
                Else
                    gGetMiscMineGlobals = "No"
                End If
    
        Case Is = "mois begin date"
                If Not IsDBNull(MineInfoDynaset.Fields("mois_begin_date").Value) Then
                    gGetMiscMineGlobals = CStr(MineInfoDynaset.Fields("mois_begin_date").Value)
                Else
                    gGetMiscMineGlobals = "12/31/8888"
                End If
    
        Case Is = "num washer sides"
                If Not IsDBNull(MineInfoDynaset.Fields("num_washer_sides").Value) Then
                    gGetMiscMineGlobals = CStr(MineInfoDynaset.Fields("num_washer_sides").Value)
                Else
                    gGetMiscMineGlobals = "0"
                End If
            
        Case Is = "prosp pb is pbip"
                If Not IsDBNull(MineInfoDynaset.Fields("prosp_pb_is_pbip").Value) Then
                    If MineInfoDynaset.Fields("prosp_pb_is_pbip").Value = 1 Then
                        gGetMiscMineGlobals = "Yes"
                    Else
                        gGetMiscMineGlobals = "No"
                    End If
                Else
                    gGetMiscMineGlobals = "No"
                End If
            
        Case Is = "has catalog reserves"
                If Not IsDBNull(MineInfoDynaset.Fields("has_ctlg_reserves").Value) Then
                    If MineInfoDynaset.Fields("has_ctlg_reserves").Value = 1 Then
                        gGetMiscMineGlobals = "Yes"
                    Else
                        gGetMiscMineGlobals = "No"
                    End If
                Else
                    gGetMiscMineGlobals = "No"
                End If
           
        Case Is = "prod tons round"
                If Not IsDBNull(MineInfoDynaset.Fields("prod_tons_round").Value) Then
                    gGetMiscMineGlobals = CStr(MineInfoDynaset.Fields("prod_tons_round").Value)
                Else
                    gGetMiscMineGlobals = "50"
                End If
            
        Case Is = "hard freeze date"
                If Not IsDBNull(MineInfoDynaset.Fields("hard_freeze_date").Value) Then
                    gGetMiscMineGlobals = Format(MineInfoDynaset.Fields("hard_freeze_date").Value, "MM/dd/yyyy")
                Else
                    gGetMiscMineGlobals = ""
                End If
            
        Case Is = "soft freeze date"
                If Not IsDBNull(MineInfoDynaset.Fields("soft_freeze_date").Value) Then
                    gGetMiscMineGlobals = Format(MineInfoDynaset.Fields("soft_freeze_date").Value, "MM/dd/yyyy")
                Else
                    gGetMiscMineGlobals = ""
                End If
            
        Case Is = "has ctlg reserves"
                If Not IsDBNull(MineInfoDynaset.Fields("has_ctlg_reserves").Value) Then
                    If MineInfoDynaset.Fields("has_ctlg_reserves").Value = 1 Then
                        gGetMiscMineGlobals = "Yes"
                    Else
                        gGetMiscMineGlobals = "No"
                    End If
                Else
                    gGetMiscMineGlobals = "No"
                End If
          
        Case Is = "catalog prosp desc"
                If Not IsDBNull(MineInfoDynaset.Fields("pctcatalog_prosp_desc").Value) Then
                    gGetMiscMineGlobals = MineInfoDynaset.Fields("pctcatalog_prosp_desc").Value
                Else
                    gGetMiscMineGlobals = ""
                End If
            
        Case Is = "100% prosp desc"
                If Not IsDBNull(MineInfoDynaset.Fields("pct100_prosp_desc").Value) Then
                    gGetMiscMineGlobals = MineInfoDynaset.Fields("pct100_prosp_desc").Value
                Else
                    gGetMiscMineGlobals = ""
                End If
            
        Case Is = "shift change date"
                If Not IsDBNull(MineInfoDynaset.Fields("mois_begin_date").Value) Then
                    gGetMiscMineGlobals = Format(MineInfoDynaset.Fields("mois_begin_date").Value, "MM/dd/yyyy")
                Else
                    gGetMiscMineGlobals = "12/31/8888"
                End If
            
        Case Is = "mass balance mode"
                If Not IsDBNull(MineInfoDynaset.Fields("massbalance_mode").Value) Then
                    gGetMiscMineGlobals = Format(MineInfoDynaset.Fields("massbalance_mode").Value, "MM/dd/yyyy")
                Else
                    gGetMiscMineGlobals = ""
                End If
            
        Case Is = "has draglines"
                If Not IsDBNull(MineInfoDynaset.Fields("has_draglines").Value) Then
                    If MineInfoDynaset.Fields("has_draglines").Value = 1 Then
                        gGetMiscMineGlobals = "Yes"
                    Else
                        gGetMiscMineGlobals = "No"
                    End If
                Else
                    gGetMiscMineGlobals = "No"
                End If
    End Select
    
    MineInfoDynaset.Close
    
    Exit Function
    
SetMiscMineGlobalsError:
        MsgBox("Error accessing mines info." & vbCrLf & _
               Err.Description, vbOKOnly + vbExclamation, _
               "Mine Globals Access Error")
    On Error Resume Next
        ClearParams(params)
    MineInfoDynaset.Close
End Function

Public Function gQtrBeginDateCurr(ByVal aDate As Date) As Date

'**********************************************************************
'Returns the first day in the quarter that aDate is in.
'
'
'**********************************************************************

    Dim Fybd As Date
    
    Dim QtrBeg1 As Date
    Dim QtrBeg2 As Date
    Dim QtrBeg3 As Date
    Dim QtrBeg4 As Date
    
    Fybd = gGetBeginOfFiscalYear(aDate)
    
    QtrBeg1 = gQtrBeginDate(Fybd, 1)
    QtrBeg2 = gQtrBeginDate(Fybd, 2)
    QtrBeg3 = gQtrBeginDate(Fybd, 3)
    QtrBeg4 = gQtrBeginDate(Fybd, 4)
    
    If aDate >= QtrBeg1 And aDate < QtrBeg2 Then
        gQtrBeginDateCurr = QtrBeg1
        Exit Function
    End If
    If aDate >= QtrBeg2 And aDate < QtrBeg3 Then
        gQtrBeginDateCurr = QtrBeg2
        Exit Function
    End If
    If aDate >= QtrBeg3 And aDate < QtrBeg4 Then
        gQtrBeginDateCurr = QtrBeg3
        Exit Function
    End If
    
    gQtrBeginDateCurr = QtrBeg4
End Function

Public Function gQtrEndDateCurr(ByVal aDate As Date) As Date

'**********************************************************************
'Returns the last day in the quarter that aDate is in.
'
'
'**********************************************************************

    Dim Fybd As Date
    
    Dim QtrBeg1 As Date
    Dim QtrBeg2 As Date
    Dim QtrBeg3 As Date
    Dim QtrBeg4 As Date
    
    Fybd = gGetBeginOfFiscalYear(aDate)
    
    QtrBeg1 = gQtrBeginDate(Fybd, 1)
    QtrBeg2 = gQtrBeginDate(Fybd, 2)
    QtrBeg3 = gQtrBeginDate(Fybd, 3)
    QtrBeg4 = gQtrBeginDate(Fybd, 4)
    
    If aDate >= QtrBeg1 And aDate < QtrBeg2 Then
            gQtrEndDateCurr = QtrBeg2.AddDays(-1)
        Exit Function
    End If
    If aDate >= QtrBeg2 And aDate < QtrBeg3 Then
            gQtrEndDateCurr = QtrBeg3.AddDays(-1)
        Exit Function
    End If
    If aDate >= QtrBeg3 And aDate < QtrBeg4 Then
            gQtrEndDateCurr = QtrBeg4.AddDays(-1)
        Exit Function
    End If
    
        gQtrEndDateCurr = QtrBeg1.AddDays(-1)
End Function

Public Function gQtrBeginDate(ByVal aDate As Date, _
                              ByVal aQuarter As Integer) As Date

'**********************************************************************
'Determines the quarter begin date for the fiscal year (determined by
'aDate) for the quarter requested (aQuarter = 1, 2, 3, 4).
'
'
'**********************************************************************
    
    Dim Fybd As Date
    Dim QtrCount As Integer
    Dim ThisMo As Integer
    Dim ThisYr As Integer
    Dim DateStr As String
    
    Fybd = gGetBeginOfFiscalYear(aDate)
    
    ThisMo = DatePart("m", Fybd)
    ThisYr = DatePart("yyyy", Fybd)
    
    If aQuarter = 1 Then
        gQtrBeginDate = Fybd
        Exit Function
    End If
    
    For QtrCount = 1 To aQuarter - 1
        ThisMo = ThisMo + 3
        If ThisMo > 12 Then
            ThisMo = ThisMo - 12
            ThisYr = ThisYr + 1
        End If
    Next QtrCount
    
    'Now "reform" a date
    DateStr = CStr(ThisMo) & "/" & "01" & "/" & CStr(ThisYr)
    gQtrBeginDate = CDate(DateStr)
End Function

Public Function gGetShiftBegDtime(ByVal aDate As Date, _
                                  ByVal aBeginHour As Integer) As Date

'**********************************************************************
'
'
'
'**********************************************************************

    'Assumes that the shifts in ShiftNames() are in time order
        gGetShiftBegDtime = aDate.AddDays(CInt(aBeginHour / 24))
End Function

Public Function gGetFirstShiftBegDtime(ByVal aDate As Date) As Date

'**********************************************************************
'
'
'
'**********************************************************************
 
    'Assumes that the shifts in ShiftNames() are in time order
        gGetFirstShiftBegDtime = aDate.AddDays(CInt(gShiftNames(1).BeginHour / 24))
End Function

Public Function gGetLastShiftBegDtime(ByVal aDate As Date) As Date

'**********************************************************************
'
'
'
'**********************************************************************

    'Assumes that the shifts in ShiftNames() are in time order
        gGetLastShiftBegDtime = aDate.AddDays _
                         (CInt(gShiftNames(UBound(gShiftNames)).BeginHour / 24))
End Function

Public Function gGetShiftStartHour(ByVal aShift As String) As Integer

'**********************************************************************
'This function returns the hour (military time) at which this shift
'begins.  If the function cannot determine a shift begin hour then the
'function returns 99.
'**********************************************************************
    
    Dim RowIdx As Integer
    Dim ShiftPos As Integer
    
    'This function may receive a shift name in any of the following
    'formats:
    '1) Day
    '2) DAY
    '3) Day shift
    '4) DAY SHIFT
    'etc.
    
    aShift = StrConv(aShift, vbUpperCase)
    ShiftPos = InStr(aShift, "SHIFT")
    
    If ShiftPos <> 0 Then
        aShift = Mid(aShift, 1, ShiftPos - 2)
    End If
        
    'Find the shift we are in (gShiftNames)
    'Assume that the shifts in gShiftInfo are in the correct order
    'timewise.
    
    gGetShiftStartHour = 99
    
    For RowIdx = 1 To UBound(gShiftNames)
        If StrConv(gShiftNames(RowIdx).ShiftName, vbUpperCase) = _
            StrConv(aShift, vbUpperCase) Then

            gGetShiftStartHour = gShiftNames(RowIdx).BeginHour
        End If
    Next RowIdx
End Function

Public Function gLastDayInQuarter(ByVal aDate As Date) As String

'**********************************************************************
'If aDate is the last date in a quarter then the function returns
'"1st", "2nd", "3rd", or "4th".  If aDate is not the last
'date in a quarter then the function returns "".
'**********************************************************************

    Dim Fybd As Date
    
    Dim QtrEnd1 As Date
    Dim QtrEnd2 As Date
    Dim QtrEnd3 As Date
    Dim QtrEnd4 As Date
        gLastDayInQuarter = String.Empty
    'Make sure we are only working with a date.
        aDate = CDate(Format(aDate, "MM/dd/yyyy"))
    Fybd = gGetBeginOfFiscalYear(aDate)
    
        QtrEnd1 = gQtrBeginDate(Fybd, 2).AddDays(-1)
        QtrEnd2 = gQtrBeginDate(Fybd, 3).AddDays(-1)
        QtrEnd3 = gQtrBeginDate(Fybd, 4).AddDays(-1)
        QtrEnd4 = gQtrBeginDate(Fybd, 1).AddDays(-1)
    
    If aDate = QtrEnd1 Then
        gLastDayInQuarter = "1st"
    End If
    If aDate = QtrEnd2 Then
        gLastDayInQuarter = "2nd"
    End If
    If aDate = QtrEnd3 Then
        gLastDayInQuarter = "3rd"
    End If
    If aDate = QtrEnd4 Then
        gLastDayInQuarter = "4th"
    End If
End Function

Public Sub gGetAllShiftsCbo(ByVal aMineName As String, _
                            ByVal aCboBox As ComboBox, _
                            ByVal aAddSelectItem As Boolean)

'**********************************************************************
'
'
'
'**********************************************************************
           
    On Error GoTo gGetAllShiftsCboError
    
    Dim ShiftNamesDynaset As OraDynaset
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt

    Dim ThisShiftName As String
    Dim ItemIdx As Integer
    
        'Set 
        params = gDBParams
      
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_shift_names
    'pMineName           IN     VARCHAR2,
    'pResult             IN OUT c_minenames)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_shift_names(:pMineName, " & _
                  ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        ShiftNamesDynaset = params("pResult").Value
        ClearParams(params)
   
    'Clear the by reference combo box
        For ItemIdx = 0 To aCboBox.Items.Count - 1
            aCboBox.Items.RemoveAt(0)
        Next ItemIdx
    
    If aAddSelectItem = True Then
            aCboBox.Items.Add("(Select shift...)")
    End If
    
    ShiftNamesDynaset.MoveFirst
    Do While Not ShiftNamesDynaset.EOF
        ThisShiftName = ShiftNamesDynaset.Fields("shift").Value
            aCboBox.Items.Add(StrConv(ThisShiftName, vbProperCase))
        ShiftNamesDynaset.MoveNext
    Loop
        
        aCboBox.Text = aCboBox.Items(0)
    
    ShiftNamesDynaset.Close
    Exit Sub
    
gGetAllShiftsCboError:
        MsgBox("Error getting shifts." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Shift Names Get Error")
        
    On Error Resume Next
        ClearParams(params)
    ShiftNamesDynaset.Close
End Sub

    Public Sub gResizeSsCols(ByRef aSpread As AxvaSpread, _
                             ByVal aCorrFactor As Single)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ColIdx As Long
        Dim MaxWidth As Single

        If aCorrFactor = 0 Then
            aCorrFactor = 1
        End If

        For ColIdx = 0 To aSpread.MaxCols
            MaxWidth = aSpread.get_MaxTextColWidth(ColIdx)
            aSpread.set_ColWidth(ColIdx, MaxWidth * aCorrFactor) ') =
        Next ColIdx
    End Sub

    Public Sub gEveryOtherGreen(ByRef aSpread As AxvaSpread)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim RowIdx As Integer

        With aSpread
            .Redraw = False
            For RowIdx = 1 To .MaxRows
                If gIsEvenNumber(RowIdx) = True Then
                    .BlockMode = True
                    .Row = RowIdx
                    .Row2 = RowIdx
                    .Col = 1
                    .Col2 = .MaxCols
                    .BackColor = Color.LightGreen ' &HD8FFD8    'Light light green
                    .BlockMode = False
                End If
                .Redraw = True
            Next RowIdx
        End With
    End Sub

Public Function gGetChoices(ByRef aChoiceDynaset As OraDynaset, _
                            ByVal aMineName As String, _
                            ByVal aCtgryName As String, _
                            ByVal aSortNumeric As Boolean) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo gGetChoicesError
    
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
                
        'Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
            
    params.Add("pCategoryName", aCtgryName, ORAPARM_INPUT)
    params("pCategoryName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    If aSortNumeric = False Then
        'PROCEDURE get_all_choices
        'pMineName           IN     VARCHAR2,
        'pCategoryName       IN     VARCHAR2,
        'pResult             IN OUT c_choices)
            'Set 
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_choices(:pMineName, " + _
                      ":pCategoryName, :pResult);end;", ORASQL_FAILEXEC)
            'Set 
            aChoiceDynaset = params("pResult").Value
    Else
        'PROCEDURE get_all_choices_num
        'pMineName           IN     VARCHAR2,
        'pCategoryName       IN     VARCHAR2,
        'pResult             IN OUT c_choices)
            'Set 
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_choices_num(:pMineName, " + _
                      ":pCategoryName, :pResult);end;", ORASQL_FAILEXEC)
            'Set 
            aChoiceDynaset = params("pResult").Value
    End If
        ClearParams(params)
    gGetChoices = True
    
    Exit Function
    
gGetChoicesError:
        MsgBox("Error getting choices." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Choices Error")
    
    On Error Resume Next
    ClearParams( params)
    gGetChoices = False
End Function

Public Function gGetChoicesActive(ByRef aChoiceDynaset As OraDynaset, _
                                  ByVal aMineName As String, _
                                  ByVal aCtgryName As String, _
                                  ByVal aSortNumeric As Boolean, _
                                  ByVal aDate As Date, _
                                  ByVal aShift As String) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo gGetChoicesError
    
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
                
        'Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
            
    params.Add("pCategoryName", aCtgryName, ORAPARM_INPUT)
    params("pCategoryName").serverType = ORATYPE_VARCHAR2
    
    params.Add("pDate", aDate, ORAPARM_INPUT)
    params("pDate").serverType = ORATYPE_DATE
    
    params.Add("pShift", StrConv(aShift, vbUpperCase), ORAPARM_INPUT)
    params("pShift").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    If aSortNumeric = False Then
        'PROCEDURE get_all_choices_active
        'pMineName           IN     VARCHAR2,
        'pCategoryName       IN     VARCHAR2,
        'pDate               IN     DATE,
        'pShift              IN     VARCHAR2,
        'pResult             IN OUT c_choices)
            'Set 
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_choices_active(:pMineName, " + _
                      ":pCategoryName, :pDate, :pShift, :pResult);end;", ORASQL_FAILEXEC)
            'Set 
            aChoiceDynaset = params("pResult").Value
    Else
        'PROCEDURE get_all_choices_num_active
        'pMineName           IN     VARCHAR2,
        'pCategoryName       IN     VARCHAR2,
        'pDate               IN     DATE,
        'pShift              IN     VARCHAR2,
        'pResult             IN OUT c_choices)
            'Set 
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_choices_num_active(:pMineName, " + _
                      ":pCategoryName, :pDate, :pShift, :pResult);end;", ORASQL_FAILEXEC)
            'Set 
            aChoiceDynaset = params("pResult").Value
    End If
        ClearParams(params)
    gGetChoicesActive = True
    
    Exit Function
    
gGetChoicesError:
        MsgBox("Error getting choices." & vbCrLf & _
           Err.Description, _
           vbOKOnly + vbExclamation, _
           "Choices Error")
    
    On Error Resume Next
        ClearParams(params)
    gGetChoicesActive = False
End Function

Public Function gGetAllMineNamesDyn(aMineNameDynaset As OraDynaset) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo gGetAllMineNamesDynError
     
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt

    'Get all existing mine names
        ' Set 
        params = gDBParams
 
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_all_mine_names
    'pResult IN OUT c_minenames)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_mine_names(:pResult);end;", ORASQL_FAILEXEC)
        'Set 
        aMineNameDynaset = params("pResult").Value
        ClearParams(params)
     
    gGetAllMineNamesDyn = True
    Exit Function
    
gGetAllMineNamesDynError:
        MsgBox("Error getting all mine names." & vbCrLf & _
        Err.Description, _
        vbOKOnly + vbExclamation, _
        "Mine Names Error")
        
    On Error Resume Next
    gGetAllMineNamesDyn = False
        ClearParams(params)
End Function

Public Function gGetAllMineNamesCbo(aMineNameCbo As ComboBox, _
                                    aIncludeSelect As Boolean, _
                                    aProspectMinesOnly As Boolean) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo gGetAllMineNamesCboError
     
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim MineNameDynaset As OraDynaset
    Dim RowIdx As Integer
        gGetAllMineNamesCbo = True
    'Get all existing mine names
        'Set 
        params = gDBParams
 
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_all_mine_names
    'pResult IN OUT c_minenames)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_mine_info(:pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MineNameDynaset = params("pResult").Value
        ClearParams(params)
    
    'Clear the Combobox
        For RowIdx = 0 To aMineNameCbo.Items.Count - 1
            aMineNameCbo.Items.RemoveAt(0)
        Next RowIdx
            
    If aIncludeSelect = True Then
            aMineNameCbo.Items.Add("(Select mine name...)")
    End If
    
    MineNameDynaset.MoveFirst
    Do While Not MineNameDynaset.EOF
        If aProspectMinesOnly = True Then
            'Only want mines with prospect data!
            If MineNameDynaset.Fields("mine_prospect").Value = 1 Then
                    aMineNameCbo.Items.Add(MineNameDynaset.Fields("mine_name").Value)
            End If
        Else
                aMineNameCbo.Items.Add(MineNameDynaset.Fields("mine_name").Value)
        End If
        
        MineNameDynaset.MoveNext
    Loop
    
    'Set to the first item in the cbo list
        aMineNameCbo.Text = aMineNameCbo.Items(0)
     
    MineNameDynaset.Close
    Exit Function
    
gGetAllMineNamesCboError:
        MsgBox("Error getting all mine names." & vbCrLf & _
        Err.Description, _
        vbOKOnly + vbExclamation, _
        "Mine Names Error")
        
    On Error Resume Next
        ClearParams(params)
    MineNameDynaset.Close
End Function

Public Function gGetTypes(ByRef aTypeDynaset As OraDynaset, _
                          ByVal aMineName As String, _
                          ByVal aCategory As String) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
 
        'Set 
        params = gDBParams
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pCategoryName", aCategory, ORAPARM_INPUT)
    params("pCategoryName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_all_ctgry_types
    'pMineName           IN     VARCHAR2,
    'pCategoryName       IN     VARCHAR2,
    'pResult             IN OUT c_types)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_ctgry_types(:pMineName," + _
                  ":pCategoryName, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        aTypeDynaset = params("pResult").Value
        ClearParams(params)
    
    gGetTypes = True
    Exit Function
    
gGetTypesError:
        MsgBox("Error getting types." & vbCrLf & _
        Err.Description, _
        vbOKOnly + vbExclamation, _
        "Types Error")
    
    On Error Resume Next
        ClearParams(params)
    gGetTypes = False
End Function

Public Function gGetEqptTypeName(ByVal aMineName As String, _
                                 ByVal aEqptName As String) As String

'**********************************************************************
'
'
'
'**********************************************************************
    
    On Error GoTo gGetEqptTypeNameError
    
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
        
        'Set 
        params = gDBParams
                  
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptName", aEqptName, ORAPARM_INPUT)
    params("pEqptName").serverType = ORATYPE_VARCHAR2
                 
        params.Add("pEqptTypeName", "", ORAPARM_OUTPUT)
    params("pEqptTypeName").serverType = ORATYPE_VARCHAR2
    
    'Procedure get_eqpt_type_name
    'pMineName     IN     VARCHAR2,
    'pEqptName     IN     VARCHAR2,
    'pEqptTypeName IN OUT VARCHAR2)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities2.get_eqpt_type_name(:pMineName," + _
                  ":pEqptName, :pEqptTypeName);end;", ORASQL_FAILEXEC)
                    
    gGetEqptTypeName = params("pEqptTypeName").Value
          
        ClearParams(params)
    
    Exit Function
    
gGetEqptTypeNameError:
        MsgBox("Error getting equipment type name." & vbCrLf & _
        Err.Description, _
        vbOKOnly + vbExclamation, _
        "Equipment Type Name Get Error")
    
    On Error Resume Next
    gGetEqptTypeName = ""
End Function

Public Function gGetProdMatlTypeName(ByVal aMeasureName As String) As String

'**********************************************************************
'
'
'
'**********************************************************************

    gGetProdMatlTypeName = ""
    
    aMeasureName = StrConv(aMeasureName, vbUpperCase)
    
    If InStr(aMeasureName, "CONCENTRATE TONS") <> 0 Then
        gGetProdMatlTypeName = "Concentrate"
    End If
    
    If InStr(aMeasureName, "PEBBLE TONS") <> 0 Then
        gGetProdMatlTypeName = "Pebble"
    End If
    
    If InStr(aMeasureName, "IP TONS") <> 0 Then
        gGetProdMatlTypeName = "IP"
    End If
End Function

Public Sub gSaveVbError(ByVal aErrNum As Integer, _
                        ByVal aErrDesc As String, _
                        ByVal aErrSource As String, _
                        ByVal aErrComment As String, _
                        ByVal aMineName As String, _
                        ByVal aLastServerErrNum As Integer, _
                        ByVal aLastServerErrText As String)

'**********************************************************************
'
'
'
'**********************************************************************
        
    'This procedure will set:
    'UserId
    'ErrDateTime
    
    On Error GoTo gSaveVbErrorError
    
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim Status As Integer
        
        'Set 
        params = gDBParams
 
    params.Add("pErrUserId", gUserName, ORAPARM_INPUT)
    params("pErrUserId").serverType = ORATYPE_VARCHAR2
    
    params.Add("pErrDateTime", Now, ORAPARM_INPUT)
    params("pErrDateTime").serverType = ORATYPE_DATE
    
    params.Add("pErrNum", aErrNum, ORAPARM_INPUT)
    params("pErrNum").serverType = ORATYPE_NUMBER

    If aErrDesc = "" Then
        aErrDesc = " "
    End If
    params.Add("pErrDesc", aErrDesc, ORAPARM_INPUT)
    params("pErrDesc").serverType = ORATYPE_VARCHAR2
    
    If aErrSource = "" Then
        aErrSource = " "
    End If
    params.Add("pErrSource", aErrSource, ORAPARM_INPUT)
    params("pErrSource").serverType = ORATYPE_VARCHAR2

    If aErrComment = "" Then
        aErrComment = " "
    End If
    params.Add("pErrComment", aErrComment, ORAPARM_INPUT)
    params("pErrComment").serverType = ORATYPE_VARCHAR2
    
    params.Add("pErrMineName", aMineName, ORAPARM_INPUT)
    params("pErrMineName").serverType = ORATYPE_VARCHAR2
    
    params.Add("pLastServerErrNum", aLastServerErrNum, ORAPARM_INPUT)
    params("pLastServerErrNum").serverType = ORATYPE_NUMBER
        
    params.Add("pLastServerErrText", aLastServerErrText, ORAPARM_INPUT)
    params("pLastServerErrText").serverType = ORATYPE_VARCHAR2
   
        params.Add("pErrResult", 0, ORAPARM_OUTPUT)
    params("pErrResult").serverType = ORATYPE_NUMBER
    
    'PROCEDURE update_vb_errors
    'pErrUserId         IN     VARCHAR2,
    'pErrDateTime       IN     DATE,
    'pErrNum            IN     NUMBER,
    'pErrDesc           IN     VARCHAR2,
    'pErrSource         IN     VARCHAR2,
    'pErrComment        IN     VARCHAR2,
    'pErrMineName       IN     VARCHAR2,
    'pLastServerErrNum  IN     NUMBER,
    'pLastServerErrText IN     VARCHAR2,
    'pErrResult         IN OUT NUMBER
        ' Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities2.update_vb_errors(:pErrUserId, " & _
                  ":pErrDateTime, :pErrNum, :pErrDesc, " & _
                  ":pErrSource, :pErrComment, :pErrMineName, " & _
                  ":pLastServerErrNum, :pLastServerErrText, :pErrResult);end;", ORASQL_FAILEXEC)
    Status = params("pErrResult").Value
     
        ClearParams(params)
 
    Exit Sub
    
gSaveVbErrorError:
    'Since this procedure is being called when an error has occurred in MOIS
    'I don't want to complicate things for the user by returning an error
    'message from trying to save the error that has already occurred to Oracle.
    'This no message box here -- just continue onwards and upwards.
    On Error Resume Next
        ClearParams(params)
End Sub

Public Function gGetDlName(ByVal aDlPitName As String) As String

'**********************************************************************
'
'
'
'**********************************************************************

    gGetDlName = ""
    
    'Dragline names are always like "Dragline #10", "Dragline #11", etc.
    'Associated dragline pits are always "Dragline #10 pit",
    '"Dragline #11 pit", etc.
    
    If Len(aDlPitName) > 4 Then
        gGetDlName = Trim(Mid(aDlPitName, 1, Len(aDlPitName) - 4))
    Else
        gGetDlName = ""
    End If
End Function

Public Function gGetDlPitName(ByVal aDlName As String) As String

'**********************************************************************
'
'
'
'**********************************************************************

    gGetDlPitName = ""
    
    'Dragline names are always like "Dragline #10", "Dragline #11", etc.
    'Associated dragline pits are always "Dragline #10 pit",
    '"Dragline #11 pit", etc.
    
    gGetDlPitName = Trim(aDlName) & " pit"
End Function

Public Sub gGetMineProducts()

'**********************************************************************
'
'
'
'**********************************************************************
    
    On Error GoTo GetMineProductsError
     
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim MineProdDynaset As OraDynaset
    Dim ThisMineProd As String
    
    'Will set:
    'gMineHasPb
    'gMineHasCn
    'gMineHasIp
    
    gMineHasPb = False
    gMineHasCn = False
    gMineHasIp = False
    
    'Get mine products
        'Set 
        params = gDBParams
 
    params.Add("pMineName", gActiveMineNameLong, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
  
    params.Add("pCategoryName", "Product material", ORAPARM_INPUT)
    params("pCategoryName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_all_ctgry_types
    'pMineName           IN     VARCHAR2,
    'pCategoryName       IN     VARCHAR2,
    'pResult             IN OUT c_types)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_ctgry_types(:pMineName, " & _
                  ":pCategoryName, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MineProdDynaset = params("pResult").Value
        ClearParams(params)
    
    MineProdDynaset.MoveFirst

    gMineHasPb = False
    gMineHasCn = False
    gMineHasIp = False

    Do While Not MineProdDynaset.EOF
        ThisMineProd = MineProdDynaset.Fields("type_name").Value
        
        If StrConv(ThisMineProd, vbUpperCase) = "PEBBLE" Then
            gMineHasPb = True
        End If
        
        If StrConv(ThisMineProd, vbUpperCase) = "CONCENTRATE" Then
            gMineHasCn = True
        End If
        
        If StrConv(ThisMineProd, vbUpperCase) = "IP" Then
            gMineHasIp = True
        End If
        MineProdDynaset.MoveNext
    Loop
    
    MineProdDynaset.Close
    
    Exit Sub
    
GetMineProductsError:
        MsgBox("Error getting mine products." & vbCrLf & _
        Err.Description, _
        vbOKOnly + vbExclamation, _
        "Mine Products Get Error")
        
    On Error Resume Next
        ClearParams(params)
    MineProdDynaset.Close
End Sub

Function gMineHasSubProducts(aMineName As String) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo gMineHasSubProductsError
 
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim EqptMatlDynaset As OraDynaset
    Dim MatlName As String
    Dim ProdMatlTypeName As String
        
        ' Set 
        params = gDBParams
    
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
 
        params.Add("pEqptTypeName", "Physical product bin", ORAPARM_OUTPUT)
    params("pEqptTypeName").serverType = ORATYPE_VARCHAR2
        
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_eqpttype_matls
    'pMineName           IN     VARCHAR2,
    'pEqptTypeName       IN     VARCHAR2,
    'pResult             IN OUT c_eqptmatls)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_eqpttype_matls(:pMineName," + _
                  ":pEqptTypeName, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        EqptMatlDynaset = params("pResult").Value
        ClearParams(params)
    
    gMineHasSubProducts = False
        
    EqptMatlDynaset.MoveFirst
    Do While Not EqptMatlDynaset.EOF
        MatlName = EqptMatlDynaset.Fields("matl_name").Value
        ProdMatlTypeName = EqptMatlDynaset.Fields("prod_matl_type_name").Value
        
        If MatlName <> ProdMatlTypeName Then
            gMineHasSubProducts = True
            Exit Do
        End If
        EqptMatlDynaset.MoveNext
    Loop
 
    EqptMatlDynaset.Close
    
    Exit Function
    
gMineHasSubProductsError:
        MsgBox("Error determining if mine has subproducts." & vbCrLf & _
        Err.Description, _
        vbOKOnly + vbExclamation, _
        "Subproduct Error")
        
    On Error Resume Next
        ClearParams(params)
    EqptMatlDynaset.Close
    gMineHasSubProducts = False
End Function

Public Function gGetMineHasIp(ByVal aMineName As String) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************
    
    On Error GoTo gGetMineHasIpError
     
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
    Dim MineProdDynaset As OraDynaset
    Dim ThisMineProd As String
     
    'Get mine products
        ' Set 
        params = gDBParams
 
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
  
    params.Add("pCategoryName", "Product material", ORAPARM_INPUT)
    params("pCategoryName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_all_ctgry_types
    'pMineName           IN     VARCHAR2,
    'pCategoryName       IN     VARCHAR2,
    'pResult             IN OUT c_types)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_ctgry_types(:pMineName, " & _
                  ":pCategoryName, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MineProdDynaset = params("pResult").Value
        ClearParams(params)
    
    MineProdDynaset.MoveFirst
    
    gGetMineHasIp = False
    Do While Not MineProdDynaset.EOF
        ThisMineProd = MineProdDynaset.Fields("type_name").Value
                
        If StrConv(ThisMineProd, vbUpperCase) = "IP" Then
            gGetMineHasIp = True
            Exit Do
        End If
        MineProdDynaset.MoveNext
    Loop
    
    MineProdDynaset.Close
    
    Exit Function
    
gGetMineHasIpError:
        MsgBox("Error getting mine products." & vbCrLf & _
        Err.Description, _
        vbOKOnly + vbExclamation, _
        "Mine Products Get Error")
        
    On Error Resume Next
        ClearParams(params)
    MineProdDynaset.Close
End Function

Public Function gGetShiftNames(ByVal aMineName As String, _
                               ByRef aShiftNames() As gShiftNamesType) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************

    On Error GoTo gGetShiftNamesError
    
    Dim ShiftNamesDynaset As OraDynaset
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt

    Dim ThisShiftName As String
    Dim ThisBeginTime As String
    Dim ThisEndTime As String
    Dim ThisShiftLength As Integer
    Dim ThisShiftOrder As Integer
    Dim ThisSampBeginTime As String
    Dim ThisSampEndTime As String
    Dim ThisBegHr As Integer
    Dim ThisBegMin As Integer
    Dim ThisEndHr As Integer
    Dim ThisEndMin As Integer
    Dim ShiftNameCnt As Integer
    Dim RecordCnt As Integer
    
    'Need to assign the following:
    '1) gShiftNames() As String
    '2) gNumShifts As Integer
    '3) gFirstShift As String
    '4) gLastShift as String
    '5) BeginHour As Integer
    '6) BeginMinute As Integer
    '7) EndHour As Integer
    '8) EndMinute as Integer

        'Set 
        params = gDBParams
      
    params.Add("pMineName", aMineName, ORAPARM_INPUT)
    params("pMineName").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
    params("pResult").serverType = ORATYPE_CURSOR
    
    'PROCEDURE get_shift_names
    'pMineName           IN     VARCHAR2,
    'pResult             IN OUT c_minenames)
        ' Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_shift_names(:pMineName, " & _
                  ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        ShiftNamesDynaset = params("pResult").Value
        ClearParams(params)
    RecordCnt = ShiftNamesDynaset.RecordCount
    ReDim aShiftNames(RecordCnt)
    
    ShiftNameCnt = 0
    
    ShiftNamesDynaset.MoveFirst
    
    Do While Not ShiftNamesDynaset.EOF
        ThisShiftName = ShiftNamesDynaset.Fields("shift").Value
        ThisBeginTime = ShiftNamesDynaset.Fields("begin_time").Value
        ThisEndTime = ShiftNamesDynaset.Fields("end_time").Value
        ThisShiftLength = ShiftNamesDynaset.Fields("shift_length_hrs").Value
        ThisShiftOrder = 0
        ThisSampBeginTime = ""
        ThisSampEndTime = ""
               
        ShiftNameCnt = ShiftNameCnt + 1
  
        'Need to get hour and minute for ThisBeginTime
        ThisBegHr = Val(Mid(ThisBeginTime, 1, 2))
        ThisBegMin = Val(Mid(ThisBeginTime, 4))
        
        'Need to get hour and minute for ThisBeginTime
        ThisEndHr = Val(Mid(ThisEndTime, 1, 2))
        ThisEndMin = Val(Mid(ThisEndTime, 4))
        
        aShiftNames(ShiftNameCnt).ShiftName = ThisShiftName
        aShiftNames(ShiftNameCnt).BeginTime = ThisBeginTime
        aShiftNames(ShiftNameCnt).EndTime = ThisEndTime
        aShiftNames(ShiftNameCnt).ShiftLength = ThisShiftLength
        aShiftNames(ShiftNameCnt).ShiftOrder = ThisShiftOrder
        aShiftNames(ShiftNameCnt).SampBeginTime = ThisSampBeginTime
        aShiftNames(ShiftNameCnt).SampEndTime = ThisSampEndTime
        
        aShiftNames(ShiftNameCnt).BeginHour = ThisBegHr
        aShiftNames(ShiftNameCnt).BeginMinute = ThisBegMin
        
        aShiftNames(ShiftNameCnt).EndHour = ThisEndHr
        aShiftNames(ShiftNameCnt).EndMinute = ThisEndMin
        
        ShiftNamesDynaset.MoveNext
    Loop
 
    ShiftNamesDynaset.Close
    gGetShiftNames = True
    Exit Function
    
gGetShiftNamesError:
        MsgBox("Error getting shift names." & vbCrLf & _
        Err.Description, _
        vbOKOnly + vbExclamation, _
        "Shift Names Get Error")
        
    On Error Resume Next
    gGetShiftNames = False
        ClearParams(params)
    ShiftNamesDynaset.Close
End Function

Public Function gGetDlHasMtxTons(ByVal aMineName As String) As Boolean

'**********************************************************************
'
'
'
'**********************************************************************
           
    On Error GoTo gGetDlHasMtxTonsError
    
    Dim params As OraParameters
    Dim SQLStmt As OraSqlStmt
        Dim MineInfoDynaset As OraDynaset

        'Get mine information from MINES_MOIS
        'Set 
        params = gDBParams
 
    params.Add("pMine", aMineName, ORAPARM_INPUT)
    params("pMine").serverType = ORATYPE_VARCHAR2
    
        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_one_mine_info(:pMine, " + _
                 ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MineInfoDynaset = params("pResult").Value
        ClearParams(params)

        'Should be only one row returned.
        MineInfoDynaset.MoveFirst()

        If MineInfoDynaset.Fields("dl_has_mtx_tons").Value = 1 Then
            gDlHasMtxTons = True
        Else
            gDlHasMtxTons = False
        End If

        MineInfoDynaset.Close()

        Return gDlHasMtxTons

        Exit Function

gGetDlHasMtxTonsError:
        MsgBox("Error getting shift names." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Shift Names Get Error")

        On Error Resume Next
        gGetDlHasMtxTons = False
        ClearParams(params)
        MineInfoDynaset.Close()
    End Function

    Public Function gGetDlMtxYdsSource(ByVal aMineName As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetDlMtxYdsSourceError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim MineInfoDynaset As OraDynaset

        'Get mine information from MINES_MOIS
        'Set 
        params = gDBParams

        params.Add("pMine", aMineName, ORAPARM_INPUT)
        params("pMine").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_one_mine_info(:pMine, " + _
                 ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MineInfoDynaset = params("pResult").Value
        ClearParams(params)

        'Should be only one row returned.
        MineInfoDynaset.MoveFirst()

        gGetDlMtxYdsSource = MineInfoDynaset.Fields("mtx_yds_source").Value

        MineInfoDynaset.Close()

        Exit Function

gGetDlMtxYdsSourceError:
        MsgBox("Error getting mtx yards source." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Mtx Yards Source Error")

        On Error Resume Next
        gGetDlMtxYdsSource = ""
        ClearParams(params)
        MineInfoDynaset.Close()
    End Function

    Public Function gGetProspGridType(ByVal aMineName As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspGridTypeError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim MineInfoDynaset As OraDynaset

        'Get mine information from MINES_MOIS
        'Set 
        params = gDBParams

        params.Add("pMine", aMineName, ORAPARM_INPUT)
        params("pMine").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_one_mine_info(:pMine, " + _
                 ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MineInfoDynaset = params("pResult").Value
        ClearParams(params)

        'Should be only one row returned.
        MineInfoDynaset.MoveFirst()

        gGetProspGridType = MineInfoDynaset.Fields("prosp_grid_type").Value

        MineInfoDynaset.Close()

        Exit Function

gGetProspGridTypeError:
        MsgBox("Error getting prospect grid type." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Prospect Grid Type Source Error")

        On Error Resume Next
        gGetProspGridType = ""
        ClearParams(params)
        MineInfoDynaset.Close()
    End Function

    Public Function gGetCalcSplitConc(ByVal aMineName As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetCalcSplitConcError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim MineInfoDynaset As OraDynaset

        'Get mine information from MINES_MOIS
        'Set 
        params = gDBParams

        params.Add("pMine", aMineName, ORAPARM_INPUT)
        params("pMine").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_one_mine_info(:pMine, " + _
                 ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MineInfoDynaset = params("pResult").Value
        ClearParams(params)

        'Should be only one row returned.
        MineInfoDynaset.MoveFirst()

        If MineInfoDynaset.Fields("calc_split_conc").Value = 1 Then
            gGetCalcSplitConc = True
        Else
            gGetCalcSplitConc = False
        End If

        MineInfoDynaset.Close()

        Exit Function

gGetCalcSplitConcError:
        MsgBox("Error getting data." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Data Get Error")

        On Error Resume Next
        gGetCalcSplitConc = False
        ClearParams(params)
        MineInfoDynaset.Close()
    End Function

    Public Function gRoundFifty(ByVal aNumber As Object) As Long

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim IsNegative As Boolean
        Dim TempNumber As Single
        Dim TempNumberInt As Single
        Dim TempNumberFra As Single
        gRoundFifty = 0
        If aNumber < 0 Then
            IsNegative = True
        Else
            IsNegative = False
        End If

        TempNumber = Abs(aNumber)
        TempNumber = Int(TempNumber)
        TempNumberInt = TempNumber / 100 - Int(TempNumber / 100)
        TempNumberFra = TempNumberInt * 100

        If TempNumberFra <= 25 Then
            gRoundFifty = TempNumber - TempNumberFra
        End If

        If TempNumberFra >= 25 And TempNumberFra <= 50 Then
            gRoundFifty = TempNumber + (50 - TempNumberFra)
        End If

        If TempNumberFra >= 50 And TempNumberFra < 75 Then
            gRoundFifty = TempNumber - TempNumberFra + 50
        End If

        If TempNumberFra >= 75 And TempNumberFra <= 99.99 Then
            gRoundFifty = TempNumber + (100 - TempNumberFra)
        End If

        If IsNegative Then
            gRoundFifty = gRoundFifty * -1
        End If
    End Function

    Public Function gRound(ByVal aNumber As Double, ByVal aRound As Integer) As Double

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        gRound = 0

        Dim Z1 As Double
        Dim Z2 As Double
        Dim Z3 As Double
        Dim Z4 As Double
        Dim Z5 As Double
        Dim Z6 As Double
        Dim Z7 As Double
        Dim Z8 As Double
        Dim Z9 As Double
        Dim Negative As Boolean

        Z1 = 0
        Z2 = 0
        Z3 = 0
        Z4 = 0
        Z5 = 0
        Z6 = 0
        Z7 = 0
        Z8 = 0
        Z9 = 0

        Negative = False
        If aNumber < 0 Then
            Negative = True
        End If

        aNumber = Abs(aNumber)

        Z8 = aNumber * (10 ^ aRound)
        Z9 = Int(Z8)

        'Z1 = (aNumber * (10 ^ aRound)) - Int(aNumber * (10 ^ aRound))
        Z1 = Z8 - Z9

        'Z2 = Int(Z1 * 10)
        Z2 = Z1

        If Z2 >= 0.4999999 Then
            Z3 = 1
        Else
            Z3 = 0
        End If

        Z6 = aNumber * (10 ^ aRound)
        Z7 = Int(Z6)

        Z4 = Z7 + Z3

        Z5 = Z4 / (10 ^ aRound)

        If Negative Then
            Z5 = -1 * Z5
        End If

        gRound = Z5
    End Function

    Public Function gRoundTen(ByVal aNumber As Object) As Long

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim Negative As Boolean

        Negative = False
        If aNumber < 0 Then
            Negative = True
        End If

        gRoundTen = gRound(aNumber / 100, 1) * 100

        If Negative Then
            gRoundTen = -1 * gRoundTen
        End If
    End Function

    Public Function gRoundFive(ByVal aNumber As Object) As Long

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim Negative As Boolean

        Negative = False
        If aNumber < 0 Then
            Negative = True
        End If

        gRoundFive = gRoundTen(aNumber * 2) / 2

        If Negative Then
            gRoundFive = -1 * gRoundFive
        End If
    End Function

    Public Function gGetShiftForDateTime(ByVal aMineName As String, _
                                         ByVal aDate As Date, _
                                         ByVal aMilitaryTime As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetShiftForDateTimeError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim DateTime As Date

        'Create a date time from aDate and aMilitaryTime
        'If IsDate(Format(aDate, "mm/dd/yyyy") & " " & aMilitaryTime) Then
        '    DateTime = CDate(Format(aDate, "mm/dd/yyyy") & " " & aMilitaryTime)
        'Else
        '    gGetShiftForDateTime = "?"
        '    Exit Function
        'End If

        ' Set 
        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pDate", aDate, ORAPARM_INPUT)
        params("pDate").serverType = ORATYPE_DATE

        params.Add("pMilitaryTime", aMilitaryTime, ORAPARM_INPUT)
        params("pMilitaryTime").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", "", ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_VARCHAR2

        'Procedure get_shift_for_datetime
        'pMineName         IN     VARCHAR2,
        'pDate             IN     DATE,
        'pMilitaryTime     IN     VARCHAR2,
        'pResult           IN OUT VARCHAR2)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities2.get_shift_for_datetime(:pMineName, " + _
                 ":pDate, :pMilitaryTime, :pResult);end;", ORASQL_FAILEXEC)
        gGetShiftForDateTime = params("pResult").Value
        ClearParams(params)

        Exit Function

gGetShiftForDateTimeError:
        MsgBox("Error getting shift." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Shift Get Error")

        On Error Resume Next
        gGetShiftForDateTime = "?"
        ClearParams(params)
    End Function

    Public Function gGetProspStandard(ByVal aProspStandard As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Select Case aProspStandard
            Case Is = "100% Prospect"
                gGetProspStandard = "100%PROSPECT"
            Case Is = "Catalog"
                gGetProspStandard = "CATALOG"
            Case Else
                gGetProspStandard = ""
        End Select
    End Function

    Public Function gGetProspStandardRev(ByVal aProspStandard As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Select Case aProspStandard
            Case Is = "100%PROSPECT"
                gGetProspStandardRev = "100% Prospect"
            Case Is = "CATALOG"
                gGetProspStandardRev = "Catalog"
            Case Else
                gGetProspStandardRev = ""
        End Select
    End Function

    Public Function gGetHasCtlgReserves(ByVal aMineName As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetHasCtlgReservesError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim MineInfoDynaset As OraDynaset

        'Get mine information from MINES_MOIS
        ' Set 
        params = gDBParams

        params.Add("pMine", aMineName, ORAPARM_INPUT)
        params("pMine").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_one_mine_info(:pMine, " + _
                 ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MineInfoDynaset = params("pResult").Value
        ClearParams(params)

        'Should be only one row returned.
        MineInfoDynaset.MoveFirst()

        If MineInfoDynaset.Fields("has_ctlg_reserves").Value = 1 Then
            gGetHasCtlgReserves = True
        Else
            gGetHasCtlgReserves = False
        End If

        MineInfoDynaset.Close()

        Exit Function

gGetHasCtlgReservesError:
        MsgBox("Error getting has catalog reserves." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Has Catalog Reserves Get Error")

        On Error Resume Next
        gGetHasCtlgReserves = False
        ClearParams(params)
        MineInfoDynaset.Close()
    End Function

    Public Function gGetInvAdjAppliedShifts(ByVal aMineName As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetInvAdjAppliedShiftsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim MineInfoDynaset As OraDynaset

        'Get mine information from MINES_MOIS
        ' Set 
        params = gDBParams

        params.Add("pMine", aMineName, ORAPARM_INPUT)
        params("pMine").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        ' Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_one_mine_info(:pMine, " + _
                 ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MineInfoDynaset = params("pResult").Value
        ClearParams(params)

        'Should be only one row returned.
        MineInfoDynaset.MoveFirst()

        If MineInfoDynaset.Fields("has_ctlg_reserves").Value = 1 Then
            gGetInvAdjAppliedShifts = True
        Else
            gGetInvAdjAppliedShifts = False
        End If

        MineInfoDynaset.Close()

        Exit Function

gGetInvAdjAppliedShiftsError:
        MsgBox("Error getting invadj applied shifts." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "InvAdj Applied Shifts Get Error")

        On Error Resume Next
        gGetInvAdjAppliedShifts = False
        ClearParams(params)
        MineInfoDynaset.Close()
    End Function

    Public Sub gMiscMoisSetup()

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        gGenDelayComment = "Always enter the TOTAL hours that the delay reason " & _
                           "occurred for the shift (adding up the hours for all " & _
                           " occurrences of " & _
                           "the delay reason for the shift).  You may enter the number " & _
                           "of times the delay occured during the shift in the " & _
                           "'Occur' column, however MOIS will not multiply the 'Occur' value " & _
                           "by the delay reason hours that you enter!"
    End Sub

    Public Function gGetBinNum(ByVal aBinNum As String) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Makes the assumption that bins are name Bin# 1, Bin #2, Bin #3, etc.
        'Will have to mak this more generic as needed.

        gGetBinNum = Val(Mid(aBinNum, 6))
    End Function

    Public Function gGetDateFromDateTime(ByVal aDateTime As Date) As Date

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        gGetDateFromDateTime = CDate(Format(aDateTime, "MM/dd/yyyy"))
    End Function

    Public Function gGetMatlCorrection(ByVal aMineName As String, _
                                       ByVal aDate As Date, _
                                       ByVal aShift As String, _
                                       ByVal aEqptTypeName As String, _
                                       ByVal aEqptName As String, _
                                       ByVal aMatltype As String, _
                                       ByVal aMatl As String) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetMatlCorrectionError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim CorrectionDynaset As OraDynaset
        Dim RecordCount As Integer

        'Set 
        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pDate", aDate, ORAPARM_INPUT)
        params("pDate").serverType = ORATYPE_DATE

        params.Add("pShift", StrConv(aShift, vbUpperCase), ORAPARM_INPUT)
        params("pShift").serverType = ORATYPE_VARCHAR2

        params.Add("pEqptTypeName", aEqptTypeName, ORAPARM_INPUT)
        params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptName", aEqptName, ORAPARM_INPUT)
        params("pEqptName").serverType = ORATYPE_VARCHAR2

    params.Add("pMatlTypeName", aMatltype, ORAPARM_INPUT)
        params("pMatlTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pMatlName", aMatl, ORAPARM_INPUT)
        params("pMatlName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_matl_correction
        'pMineName           IN     VARCHAR2,
        'pDate               IN     DATE,
        'pShift              IN     VARCHAR2,
        'pEqptTypeName       IN     VARCHAR2,
        'pEqptName           IN     VARCHAR2,
        'pMatlTypeName       IN     VARCHAR2,
        'pMatlName           IN     VARCHAR2,
        'pResult             IN OUT c_corrections)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_field.get_matl_correction(:pMineName, " & _
                                         ":pDate, :pShift, :pEqptTypeName, :pEqptName, " & _
                                         ":pMatlTypeName, :pMatlName, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        CorrectionDynaset = params("pResult").Value
        ClearParams(params)
        RecordCount = CorrectionDynaset.RecordCount

        If RecordCount = 1 Then
            CorrectionDynaset.MoveFirst()
            gGetMatlCorrection = CorrectionDynaset.Fields("correction_factor").Value
        Else
            gGetMatlCorrection = 0
        End If

        CorrectionDynaset.Close()

        Exit Function

gGetMatlCorrectionError:
        MsgBox("Error getting material correction." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Material Correction Get Error")

        On Error Resume Next
        ClearParams(params)
        CorrectionDynaset.Close()
    End Function

    Public Function gGetPeriodicEqptMsrmnt(ByVal aMineName As String, _
                                           ByVal aDate As Date, _
                                           ByVal aShift As String, _
                                           ByVal aEqptTypeName As String, _
                                           ByVal aEqptName As String, _
                                           ByVal aMeasureName As String) As Double

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gPeriodicEqptMsrError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim PeriodicEqptMsrmntDynaset As OraDynaset
        Dim RecordCount As Integer

        ' Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pDate", aDate, ORAPARM_INPUT)
        params("pDate").serverType = ORATYPE_DATE

    params.Add("pShift", StrConv(aShift, vbUpperCase), ORAPARM_INPUT)
        params("pShift").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptTypeName", aEqptTypeName, ORAPARM_INPUT)
        params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptName", aEqptName, ORAPARM_INPUT)
        params("pEqptName").serverType = ORATYPE_VARCHAR2

    params.Add("pMeasureName", aMeasureName, ORAPARM_INPUT)
        params("pMeasureName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_eqpt_prdc_msrmnt
        'pMineName           IN     VARCHAR2,
        'pDate               IN     DATE,
        'pShift              IN     VARCHAR2,
        'pEqptTypeName       IN     VARCHAR2,
        'pEqptName           IN     VARCHAR2,
        'pMeasureName        IN     VARCHAR2,
        'pResult             IN OUT c_currprdcmsrmnts);
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prdcmsrmnt.get_eqpt_prdc_msrmnt(:pMineName, " & _
                                         ":pDate, :pShift, :pEqptTypeName, :pEqptName, " & _
                                         ":pMeasureName, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        PeriodicEqptMsrmntDynaset = params("pResult").Value
        ClearParams(params)
        RecordCount = PeriodicEqptMsrmntDynaset.RecordCount

        If RecordCount = 1 Then
            PeriodicEqptMsrmntDynaset.MoveFirst()
            gGetPeriodicEqptMsrmnt = PeriodicEqptMsrmntDynaset.Fields("value").Value
        Else
            gGetPeriodicEqptMsrmnt = 0
        End If

        PeriodicEqptMsrmntDynaset.Close()

        Exit Function

gPeriodicEqptMsrError:
        MsgBox("Error getting periodic equipment measurement." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Periodic Equipment Measurement Get Error")

        On Error Resume Next
        ClearParams(params)
        PeriodicEqptMsrmntDynaset.Close()
    End Function

    Public Function gGetEqptCalcSum(ByVal aMineName As String, _
                                    ByVal aBeginDate As Date, _
                                    ByVal aBeginShift As String, _
                                    ByVal aEndDate As Date, _
                                    ByVal aEndShift As String, _
                                    ByVal aEqptTypeName As String, _
                                    ByVal aEqptName As String, _
                                    ByVal aMeasureName As String) As Double

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetEqptCalcSumError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        gGetEqptCalcSum = 0

        ' Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pBeginDate", aBeginDate, ORAPARM_INPUT)
        params("pBeginDate").serverType = ORATYPE_DATE

    params.Add("pBeginShift", StrConv(aBeginShift, vbUpperCase), ORAPARM_INPUT)
        params("pBeginShift").serverType = ORATYPE_VARCHAR2

    params.Add("pEndDate", aEndDate, ORAPARM_INPUT)
        params("pEndDate").serverType = ORATYPE_DATE

    params.Add("pEndShift", StrConv(aEndShift, vbUpperCase), ORAPARM_INPUT)
        params("pEndShift").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptTypeName", aEqptTypeName, ORAPARM_INPUT)
        params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptName", aEqptName, ORAPARM_INPUT)
        params("pEqptName").serverType = ORATYPE_VARCHAR2

    params.Add("pMeasureName", aMeasureName, ORAPARM_INPUT)
        params("pMeasureName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        'PROCEDURE get_eqpt_calc_sum
        'pMineName          IN     VARCHAR2,
        'pEqptTypeName      IN     VARCHAR2,
        'pEqptName          IN     VARCHAR2,
        'pMeasureName       IN     VARCHAR2,
        'pBeginDate         IN     DATE,
        'pBeginShift        IN     VARCHAR2,
        'pEndDate           IN     DATE,
        'pEndShift          IN     VARCHAR2,
        'pResult            IN OUT NUMBER)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities2.get_eqpt_calc_sum(:pMineName, " & _
                                         ":pEqptTypeName, :pEqptName, " & _
                                         ":pMeasureName, :pBeginDate, " & _
                                         ":pBeginShift, :pEndDate, :pEndShift, :pResult);end;", ORASQL_FAILEXEC)

        If Not IsDBNull(params("pResult").Value) Then
            gGetEqptCalcSum = params("pResult").Value
        Else
            gGetEqptCalcSum = 0
        End If

        ClearParams(params)

        Exit Function

gGetEqptCalcSumError:
        MsgBox("Error getting sum." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Eqpt Calc Sum Get Error")

        On Error Resume Next
        ClearParams(params)
    End Function

    Public Function gGetSzdFdTons(ByVal aMineName As String, _
                                  ByVal aBeginDate As Date, _
                                  ByVal aBeginShift As String, _
                                  ByVal aEndDate As Date, _
                                  ByVal aEndShift As String) As gSzdFdTonsType

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetSzdFdTonsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim SzdFdDynaset As OraDynaset
        Dim RecordCount As Integer

        Dim ThisEqptType As String
        Dim ThisEqptName As String
        Dim ThisOperHrs As Single
        Dim ThisFdTons As Double

        Dim SumFneFdTons As Double
        Dim SumFneHrs As Single
        Dim SumCrsFdTons As Double
        Dim SumCrsHrs As Single
        Dim NumFneCircs As Integer
        Dim NumCrsCircs As Integer

        SumFneFdTons = 0
        SumFneHrs = 0
        SumCrsFdTons = 0
        SumCrsHrs = 0
        NumFneCircs = 0
        NumCrsCircs = 0

        'Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pBeginDate", aBeginDate, ORAPARM_INPUT)
        params("pBeginDate").serverType = ORATYPE_DATE

    params.Add("pBeginShift", StrConv(aBeginShift, vbUpperCase), ORAPARM_INPUT)
        params("pBeginShift").serverType = ORATYPE_VARCHAR2

    params.Add("pEndDate", aEndDate, ORAPARM_INPUT)
        params("pEndDate").serverType = ORATYPE_DATE

    params.Add("pEndShift", StrConv(aEndShift, vbUpperCase), ORAPARM_INPUT)
        params("pEndShift").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'Procedure get_szdfd_ton_data
        'pMineName           IN     VARCHAR2,
        'pBeginDate          IN     DATE,
        'pBeginShift         IN     VARCHAR2,
        'pEndDate            IN     DATE,
        'pEndShift           IN     VARCHAR2,
        'pResult             IN OUT c_sizedfeedtons)
        ' Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_floatplant2.get_szdfd_ton_data(:pMineName," + _
                  ":pBeginDate, :pBeginShift, :pEndDate, :pEndShift, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        SzdFdDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = SzdFdDynaset.RecordCount

        SzdFdDynaset.MoveFirst()
        Do While Not SzdFdDynaset.EOF
            ThisEqptType = SzdFdDynaset.Fields("eqpt_type_name").Value
            ThisEqptName = SzdFdDynaset.Fields("eqpt_name").Value
            ThisOperHrs = SzdFdDynaset.Fields("operating_hours").Value
            ThisFdTons = SzdFdDynaset.Fields("sized_feed_tons").Value

            If ThisEqptType = "Float plant rougher circuit" And _
                InStr(StrConv(ThisEqptName, vbUpperCase), "FINE ROUGHER") <> 0 Then
                SumFneFdTons = SumFneFdTons + ThisFdTons
                SumFneHrs = SumFneHrs + ThisOperHrs
                NumFneCircs = NumFneCircs + 1
            End If
            If ThisEqptType = "Float plant rougher circuit" And _
                InStr(StrConv(ThisEqptName, vbUpperCase), "COARSE ROUGHER") <> 0 Then
                SumCrsFdTons = SumCrsFdTons + ThisFdTons
                SumCrsHrs = SumCrsHrs + ThisOperHrs
                NumCrsCircs = NumCrsCircs + 1
            End If

            SzdFdDynaset.MoveNext()
        Loop

        With gGetSzdFdTons
            .FneFdTons = SumFneFdTons
            .FneFdHrs = SumFneHrs
            .CrsFdTons = SumCrsFdTons
            .CrsFdHrs = SumCrsHrs
            .NumFneCircs = NumFneCircs
            .NumCrsCircs = NumCrsCircs

            If .FneFdHrs <> 0 Then
                .FneFdTph = Round(.FneFdTons / .FneFdHrs, 0)
            Else
                .FneFdTph = 0
            End If

            If .CrsFdHrs <> 0 Then
                .CrsFdTph = Round(.CrsFdTons / .CrsFdHrs, 0)
            Else
                .CrsFdTph = 0
            End If
        End With

        SzdFdDynaset.Close()

        Exit Function

gGetSzdFdTonsError:
        MsgBox("Error sized feed data." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Sized Feed Data Get Error")

        On Error Resume Next
        ClearParams(params)
        SzdFdDynaset.Close()
    End Function

    Public Function gGet2NumAvg(ByVal aNum1 As Double, _
                                ByVal aNum2 As Double, _
                                ByVal aRound As Integer) As Double

        '**********************************************************************
        ' Special average for two numbers.
        '
        '
        '**********************************************************************

        If aNum1 = 0 And aNum2 <> 0 Then
            gGet2NumAvg = Round(aNum2, aRound)
            Exit Function
        End If
        If aNum1 <> 0 And aNum2 = 0 Then
            gGet2NumAvg = Round(aNum1, aRound)
            Exit Function
        End If

        gGet2NumAvg = Round(((aNum1 + aNum2) / 2), aRound)
    End Function

    Public Function gGetDlNum(ByVal aDlName As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim PoundPos As Integer

        'Dragline names are always like "Dragline #10", "Dragline #11", etc.

        gGetDlNum = ""
        PoundPos = InStr(aDlName, "#")

        If PoundPos = 0 Then
            Exit Function
        End If

        gGetDlNum = Trim(Mid(aDlName, PoundPos + 1))
    End Function

    Public Function gGetDlNameFromNum(ByVal aDlNum As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Dragline names are always like "Dragline #10", "Dragline #11", etc.

        gGetDlNameFromNum = "Dragline #" & Trim(aDlNum)
    End Function

    Public Function gGetMonthNum(ByVal aMonthAbbrv As String) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        aMonthAbbrv = Trim(StrConv(aMonthAbbrv, vbUpperCase))

        Select Case aMonthAbbrv
            Case Is = "JAN"
                gGetMonthNum = 1
            Case Is = "FEB"
                gGetMonthNum = 2
            Case Is = "MAR"
                gGetMonthNum = 3
            Case Is = "APR"
                gGetMonthNum = 4
            Case Is = "MAY"
                gGetMonthNum = 5
            Case Is = "JUN"
                gGetMonthNum = 6
            Case Is = "JUL"
                gGetMonthNum = 7
            Case Is = "AUG"
                gGetMonthNum = 8
            Case Is = "SEP"
                gGetMonthNum = 9
            Case Is = "OCT"
                gGetMonthNum = 10
            Case Is = "NOV"
                gGetMonthNum = 11
            Case Is = "DEC"
                gGetMonthNum = 12
            Case Else
                gGetMonthNum = 0
        End Select
    End Function

    Public Function gGetMonthNum2(ByVal aMonth As String) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        aMonth = Trim(StrConv(aMonth, vbUpperCase))

        Select Case aMonth
            Case Is = "JANUARY"
                gGetMonthNum2 = 1
            Case Is = "FEBRUARY"
                gGetMonthNum2 = 2
            Case Is = "MARCH"
                gGetMonthNum2 = 3
            Case Is = "APRIL"
                gGetMonthNum2 = 4
            Case Is = "MAY"
                gGetMonthNum2 = 5
            Case Is = "JUNE"
                gGetMonthNum2 = 6
            Case Is = "JULY"
                gGetMonthNum2 = 7
            Case Is = "AUGUST"
                gGetMonthNum2 = 8
            Case Is = "SEPTEMBER"
                gGetMonthNum2 = 9
            Case Is = "OCTOBER"
                gGetMonthNum2 = 10
            Case Is = "NOVEMBER"
                gGetMonthNum2 = 11
            Case Is = "DECEMBER"
                gGetMonthNum2 = 12
            Case Else
                gGetMonthNum2 = 0
        End Select
    End Function

    Public Function gGetYearsBackDate(ByVal aDate As Date, _
                                      ByVal aYrsBack As Integer, _
                                      ByVal SetToFirst As Boolean) As Date

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim Day As Integer
        Dim Month As Integer
        Dim Year As Integer

        Day = DatePart("d", aDate)
        Month = DatePart("m", aDate)
        Year = DatePart("yyyy", aDate)

        Year = Year - aYrsBack

        If SetToFirst = True Then
            Day = 1
        End If

        gGetYearsBackDate = CDate(Format(Month, "0#") & "/" & Format(Day, "0#") & _
                            "/" & Format(Year, "####"))
    End Function

    Public Function gConvertToShiftDate2(ByVal aMineName As String, _
                                         ByVal aDate As Date, _
                                         ByVal aShift As String) As Date

        '**********************************************************************
        '   This routine converts a date and shift to a shiftdate using the
        '   mois.check_shifts Oracle function.
        '
        '   This function is not tied to the 2 shift ("Day", "Night") scheme.
        '   It will work for 3 shift mines also.
        '**********************************************************************

        Dim SQLString As String
        'Set up SQL command.

        aShift = StrConv(aShift, vbUpperCase)
        SQLString = _
             "Begin :pShiftDate := mois.check_shifts" & _
             "(:pMineName, " & _
             ":pDate, " & _
             ":pShift);  " & _
             "End;"

        'Execute command with parameter information passed in as an array
        'of arrays.
        Dim arA1() As Object = {"pMineName", aMineName, ORAPARM_INPUT, ORATYPE_VARCHAR2} ') As Array
        Dim arA2() As Object = {"pDate", aDate.ToString, ORAPARM_INPUT, ORATYPE_DATE} ') As Array
        Dim arA3() As Object = {"pShift", aShift, ORAPARM_INPUT, ORATYPE_VARCHAR2} ') As Array
        Dim arA4() As Object = {"pShiftDate", 0, ORAPARM_OUTPUT, ORATYPE_DATE} ') As Array
        gConvertToShiftDate2 = RunSPReturnDate _
            ( _
                SQLString, _
                arA1, _
                arA2, _
                arA3, _
                arA4 _
            )

        'Array("pMineName", aMineName, ORAPARM_INPUT, ORATYPE_VARCHAR2), _
        'Array("pDate", aDate, ORAPARM_INPUT, ORATYPE_DATE), _
        'Array("pShift", aShift, ORAPARM_INPUT, ORATYPE_VARCHAR2), _
        'Array("pShiftDate", 0, ORAPARM_OUTPUT, ORATYPE_DATE) _

    End Function

    Public Function gLimsProdToMoisProd(ByVal aLimsProd As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Select Case aLimsProd
            Case Is = "FC"
                gLimsProdToMoisProd = "Fine concentrate"

            Case Is = "CC"
                gLimsProdToMoisProd = "Coarse concentrate"

            Case Is = "UC"
                gLimsProdToMoisProd = "Ultra-coarse concentrate"

            Case Is = "PB"
                gLimsProdToMoisProd = "Pebble"

            Case Is = "CN"
                gLimsProdToMoisProd = "Concentrate"

            Case Is = "IP"
                gLimsProdToMoisProd = "IP"

            Case Else
                gLimsProdToMoisProd = "??"
        End Select
    End Function

    Public Function gMoisProdToLimsProd(ByVal aMoisProd As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Select Case aMoisProd
            Case Is = "Fine concentrate"
                gMoisProdToLimsProd = "FC"

            Case Is = "Coarse concentrate"
                gMoisProdToLimsProd = "CC"

            Case Is = "Ultra-coarse concentrate"
                gMoisProdToLimsProd = "UC"

            Case Is = "Pebble"
                gMoisProdToLimsProd = "PB"

            Case Is = "Concentrate"
                gMoisProdToLimsProd = "CN"

            Case Is = "IP"
                gMoisProdToLimsProd = "IP"

            Case Else
                gMoisProdToLimsProd = "??"
        End Select
    End Function

    Public Function gGetMineNameShort(ByVal aMineNameLong As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetMineNameShortError

        Dim MineInfoDynaset As OraDynaset
        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        Dim ThisMineNameShort As String

        gGetMineNameShort = "??"

        'PROCEDURE get_one_mine_info
        'pMine   IN     VARCHAR2,
        'pResult IN OUT c_mineinfo)
        'Set 
        params = gDBParams

    params.Add("pMine", aMineNameLong, ORAPARM_INPUT)
        params("pMine").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_one_mine_info(:pMine, " + _
                 ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        MineInfoDynaset = params("pResult").Value
        ClearParams(params)

        If Not IsDBNull(MineInfoDynaset.Fields("mine_abbrv").Value) Then
            gGetMineNameShort = MineInfoDynaset.Fields("mine_abbrv").Value
        Else
            gGetMineNameShort = "??"
        End If

        MineInfoDynaset.Close()
        Exit Function

gGetMineNameShortError:
        MsgBox("Error getting mine abbreviation." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Mine Abbreviation Error")

        On Error Resume Next
        gGetMineNameShort = "??"
        ClearParams(params)
        MineInfoDynaset.Close()
    End Function

    Public Function gGetMassBalanceMethod(ByVal aMineNameLong As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        '02/25/2008, lss
        'Was:
        'Select Case aMineNameLong
        '    Case Is = "South Fort Meade"
        '        gGetMassBalanceMethod = "Normal"
        '    Case Is = "Hookers Prairie"
        '        gGetMassBalanceMethod = "Normal"
        '    Case Is = "South Fort Meade"
        '        gGetMassBalanceMethod = "Normal"
        '    Case Is = "Four Corners"
        '        gGetMassBalanceMethod = "Special#1"
        '    Case Else
        '        gGetMassBalanceMethod = "??"
        'End Select

        gGetMassBalanceMethod = gGetMiscMineGlobals(aMineNameLong, "mass balance mode")
    End Function

    Public Function gStrCharCount(ByVal aString As String, _
                                  ByVal aChar As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim CharIdx As Integer
        Dim CharCnt As Integer

        gStrCharCount = 0
        CharCnt = 0

        For CharIdx = 1 To Len(aString)
            If Mid(aString, CharIdx, 1) = aChar Then
                CharCnt = CharCnt + 1
            End If
        Next CharIdx

        gStrCharCount = CharCnt
    End Function

    Public Function gGetFirstOfNextWeek(ByVal aDate As Date) As Date

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'We will add 7 days to whatever day is passed and that will put us
        'into the next week somewhere.  We will then just get the first of
        'that week.
        gGetFirstOfNextWeek = CDate(aDate.AddDays(7)).AddDays(-DatePart("w", aDate, vbMonday) + 1)
    End Function

    Public Function gGetCircBplSpec(ByVal aMineName As String, _
                                    ByVal aEqptTypeName As String, _
                                    ByVal aEqptName As String, _
                                    ByVal aDate As Date, _
                                    ByVal aShift As String, _
                                    ByVal aMeasureName As String) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetCircBplSpecError

        Dim MineInfoDynaset As OraDynaset
        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim BplValue As Single

        'Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptTypeName", aEqptTypeName, ORAPARM_INPUT)
        params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptName", aEqptName, ORAPARM_INPUT)
        params("pEqptName").serverType = ORATYPE_VARCHAR2

    params.Add("pDate", aDate, ORAPARM_INPUT)
        params("pDate").serverType = ORATYPE_DATE

    params.Add("pShift", StrConv(aShift, vbUpperCase), ORAPARM_INPUT)
        params("pShift").serverType = ORATYPE_VARCHAR2

    params.Add("pMeasureName", aMeasureName, ORAPARM_INPUT)
        params("pMeasureName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", "", ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_VARCHAR2

        'Procedure get_circ_bpl_spec
        'pMineName     IN     VARCHAR2,
        'pEqptTypeName IN     DATE,
        'pEqptName     IN     VARCHAR2,
        'pDate         IN     DATE,
        'pShift        IN     VARCHAR2,
        'pMeasureName  IN     VARCHAR2,
        'pResult       IN OUT VARCHAR2)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_floatplant2.get_circ_bpl_spec(:pMineName, " + _
                 ":pEqptTypeName, :pEqptName, :pDate, :pShift, " + _
                 ":pMeasureName, :pResult);end;", ORASQL_FAILEXEC)

        If Not IsDBNull(params("pResult").Value) Then
            gGetCircBplSpec = Val(params("pResult").Value)
        Else
            gGetCircBplSpec = -99
        End If
        ClearParams(params)

        Exit Function

gGetCircBplSpecError:
        MsgBox("Error getting circuit BPL." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Circuit BPL Error")

        On Error Resume Next
        gGetCircBplSpec = -99
        ClearParams(params)
    End Function

    Public Function gHasMultifosPotential(ByVal aActBpl As Single, _
                                          ByVal aActAl As Single, _
                                          ByVal aActIns As Single, _
                                          ByVal aDesiredIns As Single, _
                                          ByVal aMinBpl As Single, _
                                          ByVal aMaxAl As Single) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim AdjBpl As Single
        Dim AdjAl As Single

        If 100 - aActIns > 0 Then
            AdjBpl = aActBpl * (100 - aDesiredIns) / (100 - aActIns)
        Else
            AdjBpl = 0
        End If

        If aActBpl <> 0 Then
            AdjAl = aActAl * (AdjBpl / aActBpl)
        Else
            AdjAl = 0
        End If

        If AdjBpl > aMinBpl And AdjAl < aMaxAl Then
            gHasMultifosPotential = True
        Else
            gHasMultifosPotential = False
        End If
    End Function

    Public Function gGetPeriodicEqptMsrAvg2(ByVal aMineName As String, _
                                            ByVal aBeginDate As Date, _
                                            ByVal aBeginShift As String, _
                                            ByVal aEndDate As Date, _
                                            ByVal aEndShift As String, _
                                            ByVal aEqptTypeName As String, _
                                            ByVal aEqptName As String, _
                                            ByVal aMeasureName As String) As Double

        '**********************************************************************
        ' This function calls get_avg_eqpt_prdc_msrmnt which only gets the
        ' average of the periodic equipment measurements that have been
        ' assigned during the timeframe in question.
        '**********************************************************************

        On Error GoTo gPeriodicEqptMsrAvgError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim PeriodicEqptMsrmntDynaset As OraDynaset
        Dim RecordCount As Integer

        'Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pBeginDate", aBeginDate, ORAPARM_INPUT)
        params("pBeginDate").serverType = ORATYPE_DATE

    params.Add("pBeginShift", StrConv(aBeginShift, vbUpperCase), ORAPARM_INPUT)
        params("pBeginShift").serverType = ORATYPE_VARCHAR2

    params.Add("pEndDate", aEndDate, ORAPARM_INPUT)
        params("pEndDate").serverType = ORATYPE_DATE

    params.Add("pEndShift", StrConv(aEndShift, vbUpperCase), ORAPARM_INPUT)
        params("pEndShift").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptTypeName", aEqptTypeName, ORAPARM_INPUT)
        params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptName", aEqptName, ORAPARM_INPUT)
        params("pEqptName").serverType = ORATYPE_VARCHAR2

    params.Add("pMeasureName", aMeasureName, ORAPARM_INPUT)
        params("pMeasureName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'Procedure get_avg_eqpt_prdc_msrmnt
        'pMineName           IN     VARCHAR2,
        'pBeginDate          IN     DATE,
        'pBeginShift         IN     VARCHAR2,
        'pEndDate            IN     DATE,
        'pEndShift           IN     VARCHAR2,
        'pEqptTypeName       IN     VARCHAR2,
        'pEqptName           IN     VARCHAR2,
        'pMeasureName        IN     VARCHAR2,
        'pResult             IN OUT c_currprdcmsrmnts)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prdcmsrmnt.get_avg_eqpt_prdc_msrmnt(:pMineName, " & _
                                         ":pBeginDate, :pBeginShift, :pEndDate, :pEndShift, " & _
                                         ":pEqptTypeName, :pEqptName, " & _
                                         ":pMeasureName, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        PeriodicEqptMsrmntDynaset = params("pResult").Value
        ClearParams(params)
        RecordCount = PeriodicEqptMsrmntDynaset.RecordCount

        If RecordCount = 1 Then
            PeriodicEqptMsrmntDynaset.MoveFirst()
            gGetPeriodicEqptMsrAvg2 = PeriodicEqptMsrmntDynaset.Fields("avg_val").Value
        Else
            gGetPeriodicEqptMsrAvg2 = 0
        End If

        PeriodicEqptMsrmntDynaset.Close()

        Exit Function

gPeriodicEqptMsrAvgError:
        MsgBox("Error getting average periodic equipment measurement." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Average Periodic Equipment Measurement Get Error")

        On Error Resume Next
        ClearParams(params)
        gGetPeriodicEqptMsrAvg2 = 0
        PeriodicEqptMsrmntDynaset.Close()
    End Function

    Public Function gGetPeriodicEqptMsrAvg(ByVal aMineName As String, _
                                           ByVal aBeginDate As Date, _
                                           ByVal aBeginShift As String, _
                                           ByVal aEndDate As Date, _
                                           ByVal aEndShift As String, _
                                           ByVal aEqptTypeName As String, _
                                           ByVal aEqptName As String, _
                                           ByVal aMeasureName As String) As Double

        '**********************************************************************
        ' This function calls get_shift_avg_epm which determines the periodic
        ' equipment value for each shift in a time frame and then averages
        ' them.
        '**********************************************************************

        On Error GoTo gPeriodicEqptMsrAvgError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer
        Dim NumMineShifts As Integer

        NumMineShifts = gGetNumShifts2(aMineName, aBeginDate)

        ' Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pBeginDate", aBeginDate, ORAPARM_INPUT)
        params("pBeginDate").serverType = ORATYPE_DATE

    params.Add("pBeginShift", StrConv(aBeginShift, vbUpperCase), ORAPARM_INPUT)
        params("pBeginShift").serverType = ORATYPE_VARCHAR2

    params.Add("pEndDate", aEndDate, ORAPARM_INPUT)
        params("pEndDate").serverType = ORATYPE_DATE

    params.Add("pEndShift", StrConv(aEndShift, vbUpperCase), ORAPARM_INPUT)
        params("pEndShift").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptTypeName", aEqptTypeName, ORAPARM_INPUT)
        params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptName", aEqptName, ORAPARM_INPUT)
        params("pEqptName").serverType = ORATYPE_VARCHAR2

    params.Add("pMeasureName", aMeasureName, ORAPARM_INPUT)
        params("pMeasureName").serverType = ORATYPE_VARCHAR2

    params.Add("pShiftMax", NumMineShifts, ORAPARM_INPUT)
        params("pShiftMax").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        'Procedure get_shift_avg_epm
        'pMineName           IN     VARCHAR2,
        'pBeginDate          IN     DATE,
        'pBeginShift         IN     VARCHAR2,
        'pEndDate            IN     DATE,
        'pEndShift           IN     VARCHAR2,
        'pEqptTypeName       IN     VARCHAR2,
        'pEqptName           IN     VARCHAR2,
        'pMeasureName        IN     VARCHAR2,
        'pShiftMax           IN     NUMBER,
        'pResult             IN OUT NUMBER)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prdcmsrmnt.get_shift_avg_epm(:pMineName, " & _
                                         ":pBeginDate, :pBeginShift, :pEndDate, :pEndShift, " & _
                                         ":pEqptTypeName, :pEqptName, " & _
                                         ":pMeasureName, :pShiftMax, :pResult);end;", ORASQL_FAILEXEC)
        gGetPeriodicEqptMsrAvg = Round(params("pResult").Value, 5)
        ClearParams(params)

        Exit Function

gPeriodicEqptMsrAvgError:
        MsgBox("Error getting average periodic equipment measurement." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Average Periodic Equipment Measurement Get Error")

        On Error Resume Next
        gGetPeriodicEqptMsrAvg = 0
        On Error Resume Next
        ClearParams(params)
    End Function

    Public Function gGetPeriodicEqptMsrAvg3(ByVal aMineName As String, _
                                            ByVal aBeginDate As Date, _
                                            ByVal aBeginShift As String, _
                                            ByVal aEndDate As Date, _
                                            ByVal aEndShift As String, _
                                            ByVal aEqptTypeName As String, _
                                            ByVal aEqptName As String, _
                                            ByVal aMeasureName As String) As Double

        '**********************************************************************
        ' This function calls get_shift_avg_epm which determines the periodic
        ' equipment value for each shift in a time frame and then averages
        ' them.
        '**********************************************************************

        On Error GoTo gPeriodicEqptMsrAvg3Error

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer

        'Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pBeginDate", aBeginDate, ORAPARM_INPUT)
        params("pBeginDate").serverType = ORATYPE_DATE

    params.Add("pBeginShift", StrConv(aBeginShift, vbUpperCase), ORAPARM_INPUT)
        params("pBeginShift").serverType = ORATYPE_VARCHAR2

    params.Add("pEndDate", aEndDate, ORAPARM_INPUT)
        params("pEndDate").serverType = ORATYPE_DATE

    params.Add("pEndShift", StrConv(aEndShift, vbUpperCase), ORAPARM_INPUT)
        params("pEndShift").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptTypeName", aEqptTypeName, ORAPARM_INPUT)
        params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptName", aEqptName, ORAPARM_INPUT)
        params("pEqptName").serverType = ORATYPE_VARCHAR2

    params.Add("pMeasureName", aMeasureName, ORAPARM_INPUT)
        params("pMeasureName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        'Procedure get_shift_avg_epm2
        'pMineName           IN     VARCHAR2,
        'pBeginDate          IN     DATE,
        'pBeginShift         IN     VARCHAR2,
        'pEndDate            IN     DATE,
        'pEndShift           IN     VARCHAR2,
        'pEqptTypeName       IN     VARCHAR2,
        'pEqptName           IN     VARCHAR2,
        'pMeasureName        IN     VARCHAR2,
        'pResult             IN OUT NUMBER)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prdcmsrmnt.get_shift_avg_epm2(:pMineName, " & _
                                         ":pBeginDate, :pBeginShift, :pEndDate, :pEndShift, " & _
                                         ":pEqptTypeName, :pEqptName, " & _
                                         ":pMeasureName, :pResult);end;", ORASQL_FAILEXEC)
        gGetPeriodicEqptMsrAvg3 = Round(params("pResult").Value, 5)
        ClearParams(params)

        Exit Function

gPeriodicEqptMsrAvg3Error:
        MsgBox("Error getting average periodic equipment measurement." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Average Periodic Equipment Measurement Get Error")

        On Error Resume Next
        gGetPeriodicEqptMsrAvg3 = 0
        On Error Resume Next
        ClearParams(params)
    End Function

    Public Function gGetModifiedHour(ByVal aTime As String, _
                                     ByVal aNumShifts As Integer) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ThisHour As Integer

        gGetModifiedHour = 0

        If aNumShifts = 2 Then
            'If hour is >=0 and <=6 then add 24 to it.
            ThisHour = Val(Mid(aTime, 1, 2))

            If ThisHour >= 0 And ThisHour <= 6 Then
                gGetModifiedHour = ThisHour + 24
            Else
                gGetModifiedHour = ThisHour
            End If
        End If
        If aNumShifts = 3 Then
            'If hour is >=0 and <=7 then add 24 to it.
            ThisHour = Val(Mid(aTime, 1, 2))

            If ThisHour >= 0 And ThisHour <= 7 Then
                gGetModifiedHour = ThisHour + 24
            Else
                gGetModifiedHour = ThisHour
            End If
        End If
    End Function

    Public Function gGetInsAdjBpl(ByVal aActualBpl As Single, _
                                  ByVal aDesiredIns As Single, _
                                  ByVal aActualIns As Single, _
                                  ByVal aRoundVal As Integer) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        If aActualIns <> 0 Then
            gGetInsAdjBpl = Round(aActualBpl * (100 - aDesiredIns) / _
                            (100 - aActualIns), aRoundVal)
        Else
            gGetInsAdjBpl = 0
        End If
    End Function

    Public Function gGetInsAdjAl(ByVal aActualBpl As Single, _
                                 ByVal aDesiredIns As Single, _
                                 ByVal aActualIns As Single, _
                                 ByVal aActualAl As Single, _
                                 ByVal aRoundVal As Integer) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim AdjBpl As Single

        AdjBpl = gGetInsAdjBpl(aActualBpl, aDesiredIns, aActualIns, 2)

        If aActualIns <> 0 Then
            gGetInsAdjAl = Round(aActualAl * (AdjBpl / aActualBpl), aRoundVal)
        Else
            gGetInsAdjAl = 0
        End If
    End Function

    Public Function gShiftGreater(ByVal aShift1 As String, _
                                  ByVal aShift2 As String, _
                                  ByVal aNumShifts As Integer) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ShiftNum1 As Integer
        Dim ShiftNum2 As Integer

        'Determine if aShift1 is "greater" than aShift2 ie. aShift1 is later
        'than aShift2.

        'Assume that:
        'aNumShifts = 2 or 3
        'If aNumShifts = 2 then the shifts are "Day" and "Night"
        'If aNumShifts = 3 then the shifts are "1ST", "2ND", and "3RD"

        aShift1 = StrConv(aShift1, vbUpperCase)
        aShift2 = StrConv(aShift2, vbUpperCase)

        gShiftGreater = False

        If aNumShifts = 2 Then
            If aShift1 = "DAY" And aShift2 = "NIGHT" Then
                gShiftGreater = False
            End If
            If aShift1 = "NIGHT" And aShift2 = "DAY" Then
                gShiftGreater = True
            End If
        End If

        If aNumShifts = 3 Then
            ShiftNum1 = Val(Mid(aShift1, 1, 1))
            ShiftNum2 = Val(Mid(aShift2, 1, 1))

            If ShiftNum1 > ShiftNum2 Then
                gShiftGreater = True
            Else
                gShiftGreater = False
            End If
        End If
    End Function

    Public Function gTransSapphireProd(ByVal aMineName As String, _
                                       ByVal aSapphireProd As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        aSapphireProd = StrConv(aSapphireProd, vbUpperCase)
        gTransSapphireProd = "??"

        If gActiveMineNameLong = "Four Corners" Then
            Select Case aSapphireProd
                Case Is = "PEBBLE"
                    gTransSapphireProd = "PB"
                Case Is = "CONCENTRATE"
                    gTransSapphireProd = "CN"
                Case Is = "IP PRODUCT"
                    gTransSapphireProd = "IP"
            End Select
        End If

        If gActiveMineNameLong = "Hookers Prairie" Then
            Select Case aSapphireProd
                Case Is = "PEBBLE"
                    gTransSapphireProd = "PB"
                Case Is = "CONCENTRATE"
                    gTransSapphireProd = "CN"
            End Select
        End If

        If gActiveMineNameLong = "South Fort Meade" Then
            Select Case aSapphireProd
                Case Is = "PEBBLE"
                    gTransSapphireProd = "PB"
                Case Is = "COARSE CONCENTRATE"
                    gTransSapphireProd = "CC"
                Case Is = "FINE CONCENTRATE"
                    gTransSapphireProd = "FC"
                Case Is = "ULTRA COARSE CONC"
                    gTransSapphireProd = "UC"
            End Select
        End If

        If gActiveMineNameLong = "Hopewell" Then
            Select Case aSapphireProd
                Case Is = "PEBBLE"
                    gTransSapphireProd = "PB"
                Case Is = "CONCENTRATE"
                    gTransSapphireProd = "CN"
                Case Is = "IP PRODUCT"
                    gTransSapphireProd = "IP"
            End Select
        End If

        If gActiveMineNameLong = "Wingate" Then
            Select Case aSapphireProd
                Case Is = "PEBBLE"
                    gTransSapphireProd = "PB"
                Case Is = "CONCENTRATE"
                    gTransSapphireProd = "CN"
            End Select
        End If
    End Function

    Public Function gTransSapphireEqpt(ByVal aFacility As String, _
                                       ByVal aLocation As String, _
                                       ByVal aMaterial As String, _
                                       ByVal aAnalyte As String, _
                                       ByRef aMoisEqpt1 As String, _
                                       ByRef aMoisEqpt2 As String) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gTransSapphireEqptError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer
        Dim ThisEqptName As String
        Dim fLimsEqptDynaset As OraDynaset
        Dim EqptCount As Integer

        gTransSapphireEqpt = 0
        aMoisEqpt1 = ""
        aMoisEqpt2 = ""

        'Set 
        params = gDBParams

    params.Add("pFacility", aFacility, ORAPARM_INPUT)
        params("pFacility").serverType = ORATYPE_VARCHAR2

    params.Add("pLocation", aLocation, ORAPARM_INPUT)
        params("pLocation").serverType = ORATYPE_VARCHAR2

    params.Add("pMaterial", aMaterial, ORAPARM_INPUT)
        params("pMaterial").serverType = ORATYPE_VARCHAR2

    params.Add("pAnalyte", aAnalyte, ORAPARM_INPUT)
        params("pAnalyte").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE trans_lims_eqpt_to_mois
        'pFacility     IN     VARCHAR2,
        'pLocation     IN     VARCHAR2,
        'pMaterial     IN     VARCHAR2,
        'pAnalyte      IN     VARCHAR2,
        'pResult       IN OUT c_measures)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_interface.trans_lims_eqpt_to_mois(:pFacility, " & _
                                         ":pLocation, :pMaterial, :pAnalyte, " & _
                                         ":pResult);end;", ORASQL_FAILEXEC)
        'Set 
        fLimsEqptDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = fLimsEqptDynaset.RecordCount
        gTransSapphireEqpt = RecordCount
        EqptCount = 0

        If RecordCount = 1 Or RecordCount = 2 Then
            fLimsEqptDynaset.MoveFirst()
            Do While Not fLimsEqptDynaset.EOF
                EqptCount = EqptCount + 1
                If EqptCount = 1 Then
                    aMoisEqpt1 = fLimsEqptDynaset.Fields("eqpt_name").Value
                End If
                If EqptCount = 2 Then
                    aMoisEqpt2 = fLimsEqptDynaset.Fields("eqpt_name").Value
                End If

                fLimsEqptDynaset.MoveNext()
            Loop
        Else
            aMoisEqpt1 = "??"
            aMoisEqpt2 = "??"
        End If

        fLimsEqptDynaset.Close()

        Exit Function

gTransSapphireEqptError:
        MsgBox("Error translating Sapphire equipment." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Sapphire Equipment Translation Error")

        On Error Resume Next
        gTransSapphireEqpt = 0
        aMoisEqpt1 = "??"
        aMoisEqpt2 = "??"
        On Error Resume Next
        fLimsEqptDynaset.Close()
        On Error Resume Next
        ClearParams(params)
    End Function

    Public Function gGetMoisDateForExtDate(ByVal aMineName As String, _
                                           ByVal aDateTime As Date) As Date

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim FirstShiftStartTime As String
        Dim FirstShiftStartDate As Date
        Dim MidNiteDateTime As Date

        Dim GotShiftNames As Boolean
        Dim ShiftNames() As gShiftNamesType

        GotShiftNames = gGetShiftNames(aMineName, ShiftNames)

        FirstShiftStartTime = ShiftNames(1).BeginTime
        FirstShiftStartDate = CDate(Format(aDateTime, "MM/dd/yyyy") & " " & _
                              Format(FirstShiftStartTime, "hh:mm AMPM"))

        MidNiteDateTime = CDate(Format(aDateTime, "MM/dd/yyyy") & " " & _
                          "12:00:01 AM")

        If aDateTime >= MidNiteDateTime And aDateTime < FirstShiftStartDate Then
            gGetMoisDateForExtDate = CDate(Format(aDateTime, "MM/dd/yyyy")).AddDays(-1)
        Else
            gGetMoisDateForExtDate = CDate(Format(aDateTime, "MM/dd/yyyy"))
        End If
    End Function

    Public Function gGetExtDateForMoisDate(ByVal aMineName As String, _
                                           ByVal aDateTime As Date, _
                                           ByRef aShiftNames() As gShiftNamesType) As Date

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim FirstShiftStartTime As String
        Dim FirstShiftStartDate As Date
        Dim MidNiteDateTime As Date

        FirstShiftStartTime = aShiftNames(1).BeginTime
        FirstShiftStartDate = CDate(Format(aDateTime, "MM/dd/yyyy") & " " & _
                              Format(FirstShiftStartTime, "hh:mm AMPM"))

        MidNiteDateTime = CDate(Format(aDateTime, "MM/dd/yyyy") & " " & _
                          "12:00:01 AM")

        If aDateTime >= MidNiteDateTime And aDateTime < FirstShiftStartDate Then
            gGetExtDateForMoisDate = CDate(Format(aDateTime, "MM/dd/yyyy")).AddDays(1)
        Else
            gGetExtDateForMoisDate = CDate(Format(aDateTime, "MM/dd/yyyy"))
        End If
    End Function

    Public Function gAddEqptMsrmnt(ByVal aMineName As String, _
                                   ByVal aEqptTypeName As String, _
                                   ByVal aEqptName As String, _
                                   ByVal aDate As Date, _
                                   ByVal aShift As String, _
                                   ByVal aMeasureName As String, _
                                   ByVal aValue As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gAddEqptMsrmntError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim SaveStatus As Integer

        gAddEqptMsrmnt = False

        'Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptTypeName", aEqptTypeName, ORAPARM_INPUT)
        params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptName", aEqptName, ORAPARM_INPUT)
        params("pEqptName").serverType = ORATYPE_VARCHAR2

        'update_one_eqpt_msrmnt will form a shift date from aDate and aShift
    params.Add("pShiftDate", aDate, ORAPARM_INPUT)
        params("pShiftDate").serverType = ORATYPE_DATE

    params.Add("pShift", StrConv(aShift, vbUpperCase), ORAPARM_INPUT)
        params("pShift").serverType = ORATYPE_VARCHAR2

    params.Add("pMeasureName", aMeasureName, ORAPARM_INPUT)
        params("pMeasureName").serverType = ORATYPE_VARCHAR2

    params.Add("pValue", aValue, ORAPARM_INPUT)
        params("pValue").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        'PROCEDURE update_one_eqpt_msrmnt
        'pMineName         IN     VARCHAR2,
        'pEqptTypeName     IN     VARCHAR2,
        'pEqptName         IN     VARCHAR2,
        'pShiftDate        IN     DATE,
        'pShift            IN     VARCHAR2,
        'pMeasureName      IN     VARCHAR2,
        'pValue            IN     VARCHAR2,
        'pResult           IN OUT NUMBER)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities2.update_one_eqpt_msrmnt(:pMineName," + _
                  ":pEqptTypeName, :pEqptName, :pShiftDate, :pShift, :pMeasureName," + _
                  ":pValue, :pResult);end;", ORASQL_FAILEXEC)

        SaveStatus = params("pResult").Value
        ClearParams(params)
        gAddEqptMsrmnt = True

        Exit Function

gAddEqptMsrmntError:
        MsgBox("Error updating EQPT_MSRMNT." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Update Error")

        On Error Resume Next
        ClearParams(params)
        gAddEqptMsrmnt = False
    End Function

    Public Function gCboBoxChoiceIsThere(ByVal aCboBox As ComboBox, _
                                         ByVal aValue As String) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ItemIdx As Integer

        gCboBoxChoiceIsThere = False

        For ItemIdx = 0 To aCboBox.Items.Count - 1
            If aCboBox.Items(ItemIdx) = aValue Then
                gCboBoxChoiceIsThere = True
                'Exit Function
            End If
        Next ItemIdx
    End Function

    Public Function gGetShiftLengthForDate(ByVal aMineName As String, _
                                           ByVal aDate As Date) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetShiftLengthForDateError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim PeriodicEqptMsrmntDynaset As OraDynaset
        Dim RecordCount As Integer

        'This is a special correction that is needed -- "current"
        'dates have been entered as 12/31/9999 and 12/31/8888.
        'Need to compensate for the "old" 12/31/9999 ones.
        If aDate = #12/31/9999# Then
            aDate = #12/31/8888#
        End If

        'Need to get the shift length for this date based on data
        'in EQPT_PRDC_MSRMNT.

        'Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pDate", aDate, ORAPARM_INPUT)
        params("pDate").serverType = ORATYPE_DATE

    params.Add("pEqptTypeName", "Mine", ORAPARM_INPUT)
        params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptName", "Mine", ORAPARM_INPUT)
        params("pEqptName").serverType = ORATYPE_VARCHAR2

    params.Add("pMeasureName", "Shift length", ORAPARM_INPUT)
        params("pMeasureName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_eqpt_prdc_msrmnt_bydate
        'pMineName           IN     VARCHAR2,
        'pDate               IN     DATE,
        'pEqptTypeName       IN     VARCHAR2,
        'pEqptName           IN     VARCHAR2,
        'pMeasureName        IN     VARCHAR2,
        'pResult             IN OUT c_currprdcmsrmnts)
        ' Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_prdcmsrmnt.get_eqpt_prdc_msrmnt_bydate(:pMineName, " & _
                                         ":pDate, :pEqptTypeName, :pEqptName, " & _
                                         ":pMeasureName, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        PeriodicEqptMsrmntDynaset = params("pResult").Value
        ClearParams(params)
        RecordCount = PeriodicEqptMsrmntDynaset.RecordCount

        If RecordCount = 1 Then
            PeriodicEqptMsrmntDynaset.MoveFirst()
            gGetShiftLengthForDate = PeriodicEqptMsrmntDynaset.Fields("value").Value
        Else
            gGetShiftLengthForDate = 0
        End If

        PeriodicEqptMsrmntDynaset.Close()

        Exit Function

gGetShiftLengthForDateError:
        MsgBox("Error getting shift length for date." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Shift Length Get Error")

        On Error Resume Next
        ClearParams(params)
        PeriodicEqptMsrmntDynaset.Close()
    End Function

    Function gGetMineUserPermissions(ByVal aUserId As String, _
                                     ByVal aArea As String, _
                                     ByVal aMineName As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim UserPermissionDynaset As OraDynaset
        'Dim MaxRows As Integer
        'Dim CurrentRow As Integer
        Dim RecordCount As Integer

        gGetMineUserPermissions = "None"

        'Load user permissions from Oracle table"

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        'Get all permissions for user
        'Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pUserID", StrConv(aUserId, vbUpperCase), ORAPARM_INPUT)
        params("pUserID").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_user_permissions
        'pMineName        IN     VARCHAR2,
        'pUserID          IN     VARCHAR2,
        'pResult          IN OUT c_users)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_user_permissions(:pMineName," + _
                  ":pUserID, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        UserPermissionDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = UserPermissionDynaset.RecordCount

        ' 1) Field
        ' 2) Washer
        ' 3) Sizing
        ' 4) Float plant
        ' 5) Miscellaneous
        ' 6) Analysis
        ' 7) Shipping
        ' 8) Production
        ' 9) Reagent
        '10) Prospect
        '11) Survey
        '12) Mine plan
        '13) Utilities
        '14) Pump yardages
        '15) Train shipping
        '16) Dragline cables
        '17) Raw prospect chem lab
        '18) Raw prospect met lab
        '19) Utility Operator Report
        '20) DL Inspection Report
        '21) Pump Inspection Report
        '22) Piezometers
        '23) Reclamation activity
        '24) Web reports
        '25) Water samples
        '26) Pipe thickness
        '27) Decision grid
        '28) Washer Shift Report
        '29) Float Plant Shift Report
        '30) Sizing Shift Report
        '31) Reagent Shift Report
        '32) Shipping Shift Report
        '33) Absentees
        '34) Safety meetings
        '35) Pump Pack Shift Report

        UserPermissionDynaset.MoveFirst()
        Do While Not UserPermissionDynaset.EOF
            Select Case UserPermissionDynaset.Fields("permission_type_name").Value
                'Field
                Case Is = "Field input read"
                    If aArea = "Field" Then
                        gGetMineUserPermissions = "Read"
                    End If
                Case Is = "Field input write"
                    If aArea = "Field" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Washer
                Case Is = "Washer input read"
                    If aArea = "Field" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Washer input write"
                    If aArea = "Field" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Sizing
                Case Is = "Sizing input read"
                    If aArea = "Sizing" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Sizing input write"
                    If aArea = "Sizing" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Float plant
                Case Is = "Float plant input read"
                    If aArea = "Float plant" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Float plant input write"
                    If aArea = "Float plant" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Miscellaneous
                Case Is = "Misc input read"
                    If aArea = "Miscellaneous" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Misc input write"
                    If aArea = "Miscellaneous" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Analysis
                Case Is = "Analysis input read"
                    If aArea = "Analysis" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Analysis input write"
                    If aArea = "Analysis" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Shipping
                Case Is = "Shipping input read"
                    If aArea = "Shipping" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Shipping input write"
                    If aArea = "Shipping" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Production
                Case Is = "Production input read"
                    If aArea = "Production" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Production input write"
                    If aArea = "Production" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Reagent
                Case Is = "Reagent input read"
                    If aArea = "Reagent" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Reagent input write"
                    If aArea = "Reagent" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Prospect"
                Case Is = "Prospect read"
                    If aArea = "Prospect" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Prospect write"
                    If aArea = "Prospect" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Survey
                Case Is = "Survey read"
                    If aArea = "Survey" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Survey write"
                    If aArea = "Survey" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Mine plan
                Case Is = "Mine plan read"
                    If aArea = "Mine plan" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Mine plan write"
                    If aArea = "Mine plan" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Utilities
                Case Is = "Utilities read"
                    If aArea = "Utilities" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Utilities write"
                    If aArea = "Utilities" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Pump yardages
                Case Is = "Pump yardages read"
                    If aArea = "Pump yardages" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Pump yardages write"
                    If aArea = "Pump yardages" Then
                        gGetMineUserPermissions = "Write"
                    End If

                Case Is = "Pump yardages setup"
                    If aArea = "Pump yardages" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                    'Train shipping
                Case Is = "Train shipping read"
                    If aArea = "Train shipping" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Train shipping write"
                    If aArea = "Train shipping" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Dragline cables
                Case Is = "Dragline cables read"
                    If aArea = "Dragline cables" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Dragline cables write"
                    If aArea = "Dragline cables" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Raw prospect chem lab
                Case Is = "Raw prospect chem lab read"
                    If aArea = "Raw prospect chem lab" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Raw prospect chem lab write"
                    If aArea = "Raw prospect chem lab" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Raw prospect met lab
                Case Is = "Raw prospect met lab read"
                    If aArea = "Raw prospect met lab" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Raw prospect met lab write"
                    If aArea = "Raw prospect met lab" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Maps
                Case Is = "Maps read"
                    If aArea = "Maps" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Maps write"
                    If aArea = "Maps" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Utility Operator Report"
                Case Is = "Utility Operator Report read"
                    If aArea = "Utility Operator Report" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Utility Operator Report write"
                    If aArea = "Utility Operator Report" Then
                        gGetMineUserPermissions = "Write"
                    End If

                Case Is = "Utility Operator Report setup"
                    If aArea = "Utility Operator Report" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                    'DL Inspection Report
                Case Is = "DL Inspection Report read"
                    If aArea = "DL Inspection Report" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "DL Inspection Report write"
                    If aArea = "DL Inspection Report" Then
                        gGetMineUserPermissions = "Write"
                    End If

                Case Is = "DL Inspection Report setup"
                    If aArea = "DL Inspection Report" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                    'Pump Inspection Report
                Case Is = "Pump Inspection Report read"
                    If aArea = "Pump Inspection Report" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Pump Inspection Report write"
                    If aArea = "Pump Inspection Report" Then
                        gGetMineUserPermissions = "Write"
                    End If

                Case Is = "Pump Inspection Report setup"
                    If aArea = "Pump Inspection Report" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                    'Piezometers
                Case Is = "Piezometers read"
                    If aArea = "Piezometers" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Piezometers write"
                    If aArea = "Piezometers" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Reclamation
                Case Is = "Reclamation activity read"
                    If aArea = "Reclamation" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Reclamation activity write"
                    If aArea = "Rewclamation" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Web reports
                Case Is = "Web reports read"
                    If aArea = "Web reports" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Web reports write"
                    If aArea = "Web reports" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Water samples
                Case Is = "Water samples read"
                    If aArea = "Water samples" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Water samples write"
                    If aArea = "Water samples" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Pipe thickness
                Case Is = "Pipe thickness read"
                    If aArea = "Pipe thickness" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Pipe thickness write"
                    If aArea = "Pipe thickness" Then
                        gGetMineUserPermissions = "Write"
                    End If

                Case Is = "Pipe thickness setup"
                    If aArea = "Pipe thickness" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                    'Decision grid
                Case Is = "Decision grid read"
                    If aArea = "Decision grid" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Decision grid write"
                    If aArea = "Decision grid" Then
                        gGetMineUserPermissions = "Write"
                    End If

                    'Washer Shift Report
                Case Is = "Washer Shift Report read"
                    If aArea = "Washer Shift Report" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Washer Shift Report write"
                    If aArea = "Washer Shift Report" Then
                        gGetMineUserPermissions = "Write"
                    End If

                Case Is = "Washer Shift Report setup"
                    If aArea = "Washer Shift Report" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                    'Float Plant Shift Report
                Case Is = "Float Plant Shift Report read"
                    If aArea = "Float Plant Shift Report" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Float Plant Shift Report write"
                    If aArea = "Float Plant Shift Report" Then
                        gGetMineUserPermissions = "Write"
                    End If

                Case Is = "Float Plant Shift Report setup"
                    If aArea = "Float Plant Shift Report" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                    'Sizing Shift Report
                Case Is = "Sizing Shift Report read"
                    If aArea = "Sizing Shift Report" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Sizing Shift Report write"
                    If aArea = "Sizing Shift Report" Then
                        gGetMineUserPermissions = "Write"
                    End If

                Case Is = "Sizing Shift Report setup"
                    If aArea = "Sizing Shift Report" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                    'Reagent Shift Report
                Case Is = "Reagent Shift Report read"
                    If aArea = "Reagent Shift Report" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Reagent Shift Report write"
                    If aArea = "Reagent Shift Report" Then
                        gGetMineUserPermissions = "Write"
                    End If

                Case Is = "Reagent Shift Report setup"
                    If aArea = "Reagent Shift Report" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                    'Shipping Shift Report
                Case Is = "Shipping Shift Report read"
                    If aArea = "Shipping Shift Report" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Shipping Shift Report write"
                    If aArea = "Shipping Shift Report" Then
                        gGetMineUserPermissions = "Write"
                    End If

                Case Is = "Shipping Shift Report setup"
                    If aArea = "Shipping Shift Report" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                    'Absentees
                Case Is = "Absentees input"
                    If aArea = "Absentees" Then
                        gGetMineUserPermissions = "Input"
                    End If

                Case Is = "Absentees review"
                    If aArea = "Absentees" Then
                        gGetMineUserPermissions = "Review"
                    End If

                Case Is = "Absentees setup"
                    If aArea = "Absentees" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                Case Is = "Absentees read"
                    If aArea = "Absentees" Then
                        gGetMineUserPermissions = "Read"
                    End If

                    'Safety meetings
                Case Is = "Safety meetings read"
                    If aArea = "Safety meetings" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Safety meetings write"
                    If aArea = "Safety meetings" Then
                        gGetMineUserPermissions = "Write"
                    End If

                Case Is = "Safety meetings setup"
                    If aArea = "Safety meetings" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                    'Pump Pack Shift Report
                Case Is = "Pump Pack Shift Report read"
                    If aArea = "Pump Pack Shift Report" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Pump Pack Shift Report write"
                    If aArea = "Pump Pack Shift Report" Then
                        gGetMineUserPermissions = "Write"
                    End If

                Case Is = "Pump Pack Shift Report setup"
                    If aArea = "Pump Pack Shift Report" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                    'Inventory
                Case Is = "Inventory adjust"
                    If aArea = "Pump Pack Shift Report" Then
                        gGetMineUserPermissions = "True"
                    End If

                    'Raw prospect reduction
                Case Is = "Raw prospect reduction read"
                    If aArea = "Raw Prospect Reduction" Then
                        gGetMineUserPermissions = "Read"
                    End If

                Case Is = "Raw prospect reduction write"
                    If aArea = "Raw Prospect Reduction" Then
                        gGetMineUserPermissions = "Write"
                    End If

                Case Is = "Raw prospect reduction setup"
                    If aArea = "Raw Prospect Reduction" Then
                        gGetMineUserPermissions = "Setup"
                    End If

                Case Is = "Raw prospect reduction admin"
                    If aArea = "Raw Prospect Reduction" Then
                        gGetMineUserPermissions = "Admin"
                    End If
            End Select

            UserPermissionDynaset.MoveNext()
        Loop
    End Function

    Public Function gGetEqptYds(ByVal aMineName As String, _
                                ByVal aEqptTypeName As String, _
                                ByVal aEqptName As String, _
                                ByVal aBeginDate As Date, _
                                ByVal aBeginShift As String, _
                                ByVal aEndDate As Date, _
                                ByVal aEndShift As String, _
                                ByVal aYdType As String, _
                                ByVal aMtxYdsType As String) As Double

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetEqptYdsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer
        Dim EqptYdsDynaset As OraDynaset
        Dim MtxYds As Double
        Dim OvbYds As Double
        Dim MtxYdsType As String

        'aYdType will be "Matrix" or "Overburden"
        'aMtxYdsType will be "From bucket count" or "From matrix tons"

        If aMtxYdsType = "From bucket count" Then
            MtxYdsType = "Bucket count"
        End If
        If aMtxYdsType = "From matrix tons" Then
            MtxYdsType = "Matrix tons"
        End If

        MtxYds = 0
        OvbYds = 0
        gGetEqptYds = 0

        'Get actual equipment yards
        'Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptTypeName", aEqptTypeName, ORAPARM_INPUT)
        params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptName", aEqptName, ORAPARM_INPUT)
        params("pEqptName").serverType = ORATYPE_VARCHAR2

    params.Add("pBeginDate", aBeginDate, ORAPARM_INPUT)
        params("pBeginDate").serverType = ORATYPE_DATE

    params.Add("pBeginShift", aBeginShift, ORAPARM_INPUT)
        params("pBeginShift").serverType = ORATYPE_VARCHAR2

    params.Add("pEndDate", aEndDate, ORAPARM_INPUT)
        params("pEndDate").serverType = ORATYPE_DATE

    params.Add("pEndShift", aEndShift, ORAPARM_INPUT)
        params("pEndShift").serverType = ORATYPE_VARCHAR2

    params.Add("pMtxYdsType", MtxYdsType, ORAPARM_INPUT)
        params("pMtxYdsType").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_dl_yards_sel
        'pMineName           IN     VARCHAR2,
        'pEqptTypeName       IN     VARCHAR2,
        'pEqptName           IN     VARCHAR2,
        'pBeginDate          IN     DATE,
        'pBeginShift         IN     VARCHAR2,
        'pEndDate            IN     DATE,
        'pEndShift           IN     VARCHAR2,
        'pMtxYdsType         IN     VARCHAR2,
        'pResult             IN OUT c_ovbmtxyards)
        ' Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_field.get_dl_yards_sel(:pMineName," + _
                  ":pEqptTypeName, :pEqptName, " + _
                  ":pBeginDate, :pBeginShift, :pEndDate, :pEndShift, :pMtxYdsType, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        EqptYdsDynaset = params("pResult").Value
        ClearParams(params)

        RecordCount = EqptYdsDynaset.RecordCount

        If RecordCount = 1 Then
            EqptYdsDynaset.MoveFirst()
            If Not IsDBNull(EqptYdsDynaset.Fields("matrix_yards").Value) And _
                EqptYdsDynaset.Fields("matrix_yards").Value <> "" Then
                MtxYds = EqptYdsDynaset.Fields("matrix_yards").Value
            Else
                MtxYds = 0
            End If
            If Not IsDBNull(EqptYdsDynaset.Fields("overburden_yards").Value) And _
                EqptYdsDynaset.Fields("overburden_yards").Value <> "" Then
                OvbYds = EqptYdsDynaset.Fields("overburden_yards").Value
            Else
                OvbYds = 0
            End If

            If StrConv(aYdType, vbUpperCase) = "MATRIX" Then
                gGetEqptYds = MtxYds
            End If
            If StrConv(aYdType, vbUpperCase) = "OVERBURDEN" Then
                gGetEqptYds = OvbYds
            End If
        Else
            gGetEqptYds = 0
        End If

        EqptYdsDynaset.Close()

        Exit Function

gGetEqptYdsError:
        MsgBox("Error getting equipment yards." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Equipment Yards Error")

        On Error Resume Next
        ClearParams(params)
        On Error Resume Next
        EqptYdsDynaset.Close()
        On Error Resume Next
        gGetEqptYds = 0
    End Function

    Public Sub gGetShiftDataForDate(ByVal aMineName As String, _
                                    ByVal aDate As Date, _
                                    ByRef aShiftNames() As gShiftNamesType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetShiftDataForDateError

        Dim ShiftNamesDynaset As OraDynaset
        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        Dim ThisShiftName As String
        Dim ThisBeginTime As String
        Dim ThisEndTime As String
        Dim ThisShiftLength As Integer
        Dim ThisShiftOrder As Integer
        Dim ThisSampBeginTime As String
        Dim ThisSampEndTime As String
        Dim ThisBegHr As Integer
        Dim ThisBegMin As Integer
        Dim ThisEndHr As Integer
        Dim ThisEndMin As Integer
        Dim ShiftNameCnt As Integer
        Dim RecordCnt As Integer
        Dim ShiftLength As Integer

        'Over time a mine may change the number of shifts that
        'it has per day (typically 2 shifts to 3 shifts or 3 shifts
        'to 2 shifts).  Thus we need to be careful getting the shifts
        'for a mine -- it is date dependent.

        'First need to get the shift length for this date!
        ShiftLength = gGetShiftLengthForDate(aMineName, aDate)

        'Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pShiftLengthHrs", ShiftLength, ORAPARM_INPUT)
        params("pShiftLengthHrs").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_shift_names_bylength
        'pMineName           IN     VARCHAR2,
        'pShiftLengthHrs     IN     NUMBER,
        'pResult             IN OUT c_minenames)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_shift_names_bylength(:pMineName, " & _
                  ":pShiftLengthHrs, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        ShiftNamesDynaset = params("pResult").Value
        ClearParams(params)
        RecordCnt = ShiftNamesDynaset.RecordCount
        ReDim aShiftNames(RecordCnt)

        ShiftNameCnt = 0

        ShiftNamesDynaset.MoveFirst()

        Do While Not ShiftNamesDynaset.EOF
            ThisShiftName = ShiftNamesDynaset.Fields("shift").Value
            ThisBeginTime = ShiftNamesDynaset.Fields("begin_time").Value
            ThisEndTime = ShiftNamesDynaset.Fields("end_time").Value
            ThisShiftLength = ShiftNamesDynaset.Fields("shift_length_hrs").Value
            ThisShiftOrder = 0
            ThisSampBeginTime = ""
            ThisSampEndTime = ""

            ShiftNameCnt = ShiftNameCnt + 1

            'Need to get hour and minute for ThisBeginTime
            ThisBegHr = Val(Mid(ThisBeginTime, 1, 2))
            ThisBegMin = Val(Mid(ThisBeginTime, 4))

            'Need to get hour and minute for ThisBeginTime
            ThisEndHr = Val(Mid(ThisEndTime, 1, 2))
            ThisEndMin = Val(Mid(ThisEndTime, 4))

            aShiftNames(ShiftNameCnt).ShiftName = ThisShiftName
            aShiftNames(ShiftNameCnt).BeginTime = ThisBeginTime
            aShiftNames(ShiftNameCnt).EndTime = ThisEndTime
            aShiftNames(ShiftNameCnt).ShiftLength = ThisShiftLength
            aShiftNames(ShiftNameCnt).ShiftOrder = ThisShiftOrder
            aShiftNames(ShiftNameCnt).SampBeginTime = ThisSampBeginTime
            aShiftNames(ShiftNameCnt).SampEndTime = ThisSampEndTime

            aShiftNames(ShiftNameCnt).BeginHour = ThisBegHr
            aShiftNames(ShiftNameCnt).BeginMinute = ThisBegMin

            aShiftNames(ShiftNameCnt).EndHour = ThisEndHr
            aShiftNames(ShiftNameCnt).EndMinute = ThisEndMin

            ShiftNamesDynaset.MoveNext()
        Loop

        ShiftNamesDynaset.Close()
        Exit Sub

gGetShiftDataForDateError:
        MsgBox("Error getting shift data for date." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Shift Data Get Error")

        On Error Resume Next
        ClearParams(params)
        ShiftNamesDynaset.Close()
    End Sub

    Public Function gGetFirstShiftBegDtime2(ByVal aMineName As String, _
                                            ByVal aDate As Date) As Date

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ShiftNames() As gShiftNamesType

        'This is based on gGetFirstShiftBegDtime except that this function
        'does not depend on aDate being gActiveDate.

        'aDate may not be the gActiveDate thus we cannot use the info
        'in gShiftNames -- must reaccess it to be safe

        '03/08/2007, lss -- modified to go faster (added gNeedToChangeShiftNames).

        If gNeedToChangeShiftNames(gActiveMineNameLong, _
                                   gActiveDate, _
                                   aDate) = False Then
            'Assumes that the shifts in ShiftNames() are in time order
            gGetFirstShiftBegDtime2 = aDate.AddDays((gShiftNames(1).BeginHour / 24))
        Else
            gGetShiftDataForDate(aMineName, aDate, ShiftNames)
            'Assumes that the shifts in ShiftNames() are in time order
            gGetFirstShiftBegDtime2 = aDate.AddDays((ShiftNames(1).BeginHour / 24))
        End If
    End Function

    Public Function gGetLastShiftBegDtime2(ByVal aMineName As String, _
                                           ByVal aDate As Date) As Date

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ShiftNames() As gShiftNamesType

        'This is based on gGetLastShiftBegDtime except that this function
        'does not depend on aDate being gActiveDate.

        'aDate may not be the gActiveDate thus we cannot use the info
        'in gShiftNames -- must reaccess it to be safe

        '03/08/2007, lss -- modified to go faster (added gNeedToChangeShiftNames).

        If gNeedToChangeShiftNames(gActiveMineNameLong, _
                                   gActiveDate, _
                                   aDate) = False Then
            gGetLastShiftBegDtime2 = aDate.AddDays( _
                                    (gShiftNames(UBound(gShiftNames)).BeginHour / 24))
        Else
            gGetShiftDataForDate(aMineName, aDate, ShiftNames)
            'Assumes that the shifts in ShiftNames() are in time order
            gGetLastShiftBegDtime2 = aDate.AddDays( _
                                    (ShiftNames(UBound(ShiftNames)).BeginHour / 24))
        End If
    End Function

    Public Sub gAddGridAvgs(ByRef aGrid As AxvaSpread, _
                            ByVal aBegCol As Integer, _
                            ByVal aEndCol As Integer, _
                            ByVal aExcludeZeros As Boolean, _
                            ByVal aAddTotalRow As Boolean)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim ColIdx As Integer
        Dim RowIdx As Integer
        Dim SumVals As Double
        Dim NumVals As Integer
        Dim RoundVal As Integer

        'If necessary add horizontal divider and a total row.
        If aAddTotalRow = True Then
            With aGrid
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .set_RowHeight(.Row, 0.8)
                .BlockMode = True
                .Row = .MaxRows
                .Row2 = .MaxRows
                .Col = 0
                .Col2 = .MaxCols
                .CellType = CellTypeConstants.CellTypeStaticText 'SS_CELL_TYPE_STATIC_TEXT
                .Text = " "
                .TypeTextShadow = False
                .BackColor = Color.Black
                .BlockMode = False

                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 1
                .CellType = CellTypeConstants.CellTypeStaticText 'SS_CELL_TYPE_STATIC_TEXT
                .Text = "Averages"
            End With
        End If

        With aGrid
            For ColIdx = aBegCol To aEndCol
                .Col = ColIdx
                SumVals = 0
                NumVals = 0

                For RowIdx = 1 To .MaxRows - 2
                    .Row = RowIdx
                    SumVals = SumVals + .Value

                    If RowIdx = 1 Then
                        RoundVal = .TypeFloatDecimalPlaces
                    End If

                    If (aExcludeZeros = True And .Value <> 0) Or _
                       aExcludeZeros = False Then
                        NumVals = NumVals + 1
                    End If
                Next RowIdx

                'Place the average
                .Row = .MaxRows

                If NumVals <> 0 Then
                    .Value = Round(SumVals / NumVals, RoundVal)
                Else
                    .Value = 0
                End If
            Next ColIdx
        End With
    End Sub

    Public Sub gGetAllShiftsCbo2(ByVal aMineName As String, _
                                 ByVal aDate As Date, _
                                 ByVal aCboBox As ComboBox, _
                                 ByVal aAddSelectItem As Boolean)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetAllShiftsCbo2Error

        Dim ShiftNames() As gShiftNamesType
        Dim ItemIdx As Integer
        Dim RowIdx As Integer
        Dim UseHardCode As Boolean

        UseHardCode = True

        'Clear the by reference combo box
        For ItemIdx = 0 To aCboBox.Items.Count - 1
            aCboBox.Items.RemoveAt(0)
        Next ItemIdx
        If aAddSelectItem = True Then
            aCboBox.Items.Add("(Select shift...)")
        End If

        If aMineName = "Four Corners" Then
            If UseHardCode = True Then
                If aDate < gFcoChangeDate Then
                    aCboBox.Items.Add("1st")
                    aCboBox.Items.Add("2nd")
                    aCboBox.Items.Add("3rd")
                Else
                    aCboBox.Items.Add("Day")
                    aCboBox.Items.Add("Night")
                End If
            Else
                'Get the shift data for this date
                'Only Four Corners had 1st, 2nd, 3rd and then changed to Day & Night!
                gGetShiftDataForDate(aMineName, aDate, ShiftNames)
                For RowIdx = 1 To UBound(ShiftNames)
                    aCboBox.Items.Add(StrConv(ShiftNames(RowIdx).ShiftName, vbProperCase))
                Next RowIdx

                If aAddSelectItem = True Then
                    aCboBox.Items.Add("(Select shift...)")
                End If
            End If

            aCboBox.Text = aCboBox.Items(0)
        Else
            aCboBox.Items.Add("Day")
            aCboBox.Items.Add("Night")
            aCboBox.Text = aCboBox.Items(0)
        End If

        Exit Sub

gGetAllShiftsCbo2Error:
        MsgBox("Error getting shifts." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Shift Names Get Error")
    End Sub

    Public Function gGetPrevShift2(ByVal aDate As Date, _
                                   ByVal aShift As String) As gShiftInfoType

        '**********************************************************************
        'This function returns the previous shift from the date and shift
        'that it receives as input parameters.
        '
        '**********************************************************************

        Dim RowIdx As Integer
        Dim ShiftInfo As gShiftInfoType
        Dim ShiftNames() As gShiftNamesType

        'NOTE: This is a special case gGetPrevShift2 -- it is not generic!
        '      It is based on gActiveMineNameLong!

        'Find the shift we are in (gShiftNames)
        'Assume that the shifts in gShiftInfo are in the correct order
        'timewise.

        For RowIdx = 1 To UBound(gShiftNames)
            If StrConv(gShiftNames(RowIdx).ShiftName, vbUpperCase) = _
                StrConv(aShift, vbUpperCase) Then
                'We have found the shift, we want the shift that is previous
                'to this one

                If RowIdx <> 1 Then
                    'Just need to go back one shift for the current date.
                    ShiftInfo.dDate = aDate
                    ShiftInfo.Shift = gShiftNames(RowIdx - 1).ShiftName
                Else
                    'Need the last shift from the previous day.
                    'Will have to get a new set of shifts from the previous day!
                    ShiftInfo.dDate = aDate.AddDays(-1)

                    gGetShiftDataForDate(gActiveMineNameLong, _
                                         ShiftInfo.dDate, _
                                         ShiftNames)

                    ShiftInfo.Shift = ShiftNames(UBound(ShiftNames)).ShiftName
                End If
            End If
        Next RowIdx

        gGetPrevShift2 = ShiftInfo
    End Function

    Public Function gGetNextShift2(ByVal aDate As Date, _
                                   ByVal aShift As String) As gShiftInfoType

        '**********************************************************************
        'This function returns the next shift from the date and shift
        'that it receives as input parameters.
        '
        '**********************************************************************

        Dim RowIdx As Integer
        Dim ShiftInfo As gShiftInfoType
        Dim ShiftNames() As gShiftNamesType

        'NOTE: This is a special case gGetNextShift2 -- it is not generic!
        '      It is based on gActiveMineNameLong!

        'Find the shift we are in (gShiftNames)
        'Assume that the shifts in gShiftInfo are in the correct order
        'timewise.

        For RowIdx = 1 To UBound(gShiftNames)
            If StrConv(gShiftNames(RowIdx).ShiftName, vbUpperCase) = _
                StrConv(aShift, vbUpperCase) Then
                'We have found the shift, we want the shift that is next
                'to this one

                If RowIdx <> UBound(gShiftNames) Then
                    'Just need to go forward one shift for the current date.
                    ShiftInfo.dDate = aDate
                    ShiftInfo.Shift = gShiftNames(RowIdx + 1).ShiftName
                Else
                    'Need the first shift from the next day.
                    ShiftInfo.dDate = aDate.AddDays(1)

                    gGetShiftDataForDate(gActiveMineNameLong, _
                                         ShiftInfo.dDate, _
                                         ShiftNames)

                    ShiftInfo.Shift = ShiftNames(1).ShiftName
                End If
            End If
        Next RowIdx

        gGetNextShift2 = ShiftInfo
    End Function

    Public Function gGetNumShifts2(ByVal aMineName As String, _
                                   ByVal aDate As Date) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetNumShifts2Error

        Dim ShiftNamesDynaset As OraDynaset
        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ThisShiftName As String
        Dim ShiftLength As Integer

        gGetNumShifts2 = 0

        'First need to get the shift length for this date
        ShiftLength = gGetShiftLengthForDate(aMineName, aDate)

        'Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pShiftLengthHrs", ShiftLength, ORAPARM_INPUT)
        params("pShiftLengthHrs").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_shift_names_bylength
        'pMineName           IN     VARCHAR2,
        'pShiftLengthHrs     IN     NUMBER,
        'pResult             IN OUT c_minenames)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_shift_names_bylength(:pMineName, " & _
                  ":pShiftLengthHrs, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        ShiftNamesDynaset = params("pResult").Value
        ClearParams(params)

        gGetNumShifts2 = ShiftNamesDynaset.RecordCount

        ShiftNamesDynaset.Close()
        Exit Function

gGetNumShifts2Error:
        MsgBox("Error getting number of shifts." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Shift Number Error")

        On Error Resume Next
        gGetNumShifts2 = 0
        ClearParams(params)
        ShiftNamesDynaset.Close()
    End Function

    Public Function gGetShiftNames2(ByVal aMineName As String, _
                                    ByVal aDate As Date, _
                                    ByRef aShiftNames() As gShiftNamesType) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetShiftNames2Error

        Dim ShiftNamesDynaset As OraDynaset
        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        Dim ThisShiftName As String
        Dim ThisBeginTime As String
        Dim ThisEndTime As String
        Dim ThisShiftLength As Integer
        Dim ThisShiftOrder As Integer
        Dim ThisSampBeginTime As String
        Dim ThisSampEndTime As String
        Dim ThisBegHr As Integer
        Dim ThisBegMin As Integer
        Dim ThisEndHr As Integer
        Dim ThisEndMin As Integer
        Dim ShiftNameCnt As Integer
        Dim RecordCnt As Integer
        Dim ShiftLength As Integer

        'Need to assign the following:
        '1) gShiftNames() As String
        '2) gNumShifts As Integer
        '3) gFirstShift As String
        '4) gLastShift as String
        '5) BeginHour As Integer
        '6) BeginMinute As Integer
        '7) EndHour As Integer
        '8) EndMinute as Integer

        'First need to get the shift length for this date
        ShiftLength = gGetShiftLengthForDate(aMineName, aDate)

        'Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pShiftLengthHrs", ShiftLength, ORAPARM_INPUT)
        params("pShiftLengthHrs").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_shift_names_bylength
        'pMineName           IN     VARCHAR2,
        'pShiftLengthHrs     IN     NUMBER,
        'pResult             IN OUT c_minenames)
        ' Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_shift_names_bylength(:pMineName, " & _
                  ":pShiftLengthHrs, :pResult);end;", ORASQL_FAILEXEC)
        'Set 
        ShiftNamesDynaset = params("pResult").Value
        ClearParams(params)

        RecordCnt = ShiftNamesDynaset.RecordCount
        ReDim aShiftNames(RecordCnt)

        ShiftNameCnt = 0

        ShiftNamesDynaset.MoveFirst()

        Do While Not ShiftNamesDynaset.EOF
            ThisShiftName = ShiftNamesDynaset.Fields("shift").Value
            ThisBeginTime = ShiftNamesDynaset.Fields("begin_time").Value
            ThisEndTime = ShiftNamesDynaset.Fields("end_time").Value
            ThisShiftLength = ShiftNamesDynaset.Fields("shift_length_hrs").Value
            ThisShiftOrder = 0
            ThisSampBeginTime = ""
            ThisSampEndTime = ""

            ShiftNameCnt = ShiftNameCnt + 1

            'Need to get hour and minute for ThisBeginTime
            ThisBegHr = Val(Mid(ThisBeginTime, 1, 2))
            ThisBegMin = Val(Mid(ThisBeginTime, 4))

            'Need to get hour and minute for ThisBeginTime
            ThisEndHr = Val(Mid(ThisEndTime, 1, 2))
            ThisEndMin = Val(Mid(ThisEndTime, 4))

            aShiftNames(ShiftNameCnt).ShiftName = ThisShiftName
            aShiftNames(ShiftNameCnt).BeginTime = ThisBeginTime
            aShiftNames(ShiftNameCnt).EndTime = ThisEndTime
            aShiftNames(ShiftNameCnt).ShiftLength = ThisShiftLength
            aShiftNames(ShiftNameCnt).ShiftOrder = ThisShiftOrder
            aShiftNames(ShiftNameCnt).SampBeginTime = ThisSampBeginTime
            aShiftNames(ShiftNameCnt).SampEndTime = ThisSampEndTime

            aShiftNames(ShiftNameCnt).BeginHour = ThisBegHr
            aShiftNames(ShiftNameCnt).BeginMinute = ThisBegMin

            aShiftNames(ShiftNameCnt).EndHour = ThisEndHr
            aShiftNames(ShiftNameCnt).EndMinute = ThisEndMin

            ShiftNamesDynaset.MoveNext()
        Loop

        ShiftNamesDynaset.Close()
        gGetShiftNames2 = True
        Exit Function

gGetShiftNames2Error:
        MsgBox("Error getting shift names." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Shift Names Get Error")

        On Error Resume Next
        gGetShiftNames2 = False
        ClearParams(params)
        ShiftNamesDynaset.Close()
    End Function

    Public Function gGetNumShiftsRge(ByVal aMineName As String, _
                                     ByVal aBeginDate As Date, _
                                     ByVal aEndDate As Date) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetNumShiftsRgeError

        Dim ShiftNamesDynaset As OraDynaset
        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ThisShiftName As String
        Dim ShiftLength As Integer
        Dim DateIdx As Date

        gGetNumShiftsRge = 0

        'This is an ugly way to do this -- will revisit at a later date
        '(see gGetNumShiftsRge2 below).
        'gGetNumShiftsRge2 makes use of get_number_of_shifts in MOIS_UTILITIES2.
        Dim ddif As Integer = Abs(DateDiff(DateInterval.Day, aBeginDate, aEndDate))
        For i As Integer = 0 To ddif
            DateIdx = DateAdd(DateInterval.Day, i, aBeginDate)
            'First need to get the shift length for this date
            ShiftLength = gGetShiftLengthForDate(aMineName, DateIdx)

            ' Set
            params = gDBParams

            params.Add("pMineName", aMineName, ORAPARM_INPUT)
            params("pMineName").serverType = ORATYPE_VARCHAR2

            params.Add("pShiftLengthHrs", ShiftLength, ORAPARM_INPUT)
            params("pShiftLengthHrs").serverType = ORATYPE_NUMBER

            params.Add("pResult", 0, ORAPARM_OUTPUT)
            params("pResult").serverType = ORATYPE_CURSOR

            'PROCEDURE get_shift_names_bylength
            'pMineName           IN     VARCHAR2,
            'pShiftLengthHrs     IN     NUMBER,
            'pResult             IN OUT c_minenames)
            'Set 
            SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_shift_names_bylength(:pMineName, " & _
                      ":pShiftLengthHrs, :pResult);end;", ORASQL_FAILEXEC)
            'Set 
            ShiftNamesDynaset = params("pResult").Value
            ClearParams(params)

            gGetNumShiftsRge = gGetNumShiftsRge + ShiftNamesDynaset.RecordCount
            ShiftNamesDynaset.Close()
        Next i ' DateIdx

        Exit Function

gGetNumShiftsRgeError:
        MsgBox("Error getting number of shifts for range." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Shift Number Error")

        On Error Resume Next
        gGetNumShiftsRge = 0
        ClearParams(params)
        ShiftNamesDynaset.Close()
    End Function

    Public Function gGetNumShiftsRge2(ByVal aMineName As String, _
                                      ByVal aBeginDate As Date, _
                                      ByVal aEndDate As Date) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetNumShiftsRge2Error

        Dim ShiftNamesDynaset As OraDynaset
        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim FirstShift As String
        Dim LastShift As String
        Dim ShiftCount As Integer

        FirstShift = StrConv(gGetFirstShift2(aMineName, aBeginDate), vbUpperCase)
        LastShift = StrConv(gGetLastShift2(aMineName, aEndDate), vbUpperCase)

        'This function makes use of Bob's procedure get_number_of_shifts
        'in MOIS_UTILITIES2

        gGetNumShiftsRge2 = 0

        '  Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pStartDate", aBeginDate, ORAPARM_INPUT)
        params("pStartDate").serverType = ORATYPE_DATE

    params.Add("pStartShift", StrConv(FirstShift, vbUpperCase), ORAPARM_INPUT)
        params("pStartShift").serverType = ORATYPE_VARCHAR2

    params.Add("pStopDate", aEndDate, ORAPARM_INPUT)
        params("pStopDate").serverType = ORATYPE_DATE

    params.Add("pStopShift", StrConv(LastShift, vbUpperCase), ORAPARM_INPUT)
        params("pStopShift").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        'PROCEDURE get_number_of_shifts
        'pMineName           IN     VARCHAR2,
        'pStartDate          IN     DATE,
        'pStartShift         IN     VARCHAR2,
        'pStopDate           IN     DATE,
        'pStopShift          IN     VARCHAR2,
        'pResult             IN OUT INTEGER);
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities2.get_number_of_shifts(:pMineName, " & _
                  ":pStartDate, :pStartShift, :pStopDate, :pStopShift, :pResult);end;", ORASQL_FAILEXEC)
        ShiftCount = params("pResult").Value
        ClearParams(params)

        gGetNumShiftsRge2 = ShiftCount

        Exit Function

gGetNumShiftsRge2Error:
        MsgBox("Error getting number of shifts for range." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Shift Number Error")

        On Error Resume Next
        gGetNumShiftsRge2 = 0
        ClearParams(params)
        ShiftNamesDynaset.Close()
    End Function

    Public Function gGetAllMatl(ByVal aMineName As String, _
                                ByVal aMatlTypeName As String) As OraDynaset

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetAllMatlError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        'Get all existing materials for current mine & selected material type
        ' Set 
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pMatlTypeName", aMatlTypeName, ORAPARM_INPUT)
        params("pMatlTypeName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_all_matls
        'pMineName        IN     VARCHAR2,
        'pMatlTypeName    IN     VARCHAR2,
        'pResult          IN OUT c_matl)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities.get_all_matls(:pMineName," + _
                  ":pMatlTypeName, :pResult);end;", ORASQL_FAILEXEC)
        ' Set 
        gGetAllMatl = params("pResult").Value
        ClearParams(params)

        Exit Function

gGetAllMatlError:
        MsgBox("Error getting materials." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Material Get Error")
    End Function

    Public Function gGetDateFromMonthAndYear(ByVal aMoYrStr As String, _
                                             ByVal aDay As Integer) As Date

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim CommaPos As Integer
        Dim MonthStr As String
        Dim MonthVal As Integer
        Dim YearVal As Integer

        'Assume that aMoYrStr will be something like "January, 2006"

        CommaPos = InStr(aMoYrStr, ",")

        If CommaPos = 0 Then
            gGetDateFromMonthAndYear = #12/31/8888#
            Exit Function
        End If

        MonthStr = Mid(aMoYrStr, 1, CommaPos - 1)

        If IsNumeric(Mid(aMoYrStr, CommaPos + 1)) Then
            YearVal = Val(Mid(aMoYrStr, CommaPos + 1))
        Else
            gGetDateFromMonthAndYear = #12/31/8888#
            Exit Function
        End If

        MonthVal = 0

        Select Case MonthStr
            Case "January", "Jan"
                MonthVal = 1
            Case "February", "Feb"
                MonthVal = 2
            Case "March", "Mar"
                MonthVal = 3
            Case "April", "Apr"
                MonthVal = 4
            Case "May", "May"
                MonthVal = 5
            Case "June", "Jun"
                MonthVal = 6
            Case Is = "July", "Jul"
                MonthVal = 7
            Case "August", "Aug"
                MonthVal = 8
            Case "September", "Sep"
                MonthVal = 9
            Case "October", "Oct"
                MonthVal = 10
            Case "November", "Nov"
                MonthVal = 11
            Case "December", "Dec"
                MonthVal = 12
        End Select

        gGetDateFromMonthAndYear = CDate(CStr(MonthVal) & "/" & CStr(aDay) & _
                                   "/" & CStr(YearVal))
    End Function

    Public Function gPadLeftChar(ByVal aString As String, _
                                 ByVal aLength As Integer, _
                                 ByVal aChar As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        gPadLeftChar = Trim(aString)

        Do While Len(gPadLeftChar) < aLength
            gPadLeftChar = aChar + gPadLeftChar
        Loop
    End Function

    Public Function gPadRightChar(ByVal aString As String, _
                                  ByVal aLength As Integer, _
                                  ByVal aChar As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        gPadRightChar = Trim(aString)

        Do While Len(gPadRightChar) < aLength
            gPadRightChar = gPadRightChar + aChar
        Loop
    End Function

    Public Sub gGetMassBalDataAbbrv(ByVal aBegDate As Date, _
                                    ByVal aBegShift As String, _
                                    ByVal aEndDate As Date, _
                                    ByVal aEndShift As String, _
                                    ByVal aMineName As String, _
                                    ByRef aMassBalDataAbbrv As gMassBalDataAbbrvType)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim NumShifts As Long
        Dim FloatPlantCirc(20, 15) As Object
        Dim FloatPlantGmt(5, 15) As Object
        Dim AdjFdTons As Double
        Dim FdBpl As Single
        Dim PctRcvry As Single

        'The "All" in the following function calls indicates "All" crew
        'numbers.

        With aMassBalDataAbbrv
            Select Case aMineName
                Case Is = "South Fort Meade"
                    NumShifts = gGetSfFloatPlantBalanceData(FloatPlantCirc, _
                                                            FloatPlantGmt, _
                                                            aBegDate, _
                                                            StrConv(aBegShift, vbUpperCase), _
                                                            aEndDate, _
                                                            StrConv(aEndShift, vbUpperCase), _
                                                            "All", _
                                                            0)

                    .AdjFdTons = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcFdTons)
                    .FdBpl = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcFdBpl)

                    .PctRcvry = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcPctRcvry)
                    .GmtBpl = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcTlBpl)

                Case Is = "Hookers Prairie"
                    'NumShifts = gGetHpFloatPlantBalanceData(FloatPlantCirc, _
                    '                                        FloatPlantGmt, _
                    '                                        aBegDate, _
                    '                                        StrConv(aBegShift, vbUpperCase), _
                    '                                        aEndDate, _
                    '                                        StrConv(aEndShift, vbUpperCase), _
                    '                                        "All", _
                    '                                        gMassBalanceMode)

                    .AdjFdTons = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcFdTons)
                    .FdBpl = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcFdBpl)
                    .PctRcvry = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcPctRcvry)
                    .GmtBpl = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcTlBpl)

                Case Is = "Wingate"
                    'NumShifts = gGetWgFloatPlantBalanceData(FloatPlantCirc(), _
                    '                                        FloatPlantGmt(), _
                    '                                        aBegDate, _
                    '                                        StrConv(aBegShift, vbUpperCase), _
                    '                                        aEndDate, _
                    '                                        StrConv(aEndShift, vbUpperCase), _
                    '                                        "All")

                    .AdjFdTons = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcFdTons)
                    .FdBpl = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcFdBpl)
                    .PctRcvry = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcPctRcvry)
                    .GmtBpl = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcTlBpl)

                Case Is = "Four Corners"
                    NumShifts = gGetFcFloatPlantBalanceData(FloatPlantCirc, _
                                                            FloatPlantGmt, _
                                                            aBegDate, _
                                                            StrConv(aBegShift, vbUpperCase), _
                                                            aEndDate, _
                                                            StrConv(aEndShift, vbUpperCase), _
                                                            "All", _
                                                            2, _
                                                            gMassBalanceMode)

                    .AdjFdTons = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcFdTons)
                    .FdBpl = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcFdBpl)
                    .PctRcvry = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcPctRcvry)
                    .GmtBpl = FloatPlantGmt(fFloatPlantGmtRowEnum.GrCalculatedGmtBpl, fFloatPlantGmtColEnum.GcTlBpl)

                Case Is = "Hopewell"
                    'Removed functionality -- 05/13/2010
                    'Removed Hopewell mass balance.
                    ''NumShifts = gGetHwFloatPlantBalanceData(FloatPlantCirc(), _
                    ''                                        FloatPlantGmt(), _
                    ''                                        aBegDate, _
                    ''                                        StrConv(aBegShift, vbUpperCase), _
                    ''                                        aEndDate, _
                    ''                                        StrConv(aEndShift, vbUpperCase), _
                    ''                                        "All", _
                    ''                                        2)
                    ''
                    ''.AdjFdTons = FloatPlantGmt(GrCalculatedGmtBpl, GcFdTons)
                    ''.FdBpl = FloatPlantGmt(GrCalculatedGmtBpl, GcFdBpl)
                    ''.PctRcvry = FloatPlantGmt(GrCalculatedGmtBpl, GcPctRcvry)
                    ''.GmtBpl = FloatPlantGmt(GrCalculatedGmtBpl, GcTlBpl)
            End Select
        End With
    End Sub

    Public Function gNeedToChangeShiftNames(ByVal aMineName As String, _
                                            ByVal aOldDate As Date, _
                                            ByVal aNewDate As Date) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'On 08/14/2006 FourCorners switched from 3 shifts to 2 shifts
        '(1st, 2nd, 3rd to Day, Night).  It is the only mine that has this
        'issue.

        If aMineName <> "Four Corners" Then
            gNeedToChangeShiftNames = False
            Exit Function
        End If

        'Special for Four Corners Only!
        'Special for Four Corners Only!
        'Special for Four Corners Only!

        'Make sure that the dates are in the format we want
        aOldDate = CDate(Format(aOldDate, "MM/dd/yyyy"))
        aNewDate = CDate(Format(aNewDate, "MM/dd/yyyy"))

        'For Four Corners:
        '08/13/2006   1st, 2nd, 3rd
        '08/14/2006   Day, Night

        If aOldDate = aNewDate Then
            gNeedToChangeShiftNames = False
            Exit Function
        End If

        If aOldDate < aNewDate Then
            If aOldDate < gFcoChangeDate And aNewDate >= gFcoChangeDate Then
                gNeedToChangeShiftNames = True
                Exit Function
            Else
                gNeedToChangeShiftNames = False
                Exit Function
            End If
        End If

        If aOldDate > aNewDate Then
            If aOldDate >= gFcoChangeDate And aNewDate < gFcoChangeDate Then
                gNeedToChangeShiftNames = True
                Exit Function
            Else
                gNeedToChangeShiftNames = False
                Exit Function
            End If
        End If
    End Function

    Public Function gGetFirstShiftHardCode(ByVal aMineName As String, _
                                           ByVal aDate As Date) As String

        '**********************************************************************
        '
        '
        '
        '
        '**********************************************************************

        If aMineName = "Four Corners" Then
            If aDate < gFcoChangeDate Then
                gGetFirstShiftHardCode = "1ST"
            Else
                gGetFirstShiftHardCode = "DAY"
            End If
        Else
            gGetFirstShiftHardCode = "DAY"
        End If
    End Function

    Public Function gGetLastShiftHardCode(ByVal aMineName As String, _
                                          ByVal aDate As Date) As String

        '**********************************************************************
        '
        '
        '
        '
        '**********************************************************************

        If aMineName = "Four Corners" Then
            If aDate < gFcoChangeDate Then
                gGetLastShiftHardCode = "3RD"
            Else
                gGetLastShiftHardCode = "NIGHT"
            End If
        Else
            gGetLastShiftHardCode = "NIGHT"
        End If
    End Function

    Public Function gRoundHalf(ByVal aNumber As Single) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim TempNumber As Single
        Dim TempNumberInt As Single
        Dim TempNumberFra As Single

        'For use for reagent calculations at Hookers Prairie only!
        'aNumber will not be negative and will not be large!

        TempNumber = Round(aNumber, 1)
        TempNumberInt = Int(Round(aNumber, 1))

        TempNumberFra = Round(TempNumber - TempNumberInt, 1)

        If TempNumberFra <= 0.2 Then
            gRoundHalf = TempNumberInt
        End If

        If TempNumberFra >= 0.3 And TempNumberFra <= 0.5 Then
            gRoundHalf = TempNumberInt + 0.5
        End If

        If TempNumberFra >= 0.6 And TempNumberFra <= 0.9 Then
            gRoundHalf = TempNumberInt + 1
        End If
    End Function

    Public Function gGetEqptMsrmntValue(ByVal aMineName As String, _
                                        ByVal aDate As Date, _
                                        ByVal aShift As String, _
                                        ByVal aEqptTypeName As String, _
                                        ByVal aEqptName As String, _
                                        ByVal aMeasureName As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetEqptMsrmntValueError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt

        gGetEqptMsrmntValue = ""

        'Set
        params = gDBParams

    params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptTypeName", aEqptTypeName, ORAPARM_INPUT)
        params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

    params.Add("pEqptName", aEqptName, ORAPARM_INPUT)
        params("pEqptName").serverType = ORATYPE_VARCHAR2

    params.Add("pMeasureName", aMeasureName, ORAPARM_INPUT)
        params("pMeasureName").serverType = ORATYPE_VARCHAR2

    params.Add("pDate", aDate, ORAPARM_INPUT)
        params("pDate").serverType = ORATYPE_DATE

    params.Add("pShift", StrConv(aShift, vbUpperCase), ORAPARM_INPUT)
        params("pShift").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_NUMBER

        'PROCEDURE get_eqpt_msrmnt_value     (pMineName          IN     VARCHAR2,
        '                                     pEqptTypeName      IN     VARCHAR2,
        '                                     pEqptName          IN     VARCHAR2,
        '                                     pMeasureName       IN     VARCHAR2,
        '                                     pDate              IN     DATE,
        '                                     pShift             IN     VARCHAR2,
        '                                     pResult            IN OUT VARCHAR2)
        'Set 
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_utilities2.get_eqpt_msrmnt_value (:pMineName, " & _
                                         ":pEqptTypeName, :pEqptName, " & _
                                         ":pMeasureName, :pDate, " & _
                                         ":pShift, :pResult);end;", ORASQL_FAILEXEC)

        If Not IsDBNull(params("pResult").Value) Then
            gGetEqptMsrmntValue = params("pResult").Value
        Else
            gGetEqptMsrmntValue = ""
        End If

        ClearParams(params)

        Exit Function

gGetEqptMsrmntValueError:
        MsgBox("Error getting value." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "Get Error")

        On Error Resume Next
        ClearParams(params)
    End Function
    Function RunSPReturnDate(ByVal strSP As String, _
                         ByVal ParamArray params() As Object) As Date

        '********************************************************************
        '
        '
        '
        '********************************************************************

        Dim SQLStmt As OraSqlStmt
        Dim dbparams As OraParameters
        Dim a As Integer

        dbparams = gOradatabase.Parameters
        collectParams(dbparams, params)

        'Return the resultant recordset
        SQLStmt = gOradatabase.CreateSql _
               (strSP, ORASQL_FAILEXEC)
        RunSPReturnDate = dbparams(dbparams.count - 1).Value

        'Clear Params
        For a = 0 To dbparams.count - 1
            dbparams.Remove(0)
        Next

        Exit Function
    End Function
    Sub collectParams(ByRef paramlist As OraParameters, _
                      ByVal ParamArray argparams() As Object)

        '********************************************************************
        '
        '
        '
        '********************************************************************

        Dim params As Object

        params = argparams(0)
        Dim I As Integer, v As Object
        For I = LBound(params) To UBound(params)
            If TypeName(params(I)(1)) = "String" Then
                v = IIf(params(I)(1) = "", DBNull.Value, params(I)(1))
            ElseIf IsNumeric(params(I)(1)) Then
                v = IIf(params(I)(1) < 0, DBNull.Value, params(I)(1))
            Else
                v = params(I)(1)
            End If
            'Skip adding parameter if its server type value is ORATYPE_CURSOR so
            'that code will work with Oracle 8. Should work regressively. Appears
            'that explicit addition of cursor parameter was never required and
            'actually causes errors in Oracle 8.
            If params(I)(3) <> ORATYPE_CURSOR Then
                paramlist.Add(params(I)(0), params(I)(1), params(I)(2))
                paramlist(I).servertype = params(I)(3)
            End If
        Next I

        Exit Sub
    End Sub



    Public Function gGetMoisStartDate(ByVal aMineName As String) As Date

        '**********************************************************************
        '    this is from modGlobalInfo2 but added here 
        '
        '
        '**********************************************************************

        gGetMoisStartDate = #12/31/8888#

        Select Case aMineName
            Case Is = "South Fort Meade"
                gGetMoisStartDate = #12/21/1995#
            Case Is = "Hookers Prairie"
                gGetMoisStartDate = #1/3/2000#
            Case Is = "Wingate"
                gGetMoisStartDate = #1/3/2005#
            Case Is = "Four Corners"
                gGetMoisStartDate = #2/2/2006#
            Case Is = "Hopewell"
                gGetMoisStartDate = #9/27/2006#
        End Select
    End Function


End Module
