Imports OracleInProcServer
Imports System.Math
Module modAutocad
    '**********************************************************************
    'AUTOCAD MODULE
    '
    '
    '**********************************************************************
    '   Maintenance Log
    '
    '   01/24/2001, lss
    '      Added this module.
    '   01/29/2001, lss
    '      Continued working on this module.
    '   09/13/2001, lss
    '      Modified gPolyline and gPolycline.
    '   09/18/2001, lss
    '      Added gMakeColCircle, gMakeYrCircle.
    '   11/19/2001, lss
    '      Modified gOpenDxf -- added aAcadVer.
    '   03/07/2002, lss
    '      Added gEntitiesDxf.
    '   04/04/2002, lss
    '      Added gCreateProspectGrid and related stuff.
    '   04/17/2002, lss
    '      Added gViewAutoCadMap.
    '   04/22/2002, lss
    '      Added gMakeSpCircle.
    '   04/23/2002, lss
    '      Added gCreateProspectGrid (added aColor as a parameter).
    '      Added gGetAcadColor.
    '   06/06/2002, lss
    '      Modified gMakeCompBox and AddCompData.
    '   06/19/2002, lss
    '      Modified gMakeYrCircle.
    '   11/18/2002, lss
    '      Modified polyline width for gCreateProspectGrid.
    '   11/19/2002, lss
    '      Modified gMakeSpCircle.
    '   01/28/2004, lss
    '      Added gGetCurrent6MonthMap.
    '   02/27/2004, lss
    '      Changed gMakeSpCircle -- added aTextHeight.
    '   12/07/2004, lss
    '      Added aIncludeAreas to gCreateProspectGrid().
    '   03/11/2005, lss
    '      Added gMakeCompBoxSimp and AddCompDataSimp.
    '   05/16/2005, lss
    '      Modified gMakeYrCircle() --  If Len(Trim(aHole)) <= 3 Then
    '      TextHeight = 70 Else TextHeight = 50 End If.
    '   05/16/2006, lss
    '      Modified Function gCreateProspectGrid2 (6000 to 7000) for line
    '      extensions.
    '   03/15/2007, lss
    '      Removed gViewAutoCadMap functionality for now (remarked it out).
    '   08/15/2007, lss
    '      Removed Public gHey -- not really sure what this was for!
    '   05/20/2008, lss
    '      Added Sub gGetHoleCoords.
    '   04/28/2010, lss
    '      Added Public Sub gMakeColCircle2.
    '
    '**********************************************************************


    Public gEntityNumber As Long

    Dim mXcoord As Double
    Dim mYcoord As Double
    Public gProspX(0 To 16, 0 To 16) As Double
    Public gProspY(0 To 16, 0 To 16) As Double
    Public gProspIntX(0 To 17, 0 To 17) As Double
    Public gProspIntY(0 To 17, 0 To 17) As Double
    Dim mSectnCoordsDynaset As OraDynaset
    Dim mMapDynaset As OraDynaset

    Public Structure gPtCoordType
        Public X As Double
        Public Y As Double
    End Structure

    Public Sub gWriteLine(ByVal aFileNumber As Integer, _
                          ByVal aString As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        PrintLine(aFileNumber, aString)

        'My.Computer.FileSystem.WriteAllText(

    End Sub

    Public Sub gMakeCircle(ByVal aFileNumber As Integer, _
                           ByVal aXcoord As Double, _
                           ByVal aYcoord As Double, _
                           ByVal aRadius As Single, _
                           ByVal aLayer As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        '--> 2002  indicates added for AutoCAD 200, AutoCAD 2002

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "CIRCLE")
        gWriteLine(aFileNumber, "5")                 '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbCircle")        '--> 2002
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aXcoord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aYcoord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")
        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aRadius)))
    End Sub

    Public Sub gMakeColCircle(ByVal aFileNumber As Integer, _
                              ByVal aXcoord As Double, _
                              ByVal aYcoord As Double, _
                              ByVal aRadius As Single, _
                              ByVal aLayer As String, _
                              ByVal aColor As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        '--> 2002  indicates added for AutoCAD 200, AutoCAD 2002

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "CIRCLE")
        gWriteLine(aFileNumber, "5")
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbCircle")        '--> 2002
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aXcoord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aYcoord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")
        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aRadius)))

        If aColor <> 0 Then
            gWriteLine(aFileNumber, "62")
            gWriteLine(aFileNumber, Trim(Str(aColor)))
        End If
    End Sub

    Public Sub gMakeColCircle2(ByVal aFileNumber As Integer, _
                               ByVal aXcoord As Double, _
                               ByVal aYcoord As Double, _
                               ByVal aRadius As Single, _
                               ByVal aLayer As String, _
                               ByVal aColor As Integer, _
                               ByVal aText As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim TextHeight As Single
        Dim TextAlign As Single
        Dim HorAlign As Single
        Dim VerAlign As Single

        '--> 2002  indicates added for AutoCAD 200, AutoCAD 2002

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "CIRCLE")
        gWriteLine(aFileNumber, "5")
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbCircle")        '--> 2002
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aXcoord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aYcoord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")
        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aRadius)))

        If aColor <> 0 Then
            gWriteLine(aFileNumber, "62")
            gWriteLine(aFileNumber, Trim(Str(aColor)))
        End If

        '-----
        TextHeight = 9

        'MTEXT not working in AutoCAD2002 yet!
        'TextAlign = 5   'Middle Center
        'gTextMline aFileNumber, aXCoord, aYCoord, _
        '           TextHeight, TextAlign, _
        '           aHole, aLayer, aColor
        HorAlign = 1    'Center
        VerAlign = 2    'Middle
        gTextline2(aFileNumber, aXcoord, aYcoord, TextHeight, HorAlign, _
                   VerAlign, aText, aLayer, aColor)
    End Sub

    Public Sub gMakeYrCircle(ByVal aFileNumber As Integer, _
                             ByVal aXcoord As Double, _
                             ByVal aYcoord As Double, _
                             ByVal aRadius As Single, _
                             ByVal aLayer As String, _
                             ByVal aHole As String, _
                             ByVal aColor As Integer, _
                             ByVal aYear As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        '--> 2002  indicates added for AutoCAD 2000, AutoCAD 2002

        Dim TextHeight As Single
        Dim TextAlign As Single
        Dim HorAlign As Single
        Dim VerAlign As Single
        Dim YearStr As String

        'Make two circles

        'Circle #1
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "CIRCLE")
        gWriteLine(aFileNumber, "5")                    '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))     '--> 2002
        gWriteLine(aFileNumber, "100")                  '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")           '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")                  '--> 2002
        gWriteLine(aFileNumber, "AcDbCircle")           '--> 2002
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aXcoord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aYcoord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")
        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aRadius)))

        If aColor <> 0 Then
            gWriteLine(aFileNumber, "62")
            gWriteLine(aFileNumber, Trim(Str(aColor)))
        End If

        'Circle #2
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "CIRCLE")
        gWriteLine(aFileNumber, "5")                 '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbCircle")        '--> 2002
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aXcoord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aYcoord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")
        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aRadius + 5)))

        If aColor <> 0 Then
            gWriteLine(aFileNumber, "62")
            gWriteLine(aFileNumber, Trim(Str(aColor)))
        End If

        'Colors     1   Red
        '           2   Yellow
        '           3   Green
        '           4   Cyan
        '           5   Blue
        '           6   Magenta
        '           7   White
        '           40  Orange

        'Text Alignment -- for MTEXT
        'Top Left           71 = 1
        'Middle Left        71 = 4
        'Bottom Left        71 = 7
        '-------------------------
        'Top Center         71 = 2
        'Middle Center      71 = 5
        'Bottom Center      71 = 8
        '-------------------------
        'Top Right          71 = 3
        'Middle Right       71 = 6
        'Bottom Right       71 = 9

        If aYear <> "XX" Then   'Was -99 --> indicates missing hole
            If Trim(aYear) = "0" Or Trim(aYear) = "1" Or _
                Trim(aYear) = "2" Or Trim(aYear) = "3" Or _
                Trim(aYear) = "4" Or Trim(aYear) = "5" Or _
                Trim(aYear) = "6" Or Trim(aYear) = "7" Or _
                Trim(aYear) = "8" Or Trim(aYear) = "9" Then

                YearStr = "0" + Trim(Str(aYear))
            Else
                YearStr = Trim(aYear)
            End If
            TextHeight = 120

            'MTEXT not working in AutoCAD2002 yet!
            'TextAlign = 5   'Middle Center
            'gTextMline aFileNumber, aXCoord, aYCoord, _
            '           TextHeight, TextAlign, _
            '           YearStr, aLayer, aColor

            HorAlign = 1    'Center
            VerAlign = 2    'Middle
            gTextline2(aFileNumber, aXcoord, aYcoord, TextHeight, HorAlign, _
                VerAlign, YearStr, aLayer, aColor)
        Else
            If Len(Trim(aHole)) <= 3 Then
                TextHeight = 70
            Else
                TextHeight = 55
            End If

            'MTEXT not working in AutoCAD2002 yet!
            'TextAlign = 5   'Middle Center
            'gTextMline aFileNumber, aXCoord, aYCoord, _
            '           TextHeight, TextAlign, _
            '           aHole, aLayer, aColor
            HorAlign = 1    'Center
            VerAlign = 2    'Middle
            gTextline2(aFileNumber, aXcoord, aYcoord, TextHeight, HorAlign, _
                VerAlign, aHole, aLayer, aColor)
        End If
    End Sub

    Public Sub gMakeSpCircle(ByVal aFileNumber As Integer, _
                             ByVal aXcoord As Double, _
                             ByVal aYcoord As Double, _
                             ByVal aRadius As Single, _
                             ByVal aLayer As String, _
                             ByVal aHole As String, _
                             ByVal aColor As Integer, _
                             ByVal aVal As String, _
                             ByVal aShowVal As Boolean, _
                             ByVal aShowHole As Boolean, _
                             ByVal aSource As String, _
                             ByVal aTextHeight As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        '--> 2002  indicates added for AutoCAD 200, AutoCAD 2002

        Dim TextHeight As Single
        Dim HorAlign As Single
        Dim VerAlign As Single
        Dim YearStr As String

        'Make two circles

        'Circle #1
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "CIRCLE")
        gWriteLine(aFileNumber, "5")                 '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbCircle")        '--> 2002
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aXcoord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aYcoord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")
        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aRadius)))

        If aColor <> 0 Then
            gWriteLine(aFileNumber, "62")
            gWriteLine(aFileNumber, Trim(Str(aColor)))
        End If

        'Circle #2
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "CIRCLE")
        gWriteLine(aFileNumber, "5")                 '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbCircle")        '--> 2002
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aXcoord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aYcoord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")
        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aRadius + 5)))

        If aColor <> 0 Then
            gWriteLine(aFileNumber, "62")
            gWriteLine(aFileNumber, Trim(Str(aColor)))
        End If

        'Colors     1   Red
        '           2   Yellow
        '           3   Green
        '           4   Cyan
        '           5   Blue
        '           6   Magenta
        '           7   White
        '           40  Orange

        If aShowVal = False Then
            If aShowHole = True Then
                'Don't show a value -- show the hole location
                TextHeight = 75
                HorAlign = 1
                VerAlign = 2
                gTextline(aFileNumber, aXcoord, aYcoord, TextHeight, HorAlign, _
                          VerAlign, aHole, aLayer, aColor)
            End If
        Else
            'Show a value
            If aTextHeight <> 0 Then
                TextHeight = aTextHeight
            Else
                TextHeight = 75
            End If
            HorAlign = 1
            VerAlign = 2
            gTextline(aFileNumber, aXcoord, aYcoord, TextHeight, HorAlign, _
                      VerAlign, aVal, aLayer, aColor)
        End If
    End Sub

    Public Sub gPolyline(ByVal aFileNumber As Integer, _
                         ByVal aX1Coord As Double, _
                         ByVal aY1Coord As Double, _
                         ByVal aX2Coord As Double, _
                         ByVal aY2Coord As Double, _
                         ByVal aWidth As Single, _
                         ByVal aLayer As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        '--> 2002  indicates added for AutoCAD 200, AutoCAD 2002

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "POLYLINE")
        gWriteLine(aFileNumber, "5")                 '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDb2dPolyline")    '--> 2002
        gWriteLine(aFileNumber, "66")
        gWriteLine(aFileNumber, "1")
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aWidth)))
        gWriteLine(aFileNumber, "41")
        gWriteLine(aFileNumber, Trim(Str(aWidth)))
        gWriteLine(aFileNumber, "70")
        gWriteLine(aFileNumber, "1")

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "VERTEX")
        gWriteLine(aFileNumber, "5")                 '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDb2dVertex")      '--> 2002
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aX1Coord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aY1Coord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "VERTEX")
        gWriteLine(aFileNumber, "5")                 '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDb2dVertex")      '--> 2002
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aX2Coord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aY2Coord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "SEQEND")
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
    End Sub

    Public Sub gPolycline(ByVal aFileNumber As Integer, _
                          ByVal aX1Coord As Double, _
                          ByVal aY1Coord As Double, _
                          ByVal aX2Coord As Double, _
                          ByVal aY2Coord As Double, _
                          ByVal aWidth As Single, _
                          ByVal aLayer As String, _
                          ByVal aColor As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        '--> 2002  indicates added for AutoCAD 200, AutoCAD 2002

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "POLYLINE")
        gWriteLine(aFileNumber, "5")                 '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDb2dPolyline")    '--> 2002
        gWriteLine(aFileNumber, "66")
        gWriteLine(aFileNumber, "1")
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aWidth)))
        gWriteLine(aFileNumber, "41")
        gWriteLine(aFileNumber, Trim(Str(aWidth)))
        gWriteLine(aFileNumber, "70")
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "62")
        gWriteLine(aFileNumber, Trim(Str(aColor)))


        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "VERTEX")
        gWriteLine(aFileNumber, "5")                 '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDb2dVertex")      '--> 2002
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aX1Coord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aY1Coord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "VERTEX")
        gWriteLine(aFileNumber, "5")                 '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDb2dVertex")      '--> 2002
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aX2Coord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aY2Coord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "SEQEND")
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
    End Sub

    Public Sub gTextMline(ByVal aFileNumber As Integer, _
                          ByVal aX1Coord As Double, _
                          ByVal aY1Coord As Double, _
                          ByVal aHeight As Single, _
                          ByVal aTxtAlign As Single, _
                          ByVal aText As String, _
                          ByVal aLayer As String, _
                          ByVal aColor As Integer)

        '**********************************************************************
        'This subroutine does not work in AutoCAD 2002 yet.
        '
        '
        '**********************************************************************

        '--> 2002  indicates added for AutoCAD 200, AutoCAD 2002

        Dim IncludeEntities As Boolean
        IncludeEntities = True

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "MTEXT")

        If IncludeEntities = True Then
            gWriteLine(aFileNumber, "5")                 '--> 2002
            gEntityNumber = gEntityNumber + 1
            gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
            gWriteLine(aFileNumber, "100")               '--> 2002
            gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        End If

        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))

        If IncludeEntities = True Then
            gWriteLine(aFileNumber, "100")               '--> 2002
            gWriteLine(aFileNumber, "AcDbMText")         '--> 2002
        End If

        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aX1Coord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aY1Coord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aHeight)))
        'gWriteLine aFileNumber, "41"
        'gWriteLine aFileNumber, "10"
        gWriteLine(aFileNumber, "71")
        gWriteLine(aFileNumber, Trim(Str(aTxtAlign)))
        gWriteLine(aFileNumber, "72")
        gWriteLine(aFileNumber, "5")
        gWriteLine(aFileNumber, "1")
        gWriteLine(aFileNumber, aText)
        gWriteLine(aFileNumber, "7")
        gWriteLine(aFileNumber, "ROMANS")

        If aColor <> 0 Then
            gWriteLine(aFileNumber, "62")
            gWriteLine(aFileNumber, Trim(Str(aColor)))
        End If
    End Sub

    Public Sub gTextline2(ByVal aFileNumber As Integer, _
                          ByVal aX1Coord As Double, _
                          ByVal aY1Coord As Double, _
                          ByVal aHeight As Single, _
                          ByVal aHall As Single, _
                          ByVal aVall As Single, _
                          ByVal aText As String, _
                          ByVal aLayer As String, _
                          ByVal aColor As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        '--> 2002  indicates added for AutoCAD 200, AutoCAD 2002

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "TEXT")
        gWriteLine(aFileNumber, "5")                 '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbText")          '--> 2002
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aX1Coord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aY1Coord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aHeight)))
        gWriteLine(aFileNumber, "1")
        gWriteLine(aFileNumber, aText)
        gWriteLine(aFileNumber, "72")
        gWriteLine(aFileNumber, Trim(Str(aHall)))
        gWriteLine(aFileNumber, "11")
        gWriteLine(aFileNumber, Trim(Str(aX1Coord)))
        gWriteLine(aFileNumber, "21")
        gWriteLine(aFileNumber, Trim(Str(aY1Coord)))
        gWriteLine(aFileNumber, "31")
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "100")
        gWriteLine(aFileNumber, "AcDbText")
        gWriteLine(aFileNumber, "73")
        gWriteLine(aFileNumber, Trim(Str(aVall)))
        gWriteLine(aFileNumber, "7")
        gWriteLine(aFileNumber, "ROMANS")

        If aColor <> 0 Then
            gWriteLine(aFileNumber, "62")
            gWriteLine(aFileNumber, Trim(Str(aColor)))
        End If
    End Sub

    Public Sub gTextline(ByVal aFileNumber As Integer, _
                         ByVal aX1Coord As Double, _
                         ByVal aY1Coord As Double, _
                         ByVal aHeight As Single, _
                         ByVal aHall As Single, _
                         ByVal aVall As Single, _
                         ByVal aText As String, _
                         ByVal aLayer As String, _
                         ByVal aColor As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        '--> 2002  indicates added for AutoCAD 200, AutoCAD 2002

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "TEXT")
        gWriteLine(aFileNumber, "5")                 '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbText")          '--> 2002
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aX1Coord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aY1Coord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")
        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aHeight)))
        gWriteLine(aFileNumber, "72")
        gWriteLine(aFileNumber, Trim(Str(aHall)))
        gWriteLine(aFileNumber, "73")
        gWriteLine(aFileNumber, Trim(Str(aVall)))
        gWriteLine(aFileNumber, "11")
        gWriteLine(aFileNumber, Trim(Str(aX1Coord)))
        gWriteLine(aFileNumber, "21")
        gWriteLine(aFileNumber, Trim(Str(aY1Coord)))
        gWriteLine(aFileNumber, "31")
        gWriteLine(aFileNumber, "0.00")
        gWriteLine(aFileNumber, "7")
        gWriteLine(aFileNumber, "ROMANS")

        If aColor <> 0 Then
            gWriteLine(aFileNumber, "62")
            gWriteLine(aFileNumber, Trim(Str(aColor)))
        End If

        gWriteLine(aFileNumber, "1")
        gWriteLine(aFileNumber, aText)
    End Sub

    Public Sub gTextcline(ByVal aFileNumber As Integer, _
                          ByVal aX1Coord As Double, _
                          ByVal aY1Coord As Double, _
                          ByVal aHeight As Single, _
                          ByVal aHall As Single, _
                          ByVal aVall As Single, _
                          ByVal aText As String, _
                          ByVal aLayer As String, _
                          ByVal aColor As Integer, _
                          ByVal aFont As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'For use in AutoCAD

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "TEXT")
        gWriteLine(aFileNumber, "5")
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))
        gWriteLine(aFileNumber, "100")
        gWriteLine(aFileNumber, "AcDbEntity")
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")
        gWriteLine(aFileNumber, "AcDbText")

        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aX1Coord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aY1Coord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")

        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aHeight)))
        gWriteLine(aFileNumber, "72")
        gWriteLine(aFileNumber, Trim(Str(aHall)))
        gWriteLine(aFileNumber, "73")
        gWriteLine(aFileNumber, Trim(Str(aVall)))

        gWriteLine(aFileNumber, "11")
        gWriteLine(aFileNumber, Trim(Str(aX1Coord)))
        gWriteLine(aFileNumber, "21")
        gWriteLine(aFileNumber, Trim(Str(aY1Coord)))
        gWriteLine(aFileNumber, "31")
        gWriteLine(aFileNumber, "0.00")

        gWriteLine(aFileNumber, "7")
        gWriteLine(aFileNumber, Trim(aFont))

        If aColor <> 0 Then
            gWriteLine(aFileNumber, "62")
            gWriteLine(aFileNumber, Trim(Str(aColor)))
        End If

        gWriteLine(aFileNumber, "1")
        gWriteLine(aFileNumber, aText)
    End Sub

    Public Sub gTextclineL(ByVal aFileNumber As Integer, _
                           ByVal aX1Coord As Double, _
                           ByVal aY1Coord As Double, _
                           ByVal aHeight As Single, _
                           ByVal aText As String, _
                           ByVal aLayer As String, _
                           ByVal aColor As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'For use by InViso

        'Orientation is fixed 73 = 0, 72 = 0

        'Vertical alignment   (73) --> TLeft
        'Horizontal alignment (72) --> Left (Baseline)

        '      |
        '      |
        '      |
        '    x |_______

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "TEXT")
        gWriteLine(aFileNumber, "5")                 '--> 2002
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))  '--> 2002
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbEntity")        '--> 2002
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")               '--> 2002
        gWriteLine(aFileNumber, "AcDbText")          '--> 2002

        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(Str(aX1Coord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(Str(aY1Coord)))
        gWriteLine(aFileNumber, "30")
        gWriteLine(aFileNumber, "0.00")
        gWriteLine(aFileNumber, "40")
        gWriteLine(aFileNumber, Trim(Str(aHeight)))

        gWriteLine(aFileNumber, "1")
        gWriteLine(aFileNumber, aText)

        gWriteLine(aFileNumber, "72")
        gWriteLine(aFileNumber, "0")

        gWriteLine(aFileNumber, "11")
        gWriteLine(aFileNumber, Trim(Str(aX1Coord)))
        gWriteLine(aFileNumber, "21")
        gWriteLine(aFileNumber, Trim(Str(aY1Coord)))
        gWriteLine(aFileNumber, "31")
        gWriteLine(aFileNumber, "0.00")

        gWriteLine(aFileNumber, "100")
        gWriteLine(aFileNumber, "AcDbText")

        gWriteLine(aFileNumber, "73")
        gWriteLine(aFileNumber, "0")

        If aColor <> 0 Then
            gWriteLine(aFileNumber, "62")
            gWriteLine(aFileNumber, Trim(Str(aColor)))
        End If
    End Sub

    Public Sub gOpenDxf(ByVal aFileNumber As Integer, _
                        ByVal aAcadVer As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        gEntityNumber = 32

        Select Case aAcadVer
            Case Is = "AutoCAD 14"
                gWriteLine(aFileNumber, "0")
                gWriteLine(aFileNumber, "SECTION")
                gWriteLine(aFileNumber, "2")
                gWriteLine(aFileNumber, "HEADER")
                gWriteLine(aFileNumber, "9")
                gWriteLine(aFileNumber, "$ACADVER")
                gWriteLine(aFileNumber, "1")
                gWriteLine(aFileNumber, "AC1009")
                gWriteLine(aFileNumber, "0")
                gWriteLine(aFileNumber, "ENDSEC")
                gWriteLine(aFileNumber, "0")
                gWriteLine(aFileNumber, "SECTION")
                gWriteLine(aFileNumber, "2")
                gWriteLine(aFileNumber, "ENTITIES")

            Case Is = "AutoCAD 2000"
                gWriteLine(aFileNumber, "0")
                gWriteLine(aFileNumber, "SECTION")
                gWriteLine(aFileNumber, "2")
                gWriteLine(aFileNumber, "HEADER")
                gWriteLine(aFileNumber, "9")
                gWriteLine(aFileNumber, "$ACADVER")
                gWriteLine(aFileNumber, "1")
                gWriteLine(aFileNumber, "AC1015")    'AC1015 ?
                gWriteLine(aFileNumber, "9")

                gWriteLine(aFileNumber, "$ACADMAINTVER")
                gWriteLine(aFileNumber, "70")
                gWriteLine(aFileNumber, "6")
                gWriteLine(aFileNumber, "9")

                gWriteLine(aFileNumber, "ENDSEC")
                gWriteLine(aFileNumber, "0")
                gWriteLine(aFileNumber, "SECTION")
                gWriteLine(aFileNumber, "2")
                gWriteLine(aFileNumber, "ENTITIES")

            Case Is = "AutoCAD 2002"
                gWriteLine(aFileNumber, "0")
                gWriteLine(aFileNumber, "SECTION")
                gWriteLine(aFileNumber, "2")
                gWriteLine(aFileNumber, "HEADER")
                gWriteLine(aFileNumber, "9")
                gWriteLine(aFileNumber, "$ACADVER")
                gWriteLine(aFileNumber, "1")
                gWriteLine(aFileNumber, "AC1015")
                gWriteLine(aFileNumber, "9")

                gWriteLine(aFileNumber, "$ACADMAINTVER")
                gWriteLine(aFileNumber, "70")
                gWriteLine(aFileNumber, "13")
                gWriteLine(aFileNumber, "9")

                gWriteLine(aFileNumber, "ENDSEC")
                gWriteLine(aFileNumber, "0")
                gWriteLine(aFileNumber, "SECTION")
                gWriteLine(aFileNumber, "2")
                gWriteLine(aFileNumber, "ENTITIES")

            Case Else
                gWriteLine(aFileNumber, "0")
                gWriteLine(aFileNumber, "SECTION")
                gWriteLine(aFileNumber, "2")
                gWriteLine(aFileNumber, "HEADER")
                gWriteLine(aFileNumber, "9")
                gWriteLine(aFileNumber, "$ACADVER")
                gWriteLine(aFileNumber, "1")
                gWriteLine(aFileNumber, "AC1009")
                gWriteLine(aFileNumber, "0")
                gWriteLine(aFileNumber, "ENDSEC")
                gWriteLine(aFileNumber, "0")
                gWriteLine(aFileNumber, "SECTION")
                gWriteLine(aFileNumber, "2")
                gWriteLine(aFileNumber, "ENTITIES")
        End Select
    End Sub

    Public Sub gEntitiesDxf(ByVal aFileNumber As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************
        gEntityNumber = 32

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "SECTION")
        gWriteLine(aFileNumber, "2")
        gWriteLine(aFileNumber, "ENTITIES")
    End Sub

    Public Sub gCloseDxf(ByVal aFileNumber As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "ENDSEC")
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "EOF")
    End Sub

    Public Function gCreateProspectGrid(ByVal aSec As Integer, _
                                        ByVal aTwp As Integer, _
                                        ByVal aRge As Integer, _
                                        ByVal aFileNumber As Integer, _
                                        ByVal aGridInDxf As Boolean, _
                                        ByVal aMineName As String, _
                                        ByVal aProspGridLayer As String, _
                                        ByVal aIncludeDate As Boolean, _
                                        ByVal aColor As Integer, _
                                        ByVal aSecBdryWidth As String, _
                                        ByVal aAddHoleLocs As Boolean, _
                                        ByVal aIncludeAreas As Boolean) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo CreateProspectGridError

        'Cargill stel prospect grid.
        'Divide section lines into sixteen equal size segments and connect the dots!

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer

        Dim RowCount As Integer
        Dim ColCount As Integer

        Dim NWx As Double
        Dim NWy As Double
        Dim NEx As Double
        Dim NEy As Double
        Dim SEx As Double
        Dim SEy As Double
        Dim SWx As Double
        Dim SWy As Double

        Dim Bot(0 To 33, 0 To 2)
        Dim Top(0 To 33, 0 To 2)
        Dim Left(0 To 33, 0 To 2)
        Dim Right(0 To 33, 0 To 2)

        Dim Dist As Double
        Dim Angle As Single
        Dim IncDist As Double
        Dim AccumDist As Double
        Dim X As Integer
        Dim Y As Integer
        Dim Pos As Integer
        Dim Pos2 As Integer
        Dim Line As Integer
        Dim X1 As Double
        Dim Y1 As Double
        Dim X2 As Double
        Dim Y2 As Double
        Dim MapLayer As String
        Dim PolylineThick As Single
        Dim PolylineColor As Integer

        Dim SecStr As String
        Dim TwpStr As String
        Dim RgeStr As String

        Dim AddCircles As Boolean
        Dim HoleLoc As String
        Dim CharPart As String
        Dim NumPart As String
        Dim Alphabet As String

        Dim Radius As Single

        Dim VerAlign As Single
        Dim HorAlign As Single
        Dim TextHeight As Single
        Dim TextColor As Integer

        Dim SkipLine As Boolean

        Dim Ax As Double
        Dim Ay As Double
        Dim Bx As Double
        Dim By As Double
        Dim Cx As Double
        Dim Cy As Double
        Dim Dx As Double
        Dim Dy As Double
        Dim CellAreaFt As Single
        Dim CellAreaAcres As Single

        Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

        AddCircles = False

        If aSec > 9 Then
            SecStr = Trim(Str(aSec))
        Else
            SecStr = "0" & Trim(Str(aSec))
        End If
        If aTwp > 9 Then
            TwpStr = Trim(Str(aTwp))
        Else
            TwpStr = "0" & Trim(Str(aTwp))
        End If
        If aRge > 9 Then
            RgeStr = Trim(Str(aRge))
        Else
            RgeStr = "0" & Trim(Str(aRge))
        End If

        'Need to get the state-planar coordinates for this section.
        'They should be in the table SECTN_COORDS.

        'Get section state-planar coordinates
        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_sectn_coords
        'pMineName
        'pSection
        'pTownship
        'pRange
        'pResult

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_survey.get_sectn_coords(:pMineName," + _
                      ":pSection, :pTownship, :pRange, :pResult);end;", ORASQL_FAILEXEC)
        mSectnCoordsDynaset = params("pResult").Value

        RecordCount = mSectnCoordsDynaset.RecordCount
        ClearParams(params)

        If RecordCount <> 1 Then
            gCreateProspectGrid = False
            Exit Function
        End If

        'We now have the state-planar coordinates for this section
        mSectnCoordsDynaset.MoveFirst()
        NWx = mSectnCoordsDynaset.Fields("nw_x_cdnt").Value
        NWy = mSectnCoordsDynaset.Fields("nw_y_cdnt").Value
        NEx = mSectnCoordsDynaset.Fields("ne_x_cdnt").Value
        NEy = mSectnCoordsDynaset.Fields("ne_y_cdnt").Value
        SWx = mSectnCoordsDynaset.Fields("sw_x_cdnt").Value
        SWy = mSectnCoordsDynaset.Fields("sw_y_cdnt").Value
        SEx = mSectnCoordsDynaset.Fields("se_x_cdnt").Value
        SEy = mSectnCoordsDynaset.Fields("se_y_cdnt").Value

        'Create the prospect grid
        'Bottom line
        Dist = gGetDistance(SWx, SWy, SEx, SEy)
        Angle = gGetAngle(SWx, SWy, SEx, SEy, "X")
        IncDist = Round(Dist / 32, 4)
        Bot(1, 1) = SWx
        Bot(1, 2) = SWy
        Bot(33, 1) = SEx
        Bot(33, 2) = SEy
        For Pos = 1 To 31
            AccumDist = Round(Pos * IncDist, 4)
            gGetNewCoords(SWx, SWy, SEx, SEy, Angle, AccumDist, "X")
            Bot(Pos + 1, 1) = mXcoord
            Bot(Pos + 1, 2) = mYcoord
        Next Pos

        'Top line
        Dist = gGetDistance(NWx, NWy, NEx, NEy)
        Angle = gGetAngle(NWx, NWy, NEx, NEy, "X")
        IncDist = Round(Dist / 32, 4)
        Top(1, 1) = NWx
        Top(1, 2) = NWy
        Top(33, 1) = NEx
        Top(33, 2) = NEy
        For Pos = 1 To 31
            AccumDist = Round(Pos * IncDist, 4)
            gGetNewCoords(NWx, NWy, NEx, NEy, Angle, AccumDist, "X")
            Top(Pos + 1, 1) = mXcoord
            Top(Pos + 1, 2) = mYcoord
        Next Pos

        'Left line
        Dist = gGetDistance(SWx, SWy, NWx, NWy)
        Angle = gGetAngle(SWx, SWy, NWx, NWy, "Y")
        IncDist = Round(Dist / 32, 4)
        Left(1, 1) = SWx
        Left(1, 2) = SWy
        Left(33, 1) = NWx
        Left(33, 2) = NWy
        For Pos = 1 To 31
            AccumDist = Round(Pos * IncDist, 4)
            gGetNewCoords(SWx, SWy, NWx, NWy, Angle, AccumDist, "Y")
            Left(Pos + 1, 1) = mXcoord
            Left(Pos + 1, 2) = mYcoord
        Next Pos

        'Right line
        Dist = gGetDistance(SEx, SEy, NEx, NEy)
        Angle = gGetAngle(SEx, SEy, NEx, NEy, "Y")
        IncDist = Round(Dist / 32, 4)
        Right(1, 1) = SEx
        Right(1, 2) = SEy
        Right(33, 1) = NEx
        Right(33, 2) = NEy
        For Pos = 1 To 31
            AccumDist = Round(Pos * IncDist, 4)
            gGetNewCoords(SEx, SEy, NEx, NEy, Angle, AccumDist, "Y")
            Right(Pos + 1, 1) = mXcoord
            Right(Pos + 1, 2) = mYcoord
        Next Pos

        'Which layer to use for prospect grid?
        If Len(Trim(aProspGridLayer)) = 0 Then
            'Use default layer
            MapLayer = SecStr + TwpStr + RgeStr + "g"
        Else
            MapLayer = aProspGridLayer
        End If

        PolylineThick = 0
        PolylineColor = aColor   'Red

        'aSecBdryWidth -- "Thick", "Thin", "None"

        If aGridInDxf = True Then
            'Create vertical lines -- draw from bottom to top
            For Line = 1 To 33 Step 2
                If Line = 1 Or Line = 33 Then
                    Select Case aSecBdryWidth
                        Case Is = "Thick"
                            PolylineThick = 30
                            SkipLine = False
                        Case Is = "None"
                            SkipLine = True
                        Case Else
                            PolylineThick = 0
                            SkipLine = False
                    End Select
                Else
                    PolylineThick = 0
                End If

                If SkipLine = False Then
                    X1 = Bot(Line, 1)
                    Y1 = Bot(Line, 2)
                    X2 = Top(Line, 1)
                    Y2 = Top(Line, 2)
                    'Draw a red polyline
                    gPolycline(aFileNumber, X1, Y1, X2, Y2, PolylineThick, MapLayer, _
                               PolylineColor)
                End If
            Next Line

            'Create horizontal lines -- draw from left to right
            For Line = 1 To 33 Step 2
                If Line = 1 Or Line = 33 Then
                    Select Case aSecBdryWidth
                        Case Is = "Thick"
                            PolylineThick = 30
                            SkipLine = False
                        Case Is = "None"
                            SkipLine = True
                        Case Else
                            PolylineThick = 0
                            SkipLine = False
                    End Select
                Else
                    PolylineThick = 0
                End If

                X1 = Left(Line, 1)
                Y1 = Left(Line, 2)
                X2 = Right(Line, 1)
                Y2 = Right(Line, 2)
                'Draw a red polyline
                gPolycline(aFileNumber, X1, Y1, X2, Y2, PolylineThick, MapLayer, _
                           PolylineColor)
            Next Line
        End If

        'Fill in gProspX(1 To 16, 1 To 16)
        'Fill in gProspY(1 To 16, 1 To 16)
        For Pos = 2 To 32 Step 2    'Process 16 east-west lines
            X1 = Left(Pos, 1)
            Y1 = Left(Pos, 2)
            X2 = Right(Pos, 1)
            Y2 = Right(Pos, 2)
            Dist = gGetDistance(X1, Y1, X2, Y2)
            Angle = gGetAngle(X1, Y1, X2, Y2, "X")
            IncDist = Round(Dist / 32, 4)
            For Pos2 = 2 To 32 Step 2
                AccumDist = Round((Pos2 - 1) * IncDist, 4)
                gGetNewCoords(X1, Y1, X2, Y2, Angle, AccumDist, "X")
                gProspX(Pos2 / 2, Pos / 2) = mXcoord
                gProspY(Pos2 / 2, Pos / 2) = mYcoord
            Next Pos2
        Next Pos

        'Fill in gProspIntX(1 To 17, 1 To 17)
        'Fill in gProspIntY(1 To 17, 1 To 17)
        For Pos = 1 To 33 Step 2    'Process 16 east-west lines
            X1 = Left(Pos, 1)
            Y1 = Left(Pos, 2)
            X2 = Right(Pos, 1)
            Y2 = Right(Pos, 2)
            Dist = gGetDistance(X1, Y1, X2, Y2)
            Angle = gGetAngle(X1, Y1, X2, Y2, "X")
            IncDist = Round(Dist / 16, 4)

            'Set first position
            gProspIntX(1, (Pos + 1) / 2) = X1
            gProspIntY(1, (Pos + 1) / 2) = Y1

            For Pos2 = 2 To 16
                'Get first position
                AccumDist = Round((Pos2 - 1) * IncDist, 4)
                gGetNewCoords(X1, Y1, X2, Y2, Angle, AccumDist, "X")
                gProspIntX(Pos2, (Pos + 1) / 2) = mXcoord
                gProspIntY(Pos2, (Pos + 1) / 2) = mYcoord
            Next Pos2

            'For Pos2 = 2 To 32 Step 2
            '    'Get first position
            '    If Pos = 2 Then
            '        gProspIntX(1, Pos2 / 2) = X1
            '        gProspIntY(1, Pos2 / 2) = Y1
            '     End If
            '
            '    AccumDist = Round((Pos2 / 2) * IncDist, 4)
            '    gGetNewCoords X1, Y1, X2, Y2, Angle, AccumDist, "X"
            '    gProspIntX(Pos2 / 2, Pos / 2) = mXcoord
            '    gProspIntY(Pos2 / 2, Pos / 2) = mYcoord
            'Next Pos2

            'Get last position
            gProspIntX(17, (Pos + 1) / 2) = X2
            gProspIntY(17, (Pos + 1) / 2) = Y2
        Next Pos

        If aIncludeAreas = True Then
            'Alpha-numeric grids only -- will be 16 X 16.
            'gProspIntX(1 To 17, 1 To 17)
            'gProspIntY(1 To 17, 1 To 17)

            For X = 2 To 17     'Will be processing 16 cells in a 16 rows
                For Y = 2 To 17
                    Ax = gProspIntX(X - 1, Y - 1)
                    Ay = gProspIntY(X - 1, Y - 1)
                    '-----
                    Bx = gProspIntX(X - 1, Y)
                    By = gProspIntY(Y - 1, Y)
                    '-----
                    Cx = gProspIntX(X, Y)
                    Cy = gProspIntY(X, Y)
                    '-----
                    Dx = gProspIntX(X, Y - 1)
                    Dy = gProspIntY(X, Y - 1)

                    CellAreaFt = gGetQuadrilateralArea(Ax, Ay, Bx, By, _
                                 Cx, Cy, Dx, Dy, 2)
                    CellAreaAcres = Round(CellAreaFt / 43560, 2)

                    X1 = gProspX(X - 1, Y - 1) + 25
                    Y1 = gProspY(X - 1, Y - 1) - 100

                    TextHeight = 30
                    TextColor = 0   'Black
                    HorAlign = 0    'Left
                    VerAlign = 2    'Middle
                    gTextcline(aFileNumber, X1, Y1, TextHeight, HorAlign, _
                        VerAlign, Format(CellAreaAcres, "#0.00"), MapLayer, TextColor, "ROMANS")
                Next Y
            Next X
        End If

        If aGridInDxf Then
            If aIncludeDate Then
                'Add in today's date in lower left hand corner
                TextHeight = 10
                TextColor = 5   'Blue
                HorAlign = 0    'Left
                VerAlign = 2    'Middle
                X1 = gProspX(1, 1) - 160
                Y1 = gProspY(1, 1) - 160

                gTextcline(aFileNumber, X1, Y1, TextHeight, HorAlign, _
                           VerAlign, CStr(Today), MapLayer, TextColor, "ROMANS")
            End If
        End If

        If AddCircles = True Then
            For X = 1 To 16
                For Y = 1 To 16
                    X1 = gProspX(X, Y)
                    Y1 = gProspY(X, Y)
                    Radius = 40
                    gMakeCircle(aFileNumber, X1, Y1, Radius, MapLayer)
                Next Y
            Next X
        End If

        If aAddHoleLocs = True Then
            For X = 1 To 16
                For Y = 1 To 16
                    TextHeight = 40
                    TextColor = 5   'Blue
                    HorAlign = 0    'Left
                    VerAlign = 2    'Middle
                    X1 = gProspX(X, Y) - 120
                    Y1 = gProspY(X, Y) - 100

                    CharPart = Mid(Alphabet, X, 1)
                    If Y <= 9 Then
                        NumPart = "0" & Trim(CStr(Y))
                    Else
                        NumPart = Trim(CStr(Y))
                    End If
                    HoleLoc = CharPart & NumPart

                    gTextcline(aFileNumber, X1, Y1, TextHeight, HorAlign, _
                               VerAlign, HoleLoc, MapLayer, TextColor, "ROMANS")
                Next Y
            Next X
        End If

        gCreateProspectGrid = True
        Exit Function

CreateProspectGridError:
        MsgBox("Error creating prospect box." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Prospect Box Add Error")

    End Function

    Private Sub gGetNewCoords(ByVal aX1 As Double, _
                              ByVal aY1 As Double, _
                              ByVal aX2 As Double, _
                              ByVal aY2 As Double, _
                              ByVal aAngle As Single, _
                              ByVal aAccumDist As Double, _
                              ByVal aAxis As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'Need to assign form variables: mXcoord and mYcoord

        Dim Xdiff As Single
        Dim Ydiff As Single
        Dim CorrFactor As Integer

        CorrFactor = 1
        If aAxis = "X" Then
            If aY2 < aY1 Then
                CorrFactor = -1
            End If
        End If
        If aAxis = "Y" Then
            If aX2 < aX1 Then
                CorrFactor = -1
            End If
        End If

        If aAxis = "X" Then
            Xdiff = Round((Cos(aAngle) * aAccumDist), 4)
            Ydiff = Round((Sin(aAngle) * aAccumDist), 4) * CorrFactor
        End If
        If aAxis = "Y" Then
            Ydiff = Round((Cos(aAngle) * aAccumDist), 4)
            Xdiff = Round((Sin(aAngle) * aAccumDist), 4) * CorrFactor
        End If

        mXcoord = aX1 + Xdiff
        mYcoord = aY1 + Ydiff
    End Sub

    Private Function gGetDistance(ByVal aX1 As Double, _
                                  ByVal aY1 As Double, _
                                  ByVal aX2 As Double, _
                                  ByVal aY2 As Double) As Double

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim Xdiff As Double
        Dim Ydiff As Double

        Xdiff = Abs(aX2 - aX1)
        Ydiff = Abs(aY2 - aY1)

        gGetDistance = Sqrt(Xdiff ^ 2 + Ydiff ^ 2)
    End Function

    Private Function gGetAngle(ByVal aX1 As Double, _
                               ByVal aY1 As Double, _
                               ByVal aX2 As Double, _
                               ByVal aY2 As Double, _
                               ByVal aAxis As String) As Single

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim CoTan As Single
        Dim Tan As Single
        Dim AngleInRadians As Single
        Dim AngleInDegrees As Single
        Dim Xdiff As Double
        Dim Ydiff As Double
        Dim PiVal As Single

        PiVal = 4 * Atan(1.0#)

        Xdiff = Abs(aX2 - aX1)
        Ydiff = Abs(aY2 - aY1)

        If Xdiff <> 0 Then
            If aAxis = "X" Then
                Tan = Round((Ydiff / Xdiff), 5)
            End If
            If aAxis = "Y" Then
                Tan = Round((Xdiff / Ydiff), 5)
            End If
        Else
            Tan = 0
        End If

        'ArcTangent -- Angle in radians whose tangent = Tan
        AngleInRadians = Atan(Tan)

        'Degrees = radians * 180/pi
        'Radians = degrees * pi/180
        AngleInDegrees = Round(AngleInRadians * 180 / PiVal, 5)

        gGetAngle = AngleInRadians

    End Function

    Public Sub gViewAutoCadMap(ByVal aBaseMap As String, _
                               ByVal aDxfFile As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        '03/15/07, lss
        'Removed this functionality for now -- don't want to deal with
        'AutoCAD 2002, AutoCAD 2007 issues!
        MsgBox("Functionality no longer available!", vbOKOnly, _
               "Functionality Status")

        'Dim objAcadApp As AcadApplication
        'Dim objAcadDoc As AcadDocument
        'Dim InsertPoint(0 To 2) As Double
        'Dim ScaleFactor As Double
        'Dim ImportFile As String

        'InsertPoint(0) = 0
        'InsertPoint(1) = 0
        'InsertPoint(2) = 0
        'ScaleFactor = 1

        'ImportFile = aDxfFile

        'On Error Resume Next

        'IMPORTANT:  The LT versions of AutoCAD DO NOT SUPPORT
        '            customization -- ie. they do not suppport
        '            the ActiveX based code below -- for example
        '            AutoCAD LT 2002 will not work with this
        '            code!  (02/25/2004, LSS)

        'Set objAcadApp = GetObject(, "AutoCAD.Application")
        'objAcadApp.Visible = True

        'If Err Then
        '    Err.Clear

        '    Start AutoCAD if it is not running.
        '    Set objAcadApp = CreateObject("AutoCAD.Application")
        '    objAcadApp.Visible = True

        '    If Err Then
        '        MsgBox Err.Description
        '        Exit Sub
        '    End If
        'End If

        'On Error GoTo cmdViewAutoCadMapClickError

        'Get the base map selected by the user
        'Set objAcadDoc = objAcadApp.Documents.Open(aBaseMap)

        'Import .dxf file created by the user
        'objAcadApp.ActiveDocument.Import ImportFile, InsertPoint, ScaleFactor
        'objAcadDoc.SendCommand "_zoom" & vbCr & "E" & vbCr

        'Switch focus
        'AppActivate objAcadApp.Caption

        'Exit Sub

        'cmdViewAutoCadMapClickError:

        '  MsgBox "Error accessing AutoCAD." + str(Err.number) + Chr$(10) + Chr$(10) + _
        '      Err.Description, vbExclamation, "Error Accessing AutoCAD"
    End Sub

    Public Function gGetAcadColor(ByVal aColor) As Integer

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'AutoCAD colors
        '1   Red
        '2   Yellow
        '3   Green
        '4   Cyan
        '5   Blue
        '6   Magenta
        '7   White
        '40  Orange

        Select Case aColor
            Case Is = "Red"
                gGetAcadColor = 1

            Case Is = "Yellow"
                gGetAcadColor = 2

            Case Is = "Green"
                gGetAcadColor = 3

            Case Is = "Cyan"
                gGetAcadColor = 4

            Case Is = "Blue"
                gGetAcadColor = 5

            Case Is = "Magenta"
                gGetAcadColor = 6

            Case Is = "White"
                gGetAcadColor = 7

            Case Is = "Orange"
                gGetAcadColor = 40

            Case Is = "Black"
                gGetAcadColor = 0
        End Select
    End Function

    Public Sub gMakeCompBox(ByVal aFnum As Integer, _
                            ByVal Ax As Double, _
                            ByVal Ay As Double, _
                            ByVal aXoff As Double, _
                            ByVal aYoff As Double, _
                            ByVal aDrwLayer As String, _
                            ByVal aScale As Single, _
                            ByVal aMode As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim Color As Integer
        Dim Xcorr As Double
        Dim Ycorr As Double
        Dim TextHeight As Single

        'Draw main box
        gPolyline(aFnum, Ax, Ay, Ax, _
                  Ay + (79 * aScale), 0, aDrwLayer)

        gPolyline(aFnum, Ax, Ay + (79 * aScale), _
                  Ax + (132 * aScale), _
                  Ay + (79 * aScale), 0, aDrwLayer)

        gPolyline(aFnum, Ax + (132 * aScale), Ay + (79 * aScale), _
                  Ax + (132 * aScale), Ay, 0, aDrwLayer)

        gPolyline(aFnum, Ax + (132 * aScale), Ay, Ax, _
                  Ay, 0, aDrwLayer)

        'Draw individual lines
        'Draw vertical lines
        gPolyline(aFnum, Ax + (18.25 * aScale), Ay, Ax + (18.25 * aScale), _
                  Ay + (43.75 * aScale), 0, aDrwLayer)

        gPolyline(aFnum, Ax + (44.75 * aScale), Ay, Ax + (44.75 * aScale), _
                  Ay + (43.75 * aScale), 0, aDrwLayer)

        gPolyline(aFnum, Ax + (63 * aScale), Ay, Ax + (63 * aScale), _
                  Ay + (43.75 * aScale), 0, aDrwLayer)

        gPolyline(aFnum, Ax + (86 * aScale), Ay, Ax + (86 * aScale), _
                  Ay + (43.75 * aScale), 0, aDrwLayer)

        gPolyline(aFnum, Ax + (109 * aScale), Ay, Ax + (109 * aScale), _
                  Ay + (43.75 * aScale), 0, aDrwLayer)

        'Draw horizontal lines
        gPolyline(aFnum, Ax, Ay + (8.75 * aScale), Ax + (132 * aScale), _
                  Ay + (8.75 * aScale), 0, aDrwLayer)

        gPolyline(aFnum, Ax, Ay + (17.5 * aScale), Ax + (132 * aScale), _
                  Ay + (17.5 * aScale), 0, aDrwLayer)

        gPolyline(aFnum, Ax, Ay + (26.25 * aScale), Ax + (132 * aScale), _
                  Ay + (26.25 * aScale), 0, aDrwLayer)

        gPolyline(aFnum, Ax, Ay + (35 * aScale), Ax + (132 * aScale), _
                  Ay + (35 * aScale), 0.7 * aScale, aDrwLayer)

        gPolyline(aFnum, Ax, Ay + (43.75 * aScale), Ax + (132 * aScale), _
                  Ay + (43.75 * aScale), 0.7 * aScale, aDrwLayer)

        gPolyline(aFnum, Ax, Ay + (70 * aScale), Ax + (132 * aScale), _
                  Ay + (70 * aScale), 0.7 * aScale, aDrwLayer)

        'Draw text lines -- row headers
        'Draw text lines -- row headers
        'Draw text lines -- row headers

        TextHeight = 3.5 * aScale

        If aMode = "AutoCAD" Then
            'AutoCAD alignment

            gTextcline(aFnum, Ax + (9 * aScale), Ay + (4.375 * aScale), _
                       TextHeight, 1, 2, "TPrd", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (9 * aScale), Ay + (13.125 * aScale), _
                       TextHeight, 1, 2, "Conc", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (9 * aScale), Ay + (21.875 * aScale), _
                       TextHeight, 1, 2, "Feed", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (9 * aScale), Ay + (30.625 * aScale), _
                       TextHeight, 1, 2, "Pebb", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (10 * aScale), Ay + (56.875 * aScale), _
                       TextHeight, 1, 2, "Mtx", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (10 * aScale), Ay + (65.625 * aScale), _
                       TextHeight, 1, 2, "Ovb", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (31.5 * aScale), Ay + (39.375 * aScale), _
                       TextHeight, 1, 2, "TPA", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (53.875 * aScale), Ay + (39.375 * aScale), _
                       TextHeight, 1, 2, "BPL", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (74.5 * aScale), Ay + (39.375 * aScale), _
                       TextHeight, 1, 2, "Fe", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (97.5 * aScale), Ay + (39.375 * aScale), _
                       TextHeight, 1, 2, "Al", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (120.5 * aScale), Ay + (39.375 * aScale), _
                       TextHeight, 1, 2, "Mg", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (68 * aScale), Ay + (65.625 * aScale), _
                       TextHeight, 1, 2, "MtxX", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (68 * aScale), Ay + (56.875 * aScale), _
                       TextHeight, 1, 2, "TotX", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (68 * aScale), Ay + (48.125 * aScale), _
                       TextHeight, 1, 2, "%Cly", aDrwLayer, 0, "ROMANS")
        Else
            'InViso alignment
            Ycorr = 0.5 * TextHeight

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (9 * aScale) - Xcorr, Ay + (4.375 * aScale) - Ycorr, _
                        TextHeight, "TPrd", aDrwLayer, 0)

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (9 * aScale) - Xcorr, Ay + (13.125 * aScale) - Ycorr, _
                        TextHeight, "Conc", aDrwLayer, 0)

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (9 * aScale) - Xcorr, Ay + (21.875 * aScale) - Ycorr, _
                        TextHeight, "Feed", aDrwLayer, 0)

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (9 * aScale) - Xcorr, Ay + (30.625 * aScale) - Ycorr, _
                        TextHeight, "Pebb", aDrwLayer, 0)

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (10 * aScale) - Xcorr, Ay + (56.875 * aScale) - Ycorr, _
                        TextHeight, "Mtx", aDrwLayer, 0)

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (10 * aScale) - Xcorr, Ay + (65.625 * aScale) - Ycorr, _
                        TextHeight, "Ovb", aDrwLayer, 0)

            Xcorr = 0.06
            gTextclineL(aFnum, Ax + (31.5 * aScale) - Xcorr, Ay + (39.375 * aScale) - Ycorr, _
                        TextHeight, "TPA", aDrwLayer, 0)

            Xcorr = 0.06
            gTextclineL(aFnum, Ax + (53.875 * aScale) - Xcorr, Ay + (39.375 * aScale) - Ycorr, _
                        TextHeight, "BPL", aDrwLayer, 0)

            Xcorr = 0.03
            gTextclineL(aFnum, Ax + (74.5 * aScale) - Xcorr, Ay + (39.375 * aScale) - Ycorr, _
                        TextHeight, "Fe", aDrwLayer, 0)

            Xcorr = 0.03
            gTextclineL(aFnum, Ax + (97.5 * aScale) - Xcorr, Ay + (39.375 * aScale) - Ycorr, _
                        TextHeight, "Al", aDrwLayer, 0)

            Xcorr = 0.03
            gTextclineL(aFnum, Ax + (120.5 * aScale) - Xcorr, Ay + (39.375 * aScale) - Ycorr, _
                        TextHeight, "Mg", aDrwLayer, 0)

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (68 * aScale) - Xcorr, Ay + (65.625 * aScale) - Ycorr, _
                        TextHeight, "MtxX", aDrwLayer, 0)

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (68 * aScale) - Xcorr, Ay + (56.875 * aScale) - Ycorr, _
                        TextHeight, "TotX", aDrwLayer, 0)

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (68 * aScale) - Xcorr, Ay + (48.125 * aScale) - Ycorr, _
                        TextHeight, "%Cly", aDrwLayer, 0)
        End If

        Color = 0   'Black
        AddCompData(aFnum, Ax, Ay, aDrwLayer, aScale, Color, aMode)
    End Sub

    Private Sub AddCompData(ByVal aFnum As Integer, _
                            ByVal Ax As Double, _
                            ByVal Ay As Double, _
                            ByVal aDrwLayer As String, _
                            ByVal aScale As Single, _
                            ByVal aColor As String, _
                            ByVal aMode As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim Location As String
        Dim SpecColor As Integer
        Dim TextHeight As Single
        Dim Xcorr As Double
        Dim Ycorr As Double

        TextHeight = 4.25 * aScale

        If aMode = "AutoCAD" Then
            'AutoCAD alignment

            'Add hole name
            gTextcline(aFnum, Ax + (7 * aScale), Ay + (74.5 * aScale), _
                       TextHeight, 0, 2, gComposite.HoleLocation, _
                       aDrwLayer, aColor, "ROMANS")

            'Add Section, Township, Range
            Location = gGetHoleLocationShortDot(gComposite.Section, gComposite.Township, _
                       gComposite.Range)

            gTextcline(aFnum, Ax + (30 * aScale), Ay + (74.5 * aScale), _
                       TextHeight, 0, 2, Location, _
                       aDrwLayer, aColor, "ROMANS")

            'Add drill date
            gTextcline(aFnum, Ax + (70 * aScale), Ay + (74.5 * aScale), _
                       TextHeight, 0, 2, gComposite.DrillCdate, _
                       aDrwLayer, aColor, "ROMANS")

            'Add overburden thickness
            SpecColor = 5   'Blue
            gTextcline(aFnum, Ax + (33 * aScale), Ay + (65.625 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.OvbThickness, "##0.0"), _
                       aDrwLayer, SpecColor, "ROMANS")

            'Add matrix thickness
            SpecColor = 3  'Green
            gTextcline(aFnum, Ax + (33 * aScale), Ay + (56.875 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.MtxThickness, "##0.0"), _
                       aDrwLayer, SpecColor, "ROMANS")

            'Add MatrixX
            gTextcline(aFnum, Ax + (93 * aScale), Ay + (65.625 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.MtxX, "##0.0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Add TotalX
            gTextcline(aFnum, Ax + (93 * aScale), Ay + (56.875 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalX, "##0.0"), _
                       aDrwLayer, aColor, "ROMANS")

            '%Clay
            gTextcline(aFnum, Ax + (93 * aScale), Ay + (48.125 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.WasteClayWtp, "##0.0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Pebble  Pebble  Pebble  Pebble  Pebble  Pebble  Pebble
            'Pebble  Pebble  Pebble  Pebble  Pebble  Pebble  Pebble
            'Pebble  Pebble  Pebble  Pebble  Pebble  Pebble  Pebble

            'Pebble BPL
            gTextcline(aFnum, Ax + (60 * aScale), Ay + (30.625 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalPebbleBpl, "##.0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Pebble TPA
            gTextcline(aFnum, Ax + (39.75 * aScale), Ay + (30.625 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalPebbleTpa, "####0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Pebble Fe2O3
            gTextcline(aFnum, Ax + (81 * aScale), Ay + (30.625 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalPebbleFe2O3, "#0.00"), _
                       aDrwLayer, aColor, "ROMANS")

            'Pebble Al2O3
            gTextcline(aFnum, Ax + (104 * aScale), Ay + (30.625 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalPebbleAl2O3, "#0.00"), _
                       aDrwLayer, aColor, "ROMANS")

            'Pebble MgO
            gTextcline(aFnum, Ax + (127 * aScale), Ay + (30.625 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalPebbleMgO, "#0.00"), _
                       aDrwLayer, aColor, "ROMANS")

            'Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed
            'Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed
            'Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed

            'Total feed BPL
            gTextcline(aFnum, Ax + (60 * aScale), Ay + (21.875 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalFeedBpl, "#0.0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Total feed TPA
            gTextcline(aFnum, Ax + (39.75 * aScale), Ay + (21.875 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalFeedTpa, "####0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Concentrate  Concentrate  Concentrate  Concentrate  Concentrate
            'Concentrate  Concentrate  Concentrate  Concentrate  Concentrate
            'Concentrate  Concentrate  Concentrate  Concentrate  Concentrate

            'Concentrate BPL
            gTextcline(aFnum, Ax + (60 * aScale), Ay + (13.125 * aScale), _
                       4.25 * aScale, 2, 2, Format(gComposite.ConcentrateBPL, "#0.0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Concentrate TPA
            gTextcline(aFnum, Ax + (39.75 * aScale), Ay + (13.125 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.ConcentrateTPA, "####0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Concentrate Fe2O3
            gTextcline(aFnum, Ax + (81 * aScale), Ay + (13.125 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.ConcentrateFe2O3, "#0.00"), _
                       aDrwLayer, aColor, "ROMANS")

            'Concentrate Al2O3
            gTextcline(aFnum, Ax + (104 * aScale), Ay + (13.125 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.ConcentrateAl2O3, "#0.00"), _
                       aDrwLayer, aColor, "ROMANS")

            'Concentrate MgO
            gTextcline(aFnum, Ax + (127 * aScale), Ay + (13.125 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.ConcentrateMgO, "#0.00"), _
                       aDrwLayer, aColor, "ROMANS")

            'Total Product  Total Product  Total Product  Total Product
            'Total Product  Total Product  Total Product  Total Product
            'Total Product  Total Product  Total Product  Total Product

            'Total product BPL
            gTextcline(aFnum, Ax + (60 * aScale), Ay + (4.375 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalProductBpl, "#0.0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Total Product TPA
            gTextcline(aFnum, Ax + (39.75 * aScale), Ay + (4.375 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalProductTpa, "####0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Total Product Fe2O3
            gTextcline(aFnum, Ax + (81 * aScale), Ay + (4.375 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalProductFe2O3, "#0.00"), _
                       aDrwLayer, aColor, "ROMANS")

            'Total Product Al2O3
            gTextcline(aFnum, Ax + (104 * aScale), Ay + (4.375 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalProductAl2O3, "#0.00"), _
                       aDrwLayer, aColor, "ROMANS")

            'Total Product MgO
            gTextcline(aFnum, Ax + (127 * aScale), Ay + (4.375 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalProductMgO, "#0.00"), _
                       aDrwLayer, aColor, "ROMANS")
        Else
            'InViso alignment

            Ycorr = 0.5 * TextHeight

            'Add hole name
            Xcorr = 0
            gTextclineL(aFnum, Ax + (7 * aScale) - Xcorr, Ay + (74.5 * aScale) - Ycorr, _
                        TextHeight, gComposite.HoleLocation, _
                        aDrwLayer, aColor)

            'Add Section, Township, Range
            Location = gGetHoleLocationShortDot(gComposite.Section, gComposite.Township, _
                       gComposite.Range)

            Xcorr = 0
            gTextclineL(aFnum, Ax + (30 * aScale) - Xcorr, Ay + (74.5 * aScale) - Ycorr, _
                        TextHeight, Location, _
                        aDrwLayer, aColor)

            'Add drill date
            Xcorr = 0
            gTextclineL(aFnum, Ax + (70 * aScale) - Xcorr, Ay + (74.5 * aScale) - Ycorr, _
                        TextHeight, gComposite.DrillCdate, _
                        aDrwLayer, aColor)

            'Add overburden thickness
            SpecColor = 5   'Blue
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (33 * aScale) - Xcorr, Ay + (65.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.OvbThickness, "##0.0"), 5), _
                        aDrwLayer, SpecColor)

            'Add matrix thickness
            SpecColor = 3  'Green
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (33 * aScale) - Xcorr, Ay + (56.875 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.MtxThickness, "##0.0"), 5), _
                        aDrwLayer, SpecColor)

            'Add MatrixX
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (93 * aScale) - Xcorr, Ay + (65.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.MtxX, "##0.0"), 5), _
                        aDrwLayer, aColor)

            'Add TotalX
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (93 * aScale) - Xcorr, Ay + (56.875 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalX, "##0.0"), 5), _
                        aDrwLayer, aColor)

            '%Clay
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (93 * aScale) - Xcorr, Ay + (48.125 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.WasteClayWtp, "##0.0"), 5), _
                        aDrwLayer, aColor)

            'Pebble  Pebble  Pebble  Pebble  Pebble  Pebble  Pebble
            'Pebble  Pebble  Pebble  Pebble  Pebble  Pebble  Pebble
            'Pebble  Pebble  Pebble  Pebble  Pebble  Pebble  Pebble

            'Pebble BPL
            Xcorr = 0.15
            gTextclineL(aFnum, Ax + (60 * aScale) - Xcorr, Ay + (30.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalPebbleBpl, "#0.0"), 4), _
                        aDrwLayer, aColor)

            'Pebble TPA
            Xcorr = 0.22
            gTextclineL(aFnum, Ax + (39.75 * aScale) - Xcorr, Ay + (30.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalPebbleTpa, "####0"), 5), _
                        aDrwLayer, aColor)

            'Pebble Fe2O3
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (81 * aScale) - Xcorr, Ay + (30.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalPebbleFe2O3, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Pebble Al2O3
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (104 * aScale) - Xcorr, Ay + (30.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalPebbleAl2O3, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Pebble MgO
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (127 * aScale) - Xcorr, Ay + (30.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalPebbleMgO, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed
            'Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed
            'Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed

            'Total feed BPL
            Xcorr = 0.15
            gTextclineL(aFnum, Ax + (60 * aScale) - Xcorr, Ay + (21.875 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalFeedBpl, "#0.0"), 4), _
                        aDrwLayer, aColor)

            'Total feed TPA
            Xcorr = 0.22
            gTextclineL(aFnum, Ax + (39.75 * aScale) - Xcorr, Ay + (21.875 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalFeedTpa, "####0"), 5), _
                        aDrwLayer, aColor)

            'Concentrate  Concentrate  Concentrate  Concentrate  Concentrate
            'Concentrate  Concentrate  Concentrate  Concentrate  Concentrate
            'Concentrate  Concentrate  Concentrate  Concentrate  Concentrate

            'Concentrate BPL
            Xcorr = 0.15
            gTextclineL(aFnum, Ax + (60 * aScale) - Xcorr, Ay + (13.125 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.ConcentrateBPL, "#0.0"), 4), _
                        aDrwLayer, aColor)

            'Concentrate TPA
            Xcorr = 0.22
            gTextclineL(aFnum, Ax + (39.75 * aScale) - Xcorr, Ay + (13.125 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.ConcentrateTPA, "####0"), 5), _
                        aDrwLayer, aColor)

            'Concentrate Fe2O3
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (81 * aScale) - Xcorr, Ay + (13.125 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.ConcentrateFe2O3, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Concentrate Al2O3
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (104 * aScale) - Xcorr, Ay + (13.125 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.ConcentrateAl2O3, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Concentrate MgO
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (127 * aScale) - Xcorr, Ay + (13.125 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.ConcentrateMgO, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Total Product  Total Product  Total Product  Total Product
            'Total Product  Total Product  Total Product  Total Product
            'Total Product  Total Product  Total Product  Total Product

            'Total product BPL
            Xcorr = 0.15
            gTextclineL(aFnum, Ax + (60 * aScale) - Xcorr, Ay + (4.375 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalProductBpl, "#0.0"), 4), _
                        aDrwLayer, aColor)

            'Total Product TPA
            Xcorr = 0.22
            gTextclineL(aFnum, Ax + (39.75 * aScale) - Xcorr, Ay + (4.375 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalProductTpa, "####0"), 5), _
                        aDrwLayer, aColor)

            'Total Product Fe2O3
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (81 * aScale) - Xcorr, Ay + (4.375 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalProductFe2O3, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Total Product Al2O3
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (104 * aScale) - Xcorr, Ay + (4.375 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalProductAl2O3, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Total Product MgO
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (127 * aScale) - Xcorr, Ay + (4.375 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalProductMgO, "#0.00"), 5), _
                        aDrwLayer, aColor)
        End If
    End Sub

    Public Function gGetProspHoleCoords(ByVal aMineName As String, _
                                        ByVal aSec As Integer, _
                                        ByVal aTwp As Integer, _
                                        ByVal aRge As Integer, _
                                        ByVal aHole As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo gGetProspHoleCoordsError

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer

        Dim RowCount As Integer
        Dim ColCount As Integer

        Dim NWx As Double
        Dim NWy As Double
        Dim NEx As Double
        Dim NEy As Double
        Dim SEx As Double
        Dim SEy As Double
        Dim SWx As Double
        Dim SWy As Double

        Dim Bot(0 To 33, 0 To 2)
        Dim Top(0 To 33, 0 To 2)
        Dim Left(0 To 33, 0 To 2)
        Dim Right(0 To 33, 0 To 2)

        Dim Dist As Double
        Dim Angle As Single
        Dim IncDist As Double
        Dim AccumDist As Double
        Dim X As Integer
        Dim Y As Integer
        Dim Pos As Integer
        Dim Pos2 As Integer
        Dim Line As Integer
        Dim X1 As Double
        Dim Y1 As Double
        Dim X2 As Double
        Dim Y2 As Double

        Dim SecStr As String
        Dim TwpStr As String
        Dim RgeStr As String

        Dim LettString As String

        Dim Xcoord As Double
        Dim Ycoord As Double

        'This function works for alpha-numeric hole locations only!
        'A01, J12, K06, etc.  It also works only for prospect grids
        'that are developed by dividing each of the section sides into
        '16 equivalent line segments.

        If gHoleLocationOk(aHole) = False Then
            gGetProspHoleCoords = ""
            Exit Function
        End If

        If aSec > 9 Then
            SecStr = Trim(Str(aSec))
        Else
            SecStr = "0" & Trim(Str(aSec))
        End If
        If aTwp > 9 Then
            TwpStr = Trim(Str(aTwp))
        Else
            TwpStr = "0" & Trim(Str(aTwp))
        End If
        If aRge > 9 Then
            RgeStr = Trim(Str(aRge))
        Else
            RgeStr = "0" & Trim(Str(aRge))
        End If

        'Need to get the state-planar coordinates for this section.
        'They should be in the table SECTN_COORDS.

        'Get section state-planar coordinates
        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_sectn_coords
        'pMineName
        'pSection
        'pTownship
        'pRange
        'pResult

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_survey.get_sectn_coords(:pMineName," + _
                      ":pSection, :pTownship, :pRange, :pResult);end;", ORASQL_FAILEXEC)
        mSectnCoordsDynaset = params("pResult").Value

        RecordCount = mSectnCoordsDynaset.RecordCount
        ClearParams(params)

        If RecordCount <> 1 Then
            gGetProspHoleCoords = ""
            Exit Function
        End If

        'We now have the state-planar coordinates for this section
        mSectnCoordsDynaset.MoveFirst()
        NWx = mSectnCoordsDynaset.Fields("nw_x_cdnt").Value
        NWy = mSectnCoordsDynaset.Fields("nw_y_cdnt").Value
        NEx = mSectnCoordsDynaset.Fields("ne_x_cdnt").Value
        NEy = mSectnCoordsDynaset.Fields("ne_y_cdnt").Value
        SWx = mSectnCoordsDynaset.Fields("sw_x_cdnt").Value
        SWy = mSectnCoordsDynaset.Fields("sw_y_cdnt").Value
        SEx = mSectnCoordsDynaset.Fields("se_x_cdnt").Value
        SEy = mSectnCoordsDynaset.Fields("se_y_cdnt").Value

        'Create the prospect grid
        'Bottom line
        Dist = gGetDistance(SWx, SWy, SEx, SEy)
        Angle = gGetAngle(SWx, SWy, SEx, SEy, "X")
        IncDist = Round(Dist / 32, 4)
        Bot(1, 1) = SWx
        Bot(1, 2) = SWy
        Bot(33, 1) = SEx
        Bot(33, 2) = SEy
        For Pos = 1 To 31
            AccumDist = Round(Pos * IncDist, 4)
            gGetNewCoords(SWx, SWy, SEx, SEy, Angle, AccumDist, "X")
            Bot(Pos + 1, 1) = mXcoord
            Bot(Pos + 1, 2) = mYcoord
        Next Pos

        'Top line
        Dist = gGetDistance(NWx, NWy, NEx, NEy)
        Angle = gGetAngle(NWx, NWy, NEx, NEy, "X")
        IncDist = Round(Dist / 32, 4)
        Top(1, 1) = NWx
        Top(1, 2) = NWy
        Top(33, 1) = NEx
        Top(33, 2) = NEy
        For Pos = 1 To 31
            AccumDist = Round(Pos * IncDist, 4)
            gGetNewCoords(NWx, NWy, NEx, NEy, Angle, AccumDist, "X")
            Top(Pos + 1, 1) = mXcoord
            Top(Pos + 1, 2) = mYcoord
        Next Pos

        'Left line
        Dist = gGetDistance(SWx, SWy, NWx, NWy)
        Angle = gGetAngle(SWx, SWy, NWx, NWy, "Y")
        IncDist = Round(Dist / 32, 4)
        Left(1, 1) = SWx
        Left(1, 2) = SWy
        Left(33, 1) = NWx
        Left(33, 2) = NWy
        For Pos = 1 To 31
            AccumDist = Round(Pos * IncDist, 4)
            gGetNewCoords(SWx, SWy, NWx, NWy, Angle, AccumDist, "Y")
            Left(Pos + 1, 1) = mXcoord
            Left(Pos + 1, 2) = mYcoord
        Next Pos

        'Right line
        Dist = gGetDistance(SEx, SEy, NEx, NEy)
        Angle = gGetAngle(SEx, SEy, NEx, NEy, "Y")
        IncDist = Round(Dist / 32, 4)
        Right(1, 1) = SEx
        Right(1, 2) = SEy
        Right(33, 1) = NEx
        Right(33, 2) = NEy
        For Pos = 1 To 31
            AccumDist = Round(Pos * IncDist, 4)
            gGetNewCoords(SEx, SEy, NEx, NEy, Angle, AccumDist, "Y")
            Right(Pos + 1, 1) = mXcoord
            Right(Pos + 1, 2) = mYcoord
        Next Pos

        'Fill in gProspX(1 To 16, 1 To 16)
        'Fill in gProspY(1 To 16, 1 To 16)
        For Pos = 2 To 32 Step 2    'Process 16 east-west lines
            X1 = Left(Pos, 1)
            Y1 = Left(Pos, 2)
            X2 = Right(Pos, 1)
            Y2 = Right(Pos, 2)
            Dist = gGetDistance(X1, Y1, X2, Y2)
            Angle = gGetAngle(X1, Y1, X2, Y2, "X")
            IncDist = Round(Dist / 32, 4)
            For Pos2 = 2 To 32 Step 2
                AccumDist = Round((Pos2 - 1) * IncDist, 4)
                gGetNewCoords(X1, Y1, X2, Y2, Angle, AccumDist, "X")
                gProspX(Pos2 / 2, Pos / 2) = mXcoord
                gProspY(Pos2 / 2, Pos / 2) = mYcoord
            Next Pos2
        Next Pos

        'Coordinates for all holes in section are in:
        'gProspX(1 To 16, 1 To 16)
        'gProspY(1 To 16, 1 To 16)

        'Need to assign form variables: fXcoord and fYcoord

        Y = Val(Mid(aHole, 2))   '1 to 16

        LettString = "ABCDEFGHIJKLMNOP"

        For X = 1 To 16
            If Mid(LettString, X, 1) = Mid(aHole, 1, 1) Then
                Exit For
            End If
        Next X

        If X <> 0 And Y <> 0 Then
            Xcoord = gProspX(X, Y)
            Ycoord = gProspY(X, Y)
        Else
            Xcoord = 0
            Ycoord = 0
        End If

        gGetProspHoleCoords = Str(Xcoord) + "/" + Str(Ycoord)
        Exit Function

gGetProspHoleCoordsError:
        MsgBox("Error getting prospect hole coordinates." & vbCrLf & _
            Err.Description, _
            vbOKOnly + vbExclamation, _
            "Prospect Hole Coordinates Get Error")
    End Function

    Public Sub gMakeSpDonut(ByVal aFileNumber As Integer, _
                            ByVal aXcoord As Double, _
                            ByVal aYcoord As Double, _
                            ByVal aLayer As String, _
                            ByVal aInsideDiam As Single, _
                            ByVal aOutsideDiam As String, _
                            ByVal aColor As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'NOTE -- DOES NOT WORK!!!!!!!
        Dim CalcVal As Single
        Dim CalcVal2 As Single

        CalcVal = (aOutsideDiam - aInsideDiam) / 2
        CalcVal2 = aXcoord + CalcVal

        'Make a donut
        gWriteLine(aFileNumber, "0")
        gWriteLine(aFileNumber, "LWPOLYLINE")
        gWriteLine(aFileNumber, "5")
        gEntityNumber = gEntityNumber + 1
        gWriteLine(aFileNumber, Hex(gEntityNumber))
        gWriteLine(aFileNumber, "100")
        gWriteLine(aFileNumber, "AcDbEntity")
        gWriteLine(aFileNumber, "8")
        gWriteLine(aFileNumber, Trim(StrConv(aLayer, vbUpperCase)))
        gWriteLine(aFileNumber, "100")
        gWriteLine(aFileNumber, "AcDbPolyline")
        gWriteLine(aFileNumber, "90")
        gWriteLine(aFileNumber, "2")
        gWriteLine(aFileNumber, "70")
        gWriteLine(aFileNumber, "1")
        gWriteLine(aFileNumber, "43")
        gWriteLine(aFileNumber, Trim(CStr(CalcVal)))
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(CStr(aXcoord)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(CStr(aYcoord)))
        gWriteLine(aFileNumber, "42")
        gWriteLine(aFileNumber, "1.0")
        gWriteLine(aFileNumber, "10")
        gWriteLine(aFileNumber, Trim(CStr(CalcVal2)))
        gWriteLine(aFileNumber, "20")
        gWriteLine(aFileNumber, Trim(CStr(aYcoord)))
        gWriteLine(aFileNumber, "42")
        gWriteLine(aFileNumber, "1.0")

        'Colors     1   Red
        '           2   Yellow
        '           3   Green
        '           4   Cyan
        '           5   Blue
        '           6   Magenta
        '           7   White
        '           40  Orange

        If aColor <> 0 Then
            gWriteLine(aFileNumber, "62")
            gWriteLine(aFileNumber, Trim(Str(aColor)))
        End If
    End Sub

    Public Function gGetCurrent6MonthMap(ByVal aMineName As String) As String

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim ColIdx As Integer
        Dim RecordCount As Long

        On Error GoTo gGetCurrent6MonthMapError

        'Get current 6 month mine plan
        params = gDBParams

        'PROCEDURE get_all_maps
        'pMineName
        'pMapTypeName
        'pEqptTypeName
        'pEqptName
        'pResult

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pMapTypeName", "6 month mine plan", ORAPARM_INPUT)
        params("pMapTypeName").serverType = ORATYPE_VARCHAR2

        params.Add("pEqptTypeName", DBNull.Value, ORAPARM_INPUT)
        params("pEqptTypeName").serverType = ORATYPE_VARCHAR2

        params.Add("pEqptName", DBNull.Value, ORAPARM_INPUT)
        params("pEqptName").serverType = ORATYPE_VARCHAR2

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_maps.get_all_maps(:pMineName, " + _
                      ":pMapTypeName, :pEqptTypeName, :pEqptName, :pResult);end;", ORASQL_FAILEXEC)
        mMapDynaset = params("pResult").Value
        ClearParams(params)
        RecordCount = mMapDynaset.RecordCount

        If RecordCount = 0 Then
            gGetCurrent6MonthMap = ""
        Else
            'Most current 6 month mine plan will be the first one in the cursor
            mMapDynaset.MoveFirst()
            gGetCurrent6MonthMap = gPath + "\Maps\" + mMapDynaset.Fields("map_name").Value
        End If

        Exit Function

gGetCurrent6MonthMapError:
        MsgBox("Error getting current 6 month mine plan." & vbCrLf & _
        Err.Description, _
        vbOKOnly + vbExclamation, _
        "Current 6 Month Mine Plan Access Error")

        On Error Resume Next
        ClearParams(params)
    End Function

    Public Function gGetQuadrilateralArea(ByVal aAx As Double, _
                                          ByVal aAy As Double, _
                                          ByVal aBx As Double, _
                                          ByVal aBy As Double, _
                                          ByVal aCx As Double, _
                                          ByVal aCy As Double, _
                                          ByVal aDx As Double, _
                                          ByVal aDy As Double, _
                                          ByVal aRoundVal As Integer) As Double

        '**********************************************************************
        '
        '
        '
        '**********************************************************************


        Dim DistAB As Double
        Dim DistBC As Double
        Dim DistCD As Double
        Dim DistDA As Double
        Dim DistAC As Double

        Dim TriArea1 As Double
        Dim TriArea2 As Double
        Dim SideVal As Double

        DistAB = gGetDistance(aAx, aAy, aBx, aBy)
        DistBC = gGetDistance(aBx, aBy, aCx, aCy)
        DistCD = gGetDistance(aCx, aCy, aDx, aDy)
        DistDA = gGetDistance(aDx, aDy, aAx, aAy)
        DistAC = gGetDistance(aAx, aAy, aCx, aCy)

        'Area of first triangle
        SideVal = 0.5 * (DistAB + DistBC + DistAC)
        TriArea1 = Round(Sqrt(SideVal * (SideVal - DistAB) * _
                   (SideVal - DistBC) * (SideVal - DistAC)), aRoundVal)

        'Area of second triangle
        SideVal = 0.5 * (DistCD + DistDA + DistAC)
        TriArea2 = Round(Sqrt(SideVal * (SideVal - DistCD) * _
                   (SideVal - DistDA) * (SideVal - DistAC)), aRoundVal)

        gGetQuadrilateralArea = TriArea1 + TriArea2
    End Function

    Public Sub gMakeCompBoxSimp(ByVal aFnum As Integer, _
                            ByVal Ax As Double, _
                            ByVal Ay As Double, _
                            ByVal aXoff As Double, _
                            ByVal aYoff As Double, _
                            ByVal aDrwLayer As String, _
                            ByVal aScale As Single, _
                            ByVal aMode As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim Color As Integer
        Dim Xcorr As Double
        Dim Ycorr As Double
        Dim TextHeight As Single

        'Draw main box
        gPolyline(aFnum, Ax, Ay, Ax, _
                  Ay + (40 * aScale), 0, aDrwLayer)

        gPolyline(aFnum, Ax, Ay + (40 * aScale), _
                  Ax + (85 * aScale), _
                  Ay + (40 * aScale), 0, aDrwLayer)

        gPolyline(aFnum, Ax + (85 * aScale), Ay + (40 * aScale), _
                  Ax + (85 * aScale), Ay, 0, aDrwLayer)

        gPolyline(aFnum, Ax + (85 * aScale), Ay, Ax, _
                  Ay, 0, aDrwLayer)

        'Draw individual lines
        'Draw vertical lines -- don't need any vertical lines!

        'Draw horizontal lines
        gPolyline(aFnum, Ax, Ay + (8.75 * aScale), Ax + (85 * aScale), _
                  Ay + (8.75 * aScale), 0, aDrwLayer)

        gPolyline(aFnum, Ax, Ay + (31 * aScale), Ax + (85 * aScale), _
                  Ay + (31 * aScale), 0.7 * aScale, aDrwLayer)

        'Draw text lines -- row headers
        'Draw text lines -- row headers
        'Draw text lines -- row headers

        TextHeight = 3.5 * aScale

        If aMode = "AutoCAD" Then
            'AutoCAD alignment

            gTextcline(aFnum, Ax + (7 * aScale), Ay + (4.375 * aScale), _
                       TextHeight, 1, 2, "TPrd", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (7 * aScale), Ay + (27 * aScale), _
                       TextHeight, 1, 2, "Ovb", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (7 * aScale), Ay + (19.5 * aScale), _
                       TextHeight, 1, 2, "Mtx", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (40 * aScale), Ay + (27 * aScale), _
                       TextHeight, 1, 2, "MtxX", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (40 * aScale), Ay + (19.5 * aScale), _
                       TextHeight, 1, 2, "TotX", aDrwLayer, 0, "ROMANS")

            gTextcline(aFnum, Ax + (40 * aScale), Ay + (12 * aScale), _
                       TextHeight, 1, 2, "%Cly", aDrwLayer, 0, "ROMANS")
        Else
            'InViso alignment
            Ycorr = 0.5 * TextHeight

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (9 * aScale) - Xcorr, Ay + (4.375 * aScale) - Ycorr, _
                        TextHeight, "TPrd", aDrwLayer, 0)

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (10 * aScale) - Xcorr, Ay + (56.875 * aScale) - Ycorr, _
                        TextHeight, "Mtx", aDrwLayer, 0)

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (10 * aScale) - Xcorr, Ay + (65.625 * aScale) - Ycorr, _
                        TextHeight, "Ovb", aDrwLayer, 0)

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (68 * aScale) - Xcorr, Ay + (65.625 * aScale) - Ycorr, _
                        TextHeight, "MtxX", aDrwLayer, 0)

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (68 * aScale) - Xcorr, Ay + (56.875 * aScale) - Ycorr, _
                        TextHeight, "TotX", aDrwLayer, 0)

            Xcorr = 0.08
            gTextclineL(aFnum, Ax + (68 * aScale) - Xcorr, Ay + (48.125 * aScale) - Ycorr, _
                        TextHeight, "%Cly", aDrwLayer, 0)
        End If

        Color = 0   'Black
        AddCompDataSimp(aFnum, Ax, Ay, aDrwLayer, aScale, Color, aMode)
    End Sub

    Private Sub AddCompDataSimp(ByVal aFnum As Integer, _
                            ByVal Ax As Double, _
                            ByVal Ay As Double, _
                            ByVal aDrwLayer As String, _
                            ByVal aScale As Single, _
                            ByVal aColor As String, _
                            ByVal aMode As String)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim Location As String
        Dim SpecColor As Integer
        Dim TextHeight As Single
        Dim Xcorr As Double
        Dim Ycorr As Double
        Dim MtxxVal As Single

        TextHeight = 4.25 * aScale

        If aMode = "AutoCAD" Then
            'AutoCAD alignment

            'Add hole name
            gTextcline(aFnum, Ax + (1.5 * aScale), Ay + (36 * aScale), _
                       TextHeight, 0, 2, gComposite.HoleLocation, _
                       aDrwLayer, aColor, "ROMANS")

            'Add Section, Township, Range
            Location = gGetHoleLocationShortDot(gComposite.Section, gComposite.Township, _
                       gComposite.Range)

            gTextcline(aFnum, Ax + (16 * aScale), Ay + (36 * aScale), _
                       TextHeight, 0, 2, Location, _
                       aDrwLayer, aColor, "ROMANS")

            'Add drill date
            gTextcline(aFnum, Ax + (50 * aScale), Ay + (36 * aScale), _
                       3.5 * aScale, 0, 2, gComposite.DrillCdate, _
                       aDrwLayer, aColor, "ROMANS")

            'Add overburden thickness
            SpecColor = 5   'Blue
            gTextcline(aFnum, Ax + (30 * aScale), Ay + (27 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.OvbThickness, "##0.0"), _
                       aDrwLayer, SpecColor, "ROMANS")

            'Add matrix thickness
            SpecColor = 3  'Green
            gTextcline(aFnum, Ax + (30 * aScale), Ay + (19.5 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.MtxThickness, "##0.0"), _
                       aDrwLayer, SpecColor, "ROMANS")

            'Add MatrixX
            If gComposite.MtxThickness = 0 Then
                MtxxVal = 0
            Else
                MtxxVal = gComposite.MtxX
            End If

            gTextcline(aFnum, Ax + (61 * aScale), Ay + (27 * aScale), _
                       TextHeight, 2, 2, Format(MtxxVal, "##0.0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Add TotalX
            gTextcline(aFnum, Ax + (61 * aScale), Ay + (19.5 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalX, "##0.0"), _
                       aDrwLayer, aColor, "ROMANS")

            '%Clay
            gTextcline(aFnum, Ax + (61 * aScale), Ay + (12 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.WasteClayWtp, "##0.0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Total product BPL
            gTextcline(aFnum, Ax + (54.25 * aScale), Ay + (4.375 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalProductBpl, "#0.0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Total Product TPA
            gTextcline(aFnum, Ax + (34 * aScale), Ay + (4.375 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalProductTpa, "####0"), _
                       aDrwLayer, aColor, "ROMANS")

            'Total Product MgO
            gTextcline(aFnum, Ax + (74.5 * aScale), Ay + (4.375 * aScale), _
                       TextHeight, 2, 2, Format(gComposite.TotalProductMgO, "#0.00"), _
                       aDrwLayer, aColor, "ROMANS")
        Else
            'InViso alignment

            Ycorr = 0.5 * TextHeight

            'Add hole name
            Xcorr = 0
            gTextclineL(aFnum, Ax + (7 * aScale) - Xcorr, Ay + (74.5 * aScale) - Ycorr, _
                        TextHeight, gComposite.HoleLocation, _
                        aDrwLayer, aColor)

            'Add Section, Township, Range
            Location = gGetHoleLocationShortDot(gComposite.Section, gComposite.Township, _
                       gComposite.Range)

            Xcorr = 0
            gTextclineL(aFnum, Ax + (30 * aScale) - Xcorr, Ay + (74.5 * aScale) - Ycorr, _
                        TextHeight, Location, _
                        aDrwLayer, aColor)

            'Add drill date
            Xcorr = 0
            gTextclineL(aFnum, Ax + (70 * aScale) - Xcorr, Ay + (74.5 * aScale) - Ycorr, _
                        TextHeight, gComposite.DrillCdate, _
                        aDrwLayer, aColor)

            'Add overburden thickness
            SpecColor = 5   'Blue
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (33 * aScale) - Xcorr, Ay + (65.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.OvbThickness, "##0.0"), 5), _
                        aDrwLayer, SpecColor)

            'Add matrix thickness
            SpecColor = 3  'Green
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (33 * aScale) - Xcorr, Ay + (56.875 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.MtxThickness, "##0.0"), 5), _
                        aDrwLayer, SpecColor)

            'Add MatrixX
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (93 * aScale) - Xcorr, Ay + (65.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.MtxX, "##0.0"), 5), _
                        aDrwLayer, aColor)

            'Add TotalX
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (93 * aScale) - Xcorr, Ay + (56.875 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalX, "##0.0"), 5), _
                        aDrwLayer, aColor)

            '%Clay
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (93 * aScale) - Xcorr, Ay + (48.125 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.WasteClayWtp, "##0.0"), 5), _
                        aDrwLayer, aColor)

            'Pebble  Pebble  Pebble  Pebble  Pebble  Pebble  Pebble
            'Pebble  Pebble  Pebble  Pebble  Pebble  Pebble  Pebble
            'Pebble  Pebble  Pebble  Pebble  Pebble  Pebble  Pebble

            'Pebble BPL
            Xcorr = 0.15
            gTextclineL(aFnum, Ax + (60 * aScale) - Xcorr, Ay + (30.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalPebbleBpl, "#0.0"), 4), _
                        aDrwLayer, aColor)

            'Pebble TPA
            Xcorr = 0.22
            gTextclineL(aFnum, Ax + (39.75 * aScale) - Xcorr, Ay + (30.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalPebbleTpa, "####0"), 5), _
                        aDrwLayer, aColor)

            'Pebble Fe2O3
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (81 * aScale) - Xcorr, Ay + (30.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalPebbleFe2O3, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Pebble Al2O3
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (104 * aScale) - Xcorr, Ay + (30.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalPebbleAl2O3, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Pebble MgO
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (127 * aScale) - Xcorr, Ay + (30.625 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalPebbleMgO, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed
            'Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed
            'Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed  Feed

            'Total feed BPL
            Xcorr = 0.15
            gTextclineL(aFnum, Ax + (60 * aScale) - Xcorr, Ay + (21.875 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalFeedBpl, "#0.0"), 4), _
                        aDrwLayer, aColor)

            'Total feed TPA
            Xcorr = 0.22
            gTextclineL(aFnum, Ax + (39.75 * aScale) - Xcorr, Ay + (21.875 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalFeedTpa, "####0"), 5), _
                        aDrwLayer, aColor)

            'Concentrate  Concentrate  Concentrate  Concentrate  Concentrate
            'Concentrate  Concentrate  Concentrate  Concentrate  Concentrate
            'Concentrate  Concentrate  Concentrate  Concentrate  Concentrate

            'Concentrate BPL
            Xcorr = 0.15
            gTextclineL(aFnum, Ax + (60 * aScale) - Xcorr, Ay + (13.125 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.ConcentrateBPL, "#0.0"), 4), _
                        aDrwLayer, aColor)

            'Concentrate TPA
            Xcorr = 0.22
            gTextclineL(aFnum, Ax + (39.75 * aScale) - Xcorr, Ay + (13.125 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.ConcentrateTPA, "####0"), 5), _
                        aDrwLayer, aColor)

            'Concentrate Fe2O3
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (81 * aScale) - Xcorr, Ay + (13.125 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.ConcentrateFe2O3, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Concentrate Al2O3
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (104 * aScale) - Xcorr, Ay + (13.125 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.ConcentrateAl2O3, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Concentrate MgO
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (127 * aScale) - Xcorr, Ay + (13.125 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.ConcentrateMgO, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Total Product  Total Product  Total Product  Total Product
            'Total Product  Total Product  Total Product  Total Product
            'Total Product  Total Product  Total Product  Total Product

            'Total product BPL
            Xcorr = 0.15
            gTextclineL(aFnum, Ax + (60 * aScale) - Xcorr, Ay + (4.375 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalProductBpl, "#0.0"), 4), _
                        aDrwLayer, aColor)

            'Total Product TPA
            Xcorr = 0.22
            gTextclineL(aFnum, Ax + (39.75 * aScale) - Xcorr, Ay + (4.375 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalProductTpa, "####0"), 5), _
                        aDrwLayer, aColor)

            'Total Product Fe2O3
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (81 * aScale) - Xcorr, Ay + (4.375 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalProductFe2O3, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Total Product Al2O3
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (104 * aScale) - Xcorr, Ay + (4.375 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalProductAl2O3, "#0.00"), 5), _
                        aDrwLayer, aColor)

            'Total Product MgO
            Xcorr = 0.2
            gTextclineL(aFnum, Ax + (127 * aScale) - Xcorr, Ay + (4.375 * aScale) - Ycorr, _
                        TextHeight, gPadLeft(Format(gComposite.TotalProductMgO, "#0.00"), 5), _
                        aDrwLayer, aColor)
        End If
    End Sub

    Public Function gCreateProspectGrid2(ByVal aSec As Integer, _
                                         ByVal aTwp As Integer, _
                                         ByVal aRge As Integer, _
                                         ByVal aFileNumber As Integer, _
                                         ByVal aGridInDxf As Boolean, _
                                         ByVal aMineName As String, _
                                         ByVal aProspGridLayer As String, _
                                         ByVal aIncludeDate As Boolean, _
                                         ByVal aColor As Integer, _
                                         ByVal aSecBdryWidth As String, _
                                         ByVal aAddHoleLocs As Boolean, _
                                         ByVal aIncludeAreas As Boolean) As Boolean

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo CreateProspectGrid2Error

        'IMC style prospect grid.

        Dim params As OraParameters
        Dim SQLStmt As OraSqlStmt
        Dim RecordCount As Integer

        Dim LineIdx As Integer
        Dim SecLineIdx As Single

        Dim NWx As Double
        Dim NWy As Double
        Dim NEx As Double
        Dim NEy As Double
        Dim SEx As Double
        Dim SEy As Double
        Dim SWx As Double
        Dim SWy As Double

        Dim GL1 As gPtCoordType
        Dim GL2 As gPtCoordType

        Dim SL1 As gPtCoordType
        Dim SL2 As gPtCoordType

        Dim IntPtTest As gPtCoordType
        Dim IntPt1 As gPtCoordType
        Dim IntPt2 As gPtCoordType
        Dim HaveTwoIntPts As Boolean

        Dim MapLayer As String
        Dim PolylineThick As Single
        Dim PolylineColor As Integer

        Dim SecStr As String
        Dim TwpStr As String
        Dim RgeStr As String

        Dim PosX As Integer
        Dim PosY As Integer
        Dim VerAlign As Single
        Dim HorAlign As Single
        Dim TextHeight As Single
        Dim TextColor As Integer
        Dim HoleLoc As String
        Dim X1 As Double
        Dim Y1 As Double

        If aSec > 9 Then
            SecStr = Trim(Str(aSec))
        Else
            SecStr = "0" & Trim(Str(aSec))
        End If
        If aTwp > 9 Then
            TwpStr = Trim(Str(aTwp))
        Else
            TwpStr = "0" & Trim(Str(aTwp))
        End If
        If aRge > 9 Then
            RgeStr = Trim(Str(aRge))
        Else
            RgeStr = "0" & Trim(Str(aRge))
        End If

        'Need to get the state-planar coordinates for this section.
        'They should be in the table SECTN_COORDS.

        'Get section state-planar coordinates
        params = gDBParams

        params.Add("pMineName", aMineName, ORAPARM_INPUT)
        params("pMineName").serverType = ORATYPE_VARCHAR2

        params.Add("pSection", aSec, ORAPARM_INPUT)
        params("pSection").serverType = ORATYPE_NUMBER

        params.Add("pTownship", aTwp, ORAPARM_INPUT)
        params("pTownship").serverType = ORATYPE_NUMBER

        params.Add("pRange", aRge, ORAPARM_INPUT)
        params("pRange").serverType = ORATYPE_NUMBER

        params.Add("pResult", 0, ORAPARM_OUTPUT)
        params("pResult").serverType = ORATYPE_CURSOR

        'PROCEDURE get_sectn_coords
        'pMineName      IN     VARCHAR2,
        'pSection       IN     NUMBER,
        'pTownship      IN     NUMBER,
        'pRange         IN     NUMBER,
        'pResult        IN OUT c_sectncoords)
        SQLStmt = gOradatabase.CreateSql("Begin mois.mois_survey.get_sectn_coords(:pMineName," + _
                      ":pSection, :pTownship, :pRange, :pResult);end;", ORASQL_FAILEXEC)
        mSectnCoordsDynaset = params("pResult").Value

        RecordCount = mSectnCoordsDynaset.RecordCount
        ClearParams(params)

        If RecordCount <> 1 Then
            gCreateProspectGrid2 = False
            Exit Function
        End If

        'We now have the state-planar coordinates for this section
        mSectnCoordsDynaset.MoveFirst()
        NWx = mSectnCoordsDynaset.Fields("nw_x_cdnt").Value
        NWy = mSectnCoordsDynaset.Fields("nw_y_cdnt").Value
        NEx = mSectnCoordsDynaset.Fields("ne_x_cdnt").Value
        NEy = mSectnCoordsDynaset.Fields("ne_y_cdnt").Value
        SWx = mSectnCoordsDynaset.Fields("sw_x_cdnt").Value
        SWy = mSectnCoordsDynaset.Fields("sw_y_cdnt").Value
        SEx = mSectnCoordsDynaset.Fields("se_x_cdnt").Value
        SEy = mSectnCoordsDynaset.Fields("se_y_cdnt").Value

        'Set up some parameters for the drawing
        'Which layer to use for prospect grid?
        If Len(Trim(aProspGridLayer)) = 0 Then
            'Use default layer
            MapLayer = SecStr + TwpStr + RgeStr + "g"
        Else
            MapLayer = aProspGridLayer
        End If

        PolylineThick = 0
        PolylineColor = aColor   'Red

        'Create the prospect grid

        'Process through the North South prospect grid lines
        'Process through the lines from East to West

        'Start at the SE corner of the prospect grid
        'Will extend the grid lines further in the y direction by 1000
        '(this will work because the prospect grid lines are always
        'North South lines).
        GL1.X = SEx
        GL1.Y = SEy - 7000
        GL2.X = SEx
        GL2.Y = SEy + 7000

        If aGridInDxf = True Then
            For LineIdx = 1 To 21
                'Determine if the prospect grid line intersects any of the Section lines
                IntPt1.X = 0
                IntPt1.Y = 0
                IntPt2.X = 0
                IntPt2.Y = 0
                HaveTwoIntPts = False

                For SecLineIdx = 1 To 4
                    Select Case SecLineIdx
                        Case Is = 1
                            'NW to NE line
                            SL1.X = NWx
                            SL1.Y = NWy
                            SL2.X = NEx
                            SL2.Y = NEy

                        Case Is = 2
                            'SW to SE line
                            SL1.X = SWx
                            SL1.Y = SWy
                            SL2.X = SEx
                            SL2.Y = SEy

                        Case Is = 3
                            'NE to SE line
                            SL1.X = NEx
                            SL1.Y = NEy
                            SL2.X = SEx
                            SL2.Y = SEy

                        Case Is = 4
                            'NW to SW line
                            SL1.X = NWx
                            SL1.Y = NWy
                            SL2.X = SWx
                            SL2.Y = SWy
                    End Select

                    IntPtTest.X = 0
                    IntPtTest.Y = 0
                    IntPtTest = GetLineInt(GL1, GL2, SL1, SL2)

                    If IntPtTest.X <> -9999 And IntPtTest.Y <> -9999 Then
                        'We have an intersection point!
                        If IntPt1.X = 0 Then
                            IntPt1.X = IntPtTest.X
                            IntPt1.Y = IntPtTest.Y
                        Else
                            IntPt2.X = IntPtTest.X
                            IntPt2.Y = IntPtTest.Y
                            HaveTwoIntPts = True
                        End If
                    End If

                    If HaveTwoIntPts = True Then
                        'Add the line to the AutoCad dxf file
                        'Draw a polyline
                        gPolycline(aFileNumber, _
                                   IntPt1.X, _
                                   IntPt1.Y, _
                                   IntPt2.X, _
                                   IntPt2.Y, _
                                   PolylineThick, _
                                    MapLayer, _
                                  PolylineColor)
                        Exit For
                    End If
                Next SecLineIdx

                'Move the line to the west 330 feet.
                GL1.X = GL1.X - 330
                GL2.X = GL2.X - 330

                'Now process the next North South prospect grid line.
            Next LineIdx

            'Need to process to the east as well in case the Section shape is weird.
            GL1.X = SEx + 330
            GL1.Y = SEy - 7000
            GL2.X = SEx + 330
            GL2.Y = SEy + 7000

            For LineIdx = 1 To 10
                'Determine if the prospect grid line intersects any of the Section lines
                IntPt1.X = 0
                IntPt1.Y = 0
                IntPt2.X = 0
                IntPt2.Y = 0
                HaveTwoIntPts = False

                For SecLineIdx = 1 To 4
                    Select Case SecLineIdx
                        Case Is = 1
                            'NW to NE line
                            SL1.X = NWx
                            SL1.Y = NWy
                            SL2.X = NEx
                            SL2.Y = NEy

                        Case Is = 2
                            'SW to SE line
                            SL1.X = SWx
                            SL1.Y = SWy
                            SL2.X = SEx
                            SL2.Y = SEy

                        Case Is = 3
                            'NE to SE line
                            SL1.X = NEx
                            SL1.Y = NEy
                            SL2.X = SEx
                            SL2.Y = SEy

                        Case Is = 4
                            'NW to SW line
                            SL1.X = NWx
                            SL1.Y = NWy
                            SL2.X = SWx
                            SL2.Y = SWy
                    End Select

                    IntPtTest.X = 0
                    IntPtTest.Y = 0
                    IntPtTest = GetLineInt(GL1, GL2, SL1, SL2)

                    If IntPtTest.X <> -9999 And IntPtTest.Y <> -9999 Then
                        'We have an intersection point!
                        If IntPt1.X = 0 Then
                            IntPt1.X = IntPtTest.X
                            IntPt1.Y = IntPtTest.Y
                        Else
                            IntPt2.X = IntPtTest.X
                            IntPt2.Y = IntPtTest.Y
                            HaveTwoIntPts = True
                        End If
                    End If

                    If HaveTwoIntPts = True Then
                        'Add the line to the AutoCad dxf file
                        'Draw a polyline
                        gPolycline(aFileNumber, _
                                   IntPt1.X, _
                                   IntPt1.Y, _
                                   IntPt2.X, _
                                   IntPt2.Y, _
                                   PolylineThick, _
                                   MapLayer, _
                                   PolylineColor)
                        Exit For
                    End If
                Next SecLineIdx

                'Move the line to the east 330 feet.
                GL1.X = GL1.X + 330
                GL2.X = GL2.X + 330

                'Now process the next North South prospect grid line.
            Next LineIdx

            'Process through the East West prospect grid lines
            'Process through the lines from South to North

            'Start at the SE corner of the prospect grid
            'Will extend the grid lines further in the X direction by 1000
            '(this will work because the prospect grid lines are always
            'East West lines).
            GL1.X = SEx - 7000
            GL1.Y = SEy
            GL2.X = SEx + 7000
            GL2.Y = SEy

            For LineIdx = 1 To 21
                'Determine if the prospect grid line intersects any of the Section lines
                IntPt1.X = 0
                IntPt1.Y = 0
                IntPt2.X = 0
                IntPt2.Y = 0
                HaveTwoIntPts = False

                For SecLineIdx = 1 To 4
                    Select Case SecLineIdx
                        Case Is = 1
                            'NW to NE line
                            SL1.X = NWx
                            SL1.Y = NWy
                            SL2.X = NEx
                            SL2.Y = NEy

                        Case Is = 2
                            'SW to SE line
                            SL1.X = SWx
                            SL1.Y = SWy
                            SL2.X = SEx
                            SL2.Y = SEy

                        Case Is = 3
                            'NE to SE line
                            SL1.X = NEx
                            SL1.Y = NEy
                            SL2.X = SEx
                            SL2.Y = SEy

                        Case Is = 4
                            'NW to SW line
                            SL1.X = NWx
                            SL1.Y = NWy
                            SL2.X = SWx
                            SL2.Y = SWy
                    End Select

                    IntPtTest.X = 0
                    IntPtTest.Y = 0
                    IntPtTest = GetLineInt(GL1, GL2, SL1, SL2)

                    If IntPtTest.X <> -9999 And IntPtTest.Y <> -9999 Then
                        'We have an intersection point!
                        If IntPt1.X = 0 Then
                            IntPt1.X = IntPtTest.X
                            IntPt1.Y = IntPtTest.Y
                        Else
                            'If the same as the first point then we don't have 2 points yet!
                            If IntPtTest.X = IntPt1.X And IntPtTest.Y = IntPt1.Y Then
                                IntPt2.X = 0
                                IntPt2.Y = 0
                            Else
                                IntPt2.X = IntPtTest.X
                                IntPt2.Y = IntPtTest.Y
                                HaveTwoIntPts = True
                            End If
                        End If
                    End If

                    If HaveTwoIntPts = True Then
                        'Add the line to the AutoCad dxf file
                        'Draw a polyline
                        gPolycline(aFileNumber, _
                                   IntPt1.X, _
                                   IntPt1.Y, _
                                   IntPt2.X, _
                                   IntPt2.Y, _
                                   PolylineThick, _
                                   MapLayer, _
                                   PolylineColor)
                        Exit For
                    End If
                Next SecLineIdx

                'Move the line to the north 330 feet.
                GL1.Y = GL1.Y + 330
                GL2.Y = GL2.Y + 330

                'Now process the next east west prospect grid line as we
                'move north.
            Next LineIdx

            'Need to process to the south as well in case the Section shape is weird.
            GL1.X = SEx - 7000
            GL1.Y = SEy - 330
            GL2.X = SEx + 7000
            GL2.Y = SEy - 330

            For LineIdx = 1 To 11
                'Determine if the prospect grid line intersects any of the Section lines
                IntPt1.X = 0
                IntPt1.Y = 0
                IntPt2.X = 0
                IntPt2.Y = 0
                HaveTwoIntPts = False

                For SecLineIdx = 1 To 4
                    Select Case SecLineIdx
                        Case Is = 1
                            'NW to NE line
                            SL1.X = NWx
                            SL1.Y = NWy
                            SL2.X = NEx
                            SL2.Y = NEy

                        Case Is = 2
                            'SW to SE line
                            SL1.X = SWx
                            SL1.Y = SWy
                            SL2.X = SEx
                            SL2.Y = SEy

                        Case Is = 3
                            'NE to SE line
                            SL1.X = NEx
                            SL1.Y = NEy
                            SL2.X = SEx
                            SL2.Y = SEy

                        Case Is = 4
                            'NW to SW line
                            SL1.X = NWx
                            SL1.Y = NWy
                            SL2.X = SWx
                            SL2.Y = SWy
                    End Select

                    IntPtTest.X = 0
                    IntPtTest.Y = 0
                    IntPtTest = GetLineInt(GL1, GL2, SL1, SL2)

                    If IntPtTest.X <> -9999 And IntPtTest.Y <> -9999 Then
                        'We have an intersection point!
                        If IntPt1.X = 0 Then
                            IntPt1.X = IntPtTest.X
                            IntPt1.Y = IntPtTest.Y
                        Else
                            'If the same as the first point then we don't have 2 points yet!
                            If IntPtTest.X = IntPt1.X And IntPtTest.Y = IntPt1.Y Then
                                IntPt2.X = 0
                                IntPt2.Y = 0
                            Else
                                IntPt2.X = IntPtTest.X
                                IntPt2.Y = IntPtTest.Y
                                HaveTwoIntPts = True
                            End If
                        End If
                    End If

                    If HaveTwoIntPts = True Then
                        'Add the line to the AutoCad dxf file
                        'Draw a polyline
                        gPolycline(aFileNumber, _
                                   IntPt1.X, _
                                   IntPt1.Y, _
                                   IntPt2.X, _
                                   IntPt2.Y, _
                                   PolylineThick, _
                                   MapLayer, _
                                   PolylineColor)
                        Exit For
                    End If
                Next SecLineIdx

                'Move the line to the south 330 feet.
                GL1.Y = GL1.Y - 330
                GL2.Y = GL2.Y - 330

                'Now process the next east west prospect grid line as we
                'move north.
            Next LineIdx

            'Add the actually section boundaries also.
            If aSecBdryWidth = "Thick" Then
                PolylineThick = 30
            Else
                PolylineThick = 0
            End If

            'East section line
            gPolycline(aFileNumber, NWx, NWy, SWx, SWy, PolylineThick, MapLayer, PolylineColor)
            'West section line
            gPolycline(aFileNumber, NEx, NEy, SEx, SEy, PolylineThick, MapLayer, PolylineColor)
            'North section line
            gPolycline(aFileNumber, NWx, NWy, NEx, NEy, PolylineThick, MapLayer, PolylineColor)
            'South section line
            gPolycline(aFileNumber, SWx, SWy, SEx, SEy, PolylineThick, MapLayer, PolylineColor)
        End If

        For PosX = 1 To 16
            For PosY = 1 To 16
                GL1 = GetNumericGridMidLoc(SEx, SEy, PosX, PosY)
                TextHeight = 40
                TextColor = 5   'Blue
                HorAlign = 0    'Left
                VerAlign = 2    'Middle
                X1 = GL1.X - 120
                Y1 = GL1.Y - 100

                'Fill in gProspX(1 To 16, 1 To 16)
                'Fill in gProspY(1 To 16, 1 To 16)
                'Will only "capture" the normal 16 x 16 grid.  Holes that are outside this
                'normal grid for weirdly shaped sections will not be captured here at this
                'time (maybe later I will expand this).
                gProspX(1 + (15 - (PosX - 1)), PosY) = GL1.X
                gProspY(1 + (15 - (PosX - 1)), PosY) = GL1.Y

                If aAddHoleLocs = True Then
                    HoleLoc = Format(32 - ((PosX - 1) * 2), "0#") & _
                              Format(34 + ((PosY - 1) * 2), "0#")

                    gTextcline(aFileNumber, X1, Y1, TextHeight, HorAlign, _
                               VerAlign, HoleLoc, MapLayer, TextColor, "ROMANS")
                End If
            Next PosY
        Next PosX

        gCreateProspectGrid2 = True
        Exit Function

CreateProspectGrid2Error:
        MsgBox("Error creating IMC type prospect grid." & vbCrLf & _
               Err.Description, _
               vbOKOnly + vbExclamation, _
               "IMC Type Prospect Grid Add Error")
    End Function

    Private Function GetLineInt(ByRef aPt1 As gPtCoordType, _
                                ByRef aPt2 As gPtCoordType, _
                                ByRef aPt3 As gPtCoordType, _
                                ByRef aPt4 As gPtCoordType) As gPtCoordType

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim Ua As Double
        Dim Ub As Double
        Dim PtInt As gPtCoordType

        GetLineInt.X = -9999
        GetLineInt.Y = -9999

        'Line a is P1 to P2
        'Line b is P3 to P4
        'Line equations are:
        '    Pa = P1 + Ua(P2 - P1)
        '    Pb = P3 + Ub(P4 - P3)
        '
        'Solving for the point where Pa = Pb gives the following two equations in two
        'unknowns (Ua and Ub)
        '     X1 + Ua(X2 - X1) = X3 + Ub(X4 - X3)
        '     Y1 + Ua(Y2 - Y1) = Y3 + Ub(Y4 - Y3)
        '
        'Solving gives the following expressions for Ua and Ub
        '     Ua = (x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)
        '          ---------------------------------------------
        '          (y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1)
        '
        '     Ub = (x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)
        '          ---------------------------------------------
        '          (y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1)
        '
        'Substituting either of these into the corresonding equation for the line gives the
        'intersection point.  The intersection point(x,y) is:
        '     x = x1 + Ua(x2 - x1)
        '     x = x1 + Ua(x2 - x1)
        '
        'The denominators for the equations for Ua and Ub are the same.
        '
        'If the denominator for the equations for Ua and Ub is zero then the two lines are
        'parallel.
        '
        'If the denominator and numerator for the equations for Ua and Ub are 0 the the two
        'lines are coincident.

        PtInt.X = 0
        PtInt.Y = 0

        If ((aPt4.Y - aPt3.Y) * (aPt2.X - aPt1.X) - (aPt4.X - aPt3.X) * (aPt2.Y - aPt1.Y)) <> 0 Then
            Ua = ((aPt4.X - aPt3.X) * (aPt1.Y - aPt3.Y) - (aPt4.Y - aPt3.Y) * (aPt1.X - aPt3.X)) / _
                 ((aPt4.Y - aPt3.Y) * (aPt2.X - aPt1.X) - (aPt4.X - aPt3.X) * (aPt2.Y - aPt1.Y))
        Else
            Ua = 0
        End If

        If ((aPt4.Y - aPt3.Y) * (aPt2.X - aPt1.X) - (aPt4.X - aPt3.X) * (aPt2.Y - aPt1.Y)) <> 0 Then
            Ub = ((aPt2.X - aPt1.X) * (aPt1.Y - aPt3.Y) - (aPt2.Y - aPt1.Y) * (aPt1.X - aPt3.X)) / _
                 ((aPt4.Y - aPt3.Y) * (aPt2.X - aPt1.X) - (aPt4.X - aPt3.X) * (aPt2.Y - aPt1.Y))
        Else
            Ub = 0
        End If

        PtInt.X = aPt1.X + Ua * (aPt2.X - aPt1.X)
        PtInt.Y = aPt1.Y + Ua * (aPt2.Y - aPt1.Y)

        'When looking for line intersections we are only interested in
        'Ua and Ub must both lie between 0 and 1!
        If Ua <= 0 Or Ua > 1 Then
            GetLineInt.X = -9999
            GetLineInt.Y = -9999
            Exit Function
        End If

        If Ub <= 0 Or Ub > 1 Then
            GetLineInt.X = -9999
            GetLineInt.Y = -9999
            Exit Function
        End If

        GetLineInt.X = PtInt.X
        GetLineInt.Y = PtInt.Y
    End Function

    Private Function GetNumericGridMidLoc(ByVal aSeX As Double, _
                                          ByVal aSeY As Double, _
                                          ByVal aXoffset As Integer, _
                                          ByVal aYoffset As Integer) As gPtCoordType


        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        GetNumericGridMidLoc.X = aSeX - (165 + (aXoffset - 1) * 330)
        GetNumericGridMidLoc.Y = aSeY + (165 + (aYoffset - 1) * 330)
    End Function

    Public Sub gGetHoleCoords(ByVal aLoc As String, _
                              ByVal aGridType As String, _
                              ByRef aXcoord As Double, _
                              ByRef aYcoord As Double)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo GetHoleCoordsError

        Dim X As Integer
        Dim Y As Integer
        Dim LettString As String
        Dim ErrStr As String

        'Coordinates for all holes in section are in:
        'gProspX(1 To 16, 1 To 16)
        'gProspY(1 To 16, 1 To 16)

        'Need to assign form variables: fXcoord and fYcoord

        If aGridType = "Alpha-numeric" Then
            'Make sure we have an alpha-numeric hole location.
            If IsNumeric(aLoc) = True Then
                aLoc = gGetHoleLoc2(aLoc, "Char")
            End If
            If aLoc = "???" Or aLoc = "" Then
                aXcoord = 0
                aYcoord = 0
                Exit Sub
            End If

            Y = Val(Mid(aLoc, 2))   '1 to 16

            LettString = "ABCDEFGHIJKLMNOP"

            For X = 1 To 16
                If Mid(LettString, X, 1) = Mid(aLoc, 1, 1) Then
                    Exit For
                End If
            Next X
        End If

        If aGridType = "Numeric" Then
            'Make sure we have a numeric hole location.
            If IsNumeric(aLoc) = False Then
                aLoc = gGetHoleLoc2(aLoc, "Num")
            End If
            If aLoc = "???" Or aLoc = "" Then
                aXcoord = 0
                aYcoord = 0
                Exit Sub
            End If

            If Val(Mid(aLoc, 1, 2)) >= 2 And Val(Mid(aLoc, 1, 2)) <= 32 And _
                gIsEvenNumber(Val(Mid(aLoc, 1, 2))) Then
                X = Val(Mid(aLoc, 1, 2)) / 2     '02 to 32
            Else
                X = 0
            End If

            If Val(Mid(aLoc, 3)) >= 34 And Val(Mid(aLoc, 3)) <= 64 And _
                gIsEvenNumber(Val(Mid(aLoc, 3))) Then
                Y = (Val(Mid(aLoc, 3)) - 32) / 2 '34 to 64
            Else
                Y = 0
            End If
        End If

        If X <> 0 And Y <> 0 Then
            aXcoord = gProspX(X, Y)
            aYcoord = gProspY(X, Y)
        Else
            aXcoord = 0
            aYcoord = 0
        End If

        Exit Sub

GetHoleCoordsError:
        ErrStr = Err.Description & " (Location = " & aLoc & ")"
        MsgBox("Error getting hole coordinates." & vbCrLf & _
            ErrStr, _
            vbOKOnly + vbExclamation, _
            "Hole Coordinate Get Error")
    End Sub


End Module
