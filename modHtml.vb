Imports AxFPSpread

Module modHtml
    Public Function gCreateHtmlTable(ByRef aGrid As AxvaSpread,
                                 ByVal aFileNum As Integer,
                                 ByVal aTableBorder As String,
                                 ByVal aIncludeHeaders As Integer)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        'aTableBorder will be "Single" or "Double"

        Dim BeginRow As Integer
        Dim NumCols As Integer
        Dim Headings As Boolean
        Dim RowIdx As Long
        Dim ColIdx As Integer
        Dim TableWidth As Integer
        Dim TableDesc As String
        Dim ColWidth() As Single
        Dim TotWidth As Single
        Dim ThisWidth As Single
        Dim Title As String

        NumCols = aGrid.MaxCols
        ReDim ColWidth(NumCols)

        If aIncludeHeaders = 1 Then
            BeginRow = 0
        Else
            BeginRow = 1
        End If

        'HTML report header
        Title = " "
        gWriteLine(aFileNum, "<html>")
        gWriteLine(aFileNum, "<head>")
        gWriteLine(aFileNum, "<title>" + Title + "</title>")
        gWriteLine(aFileNum, "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>")
        gWriteLine(aFileNum, "</head>")
        gWriteLine(aFileNum, "<body bgcolor='#FFFFFF' text='#000000'>")

        'Figure out some table things first
        TableWidth = 670

        'Figure out column widths
        With aGrid
            .Row = 1
            TotWidth = 0
            For ColIdx = 0 To .MaxCols
                .Col = ColIdx
                TotWidth = TotWidth + .get_ColWidth(ColIdx)
            Next ColIdx

            For ColIdx = 0 To .MaxCols
                .Col = ColIdx
                ThisWidth = .get_ColWidth(ColIdx)
                ColWidth(ColIdx) = Int((ThisWidth / TotWidth) * TableWidth)
            Next ColIdx

            TableWidth = 0
            For ColIdx = 0 To UBound(ColWidth)
                TableWidth = TableWidth + ColWidth(ColIdx)
            Next ColIdx
        End With

        TableDesc = "<table width='" & CStr(TableWidth) & "' border='1' cellpadding='1' " &
                    "cellspacing='1' bordercolor='#000000'>"

        gWriteLine(aFileNum, TableDesc)

        With aGrid
            For RowIdx = BeginRow To .MaxRows
                If aIncludeHeaders = 1 And RowIdx = BeginRow Then
                    Headings = True
                Else
                    Headings = False
                End If

                AddHtmlLine(aGrid, aFileNum, RowIdx, Headings, ColWidth)
            Next RowIdx
        End With
        gWriteLine(aFileNum, "</table>")

        'Report footer
        gWriteLine(aFileNum, "</body>")
        gWriteLine(aFileNum, "</html>")
    End Function

    Private Sub AddHtmlLine(ByRef aGrid As AxvaSpread,
                            ByVal aFileNum As Integer,
                            ByVal aRowIdx As Long,
                            ByVal aHeadings As Boolean,
                            ByRef aColWidth() As Single)

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        Dim TextPropsCenter As String
        Dim TextPropsRight As String
        Dim TextPropsLeft As String
        Dim FontSize As Integer
        Dim ColIdx As Integer
        Dim GridValue As String

        FontSize = 2

        'For row 0 headings -- center
        TextPropsCenter = " align='center' valign='middle' bgcolor='#DBDBDB'>" &
                          "<font face='Times New Roman, Times, serif' size='" &
                          CStr(FontSize) & "'>"

        'For non-headings -- right
        TextPropsRight = " align='right' valign='middle'><font " &
                         "face='Times New Roman, Times, serif' size='" &
                         CStr(FontSize) & "'>"

        'For column 0 headings -- left
        TextPropsLeft = " align='left' valign='middle' bgcolor='#DBDBDB'>" &
                        "<font face='Times New Roman, Times, serif' size='" &
                        CStr(FontSize) & "'>"

        If aHeadings = True Then
            With aGrid
                .Row = aRowIdx
                gWriteLine(aFileNum, "<tr>")
                For ColIdx = 0 To .MaxCols
                    .Col = ColIdx
                    GridValue = Trim(.Text)
                    If GridValue = "" Then
                        GridValue = "--"
                    End If
                    gWriteLine(aFileNum, "<td width='" &
                               CStr(aColWidth(ColIdx)) & "'" & TextPropsCenter &
                               "<b>" & GridValue & "</b></font></td>")
                Next ColIdx
                gWriteLine(aFileNum, "</tr>")
            End With
        Else
            With aGrid
                .Row = aRowIdx
                gWriteLine(aFileNum, "<tr>")
                For ColIdx = 0 To .MaxCols
                    .Col = ColIdx
                    GridValue = Trim(.Text)
                    If GridValue = "" Then
                        GridValue = "--"
                    End If
                    If .Col = 0 Then
                        gWriteLine(aFileNum, "<td width='" &
                                   CStr(aColWidth(ColIdx)) & "'" & TextPropsLeft &
                                   "<b>" & GridValue & "</b></font></td>")
                    Else
                        gWriteLine(aFileNum, "<td width='" &
                                   CStr(aColWidth(ColIdx)) & "'" & TextPropsRight &
                                   "<b>" & GridValue & "</b></font></td>")
                    End If
                Next ColIdx
                gWriteLine(aFileNum, "</tr>")
            End With
        End If
    End Sub

End Module
