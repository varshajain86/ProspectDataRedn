Option Explicit On
Imports OracleInProcServer
Module Database
    Function RunSPReturnRS(ByVal strSP As String, _
                       ByVal ParamArray params() As Object) As OraDynaset

        '**********************************************************************
        '
        '
        '
        '**********************************************************************

        On Error GoTo errorHandler

        Dim SQLStmt As OraSqlStmt
        Dim dbparams As OraParameters
        Dim a As Integer

        dbparams = gOradatabase.Parameters
        collectParams(dbparams, params)

        ' Return the resultant recordset
        '    Set SQLStmt = gOradatabase.CreateSql _
        '           (strSP, ORASQL_FAILEXEC)

        'Oracle 7 code
        '    Set RunSPReturnRS = gOradatabase.CreatePlsqlDynaset(strSP, dpparams(dbparams.Count - 1).Name, ORADYN_DEFAULT)
        'Oracle 7 + 8  code

        RunSPReturnRS = gOradatabase.CreatePlsqlDynaset(strSP, params(dbparams.count)(0), ORADYN_DEFAULT)

        'Clear Params
        For a = 0 To dbparams.count - 1
            dbparams.Remove(0)
        Next

        Exit Function

errorHandler:
        'aiseError "Prodbininput", "RunSPReturnRS(" & strSP & ", ...)"
        For a = 0 To dbparams.count - 1
            dbparams.Remove(0)
        Next
    End Function

    Function RunSQLReturnRS(ByVal strSP As String, _
                            ByVal ParamArray params() As Object) As OraDynaset

        '**************************************************************************
        'Same as runspreturnrs except executes database method createdynaset rather
        'than createplsqldynaset.
        '
        '**************************************************************************

        On Error GoTo errorHandler

        Dim SQLStmt As OraSqlStmt
        Dim dbparams As OraParameters
        Dim a As Integer

        dbparams = gOradatabase.Parameters
        collectParams(dbparams, params)

        ' Return the resultant recordset
        '    Set SQLStmt = gOradatabase.CreateSql _
        '           (strSP, ORASQL_FAILEXEC)

        RunSQLReturnRS = gOradatabase.CreateDynaset(strSP, ORADYN_DEFAULT) ' dbparams(dbparams.Count - 1).Name, ORADYN_DEFAULT)

        'Clear Params
        For a = 0 To dbparams.count - 1
            dbparams.Remove(0)
        Next

        Exit Function

errorHandler:
        'aiseError "Prodbininput", "RunSPReturnRS(" & strSP & ", ...)"
    End Function

    Function RunSPReturnInt(ByVal strSP As String, _
                            ByVal ParamArray params() As Object) As Integer

        Dim SQLStmt As OraSqlStmt
        Dim dbparams As OraParameters
        Dim a As Integer

        dbparams = gOradatabase.Parameters
        collectParams(dbparams, params)

        'Return the resultant recordset
        SQLStmt = gOradatabase.CreateSql _
               (strSP, ORASQL_FAILEXEC)
        RunSPReturnInt = dbparams(dbparams.count - 1).Value

        'Clear Params
        For a = 0 To dbparams.count - 1
            dbparams.Remove(0)
        Next

        Exit Function
    End Function

    Function RunSPReturnIntLss(ByVal strSP As String, _
                               ByVal ParamArray params() As Object) As Integer

        '********************************************************************
        '   Assume first parameter is a number indicating batch array size.
        '
        '
        '********************************************************************

        Dim SQLStmt As OraSqlStmt
        Dim dbparams As OraParameters
        Dim a As Integer

        dbparams = gOradatabase.Parameters
        collectBatchParams(dbparams, params)

        'Return the resultant recordset
        SQLStmt = gOradatabase.CreateSql _
               (strSP, ORASQL_FAILEXEC)
        'Clear Params

        RunSPReturnIntLss = dbparams(dbparams.count - 1).Value

        For a = 0 To dbparams.count - 1
            dbparams.Remove(0)
        Next

        Exit Function

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

    Function RunSP(ByVal strSP As String, _
                   ByVal ParamArray params() As Object)

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

        'Clear Params
        For a = 0 To dbparams.count - 1
            dbparams.Remove(0)
        Next

        Exit Function

    End Function

    Sub RunBatchSP(ByVal strSP As String, _
                        ByVal ParamArray params() As Object)

        '********************************************************************
        '   Assume first parameter is a number indicating batch array size.
        '
        '
        '********************************************************************

        Dim SQLStmt As OraSqlStmt
        Dim dbparams As OraParameters = Nothing
        Dim a As Integer

        Try
            dbparams = gOradatabase.Parameters
            collectBatchParams(dbparams, params)

            'Return the resultant recordset
            SQLStmt = gOradatabase.CreateSql _
                   (strSP, ORASQL_FAILEXEC)
        Catch ex As Exception
            Throw ex
        Finally
            'Clear Params
            If Not dbparams Is Nothing Then
                For a = 0 To dbparams.Count - 1
                    dbparams.Remove(0)
                Next
            End If
        End Try

    End Sub

    Sub collectParams(ByRef paramlist As OraParameters, _
                      ByVal ParamArray argparams() As Object)

        '********************************************************************
        '
        '
        '
        '********************************************************************

        Dim params As Object

        params = argparams '(0)
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

    Sub collectBatchParams(ByRef paramlist As OraParameters, _
                           ByVal ParamArray argparams() As Object)

        '********************************************************************
        '   Assume first parameter is a number indicating batch array size.
        '
        '
        '********************************************************************

        Dim params As Object

        params = argparams '(0)
        Dim I As Integer
        Dim v As Object
        Dim LowParamIdx As Integer
        Dim ArrayIdx As Long   'Integer

        LowParamIdx = LBound(params)
        paramlist.Add(params(LowParamIdx)(0), params(LowParamIdx)(1), params(LowParamIdx)(2))
        paramlist(LowParamIdx).servertype = params(LowParamIdx)(3)

        For I = LBound(params) + 1 To UBound(params)
            If IsArray(params(I)(1)) = True Then
                If IsDBNull(params(I)(4)) Then
                    paramlist.AddTable(params(I)(0), params(I)(2), params(I)(3), _
                                        params(LowParamIdx)(1))
                Else
                    paramlist.AddTable(params(I)(0), params(I)(2), params(I)(3), _
                                         params(LowParamIdx)(1), params(I)(4))
                End If
                For ArrayIdx = 0 To params(LowParamIdx)(1) - 1
                    paramlist(params(I)(0)).put_Value(params(I)(1)(ArrayIdx), ArrayIdx)
                Next
            Else
                paramlist.Add(params(I)(0), params(I)(1), params(I)(2))
                paramlist(I).servertype = params(I)(3)
            End If
        Next I
        Exit Sub
    End Sub

End Module
