Imports OracleInProcServer

Public Class frmMDIMain
    Public frmPDR As frmProspDataReduction
    Public frmPDHR As frmProspDataHoleReduction

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        AppShared.InitializeWebServices()
        GetUser()
        'Check userid to ensure this user is setup
        If AppShared.User.ID = 0 Then
            MessageBox.Show(AppShared.User.Name & " is not setup as an MOIS user. Please call the helpdesk to have your account setup")
            End
        End If
        InitializeFromPreferences()
    End Sub

    Private Sub GetUser()
        Try
            Dim lst As IList(Of SecurityWeb.User) = My.WebServices.SecurityService.GetUserListByUserName(System.Security.Principal.WindowsIdentity.GetCurrent.Name.ToUpper)
            If lst.Count = 1 Then
                AppShared.User = New AppShared.AppUser(lst.Item(0))
            End If
            AppShared.IsUserSuperAdminRole = AppShared.User.IsInRole("MOISLegacyAdmin")
            AppShared.IsUserReadRole = AppShared.User.IsInRole("ReducedProspRead")
            AppShared.IsUserWriteRole = AppShared.User.IsInRole("ReducedProspWrite")
            AppShared.IsUserAdminRole = AppShared.User.IsInRole("ReducedProspAdmin")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub InitializeFromPreferences()
        Try
            'There needs to be settings for the following at a minimum:
            'UI_DEFAULT_FACILITY
            'If these are not set for the user, set them to a default value

            'upi = My.WebServices.UserPreferenceService.GetUserPreference(AppShared.User.ID, "UI_DEFAULT_FACILITY")
            'If upi.PreferenceValue Is Nothing Then
            '    Dim ds As New HierarchyWeb.BusinessEntityDS
            '    ds = My.WebServices.HierarchyService.GetBusinessEntityDSByHierarchyAndType("Plant", 1)
            '    If ds.Tables(0).Rows.Count > 0 Then
            '        AppShared.UI_DEFAULT_FACILITY = CType(ds.Tables(0).Rows(0), HierarchyWeb.BusinessEntityDS.BusinessEntityRow).Name
            '    End If
            'Else
            '    AppShared.UI_DEFAULT_FACILITY = upi.PreferenceValue
            'End If
            AppShared.UI_DEFAULT_FACILITY = "Four Corners"
        Catch ex As Exception
            MessageBox.Show("Exception: " & ex.Message)
        End Try
    End Sub

    Private Sub frmMDIMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            If Login() Then
                gUserName = AppShared.User.Name.Substring(AppShared.User.Name.IndexOf("\") + 1)
                gActiveMineNameLong = AppShared.UI_DEFAULT_FACILITY
                gActiveMineNameRkBkRpts = AppShared.UI_DEFAULT_FACILITY
                Select Case gActiveMineNameLong
                    Case Is = "South Fort Meade"
                        gActiveMineNameShort = "SF"
                    Case Is = "Hookers Prairie"
                        gActiveMineNameShort = "HP"
                    Case Is = "Wingate"
                        gActiveMineNameShort = "WG"
                    Case Is = "Four Corners"
                        gActiveMineNameShort = "FCO"
                    Case Is = "Fort Green"
                        gActiveMineNameShort = "FTG"
                    Case Is = "Hopewell"
                        gActiveMineNameShort = "HOP"
                End Select
                gDxfDefaultFile = "C:\TempMois\tempdxf.dxf"
                gDatDefaultFile = "C:\TempMois\tempdat.dat"
                gTxtDefaultFile = "C:\TempMois\temptxt.txt"

                gGenDelayComment = "Always enter the TOTAL hours that the delay reason " &
                   "occurred for the shift (adding up the hours for all " &
                   " occurrences of " &
                   "the delay reason for the shift).  You may enter the number " &
                   "of times the delay occured during the shift in the " &
                   "'Occur' column, however MOIS will not multiply the 'Occur' value " &
                   "by the delay reason hours that you enter!"

                UpdateLastLogon()

                gFyrBeginDate = #12/31/8888#
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Function Login() As Boolean
        Dim UserName As String = String.Empty
        Dim password As String = String.Empty
        Dim DataSource As String = String.Empty
        Dim PosNum As Single

        Try
            'All MOIS users will log into the p_site database as:
            PosNum = 1
            UserName = "moismain" '"mois" '
            password = "matthew98" '"legs2" '
            gOracleUserName = UserName
            gOracleUserPassword = password

            'ToDo: Check for Admin role and display the select Database Form

            If AppShared.IsUserSuperAdminRole Then
                Dim frmDataSource = New frmDataSourcePick() With {.DataSource = My.Settings.DataSourceDefault}
                frmDataSource.ShowDialog()
                If frmDataSource.DialogResult = DialogResult.OK Then
                    DataSource = frmDataSource.DataSource
                End If
            End If
            If DataSource = String.Empty Then
                'DataSource = "DSITE3" '"PSITE" ' .MOSAICCO.COM cboDatabase.Text
                DataSource = My.Settings.DataSourceDefault
            End If

            Dim DataEnviroment As String = String.Empty
            Dim CodeEnviroment As String = String.Empty

            If DataSource = My.Settings.DataSourceDev Then
                DataEnviroment = "Test Data"
            ElseIf DataSource = My.Settings.DataSourceProd Then
                DataEnviroment = "Production Data"
            End If
            If My.Settings.DataSourceDefault = My.Settings.DataSourceDev Then
                CodeEnviroment = "QA Application"
            ElseIf My.Settings.DataSourceDefault = My.Settings.DataSourceProd Then
                CodeEnviroment = "Production Application"
            End If

            If AppShared.IsUserSuperAdminRole Then
                Me.Text = CodeEnviroment & " - " & DataEnviroment
            Else
                Me.Text = CodeEnviroment
            End If

            PosNum = 2
            gOraSession = CreateObject("OracleInProcServer.XOrasession")

            PosNum = 3
            gOradatabase = gOraSession.OpenDatabase(DataSource, String.Format("{0}/{1}", UserName, password), &H0&)

            PosNum = 4
            gDBParams = gOradatabase.Parameters

            gConnected = True
            gDataSource = DataSource

            Return True

        Catch ex As Exception
            'If login into database is unsuccessful then the user will get
            'a error message from Oracle.
            MsgBox("Login Error: Unable to log into MOIS.  " &
                   Err.Description & "   Pos# " & CStr(PosNum))
            Return False
        End Try

    End Function

    Private Sub UpdateLastLogon()
        Dim params As OraParameters = gDBParams
        Dim LastLogonDateTime As Object
        Dim DateStr As String = String.Format("{0} {1}", Today.ToShortDateString, Now.ToShortTimeString)

        Try

            'Need to create a datetime from fThisDate & fThisTime
            LastLogonDateTime = CDate(DateStr)
            params.Add("pUserId", StrConv(gUserName, vbUpperCase), ORAPARM_INPUT)
            params("pUserId").serverType = ORATYPE_VARCHAR2
            params.Add("pLastLogon", LastLogonDateTime, ORAPARM_INPUT)
            params("pLastLogon").serverType = ORATYPE_DATE
            Dim SQLStmt As OraSqlStmt = gOradatabase.CreateSql("Begin mois.mois_pwords.update_lastlogon(:pUserId, " +
                          ":pLastLogon);end;", ORASQL_FAILEXEC)
            ClearParams(params)
        Catch
            'An error here is not critical -- don't display anything to the
            'user!
        Finally
            ClearParams(params)
        End Try

    End Sub
    Private Sub DataReductionmultiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataReductionmultiToolStripMenuItem.Click
        Try
            Me.Cursor = Cursors.WaitCursor
            If frmPDR Is Nothing OrElse frmPDR.IsDisposed Then
                frmPDR = New frmProspDataReduction
                frmPDR.MdiParent = Me

            End If
            frmPDR.Dock = DockStyle.Fill
            frmPDR.Show()
            frmPDR.Activate()
            Me.Cursor = Cursors.Arrow
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Me.Cursor = Cursors.Arrow
        End Try
    End Sub

    Private Sub DataHoleReductionsingleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataHoleReductionsingleToolStripMenuItem.Click
        Try
            Me.Cursor = Cursors.WaitCursor
            If frmPDHR Is Nothing OrElse frmPDHR.IsDisposed Then
                frmPDHR = New frmProspDataHoleReduction
                frmPDHR.MdiParent = Me
            End If
            'Dim x As Integer = (Me.Width / 2) - (frmPDHR.Width / 2)
            'Dim y As Integer = (Me.Height / 2.2) - (frmPDHR.Height / 2) '/// allow extra for the toolbox ( hence 2.2 )  
            'frmPDHR.Location = New Point(x, y) '///Center the form in it's parent. 
            With frmPDHR
                .Dock = DockStyle.Fill
                .Show()
                .Activate()
            End With
            Me.Cursor = Cursors.Arrow
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Me.Cursor = Cursors.Arrow
        End Try
    End Sub
End Class