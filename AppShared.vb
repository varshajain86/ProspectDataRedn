Public Class AppShared
    Friend Enum StatusPanel
        STATUS = 0
        USER = 1
        FACILITY = 2
        DATABASE = 3
    End Enum

    Friend Enum NagivationMode
        TOOLBAR = 1
        TREEVIEW = 2
    End Enum

    Friend Shared UI_DEFAULT_FACILITY As String
    Friend Shared UI_NAVIGATION_MODE As NagivationMode
    Friend Shared UI_STARTUP_PAGE As Long

    Friend Shared User As AppUser



    Friend Shared IsUserReadRole As Boolean
    Friend Shared IsUserWriteRole As Boolean
    Friend Shared IsUserSetUpRole As Boolean
    Friend Shared IsUserSuperAdminRole As Boolean
    Friend Shared IsUserAdminRole As Boolean



    Public Shared Sub InitializeWebServices()
        Debug.Print("Initializing OIS Web Services")
        My.WebServices.SecurityService.UseDefaultCredentials = True
        'My.WebServices.HierarchyService.UseDefaultCredentials = True
        'My.WebServices.UserPreferenceService.UseDefaultCredentials = True
    End Sub


    Friend Class AppUser
        Inherits SecurityWeb.User

        Private colRoles As Collections.Generic.IList(Of SecurityWeb.Role)

        Public Sub New(ByVal UserObj As SecurityWeb.User)
            Me.ID = UserObj.ID
            Me.Name = UserObj.Name
            colRoles = My.WebServices.SecurityService.GetListByUser(Me.ID)
        End Sub
        Public Function IsInRole(ByVal roleName As String) As Boolean
            Dim ReturnValue As Boolean = False
            If colRoles Is Nothing Then
                colRoles = My.WebServices.SecurityService.GetListByUser(Me.ID)
            End If
            For Each iRole As SecurityWeb.Role In colRoles
                If iRole.Name = roleName Then
                    ReturnValue = True
                    Exit For
                End If
            Next
            Return ReturnValue
        End Function

    End Class


End Class
