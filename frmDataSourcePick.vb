Public Class frmDataSourcePick


    Private Sub frmDataSourcePick_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.cboDatabase.Items.AddRange(New Object() {My.Settings.DataSourceDev, My.Settings.DataSourceProd})
        cboDatabase.SelectedItem = My.Settings.DataSourceDefault
    End Sub

    Public Property DataSource As String

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Me.DialogResult = DialogResult.OK
        _DataSource = cboDatabase.SelectedItem.ToString
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class