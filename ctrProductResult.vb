Public Class ctrProductResult
    Public Sub New(ByVal productResults As List(Of ViewModels.SplitProductResult))

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        If productResults.Count > 0 Then
            Dim ProductType As String = [Enum].GetName(GetType(ViewModels.ProductTypes), productResults(0).ProdType)
            For Each col As DevExpress.XtraGrid.Columns.GridColumn In Me.GridView1.Columns
                Dim ColCap As String = col.Caption.ToString()
                If Not (ColCap.Equals("T-R-S Hole") Or ColCap.Equals("Split Number")) Then
                    col.Caption = ProductType & " " & ColCap
                End If
            Next
            SplitProductResultBindingSource.DataSource = productResults
            Me.GridView1.OptionsSelection.EnableAppearanceFocusedRow = False
        End If
    End Sub
End Class
