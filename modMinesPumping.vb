Imports System.Drawing.Printing
Imports Microsoft.Office.Interop


Module modMinesPumping
    Private mImage As Image

    Public Sub gPrintScreen(ByVal aHandle As System.IntPtr)

        Try
            'Want to capture the current form and print it.
            Dim sc As New ScreenShot.ScreenCapture
            mImage = sc.CaptureWindow(aHandle)

            'Have the image -- need to print it!
            Dim dlg As New PrintDialog
            Dim pd As New PrintDocument()
            dlg.Document = pd
            dlg.AllowSelection = True
            dlg.AllowSomePages = False

            'Determine if picture should be printed in landscape or portrait
            'and set the orientation.
            If mImage.Height >= mImage.Width Then
                pd.DefaultPageSettings.Landscape = False   'Taller than wide.
            Else
                pd.DefaultPageSettings.Landscape = True    'Wider than tall.
            End If

            AddHandler pd.PrintPage, AddressOf PrintDocument_PrintPage

            If (dlg.ShowDialog = System.Windows.Forms.DialogResult.OK) Then
                pd.Print()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub PrintDocument_PrintPage(ByVal sender As System.Object, _
                                        ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        Try
            'Page bounds are in 100ths of an inch.

            Dim PageWidth As Single = e.PageBounds.Width / 100
            Dim PageHeight As Single = e.PageBounds.Height / 100
            Dim PageWidthPixels As Integer = PageWidth * e.PageSettings.PrinterResolution.X
            Dim PageHeightPixels As Integer

            'Some printers only supply 1 value
            If e.PageSettings.PrinterResolution.Y = 0 Then
                PageHeightPixels = PageHeight * e.PageSettings.PrinterResolution.X
            Else
                PageHeightPixels = PageHeight * e.PageSettings.PrinterResolution.Y
            End If

            Dim WidthRatio As Double = PageWidthPixels / mImage.Width
            Dim HeightRatio As Double = PageHeightPixels / mImage.Height
            Dim ScalingFactor As Double

            '01/15/2009, lss
            'The scaling factor seems to be a "tiny" bit too large -- some of the image
            'does not fit -- I will stick a 0.95 correction factor in here for now!!
            If WidthRatio > HeightRatio Then
                ScalingFactor = HeightRatio * 0.95
            Else
                ScalingFactor = WidthRatio * 0.95
            End If

            Dim BmapTemp As New Bitmap(mImage)

            'Some printers only supply 1 value
            If e.PageSettings.PrinterResolution.Y = 0 Then
                BmapTemp.SetResolution(e.PageSettings.PrinterResolution.X / ScalingFactor, _
                       e.PageSettings.PrinterResolution.X / ScalingFactor)
            Else
                BmapTemp.SetResolution(e.PageSettings.PrinterResolution.X / ScalingFactor, _
                                       e.PageSettings.PrinterResolution.Y / ScalingFactor)
            End If

            e.Graphics.DrawImage(BmapTemp, 0, 0)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub gViewInExcel(ByVal aCommaDelimitedFile As String)
        Try
            Dim objXl As Excel.Application

            'Start Excel and get Application object.
            objXl = CreateObject("Excel.Application")
            objXl.Visible = True
            objXl.Workbooks.Open(aCommaDelimitedFile, , , , , , , , ",")

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    '******************************************************************************************
    Public Function GetIsEvenNumber(ByVal aValue As Double) As Boolean
        GetIsEvenNumber = False

        If aValue Mod 2 = 0 Then
            GetIsEvenNumber = True
        End If
    End Function


End Module
