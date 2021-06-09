<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMDIMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMDIMain))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ProspectDataReductionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataReductionmultiToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataHoleReductionsingleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProspectDataReductionToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(969, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ProspectDataReductionToolStripMenuItem
        '
        Me.ProspectDataReductionToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DataReductionmultiToolStripMenuItem, Me.DataHoleReductionsingleToolStripMenuItem})
        Me.ProspectDataReductionToolStripMenuItem.Name = "ProspectDataReductionToolStripMenuItem"
        Me.ProspectDataReductionToolStripMenuItem.Size = New System.Drawing.Size(65, 20)
        Me.ProspectDataReductionToolStripMenuItem.Text = "Prospect"
        '
        'DataReductionmultiToolStripMenuItem
        '
        Me.DataReductionmultiToolStripMenuItem.Name = "DataReductionmultiToolStripMenuItem"
        Me.DataReductionmultiToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.DataReductionmultiToolStripMenuItem.Text = "Data Reduction (multi)"
        '
        'DataHoleReductionsingleToolStripMenuItem
        '
        Me.DataHoleReductionsingleToolStripMenuItem.Name = "DataHoleReductionsingleToolStripMenuItem"
        Me.DataHoleReductionsingleToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.DataHoleReductionsingleToolStripMenuItem.Text = "Data Hole Reduction (single)"
        '
        'frmMDIMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(969, 494)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmMDIMain"
        Me.Text = "QA - TSITE"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ProspectDataReductionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DataReductionmultiToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DataHoleReductionsingleToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
