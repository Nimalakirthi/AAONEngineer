<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUnitOutput
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.UnitReportViewer = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'UnitReportViewer
        '
        Me.UnitReportViewer.ActiveViewIndex = -1
        Me.UnitReportViewer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.UnitReportViewer.DisplayGroupTree = False
        Me.UnitReportViewer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UnitReportViewer.Location = New System.Drawing.Point(0, 0)
        Me.UnitReportViewer.Name = "UnitReportViewer"
        Me.UnitReportViewer.SelectionFormula = ""
        Me.UnitReportViewer.Size = New System.Drawing.Size(725, 666)
        Me.UnitReportViewer.TabIndex = 0
        Me.UnitReportViewer.ViewTimeSelectionFormula = ""
        '
        'frmUnitOutput
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(725, 666)
        Me.Controls.Add(Me.UnitReportViewer)
        Me.Name = "frmUnitOutput"
        Me.Text = "frmUnitOutput"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UnitReportViewer As CrystalDecisions.Windows.Forms.CrystalReportViewer
End Class
