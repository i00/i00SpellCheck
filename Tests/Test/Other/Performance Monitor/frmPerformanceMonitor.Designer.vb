<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPerformanceMonitor
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPerformanceMonitor))
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.tsiOpen = New System.Windows.Forms.ToolStripButton
        Me.ToolStripDropDownButton1 = New System.Windows.Forms.ToolStripDropDownButton
        Me.tsiAdd = New System.Windows.Forms.ToolStripDropDownButton
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.MenuTextSeperator1 = New i00SpellCheck.MenuTextSeperator
        Me.tsiGridBGColor = New i00SpellCheck.tsiColorPicker
        Me.MenuTextSeperator2 = New i00SpellCheck.MenuTextSeperator
        Me.tsiGridColor = New i00SpellCheck.tsiColorPicker
        Me.ClsGrid1 = New Test.clsGrid
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsiOpen, Me.ToolStripDropDownButton1, Me.tsiAdd})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 100)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(400, 25)
        Me.ToolStrip1.TabIndex = 0
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'tsiOpen
        '
        Me.tsiOpen.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsiOpen.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsiOpen.Image = CType(resources.GetObject("tsiOpen.Image"), System.Drawing.Image)
        Me.tsiOpen.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsiOpen.Name = "tsiOpen"
        Me.tsiOpen.Overflow = System.Windows.Forms.ToolStripItemOverflow.Never
        Me.tsiOpen.Size = New System.Drawing.Size(23, 22)
        Me.tsiOpen.Text = "Open Performance Monitor"
        Me.tsiOpen.Visible = False
        '
        'ToolStripDropDownButton1
        '
        Me.ToolStripDropDownButton1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripDropDownButton1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuTextSeperator1, Me.tsiGridBGColor, Me.MenuTextSeperator2, Me.tsiGridColor})
        Me.ToolStripDropDownButton1.Image = CType(resources.GetObject("ToolStripDropDownButton1.Image"), System.Drawing.Image)
        Me.ToolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDropDownButton1.Name = "ToolStripDropDownButton1"
        Me.ToolStripDropDownButton1.Overflow = System.Windows.Forms.ToolStripItemOverflow.Never
        Me.ToolStripDropDownButton1.Size = New System.Drawing.Size(29, 22)
        Me.ToolStripDropDownButton1.Text = "Add Performance Counter"
        '
        'tsiAdd
        '
        Me.tsiAdd.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsiAdd.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsiAdd.Image = CType(resources.GetObject("tsiAdd.Image"), System.Drawing.Image)
        Me.tsiAdd.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsiAdd.Name = "tsiAdd"
        Me.tsiAdd.Overflow = System.Windows.Forms.ToolStripItemOverflow.Never
        Me.tsiAdd.Size = New System.Drawing.Size(29, 22)
        Me.tsiAdd.Text = "Add Performance Counter"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1
        '
        'MenuTextSeperator1
        '
        Me.MenuTextSeperator1.AutoSize = False
        Me.MenuTextSeperator1.BackColor = System.Drawing.SystemColors.Control
        Me.MenuTextSeperator1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MenuTextSeperator1.Name = "MenuTextSeperator1"
        Me.MenuTextSeperator1.Size = New System.Drawing.Size(1920, 18)
        Me.MenuTextSeperator1.Text = "Background"
        '
        'tsiGridBGColor
        '
        Me.tsiGridBGColor.AutoSize = False
        Me.tsiGridBGColor.Colors = CType(resources.GetObject("tsiGridBGColor.Colors"), System.Collections.Generic.List(Of System.Drawing.Color))
        Me.tsiGridBGColor.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.None
        Me.tsiGridBGColor.Name = "tsiGridBGColor"
        Me.tsiGridBGColor.Persistent = True
        Me.tsiGridBGColor.SelectedColor = System.Drawing.Color.Empty
        '
        'MenuTextSeperator2
        '
        Me.MenuTextSeperator2.AutoSize = False
        Me.MenuTextSeperator2.BackColor = System.Drawing.SystemColors.Control
        Me.MenuTextSeperator2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MenuTextSeperator2.Name = "MenuTextSeperator2"
        Me.MenuTextSeperator2.Size = New System.Drawing.Size(1920, 18)
        Me.MenuTextSeperator2.Text = "Grid"
        '
        'tsiGridColor
        '
        Me.tsiGridColor.AutoSize = False
        Me.tsiGridColor.Colors = CType(resources.GetObject("tsiGridColor.Colors"), System.Collections.Generic.List(Of System.Drawing.Color))
        Me.tsiGridColor.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.None
        Me.tsiGridColor.Name = "tsiGridColor"
        Me.tsiGridColor.Persistent = True
        Me.tsiGridColor.SelectedColor = System.Drawing.Color.Empty
        '
        'ClsGrid1
        '
        Me.ClsGrid1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ClsGrid1.Location = New System.Drawing.Point(0, 0)
        Me.ClsGrid1.Name = "ClsGrid1"
        Me.ClsGrid1.Size = New System.Drawing.Size(400, 100)
        Me.ClsGrid1.TabIndex = 1
        '
        'frmPerformanceMonitor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(400, 125)
        Me.Controls.Add(Me.ClsGrid1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.MinimumSize = New System.Drawing.Size(410, 138)
        Me.Name = "frmPerformanceMonitor"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Performance Monitor"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsiOpen As System.Windows.Forms.ToolStripButton
    Friend WithEvents ClsGrid1 As Test.clsGrid
    Friend WithEvents tsiAdd As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents ToolStripDropDownButton1 As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents tsiGridBGColor As i00SpellCheck.tsiColorPicker
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents MenuTextSeperator1 As i00SpellCheck.MenuTextSeperator
    Friend WithEvents MenuTextSeperator2 As i00SpellCheck.MenuTextSeperator
    Friend WithEvents tsiGridColor As i00SpellCheck.tsiColorPicker
End Class
