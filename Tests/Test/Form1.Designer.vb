<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.tslStatus = New System.Windows.Forms.ToolStripLabel
        Me.tsbCodeProject = New System.Windows.Forms.ToolStripButton
        Me.tsbVBForums = New System.Windows.Forms.ToolStripButton
        Me.tsbi00Productions = New System.Windows.Forms.ToolStripButton
        Me.tsbDonate = New System.Windows.Forms.ToolStripButton
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.ToolStrip2 = New System.Windows.Forms.ToolStrip
        Me.ToolStripDropDownButton1 = New System.Windows.Forms.ToolStripDropDownButton
        Me.CrosswordSolverToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.MenuTextSeperator3 = New i00SpellCheck.MenuTextSeperator
        Me.tsbSuggestionLookup = New System.Windows.Forms.ToolStripTextBox
        Me.MenuTextSeperator1 = New i00SpellCheck.MenuTextSeperator
        Me.tstbAnagramLookup = New System.Windows.Forms.ToolStripTextBox
        Me.MenuTextSeperator2 = New i00SpellCheck.MenuTextSeperator
        Me.tstbScrabbleHelper = New System.Windows.Forms.ToolStripTextBox
        Me.tsbAbout = New System.Windows.Forms.ToolStripButton
        Me.ToolStripDropDownButton2 = New System.Windows.Forms.ToolStripDropDownButton
        Me.MenuTextSeperator4 = New i00SpellCheck.MenuTextSeperator
        Me.SpellingErrorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.tsiCPSpellingError = New i00SpellCheck.tsiColorPicker
        Me.CaseErrorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.tsiCPCaseError = New i00SpellCheck.tsiColorPicker
        Me.IgnoredWordToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.tsiCPIgnoreColor = New i00SpellCheck.tsiColorPicker
        Me.MenuTextSeperator5 = New i00SpellCheck.MenuTextSeperator
        Me.tsiDrawStyle = New System.Windows.Forms.ToolStripComboBox
        Me.MenuTextSeperator6 = New i00SpellCheck.MenuTextSeperator
        Me.ShowErrorsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ShowIgnoredToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.WhenCtrlIsPressedToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AlwaysToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.NeverToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.tsbSpellCheck = New System.Windows.Forms.ToolStripButton
        Me.tsbProperties = New System.Windows.Forms.ToolStripButton
        Me.tabSpellControls = New System.Windows.Forms.TabControl
        Me.tabRichTextBox = New System.Windows.Forms.TabPage
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.propRichTextBox = New System.Windows.Forms.PropertyGrid
        Me.tabTextBox = New System.Windows.Forms.TabPage
        Me.propTextBox = New System.Windows.Forms.PropertyGrid
        Me.tabDataGridView = New System.Windows.Forms.TabPage
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.BindingNavigator1 = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorAddNewItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel
        Me.BindingNavigatorDeleteItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator
        Me.propDataGridView = New System.Windows.Forms.PropertyGrid
        Me.ToolStrip3 = New System.Windows.Forms.ToolStrip
        Me.tsiEnabled = New System.Windows.Forms.ToolStripButton
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ToolStrip1.SuspendLayout()
        Me.ToolStrip2.SuspendLayout()
        Me.tabSpellControls.SuspendLayout()
        Me.tabRichTextBox.SuspendLayout()
        Me.tabTextBox.SuspendLayout()
        Me.tabDataGridView.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BindingNavigator1.SuspendLayout()
        Me.ToolStrip3.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslStatus, Me.tsbCodeProject, Me.tsbVBForums, Me.tsbi00Productions, Me.tsbDonate})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 417)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(624, 25)
        Me.ToolStrip1.TabIndex = 3
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'tslStatus
        '
        Me.tslStatus.Name = "tslStatus"
        Me.tslStatus.Size = New System.Drawing.Size(86, 22)
        Me.tslStatus.Text = "i00 Spell Check"
        '
        'tsbCodeProject
        '
        Me.tsbCodeProject.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbCodeProject.AutoToolTip = False
        Me.tsbCodeProject.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbCodeProject.Name = "tsbCodeProject"
        Me.tsbCodeProject.Size = New System.Drawing.Size(79, 22)
        Me.tsbCodeProject.Text = "Code Project"
        '
        'tsbVBForums
        '
        Me.tsbVBForums.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbVBForums.AutoToolTip = False
        Me.tsbVBForums.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbVBForums.Name = "tsbVBForums"
        Me.tsbVBForums.Size = New System.Drawing.Size(68, 22)
        Me.tsbVBForums.Text = "VB Forums"
        '
        'tsbi00Productions
        '
        Me.tsbi00Productions.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbi00Productions.AutoToolTip = False
        Me.tsbi00Productions.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbi00Productions.Name = "tsbi00Productions"
        Me.tsbi00Productions.Size = New System.Drawing.Size(93, 22)
        Me.tsbi00Productions.Text = "i00 Productions"
        '
        'tsbDonate
        '
        Me.tsbDonate.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbDonate.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbDonate.Name = "tsbDonate"
        Me.tsbDonate.Size = New System.Drawing.Size(49, 22)
        Me.tsbDonate.Text = "Donate"
        '
        'TextBox1
        '
        Me.TextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(3, 3)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(352, 335)
        Me.TextBox1.TabIndex = 4
        Me.TextBox1.Text = "Ths is a standrd text field that uses a dictionary to spel check the its contents" & _
            " ...  as you can se errors are underlnied in red!"
        '
        'ToolStrip2
        '
        Me.ToolStrip2.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripDropDownButton1, Me.tsbAbout, Me.ToolStripDropDownButton2})
        Me.ToolStrip2.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip2.Name = "ToolStrip2"
        Me.ToolStrip2.Size = New System.Drawing.Size(624, 25)
        Me.ToolStrip2.TabIndex = 5
        Me.ToolStrip2.Text = "ToolStrip2"
        '
        'ToolStripDropDownButton1
        '
        Me.ToolStripDropDownButton1.AutoToolTip = False
        Me.ToolStripDropDownButton1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CrosswordSolverToolStripMenuItem, Me.MenuTextSeperator3, Me.tsbSuggestionLookup, Me.MenuTextSeperator1, Me.tstbAnagramLookup, Me.MenuTextSeperator2, Me.tstbScrabbleHelper})
        Me.ToolStripDropDownButton1.Image = CType(resources.GetObject("ToolStripDropDownButton1.Image"), System.Drawing.Image)
        Me.ToolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDropDownButton1.Name = "ToolStripDropDownButton1"
        Me.ToolStripDropDownButton1.Size = New System.Drawing.Size(127, 22)
        Me.ToolStripDropDownButton1.Text = "Other Examples..."
        '
        'CrosswordSolverToolStripMenuItem
        '
        Me.CrosswordSolverToolStripMenuItem.Name = "CrosswordSolverToolStripMenuItem"
        Me.CrosswordSolverToolStripMenuItem.Size = New System.Drawing.Size(185, 22)
        Me.CrosswordSolverToolStripMenuItem.Text = "Crossword Generator"
        '
        'MenuTextSeperator3
        '
        Me.MenuTextSeperator3.AutoSize = False
        Me.MenuTextSeperator3.BackColor = System.Drawing.SystemColors.Control
        Me.MenuTextSeperator3.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.MenuTextSeperator3.Name = "MenuTextSeperator3"
        Me.MenuTextSeperator3.Size = New System.Drawing.Size(1920, 18)
        Me.MenuTextSeperator3.Text = "Suggestion Lookup:"
        '
        'tsbSuggestionLookup
        '
        Me.tsbSuggestionLookup.Name = "tsbSuggestionLookup"
        Me.tsbSuggestionLookup.Size = New System.Drawing.Size(100, 23)
        '
        'MenuTextSeperator1
        '
        Me.MenuTextSeperator1.AutoSize = False
        Me.MenuTextSeperator1.BackColor = System.Drawing.SystemColors.Control
        Me.MenuTextSeperator1.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.MenuTextSeperator1.Name = "MenuTextSeperator1"
        Me.MenuTextSeperator1.Size = New System.Drawing.Size(1920, 18)
        Me.MenuTextSeperator1.Text = "Anagram Lookup:"
        '
        'tstbAnagramLookup
        '
        Me.tstbAnagramLookup.Name = "tstbAnagramLookup"
        Me.tstbAnagramLookup.Size = New System.Drawing.Size(100, 23)
        '
        'MenuTextSeperator2
        '
        Me.MenuTextSeperator2.AutoSize = False
        Me.MenuTextSeperator2.BackColor = System.Drawing.SystemColors.Control
        Me.MenuTextSeperator2.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.MenuTextSeperator2.Name = "MenuTextSeperator2"
        Me.MenuTextSeperator2.Size = New System.Drawing.Size(1920, 18)
        Me.MenuTextSeperator2.Text = "Scrabble Helper:"
        '
        'tstbScrabbleHelper
        '
        Me.tstbScrabbleHelper.Name = "tstbScrabbleHelper"
        Me.tstbScrabbleHelper.Size = New System.Drawing.Size(100, 23)
        '
        'tsbAbout
        '
        Me.tsbAbout.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbAbout.AutoToolTip = False
        Me.tsbAbout.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbAbout.Name = "tsbAbout"
        Me.tsbAbout.Size = New System.Drawing.Size(44, 22)
        Me.tsbAbout.Text = "About"
        '
        'ToolStripDropDownButton2
        '
        Me.ToolStripDropDownButton2.AutoToolTip = False
        Me.ToolStripDropDownButton2.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuTextSeperator4, Me.SpellingErrorToolStripMenuItem, Me.CaseErrorToolStripMenuItem, Me.IgnoredWordToolStripMenuItem, Me.MenuTextSeperator5, Me.tsiDrawStyle, Me.MenuTextSeperator6, Me.ShowErrorsToolStripMenuItem, Me.ShowIgnoredToolStripMenuItem})
        Me.ToolStripDropDownButton2.Image = CType(resources.GetObject("ToolStripDropDownButton2.Image"), System.Drawing.Image)
        Me.ToolStripDropDownButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDropDownButton2.Name = "ToolStripDropDownButton2"
        Me.ToolStripDropDownButton2.Size = New System.Drawing.Size(92, 22)
        Me.ToolStripDropDownButton2.Text = "Customize"
        '
        'MenuTextSeperator4
        '
        Me.MenuTextSeperator4.AutoSize = False
        Me.MenuTextSeperator4.BackColor = System.Drawing.SystemColors.Control
        Me.MenuTextSeperator4.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.MenuTextSeperator4.Name = "MenuTextSeperator4"
        Me.MenuTextSeperator4.Size = New System.Drawing.Size(1920, 18)
        Me.MenuTextSeperator4.Text = "Colors:"
        '
        'SpellingErrorToolStripMenuItem
        '
        Me.SpellingErrorToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsiCPSpellingError})
        Me.SpellingErrorToolStripMenuItem.Name = "SpellingErrorToolStripMenuItem"
        Me.SpellingErrorToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.SpellingErrorToolStripMenuItem.Text = "Spelling Error"
        '
        'tsiCPSpellingError
        '
        Me.tsiCPSpellingError.AutoSize = False
        Me.tsiCPSpellingError.Colors = CType(resources.GetObject("tsiCPSpellingError.Colors"), System.Collections.Generic.List(Of System.Drawing.Color))
        Me.tsiCPSpellingError.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.None
        Me.tsiCPSpellingError.Name = "tsiCPSpellingError"
        Me.tsiCPSpellingError.Persistent = True
        Me.tsiCPSpellingError.SelectedColor = System.Drawing.Color.Red
        '
        'CaseErrorToolStripMenuItem
        '
        Me.CaseErrorToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsiCPCaseError})
        Me.CaseErrorToolStripMenuItem.Name = "CaseErrorToolStripMenuItem"
        Me.CaseErrorToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.CaseErrorToolStripMenuItem.Text = "Case Error"
        '
        'tsiCPCaseError
        '
        Me.tsiCPCaseError.AutoSize = False
        Me.tsiCPCaseError.Colors = CType(resources.GetObject("tsiCPCaseError.Colors"), System.Collections.Generic.List(Of System.Drawing.Color))
        Me.tsiCPCaseError.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.None
        Me.tsiCPCaseError.Name = "tsiCPCaseError"
        Me.tsiCPCaseError.Persistent = True
        Me.tsiCPCaseError.SelectedColor = System.Drawing.Color.Green
        '
        'IgnoredWordToolStripMenuItem
        '
        Me.IgnoredWordToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsiCPIgnoreColor})
        Me.IgnoredWordToolStripMenuItem.Name = "IgnoredWordToolStripMenuItem"
        Me.IgnoredWordToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.IgnoredWordToolStripMenuItem.Text = "Ignored Word"
        '
        'tsiCPIgnoreColor
        '
        Me.tsiCPIgnoreColor.AutoSize = False
        Me.tsiCPIgnoreColor.Colors = CType(resources.GetObject("tsiCPIgnoreColor.Colors"), System.Collections.Generic.List(Of System.Drawing.Color))
        Me.tsiCPIgnoreColor.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.None
        Me.tsiCPIgnoreColor.Name = "tsiCPIgnoreColor"
        Me.tsiCPIgnoreColor.Persistent = True
        Me.tsiCPIgnoreColor.SelectedColor = System.Drawing.Color.Blue
        '
        'MenuTextSeperator5
        '
        Me.MenuTextSeperator5.AutoSize = False
        Me.MenuTextSeperator5.BackColor = System.Drawing.SystemColors.Control
        Me.MenuTextSeperator5.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.MenuTextSeperator5.Name = "MenuTextSeperator5"
        Me.MenuTextSeperator5.Size = New System.Drawing.Size(1920, 18)
        Me.MenuTextSeperator5.Text = "Style:"
        '
        'tsiDrawStyle
        '
        Me.tsiDrawStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.tsiDrawStyle.Items.AddRange(New Object() {"Default", "Boxed In", "Opera", "Draft Plan"})
        Me.tsiDrawStyle.Name = "tsiDrawStyle"
        Me.tsiDrawStyle.Size = New System.Drawing.Size(121, 23)
        '
        'MenuTextSeperator6
        '
        Me.MenuTextSeperator6.AutoSize = False
        Me.MenuTextSeperator6.BackColor = System.Drawing.SystemColors.Control
        Me.MenuTextSeperator6.ForeColor = System.Drawing.SystemColors.ControlDark
        Me.MenuTextSeperator6.Name = "MenuTextSeperator6"
        Me.MenuTextSeperator6.Size = New System.Drawing.Size(1920, 18)
        Me.MenuTextSeperator6.Text = "Misc:"
        '
        'ShowErrorsToolStripMenuItem
        '
        Me.ShowErrorsToolStripMenuItem.Checked = True
        Me.ShowErrorsToolStripMenuItem.CheckOnClick = True
        Me.ShowErrorsToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ShowErrorsToolStripMenuItem.Name = "ShowErrorsToolStripMenuItem"
        Me.ShowErrorsToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.ShowErrorsToolStripMenuItem.Text = "Show Mistakes"
        '
        'ShowIgnoredToolStripMenuItem
        '
        Me.ShowIgnoredToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.WhenCtrlIsPressedToolStripMenuItem, Me.AlwaysToolStripMenuItem, Me.NeverToolStripMenuItem})
        Me.ShowIgnoredToolStripMenuItem.Name = "ShowIgnoredToolStripMenuItem"
        Me.ShowIgnoredToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.ShowIgnoredToolStripMenuItem.Text = "Show Ignored"
        '
        'WhenCtrlIsPressedToolStripMenuItem
        '
        Me.WhenCtrlIsPressedToolStripMenuItem.Checked = True
        Me.WhenCtrlIsPressedToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.WhenCtrlIsPressedToolStripMenuItem.Name = "WhenCtrlIsPressedToolStripMenuItem"
        Me.WhenCtrlIsPressedToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.WhenCtrlIsPressedToolStripMenuItem.Tag = "2"
        Me.WhenCtrlIsPressedToolStripMenuItem.Text = "When Ctrl is Pressed"
        '
        'AlwaysToolStripMenuItem
        '
        Me.AlwaysToolStripMenuItem.Name = "AlwaysToolStripMenuItem"
        Me.AlwaysToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.AlwaysToolStripMenuItem.Tag = "0"
        Me.AlwaysToolStripMenuItem.Text = "Always"
        '
        'NeverToolStripMenuItem
        '
        Me.NeverToolStripMenuItem.Name = "NeverToolStripMenuItem"
        Me.NeverToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.NeverToolStripMenuItem.Tag = "1"
        Me.NeverToolStripMenuItem.Text = "Never"
        '
        'tsbSpellCheck
        '
        Me.tsbSpellCheck.AutoToolTip = False
        Me.tsbSpellCheck.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbSpellCheck.Name = "tsbSpellCheck"
        Me.tsbSpellCheck.Size = New System.Drawing.Size(72, 22)
        Me.tsbSpellCheck.Text = "Spell Check"
        '
        'tsbProperties
        '
        Me.tsbProperties.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbProperties.AutoToolTip = False
        Me.tsbProperties.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbProperties.Name = "tsbProperties"
        Me.tsbProperties.Size = New System.Drawing.Size(64, 22)
        Me.tsbProperties.Text = "Properties"
        '
        'tabSpellControls
        '
        Me.tabSpellControls.Controls.Add(Me.tabRichTextBox)
        Me.tabSpellControls.Controls.Add(Me.tabTextBox)
        Me.tabSpellControls.Controls.Add(Me.tabDataGridView)
        Me.tabSpellControls.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabSpellControls.Location = New System.Drawing.Point(0, 50)
        Me.tabSpellControls.Name = "tabSpellControls"
        Me.tabSpellControls.SelectedIndex = 0
        Me.tabSpellControls.Size = New System.Drawing.Size(624, 367)
        Me.tabSpellControls.TabIndex = 6
        '
        'tabRichTextBox
        '
        Me.tabRichTextBox.Controls.Add(Me.RichTextBox1)
        Me.tabRichTextBox.Controls.Add(Me.propRichTextBox)
        Me.tabRichTextBox.Location = New System.Drawing.Point(4, 22)
        Me.tabRichTextBox.Name = "tabRichTextBox"
        Me.tabRichTextBox.Padding = New System.Windows.Forms.Padding(3)
        Me.tabRichTextBox.Size = New System.Drawing.Size(616, 341)
        Me.tabRichTextBox.TabIndex = 1
        Me.tabRichTextBox.Text = "RichTextBox"
        Me.tabRichTextBox.UseVisualStyleBackColor = True
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!)
        Me.RichTextBox1.Location = New System.Drawing.Point(3, 3)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(352, 335)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = resources.GetString("RichTextBox1.Text")
        '
        'propRichTextBox
        '
        Me.propRichTextBox.Dock = System.Windows.Forms.DockStyle.Right
        Me.propRichTextBox.Location = New System.Drawing.Point(355, 3)
        Me.propRichTextBox.Name = "propRichTextBox"
        Me.propRichTextBox.Size = New System.Drawing.Size(258, 335)
        Me.propRichTextBox.TabIndex = 7
        Me.propRichTextBox.Visible = False
        '
        'tabTextBox
        '
        Me.tabTextBox.Controls.Add(Me.TextBox1)
        Me.tabTextBox.Controls.Add(Me.propTextBox)
        Me.tabTextBox.Location = New System.Drawing.Point(4, 22)
        Me.tabTextBox.Name = "tabTextBox"
        Me.tabTextBox.Padding = New System.Windows.Forms.Padding(3)
        Me.tabTextBox.Size = New System.Drawing.Size(616, 341)
        Me.tabTextBox.TabIndex = 0
        Me.tabTextBox.Text = "TextBox"
        Me.tabTextBox.UseVisualStyleBackColor = True
        '
        'propTextBox
        '
        Me.propTextBox.Dock = System.Windows.Forms.DockStyle.Right
        Me.propTextBox.Location = New System.Drawing.Point(355, 3)
        Me.propTextBox.Name = "propTextBox"
        Me.propTextBox.Size = New System.Drawing.Size(258, 335)
        Me.propTextBox.TabIndex = 8
        Me.propTextBox.Visible = False
        '
        'tabDataGridView
        '
        Me.tabDataGridView.Controls.Add(Me.DataGridView1)
        Me.tabDataGridView.Controls.Add(Me.BindingNavigator1)
        Me.tabDataGridView.Controls.Add(Me.propDataGridView)
        Me.tabDataGridView.Location = New System.Drawing.Point(4, 22)
        Me.tabDataGridView.Name = "tabDataGridView"
        Me.tabDataGridView.Padding = New System.Windows.Forms.Padding(3)
        Me.tabDataGridView.Size = New System.Drawing.Size(616, 341)
        Me.tabDataGridView.TabIndex = 2
        Me.tabDataGridView.Text = "DataGridView"
        Me.tabDataGridView.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 3)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(352, 310)
        Me.DataGridView1.TabIndex = 0
        '
        'BindingNavigator1
        '
        Me.BindingNavigator1.AddNewItem = Me.BindingNavigatorAddNewItem
        Me.BindingNavigator1.CountItem = Me.BindingNavigatorCountItem
        Me.BindingNavigator1.DeleteItem = Me.BindingNavigatorDeleteItem
        Me.BindingNavigator1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BindingNavigator1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.BindingNavigator1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2, Me.BindingNavigatorAddNewItem, Me.BindingNavigatorDeleteItem})
        Me.BindingNavigator1.Location = New System.Drawing.Point(3, 313)
        Me.BindingNavigator1.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.BindingNavigator1.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.BindingNavigator1.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.BindingNavigator1.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.BindingNavigator1.Name = "BindingNavigator1"
        Me.BindingNavigator1.PositionItem = Me.BindingNavigatorPositionItem
        Me.BindingNavigator1.Size = New System.Drawing.Size(352, 25)
        Me.BindingNavigator1.TabIndex = 1
        Me.BindingNavigator1.Text = "BindingNavigator1"
        '
        'BindingNavigatorAddNewItem
        '
        Me.BindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorAddNewItem.Image = CType(resources.GetObject("BindingNavigatorAddNewItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorAddNewItem.Name = "BindingNavigatorAddNewItem"
        Me.BindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorAddNewItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorAddNewItem.Text = "Add new"
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(35, 22)
        Me.BindingNavigatorCountItem.Text = "of {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Total number of items"
        '
        'BindingNavigatorDeleteItem
        '
        Me.BindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorDeleteItem.Image = CType(resources.GetObject("BindingNavigatorDeleteItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorDeleteItem.Name = "BindingNavigatorDeleteItem"
        Me.BindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorDeleteItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorDeleteItem.Text = "Delete"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveFirstItem.Text = "Move first"
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMovePreviousItem.Text = "Move previous"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(50, 23)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Current position"
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator1"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveNextItem.Text = "Move next"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveLastItem.Text = "Move last"
        '
        'BindingNavigatorSeparator2
        '
        Me.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator2"
        Me.BindingNavigatorSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'propDataGridView
        '
        Me.propDataGridView.Dock = System.Windows.Forms.DockStyle.Right
        Me.propDataGridView.Location = New System.Drawing.Point(355, 3)
        Me.propDataGridView.Name = "propDataGridView"
        Me.propDataGridView.Size = New System.Drawing.Size(258, 335)
        Me.propDataGridView.TabIndex = 8
        Me.propDataGridView.Visible = False
        '
        'ToolStrip3
        '
        Me.ToolStrip3.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbSpellCheck, Me.tsbProperties, Me.tsiEnabled})
        Me.ToolStrip3.Location = New System.Drawing.Point(0, 25)
        Me.ToolStrip3.Name = "ToolStrip3"
        Me.ToolStrip3.Size = New System.Drawing.Size(624, 25)
        Me.ToolStrip3.TabIndex = 7
        Me.ToolStrip3.Text = "ToolStrip3"
        '
        'tsiEnabled
        '
        Me.tsiEnabled.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsiEnabled.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsiEnabled.Name = "tsiEnabled"
        Me.tsiEnabled.Size = New System.Drawing.Size(53, 22)
        Me.tsiEnabled.Text = "Enabled"
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "Column1"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(624, 442)
        Me.Controls.Add(Me.tabSpellControls)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.ToolStrip3)
        Me.Controls.Add(Me.ToolStrip2)
        Me.MinimumSize = New System.Drawing.Size(400, 300)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Spell Check"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ToolStrip2.ResumeLayout(False)
        Me.ToolStrip2.PerformLayout()
        Me.tabSpellControls.ResumeLayout(False)
        Me.tabRichTextBox.ResumeLayout(False)
        Me.tabTextBox.ResumeLayout(False)
        Me.tabTextBox.PerformLayout()
        Me.tabDataGridView.ResumeLayout(False)
        Me.tabDataGridView.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BindingNavigator1.ResumeLayout(False)
        Me.BindingNavigator1.PerformLayout()
        Me.ToolStrip3.ResumeLayout(False)
        Me.ToolStrip3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tslStatus As System.Windows.Forms.ToolStripLabel
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents tsbVBForums As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbi00Productions As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStrip2 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripDropDownButton1 As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents MenuTextSeperator1 As i00SpellCheck.MenuTextSeperator
    Friend WithEvents tstbAnagramLookup As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents MenuTextSeperator2 As i00SpellCheck.MenuTextSeperator
    Friend WithEvents tstbScrabbleHelper As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents tsbAbout As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbCodeProject As System.Windows.Forms.ToolStripButton
    Friend WithEvents MenuTextSeperator3 As i00SpellCheck.MenuTextSeperator
    Friend WithEvents tsbSuggestionLookup As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents tabSpellControls As System.Windows.Forms.TabControl
    Friend WithEvents tabTextBox As System.Windows.Forms.TabPage
    Friend WithEvents tabRichTextBox As System.Windows.Forms.TabPage
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents CrosswordSolverToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripDropDownButton2 As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents MenuTextSeperator4 As i00SpellCheck.MenuTextSeperator
    Friend WithEvents SpellingErrorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsiCPSpellingError As i00SpellCheck.tsiColorPicker
    Friend WithEvents CaseErrorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsiCPCaseError As i00SpellCheck.tsiColorPicker
    Friend WithEvents IgnoredWordToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsiCPIgnoreColor As i00SpellCheck.tsiColorPicker
    Friend WithEvents MenuTextSeperator5 As i00SpellCheck.MenuTextSeperator
    Friend WithEvents tsiDrawStyle As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents MenuTextSeperator6 As i00SpellCheck.MenuTextSeperator
    Friend WithEvents ShowErrorsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ShowIgnoredToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AlwaysToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NeverToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents WhenCtrlIsPressedToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsbSpellCheck As System.Windows.Forms.ToolStripButton
    Friend WithEvents propRichTextBox As System.Windows.Forms.PropertyGrid
    Friend WithEvents propTextBox As System.Windows.Forms.PropertyGrid
    Friend WithEvents tsbProperties As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStrip3 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsiEnabled As System.Windows.Forms.ToolStripButton
    Friend WithEvents tabDataGridView As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BindingNavigator1 As System.Windows.Forms.BindingNavigator
    Friend WithEvents BindingNavigatorAddNewItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorCountItem As System.Windows.Forms.ToolStripLabel
    Friend WithEvents BindingNavigatorDeleteItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveFirstItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents propDataGridView As System.Windows.Forms.PropertyGrid
    Friend WithEvents tsbDonate As System.Windows.Forms.ToolStripButton

End Class
