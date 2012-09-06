Imports i00SpellCheck

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''For the status bar to show the dictionary loading stats
        ''dictionary only loads 1st time you call EnableSpellCheck
        AddHandler SpellCheckFormExtension.DictionaryLoaded, AddressOf DictionaryLoaded
        tslStatus.Text = "Loading dictionary..."

        'SpellCheckControlAdded/Removed events are called whenever a spell check is created (auto or manual) on a control
        'Here this is used to process new textboxes for owner-draw events
        AddHandler SpellCheckControlExtension.SpellCheckControlAdded, AddressOf SpellCheckControlAdded
        AddHandler SpellCheckControlExtension.SpellCheckControlRemoved, AddressOf SpellCheckControlRemoved

        ''enable the spell check
        ''this will enable the spell check on ALL POSSIBLE CONTROLS ON THIS form AND ALL POSSIBLE CONTROLS ON ALL OWNED FORMS AS THEY OPEN automatically :)
        Me.EnableSpellCheck()
        ''To enable spell check on single line textboxes you will need to call:
        'TextBox1.EnableSpellCheck()

        ''if you wanted to pass in options you can do so by going:
        'Dim SpellCheckSettings As New i00SpellCheck.SpellCheckSettings
        'SpellCheckSettings.DoSubforms = True 'Specifies if owned forms should be automatically spell checked
        'SpellCheckSettings.AllowAdditions = True 'Specifies if you want to allow the user to add words to the dictionary
        'SpellCheckSettings.AllowIgnore = True 'Specifies if you want to allow the user ignore words
        'SpellCheckSettings.AllowRemovals = True 'Specifies if you want to allow users to delete words from the dictionary
        'SpellCheckSettings.AllowInMenuDefs = True 'Specifies if the in menu definitions should be shown for correctly spelled words
        'SpellCheckSettings.AllowChangeTo = True 'Specifies if "Change to..." (to change to a synonym) should be shown in the menu for correctly spelled words
        'Me.EnableSpellCheck(SpellCheckSettings)

        ''You can also enable spell checking on an individual Control (if supported):
        'TextBox1.EnableSpellCheck()

        ''To disable the spell check on a Control:
        'TextBox1.DisableSpellCheck()

        ''To see if the spell check is enabled on a Control:
        'Dim SpellCheckEnabled = TextBox1.IsSpellCheckEnabled()

        ''To change options on an individual Control:
        'TextBox1.SpellCheck.Settings.AllowAdditions = True
        'TextBox1.SpellCheck.Settings.AllowIgnore = True
        'TextBox1.SpellCheck.Settings.AllowRemovals = True
        'TextBox1.SpellCheck.Settings.ShowMistakes = True
        ''etc

        ''To show a spellcheck dialog for an individual Control:
        'Dim iSpellCheckDialog = TryCast(TextBox1.SpellCheck, i00SpellCheck.SpellCheckControlBase.iSpellCheckDialog)
        'If iSpellCheckDialog IsNot Nothing Then
        '    iSpellCheckDialog.ShowDialog()
        'End If

        ''To load a custom dictionary from a saved file:
        'Dim Dictionary = New i00SpellCheck.FlatFileDictionary("c:\Custom.dic")

        ''To create a new blank dictionary and save it as a file
        'Dim Dictionary = New i00SpellCheck.FlatFileDictionary("c:\Custom.dic", True)
        'Dictionary.Add("CustomWord1")
        'Dictionary.Add("CustomWord2")
        'Dictionary.Add("CustomWord3")
        'Dictionary.Save()

        ''To Load a custom dictionary for an individual Control:
        'TextBox1.SpellCheck.CurrentDictionary = Dictionary

        ''To Open the dictionary editor for a dictionary associated with a Control:
        ''NOTE: this should only be done after the dictionary has loaded (Control.SpellCheck.CurrentDictionary.Loading = False)
        'TextBox1.SpellCheck.CurrentDictionary.ShowUIEditor()

        ''Repaint all of the controls that use the same dictionary...
        'TextBox1.SpellCheck.InvalidateAllControlsWithSameDict()
    End Sub

    Private Sub DictionaryLoaded(ByVal sender As Object, ByVal e As EventArgs)
        tslStatus.Text = "Dictionary loaded " & Dictionary.DefaultDictionary.Count & " word" & If(Dictionary.DefaultDictionary.Count = 1, "", "s")
    End Sub

    Private Sub SpellCheckControlAdded(ByVal sender As Object, ByVal e As SpellCheckControlAddedRemovedEventArgs)
        'add the custom draw event handlers...
        AddHandler SpellCheckControls(e.Control).SpellCheckErrorPaint, AddressOf SpellCheckErrorPaint
    End Sub

    Private Sub SpellCheckControlRemoved(ByVal sender As Object, ByVal e As SpellCheckControlAddedRemovedEventArgs)
        'remove the custom draw event handlers...
        If SpellCheckControls.ContainsKey(e.Control) Then
            RemoveHandler SpellCheckControls(e.Control).SpellCheckErrorPaint, AddressOf SpellCheckErrorPaint
        End If
    End Sub

    Private Sub SpellCheckErrorPaint(ByVal sender As Object, ByRef e As SpellCheckControlBase.SpellCheckCustomPaintEventArgs)
        Dim SpellCheckControlBase = TryCast(sender, SpellCheckControlBase)
        If SpellCheckControlBase IsNot Nothing Then
            Dim Color As Color
            Select Case e.WordState
                Case Dictionary.SpellCheckWordError.CaseError
                    Color = SpellCheckControlBase.Settings.CaseMistakeColor
                Case Dictionary.SpellCheckWordError.Ignore
                    Color = SpellCheckControlBase.Settings.IgnoreColor
                Case Dictionary.SpellCheckWordError.SpellError
                    Color = SpellCheckControlBase.Settings.MistakeColor
                Case Else
                    'no mistake here...
                    Exit Sub
            End Select
            Select Case tsiDrawStyle.Text
                Case "Boxed In"
                    Using sb As New SolidBrush(Color.FromArgb(63, Color))
                        e.Graphics.FillRectangle(sb, e.Bounds)
                    End Using
                    Using p As New Pen(Color)
                        e.Graphics.DrawRectangle(p, e.Bounds)
                    End Using
                    e.DrawDefault = False
                Case "Opera"
                    Using p As New Pen(Color)
                        p.DashPattern = New Single() {1, 2}
                        p.DashStyle = Drawing2D.DashStyle.Custom
                        e.Graphics.DrawLine(p, e.Bounds.X, e.Bounds.Bottom, e.Bounds.Right, e.Bounds.Bottom)
                    End Using
                    e.DrawDefault = False
                Case "Draft Plan"
                    If e.WordState = Dictionary.SpellCheckWordError.Ignore Then
                        e.DrawDefault = True
                    Else
                        Randomize()
                        'e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                        Using gp As New Drawing2D.GraphicsPath
                            Dim Points As New List(Of PointF)
                            Dim StartX = e.Bounds.X + ((Rnd() * 16) - 8)
                            For i = StartX To e.Bounds.Right + 8 Step (Rnd() * 16) + 16
                                Dim y = ((Rnd() * 8) + ((e.Bounds.Height / 2) - 4) + e.Bounds.Top)
                                Points.Add(New PointF(i, CSng(y)))
                            Next
                            If Points.Count = 1 Then
                                'add another point
                                Dim y = ((Rnd() * 8) + ((e.Bounds.Height / 2) - 4) + e.Bounds.Top)
                                Points.Add(New PointF(Points.First.X, CSng(y)))
                            End If
                            If Points.Count > 1 Then
                                gp.AddCurve(Points.ToArray)
                                Using p As New Pen(Color, 3)
                                    Try
                                        e.Graphics.DrawPath(p, gp)
                                    Catch ex As Exception
                                        ' path produces an out of memory exception if all of the points are in the same spot ...
                                        'this happens sometimes since generating alot of random numbers in quick succession duplicates sometimes!
                                    End Try
                                End Using
                            End If
                            e.DrawDefault = False
                        End Using
                    End If
                Case Else
                    'draw default
            End Select
        End If
    End Sub

    Private Sub tstbAnagramLookup_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tstbAnagramLookup.KeyPress
        If e.KeyChar = vbCr Then
            Dim iAnagram = TryCast(i00SpellCheck.Dictionary.DefaultDictionary, i00SpellCheck.Dictionary.Interfaces.iAnagram)
            If iAnagram IsNot Nothing Then
                Dim Anagrams = iAnagram.AnagramLookup(tstbAnagramLookup.Text).ToArray
                MsgBox("Found " & Anagrams.Count & " anogram" & If(Anagrams.Count = 1, "", "s") & If(Anagrams.Count = 0, "", ":" & vbCrLf & Join(Anagrams, vbCrLf)))
            End If
        End If
    End Sub

    Private Sub tstbScrabbleHelper_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tstbScrabbleHelper.KeyPress
        If e.KeyChar = vbCr Then
            Dim iScrabble = TryCast(i00SpellCheck.Dictionary.DefaultDictionary, i00SpellCheck.Dictionary.Interfaces.iScrabble)
            Dim Scrabble = iScrabble.ScrabbleLookup(tstbScrabbleHelper.Text).ToArray
            MsgBox("Found " & Scrabble.Count & " scrabble combination" & If(Scrabble.Count = 1, "", "s") & If(Scrabble.Count = 0, "", ":" & vbCrLf & Join((From xItem In Scrabble Order By xItem.Score Descending, xItem.Word Ascending Select xItem.Word & " (" & xItem.Score & ")").ToArray, ",")), , "Scrabble combinations for " & tstbScrabbleHelper.Text)
        End If
    End Sub

    Private Sub tsbSuggestionLookup_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tsbSuggestionLookup.KeyPress
        If e.KeyChar = vbCr Then
            tsbSuggestionLookup.Text = tsbSuggestionLookup.Text.Replace(" ", "") 'remove spaces
            If i00SpellCheck.Dictionary.DefaultDictionary.SpellCheckWord(tsbSuggestionLookup.Text) = Dictionary.SpellCheckWordError.OK Then
                'Correct spelling
                'look up the meaning of the word :)
                Dim WordDefs = Join((From xItem In Definitions.DefaultDefinitions.FindWord(tsbSuggestionLookup.Text, i00SpellCheck.Dictionary.DefaultDictionary) Select Replace(Replace(xItem.Line, ";", "; "), "|", " - ")).ToArray, vbCrLf & vbCrLf)
                If WordDefs <> "" Then
                    MsgBox(tsbSuggestionLookup.Text & vbCrLf & vbCrLf & WordDefs)
                    Return
                End If
                MsgBox("Your word is spelt correctly!")
            Else
                'Incorrect - Offer suggestions
                Dim Suggestions = i00SpellCheck.Dictionary.DefaultDictionary.SpellCheckSuggestions(tsbSuggestionLookup.Text)
                If Suggestions.Count = 0 Then
                    MsgBox("Your word was not in the dictionary and no sugguestions could be made")
                    Return
                End If
                Dim TopCloseness = Suggestions.Max(Function(x As i00SpellCheck.Dictionary.SpellCheckSuggestionInfo) x.Closness)
                Dim FilteredSuggestions = (From xItem In Suggestions Order By xItem.Closness Descending, xItem.Word Ascending Where xItem.Closness >= TopCloseness * 0.75 Select xItem.Word).ToArray
                FilteredSuggestions = (From xItem In FilteredSuggestions Where Array.IndexOf(FilteredSuggestions, xItem) < 15).ToArray

                MsgBox("Found " & FilteredSuggestions.Count & " suggestion" & If(FilteredSuggestions.Count = 1, "", "s") & If(FilteredSuggestions.Count = 0, "", ":" & vbCrLf & Join(FilteredSuggestions, vbCrLf)) & vbCrLf & vbCrLf & _
                       "This sugguestion lookup uses closeness tolerance of 25% to filter results and limited to 15 results.")
            End If
        End If
    End Sub

    Private Sub CrosswordSolverToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrosswordSolverToolStripMenuItem.Click
        Using frmCrosswordGenerator As New i00SpellCheck.frmCrosswordGenerator()
            frmCrosswordGenerator.StartPosition = FormStartPosition.CenterParent
            frmCrosswordGenerator.ShowInTaskbar = False
            frmCrosswordGenerator.ShowDialog(Me)
        End Using
    End Sub

    Private Sub ShowErrorsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowErrorsToolStripMenuItem.Click
        ShowIgnoredToolStripMenuItem.Enabled = ShowErrorsToolStripMenuItem.Checked
        For Each item In SpellCheckControls
            item.Value.Settings.ShowMistakes = ShowErrorsToolStripMenuItem.Checked
        Next
        UpdatePropertyGrids()
    End Sub

    Private Sub ShowIgnoredToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim SelectedEnum = i00SpellCheck.SpellCheckSettings.ShowIgnoreState.OnKeyDown
        For Each item In ShowIgnoredToolStripMenuItem.DropDownItems.OfType(Of ToolStripMenuItem)()
            If item Is sender Then
                item.Checked = True
                SelectedEnum = CType(Val(item.Tag), i00SpellCheck.SpellCheckSettings.ShowIgnoreState)
            Else
                item.Checked = False
            End If
        Next
        For Each item In SpellCheckControls
            item.Value.Settings.ShowIgnored = SelectedEnum
        Next
        UpdatePropertyGrids()
    End Sub

    Private Sub tsiCPSpellingError_ColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsiCPSpellingError.ColorChanged
        For Each item In SpellCheckControls
            item.Value.Settings.MistakeColor = tsiCPSpellingError.SelectedColor
        Next
        UpdatePropertyGrids()
    End Sub

    Private Sub tsiCPCaseError_ColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsiCPCaseError.ColorChanged
        For Each item In SpellCheckControls
            item.Value.Settings.CaseMistakeColor = tsiCPCaseError.SelectedColor
        Next
        UpdatePropertyGrids()
    End Sub

    Private Sub tsiCPIgnoreColor_ColorChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsiCPIgnoreColor.ColorChanged
        For Each item In SpellCheckControls
            item.Value.Settings.IgnoreColor = tsiCPIgnoreColor.SelectedColor
        Next
        UpdatePropertyGrids()
    End Sub

    Private Sub tsbSpellCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbSpellCheck.Click
        Dim control As Control = tabSpellControls.SelectedTab
        Do Until control Is Nothing
            Dim iSpellCheckDialog = TryCast(control.SpellCheck, i00SpellCheck.SpellCheckControlBase.iSpellCheckDialog)
            If iSpellCheckDialog IsNot Nothing AndAlso control.SpellCheck.CurrentDictionary.Loading = False Then
                iSpellCheckDialog.ShowDialog()
                Exit Sub
            End If
            control = tabSpellControls.SelectedTab.GetNextControl(control, True)
        Loop
    End Sub

    Private Sub tsbProperties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbProperties.Click
        tsbProperties.Checked = Not tsbProperties.Checked
        UpdatePropertyGrids()
    End Sub

    Private Sub UpdatePropertyGrids()
        Dim ctl As Control = Me
        Do Until ctl Is Nothing
            Dim PropertyGrid = TryCast(ctl, PropertyGrid)
            If PropertyGrid IsNot Nothing Then
                PropertyGrid.Visible = tsbProperties.Checked
                PropertyGrid.Refresh()
            End If
            ctl = Me.GetNextControl(ctl, True)
        Loop
    End Sub

    Private Sub tsiDrawStyle_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsiDrawStyle.SelectedIndexChanged
        For Each item In SpellCheckControls
            item.Value.RepaintControl()
        Next
    End Sub

    Private Sub tabSpellControls_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabSpellControls.SelectedIndexChanged
        UpdateEnabledCheck()
        Dim selectedTabSpellCheckControl = (From xItem In tabSpellControls.SelectedTab.Controls.OfType(Of Control)() Where xItem.SpellCheck IsNot Nothing).FirstOrDefault
        If selectedTabSpellCheckControl IsNot Nothing Then
            tsbSpellCheck.Visible = TypeOf selectedTabSpellCheckControl.SpellCheck Is i00SpellCheck.SpellCheckControlBase.iSpellCheckDialog
        Else
            tsbSpellCheck.Visible = False
        End If
    End Sub

    Private Sub UpdateEnabledCheck()
        Dim selectedTabSpellCheckControl = (From xItem In tabSpellControls.SelectedTab.Controls.OfType(Of Control)() Where xItem.SpellCheck IsNot Nothing).FirstOrDefault
        If selectedTabSpellCheckControl Is Nothing Then
            tsiEnabled.Visible = False
        Else
            Dim ts = DirectCast(tsiEnabled.Owner, ToolStrip)
            Dim state As VisualStyles.CheckBoxState = If(selectedTabSpellCheckControl.IsSpellCheckEnabled, VisualStyles.CheckBoxState.CheckedNormal, VisualStyles.CheckBoxState.UncheckedNormal)
            Dim Size = System.Windows.Forms.CheckBoxRenderer.GetGlyphSize(Me.CreateGraphics, state)

            Dim bWidth As Integer, bHeight As Integer

            bWidth = ts.ImageScalingSize.Width
            bHeight = ts.ImageScalingSize.Height

            Dim Offset As New Point(0, 0)

            If Size.Width < ts.ImageScalingSize.Width Then
                Offset.X = CInt(Int((ts.ImageScalingSize.Width - Size.Width) / 2))
            Else
                bWidth = Size.Width
            End If
            If Size.Height < ts.ImageScalingSize.Height Then
                Offset.Y = CInt(Int((ts.ImageScalingSize.Height - Size.Height) / 2))
            Else
                bHeight = Size.Height
            End If


            Dim b As New Bitmap(bWidth, bHeight)
            Using g = Graphics.FromImage(b)
                g.TranslateTransform(Offset.X, Offset.Y)
                System.Windows.Forms.CheckBoxRenderer.DrawCheckBox(g, New Point(0, 0), state)
            End Using
            tsiEnabled.Image = b
            tsiEnabled.Visible = True
        End If
    End Sub

    Private Sub tsiEnabled_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsiEnabled.Click
        Dim selectedTabSpellCheckControl = (From xItem In tabSpellControls.SelectedTab.Controls.OfType(Of Control)() Where xItem.SpellCheck IsNot Nothing).FirstOrDefault
        If selectedTabSpellCheckControl IsNot Nothing Then
            If selectedTabSpellCheckControl.IsSpellCheckEnabled Then
                selectedTabSpellCheckControl.DisableSpellCheck()
            Else
                selectedTabSpellCheckControl.EnableSpellCheck()
            End If
        End If
        UpdateEnabledCheck()
    End Sub


    Private Sub tsbExamples_DropDownOpening(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbExamples.DropDownOpening
        Dim iAnagram = TryCast(i00SpellCheck.Dictionary.DefaultDictionary, i00SpellCheck.Dictionary.Interfaces.iAnagram)
        If iAnagram Is Nothing Then
            tstbAnagramLookup.Text = "Not Supported"
            tstbAnagramLookup.Enabled = False
        End If

        Dim iScrabble = TryCast(i00SpellCheck.Dictionary.DefaultDictionary, i00SpellCheck.Dictionary.Interfaces.iScrabble)
        If iScrabble Is Nothing Then
            tstbScrabbleHelper.Text = "Not Supported"
            tstbScrabbleHelper.Enabled = False
        End If
    End Sub

    Private Sub tsbEditDictionary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbEditDictionary.Click
        Dim ctl As Control = tabSpellControls.SelectedTab
        Do Until ctl Is Nothing
            If ctl.SpellCheck IsNot Nothing AndAlso ctl.SpellCheck.CurrentDictionary.Loading = False Then

                Dim controls As New List(Of Control)
                Dim AllControlsWithSameDict = (From xItem In ctl.SpellCheck.AllControlsWithSameDict() Select xItem.Control).ToArray
                If AllControlsWithSameDict.Count > 1 Then
                    Using MessageBoxManager As New MessageBoxManager
                        MessageBoxManager.Yes = "Just this"
                        MessageBoxManager.No = "All controls"
                        Select Case MsgBox("Do you want to modify the dictionary for just this control or all " & AllControlsWithSameDict.Count & " controls that share this dictionary?", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question)
                            Case MsgBoxResult.Yes 'Just this
                                controls.Add(ctl)
                            Case MsgBoxResult.No 'All controls
                                controls.AddRange(AllControlsWithSameDict)
                            Case MsgBoxResult.Cancel
                                Return
                        End Select
                    End Using
                Else
                    'this is the only control using this dict... just do this one
                    controls.Add(ctl)
                End If
                Dim NewDict = ctl.SpellCheck.CurrentDictionary.ShowUIEditor()
                For Each item In controls
                    item.SpellCheck.CurrentDictionary = NewDict
                Next

                Exit Sub
            End If

                ctl = tabSpellControls.SelectedTab.GetNextControl(ctl, True)
        Loop
    End Sub
End Class
