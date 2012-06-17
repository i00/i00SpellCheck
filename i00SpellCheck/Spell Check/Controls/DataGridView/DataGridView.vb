'i00 .Net Spell Check
'©i00 Productions All rights reserved
'Created by Kris Bennett
'----------------------------------------------------------------------------------------------------
'All property in this file is and remains the property of i00 Productions, regardless of its usage,
'unless stated otherwise in writing from i00 Productions.
'
'Anyone wishing to use this code in their projects may do so, however are required to leave a post on
'VBForums (under: http://www.vbforums.com/showthread.php?p=4075093) stating that they are doing so.
'A simple "I am using i00 Spell check in my project" will surffice.
'
'i00 is not and shall not be held accountable for any damages directly or indirectly caused by the
'use or miss-use of this product.  This product is only a component and thus is intended to be used 
'as part of other software, it is not a complete software package, thus i00 Productions is not 
'responsible for any legal ramifications that software using this product breaches.

Partial Public Class SpellCheckDataGridView
    Inherits SpellCheckControlBase

#Region "Setup"

    'Called when the control is loaded
    Overrides Sub Load()
        Dim DataGridView = TryCast(Me.Control, DataGridView)
        If DataGridView IsNot Nothing Then
            'Add control specific event handlers
            AddHandler DataGridView.CellPainting, AddressOf DataGridView_CellPainting
            RepaintControl()
        End If
    End Sub

    'Lets the EnableSpellCheck() know what ControlTypes we can spellcheck
    Public Overrides ReadOnly Property ControlType() As System.Type
        Get
            Return GetType(DataGridView)
        End Get
    End Property

    'Repaint control when settings are changed
    Private Sub SpellCheckDataGridView_SettingsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SettingsChanged
        RepaintControl()
    End Sub

#End Region

#Region "Underlying Control"

    'Make underlying control appear nicer in property grids
    <System.ComponentModel.Category("Control")> _
    <System.ComponentModel.Description("The DataGridView associated with the SpellCheckDataGridView object")> _
    <System.ComponentModel.DisplayName("Data Grid View")> _
    Public Overrides ReadOnly Property Control() As System.Windows.Forms.Control
        Get
            Return MyBase.Control
        End Get
    End Property

    'Quick access to underlying DataGridView object
    Private ReadOnly Property parentDataGridView() As DataGridView
        Get
            Return TryCast(MyBase.Control, DataGridView)
        End Get
    End Property

#End Region

#Region "Painting"

    Public Overrides Sub RepaintControl()
        parentDataGridView.Invalidate()
    End Sub

    Private Sub DataGridView_CellPainting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs)
        If CurrentDictionary IsNot Nothing AndAlso (CurrentDictionary.Loading = True OrElse CurrentDictionary.Count = 0) Then Exit Sub

        Dim NewWords As New Dictionary(Of String, Dictionary.SpellCheckWordError)
        If parentDataGridView.IsSpellCheckEnabled Then
            Dim oSender = TryCast(sender, DataGridView)
            If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 AndAlso oSender IsNot Nothing Then

                If TypeOf oSender.Columns(e.ColumnIndex) Is DataGridViewTextBoxColumn AndAlso oSender.Columns(e.ColumnIndex).ReadOnly = False AndAlso e.Value IsNot Nothing Then

                    Dim SpellCheckSettings = oSender.AutoSpellCheckSettings
                    If SpellCheckSettings IsNot Nothing AndAlso (SpellCheckSettings.ShowMistakes) Then
                        Dim CellBounds = New Rectangle(e.CellBounds.X, e.CellBounds.Y, e.CellBounds.Width - 1, e.CellBounds.Height - 1)
                        e.Handled = True
                        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

                        If (e.PaintParts And DataGridViewPaintParts.Background) = DataGridViewPaintParts.Background Then
                            e.PaintBackground(e.ClipBounds, True)
                        End If
                        e.Graphics.SetClip(CellBounds)

                        If (e.PaintParts And DataGridViewPaintParts.ContentForeground) = DataGridViewPaintParts.ContentForeground Then

                            Dim Flags As TextFormatFlags
                            Flags = TextFormatFlags.NoPrefix

                            Dim FalseFlags = Flags
                            Flags = Flags Or TextFormatFlags.Top
                            Flags = Flags Or TextFormatFlags.Left

                            Select Case e.CellStyle.Alignment
                                Case DataGridViewContentAlignment.BottomCenter
                                    Flags = Flags Or TextFormatFlags.Bottom
                                    Flags = Flags Or TextFormatFlags.HorizontalCenter
                                Case DataGridViewContentAlignment.BottomLeft
                                    Flags = Flags Or TextFormatFlags.Bottom
                                    Flags = Flags Or TextFormatFlags.Left
                                Case DataGridViewContentAlignment.BottomRight
                                    Flags = Flags Or TextFormatFlags.Bottom
                                    Flags = Flags Or TextFormatFlags.Right
                                Case DataGridViewContentAlignment.MiddleCenter
                                    Flags = Flags Or TextFormatFlags.VerticalCenter
                                    Flags = Flags Or TextFormatFlags.HorizontalCenter
                                Case DataGridViewContentAlignment.MiddleLeft
                                    Flags = Flags Or TextFormatFlags.VerticalCenter
                                    Flags = Flags Or TextFormatFlags.Left
                                Case DataGridViewContentAlignment.MiddleRight
                                    Flags = Flags Or TextFormatFlags.VerticalCenter
                                    Flags = Flags Or TextFormatFlags.Right
                                Case DataGridViewContentAlignment.TopCenter
                                    Flags = Flags Or TextFormatFlags.Top
                                    Flags = Flags Or TextFormatFlags.HorizontalCenter
                                Case DataGridViewContentAlignment.TopLeft
                                    Flags = Flags Or TextFormatFlags.Top
                                    Flags = Flags Or TextFormatFlags.Left
                                Case DataGridViewContentAlignment.TopRight
                                    Flags = Flags Or TextFormatFlags.Top
                                    Flags = Flags Or TextFormatFlags.Right
                            End Select

                            If e.CellStyle.WrapMode = DataGridViewTriState.True Then
                                Flags = Flags Or TextFormatFlags.WordBreak
                            Else
                                Flags = Flags Or TextFormatFlags.WordEllipsis
                            End If

                            e.Graphics.TranslateTransform(CellBounds.Left, CellBounds.Top)

                            'spell check bit :D
                            'changed words makes words like "kris'" > "kris'"
                            Dim ChangedWordText = Replace(e.Value.ToString, vbCrLf, vbCr)
                            ChangedWordText = Replace(ChangedWordText, vbLf, vbCr)
                            Dim ChangedWords = Dictionary.Formatting.RemoveWordBreaks(ChangedWordText)

                            Dim WordBounds = DrawingFunctions.Text.TextRendererMeasure.Measure(e.Value.ToString, e.CellStyle.Font, CellBounds.Size, Flags)
                            Dim NewWordBounds As New DrawingFunctions.Text.TextRendererMeasure.WordBounds With {.TextMargin = WordBounds.TextMargin}

                            For Each item In WordBounds
                                'use the changed word instead :)
                                item.Word = ChangedWords.Substring(item.LetterIndex, item.Word.Length).Trim

                                'Dim Newwords = Split(item.Word, " ")

                                If item.Word.Contains(" ") Then
                                    'need to break this word into parts :(

                                    Dim NewSubWordBounds = DrawingFunctions.Text.TextRendererMeasure.Measure(item.Word, e.CellStyle.Font, New Size(Integer.MaxValue, Integer.MaxValue), FalseFlags)
                                    For Each iNewSubWordBounds In NewSubWordBounds
                                        iNewSubWordBounds.Bounds.X += item.Bounds.X - WordBounds.TextMargin
                                        iNewSubWordBounds.Bounds.Y += item.Bounds.Y
                                    Next
                                    NewWordBounds.AddRange(NewSubWordBounds)
                                Else
                                    'only one word here
                                    NewWordBounds.Add(item)
                                End If

                            Next

                            For Each item In NewWordBounds

                                Dim WordState As Dictionary.SpellCheckWordError = Dictionary.SpellCheckWordError.SpellError
                                If dictCache.ContainsKey(item.Word) Then
                                    'load from cache
                                    WordState = dictCache(item.Word)
                                Else
                                    If NewWords.ContainsKey(item.Word) = False Then
                                        NewWords.Add(item.Word, Dictionary.SpellCheckWordError.OK)
                                    End If

                                    WordState = Dictionary.SpellCheckWordError.OK
                                End If

                                Dim eArgs = New SpellCheckCustomPaintEventArgs With {.Graphics = e.Graphics, .Word = item.Word, .Bounds = item.Bounds, .WordState = WordState}
                                OnSpellCheckErrorPaint(eArgs)
                                If eArgs.DrawDefault Then
                                    Select Case WordState
                                        Case Dictionary.SpellCheckWordError.Ignore
                                            If DrawIgnored() Then
                                                Using p As New Pen(Settings.IgnoreColor)
                                                    e.Graphics.DrawLine(p, item.Bounds.X, item.Bounds.Bottom + 1, item.Bounds.Right, item.Bounds.Bottom + 1)
                                                End Using
                                            End If
                                        Case Dictionary.SpellCheckWordError.CaseError
                                            DrawingFunctions.DrawWave(e.Graphics, New Point(item.Bounds.X, item.Bounds.Bottom), New Point(item.Bounds.Right, item.Bounds.Bottom), Settings.CaseMistakeColor)
                                        Case Dictionary.SpellCheckWordError.SpellError
                                            DrawingFunctions.DrawWave(e.Graphics, New Point(item.Bounds.X, item.Bounds.Bottom), New Point(item.Bounds.Right, item.Bounds.Bottom), Settings.MistakeColor)
                                    End Select
                                End If
                            Next
                            'end spell check bit :D

                            e.Graphics.ResetTransform()
                            Dim CellForeColor = e.CellStyle.ForeColor
                            If (e.State And DataGridViewElementStates.Selected) = DataGridViewElementStates.Selected Then
                                CellForeColor = e.CellStyle.SelectionForeColor
                            End If
                            TextRenderer.DrawText(e.Graphics, e.Value.ToString, e.CellStyle.Font, CellBounds, CellForeColor, Flags)
                        End If

                        e.Graphics.ResetClip()
                    End If
                End If
            End If
        End If

        If NewWords.Count > 0 Then
            AddWordsToCache(NewWords)
        End If

    End Sub

#End Region

End Class