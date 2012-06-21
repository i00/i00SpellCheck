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

Imports i00SpellCheck
Imports FastColoredTextBoxNS

Public Class SpellCheckFastColoredTextBox
    Inherits i00SpellCheck.SpellCheckControlBase

#Region "Setup"

    'Called when the control is loaded
    Overrides Sub Load()
        Dim FastColoredTextBox = TryCast(Me.Control, FastColoredTextBox)
        If FastColoredTextBox IsNot Nothing Then
            'Add control specific event handlers
            AddHandler FastColoredTextBox.TextChanged, AddressOf FastColoredTextBox_TextChanged
            AddHandler FastColoredTextBox.VisibleRangeChanged, AddressOf FastColoredTextBox_VisibleRangeChanged
            AddHandler FastColoredTextBox.VisualMarkerClick, AddressOf FastColoredTextBox_VisualMarkerClick

            SpellErrorStyle.FastColoredTextBox = FastColoredTextBox
            CaseErrorStyle.FastColoredTextBox = FastColoredTextBox
            IgnoreErrorStyle.FastColoredTextBox = FastColoredTextBox

            LoadSettingsToObjects()
        End If
    End Sub

    'Lets the EnableSpellCheck() know what ControlTypes we can spellcheck
    Public Overrides ReadOnly Property ControlType() As System.Type
        Get
            Return GetType(FastColoredTextBox)
        End Get
    End Property

    Private Sub LoadSettingsToObjects()
        SpellErrorStyle.Color = Settings.MistakeColor
        IgnoreErrorStyle.Color = Settings.IgnoreColor
        CaseErrorStyle.Color = Settings.CaseMistakeColor
    End Sub

    Private Sub SpellCheckFastColoredTextBox_DictionaryChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DictionaryChanged
        RepaintControl()
    End Sub

    'Repaint control when settings are changed
    Private Sub SpellCheckFastColoredTextBox_SettingsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SettingsChanged
        LoadSettingsToObjects()
        RepaintControl()
    End Sub

    Dim mc_SpellCheckMatch As String = ""

    <System.ComponentModel.Description("This regular expression is used to set what text gets spell checked in the FastColoredTextBox")> _
    <System.ComponentModel.DisplayName("Spell Check Match")> _
    Public Property SpellCheckMatch() As String
        Get
            Return mc_SpellCheckMatch
        End Get
        Set(ByVal value As String)
            If mc_SpellCheckMatch <> value Then
                mc_SpellCheckMatch = value
                RepaintControl()
            End If
        End Set
    End Property


#End Region

#Region "Underlying Control"

    'Make underlying control appear nicer in property grids
    <System.ComponentModel.Category("Control")> _
    <System.ComponentModel.Description("The FastColoredTextBox associated with the SpellCheckFastColoredTextBox object")> _
    <System.ComponentModel.DisplayName("FastColoredTextBox")> _
    Public Overrides ReadOnly Property Control() As System.Windows.Forms.Control
        Get
            Return MyBase.Control
        End Get
    End Property

    'Quick access to underlying FastColoredTextBox object
    Private ReadOnly Property parentFastColoredTextBox() As FastColoredTextBox
        Get
            Return TryCast(MyBase.Control, FastColoredTextBox)
        End Get
    End Property

#End Region

#Region "Click"

    Private WithEvents ErrorMenu As New ContextMenuStrip
    'Menu
    Private WithEvents SpellMenuItems As New Menu.AddSpellItemsToMenu() With {.ContextMenuStrip = ErrorMenu}

    'Error click
    Private Sub FastColoredTextBox_VisualMarkerClick(ByVal sender As Object, ByVal e As FastColoredTextBoxNS.VisualMarkerEventArgs)
        Dim ErrorStyleMarker = TryCast(e.Marker, ErrorStyle.ErrorStyleMarker)
        If ErrorStyleMarker IsNot Nothing Then
            parentFastColoredTextBox.Selection = ErrorStyleMarker.Range

            SpellMenuItems.AddItems(ErrorStyleMarker.Range.Text, CurrentDictionary, CurrentDefinitions, CurrentSynonyms, Settings)
            ErrorMenu.Show(parentFastColoredTextBox, New Point(ErrorStyleMarker.rectangle.X, ErrorStyleMarker.rectangle.Bottom))
            'MsgBox(ErrorStyleMarker.Range.Text)
        End If
    End Sub

    Private Sub ErrorMenu_Closed(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripDropDownClosedEventArgs) Handles ErrorMenu.Closed
        SpellMenuItems.RemoveSpellMenuItems()
    End Sub

    Private Sub SpellMenuItems_WordAdded(ByVal sender As Object, ByVal e As i00SpellCheck.Menu.AddSpellItemsToMenu.SpellItemEventArgs) Handles SpellMenuItems.WordAdded
        Try
            DictionaryAddWord(e.Word)
        Catch ex As Exception
            MsgBox("The following error occured adding """ & e.Word & """ to the dictionary:" & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub SpellMenuItems_WordChanged(ByVal sender As Object, ByVal e As i00SpellCheck.Menu.AddSpellItemsToMenu.SpellItemEventArgs) Handles SpellMenuItems.WordChanged
        Dim Start = parentFastColoredTextBox.Selection.Start
        parentFastColoredTextBox.InsertText(e.Word)
        parentFastColoredTextBox.Selection = New Range(parentFastColoredTextBox, Start, parentFastColoredTextBox.Selection.End)
    End Sub

    Private Sub SpellMenuItems_WordIgnored(ByVal sender As Object, ByVal e As i00SpellCheck.Menu.AddSpellItemsToMenu.SpellItemEventArgs) Handles SpellMenuItems.WordIgnored
        Try
            DictionaryIgnoreWord(e.Word)
        Catch ex As Exception
            MsgBox("The following error ignoring """ & e.Word & """:" & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub SpellMenuItems_WordRemoved(ByVal sender As Object, ByVal e As i00SpellCheck.Menu.AddSpellItemsToMenu.SpellItemEventArgs) Handles SpellMenuItems.WordRemoved
        Try
            DictionaryRemoveWord(e.Word)
        Catch ex As Exception
            MsgBox("The following error occured removing """ & e.Word & """ from the dictionary:" & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

#End Region

#Region "Painting"

#Region "Error Styles"

    Private WithEvents SpellErrorStyle As New ErrorStyle With {.WordState = Dictionary.SpellCheckWordError.SpellError}
    Private WithEvents IgnoreErrorStyle As New ErrorStyle With {.WordState = Dictionary.SpellCheckWordError.Ignore}
    Private WithEvents CaseErrorStyle As New ErrorStyle With {.WordState = Dictionary.SpellCheckWordError.CaseError}

    'Owner Draw Errors
    Private Sub ErrorStyle_SpellCheckErrorPaint(ByVal sender As Object, ByRef e As i00SpellCheck.SpellCheckControlBase.SpellCheckCustomPaintEventArgs) Handles SpellErrorStyle.SpellCheckErrorPaint, CaseErrorStyle.SpellCheckErrorPaint, IgnoreErrorStyle.SpellCheckErrorPaint
        OnSpellCheckErrorPaint(e)
    End Sub

#Region "Error Style"

    Private Class ErrorStyle
        Inherits Style

#Region "Marker"

        Public Class ErrorStyleMarker
            Inherits FastColoredTextBoxNS.StyleVisualMarker

            Public Overrides ReadOnly Property Cursor() As System.Windows.Forms.Cursor
                Get
                    Return Cursors.Default
                End Get
            End Property

            Public Range As Range

            Public Sub New(ByVal Range As Range, ByVal rectangle As Rectangle, ByVal style As FastColoredTextBoxNS.Style)
                MyBase.New(rectangle, style)
                Me.Range = Range
            End Sub
        End Class

#End Region

        Public Overrides Sub OnVisualMarkerClick(ByVal tb As FastColoredTextBoxNS.FastColoredTextBox, ByVal args As FastColoredTextBoxNS.VisualMarkerEventArgs)
            MyBase.OnVisualMarkerClick(tb, args)
        End Sub

        Public FastColoredTextBox As FastColoredTextBoxNS.FastColoredTextBox

        Public Event SpellCheckErrorPaint(ByVal sender As Object, ByRef e As SpellCheckCustomPaintEventArgs)
        Protected Sub OnSpellCheckErrorPaint(ByRef e As SpellCheckCustomPaintEventArgs)
            RaiseEvent SpellCheckErrorPaint(Me, e)
        End Sub

        Public WordState As Dictionary.SpellCheckWordError

        Public Overrides Sub Draw(ByVal gr As Graphics, ByVal position As Point, ByVal range As Range)
            Dim size As Size = Style.GetSizeOfRange(range)
            Dim rect As Rectangle = New Rectangle(position, size)

            Dim eArgs = New SpellCheckCustomPaintEventArgs With {.Graphics = gr, .Word = range.Text, .Bounds = rect, .WordState = WordState}
            OnSpellCheckErrorPaint(eArgs)
            If eArgs.DrawDefault Then
                Select Case WordState
                    Case Dictionary.SpellCheckWordError.CaseError, Dictionary.SpellCheckWordError.SpellError
                        DrawingFunctions.DrawWave(gr, New Point(rect.Left, rect.Bottom), New Point(rect.Right, rect.Bottom), Color)
                        AddVisualMarker(FastColoredTextBox, New ErrorStyleMarker(range, rect, Me))
                    Case Dictionary.SpellCheckWordError.Ignore
                        Using p As New Pen(Color)
                            gr.DrawLine(p, New Point(rect.Left, rect.Bottom), New Point(rect.Right, rect.Bottom))
                        End Using
                        AddVisualMarker(FastColoredTextBox, New ErrorStyleMarker(range, rect, Me))
                End Select
            End If
        End Sub

        Public Color As Color
    End Class

#End Region

#End Region

    Public Overrides Sub RepaintControl()
        'qwertyuiop - will probably look at a way to pass back new errors eventually so we can just paint those as this would be able to SetStyle on those new errors ...
        'this has not been an issue before... as most controls repaint as a whole...
        UpdateErrors(parentFastColoredTextBox.VisibleRange)
    End Sub

    Private Sub FastColoredTextBox_VisibleRangeChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        UpdateErrors(parentFastColoredTextBox.VisibleRange)
    End Sub

    Private Sub FastColoredTextBox_TextChanged(ByVal sender As System.Object, ByVal e As FastColoredTextBoxNS.TextChangedEventArgs)
        UpdateErrors(e.ChangedRange)
    End Sub

    Private Sub UpdateErrors(ByVal Range As FastColoredTextBoxNS.Range)
        Range.ClearStyle(IgnoreErrorStyle, CaseErrorStyle, SpellErrorStyle)
        If OKToDraw = False Then Exit Sub
        'clear errors...

        Dim text = Range.Text
        Dim RangesToSpellCheck As New List(Of Range)
        RangesToSpellCheck.Add(Range)
        If mc_SpellCheckMatch <> "" Then
            Dim matches = System.Text.RegularExpressions.Regex.Matches(text, mc_SpellCheckMatch, System.Text.RegularExpressions.RegexOptions.Multiline)
            text = Join((From xItem In matches.OfType(Of System.Text.RegularExpressions.Match)() Select xItem.Value).ToArray, " ")
            RangesToSpellCheck = Range.GetRanges(mc_SpellCheckMatch, System.Text.RegularExpressions.RegexOptions.Multiline).ToList
        End If

        text = Dictionary.Formatting.RemoveWordBreaks(text)
        Dim Words = Split(text, " ")

        Dim NewWords As New Dictionary(Of String, Dictionary.SpellCheckWordError)
        For Each word In Words
            If word <> "" Then
                Dim WordState As Dictionary.SpellCheckWordError = Dictionary.SpellCheckWordError.SpellError
                If dictCache.ContainsKey(word) Then
                    'load from cache
                    WordState = dictCache(word)
                Else
                    If NewWords.ContainsKey(word) = False Then
                        NewWords.Add(word, Dictionary.SpellCheckWordError.OK)
                    End If

                    WordState = Dictionary.SpellCheckWordError.OK
                End If
                Dim style As Style = Nothing
                Select Case WordState
                    Case Dictionary.SpellCheckWordError.Ignore
                        If DrawIgnored() Then
                            style = IgnoreErrorStyle
                        End If
                    Case Dictionary.SpellCheckWordError.CaseError
                        style = CaseErrorStyle
                    Case Dictionary.SpellCheckWordError.SpellError
                        style = SpellErrorStyle
                End Select
                If style IsNot Nothing Then
                    For Each RangeToSpellCheck In RangesToSpellCheck
                        RangeToSpellCheck.SetStyle(style, "\b" & System.Text.RegularExpressions.Regex.Escape(word) & "\b")
                    Next
                End If
            End If
        Next
        If NewWords.Count > 0 Then
            AddWordsToCache(NewWords)
        End If
    End Sub

#End Region

End Class
