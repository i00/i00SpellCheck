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
    Implements iTestHarness

#Region "Setup"

    'Called when the control is loaded
    Overrides Sub Load()
        Dim FastColoredTextBox = TryCast(Me.Control, FastColoredTextBox)
        If FastColoredTextBox IsNot Nothing Then
            'Add control specific event handlers
            AddHandler FastColoredTextBox.TextChanged, AddressOf FastColoredTextBox_TextChanged
            AddHandler FastColoredTextBox.VisibleRangeChanged, AddressOf FastColoredTextBox_VisibleRangeChanged
            AddHandler FastColoredTextBox.VisualMarkerClick, AddressOf FastColoredTextBox_VisualMarkerClick
            AddHandler FastColoredTextBox.Disposed, AddressOf FastColoredTextBox_Disposed

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
        If parentFastColoredTextBox.ReadOnly Then Exit Sub
        Dim ErrorStyleMarker = TryCast(e.Marker, ErrorStyle.ErrorStyleMarker)
        If ErrorStyleMarker IsNot Nothing Then
            parentFastColoredTextBox.Selection = ErrorStyleMarker.Range

            SpellMenuItems.AddItems(ErrorStyleMarker.Range.Text, CurrentDictionary, CurrentDefinitions, CurrentSynonyms, Settings)
            ErrorMenu.Show(parentFastColoredTextBox, New Point(ErrorStyleMarker.rectangle.X, ErrorStyleMarker.rectangle.Bottom))
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
                Select Case WordState
                Case Dictionary.SpellCheckWordError.CaseError, Dictionary.SpellCheckWordError.SpellError
                    If eArgs.DrawDefault Then
                        DrawingFunctions.DrawWave(gr, New Point(rect.Left, rect.Bottom), New Point(rect.Right, rect.Bottom), Color)
                    End If
                    AddVisualMarker(FastColoredTextBox, New ErrorStyleMarker(range, rect, Me))
                Case Dictionary.SpellCheckWordError.Ignore
                    If eArgs.DrawDefault Then
                        Using p As New Pen(Color)
                            gr.DrawLine(p, New Point(rect.Left, rect.Bottom), New Point(rect.Right, rect.Bottom))
                        End Using
                    End If
                    AddVisualMarker(FastColoredTextBox, New ErrorStyleMarker(range, rect, Me))
            End Select
        End Sub

        Public Color As Color
    End Class

#End Region

#End Region

    Private WithEvents tmrRepaint As New Timer With {.Interval = 1, .Enabled = False}
    Public Overrides Sub RepaintControl()
        'qwertyuiop - will probably look at a way to pass back new errors eventually so we can just paint those as this would be able to SetStyle on those new errors ...
        'this has not been an issue before... as most controls repaint as a whole...
        UpdateErrors(parentFastColoredTextBox.VisibleRange)
    End Sub

    Private Sub tmrRepaint_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrRepaint.Tick
        UpdateErrors(parentFastColoredTextBox.VisibleRange)
        tmrRepaint.Enabled = False
    End Sub

    Private Sub FastColoredTextBox_Disposed(ByVal sender As Object, ByVal e As System.EventArgs)
        tmrRepaint.Dispose()
        SpellErrorStyle.Dispose()
        IgnoreErrorStyle.Dispose()
        CaseErrorStyle.Dispose()
    End Sub

    Private Sub FastColoredTextBox_VisibleRangeChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        tmrRepaint.Stop()
        tmrRepaint.Start()
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

#Region "Test Harness"

    Public Function SetupControl(ByVal Control As System.Windows.Forms.Control) As Control Implements i00SpellCheck.iTestHarness.SetupControl
        Dim FastColoredTextBox = TryCast(Control, FastColoredTextBox)
        If FastColoredTextBox IsNot Nothing Then
            FastColoredTextBox.LeftBracket = "("
            FastColoredTextBox.RightBracket = ")"

            DirectCast(FastColoredTextBox.SpellCheck(), SpellCheckFastColoredTextBox).SpellCheckMatch = "('.*$|"".*?"")"

            AddHandler FastColoredTextBox.TextChanged, AddressOf TestHarness_TextChanged

            FastColoredTextBox.Text = "'Simple test to check spelling with 3rd party controls" & vbCrLf & _
                   "'This test is done on the FastColoredTextBox (open source control) that is hosted on CodeProject" & vbCrLf & _
                   "'The article can be found at: http://www.codeproject.com/Articles/161871/Fast-Colored-TextBox-for-syntax-highlighting" & vbCrLf & _
                   "" & vbCrLf & _
                   "'i00 Does not take credit for any work in the FastColoredTextBox.dll" & vbCrLf & _
                   "'i00 is however solely responsible for the spellcheck plugin to interface with FastColoredTextBox" & vbCrLf & _
                   "" & vbCrLf & _
                   "'As you can see only comments and string blocks are corrected" & vbCrLf & _
                   "'This is due to the SpellCheckMatch property being set" & vbCrLf & _
                   "" & vbCrLf & _
                   "'Click on a missspelled word to correct..." & vbCrLf & _
                   "" & vbCrLf & _
                   "Dim test = ""Test with some bad spellling!""" & vbCrLf & _
                   "" & vbCrLf & _
                   "#Region ""Char""" & vbCrLf & _
                   "   " & vbCrLf & _
                   "   ''' <summary>" & vbCrLf & _
                   "   ''' Char and style" & vbCrLf & _
                   "   ''' </summary>" & vbCrLf & _
                   "   Public Structure CharStyle" & vbCrLf & _
                   "       Public c As Char" & vbCrLf & _
                   "       Public style As StyleIndex" & vbCrLf & _
                   "   " & vbCrLf & _
                   "       Public Sub CharStyle(ByVal ch As Char)" & vbCrLf & _
                   "           c = ch" & vbCrLf & _
                   "           Style = StyleIndex.None" & vbCrLf & _
                   "       End Sub" & vbCrLf & _
                   "   " & vbCrLf & _
                   "   End Structure" & vbCrLf & _
                   "   " & vbCrLf & _
                   "#End Region"
            Return FastColoredTextBox
        Else
            Return Nothing
        End If

    End Function

    Dim KeywordStyle As TextStyle = New TextStyle(Brushes.Blue, Nothing, FontStyle.Regular)
    Dim CommentStyle As TextStyle = New TextStyle(Brushes.Green, Nothing, FontStyle.Regular)
    Dim StringStyle As TextStyle = New TextStyle(Brushes.Brown, Nothing, FontStyle.Regular)

    Private Sub TestHarness_TextChanged(ByVal sender As System.Object, ByVal e As FastColoredTextBoxNS.TextChangedEventArgs)
        'clear style of changed range
        e.ChangedRange.ClearStyle(KeywordStyle, CommentStyle, StringStyle)

        e.ChangedRange.SetStyle(StringStyle, """.*?""")
        e.ChangedRange.SetStyle(CommentStyle, "'.*$|(\s|^)rem\s.*$", System.Text.RegularExpressions.RegexOptions.Multiline Or System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        'DataTypes
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Boolean|Byte|Char|Date|Decimal|Double|Integer|Long|Object|SByte|Short|Single|String|UInteger|ULong|UShort|Variant)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'Operators
        e.ChangedRange.SetStyle(KeywordStyle, "\b(AddressOf|And|AndAlso|Is|IsNot|Like|Mod|New|Not|Or|OrElse|Xor)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'Constants
        e.ChangedRange.SetStyle(KeywordStyle, "\b(False|Me|MyBase|MyClass|Nothing|True)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'CommonKeywords
        e.ChangedRange.SetStyle(KeywordStyle, "\b(As|Of|New|End)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'CommonKeywords
        e.ChangedRange.SetStyle(KeywordStyle, "\b(As|Of|New|End)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'FunctionKeywords
        e.ChangedRange.SetStyle(KeywordStyle, "\b(CBool|CByte|CChar|CDate|CDec|CDbl|CInt|CLng|CObj|CSByte|CShort|CSng|CStr|CType|CUInt|CULng|CUShort|DirectCast|GetType|TryCast|TypeOf)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'ParamModifiers
        e.ChangedRange.SetStyle(KeywordStyle, "\b(ByRef|ByVal|Optional|ParamArray)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'AccessModifiers
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Friend|Private|Protected|Public)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'OtherModifiers
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Const|Custom|Default|Global|MustInherit|MustOverride|Narrowing|NotInheritable|NotOverridable|Overloads|Overridable|Overrides|Partial|ReadOnly|Shadows|Shared|Static|Widening|WithEvents|WriteOnly)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'Statements
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Throw|Stop|Return|Resume|AddHandler|RemoveHandler|RaiseEvent|Option|Let|GoTo|GoSub|Call|Continue|Dim|ReDim|Erase|On|Error|Exit)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'GlobalConstructs
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Namespace|Class|Imports|Implements|Inherits|Interface|Delegate|Module|Structure|Enum)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'TypeLevelConstructs
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Sub|Function|Handles|Declare|Lib|Alias|Get|Set|Property|Operator|Event)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'Constructs
        e.ChangedRange.SetStyle(KeywordStyle, "\b(SyncLock|Using|With|Do|While|Loop|Wend|Try|Catch|When|Finally|If|Then|Else|For|To|Step|Each|In|Next|Select|Case)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'ContextKeywords
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Ansi|Auto|Unicode|Preserve|Until)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        'Region
        e.ChangedRange.SetStyle(KeywordStyle, "^\s{0,}#region\b|^\s{0,}#end\s{1,}region\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Multiline)

        'clear folding markers
        e.ChangedRange.ClearFoldingMarkers()
        'set folding markers
        e.ChangedRange.SetFoldingMarkers("^\s{0,}#region\b", "^\s{0,}#end\s{1,}region\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Singleline) 'allow to collapse #region blocks

    End Sub

#End Region

End Class
