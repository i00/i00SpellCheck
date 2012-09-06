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

'NativeWindow thinks it can be designed... so disable the designer 
<System.ComponentModel.DesignerCategory("")> _
Public Class SpellCheckTextBox
    Inherits SpellCheckControlBase
    Implements iSpellCheckDialog
    Implements iTestHarness

#Region "Text Box"

#Region "Scrollbar location"

    <Runtime.InteropServices.DllImport("user32.dll")> _
    Friend Shared Function SendMessage(ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    End Function

    <Runtime.InteropServices.DllImport("user32.dll")> _
    Friend Shared Function GetScrollPos(ByVal hwnd As IntPtr, ByVal nBar As Integer) As Integer
    End Function

    <Runtime.InteropServices.DllImport("user32.dll")> _
    Friend Shared Function LockWindowUpdate(ByVal hWndLock As IntPtr) As Boolean
    End Function

    Private Const SB_HORZ As Integer = 0
    Friend Const SB_VERT As Integer = 1
    Friend Const SB_TOP As Integer = 6
    Friend Const EM_SCROLL As Integer = &HB5
    Friend Const EM_LINESCROLL As Integer = &HB6
    Private Const EM_SCROLLCARET As Integer = &HB7

#End Region

#Region "To get the height of a given line / chr position"

    Dim RTBContents As Rectangle
    Private Sub parentRichTextBox_ContentsResized(ByVal sender As Object, ByVal e As System.Windows.Forms.ContentsResizedEventArgs)
        RTBContents = e.NewRectangle
    End Sub

    'call GetLineHeightFromCharPosition(TextBox.GetFirstCharIndexFromLine(LineIndex))
    'to do this by line
    Protected Function GetLineHeightFromCharPosition(ByVal LetterIndex As Integer) As Integer
        Dim TextHeight As Integer = System.Windows.Forms.TextRenderer.MeasureText("Ag", parentTextBox.Font).Height
        If parentRichTextBox IsNot Nothing Then
            'rich text box :(... need to get the height for each bit...
            Dim OldStart = parentRichTextBox.SelectionStart
            Dim OldLength = parentRichTextBox.SelectionLength
            Dim ThisLine = parentRichTextBox.GetLineFromCharIndex(LetterIndex)
            Dim LineBelow = parentRichTextBox.GetPositionFromCharIndex(parentRichTextBox.GetFirstCharIndexFromLine(ThisLine + 1))
            If LineBelow.IsEmpty Then
                If RTBContents.IsEmpty Then
                    'qwertyuiop ....
                    'TODO: hrm ... need some way to make the RTB fire the ContentsResized event without interfering with the user...
                    'use the standard text height 4 now...
                    Return TextHeight
                Else
                    Return (RTBContents.Height + parentRichTextBox.GetPositionFromCharIndex(0).Y) - parentRichTextBox.GetPositionFromCharIndex(LetterIndex).Y
                End If
            Else
                Return LineBelow.Y - parentRichTextBox.GetPositionFromCharIndex(LetterIndex).Y
            End If
        Else
            Return TextHeight
        End If
    End Function

#End Region

#Region "Change case"

    Public Shared Function SentenceCase(ByVal Input As String) As String
        Dim SentenceBreakers As String = System.Text.RegularExpressions.Regex.Escape(".!?" & vbCrLf)
        Dim Pattern As String = "((?<=^[^" & SentenceBreakers & "\w]{0,})[a-z]|(?<![" & SentenceBreakers & "])(?<=[" & SentenceBreakers & "][^" & SentenceBreakers & "\w]{0,})[a-z])"
        Return System.Text.RegularExpressions.Regex.Replace(Input, Pattern, Function(m) m.Value(0).ToString().ToUpper() & m.Value.Substring(1))
    End Function

    Private Enum Cases
        Origional
        Sentence
        Propper
        Upper
        Lower
    End Enum

    Dim TipCase As New HTMLToolTip With {.IsBalloon = True, .ToolTipIcon = ToolTipIcon.Info, .ToolTipTitle = "Case changed", .ToolTipOrientation = HTMLToolTip.ToolTipOrientations.TopLeft}

    Private Sub parentTextBox_ForChangeCase_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_Control.Disposed
        TipCase.Dispose()
    End Sub

    Private Sub CycleCase()
        If parentTextBox.SelectedText = "" Then Return

        Static LastSelectedText As String = ""
        Static CurrentCase As Cases
        Static OrigionalRTF As String
        If UCase(LastSelectedText) <> UCase(parentTextBox.SelectedText) Then
            'new text selected
            LastSelectedText = parentTextBox.SelectedText
            CurrentCase = Cases.Origional
            LastSelectedText = parentTextBox.SelectedText
            If parentRichTextBox IsNot Nothing Then OrigionalRTF = parentRichTextBox.SelectedRtf
        End If

        CurrentCase = CType((CInt(CurrentCase) + 1) Mod (UBound([Enum].GetValues(GetType(Cases))) + 1), Cases)

        If parentRichTextBox IsNot Nothing Then
            'this section is used to preserve RTB formatting when changing case...
            Dim ToBeText As String = OrigionalRTF
            Dim RTFMatchPatern = "(?<=\s).(?<![\s|\\]).*?((?=\s)|(?=((?<!\\)\\[^\\]))(?=((?<!\\)\\[^}]))(?=((?<!\\)\\[^{])))"
            'oldregex = "(?<=( |\t|\r|\n)).(?<!( |\t|\r|\n|\\)).*?((?=})|(?=\r)|(?=\n)|(?= )|(?=((?<!\\)\\[^\\]))(?=((?<!\\)\\[^}]))(?=((?<!\\)\\[^{])))"
            Dim EndSentenceChars() As Char = {"."c, "!"c, "?"c}
            Select Case CurrentCase
                Case Cases.Origional

                Case Cases.Sentence
                    Dim mc = System.Text.RegularExpressions.Regex.Matches(ToBeText, "(\r|\n|" & RTFMatchPatern & ")")

                    Dim NewSentence As Boolean = True
                    For Each match In mc.OfType(Of System.Text.RegularExpressions.Match)()
                        If isInRTFTextSegment(ToBeText, match) Then
                            'we can change this as we are true text...
                            If NewSentence Then
                                'capitalize the 1st letter...
                                Dim CapMatch = System.Text.RegularExpressions.Regex.Match(match.Value, "\w")
                                If CapMatch.Success Then
                                    'yay
                                    Dim NewWord = match.Value.Substring(0, CapMatch.Index) & CapMatch.Value.ToUpper & match.Value.Substring(CapMatch.Index + CapMatch.Length)
                                    ToBeText = ToBeText.Substring(0, match.Index) & NewWord & ToBeText.Substring(match.Index + match.Length)
                                    NewSentence = False
                                Else
                                    'this word needs cap... but no word start found...
                                    Continue For
                                End If
                            End If
                            If match.Value = vbCr OrElse match.Value = vbLf Then
                                NewSentence = True
                            Else
                                If match.Value <> "" AndAlso EndSentenceChars.Contains(match.Value.Last) Then
                                    NewSentence = True
                                End If
                            End If
                        End If
                    Next
                Case Cases.Propper
                    ToBeText = System.Text.RegularExpressions.Regex.Replace(ToBeText, RTFMatchPatern, Function(m) If(isInRTFTextSegment(ToBeText, m), StrConv(m.Value, VbStrConv.ProperCase), m.Value))
                Case Cases.Upper
                    ToBeText = System.Text.RegularExpressions.Regex.Replace(ToBeText, RTFMatchPatern, Function(m) If(isInRTFTextSegment(ToBeText, m), m.Value.ToUpper(), m.Value))
                Case Cases.Lower
                    ToBeText = System.Text.RegularExpressions.Regex.Replace(ToBeText, RTFMatchPatern, Function(m) If(isInRTFTextSegment(ToBeText, m), m.Value.ToLower(), m.Value))
            End Select
            Dim OldSelectionStart = parentTextBox.SelectionStart
            Dim OldSelectionLen = parentTextBox.SelectionLength
            Dim SetWholeText As Boolean  'gets round the empty line rtb bug...
            '                qwertyuiop - unfortunatly it creates another... the undo stack is cleared :(
            If OldSelectionStart = 0 AndAlso OldSelectionLen = parentTextBox.TextLength Then SetWholeText = True
            If SetWholeText Then
                parentRichTextBox.Rtf = ToBeText.TrimEnd(CChar(vbCr), CChar(vbLf))
            Else
                parentRichTextBox.SelectedRtf = ToBeText
            End If
            parentTextBox.SelectionStart = OldSelectionStart
            parentTextBox.SelectionLength = OldSelectionLen
        Else
            Dim ToBeText As String = LastSelectedText
            Select Case CurrentCase
                Case Cases.Origional

                Case Cases.Sentence
                    ToBeText = SentenceCase(ToBeText)
                Case Cases.Propper
                    ToBeText = StrConv(ToBeText, VbStrConv.ProperCase)
                Case Cases.Upper
                    ToBeText = UCase(ToBeText)
                Case Cases.Lower
                    ToBeText = LCase(ToBeText)
            End Select
            Dim OldSelectionStart = parentTextBox.SelectionStart
            Dim OldSelectionLen = parentTextBox.SelectionLength
            parentTextBox.SelectedText = ToBeText
            parentTextBox.SelectionStart = OldSelectionStart
            parentTextBox.SelectionLength = OldSelectionLen
        End If

        'tooltip

        'no longer to dispose 1st (to stop the dissapearing not being 5000 if they press it multiple times quickly) due to my new tooltip

        'scroll the caret into view
        SendMessage(Me.parentTextBox.Handle, EM_SCROLLCARET, 0, 0)
        Dim ChrPos = Me.parentTextBox.GetPositionFromCharIndex(parentTextBox.SelectionStart)
        Dim LineHeight As Integer = GetLineHeightFromCharPosition(parentTextBox.SelectionStart)
        ChrPos.Y += LineHeight
        ChrPos.X += 8

        If ChrPos.Y < 0 Then ChrPos.Y = 0
        If ChrPos.Y > parentTextBox.Height Then ChrPos.Y = parentTextBox.Height
        TipCase.ShowHTML("to " & CurrentCase.ToString & " " & If(CurrentCase = Cases.Origional, "state", "case"), parentTextBox, ChrPos, 2500)

    End Sub

    Private Function isInRTFTextSegment(ByVal CompleteString As String, ByVal m As System.Text.RegularExpressions.Match) As Boolean
        Return System.Text.RegularExpressions.Regex.Matches(Left(CompleteString, m.Index), "(?<!\\){").Count - System.Text.RegularExpressions.Regex.Matches(Left(CompleteString, m.Index), "(?<!\\)}").Count = 1
    End Function

#End Region

#Region "Key handlers"

    Private Sub parentTextBox_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mc_Control.KeyUp
        Select Case e.KeyCode
            Case Keys.Apps 'right click menu
                If ContextMenuStrip.Visible Then
                    'menu is open already ... so hide
                    ContextMenuStrip.Close()
                Else
                    MenuSpellClickReturn = New MenuSpellClickReturnItem(parentTextBox, False)
                End If
            Case Keys.F3
                If Settings.AllowF3 Then
                    'cycle through case's ... sentence, propper, upper, lower, origional
                    CycleCase()
                End If
        End Select

        If parentTextBox.IsSpellCheckEnabled = False Then Return

        Select Case e.KeyCode
            Case Keys.F7
                'spell check dialogue
                If Settings.AllowF7 Then
                    ShowDialog()
                    'update the ignored/added/removed words...
                    RepaintControl()
                End If
        End Select
    End Sub

#End Region

#Region "Custom Settings"

    Private mc_RenderCompatibility As Boolean = False
    <System.ComponentModel.DefaultValue(False)> _
    <System.ComponentModel.DisplayName("Compatible Rendering")> _
    <System.ComponentModel.Description("Compatible rendering increases drawing compatibility but also adds flickering upon redraw, only enable if there are redraw/layering problems")> _
    Public Property RenderCompatibility() As Boolean
        Get
            Return mc_RenderCompatibility
        End Get
        Set(ByVal value As Boolean)
            If mc_RenderCompatibility <> value Then
                mc_RenderCompatibility = value
                parentTextBox.Invalidate() 'OnRedraw()
            End If
        End Set
    End Property


#End Region

    Private Sub parentTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_Control.TextChanged
        'Standard TextBoxes do not fire the WM_PAINT event when adding text so need to do it this way
        If TypeOf parentTextBox Is TextBox Then
            parentTextBox.Invalidate()
        End If
    End Sub

    Private Sub parentTextBox_MultilineChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim senderTextBox = TryCast(sender, TextBoxBase)
        If senderTextBox IsNot Nothing Then
            If senderTextBox.Multiline Then
                senderTextBox.EnableSpellCheck()
            Else
                senderTextBox.DisableSpellCheck()
            End If
        End If
    End Sub


    Private Const WM_VSCROLL As Integer = &H115
    Private Const WM_MOUSEWHEEL As Integer = &H20A
    Private Const WM_PAINT As Integer = &HF

    Protected Overrides Sub WndProc(ByRef m As Message)
        Select Case m.Msg
            Case WM_PAINT
                If OKToDraw Then
                    If Me.RenderCompatibility Then
                        'old draw method...
                        parentTextBox.Invalidate()
                        CloseOverlay()
                        MyBase.WndProc(m)
                        Me.CustomPaint()
                    Else
                        MyBase.WndProc(m)
                        RepaintControl()
                        'OpenOverlay()
                        'Me.CustomPaint()
                    End If
                Else
                    MyBase.WndProc(m)
                End If
            Case Else
                MyBase.WndProc(m)
        End Select
    End Sub

#End Region

#Region "Constructor"

    Overrides Sub Load()
        If parentTextBox IsNot Nothing Then

            'setup control specific event handlers
            AddHandler parentTextBox.MultilineChanged, AddressOf parentTextBox_MultilineChanged
            If parentRichTextBox IsNot Nothing Then
                AddHandler parentRichTextBox.ContentsResized, AddressOf parentRichTextBox_ContentsResized
            End If

            'parentTextBox = TextBoxBase
            parentTextBox_ContextMenuStripChanged(parentTextBox, EventArgs.Empty)
            parentTextBox_SizeChanged(parentTextBox, EventArgs.Empty)

            If parentTextBox.Multiline = False Then parentTextBox.DisableSpellCheck()
            RepaintControl()
            'parentTextBox.Invalidate()
        End If
    End Sub

#End Region

#Region "Test Harness"

    Public Overridable Function SetupControl(ByVal Control As System.Windows.Forms.Control) As Control Implements iTestHarness.SetupControl
        If Control.GetType Is GetType(TextBox) Then
            Dim TextBox = DirectCast(Control, TextBox)

            TextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!)
            TextBox.Multiline = True
            TextBox.ScrollBars = ScrollBars.Vertical
            TextBox.Text = "Ths is a standrd text field that uses a dictionary to spel check the its contents ...  as you can se errors are underlnied in red!"

            TextBox.SelectionStart = 0
            TextBox.SelectionLength = 0

            Return TextBox
        ElseIf Control.GetType Is GetType(RichTextBox) Then
            Dim RichTextBox = DirectCast(Control, RichTextBox)

            RichTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!)

            RichTextBox.Text = "i00 .Net Spell Check has built in support for RichTextBoxes!" & vbCrLf & _
                               "" & vbCrLf & _
                               "The quic brown fox junped ovr the lazy dog!" & vbCrLf & _
                               "" & vbCrLf & _
                               "You can right click to see spelling sugguestions for words and to add/ignore/remove words from the dictionary." & vbCrLf & _
                               "" & vbCrLf & _
                               "If you ignre a word you can hold ctrl down to underlne all ignored words!" & vbCrLf & _
                               "" & vbCrLf & _
                               "The initial dictionary may take a little while to load ... it holds more than 150 000 words!"

            RichTextBox.Select(RichTextBox.Text.IndexOf("i00 .Net Spell Check"), Len("i00 .Net Spell Check"))
            RichTextBox.SelectionFont = New Font(RichTextBox.Font, FontStyle.Bold)
            RichTextBox.Select(RichTextBox.Text.IndexOf("RichTextBoxes!"), Len("RichTextBoxes!"))
            RichTextBox.SelectionFont = New Font(RichTextBox.Font.Name, CSng(RichTextBox.Font.Size * 1.5), FontStyle.Bold)
            RichTextBox.Select(RichTextBox.Text.IndexOf("Rich"), Len("Rich"))
            RichTextBox.SelectionColor = Color.Red
            RichTextBox.Select(RichTextBox.Text.IndexOf("Text"), Len("Text"))
            RichTextBox.SelectionColor = Color.Green
            RichTextBox.Select(RichTextBox.Text.IndexOf("Boxes"), Len("Boxes"))
            RichTextBox.SelectionColor = Color.Blue
            RichTextBox.Select(0, 0)
            RichTextBox.ClearUndo()

            Return RichTextBox
        Else
            Return Nothing
        End If
    End Function

#End Region

End Class