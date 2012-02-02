﻿'i00 .Net Spell Check
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
    Inherits NativeWindow

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
    Private Sub mc_parentRichTextBox_ContentsResized(ByVal sender As Object, ByVal e As System.Windows.Forms.ContentsResizedEventArgs) Handles mc_parentRichTextBox.ContentsResized
        RTBContents = e.NewRectangle
    End Sub

    'call GetLineHeightFromCharPosition(TextBox.GetFirstCharIndexFromLine(LineIndex))
    'to do this by line
    Private Function GetLineHeightFromCharPosition(ByVal LetterIndex As Integer) As Integer
        Dim TextHeight As Integer = System.Windows.Forms.TextRenderer.MeasureText("Ag", parentTextBox.Font).Height
        If mc_parentRichTextBox IsNot Nothing Then
            'rich text box :(... need to get the height for each bit...
            Dim OldStart = mc_parentRichTextBox.SelectionStart
            Dim OldLength = mc_parentRichTextBox.SelectionLength
            Dim ThisLine = mc_parentRichTextBox.GetLineFromCharIndex(LetterIndex)
            Dim LineBelow = mc_parentRichTextBox.GetPositionFromCharIndex(mc_parentRichTextBox.GetFirstCharIndexFromLine(ThisLine + 1))
            If LineBelow.IsEmpty Then
                If RTBContents.IsEmpty Then
                    'qwertyuiop ....
                    'TODO: hrm ... need some way to make the RTB fire the ContentsResized event without interfering with the user...
                    'use the standard text height 4 now...
                    Return TextHeight
                Else
                    Return (RTBContents.Height + mc_parentRichTextBox.GetPositionFromCharIndex(0).Y) - mc_parentRichTextBox.GetPositionFromCharIndex(LetterIndex).Y
                End If
            Else
                Return LineBelow.Y - mc_parentRichTextBox.GetPositionFromCharIndex(LetterIndex).Y
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

    Private Sub parentTextBox_ForChangeCase_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_parentTextBox.Disposed
        TipCase.Dispose()
    End Sub

    Private Sub CycleCase()
        If mc_parentTextBox.SelectedText = "" Then Return

        Static LastSelectedText As String = ""
        Static CurrentCase As Cases
        Static OrigionalRTF As String
        If UCase(LastSelectedText) <> UCase(mc_parentTextBox.SelectedText) Then
            'new text selected
            LastSelectedText = mc_parentTextBox.SelectedText
            CurrentCase = Cases.Origional
            LastSelectedText = mc_parentTextBox.SelectedText
            If mc_parentRichTextBox IsNot Nothing Then OrigionalRTF = mc_parentRichTextBox.SelectedRtf
        End If

        CurrentCase = CType((CInt(CurrentCase) + 1) Mod (UBound([Enum].GetValues(GetType(Cases))) + 1), Cases)

        If mc_parentRichTextBox IsNot Nothing Then
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
            Dim OldSelectionStart = mc_parentTextBox.SelectionStart
            Dim OldSelectionLen = mc_parentTextBox.SelectionLength
            Dim SetWholeText As Boolean  'gets round the empty line rtb bug...
            '                qwertyuiop - unfortunatly it creates another... the undo stack is cleared :(
            If OldSelectionStart = 0 AndAlso OldSelectionLen = mc_parentTextBox.TextLength Then SetWholeText = True
            If SetWholeText Then
                mc_parentRichTextBox.Rtf = ToBeText.TrimEnd(CChar(vbCr), CChar(vbLf))
            Else
                mc_parentRichTextBox.SelectedRtf = ToBeText
            End If
            mc_parentTextBox.SelectionStart = OldSelectionStart
            mc_parentTextBox.SelectionLength = OldSelectionLen
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
            Dim OldSelectionStart = mc_parentTextBox.SelectionStart
            Dim OldSelectionLen = mc_parentTextBox.SelectionLength
            mc_parentTextBox.SelectedText = ToBeText
            mc_parentTextBox.SelectionStart = OldSelectionStart
            mc_parentTextBox.SelectionLength = OldSelectionLen
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
        If ChrPos.Y > mc_parentTextBox.Height Then ChrPos.Y = mc_parentTextBox.Height
        TipCase.ShowHTML("to " & CurrentCase.ToString & " " & If(CurrentCase = Cases.Origional, "state", "case"), mc_parentTextBox, ChrPos, 2500)

    End Sub

    Private Function isInRTFTextSegment(ByVal CompleteString As String, ByVal m As System.Text.RegularExpressions.Match) As Boolean
        Return System.Text.RegularExpressions.Regex.Matches(Left(CompleteString, m.Index), "(?<!\\){").Count - System.Text.RegularExpressions.Regex.Matches(Left(CompleteString, m.Index), "(?<!\\)}").Count = 1
    End Function

#End Region

    <System.ComponentModel.Description("The text box associated with the SpellCheckTextBox object")> _
    Public ReadOnly Property TextBox() As TextBoxBase
        Get
            Return mc_parentTextBox
        End Get
    End Property

    Private WithEvents mc_parentTextBox As TextBoxBase
    Private WithEvents mc_parentRichTextBox As RichTextBox 'for specific rich text box events :)
    Private Property parentTextBox() As TextBoxBase
        Get
            Return mc_parentTextBox
        End Get
        Set(ByVal value As TextBoxBase)
            mc_parentRichTextBox = TryCast(value, RichTextBox)
            mc_parentTextBox = value
        End Set
    End Property

    Private Sub parentTextBox_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles mc_parentTextBox.KeyUp
        Select Case e.KeyCode
            Case Keys.Apps 'right click menu
                If ContextMenuStrip.Visible Then
                    'menu is open already ... so hide
                    ContextMenuStrip.Close()
                Else
                    MenuSpellClickReturn = New MenuSpellClickReturnItem(parentTextBox, False)
                End If
            Case Keys.F7
                'spell check dialogue
                If Settings.AllowF7 Then
                    ShowDialog()
                    'update the ignored/added/removed words...
                    RepaintTextBox()
                End If
            Case Keys.F3
                If Settings.AllowF3 Then
                    'cycle through case's ... sentence, propper, upper, lower, origional
                    CycleCase()
                End If
        End Select
    End Sub

    Public Sub ShowDialog()
        If CurrentDictionary IsNot Nothing AndAlso CurrentDictionary.Loading = False Then
            Using SpellCheckDialog As New SpellCheckDialog
                SpellCheckDialog.ShowDialog(parentTextBox, Me)
            End Using
        End If
    End Sub

    Private Sub parentTextBox_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_parentTextBox.SizeChanged
        textBoxGraphics = Graphics.FromHwnd(parentTextBox.Handle)
    End Sub

    Private Sub parentTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_parentTextBox.TextChanged
        'RepaintTextBox()
        'parentTextBox.Invalidate()
    End Sub

    Private Const WM_VSCROLL As Integer = &H115
    Private Const WM_MOUSEWHEEL As Integer = &H20A
    Private Const WM_PAINT As Integer = &HF

    Protected Overrides Sub WndProc(ByRef m As Message)
        'Debug.Print(Now & " - " & m.ToString)
        If Settings.ShowMistakes Then
            Select Case m.Msg
                Case WM_PAINT
                    If Me.Settings.RenderCompatibility Then
                        'old draw method...
                        parentTextBox.Invalidate()
                        CloseOverlay()
                        MyBase.WndProc(m)
                        Me.CustomPaint()
                    Else
                        MyBase.WndProc(m)
                        RepaintTextBox()
                        'OpenOverlay()
                        'Me.CustomPaint()
                    End If
                Case Else
                    MyBase.WndProc(m)
            End Select
        Else
            MyBase.WndProc(m)
        End If
    End Sub

#End Region

#Region "Constructor"

    Public Sub New(ByVal textBox As TextBoxBase, Optional ByVal Dictionary As Dictionary = Nothing)
        mc_CurrentDictionary = Dictionary
        parentTextBox = textBox
        parentTextBox_ContextMenuStripChanged(parentTextBox, EventArgs.Empty)
        parentTextBox_SizeChanged(parentTextBox, EventArgs.Empty)
        Try
            AssignHandle(textBox.Handle)
        Catch ex As Exception

        End Try
        RepaintTextBox()
        'parentTextBox.Invalidate()
        Application.AddMessageFilter(Me)
    End Sub

#End Region

#Region "Update the WndProc Handle when the textbox gets a new handle - for when properties like RightToLeft are changed"

    Private Sub mc_parentTextBox_HandleCreated(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_parentTextBox.HandleCreated
        AssignHandle(TextBox.Handle)
    End Sub

#End Region

End Class