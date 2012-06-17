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

Partial Class SpellCheckTextBox

#Region "Custom Tool Strip Items"

    'empty interface so we can identify our toolstrip item for spelling correction over custom ones
    Private Interface tsiSpelling

    End Interface

    Private Class tsiTextSpellingSeparator
        Inherits MenuTextSeperator
        Implements tsiSpelling
    End Class

    Private Class tsiSpellingSeparator
        Inherits ToolStripSeparator
        Implements tsiSpelling
    End Class

    Friend Class tsiSpellingSuggestion
        Inherits ToolStripMenuItem
        Implements tsiSpelling
        Public UnderlyingValue As String
        Public Sub New(ByVal Text As String, ByVal UnderlyingValue As String, Optional ByVal Image As Image = Nothing)
            Me.Text = Text
            Me.UnderlyingValue = UnderlyingValue
            Me.Image = Image
        End Sub
        Dim LastStateSelected As Boolean = False
        Public Event SelectionChanged(ByVal sender As Object, ByVal e As EventArgs)
        Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
            MyBase.OnPaint(e)
            If LastStateSelected <> Me.Selected Then
                RaiseEvent SelectionChanged(Me, EventArgs.Empty)
                LastStateSelected = Me.Selected
            End If
        End Sub
    End Class

    Private Class tsiDefinition
        Inherits HTMLMenuItem
        Implements tsiSpelling
        Public Sub New(ByVal HTMLText As String)
            MyBase.New(HTMLText)
        End Sub
    End Class

#End Region

#Region "Class to get info on cursor/word location when menu is opened"

    'Used to get information on the location that was clicked on from a content menu
    Private Class MenuSpellClickReturnItem

        Public Word As String
        Public WordStart As Integer
        Public WordEnd As Integer
        Public UseMouseLocation As Boolean

        Public Sub New(ByVal TextBox As TextBoxBase, Optional ByVal UseMouseLocation As Boolean = False)
            Dim CharIndex As Integer
            Me.UseMouseLocation = UseMouseLocation
            If UseMouseLocation = True Then
                'Work out the location that was clicked on
                Dim Point0 As New Point(0, 0)
                Dim cmLocationRelToTxtLocationX = System.Windows.Forms.Control.MousePosition.X - TextBox.PointToScreen(Point0).X - (TextBox.ClientRectangle.X)
                Dim cmLocationRelToTxtLocationY = System.Windows.Forms.Control.MousePosition.Y - TextBox.PointToScreen(Point0).Y
                'Get the char index for the location we clicked on
                CharIndex = TextBox.GetCharIndexFromPosition(New Point(cmLocationRelToTxtLocationX, cmLocationRelToTxtLocationY))
            Else
                'Use the location that the caret is...
                If TextBox.SelectionLength = 0 Then
                    'ok
                    CharIndex = TextBox.SelectionStart
                Else
                    'multi text selected
                    Exit Sub
                End If
            End If

            'Get the word that is @ the CharIndex
            Dim theText As String = Dictionary.Formatting.RemoveWordBreaks(TextBox.Text)
            Dim LeftSide As String = Left(theText, CharIndex)
            Dim RightSide As String = Right(theText, Len(theText) - CharIndex)
            LeftSide = LeftSide.Split(" "c).Last
            RightSide = RightSide.Split(" "c).First
            Word = LeftSide & RightSide



            'and fill the chr indexes
            WordStart = CharIndex - Len(LeftSide)
            WordEnd = CharIndex + Len(RightSide)

        End Sub

    End Class

#End Region

#Region "Standard context menu"

    Public Class StandardContextMenuStrip
        Inherits ContextMenuStrip

#Region "Buttons"

        Private WithEvents Undo As ToolStripMenuItem
        Private Sub Undo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Undo.Click
            Dim theTextBox = TryCast(Me.SourceControl, TextBoxBase)
            If theTextBox IsNot Nothing Then
                theTextBox.Undo()
            End If
        End Sub

        Private WithEvents Cut As ToolStripMenuItem
        Private Sub Cut_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cut.Click
            Dim theTextBox = TryCast(Me.SourceControl, TextBoxBase)
            If theTextBox IsNot Nothing Then
                theTextBox.Cut()
            End If
        End Sub

        Private WithEvents Copy As ToolStripMenuItem
        Private Sub Copy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Copy.Click
            Dim theTextBox = TryCast(Me.SourceControl, TextBoxBase)
            If theTextBox IsNot Nothing Then
                theTextBox.Copy()
            End If
        End Sub

        Private WithEvents Paste As ToolStripMenuItem
        Private Sub Paste_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Paste.Click
            Dim theTextBox = TryCast(Me.SourceControl, TextBoxBase)
            If theTextBox IsNot Nothing Then
                theTextBox.Paste()
            End If
        End Sub

        Private WithEvents Delete As ToolStripMenuItem
        Private Sub Delete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Delete.Click
            Dim theTextBox = TryCast(Me.SourceControl, TextBoxBase)
            If theTextBox IsNot Nothing Then

                'lock window from updating
                LockWindowUpdate(theTextBox.Handle)
                theTextBox.SuspendLayout()

                Dim SelStart = theTextBox.SelectionStart

                Dim mc_parentRichTextBox = TryCast(theTextBox, RichTextBox)
                If mc_parentRichTextBox IsNot Nothing Then
                    'rich text box :(... use alternate method ... this will ensure that the formatting isn't lost
                    mc_parentRichTextBox.Select(SelStart, theTextBox.SelectionLength)
                    mc_parentRichTextBox.SelectedText = ""
                Else
                    Dim OldVertPos = GetScrollPos(theTextBox.Handle, SB_VERT)

                    'replace the text
                    theTextBox.Text = Strings.Left(theTextBox.Text, SelStart) & _
                                      Strings.Right(theTextBox.Text, Len(theTextBox.Text) - (theTextBox.SelectionStart + theTextBox.SelectionLength))
                    '... and select the replaced text
                    theTextBox.SelectionStart = SelStart

                    'Set scroll bars to what they were
                    SendMessage(theTextBox.Handle, EM_SCROLL, SB_TOP, 0) ' reset Vscroll to top
                    SendMessage(theTextBox.Handle, EM_LINESCROLL, 0, OldVertPos) ' set Vscroll to last saved pos.
                End If

                'unlock window updates
                theTextBox.ResumeLayout()
                LockWindowUpdate(IntPtr.Zero)

            End If
        End Sub

        Private WithEvents SelectAll As ToolStripMenuItem
        Private Sub SelectAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SelectAll.Click
            Dim theTextBox = TryCast(Me.SourceControl, TextBoxBase)
            If theTextBox IsNot Nothing Then
                theTextBox.SelectAll()
            End If
        End Sub

#End Region

        Public Sub New()
            Undo = New ToolStripMenuItem("&Undo", My.Resources.Undo)
            Me.Items.Add(Undo)

            Me.Items.Add(New ToolStripSeparator)

            If System.Threading.Thread.CurrentThread.GetApartmentState = Threading.ApartmentState.STA Then
                'allow cut/copy/paste - doesn't work unless STA thread
                Cut = New ToolStripMenuItem("Cu&t", My.Resources.Cut)
                Me.Items.Add(Cut)
                Copy = New ToolStripMenuItem("&Copy", My.Resources.Copy)
                Me.Items.Add(Copy)
                Paste = New ToolStripMenuItem("&Paste", My.Resources.Paste)
                Me.Items.Add(Paste)
            End If

            Delete = New ToolStripMenuItem("&Delete", My.Resources.Delete)
            Me.Items.Add(Delete)

            Me.Items.Add(New ToolStripSeparator)

            SelectAll = New ToolStripMenuItem("Select &All", My.Resources.SelectAll)
            Me.Items.Add(SelectAll)
        End Sub

        Private Sub StandardContextMenuStrip_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Opening
            Dim theTextBox = TryCast(Me.SourceControl, TextBoxBase)
            If theTextBox IsNot Nothing Then
                Undo.Enabled = theTextBox.CanUndo

                Cut.Enabled = theTextBox.SelectionLength > 0
                Copy.Enabled = theTextBox.SelectionLength > 0
                Paste.Enabled = Clipboard.GetText <> ""
                Delete.Enabled = theTextBox.SelectionLength > 0

                SelectAll.Enabled = theTextBox.SelectionLength <> Len(theTextBox.Text)

            End If
        End Sub

    End Class

#End Region

#Region "Actual Menu"

#Region "Menu Item Events"

    Private Sub SpellMenuItems_WordAdded(ByVal sender As Object, ByVal e As AddSpellItemsToMenu.SpellItemEventArgs) Handles SpellMenuItems.WordAdded
        Try
            DictionaryAddWord(e.Word)
        Catch ex As Exception
            MsgBox("The following error occured adding """ & e.Word & """ to the dictionary:" & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub SpellMenuItems_WordChanged(ByVal sender As Object, ByVal e As AddSpellItemsToMenu.SpellItemEventArgs) Handles SpellMenuItems.WordChanged
        'lock window from updating
        LockWindowUpdate(parentTextBox.Handle)
        parentTextBox.SuspendLayout()

        If parentRichTextBox IsNot Nothing Then
            'rich text box :(... use alternate method ... this will ensure that the formatting isn't lost
            parentRichTextBox.Select(LastMenuSpellClickReturn.WordStart, LastMenuSpellClickReturn.Word.Length)
            parentRichTextBox.SelectedText = e.Word
        Else
            'standard text box .. can just replace all of the text

            'Get old scroll bar position
            Dim OldVertPos = GetScrollPos(parentTextBox.Handle, SB_VERT)

            'replace the text
            parentTextBox.Text = Left(parentTextBox.Text, LastMenuSpellClickReturn.WordStart) & _
                                 e.Word & _
                                 Right(parentTextBox.Text, Len(parentTextBox.Text) - LastMenuSpellClickReturn.WordEnd)

            'Set scroll bars to what they were
            SendMessage(parentTextBox.Handle, EM_SCROLL, SB_TOP, 0) ' reset Vscroll to top
            SendMessage(parentTextBox.Handle, EM_LINESCROLL, 0, OldVertPos) ' set Vscroll to last saved pos.
        End If

        '... and select the replaced text
        parentTextBox.SelectionStart = LastMenuSpellClickReturn.WordStart
        parentTextBox.SelectionLength = Len(e.Word)

        'unlock window updates
        parentTextBox.ResumeLayout()
        LockWindowUpdate(IntPtr.Zero)
    End Sub

    Private Sub SpellMenuItems_WordIgnored(ByVal sender As Object, ByVal e As AddSpellItemsToMenu.SpellItemEventArgs) Handles SpellMenuItems.WordIgnored
        Try
            DictionaryIgnoreWord(e.Word)
        Catch ex As Exception
            MsgBox("The following error ignoring """ & e.Word & """:" & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub SpellMenuItems_WordRemoved(ByVal sender As Object, ByVal e As AddSpellItemsToMenu.SpellItemEventArgs) Handles SpellMenuItems.WordRemoved
        Try
            DictionaryRemoveWord(e.Word)
        Catch ex As Exception
            MsgBox("The following error occured removing """ & e.Word & """ from the dictionary:" & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

#End Region

    Dim MenuSpellClickReturn As MenuSpellClickReturnItem
    Dim LastMenuSpellClickReturn As MenuSpellClickReturnItem

    Private WithEvents ContextMenuStrip As ContextMenuStrip

    Private Sub parentTextBox_ContextMenuStripChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_Control.ContextMenuStripChanged
        ContextMenuStrip = parentTextBox.ContextMenuStrip
        If ContextMenuStrip Is Nothing Then
            ContextMenuStrip = New StandardContextMenuStrip
            parentTextBox.ContextMenuStrip = ContextMenuStrip
        End If
    End Sub

    Private Sub ContextMenuStrip_Closed(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripDropDownClosedEventArgs) Handles ContextMenuStrip.Closed
        SpellMenuItems.RemoveSpellMenuItems()
        LastMenuSpellClickReturn = MenuSpellClickReturn
        MenuSpellClickReturn = Nothing
    End Sub

    Private Sub ContextMenuStrip_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ContextMenuStrip.LocationChanged
        Static LocationChanging As Boolean
        If LocationChanging Then Exit Sub
        LocationChanging = True
        If CurrentDictionary IsNot Nothing AndAlso CurrentDictionary.Loading = True Then Exit Sub
        If ContextMenuStrip.SourceControl Is parentTextBox Then
            If MenuSpellClickReturn IsNot Nothing AndAlso MenuSpellClickReturn.UseMouseLocation = False Then 'If keyboard context menu was pressed ...
                Dim txtBoxScreenPos = Me.parentTextBox.PointToScreen(Point.Empty)
                Dim AddX = txtBoxScreenPos.X
                Dim AddY = txtBoxScreenPos.Y

                'scroll the caret into view
                SendMessage(Me.parentTextBox.Handle, EM_SCROLLCARET, 0, 0)

                Dim ChrPos = Me.parentTextBox.GetPositionFromCharIndex(parentTextBox.SelectionStart)
                Dim LineHeight As Integer = GetLineHeightFromCharPosition(parentTextBox.SelectionStart)

                ContextMenuStrip.Top = AddY + ChrPos.Y + LineHeight
                ContextMenuStrip.Left = AddX + ChrPos.X
            End If
        End If

        CheckContextMenuLocation()
        LocationChanging = False
    End Sub

    Private Sub ContextMenuStrip_ItemAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemEventArgs) Handles ContextMenuStrip.ItemAdded
        CheckContextMenuLocation()
    End Sub

    Private Sub CheckContextMenuLocation()
        Static isWorking As Boolean = False
        If isWorking Then Exit Sub
        isWorking = True
        Dim ScreenAtPoint = Screen.FromPoint(New Point(ContextMenuStrip.Left, ContextMenuStrip.Top))
        Dim PopupLocation = Control.MousePosition
        If MenuSpellClickReturn IsNot Nothing AndAlso MenuSpellClickReturn.UseMouseLocation = False Then
            'If keyboard context menu was pressed ... use carrot location instead
            'qwertyuiop hrm doesn't seem to work atm ... 
            'PopupLocation = Me.parentTextBox.GetPositionFromCharIndex(parentTextBox.SelectionStart)
            'PopupLocation = Me.parentTextBox.PointToScreen(PopupLocation)
            GoTo AllDone
        End If
        If ContextMenuStrip.Right >= ScreenAtPoint.WorkingArea.Right - 16 Then
            ContextMenuStrip.Left = PopupLocation.X - ContextMenuStrip.Width
        End If
        If ContextMenuStrip.Bottom >= ScreenAtPoint.WorkingArea.Bottom - 16 Then
            ContextMenuStrip.Top = PopupLocation.Y - ContextMenuStrip.Height
        End If
AllDone:
        isWorking = False
    End Sub

    Private Sub ContextMenuStrip_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ContextMenuStrip.SizeChanged
        CheckContextMenuLocation()
    End Sub

    Private Sub ContextMenuStrip_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip.Opening
        If ContextMenuStrip.Visible = True Then
            'already showing ... no need to run this...
            Exit Sub
        End If
        If CurrentDictionary IsNot Nothing AndAlso CurrentDictionary.Loading Then Exit Sub
        If ContextMenuStrip.SourceControl Is parentTextBox Then
            If MenuSpellClickReturn IsNot Nothing Then
                'use the existing MenuSpellClickReturn - right click button was pressed on keyboard
                'show it at location of the cursor instead of the center?
                'qwertyuiop - DOESN'T WORK HERE?
                'moved to ContextMenuStrip_LocationChanged
            Else
                'use mouse location as this was fired with a right click
                MenuSpellClickReturn = New MenuSpellClickReturnItem(parentTextBox, True)
            End If

            If MenuSpellClickReturn.Word <> "" AndAlso parentTextBox.IsSpellCheckEnabled Then
                SpellMenuItems.ContextMenuStrip = ContextMenuStrip
                SpellMenuItems.AddItems(MenuSpellClickReturn.Word, CurrentDictionary, CurrentDefinitions, CurrentSynonyms, Settings)
            End If
        Else
            'no word clicked on
        End If
    End Sub

    Private WithEvents SpellMenuItems As New AddSpellItemsToMenu()

#Region "Add spell items class"

    Public Class AddSpellItemsToMenu
        Implements IDisposable

        Public Sub RemoveSpellMenuItems()
            'remove all the previous Suggestions from a menu
            If ContextMenuStrip IsNot Nothing Then
                For Each item In ContextMenuStrip.Items.OfType(Of tsiSpelling)().ToArray
                    ContextMenuStrip.Items.Remove(TryCast(item, ToolStripItem))
                Next
            End If
        End Sub

        Private WithEvents SpellToolTip As New HTMLToolTip

        Dim DefinitionSet As Definitions
        Dim Dictionary As Dictionary

        Dim mtTip As System.Threading.Thread

        Private Sub LookupDefForTip(ByVal oTsi As Object)
            Dim tsi = TryCast(oTsi, tsiSpellingSuggestion)
            If tsi IsNot Nothing Then
                Dim WordDef = DefinitionSet.FindWord(tsi.UnderlyingValue, Dictionary).ToString
                If WordDef <> "" Then
                    'had to create a new one each time ... otherwise it doesn't fade when moving between items
                    SpellToolTip.Dispose()
                    SpellToolTip = New HTMLToolTip
                    ShowTip(WordDef, tsi)
                Else
                    HideTip(tsi)
                End If
            End If
        End Sub

        Private Delegate Sub ShowTip_cb(ByVal WordDef As String, ByVal tsi As tsiSpellingSuggestion)
        Private Sub ShowTip(ByVal WordDef As String, ByVal tsi As tsiSpellingSuggestion)
            If tsi.Owner IsNot Nothing Then
                If tsi.Owner.InvokeRequired Then
                    Dim ShowTip_cb As New ShowTip_cb(AddressOf ShowTip)
                    tsi.Owner.Invoke(ShowTip_cb, WordDef, tsi)
                Else
                    Dim ToolTipPoint = New Point(tsi.Bounds.Right, tsi.Bounds.Top)
                    SpellToolTip.ShowHTML(WordDef, tsi.Owner, ToolTipPoint)
                End If
            End If
        End Sub

        Private Delegate Sub HideTip_cb(ByVal tsi As tsiSpellingSuggestion)
        Private Sub HideTip(ByVal tsi As tsiSpellingSuggestion)
            If tsi.Owner IsNot Nothing Then
                If tsi.Owner.InvokeRequired Then
                    Dim HideTip_cb As New HideTip_cb(AddressOf HideTip)
                    tsi.Owner.Invoke(HideTip_cb, tsi)
                Else
                    Dim ToolTipPoint = New Point(tsi.Bounds.Right, tsi.Bounds.Top)
                    SpellToolTip.Hide(tsi.Owner)
                End If
            End If
        End Sub

        Private Sub SpellToolTip_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim tsi = TryCast(sender, tsiSpellingSuggestion)
            If tsi IsNot Nothing Then
                If mtTip IsNot Nothing AndAlso mtTip.IsAlive Then
                    mtTip.Abort()
                End If
                mtTip = New System.Threading.Thread(AddressOf LookupDefForTip)
                mtTip.Name = "Definition Tooltip - " & tsi.UnderlyingValue
                mtTip.IsBackground = True
                mtTip.Start(tsi)
            End If
        End Sub

        Private Sub ContextMenuStrip_Closed(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripDropDownClosedEventArgs) Handles ContextMenuStrip.Closed
            Dim ContextMenuStrip = TryCast(sender, ContextMenuStrip)
            If ContextMenuStrip IsNot Nothing Then
                SpellToolTip.Hide(ContextMenuStrip)
            End If
        End Sub

        Private Sub SpellToolTip_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
            If mtTip IsNot Nothing AndAlso mtTip.IsAlive Then
                mtTip.Abort()
            End If
            Dim tsi = TryCast(sender, ToolStripItem)
            If tsi IsNot Nothing Then
                SpellToolTip.Hide(tsi.Owner)
            End If
        End Sub

        Public WithEvents ContextMenuStrip As ContextMenuStrip

        Public Sub New()

        End Sub

        Public Sub New(ByVal ContextMenuStrip As ContextMenuStrip)
            Me.ContextMenuStrip = ContextMenuStrip
        End Sub

        Public Class SpellItemEventArgs
            Inherits EventArgs
            Public Word As String
        End Class

        Public Event WordChanged(ByVal sender As Object, ByVal e As SpellItemEventArgs)
        Public Event WordRemoved(ByVal sender As Object, ByVal e As SpellItemEventArgs)
        Public Event WordAdded(ByVal sender As Object, ByVal e As SpellItemEventArgs)
        Public Event WordIgnored(ByVal sender As Object, ByVal e As SpellItemEventArgs)

        Private Sub SpellItemClick(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim tsiSpellingSuggestion = TryCast(sender, tsiSpellingSuggestion)
            If tsiSpellingSuggestion IsNot Nothing Then
                RaiseEvent WordChanged(Me, New SpellItemEventArgs() With {.Word = tsiSpellingSuggestion.UnderlyingValue})
            End If
        End Sub

        Private Sub SpellRemoveWordClick(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim tsiSpellingSuggestion = TryCast(sender, tsiSpellingSuggestion)
            If tsiSpellingSuggestion IsNot Nothing Then
                RaiseEvent WordRemoved(Me, New SpellItemEventArgs() With {.Word = tsiSpellingSuggestion.UnderlyingValue})
            End If
        End Sub

        Private Sub SpellAddNewWordClick(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim tsiSpellingSuggestion = TryCast(sender, tsiSpellingSuggestion)
            If tsiSpellingSuggestion IsNot Nothing Then
                RaiseEvent WordAdded(Me, New SpellItemEventArgs() With {.Word = tsiSpellingSuggestion.UnderlyingValue})
            End If
        End Sub

        Private Sub SpellIgnoreWordClick(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim tsiSpellingSuggestion = TryCast(sender, tsiSpellingSuggestion)
            If tsiSpellingSuggestion IsNot Nothing Then
                RaiseEvent WordIgnored(Me, New SpellItemEventArgs() With {.Word = tsiSpellingSuggestion.UnderlyingValue})
            End If
        End Sub

        Private Sub AddSeperatorIfHasMenuItems()
            'qwertyuiop - doesn't work as required ... :(
            'If (From xItem In ContextMenuStrip.Items.OfType(Of ToolStripItem)() Where xItem.Visible = True).FirstOrDefault IsNot Nothing Then
            'we have @ least one item visible :)... so add the seperator
            ContextMenuStrip.Items.Add(New tsiSpellingSeparator())
            'End If
        End Sub

        Public Sub AddItems(ByVal Word As String, ByVal Dictionary As Dictionary, ByVal DefinitionSet As Definitions, ByVal Synonyms As Synonyms, ByVal Settings As SpellCheckSettings)
            If Dictionary IsNot Nothing AndAlso Dictionary.Loading Then Exit Sub
            Me.DefinitionSet = DefinitionSet
            Me.Dictionary = Dictionary

            Dim Result = Dictionary.SpellCheckWord(Word)
            Dim NiceWord As String = Dictionary.Formatting.RemoveApoS(Word)
            If NiceWord = "" Then Exit Sub 'shouldn't happen
            If Result = Dictionary.SpellCheckWordError.OK OrElse Result = Dictionary.SpellCheckWordError.CaseError OrElse Dictionary.Count = 0 Then
                'word is in the dictionary... regardless of case...
                'Lookup the def
                If Settings.AllowInMenuDefs Then
                    Dim WordDef = DefinitionSet.FindWord(Word, Dictionary, System.Drawing.ColorTranslator.ToHtml(DrawingFunctions.BlendColor(Color.FromKnownColor(KnownColor.MenuText), Color.FromKnownColor(KnownColor.Menu)))).ToString
                    If WordDef <> "" Then
                        AddSeperatorIfHasMenuItems()
                        Dim tsib As New tsiDefinition(WordDef)
                        ContextMenuStrip.Items.Add(tsib)
                    End If
                End If

                'Sugguest other Synonyms
                If Settings.AllowChangeTo Then
                    Dim MatchedSynonyms = Synonyms.FindWord(Word)
                    If MatchedSynonyms IsNot Nothing Then
                        'Add change to...
                        AddSeperatorIfHasMenuItems()
                        Dim tsiMain = New tsiSpellingSuggestion("Change to ...", "")
                        ContextMenuStrip.Items.Add(tsiMain)
                        For Each item In MatchedSynonyms
                            Dim tsiSeparator As New tsiTextSpellingSeparator() With {.Text = item.TypeDescription & If(item.WordType <> Synonyms.FindWordReturn.WordTypes.Other, " (" & item.WordType.ToString & ")", "")}
                            tsiMain.DropDownItems.Add(tsiSeparator)
                            For Each iSyn In item
                                Dim tsi = New tsiSpellingSuggestion(iSyn, iSyn)
                                AddHandler tsi.Click, AddressOf SpellItemClick
                                AddHandler tsi.MouseEnter, AddressOf SpellToolTip_MouseEnter
                                AddHandler tsi.MouseLeave, AddressOf SpellToolTip_MouseLeave
                                'AddHandler tsi.SelectionChanged, AddressOf SpellToolTip_SelectionChanged
                                tsiMain.DropDownItems.Add(tsi)
                            Next
                        Next
                    End If
                End If

            End If

            If Dictionary.Count = 0 Then Exit Sub

            Select Case Result
                Case Dictionary.SpellCheckWordError.OK
                    If Settings.AllowRemovals Then
                        AddSeperatorIfHasMenuItems()
                        Dim tsi = New tsiSpellingSuggestion("Remove " & NiceWord & " from dictionary", Word, My.Resources.RemoveWord)
                        AddHandler tsi.Click, AddressOf SpellRemoveWordClick
                        ContextMenuStrip.Items.Add(tsi)
                    End If
                Case Else
                    Dim Suggestions As List(Of Dictionary.SpellCheckSuggestionInfo) = Nothing

                    Dim Ucase1stLetter As Boolean
                    If Char.IsUpper(NiceWord(0)) Then
                        'upper case - so lets make suggestions 1st letter in ucase :)
                        Ucase1stLetter = True
                    End If
                    Select Case Result
                        Case Dictionary.SpellCheckWordError.CaseError
                            Suggestions = (From xItem In Dictionary Where LCase(xItem.Entry) = Word.ToLower Select New Dictionary.SpellCheckSuggestionInfo(100, xItem.Entry)).ToList
                        Case Dictionary.SpellCheckWordError.SpellError
                            Suggestions = (From xItem In Dictionary.SpellCheckSuggestions(Word) Order By xItem.Closness Descending).ToList
                    End Select

                    If Ucase1stLetter AndAlso Suggestions IsNot Nothing Then
                        For Each item In Suggestions
                            Dim arrWord = item.Word.ToArray
                            arrWord(0) = Char.ToUpper(arrWord(0))
                            item.Word = CStr(arrWord)
                        Next
                    End If

                    If Suggestions IsNot Nothing Then
                        'word is not in the dictionary ... so suggest
                        AddSeperatorIfHasMenuItems()

                        If Settings.AllowAdditions = True AndAlso Result <> Dictionary.SpellCheckWordError.CaseError Then
                            Dim theWord As String = Word
                            If NiceWord = LCase(NiceWord) Then
                                'not case sensitive
                                Dim tsi = New tsiSpellingSuggestion("Add " & NiceWord & " to dictionary", NiceWord, My.Resources.AddWord)
                                AddHandler tsi.Click, AddressOf SpellAddNewWordClick
                                ContextMenuStrip.Items.Add(tsi)
                            Else
                                'may be case sensitive
                                Dim tsi = New tsiSpellingSuggestion("Add to dictionary", "", My.Resources.AddWord)
                                Dim tsiCase = New tsiSpellingSuggestion("Case sensitive: " & NiceWord, NiceWord, My.Resources.CaseSensitive)
                                AddHandler tsiCase.Click, AddressOf SpellAddNewWordClick
                                tsi.DropDownItems.Add(tsiCase)
                                tsiCase = New tsiSpellingSuggestion("Case insensitive: " & LCase(NiceWord), LCase(NiceWord), My.Resources.CaseInsensitive)
                                AddHandler tsiCase.Click, AddressOf SpellAddNewWordClick
                                tsi.DropDownItems.Add(tsiCase)
                                ContextMenuStrip.Items.Add(tsi)
                            End If
                        End If

                        If Settings.AllowIgnore = True Then
                            Dim tsi = New tsiSpellingSuggestion("Ignore " & NiceWord, NiceWord, My.Resources.Ignore)
                            AddHandler tsi.Click, AddressOf SpellIgnoreWordClick
                            ContextMenuStrip.Items.Add(tsi)
                        End If

                        If Suggestions.Count = 0 Then
                            Dim tsi = New tsiSpellingSuggestion("No suggestions", "", My.Resources.SpellCheck)
                            tsi.Enabled = False
                            ContextMenuStrip.Items.Add(tsi)
                        Else
                            Dim TopCloseness = Suggestions.First.Closness
                            Dim FilteredSuggestions = (From xItem In Suggestions Order By xItem.Closness Descending, xItem.Word Ascending Where xItem.Closness >= TopCloseness * 0.75).ToArray
                            FilteredSuggestions = (From xItem In FilteredSuggestions Where Array.IndexOf(FilteredSuggestions, xItem) < 15).ToArray

                            For Each item In FilteredSuggestions
                                'hrm interesting ... had in the line below If(item Is Suggestions.First.... but this wouldn't work for capitialisation
                                Dim tsi = New tsiSpellingSuggestion(item.Word, item.Word, If(item.Word = Suggestions.First.Word, My.Resources.SpellCheck, Nothing))
                                AddHandler tsi.Click, AddressOf SpellItemClick
                                AddHandler tsi.MouseEnter, AddressOf SpellToolTip_MouseEnter
                                AddHandler tsi.MouseLeave, AddressOf SpellToolTip_MouseLeave
                                'AddHandler tsi.SelectionChanged, AddressOf SpellToolTip_SelectionChanged
                                ContextMenuStrip.Items.Add(tsi)
                            Next
                        End If
                    End If
            End Select
        End Sub


        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: free other state (managed objects).
                    'clean up objects
                    SpellToolTip.Dispose()
                    If mtTip IsNot Nothing AndAlso mtTip.IsAlive Then
                        mtTip.Abort()
                    End If
                End If

                ' TODO: free your own state (unmanaged objects).
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

#Region " IDisposable Support "
        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

#End Region

#End Region

End Class