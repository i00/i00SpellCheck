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

    Private Sub SpellMenuItems_WordAdded(ByVal sender As Object, ByVal e As Menu.AddSpellItemsToMenu.SpellItemEventArgs) Handles SpellMenuItems.WordAdded
        Try
            DictionaryAddWord(e.Word)
        Catch ex As Exception
            MsgBox("The following error occured adding """ & e.Word & """ to the dictionary:" & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub SpellMenuItems_WordChanged(ByVal sender As Object, ByVal e As Menu.AddSpellItemsToMenu.SpellItemEventArgs) Handles SpellMenuItems.WordChanged
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

    Private Sub SpellMenuItems_WordUnIgnored(ByVal sender As Object, ByVal e As Menu.AddSpellItemsToMenu.SpellItemEventArgs) Handles SpellMenuItems.WordUnIgnored
        Try
            DictionaryUnIgnoreWord(e.Word)
        Catch ex As Exception
            MsgBox("The following error ignoring """ & e.Word & """:" & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub SpellMenuItems_WordIgnored(ByVal sender As Object, ByVal e As Menu.AddSpellItemsToMenu.SpellItemEventArgs) Handles SpellMenuItems.WordIgnored
        Try
            DictionaryIgnoreWord(e.Word)
        Catch ex As Exception
            MsgBox("The following error ignoring """ & e.Word & """:" & vbCrLf & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub SpellMenuItems_WordRemoved(ByVal sender As Object, ByVal e As Menu.AddSpellItemsToMenu.SpellItemEventArgs) Handles SpellMenuItems.WordRemoved
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

                ScrollToCaret()

                Dim ChrPos = Me.parentTextBox.GetPositionFromCharIndex(parentTextBox.SelectionStart)
                Dim LineHeight As Integer = GetLineHeightFromCharPosition(parentTextBox.SelectionStart)

                ContextMenuStrip.Top = AddY + ChrPos.Y + LineHeight
                ContextMenuStrip.Left = AddX + ChrPos.X
            End If
        End If

        CheckContextMenuLocation()
        LocationChanging = False
    End Sub

    'scroll the caret into view
    Protected Sub ScrollToCaret()
        SendMessage(Me.parentTextBox.Handle, EM_SCROLLCARET, 0, 0)
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

            If OKToSpellCheck Then
                SpellMenuItems.ContextMenuStrip = ContextMenuStrip
                SpellMenuItems.AddItems(MenuSpellClickReturn.Word, CurrentDictionary, CurrentDefinitions, CurrentSynonyms, Settings)
            End If
        Else
            'no word clicked on
        End If
    End Sub

    Private WithEvents SpellMenuItems As New Menu.AddSpellItemsToMenu()

#End Region

End Class