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

Public Class frmDictionaryEditor

    Dim mc_Dictionary As FlatFileDictionary
    Private Property Dictionary() As FlatFileDictionary
        Get
            Return mc_Dictionary
        End Get
        Set(ByVal value As FlatFileDictionary)
            mc_Dictionary = value
            UpdateLV()
        End Set
    End Property

    Public Overloads Function ShowDialog(ByVal dict As FlatFileDictionary) As Dictionary
        If dict Is Nothing Then dict = New FlatFileDictionary
        If dict.Filename <> "" Then
            Me.Text = "Dictionary Editor " & " - " & dict.Filename
        End If
        Me.Dictionary = dict
        Me.StartPosition = FormStartPosition.CenterParent
        MyBase.ShowDialog()
        'rebuild the dictionary from the new data
        ReBuildDictionary()
        Return dict
    End Function

    Private Sub ReBuildDictionary()
        mc_Dictionary.WordList.Clear()

        'non-user
        mc_Dictionary.WordList.AddRange((From xItem In blDictItems.FullList Where xItem.User = False))

        'user-deleted
        mc_Dictionary.WordList.AddRange((From xItem In blDictItems.FullList Where xItem.User = False AndAlso xItem.ItemState = FlatFileDictionary.DictionaryItem.eItemState.Delete Select New FlatFileDictionary.DictionaryItem(xItem.Entry, FlatFileDictionary.DictionaryItem.eItemState.Delete) With {.User = True}))

        'user-ignored / added
        mc_Dictionary.WordList.AddRange((From xItem In blDictItems.FullList Where xItem.User = True))

    End Sub

    Private WithEvents blDictItems As New i00BindingList.BindingListView(Of FlatFileDictionary.DictionaryItem)

    Private Sub UpdateLV()
        blDictItems.Clear()
        blDictItems.AddRange((From xItem In mc_Dictionary.WordList Where Not (xItem.User = True AndAlso xItem.ItemState = FlatFileDictionary.DictionaryItem.eItemState.Delete)))

        If dgvDictItems.DataSource Is Nothing Then
            dgvDictItems.DataSource = blDictItems.BindingSource
            Dim bn As New i00BindingList.BindingNavigatorWithFilter
            bn.Dock = DockStyle.Bottom
            bn.GripStyle = ToolStripGripStyle.Hidden
            Me.Controls.Add(bn)
            bn.BindingSource = blDictItems.BindingSource
        End If
    End Sub

    'prevent changing of non-user items
    Private Sub dgvDictItems_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvDictItems.CellBeginEdit
        If e.RowIndex <> -1 AndAlso e.ColumnIndex <> -1 Then
            Dim DictItem = TryCast(dgvDictItems.Rows(e.RowIndex).DataBoundItem, i00SpellCheck.FlatFileDictionary.DictionaryItem)
            If DictItem IsNot Nothing Then
                If DictItem.User = False Then
                    e.Cancel = True
                End If
            End If
        End If
    End Sub

    'color rows / don't paint checkboxes on non-user rows
    Private Sub dgvDictItems_CellPainting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellPaintingEventArgs) Handles dgvDictItems.CellPainting
        If e.RowIndex <> -1 AndAlso e.ColumnIndex <> -1 Then
            Dim DictItem = TryCast(dgvDictItems.Rows(e.RowIndex).DataBoundItem, i00SpellCheck.FlatFileDictionary.DictionaryItem)
            If DictItem IsNot Nothing Then
                If DictItem.User Then
                    If DictItem.ItemState = FlatFileDictionary.DictionaryItem.eItemState.Ignore Then
                        e.CellStyle.BackColor = DrawingFunctions.BlendColor(Color.Blue, Color.FromKnownColor(KnownColor.Window), 31)
                    End If
                Else
                    If DictItem.ItemState = FlatFileDictionary.DictionaryItem.eItemState.Delete Then
                        e.CellStyle.BackColor = DrawingFunctions.BlendColor(Color.Red, Color.FromKnownColor(KnownColor.Window), 31)
                    Else
                        e.CellStyle.BackColor = DrawingFunctions.BlendColor(Color.Black, Color.FromKnownColor(KnownColor.Window), 31)
                    End If
                    If dgvDictItems.Columns(e.ColumnIndex).DataPropertyName = "pIgnore" Then
                        e.Handled = True
                        e.PaintBackground(e.ClipBounds, True)
                    End If
                End If
            Else
                If dgvDictItems.Columns(e.ColumnIndex).DataPropertyName = "pIgnore" Then
                    e.Handled = True
                    e.PaintBackground(e.ClipBounds, True)
                End If
            End If
        End If
    End Sub

    Private Sub DictionaryEditor_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        dgvDictItems.EndEdit()
    End Sub

    Private Sub tsbSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbSave.Click
        Using sfd As New SaveFileDialog
            sfd.FileName = mc_Dictionary.Filename
            sfd.Filter = "Dictionary Files (*.dic)|*.dic|All Files (*.*)|*.*"
            sfd.FilterIndex = 0
            If sfd.ShowDialog = Windows.Forms.DialogResult.OK Then
                Me.Text = "Dictionary Editor " & " - " & sfd.FileName
                dgvDictItems.EndEdit()
                ReBuildDictionary()
                mc_Dictionary.Save(sfd.FileName, True)
            End If
        End Using
    End Sub

    Private Sub tsbOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbOpen.Click
        Using ofd As New OpenFileDialog
            ofd.FileName = mc_Dictionary.Filename
            ofd.Filter = "Dictionary Files (*.dic)|*.dic|All Files (*.*)|*.*"
            ofd.FilterIndex = 0
            If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
                mc_Dictionary.LoadFromFile(ofd.FileName)
                'to update the binding source
                UpdateLV()
                Me.Text = "Dictionary Editor " & " - " & ofd.FileName
            End If
        End Using
    End Sub

    Private Sub dgvDictItems_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDictItems.Resize
        'qwertyuiop - what the? can't get the word column width to change?
        'tbcWord.Width = dgvDictItems.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - cbcIgnore.Width - dgvDictItems.RowHeadersWidth
    End Sub

    'Prevent user from deleting non-user items
    Private Sub dgvDictItems_UserDeletingRow(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs) Handles dgvDictItems.UserDeletingRow
        Dim DictItem = DirectCast(e.Row.DataBoundItem, i00SpellCheck.FlatFileDictionary.DictionaryItem)
        If DictItem.User = False Then
            If My.Computer.Keyboard.ShiftKeyDown Then
                DictItem.ItemState = FlatFileDictionary.DictionaryItem.eItemState.Correct
            Else
                DictItem.ItemState = FlatFileDictionary.DictionaryItem.eItemState.Delete
            End If
            e.Cancel = True

        End If
    End Sub

#Region "Keys / Key filtering"

    Public Class KeyList
        Inherits List(Of KeyItems)

        Public Class KeyItems
            Public KeyType As KeyTypes
            Public Color As Color
            Public Description As String
            Public Filter As String
            Public Sub New(ByVal KeyType As KeyTypes, ByVal Color As Color, ByVal Description As String, ByVal Filter As String)
                Me.KeyType = KeyType
                Me.Color = Color
                Me.Description = Description
                Me.Filter = Filter
            End Sub
            Public Overrides Function ToString() As String
                Return Replace(Me.KeyType.ToString, "_", " ")
            End Function
        End Class

        Public Enum KeyTypes
            Builtin_Word
            Removed_Word
            User_Word
            Ignored_Word
        End Enum

        Private Function GetEnumFullPath(ByVal [Enum] As [Enum]) As String
            Return Replace([Enum].GetType.FullName, "+", ".") & "." & [Enum].ToString
        End Function

        Public Sub New()
            Dim ShortcutKeyColor = System.Drawing.ColorTranslator.ToHtml(DrawingFunctions.BlendColor(Color.FromKnownColor(KnownColor.ControlText), Color.FromKnownColor(KnownColor.Control)))

            Me.Add(New KeyItems(KeyTypes.Builtin_Word, DrawingFunctions.BlendColor(Color.Black, Color.FromKnownColor(KnownColor.Window), 31), "These words are in the non-user dictionary, items cannot be edited in it, they can only be marked as deleted.", "[ItemState] = " & GetEnumFullPath(FlatFileDictionary.DictionaryItem.eItemState.Correct) & " AndAlso [User]=False"))
            Me.Add(New KeyItems(KeyTypes.Removed_Word, DrawingFunctions.BlendColor(Color.Red, Color.FromKnownColor(KnownColor.Window), 31), "These words are in the non-user dictionary that have been marked as removed by the user dictionary." & vbCrLf & vbCrLf & "Select the row selectors and press <i><font color=" & ShortcutKeyColor & ">&lt;Delete&gt;</font></i> to mark the item as removed" & vbCrLf & "To unmark an item as removed press <i><font color=" & ShortcutKeyColor & ">&lt;Shift&gt;</font></i> + <i><font color=" & ShortcutKeyColor & ">&lt;Delete&gt;</font></i>", "[ItemState] = " & GetEnumFullPath(FlatFileDictionary.DictionaryItem.eItemState.Delete)))
            Me.Add(New KeyItems(KeyTypes.User_Word, Color.FromKnownColor(KnownColor.Window), "These are words have been added to the user dictionary.", "[User] = True AndAlso [ItemState] <> " & GetEnumFullPath(FlatFileDictionary.DictionaryItem.eItemState.Ignore)))
            Me.Add(New KeyItems(KeyTypes.Ignored_Word, DrawingFunctions.BlendColor(Color.Blue, Color.FromKnownColor(KnownColor.Window), 31), "These are words have been marked in the user dictionary to be ignored by the i00 Spell Check engine.", "[User] = True AndAlso [ItemState] = " & GetEnumFullPath(FlatFileDictionary.DictionaryItem.eItemState.Ignore)))
        End Sub
    End Class

    Dim Keys As New KeyList

    Dim tsbFilterButtons As New List(Of ToolStripButton)

    Private Sub frmDictionaryEditor_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        For Each item In Keys
            Dim tsb As New ToolStripButton
            tsb.AutoToolTip = False
            tsb.Text = item.ToString
            tsb.Tag = item.Description
            AddHandler tsb.Click, AddressOf tsbKey_Click
            Dim b As New Bitmap(14, 14)
            Using g = Graphics.FromImage(b)
                g.Clear(item.Color)
                Using p As New Pen(DrawingFunctions.AlphaColor(Color.FromKnownColor(KnownColor.MenuText), 63))
                    g.DrawRectangle(p, New Rectangle(0, 0, b.Width - 1, b.Height - 1))
                End Using
            End Using
            tsb.Image = b
            tsKeys.Items.Add(tsb)

            Dim tsbFilter As New ToolStripButton
            tsbFilter.DisplayStyle = ToolStripItemDisplayStyle.Image
            tsbFilter.Image = My.Resources.Filter
            tsbFilter.Text = "Filter by " & item.ToString
            tsbFilter.Tag = item.Filter
            AddHandler tsbFilter.Click, AddressOf tsbKeyFilter_Click
            tsbFilterButtons.Add(tsbFilter)
            tsKeys.Items.Add(tsbFilter)

            If item IsNot Keys.Last Then
                tsKeys.Items.Add(New ToolStripSeparator)
            End If

        Next
    End Sub

    Public WithEvents KeyTooltip As New HTMLToolTip

    Private Sub tsbKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ToolStripButton = TryCast(sender, ToolStripButton)
        If ToolStripButton IsNot Nothing Then
            KeyTooltip.Dispose()
            KeyTooltip = New HTMLToolTip With {.IsBalloon = True, .ToolTipTitle = ToolStripButton.Text, .ToolTipIcon = ToolTipIcon.Info}
            KeyTooltip.ShowHTML(ToolStripButton.Tag.ToString, tsKeys, New Point(CInt(ToolStripButton.Bounds.Left + (ToolStripButton.Width / 2)), CInt(ToolStripButton.Bounds.Top + (ToolStripButton.Height / 2))), 5000)
        End If
    End Sub

    Private Sub tsbKeyFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ToolStripButton = TryCast(sender, ToolStripButton)
        If ToolStripButton IsNot Nothing Then
            blDictItems.BindingSource.Filter = ToolStripButton.Tag.ToString
        End If
    End Sub

    Private Sub blDictItems_FilterChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles blDictItems.FilterChanged
        For Each item In tsbFilterButtons
            If item.Tag.ToString = blDictItems.Filter Then
                item.Checked = True
            Else
                item.Checked = False
            End If
        Next
    End Sub

#End Region

End Class