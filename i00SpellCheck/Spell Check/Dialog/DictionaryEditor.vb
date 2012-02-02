Public Class DictionaryEditor

    Dim mc_Dictionary As SpellCheckTextBox.Dictionary
    Private Property Dictionary() As SpellCheckTextBox.Dictionary
        Get
            Return mc_Dictionary
        End Get
        Set(ByVal value As SpellCheckTextBox.Dictionary)
            mc_Dictionary = value
            UpdateLV()
        End Set
    End Property

    Public Overloads Function ShowDialog(ByVal dict As SpellCheckTextBox.Dictionary) As SpellCheckTextBox.Dictionary
        If dict Is Nothing Then dict = New SpellCheckTextBox.Dictionary
        Me.Dictionary = dict
        Me.StartPosition = FormStartPosition.CenterParent
        MyBase.ShowDialog()
        Return Dictionary
    End Function

    Private WithEvents bs As New BindingSource

    Private Sub UpdateLV()
        bs.DataSource = mc_Dictionary
        bs.AllowNew = True
        dgvDictItems.DataSource = bs
        BindingNavigator1.BindingSource = bs
    End Sub

    Private Sub DictionaryEditor_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        dgvDictItems.EndEdit()
        RemoveInvalidItems
    End Sub

    Private Sub RemoveInvalidItems()
        Dim lstRemoveItems = (From xItem In mc_Dictionary Where xItem.Entry = "" OrElse xItem.Entry.Contains(" ")).ToArray
        For Each item In lstRemoveItems
            mc_Dictionary.Remove(item)
        Next
        'update the dgv binding...
        Dictionary = Dictionary
    End Sub


    Private Sub tsbSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbSave.Click
        dgvDictItems.EndEdit()
        RemoveInvalidItems()
        Using sfd As New SaveFileDialog
            sfd.FileName = mc_Dictionary.Filename
            sfd.Filter = "Dictionary Files (*.dic)|*.dic|All Files (*.*)|*.*"
            sfd.FilterIndex = 0
            If sfd.ShowDialog = Windows.Forms.DialogResult.OK Then
                mc_Dictionary.Save(sfd.FileName, , True)
            End If
        End Using
    End Sub

    Private Sub tsbOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbOpen.Click
        Using ofd As New OpenFileDialog
            ofd.FileName = mc_Dictionary.Filename
            ofd.Filter = "Dictionary Files (*.dic)|*.dic|All Files (*.*)|*.*"
            ofd.FilterIndex = 0
            If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
                dgvDictItems.DataSource = Nothing
                mc_Dictionary.LoadFromFile(ofd.FileName)
                'to update the binding source
                UpdateLV()
            End If
        End Using
    End Sub

    Private Sub dgvDictItems_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDictItems.CellContentClick

    End Sub

    Private Sub dgvDictItems_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDictItems.Resize
        'qwertyuiop - what the? can't get the word column width to change?
        'tbcWord.Width = dgvDictItems.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - cbcIgnore.Width - dgvDictItems.RowHeadersWidth
    End Sub

End Class