Imports i00SpellCheck

Partial Class Form1

    Const WM_SYSCOMMAND As Integer = &H112
    Const SC_KEYMENU As Integer = &HF100

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        Select Case m.Msg
            Case WM_SYSCOMMAND
                Select Case m.WParam.ToInt32
                    Case SC_KEYMENU
                        Static LastControl As Control
                        If ToolStripDropDownButton1.GetCurrentParent.Focused() Then
                            If LastControl IsNot Nothing Then
                                Try
                                    LastControl.Focus()
                                Catch ex As Exception

                                End Try
                            End If
                        Else
                            LastControl = Me.ActiveControl
                            ToolStripDropDownButton1.GetCurrentParent.Focus()
                            ToolStripDropDownButton1.Select()
                        End If
                    Case Else
                        MyBase.WndProc(m)
                End Select
            Case Else
                MyBase.WndProc(m)
        End Select

    End Sub

    Private Sub Form1_Load_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''do not select the text
        TextBox1.SelectionStart = 0
        TextBox1.SelectionLength = 0

        Dim ThisProjectReferences = System.Reflection.Assembly.GetExecutingAssembly.GetReferencedAssemblies
        'find spell check
        Dim SpellCheckAssemblyName = (From xItem In ThisProjectReferences Where xItem.Name.ToLower = "i00spellcheck").FirstOrDefault
        If SpellCheckAssemblyName IsNot Nothing Then
            'qwertyuiop - should use something like: ... but it doesn't work for UNC's
            '... = System.Drawing.Icon.ExtractAssociatedIcon(Assembly)
            tsbAbout.ToolTipText = "About " & SpellCheckAssemblyName.Name & " " & SpellCheckAssemblyName.Version.ToString
            Dim a = System.Reflection.Assembly.Load(SpellCheckAssemblyName.FullName)
            tsbAbout.Image = IconExtraction.GetDefaultIcon(a.Location, IconExtraction.IconSize.SmallIcon).ToBitmap
            tsbSpellCheck.Image = tsbAbout.Image
            Me.Icon = IconExtraction.GetDefaultIcon(a.Location, IconExtraction.IconSize.LargeIcon)
        End If

        Dim propToolBoxIcon As New ToolboxBitmapAttribute(GetType(PropertyGrid))
        tsbProperties.Image = propToolBoxIcon.GetImage(GetType(PropertyGrid), False)

        Dim URLIcon = IconExtraction.GetDefaultIcon(".url", IconExtraction.IconSize.SmallIcon).ToBitmap
        tsbi00Productions.Image = URLIcon
        tsbVBForums.Image = URLIcon
        tsbCodeProject.Image = URLIcon

        'format rich text
        RichTextBox1.Select(RichTextBox1.Text.IndexOf("i00 .Net Spell Check"), Len("i00 .Net Spell Check"))
        RichTextBox1.SelectionFont = New Font(RichTextBox1.Font, FontStyle.Bold)
        RichTextBox1.Select(RichTextBox1.Text.IndexOf("RichTextBoxes!"), Len("RichTextBoxes!"))
        RichTextBox1.SelectionFont = New Font(RichTextBox1.Font.Name, CSng(RichTextBox1.Font.Size * 1.5), FontStyle.Bold)
        RichTextBox1.Select(RichTextBox1.Text.IndexOf("Rich"), Len("Rich"))
        RichTextBox1.SelectionColor = Color.Red
        RichTextBox1.Select(RichTextBox1.Text.IndexOf("Text"), Len("Text"))
        RichTextBox1.SelectionColor = Color.Green
        RichTextBox1.Select(RichTextBox1.Text.IndexOf("Boxes"), Len("Boxes"))
        RichTextBox1.SelectionColor = Color.Blue
        RichTextBox1.Select(0, 0)
        RichTextBox1.ClearUndo()

        LoadFavicon("http://www.paypal.com/favicon.ico", tsbDonate)
        LoadFavicon("http://vbforums.com/favicon.ico", tsbVBForums)
        LoadFavicon("http://i00Productions.org/favicon.ico", tsbi00Productions)
        LoadFavicon("http://codeproject.com/favicon.ico", tsbCodeProject)

        For Each item In ShowIgnoredToolStripMenuItem.DropDownItems.OfType(Of ToolStripItem)()
            AddHandler item.Click, AddressOf ShowIgnoredToolStripMenuItem_Click
        Next

        tsiDrawStyle.SelectedIndex = 0

        propRichTextBox.SelectedObject = RichTextBox1.SpellCheck
        propTextBox.SelectedObject = TextBox1.SpellCheck
        propDataGridView.SelectedObject = DataGridView1.SpellCheck

        UpdateEnabledCheck()

        Dim GridData As New System.ComponentModel.BindingList(Of GridViewData)
        GridData.Add(New GridViewData("This is a grid view example to demonistrate that i00 Spell Check can be used in grids!"))
        GridData.Add(New GridViewData("So comeon and edit a cell!"))

        DataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        Dim bs As New BindingSource
        bs.DataSource = GridData
        bs.AllowNew = True
        DataGridView1.DataSource = bs
        BindingNavigator1.BindingSource = bs

    End Sub

    Private Class GridViewData
        Dim mc_GridViewExample As String
        <System.ComponentModel.DisplayName("Data Grid View Example")> _
        Public Property GridViewExample() As String
            Get
                Return mc_GridViewExample
            End Get
            Set(ByVal value As String)
                mc_GridViewExample = value
            End Set
        End Property
        Public Sub New(ByVal GridViewExample As String)
            Me.mc_GridViewExample = GridViewExample
        End Sub
        Public Sub New()

        End Sub
    End Class

    Private Class FavIconData
        Public URL As String
        Public tsi As ToolStripItem
        Public Sub New(ByVal URL As String, ByVal tsi As ToolStripItem)
            Me.URL = URL
            Me.tsi = tsi
        End Sub
    End Class

    Public Sub LoadFavicon(ByVal URL As String, ByVal tsi As ToolStripItem)
        Dim t As New System.Threading.Thread(AddressOf MT_DownloadImage)
        t.IsBackground = True
        t.Name = "Downloading " & URL
        t.Start(New FavIconData(URL, tsi))
    End Sub

    Public Sub MT_DownloadImage(ByVal oFavIconData As Object)
        Try
            Dim FavIconData = TryCast(oFavIconData, FavIconData)
            If FavIconData IsNot Nothing Then
                Dim request = System.Net.HttpWebRequest.Create(FavIconData.URL)
                Dim response = request.GetResponse()
                Dim Image As Image = Nothing
                'Try standard Image
                Using stream = response.GetResponseStream()
                    Try
                        Image = Image.FromStream(stream)
                    Catch ex As Exception
                    End Try
                End Using
                'Try icon
                If Image Is Nothing Then
                    Using stream = response.GetResponseStream()
                        Try
                            Using tempStream As New IO.MemoryStream
                                Const BUFFER_SIZE As Integer = 1024
                                Dim buffer(BUFFER_SIZE - 1) As Byte
                                Dim byteCount As Integer
                                Do
                                    byteCount = stream.Read(buffer, 0, BUFFER_SIZE)
                                    tempStream.Write(buffer, 0, byteCount)
                                Loop While byteCount = BUFFER_SIZE

                                tempStream.Seek(0, IO.SeekOrigin.Begin)
                                Image = New Icon(tempStream).ToBitmap
                            End Using
                        Catch ex As Exception
                        End Try
                    End Using
                End If
                If Image IsNot Nothing Then
                    FavIconData.tsi.Image = Image
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub tsbi00Productions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbi00Productions.Click
        System.Diagnostics.Process.Start("http://i00Productions.org")
    End Sub

    Private Sub tsbVBForums_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbVBForums.Click
        System.Diagnostics.Process.Start("http://www.vbforums.com/showthread.php?p=4075093")
    End Sub

    Private Sub tsbCodeProject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbCodeProject.Click
        System.Diagnostics.Process.Start("http://www.codeproject.com/KB/edit/i00VBNETSpellCheck.aspx")
    End Sub

    Private Sub tsbDonate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsbDonate.Click
        System.Diagnostics.Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=6C7Y9ECSD2YPA")
    End Sub

    Private Sub tsbAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbAbout.Click
        Using frmAbout As New i00SpellCheck.AboutScreen
            frmAbout.ShowInTaskbar = False
            frmAbout.StartPosition = FormStartPosition.CenterParent
            frmAbout.ShowDialog(Me)
        End Using
    End Sub

    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Using frmSplash As New frmSplash
            If frmSplash.ShowDialog(Me) = False Then
                Close()
            End If
        End Using
    End Sub

    Private Sub prop_PropertyValueChanged(ByVal s As Object, ByVal e As System.Windows.Forms.PropertyValueChangedEventArgs) Handles propRichTextBox.PropertyValueChanged, propTextBox.PropertyValueChanged
        'If e.ChangedItem.Parent IsNot Nothing AndAlso TypeOf e.ChangedItem.Parent.Value Is TextBoxBase Then
        '    Dim UnSupportedDynamicPropertyNames As String() = {}
        '    If UnSupportedDynamicPropertyNames.Contains(e.ChangedItem.PropertyDescriptor.Name) Then
        '        MsgBox("i00 .Net Spell Check does not (currently) support DYNAMIC changing of the following properties:" & vbCrLf & vbCrLf & Join((From xItem In UnSupportedDynamicPropertyNames Select xItem Order By xItem).ToArray, ", ") & vbCrLf & vbCrLf & "These properties can however be set prior to initializing the spellcheck.", MsgBoxStyle.Exclamation)
        '    End If
        'End If
    End Sub

End Class
