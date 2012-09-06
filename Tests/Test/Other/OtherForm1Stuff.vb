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
                        If tsbExamples.GetCurrentParent.Focused() Then
                            If LastControl IsNot Nothing Then
                                Try
                                    LastControl.Focus()
                                Catch ex As Exception

                                End Try
                            End If
                        Else
                            LastControl = Me.ActiveControl
                            tsbExamples.GetCurrentParent.Focus()
                            tsbExamples.Select()
                        End If
                    Case Else
                        MyBase.WndProc(m)
                End Select
            Case Else
                MyBase.WndProc(m)
        End Select

    End Sub

    Private Function ExtractIcon(ByVal file As String, ByVal Large As Boolean) As Bitmap
        Using icon As Icon = icon.ExtractAssociatedIcon(file)
            If Large Then
                Return icon.ToBitmap
            Else
                ExtractIcon = New Bitmap(16, 16)
                Using g = Graphics.FromImage(ExtractIcon)
                    g.InterpolationMode = Drawing2D.InterpolationMode.High
                    g.DrawIcon(icon, New Rectangle(0, 0, ExtractIcon.Width, ExtractIcon.Height))
                End Using
            End If
        End Using
    End Function

    Private Sub Form1_Load_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim SpellCheckAssembly = System.Reflection.Assembly.Load("i00SpellCheck")
        If SpellCheckAssembly IsNot Nothing Then
            'qwertyuiop - should use something like: ... but it doesn't work for UNC's
            '... = System.Drawing.Icon.ExtractAssociatedIcon(Assembly)
            tsbAbout.ToolTipText = "About " & SpellCheckAssembly.GetName.Name & " " & SpellCheckAssembly.GetName.Version.ToString

            tsbAbout.Image = ExtractIcon(SpellCheckAssembly.Location, False)
            tsbSpellCheck.Image = tsbAbout.Image
            Me.Icon = Icon.ExtractAssociatedIcon(SpellCheckAssembly.Location)
        End If

        Dim ToolBoxIcon As New ToolboxBitmapAttribute(GetType(PropertyGrid))
        tsbProperties.Image = ToolBoxIcon.GetImage(GetType(PropertyGrid), False)



        Dim URLIcon = IconExtraction.GetDefaultIcon(".url", IconExtraction.IconSize.SmallIcon).ToBitmap
        tsbi00Productions.Image = URLIcon
        tsbVBForums.Image = URLIcon
        tsbCodeProject.Image = URLIcon

        LoadFavicon("http://www.paypal.com/favicon.ico", tsbDonate)
        LoadFavicon("http://vbforums.com/favicon.ico", tsbVBForums)
        LoadFavicon("http://i00Productions.org/favicon.ico", tsbi00Productions)
        LoadFavicon("http://codeproject.com/favicon.ico", tsbCodeProject)

        For Each item In ShowIgnoredToolStripMenuItem.DropDownItems.OfType(Of ToolStripItem)()
            AddHandler item.Click, AddressOf ShowIgnoredToolStripMenuItem_Click
        Next

        tsiDrawStyle.SelectedIndex = 0

        'add any extra plugins as extra tabs...
        Dim SpellCheckControlBases = i00SpellCheck.PluginManager(Of i00SpellCheck.SpellCheckControlBase).GetPlugins
        Dim SpellControls As New List(Of Control) ' LoadSpellCheckPlugins.Controls()
        For Each item In SpellCheckControlBases
            Try
                If item.ControlType.IsAbstract Then
                    'cannot create instance for eg TextBoxBase... lets go through everything and try to create a control that comes from this
                    SpellControls.AddRange(i00SpellCheck.PluginManager(Of Control).GetAllPluginsInReferencedAssemblies(item.GetType.Assembly, item.ControlType))
                Else
                    Dim ctl = TryCast(item.ControlType.Module.Assembly.CreateInstance(item.ControlType.FullName), Control)
                    SpellControls.Add(ctl)
                End If
            Catch ex As Exception

            End Try
        Next

        'remove the duplicate controls...
        SpellControls = (From xItem In (From xItem In SpellControls Group xItem By ControlType = xItem.GetType Into Group) Select xItem.Group.First).ToList

        For Each item In (From xItem In SpellControls Order By xItem.GetType.Name).ToList
            Dim InsertControl = item
            Dim iTestHarness = TryCast(item.SpellCheck, iTestHarness)
            If iTestHarness IsNot Nothing Then
                InsertControl = iTestHarness.SetupControl(item)
                If InsertControl Is Nothing Then Continue For
            End If

            Dim TabPage = New TabPage(item.GetType.Name)
            ToolBoxIcon = New ToolboxBitmapAttribute(item.GetType)
            ilTabSpellControls.Images.Add(item.GetType.FullName, ToolBoxIcon.GetImage(item.GetType, False))
            TabPage.ImageIndex = ilTabSpellControls.Images.IndexOfKey(item.GetType.FullName)
            tabSpellControls.TabPages.Add(TabPage)


            InsertControl.Dock = DockStyle.Fill
            Dim prop As New PropertyGrid
            prop.Width = 250
            prop.SelectedObject = item.SpellCheck
            prop.Dock = DockStyle.Right
            prop.Visible = False
            'If item IsNot Nothing Then
            TabPage.Controls.Add(InsertControl)
            TabPage.Controls.Add(prop)
            'End If
        Next

        If tabSpellControls.TabCount = 0 Then
            Dim TabPage = New TabPage("")
            tabSpellControls.TabPages.Add(TabPage)
            Dim lblNoPlugins As New Label
            lblNoPlugins.Text = "No plugins could be loaded"
            lblNoPlugins.AutoSize = False
            lblNoPlugins.Dock = DockStyle.Fill
            lblNoPlugins.TextAlign = ContentAlignment.MiddleCenter
            TabPage.Controls.Add(lblNoPlugins)
        End If

        UpdateEnabledCheck()
        tabSpellControls_SelectedIndexChanged(tabSpellControls, EventArgs.Empty)
    End Sub

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

    'Private Sub prop_PropertyValueChanged(ByVal s As Object, ByVal e As System.Windows.Forms.PropertyValueChangedEventArgs) Handles propRichTextBox.PropertyValueChanged, propTextBox.PropertyValueChanged
    '    'If e.ChangedItem.Parent IsNot Nothing AndAlso TypeOf e.ChangedItem.Parent.Value Is TextBoxBase Then
    '    '    Dim UnSupportedDynamicPropertyNames As String() = {}
    '    '    If UnSupportedDynamicPropertyNames.Contains(e.ChangedItem.PropertyDescriptor.Name) Then
    '    '        MsgBox("i00 .Net Spell Check does not (currently) support DYNAMIC changing of the following properties:" & vbCrLf & vbCrLf & Join((From xItem In UnSupportedDynamicPropertyNames Select xItem Order By xItem).ToArray, ", ") & vbCrLf & vbCrLf & "These properties can however be set prior to initializing the spellcheck.", MsgBoxStyle.Exclamation)
    '    '    End If
    '    'End If
    'End Sub

End Class
