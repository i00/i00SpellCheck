Imports i00SpellCheck

Public Class frmPerformanceMonitor

#Region "Grid Setup"

    Private Sub frmPerformanceMonitor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            tsiOpen.Image = IconExtraction.GetDefaultIcon(IO.Path.Combine(Environment.SystemDirectory, "perfmon.exe"), IconExtraction.IconSize.SmallIcon).ToBitmap
            tsiOpen.Visible = True
        Catch ex As Exception

        End Try

        tsiGridBGColor.SelectedColor = Color.Black
        ClsGrid1.BackColor = tsiGridBGColor.SelectedColor

        tsiGridColor.SelectedColor = Color.Green
        ClsGrid1.ForeColor = tsiGridColor.SelectedColor

        AddPerformanceCounter("Process", "% Processor Time", IO.Path.GetFileNameWithoutExtension(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName), Color.Lime, 60000).GridData.DataStyle = clsGrid.GridSetProperty.DataStyles.FilledLine

        AddPerformanceCounter(i00SpellCheck.DictionaryPerformanceCounter.CatName, i00SpellCheck.DictionaryPerformanceCounter.WordCheckCounterName, "", Color.Red, 60000)
        AddPerformanceCounter(i00SpellCheck.DictionaryPerformanceCounter.CatName, i00SpellCheck.DictionaryPerformanceCounter.SugguestionLookupCounterName, "", Color.Aqua, 60000)

        If Me.Owner IsNot Nothing Then
            Me.StartPosition = FormStartPosition.Manual
            Me.Location = New Point(Me.Owner.Right - Me.Width, Me.Owner.Bottom - Me.Height)
        End If
    End Sub

    Private Sub tsiGridBGColor_ColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsiGridBGColor.ColorChanged
        ClsGrid1.BackColor = tsiGridBGColor.SelectedColor
    End Sub

    Private Sub tsiGridColor_ColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsiGridColor.ColorChanged
        ClsGrid1.ForeColor = tsiGridColor.SelectedColor
    End Sub

#End Region

#Region "Open Perfmon"

    <System.Runtime.InteropServices.DllImport("User32.dll")> _
    Private Shared Function SetForegroundWindow(ByVal handle As IntPtr) As Boolean
    End Function
    <System.Runtime.InteropServices.DllImport("User32.dll")> _
    Private Shared Function ShowWindow(ByVal handle As IntPtr, ByVal nCmdShow As Integer) As Boolean
    End Function
    Private Const SW_RESTORE As Integer = 9
    <System.Runtime.InteropServices.DllImport("User32.dll")> _
    Private Shared Function IsIconic(ByVal handle As IntPtr) As Boolean
    End Function

    Dim PerfmonProcess As Process
    Private Sub tsiOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsiOpen.Click
        Dim ShellFile = IO.Path.Combine(Environment.SystemDirectory, "perfmon.exe")
        'qwertyuiop - wanted to bring the existing process to the foreground ... but it just opens mmc so it does not stay open :(
        If PerfmonProcess Is Nothing OrElse PerfmonProcess.HasExited Then
            If FileIO.FileSystem.FileExists(ShellFile) Then
                PerfmonProcess = Process.Start(ShellFile)
            End If
        Else
            Dim hdl = PerfmonProcess.MainWindowHandle
            If IsIconic(hdl) Then
                ShowWindow(hdl, SW_RESTORE)
            End If
            SetForegroundWindow(hdl)
        End If
    End Sub

#End Region

#Region "Performance counters"

    Dim tListCounters As System.Threading.Thread

    Private Sub tsiAdd_DropDownOpening(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsiAdd.DropDownOpening
        If tsiAdd.DropDownItems.Count = 0 Then
            ListCounters()
        End If
    End Sub

    Dim tsiLoading As New ToolStripMenuItem("Loading...") With {.Enabled = False}
    Private Sub ListCounters()
        tsiAdd.DropDownItems.Clear()
        tsiAdd.DropDownItems.Add(tsiLoading)

        If tListCounters IsNot Nothing AndAlso tListCounters.IsAlive Then
            tListCounters.Abort()
        End If
        tListCounters = New System.Threading.Thread(AddressOf LoadPerformanceCounters)
        tListCounters.IsBackground = True
        tListCounters.Start()
    End Sub

    Private Class PerformanceListItem
        Public Name As String
        Public Counters As New List(Of CounterListItem)
        Public Class CounterListItem
            Public Name As String
            Public Counters As New List(Of String)
        End Class
    End Class

    Private Sub LoadPerformanceCounters()
        Dim PerformanceListItems As New List(Of PerformanceListItem)
        For Each cat In (From xItem In PerformanceCounterCategory.GetCategories() Order By xItem.CategoryName).ToArray
            Dim PerformanceListItem As New PerformanceListItem With {.Name = cat.CategoryName}
            If cat.CategoryType = PerformanceCounterCategoryType.SingleInstance Then
                For Each counter In (From xItem In cat.GetCounters Order By xItem.CounterName)
                    PerformanceListItem.Counters.Add(New PerformanceListItem.CounterListItem() With {.Name = counter.CounterName})
                Next
            ElseIf cat.CategoryType = PerformanceCounterCategoryType.MultiInstance Then
                For Each instance In (From xItem In cat.GetInstanceNames Order By xItem)
                    Dim cli = New PerformanceListItem.CounterListItem() With {.Name = instance}
                    Try
                        cli.Counters.AddRange((From xItem In cat.GetCounters(instance) Order By xItem.CounterName Select xItem.CounterName).ToArray)
                    Catch ex As Exception

                    End Try
                    PerformanceListItem.Counters.Add(cli)
                Next
            End If

            PerformanceListItems.Add(PerformanceListItem)
        Next
        PerformanceCountersLoaded(PerformanceListItems)
    End Sub

    Private Class tsiCounterAdd
        Inherits ToolStripMenuItem
        Public Category As String
        Public Counter As String
        Public Instance As String
        Public Sub New(ByVal Category As String, ByVal Counter As String, Optional ByVal Instance As String = "")
            Me.Category = Category
            Me.Counter = Counter
            Me.Text = Counter
            Me.Instance = Instance
        End Sub
    End Class

    Private Delegate Sub PerformanceCountersLoaded_cb(ByVal PerformanceListItems As List(Of PerformanceListItem))
    Private Sub PerformanceCountersLoaded(ByVal PerformanceListItems As List(Of PerformanceListItem))
        If Me.InvokeRequired Then
            Dim PerformanceCountersLoaded_cb As New PerformanceCountersLoaded_cb(AddressOf PerformanceCountersLoaded)
            Me.Invoke(PerformanceCountersLoaded_cb, PerformanceListItems)
        Else
            For Each cat In PerformanceListItems
                Dim tsi As New ToolStripMenuItem(cat.Name)
                For Each counter In cat.Counters
                    If counter.Counters.Count = 0 Then
                        Dim tsiSub As New tsiCounterAdd(cat.Name, counter.Name)
                        AddHandler tsiSub.Click, AddressOf AddToolStripMenuItem_Click
                        tsi.DropDownItems.Add(tsiSub)
                    Else
                        Dim tsiSub As New ToolStripMenuItem(counter.Name)
                        For Each item In counter.Counters
                            Dim tsiSub2 As New tsiCounterAdd(cat.Name, item, counter.Name)
                            AddHandler tsiSub2.Click, AddressOf AddToolStripMenuItem_Click
                            tsiSub.DropDownItems.Add(tsiSub2)
                        Next
                        tsi.DropDownItems.Add(tsiSub)

                    End If
                Next
                tsiAdd.DropDownItems.Add(tsi)
            Next
            tsiAdd.DropDownItems.Remove(tsiLoading)
        End If
    End Sub

    Private Sub AddToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim tsiCounterAdd = TryCast(sender, tsiCounterAdd)
        If tsiCounterAdd IsNot Nothing Then
            AddPerformanceCounter(tsiCounterAdd.Category, tsiCounterAdd.Counter, tsiCounterAdd.Instance, 60000)
        End If
    End Sub

    Private Class tsiCounter
        Inherits ToolStripDropDownButton

#Region "Shared controls"

#Region "Common"

#Region "Line Style tsi"

        Public Class ctlTsiLineStyle
            Inherits ToolStripDropDownItem

            Public LineWidth As Single
            Public DashStyle As Drawing2D.DashStyle = Drawing.Drawing2D.DashStyle.Solid
            Public Color As Color

            Public Sub New()
                MyBase.AutoToolTip = False
                MyBase.AutoSize = False
                MyBase.Width = 100
            End Sub

            Public Checked As Boolean

            Private Sub ctlTsiLineStyle_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
                e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

                If Me.Selected OrElse Checked Then
                    Using tsi As New ToolStripButton
                        tsi.AutoSize = False
                        tsi.Width = Me.ContentRectangle.Width
                        tsi.Height = Me.ContentRectangle.Height
                        tsi.Select()
                        Using b As New Bitmap(tsi.Width, tsi.Height)
                            Using g = Graphics.FromImage(b)
                                Me.Parent.Renderer.DrawButtonBackground(New ToolStripItemRenderEventArgs(g, tsi))
                                If Checked = False Then
                                    'alpha image
                                    b.Filters.Alpha()
                                Else
                                    'just draw
                                End If
                                e.Graphics.DrawImageUnscaled(b, Me.ContentRectangle.Location)
                            End Using
                        End Using

                        'Me.Parent.Renderer.DrawButtonBackground(New ToolStripItemRenderEventArgs(e.Graphics, tsi))

                    End Using
                End If

                Using p As New Pen(Color, LineWidth)
                    p.DashStyle = DashStyle
                    Dim y As Single = CSng(Me.ContentRectangle.X + ((Me.ContentRectangle.Height - 1) / 2))
                    e.Graphics.DrawLine(p, Me.ContentRectangle.X, y, Me.ContentRectangle.Right - 5, y)
                End Using
            End Sub
        End Class

#End Region

#Region "Setup"

        Public Shared tsiCounterOwner As tsiCounter

        Private Sub tsiCounter_DropDownOpening(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DropDownOpening
            tsiCounterOwner = Me

            Me.DropDownItems.Add(tsiHeader)
            tsiHeader.Text = Replace(Me.Text, vbCrLf, " - ")

            Me.DropDownItems.Add(tsiStatus)
            UpdateTsiStatus(True)

            Me.DropDownItems.Add(tsiSeparator)

            Me.DropDownItems.Add(tsiColor)
            tsiColor.SelectedColor = Me.ItemColor

            Me.DropDownItems.Add(tsiLineWidth)
            For Each item In tsiLineWidth.DropDownItems.OfType(Of ctlTsiLineStyle)()
                item.Checked = item.LineWidth = Me.GridData.PenWidth
                item.Color = tsiCounterOwner.GridData.Color
                item.DashStyle = tsiCounterOwner.GridData.DashStyle
            Next

            Me.DropDownItems.Add(tsiLineStyle)
            For Each item In tsiLineStyle.DropDownItems.OfType(Of ctlTsiLineStyle)()
                item.Checked = item.DashStyle = Me.GridData.DashStyle
                item.Color = tsiCounterOwner.GridData.Color
                item.LineWidth = tsiCounterOwner.GridData.PenWidth
            Next

            Me.DropDownItems.Add(tsiUpdateInterval)
            For Each item In tsiUpdateInterval.DropDownItems.OfType(Of ToolStripMenuItem)()
                item.Checked = Val(item.Tag) = Me.GridData.TimeSpanMS
            Next

            Me.DropDownItems.Add(tsiRemove)

        End Sub

#End Region

#End Region

#Region "Header"

        Private Shared ReadOnly Property tsiHeader() As i00SpellCheck.MenuTextSeperator
            Get
                Static mc_tsiHeader As i00SpellCheck.MenuTextSeperator
                If mc_tsiHeader Is Nothing OrElse mc_tsiHeader.IsDisposed Then
                    mc_tsiHeader = New i00SpellCheck.MenuTextSeperator
                End If

                Return mc_tsiHeader
            End Get
        End Property

#End Region

#Region "Status"

        Public Shared ReadOnly Property tsiStatus() As i00SpellCheck.HTMLMenuItem
            Get
                Static mc_tsiStatus As i00SpellCheck.HTMLMenuItem
                If mc_tsiStatus Is Nothing OrElse mc_tsiStatus.IsDisposed Then
                    mc_tsiStatus = New i00SpellCheck.HTMLMenuItem(" ")
                    mc_tsiStatus.SmallHeight = 1000
                End If

                Return mc_tsiStatus
            End Get
        End Property


        Public Shared Sub UpdateTsiStatus(Optional ByVal ForceUpdate As Boolean = False)
            If tsiCounterOwner IsNot Nothing AndAlso (tsiStatus.Visible OrElse ForceUpdate) Then
                If tsiCounterOwner.GridData.GridValues.Count >= 1 Then
                    Dim Average As String = tsiCounterOwner.GridData.GridValues.Average(Function(x As clsGrid.GridSetProperty.GridValueItem) x.Value).ToString
                    Dim AverageExcluding0 As String
                    Dim ItemsNot0 = (From xItem In tsiCounterOwner.GridData.GridValues Where xItem.Value <> 0 Select xItem.Value)
                    If ItemsNot0.Count > 0 Then
                        AverageExcluding0 = ItemsNot0.Average.ToString
                    Else
                        Dim ShortcutKeyColor = System.Drawing.ColorTranslator.ToHtml(i00SpellCheck.DrawingFunctions.BlendColor(Color.FromKnownColor(KnownColor.ControlText), Color.FromKnownColor(KnownColor.Control)))
                        AverageExcluding0 = "<i><font color=" & ShortcutKeyColor & ">All recorded data is 0</font></i>"
                    End If
                    Dim WatchingSince As String = (From xItem In tsiCounterOwner.GridData.GridValues Order By xItem.Time Select (xItem.Time.ToString)).FirstOrDefault
                    Dim RecordValue As String = tsiCounterOwner.GridData.GridValues.Max(Function(x As clsGrid.GridSetProperty.GridValueItem) x.Value).ToString
                    tsiStatus.HTMLText = "<b>Current Value:</b> " & tsiCounterOwner.GridData.GridValues.Last.Value & vbCrLf & _
                                         "<b>Highest Recorded Value:</b> " & RecordValue & vbCrLf & _
                                         "<b>Graph Max:</b> " & tsiCounterOwner.GridData.MaxValue & vbCrLf & _
                                         "<b>Average:</b> " & Average & vbCrLf & _
                                         "<b>Average Excluding 0 Values:</b> " & AverageExcluding0 & vbCrLf & _
                                         "<b>Since:</b> " & WatchingSince & vbCrLf & _
                                         "<b>Total Data Recorded:</b> " & tsiCounterOwner.GridData.GridValues.Count
                    tsiStatus.Invalidate()
                Else
                    Dim ShortcutKeyColor = System.Drawing.ColorTranslator.ToHtml(i00SpellCheck.DrawingFunctions.BlendColor(Color.FromKnownColor(KnownColor.ControlText), Color.FromKnownColor(KnownColor.Control)))
                    tsiStatus.HTMLText = "<i><font color=" & ShortcutKeyColor & ">No Data</font></i>"
                End If
            End If
        End Sub

#End Region

#Region "Separator"

        Private Shared ReadOnly Property tsiSeparator() As ToolStripSeparator
            Get
                Static mc_tsiSeparator As ToolStripSeparator
                If mc_tsiSeparator Is Nothing OrElse mc_tsiSeparator.IsDisposed Then
                    mc_tsiSeparator = New ToolStripSeparator
                End If

                Return mc_tsiSeparator
            End Get
        End Property

#End Region

#Region "Color"

        Private Shared WithEvents mc_tsiColor As i00SpellCheck.tsiColorPicker
        Private Shared ReadOnly Property tsiColor() As i00SpellCheck.tsiColorPicker
            Get
                If mc_tsiColor Is Nothing OrElse mc_tsiColor.IsDisposed Then
                    mc_tsiColor = New i00SpellCheck.tsiColorPicker With {.Persistent = True}
                End If

                Return mc_tsiColor
            End Get
        End Property

        Private Shared Sub tsiColor_ColorChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_tsiColor.ColorChanged
            tsiCounterOwner.ItemColor = tsiColor.SelectedColor
        End Sub

#End Region

#Region "Line Width"

        Private Shared Sub tsiLineWidth_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim tsi = TryCast(sender, ctlTsiLineStyle)
            If tsi IsNot Nothing Then
                tsiCounterOwner.GridData.PenWidth = tsi.LineWidth
            End If
        End Sub

        Private Shared ReadOnly Property tsiLineWidth() As ToolStripMenuItem
            Get
                Static mc_tsiLineWidth As ToolStripMenuItem
                If mc_tsiLineWidth Is Nothing OrElse mc_tsiLineWidth.IsDisposed Then
                    mc_tsiLineWidth = New ToolStripMenuItem("Line Width")
                End If

                If mc_tsiLineWidth.DropDownItems.Count = 0 Then
                    For i = 1 To 4
                        Dim tsi As New ctlTsiLineStyle() With {.LineWidth = i}
                        AddHandler tsi.Click, AddressOf tsiLineWidth_Click
                        mc_tsiLineWidth.DropDownItems.Add(tsi)
                    Next
                End If
                Return mc_tsiLineWidth
            End Get
        End Property

#End Region

#Region "Line Draw Style"

        Private Shared Sub tsiLineStyle_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim tsi = TryCast(sender, ctlTsiLineStyle)
            If tsi IsNot Nothing Then
                tsiCounterOwner.GridData.DashStyle = tsi.DashStyle
            End If
        End Sub

        Private Shared ReadOnly Property tsiLineStyle() As ToolStripMenuItem
            Get
                Static mc_LineStyle As ToolStripMenuItem
                If mc_LineStyle Is Nothing OrElse mc_LineStyle.IsDisposed Then
                    mc_LineStyle = New ToolStripMenuItem("Line Style")
                End If

                If mc_LineStyle.DropDownItems.Count = 0 Then
                    For Each item In [Enum].GetValues(GetType(Drawing2D.DashStyle))
                        If CInt(item) <> Drawing2D.DashStyle.Custom Then
                            Dim tsi As New ctlTsiLineStyle()
                            tsi.DashStyle = DirectCast(item, Drawing2D.DashStyle)
                            AddHandler tsi.Click, AddressOf tsiLineStyle_Click
                            mc_LineStyle.DropDownItems.Add(tsi)
                        End If
                    Next
                End If
                Return mc_LineStyle
            End Get
        End Property

#End Region

#Region "Update Interval"

        Private Shared Sub tsiUpdateInterval_Click(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim tsi = TryCast(sender, ToolStripItem)
            If tsi IsNot Nothing Then
                tsiCounterOwner.GridData.TimeSpanMS = CLng(tsi.Tag)
            End If
        End Sub

        Private Shared ReadOnly Property tsiUpdateInterval() As ToolStripMenuItem
            Get
                Static mc_tsiUpdateInterval As ToolStripMenuItem
                If mc_tsiUpdateInterval Is Nothing OrElse mc_tsiUpdateInterval.IsDisposed Then
                    mc_tsiUpdateInterval = New ToolStripMenuItem("Time Range")
                End If

                If mc_tsiUpdateInterval.DropDownItems.Count = 0 Then
                    mc_tsiUpdateInterval.DropDownItems.Add("1 minute", Nothing, AddressOf tsiUpdateInterval_Click).Tag = 60000
                    mc_tsiUpdateInterval.DropDownItems.Add("2 minutes", Nothing, AddressOf tsiUpdateInterval_Click).Tag = 120000
                    mc_tsiUpdateInterval.DropDownItems.Add("5 minutes", Nothing, AddressOf tsiUpdateInterval_Click).Tag = 300000
                    mc_tsiUpdateInterval.DropDownItems.Add("10 minutes", Nothing, AddressOf tsiUpdateInterval_Click).Tag = 600000
                End If
                Return mc_tsiUpdateInterval
            End Get
        End Property

#End Region

#Region "Remove"

        Private Shared WithEvents mc_tsiRemove As ToolStripMenuItem
        Private Shared ReadOnly Property tsiRemove() As ToolStripMenuItem
            Get
                If mc_tsiRemove Is Nothing OrElse mc_tsiRemove.IsDisposed Then
                    mc_tsiRemove = New ToolStripMenuItem("Remove counter", My.Resources.Delete)
                End If

                Return mc_tsiRemove
            End Get
        End Property

        Private Shared Sub tsiRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_tsiRemove.Click
            'qwertyuiop - had to show before remove otherwise would get an error about the toolstrip items being read only!
            tsiCounterOwner.Overflow = ToolStripItemOverflow.Never
            tsiCounterOwner.Parent.Items.Remove(tsiCounterOwner)
        End Sub

#End Region

#End Region

        Dim mc_ItemColor As Color
        Public Property ItemColor() As Color
            Get
                Return mc_ItemColor
            End Get
            Set(ByVal value As Color)
                mc_ItemColor = value
                GridData.Color = value
                Me.Invalidate()
            End Set
        End Property

        Public Performance As String
        Public Counter As String
        Public Instance As String

        Dim mc_inError As Boolean
        Private Property inError() As Boolean
            Get
                Return mc_inError
            End Get
            Set(ByVal value As Boolean)
                If mc_inError <> value Then
                    mc_inError = value
                    Me.Invalidate()
                End If
            End Set
        End Property

        Public Sub CommitValue()
            Try
                Dim ThisValue As Single = PerformanceCounter.NextValue
                GridData.GridValues.Add(New clsGrid.GridSetProperty.GridValueItem(ThisValue))
                inError = False
            Catch ex As Exception
                'performance counter error...
                inError = True
            End Try
        End Sub

        Public ReadOnly Property PerformanceCounter() As PerformanceCounter
            Get
                Static mc_PerformanceCounter As PerformanceCounter
                If mc_PerformanceCounter Is Nothing Then
                    mc_PerformanceCounter = New PerformanceCounter(Performance, Counter, Instance, True)
                End If
                Return mc_PerformanceCounter
            End Get
        End Property

        Public Sub New(ByVal Performance As String, ByVal Counter As String, ByVal Instance As String, ByVal Color As Color, ByVal Interval As Long)
            Me.ItemColor = Color
            NewInternal(Performance, Counter, Instance, Interval)
        End Sub

        Public Sub New(ByVal Performance As String, ByVal Counter As String, ByVal Instance As String, ByVal Interval As Long)
            NewInternal(Performance, Counter, Instance, Interval)
        End Sub

        Private Shared ReadOnly Property BlankImage() As Bitmap
            Get
                Static mc_BlankImage As New Bitmap(16, 16)
                Return mc_BlankImage
            End Get
        End Property

        Private Sub NewInternal(ByVal Performance As String, ByVal Counter As String, ByVal Instance As String, ByVal Interval As Long)
            Me.Performance = Performance
            Me.Counter = Counter
            Me.Instance = Instance

            Me.Image = BlankImage
            Me.DisplayStyle = ToolStripItemDisplayStyle.Image
            Me.Text = Performance & vbCrLf & If(Instance <> "", Instance & vbCrLf, "") & Counter

            If Me.ItemColor.IsEmpty Then
                Randomize()
                Me.ItemColor = (From xItem In tsiColor.Colors Order By Rnd()).First
            End If

            GridData.Color = Me.ItemColor
            GridData.TimeSpanMS = Interval
        End Sub

        Public GridData As New clsGrid.GridSetProperty

        Private Sub tsiCounter_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
            Dim rect = Me.ContentRectangle

            'we only want the rect 13 x 13
            Dim RectSizeDiff = New Size(rect.Width - 13, rect.Height - 13)
            rect.X += CInt(RectSizeDiff.Width / 2)
            rect.Width = 13
            rect.Y += CInt(RectSizeDiff.Height / 2)
            rect.Height = 13

            Using sb As New SolidBrush(mc_ItemColor)
                e.Graphics.FillRectangle(sb, rect)
            End Using
            'draw border
            Using p As New Pen(i00SpellCheck.DrawingFunctions.AlphaColor(Color.FromKnownColor(KnownColor.MenuText), 63))
                e.Graphics.DrawRectangle(p, rect.X, rect.Y, rect.Width, rect.Height)
            End Using

            If mc_inError Then
                rect = Me.ContentRectangle
                e.Graphics.InterpolationMode = Drawing2D.InterpolationMode.High
                rect.Width = CInt(rect.Width / 2)
                rect.Height = CInt(rect.Height / 2)
                rect.X += rect.Width
                rect.Y += rect.Height
                e.Graphics.DrawImage(My.Resources.Delete, rect)
            End If
        End Sub
    End Class

    Private Function AddPerformanceCounter(ByVal Performance As String, ByVal Counter As String, ByVal Instance As String, ByVal interval As Long) As tsiCounter
        AddPerformanceCounter = New tsiCounter(Performance, Counter, Instance, interval)
        ToolStrip1.Items.Add(AddPerformanceCounter)
        ClsGrid1.GridSets.Add(AddPerformanceCounter.GridData)
    End Function

    Private Function AddPerformanceCounter(ByVal Performance As String, ByVal Counter As String, ByVal Instance As String, ByVal Color As Color, ByVal interval As Long) As tsiCounter
        AddPerformanceCounter = New tsiCounter(Performance, Counter, Instance, Color, interval)
        ToolStrip1.Items.Add(AddPerformanceCounter)
        ClsGrid1.GridSets.Add(AddPerformanceCounter.GridData)
    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Static lastTime As Integer
        Dim ThisTime = CInt(Now.TimeOfDay.TotalSeconds)
        If lastTime <> ThisTime Then
            For Each item In ToolStrip1.Items.OfType(Of tsiCounter)().ToArray
                item.CommitValue()
            Next
            lastTime = ThisTime

            'update the status menu item
            tsiCounter.UpdateTsiStatus()
        End If

        ClsGrid1.Invalidate()
    End Sub

#End Region

End Class