Public Class clsGrid
    Inherits i00SpellCheck.BufferedPanel

    Public Class GridSetProperty
        Public Color As Color
        Public SetName As String
        Public TimeSpanMS As Long = 300000 '5 mins

        Public Sub New(ByVal SetName As String, ByVal Color As Color)
            Me.SetName = SetName
            Me.Color = Color
        End Sub

        Public DashStyle As System.Drawing.Drawing2D.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid

        Public PenWidth As Single = 2

        Public Sub New()

        End Sub

        Public MaxValue As Single = 100

        Public AutoMax As Boolean = True

        Public GridValues As New List(Of GridValueItem)
        Public Sub AddGridValue(ByVal GridValueItem As GridValueItem)
            GridValues.Add(GridValueItem)
        End Sub

        Public Enum DataStyles
            Line
            Bar
            FilledLine
        End Enum

        Public DataStyle As DataStyles = DataStyles.Line

        Public Sub DrawPath(ByVal rect As Rectangle, ByVal g As Graphics)
            Try
                Dim Ps As New List(Of System.Drawing.PointF)
                Dim FirstDate = (From xItem In GridValues Order By xItem.Time Descending Where Now.Subtract(xItem.Time).TotalMilliseconds > TimeSpanMS).FirstOrDefault  'first date before our timespan...
                Dim GraphValues As List(Of GridValueItem)

                If FirstDate Is Nothing Then
                    'we have no entries older than the timespan ... use all items..
                    GraphValues = (From xItem In GridValues Order By xItem.Time).ToList
                Else
                    'only select entries newer or equal to the firstDate
                    GraphValues = (From xItem In GridValues Where xItem.Time >= FirstDate.Time Order By xItem.Time).ToList
                End If



                If GraphValues.Count <= 1 Then
                    'can't draw one point
                Else
                    If AutoMax = True Then
                        'set the maxvalue to the highest value recorded...
                        'but only if it is over the current max value...
                        Dim Max = GraphValues.Max(Function(x As GridValueItem) x.Value)
                        If Max > MaxValue Then
                            MaxValue = Max
                        End If
                    End If

                    For Each item In GraphValues
                        'find out the x location
                        Dim x As Single = (1 - (CSng(Now.Subtract(item.Time).TotalMilliseconds) / TimeSpanMS)) * rect.Width
                        Dim y As Single = (1 - (item.Value / MaxValue)) * (rect.Height - 1)
                        If DataStyle = DataStyles.Bar Then
                            If Ps.Count > 0 Then
                                Ps.Add(New System.Drawing.PointF(x, Ps.Last.Y))
                            End If
                            Ps.Add(New System.Drawing.PointF(x, y))
                        Else
                            Ps.Add(New System.Drawing.PointF(x, y))
                        End If
                    Next
                    Using gp As New System.Drawing.Drawing2D.GraphicsPath
                        If DataStyle = DataStyles.Bar Then
                            gp.AddLines(Ps.ToArray)
                        Else
                            gp.AddCurve(Ps.ToArray)
                        End If
                        Select Case DataStyle
                            Case DataStyles.Line
                                Using p As New Pen(Me.Color, PenWidth)
                                    p.DashStyle = DashStyle
                                    g.DrawPath(p, gp)
                                End Using
                            Case DataStyles.FilledLine, DataStyles.Bar
                                'finish the path
                                gp.AddLine(gp.GetLastPoint, New PointF(gp.GetLastPoint.X, rect.Bottom))
                                gp.AddLine(gp.GetLastPoint, New PointF(gp.PathPoints.First.X, rect.Bottom))
                                gp.CloseFigure()
                                Using sb As New SolidBrush(Color.FromArgb(127, Me.Color))
                                    g.FillPath(sb, gp)
                                End Using
                        End Select
                    End Using
                End If
             
            Catch ex As Exception

            End Try
        End Sub

        Public Class GridValueItem
            Public Time As DateTime
            Public Value As Single

            Public Sub New(ByVal Value As Single, Optional ByVal Time As DateTime = Nothing)
                If Time = New DateTime Then Time = Now

                Me.Value = Value
                Me.Time = Time
            End Sub
        End Class
    End Class

    Private mc_GridSets As New List(Of GridSetProperty)
    Public ReadOnly Property GridSets() As List(Of GridSetProperty)
        Get
            Return mc_GridSets
        End Get
    End Property

    Private Sub clsGrid_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

        Using b As New Bitmap(12, 12)
            Using g = Graphics.FromImage(b)
                Using p As New Pen(Me.ForeColor)
                    g.DrawLine(p, 0, 0, b.Width, 0)
                    g.DrawLine(p, 0, 0, 0, b.Height)
                End Using
            End Using
            Using tb As New TextureBrush(b)
                e.Graphics.FillRectangle(tb, e.ClipRectangle)
            End Using
        End Using

        For Each item In mc_GridSets
            item.DrawPath(e.ClipRectangle, e.Graphics)
        Next
    End Sub
End Class
