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
    Implements IMessageFilter

#Region "For Keypress to Make Ignore Underline Appear"

    Const WM_KEYDOWN As Integer = &H100
    Const WM_KEYUP As Integer = &H101

    Public Function PreFilterMessage(ByRef m As System.Windows.Forms.Message) As Boolean Implements System.Windows.Forms.IMessageFilter.PreFilterMessage
        Static ctlPressed As Boolean
        Select Case m.Msg
            Case WM_KEYDOWN
                Select Case m.WParam.ToInt32
                    Case Keys.ControlKey
                        If ctlPressed = False Then
                            ctlPressed = True
                            RepaintTextBox()
                            'parentTextBox.Invalidate()
                        End If
                End Select
            Case WM_KEYUP
                Select Case m.WParam.ToInt32
                    Case Keys.ControlKey
                        ctlPressed = False
                        RepaintTextBox()
                        'parentTextBox.Invalidate()
                End Select
        End Select
    End Function

#End Region

#Region "Dictionary cache object + Thread to add to cache"

    Friend dictCache As New Dictionary(Of String, Dictionary.SpellCheckWordError)

    Private Sub parentTextBox_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_parentTextBox.Disposed
        'kill all running threads
        AbortSpellCheckThreads()
    End Sub

    Private Sub AbortSpellCheckThreads()
        SyncLock AddWordThreads
            For Each item In AddWordThreads
                If item.IsAlive Then
                    item.Abort()
                End If
            Next
        End SyncLock
    End Sub

    Dim AddWordThreads As New List(Of Threading.Thread)
    Private Sub AddWordsToCache(ByVal oDictionary_String_SpellCheckWordError As Object)
        'use DirectCast so that error is thrown if wrong object is passed in
        Dim NewItems = DirectCast(oDictionary_String_SpellCheckWordError, Dictionary(Of String, Dictionary.SpellCheckWordError))
        If parentTextBox.InvokeRequired = False Then
            'open in another thread to the textbox
            Dim mtSpellCheck As New Threading.Thread(AddressOf AddWordsToCache)
            mtSpellCheck.Name = "Spell check"
            mtSpellCheck.IsBackground = True 'make it abort when the app is killed
            SyncLock AddWordThreads
                AddWordThreads.Add(mtSpellCheck)
            End SyncLock
            mtSpellCheck.Start(NewItems)
        Else
            'we are on another thread
            SyncLock dictCache
                For Each item In NewItems
                    'Debug.Print("Spell checking: " & item.Key)
                    If dictCache.ContainsKey(item.Key) = False Then
                        dictCache.Add(item.Key, CurrentDictionary.SpellCheckWord(item.Key))
                    End If
                Next
            End SyncLock
            'all words added
            Dim AddWordThreadsCount As Integer
            SyncLock AddWordThreads
                AddWordThreads.Remove(Threading.Thread.CurrentThread)
                AddWordThreadsCount = AddWordThreads.Count
            End SyncLock
            If AddWordThreadsCount = 0 Then
                RepaintTextBox()
            End If
        End If
    End Sub

    Private Delegate Sub cb_RepaintTextBox()

    Public Sub RepaintTextBox()
        Try
            If parentTextBox.InvokeRequired Then
                If parentTextBox.Visible Then
                    Dim cb As New cb_RepaintTextBox(AddressOf RepaintTextBox)
                    parentTextBox.Invoke(cb)
                End If
            Else
                If Me.Settings.RenderCompatibility Then
                    If Me.Settings.ShowMistakes AndAlso Me.parentTextBox.IsSpellCheckEnabled Then
                        parentTextBox.Invalidate()
                    End If
                Else
                    'only redraw if the overlay hasn't been created yet or if the overlay is visible
                    If DrawOverlayForm Is Nothing Then
                        OpenOverlay()
                    End If
                    If Me.Settings.ShowMistakes = False OrElse Me.parentTextBox.IsSpellCheckEnabled = False Then
                        DrawOverlayForm.Visible = False
                    Else
                        DrawOverlayForm.Visible = mc_parentTextBox.Visible
                    End If
                    If DrawOverlayForm.Visible Then
                        Me.CustomPaint()
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "Drawing"

    Public Sub Invalidate()
        Me.dictCache.Clear()
        If Me.TextBox IsNot Nothing Then
            RepaintTextBox()
            'Me.TextBox.Invalidate()
        End If
    End Sub

    Private textBoxGraphics As Graphics

    Public Class SpellCheckCustomPaintEventArgs
        Inherits EventArgs
        Public Bounds As Rectangle
        Public Graphics As Graphics
        Public DrawDefault As Boolean = True
        Public Word As String
        Public WordState As Dictionary.SpellCheckWordError
    End Class

    Public Event SpellCheckErrorPaint(ByVal sender As Object, ByRef e As SpellCheckCustomPaintEventArgs)

    Private Sub CustomPaint()
        If CurrentDictionary IsNot Nothing AndAlso (CurrentDictionary.Loading = True OrElse CurrentDictionary.Count = 0) Then Exit Sub

        Dim TextHeight As Integer = System.Windows.Forms.TextRenderer.MeasureText("Ag", parentTextBox.Font).Height
        Dim BufferWidth As Integer = System.Windows.Forms.TextRenderer.MeasureText("--", parentTextBox.Font).Width

        'for drawing underlines below the textbox drawing bounds when on a single line text box
        Dim bHeight = parentTextBox.ClientSize.Height
        If DrawOverlayForm IsNot Nothing Then
            bHeight = DrawOverlayForm.Height
        End If

        Using b As New Bitmap(parentTextBox.ClientSize.Width, bHeight)
            Using g = Graphics.FromImage(b)
                g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

                'Using sb As New SolidBrush(Color.FromArgb(127, Color.Blue))
                '    g.FillRectangle(sb, New RectangleF(0, 0, parentTextBox.Width, parentTextBox.Height))
                'End Using

                Dim FromChar = parentTextBox.GetCharIndexFromPosition(New Point(0, 0))
                Dim ToChar = parentTextBox.GetCharIndexFromPosition(New Point(parentTextBox.ClientRectangle.Width - 1, parentTextBox.ClientRectangle.Height - 1))

                Dim theText As String = RemoveWordBreaks(parentTextBox.Text)

                Dim LetterIndex = FromChar
                Dim LeftSide = Left(theText, FromChar)
                LeftSide = LeftSide.Split(" "c).Last
                Dim RightSide = Right(theText, Len(theText) - ToChar)
                RightSide = RightSide.Split(" "c).First
                FromChar -= Len(LeftSide)
                ToChar += Len(RightSide)

                Dim VisibleText = Mid(theText, FromChar + 1, ToChar - FromChar)

                'If parentTextBox.Multiline = False Then
                'g.TranslateTransform(-System.Windows.Forms.TextRenderer.MeasureText(LeftSide, parentTextBox.Font).Width, -5)
                'End If

                Dim NewWords As New Dictionary(Of String, Dictionary.SpellCheckWordError)

                If Trim(VisibleText) <> "" Then
                    Dim words = Replace(Replace(VisibleText, vbCr, " "), vbLf, " ").Split(" "c)

                    For iWord = LBound(words) To UBound(words)
                        If words(iWord) <> "" Then
                            Dim P1 = parentTextBox.GetPositionFromCharIndex(LetterIndex)
                            Dim P1OffsetPlus As Integer = 0
                            If iWord = 0 AndAlso P1.X >= TextBox.ClientSize.Width Then
                                P1OffsetPlus = 1
                                P1.X = 0
                            End If
                            If P1.Y < parentTextBox.Height Then
                                Dim WordState As Dictionary.SpellCheckWordError = Dictionary.SpellCheckWordError.SpellError
                                If dictCache.ContainsKey(words(iWord)) Then
                                    'load from cache
                                    WordState = dictCache(words(iWord))
                                Else
                                    ''item is not in the dict cache
                                    'WordState = SpellCheckWord(words(iWord))
                                    'If dictCache.ContainsKey(words(iWord)) = False Then
                                    '    dictCache.Add(words(iWord), WordState)
                                    'End If
                                    'assume ok for now and add word to be processed
                                    If NewWords.ContainsKey(words(iWord)) = False Then
                                        NewWords.Add(words(iWord), Dictionary.SpellCheckWordError.OK)
                                    End If

                                    WordState = Dictionary.SpellCheckWordError.OK
                                End If
                                If WordState = Dictionary.SpellCheckWordError.OK Then

                                Else
                                    If WordState = Dictionary.SpellCheckWordError.Ignore Then
                                        If Settings.ShowIgnored = SpellCheckSettings.ShowIgnoreState.OnKeyDown AndAlso My.Computer.Keyboard.CtrlKeyDown = False Then
                                            GoTo ContinueFor
                                        ElseIf Settings.ShowIgnored = SpellCheckSettings.ShowIgnoreState.AlwaysHide Then
                                            GoTo ContinueFor
                                        End If
                                    End If

                                    Dim P2 = parentTextBox.GetPositionFromCharIndex(LetterIndex + Len(words(iWord)))
                                    If LeftSide <> "" AndAlso iWord = 0 Then
                                        Dim NormalStringWidth = g.MeasureString(Mid(words(iWord), Len(LeftSide) + 1 + P1OffsetPlus), parentTextBox.Font).Width
                                        Dim XOffsetDiff = g.MeasureString("-" & Mid(words(iWord), Len(LeftSide) + 1 + P1OffsetPlus) & "-", parentTextBox.Font).Width - NormalStringWidth
                                        P2.X = CInt(parentTextBox.GetPositionFromCharIndex(LetterIndex + P1OffsetPlus).X + (NormalStringWidth - XOffsetDiff))
                                    End If
                                    If P2.X = 0 Then
                                        'we are the last char ... :(
                                        P2 = parentTextBox.GetPositionFromCharIndex(LetterIndex + Len(words(iWord)) - 1)
                                        P2.X += System.Windows.Forms.TextRenderer.MeasureText("-" & Right(words(iWord), 1) & "-", parentTextBox.Font).Width - BufferWidth
                                    End If
                                    Dim LineHeight As Integer = GetLineHeightFromCharPosition(LetterIndex)
                                    'P1.Y += LineHeight
                                    P2.Y = P1.Y + LineHeight
                                    'P2.Y = P1.Y

                                    Dim e = New SpellCheckCustomPaintEventArgs With {.Graphics = g, .Word = words(iWord), .Bounds = New Rectangle(P1.X, P1.Y, P2.X - P1.X, P2.Y - P1.Y), .WordState = WordState}
                                    RaiseEvent SpellCheckErrorPaint(Me, e)
                                    If e.DrawDefault Then
                                        Select Case WordState
                                            Case Dictionary.SpellCheckWordError.Ignore
                                                Using p As New Pen(Settings.IgnoreColor)
                                                    g.DrawLine(p, P1.X, P2.Y + 1, P2.X, P2.Y + 1)
                                                End Using
                                            Case Dictionary.SpellCheckWordError.CaseError
                                                DrawWave(g, P1, P2, Settings.CaseMistakeColor)
                                            Case Dictionary.SpellCheckWordError.SpellError
                                                DrawWave(g, P1, P2, Settings.MistakeColor)
                                        End Select
                                    End If
                                End If
                            End If
                        End If
ContinueFor:
                        If LeftSide <> "" AndAlso iWord = 0 Then
                            LetterIndex -= Len(LeftSide)
                        End If
                        LetterIndex += 1 + Len(words(iWord))
                    Next
                End If

Draw:
                If DrawOverlayForm IsNot Nothing Then
                    DrawOverlayForm.SetBitmap(b, 255)
                Else
                    textBoxGraphics.DrawImageUnscaled(b, 0, 0)
                End If

                If NewWords.Count > 0 Then
                    AddWordsToCache(NewWords)
                End If

            End Using
        End Using

    End Sub

    Private Sub DrawWave(ByVal g As Graphics, ByVal startPoint As Point, ByVal endPoint As Point, ByVal Color As Color)
        If endPoint.X - startPoint.X < 4 Then
            endPoint.X = startPoint.X + 4
        End If
        Dim points() As PointF = Nothing  'New List(Of PointF)
        For i = startPoint.X To endPoint.X Step 2
            If points Is Nothing Then
                ReDim Preserve points(0)
            Else
                ReDim Preserve points(UBound(points) + 1)
            End If
            If (i - startPoint.X) Mod 4 = 0 Then
                points(UBound(points)) = New PointF(i, endPoint.Y + 2)
                'points.Add(New PointF(i, startPoint.Y))
            Else
                points(UBound(points)) = New PointF(i, endPoint.Y)
                'points.Add(New PointF(i, startPoint.Y + 2))
            End If
        Next
        Using p As New System.Drawing.Drawing2D.GraphicsPath
            p.AddCurve(points)
            Using pen As Pen = New Pen(Color)
                g.DrawPath(pen, p)
            End Using
        End Using
    End Sub

#End Region

#Region "Overlay"

    Dim DrawOverlayForm As PerPixelAlphaForm

#Region "Create"

#Region "APIs for click through"

    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)> _
    Private Shared Function GetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)> _
    Private Shared Function SetWindowLong(ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    End Function

    Private Const GWL_EXSTYLE As Integer = -20
    Private Const WS_EX_TRANSPARENT As Integer = &H20

#End Region

    Private Sub SetOverlayBounds()
        If DrawOverlayForm IsNot Nothing Then
            'for drawing underlines below the textbox drawing bounds when on a single line text box
            Dim toBeHeight = Me.parentTextBox.ClientSize.Height
            If parentTextBox.Multiline = False Then
                'add wave height
                toBeHeight += 3
            End If
            DrawOverlayForm.Bounds = Me.parentTextBox.RectangleToScreen(New Rectangle(0, 0, Me.parentTextBox.ClientSize.Width, toBeHeight))
        End If
    End Sub

    Private Sub OpenOverlay()
        If DrawOverlayForm Is Nothing Then
            DrawOverlayForm = New PerPixelAlphaForm

            DrawOverlayForm.ShowInTaskbar = False
            DrawOverlayForm.ShowNoFocus(Me.parentTextBox.Parent)

            'make the overlay click-through
            Dim exstyle2 As Integer = GetWindowLong(DrawOverlayForm.Handle, GWL_EXSTYLE)
            exstyle2 = exstyle2 Or WS_EX_TRANSPARENT
            SetWindowLong(DrawOverlayForm.Handle, GWL_EXSTYLE, exstyle2)

            SetOverlayBounds()
            mc_parentTextBox_ParentChanged(Me.parentTextBox, EventArgs.Empty)
        End If
    End Sub

#End Region

#Region "Destory"

    Private Sub CloseOverlay()
        If DrawOverlayForm IsNot Nothing Then
            DrawOverlayForm.Close()
            DrawOverlayForm.Dispose()
            DrawOverlayForm = Nothing
            RemoveAllOverlayHandlers()
        End If
    End Sub

    Private Sub parentTextBox_ForOverlay_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_parentTextBox.Disposed
        CloseOverlay()
    End Sub

#End Region

#Region "Events for overlay"

    Private Sub parentTextBox_ForOverlay_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_parentTextBox.SizeChanged
        If DrawOverlayForm IsNot Nothing Then
            SetOverlayBounds()
            RepaintTextBox()
        End If
    End Sub

    Private Sub mc_parentTextBox_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_parentTextBox.LocationChanged
        SetOverlayBounds()
    End Sub

    Private Sub mc_parentTextBox_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_parentTextBox.VisibleChanged
        If DrawOverlayForm IsNot Nothing AndAlso DrawOverlayForm.IsDisposed = False Then
            If DrawOverlayForm.Visible <> mc_parentTextBox.Visible Then
                DrawOverlayForm.Visible = mc_parentTextBox.Visible
                If mc_parentTextBox.Visible Then
                    SetOverlayBounds()
                    RepaintTextBox()
                End If
            End If
        End If
    End Sub

    Private Sub RemoveAllOverlayHandlers()
        For Each item In OverlayHandlerControls
            RemoveHandler item.LocationChanged, AddressOf mc_parentTextBox_LocationChanged
            RemoveHandler item.VisibleChanged, AddressOf mc_parentTextBox_VisibleChanged
        Next
    End Sub

    Public OverlayHandlerControls As New List(Of Control)

    Private Sub mc_parentTextBox_ParentChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_parentTextBox.ParentChanged
        If DrawOverlayForm IsNot Nothing Then
            SetOverlayBounds()
            Try
                DrawOverlayForm.Visible = mc_parentTextBox.Visible
            Catch ex As ObjectDisposedException

            End Try
            Dim parentControl As Control = Me.parentTextBox.Parent
            RemoveAllOverlayHandlers()
            Do Until parentControl Is Nothing
                OverlayHandlerControls.Add(parentControl)
                AddHandler parentControl.LocationChanged, AddressOf mc_parentTextBox_LocationChanged
                AddHandler parentControl.VisibleChanged, AddressOf mc_parentTextBox_VisibleChanged
                parentControl = parentControl.Parent
            Loop
        End If
    End Sub

#End Region

#End Region

End Class