'i00 .Net HTML Tool Tip
'©i00 Productions All rights reserved
'Created by Kris Bennett
'----------------------------------------------------------------------------------------------------
'All property in this file is and remains the property of i00 Productions, regardless of its usage,
'unless stated otherwise in writing from i00 Productions.
'
'i00 is not and shall not be held accountable for any damages directly or indirectly caused by the
'use or miss-use of this product.  This product is only a component and thus is intended to be used 
'as part of other software, it is not a complete software package, thus i00 Productions is not 
'responsible for any legal ramifications that software using this product breaches.

<System.ComponentModel.DesignerCategory("")> _
Public Class HTMLToolTip
    Inherits ToolTip

    Public Class TipClickEventArgs
        Public mouse As MouseEventArgs
        Public handled As Boolean
    End Class

    
    Private WithEvents MasterForm As Form

    Private window As IWin32Window

#Region "Show"

    Overloads Sub Show(ByVal text As String, ByVal window As System.Windows.Forms.IWin32Window)
        Me.window = window
        MyBase.Show(text, window)
    End Sub

    Overloads Sub Show(ByVal text As String, ByVal window As System.Windows.Forms.IWin32Window, ByVal duration As Integer)
        Me.window = window
        MyBase.Show(text, window, duration)
    End Sub

    Overloads Sub Show(ByVal text As String, ByVal window As System.Windows.Forms.IWin32Window, ByVal x As Integer, ByVal y As Integer)
        Me.window = window
        MyBase.Show(text, window, x, y)
    End Sub

    Overloads Sub Show(ByVal text As String, ByVal window As System.Windows.Forms.IWin32Window, ByVal x As Integer, ByVal y As Integer, ByVal duration As Integer)
        Me.window = window
        MyBase.Show(text, window, x, y, duration)
    End Sub

    Overloads Sub Show(ByVal text As String, ByVal window As System.Windows.Forms.IWin32Window, ByVal point As System.Drawing.Point)
        Me.window = window
        MyBase.Show(text, window, point)
    End Sub

    Overloads Sub Show(ByVal text As String, ByVal window As System.Windows.Forms.IWin32Window, ByVal point As System.Drawing.Point, ByVal duration As Integer)
        Me.window = window
        MyBase.Show(text, window, point, duration)
    End Sub

#End Region


    Public Overloads Sub Hide()
        ToolTipWindow.hide(MyBase.UseAnimation OrElse MyBase.UseFading)
        'MyBase.Hide(window)
    End Sub

    Public Overloads Sub Hide(ByVal win As IWin32Window)
        ToolTipWindow.hide(MyBase.UseAnimation OrElse MyBase.UseFading)
        'MyBase.Hide(win)
    End Sub

    Private Sub MasterForm_Move_ResizeBegin(ByVal sender As Object, ByVal e As System.EventArgs) Handles MasterForm.Move, MasterForm.ResizeBegin
        'try hiding the tooltip if the window is moved - not normally picked up when a tooltip is shown
        ToolTipWindow.hide(MyBase.UseAnimation OrElse MyBase.UseFading)
    End Sub

    Dim ThisPopupOrientation As ToolTipOrientations

    Dim mc_LastText As String
    Public ReadOnly Property LastText() As String
        Get
            Return mc_LastText
        End Get
    End Property

    Public Sub ShowHTML(ByVal text As String, ByVal window As IWin32Window, ByVal point As Point, Optional ByVal Duration As Integer = Integer.MaxValue)
        mc_LastText = text
        If ToolTipWindow.TipImage IsNot Nothing Then
            Dim TipImage = ToolTipWindow.TipImage
            ToolTipWindow.TipImage = Nothing
            TipImage.Dispose()
        End If

        If ToolTipTitle <> "" Then
            text = "<font face='" & System.Drawing.SystemFonts.StatusFont.Name & "' size='" & CInt(System.Drawing.SystemFonts.StatusFont.Size * 1.2) & "' color='" & System.Drawing.ColorTranslator.ToHtml(DrawingFunctions.BlendColor(Color.FromKnownColor(KnownColor.Highlight), Color.FromKnownColor(KnownColor.ControlText))) & "'>" & ToolTipTitle & "</font>" & vbCrLf & text
        End If

        Dim ctl = TryCast(window, Control)
        If ctl IsNot Nothing Then
            If Me.ToolTipOrientation = ToolTipOrientations.Auto Then
                Dim p = ctl.PointToScreen(point.Empty)
                p.X += point.X
                p.Y += point.Y
                Dim ThisScreen = Screen.FromPoint(p)
                p.X -= ThisScreen.Bounds.X
                p.Y -= ThisScreen.Bounds.Y
                If p.X < ThisScreen.Bounds.Width / 2 Then
                    'left
                    If p.Y < ThisScreen.Bounds.Height / 2 Then
                        'top
                        ThisPopupOrientation = ToolTipOrientations.TopLeft
                    Else
                        'bot
                        ThisPopupOrientation = ToolTipOrientations.LowLeft
                    End If
                Else
                    'right
                    If p.Y < ThisScreen.Bounds.Height / 2 Then
                        'top
                        ThisPopupOrientation = ToolTipOrientations.TopRight
                    Else
                        'bot
                        ThisPopupOrientation = ToolTipOrientations.LowRight
                    End If
                End If
            Else
                ThisPopupOrientation = ToolTipOrientation
            End If

            MasterForm = ctl.FindForm()
        End If

        Dim TipSize As Size

        'If IsBalloon Then
        '    'no max width
        '    TipSize = SetTipSize(Nothing, text, point)
        'Else
        TipSize = SetTipSize(ctl, text, point)
        'End If


        If IsBalloon Then
            'offset so that the pointy bit points to the point instead of the corner
            Select Case ThisPopupOrientation
                Case ToolTipOrientations.LowLeft
                    point.X -= BalloonPointerStartAt
                    point.Y -= TipSize.Height - BalloonShadowDepth - CInt(BalloonShadowBlur / 2)
                Case ToolTipOrientations.TopRight
                    point.X -= TipSize.Width - BalloonPointerStartAt - BalloonShadowDepth - CInt(BalloonShadowBlur / 2)
                Case ToolTipOrientations.LowRight
                    point.X -= TipSize.Width - BalloonPointerStartAt - BalloonShadowDepth - CInt(BalloonShadowBlur / 2)
                    point.Y -= TipSize.Height - BalloonShadowDepth - CInt(BalloonShadowBlur / 2)
                Case ToolTipOrientations.TopLeft
                    point.X -= BalloonPointerStartAt
            End Select
        End If

        'MyBase.Show(text, window, point, Duration)
        'Dim ToolTipPopup As New ToolTipPopup

        Dim Control = TryCast(window, Control)
        If Control IsNot Nothing Then
            point.X += Control.PointToScreen(point.Empty).X
            point.Y += Control.PointToScreen(point.Empty).Y
        End If
        ToolTipWindow.Location = point

        Using b As New Bitmap(TipSize.Width, TipSize.Height)
            Using g = Graphics.FromImage(b)
                HTMLToolTipDraw(g, TipSize, text)
            End Using
            ToolTipWindow.Show(window, MyBase.UseAnimation OrElse MyBase.UseFading, Duration, b)
        End Using

    End Sub

    'qwertyuiop - cannot set ToolTipWindow with topmost or the ShowNoFocus doesn't work :(
    Public WithEvents ToolTipWindow As New ToolTipPopup With {.ShowInTaskbar = False, .StartPosition = FormStartPosition.Manual}

    Public Event TipClosed(ByVal sender As Object, ByVal e As EventArgs)
    Public Event TipClick(ByVal sender As Object, ByVal e As TipClickEventArgs)

    Private Sub ToolTipWindow_TipClick(ByVal sender As Object, ByVal e As TipClickEventArgs) Handles ToolTipWindow.TipClick
        RaiseEvent TipClick(Me, e)
    End Sub
    Private Sub ToolTipWindow_TipClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolTipWindow.TipClosed
        RaiseEvent TipClosed(Me, EventArgs.Empty)
    End Sub

    Public Class ToolTipPopup
        Inherits PerPixelAlphaForm

        Public Event TipClosed(ByVal sender As Object, ByVal e As EventArgs)
        Public Event TipClick(ByVal sender As Object, ByVal e As TipClickEventArgs)

        Public Const WM_MOUSEACTIVATE As Integer = &H21
        Public Const MA_NOACTIVATE As Integer = 3
        Private Shared ReadOnly noActivate As New IntPtr(MA_NOACTIVATE)

        Dim DoFadeOut As Boolean

        Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
            MyBase.WndProc(m)
            If m.Msg = WM_MOUSEACTIVATE Then
                'Tell Windows not to activate the window on a mouse click.
                m.Result = noActivate
            End If
        End Sub

        Private WithEvents TimerClose As New Timer

        Public TipImage As Bitmap

        Private Sub ResetCloseTimer(ByVal Duration As Integer)
            If Duration = -1 Then
                TimerClose.Enabled = False
            Else
                TimerClose.Interval = Duration
                TimerClose.Enabled = False
                TimerClose.Enabled = True
            End If
        End Sub

        Private WithEvents FadeTimer As New Timer With {.Interval = 1}
        Dim FadeFrom As Byte
        Dim FadeTo As Byte
        Dim StartFadeTime As Integer
        Const FadeTime As Integer = 250

        Dim CurrentFade As Byte
        Public Overrides Sub SetBitmap(ByVal bitmap As System.Drawing.Bitmap, ByVal opacity As Byte)
            CurrentFade = opacity
            If MyBase.IsDisposed = False Then
                MyBase.SetBitmap(bitmap, opacity)
            End If
        End Sub

        Public Overloads Sub show(ByVal IWin32Window As IWin32Window, ByVal Animate As Boolean, ByVal Duration As Integer, ByVal Image As Bitmap)
            If Me.Visible Then
                DoFadeOut = Animate
                TipImage = DirectCast(Image.Clone(), Bitmap)
                If Animate Then
                    'make it fade in... from what it was if we are already fading...
                    If FadeTimer.Enabled Then
                        FadeFrom = CurrentFade
                        FadeTo = 255
                        SetBitmap(TipImage, 0)
                        StartFadeTime = Environment.TickCount
                        FadeTimer.Enabled = True
                    Else
                        SetBitmap(TipImage, 255)
                    End If
                Else
                    SetBitmap(TipImage, 255)
                    FadeTimer.Enabled = False
                End If
            Else
                DoFadeOut = Animate
                TipImage = DirectCast(Image.Clone(), Bitmap)
                If Animate Then
                    'make it fade in
                    FadeFrom = 0
                    FadeTo = 255
                    SetBitmap(TipImage, 0)
                    StartFadeTime = Environment.TickCount
                    MyBase.ShowNoFocus(IWin32Window)
                    FadeTimer.Enabled = True
                Else
                    MyBase.ShowNoFocus(IWin32Window)
                    SetBitmap(TipImage, 255)
                    FadeTimer.Enabled = False
                End If
            End If
            ResetCloseTimer(Duration)
        End Sub
        Public Overloads Sub hide(Optional ByVal Animate As Boolean = True)
            If Animate Then
                'make it fade out
                FadeFrom = CurrentFade
                FadeTo = 0
                StartFadeTime = Environment.TickCount
                FadeTimer.Enabled = True
            Else
                MyBase.Hide()
                RaiseEvent TipClosed(Me, EventArgs.Empty)
            End If
        End Sub

        Private Sub Timer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles TimerClose.Tick
            Me.hide(DoFadeOut)
            TimerClose.Enabled = False
        End Sub

        Private Sub FadeTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles FadeTimer.Tick
            If TipImage IsNot Nothing Then
                Dim ThisTime = Environment.TickCount
                If ThisTime - StartFadeTime > FadeTime Then
                    If FadeTo = 0 Then
                        MyBase.Hide()
                        RaiseEvent TipClosed(Me, EventArgs.Empty)
                    Else
                        SetBitmap(TipImage, FadeTo)
                    End If
                    FadeTimer.Enabled = False
                Else
                    Dim AnimationPercent = ((ThisTime - StartFadeTime) / FadeTime)
                    Dim FadeAmount As Byte
                    FadeAmount = CByte((FadeFrom + ((CInt(FadeTo) - CInt(FadeFrom)) * AnimationPercent)))
                    SetBitmap(TipImage, FadeAmount)
                End If
            End If
        End Sub

        Private Sub ToolTipPopup_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
            If TipImage IsNot Nothing Then
                TipImage.Dispose()
            End If
            TimerClose.Dispose()
            FadeTimer.Dispose()

            If TipImage IsNot Nothing Then
                TipImage.Dispose()
            End If

        End Sub

        Private Sub ToolTipPopup_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick
            Dim theE = New TipClickEventArgs() With {.mouse = e}
            RaiseEvent TipClick(sender, theE)
            If theE.handled = False Then
                Me.hide(DoFadeOut)
            End If
        End Sub
    End Class

    Dim LastMaxTipWidth As Integer

    Private Function SetTipSize(ByVal control As Control, ByVal TipText As String, ByVal ToolTipPoint As Point) As Size
        If control Is Nothing Then
            Dim TipSizeF = HTMLParser.PaintHTML(TipText, Nothing, , HTMLRenderStatus)
            SetTipSize = New Size(CInt(TipSizeF.Width) + 1, CInt(TipSizeF.Height) + 1)
        Else
            Dim RealTipPoint = control.PointToScreen(New Point(0, 0))
            RealTipPoint.X += ToolTipPoint.X
            RealTipPoint.Y += ToolTipPoint.Y

            Dim Onscreen = Screen.FromPoint(RealTipPoint)

            Dim MaxTipWidth = RealTipPoint.X
            If IsBalloon And (ThisPopupOrientation = ToolTipOrientations.LowRight OrElse ThisPopupOrientation = ToolTipOrientations.TopRight) Then
            Else
                MaxTipWidth = Onscreen.Bounds.Right - MaxTipWidth
                If IsBalloon Then MaxTipWidth -= (BalloonBorderMargin * 2)
            End If
            MaxTipWidth -= If(Me.ToolTipIcon <> Windows.Forms.ToolTipIcon.None, 22, 0)


            LastMaxTipWidth = MaxTipWidth

            Dim TipSizeF = HTMLParser.PaintHTML(TipText, Nothing, MaxTipWidth, HTMLRenderStatus)
            SetTipSize = New Size(CInt(TipSizeF.Width) + 1, CInt(TipSizeF.Height) + 1)

            If SetTipSize.Width > MaxTipWidth Then
                SetTipSize.Width = MaxTipWidth
            End If
            'If RealTipPoint.Y + SetTipSize.Height > Onscreen.Bounds.Bottom Then
            '    SetTipSize.Height = Onscreen.Bounds.Bottom - RealTipPoint.Y
            'End If
        End If
        If IsBalloon Then
            SetTipSize.Height += (BalloonBorderMargin * 2) + BalloonPointerHeight
            SetTipSize.Width += (BalloonBorderMargin * 2)
        End If
        'add room for the shadow
        SetTipSize.Height += BalloonShadowDepth + CInt(BalloonShadowBlur / 2)
        SetTipSize.Width += BalloonShadowDepth + CInt(BalloonShadowBlur / 2)

        'add some right padding
        SetTipSize.Width += RightPadding

        If Me.ToolTipIcon <> Windows.Forms.ToolTipIcon.None Then
            SetTipSize.Width += 22
        End If
    End Function

#Region "APIs"

    <System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="GetClassLong")> _
    Public Shared Function GetClassLongPtr32(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SetClassLong")> _
    Public Shared Function SetClassLongPtr32(ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    End Function

    Private Const GCL_STYLE As Integer = -26
    Private Const CS_DROPSHADOW As Integer = &H20000

#End Region

    Public ReadOnly Property Handle() As IntPtr
        Get
            Dim obj As Object
            Dim hwnd As IntPtr
            Try
                hwnd = IntPtr.Zero
                obj = GetType(ToolTip).InvokeMember("Handle", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.GetProperty, Nothing, Me, Nothing)
                hwnd = CType(obj, IntPtr)
            Catch ex As Exception

            End Try
            Return hwnd
        End Get
    End Property

    Public Shared Function GenBalloonPath(ByVal rect As Rectangle, ByVal Orientation As ToolTipOrientations) As Drawing2D.GraphicsPath
        GenBalloonPath = New Drawing2D.GraphicsPath()
        GenBalloonPath.AddArc(New Rectangle(0, 0, 6, 6), 180, 90)
        GenBalloonPath.AddArc(New Rectangle(rect.Right - 6, 0, 6, 6), 270, 90)
        GenBalloonPath.AddArc(New Rectangle(rect.Right - 6, rect.Height - 6 - BalloonPointerHeight, 6, 6), 0, 90)
        GenBalloonPath.AddLine(BalloonPointerStartAt + BalloonPointerHeight, rect.Height - BalloonPointerHeight, BalloonPointerStartAt, rect.Height)
        GenBalloonPath.AddLine(BalloonPointerStartAt, rect.Height, BalloonPointerStartAt, rect.Height - BalloonPointerHeight)
        GenBalloonPath.AddArc(New Rectangle(0, rect.Height - 6 - BalloonPointerHeight, 6, 6), 90, 90)
        GenBalloonPath.CloseFigure()
    End Function

    Public Enum ShadowModes
        None
        Basic
        Complex
    End Enum

    Dim mc_ShadowMode As ShadowModes = ShadowModes.Basic
    Public Property ShadowMode() As ShadowModes
        Get
            Return mc_ShadowMode
        End Get
        Set(ByVal value As ShadowModes)
            mc_ShadowMode = value
        End Set
    End Property

    Dim mc_Image As Bitmap
    Public Property Image() As Bitmap
        Get
            Return mc_Image
        End Get
        Set(ByVal value As Bitmap)
            mc_Image = value
        End Set
    End Property

    Dim mc_ShadowBaseColor As Color = Color.Black
    Public Property ShadowBaseColor() As Color
        Get
            Return mc_ShadowBaseColor
        End Get
        Set(ByVal value As Color)
            mc_ShadowBaseColor = value
        End Set
    End Property

    Private Sub HTMLToolTipDraw(ByVal g As Graphics, ByVal TipSize As Size, ByVal TipText As String)
        Dim Rect As New Rectangle(0, 0, TipSize.Width - 1, TipSize.Height - 1)
        g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        g.InterpolationMode = Drawing2D.InterpolationMode.High

        Dim DataOffsetPoint = New Point(0, 0)

        Dim IconImage As Bitmap = mc_Image
        If IconImage Is Nothing Then
            Select Case MyBase.ToolTipIcon
                Case Windows.Forms.ToolTipIcon.Error
                    IconImage = My.Resources._error
                Case Windows.Forms.ToolTipIcon.Info
                    IconImage = My.Resources.info
                Case Windows.Forms.ToolTipIcon.Warning
                    IconImage = My.Resources.Warning
            End Select
        End If

        Dim BalloonRect = New Rectangle(Rect.X, Rect.Y, Rect.Width - BalloonShadowDepth - CInt(BalloonShadowBlur / 2), Rect.Height - BalloonShadowDepth - CInt(BalloonShadowBlur / 2))

        If IsBalloon Then
            DataOffsetPoint.X += BalloonBorderMargin
            DataOffsetPoint.Y += BalloonBorderMargin

            Using p = GenBalloonPath(BalloonRect, ToolTipOrientations.LowLeft)
                Select Case ThisPopupOrientation
                    Case ToolTipOrientations.LowLeft
                        'default :)
                    Case ToolTipOrientations.LowRight
                        Using m As New Drawing2D.Matrix(-1, 0, 0, 1, 0, 0)
                            m.Translate(-BalloonRect.Width, 0)
                            p.Transform(m)
                        End Using
                    Case ToolTipOrientations.TopLeft
                        Using m As New Drawing2D.Matrix(1, 0, 0, -1, 0, 0)
                            m.Translate(0, -BalloonRect.Height)
                            p.Transform(m)
                        End Using
                        DataOffsetPoint.Y += BalloonPointerHeight
                    Case ToolTipOrientations.TopRight
                        Using m As New Drawing2D.Matrix
                            m.RotateAt(180, New PointF(CSng(BalloonRect.Width / 2), CSng(BalloonRect.Height / 2)))
                            p.Transform(m)
                        End Using
                        DataOffsetPoint.Y += BalloonPointerHeight
                End Select

                Select Case ShadowMode
                    Case ShadowModes.Basic
                        For i = 0 To BalloonShadowDepth
                            g.TranslateTransform(1, 1)
                            Using b As New SolidBrush(Color.FromArgb(CInt((255 - ((i / BalloonShadowDepth) * 255)) / BalloonShadowDepth), ShadowBaseColor))
                                g.FillPath(b, p)
                            End Using
                        Next
                        g.ResetTransform()
                    Case ShadowModes.Complex
                        Using bShadow As New Bitmap(BalloonRect.Width + BalloonShadowBlur, BalloonRect.Height + BalloonShadowBlur)
                            Using gShadow = Graphics.FromImage(bShadow)
                                gShadow.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                                Using sb As New SolidBrush(Color.FromArgb(63, ShadowBaseColor))
                                    'gShadow.TranslateTransform(BalloonShadowBlur / 2, BalloonShadowBlur / 2)
                                    gShadow.FillPath(sb, p)
                                End Using
                            End Using
                            bShadow.Filters.GausianBlur(BalloonShadowBlur)
                            g.DrawImage(bShadow, Rect.Width - bShadow.Width, Rect.Height - bShadow.Height)
                        End Using
                End Select

                ''shadow method 2
                'e.Graphics.TranslateTransform(BalloonShadowDepth, BalloonShadowDepth)
                'Using pgb As New Drawing.Drawing2D.PathGradientBrush(p)
                '    pgb.WrapMode = Drawing2D.WrapMode.Clamp
                '    Dim _ColorBlend As New Drawing2D.ColorBlend(3)
                '    _ColorBlend.Colors = New Color() {Color.Transparent, Color.FromArgb(63, ShadowBaseColor), Color.FromArgb(63, ShadowBaseColor)}
                '    _ColorBlend.Positions = New Single() {0, 0.1, 1}
                '    pgb.InterpolationColors = _ColorBlend
                '    e.Graphics.FillPath(pgb, p)
                'End Using
                'e.Graphics.ResetTransform()

                'shadow method 3

                Using lgb As New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 0), New Point(0, TipSize.Height), Color.FromKnownColor(KnownColor.ControlLightLight), Color.FromKnownColor(KnownColor.ControlLight))
                    g.FillPath(lgb, p)
                End Using

                'should be the same as below...
                If IconImage IsNot Nothing Then g.DrawImage(IconImage, New Rectangle(DataOffsetPoint.X, DataOffsetPoint.Y + 2, 16, 16))

                If Me.ToolTipIcon <> Windows.Forms.ToolTipIcon.None Then
                    DataOffsetPoint.X += 22
                End If

                g.TranslateTransform(DataOffsetPoint.X, DataOffsetPoint.Y)
                HTMLParser.PaintHTML(TipText, g, LastMaxTipWidth, HTMLRenderStatus)
                g.ResetTransform()
                'end same as below

                Using pen As New Pen(Color.FromKnownColor(KnownColor.ControlDark))
                    g.DrawPath(pen, p)
                End Using

            End Using
        Else
            Select Case ShadowMode
                Case ShadowModes.Basic
                    For i = 0 To BalloonShadowDepth
                        g.TranslateTransform(1, 1)
                        Using b As New SolidBrush(Color.FromArgb(CInt((255 - ((i / BalloonShadowDepth) * 255)) / BalloonShadowDepth), ShadowBaseColor))
                            g.FillRectangle(b, BalloonRect)
                        End Using
                    Next
                    g.ResetTransform()
                Case ShadowModes.Complex
                    Using bShadow As New Bitmap(BalloonRect.Width + BalloonShadowBlur, BalloonRect.Height + BalloonShadowBlur)
                        Using gShadow = Graphics.FromImage(bShadow)
                            gShadow.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                            Using sb As New SolidBrush(Color.FromArgb(63, ShadowBaseColor))
                                gShadow.TranslateTransform(BalloonShadowBlur / 2, BalloonShadowBlur / 2)
                                gShadow.FillRectangle(sb, BalloonRect)
                            End Using
                        End Using
                        bShadow.Filters.GausianBlur(BalloonShadowBlur)
                        g.DrawImage(bShadow, Rect.Width - bShadow.Width, Rect.Height - bShadow.Height)
                    End Using
            End Select

            Using lgb As New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 0), New Point(0, TipSize.Height), Color.FromKnownColor(KnownColor.ControlLightLight), Color.FromKnownColor(KnownColor.ControlLight))
                g.FillRectangle(lgb, BalloonRect)
            End Using

            'should be the same as above...
            '...cept for the plus 3 below...
            If IconImage IsNot Nothing Then g.DrawImage(IconImage, New Rectangle(DataOffsetPoint.X + 3, DataOffsetPoint.Y + 2, 16, 16))

            If Me.ToolTipIcon <> Windows.Forms.ToolTipIcon.None Then
                DataOffsetPoint.X += 22
            End If
            g.TranslateTransform(DataOffsetPoint.X + 3, DataOffsetPoint.Y)
            HTMLParser.PaintHTML(TipText, g, TipSize.Width - RightPadding, HTMLRenderStatus)
            g.ResetTransform()
            'end same as above

            Using p As New Pen(DrawingFunctions.BlendColor(Color.FromKnownColor(KnownColor.Control), Color.FromKnownColor(KnownColor.ControlText)))
                g.DrawRectangle(p, BalloonRect)
            End Using
        End If
    End Sub

    Public Function HTMLRenderStatus() As HTMLParser.Status
        HTMLRenderStatus = New HTMLParser.Status
        HTMLRenderStatus.Font = New HTMLParser.STRFont(System.Drawing.SystemFonts.StatusFont)
        HTMLRenderStatus.Brush = New HTMLParser.STRBrush(Color.FromKnownColor(KnownColor.ControlText))
    End Function

    Public ToolTipOrientation As ToolTipOrientations = ToolTipOrientations.Auto

    Public Enum ToolTipOrientations
        Auto
        TopLeft
        TopRight
        LowLeft
        LowRight
    End Enum

    Const BalloonBorderMargin As Integer = 12
    Const BalloonPointerHeight As Integer = 21
    Const BalloonPointerStartAt As Integer = 16
    Const BalloonShadowDepth As Integer = 4
    Const BalloonShadowBlur As Integer = 4
    Const RightPadding As Integer = 4

    Delegate Sub HTMLToolTip_Disposed_cb(ByVal sender As Object, ByVal e As System.EventArgs)
    Private Sub HTMLToolTip_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        'qwertyuiop - on some reason the ToolTipWindow.invoke fails on occasion... when this happens the .dispose method seems to work .. even though it says it needs invoking :S
        Try
            If ToolTipWindow.InvokeRequired Then
                Dim HTMLToolTip_Disposed_cb As New HTMLToolTip_Disposed_cb(AddressOf HTMLToolTip_Disposed)
                ToolTipWindow.Invoke(HTMLToolTip_Disposed_cb, sender, e)
            Else
                ToolTipWindow.Dispose()
            End If
        Catch ex As Exception
            Try
                ToolTipWindow.Dispose()
            Catch ex2 As Exception
            End Try
        End Try
    End Sub
End Class
