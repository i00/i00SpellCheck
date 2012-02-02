Imports System.Drawing.Drawing2D

Public Class frmSplash

    Private Sub bpBackGround_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles bpBackGround.Paint
        e.Graphics.SmoothingMode = SmoothingMode.HighQuality

        Using gb As New LinearGradientBrush(New Point(0, 0), New Point(0, bpBackGround.ClientSize.Height), Color.Transparent, i00SpellCheck.DrawingFunctions.AlphaColor(Color.Black, 31))
            e.Graphics.FillRectangle(gb, bpBackGround.ClientRectangle)
        End Using
        'e.Graphics.InterpolationMode = Drawing2D.InterpolationMode.High
        'e.Graphics.DrawImage(My.Resources.Icon, New Point(CInt((bpLogo.ClientSize.Width - My.Resources.Icon.Width) / 2), CInt((bpLogo.ClientSize.Height - My.Resources.Icon.Height) / 2)))
        Using b As New SolidBrush(Color.FromArgb(255, 120, 154, 255))
            'Dim LogoSize As Single = Me.ClientSize.Width
            Dim LogoSize = CSng(Me.ClientSize.Height * 1.25)
            Dim LeftPos = CSng(Me.ClientSize.Width - (LogoSize * 0.75))
            If LeftPos < -(LogoSize * 0.25) Then LeftPos = CSng(-(LogoSize * 0.25))
            i00SpellCheck.DrawingFunctions.DrawLogo(e.Graphics, b, New RectangleF(LeftPos, CSng((bpBackGround.ClientSize.Height - LogoSize) / 2), LogoSize, LogoSize))
        End Using
    End Sub

    Private Sub frmSplash_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ControlColor = i00SpellCheck.DrawingFunctions.AlphaColor(Color.FromKnownColor(KnownColor.White), 127)
        bpDocumentHolder.BackColor = ControlColor

        Dim ControlColorBot = i00SpellCheck.DrawingFunctions.AlphaColor(Color.FromKnownColor(KnownColor.ControlDark), 63)
        bpCommandBar.BackColor = ControlColorBot

    End Sub

    Private Sub ProductLink_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ProductLink.Click
        System.Diagnostics.Process.Start("http://www.vbforums.com/showthread.php?p=4075093")
    End Sub

End Class