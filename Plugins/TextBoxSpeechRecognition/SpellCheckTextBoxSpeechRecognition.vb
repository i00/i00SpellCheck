Imports i00SpellCheck

'The weight is the priority order of the plugins ... if there are multiple plugins that extend the same control ...
'in this case TextBoxBase the plugin with the heighest weight gets used ...
'all built in plugins in i00SpellCheck have a weight of 0, if you need to extend their functionality you should use a higher weight
<i00SpellCheck.PluginWeight(1)> _
Public Class SpellCheckTextBoxSpeechRecognition
    Inherits i00SpellCheck.SpellCheckTextBox

    'called when the control is loaded...
    Public Overrides Sub Load()
        MyBase.Load()
        If parentTextBox IsNot Nothing Then
            AddHandler parentTextBox.KeyUp, AddressOf parentTextBox_KeyUp
            AddHandler parentTextBox.LostFocus, AddressOf parentTextBox_LostFocus
            AddHandler parentTextBox.TextChanged, AddressOf parentTextBox_TextChanged
            AddHandler parentTextBox.MouseDown, AddressOf parentTextBox_MouseDown
            AddHandler parentTextBox.KeyPress, AddressOf parentTextBox_KeyPress
        End If
    End Sub

    Private WithEvents Recogniser As Speech.Recognition.SpeechRecognitionEngine

#Region "Images"

    Private Shared ReadOnly Property MicImage() As Bitmap
        Get
            Static b As Bitmap
            If b Is Nothing Then
                b = New Bitmap(16, 16)
                Using g = Graphics.FromImage(b)
                    g.InterpolationMode = Drawing2D.InterpolationMode.High
                    g.DrawImage(My.Resources.Microphone, New Rectangle(0, 0, b.Width, b.Height))
                End Using
            End If
            Return b
        End Get
    End Property

    Private Shared ReadOnly Property MicImageBW() As Bitmap
        Get
            Static b As Bitmap
            If b Is Nothing Then
                b = DirectCast(MicImage.Clone, Bitmap)
                b.Filters.AlphaMask(Color.Transparent, Color.FromArgb(127, 0, 0, 0))
            End If
            Return b
        End Get
    End Property

#End Region

    Private WithEvents DictateToolTip As i00SpellCheck.HTMLToolTip
    Dim ToolTipLocation As Point
    Dim ExistingText As String
    Dim tmpText As String

    Dim DictateCancel As Boolean

    Dim ListeningColor As String = System.Drawing.ColorTranslator.ToHtml(i00SpellCheck.DrawingFunctions.BlendColor(Color.FromKnownColor(KnownColor.ControlText), Color.FromKnownColor(KnownColor.Control)))
    Private Sub parentTextBox_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Static LastF12Press As New Date
        If e.KeyCode = Keys.F12 Then
            If (Now.Subtract(LastF12Press).TotalMilliseconds < SystemInformation.DoubleClickTime) Then
                'do dictate
                '... unless we are already...
                If DictateToolTip Is Nothing Then
                    'get position for tooltip
                    ScrollToCaret()
                    ToolTipLocation = Me.parentTextBox.GetPositionFromCharIndex(parentTextBox.SelectionStart)
                    Dim lineHeight = GetLineHeightFromCharPosition(parentTextBox.SelectionStart)
                    ToolTipLocation.Y += lineHeight

                    Dim [Error] As Exception = Nothing
                    Try
                        'dictate
                        tmpText = ""
                        ExistingText = ""
                        Recogniser = New Speech.Recognition.SpeechRecognitionEngine
                        Recogniser.SetInputToDefaultAudioDevice()
                        Recogniser.LoadGrammar(New System.Speech.Recognition.DictationGrammar)
                        Recogniser.RecognizeAsync(Speech.Recognition.RecognizeMode.Multiple)
                    Catch ex As Exception
                        [Error] = ex
                    End Try

                    If [Error] Is Nothing Then
                        'tooltip
                        DictateCancel = False
                        LastAudioLevel = 0
                        DictateToolTip = New i00SpellCheck.HTMLToolTip With {.IsBalloon = True, .ToolTipTitle = "Dictate what you want typed", .ToolTipIcon = ToolTipIcon.Info}
                        DictateToolTip.ToolTipOrientation = i00SpellCheck.HTMLToolTip.ToolTipOrientations.TopLeft
                        DictateToolTip.Image = DirectCast(MicImageBW.Clone, Bitmap)
                        DictateToolTip.ShowHTML("<i><font color=" & ListeningColor & ">Listening<br></i>...click this balloon or wait after you are done talking to confirm<br>...click on the textbox or press <i>&lt;Escape&gt;</i> to cancel</font>", parentTextBox, ToolTipLocation, 5000)
                        'do something like?: '<br>...or right click the balloon when finished for corrections
                    Else
                        'show the error
                        LastAudioLevel = 0
                        DictateToolTip = New i00SpellCheck.HTMLToolTip With {.IsBalloon = True, .ToolTipTitle = "Error starting dictate", .ToolTipIcon = ToolTipIcon.Error}
                        DictateToolTip.ToolTipOrientation = i00SpellCheck.HTMLToolTip.ToolTipOrientations.TopLeft
                        DictateToolTip.Image = Nothing
                        DictateToolTip.ShowHTML("Please make sure your soundcard is working correctly", parentTextBox, ToolTipLocation, 5000)
                    End If

                End If

                LastF12Press = New Date
            Else
                LastF12Press = Now
            End If
        Else
            LastF12Press = New Date
        End If
    End Sub

    Private Sub parentTextBox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        'cancel
        CancelRecognition()
    End Sub

    Private Sub parentTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'cancel
        CancelRecognition()
    End Sub

    Private Sub parentTextBox_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        'cancel
        CancelRecognition()
    End Sub

    Private Sub parentTextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Asc(e.KeyChar) = Keys.Escape Then
            CancelRecognition()
        End If
    End Sub

    Private Sub CancelRecognition()
        DictateCancel = True
        If DictateToolTip IsNot Nothing Then DictateToolTip.Hide()
    End Sub


    Private Sub CommitRecognition()
        If DictateCancel = False Then
            'write the spoken text
            If tmpText <> "" Then
                ExistingText &= " " & tmpText
            End If
            parentTextBox.SelectedText = Trim(ExistingText) & " "
            DictateCancel = True
        End If
    End Sub

#Region "For the test control"

    Public Overrides Function SetupControl(ByVal Control As System.Windows.Forms.Control) As Control
        SetupControl = MyBase.SetupControl(Control)
        Dim TextBoxBase = TryCast(SetupControl, TextBoxBase)
        If TextBoxBase IsNot Nothing Then
            TextBoxBase.AppendText(vbCrLf & vbCrLf & "The TextBoxSpeechRecognition project now adds speech recognition functionality to any TextBox, simply press F12 twice (quickly) to dictate what you want written!")
        End If
    End Function

#End Region

    Private Sub DictateToolTip_TipClick(ByVal sender As Object, ByVal e As i00SpellCheck.HTMLToolTip.TipClickEventArgs) Handles DictateToolTip.TipClick
        CommitRecognition()
    End Sub

    Private Sub DictateToolTip_TipClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles DictateToolTip.TipClosed

        CommitRecognition()

        If DictateToolTip IsNot Nothing Then
            Try
                DictateToolTip.Dispose()
            Catch ex As Exception

            End Try
        End If
        DictateToolTip = Nothing

        If Recogniser IsNot Nothing Then
            Try
                Recogniser.RecognizeAsyncStop()
            Catch ex As Exception

            End Try
            Recogniser = Nothing
        End If
    End Sub

    Dim LastAudioLevel As Integer
    Private Sub Recogniser_AudioLevelUpdated(ByVal sender As Object, ByVal e As System.Speech.Recognition.AudioLevelUpdatedEventArgs) Handles Recogniser.AudioLevelUpdated
        If DictateCancel Then Exit Sub
        Dim AudioLevel = CInt(Int((e.AudioLevel / 100) * 16))

        If LastAudioLevel <> AudioLevel Then
            'redraw...
            Dim b As New Bitmap(16, 16)
            Using g = Graphics.FromImage(b)
                g.DrawImageUnscaled(MicImageBW, New Point(0, 0))
                g.SetClip(New Rectangle(0, 16 - AudioLevel, 16, AudioLevel))
                g.DrawImageUnscaled(MicImage, New Point(0, 0))
            End Using

            If DictateToolTip.Image IsNot Nothing Then
                DictateToolTip.Image.Dispose()
                DictateToolTip.Image = Nothing
            End If

            DictateToolTip.Image = b
            DictateToolTip.ShowHTML(DictateToolTip.LastText, parentTextBox, ToolTipLocation, 3000)

            LastAudioLevel = AudioLevel
        End If

    End Sub

    Dim strYouSaid As String = "<i><font color=" & ListeningColor & ">You Said:</font></i> "

    Private Sub Recogniser_SpeechHypothesized(ByVal sender As Object, ByVal e As System.Speech.Recognition.SpeechHypothesizedEventArgs) Handles Recogniser.SpeechHypothesized
        If DictateCancel Then Exit Sub
        tmpText = e.Result.Text
        DictateToolTip.ShowHTML(strYouSaid & ExistingText & "<font color=" & ListeningColor & ">" & tmpText & "</font>", parentTextBox, ToolTipLocation)
    End Sub

    'Dim InCorrectColor As String = System.Drawing.ColorTranslator.ToHtml(i00SpellCheck.DrawingFunctions.BlendColor(Color.Red, Color.FromKnownColor(KnownColor.ControlText), 127))

    Private Sub Recogniser_SpeechRecognized(ByVal sender As Object, ByVal e As System.Speech.Recognition.SpeechRecognizedEventArgs) Handles Recogniser.SpeechRecognized
        If DictateCancel Then Exit Sub
        tmpText = ""
        If Trim(e.Result.Text) <> "" Then
            ''for error coloring
            'For Each item In e.Result.Words
            '    If (item.DisplayAttributes And Speech.Recognition.DisplayAttributes.ConsumeLeadingSpaces) = Speech.Recognition.DisplayAttributes.ConsumeLeadingSpaces AndAlso ExistingText <> "" Then
            '        ExistingText = ExistingText.TrimEnd()
            '    End If
            '    ExistingText &= If(item.Confidence < 0.25, "<font color=" & InCorrectColor & "><b>", "") & item.Text & If(item.Confidence < 0.25, "</b></font>", "")
            '    If (item.DisplayAttributes And Speech.Recognition.DisplayAttributes.OneTrailingSpace) = Speech.Recognition.DisplayAttributes.OneTrailingSpace Then
            '        ExistingText &= " "
            '    ElseIf (item.DisplayAttributes And Speech.Recognition.DisplayAttributes.TwoTrailingSpaces) = Speech.Recognition.DisplayAttributes.TwoTrailingSpaces Then
            '        ExistingText &= "  "
            '    End If
            'Next

            ExistingText &= e.Result.Text & " "

            ''to play audio
            'Dim s As New IO.MemoryStream
            'e.Result.Audio.WriteToWaveStream(s)
            's.Position = 0
            'My.Computer.Audio.Play(s, AudioPlayMode.Background)

            DictateToolTip.ShowHTML(strYouSaid & ExistingText, parentTextBox, ToolTipLocation, 3000)
        End If
    End Sub
End Class
