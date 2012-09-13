'i00 .Net Spell Check - TextBoxSpeechRecognition
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

Imports i00SpellCheck

Partial Class SpellCheckTextBoxSpeechRecognition

    Private Class SpeechEngineHelper

        Private Shared WithEvents TextBox As TextBoxBase

        Private Shared Sub TextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox.KeyPress
            Select Case Asc(e.KeyChar)
                Case Keys.Enter, Keys.Return
                    'qwertyuiop - doesnt work ????
                    'oh well can still use timer and balloon click
                    e.Handled = True
                    CommitRecognition()
                Case Keys.Escape
                    CancelRecognition()
            End Select
        End Sub

#Region "Dictate"

        Friend Shared WithEvents Recogniser As Speech.Recognition.SpeechRecognitionEngine

        Friend Shared WithEvents DictateToolTip As i00SpellCheck.HTMLToolTip

        Friend Shared DictateCancel As Boolean

        Private Shared LastAudioLevel As Integer

        Friend Shared Sub CancelRecognition()
            TextBox = Nothing
            DictateCancel = True
            If DictateToolTip IsNot Nothing Then DictateToolTip.Hide()
        End Sub

        Private Shared Sub Recogniser_AudioLevelUpdated(ByVal sender As Object, ByVal e As System.Speech.Recognition.AudioLevelUpdatedEventArgs) Handles Recogniser.AudioLevelUpdated
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
                DictateToolTip.ShowHTML(DictateToolTip.LastText, TextBox, ToolTipLocation, 3000)

                LastAudioLevel = AudioLevel
            End If
        End Sub

        Private Shared ToolTipLocation As Point
        Private Shared ExistingText As String
        Private Shared tmpText As String
        Private Shared ListeningColor As String = System.Drawing.ColorTranslator.ToHtml(i00SpellCheck.DrawingFunctions.BlendColor(Color.FromKnownColor(KnownColor.ControlText), Color.FromKnownColor(KnownColor.Control)))

        Friend Shared Sub DoDictate(ByVal TextBox As TextBoxBase)
            'do dictate
            '... unless we are already...
            If SpeechEngineHelper.DictateToolTip Is Nothing Then
                'cancel speech
                CancelSynthesis()

                'speech balloon pos...
                Dim TextBoxSpellCheck = DirectCast(TextBox.SpellCheck, i00SpellCheck.SpellCheckTextBox)
                TextBoxSpellCheck.ScrollToCaret()
                SpeechEngineHelper.ToolTipLocation = TextBox.GetPositionFromCharIndex(TextBox.SelectionStart)
                Dim lineHeight = TextBoxSpellCheck.GetLineHeightFromCharPosition(TextBox.SelectionStart)
                ToolTipLocation.Y += lineHeight


                'dictate
                SpeechEngineHelper.TextBox = TextBox
                'SpeechEngineHelper.ToolTipLocation = ToolTipLocation
                tmpText = ""
                ExistingText = ""
                Dim [Error] As Exception = Nothing
                Try
                    Recogniser = New Speech.Recognition.SpeechRecognitionEngine
                    Try
                        Recogniser.SetInputToDefaultAudioDevice()
                    Catch ex As InvalidOperationException
                        Throw New Exception("Could not set the default audio device" & vbCrLf & "Please make sure your soundcard and microphone are working correctly")
                    End Try
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
                    DictateToolTip.ShowHTML("<i><font color=" & ListeningColor & ">Listening<br></i>...click this balloon or wait after you are done talking to confirm<br>...click on the textbox or press <i>&lt;Escape&gt;</i> to cancel</font>", TextBox, ToolTipLocation, 5000)
                    'do something like?: '<br>...or right click the balloon when finished for corrections
                Else
                    'show the error
                    LastAudioLevel = 0
                    DictateToolTip = New i00SpellCheck.HTMLToolTip With {.IsBalloon = True, .ToolTipTitle = "Error starting dictate", .ToolTipIcon = ToolTipIcon.Error}
                    DictateToolTip.ToolTipOrientation = i00SpellCheck.HTMLToolTip.ToolTipOrientations.TopLeft
                    DictateToolTip.Image = Nothing
                    DictateToolTip.ShowHTML("<b>Error: </b>" & [Error].Message, TextBox, ToolTipLocation, 5000)
                End If

            End If
        End Sub

        Private Shared Sub CommitRecognition()
            If DictateCancel = False Then
                'write the spoken text
                If tmpText <> "" Then
                    ExistingText &= " " & tmpText
                End If
                TextBox.SelectedText = Trim(ExistingText) & " "
                DictateCancel = True
            End If
            TextBox = Nothing
        End Sub

        Private Shared Sub TextBox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox.LostFocus
            CancelRecognition()
        End Sub

        Private Shared Sub TextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox.TextChanged
            CancelRecognition()
        End Sub

        Private Shared Sub TextBox_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox.MouseDown
            CancelRecognition()
        End Sub

        Private Shared Sub DictateToolTip_TipClick(ByVal sender As Object, ByVal e As i00SpellCheck.HTMLToolTip.TipClickEventArgs) Handles DictateToolTip.TipClick
            CommitRecognition()
        End Sub

        Private Shared Sub DictateToolTip_TipClosed(ByVal sender As Object, ByVal e As System.EventArgs) Handles DictateToolTip.TipClosed
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

        Private Shared strYouSaid As String = "<i><font color=" & ListeningColor & ">You Said:</font></i> "

        Private Shared Sub Recogniser_SpeechHypothesized(ByVal sender As Object, ByVal e As System.Speech.Recognition.SpeechHypothesizedEventArgs) Handles Recogniser.SpeechHypothesized
            If DictateCancel Then Exit Sub
            tmpText = e.Result.Text
            DictateToolTip.ShowHTML(strYouSaid & ExistingText & "<font color=" & ListeningColor & ">" & tmpText & "</font>", TextBox, ToolTipLocation)
        End Sub

        Private Shared Sub Recogniser_SpeechRecognized(ByVal sender As Object, ByVal e As System.Speech.Recognition.SpeechRecognizedEventArgs) Handles Recogniser.SpeechRecognized
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

                If DictateToolTip IsNot Nothing Then
                    DictateToolTip.ShowHTML(strYouSaid & ExistingText, TextBox, ToolTipLocation, 3000)
                End If
            End If
        End Sub

#End Region

#Region "Synthesis"

        Friend Shared WithEvents Synthesizer As Speech.Synthesis.SpeechSynthesizer

        Friend Shared Sub CancelSynthesis()
            StopSpeeking()
            If SpeechEngineHelper.Synthesizer IsNot Nothing Then
                SpeechEngineHelper.Synthesizer.SpeakAsyncCancelAll()
            End If
        End Sub

        Friend Shared Sub DoSynthesis(ByVal TextBox As TextBoxBase)
            If SpeechEngineHelper.Synthesizer Is Nothing Then
                SpeechEngineHelper.Synthesizer = New Speech.Synthesis.SpeechSynthesizer
            Else
                SpeechEngineHelper.Synthesizer.SpeakAsyncCancelAll()
            End If
            SpeechProgress = 0
            If TextBox.SelectionLength = 0 Then
                LastSpeechText = TextBox.Text
            Else
                LastSpeechText = TextBox.SelectedText
            End If
            SpeechEngineHelper.Synthesizer.SpeakAsync(LastSpeechText)
        End Sub

        Private Shared LastSpeechText As String
        Friend Shared SpeechProgress As Single

        Private Shared Sub Synthesizer_SpeakCompleted(ByVal sender As Object, ByVal e As System.Speech.Synthesis.SpeakCompletedEventArgs) Handles Synthesizer.SpeakCompleted
            StopSpeeking()
        End Sub

        Private Shared Sub StopSpeeking()
            SpeechProgress = 0
            If tsiSpeechProgress IsNot Nothing AndAlso tsiSpeechProgress.IsDisposed = False Then
                tsiStopSpeak.Visible = False
            End If
            If tsiSpeechProgress IsNot Nothing AndAlso tsiSpeechProgress.IsDisposed = False Then
                tsiSpeechProgress.Visible = False
            End If
        End Sub

        Private Shared Sub Synthesizer_SpeakProgress(ByVal sender As Object, ByVal e As System.Speech.Synthesis.SpeakProgressEventArgs) Handles Synthesizer.SpeakProgress
            SpeechProgress = CSng(e.CharacterPosition / LastSpeechText.Length)
            If tsiSpeechProgress IsNot Nothing AndAlso tsiSpeechProgress.IsDisposed = False Then
                tsiSpeechProgress.Progress = SpeechProgress
            End If
            ''estimate time remaining
            'If SpeechProgress <> 0 Then
            '    Dim ts = New TimeSpan(CLng(TimeSpan.TicksPerMillisecond * e.AudioPosition.TotalMilliseconds * (1 / SpeechProgress)))
            '    Debug.Print(e.AudioPosition.ToString & " / " & ts.ToString)
            'End If
        End Sub

        Public Shared tsiStopSpeak As i00SpellCheck.SpellCheckTextBox.StandardContextMenuStrip.StandardToolStripMenuItem
        Public Shared tsiSpeechProgress As SpellCheckTextBoxSpeechRecognition.tsiStandardProgress

#End Region

    End Class

End Class
