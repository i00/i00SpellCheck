﻿'i00 .Net Spell Check
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

Partial Public Class SpellCheckLabel
    Inherits i00SpellCheck.SpellCheckControlBase

#Region "Setup"

    Public Shared CompatiblityWarningDisplayed As Boolean

    'Called when the control is loaded
    Overrides Sub Load()
        Dim Label = TryCast(Me.Control, Label)
        If Label IsNot Nothing Then
            'Add control specific event handlers
            If Label.UseCompatibleTextRendering Then
                If CompatiblityWarningDisplayed = False Then
                    Dim a = System.Reflection.Assembly.GetEntryAssembly
                    Dim myAppType = a.GetTypes().Single(Function(t) t.Name = "MyApplication")
                    Dim ReasonForCompatiblityWarning As String = ""
                    If myAppType.GetProperty("ApplicationContext") Is Nothing Then
                        ReasonForCompatiblityWarning = "this appears to be because the application framework is disabled, you can call Application.SetCompatibleTextRenderingDefault(False) before the first form loads to resolve this, "
                    End If
                    Debug.Print("i00 Spell Check Label - does not work with UseCompatibleTextRendering set to True, " & ReasonForCompatiblityWarning & "controls will automatically be updated...")
                    CompatiblityWarningDisplayed = True
                End If
                Label.UseCompatibleTextRendering = False
            End If
            AddHandler Label.Paint, AddressOf Label_Paint
            RepaintControl()
        End If
    End Sub

    'Lets the EnableSpellCheck() know what ControlTypes we can spellcheck
    Public Overrides ReadOnly Property ControlType() As System.Type
        Get
            Return GetType(Label)
        End Get
    End Property

    'Repaint control when settings are changed
    Private Sub SpellCheckLabel_SettingsChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SettingsChanged
        RepaintControl()
    End Sub

#End Region

#Region "Underlying Control"

    'Make underlying control appear nicer in property grids
    <System.ComponentModel.Category("Control")> _
    <System.ComponentModel.Description("The Label associated with the SpellCheckLabel object")> _
    <System.ComponentModel.DisplayName("Label")> _
    Public Overrides ReadOnly Property Control() As System.Windows.Forms.Control
        Get
            Return MyBase.Control
        End Get
    End Property

    'Quick access to underlying Label object
    Private ReadOnly Property parentLabel() As Label
        Get
            Return TryCast(MyBase.Control, Label)
        End Get
    End Property

#End Region

#Region "Painting"

    Public Overrides Sub RepaintControl()
        parentLabel.Invalidate()
    End Sub

    Private Sub Label_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)
        If CurrentDictionary IsNot Nothing AndAlso (CurrentDictionary.Loading = True OrElse CurrentDictionary.Count = 0) Then Exit Sub
        If parentLabel.IsSpellCheckEnabled Then
            Dim oSender = TryCast(sender, Label)
            If oSender IsNot Nothing Then

                Dim SpellCheckSettings = oSender.AutoSpellCheckSettings
                If SpellCheckSettings IsNot Nothing AndAlso (SpellCheckSettings.ShowMistakes) Then

                    Dim NewWords As New Dictionary(Of String, Dictionary.SpellCheckWordError)

                    If oSender.UseCompatibleTextRendering Then Exit Sub 'not supported with CompatibleTextRendering :(

                    e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

                    Dim Flags As TextFormatFlags
                    If oSender.UseMnemonic = False Then
                        'for the Mnemonic's ... eg. &Help to underline the H
                        Flags = Flags Or TextFormatFlags.NoPrefix
                    End If
                    Flags = Flags Or TextFormatFlags.WordBreak

                    Dim FalseFlags = Flags
                    Flags = Flags Or TextFormatFlags.Top
                    Flags = Flags Or TextFormatFlags.Left

                    Select Case oSender.TextAlign
                        Case ContentAlignment.BottomCenter
                            Flags = Flags Or TextFormatFlags.Bottom
                            Flags = Flags Or TextFormatFlags.HorizontalCenter
                        Case ContentAlignment.BottomLeft
                            Flags = Flags Or TextFormatFlags.Bottom
                            Flags = Flags Or TextFormatFlags.Left
                        Case ContentAlignment.BottomRight
                            Flags = Flags Or TextFormatFlags.Bottom
                            Flags = Flags Or TextFormatFlags.Right
                        Case ContentAlignment.MiddleCenter
                            Flags = Flags Or TextFormatFlags.VerticalCenter
                            Flags = Flags Or TextFormatFlags.HorizontalCenter
                        Case ContentAlignment.MiddleLeft
                            Flags = Flags Or TextFormatFlags.VerticalCenter
                            Flags = Flags Or TextFormatFlags.Left
                        Case ContentAlignment.MiddleRight
                            Flags = Flags Or TextFormatFlags.VerticalCenter
                            Flags = Flags Or TextFormatFlags.Right
                        Case ContentAlignment.TopCenter
                            Flags = Flags Or TextFormatFlags.Top
                            Flags = Flags Or TextFormatFlags.HorizontalCenter
                        Case ContentAlignment.TopLeft
                            Flags = Flags Or TextFormatFlags.Top
                            Flags = Flags Or TextFormatFlags.Left
                        Case ContentAlignment.TopRight
                            Flags = Flags Or TextFormatFlags.Top
                            Flags = Flags Or TextFormatFlags.Right
                    End Select

                    'spell check bit :D
                    'changed words makes words like "kris'" > "kris'"
                    Dim ChangedWordText = Replace(oSender.Text, vbCrLf, vbCr)
                    ChangedWordText = Replace(ChangedWordText, vbLf, vbCr)
                    Dim ChangedWords = Dictionary.Formatting.RemoveWordBreaks(ChangedWordText)

                    Dim TextRectSize = New Size(oSender.ClientRectangle.Width - (oSender.Padding.Left + oSender.Padding.Right), oSender.ClientRectangle.Height - (oSender.Padding.Top + oSender.Padding.Bottom))
                    e.Graphics.TranslateTransform(oSender.Padding.Left, oSender.Padding.Top)

                    Dim WordBounds = DrawingFunctions.Text.TextRendererMeasure.Measure(oSender.Text, oSender.Font, TextRectSize, Flags)
                    
                    Dim NewWordBounds As New DrawingFunctions.Text.TextRendererMeasure.WordBounds With {.TextMargin = WordBounds.TextMargin}

                    For Each item In WordBounds
                        'use the changed word instead :)
                        item.Word = ChangedWords.Substring(item.LetterIndex, item.Word.Length).Trim

                        If item.Word.Contains(" ") Then
                            'need to break this word into parts :(

                            Dim NewSubWordBounds = DrawingFunctions.Text.TextRendererMeasure.Measure(item.Word, oSender.Font, New Size(Integer.MaxValue, Integer.MaxValue), FalseFlags)
                            For Each iNewSubWordBounds In NewSubWordBounds
                                iNewSubWordBounds.Bounds.X += item.Bounds.X - WordBounds.TextMargin
                                iNewSubWordBounds.Bounds.Y += item.Bounds.Y
                            Next
                            NewWordBounds.AddRange(NewSubWordBounds)
                        Else
                            'only one word here
                            NewWordBounds.Add(item)
                        End If

                    Next

                    For Each item In NewWordBounds

                        Dim WordState As Dictionary.SpellCheckWordError = Dictionary.SpellCheckWordError.SpellError
                        If dictCache.ContainsKey(item.Word) Then
                            'load from cache
                            WordState = dictCache(item.Word)
                        Else
                            If NewWords.ContainsKey(item.Word) = False Then
                                NewWords.Add(item.Word, Dictionary.SpellCheckWordError.OK)
                            End If

                            WordState = Dictionary.SpellCheckWordError.OK
                        End If
                        If item.Bounds.Bottom >= oSender.ClientSize.Height - oSender.Padding.Top Then
                            item.Bounds.Height = item.Bounds.Height - 1 'oSender.ClientSize.Height - item.Bounds.Top - 1
                        End If

                        'for custom drawing...
                        Dim eArgs = New SpellCheckCustomPaintEventArgs With {.Graphics = e.Graphics, .Word = item.Word, .Bounds = item.Bounds, .WordState = WordState}
                        OnSpellCheckErrorPaint(eArgs)
                        If eArgs.DrawDefault Then
                            Select Case WordState
                                Case Dictionary.SpellCheckWordError.Ignore
                                    If DrawIgnored() Then
                                        Using p As New Pen(Settings.IgnoreColor)
                                            e.Graphics.DrawLine(p, item.Bounds.X, item.Bounds.Bottom + 1, item.Bounds.Right, item.Bounds.Bottom + 1)
                                        End Using
                                    End If
                                Case Dictionary.SpellCheckWordError.CaseError
                                    DrawingFunctions.DrawWave(e.Graphics, New Point(CInt(item.Bounds.X), CInt(item.Bounds.Bottom)), New Point(CInt(item.Bounds.Right), CInt(item.Bounds.Bottom)), Settings.CaseMistakeColor)
                                Case Dictionary.SpellCheckWordError.SpellError
                                    DrawingFunctions.DrawWave(e.Graphics, New Point(CInt(item.Bounds.X), CInt(item.Bounds.Bottom)), New Point(CInt(item.Bounds.Right), CInt(item.Bounds.Bottom)), Settings.MistakeColor)
                            End Select
                        End If

                        'to see the text spacing...
                        'TextRenderer.DrawText(e.Graphics, item.Word, oSender.Font, New Point(item.Bounds.X + oSender.Padding.Left - WordBounds.TextMargin, item.Bounds.Y + oSender.Padding.Top), Color.Pink)
                    Next
                    e.Graphics.ResetTransform()

                    'end spell check bit :D

                    'To see the whole text rendered..
                    'TextRenderer.DrawText(e.Graphics, oSender.Text, oSender.Font, New Rectangle(New Point(oSender.Padding.Left, oSender.Padding.Top), TextRectSize), Color.Pink, Flags)
                    'e.Graphics.DrawRectangle(Pens.Pink, New Rectangle(New Point(), TextRectSize))

                    If NewWords.Count > 0 Then
                        AddWordsToCache(NewWords)
                    End If

                End If
            End If
        End If


    End Sub

#End Region

End Class