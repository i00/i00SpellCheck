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

#Region "TextBox.SpellCheck"

Public Module SpellCheckTextBoxExtension

    'store a list of the SpellCheckTextBox items that are associated with each text box
    Public SpellCheckTextBoxes As New Dictionary(Of TextBoxBase, SpellCheckTextBox)

    Public Class TextBoxAddedRemovedEventArgs
        Inherits EventArgs
        Public TextBoxBase As TextBoxBase
    End Class

    Public Event SpellCheckTextBoxAdded(ByVal sender As Object, ByVal e As TextBoxAddedRemovedEventArgs)
    Public Event SpellCheckTextBoxRemoved(ByVal sender As Object, ByVal e As TextBoxAddedRemovedEventArgs)

    <System.Runtime.CompilerServices.Extension()> _
    Public Function SpellCheck(ByVal sender As TextBoxBase, Optional ByVal AutoCreate As Boolean = True) As SpellCheckTextBox
        If SpellCheckTextBoxes.ContainsKey(sender) Then
            'exists
        Else
            If AutoCreate Then
                'create SpellCheckTextBox object and send it back...
                Dim SpellCheckTextBox = New SpellCheckTextBox(sender)
                AddHandler sender.Disposed, AddressOf TextBox_Disposed
                SpellCheckTextBoxes.Add(sender, SpellCheckTextBox)
                RaiseEvent SpellCheckTextBoxAdded(Nothing, New TextBoxAddedRemovedEventArgs With {.TextBoxBase = sender})
            Else
                Return Nothing
            End If
        End If
        Return SpellCheckTextBoxes(sender)
    End Function

    Private Sub TextBox_Disposed(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim txtSender = TryCast(sender, TextBoxBase)
        If txtSender IsNot Nothing Then
            RaiseEvent SpellCheckTextBoxRemoved(Nothing, New TextBoxAddedRemovedEventArgs With {.TextBoxBase = txtSender})
            SpellCheckTextBoxes.Remove(txtSender)
            DisabledSpellCheckTextBoxes.Remove(txtSender)
        End If
    End Sub

End Module

#End Region

#Region "Control.EnableSpellCheck / Control.AutoSpellCheckSettings"

Public Module SpellCheckFormExtension

#Region "Control.AutoSpellCheckSettings"

    'store a list of the SpellCheckTextBox items that are associated with each text box
    Private ControlSpellCheckSettings As New Dictionary(Of Control, SpellCheckTextBox.SpellCheckSettings)

    <System.Runtime.CompilerServices.Extension()> _
    Friend Function AutoSpellCheckSettings(ByVal sender As Control, Optional ByVal SpellCheckSettings As SpellCheckTextBox.SpellCheckSettings = Nothing) As SpellCheckTextBox.SpellCheckSettings
        If SpellCheckSettings IsNot Nothing Then
            If ControlSpellCheckSettings.ContainsKey(sender) Then
                'exists
            Else
                'create ControlSpellCheckSettings object and send it back...
                AddHandler sender.Disposed, AddressOf Control_Disposed
                ControlSpellCheckSettings.Add(sender, SpellCheckSettings)
            End If
        End If
        If ControlSpellCheckSettings.ContainsKey(sender) Then
            Return ControlSpellCheckSettings(sender)
        Else
            Return Nothing
        End If
    End Function

    Private Sub Control_Disposed(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim controlSender As Control = TryCast(sender, Control)
        If controlSender IsNot Nothing Then
            ControlSpellCheckSettings.Remove(controlSender)
        End If
    End Sub

#End Region

#Region "Extension"

    Public Event DictionaryLoaded(ByVal sender As Object, ByVal e As EventArgs)
    'Public Event DefinitionsLoaded(ByVal sender As Object, ByVal e As EventArgs)

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub EnableSpellCheck(ByVal sender As Control, Optional ByVal SpellCheckSettings As SpellCheckTextBox.SpellCheckSettings = Nothing)
        Dim tbb = TryCast(sender, TextBoxBase)
        If tbb IsNot Nothing AndAlso SpellCheckTextBoxes.ContainsKey(tbb) Then
            If DisabledSpellCheckTextBoxes.Contains(tbb) Then
                DisabledSpellCheckTextBoxes.Remove(tbb)
            End If
            If SpellCheckSettings IsNot Nothing Then
                'Dim SpellCheckSettingsClone As SpellCheckTextBox.SpellCheckSettings
                tbb.SpellCheck.Settings = SpellCheckSettings
            End If
            tbb.SpellCheck.RepaintTextBox()
        Else
            If SpellCheckSettings Is Nothing Then
                SpellCheckSettings = New SpellCheckTextBox.SpellCheckSettings
            End If
            SpellCheckSettings.MasterControl = sender
            'load default dictionary... on a thread ... so we don't hold things up...
            'after the dictionary is loaded it automatically invokes the rest...
            Dim t As New Threading.Thread(AddressOf LoadDictionary)
            t.Name = "Load spell check dictionary"

            t.IsBackground = True 'make it abort when the app is killed
            t.Start(SpellCheckSettings)

            'Dim t2 As New Threading.Thread(AddressOf LoadDefs)
            't2.Name = "Load spell check definitions"
            't2.IsBackground = True
            't2.Start(SpellCheckSettings)
        End If
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub DisableSpellCheck(ByVal sender As TextBoxBase)
        If SpellCheckTextBoxes.ContainsKey(sender) Then
            If DisabledSpellCheckTextBoxes.Contains(sender) Then
                DisabledSpellCheckTextBoxes.Remove(sender)
            Else
                DisabledSpellCheckTextBoxes.Add(sender)
            End If
            sender.SpellCheck.RepaintTextBox()
        End If
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsSpellCheckEnabled(ByVal sender As TextBoxBase) As Boolean
        Return SpellCheckTextBoxes.ContainsKey(sender) AndAlso DisabledSpellCheckTextBoxes.Contains(sender) = False
    End Function
#End Region

#Region "DisableSpellCheck"

    Friend DisabledSpellCheckTextBoxes As New List(Of TextBoxBase)

#End Region

#Region "Load dictionary/Definitions"

    Private Delegate Sub ReloadTextBoxes_cb(ByVal SpellCheckSettings As SpellCheckTextBox.SpellCheckSettings)
    Private Sub ReloadTextBoxes(ByVal SpellCheckSettings As SpellCheckTextBox.SpellCheckSettings)
        If SpellCheckSettings.MasterControl.InvokeRequired Then
            Dim ReloadTextBoxes_cb As New ReloadTextBoxes_cb(AddressOf ReloadTextBoxes)
            SpellCheckSettings.MasterControl.Invoke(ReloadTextBoxes_cb, SpellCheckSettings)
        Else
            SetupControl(SpellCheckSettings)
            RaiseEvent DictionaryLoaded(Nothing, EventArgs.Empty)
            For Each item In SpellCheckTextBoxes
                item.Value.RepaintTextBox()
                'old method:
                'item.Value.CustomPaint()
            Next
        End If
    End Sub

    Private Sub LoadDictionary(ByVal oSpellCheckSettings As Object)
        'use DirectCast so that error is thrown if wrong object is passed in
        Dim SpellCheckSettings = DirectCast(oSpellCheckSettings, SpellCheckTextBox.SpellCheckSettings)

        SpellCheckTextBox.LoadDefaultDictionary()

        'qwertyuiop - check SpellCheckSettings.MasterControl.IsHandleCreated = true here?
        'force create the handle
        'Dim hdl = SpellCheckSettings.MasterControl.Handle

        ReloadTextBoxes(SpellCheckSettings)
    End Sub

    'Private Sub LoadDefs(ByVal oSpellCheckSettings As Object)
    '    Dim DefaultDefs = SpellCheckTextBox.Definitions.DefaultDefs
    '    RaiseEvent DefinitionsLoaded(Nothing, EventArgs.Empty)
    'End Sub

#End Region

#Region "Owned Forms Poller"

    'store a list of the SpellCheckTextBox items that are associated with each text box

    Private Class OwnedFormsPoll
        Public Shared OwnedFormsPolls As New Dictionary(Of Control, OwnedFormsPoll)

        'Private WithEvents MasterForm As Form
        Private LastOwnedForms As Form()
        Private tt As System.Threading.Timer
        Private SpellCheckSettings As SpellCheckTextBox.SpellCheckSettings

        Friend Shared Sub CreateOwnedFormsPoll(ByVal SpellCheckSettings As SpellCheckTextBox.SpellCheckSettings)
            If TypeOf SpellCheckSettings.MasterControl Is Form Then
                If OwnedFormsPolls.ContainsKey(SpellCheckSettings.MasterControl) Then
                    'already exists
                Else
                    'add it
                    Dim ofp As New OwnedFormsPoll(SpellCheckSettings)
                    OwnedFormsPolls.Add(SpellCheckSettings.MasterControl, ofp)
                End If
            End If
        End Sub

        Private Sub New(ByVal SpellCheckSettings As SpellCheckTextBox.SpellCheckSettings)
            LastOwnedForms = CType(SpellCheckSettings.MasterControl, Form).OwnedForms
            AddHandler SpellCheckSettings.MasterControl.Disposed, AddressOf Form_Disposed
            Me.SpellCheckSettings = SpellCheckSettings
            Dim cb As System.Threading.TimerCallback = AddressOf CheckNewForms
            tt = New System.Threading.Timer(cb, Nothing, 1000, 1000)
        End Sub


        Private Delegate Sub cb_SetupNewControl(ByVal NewForm As Form)
        Private Sub SetupNewControl(ByVal NewForm As Form)
            If NewForm.IsDisposed Then
                Exit Sub
            End If
            If NewForm.InvokeRequired Then
                Try
                    'form could have been disposed between above ^ and now
                    Dim cb_SetupNewControl As New cb_SetupNewControl(AddressOf SetupNewControl)
                    NewForm.Invoke(cb_SetupNewControl, NewForm)
                Catch ex As Exception

                End Try
            Else
                Debug.Print(SpellCheckSettings.MasterControl.GetType.ToString & " - is opening - " & NewForm.GetType.ToString)
                SetupControl(SpellCheckTextBox.SpellCheckSettings.NewClone(NewForm, SpellCheckSettings))
            End If
        End Sub

        Private Sub CheckNewForms(ByVal state As Object)
            Dim NewForms = (From xItem In CType(SpellCheckSettings.MasterControl, Form).OwnedForms Where LastOwnedForms.Contains(xItem) = False).ToArray
            If NewForms.Count > 0 Then
                For Each item In NewForms
                    SetupNewControl(item)
                Next
            End If
            LastOwnedForms = CType(SpellCheckSettings.MasterControl, Form).OwnedForms
        End Sub

        Private Sub Form_Disposed(ByVal sender As Object, ByVal e As System.EventArgs)
            tt.Dispose()
            If OwnedFormsPolls.ContainsKey(SpellCheckSettings.MasterControl) Then
                OwnedFormsPolls.Remove(SpellCheckSettings.MasterControl)
            End If
        End Sub
    End Class

#End Region

#Region "Configure controls"

    Private Sub SetupControl(ByVal SpellCheckSettings As SpellCheckTextBox.SpellCheckSettings)
        If TypeOf SpellCheckSettings.MasterControl Is Form AndAlso SpellCheckSettings.DoSubforms = True Then
            'Add .ownedforms change poll here
            OwnedFormsPoll.CreateOwnedFormsPoll(SpellCheckSettings)
        End If

        SpellCheckSettings.MasterControl.AutoSpellCheckSettings(SpellCheckSettings)

        If TypeOf SpellCheckSettings.MasterControl Is TextBoxBase Then
            Dim TextBox = DirectCast(SpellCheckSettings.MasterControl, TextBoxBase)
            RemoveHandler TextBox.MultilineChanged, AddressOf TextBox_MultilineChanged
            AddHandler TextBox.MultilineChanged, AddressOf TextBox_MultilineChanged
            If TextBox.Multiline = True Then
                TextBox.SpellCheck()
                ApplyTextBoxSpellingSettings(SpellCheckSettings)
            End If
        ElseIf TypeOf SpellCheckSettings.MasterControl Is SplitContainer Then
            Dim SplitContainer = DirectCast(SpellCheckSettings.MasterControl, SplitContainer)
            SetupControl(SpellCheckTextBox.SpellCheckSettings.NewClone(SplitContainer.Panel1, SpellCheckSettings))
            SetupControl(SpellCheckTextBox.SpellCheckSettings.NewClone(SplitContainer.Panel2, SpellCheckSettings))
        End If
        RemoveHandler SpellCheckSettings.MasterControl.ControlAdded, AddressOf Control_ControlAdded
        AddHandler SpellCheckSettings.MasterControl.ControlAdded, AddressOf Control_ControlAdded

        For Each ctl In SpellCheckSettings.MasterControl.Controls.OfType(Of Control)()
            SetupControl(SpellCheckTextBox.SpellCheckSettings.NewClone(ctl, SpellCheckSettings))
        Next
    End Sub

    Private Sub ApplyTextBoxSpellingSettings(ByVal AutoSpellCheckSettings As SpellCheckTextBox.SpellCheckSettings)
        If AutoSpellCheckSettings IsNot Nothing AndAlso TypeOf AutoSpellCheckSettings.MasterControl Is TextBoxBase Then
            Dim TextBox As TextBoxBase = DirectCast(AutoSpellCheckSettings.MasterControl, TextBoxBase)
            TextBox.SpellCheck.Settings.AllowAdditions = AutoSpellCheckSettings.AllowAdditions
            TextBox.SpellCheck.Settings.AllowIgnore = AutoSpellCheckSettings.AllowIgnore
            TextBox.SpellCheck.Settings.AllowRemovals = AutoSpellCheckSettings.AllowRemovals
            TextBox.SpellCheck.Settings.AllowInMenuDefs = AutoSpellCheckSettings.AllowInMenuDefs
            TextBox.SpellCheck.Settings.AllowChangeTo = AutoSpellCheckSettings.AllowChangeTo
        End If
    End Sub

    Private Sub TextBox_MultilineChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim senderTextBox = TryCast(sender, TextBox)
        If senderTextBox IsNot Nothing Then
            If senderTextBox.Multiline = True Then
                senderTextBox.SpellCheck()
                ApplyTextBoxSpellingSettings(senderTextBox.AutoSpellCheckSettings)
            End If
        End If
    End Sub

    Private Sub Control_ControlAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.ControlEventArgs)
        Dim senderControl = TryCast(sender, Control)
        If senderControl IsNot Nothing Then
            If OwnedFormsPoll.OwnedFormsPolls.ContainsKey(senderControl) Then
                'don't add it this control is a form that has been added through form polling!
            Else
                SetupControl(SpellCheckTextBox.SpellCheckSettings.NewClone(e.Control, senderControl.AutoSpellCheckSettings))
            End If
        End If
    End Sub

#End Region

End Module

#End Region