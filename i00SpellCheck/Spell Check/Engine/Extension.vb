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

#Region "Control.SpellCheck"

Public Module SpellCheckControlExtension

    'store a list of the SpellCheckControlBase items that are associated with each control
    Public SpellCheckControls As New Dictionary(Of Control, SpellCheckControlBase)

    Public Class SpellCheckControlAddedRemovedEventArgs
        Inherits EventArgs
        Public Control As Control
    End Class

    Public Class SpellCheckControlAddingEventArgs
        Inherits EventArgs
        Public Cancel As Boolean
    End Class

    Public Event SpellCheckControlAdded(ByVal sender As Object, ByVal e As SpellCheckControlAddedRemovedEventArgs)
    Public Event SpellCheckControlAdding(ByVal sender As Object, ByVal e As SpellCheckControlAddingEventArgs)
    Public Event SpellCheckControlRemoved(ByVal sender As Object, ByVal e As SpellCheckControlAddedRemovedEventArgs)

    <System.Runtime.CompilerServices.Extension()> _
    Public Function SpellCheck(ByVal sender As Control, Optional ByVal AutoCreate As Boolean = True, Optional ByVal SpellCheckSettings As SpellCheckSettings = Nothing) As SpellCheckControlBase
        If SpellCheckControls.ContainsKey(sender) Then
            'exists

        Else
            If AutoCreate Then
                ''create SpellCheckControlBase object and send it back...

                Static plugins As List(Of SpellCheckControlBase) = PluginManager(Of SpellCheckControlBase).GetPlugins

                Dim AcceptedClass = (From xItem In plugins Where xItem.ControlType.IsAssignableFrom(sender.GetType)).FirstOrDefault
                If AcceptedClass IsNot Nothing Then
                    Dim SpellCheckControlAddingEventArgs As New SpellCheckControlAddingEventArgs
                    RaiseEvent SpellCheckControlAdding(Nothing, SpellCheckControlAddingEventArgs)

                    If SpellCheckControlAddingEventArgs.Cancel = True Then
                        Return Nothing
                    End If
                    'create a new instance of the plugin class
                    Dim o = DirectCast(System.Activator.CreateInstance(AcceptedClass.GetType), SpellCheckControlBase) 'TryCast(AcceptedClass.CreateObject, SpellCheckControlBase)
                    If o IsNot Nothing Then
                        SpellCheckControls.Add(sender, o)
                        AddHandler sender.Disposed, AddressOf Control_Disposed
                        o.mc_Control = sender
                        o.DoLoad()
                        RaiseEvent SpellCheckControlAdded(Nothing, New SpellCheckControlAddedRemovedEventArgs With {.Control = sender})
                    Else
                        Return Nothing
                    End If
                Else
                    'no plugins for this control type
                    Return Nothing
                End If
            Else
                'we don't want to automatically this control to check spelling, and it is not enabled already so return nothing
                Return Nothing
            End If
        End If
        If SpellCheckSettings IsNot Nothing Then
            SpellCheckControls(sender).Settings = SpellCheckSettings
        End If
        Return SpellCheckControls(sender)
    End Function

    Private Sub Control_Disposed(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ctlSender = TryCast(sender, Control)
        If ctlSender IsNot Nothing Then
            RaiseEvent SpellCheckControlRemoved(Nothing, New SpellCheckControlAddedRemovedEventArgs With {.Control = ctlSender})
            SpellCheckControls.Remove(ctlSender)
            DisabledSpellCheckControls.Remove(ctlSender)
        End If
    End Sub

End Module

#End Region

#Region "Control.EnableSpellCheck / Control.AutoSpellCheckSettings"

Public Module SpellCheckFormExtension

#Region "Control.AutoSpellCheckSettings"

    'store a list of the SpellCheckControl items that are associated with each control
    Private ControlSpellCheckSettings As New Dictionary(Of Control, i00SpellCheck.SpellCheckSettings)

    <System.Runtime.CompilerServices.Extension()> _
    Public Function AutoSpellCheckSettings(ByVal sender As Control, Optional ByVal SpellCheckSettings As i00SpellCheck.SpellCheckSettings = Nothing) As i00SpellCheck.SpellCheckSettings
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
    Public Sub EnableSpellCheck(ByVal sender As Control, Optional ByVal SpellCheckSettings As i00SpellCheck.SpellCheckSettings = Nothing)
        If sender.SpellCheck IsNot Nothing Then
            If DisabledSpellCheckControls.Contains(sender) Then
                DisabledSpellCheckControls.Remove(sender)
            End If
            If SpellCheckSettings IsNot Nothing Then
                sender.SpellCheck.Settings = SpellCheckSettings
            End If
            sender.SpellCheck.RepaintControl()
        Else
            If SpellCheckSettings Is Nothing Then
                SpellCheckSettings = New i00SpellCheck.SpellCheckSettings
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
    Public Sub DisableSpellCheck(ByVal sender As Control)
        If SpellCheckControls.ContainsKey(sender) Then
            If DisabledSpellCheckControls.Contains(sender) Then
                DisabledSpellCheckControls.Remove(sender)
            Else
                DisabledSpellCheckControls.Add(sender)
            End If
            sender.SpellCheck.RepaintControl()
        End If
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsSpellCheckEnabled(ByVal sender As Control) As Boolean
        Return SpellCheckControls.ContainsKey(sender) AndAlso DisabledSpellCheckControls.Contains(sender) = False
    End Function
#End Region

#Region "DisableSpellCheck"

    Friend DisabledSpellCheckControls As New List(Of Control)

#End Region

#Region "Load dictionary/Definitions"

    Private Delegate Sub ReloadSpellCheckControls_cb(ByVal SpellCheckSettings As i00SpellCheck.SpellCheckSettings)
    Private Sub ReloadSpellCheckControls(ByVal SpellCheckSettings As i00SpellCheck.SpellCheckSettings)
        If SpellCheckSettings.MasterControl.InvokeRequired Then
            Dim ReloadSpellCheckControls_cb As New ReloadSpellCheckControls_cb(AddressOf ReloadSpellCheckControls)
            SpellCheckSettings.MasterControl.Invoke(ReloadSpellCheckControls_cb, SpellCheckSettings)
        Else
            SetupControl(SpellCheckSettings)
            RaiseEvent DictionaryLoaded(Nothing, EventArgs.Empty)

            For Each item In SpellCheckControls
                item.Value.RepaintControl()
                'old method:
                'item.Value.CustomPaint()
            Next
        End If
    End Sub

    Private Sub LoadDictionary(ByVal oSpellCheckSettings As Object)
        'use DirectCast so that error is thrown if wrong object is passed in
        Dim SpellCheckSettings = DirectCast(oSpellCheckSettings, i00SpellCheck.SpellCheckSettings)

        If Dictionary.DefaultDictionary Is Nothing Then
            'load flat file as default dictionary
            FlatFileDictionary.LoadDefaultDictionary()
        End If

        'qwertyuiop - check SpellCheckSettings.MasterControl.IsHandleCreated = true here?
        'force create the handle
        'Dim hdl = SpellCheckSettings.MasterControl.Handle

        ReloadSpellCheckControls(SpellCheckSettings)
    End Sub

#End Region

#Region "Owned Forms Poller"


    Private Class OwnedFormsPoll
        Public Shared OwnedFormsPolls As New Dictionary(Of Control, OwnedFormsPoll)

        'Private WithEvents MasterForm As Form
        Private LastOwnedForms As Form()
        Private tt As System.Threading.Timer
        Private SpellCheckSettings As i00SpellCheck.SpellCheckSettings

        Friend Shared Sub CreateOwnedFormsPoll(ByVal SpellCheckSettings As i00SpellCheck.SpellCheckSettings)
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

        Private Sub New(ByVal SpellCheckSettings As i00SpellCheck.SpellCheckSettings)
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
                SetupControl(i00SpellCheck.SpellCheckSettings.NewClone(NewForm, SpellCheckSettings))
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
                Debug.Print("Closing - " & SpellCheckSettings.MasterControl.GetType.ToString)
                OwnedFormsPolls.Remove(SpellCheckSettings.MasterControl)
                If OwnedFormsPolls.Count = 0 Then
                    DictionaryPerformanceCounter.Remove()
                End If
            End If
        End Sub
    End Class

#End Region

#Region "Configure controls"

    Private Sub SetupControl(ByVal SpellCheckSettings As i00SpellCheck.SpellCheckSettings)
        If TypeOf SpellCheckSettings.MasterControl Is Form AndAlso SpellCheckSettings.DoSubforms = True Then
            'Add .ownedforms change poll here
            OwnedFormsPoll.CreateOwnedFormsPoll(SpellCheckSettings)
        End If

        SpellCheckSettings.MasterControl.AutoSpellCheckSettings(SpellCheckSettings)

        If TypeOf SpellCheckSettings.MasterControl Is SplitContainer Then
            Dim SplitContainer = DirectCast(SpellCheckSettings.MasterControl, SplitContainer)
            SetupControl(i00SpellCheck.SpellCheckSettings.NewClone(SplitContainer.Panel1, SpellCheckSettings))
            SetupControl(i00SpellCheck.SpellCheckSettings.NewClone(SplitContainer.Panel2, SpellCheckSettings))
        ElseIf TypeOf SpellCheckSettings.MasterControl Is PropertyGrid Then
            'don't spell check property grids
            Return
        Else
            SpellCheckSettings.MasterControl.SpellCheck(, SpellCheckSettings)
        End If

        RemoveHandler SpellCheckSettings.MasterControl.ControlAdded, AddressOf Control_ControlAdded
        AddHandler SpellCheckSettings.MasterControl.ControlAdded, AddressOf Control_ControlAdded

        For Each ctl In SpellCheckSettings.MasterControl.Controls.OfType(Of Control)()
            SetupControl(i00SpellCheck.SpellCheckSettings.NewClone(ctl, SpellCheckSettings))
        Next
    End Sub

    Private Sub Control_ControlAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.ControlEventArgs)
        Dim senderControl = TryCast(sender, Control)
        If senderControl IsNot Nothing Then
            If OwnedFormsPoll.OwnedFormsPolls.ContainsKey(senderControl) Then
                'don't add it this control is a form that has been added through form polling!
            Else
                SetupControl(i00SpellCheck.SpellCheckSettings.NewClone(e.Control, senderControl.AutoSpellCheckSettings))
            End If
        End If
    End Sub

#End Region

End Module

#End Region