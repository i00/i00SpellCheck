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

Public MustInherit Class SpellCheckControlBase
    Inherits NativeWindow
    Implements IMessageFilter

#Region "Ignore"

#Region "Update the WndProc Handle when the control gets a new handle - for when properties like RightToLeft are changed"

    Private Sub mc_Control_HandleCreated(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_Control.HandleCreated
        AssignHandle(Control.Handle)
    End Sub

#End Region

#Region "IMessageFilter - For Keypress to Make Ignore Underline Appear"

    Const WM_KEYDOWN As Integer = &H100
    Const WM_KEYUP As Integer = &H101

    Public Function PreFilterMessage(ByRef m As System.Windows.Forms.Message) As Boolean Implements System.Windows.Forms.IMessageFilter.PreFilterMessage
        Static ctlPressed As Boolean
        Select Case m.Msg
            Case WM_KEYDOWN
                Select Case m.WParam.ToInt32
                    Case Keys.ControlKey
                        If Settings.ShowIgnored = SpellCheckSettings.ShowIgnoreState.OnKeyDown AndAlso ctlPressed = False Then
                            ctlPressed = True
                            RepaintControl()
                        End If
                End Select
            Case WM_KEYUP
                Select Case m.WParam.ToInt32
                    Case Keys.ControlKey
                        If Settings.ShowIgnored = SpellCheckSettings.ShowIgnoreState.OnKeyDown AndAlso ctlPressed = True Then
                            ctlPressed = False
                            RepaintControl()
                        End If
                End Select
        End Select
    End Function

#End Region

    Protected Function DrawIgnored() As Boolean
        If Settings.ShowIgnored = SpellCheckSettings.ShowIgnoreState.OnKeyDown AndAlso My.Computer.Keyboard.CtrlKeyDown Then
            Return True
        ElseIf Settings.ShowIgnored = SpellCheckSettings.ShowIgnoreState.AlwaysShow Then
            Return True
        End If
    End Function

#End Region

    Public Event DictionaryChanged(ByVal sender As Object, ByVal e As EventArgs)
    Dim mc_CurrentDictionary As Dictionary
    <System.ComponentModel.Editor(GetType(Dictionary.dicFile_UITypeEditor), GetType(System.Drawing.Design.UITypeEditor))> _
    <System.ComponentModel.Description("The current items in the spell check dictionary" & vbCrLf & ".Save() must be called to commit changes to the underlying file")> _
    <System.ComponentModel.TypeConverter(GetType(CollectionToStringConverter))> _
    <System.ComponentModel.DisplayName("Dictionary")> _
    <System.ComponentModel.Category("Dictionaries")> _
    Public Property CurrentDictionary() As Dictionary
        Get
            If mc_CurrentDictionary Is Nothing Then
                'return the default :)
                If Dictionary.DefaultDictionary Is Nothing Then
                    'load flat file as default dictionary
                    FlatFileDictionary.LoadDefaultDictionary()
                End If
                Return Dictionary.DefaultDictionary
            Else
                Return mc_CurrentDictionary
            End If
        End Get
        Set(ByVal value As Dictionary)
            If mc_CurrentDictionary IsNot value Then
                mc_CurrentDictionary = value
                RaiseEvent DictionaryChanged(Me, EventArgs.Empty)
            End If
        End Set
    End Property

    Dim mc_CurrentDefinitions As Definitions = Nothing
    <System.ComponentModel.DisplayName("Definitions")> _
    <System.ComponentModel.Category("Dictionaries")> _
    Public Property CurrentDefinitions() As Definitions
        Get
            If mc_CurrentDefinitions Is Nothing Then
                Definitions.LoadDefaultDefinitions()
                mc_CurrentDefinitions = Definitions.DefaultDefinitions
            End If
            Return mc_CurrentDefinitions
        End Get
        Set(ByVal value As Definitions)
            mc_CurrentDefinitions = value
        End Set
    End Property

    Private Sub mc_Settings_Redraw(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_Settings.Redraw
        RaiseEvent SettingsChanged(Me, EventArgs.Empty)
    End Sub

    Protected ReadOnly Property OKToSpellCheck() As Boolean
        Get
            If Control.IsSpellCheckEnabled = False Then Return False
            If CurrentDictionary Is Nothing Then Return False
            If CurrentDictionary IsNot Nothing AndAlso (CurrentDictionary.Loading = True OrElse CurrentDictionary.Count = 0) Then Return False
            Return True
        End Get
    End Property

    Protected ReadOnly Property OKToDraw() As Boolean
        Get
            If OKToSpellCheck = False Then Return False
            If Settings.ShowMistakes = False Then Return False
            Return True
        End Get
    End Property

    Public Event SettingsChanged(ByVal sender As Object, ByVal e As EventArgs)
    Private WithEvents mc_Settings As SpellCheckSettings = New SpellCheckSettings
    Public Property Settings() As SpellCheckSettings
        Get
            If mc_Settings Is Nothing Then
                mc_Settings = New SpellCheckSettings
            End If
            Return mc_Settings
        End Get
        Set(ByVal value As SpellCheckSettings)
            If mc_Settings IsNot value Then
                mc_Settings = value
                RaiseEvent SettingsChanged(Me, EventArgs.Empty)
            End If
        End Set
    End Property

    Public MustOverride Sub RepaintControl()

    Public Class SpellCheckCustomPaintEventArgs
        Inherits EventArgs
        Public Bounds As Rectangle
        Public Graphics As Graphics
        Public DrawDefault As Boolean = True
        Public Word As String
        Public WordState As Dictionary.SpellCheckWordError
    End Class

    Public Event SpellCheckErrorPaint(ByVal sender As Object, ByRef e As SpellCheckCustomPaintEventArgs)

    Protected Sub OnSpellCheckErrorPaint(ByRef e As SpellCheckCustomPaintEventArgs)
        RaiseEvent SpellCheckErrorPaint(Me, e)
    End Sub

    Public Sub InvalidateAllControlsWithSameDict(Optional ByVal IncludeThis As Boolean = True)
        For Each iSellClassWithSameDict In (From xItem In AllControlsWithSameDict() Where (IncludeThis OrElse xItem IsNot Me))
            'remove item from cache
            iSellClassWithSameDict.RepaintControl()
        Next
    End Sub

    Public Function AllControlsWithSameDict() As IEnumerable(Of SpellCheckControlBase)
        Return (From xItem In SpellCheckControls Where xItem.Value.CurrentDictionary Is CurrentDictionary Select xItem.Value)
    End Function

    Dim mc_CurrentSynonyms As Synonyms = Nothing
    <System.ComponentModel.DisplayName("Synonyms")> _
    <System.ComponentModel.Category("Dictionaries")> _
    Public Property CurrentSynonyms() As Synonyms
        Get
            If mc_CurrentSynonyms Is Nothing Then
                Synonyms.LoadDefaultSynonyms()
                mc_CurrentSynonyms = Synonyms.DefaultSynonyms
            End If
            Return mc_CurrentSynonyms
        End Get
        Set(ByVal value As Synonyms)
            mc_CurrentSynonyms = value
        End Set
    End Property

    <System.ComponentModel.Category("Control")> _
    <System.ComponentModel.Description("The type of control that will be automatically spellchecked by this class")> _
    <System.ComponentModel.DisplayName("Control Type")> _
    Public MustOverride ReadOnly Property ControlType() As Type

    Protected Friend WithEvents mc_Control As Control
    <System.ComponentModel.Category("Control")> _
    <System.ComponentModel.Description("The Control associated with the SpellCheckControlBase object")> _
    <System.ComponentModel.DisplayName("Control")> _
    Public Overridable ReadOnly Property Control() As Control
        Get
            Return mc_Control
        End Get
    End Property

    Friend Sub DoLoad()
        'AssignHandle allows the WndProc to be fired :)
        Try
            AssignHandle(Control.Handle)
        Catch ex As Exception

        End Try
        Application.AddMessageFilter(Me)
        Load()
    End Sub

    MustOverride Sub Load()

    Public Interface iSpellCheckDialog
        Sub ShowDialog()
    End Interface


    Private Sub mc_Control_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles mc_Control.Disposed
        'kill all running threads
        AbortSpellCheckThreads()

        'remove the meesage filter...
        Application.RemoveMessageFilter(Me)
    End Sub

#Region "Dictionary Functions"

    Public Class DictionarySaveException
        Inherits Exception
        Public Sub New(ByVal ex As Exception)
            MyBase.New(ex.Message, ex.InnerException)
        End Sub
    End Class

    Public Sub DictionaryUnIgnoreWord(ByVal Word As String)
        CurrentDictionary.UnIgnore(Word)
        'for all items with the same dict
        For Each iSellClassWithSameDict In AllControlsWithSameDict()
            'remove item from cache
            For Each item In (From xItem In iSellClassWithSameDict.dictCache Where LCase(xItem.Key) = LCase(Word) OrElse LCase(xItem.Key) = LCase(Word) & "'s" OrElse LCase(xItem.Key) = LCase(Word) & Chr(146) & "s").ToArray
                iSellClassWithSameDict.dictCache.Remove(item.Key)
                iSellClassWithSameDict.dictCache.Add(item.Key, Dictionary.SpellCheckWordError.SpellError)
            Next
            iSellClassWithSameDict.RepaintControl()
        Next
        If CurrentDictionary.Filename <> "" Then
            Try
                CurrentDictionary.Save()
            Catch ex As Exception
                Throw New DictionarySaveException(ex)
            End Try
        End If
    End Sub

    Public Sub DictionaryIgnoreWord(ByVal Word As String)
        CurrentDictionary.Ignore(Word)
        'for all items with the same dict
        For Each iSellClassWithSameDict In AllControlsWithSameDict()
            'remove item from cache
            For Each item In (From xItem In iSellClassWithSameDict.dictCache Where LCase(xItem.Key) = LCase(Word) OrElse LCase(xItem.Key) = LCase(Word) & "'s" OrElse LCase(xItem.Key) = LCase(Word) & Chr(146) & "s").ToArray
                iSellClassWithSameDict.dictCache.Remove(item.Key)
                iSellClassWithSameDict.dictCache.Add(item.Key, Dictionary.SpellCheckWordError.Ignore)
            Next
            iSellClassWithSameDict.RepaintControl()
        Next
        If CurrentDictionary.Filename <> "" Then
            Try
                CurrentDictionary.Save()
            Catch ex As Exception
                Throw New DictionarySaveException(ex)
            End Try
        End If
    End Sub

    Public Sub DictionaryRemoveWord(ByVal Word As String)
        CurrentDictionary.Remove(Word)

        'for all items with the same dict
        For Each iSellClassWithSameDict In AllControlsWithSameDict()
            'remove item from cache
            For Each item In (From xItem In iSellClassWithSameDict.dictCache Where LCase(xItem.Key) = LCase(Word) OrElse LCase(xItem.Key) = LCase(Word) & "'s" OrElse LCase(xItem.Key) = LCase(Word) & Chr(146) & "s").ToArray
                dictCache.Remove(item.Key)
                dictCache.Add(item.Key, Dictionary.SpellCheckWordError.SpellError)
            Next
            iSellClassWithSameDict.RepaintControl()
        Next

        If CurrentDictionary.Filename <> "" Then
            Try
                CurrentDictionary.Save()
            Catch ex As Exception
                Throw New DictionarySaveException(ex)
            End Try
        End If
    End Sub

    Public Sub DictionaryAddWord(ByVal Word As String)
        CurrentDictionary.Add(Word)
        'for all items with the same dict
        For Each iSellClassWithSameDict In AllControlsWithSameDict()
            'remove item from cache
            For Each item In (From xItem In iSellClassWithSameDict.dictCache Where LCase(xItem.Key) = LCase(Word) OrElse LCase(xItem.Key) = LCase(Word) & "'s" OrElse LCase(xItem.Key) = LCase(Word) & Chr(146) & "s").ToArray
                iSellClassWithSameDict.dictCache.Remove(item.Key)
            Next
            iSellClassWithSameDict.RepaintControl()
        Next
        If CurrentDictionary.Filename <> "" Then
            Try
                CurrentDictionary.Save()
            Catch ex As Exception
                Throw New DictionarySaveException(ex)
            End Try
        End If
    End Sub

#End Region

#Region "Dictionary cache"

    Protected dictCache As New Dictionary(Of String, Dictionary.SpellCheckWordError)

    Private Sub AbortSpellCheckThreads()
        If AddWordsThread IsNot Nothing AndAlso AddWordsThread.IsAlive Then
            AddWordsThread.Abort()
        End If
    End Sub

    Public Sub AddWordsToCache(ByVal Dictionary As Dictionary(Of String, Dictionary.SpellCheckWordError), Optional ByVal Thread As Boolean = True)
        'qwertyuiop - SyncLock should not be needed here anymore since there is only one thread doing the checking for each control now
        '... but it doesnt work for some reason without it :S
        SyncLock WordsToCheck
            For Each item In Dictionary
                If WordsToCheck.ContainsKey(item.Key) = False Then
                    WordsToCheck.Add(item.Key, item.Value)
                End If
            Next
        End SyncLock
        If AddWordsThread Is Nothing Then
            AddWordsThread = New Threading.Thread(AddressOf mtAddWordsToCache)
            AddWordsThread.IsBackground = True
            AddWordsThread.Name = "Spell checking " & Me.Control.Name & "(" & Me.GetType.Name & " - " & Me.Control.GetType.Name & ")"
            AddWordsThread.Start()
        End If
    End Sub

    Dim WordsToCheck As New Dictionary(Of String, Dictionary.SpellCheckWordError)
    Dim AddWordsThread As Threading.Thread

    Private Sub mtAddWordsToCache(ByVal oDictionary_String_SpellCheckWordError As Object)

        Do Until WordsToCheck.Count = 0
            Dim arrToCheck() As KeyValuePair(Of String, Dictionary.SpellCheckWordError)
            Try
                arrToCheck = WordsToCheck.ToArray
            Catch ex As ArgumentException
                Continue Do
            End Try
            For Each item In arrToCheck
                'Dim item = WordsToCheck.First
                If dictCache.ContainsKey(item.Key) = False Then
                    dictCache.Add(item.Key, CurrentDictionary.SpellCheckWord(item.Key))
                End If
                WordsToCheck.Remove(item.Key)
            Next
        Loop

        AddWordsThread = Nothing

        InvokeRepaint()
    End Sub

    Private Delegate Sub cb_InvokeRepaint()
    Public Sub InvokeRepaint()
        If Control.InvokeRequired Then
            'If Control.Visible Then
            Try
                Dim cb As New cb_InvokeRepaint(AddressOf InvokeRepaint)
                Control.Invoke(cb)
            Catch ex As ObjectDisposedException
                'control was disposed
            End Try
            'End If
        Else
            RepaintControl()
        End If

    End Sub

#End Region

    Private Sub SpellCheckControlBase_DictionaryChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DictionaryChanged
        dictCache.Clear()
        InvokeRepaint()
    End Sub
End Class