Public Class Hunspell
    Inherits i00SpellCheck.UserDictionaryBase

    Dim NHunspell As NHunspell.Hunspell

    Public Overrides ReadOnly Property Count() As Integer
        Get
            If NHunspell Is Nothing Then
                Return 0
            Else
                Return -1
            End If
        End Get
    End Property

    Private Function GetAffFile(ByVal Filename As String) As String
        Return IO.Path.Combine(IO.Path.GetDirectoryName(Filename), IO.Path.GetFileNameWithoutExtension(Filename)) & ".aff"
    End Function

    Public Overrides Sub LoadFromFileInternal(ByVal Filename As String)
        NHunspell = New NHunspell.Hunspell(GetAffFile(Filename), Filename)

        'user dict
        LoadUserDictionary(Filename)
        'add to the sugguestions...
        For Each item In (From xItem In UserWordList Where xItem.State = SpellCheckWordError.OK).ToArray
            NHunspell.Add(item.Word)
        Next
    End Sub

    Public Overrides Sub SaveInternal(ByVal Filename As String, Optional ByVal ForceFullSave As Boolean = False)
        'copy file elsewhere?
        SaveUserDictionary(Filename, ForceFullSave)
    End Sub

    Public Overrides Function SpellCheckSuggestionsNonUser(ByVal Word As String) As System.Collections.Generic.List(Of i00SpellCheck.Dictionary.SpellCheckSuggestionInfo)
        Return (From xItem In NHunspell.Suggest(Word) Select New i00SpellCheck.Dictionary.SpellCheckSuggestionInfo(0, xItem)).ToList
    End Function

    Public Overrides Function SpellCheckWordNonUser(ByVal Word As String) As i00SpellCheck.Dictionary.SpellCheckWordError
        If NHunspell.Spell(Word) Then
            Return SpellCheckWordError.OK
        Else
            Return SpellCheckWordError.SpellError
        End If
    End Function

    Private Sub Hunspell_WordAdded(ByVal Item As String) Handles Me.WordAdded
        NHunspell.Add(Item)
    End Sub

    Private Sub Hunspell_WordRemoved(ByVal Item As String) Handles Me.WordRemoved
        'NHunspell - doesn't support removal :(
    End Sub

    Public Overrides ReadOnly Property DicFileFilter() As String
        Get
            Return "*_*.dic"
        End Get
    End Property
End Class
