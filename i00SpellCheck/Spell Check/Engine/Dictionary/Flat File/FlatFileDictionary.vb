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

<Serializable()> _
Public Class FlatFileDictionary
    Inherits Dictionary

    Friend WordList As New List(Of DictionaryItem)

    Public Overrides ReadOnly Property Count() As Integer
        Get
            Return WordList.Count
        End Get
    End Property

    Public Overloads Overrides Function ToString() As String
        Return WordList.Count.ToString & " word" & If(WordList.Count = 1, "", "s")
    End Function

#Region "Default dic"

    Friend Shared ReadOnly Property DefaultDictFile() As String
        Get
            Return IO.Path.Combine(IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location()), "dic.dic")
        End Get
    End Property

    Public Shared Sub LoadDefaultDictionary()
        If DefaultDictionary Is Nothing Then
            'load from file
            If FileIO.FileSystem.FileExists(DefaultDictFile) Then
                'load this
                Try
                    DefaultDictionary = New FlatFileDictionary()
                    DefaultDictionary.LoadFromFile(DefaultDictFile)
                Catch ex As Exception
                    'failed ... load blank one 
                    DefaultDictionary = New FlatFileDictionary(DefaultDictFile, True)
                End Try
            Else
                'file not found ... load blank one
                DefaultDictionary = New FlatFileDictionary(DefaultDictFile, True)
            End If
        End If
    End Sub

#End Region

    Dim mc_Filename As String = ""
    Public Overrides ReadOnly Property Filename() As String
        Get
            Return mc_Filename
        End Get
    End Property

    Public Sub New()

    End Sub

    Private mc_Loading As Boolean
    Public Overrides ReadOnly Property Loading() As Boolean
        Get
            Return mc_Loading
        End Get
    End Property

    Public Overrides Sub LoadFromFile(ByVal Filename As String)
        LoadFromFileInternal(Filename)
    End Sub

    Private Sub LoadFromFileInternal(ByVal Filename As String, Optional ByVal LoadNonUser As Boolean = True, Optional ByVal LoadUser As Boolean = True)
        mc_Loading = True
        WordList.Clear()
        IndexDictionarySections.Clear()

        'zxcvbnm - need way to re-index?

        'try loading from a flat file
        If LoadNonUser = True Then
            Dim fileData = My.Computer.FileSystem.ReadAllText(Filename)
            Dim lines = fileData.Split(New String() {vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries).ToArray  'fileData.Split(CChar(vbLf))
            WordList.AddRange((From xItem In lines Where xItem <> "" Select New DictionaryItem With {.Entry = xItem, .ItemState = DictionaryItem.eItemState.Correct, .Committed = True, .User = False}))

            'index the first letters
            WordList.Sort(Function(p1, p2) p1.Entry.CompareTo(p2.Entry))
            'now find the location of each letter...
            Dim SplitByLetters = (From xItem In WordList _
                                 Group xItem By StartChar = xItem.Entry.ToLower.First Into Group).ToArray
            For Each LetterGroup In SplitByLetters
                IndexDictionarySections.Add(LetterGroup.StartChar, New IndexDictionarySection() With {.StartIndex = WordList.IndexOf(LetterGroup.Group.First), .Length = LetterGroup.Group.Count})
            Next
            UserDictStart = WordList.Count
        End If

        If LoadUser = True Then
            'now load the user dictionary
            Dim UserDicFile As String = Filename & ".user"
            If FileIO.FileSystem.FileExists(UserDicFile) Then
                Dim fileData = My.Computer.FileSystem.ReadAllText(UserDicFile)
                Dim lines = fileData.Split(New String() {vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries).ToArray  'fileData.Split(CChar(vbLf))

                WordList.AddRange((From xItem In lines Where xItem <> "" AndAlso (xItem.First = "+"c OrElse xItem.First = "*"c) Select New DictionaryItem With {.Entry = xItem.Substring(1), .ItemState = If(xItem.First = "*", DictionaryItem.eItemState.Ignore, DictionaryItem.eItemState.Correct), .Committed = True}))

                Dim DeletedWords = (From xItem In lines Where xItem <> "" AndAlso (xItem.First = "-"c) Select xItem.Substring(1)).ToArray
                'zxcvbnm - should find a faster way to do this ... linq join?
                For Each DeletedWordItem In DeletedWords
                    pRemove(DeletedWordItem, True).User = True
                Next

            End If
        End If

        Me.mc_Filename = Filename

        mc_Loading = False
    End Sub


    Dim UserDictStart As Integer
    Dim IndexDictionarySections As New Dictionary(Of Char, IndexDictionarySection)

    Private Class IndexDictionarySection
        Public StartIndex As Integer
        Public Length As Integer
    End Class

    Private Function GetSectionDict(ByVal Word As String) As DictionaryItem()
        If IndexDictionarySections.ContainsKey(Word.ToLower.First) Then
            Dim IndexItem = IndexDictionarySections(Word.ToLower.First)
            Dim SectionSegment(IndexItem.Length - 1) As DictionaryItem
            WordList.CopyTo(IndexItem.StartIndex, SectionSegment, 0, SectionSegment.Count)
            Return SectionSegment
        Else
            Return Nothing
        End If
    End Function

    Private Function GetUserDict() As DictionaryItem()
        Dim UserDict() As DictionaryItem = Nothing

        Dim UserDictSize = WordList.Count - UserDictStart
        If UserDictSize >= 1 Then
            Dim SectionSegment(UserDictSize - 1) As DictionaryItem
            WordList.CopyTo(UserDictStart, SectionSegment, 0, SectionSegment.Length)
            UserDict = SectionSegment
        End If
        Return UserDict
    End Function

    Private Sub RemoveFromUserDict(ByVal Word As String)
        Dim UserDict = GetUserDict()
        If UserDict IsNot Nothing Then
            Dim ExistingUserItems = (From xItem In UserDict Where xItem.Entry = Word).ToArray
            For Each UserItem In ExistingUserItems
                WordList.Remove(UserItem)
            Next
        End If
    End Sub

    Private Sub AddIgnore(ByVal Item As String, Optional ByVal IgnoreForNow As Boolean = False)
        If Trim(Item) <> "" Then
            RemoveFromUserDict(Item)
            Dim ItemState As DictionaryItem.eItemState = DictionaryItem.eItemState.Correct
            If IgnoreForNow Then ItemState = DictionaryItem.eItemState.Ignore
            For Each CurrentItem In (From xItem In WordList Where xItem.Entry = Item).ToArray
                'update existing
                CurrentItem.ItemState = ItemState
            Next
            Dim NewItem = New DictionaryItem(Item, ItemState)
            WordList.Add(NewItem)
        End If
    End Sub

    Public Overrides Sub Ignore(ByVal Item As String)
        AddIgnore(Item, True)
    End Sub

    Public Overrides Sub Add(ByVal Item As String)
        AddIgnore(Item)
    End Sub

    Private Function pRemove(ByVal Item As String, Optional ByVal Commited As Boolean = False) As DictionaryItem
        RemoveFromUserDict(Item)
        Dim NoApoS = Dictionary.Formatting.RemoveApoS(Item)
        For Each DelItem In (From xItem In WordList Where LCase(xItem.Entry) = LCase(Item) OrElse LCase(xItem.Entry) = LCase(NoApoS))
            DelItem.ItemState = DictionaryItem.eItemState.Delete
        Next
        pRemove = New DictionaryItem(Item, DictionaryItem.eItemState.Delete) With {.Committed = Commited}
        WordList.Add(pRemove)
    End Function

    Public Overrides Sub Remove(ByVal Item As String)
        pRemove(Item)
    End Sub

    Public Sub New(ByVal Filename As String, ByVal CreateNewDict As Boolean)
        If CreateNewDict Then
            Me.mc_Filename = Filename
        Else
            LoadFromFile(Filename)
        End If
    End Sub

    Public Sub New(ByVal Filename As String)
        LoadFromFile(Filename)
    End Sub

    <Serializable()> _
    Public Class DictionaryItem
        'public for speed...
        Public Entry As String
        'and property for data source binding...
        Public Property pEntry() As String
            Get
                Return Entry
            End Get
            Set(ByVal value As String)
                Entry = value
            End Set
        End Property
        Public User As Boolean = True
        Public Committed As Boolean
        Public Enum eItemState
            Correct
            Ignore
            Delete
        End Enum
        'property for data source binding...
        Public Property pIgnore() As Boolean
            Get
                Return ItemState = eItemState.Ignore
            End Get
            Set(ByVal value As Boolean)
                If value = True Then
                    ItemState = eItemState.Ignore
                Else
                    ItemState = eItemState.Correct
                End If
            End Set
        End Property
        <NonSerialized()> _
        Public ItemState As eItemState = eItemState.Correct
        Public Sub New()

        End Sub
        Public Sub New(ByVal Entry As String, ByVal ItemState As eItemState)
            Me.Entry = Entry
            Me.ItemState = ItemState
        End Sub
        Public Overrides Function ToString() As String
            Return Me.Entry & " - " & Me.ItemState.ToString
        End Function
    End Class

    Public Overrides Sub Save(Optional ByVal FileName As String = "", Optional ByVal ForceFullSave As Boolean = False)
        'save the user items to a file...
        
        Dim SameFile As Boolean = False
        If FileName = "" AndAlso mc_Filename <> "" Then
            SameFile = True
        Else
            If mc_Filename = "" Then
                SameFile = False
            Else
                If LCase(FileIO.FileSystem.GetFileInfo(FileName).FullName) = LCase(LCase(FileIO.FileSystem.GetFileInfo(mc_Filename).FullName)) Then
                    SameFile = True
                End If
            End If
        End If

        If FileName = "" Then FileName = mc_Filename
        If FileName = "" Then
            Throw New Exception("No filename has been specified to save the dictionary to")
        End If


        Dim UserDict = GetUserDict()
        Dim UserDicFile As String = FileName & ".user"


        If ForceFullSave Or SameFile = False Then
            'we need to save the non-user dictionary too...
            Dim FileData = Join((From xItem In WordList Where xItem.User = False Select xItem.Entry).ToArray, vbCrLf)
            My.Computer.FileSystem.WriteAllText(FileName, FileData, False)
        End If

        If SameFile AndAlso FileIO.FileSystem.FileExists(UserDicFile) AndAlso ForceFullSave = False Then
            'just add the extra items to the exiting file so that multiple users can access (and add) to the same dictionary simultaneously
            '... if we can ... otherwise just re-save the whole file
            Try
                'load the user file
                Dim CompareDict As New FlatFileDictionary()
                CompareDict.LoadFromFileInternal(FileName, False, True)
                For Each item In UserDict
                    'find the word in the compare dict .. 
                    Dim theItem = item
                    Dim CompareDictWord = (From xItem In CompareDict.WordList Where xItem.Entry = theItem.Entry).FirstOrDefault
                    If CompareDictWord IsNot Nothing Then
                        'update it
                        'for new case
                        CompareDictWord.Entry = item.Entry
                        'for item state
                        CompareDictWord.ItemState = item.ItemState
                    Else
                        'add it
                        CompareDict.WordList.Add(item)
                    End If
                Next

                UserDict = CompareDict.WordList.ToArray
            Catch ex As Exception

            End Try
        End If

        Dim UserDicContent As String = Microsoft.VisualBasic.Strings.Join((From xItem In UserDict Where Trim(xItem.Entry) <> "" Order By xItem.Entry Select If(xItem.ItemState = DictionaryItem.eItemState.Delete, "-", If(xItem.ItemState = DictionaryItem.eItemState.Ignore, "*", "+")) & xItem.Entry).ToArray, vbCrLf)
        My.Computer.FileSystem.WriteAllText(UserDicFile, UserDicContent, False)

        Dim DicItemsToCommit = (From xItem In WordList Where xItem.Committed = False).ToArray
        For Each DicItemToCommit In DicItemsToCommit
            DicItemToCommit.Committed = True
        Next

    End Sub

End Class