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
Public Class Dictionary
    Inherits List(Of DictionaryItem)

    Public Overloads Overrides Function ToString() As String
        Return MyBase.Count.ToString & " word" & If(MyBase.Count = 1, "", "s")
    End Function

#Region "Default dic"

    Public Shared DefaultDictionary As Dictionary

    Public Shared ReadOnly Property DefaultDictFile() As String
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
                    DefaultDictionary = New Dictionary()
                    DefaultDictionary.LoadFromFile(DefaultDictFile)
                Catch ex As Exception
                    'failed ... load blank one 
                    DefaultDictionary = New Dictionary(DefaultDictFile, True)
                End Try
            Else
                'file not found ... load blank one
                DefaultDictionary = New Dictionary(DefaultDictFile, True)
            End If
        End If
    End Sub

#Region "PropertyGrid UI Def File Selector"
    Public Class dicFile_UITypeEditor
        Inherits System.Drawing.Design.UITypeEditor

        Public Overloads Overrides Function EditValue(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal provider As IServiceProvider, ByVal value As Object) As Object
            Dim Dict = TryCast(value, Dictionary)
            Using DictionaryEditor As New DictionaryEditor
                Dict = DictionaryEditor.ShowDialog(Dict)
                'qwertyuiop - needed to create a new dictionary object - otherwise it would not run "set" in the underlying property
                Dim DictOut As New Dictionary
                DictOut.AddRange(Dict)
                DictOut.Filename = Dict.Filename
                Return DictOut
            End Using
        End Function

        Public Overloads Overrides Function GetEditStyle(ByVal context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
            Return System.Drawing.Design.UITypeEditorEditStyle.Modal
        End Function
    End Class
#End Region

#End Region

#Region "Find Base Word"

    'makes countries becomes country etc

    Public Class FindBaseWordReturn
        Public Found As Boolean
        Public WordBase As String
        Public Enum BaseTypes
            None = 0 'base word
            Plural = 1 's / ies
            Progressive = 2 'ing
            PastTense = 4 'ed
            Comparative = 8 'er
            Superlative = 16 'ist
            PastParticiple = 32 'en
            'tion - sugguestion

            'need stuff for??:
            'ian - comedian, politic-ian, utop-ian, austral-ian, as-ian - EnumName: Origin
            'ise, ize - neutralise, plagiarised + ed in this case...
            'ify - mistify, acidify - EnumName: Produce
            'ful - playful
            'ly - playfully - ful+ly, happily
            'able - movable, blamable - drop e
            'ible - corruptible, may also start with in or irr eg incorruptible / irresistible (readd r in this case)
            'ness - happiness
            'less - careless
            'ism - industrialism
            'al - industryal
            'ment - enjoyment
            'ist - perfectionist, artist
            'ish - English

            'prefixes:
            'irr - irregular
            'un - unnatural
            'ab - abnormal
            'de - detoxify, degrade, demote - EnumName: Reverse
            're - redo, reconstitute, reopen - EnumName: Repetitive
        End Enum

        Public BaseType As BaseTypes
        Public ReadOnly Property BaseTypeToArray() As String()
            Get
                Return (From xItem In [Enum].GetValues(GetType(BaseTypes)).OfType(Of BaseTypes)() Where xItem <> BaseTypes.None AndAlso (BaseType And xItem) = xItem Select System.Text.RegularExpressions.Regex.Replace([Enum].GetName(GetType(BaseTypes), xItem), "(?=(?<!^)[A-Z])", " ")).ToArray
            End Get
        End Property
    End Class

    Public Function FindBaseWord(ByVal Word As String) As FindBaseWordReturn
        Word = Word.Trim("'"c)

        Dim FindBaseWordReturn As New FindBaseWordReturn
        FindBaseWordReturn.Found = False
        FindBaseWordReturn.WordBase = Word

        If Word.EndsWith("ing", StringComparison.OrdinalIgnoreCase) Then
            FindBaseWordReturn.BaseType = Dictionary.FindBaseWordReturn.BaseTypes.Progressive
        ElseIf Word.EndsWith("er", StringComparison.OrdinalIgnoreCase) Then
            FindBaseWordReturn.BaseType = Dictionary.FindBaseWordReturn.BaseTypes.Comparative
        ElseIf Word.EndsWith("ed", StringComparison.OrdinalIgnoreCase) Then
            FindBaseWordReturn.BaseType = Dictionary.FindBaseWordReturn.BaseTypes.PastTense
        ElseIf Word.EndsWith("ist", StringComparison.OrdinalIgnoreCase) Then
            FindBaseWordReturn.BaseType = Dictionary.FindBaseWordReturn.BaseTypes.Superlative
        ElseIf Word.EndsWith("en", StringComparison.OrdinalIgnoreCase) Then
            FindBaseWordReturn.BaseType = Dictionary.FindBaseWordReturn.BaseTypes.PastParticiple
        End If

        Dim ReTryWord As String = ""
        If Word.EndsWith("ing", StringComparison.OrdinalIgnoreCase) OrElse Word.EndsWith("ed", StringComparison.OrdinalIgnoreCase) OrElse Word.EndsWith("er", StringComparison.OrdinalIgnoreCase) OrElse Word.EndsWith("est", StringComparison.OrdinalIgnoreCase) OrElse Word.EndsWith("en", StringComparison.OrdinalIgnoreCase) Then
            Dim theWord = System.Text.RegularExpressions.Regex.Replace(Word, "(ing|ed|er|est|en)$", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            'see if the word exists in the dictionary
            If ContainsWord(theWord, True) Then
                'found
                FindBaseWordReturn.Found = True
                FindBaseWordReturn.WordBase = theWord
            Else
                'drop 1 letter if double letters at the back eg cancelled... on ed/ing words
                If Word.EndsWith("ing", StringComparison.OrdinalIgnoreCase) OrElse Word.EndsWith("ed", StringComparison.OrdinalIgnoreCase) Then
                    If theWord.Length >= 2 AndAlso theWord.ToLower.EndsWith(New String(theWord.ToLower.Last, 2)) Then
                        Dim OrigTheWord = theWord
                        theWord = theWord.Remove(theWord.Length - 1, 1)
                        If ContainsWord(theWord, True) Then
                            'found
                            FindBaseWordReturn.Found = True
                            FindBaseWordReturn.WordBase = theWord
                        Else
                            'not found put the letter back and continue
                            theWord = OrigTheWord
                        End If
                    End If
                End If

                If FindBaseWordReturn.Found = False Then
                    'add the e back? - for words like ignoring/ignored = ignore, wider = wide, shaven = shave
                    ReTryWord = theWord
                    theWord &= "e"
                    If ContainsWord(theWord, True) Then
                        'found
                        FindBaseWordReturn.Found = True
                        FindBaseWordReturn.WordBase = theWord
                    Else
                        If Word.EndsWith("ier", StringComparison.OrdinalIgnoreCase) OrElse Word.EndsWith("iest", StringComparison.OrdinalIgnoreCase) Then
                            'for words like happier/happiest = happy
                            theWord = System.Text.RegularExpressions.Regex.Replace(Word, "(ier|iest)$", "y", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                            If ContainsWord(theWord, True) Then
                                FindBaseWordReturn.Found = True
                                FindBaseWordReturn.WordBase = theWord
                            Else
                                'not found...
                                'cannot extend off ist or ier
                                theWord = ""
                            End If
                        Else
                            'not found...
                        End If
                    End If
                End If
            End If
        ElseIf Word.EndsWith("ies", StringComparison.OrdinalIgnoreCase) Then
            Dim theWord = System.Text.RegularExpressions.Regex.Replace(Word, "ies$", "y", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If ContainsWord(theWord, True) Then
                FindBaseWordReturn.Found = True
                FindBaseWordReturn.WordBase = theWord
                FindBaseWordReturn.BaseType = Dictionary.FindBaseWordReturn.BaseTypes.Plural
            Else
                'not found...
                ReTryWord = theWord
            End If
        ElseIf Word.EndsWith("s", StringComparison.OrdinalIgnoreCase) OrElse Word.EndsWith("'s", StringComparison.OrdinalIgnoreCase) Then ' OrElse Word.EndsWith("'", StringComparison.OrdinalIgnoreCase) Then
            Dim theWord = System.Text.RegularExpressions.Regex.Replace(Word, "('?s|')$", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase) 'Old match for words ending in ' "('?s|')$"
            If ContainsWord(theWord, True) Then
                FindBaseWordReturn.Found = True
                FindBaseWordReturn.WordBase = theWord
                FindBaseWordReturn.BaseType = Dictionary.FindBaseWordReturn.BaseTypes.Plural
            Else
                FindBaseWordReturn = FindBaseWord(theWord)
                FindBaseWordReturn.BaseType = FindBaseWordReturn.BaseType Or Dictionary.FindBaseWordReturn.BaseTypes.Plural
            End If
        Else
            'not found... and no special pre/suf-ixes
        End If
        'do multiple times ... for words like quickened
        If ReTryWord <> "" AndAlso FindBaseWordReturn.Found = False Then
            Dim OldBaseWordReturn = FindBaseWordReturn
            FindBaseWordReturn = FindBaseWord(ReTryWord)
            FindBaseWordReturn.BaseType = FindBaseWordReturn.BaseType Or OldBaseWordReturn.BaseType
        End If
        If FindBaseWordReturn.Found = False Then FindBaseWordReturn.BaseType = Nothing
        Return FindBaseWordReturn
    End Function

#End Region

#Region "Anagram Lookup"

    Public Function AnagramLookup(ByVal Letters As String) As List(Of String)
        AnagramLookup = New List(Of String)

        Letters = Letters.ToLower

        Dim Dict() As String = (From xItem In Me Where xItem.ItemState <> DictionaryItem.eItemState.Delete Select xItem.Entry.ToLower).ToArray 'fill dict here

        Dim WordsWithMatchingChars = (From xItem In Dict Where xItem.Length = Letters.Length Select New With {.Word = xItem, .LetterCount = (From xItemWordMatch In xItem Where Letters.Contains(xItemWordMatch)).Count}).ToArray

        Dim MatchingLenthWords = (From xItem In WordsWithMatchingChars Where xItem.LetterCount = Letters.Count Select xItem.Word).ToArray

        'final check to check each char occurs only for each char:
        For Each iMatchedWord In MatchingLenthWords
            Dim CheckLetters = Letters
            For Each iMatchedWordLetter In iMatchedWord
                'find and remove the letter from check
                Dim index = CheckLetters.IndexOf(iMatchedWordLetter)
                If index <> -1 Then
                    CheckLetters = CheckLetters.Remove(index, 1)
                End If
            Next
            If CheckLetters.Count = 0 Then
                'all matched
                AnagramLookup.Add(iMatchedWord)
            End If
        Next
    End Function

#End Region

#Region "Scrabble Solver"

#Region "Scrabble result"

    Public Class ScrabbleResult
        Public Word As String
        Public Shared ReadOnly Property ScrabbleScoreLetters() As Dictionary(Of Char, Integer)
            Get
                Static mc_ScrabbleScoreLetters As Dictionary(Of Char, Integer)
                If mc_ScrabbleScoreLetters Is Nothing Then
                    mc_ScrabbleScoreLetters = New Dictionary(Of Char, Integer)()
                    mc_ScrabbleScoreLetters.Add("A"c, 1)
                    mc_ScrabbleScoreLetters.Add("B"c, 3)
                    mc_ScrabbleScoreLetters.Add("C"c, 3)
                    mc_ScrabbleScoreLetters.Add("D"c, 2)
                    mc_ScrabbleScoreLetters.Add("E"c, 1)
                    mc_ScrabbleScoreLetters.Add("F"c, 4)
                    mc_ScrabbleScoreLetters.Add("G"c, 2)
                    mc_ScrabbleScoreLetters.Add("H"c, 4)
                    mc_ScrabbleScoreLetters.Add("I"c, 1)
                    mc_ScrabbleScoreLetters.Add("J"c, 8)
                    mc_ScrabbleScoreLetters.Add("K"c, 5)
                    mc_ScrabbleScoreLetters.Add("L"c, 1)
                    mc_ScrabbleScoreLetters.Add("M"c, 3)
                    mc_ScrabbleScoreLetters.Add("N"c, 1)
                    mc_ScrabbleScoreLetters.Add("O"c, 1)
                    mc_ScrabbleScoreLetters.Add("P"c, 3)
                    mc_ScrabbleScoreLetters.Add("Q"c, 10)
                    mc_ScrabbleScoreLetters.Add("R"c, 1)
                    mc_ScrabbleScoreLetters.Add("S"c, 1)
                    mc_ScrabbleScoreLetters.Add("T"c, 1)
                    mc_ScrabbleScoreLetters.Add("U"c, 1)
                    mc_ScrabbleScoreLetters.Add("V"c, 4)
                    mc_ScrabbleScoreLetters.Add("W"c, 4)
                    mc_ScrabbleScoreLetters.Add("X"c, 8)
                    mc_ScrabbleScoreLetters.Add("Y"c, 4)
                    mc_ScrabbleScoreLetters.Add("Z"c, 10)
                End If
                Return mc_ScrabbleScoreLetters
            End Get
        End Property
        Public ReadOnly Property Score() As Integer
            Get
                Return (From xItem In Word.ToUpper Where ScrabbleResult.ScrabbleScoreLetters.ContainsKey(xItem) Select ScrabbleResult.ScrabbleScoreLetters(xItem)).Sum
            End Get
        End Property
        Public Sub New(ByVal Word As String)
            Me.Word = Word
        End Sub
    End Class

#End Region

    Public Function ScrabbleLookup(ByVal Letters As String) As List(Of ScrabbleResult)
        ScrabbleLookup = New List(Of ScrabbleResult)

        Dim Dict() As String = (From xItem In Me Where xItem.ItemState <> DictionaryItem.eItemState.Delete Select xItem.Entry.ToLower).ToArray

        For Each iDict In Dict
            Dim CheckLetters = Letters
            For Each iDictLetter In iDict 'for each letter in our dictionary
                Dim index = CheckLetters.IndexOf(iDictLetter)
                If index = -1 Then
                    'this letter is not in our Letters
                    GoTo NextWord
                Else
                    'take the letter out
                    CheckLetters = CheckLetters.Remove(index, 1)
                End If
            Next
            'if we got this far we were can add the word to our return :)
            ScrabbleLookup.Add(New ScrabbleResult(iDict))
NextWord:
        Next
    End Function

#End Region

    Dim mc_Filename As String = ""
    Public Property Filename() As String
        Get
            Return mc_Filename
        End Get
        Set(ByVal value As String)
            mc_Filename = value
        End Set
    End Property

    Public Sub New()

    End Sub

    Private mc_Loading As Boolean
    Public ReadOnly Property Loading() As Boolean
        Get
            Return mc_Loading
        End Get
    End Property

    Public Function ContainsWord(ByVal Word As String, ByVal CaseInsensitive As Boolean) As Boolean
        ContainsWord = (From xItem In Me Where If(CaseInsensitive = False, xItem.Entry = Word, xItem.Entry.ToLower = Word.ToLower)).Count > 0
    End Function

    Public Sub LoadFromFile(ByVal Filename As String)
        mc_Loading = True
        Dim rnd = Guid.NewGuid
        'Debug.Print(rnd.ToString & " - Loading dictionary " & Filename)
        'Suppoers loading of Open office dictionaries, these can be downloaded @ http://wiki.services.openoffice.org/wiki/Dictionaries#English_.28AU.2CCA.2CGB.2CNZ.2CUS.2CZA.29

        Me.Clear()

        'try loading from a serialized object...
        Try
            Dim DictFile = TryCast(Serialize.FileDeserialize(Filename), Dictionary)
            If DictFile IsNot Nothing Then
                Me.AddRange(DictFile.ToArray)
            End If
        Catch ex As Exception
            'try loading from open office file - also supports flat files :)
            Dim fileData = My.Computer.FileSystem.ReadAllText(Filename)
            Dim lines = fileData.Split(New String() {vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries).ToList  'fileData.Split(CChar(vbLf))
            If lines.Count > 0 Then lines(0) = ""
            Dim linesLinq = (From xItem In lines Select xItem.Split("/".ToArray, 2)(0).Trim(CChar(vbCr))).ToArray
            Me.AddRange((From xItem In linesLinq Where xItem <> "" Select New DictionaryItem With {.Entry = xItem, .ItemState = DictionaryItem.eItemState.Existing}))

        End Try
        Me.mc_Filename = Filename

        mc_Loading = False
    End Sub

    Public Overloads Sub Add(ByVal Item As String, Optional ByVal IgnoreForNow As Boolean = False)
        If Trim(Item) <> "" Then
            Dim ItemState As DictionaryItem.eItemState = DictionaryItem.eItemState.AddNew
            If IgnoreForNow Then ItemState = DictionaryItem.eItemState.Ignore
            Dim CurrentItem = (From xItem In Me Where xItem.Entry = Item).FirstOrDefault
            If CurrentItem Is Nothing Then
                'add new
                CurrentItem = New DictionaryItem(Item, ItemState)
                MyBase.Add(CurrentItem)
            Else
                'update existing
                If (CurrentItem.ItemState = DictionaryItem.eItemState.Delete OrElse CurrentItem.ItemState = DictionaryItem.eItemState.Ignore) AndAlso IgnoreForNow = False Then
                    CurrentItem.ItemState = DictionaryItem.eItemState.AddNew
                End If
            End If
        End If
    End Sub

    Public Overloads Sub Remove(ByVal Item As String)
        For Each i In (From xItem In Me Where LCase(xItem.Entry) = LCase(Item))
            i.ItemState = DictionaryItem.eItemState.Delete
        Next
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
        Public Enum eItemState
            Existing
            AddNew
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
                    ItemState = eItemState.Existing
                End If
            End Set
        End Property
        <NonSerialized()> _
        Public ItemState As eItemState = eItemState.AddNew
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

    Public Sub BaseRemove(ByVal item As DictionaryItem)
        MyBase.Remove(item)
    End Sub

    Private Sub CommitChanges()
        'mark the items in .net as old since they have now been commited
        For Each i In (From xItem In Me Where xItem.ItemState = DictionaryItem.eItemState.AddNew).ToArray
            i.ItemState = DictionaryItem.eItemState.Existing
        Next
        'remove the deleted items
        For Each i In (From xItem In Me Where xItem.ItemState = DictionaryItem.eItemState.Delete).ToArray
            MyBase.Remove(i)
        Next
    End Sub

    Public Function Save(Optional ByVal FileName As String = "", Optional ByVal SaveSerialized As Boolean = False, Optional ByVal ForceFullSave As Boolean = False) As Boolean
        Save = False
        Dim SameFile As Boolean = False
        If FileName = "" AndAlso mc_Filename <> "" Then
            SameFile = True
        Else
            If LCase(FileIO.FileSystem.GetFileInfo(FileName).FullName) = LCase(LCase(FileIO.FileSystem.GetFileInfo(mc_Filename).FullName)) Then
                SameFile = True
            End If
        End If

        If FileName = "" Then FileName = mc_Filename
        If FileName = "" Then
            Throw New Exception("No filename has been specified to save the dictionary to")
        End If

        If ForceFullSave = False AndAlso SameFile AndAlso FileIO.FileSystem.FileExists(FileName) Then
            'just add the extra items to the exiting file so that multiple users can access (and add) to the same dictionary simultaneously
            '... if we can ... otherwise just re-save the whole file
            Try
                'load the file
                Dim DictFile As New Dictionary
                DictFile.LoadFromFile(FileName)
                'TryCast(Serialize.FileDeserialize(FileName), Dictionary)
                If DictFile IsNot Nothing Then
                    'add new items
                    DictFile.AddRange((From xItem In Me Where xItem.ItemState = DictionaryItem.eItemState.AddNew).ToArray)
                    'remove the deleted items
                    For Each i In (From xItem In Me _
                                   Group Join f In DictFile On xItem.Entry Equals f.Entry Into Group _
                                   Where xItem.ItemState = DictionaryItem.eItemState.Delete).ToArray
                        DictFile.BaseRemove(i.Group.FirstOrDefault)
                    Next
                    're-save the file
                    If SaveSerialized Then
                        Serialize.FileSerialize(FileName, DictFile)
                    Else
                        My.Computer.FileSystem.WriteAllText(FileName, "WordText" & vbCrLf & Strings.Join((From xItem In DictFile Select xItem.Entry).ToArray, vbCrLf), False)
                    End If
                    CommitChanges()
                    Save = True
                End If
            Catch ex As Exception
            End Try
        End If

        If Save = False Then
            'has not been saved yet or couldn't be merged... save with complete save...
            Try
                CommitChanges()
                If SaveSerialized Then
                    Dim tmpDictionary As New Dictionary
                    tmpDictionary.AddRange((From xItem In Me Where xItem.ItemState <> DictionaryItem.eItemState.Ignore))
                    Serialize.FileSerialize(FileName, tmpDictionary)
                Else
                    My.Computer.FileSystem.WriteAllText(FileName, "WordText" & vbCrLf & Strings.Join((From xItem In Me Where xItem.ItemState <> DictionaryItem.eItemState.Ignore Select xItem.Entry).ToArray, vbCrLf), False)
                End If
                Save = True
            Catch ex As Exception

            End Try
        End If

    End Function

End Class