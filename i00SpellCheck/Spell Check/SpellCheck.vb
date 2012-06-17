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

Partial Class Dictionary

#Region "Check that Word is in the Dictionary"

    Public Enum SpellCheckWordError
        OK
        SpellError
        CaseError
        Ignore
    End Enum

    Public Function SpellCheckWord(ByVal Word As String) As SpellCheckWordError
        'Doing this directly on the word object didnot work????:
        Dim theWord = Word

        'Strip 's
        Dim OldWord = theWord
        theWord = Dictionary.Formatting.RemoveApoS(theWord)

        'ignore numbers
        Dim NumericWord = CStr((From xItem In theWord Select xItem Where xItem <> "$" AndAlso xItem <> "." AndAlso xItem <> "%" AndAlso xItem <> "#").ToArray)
        If IsNumeric(NumericWord) Then
            Return SpellCheckWordError.OK 'not in dic
        End If

        SpellCheckWord = SpellCheckWordError.SpellError
        Dim DicWords() As DictionaryItem
        DicWords = (From xItem In Me.ToArray Where LCase(xItem.Entry) = LCase(theWord) AndAlso xItem.ItemState <> Dictionary.DictionaryItem.eItemState.Delete Select xItem).ToArray

        For Each iDicWord In DicWords
            'words found
            If iDicWord.Entry = LCase(iDicWord.Entry) Then
                'word is not case sensitive
                If iDicWord.ItemState = DictionaryItem.eItemState.Ignore Then
                    Return SpellCheckWordError.Ignore
                Else
                    'check for funny case ... if we have more than one letter and the word isn't in upper case
                    If Word.Length > 1 AndAlso Word.ToUpper <> Word Then
                        Dim AllButFirst = Word.Substring(1, Word.Length - 1)
                        If AllButFirst <> AllButFirst.ToLower Then
                            'interesting case
                            SpellCheckWord = SpellCheckWordError.CaseError
                        Else
                            SpellCheckWord = SpellCheckWordError.OK
                        End If
                    Else
                        Return SpellCheckWordError.OK
                    End If
                End If
            Else
                'word is case sensitive :(
                For iLetter = 1 To Len(iDicWord.Entry)
                    Dim DicLetter = CChar(Mid(iDicWord.Entry, iLetter, 1))
                    Dim Letter = CChar(Mid(theWord, iLetter, 1))
                    If DicLetter = UCase(DicLetter) Then
                        'this letter needs to be in caps... so check it in the orig word
                        If Letter = DicLetter Then
                            'we are ok
                        Else
                            'we couldn't match this iDicWord
                            SpellCheckWord = SpellCheckWordError.CaseError
                            GoTo NextWord
                        End If
                    End If
                Next
                'if we made it this far all caps are matched :)
                If iDicWord.ItemState = DictionaryItem.eItemState.Ignore Then
                    Return SpellCheckWordError.Ignore
                Else
                    Return SpellCheckWordError.OK
                End If
            End If
NextWord:
        Next

        'we are not in the dictionary :(

        '... check for special cases... eg all in caps - CEO
        Dim ChrsInWord = CStr((From xItem In theWord Select xItem Where Asc(LCase(xItem)) >= 97 AndAlso Asc(LCase(xItem)) <= 122).ToArray)
        If ChrsInWord = UCase(ChrsInWord) Then
            'all letters in this word are caps
            Return SpellCheckWordError.OK 'not in dic
        End If

    End Function

#End Region

#Region "Spelling Suggestions"

    Public Class SpellCheckSuggestionInfo
        Public Closness As Integer
        Public Word As String
        Public Sub New(ByVal Closness As Integer, ByVal Word As String)
            Me.Closness = Closness
            Me.Word = Word
        End Sub
    End Class

    Public Function SpellCheckSuggestions(ByVal Word As String) As List(Of SpellCheckSuggestionInfo)
        Dim leewaynum As Integer
        Dim leewaypct As Double

        Dim theWord = Word
        Dim OldWord = theWord
        theWord = System.Text.RegularExpressions.Regex.Replace(theWord, "'s$", "") '.. remove 's ... can't use SpellCheckTextBox.RemoveApoS(theWord) as we also want to remove them if we have chris's
        Dim ApoSRemoved = False
        If OldWord <> theWord Then
            ApoSRemoved = True
        End If

        Dim txtlen As Integer = theWord.Length
        Select Case txtlen
            Case Is < 5
                leewaynum = 2
                leewaypct = 0.75
            Case 5, 6, 7
                leewaynum = 3
                leewaypct = 0.6
            Case 8, 9, 10, 11
                leewaynum = 4
                leewaypct = 0.5
            Case Else
                leewaynum = 5
                leewaypct = 0.45
        End Select

        'this makes words such as runnning match running 1st then everything else
        Dim theWordNoDups = System.Text.RegularExpressions.Regex.Replace(theWord.ToLower, "(.)(\1)+", "$1")

        Dim CutDownDict = (From xItem In Me Where xItem.ItemState <> Dictionary.DictionaryItem.eItemState.Delete AndAlso xItem.Entry.ToLower.StartsWith(Left(theWord.ToLower, 1)) AndAlso Len(xItem.Entry) > txtlen - leewaynum AndAlso Len(xItem.Entry) < txtlen + leewaynum Group xItem By Entry = xItem.Entry Into g = Group Select g(0).Entry)

        SpellCheckSuggestions = New List(Of SpellCheckSuggestionInfo)

        'Dim StartTime = Environment.TickCount
        For Each iWord In CutDownDict
            Dim nummat As Integer = 0
            Dim allmat As Integer = 0
            Dim firstfewmat As Integer = 0

            'If iWord.StartsWith(Left(theWord, 1), StringComparison.OrdinalIgnoreCase) Then
            If theWord.Contains(Left$(iWord, CInt(leewaypct * txtlen))) Then
                '1st leewaypct of characters match (Sliding scale percentage based on theWord len)
                firstfewmat = CInt(4 * txtlen)    'if first 3 of 4 letters matches, weighting would be an extra 12
            End If
            'If txtlen > 5 And (InStr(1, theWord, Left$(iWord, 3), CompareMethod.Text)) > 0 Then
            '    '1st leewaypct of characters match (Sliding scale percentage based on theWord len)
            '    firstfewmat = firstfewmat + 5    'if first 3 of 4 letters matches, weighting would be an extra 12
            'End If
            If iWord.StartsWith(Left(theWord, 3), StringComparison.OrdinalIgnoreCase) AndAlso iWord.EndsWith(Right(theWord, 2), StringComparison.OrdinalIgnoreCase) Then
                '1st leewaypct of characters match (Sliding scale percentage based on theWord len)
                firstfewmat = firstfewmat + 10    'if first 3 of 4 letters matches, weighting would be an extra 12
            End If
            If iWord.EndsWith("cause", StringComparison.OrdinalIgnoreCase) AndAlso theWord.EndsWith("cose", StringComparison.OrdinalIgnoreCase) Then
                'give extra weight to this common mis-spelling
                firstfewmat = firstfewmat + 20    'if first 3 of 4 letters matches, weighting would be an extra 12
            End If
            If iWord.EndsWith("ds", StringComparison.OrdinalIgnoreCase) AndAlso theWord.EndsWith("des", StringComparison.OrdinalIgnoreCase) Then
                'give extra weight to this common mis-spelling
                firstfewmat = firstfewmat + 20    'if first 3 of 4 letters matches, weighting would be an extra 12
            End If
            If txtlen > 5 AndAlso iWord.EndsWith(Right(theWord, 3), StringComparison.OrdinalIgnoreCase) Then
                'last 3 letters match, give this a bit more weight
                firstfewmat = firstfewmat + txtlen    'if first 3 of 4 letters matches, weighting would be an extra 12
            End If

            For i = 1 To Len(theWord)
                If InStr(If(i - 1 > 1, i - 1, i), iWord, Mid(theWord, i, 1), CompareMethod.Text) > 0 Then 'i-1 to cover transpositions
                    'If InStr(IIf(i - 1 > 1, i - 1, i), theWord, Mid$(iWord, i, 1), 1) > 0 Then 'i-1 to cover transpositions
                    nummat = nummat + 1
                End If
            Next i
            If nummat = txtlen Then
                If txtlen = iWord.Length Then
                    allmat = 100    'extra extra weight for all matches, this would probably be a transposition
                Else
                    allmat = 50 'was 20
                End If
            ElseIf Math.Abs(nummat - txtlen) = 1 Then
                'almost all characters were mached
                allmat = 25
            ElseIf Math.Abs(nummat - txtlen) = 2 Then
                allmat = 15
            ElseIf Math.Abs(nummat - txtlen) = 3 Then
                allmat = 10
            End If
            If nummat + allmat + firstfewmat > 0 Then
                Dim SugguestionTxt = iWord
                If ApoSRemoved Then
                    'add the 's back...
                    If SugguestionTxt.EndsWith("'s", StringComparison.OrdinalIgnoreCase) = False OrElse SugguestionTxt.EndsWith("'") = False Then
                        If SugguestionTxt.EndsWith("s", StringComparison.OrdinalIgnoreCase) Then
                            SugguestionTxt &= "'"
                        Else
                            SugguestionTxt &= "'s"
                        End If
                    End If
                End If
                Dim Closeness = nummat + allmat + firstfewmat
                If System.Text.RegularExpressions.Regex.Replace(iWord.ToLower, "(.)(\1)+", "$1") = theWordNoDups Then
                    Closeness = -1
                End If
                SpellCheckSuggestions.Add(New SpellCheckSuggestionInfo(Closeness, SugguestionTxt))
            End If

            'End If
        Next
        If SpellCheckSuggestions.Count > 0 Then
            Dim MaxCloseness = SpellCheckSuggestions.Max(Function(x As SpellCheckSuggestionInfo) x.Closness)
            For Each iSuggest In (From xItem In SpellCheckSuggestions Where xItem.Closness = -1).ToArray
                iSuggest.Closness = MaxCloseness + 1
            Next
        End If

        'Debug.Print((Environment.TickCount - StartTime).ToString)

    End Function

#End Region

    Partial Class Formatting

#Region "Text Formating"

#Region "Word Breaks"

        Public Shared Function RemoveWordBreaks(ByVal Text As String) As String
            If Text <> "" Then
                'replace all chrs
                'Dim arr = (From xItem In WordBreakChrs Select System.Text.RegularExpressions.Regex.Escape(xItem)).ToArray
                Text = System.Text.RegularExpressions.Regex.Replace(Text, regexWordBreakChrs, New System.Text.RegularExpressions.MatchEvaluator(Function(x As System.Text.RegularExpressions.Match) New String(" "c, x.Length)))

                'regex below removes "'" but only for words ending in this eg "chris'" = "chris ", but "heather's" <> "heather s"
                Text = System.Text.RegularExpressions.Regex.Replace(Text, "'(?![A-Z0-9_])", " ", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            End If
            Return Text
        End Function

        'These do not work:
        ' Chr(146) & " "    "’ "

        Public Shared Function regexWordBreakChrs() As String
            Dim arr = (From xItem In WordBreakChrs Select System.Text.RegularExpressions.Regex.Escape(xItem)).ToArray
            'qwertyuiop - below should be "[\s|^|$|\t|\r|\n]{1,}" but doesn't work for start of string starting with ' for eg???
            arr = (From xItem In arr Select Replace(xItem, "\ ", "([\s|^|$|\t|\r|\n]{1,}|^{1}){1}")).ToArray
            regexWordBreakChrs = Join(arr, "|")

        End Function

        Private Shared WordBreakChrs As String() = New String() {vbCr, vbLf, ".", ",", ";", ":", "?", "!", ")", "]", "}", Chr(146) & " ", """", "/", "\", "(", "[", "{", " '", " " & Chr(145), "-", Chr(147), Chr(148)}

#End Region

#Region "For words with 's"

        Friend Shared Function RemoveApoS(ByVal text As String) As String
            Return System.Text.RegularExpressions.Regex.Replace(text, "(?<!s)['\x92]s", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            'Dim Last2 = LCase(Right(Word, 2))
            'If Last2 = "'s" OrElse Last2 = Chr(146) & "s" Then
            '    Word = Left(Word, Len(Word) - 2)
            'End If
            'Return Word
        End Function

#End Region

#End Region

    End Class

End Class