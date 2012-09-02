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

Partial Public MustInherit Class Dictionary

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
