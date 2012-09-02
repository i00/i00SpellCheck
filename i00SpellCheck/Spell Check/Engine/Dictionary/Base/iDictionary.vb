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

#Region "Loading / Saving"

    Public MustOverride ReadOnly Property Loading() As Boolean

    Public MustOverride Sub LoadFromFile(ByVal Filename As String)

    Public MustOverride Sub Save(Optional ByVal FileName As String = "", Optional ByVal ForceFullSave As Boolean = False)

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

    Public MustOverride Function SpellCheckSuggestions(ByVal Word As String) As List(Of SpellCheckSuggestionInfo)

    Public Enum SpellCheckWordError
        OK
        SpellError
        CaseError
        Ignore
    End Enum

    Public MustOverride Function SpellCheckWord(ByVal Word As String) As SpellCheckWordError

#End Region

    Public MustOverride ReadOnly Property Filename() As String

    Public MustOverride Sub Add(ByVal Item As String)

    Public MustOverride Sub Ignore(ByVal Item As String)

    Public MustOverride Sub Remove(ByVal Item As String)

    Public Shared DefaultDictionary As Dictionary

    Public MustOverride ReadOnly Property Count() As Integer

End Class
