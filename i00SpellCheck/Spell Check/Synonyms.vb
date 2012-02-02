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

Partial Class SpellCheckTextBox

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

    <System.ComponentModel.TypeConverter(GetType(System.ComponentModel.ExpandableObjectConverter))> _
    Public Class Synonyms
        '        Inherits System.ComponentModel.ExpandableObjectConverter

        '#Region "PropertyGrid FromString (ExpandableObjectConverter)"

        '        Public Overrides Function CanConvertFrom( _
        '                                      ByVal context As System.ComponentModel.ITypeDescriptorContext, _
        '                                      ByVal sourceType As System.Type) As Boolean
        '            If (sourceType Is GetType(String)) Then
        '                Return True
        '            End If
        '            Return MyBase.CanConvertFrom(context, sourceType)
        '        End Function

        '        Public Overrides Function GetCreateInstanceSupported(ByVal context As System.ComponentModel.ITypeDescriptorContext) As Boolean
        '            Return True
        '        End Function

        '        Public Overrides Function CanConvertTo(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal destinationType As System.Type) As Boolean
        '            If (destinationType Is GetType(String)) OrElse destinationType Is GetType(Synonyms) Then
        '                Return True
        '            End If
        '        End Function

        '        Public Overrides Function ConvertTo(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As System.Type) As Object
        '            If (destinationType Is GetType(String)) Then
        '                Return Me.File
        '            ElseIf (destinationType Is GetType(Synonyms)) Then
        '                Return Me
        '            End If
        '            Return Nothing
        '        End Function

        '        Public Overrides Function ConvertFrom( _
        '                              ByVal context As System.ComponentModel.ITypeDescriptorContext, _
        '                              ByVal culture As System.Globalization.CultureInfo, _
        '                              ByVal value As Object) As Object
        '            If TypeOf value Is String Then
        '                Dim File = DirectCast(value, String)
        '                If FileIO.FileSystem.FileExists(File) Then
        '                    Return New Synonyms With {.File = File}
        '                End If
        '            End If
        '            Return Nothing
        '        End Function

        '#End Region

#Region "Default Synonyms"

        Public Shared ReadOnly Property DefaultSynonymsFile() As String
            Get
                Return IO.Path.Combine(IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location()), "syn.syn")
            End Get
        End Property

        Public Shared DefaultSynonyms As Synonyms
        Public Shared Sub LoadDefaultSynonyms()
            If DefaultSynonyms Is Nothing Then
                DefaultSynonyms = New Synonyms With {.File = DefaultSynonymsFile}
            End If
        End Sub

#End Region

        Public Overrides Function ToString() As String
            Return If(File <> "", File, "No File")
        End Function

        Dim mc_File As String
        <System.ComponentModel.DisplayName("Filename")> _
        <System.ComponentModel.Editor(GetType(symFile_UITypeEditor), GetType(System.Drawing.Design.UITypeEditor))> _
        Public Property File() As String
            Get
                Return mc_File
            End Get
            Set(ByVal value As String)
                mc_File = value
            End Set
        End Property

#Region "PropertyGrid UI Syn File Selector"
        Public Class symFile_UITypeEditor
            Inherits System.Drawing.Design.UITypeEditor

            Public Overloads Overrides Function EditValue(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal provider As IServiceProvider, ByVal value As Object) As Object
                Using ofd As New OpenFileDialog
                    ofd.Filter = "Synonym Files (*.syn)|*.syn|All Files (*.*)|*.*"
                    If ofd.ShowDialog() = DialogResult.OK Then
                        value = ofd.FileName
                    End If
                    Return value
                End Using
            End Function

            Public Overloads Overrides Function GetEditStyle(ByVal context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
                Return System.Drawing.Design.UITypeEditorEditStyle.Modal
            End Function
        End Class
#End Region

        <System.ComponentModel.DisplayName("Count")> _
        <System.ComponentModel.Description("The count of words that have synonyms specified in the selected syn file")> _
        Public ReadOnly Property SynonymCount() As Long
            Get
                If FileIO.FileSystem.FileExists(File) Then
                    Using sr As New IO.StreamReader(File)
                        '    Return System.Text.RegularExpressions.Regex.Matches(sr.ReadToEnd(), "\r\n|\n\r|\r|\n").Count
                        Do Until sr.EndOfStream
                            Dim ThisLine = sr.ReadLine
                            SynonymCount += 1
                        Loop
                    End Using
                End If
            End Get
        End Property

        Public Function FindWord(ByVal Word As String) As List(Of FindWordReturn)
            If FileIO.FileSystem.FileExists(File) Then
                Using sr As New IO.StreamReader(File)
                    Do Until sr.EndOfStream
                        Dim ThisLine = sr.ReadLine
                        If ThisLine.ToLower.StartsWith(Word.ToLower & "|") Then
                            'found what we want
                            Return FindWordReturn.FromFileLine(ThisLine)
                        End If
                    Loop
                End Using
            End If
            Return Nothing
        End Function

        Public Class FindWordReturn
            Inherits List(Of String)
            Private Sub New()

            End Sub
            Public Shared Function FromFileLine(ByVal Line As String) As List(Of FindWordReturn)
                Try
                    Dim Output = New List(Of FindWordReturn)
                    Dim Defs = Line.Split("|"c).ToList
                    Defs.Remove(Defs.First)
                    For Each iDef In Defs
                        Dim FindWordReturn As New FindWordReturn
                        Dim SplitDef = iDef.Split(New Char() {">"c}, 2)
                        Dim SplitDefWord = SplitDef(0).Split(";"c)
                        FindWordReturn.TypeDescription = SplitDefWord(0)
                        FindWordReturn.WordType = DirectCast(CInt(SplitDefWord(1)), WordTypes)
                        FindWordReturn.AddRange(SplitDef(1).Split(";"c))
                        If FindWordReturn.Count > 0 Then
                            Output.Add(FindWordReturn)
                        End If
                    Next
                    If Output.Count > 0 Then
                        Return Output
                    Else
                        Return Nothing
                    End If
                Catch ex As Exception
                    Return Nothing
                End Try
            End Function
            Public TypeDescription As String
            Public Enum WordTypes
                Adjective = 0
                Noun = 1
                Adverb = 2
                Verb = 3
                Pronoun = 4
                Conjunction = 5
                Preposition = 6
                Interjection = 7
                Idiom = 8
                Other = 9
            End Enum
            Public WordType As WordTypes
        End Class

    End Class

End Class
