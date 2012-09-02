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

#Region "Show spellcheck dialog"

    Public Sub ShowDialog() Implements iSpellCheckDialog.ShowDialog
        If CurrentDictionary IsNot Nothing AndAlso CurrentDictionary.Loading = False Then
            Using SpellCheckDialog As New SpellCheckDialog
                SpellCheckDialog.ShowDialog(parentTextBox, Me)
            End Using
        End If
    End Sub

#End Region

#Region "Ease of access"

    Protected Friend ReadOnly Property parentTextBox() As TextBoxBase
        Get
            Return TryCast(MyBase.Control, TextBoxBase)
        End Get
    End Property

    Protected Friend ReadOnly Property parentRichTextBox() As RichTextBox
        Get
            Return TryCast(MyBase.Control, RichTextBox)
        End Get
    End Property

#End Region

#Region "Misc"

    <System.ComponentModel.Category("Control")> _
    <System.ComponentModel.Description("The TextBox associated with the SpellCheckTextBox object")> _
    <System.ComponentModel.DisplayName("Text Box")> _
    Public Overrides ReadOnly Property Control() As System.Windows.Forms.Control
        Get
            Return MyBase.Control
        End Get
    End Property

    Public Overrides ReadOnly Property ControlType() As System.Type
        Get
            Return GetType(TextBoxBase)
        End Get
    End Property

#End Region

End Class