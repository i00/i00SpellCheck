Imports i00SpellCheck

Public Class Form1

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.Icon = My.Resources.Icon

        Dim Hunspell As New Hunspell
        Hunspell.LoadFromFile("en_US.dic")
        i00SpellCheck.Dictionary.DefaultDictionary = Hunspell
        Me.EnableSpellCheck()

    End Sub
End Class
