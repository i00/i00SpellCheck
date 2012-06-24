Imports FastColoredTextBoxNS
Imports i00SpellCheck

Public Class frmTest

    Private Sub frmTest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DirectCast(FastColoredTextBox1.SpellCheck(), iTestHarness).SetupControl(FastColoredTextBox1)

        Me.EnableSpellCheck()


        PropertyGrid1.SelectedObject = FastColoredTextBox1.SpellCheck

        Dim propToolBoxIcon As New ToolboxBitmapAttribute(GetType(FastColoredTextBox))
        Using b As Bitmap = DirectCast(propToolBoxIcon.GetImage(GetType(FastColoredTextBox), True), Bitmap)
            Me.Icon = Icon.FromHandle(b.GetHicon)
        End Using
    End Sub

End Class
