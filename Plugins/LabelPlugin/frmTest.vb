Imports i00SpellCheck

Public Class frmTest

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.EnableSpellCheck()
        PropertyGrid1.SelectedObject = Label1.SpellCheck

        Dim propToolBoxIcon As New ToolboxBitmapAttribute(GetType(Label))
        Using b As Bitmap = DirectCast(propToolBoxIcon.GetImage(GetType(Label), True), Bitmap)
            Me.Icon = Icon.FromHandle(b.GetHicon)
        End Using
    End Sub

End Class
