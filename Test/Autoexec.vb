Public Class Autoexec

    Public Shared Sub Main()
        Application.EnableVisualStyles()

        Using Form1 As New Form1
            Form1.ShowDialog()
        End Using

        My.Settings.Save()
    End Sub

End Class
