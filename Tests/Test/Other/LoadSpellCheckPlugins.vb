Public Class LoadSpellCheckPlugins

    Public Class Types(Of T)
        Public Shared Function GetClassesOfType(ByVal a As System.Reflection.Assembly) As List(Of Types(Of T))
            Return (From xItem In a.GetTypes() Where GetType(T).IsAssignableFrom(xItem) AndAlso xItem.IsAbstract = False AndAlso xItem.IsPublic = True Select New Types(Of T) With {.Type = xItem}).ToList
        End Function

        Public Type As Type
        Public Function CreateObject() As T
            Return DirectCast(System.Activator.CreateInstance(Type), T)
        End Function
    End Class

    Public Shared Function Controls() As List(Of Control)
        Controls = New List(Of Control)
        Dim aMasterSpellCheck = System.Reflection.Assembly.Load("i00SpellCheck")
        For Each file In FileIO.FileSystem.GetFiles(IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location()), FileIO.SearchOption.SearchTopLevelOnly, "*.dll", "*.exe")
            Try
                Dim a = System.Reflection.Assembly.LoadFile(file)
                Dim t As Type = GetType(i00SpellCheck.SpellCheckControlBase)

                Dim Classes = Types(Of i00SpellCheck.SpellCheckControlBase).GetClassesOfType(a)
                For Each item In Classes
                    Try
                        Dim o = item.CreateObject
                        If o.ControlType.IsAbstract Then
                            'cannot create instance ... lets go through everything and try to create a control that comes from this
                            Dim ctls = OtherControls(a, o.ControlType)
                            Controls.AddRange(ctls)
                        Else
                            Dim ctl = TryCast(o.ControlType.Module.Assembly.CreateInstance(o.ControlType.FullName), Control)
                            Controls.Add(ctl)
                        End If
                    Catch ex As Exception

                    End Try
                Next
            Catch ex As Exception

            End Try
        Next

    End Function

    Private Shared Function OtherControls(ByVal a As System.Reflection.Assembly, ByVal BaseType As Type) As List(Of Control)
        OtherControls = New List(Of Control)
        Dim AssembliesToCheck = a.GetReferencedAssemblies.ToList
        AssembliesToCheck.Add(a.GetName)

        For Each item In AssembliesToCheck
            Try
                Dim AssemblyToCheck = System.Reflection.Assembly.Load(item.FullName)

                Dim CreatableControls = (From xItem In AssemblyToCheck.GetTypes() Where BaseType.IsAssignableFrom(xItem) AndAlso xItem.IsAbstract = False AndAlso xItem.IsPublic = True).ToList
                For Each CreateItem In CreatableControls
                    Try
                        Dim ctl = DirectCast(System.Activator.CreateInstance(CreateItem), Control)
                        OtherControls.Add(ctl)
                    Catch ex As Exception

                    End Try
                Next
            Catch ex As Exception

            End Try
        Next
    End Function

End Class
