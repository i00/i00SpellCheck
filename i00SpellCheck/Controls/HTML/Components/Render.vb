Partial Class HTMLParser
    Public Shared Function PaintHTML(ByVal HTML As String, Optional ByVal g As Graphics = Nothing, Optional ByVal MaxWidth As Single = -1, Optional ByVal status As Status = Nothing) As SizeF
        HTML = Replace(HTML, vbCrLf, "<br>")
        Dim _textLines As New List(Of TextLine)()
        Dim currWdth As Single = 0, currHgth As Single = 0
        If status Is Nothing Then
            status = New Status()
            status.Font = New STRFont(System.Drawing.SystemFonts.DefaultFont)
            status.Brush = New STRBrush(Color.FromKnownColor(KnownColor.ControlText))
        End If
        Dim font As Font = status.Font.GetRealFont()
        Dim values As IList(Of Element) = Nothing
        Dim lastSpaceSize As Double = 0
        Dim _totalHeight As Single = 0
        Dim _labelsPositions As New Dictionary(Of String, Element)
        Dim _totalWidth As Single = 0

        Dim b As Bitmap = Nothing
        If g Is Nothing Then
            b = New Bitmap(1, 1)
            g = Graphics.FromImage(b)
        End If
        Try
            If MaxWidth = -1 Then
                status.WordWrap = False
            End If

            Dim _HTMLElements As New Elements
            _HTMLElements.Parse(HTML, status)


            values = _HTMLElements.Value

            Dim newLine As Boolean = False
            For i As Integer = 1 To values.Count - 1
                If values(i).Type = ElementType.Status Then
                    If values(i).Status.Image IsNot Nothing AndAlso values(i).Status.Image.Image IsNot Nothing Then
                        If values(i).Status.Image.Size.IsEmpty = False Then
                            'use this size
                            values(i).Size = values(i).Status.Image.Size
                        Else
                            values(i).Size = values(i).Status.Image.Image.Size
                        End If
                        GoTo SizeElement
                    Else
                        newLine = (status.Alignment <> values(i).Status.Alignment) OrElse (status.WordWrap <> values(i).Status.WordWrap)
                        status = values(i).Status
                        font = status.Font.GetRealFont()
                        newLine = newLine Or status.NewLine

                        If newLine Then
                            newLine = False
                            _textLines.Add(New TextLine(currWdth, currHgth, i - 1))
                            _totalHeight += currHgth
                            currWdth = 0
                            currHgth = 0
                        End If
                    End If
                ElseIf values(i).Type = ElementType.HTML Then
                    If values(i).HTML.Type = HTMLParser.PartType.Text Then
                        values(i).Size = g.MeasureString(System.Web.HttpUtility.HtmlDecode(values(i).HTML.Value), font)
                    Else
                        If (values(i).Size.Width < 0) OrElse (values(i).Size.Height < 0) Then
                            Dim txt As String = values(i).HTML.Value
                            If txt = "" Then
                                txt = ""
                            End If
                            Dim size As SizeF = g.MeasureString(System.Web.HttpUtility.HtmlDecode(txt), font)
                            If values(i).Size.Width > 0 Then
                                size.Width = values(i).Size.Width
                            End If
                            If values(i).Size.Height > 0 Then
                                size.Height = values(i).Size.Height
                            End If
                            values(i).Size = size
                        End If
                    End If
SizeElement:
                    If i = 0 Then
                        currWdth = values(i).Size.Width
                        currHgth = values(i).Size.Height
                    Else
                        Dim spaceSize As Single = 0
                        If spaceSize < 0 Then
                            spaceSize = CSng(values(i).SpaceSize + lastSpaceSize) / 2
                        End If

                        If ((status.WordWrap) AndAlso (currWdth + values(i).Size.Width >= MaxWidth)) Then
                            newLine = False
                            _textLines.Add(New TextLine(currWdth, currHgth, i - 1))
                            _totalHeight += currHgth
                            currWdth = values(i).Size.Width
                            If currWdth > _totalWidth Then _totalWidth = currWdth
                            currHgth = values(i).Size.Height
                        Else
                            If currWdth > 0 Then
                                currWdth += spaceSize
                            End If
                            currWdth += values(i).Size.Width
                            If currWdth > _totalWidth Then _totalWidth = currWdth
                            currHgth = Math.Max(currHgth, values(i).Size.Height)
                        End If
                    End If
                    lastSpaceSize = values(i).SpaceSize
                End If
            Next
            _textLines.Add(New TextLine(currWdth, currHgth, values.Count - 1))
            _totalHeight += currHgth

        Catch ex As Exception
        Finally
            If b IsNot Nothing Then
                'we were using a graphics buffer
                g.Dispose()
                b.Dispose()
            End If
        End Try

        PaintHTML = New SizeF(_totalWidth, _totalHeight)
        If b IsNot Nothing Then
            'don't paint
            Exit Function
        End If

        If values Is Nothing Then Exit Function


        Dim brush As Brush = status.Brush.GetRealBrush()
        font = status.Font.GetRealFont()

        Dim currElement As Integer = 1
        Dim left As Single = 0, top As Single = 0

        For Each line As TextLine In _textLines
            If values(currElement).Type = ElementType.Status Then
                status = values(currElement).Status
                brush = status.Brush.GetRealBrush()
                font = status.Font.GetRealFont()
            End If

            Select Case status.Alignment
                Case ContentAlignment.TopCenter
                    left = (MaxWidth - line.Width) / 2
                    Exit Select
                Case ContentAlignment.TopRight
                    left = MaxWidth - line.Width
                    Exit Select
                Case Else
                    left = 0
                    Exit Select
            End Select

            lastSpaceSize = 0
            While currElement <= line.LastElement
                If values(currElement).Type = ElementType.Status Then
                    If values(currElement).Status.Image IsNot Nothing AndAlso values(currElement).Status.Image.Image IsNot Nothing Then
                        Dim irect As New Rectangle(CInt(Math.Truncate(left)), CInt(Math.Truncate(top + line.Height - values(currElement).Size.Height)), CInt(Math.Truncate(values(currElement).Size.Width)), CInt(Math.Truncate(values(currElement).Size.Height)))
                        values(currElement).DisplayedRect = irect
                        g.DrawImage(values(currElement).Status.Image.Image, values(currElement).DisplayedRect)
                        If lastSpaceSize = 0 Then
                            lastSpaceSize = values(currElement).SpaceSize
                        End If
                        Dim spaceSize As Single = 0
                        left += values(currElement).Size.Width + spaceSize
                        lastSpaceSize = values(currElement).SpaceSize
                    Else
                        status = values(currElement).Status
                        brush = status.Brush.GetRealBrush()
                        font = status.Font.GetRealFont()
                    End If
                End If

                If values(currElement).Type = ElementType.HTML Then
                    If values(currElement).HTML.Type = HTMLParser.PartType.Text Then
                        g.DrawString(System.Web.HttpUtility.HtmlDecode(values(currElement).HTML.Value), font, brush, left, top + line.Height - values(currElement).Size.Height)
                    Else
                        Dim label As HTMLParser.Label = TryCast(values(currElement).HTML, HTMLParser.Label)
                        Dim rect As New RectangleF(left, top + line.Height - values(currElement).Size.Height, values(currElement).Size.Width, values(currElement).Size.Height)
                        Dim irect As New Rectangle(CInt(Math.Truncate(left)), CInt(Math.Truncate(top + line.Height - values(currElement).Size.Height)), CInt(Math.Truncate(values(currElement).Size.Width)), CInt(Math.Truncate(values(currElement).Size.Height)))

                        g.DrawString(System.Web.HttpUtility.HtmlDecode(label.Value), font, brush, rect)

                        values(currElement).DisplayedRect = irect
                        _labelsPositions.Add(label.ID, values(currElement))
                    End If

                    If lastSpaceSize = 0 Then
                        lastSpaceSize = values(currElement).SpaceSize
                    End If

                    Dim spaceSize As Single = 0
                    left += values(currElement).Size.Width + spaceSize
                    lastSpaceSize = values(currElement).SpaceSize
                End If
                currElement += 1
            End While
            top += line.Height
        Next
    End Function

End Class
