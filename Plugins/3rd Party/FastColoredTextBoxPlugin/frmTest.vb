Imports FastColoredTextBoxNS
Imports i00SpellCheck

Public Class frmTest

#Region "FastColoredTextBox Formatting"

    Dim KeywordStyle As TextStyle = New TextStyle(Brushes.Blue, Nothing, FontStyle.Regular)
    Dim CommentStyle As TextStyle = New TextStyle(Brushes.Green, Nothing, FontStyle.Regular)
    Dim StringStyle As TextStyle = New TextStyle(Brushes.Brown, Nothing, FontStyle.Regular)

    Private Sub CSharpSyntaxHighlight(ByVal e As TextChangedEventArgs)
        'clear style of changed range
        e.ChangedRange.ClearStyle(KeywordStyle, CommentStyle, StringStyle)

        e.ChangedRange.SetStyle(StringStyle, """.*?""")
        e.ChangedRange.SetStyle(CommentStyle, "'.*$|(\s|^)rem\s.*$", System.Text.RegularExpressions.RegexOptions.Multiline Or System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        'DataTypes
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Boolean|Byte|Char|Date|Decimal|Double|Integer|Long|Object|SByte|Short|Single|String|UInteger|ULong|UShort|Variant)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'Operators
        e.ChangedRange.SetStyle(KeywordStyle, "\b(AddressOf|And|AndAlso|Is|IsNot|Like|Mod|New|Not|Or|OrElse|Xor)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'Constants
        e.ChangedRange.SetStyle(KeywordStyle, "\b(False|Me|MyBase|MyClass|Nothing|True)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'CommonKeywords
        e.ChangedRange.SetStyle(KeywordStyle, "\b(As|Of|New|End)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'CommonKeywords
        e.ChangedRange.SetStyle(KeywordStyle, "\b(As|Of|New|End)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'FunctionKeywords
        e.ChangedRange.SetStyle(KeywordStyle, "\b(CBool|CByte|CChar|CDate|CDec|CDbl|CInt|CLng|CObj|CSByte|CShort|CSng|CStr|CType|CUInt|CULng|CUShort|DirectCast|GetType|TryCast|TypeOf)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'ParamModifiers
        e.ChangedRange.SetStyle(KeywordStyle, "\b(ByRef|ByVal|Optional|ParamArray)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'AccessModifiers
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Friend|Private|Protected|Public)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'OtherModifiers
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Const|Custom|Default|Global|MustInherit|MustOverride|Narrowing|NotInheritable|NotOverridable|Overloads|Overridable|Overrides|Partial|ReadOnly|Shadows|Shared|Static|Widening|WithEvents|WriteOnly)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'Statements
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Throw|Stop|Return|Resume|AddHandler|RemoveHandler|RaiseEvent|Option|Let|GoTo|GoSub|Call|Continue|Dim|ReDim|Erase|On|Error|Exit)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'GlobalConstructs
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Namespace|Class|Imports|Implements|Inherits|Interface|Delegate|Module|Structure|Enum)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'TypeLevelConstructs
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Sub|Function|Handles|Declare|Lib|Alias|Get|Set|Property|Operator|Event)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'Constructs
        e.ChangedRange.SetStyle(KeywordStyle, "\b(SyncLock|Using|With|Do|While|Loop|Wend|Try|Catch|When|Finally|If|Then|Else|For|To|Step|Each|In|Next|Select|Case)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        'ContextKeywords
        e.ChangedRange.SetStyle(KeywordStyle, "\b(Ansi|Auto|Unicode|Preserve|Until)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        'Region
        e.ChangedRange.SetStyle(KeywordStyle, "^\s{0,}#region\b|^\s{0,}#end\s{1,}region\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Multiline)

        'clear folding markers
        e.ChangedRange.ClearFoldingMarkers()
        'set folding markers
        'e.ChangedRange.SetFoldingMarkers("^\s{0,}#region\b", "^\s{0,}#end\s{1,}region\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase) 'allow to collapse #region blocks
    End Sub

    Private Sub FastColoredTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As FastColoredTextBoxNS.TextChangedEventArgs) Handles FastColoredTextBox1.TextChanged
        CSharpSyntaxHighlight(e)
    End Sub

#End Region

#Region "Setup"

    Private Sub SetupFastColoredTextBox()
        FastColoredTextBox1.LeftBracket = "("
        FastColoredTextBox1.RightBracket = ")"

        FastColoredTextBox1.Text = "'Simple test to check spelling with 3rd party controls" & vbCrLf & _
               "'This test is done on the FastColoredTextBox (open source control) that is hosted on CodeProject" & vbCrLf & _
               "'The article can be found at: http://www.codeproject.com/Articles/161871/Fast-Colored-TextBox-for-syntax-highlighting" & vbCrLf & _
               "" & vbCrLf & _
               "'i00 Does not take credit for any work in the FastColoredTextBox.dll" & vbCrLf & _
               "'i00 is however solely responsible for the spellcheck plugin to interface with FastColoredTextBox" & vbCrLf & _
               "" & vbCrLf & _
               "'As you can see only comments and string blocks are corrected" & vbCrLf & _
               "'This is due to the SpellCheckMatch property being set" & vbCrLf & _
               "" & vbCrLf & _
               "'Click on a missspelled word to correct..." & vbCrLf & _
               "" & vbCrLf & _
               "Dim test = ""Test with some bad spellling!""" & vbCrLf & _
               "" & vbCrLf & _
               "#Region ""Char""" & vbCrLf & _
               "   " & vbCrLf & _
               "   ''' <summary>" & vbCrLf & _
               "   ''' Char and style" & vbCrLf & _
               "   ''' </summary>" & vbCrLf & _
               "   Public Structure CharStyle" & vbCrLf & _
               "       Public c As Char" & vbCrLf & _
               "       Public style As StyleIndex" & vbCrLf & _
               "   " & vbCrLf & _
               "       Public Sub CharStyle(ByVal ch As Char)" & vbCrLf & _
               "           c = ch" & vbCrLf & _
               "           Style = StyleIndex.None" & vbCrLf & _
               "       End Sub" & vbCrLf & _
               "   " & vbCrLf & _
               "   End Structure" & vbCrLf & _
               "   " & vbCrLf & _
               "#End Region"
    End Sub

#End Region

    Private Sub frmTest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetupFastColoredTextBox()

        Me.EnableSpellCheck()

        DirectCast(FastColoredTextBox1.SpellCheck(), SpellCheckFastColoredTextBox).SpellCheckMatch = "('.*$|"".*?"")"

        PropertyGrid1.SelectedObject = FastColoredTextBox1.SpellCheck

        Dim propToolBoxIcon As New ToolboxBitmapAttribute(GetType(FastColoredTextBox))
        Using b As Bitmap = DirectCast(propToolBoxIcon.GetImage(GetType(FastColoredTextBox), True), Bitmap)
            Me.Icon = Icon.FromHandle(b.GetHicon)
        End Using
    End Sub

End Class
