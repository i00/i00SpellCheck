Imports System.IO
Imports System.Runtime.Serialization

Public Class Serialize

#Region "Serialization classes for save features"

    <Serializable()> _
    Private Class CompressedObject
        Public CompressedObject As Byte()
        Public Sub New(ByVal CompressedObject As Byte())
            Me.CompressedObject = CompressedObject
        End Sub
    End Class

    <Serializable()> _
    Private Class EncryptedObject
        Public EncryptedObject As Byte()
        Public Sub New(ByVal EncryptedObject As Byte())
            Me.EncryptedObject = EncryptedObject
        End Sub
    End Class

#End Region

    Private Shared Sub StreamSerialize(ByRef s As Stream, ByVal Obj As Object, Optional ByVal Compress As Boolean = True, Optional ByVal EncryptKey As String = Nothing)
        Dim b As New Formatters.Binary.BinaryFormatter
        If Compress = True Then
            'we want to compress this object to a byte array
            Obj = New CompressedObject(CompressionFunctions.CompressByte(ByteArrSerialize(Obj)))
        End If
        If EncryptKey <> "" Then
            Obj = New EncryptedObject(Encryption.EncryptByte(ByteArrSerialize(Obj), EncryptKey))
        End If
        b.Serialize(s, Obj)
    End Sub

    Public Shared Sub FileSerialize(ByVal File As String, ByVal Obj As Object, Optional ByVal Compress As Boolean = True, Optional ByVal EncryptKey As String = Nothing)
        Using s As New FileStream(File, FileMode.Create, FileAccess.ReadWrite)
            StreamSerialize(s, Obj, Compress, EncryptKey)
            s.Close()
        End Using
    End Sub

    Public Shared Function FileDeserialize(ByVal File As String, Optional ByVal DecryptKey As String = Nothing) As Object
        FileDeserialize = Nothing
        Dim theError As Exception = Nothing
        Using s As New FileStream(File, FileMode.Open, FileAccess.Read)
            FileDeserialize = StreamDeserialize(s, DecryptKey)
            s.Close()
        End Using
    End Function

    Private Shared Function StreamDeserialize(ByVal s As Stream, Optional ByVal DecryptKey As String = Nothing) As Object
        Dim b As New Formatters.Binary.BinaryFormatter
        StreamDeserialize = b.Deserialize(s)

StartFileTypeCheck:
        'lets see if the file was compressed or Encrypted 
        If TypeOf StreamDeserialize Is CompressedObject Then
            'we were compressed - so lets decompress
            StreamDeserialize = ByteArrDeserialize(CompressionFunctions.DecompressByte(CType(StreamDeserialize, CompressedObject).CompressedObject))
            GoTo StartFileTypeCheck
        Else
            'no compression
            'return as is
        End If

        If TypeOf StreamDeserialize Is EncryptedObject Then
            'we were compressed - so lets decompress
            StreamDeserialize = ByteArrDeserialize(Encryption.DecryptByte(CType(StreamDeserialize, EncryptedObject).EncryptedObject, DecryptKey))
            GoTo StartFileTypeCheck
        Else
            'no compression
            'return as is
        End If
    End Function

    Private Shared Function ByteArrSerialize(ByVal Obj As Object) As Byte()
        Using MS As New MemoryStream
            Dim BF As New Formatters.Binary.BinaryFormatter
            BF.Serialize(MS, Obj)
            ByteArrSerialize = MS.ToArray
            MS.Close()
        End Using
    End Function

    Private Shared Function ByteArrDeserialize(ByVal SerializedData() As Byte) As Object
        Using MS As New MemoryStream(SerializedData)
            Dim BF As New Formatters.Binary.BinaryFormatter
            ByteArrDeserialize = BF.Deserialize(MS)
            MS.Close()
        End Using
    End Function

End Class
