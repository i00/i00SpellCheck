Imports System.Security.Cryptography
Imports System.Text
Imports System.IO

Public Class Encryption

    Public Shared Function EncryptByte(ByVal Data As Byte(), ByVal Key As String) As Byte()
        Key = FixKey(Key, 16, 24, 32)
        Dim bytes() As Byte = ASCIIEncoding.ASCII.GetBytes(Key)
        Dim ms As New MemoryStream()
        Dim alg As Rijndael = Rijndael.Create()
        alg.Key = bytes
        alg.IV = New Byte() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}
        Dim cs As New CryptoStream(ms, alg.CreateEncryptor(), CryptoStreamMode.Write)
        cs.Write(Data, 0, Data.Length)
        cs.Close()
        Dim encryptedData As Byte() = ms.ToArray()
        Return encryptedData
    End Function

    Public Shared Function DecryptByte(ByVal cipherData As Byte(), ByVal Key As String) As Byte()
        Try
            Key = FixKey(Key, 16, 24, 32)
            Dim bytes() As Byte = ASCIIEncoding.ASCII.GetBytes(Key)
            Dim ms As New MemoryStream()
            Dim alg As Rijndael = Rijndael.Create()
            alg.Key = bytes
            alg.IV = New Byte() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}
            Dim cs As New CryptoStream(ms, alg.CreateDecryptor(), CryptoStreamMode.Write)
            cs.Write(cipherData, 0, cipherData.Length)
            cs.Close()
            Dim decryptedData As Byte() = ms.ToArray()
            Return decryptedData
        Catch ex As Exception
            Throw New Exception("Error decrypting - please make sure the key is correct")
        End Try
    End Function

    Private Shared Function FixKey(ByVal Key As String, ByVal ParamArray FixLength() As Integer) As String
        Dim SortedLength = From xItem In FixLength Order By xItem
        For Each item In SortedLength
            If Len(Key) <= item Then
                If Key Is Nothing Then Key = " "
                Key = Key.PadRight(item)
                Exit For
            End If
        Next
        If Key.Length > FixLength(UBound(FixLength)) Then
            Key = Strings.Left(Key, FixLength(UBound(FixLength)))
        End If
        FixKey = Key
    End Function

End Class
