Imports System.IO
Imports System.IO.Compression

Public Class CompressionFunctions

    Public Shared Function CompressByte(ByVal byteSource() As Byte) As Byte()
        ' Create a GZipStream object and memory stream object to store compressed stream
        Dim objMemStream As New MemoryStream()
        Dim objGZipStream As New GZipStream(objMemStream, CompressionMode.Compress, True)
        objGZipStream.Write(byteSource, 0, byteSource.Length)
        objGZipStream.Dispose()
        objMemStream.Position = 0
        ' Write compressed memory stream into byte array
        Dim buffer(CInt(objMemStream.Length)) As Byte
        objMemStream.Read(buffer, 0, buffer.Length)
        objMemStream.Dispose()
        Return buffer
    End Function

    Public Shared Function DecompressByte(ByVal byteCompressed() As Byte) As Byte()

        Try
            ' Initialize memory stream with byte array.
            Dim objMemStream As New MemoryStream(byteCompressed)

            ' Initialize GZipStream object with memory stream.
            Dim objGZipStream As New GZipStream(objMemStream, CompressionMode.Decompress)

            ' Define a byte array to store header part from compressed stream.
            Dim sizeBytes(3) As Byte

            ' Read the size of compressed stream.
            objMemStream.Position = objMemStream.Length - 5
            objMemStream.Read(sizeBytes, 0, 4)

            Dim iOutputSize As Integer = BitConverter.ToInt32(sizeBytes, 0)

            ' Posistion the to point at beginning of the memory stream to read
            ' compressed stream for decompression.
            objMemStream.Position = 0

            Dim decompressedBytes(iOutputSize - 1) As Byte

            ' Read the decompress bytes and write it into result byte array.
            objGZipStream.Read(decompressedBytes, 0, iOutputSize)

            objGZipStream.Dispose()
            objMemStream.Dispose()

            Return decompressedBytes

        Catch ex As Exception
            Return Nothing
        End Try

    End Function

End Class
