Friend Class FileIO
    Public Shared Function DirectoryExists(ByVal ParamArray Names() As String) As Boolean
        Return IO.Directory.Exists(String.Join("\", Names).Replace("/", "\").Replace("\\", "\"))
    End Function
    Public Shared Function ProvideDirectory(ByVal ParamArray Names() As String) As String
        Dim retval As String = String.Join("\", Names).Replace("/", "\").Replace("\\", "\")
        If Not IO.Directory.Exists(retval) Then
            Try
                IO.Directory.CreateDirectory(retval)
                Return retval
            Catch
                My.Log.WriteLine("Error: Could not create directory <" & retval & ">")
                Return Nothing
            End Try
        Else
            Return retval
        End If
    End Function

    Public Shared Function WriteText(ByVal Text As String, ByVal FileName As String) As Boolean
        My.Log.Write("Storing <" & IO.Path.GetFileName(FileName) & ">... ")
        Dim fws As IO.FileStream = Nothing
        Dim fwr As IO.StreamWriter = Nothing
        Try
            fws = New IO.FileStream(FileName, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.Read)
            fwr = New IO.StreamWriter(fws, System.Text.Encoding.GetEncoding(1251))
            fwr.WriteLine(Text)
            My.Log.WriteLine("done.")
            Return True
        Catch
            My.Log.Write("error: Could not write.")
            Return False
        Finally
            If Not fwr Is Nothing Then : fwr.Close() : fwr.Dispose() : End If
            If Not fws Is Nothing Then : fws.Close() : fws.Dispose() : End If
        End Try
    End Function

    Public Shared Function ReadText(ByVal FileName As String) As String
        Dim fs As IO.FileStream = Nothing
        Dim sr As IO.StreamReader = Nothing
        Try
            fs = New IO.FileStream(FileName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
            sr = New IO.StreamReader(fs, System.Text.Encoding.GetEncoding(1251))
            ReadText = sr.ReadToEnd
        Catch
            My.Log.WriteLine("Error: Could not read <" & IO.Path.GetFileName(FileName) & ">")
            ReadText = String.Empty
        Finally
            If Not sr Is Nothing Then : sr.Close() : sr.Dispose() : End If
            If Not fs Is Nothing Then : fs.Close() : fs.Dispose() : End If
        End Try
    End Function

    Public Shared Function GetMIMEType(ByVal FileName As String) As String
        Select Case IO.Path.GetExtension(FileName)
            Case ".pdf"
                Return "application/pdf"
            Case ".doc"
                Return "application/msword"
            Case ".rtf"
                Return "application/msword"
            Case ".zip"
                Return "application/x-zip-compressed"
            Case ".jpg"
                Return "image/pjpeg"
            Case Else
                Return String.Empty
        End Select
    End Function
End Class