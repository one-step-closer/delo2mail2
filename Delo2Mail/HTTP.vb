Imports System.Net

Friend Class HTTPReader
    Private ckSession As Cookie, isRegd As Byte

    Public Function DownloadPage(ByVal PageName As String, ByVal PageUrl As String, ByVal RetryCount As Byte) As String
        Dim rcnt As Integer = 0
        Dim retval As String
        My.Log.Write("Downloading message page <" & PageName & ">... ", ConsoleColor.Magenta)
        While rcnt < RetryCount
            retval = DownloadPage(PageUrl)
            If retval <> "" Then
                Return retval
            End If
            My.Log.Write("Retrying message page download... ")
            rcnt += 1
        End While
        My.Log.WriteLine("Could not retrieve page. Aborted")
        Return String.Empty
    End Function
    Private Function DownloadPage(ByVal Url As String) As String
        Dim strRetval As String = String.Empty

        Dim hReq As HttpWebRequest = Nothing
        Dim hResp As HttpWebResponse = Nothing
        Dim sr As IO.StreamReader = Nothing

        Try
            hReq = DirectCast(WebRequest.Create(Url), HttpWebRequest)
            hReq.Method = "GET"
            hReq.Accept = "*/*"
            hReq.Headers.Add("Accept-Language", "ru")
            hReq.CookieContainer = New CookieContainer
            hReq.CookieContainer.Add(ckSession)

            hResp = DirectCast(hReq.GetResponse(), HttpWebResponse)

            sr = New IO.StreamReader(hResp.GetResponseStream, System.Text.Encoding.GetEncoding(1251))
            strRetval = sr.ReadToEnd()
            My.Log.WriteLine(hResp.ContentLength.ToString & " bytes received.")
        Catch ex As Exception
            My.Log.WriteLine("Error: " & ex.Message)
        Finally
            If Not hResp Is Nothing Then hResp.Close()
            If Not sr Is Nothing Then sr.Close()
        End Try
        Return strRetval
    End Function

    Public Function DownloadFile(ByVal Url As String, ByVal FileName As String, ByVal RetryCount As Byte) As Boolean
        Dim rcnt As Integer = 0
        Dim retval As Boolean
        My.Log.Write("Downloading file <" & IO.Path.GetFileName(FileName) & "> ")
        While rcnt < RetryCount
            retval = DownloadFile(Url, FileName)
            If retval Then Return retval
            My.Log.Write("Retrying file download... ")
            rcnt += 1
        End While
        My.Log.WriteLine("Could not retrieve file. Aborted.")
        Return retval
    End Function
    Private Function DownloadFile(ByVal Url As String, ByVal LocalPath As String) As Boolean
        Dim retval As Boolean
        Dim hreq As HttpWebRequest = Nothing
        Dim hresp As HttpWebResponse = Nothing
        Dim fsresp As IO.FileStream = Nothing
        Try
            hreq = DirectCast(WebRequest.Create(Url), HttpWebRequest)
            hreq.Method = "GET"
            hreq.CookieContainer = New CookieContainer
            hreq.CookieContainer.Add(ckSession)

            hresp = DirectCast(hreq.GetResponse(), HttpWebResponse)
            If hresp.StatusCode = HttpStatusCode.OK Then
                My.Log.Write("(" & hresp.ContentLength.ToString & " bytes)... ")
                fsresp = New IO.FileStream(LocalPath, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.Read)
                Dim bts(1024) As Byte, btsRead As Integer
                btsRead = hresp.GetResponseStream.Read(bts, 0, bts.Length)
                Do While btsRead > 0
                    fsresp.Write(bts, 0, btsRead) : fsresp.Flush()
                    btsRead = hresp.GetResponseStream.Read(bts, 0, bts.Length)
                Loop
                My.Log.WriteLine("success.")
                retval = True
            Else
                My.Log.WriteLine("Error: File not found.")
            End If
        Catch ex As Exception
            My.Log.WriteLine("Error: " & ex.Message)
        Finally
            If Not hresp Is Nothing Then hresp.Close()
            If Not fsresp Is Nothing Then
                fsresp.Close() : fsresp.Dispose()
            End If
        End Try
        Return retval
    End Function


    Public Function Register(ByVal LoginUrl As String, ByVal Credentials As String, ByVal RetryCount As Byte) As Boolean
        Dim rcnt As Integer = 0
        ckSession = Nothing
        My.Log.Write("Logging on to server... ")
        While rcnt < RetryCount
            If Register(LoginUrl, Credentials) Then Return True
            My.Log.Write("Retrying register... ")
            rcnt += 1
        End While
        My.Log.WriteLine("Could not register. Aborted.")
        Return False
    End Function
    Private Function Register(ByVal LoginUrl As String, ByVal Credentials As String) As Boolean

        Dim hreq As HttpWebRequest = Nothing
        Dim hresp As HttpWebResponse = Nothing
        Dim hsr As IO.StreamWriter = Nothing
        Dim bErr As Byte

        Try
            hreq = DirectCast(WebRequest.Create(LoginUrl), HttpWebRequest)
            hreq.Method = "POST"
            hreq.ContentType = "application/x-www-form-urlencoded"
            hreq.ContentLength = Credentials.Length
            hreq.CookieContainer = New CookieContainer

            hsr = New IO.StreamWriter(hreq.GetRequestStream(), System.Text.Encoding.ASCII)
            hsr.Write(Credentials)
            hsr.Close()

        Catch
            My.Log.WriteLine("Error: Could not send registration data.")
            bErr = 1
        Finally
            If Not hsr Is Nothing Then : hsr.Close() : hsr.Dispose() : End If
        End Try

        If bErr = 0 Then
            Try
                hresp = DirectCast(hreq.GetResponse(), HttpWebResponse)
            Catch
                My.Log.WriteLine("Error: Could not get answer from server.")
                bErr = 1
            End Try
        End If

        If bErr = 0 Then
            Dim cookies As CookieCollection = hreq.CookieContainer.GetCookies(New Uri(GetBasicUrl(LoginUrl)))
            If cookies.Count > 0 Then
                My.Log.WriteLine("success.")
                ckSession = cookies.Item(cookies.Count - 1)
                ckSession.Expires = DateTime.Now.AddDays(1)
                Return True
            Else
                My.Log.WriteLine("Error: no cookies retrived.")
            End If
        End If
        Return False
    End Function

    Public ReadOnly Property Registered() As Boolean
        Get
            Return (isRegd = 1)
        End Get
    End Property


    Private Function GetBasicUrl(ByVal LoginUrl As String) As String
        Dim ios As Integer
        If LoginUrl.IndexOf("http://") > -1 Then
            ios = LoginUrl.IndexOf("/"c, 7)
        Else
            ios = LoginUrl.IndexOf("/"c)
        End If
        If ios = -1 Then Return LoginUrl
        Return LoginUrl.Substring(0, ios)
    End Function
End Class