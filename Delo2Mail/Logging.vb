Imports System.IO
Namespace My
    Friend Class Log
        Private Shared lwr As LogWriter
        Private Shared isLwr As Byte

        Public Shared Sub Write(ByVal Text As String, Color As ConsoleColor)
            If String.IsNullOrEmpty(Text.Trim) Then Exit Sub
            If isLwr = 0 Then
                Try
                    lwr = New LogWriter : lwr.Open("Logs")
                    isLwr = 1
                Catch
                    Console.ForegroundColor = ConsoleColor.DarkRed
                    Console.WriteLine("Error: could not create log file")
                    isLwr = 2
                End Try
            End If
            If isLwr = 1 Then lwr.Write(Text)
            Console.ForegroundColor = Color
            Console.Write(Text)
            If Text.EndsWith(Environment.NewLine) Then Console.ForegroundColor = ConsoleColor.White
        End Sub
        Public Shared Sub Write(ByVal Text As String)
            Write(Text, ConsoleColor.White)
        End Sub
        Public Shared Sub WriteLine(ByVal Text As String, Color As ConsoleColor)
            Write(Text & If(Text.EndsWith(Environment.NewLine), "", Environment.NewLine))
        End Sub
        Public Shared Sub WriteLine(ByVal Text As String)
            WriteLine(Text, ConsoleColor.White)
        End Sub
        Public Shared Sub [End]()
            If Not lwr Is Nothing Then : lwr.Close() : lwr = Nothing : End If
        End Sub
    End Class

    Friend Class LogWriter
        Dim fs As FileStream, sr As StreamWriter
        Dim tmFlg As Boolean

        Public Sub Open(ByVal LogDirectory As String)
            If Not IO.Directory.Exists(LogDirectory) Then IO.Directory.CreateDirectory(LogDirectory)
            If LogDirectory.EndsWith("\") = False Then LogDirectory &= "\"
            LogDirectory &= Date.Now.ToString("dd-MM-yyyy HH-mm-ss") & ".txt"
            fs = New FileStream(LogDirectory, FileMode.Create, FileAccess.Write, FileShare.Read)
            sr = New StreamWriter(fs, System.Text.Encoding.GetEncoding(1251))
            tmFlg = True
        End Sub
        Public Sub Write(ByVal Text As String)
            If tmFlg Then
                Text = "[" & Date.Now.ToString("HH:mm:ss") & "]  " & Text
            End If
            sr.Write(Text) : sr.Flush()
            tmFlg = If(Text.EndsWith(Environment.NewLine), True, False)
        End Sub

        Public Sub Close()
            sr.Flush() : sr.Close() : sr.Dispose()
            fs.Close() : fs.Dispose()
        End Sub
    End Class
End Namespace