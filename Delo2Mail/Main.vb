Imports System.Text.RegularExpressions
Imports System.Net
Module Delo2Mail

    Sub Main()
        'Console.WriteLine(Settings.String("Message Pages\SMTPPassword", "default"))
        If Array.IndexOf(Environment.GetCommandLineArgs, "-sendonly") = -1 Then DownloadMessages()
        'If Array.IndexOf(Environment.GetCommandLineArgs, "-downloadonly") = -1 Then SendMessages()
        My.Log.End()
    End Sub

    Private Sub DownloadMessages()
        Dim hr As New HTTPReader
        Dim iMsgs As Integer = 0

        If hr.Register(My.Settings.LoginUrl, My.Settings.LoginString, 5) = True Then
            Dim sdUrls As Dictionary(Of String, String) = My.Settings.MessageListUrl
            Dim ini_ser As New IniSerializer(Of Message)
            '<Message pages iteration>
            For Each k As String In sdUrls.Keys
                Dim msgs As List(Of Message) = ParseMessageListPage(hr.DownloadPage(k, sdUrls(k), 5))
                If msgs.Count > 0 Then
                    '<Messages iteration>
                    For Each msg As Message In msgs
                        ' skip old messages
                        If msg.Date < My.Settings.EarliestDateFilter Then Continue For
                        '<New message download and parsing>
                        If Not FileIO.DirectoryExists("Messages", msg.ID) Then
                            ParseMessageDetailPage( _
                                                   msg, _
                                                   hr.DownloadPage(msg.Subject, My.Settings.InnerUrlPrefix & msg.Href, 5) _
                                                   )
                            msg.Category = k
                            Dim strStorePath As String = FileIO.ProvideDirectory("Messages", msg.ID)
                            '<New message files store>
                            If String.IsNullOrEmpty(strStorePath) = False Then
                                'store info file
                                Try
                                    My.Log.Write("Storing <" & strStorePath & "\_message.info" & ">... ")
                                    Using fs As New IO.FileStream(strStorePath & "\_message.info", IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.Read)
                                        ini_ser.Serialize(fs, msg)
                                        My.Log.WriteLine("done.")
                                    End Using
                                Catch ex As Exception
                                    My.Log.WriteLine("error: Could not write.")
                                End Try

                                ' store attachment files
                                For Each a As String In msg.Attachments.Keys
                                    If hr.DownloadFile(My.Settings.InnerUrlPrefix & a, strStorePath & "\" & msg.Attachments(a), 5) = False Then
                                        msg.ErrorMessage = String.Format("Файлы-вложения не были загружены. Пройдите по ссылке <a href='{0}'>{0}</a>", _
                                                                         My.Settings.InnerUrlPrefix & msg.Href)
                                    End If
                                Next
                                'FileIO.WriteText(MessageToHTML(msg), strStorePath & "\_message.htm")
                            End If
                            '</New message files store>
                            iMsgs += 1
                        End If
                        '</New message download and parsing>

                    Next
                    '</Messages iteration>

                End If
                msgs.Clear() : msgs = Nothing
            Next
            '</Message pages iteration>
            sdUrls.Clear() : sdUrls = Nothing
        End If
        My.Log.WriteLine(iMsgs.ToString & " new message(s) downloaded." & Environment.NewLine)
    End Sub


#Region "Parsing"
    Private Function ParseMessageListPage(ByVal Text As String) As List(Of Message)
        Dim retval As New List(Of Message)
        If String.IsNullOrEmpty(Text) Then Return retval

        ' get message table from page html 
        Dim strMsgTbl As String = Regex.Match( _
                                    Text, _
                                    My.Settings.MessageTableTemplate, _
                                    RegexOptions.Singleline Or RegexOptions.IgnoreCase _
                                              ).Value
        If String.IsNullOrEmpty(strMsgTbl) Then Return retval

        ' now get message entries
        Dim m As MatchCollection = Regex.Matches( _
                                    strMsgTbl, _
                                    My.Settings.MessageRowTemplate, _
                                    RegexOptions.Singleline Or RegexOptions.IgnoreCase _
                                                )
        If m.Count = 0 Then
            strMsgTbl = Nothing
            m = Nothing
            Return retval
        End If

        ' create a dictionary for each message entry
        Dim requisites1Turn As Dictionary(Of String, String) = My.Settings.FirstPageRequisiteTemplates
        For itm As Integer = 0 To m.Count - 1
            Dim msg As New Message
            For Each k As String In requisites1Turn.Keys
                Try
                    msg.Detail.Add( _
                            k, _
                            Regex.Match(m(itm).Value, requisites1Turn(k), RegexOptions.Singleline Or RegexOptions.IgnoreCase).Value _
                            )
                Catch : End Try
            Next
            retval.Add(msg)
        Next

        strMsgTbl = Nothing
        requisites1Turn.Clear() : requisites1Turn = Nothing
        m = Nothing

        Return retval
    End Function
    Private Sub ParseMessageDetailPage(Message As Message, ByVal Detail As String)
        If String.IsNullOrEmpty(Detail) Then Exit Sub

        Dim kiReqs2Turn As Dictionary(Of String, String) = My.Settings.SecondPageRequisiteTemplates
        Dim strDetTbl As String = Regex.Match(Detail, _
                                              My.Settings.MessageDetailTemplate, _
                                              RegexOptions.Singleline Or RegexOptions.IgnoreCase).Value
        For Each k As String In kiReqs2Turn.Keys
            Message.Detail.Add(k, Regex.Replace( _
                                              Regex.Match(strDetTbl, kiReqs2Turn(k), RegexOptions.Singleline Or RegexOptions.IgnoreCase).Value, _
                                              "\s+", _
                                              " ") _
                                )
        Next
        strDetTbl = Nothing
        kiReqs2Turn.Clear() : kiReqs2Turn = Nothing

        Dim m As MatchCollection = Regex.Matches(Detail, My.Settings.AttachmentTemplate, RegexOptions.Singleline Or RegexOptions.IgnoreCase)
        For itm As Integer = 0 To m.Count - 1
            Message.Attachments.Add( _
                                    m(itm).Value, _
                                    Regex.Match(m(itm).Value, My.Settings.AttachmentFileNameTemplate, RegexOptions.Singleline Or RegexOptions.IgnoreCase).Value _
                                    )
        Next
    End Sub
#End Region

#Region "Converting"
    Private Function MessageToHTML(ByVal Message As Dictionary(Of String, String)) As String
        Dim retval As New System.Text.StringBuilder(My.Settings.HTMLStoreTemplate)
        For Each k As String In Message.Keys
            retval = retval.Replace("{" & k & "}", If(Message.ContainsKey(k), Message(k), "&nbsp;"))
        Next
        Return retval.ToString
    End Function
#End Region
End Module

<Serializable> Friend Class Message
    <IniValue("Detail", IniSerializationType.Dictionary)> Private _detail As Dictionary(Of String, String)
    Private _fileUrls As Dictionary(Of String, String)
    Sub New()
        _detail = New Dictionary(Of String, String)
        _fileUrls = New Dictionary(Of String, String)
    End Sub
    Public Property Category As String
    Public Property ErrorMessage As String
    Public ReadOnly Property ID As String
        Get
            If _detail.ContainsKey("ID") = False Then
                _detail.Add("ID", System.DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss-ffff"))
            End If
            Return _detail("ID")
        End Get
    End Property
    Public ReadOnly Property Href As String
        Get
            If _detail.ContainsKey("Href") Then Return _detail("Href")
            Return String.Empty
        End Get
    End Property
    Public ReadOnly Property Subject As String
        Get
            If _detail.ContainsKey("Subject") Then Return _detail("Subject")
            Return String.Empty
        End Get
    End Property
    Public ReadOnly Property [Date] As Date
        Get
            Dim d As Date = System.DateTime.Now
            If _detail.ContainsKey("Date") Then
                System.DateTime.TryParse(_detail("Date"), d)
            End If
            Return d
        End Get
    End Property
    Public ReadOnly Property Detail As Dictionary(Of String, String)
        Get
            Return _detail
        End Get
    End Property
    Public ReadOnly Property Attachments As Dictionary(Of String, String)
        Get
            Return _fileUrls
        End Get
    End Property
End Class