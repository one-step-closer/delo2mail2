Imports Word = Microsoft.Office.Interop.Word

Namespace My
    Module SettingsModule
        Private so As SettingsObject
        Private ReadOnly Property SettingsPath As String
            Get
                Dim myIniPath As String = Reflection.Assembly.GetCallingAssembly().Location
                If myIniPath.EndsWith(".exe", StringComparison.InvariantCultureIgnoreCase) Then
                    myIniPath = myIniPath.Substring(0, myIniPath.Length - 4) & ".ini"
                    If IO.File.Exists(myIniPath) Then
                        Return myIniPath
                    End If
                End If
                myIniPath = Nothing
                Return Nothing
            End Get
        End Property
        Public ReadOnly Property Settings As SettingsObject
            Get
                If so IsNot Nothing Then Return so
                If String.IsNullOrEmpty(SettingsPath) Then
                    so = New SettingsObject : Return so
                End If

                Try
                    Using sr As New IO.FileStream(SettingsPath, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
                        Using serializer As New IniSerializer(Of SettingsObject)
                            so = DirectCast(serializer.Deserialize(sr), SettingsObject)
                        End Using
                    End Using
                Catch
                    so = New SettingsObject
                End Try
                Return so
            End Get
        End Property
    End Module


    <Serializable> Friend Class SettingsObject
        Sub New()
            dictMessageListUrls = New Dictionary(Of String, String)
            dictFirstPageRequisiteTemplates = New Dictionary(Of String, String)
            dictSecondPageRequisiteTemplates = New Dictionary(Of String, String)
        End Sub
        ' --------------
        ' Authentication
        ' --------------
        <IniValue("Authentication", "LoginUrl", IniSerializationType.String)> Private strLoginUrl As String
        Public ReadOnly Property LoginUrl As String
            Get
                Return strLoginUrl
            End Get
        End Property
        <IniValue("Authentication", "LoginString", IniSerializationType.String)> Private strLoginString As String
        Public ReadOnly Property LoginString As String
            Get
                Return strLoginString
            End Get
        End Property

        ' --------------
        ' Message Pages
        ' --------------
        <IniValue("Message Pages", "InnerUrlPrefix", IniSerializationType.String)> Private strInnerUrlPrefix As String
        Public ReadOnly Property InnerUrlPrefix As String
            Get
                Return strInnerUrlPrefix
            End Get
        End Property
        <IniValue("Message Pages", "MessageListUrls", IniSerializationType.Dictionary)> Private dictMessageListUrls As Dictionary(Of String, String)
        Public ReadOnly Property MessageListUrl As Dictionary(Of String, String)
            Get
                Return dictMessageListUrls
            End Get
        End Property
        <IniValue("Message Pages", "EarliestDateFilter", IniSerializationType.Date)> Private dEarliestDateFilter As Date
        Public ReadOnly Property EarliestDateFilter As Date
            Get
                Return dEarliestDateFilter
            End Get
        End Property

        ' ----------------
        ' Message Parsing
        ' ----------------
        <IniValue("Message Parsing", "MessageTableTemplate", IniSerializationType.String)> Private strMessageTableTemplate As String
        Public ReadOnly Property MessageTableTemplate As String
            Get
                Return strMessageTableTemplate
            End Get
        End Property
        <IniValue("Message Parsing", "MessageRowTemplate", IniSerializationType.String)> Private strMessageRowTemplate As String
        Public ReadOnly Property MessageRowTemplate As String
            Get
                Return strMessageRowTemplate
            End Get
        End Property
        <IniValue("Message Parsing", "MessageDetailTemplate", IniSerializationType.String)> Private strMessageDetailTemplate As String
        Public ReadOnly Property MessageDetailTemplate As String
            Get
                Return strMessageDetailTemplate
            End Get
        End Property
        <IniValue("Message Parsing", "FirstPageRequisiteTemplates", IniSerializationType.Dictionary)> Private dictFirstPageRequisiteTemplates As Dictionary(Of String, String)
        Public ReadOnly Property FirstPageRequisiteTemplates As Dictionary(Of String, String)
            Get
                Return dictFirstPageRequisiteTemplates
            End Get
        End Property
        <IniValue("Message Parsing", "SecondPageRequisiteTemplates", IniSerializationType.Dictionary)> Private dictSecondPageRequisiteTemplates As Dictionary(Of String, String)
        Public ReadOnly Property SecondPageRequisiteTemplates As Dictionary(Of String, String)
            Get
                Return dictSecondPageRequisiteTemplates
            End Get
        End Property
        <IniValue("Message Parsing", "AttachmentTemplate", IniSerializationType.String)> Private strAttachmentTemplate As String
        Public ReadOnly Property AttachmentTemplate As String
            Get
                Return strAttachmentTemplate
            End Get
        End Property
        <IniValue("Message Parsing", "AttachmentFileNameTemplate", IniSerializationType.String)> Private strAttachmentFileNameTemplate As String
        Public ReadOnly Property AttachmentFileNameTemplate As String
            Get
                Return strAttachmentFileNameTemplate
            End Get
        End Property

        ' -------------
        ' Mail Sending
        ' -------------
        <IniValue("Mail Sending", "SMTPAddress", IniSerializationType.String)> Private strSMTPAddress As String
        Public ReadOnly Property SMTPAddress As String
            Get
                Return strSMTPAddress
            End Get
        End Property
        <IniValue("Mail Sending", "SMTPPort", IniSerializationType.Integer)> Private iSMTPPort As Integer
        Public ReadOnly Property SMTPPort As Integer
            Get
                Return iSMTPPort
            End Get
        End Property
        <IniValue("Mail Sending", "SMTPLogin", IniSerializationType.String)> Private strSMTPLogin As String
        Public ReadOnly Property SMTPLogin As String
            Get
                Return strSMTPLogin
            End Get
        End Property
        <IniValue("Mail Sending", "SMTPPassword", IniSerializationType.String)> Private strSMTPPassword As String
        Public ReadOnly Property SMTPPassword As String
            Get
                Return strSMTPPassword
            End Get
        End Property
        <IniValue("Mail Sending", "UseSSL", IniSerializationType.Boolean)> Private bUseSSL As Boolean
        Public ReadOnly Property UseSSL As Boolean
            Get
                Return bUseSSL
            End Get
        End Property
        <IniValue("Mail Sending", "MailFromAddress", IniSerializationType.String)> Private strMailFromAddress As String
        Public ReadOnly Property MailFromAddress As String
            Get
                Return strMailFromAddress
            End Get
        End Property
        <IniValue("Mail Sending", "MailToAddress", IniSerializationType.String)> Private strMailToAddress As String
        Public ReadOnly Property MailToAddress As String
            Get
                Return strMailToAddress
            End Get
        End Property

        ' -----
        ' Store
        ' -----
        <IniValue("Store", "HTMLStoreTemplate", IniSerializationType.String)> Private strHTMLStoreTemplate As String
        Public ReadOnly Property HTMLStoreTemplate As String
            Get
                Return strHTMLStoreTemplate
            End Get
        End Property

    End Class
End Namespace