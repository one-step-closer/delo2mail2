Option Strict On
Imports System.Text.RegularExpressions
Imports System.Reflection
Imports System.Runtime.Serialization

Public Class IniSerializer(Of T)
    Implements IFormatter, IDisposable
    Private _binder As SerializationBinder
    Private _context As StreamingContext
    Private _srgSelector As ISurrogateSelector
    Private disposedValue As Boolean

    Sub New()
        _context = New StreamingContext(StreamingContextStates.All)
    End Sub
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
            End If
            _context = Nothing
            _binder = Nothing
            _srgSelector = Nothing
        End If
        Me.disposedValue = True
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Public Property Binder As Runtime.Serialization.SerializationBinder Implements Runtime.Serialization.IFormatter.Binder
        Get
            Return _binder
        End Get
        Set(value As Runtime.Serialization.SerializationBinder)
            _binder = value
        End Set
    End Property
    Public Property Context As Runtime.Serialization.StreamingContext Implements Runtime.Serialization.IFormatter.Context
        Get
            Return _context
        End Get
        Set(value As Runtime.Serialization.StreamingContext)
            _context = value
        End Set
    End Property
    Public Property SurrogateSelector As Runtime.Serialization.ISurrogateSelector Implements Runtime.Serialization.IFormatter.SurrogateSelector
        Get
            Return _srgSelector
        End Get
        Set(value As Runtime.Serialization.ISurrogateSelector)
            _srgSelector = value
        End Set
    End Property

    Public Sub Serialize(serializationStream As IO.Stream, graph As Object) Implements Runtime.Serialization.IFormatter.Serialize
        If graph Is Nothing Then Exit Sub
        If GetType(T).IsAssignableFrom(graph.GetType) = False Then
            Throw New SerializationException("Object type mismatch")
            Exit Sub
        End If

        Dim ca As IniValueAttribute = Nothing, strSect As String, strKey As String
        Dim dict As New Dictionary(Of String, Dictionary(Of String, String))
        Dim key_length As Integer = 0

        Dim minfo() As MemberInfo = FormatterServices.GetSerializableMembers(graph.GetType, _context)
        Dim mdata() As Object = FormatterServices.GetObjectData(graph, minfo)
        For itm = 0 To minfo.Length - 1
            If minfo(itm).GetCustomAttributes(False).Any(Function(c) c.GetType Is GetType(IniValueAttribute)) Then
                ca = DirectCast(minfo(itm).GetCustomAttributes(False).First(), IniValueAttribute)
            Else
                ca = Nothing
            End If
            If ca IsNot Nothing Then
                strSect = ca.Section
                strKey = ca.Key
            ElseIf DirectCast(minfo(itm), FieldInfo).IsPrivate Then
                Continue For
            Else
                strSect = String.Empty
                strKey = minfo(itm).Name
            End If
            If strKey.Length > key_length Then key_length = strKey.Length
            If dict.ContainsKey(strSect) = False Then
                dict.Add(strSect, New Dictionary(Of String, String))
            End If
            If dict(strSect).ContainsKey(strKey) = False Then
                dict(strSect).Add(strKey, Nothing)
            End If
            If ca IsNot Nothing Then
                dict(strSect)(strKey) = ca.ToString(mdata(itm))
            ElseIf mdata(itm) IsNot Nothing Then
                dict(strSect)(strKey) = mdata(itm).ToString
            End If
        Next
        Erase minfo : minfo = Nothing
        Erase mdata : mdata = Nothing
        ca = Nothing : strSect = Nothing : strKey = Nothing

        Using swr As New IO.StreamWriter(serializationStream)
            For Each s As String In dict.Keys
                If String.IsNullOrEmpty(s) = False Then swr.WriteLine("[" & s & "]")
                For Each k As String In dict(s).Keys
                    swr.WriteLine( _
                        (k & " = ").PadRight(key_length + 3) & _
                        dict(s)(k).Replace(Environment.NewLine, Environment.NewLine & New String(" "c, key_length + 3)) _
                                )
                Next
            Next
        End Using

        dict.Clear() : dict = Nothing
    End Sub
    Public Function Deserialize(serializationStream As IO.Stream) As Object Implements Runtime.Serialization.IFormatter.Deserialize
        Dim retval As Object = Activator.CreateInstance(Of T)()
        Dim minfo() As MemberInfo = FormatterServices.GetSerializableMembers(GetType(T), _context)
        Dim mdata(minfo.GetUpperBound(0)) As Object

        Using fr As New IO.StreamReader(serializationStream)
            Dim l As String, lastSection As String = String.Empty, lastKey As String = Nothing, lastValue As String = Nothing
            While Not fr.EndOfStream
                l = fr.ReadLine.Trim
                If String.IsNullOrWhiteSpace(l) OrElse l.StartsWith("//") Then Continue While
                If Regex.IsMatch(l, "^\[\w[\w\s]*\]") Then
                    lastSection = l.Trim("["c, "]"c, " "c).ToLower
                    lastKey = Nothing
                    lastValue = Nothing
                ElseIf Regex.IsMatch(l, "^\w+\s*\=") Then
                    lastKey = l.Substring(0, l.IndexOf("="c)).ToLower.Trim
                    lastValue = l.Substring(l.IndexOf("="c) + 1).Trim
                ElseIf String.IsNullOrEmpty(l) = False Then
                    If String.IsNullOrEmpty(lastKey) = False Then
                        If String.IsNullOrEmpty(lastValue) Then
                            lastValue = l.Trim
                        Else
                            lastValue &= Environment.NewLine & l.Trim
                        End If
                    End If
                End If
                If String.IsNullOrEmpty(lastKey) = False _
                    And String.IsNullOrWhiteSpace(lastValue) = False Then

                    Dim ca As IniValueAttribute = Nothing
                    Dim found As Boolean
                    For it As Integer = 0 To minfo.Count - 1
                        If minfo(it).GetCustomAttributes(False).Any(Function(c) c.GetType Is GetType(IniValueAttribute)) Then
                            ca = DirectCast(minfo(it).GetCustomAttributes(False).First(), IniValueAttribute)
                        End If
                        If ca IsNot Nothing AndAlso _
                            ca.Section.Equals(lastSection, StringComparison.InvariantCultureIgnoreCase) And _
                            ca.Key.Equals(lastKey, StringComparison.InvariantCultureIgnoreCase) Then
                            found = True
                            mdata(it) = ca.ToObject(lastValue)
                            Exit For
                        End If
                    Next
                    If found = False Then
                        For it = 0 To minfo.Count - 1
                            If minfo(it).Name.Equals(lastKey, StringComparison.InvariantCultureIgnoreCase) Then
                                mdata(it) = System.Convert.ChangeType(lastValue, DirectCast(minfo(it), FieldInfo).FieldType)
                                Exit For
                            End If
                        Next
                    End If
                    ca = Nothing
                End If
            End While
        End Using

        retval = FormatterServices.PopulateObjectMembers(retval, minfo, mdata)
        Erase minfo : minfo = Nothing
        Erase mdata : mdata = Nothing
        Return retval
    End Function


    'Private Shared Sub InitializePropWithDefaultValue(O As Object, Prop As PropertyInfo)
    '    Dim ca As Attribute = Prop.GetCustomAttribute(GetType(IniValueAttribute))
    '    Dim pi As PropertyInfo = ca.GetType.GetProperty("Default")
    '    If pi IsNot Nothing Then
    '        Prop.SetValue(O, pi.GetValue(ca))
    '    End If
    '    pi = Nothing
    '    ca = Nothing
    'End Sub

End Class
Public Enum IniSerializationType As Byte
    [String] = 0
    [Integer]
    [Double]
    [Boolean]
    [Date]
    [Dictionary]
    Invalid = 255
End Enum
<AttributeUsage(AttributeTargets.Field)> Public Class IniValueAttribute
    Inherits Attribute
    Private _sect As String, _key As String, _type As IniSerializationType
    Public Property Section As String
        Get
            Return _sect
        End Get
        Set(value As String)
            _sect = Regex.Replace(value.Trim, "\s+", " ")
        End Set
    End Property
    Public Property Key As String
        Get
            Return _key
        End Get
        Set(value As String)
            _key = Regex.Replace(value.Trim, "\s+", "_")
        End Set
    End Property

    Sub New(Key As String, Type As IniSerializationType)
        Me.New(String.Empty, Key, Type)
    End Sub
    Sub New(Section As String, Key As String, Type As IniSerializationType)
        If String.IsNullOrEmpty(Key) Then Throw New ArgumentNullException
        Me._sect = Section
        Me._key = Key
        Me._type = Type
    End Sub
    Public Function ToObject(Text As String) As Object
        Select Case _type
            Case IniSerializationType.String
                Return Text
            Case IniSerializationType.Integer
                Dim retval As Integer
                If Integer.TryParse(Text, retval) Then Return retval
                Return 0
            Case IniSerializationType.Double
                Dim retval As Double
                If Double.TryParse(Text, retval) Then Return retval
                Return 0
            Case IniSerializationType.Boolean
                Dim retval As Boolean
                If Boolean.TryParse(Text, retval) Then Return retval
                Return False
            Case IniSerializationType.Date
                Dim retval As Date
                If Date.TryParse(Text, retval) Then Return retval
                Return #12:00:00 AM#
            Case IniSerializationType.Dictionary
                Dim retval As New Dictionary(Of String, String)
                For Each l As String In Text.Split(Chr(13))
                    Dim ioc As Integer = l.IndexOf(":"c)
                    If ioc > -1 Then
                        retval.Add(l.Substring(0, ioc).Trim, l.Substring(ioc + 1).Trim)
                    End If
                Next
                Return retval
        End Select
        Return Nothing
    End Function
    Public Overloads Function ToString(Value As Object) As String
        Select Case _type
            Case IniSerializationType.String
                If Value Is Nothing Then Return String.Empty
                Return CStr(Value)
            Case IniSerializationType.Integer
                If Value Is Nothing Then Return "0"
                Return CInt(Value).ToString
            Case IniSerializationType.Double
                If Value Is Nothing Then Return "0"
                Return CDbl(Value).ToString
            Case IniSerializationType.Boolean
                If Value Is Nothing Then Return "False"
                Return CBool(Value).ToString
            Case IniSerializationType.Date
                If Value Is Nothing Then Return String.Empty
                Return CDate(Value).ToString("dd.MM.yyyy HH:mm")
            Case IniSerializationType.Dictionary
                If Value Is Nothing Then Return String.Empty
                Dim retval As New System.Text.StringBuilder
                Dim e As IDictionaryEnumerator = DirectCast(Value, IDictionary).GetEnumerator
                While e.MoveNext
                    retval.AppendLine(e.Key.ToString & ":" & e.Value.ToString)
                End While
                e = Nothing
                Return retval.ToString.Trim
        End Select
        Return String.Empty
    End Function
End Class





