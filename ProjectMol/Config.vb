Imports System.Configuration
Public Class Config
    Public Property smtpServer As String
    Public Property smtpPort As Int32
    Public Property smtpTimeOut As Int32
    Public Property sender As String
    Public Property senderPassword As String
    Public Property recipients As List(Of String)
    Public Property cc As List(Of String)
    Public Property subject As String
    Public Property body As String
    Public Property attachmentDir As List(Of String)

    Public Shared Function FromConfigFile() As Config
        Dim config = New Config
        Dim settings = ConfigurationManager.AppSettings
        config.smtpServer = settings("smtpServer")
        config.smtpPort = settings("smtpPort")
        config.smtpTimeOut = settings("smtpTimeOut")
        config.sender = settings("sender")
        config.senderPassword = settings("senderPassword")

        config.recipients = New List(Of String)
        config.recipients.AddRange(config.ExtractEmails(settings("recipients")))

        config.cc = New List(Of String)
        config.cc.AddRange(config.ExtractEmails(settings("cc")))

        config.attachmentDir = New List(Of String)
        config.attachmentDir.AddRange(config.ExtractEmails(settings("attachmentDir")))
        Return config
    End Function

    Public Function ExtractEmails(ByVal value As String) As List(Of String)
        Dim v = value.Split(";")
        Return v.ToList()
    End Function

End Class
