
Imports System.Net.Mail
Imports System.Security.AccessControl
Imports System.ComponentModel
Imports System.IO

'Imports System.Web.Mail

Module EdiSender
    
    Sub Main()
        WriteToLog(Environment.NewLine)
        WriteToLog(DateTime.Now.ToString)

        WriteToLog("Starting ProjectMol EdiSender..........")
        Dim configs = Config.FromConfigFile()
        Try

            Dim attachements As New List(Of String)
            For Each item In configs.attachmentDir
                Dim attachment = GetFileInDirectory(item.Trim, ".txt", "n")
                If Not String.IsNullOrEmpty(attachment) Then
                    attachements.Add(attachment)
                End If

            Next

            For Each item In attachements
                WriteToLog("Attachment File : " + item)
            Next

            If attachements.Count = 0 Then
                WriteToLog("No attachments were found. Exiting mail sending process")
                Environment.Exit(0)
                Return
            End If

            Using smptClient As New SmtpClient()
                smptClient.Credentials = New Net.NetworkCredential(configs.senderPassword, configs.senderPassword)
                smptClient.Port = CInt(configs.smtpPort)
                smptClient.Host = configs.smtpServer
                smptClient.EnableSsl = True
                smptClient.Timeout = configs.smtpTimeOut


                Using mail As New MailMessage()

                    mail.From = New MailAddress(configs.sender)
                    For Each item In configs.recipients
                        mail.To.Add(item.Trim)
                    Next
                    mail.Subject = configs.subject
                    mail.Body = configs.body
                    For Each item In configs.cc
                        mail.CC.Add(item.Trim)
                    Next
                    mail.IsBodyHtml = True
                    For Each item In attachements
                        mail.Attachments.Add(New Attachment(item))
                    Next
                    smptClient.Send(mail)
                    WriteToLog("Message sent.")

                End Using
            End Using

            For Each item In attachements
                CopyFileIntoFolder("archive", item)
                DeleteFile(item)
            Next


        Catch ex As Exception
            WriteToLog(ex.Message)
            WriteToLog(ex.StackTrace)
        Finally

        End Try

        Environment.Exit(0)

    End Sub



    Public Function CreateLogStream() As StreamWriter
        Dim d = DateTime.Now.Date.ToString("yyyMMdd")
        Dim info = New IO.DirectoryInfo("./logs/" + d)
        If info.Exists = False Then
            info.Create()
        End If
        Dim file = New IO.FileInfo(info.FullName + "/log.txt")
        If file.Exists = False Then
            file.Create()
        End If
        Dim stream = IO.File.AppendText(file.FullName)
        Return stream
       End Function

    Public Sub WriteToLog(ByVal message As String)
        Dim oldOut = Console.Out
        Using stream = CreateLogStream()
            Console.SetOut(stream)
            Console.WriteLine(message)
        End Using
        Console.SetOut(oldOut)
    End Sub
    Public Function GetFileInDirectory(ByVal directoryName As String, ByVal fileExtension As String, ByVal partern As String) As String

        Dim directory = New IO.DirectoryInfo(directoryName)
        For Each file In directory.GetFiles
            Dim filePath As String = file.FullName
            Dim fileName As String = file.Name
            Dim extension As String = file.Extension
            If extension.Equals(fileExtension) And fileName.StartsWith(partern) Then
                Return filePath 'And InStr(fileName, partern)
            End If
        Next
        Return ""
    End Function

    Public Sub CopyFileIntoFolder(ByVal destinationFolder As String, ByVal file As String)
        WriteToLog("Copying " + file + " to " + destinationFolder)
        Dim fileInfo = New IO.FileInfo(file)

        If fileInfo.Exists = False Then
            WriteToLog("The specified file does not exist")
            Return
        End If
        Dim parentDirectory = fileInfo.Directory.Parent
        Dim finalDirectory = parentDirectory.FullName + "\" + destinationFolder + "\"
        Dim finalDirInfo = New IO.DirectoryInfo(finalDirectory)
        If finalDirInfo.Exists = False Then
            finalDirInfo.Create()
        End If
        fileInfo.CopyTo(finalDirectory + fileInfo.Name, True)
    End Sub
    Public Sub DeleteFile(ByVal file As String)
        WriteToLog("Deleting file " + file)
        Dim fileInfo = New IO.FileInfo(file)
        If fileInfo.Exists = False Then
            WriteToLog("The specified file does not exist")
            Return
        End If
        fileInfo.Delete()
    End Sub
End Module
